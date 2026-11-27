function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sales Performance Dashboard 2026')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ============================================================
// ANALYTICS ROUTING + SAFE REFRESH — fungsi inti dipertahankan di Code.gs
// ============================================================
const ANALYTICS_BUILD_PREFIX_ = 'MANTUL_ANALYTICS_BUILD_';
const P1_FACT_HEADERS_ = ['YEAR','MONTH','SUBDIST','SUBBRAND','BRAND','CATEGORY','SBU','PIC','PF_MONTHLY','PF_QUARTERLY','VAL','SALES_TYPE','TARGET','BE'];

function analyticsAssertHeaders_(sheet, expected) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < expected.length) {
    throw new Error('Sheet ' + (sheet ? sheet.getName() : '(tidak ada)') + ' tidak memiliki header yang diperlukan.');
  }
  const normalize = value => String(value == null ? '' : value).trim().toUpperCase().replace(/[ _-]+/g, ' ');
  const actual = sheet.getRange(1, 1, 1, expected.length).getDisplayValues()[0].map(normalize);
  const wanted = expected.map(normalize);
  if (!actual.every((value, index) => value === wanted[index])) {
    throw new Error('Header ' + sheet.getName() + ' tidak sesuai. Diharapkan: ' + expected.join(' | ') + '. Terbaca: ' + actual.join(' | '));
  }
}

function analyticsGetReadySheet_(analyticsSheetName, sourceSheetName, phase) {
  const props = PropertiesService.getScriptProperties();
  const prefix = ANALYTICS_BUILD_PREFIX_;
  if (props.getProperty(prefix + 'STATUS') !== 'COMPLETE') return null;
  const ss = getDatabase_();
  if (ss.getId() !== String(props.getProperty(prefix + 'SOURCE_ID') || '')) return null;
  if (analyticsSheetName !== sourceSheetName) return null;
  const sheet = ss.getSheetByName(analyticsSheetName);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const expected = Number(props.getProperty(prefix + phase + '_WRITTEN') || 0);
  return sheet.getLastRow() - 1 === expected ? sheet : null;
}

function mantulV2ValidateLiveDatabase_() {
  const ss = getDatabase_();
  const p1 = ss.getSheetByName('P1_FACT');
  const raw = ss.getSheetByName(RAW_DATA_SHEET_NAME);
  if (!p1 || p1.getLastRow() < 2) throw new Error('P1_FACT belum tersedia atau belum berisi data.');
  if (!raw || raw.getLastRow() < 2) throw new Error('RAW_DATA_SALES belum tersedia atau belum berisi data.');
  analyticsAssertHeaders_(p1, P1_FACT_HEADERS_);
  validateRawDatabaseHeaders_(raw.getRange(1, 1, 1, RAW_SALES_COLUMN_COUNT).getDisplayValues()[0]);
  return {ok:true, version:'RAW_DIRECT', spreadsheetId:ss.getId(), sheets:{
    P1_FACT:{rows:p1.getLastRow()-1,columns:p1.getLastColumn()},
    RAW_DATA_SALES:{rows:raw.getLastRow()-1,columns:raw.getLastColumn()}
  }};
}

function refreshMantulDashboard() {
  const props = PropertiesService.getScriptProperties();
  const oldSales = props.getProperty('SALES_LAST_UPDATED') || '';
  const oldP2 = props.getProperty('PAGE2_RAW_UPDATED') || '';
  const started = Date.now();
  try {
    const validation = mantulV2ValidateLiveDatabase_();
    const readIndex = refreshMantulReadIndexMonthly_();
    let materializedFacts = {ok:true,status:'SKIPPED'};
    if (readIndex.status === 'UPDATED' || readIndex.status === 'REBUILT') {
      // Existing FACT writers are full-build-safe but their append-only path is not
      // overwrite-safe. Invalidate readiness and rebuild both facts from the verified
      // A:Y index so stale rows can never be reported COMPLETE.
      props.setProperties({PAGE2_FACT_STATUS:'STALE',PAGE34_FACT_STATUS:'STALE'});
    }
    try {
      let p2 = refreshPage2OutletFactIncremental_();
      if (p2.status === 'NEEDS_FULL_BUILD') p2 = buildPage2OutletFact();
      let p34 = refreshPage34EventFactIncremental_();
      if (p34.status === 'NEEDS_FULL_BUILD') p34 = buildPage34EventFact();
      const p2Valid = p2.ok && (p2.validated === true || p2.status === 'UNCHANGED');
      const p34Valid = p34.ok && (p34.validated === true || p34.status === 'UNCHANGED');
      if (!p2Valid || !p34Valid) {
        throw new Error('Satu atau lebih materialized FACT belum tervalidasi COMPLETE.');
      }
      materializedFacts = {ok:true,page2:p2,page34:p34};
    } catch (factErr) {
      props.setProperties({PAGE2_FACT_STATUS:'FAILED',PAGE34_FACT_STATUS:'FAILED'});
      throw factErr;
    }
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
    props.setProperty('SALES_LAST_UPDATED', stamp);
    props.setProperty('PAGE2_RAW_UPDATED', stamp);
    props.deleteProperty('MANTUL_V2_DATA_STALE');
    cacheDashboardData_();
    const result = {ok:true,status:'COMPLETE',buildExecuted:false,
      message:'Refresh selesai. Page 1 tetap memakai P1_FACT; Page 2–4 memakai materialized facts tervalidasi dari RAW_DATA_SALES A:Y.',
      validation:validation,readIndex:readIndex,materializedFacts:materializedFacts,updatedAt:stamp,elapsedMs:Date.now()-started};
    Logger.log('MANTUL SAFE REFRESH\\n' + JSON.stringify(result, null, 2));
    return result;
  } catch (err) {
    if (oldSales) props.setProperty('SALES_LAST_UPDATED', oldSales); else props.deleteProperty('SALES_LAST_UPDATED');
    if (oldP2) props.setProperty('PAGE2_RAW_UPDATED', oldP2); else props.deleteProperty('PAGE2_RAW_UPDATED');
    Logger.log('MANTUL SAFE REFRESH FAILED: ' + (err && err.stack || err));
    throw err;
  }
}

function validateMantulDashboardSources() {
  return Object.assign({readOnly:true}, mantulV2ValidateLiveDatabase_());
}


// ============================================================
// MANTUL READ INDEX — snapshot terverifikasi RAW_DATA_SALES A:Y.
// Page 1/P1_FACT tidak disentuh. Index dibangun secara staging,
// diverifikasi row count + total VAL, lalu dipublikasikan atomik.
// ============================================================
const MANTUL_READ_INDEX_SHEET_ = 'P2_EVENT_INDEX';
const MANTUL_READ_INDEX_STAGE_ = 'P2_EVENT_INDEX_BUILD';

function buildMantulReadIndex() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  const ss = getDatabase_();
  let stage = null;
  try {
    const raw = ss.getSheetByName(RAW_DATA_SHEET_NAME);
    if (!raw || raw.getLastRow() < 2) throw new Error('RAW_DATA_SALES kosong atau tidak ditemukan.');
    validateRawDatabaseHeaders_(raw.getRange(1, 1, 1, RAW_SALES_COLUMN_COUNT).getDisplayValues()[0]);
    if (raw.getLastColumn() < RAW_SALES_COLUMN_COUNT) {
      throw new Error('RAW_DATA_SALES minimal harus memiliki 25 kolom A:Y. Terbaca ' + raw.getLastColumn() + ' kolom. Kolom setelah Y diabaikan oleh read index.');
    }

    const rowCount = raw.getLastRow() - 1;
    const headers = raw.getRange(1, 1, 1, RAW_SALES_COLUMN_COUNT).getValues();
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const stageName = MANTUL_READ_INDEX_STAGE_;
    const oldStage = ss.getSheetByName(stageName);
    if (oldStage) ss.deleteSheet(oldStage);
    stage = ss.insertSheet(stageName);
    stage.getRange(1, 1, 1, RAW_SALES_COLUMN_COUNT).setValues(headers);

    let sourceValue = 0, indexValue = 0, copied = 0;
    const batchSize = 5000;
    for (let offset = 0; offset < rowCount; offset += batchSize) {
      const count = Math.min(batchSize, rowCount - offset);
      const values = raw.getRange(offset + 2, 1, count, RAW_SALES_COLUMN_COUNT).getValues();
      for (let i = 0; i < values.length; i++) sourceValue += Number(values[i][RAW_SALES_COL.VALUE]) || 0;
      stage.getRange(offset + 2, 1, count, RAW_SALES_COLUMN_COUNT).setValues(values);
      copied += count;
    }
    SpreadsheetApp.flush();

    if (copied !== rowCount || stage.getLastRow() !== rowCount + 1 || stage.getLastColumn() !== RAW_SALES_COLUMN_COUNT) {
      throw new Error('Validasi indeks gagal: jumlah baris/kolom tidak cocok.');
    }
    for (let offset = 0; offset < rowCount; offset += batchSize) {
      const count = Math.min(batchSize, rowCount - offset);
      const values = stage.getRange(offset + 2, 1, count, RAW_SALES_COLUMN_COUNT).getValues();
      for (let i = 0; i < values.length; i++) indexValue += Number(values[i][RAW_SALES_COL.VALUE]) || 0;
    }
    if (Math.abs(sourceValue - indexValue) > Math.max(0.01, Math.abs(sourceValue) * 1e-10)) {
      throw new Error('Validasi total VAL gagal. RAW=' + sourceValue + '; INDEX=' + indexValue);
    }

    const current = ss.getSheetByName(MANTUL_READ_INDEX_SHEET_);
    const backupName = 'P2_EVENT_INDEX_OLD_' + stamp;
    if (current) current.setName(backupName);
    try {
      stage.setName(MANTUL_READ_INDEX_SHEET_);
      stage.setFrozenRows(1);
      PropertiesService.getScriptProperties().setProperties({
        MANTUL_READ_INDEX_STATUS: 'COMPLETE',
        MANTUL_READ_INDEX_SOURCE_ID: ss.getId(),
        MANTUL_READ_INDEX_SOURCE_ROWS: String(rowCount),
        MANTUL_READ_INDEX_SOURCE_LASTROW: String(raw.getLastRow()),
        MANTUL_READ_INDEX_VAL: String(indexValue),
        MANTUL_READ_INDEX_BUILT_AT: new Date().toISOString()
      });
      if (current) ss.deleteSheet(ss.getSheetByName(backupName));
    } catch (publishError) {
      const published = ss.getSheetByName(MANTUL_READ_INDEX_SHEET_);
      if (published && published.getSheetId() === stage.getSheetId()) ss.deleteSheet(published);
      const backup = ss.getSheetByName(backupName);
      if (backup && !ss.getSheetByName(MANTUL_READ_INDEX_SHEET_)) backup.setName(MANTUL_READ_INDEX_SHEET_);
      throw publishError;
    }
    const result = {ok:true,status:'COMPLETE',sheet:MANTUL_READ_INDEX_SHEET_,rows:rowCount,columns:RAW_SALES_COLUMN_COUNT,sourceValue:sourceValue,indexValue:indexValue,validated:true};
    Logger.log('MANTUL READ INDEX BUILD\\n' + JSON.stringify(result, null, 2));
    return result;
  } catch (err) {
    if (stage) {
      const stillThere = ss.getSheetByName(MANTUL_READ_INDEX_STAGE_);
      if (stillThere && stillThere.getSheetId() === stage.getSheetId()) ss.deleteSheet(stillThere);
    }
    PropertiesService.getScriptProperties().setProperty('MANTUL_READ_INDEX_STATUS', 'FAILED');
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Refresh read index append-only. Asumsi: histori lama tidak diedit,
 * transaksi baru selalu ditambahkan di bawah. Jika tidak valid, full rebuild.
 */
/**
 * Monthly-tail synchronizer for RAW_DATA_SALES A:Y.
 * Only reads the contiguous tail month to fingerprint changes. First run or
 * unprovable boundaries fall back to the validated full index builder.
 */
function mantulTailPartition_(sheet) {
  const last=sheet.getLastRow();
  if(last<2)throw new Error(sheet.getName()+' tidak memiliki data.');
  const tail=sheet.getRange(last,1,1,2).getValues()[0];
  const year=String(tail[0] instanceof Date?tail[0].getFullYear():tail[0]).trim();
  const month=page2Month_(tail[1]);
  if(!year||!month)throw new Error('YEAR/MONTH baris terakhir tidak valid pada '+sheet.getName()+'.');
  let start=last;
  const block=4000;
  while(start>2){
    const from=Math.max(2,start-block),n=start-from;
    const pairs=sheet.getRange(from,1,n,2).getValues();
    let found=-1;
    for(let i=pairs.length-1;i>=0;i--){
      const y=String(pairs[i][0] instanceof Date?pairs[i][0].getFullYear():pairs[i][0]).trim();
      const m=page2Month_(pairs[i][1]);
      if(y===year&&m===month)found=i;
      else if(found>=0)break;
      else break;
    }
    if(found<0)break;
    start=from+found;
    if(found>0)break;
  }
  const n=last-start+1,values=sheet.getRange(start,1,n,RAW_SALES_COLUMN_COUNT).getValues();
  for(let i=0;i<values.length;i++){
    const y=String(values[i][0] instanceof Date?values[i][0].getFullYear():values[i][0]).trim();
    if(y!==year||page2Month_(values[i][1])!==month)throw new Error('Blok bulan terakhir tidak kontigu pada '+sheet.getName()+'.');
  }
  const digest=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(values));
  return {year:year,month:month,startRow:start,lastRow:last,rows:n,values:values,
    hash:Utilities.base64Encode(digest),value:values.reduce((s,r)=>s+(Number(r[RAW_SALES_COL.VALUE])||0),0)};
}

function refreshMantulReadIndexMonthly_() {
  const lock=LockService.getScriptLock();lock.waitLock(30000);let lockHeld=true;
  try{
    const ss=getDatabase_(),raw=ss.getSheetByName(RAW_DATA_SHEET_NAME),index=ss.getSheetByName(MANTUL_READ_INDEX_SHEET_);
    if(!raw||raw.getLastRow()<2)throw new Error('RAW_DATA_SALES kosong atau tidak ditemukan.');
    validateRawDatabaseHeaders_(raw.getRange(1,1,1,RAW_SALES_COLUMN_COUNT).getDisplayValues()[0]);
    if(raw.getLastColumn()<RAW_SALES_COLUMN_COUNT)throw new Error('RAW_DATA_SALES harus memiliki kolom A:Y.');
    const tail=mantulTailPartition_(raw),props=PropertiesService.getScriptProperties();
    const ready=index&&props.getProperty('MANTUL_READ_INDEX_STATUS')==='COMPLETE'&&
      props.getProperty('MANTUL_READ_INDEX_SOURCE_ID')===ss.getId()&&
      Number(props.getProperty('MANTUL_READ_INDEX_SOURCE_LASTROW'))===index.getLastRow()&&
      Number(props.getProperty('MANTUL_READ_INDEX_TAIL_START'))>=2&&
      props.getProperty('MANTUL_READ_INDEX_TAIL_PERIOD');
    if(!ready){
      lock.releaseLock();lockHeld=false;
      const built=buildMantulReadIndex();
      const current=ss.getSheetByName(MANTUL_READ_INDEX_SHEET_);
      const saved=mantulTailPartition_(current);
      props.setProperties({MANTUL_READ_INDEX_TAIL_START:String(saved.startRow),
        MANTUL_READ_INDEX_TAIL_PERIOD:saved.year+'-'+saved.month,
        MANTUL_READ_INDEX_TAIL_HASH:saved.hash,MANTUL_READ_INDEX_TAIL_VAL:String(saved.value)});
      return Object.assign({},built,{status:'REBUILT',tailPeriod:saved.year+'-'+saved.month,tailRows:saved.rows});
    }
    const oldStart=Number(props.getProperty('MANTUL_READ_INDEX_TAIL_START'));
    const oldPeriod=props.getProperty('MANTUL_READ_INDEX_TAIL_PERIOD');
    const oldTailValue=Number(props.getProperty('MANTUL_READ_INDEX_TAIL_VAL')||0);
    const oldHash=props.getProperty('MANTUL_READ_INDEX_TAIL_HASH')||'';
    const indexTail=mantulTailPartition_(index);
    if(indexTail.startRow!==oldStart||indexTail.year+'-'+indexTail.month!==oldPeriod)
      throw new Error('Batas partisi INDEX tidak sesuai metadata; jalankan rebuild index terkontrol.');
    if(tail.year+'-'+tail.month===oldPeriod&&tail.hash===oldHash&&raw.getLastRow()===index.getLastRow()){
      return {ok:true,status:'UNCHANGED',sheet:MANTUL_READ_INDEX_SHEET_,rows:raw.getLastRow()-1,
        tailPeriod:oldPeriod,tailRows:tail.rows,indexValue:Number(props.getProperty('MANTUL_READ_INDEX_VAL')||0)};
    }
    const samePeriod=tail.year+'-'+tail.month===oldPeriod;
    if(!samePeriod&&tail.startRow!==index.getLastRow()+1)
      throw new Error('Bulan baru tidak dimulai tepat setelah partisi INDEX terakhir; rebuild diperlukan.');
    const writeStart=samePeriod?oldStart:tail.startRow;
    if(samePeriod&&tail.startRow!==oldStart)
      throw new Error('Awal partisi bulan berjalan berubah; rebuild diperlukan demi keamanan.');
    const oldLast=index.getLastRow();
    if(samePeriod&&oldLast>=writeStart)index.getRange(writeStart,1,oldLast-writeStart+1,RAW_SALES_COLUMN_COUNT).clearContent();
    if(tail.rows>0)index.getRange(writeStart,1,tail.rows,RAW_SALES_COLUMN_COUNT).setValues(tail.values);
    SpreadsheetApp.flush();
    if(index.getLastRow()!==raw.getLastRow())throw new Error('Validasi jumlah baris INDEX setelah refresh gagal.');
    const check=mantulTailPartition_(index);
    if(check.rows!==tail.rows||check.hash!==tail.hash||Math.abs(check.value-tail.value)>Math.max(.01,Math.abs(tail.value)*1e-10))
      throw new Error('Validasi fingerprint/VAL partisi INDEX gagal.');
    let totalValue=Number(props.getProperty('MANTUL_READ_INDEX_VAL')||0);
    totalValue=samePeriod?totalValue-oldTailValue+tail.value:totalValue+tail.value;
    props.setProperties({
      MANTUL_READ_INDEX_STATUS:'COMPLETE',MANTUL_READ_INDEX_SOURCE_ID:ss.getId(),
      MANTUL_READ_INDEX_SOURCE_ROWS:String(raw.getLastRow()-1),MANTUL_READ_INDEX_SOURCE_LASTROW:String(raw.getLastRow()),
      MANTUL_READ_INDEX_VAL:String(totalValue),MANTUL_READ_INDEX_BUILT_AT:new Date().toISOString(),
      MANTUL_READ_INDEX_TAIL_START:String(tail.startRow),MANTUL_READ_INDEX_TAIL_PERIOD:tail.year+'-'+tail.month,
      MANTUL_READ_INDEX_TAIL_HASH:tail.hash,MANTUL_READ_INDEX_TAIL_VAL:String(tail.value)
    });
    return {ok:true,status:'UPDATED',sheet:MANTUL_READ_INDEX_SHEET_,tailPeriod:tail.year+'-'+tail.month,
      tailRows:tail.rows,previousTailRows:oldLast-oldStart+1,sourceRows:raw.getLastRow()-1,indexValue:totalValue,validated:true};
  }finally{if(lockHeld)lock.releaseLock();}
}

function refreshMantulReadIndexIncremental_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let lockHeld = true;
  const ss = getDatabase_();
  const props = PropertiesService.getScriptProperties();
  try {
    const raw = ss.getSheetByName(RAW_DATA_SHEET_NAME);
    if (!raw || raw.getLastRow() < 2) throw new Error('RAW_DATA_SALES kosong atau tidak ditemukan.');
    if (raw.getLastColumn() < RAW_SALES_COLUMN_COUNT) throw new Error('RAW_DATA_SALES minimal harus memiliki 25 kolom A:Y. Terbaca ' + raw.getLastColumn() + ' kolom. Kolom setelah Y diabaikan oleh read index.');
    validateRawDatabaseHeaders_(raw.getRange(1, 1, 1, RAW_SALES_COLUMN_COUNT).getDisplayValues()[0]);

    const index = ss.getSheetByName(MANTUL_READ_INDEX_SHEET_);
    const previousLastRow = Number(props.getProperty('MANTUL_READ_INDEX_SOURCE_LASTROW') || 0);
    const previousRows = Number(props.getProperty('MANTUL_READ_INDEX_SOURCE_ROWS') || 0);
    const previousValue = Number(props.getProperty('MANTUL_READ_INDEX_VAL') || 0);
    const ready = props.getProperty('MANTUL_READ_INDEX_STATUS') === 'COMPLETE' &&
      props.getProperty('MANTUL_READ_INDEX_SOURCE_ID') === ss.getId() &&
      index && previousLastRow >= 2 && index.getLastRow() === previousLastRow &&
      previousRows === previousLastRow - 1;

    if (!ready || raw.getLastRow() < previousLastRow) {
      lock.releaseLock(); lockHeld = false;
      return buildMantulReadIndex();
    }
    if (raw.getLastRow() === previousLastRow) {
      return {ok:true,status:'UNCHANGED',sheet:MANTUL_READ_INDEX_SHEET_,rows:previousRows,
        columns:RAW_SALES_COLUMN_COUNT,indexValue:previousValue,appendedRows:0};
    }

    const appendCount = raw.getLastRow() - previousLastRow;
    const batchSize = 4000;
    let appendedValue = 0, verifiedValue = 0, written = 0;
    try {
      for (let offset = 0; offset < appendCount; offset += batchSize) {
        const count = Math.min(batchSize, appendCount - offset);
        const values = raw.getRange(previousLastRow + offset + 1, 1, count, RAW_SALES_COLUMN_COUNT).getValues();
        values.forEach(row => appendedValue += Number(row[RAW_SALES_COL.VALUE]) || 0);
        index.getRange(previousLastRow + offset + 1, 1, count, RAW_SALES_COLUMN_COUNT).setValues(values);
        written += count;
      }
      SpreadsheetApp.flush();
      if (written !== appendCount || index.getLastRow() !== raw.getLastRow() || index.getLastColumn() !== RAW_SALES_COLUMN_COUNT) {
        throw new Error('Validasi append index gagal: jumlah baris/kolom tidak cocok.');
      }
      for (let offset = 0; offset < appendCount; offset += batchSize) {
        const count = Math.min(batchSize, appendCount - offset);
        index.getRange(previousLastRow + offset + 1, 1, count, RAW_SALES_COLUMN_COUNT).getValues()
          .forEach(row => verifiedValue += Number(row[RAW_SALES_COL.VALUE]) || 0);
      }
      if (Math.abs(appendedValue - verifiedValue) > Math.max(0.01, Math.abs(appendedValue) * 1e-10)) {
        throw new Error('Validasi total VAL baris baru gagal. RAW=' + appendedValue + '; INDEX=' + verifiedValue);
      }
      const totalValue = previousValue + verifiedValue;
      props.setProperties({
        MANTUL_READ_INDEX_STATUS:'COMPLETE', MANTUL_READ_INDEX_SOURCE_ID:ss.getId(),
        MANTUL_READ_INDEX_SOURCE_ROWS:String(raw.getLastRow()-1),
        MANTUL_READ_INDEX_SOURCE_LASTROW:String(raw.getLastRow()),
        MANTUL_READ_INDEX_VAL:String(totalValue), MANTUL_READ_INDEX_BUILT_AT:new Date().toISOString()
      });
      return {ok:true,status:'APPENDED',sheet:MANTUL_READ_INDEX_SHEET_,rows:raw.getLastRow()-1,
        columns:RAW_SALES_COLUMN_COUNT,appendedRows:appendCount,appendedValue:appendedValue,
        indexValue:totalValue,validated:true};
    } catch (err) {
      const last=index.getLastRow();
      if(last>previousLastRow) index.getRange(previousLastRow+1,1,last-previousLastRow,RAW_SALES_COLUMN_COUNT).clearContent();
      throw err;
    }
  } finally {
    if (lockHeld) lock.releaseLock();
  }
}

function mantulGetReadyReadIndex_() {
  const ss = getDatabase_();
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('MANTUL_READ_INDEX_STATUS') !== 'COMPLETE' ||
      props.getProperty('MANTUL_READ_INDEX_SOURCE_ID') !== ss.getId()) return null;
  const raw = ss.getSheetByName(RAW_DATA_SHEET_NAME);
  const index = ss.getSheetByName(MANTUL_READ_INDEX_SHEET_);
  if (!raw || !index || raw.getLastRow() !== Number(props.getProperty('MANTUL_READ_INDEX_SOURCE_LASTROW')) ||
      index.getLastRow() !== raw.getLastRow() || index.getLastColumn() !== RAW_SALES_COLUMN_COUNT) return null;
  return index;
}


/**
 * AUDIT READ-ONLY: membandingkan RAW, read index, dan materialized facts
 * per bulan. Tidak menulis sheet, Script Properties, cache, atau trigger.
 * Jalankan saat angka sumber dan dashboard tidak cocok.
 */
function auditMantulDashboardSync() {
  const started = Date.now();
  const ss = getDatabase_();
  const props = PropertiesService.getScriptProperties();
  const raw = ss.getSheetByName(RAW_DATA_SHEET_NAME);
  const index = ss.getSheetByName(MANTUL_READ_INDEX_SHEET_);
  const p2 = ss.getSheetByName('P2_OUTLET_FACT');
  const p34 = ss.getSheetByName('P34_EVENT_FACT');
  if (!raw) throw new Error('RAW_DATA_SALES tidak ditemukan.');

  function summarize(sheet, valueCol, sourceCountCol) {
    if (!sheet || sheet.getLastRow() < 2) return {exists:!!sheet, physicalRows:0, sourceRows:0, totalValue:0, latestDate:'', byMonth:{}};
    const last = sheet.getLastRow(), batch = 5000;
    const out = {exists:true, physicalRows:last-1, sourceRows:0, totalValue:0, latestDate:'', byMonth:{}};
    for (let start=2; start<=last; start+=batch) {
      const n=Math.min(batch,last-start+1);
      const dims=sheet.getRange(start,1,n,3).getValues();
      const vals=sheet.getRange(start,valueCol,n,1).getValues();
      const counts=sourceCountCol?sheet.getRange(start,sourceCountCol,n,1).getValues():null;
      for(let i=0;i<n;i++){
        const y=String(Number(dims[i][0])||dims[i][0]||'').trim();
        const m=page2Month_(dims[i][1]);
        const d=page2Day_(dims[i][2]);
        const v=Number(vals[i][0])||0;
        const c=counts?(Number(counts[i][0])||0):1;
        out.totalValue+=v;
        out.sourceRows+=c;
        if(y&&m){
          const key=y+'-'+String(PAGE2_MONTHS_.indexOf(m)+1).padStart(2,'0');
          if(!out.byMonth[key])out.byMonth[key]={rows:0,value:0};
          out.byMonth[key].rows+=c;
          out.byMonth[key].value+=v;
          const date=page2Date_(y,m,d);
          if(date&&date>out.latestDate)out.latestDate=date;
        }
      }
    }
    return out;
  }

  let revision='unavailable';
  try { revision=getDatabaseRevision_(); } catch(e) { revision='ERROR: '+String(e.message||e); }
  let triggers=[];
  try {
    triggers=ScriptApp.getProjectTriggers().map(t=>({
      handler:t.getHandlerFunction(),
      eventType:String(t.getEventType()),
      source:String(t.getTriggerSource())
    }));
  } catch(e) { triggers=[{error:String(e.message||e)}]; }

  const report={
    ok:true,readOnly:true,spreadsheetId:ss.getId(),elapsedMs:0,
    sourceRevision:revision,
    scriptProperties:{
      indexStatus:props.getProperty('MANTUL_READ_INDEX_STATUS')||'',
      indexSourceLastRow:Number(props.getProperty('MANTUL_READ_INDEX_SOURCE_LASTROW')||0),
      indexSourceRows:Number(props.getProperty('MANTUL_READ_INDEX_SOURCE_ROWS')||0),
      indexBuiltAt:props.getProperty('MANTUL_READ_INDEX_BUILT_AT')||'',
      page2FactStatus:props.getProperty('PAGE2_FACT_STATUS')||'',
      page2FactSourceRows:Number(props.getProperty('PAGE2_FACT_SOURCE_ROWS')||0),
      page2FactWritten:Number(props.getProperty('PAGE2_FACT_WRITTEN')||0),
      page2FactBuiltAt:props.getProperty('PAGE2_FACT_BUILT_AT')||'',
      page34FactStatus:props.getProperty('PAGE34_FACT_STATUS')||'',
      page34FactSourceRows:Number(props.getProperty('PAGE34_FACT_SOURCE_ROWS')||0),
      page34FactWritten:Number(props.getProperty('PAGE34_FACT_WRITTEN')||0),
      page34FactBuiltAt:props.getProperty('PAGE34_FACT_BUILT_AT')||''
    },
    sheets:{
      RAW_DATA_SALES:summarize(raw,13,null),
      P2_EVENT_INDEX:summarize(index,13,null),
      P2_OUTLET_FACT:summarize(p2,15,16),
      P34_EVENT_FACT:summarize(p34,15,19)
    },
    triggers:triggers
  };
  const rawNov=report.sheets.RAW_DATA_SALES.byMonth['2026-11']||{rows:0,value:0};
  report.november2026={
    raw:rawNov,
    index:report.sheets.P2_EVENT_INDEX.byMonth['2026-11']||{rows:0,value:0},
    page2:report.sheets.P2_OUTLET_FACT.byMonth['2026-11']||{rows:0,value:0},
    page34:report.sheets.P34_EVENT_FACT.byMonth['2026-11']||{rows:0,value:0},
    expectedRawValue:1988904755
  };
  report.elapsedMs=Date.now()-started;
  Logger.log('MANTUL READ-ONLY SYNC AUDIT\\n'+JSON.stringify(report,null,2));
  return report;
}
