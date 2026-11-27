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
  const props=PropertiesService.getScriptProperties();
  const oldSales=props.getProperty('SALES_LAST_UPDATED')||'';
  const oldP2=props.getProperty('PAGE2_RAW_UPDATED')||'';
  const started=Date.now();
  try {
    const validation=mantulV2ValidateLiveDatabase_();
    if(mantulBootstrapLoad_()){
      mantulStartOrResumeBootstrap_();
      return Object.assign({validation:validation},mantulContinueDashboardBuild_());
    }
    if(!mantulHasMonthlyRefreshBase_()){
      mantulStartOrResumeBootstrap_();
      return Object.assign({validation:validation},mantulContinueDashboardBuild_());
    }
    let readIndex;
    try{
      readIndex=refreshMantulReadIndexMonthly_();
    }catch(indexErr){
      const message=String(indexErr&&indexErr.message||indexErr);
      if(/INDEX dan RAW berbeda pada partisi bulan terakhir|Batas partisi INDEX tidak sesuai metadata|Awal partisi bulan berjalan berubah/.test(message)){
        props.setProperties({MANTUL_READ_INDEX_STATUS:'STALE',PAGE2_FACT_STATUS:'STALE',PAGE34_FACT_STATUS:'STALE'});
        mantulStartOrResumeBootstrap_(true);
        return Object.assign({validation:validation,repairReason:message},mantulContinueDashboardBuild_());
      }
      throw indexErr;
    }
    let materializedFacts={ok:true,status:'UNCHANGED'};
    if(readIndex.status==='UPDATED'||readIndex.status==='REBUILT'){
      const p2=refreshPage2OutletFactMonthly_(readIndex);
      const p34=refreshPage34EventFactMonthly_(readIndex);
      if(p2.status==='NEEDS_FULL_BUILD'||p34.status==='NEEDS_FULL_BUILD'){
        props.setProperties({PAGE2_FACT_STATUS:'STALE',PAGE34_FACT_STATUS:'STALE'});
        mantulStartOrResumeBootstrap_();
        return Object.assign({validation:validation,readIndex:readIndex,monthlyFacts:{page2:p2,page34:p34}},mantulContinueDashboardBuild_());
      }
      materializedFacts={ok:true,page2:p2,page34:p34};
    }
    const stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd MMM yyyy HH:mm');
    props.setProperty('SALES_LAST_UPDATED',stamp);
    props.setProperty('PAGE2_RAW_UPDATED',stamp);
    props.deleteProperty('MANTUL_V2_DATA_STALE');
    cacheDashboardData_();
    const result={ok:true,status:'COMPLETE',buildExecuted:false,
      message:'Refresh selesai. Hanya partisi bulan aktif yang disinkronkan ke INDEX dan FACT; Page 1 tetap memakai P1_FACT.',
      validation:validation,readIndex:readIndex,materializedFacts:materializedFacts,updatedAt:stamp,elapsedMs:Date.now()-started};
    Logger.log('MANTUL MONTHLY REFRESH\\n'+JSON.stringify(result,null,2));
    return result;
  }catch(err){
    if(oldSales)props.setProperty('SALES_LAST_UPDATED',oldSales);else props.deleteProperty('SALES_LAST_UPDATED');
    if(oldP2)props.setProperty('PAGE2_RAW_UPDATED',oldP2);else props.deleteProperty('PAGE2_RAW_UPDATED');
    Logger.log('MANTUL MONTHLY REFRESH FAILED: '+(err&&err.stack||err));
    throw err;
  }
}


function mantulHasMonthlyRefreshBase_(){
 const ss=getDatabase_(),p=PropertiesService.getScriptProperties();
 const ix=ss.getSheetByName(MANTUL_READ_INDEX_SHEET_),p2=ss.getSheetByName(PAGE2_FACT_SHEET_),p34=ss.getSheetByName(PAGE34_FACT_SHEET_);
 if(!ix||!p2||!p34||ix.getLastRow()<2||ix.getLastColumn()!==RAW_SALES_COLUMN_COUNT)return false;
 const sourceRows=ix.getLastRow()-1;
 if(p.getProperty('MANTUL_READ_INDEX_STATUS')!=='COMPLETE'||p.getProperty('MANTUL_READ_INDEX_SOURCE_ID')!==ss.getId()||
    Number(p.getProperty('MANTUL_READ_INDEX_SOURCE_LASTROW'))!==ix.getLastRow())return false;
 const h2=p2.getLastRow()>1&&p2.getLastColumn()>=PAGE2_FACT_HEADERS_.length?p2.getRange(1,1,1,PAGE2_FACT_HEADERS_.length).getDisplayValues()[0].map(page2HeaderKey_):[];
 const h34=p34.getLastRow()>1&&p34.getLastColumn()>=PAGE34_FACT_HEADERS_.length?p34.getRange(1,1,1,PAGE34_FACT_HEADERS_.length).getDisplayValues()[0].map(page2HeaderKey_):[];
 return p.getProperty('PAGE2_FACT_STATUS')==='COMPLETE'&&p.getProperty('PAGE2_FACT_SOURCE_ID')===ss.getId()&&
  Number(p.getProperty('PAGE2_FACT_SOURCE_ROWS'))===sourceRows&&Number(p.getProperty('PAGE2_FACT_WRITTEN'))===p2.getLastRow()-1&&
  PAGE2_FACT_HEADERS_.every((v,i)=>h2[i]===page2HeaderKey_(v))&&
  p.getProperty('PAGE34_FACT_STATUS')==='COMPLETE'&&p.getProperty('PAGE34_FACT_SOURCE_ID')===ss.getId()&&
  Number(p.getProperty('PAGE34_FACT_SOURCE_ROWS'))===sourceRows&&Number(p.getProperty('PAGE34_FACT_WRITTEN'))===p34.getLastRow()-1&&
  PAGE34_FACT_HEADERS_.every((v,i)=>h34[i]===page2HeaderKey_(v));
}
function mantulFactPeriod_(year,month){
 const y=String(Number(year)||year||'').trim(),m=page2Month_(month);
 return y&&m?y+'-'+String(PAGE2_MONTHS_.indexOf(m)+1).padStart(2,'0'):'';
}
function mantulDiscoverFactTail_(sheet,period,sourceCol,valueCol,cartonCol){
 const last=sheet.getLastRow();if(last<2)throw new Error(sheet.getName()+' belum memiliki baris fact.');
 const pairs=sheet.getRange(2,1,last-1,2).getValues();let first=-1;
 for(let i=pairs.length-1;i>=0;i--){const p=mantulFactPeriod_(pairs[i][0],pairs[i][1]);if(p===period)first=i;else if(first>=0)break;}
 if(first<0)throw new Error('Partisi bulan aktif '+period+' tidak ditemukan di '+sheet.getName()+'.');
 for(let i=0;i<first;i++)if(mantulFactPeriod_(pairs[i][0],pairs[i][1])===period)throw new Error('Partisi '+period+' tidak kontigu di '+sheet.getName()+'.');
 const startRow=first+2,count=last-startRow+1,values=sheet.getRange(startRow,1,count,sheet.getLastColumn()).getValues();
 let sourceCount=0,value=0,carton=0;
 values.forEach(r=>{sourceCount+=Number(r[sourceCol])||0;value+=Number(r[valueCol])||0;if(cartonCol!=null)carton+=Number(r[cartonCol])||0;});
 return {period:period,startRow:startRow,rows:count,sourceCount:sourceCount,value:value,carton:carton};
}
function mantulWriteRows_(sheet,startRow,rows,columns){
 const batch=2000;for(let i=0;i<rows.length;i+=batch){const part=rows.slice(i,i+batch);sheet.getRange(startRow+i,1,part.length,columns).setValues(part);}
}
function mantulPublishMonthlyFact_(kind,tail,rows,newMetrics,indexRows){
 const isP2=kind==='P2',ss=getDatabase_(),props=PropertiesService.getScriptProperties();
 const target=isP2?PAGE2_FACT_SHEET_:PAGE34_FACT_SHEET_,stageName=isP2?'P2_OUTLET_FACT_MONTH_BUILD':'P34_EVENT_FACT_MONTH_BUILD';
 const headers=isP2?PAGE2_FACT_HEADERS_:PAGE34_FACT_HEADERS_,prefix=isP2?'PAGE2_FACT_':'PAGE34_FACT_',fact=ss.getSheetByName(target);
 if(!fact||fact.getLastRow()<2)return {ok:true,status:'NEEDS_FULL_BUILD',reason:'FACT_MISSING'};
 const oldStart=Number(props.getProperty(prefix+'TAIL_START')||0),oldPeriod=props.getProperty(prefix+'TAIL_PERIOD')||'';
 if(oldStart<2||!oldPeriod)return {ok:true,status:'NEEDS_FULL_BUILD',reason:'TAIL_METADATA_MISSING'};
 const newPeriod=tail.year+'-'+String(PAGE2_MONTHS_.indexOf(tail.month)+1).padStart(2,'0');
 if(newPeriod<oldPeriod)return {ok:true,status:'NEEDS_FULL_BUILD',reason:'HISTORICAL_MONTH_REQUIRES_REBUILD'};
 const replaceTail=newPeriod===oldPeriod;
 if(oldStart>fact.getLastRow())return {ok:true,status:'NEEDS_FULL_BUILD',reason:'TAIL_BOUNDARY_INVALID'};
 const oldTailSource=Number(props.getProperty(prefix+'TAIL_SOURCE_COUNT')||0),oldTailValue=Number(props.getProperty(prefix+'TAIL_VAL')||0),oldTailCarton=Number(props.getProperty(prefix+'TAIL_CARTON')||0);
 const prefixLast=replaceTail?oldStart-1:fact.getLastRow(),oldStage=ss.getSheetByName(stageName);
 if(oldStage)ss.deleteSheet(oldStage);
 const stage=ss.insertSheet(stageName);
 try{
  stage.getRange(1,1,1,headers.length).setValues([headers]);
  if(prefixLast>=1){
   const needed=prefixLast+rows.length;if(stage.getMaxRows()<needed)stage.insertRowsAfter(stage.getMaxRows(),needed-stage.getMaxRows());
   fact.getRange(1,1,prefixLast,headers.length).copyTo(stage.getRange(1,1),{contentsOnly:true});
  }
  if(rows.length)mantulWriteRows_(stage,prefixLast+1,rows,headers.length);
  SpreadsheetApp.flush();
  if(stage.getLastRow()!==prefixLast+rows.length)throw new Error('Jumlah baris staging '+target+' tidak sesuai.');
  let checkCount=0,checkValue=0,checkCarton=0;
  if(rows.length)stage.getRange(prefixLast+1,1,rows.length,headers.length).getValues().forEach(r=>{
   checkCount+=Number(r[isP2?15:18])||0;checkValue+=Number(r[14])||0;if(!isP2)checkCarton+=Number(r[15])||0;
  });
  const close=(a,b)=>Math.abs(a-b)<=Math.max(.01,Math.abs(a)*1e-10);
  if(checkCount!==newMetrics.sourceCount||!close(checkValue,newMetrics.value)||(!isP2&&!close(checkCarton,newMetrics.carton)))
   throw new Error('Validasi partisi bulan '+newPeriod+' pada '+target+' gagal.');
  const totalSource=Number(props.getProperty(prefix+'SOURCE_COUNT')||0)+(replaceTail?-oldTailSource:0)+checkCount;
  const totalValue=Number(props.getProperty(prefix+'VAL')||0)+(replaceTail?-oldTailValue:0)+checkValue;
  const totalCarton=isP2?0:Number(props.getProperty(prefix+'CARTON')||0)+(replaceTail?-oldTailCarton:0)+checkCarton;
  const newStart=prefixLast+1,stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd_HHmmss'),backup=target+'_OLD_'+stamp;
  props.setProperty(prefix+'STATUS','BUILDING');fact.setName(backup);
  try{
   stage.setName(target);stage.setFrozenRows(1);
   const values={
    [prefix+'STATUS']:'COMPLETE',[prefix+'SOURCE_ID']:ss.getId(),[prefix+'SOURCE_ROWS']:String(indexRows),
    [prefix+'WRITTEN']:String(stage.getLastRow()-1),[prefix+'SOURCE_COUNT']:String(totalSource),[prefix+'VAL']:String(totalValue),
    [prefix+'TAIL_START']:String(newStart),[prefix+'TAIL_PERIOD']:newPeriod,[prefix+'TAIL_SOURCE_COUNT']:String(checkCount),
    [prefix+'TAIL_VAL']:String(checkValue),[prefix+'BUILT_AT']:new Date().toISOString()
   };
   if(!isP2){values[prefix+'CARTON']=String(totalCarton);values[prefix+'TAIL_CARTON']=String(checkCarton);}
   props.setProperties(values);const old=ss.getSheetByName(backup);if(old)ss.deleteSheet(old);
  }catch(e){
   const live=ss.getSheetByName(target);if(live&&live.getSheetId()===stage.getSheetId())ss.deleteSheet(live);
   const back=ss.getSheetByName(backup);if(back&&!ss.getSheetByName(target))back.setName(target);
   props.setProperty(prefix+'STATUS','STALE');throw e;
  }
  return {ok:true,status:'UPDATED',sheet:target,period:newPeriod,sourceRows:checkCount,factRows:rows.length,sourceValue:checkValue,sourceCarton:isP2?undefined:checkCarton,validated:true};
 }catch(err){
  const leftover=ss.getSheetByName(stageName);if(leftover)ss.deleteSheet(leftover);
  if(props.getProperty(prefix+'STATUS')==='BUILDING')props.setProperty(prefix+'STATUS','STALE');
  throw err;
 }
}
function refreshPage2OutletFactMonthly_(indexRefresh){
 const index=mantulGetReadyReadIndex_(),tail=mantulTailPartition_(index),incoming=page2Read_(index,tail.startRow,tail.rows),grouped=new Map();
 let sourceCount=0,value=0;
 incoming.forEach(r=>{const key=JSON.stringify(r.slice(0,14));let item=grouped.get(key);if(!item){item=r.slice(0,14).concat([0,0]);grouped.set(key,item);}
  item[14]+=Number(r[14])||0;item[15]+=Number(r[15])||1;sourceCount+=Number(r[15])||1;value+=Number(r[14])||0;});
 return mantulPublishMonthlyFact_('P2',tail,Array.from(grouped.values()),{sourceCount:sourceCount,value:value},index.getLastRow()-1);
}
function refreshPage34EventFactMonthly_(indexRefresh){
 const index=mantulGetReadyReadIndex_(),tail=mantulTailPartition_(index),incoming=page34ReadRawIndex_(tail.startRow,tail.rows),grouped=new Map();
 let sourceCount=0,value=0,carton=0;
 incoming.rows.forEach(r=>{const key=JSON.stringify(r.slice(0,14).concat(r.slice(16,18)));let item=grouped.get(key);if(!item){item=r.slice();item[14]=0;item[15]=0;item[18]=0;grouped.set(key,item);}
  item[18]+=Number(r[18])||1;item[14]+=Number(r[14])||0;item[15]+=Number(r[15])||0;sourceCount+=Number(r[18])||1;value+=Number(r[14])||0;carton+=Number(r[15])||0;});
 return mantulPublishMonthlyFact_('P34',tail,Array.from(grouped.values()),{sourceCount:sourceCount,value:value,carton:carton},index.getLastRow()-1);
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
      // Never fall back to a full 130k-row rebuild from the foreground refresh.
      // If the existing index is structurally current, seed missing tail metadata
      // only after its final-month fingerprint matches RAW_DATA_SALES.
      const indexUsable=index&&index.getLastRow()===raw.getLastRow()&&
        index.getLastColumn()===RAW_SALES_COLUMN_COUNT&&
        props.getProperty('MANTUL_READ_INDEX_STATUS')==='COMPLETE'&&
        props.getProperty('MANTUL_READ_INDEX_SOURCE_ID')===ss.getId()&&
        Number(props.getProperty('MANTUL_READ_INDEX_SOURCE_LASTROW'))===raw.getLastRow();
      if(indexUsable){
        const indexTail=mantulTailPartition_(index);
        if(indexTail.hash!==tail.hash||indexTail.rows!==tail.rows)
          throw new Error('INDEX dan RAW berbeda pada partisi bulan terakhir. Jalankan bootstrap bertahap; full rebuild otomatis dinonaktifkan.');
        props.setProperties({MANTUL_READ_INDEX_TAIL_START:String(indexTail.startRow),
          MANTUL_READ_INDEX_TAIL_PERIOD:indexTail.year+'-'+indexTail.month,
          MANTUL_READ_INDEX_TAIL_HASH:indexTail.hash,MANTUL_READ_INDEX_TAIL_VAL:String(indexTail.value)});
        return {ok:true,status:'UNCHANGED',sheet:MANTUL_READ_INDEX_SHEET_,
          rows:raw.getLastRow()-1,tailPeriod:indexTail.year+'-'+indexTail.month,
          tailRows:indexTail.rows,indexValue:Number(props.getProperty('MANTUL_READ_INDEX_VAL')||0),
          metadataSeeded:true};
      }
      throw new Error('Read index belum siap. Bootstrap bertahap harus dijalankan; full rebuild satu eksekusi tidak diizinkan.');
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
    props.setProperty('MANTUL_READ_INDEX_STATUS','BUILDING');
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
      tailRows:tail.rows,tailStartRow:tail.startRow,previousTailRows:oldLast-oldStart+1,sourceRows:raw.getLastRow()-1,indexValue:totalValue,validated:true};
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



/**
 * Read-only, row-by-row reconciliation of RAW_DATA_SALES A:Y against
 * P2_EVENT_INDEX A:Y. Does not write sheets, properties, caches or triggers.
 * Run only after clearing any sheet filters for human visual inspection;
 * Apps Script getValues() itself reads filtered/hidden rows too.
 */
function auditMantulRawIndexParity() {
  const started = Date.now();
  const ss = getDatabase_();
  const raw = ss.getSheetByName(RAW_DATA_SHEET_NAME);
  const index = ss.getSheetByName(MANTUL_READ_INDEX_SHEET_);
  if (!raw || !index) throw new Error('RAW_DATA_SALES atau P2_EVENT_INDEX tidak ditemukan.');
  const columns = RAW_SALES_COLUMN_COUNT;
  const rawLast = raw.getLastRow(), indexLast = index.getLastRow();
  const rawHeaders = raw.getRange(1, 1, 1, columns).getValues()[0];
  const indexHeaders = index.getRange(1, 1, 1, columns).getValues()[0];
  const headerDiffs = [];
  for (let c = 0; c < columns; c++) {
    if (JSON.stringify(rawHeaders[c]) !== JSON.stringify(indexHeaders[c])) {
      headerDiffs.push({column:c + 1, raw:String(rawHeaders[c]), index:String(indexHeaders[c])});
    }
  }

  function monthKey(row) {
    const y = String(row[0] instanceof Date ? row[0].getFullYear() : row[0]).trim();
    const m = page2Month_(row[1]);
    return y && m ? y + '-' + String(PAGE2_MONTHS_.indexOf(m) + 1).padStart(2, '0') : 'INVALID_YEAR_MONTH';
  }
  function norm(value) {
    if (value instanceof Date) return 'DATE:' + value.toISOString();
    if (value === null || value === undefined) return '';
    if (typeof value === 'number' && !isFinite(value)) return String(value);
    return value;
  }
  function summarizeMonth(map, row) {
    const key = monthKey(row);
    if (!map[key]) map[key] = {rows:0, value:0};
    map[key].rows++;
    map[key].value += Number(row[RAW_SALES_COL.VALUE]) || 0;
  }

  const rawMonths = {}, indexMonths = {};
  const firstDifferences = [];
  const columnDifferenceCounts = {};
  let comparedRows = 0, differentRows = 0, comparedCells = 0, differentCells = 0;
  const commonLast = Math.min(rawLast, indexLast);
  const batchSize = 2500;
  for (let start = 2; start <= commonLast; start += batchSize) {
    const count = Math.min(batchSize, commonLast - start + 1);
    const rawRows = raw.getRange(start, 1, count, columns).getValues();
    const indexRows = index.getRange(start, 1, count, columns).getValues();
    for (let i = 0; i < count; i++) {
      const rr = rawRows[i], ir = indexRows[i];
      summarizeMonth(rawMonths, rr);
      summarizeMonth(indexMonths, ir);
      comparedRows++;
      let rowDifferent = false;
      for (let c = 0; c < columns; c++) {
        comparedCells++;
        if (JSON.stringify(norm(rr[c])) !== JSON.stringify(norm(ir[c]))) {
          differentCells++;
          rowDifferent = true;
          const colName = String(rawHeaders[c] || ('COL_' + (c + 1)));
          columnDifferenceCounts[colName] = (columnDifferenceCounts[colName] || 0) + 1;
          if (firstDifferences.length < 30) {
            firstDifferences.push({
              row:start + i, column:c + 1, header:colName,
              raw:rr[c] instanceof Date ? rr[c].toISOString() : rr[c],
              index:ir[c] instanceof Date ? ir[c].toISOString() : ir[c]
            });
          }
        }
      }
      if (rowDifferent) differentRows++;
    }
  }

  const monthKeys = Array.from(new Set(Object.keys(rawMonths).concat(Object.keys(indexMonths)))).sort();
  const monthComparison = {};
  monthKeys.forEach(key => {
    const r = rawMonths[key] || {rows:0,value:0};
    const x = indexMonths[key] || {rows:0,value:0};
    monthComparison[key] = {
      rawRows:r.rows, indexRows:x.rows, rowDelta:r.rows - x.rows,
      rawValue:r.value, indexValue:x.value, valueDelta:r.value - x.value
    };
  });
  const report = {
    ok:true, readOnly:true, comparison:'RAW_DATA_SALES A:Y vs P2_EVENT_INDEX A:Y',
    spreadsheetId:ss.getId(), elapsedMs:0,
    rawLastRow:rawLast, indexLastRow:indexLast,
    rawDataRows:Math.max(0, rawLast - 1), indexDataRows:Math.max(0, indexLast - 1),
    rowCountEqual:rawLast === indexLast,
    headerDifferences:headerDiffs,
    comparedRows:comparedRows, comparedCells:comparedCells,
    differentRows:differentRows, differentCells:differentCells,
    firstDifferences:firstDifferences,
    columnDifferenceCounts:columnDifferenceCounts,
    byMonth:monthComparison,
    unpairedRows:Math.abs(rawLast - indexLast),
    note:'Read-only. Filters and hidden rows do not exclude getValues() results.'
  };
  report.elapsedMs = Date.now() - started;
  Logger.log('MANTUL RAW-INDEX PARITY AUDIT READ-ONLY\\n' + JSON.stringify(report, null, 2));
  return report;
}

// ============================================================
// RESUMABLE FIRST-TIME BOOTSTRAP — INDEX -> PAGE 2 FACT -> PAGE 3/4 FACT
// Builds into staging sheets in deterministic chunks. Existing published sheets
// remain untouched until the corresponding staging output has passed validation.
// A one-minute time trigger resumes work after each bounded execution.
// ============================================================
const MANTUL_BOOTSTRAP_STATE_KEY_ = 'MANTUL_BOOTSTRAP_STATE_V1';
const MANTUL_BOOTSTRAP_BATCH_ = 2500;
const MANTUL_BOOTSTRAP_VERIFY_BATCH_ = 5000;
const MANTUL_BOOTSTRAP_RUNTIME_MS_ = 210000;
const MANTUL_BOOTSTRAP_INDEX_STAGE_ = 'P2_EVENT_INDEX_BUILD';
const MANTUL_BOOTSTRAP_P2_STAGE_ = 'P2_OUTLET_FACT_BUILD';
const MANTUL_BOOTSTRAP_P2_PACK_ = 'P2_OUTLET_FACT_PACK';
const MANTUL_BOOTSTRAP_P34_STAGE_ = 'P34_EVENT_FACT_BUILD';
const MANTUL_BOOTSTRAP_P34_PACK_ = 'P34_EVENT_FACT_PACK';
const MANTUL_BOOTSTRAP_TRIGGER_ = 'mantulContinueDashboardBuild_';

function mantulBootstrapRequired_() {
  const props = PropertiesService.getScriptProperties();
  const state = props.getProperty(MANTUL_BOOTSTRAP_STATE_KEY_);
  if (state) return true;
  const ss = getDatabase_();
  const index = mantulGetReadyReadIndex_();
  const p2 = ss.getSheetByName(PAGE2_FACT_SHEET_);
  const p34 = ss.getSheetByName(PAGE34_FACT_SHEET_);
  let p2HeaderReady=false,p34HeaderReady=false;
  if(p2&&p2.getLastRow()>1&&p2.getLastColumn()>=PAGE2_FACT_HEADERS_.length){
    const h=p2.getRange(1,1,1,PAGE2_FACT_HEADERS_.length).getDisplayValues()[0].map(page2HeaderKey_);
    p2HeaderReady=PAGE2_FACT_HEADERS_.every((v,i)=>h[i]===page2HeaderKey_(v));
  }
  if(p34&&p34.getLastRow()>1&&p34.getLastColumn()>=PAGE34_FACT_HEADERS_.length){
    const h=p34.getRange(1,1,1,PAGE34_FACT_HEADERS_.length).getDisplayValues()[0].map(page2HeaderKey_);
    p34HeaderReady=PAGE34_FACT_HEADERS_.every((v,i)=>h[i]===page2HeaderKey_(v));
  }
  const p2Ready = index && p2HeaderReady &&
    props.getProperty('PAGE2_FACT_STATUS') === 'COMPLETE' &&
    props.getProperty('PAGE2_FACT_SOURCE_ID') === ss.getId() &&
    Number(props.getProperty('PAGE2_FACT_SOURCE_ROWS')) === index.getLastRow() - 1 &&
    Number(props.getProperty('PAGE2_FACT_WRITTEN')) === p2.getLastRow() - 1;
  const p34Ready = index && p34HeaderReady &&
    props.getProperty('PAGE34_FACT_STATUS') === 'COMPLETE' &&
    props.getProperty('PAGE34_FACT_SOURCE_ID') === ss.getId() &&
    Number(props.getProperty('PAGE34_FACT_SOURCE_ROWS')) === index.getLastRow() - 1 &&
    Number(props.getProperty('PAGE34_FACT_WRITTEN')) === p34.getLastRow() - 1;
  return !index || !p2Ready || !p34Ready;
}

function mantulBootstrapSave_(state) {
  PropertiesService.getScriptProperties().setProperty(MANTUL_BOOTSTRAP_STATE_KEY_, JSON.stringify(state));
}

function mantulBootstrapLoad_() {
  const raw = PropertiesService.getScriptProperties().getProperty(MANTUL_BOOTSTRAP_STATE_KEY_);
  return raw ? JSON.parse(raw) : null;
}

function mantulBootstrapEnsureTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction() === MANTUL_BOOTSTRAP_TRIGGER_);
  if (!exists) ScriptApp.newTrigger(MANTUL_BOOTSTRAP_TRIGGER_).timeBased().everyMinutes(1).create();
}

function mantulBootstrapRemoveTrigger_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === MANTUL_BOOTSTRAP_TRIGGER_) ScriptApp.deleteTrigger(t);
  });
}

function mantulBootstrapEnsureRows_(sheet, requiredLastRow, requiredColumns) {
  const maxRows=sheet.getMaxRows(),maxCols=sheet.getMaxColumns();
  if(maxRows<requiredLastRow)sheet.insertRowsAfter(maxRows,requiredLastRow-maxRows);
  if(requiredColumns&&maxCols<requiredColumns)sheet.insertColumnsAfter(maxCols,requiredColumns-maxCols);
}

function mantulBootstrapNewSheet_(ss, name, headers) {
  const old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  const sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

function mantulBootstrapPublish_(ss, stage, targetName, prefix, properties) {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const current = ss.getSheetByName(targetName);
  const backupName = targetName + '_OLD_' + stamp;
  if (current) current.setName(backupName);
  try {
    stage.setName(targetName);
    stage.setFrozenRows(1);
    PropertiesService.getScriptProperties().setProperties(properties);
    if (current) {
      const backup = ss.getSheetByName(backupName);
      if (backup) ss.deleteSheet(backup);
    }
  } catch (err) {
    const published = ss.getSheetByName(targetName);
    if (published && published.getSheetId() === stage.getSheetId()) ss.deleteSheet(published);
    const backup = ss.getSheetByName(backupName);
    if (backup && !ss.getSheetByName(targetName)) backup.setName(targetName);
    throw err;
  }
}

function mantulStartOrResumeBootstrap_(forceIndex) {
  const props = PropertiesService.getScriptProperties();
  let state = mantulBootstrapLoad_();
  if(state&&forceIndex)throw new Error('Bootstrap sedang berjalan. Selesaikan atau audit proses aktif sebelum memulai rebuild paksa.');
  if (state) {
    if (state.sourceId !== getDatabase_().getId()) throw new Error('Bootstrap aktif berasal dari spreadsheet berbeda. Hentikan dan audit sebelum melanjutkan.');
    mantulBootstrapEnsureTrigger_();
    return {ok:true,status:'RESUMING',phase:state.phase,offset:state.offset || 0};
  }
  const ss = getDatabase_(), raw = ss.getSheetByName(RAW_DATA_SHEET_NAME);
  if (!raw || raw.getLastRow() < 2 || raw.getLastColumn() < RAW_SALES_COLUMN_COUNT) throw new Error('RAW_DATA_SALES A:Y belum siap.');
  validateRawDatabaseHeaders_(raw.getRange(1,1,1,RAW_SALES_COLUMN_COUNT).getDisplayValues()[0]);
  const readyIndex = forceIndex ? null : mantulGetReadyReadIndex_();
  state = {
    version:1, sourceId:ss.getId(), sourceLastRow:raw.getLastRow(),
    sourceRows:raw.getLastRow()-1, sourceTailHash:mantulTailPartition_(raw).hash,
    phase:readyIndex?'P2_INIT':'INDEX_INIT',
    offset:0, sourceValue:0, verifiedValue:0, chunkRows:[],
    sourceCount:0, sourceCarton:0, packedRows:0, packIndex:0
  };
  props.deleteProperty('MANTUL_BOOTSTRAP_ERROR');
  props.setProperties({
    PAGE2_FACT_STATUS:'BUILDING',
    PAGE34_FACT_STATUS:'BUILDING'
  });
  if (!readyIndex) props.setProperty('MANTUL_READ_INDEX_STATUS','BUILDING');
  mantulBootstrapSave_(state);
  mantulBootstrapEnsureTrigger_();
  return {ok:true,status:'STARTED',phase:state.phase,sourceRows:state.sourceRows,batchSize:MANTUL_BOOTSTRAP_BATCH_};
}

function mantulBootstrapCheckSource_(state) {
  const ss = getDatabase_(), raw = ss.getSheetByName(RAW_DATA_SHEET_NAME);
  if (!raw || ss.getId() !== state.sourceId || raw.getLastRow() !== state.sourceLastRow) {
    throw new Error('RAW_DATA_SALES berubah selama bootstrap. Proses dihentikan; audit sumber sebelum melanjutkan.');
  }
  if(state.sourceTailHash && mantulTailPartition_(raw).hash!==state.sourceTailHash){
    throw new Error('Isi partisi bulan terakhir RAW_DATA_SALES berubah selama bootstrap. Proses dihentikan agar INDEX/FACT tidak mencampur dua versi sumber.');
  }
  return {ss:ss,raw:raw};
}

function mantulBootstrapIndexStep_(state, ss, raw) {
  const props = PropertiesService.getScriptProperties();
  if (state.phase === 'INDEX_INIT') {
    const stage = mantulBootstrapNewSheet_(ss, MANTUL_BOOTSTRAP_INDEX_STAGE_, RAW_SALES_HEADERS);
    state.phase = 'INDEX_COPY'; state.offset = 0; state.sourceValue = 0;
    mantulBootstrapSave_(state);
    return;
  }
  const stage = ss.getSheetByName(MANTUL_BOOTSTRAP_INDEX_STAGE_);
  if (!stage) throw new Error('Sheet staging INDEX hilang; bootstrap tidak dapat dilanjutkan dengan aman.');
  if (state.phase === 'INDEX_COPY') {
    if (state.offset < state.sourceRows) {
      const n = Math.min(MANTUL_BOOTSTRAP_BATCH_, state.sourceRows - state.offset);
      const values = raw.getRange(state.offset + 2,1,n,RAW_SALES_COLUMN_COUNT).getValues();
      let partValue = 0;
      values.forEach(r => partValue += Number(r[RAW_SALES_COL.VALUE]) || 0);
      mantulBootstrapEnsureRows_(stage,state.offset+1+n,RAW_SALES_COLUMN_COUNT);
      stage.getRange(state.offset + 2,1,n,RAW_SALES_COLUMN_COUNT).setValues(values);
      SpreadsheetApp.flush();
      state.sourceValue += partValue;
      state.offset += n;
      mantulBootstrapSave_(state);
      return;
    }
    if (stage.getLastRow() !== state.sourceRows + 1) throw new Error('Jumlah baris staging INDEX tidak cocok setelah penyalinan.');
    state.phase = 'INDEX_VERIFY'; state.offset = 0; state.verifiedValue = 0;
    mantulBootstrapSave_(state);
    return;
  }
  if (state.phase === 'INDEX_VERIFY') {
    if (state.offset < state.sourceRows) {
      const n = Math.min(MANTUL_BOOTSTRAP_VERIFY_BATCH_, state.sourceRows - state.offset);
      const sourceBatch=raw.getRange(state.offset+2,1,n,RAW_SALES_COLUMN_COUNT).getValues();
      const stagedBatch=stage.getRange(state.offset+2,1,n,RAW_SALES_COLUMN_COUNT).getValues();
      for(let i=0;i<n;i++){
        for(let c=0;c<RAW_SALES_COLUMN_COUNT;c++){
          const a=sourceBatch[i][c],b=stagedBatch[i][c];
          const equal=(a instanceof Date||b instanceof Date)?(a instanceof Date&&b instanceof Date&&a.getTime()===b.getTime()):a===b;
          if(!equal)throw new Error('Validasi INDEX gagal: konten A:Y berbeda pada baris '+(state.offset+i+2)+', kolom '+(c+1)+'.');
        }
        state.verifiedValue+=Number(stagedBatch[i][RAW_SALES_COL.VALUE])||0;
      }
      state.offset += n;
      mantulBootstrapSave_(state);
      return;
    }
    const close = (a,b) => Math.abs(a-b) <= Math.max(0.01,Math.abs(a)*1e-10);
    if (!close(state.sourceValue,state.verifiedValue)) throw new Error('Validasi INDEX gagal: total VAL sumber dan staging berbeda.');
    mantulBootstrapPublish_(ss,stage,MANTUL_READ_INDEX_SHEET_,'INDEX',{
      MANTUL_READ_INDEX_STATUS:'COMPLETE',
      MANTUL_READ_INDEX_SOURCE_ID:ss.getId(),
      MANTUL_READ_INDEX_SOURCE_ROWS:String(state.sourceRows),
      MANTUL_READ_INDEX_SOURCE_LASTROW:String(state.sourceLastRow),
      MANTUL_READ_INDEX_VAL:String(state.verifiedValue),
      MANTUL_READ_INDEX_BUILT_AT:new Date().toISOString()
    });
    const index = ss.getSheetByName(MANTUL_READ_INDEX_SHEET_);
    const tail = mantulTailPartition_(index);
    props.setProperties({
      MANTUL_READ_INDEX_TAIL_START:String(tail.startRow),
      MANTUL_READ_INDEX_TAIL_PERIOD:tail.year+'-'+tail.month,
      MANTUL_READ_INDEX_TAIL_HASH:tail.hash,
      MANTUL_READ_INDEX_TAIL_VAL:String(tail.value)
    });
    state.phase = 'P2_INIT'; state.offset = 0;
    mantulBootstrapSave_(state);
  }
}

function mantulBootstrapFactInit_(state, ss, kind) {
  const isP2 = kind === 'P2';
  const stageName = isP2 ? MANTUL_BOOTSTRAP_P2_STAGE_ : MANTUL_BOOTSTRAP_P34_STAGE_;
  const headers = isP2 ? PAGE2_FACT_HEADERS_ : PAGE34_FACT_HEADERS_;
  mantulBootstrapNewSheet_(ss,stageName,headers);
  state.phase = isP2 ? 'P2_BUILD' : 'P34_BUILD';
  state.offset = 0; state.chunkRows = [];
  state.sourceCount = 0; state.sourceValue = 0; state.sourceCarton = 0;
  state.packedRows = 0; state.packIndex = 0;
  mantulBootstrapSave_(state);
}

function mantulBootstrapP2Step_(state, ss) {
  if (state.phase === 'P2_INIT') { mantulBootstrapFactInit_(state,ss,'P2'); return; }
  const stage = ss.getSheetByName(MANTUL_BOOTSTRAP_P2_STAGE_);
  if (!stage) throw new Error('Staging P2_OUTLET_FACT hilang.');
  if (state.phase === 'P2_BUILD') {
    if (state.offset < state.sourceRows) {
      const n = Math.min(MANTUL_BOOTSTRAP_BATCH_,state.sourceRows-state.offset);
      const incoming = page2Read_(ss.getSheetByName(MANTUL_READ_INDEX_SHEET_),state.offset+2,n);
      const grouped = new Map(); let partValue=0,partCount=0;
      incoming.forEach(r=>{
        const key=JSON.stringify(r.slice(0,14));let item=grouped.get(key);
        if(!item){item=r.slice(0,14).concat([0,0]);grouped.set(key,item);}
        item[14]+=Number(r[14])||0;item[15]+=Number(r[15])||1;
        partValue+=Number(r[14])||0;partCount+=Number(r[15])||1;
      });
      const rows=Array.from(grouped.values());
      if(rows.length){mantulBootstrapEnsureRows_(stage,state.offset+1+rows.length,PAGE2_FACT_HEADERS_.length);stage.getRange(state.offset+2,1,rows.length,PAGE2_FACT_HEADERS_.length).setValues(rows);}
      state.chunkRows.push(rows.length);
      state.sourceValue+=partValue;state.sourceCount+=partCount;state.offset+=n;
      SpreadsheetApp.flush();mantulBootstrapSave_(state);return;
    }
    if(state.sourceCount!==state.sourceRows)throw new Error('Jumlah source rows P2 tidak cocok dengan INDEX.');
    state.phase='P2_PACK_INIT';state.offset=0;state.packIndex=0;state.packedRows=0;
    mantulBootstrapSave_(state);return;
  }
  if (state.phase === 'P2_PACK_INIT') {
    mantulBootstrapNewSheet_(ss,MANTUL_BOOTSTRAP_P2_PACK_,PAGE2_FACT_HEADERS_);
    state.phase='P2_PACK';mantulBootstrapSave_(state);return;
  }
  if (state.phase === 'P2_PACK') {
    if(state.packIndex<state.chunkRows.length){
      const count=state.chunkRows[state.packIndex],sourceStart=2+state.packIndex*MANTUL_BOOTSTRAP_BATCH_;
      if(count){
        const rows=stage.getRange(sourceStart,1,count,PAGE2_FACT_HEADERS_.length).getValues();
        const pack=ss.getSheetByName(MANTUL_BOOTSTRAP_P2_PACK_);mantulBootstrapEnsureRows_(pack,state.packedRows+1+count,PAGE2_FACT_HEADERS_.length);pack.getRange(state.packedRows+2,1,count,PAGE2_FACT_HEADERS_.length).setValues(rows);
      }
      state.packedRows+=count;state.packIndex++;SpreadsheetApp.flush();mantulBootstrapSave_(state);return;
    }
    state.phase='P2_VERIFY';state.offset=0;state.verifiedValue=0;state.verifiedCount=0;
    mantulBootstrapSave_(state);return;
  }
  if(state.phase==='P2_VERIFY'){
    const pack=ss.getSheetByName(MANTUL_BOOTSTRAP_P2_PACK_);
    if(state.offset<state.packedRows){
      const n=Math.min(MANTUL_BOOTSTRAP_VERIFY_BATCH_,state.packedRows-state.offset);
      pack.getRange(state.offset+2,1,n,PAGE2_FACT_HEADERS_.length).getValues().forEach(r=>{
        state.verifiedValue+=Number(r[14])||0;state.verifiedCount+=Number(r[15])||0;
      });
      state.offset+=n;mantulBootstrapSave_(state);return;
    }
    const close=(a,b)=>Math.abs(a-b)<=Math.max(.01,Math.abs(a)*1e-10);
    if(state.verifiedCount!==state.sourceCount||!close(state.verifiedValue,state.sourceValue))throw new Error('Validasi P2_OUTLET_FACT gagal (SOURCE ROWS/VAL).');
    const packRows=state.packedRows;
    if(pack.getLastRow()!==packRows+1)throw new Error('Jumlah baris packed P2 tidak cocok.');
    mantulBootstrapPublish_(ss,pack,PAGE2_FACT_SHEET_,'P2',{
      PAGE2_FACT_STATUS:'COMPLETE',PAGE2_FACT_SOURCE_ID:ss.getId(),
      PAGE2_FACT_SOURCE_ROWS:String(state.sourceRows),PAGE2_FACT_WRITTEN:String(packRows),
      PAGE2_FACT_SOURCE_COUNT:String(state.verifiedCount),PAGE2_FACT_VAL:String(state.verifiedValue),
      PAGE2_FACT_BUILT_AT:new Date().toISOString()
    });
    const oldStage=ss.getSheetByName(MANTUL_BOOTSTRAP_P2_STAGE_);if(oldStage)ss.deleteSheet(oldStage);
    state.phase='P34_INIT';state.offset=0;state.chunkRows=[];
    mantulBootstrapSave_(state);
  }
}

function mantulBootstrapP34Step_(state, ss) {
  if(state.phase==='P34_INIT'){mantulBootstrapFactInit_(state,ss,'P34');return;}
  const stage=ss.getSheetByName(MANTUL_BOOTSTRAP_P34_STAGE_);
  if(!stage)throw new Error('Staging P34_EVENT_FACT hilang.');
  if(state.phase==='P34_BUILD'){
    if(state.offset<state.sourceRows){
      const n=Math.min(MANTUL_BOOTSTRAP_BATCH_,state.sourceRows-state.offset);
      const incoming=page34ReadRawIndex_(state.offset+2,n),grouped=new Map();
      let partCount=0,partValue=0,partCarton=0;
      incoming.rows.forEach(r=>{
        const key=JSON.stringify(r.slice(0,14).concat(r.slice(16,18)));let item=grouped.get(key);
        if(!item){item=r.slice();item[14]=0;item[15]=0;item[18]=0;grouped.set(key,item);}
        item[18]+=Number(r[18])||1;item[14]+=Number(r[14])||0;item[15]+=Number(r[15])||0;
        partCount+=Number(r[18])||1;partValue+=Number(r[14])||0;partCarton+=Number(r[15])||0;
      });
      const rows=Array.from(grouped.values());
      if(rows.length){mantulBootstrapEnsureRows_(stage,state.offset+1+rows.length,PAGE34_FACT_HEADERS_.length);stage.getRange(state.offset+2,1,rows.length,PAGE34_FACT_HEADERS_.length).setValues(rows);}
      state.chunkRows.push(rows.length);state.sourceCount+=partCount;
      state.sourceValue+=partValue;state.sourceCarton+=partCarton;state.offset+=n;
      SpreadsheetApp.flush();mantulBootstrapSave_(state);return;
    }
    if(state.sourceCount!==state.sourceRows)throw new Error('Jumlah source rows P34 tidak cocok dengan INDEX.');
    state.phase='P34_PACK_INIT';state.offset=0;state.packIndex=0;state.packedRows=0;
    mantulBootstrapSave_(state);return;
  }
  if(state.phase==='P34_PACK_INIT'){
    mantulBootstrapNewSheet_(ss,MANTUL_BOOTSTRAP_P34_PACK_,PAGE34_FACT_HEADERS_);
    state.phase='P34_PACK';mantulBootstrapSave_(state);return;
  }
  if(state.phase==='P34_PACK'){
    if(state.packIndex<state.chunkRows.length){
      const count=state.chunkRows[state.packIndex],sourceStart=2+state.packIndex*MANTUL_BOOTSTRAP_BATCH_;
      if(count){
        const rows=stage.getRange(sourceStart,1,count,PAGE34_FACT_HEADERS_.length).getValues();
        const pack=ss.getSheetByName(MANTUL_BOOTSTRAP_P34_PACK_);mantulBootstrapEnsureRows_(pack,state.packedRows+1+count,PAGE34_FACT_HEADERS_.length);pack.getRange(state.packedRows+2,1,count,PAGE34_FACT_HEADERS_.length).setValues(rows);
      }
      state.packedRows+=count;state.packIndex++;SpreadsheetApp.flush();mantulBootstrapSave_(state);return;
    }
    state.phase='P34_VERIFY';state.offset=0;state.verifiedValue=0;state.verifiedCarton=0;state.verifiedCount=0;
    mantulBootstrapSave_(state);return;
  }
  if(state.phase==='P34_VERIFY'){
    const pack=ss.getSheetByName(MANTUL_BOOTSTRAP_P34_PACK_);
    if(state.offset<state.packedRows){
      const n=Math.min(MANTUL_BOOTSTRAP_VERIFY_BATCH_,state.packedRows-state.offset);
      pack.getRange(state.offset+2,1,n,PAGE34_FACT_HEADERS_.length).getValues().forEach(r=>{
        state.verifiedCount+=Number(r[18])||0;state.verifiedValue+=Number(r[14])||0;state.verifiedCarton+=Number(r[15])||0;
      });
      state.offset+=n;mantulBootstrapSave_(state);return;
    }
    const close=(a,b)=>Math.abs(a-b)<=Math.max(.01,Math.abs(a)*1e-10);
    if(state.verifiedCount!==state.sourceCount||!close(state.verifiedValue,state.sourceValue)||!close(state.verifiedCarton,state.sourceCarton))throw new Error('Validasi P34_EVENT_FACT gagal (SOURCE ROWS/VAL/CARTON).');
    const packRows=state.packedRows;if(pack.getLastRow()!==packRows+1)throw new Error('Jumlah baris packed P34 tidak cocok.');
    mantulBootstrapPublish_(ss,pack,PAGE34_FACT_SHEET_,'P34',{
      PAGE34_FACT_STATUS:'COMPLETE',PAGE34_FACT_SOURCE_ID:ss.getId(),
      PAGE34_FACT_SOURCE_ROWS:String(state.sourceRows),PAGE34_FACT_WRITTEN:String(packRows),
      PAGE34_FACT_SOURCE_COUNT:String(state.verifiedCount),PAGE34_FACT_VAL:String(state.verifiedValue),
      PAGE34_FACT_CARTON:String(state.verifiedCarton),PAGE34_FACT_BUILT_AT:new Date().toISOString()
    });
    const oldStage=ss.getSheetByName(MANTUL_BOOTSTRAP_P34_STAGE_);if(oldStage)ss.deleteSheet(oldStage);
    state.phase='FINALIZE';mantulBootstrapSave_(state);
  }
}

function mantulContinueDashboardBuild_() {
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(1000))return {ok:true,status:'BUSY',message:'Bootstrap lain sedang berjalan.'};
  const started=Date.now();
  try{
    let state=mantulBootstrapLoad_();
    if(!state)return {ok:true,status:'IDLE'};
    const context=mantulBootstrapCheckSource_(state);
    while(Date.now()-started<MANTUL_BOOTSTRAP_RUNTIME_MS_){
      state=mantulBootstrapLoad_();
      if(!state)break;
      if(state.phase==='INDEX_INIT'||state.phase==='INDEX_COPY'||state.phase==='INDEX_VERIFY'){
        mantulBootstrapIndexStep_(state,context.ss,context.raw);
      }else if(state.phase.indexOf('P2_')===0){
        mantulBootstrapP2Step_(state,context.ss);
      }else if(state.phase.indexOf('P34_')===0){
        mantulBootstrapP34Step_(state,context.ss);
      }else if(state.phase==='FINALIZE'){
        const props=PropertiesService.getScriptProperties();
        const activeTail=mantulTailPartition_(context.ss.getSheetByName(MANTUL_READ_INDEX_SHEET_));
        const activePeriod=activeTail.year+'-'+String(PAGE2_MONTHS_.indexOf(activeTail.month)+1).padStart(2,'0');
        const p2Tail=mantulDiscoverFactTail_(context.ss.getSheetByName(PAGE2_FACT_SHEET_),activePeriod,15,14,null);
        const p34Tail=mantulDiscoverFactTail_(context.ss.getSheetByName(PAGE34_FACT_SHEET_),activePeriod,18,14,15);
        props.setProperties({
          PAGE2_FACT_TAIL_START:String(p2Tail.startRow),PAGE2_FACT_TAIL_PERIOD:p2Tail.period,
          PAGE2_FACT_TAIL_SOURCE_COUNT:String(p2Tail.sourceCount),PAGE2_FACT_TAIL_VAL:String(p2Tail.value),
          PAGE34_FACT_TAIL_START:String(p34Tail.startRow),PAGE34_FACT_TAIL_PERIOD:p34Tail.period,
          PAGE34_FACT_TAIL_SOURCE_COUNT:String(p34Tail.sourceCount),PAGE34_FACT_TAIL_VAL:String(p34Tail.value),
          PAGE34_FACT_TAIL_CARTON:String(p34Tail.carton)
        });
        const stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd MMM yyyy HH:mm');
        props.setProperties({SALES_LAST_UPDATED:stamp,PAGE2_RAW_UPDATED:stamp});
        props.deleteProperty('MANTUL_V2_DATA_STALE');
        props.deleteProperty(MANTUL_BOOTSTRAP_STATE_KEY_);
        cacheDashboardData_();
        mantulBootstrapRemoveTrigger_();
        return {ok:true,status:'COMPLETE',message:'Bootstrap INDEX, Page 2 FACT, dan Page 3/4 FACT selesai dan tervalidasi.',sourceRows:state.sourceRows,elapsedMs:Date.now()-started};
      }else{
        throw new Error('Fase bootstrap tidak dikenal: '+state.phase);
      }
    }
    state=mantulBootstrapLoad_();
    return {ok:true,status:'IN_PROGRESS',phase:state?state.phase:'FINALIZE',offset:state?state.offset:0,
      sourceRows:state?state.sourceRows:0,elapsedMs:Date.now()-started,
      message:'Batch selesai untuk sesi ini. Trigger akan melanjutkan otomatis; fungsi ini juga aman dijalankan ulang untuk resume.'};
  }catch(err){
    const props=PropertiesService.getScriptProperties();
    props.setProperty('MANTUL_BOOTSTRAP_ERROR',String(err&&err.stack||err));
    Logger.log('MANTUL BOOTSTRAP FAILED: '+(err&&err.stack||err));
    // Keep state and staging sheets for diagnosis/resume; do not delete published data.
    mantulBootstrapRemoveTrigger_();
    throw err;
  }finally{
    lock.releaseLock();
  }
}

// Trigger continuation and manual resume entry point.
/** Controlled full recovery from RAW A:Y. Rebuilds INDEX, P2 FACT and P34 FACT in resumable staging. */
function rebuildMantulAllIndexesVerified(){
 if(mantulBootstrapLoad_())throw new Error('Bootstrap sedang aktif. Periksa getMantulDashboardBuildStatus() sebelum memulai rebuild.');
 const started=mantulStartOrResumeBootstrap_(true);
 return Object.assign(started,mantulContinueDashboardBuild_());
}

function mantulContinueDashboardBuild() {
  return mantulContinueDashboardBuild_();
}

function getMantulDashboardBuildStatus() {
  const state=mantulBootstrapLoad_();
  const triggers=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===MANTUL_BOOTSTRAP_TRIGGER_).length;
  if(!state)return {ok:true,status:'IDLE',triggerCount:triggers};
  return {ok:true,status:'IN_PROGRESS',phase:state.phase,offset:state.offset||0,
    sourceRows:state.sourceRows,completedSourcePercent:state.sourceRows?Math.round((state.offset||0)/state.sourceRows*1000)/10:0,
    chunkCount:(state.chunkRows||[]).length,packedRows:state.packedRows||0,triggerCount:triggers,
    lastError:PropertiesService.getScriptProperties().getProperty('MANTUL_BOOTSTRAP_ERROR')||''};
}
