/**
 * MANTUL — MONTHLY PARTITION IMPORT
 *
 * Import bulanan: P1_FACT tetap menjadi sumber Page 1, sedangkan
 * Page 2–4 membaca transaksi langsung dari RAW_DATA_SALES.
 * P2/P3 tidak diimpor; indeks turunan dibangun ulang setelah import selesai.
 * Kolom helper RAW di luar A:Z dipertahankan.
 */
const MONTHLY_IMPORT_PREFIX_ = 'MANTUL_MONTHLY_IMPORT_';
const MONTHLY_RAW_HEADERS_ = RAW_SALES_HEADERS.slice();
const MONTHLY_STAGE_NAMES_ = {
  p1:'P1_MONTH_STAGING',
  raw:'RAW_MONTH_STAGING'
};
const MONTHLY_BACKUP_NAMES_ = {
  p1:'P1_MONTH_BACKUP',
  raw:'RAW_MONTH_BACKUP'
};

function monthlyImportMonth_(value) {
  const s = String(value == null ? '' : value).trim().toUpperCase();
  const aliases = {
    JANUARI:'JAN',FEBRUARI:'FEB',MARET:'MAR',APRIL:'APR',MEI:'MAY',
    JUNI:'JUN',JULI:'JUL',AGUSTUS:'AUG',SEPTEMBER:'SEP',
    OKTOBER:'OCT',NOVEMBER:'NOV',DESEMBER:'DEC',
    JANUARY:'JAN',FEBRUARY:'FEB',MARCH:'MAR',JUNE:'JUN',
    JULY:'JUL',AUGUST:'AUG',OCTOBER:'OCT',DECEMBER:'DEC'
  };
  if (aliases[s]) return aliases[s];
  if (/^\d{1,2}$/.test(s)) {
    const n = Number(s);
    return n >= 1 && n <= 12 ? PAGE2_MONTHS_[n - 1] : '';
  }
  return s.slice(0,3);
}

function monthlyImportAssertHeaders_(sheet, expected) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < expected.length) {
    throw new Error('Sheet ' + (sheet ? sheet.getName() : '(tidak ada)') + ' tidak memiliki kolom yang diperlukan.');
  }
  const actual = sheet.getRange(1,1,1,expected.length).getDisplayValues()[0]
    .map(v => String(v == null ? '' : v).trim().toUpperCase().replace(/[_-]+/g,' ').replace(/\s+/g,' '));
  const wanted = expected.map(v => String(v).trim().toUpperCase().replace(/[_-]+/g,' ').replace(/\s+/g,' '));
  if (!actual.every((v,i) => v === wanted[i])) {
    throw new Error('Header ' + sheet.getName() + ' tidak sesuai. Terbaca: ' + actual.join(' | '));
  }
}

function beginMonthlyDataImport(meta) {
  assertUploadPermission_(); meta=meta||{};
  const year=Number(meta.year),month=monthlyImportMonth_(meta.month),
    p1Total=Number(meta.p1Total),rawTotal=Number(meta.rawTotal);
  if(!Number.isInteger(year)||year<2000||year>2100||PAGE2_MONTHS_.indexOf(month)<0||
     !Number.isInteger(p1Total)||p1Total<1||!Number.isInteger(rawTotal)||rawTotal<1)
    throw new Error('Periode atau jumlah baris P1/RAW tidak valid.');
  if(p1Total>5000||rawTotal>30000)throw new Error('Jumlah baris periode melebihi batas aman import bulanan.');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const props=PropertiesService.getScriptProperties(),prefix=MONTHLY_IMPORT_PREFIX_,
      existing=props.getProperty(prefix+'TOKEN'),started=Number(props.getProperty(prefix+'STARTED')||0);
    if(existing&&started&&Date.now()-started<2*60*60*1000)throw new Error('Import bulanan lain masih berlangsung.');
    const ss=getDatabase_(),p1=ss.getSheetByName('P1_FACT'),raw=ss.getSheetByName('RAW_DATA_SALES');
    if(!p1||!raw)throw new Error('P1_FACT dan RAW_DATA_SALES wajib tersedia.');
    analyticsAssertHeaders_(p1,P1_FACT_HEADERS_);mantulV2AssertRawHeaders_(raw);
    [['p1',P1_FACT_HEADERS_],['raw',MONTHLY_RAW_HEADERS_]].forEach(([key,headers])=>{
      const stage=getOrCreateSheet_(MONTHLY_STAGE_NAMES_[key]);stage.clearContents();
      ensureSheetGrid_(stage,1,headers.length);stage.getRange(1,1,1,headers.length).setValues([headers]);
    });
    const token=Utilities.getUuid();
    props.setProperties({
      [prefix+'TOKEN']:token,[prefix+'YEAR']:String(year),[prefix+'MONTH']:month,
      [prefix+'P1_EXPECTED']:String(p1Total),[prefix+'RAW_EXPECTED']:String(rawTotal),
      [prefix+'P1_WRITTEN']:'0',[prefix+'RAW_WRITTEN']:'0',
      [prefix+'STARTED']:String(Date.now()),[prefix+'FILENAME']:safeText_(meta.filename||'Workbook sales',200)
    });
    return {ok:true,token:token,year:year,month:month,p1Expected:p1Total,rawExpected:rawTotal};
  }finally{lock.releaseLock();}
}

function monthlyImportAssertSession_(props, token) {
  const prefix = MONTHLY_IMPORT_PREFIX_;
  const saved = props.getProperty(prefix+'TOKEN');
  const started = Number(props.getProperty(prefix+'STARTED') || 0);
  if (!saved || !token || saved !== token) throw new Error('Sesi import bulanan tidak valid atau sudah diganti.');
  if (!started || Date.now()-started >= 2*60*60*1000) {
    monthlyImportClearSession_(props);
    throw new Error('Sesi import bulanan kedaluwarsa. Silakan upload ulang.');
  }
}

function monthlyImportClearSession_(props) {
  const prefix = MONTHLY_IMPORT_PREFIX_;
  ['TOKEN','YEAR','MONTH','P1_EXPECTED','RAW_EXPECTED',
   'P1_WRITTEN','RAW_WRITTEN','STARTED','FILENAME']
    .forEach(key => props.deleteProperty(prefix+key));
}

function monthlyImportValidateChunk_(rows, width, year, month, sheetLabel, yearIndex, monthIndex) {
  (rows || []).forEach((row,i) => {
    if (!Array.isArray(row) || row.length !== width) {
      throw new Error(sheetLabel + ' chunk baris ' + (i+1) + ' harus memiliki ' + width + ' kolom.');
    }
    if (Number(row[yearIndex]) !== year || monthlyImportMonth_(row[monthIndex]) !== month) {
      throw new Error(sheetLabel + ' berisi baris di luar periode ' + year + '-' + month + '.');
    }
  });
}

function appendMonthlyDataImportChunk(payload) {
  assertUploadPermission_();payload=payload||{};
  const p1Chunk=Array.isArray(payload.p1Chunk)?payload.p1Chunk:[],
    rawChunk=Array.isArray(payload.rawChunk)?payload.rawChunk:[];
  if(!p1Chunk.length&&!rawChunk.length)return {written:0};
  if(p1Chunk.length>1000||rawChunk.length>1000)throw new Error('Maksimal 1.000 baris per sheet per batch.');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const props=PropertiesService.getScriptProperties();monthlyImportAssertSession_(props,payload.token);
    const prefix=MONTHLY_IMPORT_PREFIX_,year=Number(props.getProperty(prefix+'YEAR')),month=props.getProperty(prefix+'MONTH');
    [['p1',p1Chunk,14,'P1_FACT'],['raw',rawChunk,26,'RAW_DATA_SALES']].forEach(([key,rows,width,label])=>{
      if(!rows.length)return;
      const expected=Number(props.getProperty(prefix+key.toUpperCase()+'_EXPECTED')||0),
        written=Number(props.getProperty(prefix+key.toUpperCase()+'_WRITTEN')||0);
      if(written+rows.length>expected)throw new Error(label+' melebihi jumlah baris yang didaftarkan.');
      monthlyImportValidateChunk_(rows,width,year,month,label,0,1);
      if(key==='raw'&&rows.some(r=>r.length!==26))throw new Error('RAW_DATA_SALES wajib 26 kolom sesuai schema V2.');
      const stage=getOrCreateSheet_(MONTHLY_STAGE_NAMES_[key]);ensureSheetGrid_(stage,written+rows.length+1,width);
      stage.getRange(written+2,1,rows.length,width).setValues(rows);
      props.setProperty(prefix+key.toUpperCase()+'_WRITTEN',String(written+rows.length));
    });
    return {p1Written:Number(props.getProperty(prefix+'P1_WRITTEN')||0),p1Expected:Number(props.getProperty(prefix+'P1_EXPECTED')||0),
      rawWritten:Number(props.getProperty(prefix+'RAW_WRITTEN')||0),rawExpected:Number(props.getProperty(prefix+'RAW_EXPECTED')||0)};
  }finally{lock.releaseLock();}
}

function monthlyImportReadStage_(key, expected, width) {
  const sheet = getDatabase_().getSheetByName(MONTHLY_STAGE_NAMES_[key]);
  if (!sheet || sheet.getLastRow() !== expected+1) {
    throw new Error('Staging ' + key + ' tidak lengkap. Diharapkan ' + expected + ' baris.');
  }
  return expected ? sheet.getRange(2,1,expected,width).getValues() : [];
}

function monthlyImportFindBlocks_(sheet, year, month) {
  const lastRow = monthlyImportLastDataRow_(sheet);
  if (lastRow <= 1) return [];
  const keys = sheet.getRange(2,1,lastRow-1,2).getValues();
  const blocks = [];
  let start = -1, count = 0;
  for (let i=0;i<keys.length;i++) {
    const match = Number(keys[i][0]) === year && monthlyImportMonth_(keys[i][1]) === month;
    if (match) {
      if (start < 0) start = i+2;
      count++;
    } else if (start >= 0) {
      blocks.push([start,count]);
      start = -1;
      count = 0;
    }
  }
  if (start >= 0) blocks.push([start,count]);
  return blocks;
}

function monthlyImportLastDataRow_(sheet) {
  if(!sheet)return 1;
  // RAW helper formulas may extend far below the actual transaction rows.
  // Locate the last non-empty cell in column A (YEAR), not the sheet's global last row.
  const last=sheet.getRange(sheet.getMaxRows(),1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  return Math.max(1,last);
}

function monthlyImportReadPeriodRows_(sheet, year, month, width) {
  const blocks = monthlyImportFindBlocks_(sheet,year,month);
  const rows = [];
  blocks.forEach(([start,count]) => {
    const part = sheet.getRange(start,1,count,width).getValues();
    Array.prototype.push.apply(rows,part);
  });
  return {blocks:blocks,rows:rows,width:width};
}

function monthlyImportWriteBackup_(key, headers, rows) {
  const backup = getOrCreateSheet_(MONTHLY_BACKUP_NAMES_[key]);
  backup.clearContents();
  ensureSheetGrid_(backup,rows.length+1,headers.length);
  backup.getRange(1,1,1,headers.length).setValues([headers]);
  if (rows.length) backup.getRange(2,1,rows.length,headers.length).setValues(rows);
  return rows;
}

function monthlyImportClearPeriodValues_(sheet,year,month,width){
  const blocks=monthlyImportFindBlocks_(sheet,year,month);
  blocks.forEach(([start,count])=>sheet.getRange(start,1,count,width).clearContent());
  return blocks.reduce((sum,x)=>sum+x[1],0);
}

function monthlyImportDeletePeriod_(sheet, year, month) {
  const blocks = monthlyImportFindBlocks_(sheet,year,month).sort((a,b)=>b[0]-a[0]);
  let removed = 0;
  blocks.forEach(([start,count]) => {
    sheet.deleteRows(start,count);
    removed += count;
  });
  return removed;
}

function monthlyImportAppendRows_(sheet, rows, width) {
  if (!rows.length) return;
  // RAW_DATA_SALES may have helper formulas in columns beyond A:Z.
  // Use the last populated YEAR row, not getLastRow() from helper formulas.
  const start = monthlyImportLastDataRow_(sheet)+1;
  ensureSheetGrid_(sheet,start+rows.length-1,width);
  sheet.getRange(start,1,rows.length,width).setValues(rows);
}

function finishMonthlyDataImport(payload) {
  assertUploadPermission_();payload=payload||{};
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const props=PropertiesService.getScriptProperties();monthlyImportAssertSession_(props,payload.token);
    const prefix=MONTHLY_IMPORT_PREFIX_,year=Number(props.getProperty(prefix+'YEAR')),month=props.getProperty(prefix+'MONTH'),
      p1Expected=Number(props.getProperty(prefix+'P1_EXPECTED')||0),rawExpected=Number(props.getProperty(prefix+'RAW_EXPECTED')||0);
    if(Number(props.getProperty(prefix+'P1_WRITTEN')||0)!==p1Expected||Number(props.getProperty(prefix+'RAW_WRITTEN')||0)!==rawExpected)
      throw new Error('Import belum lengkap. P1_FACT dan RAW_DATA_SALES belum mencapai jumlah baris yang didaftarkan.');
    const p1Rows=monthlyImportReadStage_('p1',p1Expected,14),rawRows=monthlyImportReadStage_('raw',rawExpected,26);
    monthlyImportValidateChunk_(p1Rows,14,year,month,'P1_FACT',0,1);
    monthlyImportValidateChunk_(rawRows,26,year,month,'RAW_DATA_SALES',0,1);
    const p1Soo=p1Rows.filter(r=>page2Key_(r[11])==='SOO').reduce((a,r)=>a+(Number(r[10])||0),0),
      rawSoo=rawRows.filter(r=>page2Key_(r[14])==='SOO').reduce((a,r)=>a+(Number(r[12])||0),0),
      p1So=p1Rows.filter(r=>page2Key_(r[11])==='SO').reduce((a,r)=>a+(Number(r[10])||0),0),
      rawSo=rawRows.filter(r=>page2Key_(r[14])==='SO').reduce((a,r)=>a+(Number(r[12])||0),0);
    const ss=getDatabase_(),live={p1:ss.getSheetByName('P1_FACT'),raw:ss.getSheetByName('RAW_DATA_SALES')},
      headers={p1:P1_FACT_HEADERS_,raw:MONTHLY_RAW_HEADERS_},widths={p1:14,raw:26},newRows={p1:p1Rows,raw:rawRows},oldRows={};
    Object.keys(live).forEach(key=>{
      if(!live[key])throw new Error('Sheet live '+key+' tidak ditemukan.');
      const captured=monthlyImportReadPeriodRows_(live[key],year,month,widths[key]);oldRows[key]=captured.rows;
      monthlyImportWriteBackup_(key,headers[key],captured.rows);
    });
    const oldSales=props.getProperty('SALES_LAST_UPDATED')||'',oldP2=props.getProperty('PAGE2_RAW_UPDATED')||'',
      oldStale=props.getProperty('MANTUL_V2_DATA_STALE')||'';
    try{
      Object.keys(live).forEach(key=>key==='raw'?monthlyImportClearPeriodValues_(live[key],year,month,widths[key]):monthlyImportDeletePeriod_(live[key],year,month));
      Object.keys(live).forEach(key=>monthlyImportAppendRows_(live[key],newRows[key],widths[key]));
      mantulV2AssertRawHeaders_(live.raw);
      const now=new Date(),formatted=Utilities.formatDate(now,Session.getScriptTimeZone(),'dd MMM yyyy HH:mm');
      props.setProperty('SALES_LAST_UPDATED',formatted);props.setProperty('PAGE2_RAW_UPDATED',formatted);
      // Page 2–4 membaca RAW_DATA_SALES secara langsung; tidak perlu rebuild V2.
      if(typeof cacheDashboardData_==='function')cacheDashboardData_();
      const log=getOrCreateSheet_('MONTHLY_IMPORT_LOG');
      if(log.getLastRow()===0)log.appendRow(['WAKTU','FILE','PERIODE','P1_FACT ROWS','RAW_DATA_SALES ROWS','P1 SOO','RAW SOO','SELISIH SOO','P1 SO','RAW SO','UPLOADER','STATUS']);
      log.appendRow([now,props.getProperty(prefix+'FILENAME')||'',year+'-'+month,p1Rows.length,rawRows.length,p1Soo,rawSoo,p1Soo-rawSoo,p1So,rawSo,
        Session.getActiveUser().getEmail()||'email tidak tersedia','SUCCESS — RAW DIRECT']);
      ['p1','raw'].forEach(key=>{const stage=ss.getSheetByName(MONTHLY_STAGE_NAMES_[key]);if(stage)stage.clearContents();});
      monthlyImportClearSession_(props);
      return {ok:true,year:year,month:month,p1Rows:p1Rows.length,rawRows:rawRows.length,p1Soo:p1Soo,rawSoo:rawSoo,
        p1SooDiff:p1Soo-rawSoo,p1So:p1So,rawSo:rawSo,updatedAt:formatted,
        v2RebuildRequired:false,message:'P1_FACT dan RAW_DATA_SALES berhasil diperbarui. Page 2–4 membaca RAW_DATA_SALES langsung.'};
    }catch(publishError){
      let rollbackError=null;
      try{
        Object.keys(live).forEach(key=>key==='raw'?monthlyImportClearPeriodValues_(live[key],year,month,widths[key]):monthlyImportDeletePeriod_(live[key],year,month));
        Object.keys(live).forEach(key=>monthlyImportAppendRows_(live[key],oldRows[key],widths[key]));
        props.setProperty('SALES_LAST_UPDATED',oldSales);props.setProperty('PAGE2_RAW_UPDATED',oldP2);
        if(oldStale)props.setProperty('MANTUL_V2_DATA_STALE',oldStale);else props.deleteProperty('MANTUL_V2_DATA_STALE');
        if(typeof cacheDashboardData_==='function')cacheDashboardData_();
      }catch(err){rollbackError=err;}
      if(rollbackError)throw new Error('Import gagal: '+publishError.message+'. Rollback juga gagal: '+rollbackError.message+'. Periksa backup P1_MONTH_BACKUP dan RAW_MONTH_BACKUP.');
      throw new Error('Import gagal dan data periode sebelumnya sudah dipulihkan: '+publishError.message);
    }
  }finally{lock.releaseLock();}
}

function abortMonthlyDataImport(payload) {
  assertUploadPermission_();
  payload = payload || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    monthlyImportAssertSession_(props,payload.token);
    const ss = getDatabase_();
    Object.keys(MONTHLY_STAGE_NAMES_).forEach(key => {
      const stage = ss.getSheetByName(MONTHLY_STAGE_NAMES_[key]);
      if (stage) stage.clearContents();
    });
    monthlyImportClearSession_(props);
    return {ok:true,message:'Sesi import dibatalkan; data live tidak diubah.'};
  } finally {
    lock.releaseLock();
  }
}
