/**
 * MANTUL ANALYTICS — IN-PLACE BUILDER
 * Output P1_FACT dan P2_EVENT_INDEX disimpan di spreadsheet sumber yang sama.
 * Sumber utama: DATA_SALES + RAW_DATA_SALES.
 * Jalankan analyticsStartInitialBuild(), lalu ulangi analyticsBuildNextChunk()
 * sampai status COMPLETE.
 */
const ANALYTICS_BUILD_PREFIX_ = 'MANTUL_ANALYTICS_BUILD_';
const ANALYTICS_BATCH_SIZE_ = 4000;

const P1_FACT_HEADERS_ = [
  'YEAR','MONTH','SUBDIST','SUBBRAND','BRAND','CATEGORY','SBU','PIC',
  'PF_MONTHLY','PF_QUARTERLY','VAL','SALES_TYPE','TARGET','BE'
];
const P2_EVENT_HEADERS_ = [
  'YEAR','MONTH','DD','CUSTOMER_CODE','CUSTOMER_NAME','SUBDIST','PIC',
  'BRAND','SUBBRAND','CATEGORY','SBU','SALES_TYPE','SALESMAN','TEAM','VAL'
];

/** Memulai ulang build. Hanya sheet output analytics yang dibersihkan. */
function analyticsStartInitialBuild() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const source = getDatabase_();
    const p1 = source.getSheetByName(DATA_SHEET_NAME);
    const p2 = source.getSheetByName(RAW_DATA_SHEET_NAME);
    if (!p1 || p1.getLastRow() < 2) throw new Error('DATA_SALES tidak ditemukan atau kosong pada spreadsheet sumber.');
    if (!p2 || p2.getLastRow() < 2) throw new Error('RAW_DATA_SALES tidak ditemukan atau kosong pada spreadsheet sumber.');

    const p1Width = Math.max(p1.getLastColumn(), SALES_HEADERS.length);
    const p1Headers = p1.getRange(1,1,1,p1Width).getValues()[0];
    validateDatabaseHeaders_(p1Headers);
    analyticsResolveP1Columns_(p1Headers);
    const rawHeaders = p2.getRange(1,1,1,RAW_SALES_COLUMN_COUNT).getValues()[0];
    validateRawDatabaseHeaders_(rawHeaders);

    // Pastikan output berada pada file sumber yang sama; tidak membuat file baru.
    analyticsPrepareSheet_(source, 'P1_FACT', P1_FACT_HEADERS_);
    analyticsPrepareSheet_(source, 'P2_EVENT_INDEX', P2_EVENT_HEADERS_);
    analyticsPrepareSheet_(source, 'SYSTEM_STATE', ['KEY','VALUE','UPDATED_AT']);

    const props = PropertiesService.getScriptProperties();
    const prefix = ANALYTICS_BUILD_PREFIX_;
    const now = new Date().toISOString();
    props.setProperties({
      [prefix+'STATUS']:'RUNNING',
      [prefix+'PHASE']:'P1',
      [prefix+'P1_NEXT_ROW']:'2',
      [prefix+'P2_NEXT_ROW']:'2',
      [prefix+'P1_SOURCE_ROWS']:String(p1.getLastRow()-1),
      [prefix+'P2_SOURCE_ROWS']:String(p2.getLastRow()-1),
      [prefix+'P1_WRITTEN']:'0',
      [prefix+'P2_WRITTEN']:'0',
      [prefix+'P2_SOURCE_UPDATED_AT']:String(props.getProperty('PAGE2_RAW_UPDATED')||''),
      [prefix+'STARTED_AT']:now,
      [prefix+'SOURCE_ID']:source.getId()
    });
    source.getSheetByName('SYSTEM_STATE').getRange(2,1,6,3).setValues([
      ['BUILD_STATUS','RUNNING',now],
      ['SOURCE_SPREADSHEET_ID',source.getId(),now],
      ['SOURCE_DATA_SALES_ROWS',String(p1.getLastRow()-1),now],
      ['SOURCE_RAW_DATA_SALES_ROWS',String(p2.getLastRow()-1),now],
      ['SCHEMA_VERSION','3',now],
      ['BUILD_PHASE','P1',now]
    ]);
    return analyticsGetBuildStatus();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Satu kali klik untuk rebuild dan melanjutkan batch otomatis.
 * Jika mendekati batas waktu Apps Script, trigger akan melanjutkan proses.
 */
function analyticsRebuildAll() {
  analyticsStartInitialBuild();
  return analyticsRunBuildWorker();
}

function analyticsRunBuildWorker() {
  const started = Date.now();
  const budgetMs = 4 * 60 * 1000;
  let status = analyticsGetBuildStatus();
  while (status.status === 'RUNNING' && Date.now() - started < budgetMs) {
    status = analyticsBuildNextChunk();
  }
  if (status.status === 'COMPLETE') {
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getHandlerFunction() === 'analyticsRunBuildWorker') ScriptApp.deleteTrigger(trigger);
    });
    return status;
  }
  const exists = ScriptApp.getProjectTriggers().some(trigger =>
    trigger.getHandlerFunction() === 'analyticsRunBuildWorker'
  );
  if (!exists) {
    ScriptApp.newTrigger('analyticsRunBuildWorker').timeBased().after(60 * 1000).create();
  }
  return Object.assign(status, {continuationScheduled:true});
}

/** Memproses satu batch 4.000 baris; aman dijalankan berulang. */
function analyticsBuildNextChunk() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const prefix = ANALYTICS_BUILD_PREFIX_;
    if (props.getProperty(prefix+'STATUS') !== 'RUNNING') {
      throw new Error('Build belum aktif. Jalankan analyticsStartInitialBuild() terlebih dahulu.');
    }
    const source = getDatabase_();
    if (source.getId() !== props.getProperty(prefix+'SOURCE_ID')) {
      throw new Error('Spreadsheet sumber berubah. Build dihentikan untuk mencegah data tercampur.');
    }
    const phase = props.getProperty(prefix+'PHASE') || 'P1';
    const started = Date.now();

    if (phase === 'P1') {
      const sh = source.getSheetByName(DATA_SHEET_NAME);
      const next = Number(props.getProperty(prefix+'P1_NEXT_ROW')||2);
      const last = Number(props.getProperty(prefix+'P1_SOURCE_ROWS')||0)+1;
      if (next > last) {
        props.setProperty(prefix+'PHASE','P2');
        return Object.assign(analyticsGetBuildStatus(),{lastBatch:{phase:'P1',written:0,complete:true}});
      }
      const count = Math.min(ANALYTICS_BATCH_SIZE_,last-next+1);
      const width = Math.max(sh.getLastColumn(),SALES_HEADERS.length);
      const header = sh.getRange(1,1,1,width).getValues()[0];
      const ix = analyticsResolveP1Columns_(header);
      const values = sh.getRange(next,1,count,width).getValues();
      const output = [];
      values.forEach(r => {
        if (!r.some(v=>v!==''&&v!==null)) return;
        const year=Number(r[ix.y]), val=Number(r[ix.v]);
        if (!Number.isFinite(year)||!Number.isFinite(val)) return;
        output.push([
          year,String(r[ix.m]||'').trim().toUpperCase(),String(r[ix.d]||''),
          String(r[ix.s]||''),String(r[ix.b]||''),String(r[ix.c]||''),
          String(r[ix.u]||''),String(r[ix.p]||''),String(r[ix.fm]||''),
          String(r[ix.fq]||''),val,String(r[ix.t]||'').trim().toUpperCase(),
          ix.tg<0||r[ix.tg]===''||r[ix.tg]==null?'':Number(r[ix.tg]),
          ix.be<0||r[ix.be]===''||r[ix.be]==null?'':Number(r[ix.be])
        ]);
      });
      if(output.length) analyticsAppendRows_(source.getSheetByName('P1_FACT'),output,P1_FACT_HEADERS_.length);
      props.setProperties({
        [prefix+'P1_NEXT_ROW']:String(next+count),
        [prefix+'P1_WRITTEN']:String(Number(props.getProperty(prefix+'P1_WRITTEN')||0)+output.length)
      });
      return Object.assign(analyticsGetBuildStatus(),{lastBatch:{phase:'P1',sourceStart:next,sourceEnd:next+count-1,written:output.length,elapsedMs:Date.now()-started}});
    }

    if (phase === 'P2') {
      const sh = source.getSheetByName(RAW_DATA_SHEET_NAME);
      const next = Number(props.getProperty(prefix+'P2_NEXT_ROW')||2);
      const last = Number(props.getProperty(prefix+'P2_SOURCE_ROWS')||0)+1;
      if (next > last) {
        props.setProperty(prefix+'STATUS','COMPLETE');
        props.setProperty(prefix+'PHASE','COMPLETE');
        const now = new Date().toISOString();
        const state = source.getSheetByName('SYSTEM_STATE');
        state.getRange(2,1,7,3).setValues([
          ['BUILD_STATUS','COMPLETE',now],
          ['SOURCE_SPREADSHEET_ID',source.getId(),now],
          ['P1_FACT_ROWS',props.getProperty(prefix+'P1_WRITTEN')||'0',now],
          ['P2_EVENT_INDEX_ROWS',props.getProperty(prefix+'P2_WRITTEN')||'0',now],
          ['SCHEMA_VERSION','3',now],
          ['BUILD_COMPLETED_AT',now,now],
          ['SOURCE_VERSION',source.getId()+'|'+sh.getLastRow()+'|'+sh.getSheetId(),now]
        ]);
        return analyticsGetBuildStatus();
      }
      const count=Math.min(ANALYTICS_BATCH_SIZE_,last-next+1);
      const values=sh.getRange(next,1,count,RAW_SALES_COLUMN_COUNT).getValues();
      const output=[];
      values.forEach(r=>{
        if(String(r[RAW_SALES_COL.SALES_TYPE]||'').trim().toUpperCase()!=='SOO') return;
        output.push([
          r[RAW_SALES_COL.YEAR],String(r[RAW_SALES_COL.MONTH]||'').trim().toUpperCase(),
          r[RAW_SALES_COL.DAY],String(r[RAW_SALES_COL.CUSTOMER_CODE]||'').trim(),
          String(r[RAW_SALES_COL.CUSTOMER_NAME]||'').trim(),String(r[RAW_SALES_COL.SUBDIST]||'').trim(),
          String(r[RAW_SALES_COL.PIC]||'').trim(),String(r[RAW_SALES_COL.BRAND]||'').trim(),
          String(r[RAW_SALES_COL.SUBBRAND]||'').trim(),String(r[RAW_SALES_COL.CATEGORY]||'').trim(),
          String(r[RAW_SALES_COL.SBU]||'').trim(),'SOO',
          String(r[RAW_SALES_COL.SALESMAN]||'').trim(),String(r[RAW_SALES_COL.TEAM]||'').trim(),
          Number(r[RAW_SALES_COL.VALUE])||0
        ]);
      });
      if(output.length) analyticsAppendRows_(source.getSheetByName('P2_EVENT_INDEX'),output,P2_EVENT_HEADERS_.length);
      props.setProperties({
        [prefix+'P2_NEXT_ROW']:String(next+count),
        [prefix+'P2_WRITTEN']:String(Number(props.getProperty(prefix+'P2_WRITTEN')||0)+output.length)
      });
      return Object.assign(analyticsGetBuildStatus(),{lastBatch:{phase:'P2',sourceStart:next,sourceEnd:next+count-1,written:output.length,elapsedMs:Date.now()-started}});
    }
    return analyticsGetBuildStatus();
  } finally {
    lock.releaseLock();
  }
}

function analyticsGetBuildStatus() {
  const props=PropertiesService.getScriptProperties(),p=ANALYTICS_BUILD_PREFIX_;
  const sourceId=props.getProperty(p+'SOURCE_ID')||'';
  return {
    status:props.getProperty(p+'STATUS')||'NOT_STARTED',
    phase:props.getProperty(p+'PHASE')||'',
    sourceSpreadsheetId:sourceId,
    outputLocation:'SAME_SOURCE_SPREADSHEET',
    sourceUrl:sourceId?'https://docs.google.com/spreadsheets/d/'+sourceId+'/edit':'',
    p1:{sourceRows:Number(props.getProperty(p+'P1_SOURCE_ROWS')||0),processedThrough:Number(props.getProperty(p+'P1_NEXT_ROW')||2)-2,writtenRows:Number(props.getProperty(p+'P1_WRITTEN')||0)},
    p2:{sourceRows:Number(props.getProperty(p+'P2_SOURCE_ROWS')||0),processedThrough:Number(props.getProperty(p+'P2_NEXT_ROW')||2)-2,writtenRows:Number(props.getProperty(p+'P2_WRITTEN')||0)}
  };
}

function analyticsPrepareSheet_(ss,name,headers) {
  let sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name);
  sh.clearContents();
  const neededColumns=headers.length;
  if(sh.getMaxColumns()<neededColumns) sh.insertColumnsAfter(sh.getMaxColumns(),neededColumns-sh.getMaxColumns());
  sh.getRange(1,1,1,neededColumns).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

function analyticsAppendRows_(sheet,rows,width) {
  if(!rows.length) return;
  const start=sheet.getLastRow()+1,needed=start+rows.length-1;
  if(sheet.getMaxRows()<needed) sheet.insertRowsAfter(sheet.getMaxRows(),needed-sheet.getMaxRows());
  sheet.getRange(start,1,rows.length,width).setValues(rows);
}

function analyticsResolveP1Columns_(headers) {
  const key=v=>String(v==null?'':v).trim().toLowerCase().replace(/[_\\s]+/g,' ');
  const map={};
  headers.forEach((v,i)=>map[key(v)]=i);
  const col=(aliases,required)=>{
    for(const a of aliases) if(map[key(a)]!==undefined) return map[key(a)];
    if(required) throw new Error('Kolom DATA_SALES tidak ditemukan: '+aliases[0]);
    return -1;
  };
  return {
    y:col(['YEAR','y'],true),m:col(['MONTH','m'],true),d:col(['SUBDIST','d'],true),
    s:col(['SUBBRAND','s'],true),b:col(['BRAND','b'],true),c:col(['CATEGORY','c'],true),
    u:col(['SBU','u'],true),p:col(['PIC','p'],true),fm:col(['PF Bulanan','PF_MONTHLY','fm'],true),
    fq:col(['PF Triwulan','PF_QUARTERLY','fq'],true),v:col(['VAL','v'],true),
    t:col(['JENIS SALES','SALES_TYPE','t'],true),tg:col(['TARGET','tg'],false),be:col(['BE','be'],false)
  };
}

/** Kembalikan index hanya jika build selesai dan cocok dengan sumber aktif. */
function analyticsGetReadySheet_(analyticsSheetName,sourceSheetName,phase) {
  const props=PropertiesService.getScriptProperties(),p=ANALYTICS_BUILD_PREFIX_;
  if(props.getProperty(p+'STATUS')!=='COMPLETE') return null;
  const source=getDatabase_();
  if(source.getId()!==String(props.getProperty(p+'SOURCE_ID')||'')) return null;
  const sourceSheet=source.getSheetByName(sourceSheetName);
  if(!sourceSheet||sourceSheet.getLastRow()<2) return null;
  if(sourceSheet.getLastRow()-1!==Number(props.getProperty(p+phase+'_SOURCE_ROWS')||0)) return null;
  if(phase==='P2') {
    const indexed=String(props.getProperty(p+'P2_SOURCE_UPDATED_AT')||'');
    const current=String(props.getProperty('PAGE2_RAW_UPDATED')||'');
    if(indexed!==current&&(indexed||current)) return null;
  }
  const target=source.getSheetByName(analyticsSheetName);
  const expected=Number(props.getProperty(p+phase+'_WRITTEN')||0);
  if(!target||target.getLastRow()-1!==expected) return null;
  if(analyticsSheetName==='P1_FACT'&&target.getLastColumn()<P1_FACT_HEADERS_.length) return null;
  if(analyticsSheetName==='P2_EVENT_INDEX'&&target.getLastColumn()<P2_EVENT_HEADERS_.length) return null;
  return target;
}

/** Pemeriksaan koneksi dan lokasi output; read-only. */
function testAnalyticsDashboardRouting() {
  const source=getDatabase_();
  const p1=analyticsGetReadySheet_('P1_FACT',DATA_SHEET_NAME,'P1');
  const p2=analyticsGetReadySheet_('P2_EVENT_INDEX',RAW_DATA_SHEET_NAME,'P2');
  const result={
    buildStatus:PropertiesService.getScriptProperties().getProperty(ANALYTICS_BUILD_PREFIX_+'STATUS')||'NOT_STARTED',
    sourceSpreadsheetId:source.getId(),
    outputSpreadsheetId:source.getId(),
    sameSpreadsheet:true,
    p1:{route:p1?'P1_FACT':'DATA_SALES fallback',sourceRows:source.getSheetByName(DATA_SHEET_NAME)?.getLastRow()-1||0,analyticsRows:p1?.getLastRow()-1||0},
    p2:{route:p2?'P2_EVENT_INDEX':'RAW_DATA_SALES fallback',sourceRows:source.getSheetByName(RAW_DATA_SHEET_NAME)?.getLastRow()-1||0,analyticsRows:p2?.getLastRow()-1||0}
  };
  Logger.log(JSON.stringify(result,null,2));
  return result;
}
