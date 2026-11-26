/**
 * MANTUL DATABASE V2 — RESUMABLE STAGING BUILDER
 *
 * Safety: this builder writes only *_V2_STAGE sheets. It never clears or
 * overwrites live P1_FACT, P2_EVENT_INDEX, P2_DASHBOARD_FACT or P3_OUTLET_MONTHLY.
 * Run mantulV2BuildStart(), then repeat mantulV2BuildContinue() until READY.
 */
const MANTUL_V2_BUILD_ = Object.freeze({
  PREFIX: 'MANTUL_V2_BUILD_',
  CHUNK_ROWS: 4000,
  EVENT_STAGE: 'P2_EVENT_INDEX_V2_STAGE',
  P2_STAGE: 'P2_DASHBOARD_FACT_V2_STAGE',
  P3_STAGE: 'P3_OUTLET_MONTHLY_V2_STAGE',
  P3_PARTIAL_STAGE: 'P3_OUTLET_PARTIAL_V2_STAGE',
  PARTIAL_STAGE: 'P2_FACT_PARTIAL_V2_STAGE'
});

function mantulV2BuildStart() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getDatabase_();
    const raw = ss.getSheetByName('RAW_DATA_SALES');
    if (!raw || raw.getLastRow() < 2) throw new Error('RAW_DATA_SALES kosong atau tidak ditemukan.');
    mantulV2AssertRawHeaders_(raw);
    const props = PropertiesService.getScriptProperties();
    const p = MANTUL_V2_BUILD_.PREFIX;
    const oldStatus = props.getProperty(p + 'STATUS');
    if (oldStatus === 'RUNNING') throw new Error('Build V2 masih berjalan. Lanjutkan dengan mantulV2BuildContinue().');
    const specs = [
      [MANTUL_V2_BUILD_.EVENT_STAGE, MANTUL_V2_SCHEMA_.P2_EVENT_INDEX],
      [MANTUL_V2_BUILD_.P2_STAGE, MANTUL_V2_SCHEMA_.P2_DASHBOARD_FACT],
      [MANTUL_V2_BUILD_.P3_STAGE, MANTUL_V2_SCHEMA_.P3_OUTLET_MONTHLY],
      [MANTUL_V2_BUILD_.P3_PARTIAL_STAGE, MANTUL_V2_SCHEMA_.P3_OUTLET_MONTHLY],
      [MANTUL_V2_BUILD_.PARTIAL_STAGE, MANTUL_V2_SCHEMA_.P2_DASHBOARD_FACT]
    ];
    specs.forEach(([name, headers]) => {
      const sh = ss.getSheetByName(name) || ss.insertSheet(name);
      sh.clearContents();
      ensureSheetGrid_(sh, 1, headers.length);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    });
    const total = mantulV2LastRawDataRow_(raw) - 1;
    props.setProperties({
      [p+'STATUS']:'RUNNING',
      [p+'PHASE']:'EVENTS',
      [p+'SOURCE_ID']:ss.getId(),
      [p+'SOURCE_ROWS']:String(total),
      [p+'NEXT_ROW']:'2',
      [p+'EVENT_ROWS']:'0',
      [p+'NONEMPTY_ROWS']:'0',
      [p+'P3_ROWS']:'0',
      [p+'PARTIAL_ROWS']:'0',
      [p+'STARTED_AT']:new Date().toISOString(),
      [p+'ERROR']:'',
      MANTUL_V2_DATA_STALE:'TRUE'
    });
    return {status:'RUNNING', sourceRows:total, nextRow:2, chunkRows:MANTUL_V2_BUILD_.CHUNK_ROWS};
  } finally { lock.releaseLock(); }
}

function mantulV2BuildContinue() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const p = MANTUL_V2_BUILD_.PREFIX;
    if (props.getProperty(p+'STATUS') !== 'RUNNING') throw new Error('Tidak ada build V2 berstatus RUNNING. Jalankan mantulV2BuildStart().');
    const ss = getDatabase_();
    if (ss.getId() !== props.getProperty(p+'SOURCE_ID')) throw new Error('Spreadsheet sumber berubah sejak build dimulai.');
    const raw = ss.getSheetByName('RAW_DATA_SALES');
    mantulV2AssertRawHeaders_(raw);
    const total = Number(props.getProperty(p+'SOURCE_ROWS') || 0);
    const start = Number(props.getProperty(p+'NEXT_ROW') || 2);
    if (start > total + 1) return mantulV2BuildStatus();
    const count = Math.min(MANTUL_V2_BUILD_.CHUNK_ROWS, total - start + 2);
    const rows = raw.getRange(start, 1, count, MANTUL_V2_SCHEMA_.RAW.length).getValues();
    const eventRows = [], p3Grouped = new Map(), grouped = new Map();
    rows.forEach((r, i) => {
      const sourceRow = start + i;
      if (r.every(v => v === '' || v === null)) return;
      const ev = mantulV2MapRawToP2Event_(r, sourceRow);
      eventRows.push(ev);
      const p3 = mantulV2MapRawToP3Outlet_(r);
      const p3Key = p3.slice(0,15).map(mantulV2KeyPart_).join('\u001f');
      let p3g = p3Grouped.get(p3Key);
      if (!p3g) { p3g = p3.slice(0,15).concat([0,0,0]); p3Grouped.set(p3Key,p3g); }
      p3g[15] += Number(p3[15]) || 0;
      p3g[16] += Number(p3[16]) || 0;
      p3g[17] += Number(p3[17]) || 1;
      const key = ev.slice(0,14).map(mantulV2KeyPart_).join('\u001f');
      let g = grouped.get(key);
      if (!g) { g = ev.slice(0,14).concat([0,0,0]); grouped.set(key,g); }
      g[14] += Number(ev[14]) || 0;
      g[15] += Number(ev[15]) || 0;
      g[16] += Number(ev[16]) || 1;
    });
    mantulV2AppendStage_(ss, MANTUL_V2_BUILD_.EVENT_STAGE, eventRows, MANTUL_V2_SCHEMA_.P2_EVENT_INDEX.length);
    const p3PartialRows = Array.from(p3Grouped.values());
    mantulV2AppendStage_(ss, MANTUL_V2_BUILD_.P3_PARTIAL_STAGE, p3PartialRows, MANTUL_V2_SCHEMA_.P3_OUTLET_MONTHLY.length);
    const partialRows = Array.from(grouped.values());
    mantulV2AppendStage_(ss, MANTUL_V2_BUILD_.PARTIAL_STAGE, partialRows, MANTUL_V2_SCHEMA_.P2_DASHBOARD_FACT.length);
    const eventWritten = Number(props.getProperty(p+'EVENT_ROWS')||0)+eventRows.length;
    const p3Written = Number(props.getProperty(p+'P3_ROWS')||0)+p3PartialRows.reduce((n,r)=>n+Number(r[17]||0),0);
    const partialWritten = Number(props.getProperty(p+'PARTIAL_ROWS')||0)+partialRows.length;
    const next = start + count;
    props.setProperties({
      [p+'NEXT_ROW']:String(next),
      [p+'EVENT_ROWS']:String(eventWritten),
      [p+'NONEMPTY_ROWS']:String(Number(props.getProperty(p+'NONEMPTY_ROWS')||0)+eventRows.length),
      [p+'P3_ROWS']:String(p3Written),
      [p+'PARTIAL_ROWS']:String(partialWritten),
      [p+'PHASE']:next>total?'PARTIALS_WRITTEN':'EVENTS'
    });
    if (next > total) {
      props.setProperty(p+'STATUS','STAGING_PARTIALS_READY');
      props.setProperty(p+'PHASE','AGGREGATION_PENDING');
    }
    return mantulV2BuildStatus();
  } catch (err) {
    const props = PropertiesService.getScriptProperties(), p=MANTUL_V2_BUILD_.PREFIX;
    props.setProperty(p+'ERROR',String(err && err.message || err));
    throw err;
  } finally { lock.releaseLock(); }
}

function mantulV2KeyPart_(v) {
  if (v instanceof Date) return String(v.getTime());
  return String(v == null ? '' : v).trim().toUpperCase();
}

function mantulV2AppendStage_(ss, name, rows, width) {
  if (!rows.length) return;
  const sh = ss.getSheetByName(name);
  const start = Math.max(2, sh.getLastRow()+1);
  ensureSheetGrid_(sh, start+rows.length-1, width);
  sh.getRange(start,1,rows.length,width).setValues(rows);
}

function mantulV2BuildStatus() {
  const props=PropertiesService.getScriptProperties(),p=MANTUL_V2_BUILD_.PREFIX;
  const ss=getDatabase_();
  const names=[MANTUL_V2_BUILD_.EVENT_STAGE,MANTUL_V2_BUILD_.P2_STAGE,MANTUL_V2_BUILD_.P3_STAGE,MANTUL_V2_BUILD_.P3_PARTIAL_STAGE,MANTUL_V2_BUILD_.PARTIAL_STAGE];
  const sheets={};
  names.forEach(n=>{const sh=ss.getSheetByName(n);sheets[n]=sh?Math.max(0,sh.getLastRow()-1):0;});
  return {
    status:props.getProperty(p+'STATUS')||'NOT_STARTED',
    phase:props.getProperty(p+'PHASE')||'',
    sourceRows:Number(props.getProperty(p+'SOURCE_ROWS')||0),
    nextSourceRow:Number(props.getProperty(p+'NEXT_ROW')||2),
    eventRows:Number(props.getProperty(p+'EVENT_ROWS')||0),
    p3Rows:Number(props.getProperty(p+'P3_ROWS')||0),
    partialFactRows:Number(props.getProperty(p+'PARTIAL_ROWS')||0),
    stageRows:sheets,
    error:props.getProperty(p+'ERROR')||''
  };
}

/**
 * Finalize staged aggregate tables with Sheets QUERY. This is a separate,
 * explicit step so chunk ingestion can resume without a long single execution.
 */
function mantulV2BuildFinalize() {
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const status=mantulV2BuildStatus();
    if(status.status!=='STAGING_PARTIALS_READY') throw new Error('Staging belum lengkap. Status: '+status.status);
    const ss=getDatabase_(), props=PropertiesService.getScriptProperties(), p=MANTUL_V2_BUILD_.PREFIX;
    const event=ss.getSheetByName(MANTUL_V2_BUILD_.EVENT_STAGE);
    const p2partial=ss.getSheetByName(MANTUL_V2_BUILD_.PARTIAL_STAGE);
    const p3partial=ss.getSheetByName(MANTUL_V2_BUILD_.P3_PARTIAL_STAGE);
    const p2=ss.getSheetByName(MANTUL_V2_BUILD_.P2_STAGE);
    const p3=ss.getSheetByName(MANTUL_V2_BUILD_.P3_STAGE);
    const expectedNonempty=Number(props.getProperty(p+'NONEMPTY_ROWS')||0);
    if(event.getLastRow()-1!==status.eventRows || status.eventRows!==expectedNonempty)
      throw new Error('Jumlah event staging tidak sama dengan jumlah baris sumber non-kosong yang diproses. Event='+status.eventRows+', nonempty='+expectedNonempty+'.');
    p2.clearContents(); p3.clearContents();
    const q2="=QUERY('"+MANTUL_V2_BUILD_.PARTIAL_STAGE+"'!A:Q,\"select A,B,C,D,E,F,G,H,I,J,K,L,M,N,sum(O),sum(P),sum(Q) where A is not null group by A,B,C,D,E,F,G,H,I,J,K,L,M,N label sum(O) 'VAL',sum(P) 'CARTON',sum(Q) 'SOURCE_ROWS'\",1)";
    const q3="=QUERY('"+MANTUL_V2_BUILD_.P3_PARTIAL_STAGE+"'!A:R,\"select A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,sum(P),sum(Q),sum(R) where A is not null group by A,B,C,D,E,F,G,H,I,J,K,L,M,N,O label sum(P) 'VAL',sum(Q) 'CARTON',sum(R) 'SOURCE_ROWS'\",1)";
    p2.getRange(1,1).setFormula(q2); p3.getRange(1,1).setFormula(q3);
    SpreadsheetApp.flush();
    const p2Rows=Math.max(0,p2.getLastRow()-1), p3Rows=Math.max(0,p3.getLastRow()-1);
    if(!p2Rows||!p3Rows) throw new Error('Hasil agregasi P2/P3 kosong. Periksa formula QUERY dan data staging.');
    const eventTotals=mantulV2SheetTotals_(event,MANTUL_V2_SCHEMA_.P2_EVENT_INDEX.length,15,16);
    const p2Totals=mantulV2SheetTotals_(p2,MANTUL_V2_SCHEMA_.P2_DASHBOARD_FACT.length,15,16);
    const p3Totals=mantulV2SheetTotals_(p3,MANTUL_V2_SCHEMA_.P3_OUTLET_MONTHLY.length,16,17);
    const raw=ss.getSheetByName('RAW_DATA_SALES');
    const rawTotals=mantulV2SheetTotals_(raw,MANTUL_V2_SCHEMA_.RAW.length,13,8);
    const close=(a,b)=>Math.abs(a-b)<=Math.max(0.01,Math.abs(b)*1e-9);
    if(event.getLastRow()-1!==status.eventRows || eventTotals.rows!==expectedNonempty ||
       !close(eventTotals.value,rawTotals.value)||!close(eventTotals.carton,rawTotals.carton)||
       !close(p2Totals.value,rawTotals.value)||!close(p2Totals.carton,rawTotals.carton)||
       !close(p3Totals.value,rawTotals.value)||!close(p3Totals.carton,rawTotals.carton)) {
      throw new Error('Rekonsiliasi gagal. RAW='+JSON.stringify(rawTotals)+', EVENT='+JSON.stringify(eventTotals)+', P2='+JSON.stringify(p2Totals)+', P3='+JSON.stringify(p3Totals));
    }
    props.setProperties({
      [p+'STATUS']:'READY_FOR_REVIEW',
      [p+'PHASE']:'RECONCILED',
      [p+'P2_FACT_ROWS']:String(p2Rows),
      [p+'P3_FACT_ROWS']:String(p3Rows),
      [p+'RAW_VALUE']:String(rawTotals.value),
      [p+'RAW_CARTON']:String(rawTotals.carton),
      [p+'COMPLETED_AT']:new Date().toISOString()
    });
    return Object.assign(mantulV2BuildStatus(),{p2FactRows:p2Rows,p3FactRows:p3Rows,reconciliation:{raw:rawTotals,event:eventTotals,p2:p2Totals,p3:p3Totals}});
  } finally {lock.releaseLock();}
}


/**
 * Publish only a fully reconciled V2 build. Existing live sheets are renamed
 * to timestamped backups; P1_FACT and RAW_DATA_SALES are never replaced.
 */
function mantulV2PublishStaging() {
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  const ss=getDatabase_(),props=PropertiesService.getScriptProperties(),prefix=MANTUL_V2_BUILD_.PREFIX;
  const renamed=[],created=[],published=[];
  try{
    const status=mantulV2BuildStatus();
    if(status.status!=='READY_FOR_REVIEW')throw new Error('Publish ditahan. Status build harus READY_FOR_REVIEW, terbaca '+status.status+'.');
    const specs=[
      {stage:MANTUL_V2_BUILD_.EVENT_STAGE,target:'P2_EVENT_INDEX',headers:MANTUL_V2_SCHEMA_.P2_EVENT_INDEX,value:15,carton:16},
      {stage:MANTUL_V2_BUILD_.P2_STAGE,target:'P2_DASHBOARD_FACT',headers:MANTUL_V2_SCHEMA_.P2_DASHBOARD_FACT,value:15,carton:16},
      {stage:MANTUL_V2_BUILD_.P3_STAGE,target:'P3_OUTLET_MONTHLY',headers:MANTUL_V2_SCHEMA_.P3_OUTLET_MONTHLY,value:16,carton:17}
    ];
    const stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd_HHmmss');
    const temp=[];
    specs.forEach(spec=>{
      const source=ss.getSheetByName(spec.stage);
      if(!source||source.getLastRow()<2)throw new Error('Stage '+spec.stage+' kosong.');
      const actual=source.getRange(1,1,1,spec.headers.length).getDisplayValues()[0].map(v=>String(v||'').trim().toUpperCase());
      if(!spec.headers.every((v,i)=>actual[i]===String(v).toUpperCase()))throw new Error('Header stage '+spec.stage+' tidak sesuai schema.');
      const name=(spec.target+'_V2_TMP_'+stamp).slice(0,99),copy=ss.insertSheet(name);
      created.push(name);
      ensureSheetGrid_(copy,source.getLastRow(),spec.headers.length);
      // Materialize evaluated QUERY results as values; do not retain formulas
      // that point to partial staging sheets in the published database.
      const totalRows=source.getLastRow();
      for(let offset=0;offset<totalRows;offset+=MANTUL_V2_BUILD_.CHUNK_ROWS){
        const n=Math.min(MANTUL_V2_BUILD_.CHUNK_ROWS,totalRows-offset);
        const values=source.getRange(offset+1,1,n,spec.headers.length).getValues();
        copy.getRange(offset+1,1,n,spec.headers.length).setValues(values);
      }
      copy.setFrozenRows(1);
      temp.push({spec:spec,sheet:copy});
    });
    const raw=ss.getSheetByName('RAW_DATA_SALES'),rawTotals=mantulV2SheetTotals_(raw,MANTUL_V2_SCHEMA_.RAW.length,13,8);
    const totals=temp.map(x=>mantulV2SheetTotals_(x.sheet,x.spec.headers.length,x.spec.value,x.spec.carton));
    const close=(a,b)=>Math.abs(a-b)<=Math.max(.01,Math.abs(b)*1e-9);
    if(totals[0].rows!==status.eventRows||
       totals.some(t=>!close(t.value,rawTotals.value)||!close(t.carton,rawTotals.carton)))
      throw new Error('Validasi sebelum publish gagal. RAW='+JSON.stringify(rawTotals)+'; staged='+JSON.stringify(totals));
    temp.forEach(({spec,sheet})=>{
      const old=ss.getSheetByName(spec.target);
      if(old){const backup=(spec.target+'_V2_BAK_'+stamp).slice(0,99);old.setName(backup);renamed.push({target:spec.target,backup:backup});}
      published.push({target:spec.target,hadOld:!!old});
      sheet.setName(spec.target);
    });
    props.deleteProperty('MANTUL_V2_DATA_STALE');
    props.setProperties({
      [prefix+'STATUS']:'PUBLISHED',
      [prefix+'PHASE']:'LIVE',
      [prefix+'PUBLISHED_AT']:new Date().toISOString()
    });
    return {ok:true,status:'PUBLISHED',publishedAt:props.getProperty(prefix+'PUBLISHED_AT'),
      sheets:specs.map(x=>x.target),backups:renamed.map(x=>x.backup),
      rows:{P2_EVENT_INDEX:totals[0].rows,P2_DASHBOARD_FACT:totals[1].rows,P3_OUTLET_MONTHLY:totals[2].rows},
      reconciliation:{raw:rawTotals,event:totals[0],p2:totals[1],p3:totals[2]}};
  }catch(err){
    // Restore every original live sheet if any rename in the publish sequence failed.
    published.slice().reverse().forEach(x=>{
      const current=ss.getSheetByName(x.target),backup=renamed.find(b=>b.target===x.target);
      if(current){
        if(backup){const failed=(x.target+'_V2_FAILED_'+Date.now()).slice(0,99);current.setName(failed);}
        else ss.deleteSheet(current);
      }
      if(backup){const oldSheet=ss.getSheetByName(backup.backup);if(oldSheet&&!ss.getSheetByName(x.target))oldSheet.setName(x.target);}
    });
    created.forEach(name=>{const sh=ss.getSheetByName(name);if(sh)ss.deleteSheet(sh);});
    throw err;
  }finally{lock.releaseLock();}
}

function mantulV2LastRawDataRow_(sheet) {
  if(!sheet)return 1;
  const last=sheet.getRange(sheet.getMaxRows(),1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  return Math.max(1,last);
}

function mantulV2SheetTotals_(sheet,width,valueCol,cartonCol) {
  const last=sheet.getName()==='RAW_DATA_SALES'?mantulV2LastRawDataRow_(sheet):sheet.getLastRow();
  let value=0,carton=0;
  for(let row=2;row<=last;row+=5000){
    const n=Math.min(5000,last-row+1);
    const vals=sheet.getRange(row,1,n,width).getValues();
    vals.forEach(r=>{value+=Number(r[valueCol-1])||0;carton+=Number(r[cartonCol-1])||0;});
  }
  return {rows:Math.max(0,last-1),value:value,carton:carton};
}
