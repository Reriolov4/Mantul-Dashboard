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
    const total = raw.getLastRow() - 1;
    props.setProperties({
      [p+'STATUS']:'RUNNING',
      [p+'PHASE']:'EVENTS',
      [p+'SOURCE_ID']:ss.getId(),
      [p+'SOURCE_ROWS']:String(total),
      [p+'NEXT_ROW']:'2',
      [p+'EVENT_ROWS']:'0',
      [p+'P3_ROWS']:'0',
      [p+'PARTIAL_ROWS']:'0',
      [p+'STARTED_AT']:new Date().toISOString(),
      [p+'ERROR']:''
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
      const p3Key = p3.slice(0,15).map(mantulV2KeyPart_).join('\\u001f');
      let p3g = p3Grouped.get(p3Key);
      if (!p3g) { p3g = p3.slice(0,15).concat([0,0,0]); p3Grouped.set(p3Key,p3g); }
      p3g[15] += Number(p3[15]) || 0;
      p3g[16] += Number(p3[16]) || 0;
      p3g[17] += 1;
      const key = ev.slice(0,14).map(mantulV2KeyPart_).join('\u001f');
      let g = grouped.get(key);
      if (!g) { g = ev.slice(0,14).concat([0,0,0]); grouped.set(key,g); }
      g[14] += Number(ev[14]) || 0;
      g[15] += Number(ev[15]) || 0;
      g[16] += 1;
    });
    mantulV2AppendStage_(ss, MANTUL_V2_BUILD_.EVENT_STAGE, eventRows, 19);
    const p3PartialRows = Array.from(p3Grouped.values());
    mantulV2AppendStage_(ss, MANTUL_V2_BUILD_.P3_PARTIAL_STAGE, p3PartialRows, 18);
    const partialRows = Array.from(grouped.values());
    mantulV2AppendStage_(ss, MANTUL_V2_BUILD_.PARTIAL_STAGE, partialRows, 17);
    const eventWritten = Number(props.getProperty(p+'EVENT_ROWS')||0)+eventRows.length;
    const p3Written = Number(props.getProperty(p+'P3_ROWS')||0)+p3PartialRows.reduce((n,r)=>n+Number(r[17]||0),0);
    const partialWritten = Number(props.getProperty(p+'PARTIAL_ROWS')||0)+partialRows.length;
    const next = start + count;
    props.setProperties({
      [p+'NEXT_ROW']:String(next),
      [p+'EVENT_ROWS']:String(eventWritten),
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
    if(event.getLastRow()-1!==status.eventRows) throw new Error('Jumlah event staging tidak sama dengan jumlah baris sumber non-kosong yang diproses.');
    p2.clearContents(); p3.clearContents();
    const q2="=QUERY('"+MANTUL_V2_BUILD_.PARTIAL_STAGE+"'!A:Q,\"select A,B,C,D,E,F,G,H,I,J,K,L,M,N,sum(O),sum(P),sum(Q) where A is not null group by A,B,C,D,E,F,G,H,I,J,K,L,M,N\",1)";
    const q3="=QUERY('"+MANTUL_V2_BUILD_.P3_PARTIAL_STAGE+"'!A:R,\"select A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,sum(P),sum(Q),sum(R) where A is not null group by A,B,C,D,E,F,G,H,I,J,K,L,M,N,O\",1)";
    p2.getRange(1,1).setFormula(q2); p3.getRange(1,1).setFormula(q3);
    SpreadsheetApp.flush();
    const p2Rows=Math.max(0,p2.getLastRow()-1), p3Rows=Math.max(0,p3.getLastRow()-1);
    if(!p2Rows||!p3Rows) throw new Error('Hasil agregasi P2/P3 kosong. Periksa formula QUERY dan data staging.');
    const eventTotals=mantulV2SheetTotals_(event,19,15,16);
    const p2Totals=mantulV2SheetTotals_(p2,17,15,16);
    const p3Totals=mantulV2SheetTotals_(p3,18,16,17);
    const raw=ss.getSheetByName('RAW_DATA_SALES');
    const rawTotals=mantulV2SheetTotals_(raw,27,12,7);
    const close=(a,b)=>Math.abs(a-b)<=Math.max(0.01,Math.abs(b)*1e-9);
    if(event.getLastRow()-1!==status.eventRows ||
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

function mantulV2SheetTotals_(sheet,width,valueCol,cartonCol) {
  const last=sheet.getLastRow(); let value=0,carton=0;
  for(let row=2;row<=last;row+=5000){
    const n=Math.min(5000,last-row+1);
    const vals=sheet.getRange(row,1,n,width).getValues();
    vals.forEach(r=>{value+=Number(r[valueCol-1])||0;carton+=Number(r[cartonCol-1])||0;});
  }
  return {rows:Math.max(0,last-1),value:value,carton:carton};
}
