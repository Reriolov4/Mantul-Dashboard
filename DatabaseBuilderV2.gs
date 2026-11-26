/**
 * MANTUL DATABASE V2 — RESUMABLE STAGING BUILDER
 *
 * Safety: this builder writes only *_V2_STAGE sheets. It never clears or
 * overwrites live P1_FACT, P2_EVENT_INDEX, P2_DASHBOARD_FACT or P3_OUTLET_MONTHLY.
 * Run mantulV2BuildStart(), then repeat mantulV2BuildContinue() until READY.
 */
const MANTUL_V2_BUILD_ = Object.freeze({
  PREFIX: 'MANTUL_V2_BUILD_',
  CHUNK_ROWS: 2500,
  EVENT_STAGE: 'P2_EVENT_INDEX_V2_STAGE',
  P2_STAGE: 'P2_DASHBOARD_FACT_V2_STAGE',
  P3_STAGE: 'P3_OUTLET_MONTHLY_V2_STAGE',
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
    const eventRows = [], outletRows = [], grouped = new Map();
    rows.forEach((r, i) => {
      const sourceRow = start + i;
      if (r.every(v => v === '' || v === null)) return;
      const ev = mantulV2MapRawToP2Event_(r, sourceRow);
      eventRows.push(ev);
      const p3 = mantulV2MapRawToP3Outlet_(r);
      outletRows.push(p3);
      const key = ev.slice(0,14).map(mantulV2KeyPart_).join('\u001f');
      let g = grouped.get(key);
      if (!g) { g = ev.slice(0,14).concat([0,0,0]); grouped.set(key,g); }
      g[14] += Number(ev[14]) || 0;
      g[15] += Number(ev[15]) || 0;
      g[16] += 1;
    });
    mantulV2AppendStage_(ss, MANTUL_V2_BUILD_.EVENT_STAGE, eventRows, 19);
    mantulV2AppendStage_(ss, MANTUL_V2_BUILD_.P3_STAGE, outletRows, 18);
    const partialRows = Array.from(grouped.values());
    mantulV2AppendStage_(ss, MANTUL_V2_BUILD_.PARTIAL_STAGE, partialRows, 17);
    const eventWritten = Number(props.getProperty(p+'EVENT_ROWS')||0)+eventRows.length;
    const p3Written = Number(props.getProperty(p+'P3_ROWS')||0)+outletRows.length;
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
  const names=[MANTUL_V2_BUILD_.EVENT_STAGE,MANTUL_V2_BUILD_.P2_STAGE,MANTUL_V2_BUILD_.P3_STAGE,MANTUL_V2_BUILD_.PARTIAL_STAGE];
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
