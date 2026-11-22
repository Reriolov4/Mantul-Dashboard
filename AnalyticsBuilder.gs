/**
 * MANTUL ANALYTICS — Phase 1: normalized analytics store + resumable initial build.
 *
 * This phase is intentionally isolated from dashboard APIs. It creates a separate
 * spreadsheet and copies normalized, dashboard-relevant facts in bounded batches.
 * It does NOT change Page 1/Page 2 rendering or KPI calculations.
 *
 * Run manually:
 *   1) analyticsStartInitialBuild()
 *   2) Repeat analyticsBuildNextChunk() until status.complete === true
 *   3) analyticsGetBuildStatus()
 */

const ANALYTICS_DB_PROPERTY_ = 'MANTUL_ANALYTICS_SPREADSHEET_ID';
const ANALYTICS_BUILD_PREFIX_ = 'MANTUL_ANALYTICS_BUILD_';
const ANALYTICS_BATCH_SIZE_ = 4000;

const P1_FACT_HEADERS_ = [
  'YEAR','MONTH','SUBDIST','SUBBRAND','BRAND','CATEGORY','SBU','PIC',
  'PF_MONTHLY','PF_QUARTERLY','VAL','SALES_TYPE','TARGET','BE'
];

const P2_EVENT_HEADERS_ = [
  'YEAR','MONTH','DD','CUSTOMER_CODE','CUSTOMER_NAME','SUBDIST','PIC',
  'BRAND','SUBBRAND','CATEGORY','SBU','SALES_TYPE','SALESMAN','TEAM'
];

function analyticsStartInitialBuild() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const source = getDatabase_();
    const p1 = source.getSheetByName(DATA_SHEET_NAME);
    const p2 = source.getSheetByName(RAW_DATA_SHEET_NAME);
    if (!p1 || p1.getLastRow() < 2) throw new Error('DATA_SALES tidak ditemukan atau kosong.');
    if (!p2 || p2.getLastRow() < 2) throw new Error('RAW_DATA_SALES tidak ditemukan atau kosong.');

    const props = PropertiesService.getScriptProperties();
    let analyticsId = String(props.getProperty(ANALYTICS_DB_PROPERTY_) || '').trim();
    let analytics = null;
    if (analyticsId) {
      try { analytics = SpreadsheetApp.openById(analyticsId); } catch (e) { analytics = null; }
    }
    if (!analytics) {
      analytics = SpreadsheetApp.create('MANTUL_ANALYTICS_DB');
      analyticsId = analytics.getId();
      props.setProperty(ANALYTICS_DB_PROPERTY_, analyticsId);
    }

    analyticsPrepareSheet_(analytics, 'P1_FACT', P1_FACT_HEADERS_);
    analyticsPrepareSheet_(analytics, 'P2_EVENT_INDEX', P2_EVENT_HEADERS_);
    analyticsPrepareSheet_(analytics, 'SYSTEM_STATE', ['KEY','VALUE','UPDATED_AT']);

    const p1Headers = p1.getRange(1,1,1,p1.getLastColumn()).getValues()[0];
    validateDatabaseHeaders_(p1Headers);
    const p1Ix = analyticsResolveP1Columns_(p1Headers);
    const rawHeaders = p2.getRange(1,1,1,RAW_SALES_COLUMN_COUNT).getValues()[0];
    validateRawDatabaseHeaders_(rawHeaders);

    const now = new Date().toISOString();
    props.setProperties({
      [ANALYTICS_BUILD_PREFIX_ + 'STATUS']: 'RUNNING',
      [ANALYTICS_BUILD_PREFIX_ + 'PHASE']: 'P1',
      [ANALYTICS_BUILD_PREFIX_ + 'P1_NEXT_ROW']: '2',
      [ANALYTICS_BUILD_PREFIX_ + 'P2_NEXT_ROW']: '2',
      [ANALYTICS_BUILD_PREFIX_ + 'P1_SOURCE_ROWS']: String(p1.getLastRow()-1),
      [ANALYTICS_BUILD_PREFIX_ + 'P2_SOURCE_ROWS']: String(p2.getLastRow()-1),
      [ANALYTICS_BUILD_PREFIX_ + 'P1_WRITTEN']: '0',
      [ANALYTICS_BUILD_PREFIX_ + 'P2_WRITTEN']: '0',
      [ANALYTICS_BUILD_PREFIX_ + 'STARTED_AT']: now,
      [ANALYTICS_BUILD_PREFIX_ + 'SOURCE_ID']: source.getId()
    });
    analytics.getSheetByName('SYSTEM_STATE').getRange(2,1,6,2).setValues([
      ['BUILD_STATUS','RUNNING'],['BUILD_STARTED_AT',now],
      ['SOURCE_SPREADSHEET_ID',source.getId()],
      ['SOURCE_DATA_SALES_ROWS',String(p1.getLastRow()-1)],
      ['SOURCE_RAW_DATA_SALES_ROWS',String(p2.getLastRow()-1)],
      ['SCHEMA_VERSION','1']
    ]);
    return analyticsGetBuildStatus();
  } finally {
    lock.releaseLock();
  }
}

function analyticsBuildNextChunk() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const prefix = ANALYTICS_BUILD_PREFIX_;
    if (props.getProperty(prefix+'STATUS') !== 'RUNNING') {
      throw new Error('Belum ada initial build aktif. Jalankan analyticsStartInitialBuild() terlebih dahulu.');
    }
    const sourceId = props.getProperty(prefix+'SOURCE_ID');
    const source = SpreadsheetApp.openById(sourceId);
    const analytics = analyticsOpenDb_();
    let phase = props.getProperty(prefix+'PHASE') || 'P1';
    const started = Date.now();

    if (phase === 'P1') {
      const sh = source.getSheetByName(DATA_SHEET_NAME);
      const next = Number(props.getProperty(prefix+'P1_NEXT_ROW') || 2);
      const last = Number(props.getProperty(prefix+'P1_SOURCE_ROWS') || 0) + 1;
      if (next > last) {
        props.setProperty(prefix+'PHASE','P2');
        phase = 'P2';
      } else {
        const count = Math.min(ANALYTICS_BATCH_SIZE_, last-next+1);
        const width = Math.max(sh.getLastColumn(), SALES_HEADERS.length);
        const values = sh.getRange(next,1,count,width).getValues();
        const header = sh.getRange(1,1,1,width).getValues()[0];
        const ix = analyticsResolveP1Columns_(header);
        const output = [];
        values.forEach(r => {
          if (!r.some(v => v !== '' && v !== null)) return;
          const year = Number(r[ix.y]), val = Number(r[ix.v]);
          if (!Number.isFinite(year) || !Number.isFinite(val)) return;
          output.push([
            year, String(r[ix.m]||'').trim().toUpperCase(), String(r[ix.d]||''),
            String(r[ix.s]||''), String(r[ix.b]||''), String(r[ix.c]||''),
            String(r[ix.u]||''), String(r[ix.p]||''), String(r[ix.fm]||''),
            String(r[ix.fq]||''), val, String(r[ix.t]||'').trim().toUpperCase(),
            ix.tg < 0 || r[ix.tg] === '' || r[ix.tg] == null ? '' : Number(r[ix.tg]),
            ix.be < 0 || r[ix.be] === '' || r[ix.be] == null ? '' : Number(r[ix.be])
          ]);
        });
        if (output.length) analyticsAppendRows_(analytics.getSheetByName('P1_FACT'), output, P1_FACT_HEADERS_.length);
        props.setProperties({
          [prefix+'P1_NEXT_ROW']:String(next+count),
          [prefix+'P1_WRITTEN']:String(Number(props.getProperty(prefix+'P1_WRITTEN')||0)+output.length)
        });
        return Object.assign(analyticsGetBuildStatus(), {lastBatch:{phase:'P1',sourceStart:next,sourceEnd:next+count-1,written:output.length,elapsedMs:Date.now()-started}});
      }
    }

    if (phase === 'P2') {
      const sh = source.getSheetByName(RAW_DATA_SHEET_NAME);
      const next = Number(props.getProperty(prefix+'P2_NEXT_ROW') || 2);
      const last = Number(props.getProperty(prefix+'P2_SOURCE_ROWS') || 0) + 1;
      if (next > last) {
        props.setProperty(prefix+'STATUS','COMPLETE');
        props.setProperty(prefix+'PHASE','COMPLETE');
        const state = analytics.getSheetByName('SYSTEM_STATE');
        state.getRange(2,1,8,2).setValues([
          ['BUILD_STATUS','COMPLETE'],
          ['BUILD_STARTED_AT',props.getProperty(prefix+'STARTED_AT')||''],
          ['BUILD_COMPLETED_AT',new Date().toISOString()],
          ['SOURCE_SPREADSHEET_ID',sourceId],
          ['P1_FACT_ROWS',props.getProperty(prefix+'P1_WRITTEN')||'0'],
          ['P2_EVENT_INDEX_ROWS',props.getProperty(prefix+'P2_WRITTEN')||'0'],
          ['SCHEMA_VERSION','1'],
          ['SOURCE_VERSION',source.getId()+'|'+sh.getLastRow()+'|'+source.getSheetId()]
        ]);
        return analyticsGetBuildStatus();
      }
      const count = Math.min(ANALYTICS_BATCH_SIZE_, last-next+1);
      const values = sh.getRange(next,1,count,RAW_SALES_COLUMN_COUNT).getValues();
      const output = [];
      values.forEach(r => {
        const salesType = String(r[RAW_SALES_COL.SALES_TYPE]||'').trim().toUpperCase();
        if (salesType !== 'SOO') return;
        output.push([
          r[RAW_SALES_COL.YEAR], String(r[RAW_SALES_COL.MONTH]||'').trim().toUpperCase(),
          r[RAW_SALES_COL.DAY], String(r[RAW_SALES_COL.CUSTOMER_CODE]||'').trim(),
          String(r[RAW_SALES_COL.CUSTOMER_NAME]||'').trim(),
          String(r[RAW_SALES_COL.SUBDIST]||'').trim(),
          String(r[RAW_SALES_COL.PIC]||'').trim(),
          String(r[RAW_SALES_COL.BRAND]||'').trim(),
          String(r[RAW_SALES_COL.SUBBRAND]||'').trim(),
          String(r[RAW_SALES_COL.CATEGORY]||'').trim(),
          String(r[RAW_SALES_COL.SBU]||'').trim(),
          salesType, String(r[RAW_SALES_COL.SALESMAN]||'').trim(),
          String(r[RAW_SALES_COL.TEAM]||'').trim()
        ]);
      });
      if (output.length) analyticsAppendRows_(analytics.getSheetByName('P2_EVENT_INDEX'), output, P2_EVENT_HEADERS_.length);
      props.setProperties({
        [prefix+'P2_NEXT_ROW']:String(next+count),
        [prefix+'P2_WRITTEN']:String(Number(props.getProperty(prefix+'P2_WRITTEN')||0)+output.length)
      });
      return Object.assign(analyticsGetBuildStatus(), {lastBatch:{phase:'P2',sourceStart:next,sourceEnd:next+count-1,written:output.length,elapsedMs:Date.now()-started}});
    }
    return analyticsGetBuildStatus();
  } finally {
    lock.releaseLock();
  }
}

function analyticsGetBuildStatus() {
  const props = PropertiesService.getScriptProperties();
  const p = ANALYTICS_BUILD_PREFIX_;
  const analyticsId = props.getProperty(ANALYTICS_DB_PROPERTY_) || '';
  const p1Total = Number(props.getProperty(p+'P1_SOURCE_ROWS')||0);
  const p2Total = Number(props.getProperty(p+'P2_SOURCE_ROWS')||0);
  const p1Written = Number(props.getProperty(p+'P1_WRITTEN')||0);
  const p2Written = Number(props.getProperty(p+'P2_WRITTEN')||0);
  return {
    status:props.getProperty(p+'STATUS')||'NOT_STARTED',
    phase:props.getProperty(p+'PHASE')||'',
    analyticsSpreadsheetId:analyticsId,
    analyticsUrl:analyticsId?'https://docs.google.com/spreadsheets/d/'+analyticsId+'/edit':'',
    p1:{sourceRows:p1Total,processedThrough:Number(props.getProperty(p+'P1_NEXT_ROW')||2)-2,writtenRows:p1Written},
    p2:{sourceRows:p2Total,processedThrough:Number(props.getProperty(p+'P2_NEXT_ROW')||2)-2,writtenRows:p2Written}
  };
}

function analyticsOpenDb_() {
  const id = String(PropertiesService.getScriptProperties().getProperty(ANALYTICS_DB_PROPERTY_)||'').trim();
  if (!id) throw new Error('Analytics spreadsheet belum dibuat.');
  return SpreadsheetApp.openById(id);
}

function analyticsPrepareSheet_(ss,name,headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();
  const maxRows = sh.getMaxRows();
  if (maxRows < 2) sh.insertRowsAfter(maxRows,2-maxRows);
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

function analyticsAppendRows_(sheet,rows,width) {
  if (!rows.length) return;
  const start = sheet.getLastRow()+1;
  const needed = start+rows.length-1;
  const maxRows = sheet.getMaxRows();
  if (maxRows < needed) sheet.insertRowsAfter(maxRows,needed-maxRows);
  sheet.getRange(start,1,rows.length,width).setValues(rows);
}

function analyticsResolveP1Columns_(headers) {
  const key = v => String(v==null?'':v).trim().toLowerCase().replace(/[_\\s]+/g,' ');
  const map = {};
  headers.forEach((v,i)=>map[key(v)]=i);
  const col = (aliases,required) => {
    for (const a of aliases) if (map[key(a)] !== undefined) return map[key(a)];
    if (required) throw new Error('Kolom DATA_SALES tidak ditemukan: '+aliases[0]);
    return -1;
  };
  return {
    y:col(['YEAR','y'],true),m:col(['MONTH','m'],true),d:col(['SUBDIST','d'],true),
    s:col(['Subbrand','SUBBRAND','s'],true),b:col(['Brand','BRAND','b'],true),
    c:col(['CATEGORY','c'],true),u:col(['SBU','u'],true),p:col(['PIC','p'],true),
    fm:col(['PF Bulanan','fm'],true),fq:col(['PF Triwulan','fq'],true),
    v:col(['VAL','v'],true),t:col(['JENIS SALES','t'],true),
    tg:col(['TARGET','tg'],false),be:col(['BE','be'],false)
  };
}


/**
 * Return a verified analytics sheet for dashboard reads.
 * The source row-count check prevents using an index after append-only source updates.
 * Returns null when the build is missing, incomplete, or stale; callers then use live sheets.
 */
function analyticsGetReadySheet_(analyticsSheetName, sourceSheetName, phase) {
  const props = PropertiesService.getScriptProperties();
  const prefix = ANALYTICS_BUILD_PREFIX_;
  if (props.getProperty(prefix + 'STATUS') !== 'COMPLETE') return null;

  const sourceId = String(props.getProperty(prefix + 'SOURCE_ID') || '').trim();
  if (!sourceId) return null;

  const source = getDatabase_();
  if (source.getId() !== sourceId) return null;

  const sourceSheet = source.getSheetByName(sourceSheetName);
  if (!sourceSheet || sourceSheet.getLastRow() < 2) return null;

  const sourceCountKey = prefix + phase + '_SOURCE_ROWS';
  const writtenKey = prefix + phase + '_WRITTEN';
  const expectedSourceRows = Number(props.getProperty(sourceCountKey) || 0);
  const expectedWrittenRows = Number(props.getProperty(writtenKey) || 0);
  if (sourceSheet.getLastRow() - 1 !== expectedSourceRows) return null;

  let analytics;
  try {
    analytics = analyticsOpenDb_();
  } catch (err) {
    console.error('Analytics DB tidak dapat dibuka; dashboard memakai sumber live: ' + err.message);
    return null;
  }
  const target = analytics.getSheetByName(analyticsSheetName);
  if (!target || target.getLastRow() - 1 !== expectedWrittenRows) return null;
  return target;
}

/** Read-only route diagnostic for the dashboard migration. */
function testAnalyticsDashboardRouting() {
  const source = getDatabase_();
  const p1Source = source.getSheetByName(DATA_SHEET_NAME);
  const p2Source = source.getSheetByName(RAW_DATA_SHEET_NAME);
  const p1 = analyticsGetReadySheet_('P1_FACT', DATA_SHEET_NAME, 'P1');
  const p2 = analyticsGetReadySheet_('P2_EVENT_INDEX', RAW_DATA_SHEET_NAME, 'P2');
  const result = {
    buildStatus: PropertiesService.getScriptProperties().getProperty(ANALYTICS_BUILD_PREFIX_ + 'STATUS') || 'NOT_STARTED',
    sourceSpreadsheetId: source.getId(),
    p1: {
      route: p1 ? 'P1_FACT' : 'DATA_SALES fallback',
      sourceRows: p1Source ? p1Source.getLastRow() - 1 : 0,
      analyticsRows: p1 ? p1.getLastRow() - 1 : 0
    },
    p2: {
      route: p2 ? 'P2_EVENT_INDEX' : 'RAW_DATA_SALES fallback',
      sourceRows: p2Source ? p2Source.getLastRow() - 1 : 0,
      analyticsRows: p2 ? p2.getLastRow() - 1 : 0
    }
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}


/**
 * Read-only parity test: compare the current analytics facts with the original
 * DATA_SALES / RAW_DATA_SALES calculations before relying on the new route.
 * This does not write to either spreadsheet.
 */
function testAnalyticsDashboardDataParity() {
  const source = getDatabase_();
  const analytics = analyticsOpenDb_();
  const p1Source = source.getSheetByName(DATA_SHEET_NAME);
  const p2Source = source.getSheetByName(RAW_DATA_SHEET_NAME);
  const p1Fact = analytics.getSheetByName('P1_FACT');
  const p2Index = analytics.getSheetByName('P2_EVENT_INDEX');
  if (!p1Source || !p2Source || !p1Fact || !p2Index) {
    throw new Error('Sheet sumber atau Analytics DB tidak lengkap.');
  }

  // PAGE 1: compare every normalized fact row and all 14 fields in order.
  const p1Last = p1Source.getLastRow();
  const p1Width = Math.max(p1Source.getLastColumn(), SALES_HEADERS.length);
  const p1Headers = p1Source.getRange(1, 1, 1, p1Width).getValues()[0];
  const p1Ix = analyticsResolveP1Columns_(p1Headers);
  const p1RawValues = p1Last > 1
    ? p1Source.getRange(2, 1, p1Last - 1, p1Width).getValues()
    : [];
  const p1Expected = [];
  p1RawValues.forEach(r => {
    if (!r.some(v => v !== '' && v !== null)) return;
    const year = Number(r[p1Ix.y]), val = Number(r[p1Ix.v]);
    if (!Number.isFinite(year) || !Number.isFinite(val)) return;
    p1Expected.push([
      year, String(r[p1Ix.m] || '').trim().toUpperCase(),
      String(r[p1Ix.d] || ''), String(r[p1Ix.s] || ''),
      String(r[p1Ix.b] || ''), String(r[p1Ix.c] || ''),
      String(r[p1Ix.u] || ''), String(r[p1Ix.p] || ''),
      String(r[p1Ix.fm] || ''), String(r[p1Ix.fq] || ''), val,
      String(r[p1Ix.t] || '').trim().toUpperCase(),
      p1Ix.tg < 0 || r[p1Ix.tg] === '' || r[p1Ix.tg] == null ? '' : Number(r[p1Ix.tg]),
      p1Ix.be < 0 || r[p1Ix.be] === '' || r[p1Ix.be] == null ? '' : Number(r[p1Ix.be])
    ]);
  });
  const p1Actual = p1Fact.getLastRow() > 1
    ? p1Fact.getRange(2, 1, p1Fact.getLastRow() - 1, P1_FACT_HEADERS_.length).getValues()
    : [];
  let p1Mismatch = -1;
  const p1Count = Math.min(p1Expected.length, p1Actual.length);
  for (let i = 0; i < p1Count; i++) {
    if (JSON.stringify(p1Expected[i]) !== JSON.stringify(p1Actual[i])) {
      p1Mismatch = i + 1;
      break;
    }
  }
  const p1Pass = p1Expected.length === p1Actual.length && p1Mismatch < 0;

  // PAGE 2: calculate the same KPI/chart payload from the live source and
  // the normalized SOO index, using the existing production aggregation logic.
  const p2SourceRows = page2Read_(p2Source);
  const p2IndexRows = page2Read_(p2Index);
  const years = Array.from(new Set(p2IndexRows.map(r => Number(r[0])).filter(Number.isFinite))).sort((a,b) => a-b);
  const p2Checks = years.map(year => {
    const live = page2BuildDashboardData_({year:year}, p2SourceRows);
    const indexed = page2BuildDashboardData_({year:year}, p2IndexRows);
    const sections = ['kpi','monthly','daily','donut','sbu','byTeam'];
    const mismatches = sections.filter(key =>
      JSON.stringify(live[key]) !== JSON.stringify(indexed[key])
    );
    return {
      year:year,
      status:mismatches.length ? 'MISMATCH' : 'PASS',
      mismatchedSections:mismatches,
      liveKpi:live.kpi,
      indexedKpi:indexed.kpi
    };
  });
  const result = {
    sourceSpreadsheetId:source.getId(),
    page1:{
      status:p1Pass ? 'PASS' : 'MISMATCH',
      sourceNormalizedRows:p1Expected.length,
      analyticsRows:p1Actual.length,
      firstMismatchRow:p1Mismatch < 0 ? null : p1Mismatch
    },
    page2:{
      sourceRows:p2Source.getLastRow() - 1,
      sourceSooRows:p2SourceRows.filter(r => page2Key_(r[7]) === 'SOO').length,
      analyticsRows:p2IndexRows.length,
      checks:p2Checks,
      status:p2Checks.every(x => x.status === 'PASS') ? 'PASS' : 'MISMATCH'
    }
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}


/**
 * Read-only audit for Page 2 EC discrepancies in 2026.
 * Reports current production-rule EC keys and records excluded by each guard.
 * Does not write to either spreadsheet or change dashboard calculations.
 */
function analyticsAuditEc2026Detail() {
  const source = getDatabase_();
  const sh = source.getSheetByName(RAW_DATA_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) throw new Error('RAW_DATA_SALES tidak ditemukan atau kosong.');

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, RAW_SALES_COLUMN_COUNT).getValues();
  const norm = value => String(value == null ? '' : value).trim().toUpperCase().replace(/\s+/g, ' ');
  const months = PAGE2_MONTHS_;
  const overall = {
    sooRows:0, blankCodeRows:0, othersbRows:0, validCodeRows:0,
    blankCustomerNameRows:0, invalidDayRows:0, blankSubdistRows:0,
    validEcRows:0, ecIfNameNotRequired:0, validAoRows:0
  };
  const ecKeys = new Set(), ecWithoutName = new Set(), aoKeys = new Set();
  const monthly = {};
  months.forEach(m => monthly[m] = {
    rows:0,validCodeRows:0,ec:0,ecIfNameNotRequired:0,ao:0,
    blankCodeRows:0,othersbRows:0,blankCustomerNameRows:0,invalidDayRows:0,blankSubdistRows:0
  });
  const samples = {blankCustomerName:[],blankSubdist:[],invalidDay:[],othersb:[]};

  rows.forEach((r, index) => {
    if (Number(r[RAW_SALES_COL.YEAR]) !== 2026 || norm(r[RAW_SALES_COL.SALES_TYPE]) !== 'SOO') return;
    overall.sooRows++;
    const month = norm(r[RAW_SALES_COL.MONTH]).slice(0,3);
    const m = monthly[month] || null;
    if (m) m.rows++;

    const code = norm(r[RAW_SALES_COL.CUSTOMER_CODE]);
    const name = norm(r[RAW_SALES_COL.CUSTOMER_NAME]);
    const subdist = norm(r[RAW_SALES_COL.SUBDIST]);
    const day = page2Day_(r[RAW_SALES_COL.DAY]);
    if (!code) {
      overall.blankCodeRows++;
      if (m) m.blankCodeRows++;
      return;
    }
    if (code === 'OTHERSB') {
      overall.othersbRows++;
      if (m) m.othersbRows++;
      if (samples.othersb.length < 10) samples.othersb.push({sheetRow:index+2,month:month,day:r[RAW_SALES_COL.DAY],customer:name,subdist:subdist,code:code});
      return;
    }

    overall.validCodeRows++;
    if (m) m.validCodeRows++;
    if (!name) {
      overall.blankCustomerNameRows++;
      if (m) m.blankCustomerNameRows++;
      if (samples.blankCustomerName.length < 10) samples.blankCustomerName.push({sheetRow:index+2,month:month,day:r[RAW_SALES_COL.DAY],code:code,subdist:subdist});
    }
    if (!day) {
      overall.invalidDayRows++;
      if (m) m.invalidDayRows++;
      if (samples.invalidDay.length < 10) samples.invalidDay.push({sheetRow:index+2,month:month,customer:name,code:code,subdist:subdist,rawDay:r[RAW_SALES_COL.DAY]});
      return;
    }

    const eventKey = page2CompositeKey_([day,code,subdist]);
    const aoKey = page2CompositeKey_([code,subdist]);
    if (m) {
      // Alternative count without the production customer-name guard.
      if (!m._ecAll) m._ecAll = new Set();
      m._ecAll.add(eventKey);
    }
    if (!name) {
      ecWithoutName.add(eventKey);
      return; // Production KPI excludes rows without CUSTOMER NAME.
    }

    ecKeys.add(eventKey);
    aoKeys.add(aoKey);
    if (m) {
      if (!m._ec) m._ec = new Set();
      if (!m._ao) m._ao = new Set();
      m._ec.add(eventKey);
      m._ao.add(aoKey);
    }
    if (!subdist) {
      overall.blankSubdistRows++;
      if (m) m.blankSubdistRows++;
      if (samples.blankSubdist.length < 10) samples.blankSubdist.push({sheetRow:index+2,month:month,day:day,customer:name,code:code});
    }
  });

  overall.validEcRows = ecKeys.size;
  overall.ecIfNameNotRequired = new Set([...ecKeys, ...ecWithoutName]).size;
  overall.validAoRows = aoKeys.size;
  months.forEach(month => {
    const m = monthly[month];
    m.ec = m._ec ? m._ec.size : 0;
    m.ecIfNameNotRequired = m._ecAll ? m._ecAll.size : 0;
    m.ao = m._ao ? m._ao.size : 0;
    delete m._ec; delete m._ecAll; delete m._ao;
  });

  const result = {
    sourceSpreadsheetId:source.getId(),
    year:2026,
    rule:'SOO; valid KODE CUSTOMER (not blank/OTHERSB); EC distinct DD + KODE CUSTOMER + SUBDIST; customer name and day required, matching Page2 production logic',
    overall:overall,
    monthly:monthly,
    excludedOrEdgeCaseSamples:samples
  };
  Logger.log(JSON.stringify(result,null,2));
  return result;
}