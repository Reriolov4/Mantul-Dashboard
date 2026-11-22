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
