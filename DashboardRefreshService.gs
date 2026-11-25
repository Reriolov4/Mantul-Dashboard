/**
 * MANTUL — MANUAL SOURCE REFRESH
 *
 * P1_FACT, P2_EVENT_INDEX, and RAW_DATA_SALES are manually maintained inputs.
 * This function validates them and rebuilds only the derived Page 2 / Page 3
 * read models. It never writes to or clears any of the three input sheets.
 */
function refreshMantulDashboard() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  const props = PropertiesService.getScriptProperties();
  const oldSalesUpdated = props.getProperty('SALES_LAST_UPDATED') || '';
  const oldPage2Updated = props.getProperty('PAGE2_RAW_UPDATED') || '';
  const started = Date.now();

  try {
    const ss = getDatabase_();
    const p1 = ss.getSheetByName('P1_FACT');
    const p2 = ss.getSheetByName('P2_EVENT_INDEX');
    const raw = ss.getSheetByName('RAW_DATA_SALES');

    if (!p1 || p1.getLastRow() < 2) throw new Error('P1_FACT belum tersedia atau belum berisi data.');
    if (!p2 || p2.getLastRow() < 2) throw new Error('P2_EVENT_INDEX belum tersedia atau belum berisi data.');
    if (!raw || raw.getLastRow() < 2) throw new Error('RAW_DATA_SALES belum tersedia atau belum berisi data.');

    analyticsAssertHeaders_(p1, P1_FACT_HEADERS_);
    analyticsAssertHeaders_(p2, P2_EVENT_HEADERS_);
    const rawCheck = analyticsAssertRawSheet_(raw);

    // Keep the existing source tables untouched. The initial build routine
    // only validates/registers their routing; it does not rewrite their data.
    const sourceValidation = analyticsStartInitialBuild();

    // Rebuild the two derived read models from the manually maintained P2 index.
    const indexes = analyticsBuildDashboardIndexes();

    const now = new Date();
    const stamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
    props.setProperty('SALES_LAST_UPDATED', stamp);
    props.setProperty('PAGE2_RAW_UPDATED', stamp);

    // Remove legacy Page 1 cache keys. Page 1's current cache is also protected
    // by a spreadsheet revision fingerprint; Page 2/3 use PAGE2_RAW_UPDATED.
    if (typeof cacheDashboardData_ === 'function') cacheDashboardData_();

    const result = {
      ok: true,
      status: 'COMPLETE',
      message: 'Refresh selesai. P1_FACT, P2_EVENT_INDEX, dan RAW_DATA_SALES tidak diubah.',
      spreadsheetId: ss.getId(),
      sources: {
        P1_FACT: { rows: p1.getLastRow() - 1, columns: p1.getLastColumn() },
        P2_EVENT_INDEX: { rows: p2.getLastRow() - 1, columns: p2.getLastColumn() },
        RAW_DATA_SALES: { rows: raw.getLastRow() - 1, columns: rawCheck.columns }
      },
      derived: {
        P2_DASHBOARD_FACT: { rows: indexes.p2SummaryRows },
        P3_OUTLET_MONTHLY: { rows: indexes.p3SummaryRows }
      },
      updatedAt: stamp,
      elapsedMs: Date.now() - started
    };
    Logger.log('MANTUL DASHBOARD REFRESH RESULT\\n' + JSON.stringify(result, null, 2));
    console.log('MANTUL DASHBOARD REFRESH RESULT', result);
    return result;
  } catch (err) {
    // Do not leave a misleading freshness timestamp after a failed refresh.
    if (oldSalesUpdated) props.setProperty('SALES_LAST_UPDATED', oldSalesUpdated);
    else props.deleteProperty('SALES_LAST_UPDATED');
    if (oldPage2Updated) props.setProperty('PAGE2_RAW_UPDATED', oldPage2Updated);
    else props.deleteProperty('PAGE2_RAW_UPDATED');
    Logger.log('MANTUL DASHBOARD REFRESH FAILED: ' + (err && err.stack || err));
    console.error('MANTUL DASHBOARD REFRESH FAILED', err);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/** Read-only validation; useful before a scheduled or manual refresh. */
function validateMantulDashboardSources() {
  const ss = getDatabase_();
  const p1 = ss.getSheetByName('P1_FACT');
  const p2 = ss.getSheetByName('P2_EVENT_INDEX');
  const raw = ss.getSheetByName('RAW_DATA_SALES');
  if (!p1 || p1.getLastRow() < 2) throw new Error('P1_FACT belum tersedia atau belum berisi data.');
  if (!p2 || p2.getLastRow() < 2) throw new Error('P2_EVENT_INDEX belum tersedia atau belum berisi data.');
  if (!raw || raw.getLastRow() < 2) throw new Error('RAW_DATA_SALES belum tersedia atau belum berisi data.');
  analyticsAssertHeaders_(p1, P1_FACT_HEADERS_);
  analyticsAssertHeaders_(p2, P2_EVENT_HEADERS_);
  const rawCheck = analyticsAssertRawSheet_(raw);
  const result = {
    ok: true,
    readOnly: true,
    sources: {
      P1_FACT: p1.getLastRow() - 1,
      P2_EVENT_INDEX: p2.getLastRow() - 1,
      RAW_DATA_SALES: raw.getLastRow() - 1
    },
    rawColumns: rawCheck.columns,
    spreadsheetId: ss.getId()
  };
  Logger.log('MANTUL SOURCE VALIDATION\\n' + JSON.stringify(result, null, 2));
  return result;
}
