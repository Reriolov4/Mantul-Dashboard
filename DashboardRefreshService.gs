/**
 * MANTUL — SAFE DASHBOARD REFRESH
 * Page 1 reads P1_FACT; Pages 2–4 read RAW_DATA_SALES directly.
 * This validator never requires or rebuilds Database V2 tables.
 */
function mantulV2ValidateLiveDatabase_() {
  const ss = getDatabase_();
  const p1 = ss.getSheetByName('P1_FACT');
  const raw = ss.getSheetByName(RAW_DATA_SHEET_NAME);
  if (!p1 || p1.getLastRow() < 2) throw new Error('P1_FACT belum tersedia atau belum berisi data.');
  if (!raw || raw.getLastRow() < 2) throw new Error('RAW_DATA_SALES belum tersedia atau belum berisi data.');
  analyticsAssertHeaders_(p1, P1_FACT_HEADERS_);
  mantulV2AssertRawHeaders_(raw);
  return {
    ok:true,
    version:'RAW_DIRECT',
    spreadsheetId:ss.getId(),
    sheets:{
      P1_FACT:{rows:p1.getLastRow()-1,columns:p1.getLastColumn()},
      RAW_DATA_SALES:{rows:raw.getLastRow()-1,columns:RAW_SALES_COLUMN_COUNT}
    }
  };
}

function refreshMantulDashboard() {
  const props = PropertiesService.getScriptProperties();
  const oldSales = props.getProperty('SALES_LAST_UPDATED') || '';
  const oldP2 = props.getProperty('PAGE2_RAW_UPDATED') || '';
  const started = Date.now();
  try {
    const validation = mantulV2ValidateLiveDatabase_();
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
    props.setProperty('SALES_LAST_UPDATED', stamp);
    props.setProperty('PAGE2_RAW_UPDATED', stamp);
    props.deleteProperty('MANTUL_V2_DATA_STALE');
    if (typeof cacheDashboardData_ === 'function') cacheDashboardData_();
    const result = {
      ok:true,
      status:'COMPLETE',
      buildExecuted:false,
      message:'Refresh selesai. Page 1 divalidasi dari P1_FACT dan Page 2–4 menggunakan RAW_DATA_SALES langsung. Tidak ada tabel sumber yang dibangun ulang atau ditimpa.',
      validation:validation,
      updatedAt:stamp,
      elapsedMs:Date.now()-started
    };
    Logger.log('MANTUL SAFE REFRESH\\n' + JSON.stringify(result, null, 2));
    return result;
  } catch (err) {
    if (oldSales) props.setProperty('SALES_LAST_UPDATED', oldSales);
    else props.deleteProperty('SALES_LAST_UPDATED');
    if (oldP2) props.setProperty('PAGE2_RAW_UPDATED', oldP2);
    else props.deleteProperty('PAGE2_RAW_UPDATED');
    Logger.log('MANTUL SAFE REFRESH FAILED: ' + (err && err.stack || err));
    throw err;
  }
}

function validateMantulDashboardSources() {
  return Object.assign({readOnly:true}, mantulV2ValidateLiveDatabase_());
}