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
    cacheDashboardData_();
    const result = {ok:true,status:'COMPLETE',buildExecuted:false,
      message:'Refresh selesai. Page 1 divalidasi dari P1_FACT dan Page 2–4 membaca RAW_DATA_SALES langsung. Tidak ada tabel sumber yang dibangun ulang atau ditimpa.',
      validation:validation,updatedAt:stamp,elapsedMs:Date.now()-started};
    Logger.log('MANTUL SAFE REFRESH\\n' + JSON.stringify(result, null, 2));
    return result;
  } catch (err) {
    if (oldSales) props.setProperty('SALES_LAST_UPDATED', oldSales); else props.deleteProperty('SALES_LAST_UPDATED');
    if (oldP2) props.setProperty('PAGE2_RAW_UPDATED', oldP2); else props.deleteProperty('PAGE2_RAW_UPDATED');
    Logger.log('MANTUL SAFE REFRESH FAILED: ' + (err && err.stack || err));
    throw err;
  }
}

function validateMantulDashboardSources() {
  return Object.assign({readOnly:true}, mantulV2ValidateLiveDatabase_());
}
