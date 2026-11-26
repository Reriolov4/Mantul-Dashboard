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
    const readIndex = refreshMantulReadIndexIncremental_();
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
    props.setProperty('SALES_LAST_UPDATED', stamp);
    props.setProperty('PAGE2_RAW_UPDATED', stamp);
    props.deleteProperty('MANTUL_V2_DATA_STALE');
    cacheDashboardData_();
    const result = {ok:true,status:'COMPLETE',buildExecuted:false,
      message:'Refresh selesai. Page 1 tetap memakai P1_FACT; Page 2–4 memakai P2_EVENT_INDEX hasil salinan terverifikasi dari RAW_DATA_SALES A:Y.',
      validation:validation,readIndex:readIndex,updatedAt:stamp,elapsedMs:Date.now()-started};
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


// ============================================================
// MANTUL READ INDEX — snapshot terverifikasi RAW_DATA_SALES A:Y.
// Page 1/P1_FACT tidak disentuh. Index dibangun secara staging,
// diverifikasi row count + total VAL, lalu dipublikasikan atomik.
// ============================================================
const MANTUL_READ_INDEX_SHEET_ = 'P2_EVENT_INDEX';
const MANTUL_READ_INDEX_STAGE_ = 'P2_EVENT_INDEX_BUILD';

function buildMantulReadIndex() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  const ss = getDatabase_();
  let stage = null;
  try {
    const raw = ss.getSheetByName(RAW_DATA_SHEET_NAME);
    if (!raw || raw.getLastRow() < 2) throw new Error('RAW_DATA_SALES kosong atau tidak ditemukan.');
    validateRawDatabaseHeaders_(raw.getRange(1, 1, 1, RAW_SALES_COLUMN_COUNT).getDisplayValues()[0]);
    if (raw.getLastColumn() !== RAW_SALES_COLUMN_COUNT) {
      throw new Error('RAW_DATA_SALES harus tepat 25 kolom A:Y. Terbaca ' + raw.getLastColumn() + ' kolom.');
    }

    const rowCount = raw.getLastRow() - 1;
    const headers = raw.getRange(1, 1, 1, RAW_SALES_COLUMN_COUNT).getValues();
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const stageName = MANTUL_READ_INDEX_STAGE_;
    const oldStage = ss.getSheetByName(stageName);
    if (oldStage) ss.deleteSheet(oldStage);
    stage = ss.insertSheet(stageName);
    stage.getRange(1, 1, 1, RAW_SALES_COLUMN_COUNT).setValues(headers);

    let sourceValue = 0, indexValue = 0, copied = 0;
    const batchSize = 5000;
    for (let offset = 0; offset < rowCount; offset += batchSize) {
      const count = Math.min(batchSize, rowCount - offset);
      const values = raw.getRange(offset + 2, 1, count, RAW_SALES_COLUMN_COUNT).getValues();
      for (let i = 0; i < values.length; i++) sourceValue += Number(values[i][RAW_SALES_COL.VALUE]) || 0;
      stage.getRange(offset + 2, 1, count, RAW_SALES_COLUMN_COUNT).setValues(values);
      copied += count;
    }
    SpreadsheetApp.flush();

    if (copied !== rowCount || stage.getLastRow() !== rowCount + 1 || stage.getLastColumn() !== RAW_SALES_COLUMN_COUNT) {
      throw new Error('Validasi indeks gagal: jumlah baris/kolom tidak cocok.');
    }
    for (let offset = 0; offset < rowCount; offset += batchSize) {
      const count = Math.min(batchSize, rowCount - offset);
      const values = stage.getRange(offset + 2, 1, count, RAW_SALES_COLUMN_COUNT).getValues();
      for (let i = 0; i < values.length; i++) indexValue += Number(values[i][RAW_SALES_COL.VALUE]) || 0;
    }
    if (Math.abs(sourceValue - indexValue) > Math.max(0.01, Math.abs(sourceValue) * 1e-10)) {
      throw new Error('Validasi total VAL gagal. RAW=' + sourceValue + '; INDEX=' + indexValue);
    }

    const current = ss.getSheetByName(MANTUL_READ_INDEX_SHEET_);
    const backupName = 'P2_EVENT_INDEX_OLD_' + stamp;
    if (current) current.setName(backupName);
    try {
      stage.setName(MANTUL_READ_INDEX_SHEET_);
      stage.setFrozenRows(1);
      PropertiesService.getScriptProperties().setProperties({
        MANTUL_READ_INDEX_STATUS: 'COMPLETE',
        MANTUL_READ_INDEX_SOURCE_ID: ss.getId(),
        MANTUL_READ_INDEX_SOURCE_ROWS: String(rowCount),
        MANTUL_READ_INDEX_SOURCE_LASTROW: String(raw.getLastRow()),
        MANTUL_READ_INDEX_VAL: String(indexValue),
        MANTUL_READ_INDEX_BUILT_AT: new Date().toISOString()
      });
      if (current) ss.deleteSheet(ss.getSheetByName(backupName));
    } catch (publishError) {
      const published = ss.getSheetByName(MANTUL_READ_INDEX_SHEET_);
      if (published && published.getSheetId() === stage.getSheetId()) ss.deleteSheet(published);
      const backup = ss.getSheetByName(backupName);
      if (backup && !ss.getSheetByName(MANTUL_READ_INDEX_SHEET_)) backup.setName(MANTUL_READ_INDEX_SHEET_);
      throw publishError;
    }
    const result = {ok:true,status:'COMPLETE',sheet:MANTUL_READ_INDEX_SHEET_,rows:rowCount,columns:RAW_SALES_COLUMN_COUNT,sourceValue:sourceValue,indexValue:indexValue,validated:true};
    Logger.log('MANTUL READ INDEX BUILD\\n' + JSON.stringify(result, null, 2));
    return result;
  } catch (err) {
    if (stage) {
      const stillThere = ss.getSheetByName(MANTUL_READ_INDEX_STAGE_);
      if (stillThere && stillThere.getSheetId() === stage.getSheetId()) ss.deleteSheet(stillThere);
    }
    PropertiesService.getScriptProperties().setProperty('MANTUL_READ_INDEX_STATUS', 'FAILED');
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Refresh read index append-only. Asumsi: histori lama tidak diedit,
 * transaksi baru selalu ditambahkan di bawah. Jika tidak valid, full rebuild.
 */
function refreshMantulReadIndexIncremental_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let lockHeld = true;
  const ss = getDatabase_();
  const props = PropertiesService.getScriptProperties();
  try {
    const raw = ss.getSheetByName(RAW_DATA_SHEET_NAME);
    if (!raw || raw.getLastRow() < 2) throw new Error('RAW_DATA_SALES kosong atau tidak ditemukan.');
    if (raw.getLastColumn() !== RAW_SALES_COLUMN_COUNT) throw new Error('RAW_DATA_SALES harus tepat 25 kolom A:Y. Terbaca ' + raw.getLastColumn() + ' kolom.');
    validateRawDatabaseHeaders_(raw.getRange(1, 1, 1, RAW_SALES_COLUMN_COUNT).getDisplayValues()[0]);

    const index = ss.getSheetByName(MANTUL_READ_INDEX_SHEET_);
    const previousLastRow = Number(props.getProperty('MANTUL_READ_INDEX_SOURCE_LASTROW') || 0);
    const previousRows = Number(props.getProperty('MANTUL_READ_INDEX_SOURCE_ROWS') || 0);
    const previousValue = Number(props.getProperty('MANTUL_READ_INDEX_VAL') || 0);
    const ready = props.getProperty('MANTUL_READ_INDEX_STATUS') === 'COMPLETE' &&
      props.getProperty('MANTUL_READ_INDEX_SOURCE_ID') === ss.getId() &&
      index && previousLastRow >= 2 && index.getLastRow() === previousLastRow &&
      previousRows === previousLastRow - 1;

    if (!ready || raw.getLastRow() < previousLastRow) {
      lock.releaseLock(); lockHeld = false;
      return buildMantulReadIndex();
    }
    if (raw.getLastRow() === previousLastRow) {
      return {ok:true,status:'UNCHANGED',sheet:MANTUL_READ_INDEX_SHEET_,rows:previousRows,
        columns:RAW_SALES_COLUMN_COUNT,indexValue:previousValue,appendedRows:0};
    }

    const appendCount = raw.getLastRow() - previousLastRow;
    const batchSize = 4000;
    let appendedValue = 0, verifiedValue = 0, written = 0;
    try {
      for (let offset = 0; offset < appendCount; offset += batchSize) {
        const count = Math.min(batchSize, appendCount - offset);
        const values = raw.getRange(previousLastRow + offset + 1, 1, count, RAW_SALES_COLUMN_COUNT).getValues();
        values.forEach(row => appendedValue += Number(row[RAW_SALES_COL.VALUE]) || 0);
        index.getRange(previousLastRow + offset + 1, 1, count, RAW_SALES_COLUMN_COUNT).setValues(values);
        written += count;
      }
      SpreadsheetApp.flush();
      if (written !== appendCount || index.getLastRow() !== raw.getLastRow() || index.getLastColumn() !== RAW_SALES_COLUMN_COUNT) {
        throw new Error('Validasi append index gagal: jumlah baris/kolom tidak cocok.');
      }
      for (let offset = 0; offset < appendCount; offset += batchSize) {
        const count = Math.min(batchSize, appendCount - offset);
        index.getRange(previousLastRow + offset + 1, 1, count, RAW_SALES_COLUMN_COUNT).getValues()
          .forEach(row => verifiedValue += Number(row[RAW_SALES_COL.VALUE]) || 0);
      }
      if (Math.abs(appendedValue - verifiedValue) > Math.max(0.01, Math.abs(appendedValue) * 1e-10)) {
        throw new Error('Validasi total VAL baris baru gagal. RAW=' + appendedValue + '; INDEX=' + verifiedValue);
      }
      const totalValue = previousValue + verifiedValue;
      props.setProperties({
        MANTUL_READ_INDEX_STATUS:'COMPLETE', MANTUL_READ_INDEX_SOURCE_ID:ss.getId(),
        MANTUL_READ_INDEX_SOURCE_ROWS:String(raw.getLastRow()-1),
        MANTUL_READ_INDEX_SOURCE_LASTROW:String(raw.getLastRow()),
        MANTUL_READ_INDEX_VAL:String(totalValue), MANTUL_READ_INDEX_BUILT_AT:new Date().toISOString()
      });
      return {ok:true,status:'APPENDED',sheet:MANTUL_READ_INDEX_SHEET_,rows:raw.getLastRow()-1,
        columns:RAW_SALES_COLUMN_COUNT,appendedRows:appendCount,appendedValue:appendedValue,
        indexValue:totalValue,validated:true};
    } catch (err) {
      const last=index.getLastRow();
      if(last>previousLastRow) index.getRange(previousLastRow+1,1,last-previousLastRow,RAW_SALES_COLUMN_COUNT).clearContent();
      throw err;
    }
  } finally {
    if (lockHeld) lock.releaseLock();
  }
}

function mantulGetReadyReadIndex_() {
  const ss = getDatabase_();
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('MANTUL_READ_INDEX_STATUS') !== 'COMPLETE' ||
      props.getProperty('MANTUL_READ_INDEX_SOURCE_ID') !== ss.getId()) return null;
  const raw = ss.getSheetByName(RAW_DATA_SHEET_NAME);
  const index = ss.getSheetByName(MANTUL_READ_INDEX_SHEET_);
  if (!raw || !index || raw.getLastRow() !== Number(props.getProperty('MANTUL_READ_INDEX_SOURCE_LASTROW')) ||
      index.getLastRow() !== raw.getLastRow() || index.getLastColumn() !== RAW_SALES_COLUMN_COUNT) return null;
  return index;
}
