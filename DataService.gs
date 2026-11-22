/**
 * Data service for Sales Performance Dashboard.
 * Import maintains two independent schemas and publishes both together.
 * A failed import leaves the currently published DATA_SALES and RAW_DATA_SALES intact.
 */

function getDatabase_() {
  const configuredId = String(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '').trim();
  const spreadsheetId = configuredId || String(DEFAULT_SPREADSHEET_ID || '').trim();
  if (!spreadsheetId) {
    throw new Error('ID Spreadsheet belum dikonfigurasi. Isi Script Properties SPREADSHEET_ID atau DEFAULT_SPREADSHEET_ID di Config.gs.');
  }
  try {
    return SpreadsheetApp.openById(spreadsheetId);
  } catch (err) {
    throw new Error('Gagal membuka Spreadsheet dengan ID yang dikonfigurasi (' +
      (configuredId ? 'Script Properties' : 'fallback Config.gs') +
      '). Pastikan akun pemilik Apps Script memiliki akses ke Spreadsheet. Detail: ' + err.message);
  }
}

/**
 * Jalankan manual dari editor Apps Script untuk memverifikasi koneksi database.
 * Mengembalikan nama file dan jumlah baris sheet utama tanpa mengubah data.
 */
function testMantulDatabaseConnection() {
  const ss = getDatabase_();
  const raw = ss.getSheetByName(RAW_DATA_SHEET_NAME);
  const sales = ss.getSheetByName(DATA_SHEET_NAME);
  const result = {
    ok: true,
    spreadsheetName: ss.getName(),
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    activeSource: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') ? 'Script Properties' : 'Config.gs fallback',
    rawDataSheetFound: !!raw,
    rawDataRows: raw ? Math.max(0, raw.getLastRow() - 1) : 0,
    rawDataLastColumn: raw ? raw.getLastColumn() : 0,
    dataSalesSheetFound: !!sales,
    dataSalesRows: sales ? Math.max(0, sales.getLastRow() - 1) : 0
  };
  Logger.log(JSON.stringify(result));
  return result;
}

function getOrCreateSheet_(name) {
  const ss = getDatabase_();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getDashboardData() {
  const ss = getDatabase_();
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return { rows: [], updatedAt: getLastUpdated_(), rowCount: 0 };
  }

  // Source fingerprint prevents stale Page 1 cache after append-only updates.
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const signature = [ss.getId(), sheet.getSheetId(), lastRow, lastColumn, getLastUpdated_()].join('|');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, signature, Utilities.Charset.UTF_8);
  const cacheKey = 'dashboard_rows_v5_' + Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, '');
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); }
    catch (err) { cache.remove(cacheKey); }
  }

  const values = sheet.getRange(1, 1, lastRow, Math.max(sheet.getLastColumn(), SALES_HEADERS.length)).getValues();
  const headers = values[0] || [];
  validateDatabaseHeaders_(headers);

  // Resolve DATA_SALES by header name. This supports the workbook's
  // human-readable 14-column schema and the compact internal schema.
  const headerKey = value => String(value == null ? '' : value).trim().toLowerCase().replace(/[_\\s]+/g, ' ');
  const headerIndex = {};
  headers.forEach((value, index) => { headerIndex[headerKey(value)] = index; });
  const col = (aliases, required) => {
    for (const alias of aliases) {
      const index = headerIndex[headerKey(alias)];
      if (index !== undefined) return index;
    }
    if (required) throw new Error('Kolom DATA_SALES tidak ditemukan: ' + aliases[0]);
    return -1;
  };
  const ix = {
    y:col(['YEAR','y'],true), m:col(['MONTH','m'],true),
    d:col(['SUBDIST','d'],true), s:col(['Subbrand','SUBBRAND','s'],true),
    b:col(['Brand','BRAND','b'],true), c:col(['CATEGORY','c'],true),
    u:col(['SBU','u'],true), p:col(['PIC','p'],true),
    fm:col(['PF Bulanan','fm'],true), fq:col(['PF Triwulan','fq'],true),
    v:col(['VAL','v'],true), t:col(['JENIS SALES','t'],true),
    tg:col(['TARGET','tg'],false), be:col(['BE','be'],false)
  };

  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r.some(v => v !== '' && v !== null)) continue;
    const year = Number(r[ix.y]);
    const amount = Number(r[ix.v]);
    if (!Number.isFinite(year) || !Number.isFinite(amount)) continue;

    rows.push({
      y: year,
      m: String(r[ix.m] || '').toUpperCase(),
      d: String(r[ix.d] || ''),
      s: String(r[ix.s] || ''),
      b: String(r[ix.b] || ''),
      c: String(r[ix.c] || ''),
      u: String(r[ix.u] || ''),
      p: String(r[ix.p] || ''),
      fm: String(r[ix.fm] || ''),
      fq: String(r[ix.fq] || ''),
      v: amount,
      t: String(r[ix.t] || '').toUpperCase(),
      tg: ix.tg < 0 || r[ix.tg] === '' || r[ix.tg] === null ? null : Number(r[ix.tg]),
      be: ix.be < 0 || r[ix.be] === '' || r[ix.be] === null ? null : Number(r[ix.be])
    });
  }

  const result = { rows: rows, updatedAt: getLastUpdated_(), rowCount: rows.length };
  const json = JSON.stringify(result);
  if (json.length <= 90000) {
    try { cache.put(cacheKey, json, 300); } catch (err) { /* cache is optional */ }
  }
  return result;
}
function getLastUpdated_() {
  return PropertiesService.getScriptProperties().getProperty('SALES_LAST_UPDATED') || '';
}

function beginSalesImport(meta) {
  assertUploadPermission_();
  if (!meta || !Number.isInteger(Number(meta.total)) || Number(meta.total) < 1) {
    throw new Error('Jumlah baris DATA_SALES tidak valid.');
  }
  if (!Number.isInteger(Number(meta.rawTotal)) || Number(meta.rawTotal) < 1) {
    throw new Error('Sheet RAW_DATA_SALES wajib tersedia dan berisi data.');
  }
  if (Number(meta.total) > 500000 || Number(meta.rawTotal) > 500000) {
    throw new Error('Jumlah baris melebihi batas aman 500.000 baris per sheet.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const existingToken = props.getProperty('SALES_IMPORT_TOKEN');
    const startedAt = Number(props.getProperty('SALES_IMPORT_STARTED') || 0);
    const sessionActive = existingToken && startedAt && (Date.now() - startedAt < 2 * 60 * 60 * 1000);
    if (sessionActive) {
      throw new Error('Import lain sedang berlangsung. Selesaikan atau tunggu sesi sebelumnya kedaluwarsa (2 jam).');
    }

    const ss = getDatabase_();
    // The workbook is close to Google Sheets' 10-million-cell ceiling.
    // Keep only the current live tables plus the new staging pair during upload.
    cleanupOldBackups_(ss, 0);
    const stage = getOrCreateSheet_(STAGING_SHEET_NAME);
    const rawStage = getOrCreateSheet_(RAW_STAGING_SHEET_NAME);
    stage.clearContents();
    rawStage.clearContents();
    stage.getRange(1, 1, 1, SALES_HEADERS.length).setValues([SALES_HEADERS]);
    rawStage.getRange(1, 1, 1, RAW_SALES_HEADERS.length).setValues([RAW_SALES_HEADERS]);

    const token = Utilities.getUuid();
    props.setProperties({
      SALES_IMPORT_TOKEN: token,
      SALES_IMPORT_EXPECTED: String(Number(meta.total)),
      SALES_IMPORT_WRITTEN: '0',
      RAW_IMPORT_EXPECTED: String(Number(meta.rawTotal)),
      RAW_IMPORT_WRITTEN: '0',
      SALES_IMPORT_STARTED: String(Date.now()),
      SALES_IMPORT_FILENAME: safeText_(meta.filename || 'Workbook Excel', 200)
    });
    return { ok: true, token: token, expected: Number(meta.total), rawExpected: Number(meta.rawTotal) };
  } finally {
    lock.releaseLock();
  }
}

function appendSalesImportChunk(payload) {
  assertUploadPermission_();
  if (!payload) throw new Error('Payload import tidak valid.');
  const chunk = Array.isArray(payload.chunk) ? payload.chunk : [];
  const rawChunk = Array.isArray(payload.rawChunk) ? payload.rawChunk : [];
  if (!chunk.length && !rawChunk.length) return { written: 0, rawWritten: 0 };
  if (chunk.length > 1000 || rawChunk.length > 1000) throw new Error('Maksimal 1.000 baris per sheet per batch.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    assertImportSession_(props, payload.token);
    const expected = Number(props.getProperty('SALES_IMPORT_EXPECTED') || 0);
    const writtenBefore = Number(props.getProperty('SALES_IMPORT_WRITTEN') || 0);
    const rawExpected = Number(props.getProperty('RAW_IMPORT_EXPECTED') || 0);
    const rawWrittenBefore = Number(props.getProperty('RAW_IMPORT_WRITTEN') || 0);
    if (writtenBefore + chunk.length > expected || rawWrittenBefore + rawChunk.length > rawExpected) {
      throw new Error('Jumlah baris yang dikirim melebihi jumlah yang didaftarkan.');
    }

    if (chunk.length) {
      const rows = chunk.map((o, index) => normalizeImportRow_(o, writtenBefore + index + 1));
      const stage = getOrCreateSheet_(STAGING_SHEET_NAME);
      ensureSheetGrid_(stage, writtenBefore + rows.length + 1, SALES_HEADERS.length);
      stage.getRange(writtenBefore + 2, 1, rows.length, SALES_HEADERS.length).setValues(rows);
      props.setProperty('SALES_IMPORT_WRITTEN', String(writtenBefore + rows.length));
    }
    if (rawChunk.length) {
      const rawRows = rawChunk.map((row, index) => normalizeRawImportRow_(row, rawWrittenBefore + index + 1));
      const rawStage = getOrCreateSheet_(RAW_STAGING_SHEET_NAME);
      ensureSheetGrid_(rawStage, rawWrittenBefore + rawRows.length + 1, RAW_SALES_HEADERS.length);
      rawStage.getRange(rawWrittenBefore + 2, 1, rawRows.length, RAW_SALES_HEADERS.length).setValues(rawRows);
      props.setProperty('RAW_IMPORT_WRITTEN', String(rawWrittenBefore + rawRows.length));
    }
    return {
      written: Number(props.getProperty('SALES_IMPORT_WRITTEN') || 0),
      expected: expected,
      rawWritten: Number(props.getProperty('RAW_IMPORT_WRITTEN') || 0),
      rawExpected: rawExpected
    };
  } finally {
    lock.releaseLock();
  }
}

function finishSalesImport(payload) {
  assertUploadPermission_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    assertImportSession_(props, payload && payload.token);
    const expected = Number(props.getProperty('SALES_IMPORT_EXPECTED') || 0);
    const written = Number(props.getProperty('SALES_IMPORT_WRITTEN') || 0);
    const rawExpected = Number(props.getProperty('RAW_IMPORT_EXPECTED') || 0);
    const rawWritten = Number(props.getProperty('RAW_IMPORT_WRITTEN') || 0);
    if (!expected || written !== expected || !rawExpected || rawWritten !== rawExpected) {
      throw new Error('Import belum lengkap. DATA_SALES: ' + written + '/' + expected + '; RAW_DATA_SALES: ' + rawWritten + '/' + rawExpected + ' baris.');
    }

    const ss = getDatabase_();
    const stage = getOrCreateSheet_(STAGING_SHEET_NAME);
    const rawStage = getOrCreateSheet_(RAW_STAGING_SHEET_NAME);
    if (stage.getLastRow() !== expected + 1 || rawStage.getLastRow() !== rawExpected + 1) {
      throw new Error('Jumlah baris staging tidak sesuai. Database lama tetap dipertahankan; silakan ulangi import.');
    }
    validateDatabaseHeaders_(stage.getRange(1, 1, 1, SALES_HEADERS.length).getValues()[0]);
    validateRawDatabaseHeaders_(rawStage.getRange(1, 1, 1, RAW_SALES_HEADERS.length).getValues()[0]);

    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '_' + Utilities.getUuid().slice(0, 6);
    const nextDataName = 'DATA_NEXT_' + stamp;
    const nextRawName = 'RAW_NEXT_' + stamp;
    const backupDataName = 'DATA_BACKUP_' + stamp;
    const backupRawName = 'RAW_BACKUP_' + stamp;
    let nextData = null, nextRaw = null;
    const currentData = ss.getSheetByName(DATA_SHEET_NAME);
    const currentRaw = ss.getSheetByName(RAW_DATA_SHEET_NAME);

    try {
      // Reuse the fully verified staging sheets as the replacement sheets.
      // This avoids duplicating the 130k+ x 34 raw table and exceeding the
      // spreadsheet's 10-million-cell limit.
      nextData = stage;
      nextRaw = rawStage;
      nextData.setFrozenRows(1);
      nextRaw.setFrozenRows(1);
      nextData.getRange(1, 1, 1, SALES_HEADERS.length).setFontWeight('bold').setBackground('#eaf1ff');
      nextRaw.getRange(1, 1, 1, RAW_SALES_HEADERS.length).setFontWeight('bold').setBackground('#eaf1ff');
      SpreadsheetApp.flush();

      if (nextData.getLastRow() !== expected + 1 || nextData.getLastColumn() < SALES_HEADERS.length ||
          nextRaw.getLastRow() !== rawExpected + 1 || nextRaw.getLastColumn() < RAW_SALES_HEADERS.length) {
        throw new Error('Verifikasi salah satu database pengganti gagal. Database lama tetap dipertahankan.');
      }

      if (currentData) currentData.setName(backupDataName);
      if (currentRaw) currentRaw.setName(backupRawName);
      nextData.setName(DATA_SHEET_NAME);
      nextRaw.setName(RAW_DATA_SHEET_NAME);
    } catch (commitError) {
      // Roll back both sheets as a pair if any part of the publish fails.
      const publishedData = ss.getSheetByName(DATA_SHEET_NAME);
      const publishedRaw = ss.getSheetByName(RAW_DATA_SHEET_NAME);
      if (publishedData && nextData && publishedData.getSheetId() === nextData.getSheetId()) ss.deleteSheet(publishedData);
      if (publishedRaw && nextRaw && publishedRaw.getSheetId() === nextRaw.getSheetId()) ss.deleteSheet(publishedRaw);
      const dataBackup = ss.getSheetByName(backupDataName);
      const rawBackup = ss.getSheetByName(backupRawName);
      if (dataBackup && !ss.getSheetByName(DATA_SHEET_NAME)) dataBackup.setName(DATA_SHEET_NAME);
      if (rawBackup && !ss.getSheetByName(RAW_DATA_SHEET_NAME)) rawBackup.setName(RAW_DATA_SHEET_NAME);
      [STAGING_SHEET_NAME, RAW_STAGING_SHEET_NAME, nextDataName, nextRawName].forEach(name => {
        const sh = ss.getSheetByName(name); if (sh) ss.deleteSheet(sh);
      });
      throw commitError;
    }

    // The two published tables now represent the same uploaded workbook.
    try { cleanupOldBackups_(ss, 3); } catch (cleanupError) { console.error('Pembersihan backup tertunda: ' + cleanupError.message); }
    const now = new Date();
    const formatted = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
    props.setProperty('SALES_LAST_UPDATED', formatted);
    props.setProperty('PAGE2_RAW_UPDATED', formatted);
    try {
      const log = getOrCreateSheet_(IMPORT_LOG_SHEET_NAME);
      if (log.getLastRow() === 0) log.appendRow(['WAKTU', 'FILE', 'BARIS DATA_SALES', 'BARIS RAW_DATA_SALES', 'UPLOADER', 'STATUS']);
      log.appendRow([now, props.getProperty('SALES_IMPORT_FILENAME') || '', written, rawWritten,
        Session.getActiveUser().getEmail() || 'email tidak tersedia', 'SUCCESS - BOTH DATABASES']);
    } catch (logError) {
      console.error('Import berhasil tetapi pencatatan log gagal: ' + logError.message);
    }

    clearImportSession_(props);
    cacheDashboardData_();
    return getDashboardData();
  } finally {
    lock.releaseLock();
  }
}

function normalizeRawImportRow_(row, rowNumber) {
  if (!Array.isArray(row) || row.length < RAW_SALES_HEADERS.length) {
    throw new Error('RAW_DATA_SALES baris ' + rowNumber + ' tidak memiliki 27 kolom A:AA.');
  }
  const out = row.slice(0, RAW_SALES_HEADERS.length).map(value => {
    if (value === undefined || value === null) return '';
    if (typeof value === 'string') return safeText_(value, 500);
    return value;
  });
  const year = Number(out[0]);
  const month = String(out[1] == null ? '' : out[1]).trim();
  const amount = Number(out[12]);
  const salesType = String(out[18] == null ? '' : out[18]).trim().toUpperCase();
  if (!Number.isInteger(year) || year < 2000 || year > 2100 || !month || !Number.isFinite(amount) ||
      (salesType !== 'SO' && salesType !== 'SOO')) {
    throw new Error('TAHUN, BULAN, VAL, atau JENIS SALES tidak valid pada RAW_DATA_SALES baris ' + rowNumber + '.');
  }
  return out;
}

function validateRawDatabaseHeaders_(headers) {
  const normalize = value => String(value == null ? '' : value).trim().toLowerCase().replace(/\s+/g, ' ');
  const actual = (headers || []).slice(0, RAW_SALES_HEADERS.length).map(normalize);
  const expected = RAW_SALES_HEADERS.map(normalize);
  if (actual.length !== expected.length || !actual.every((value, index) => value === expected[index])) {
    throw new Error('Header RAW_DATA_SALES tidak sesuai struktur A:AA pada workbook terbaru. Tidak ada database yang dipublikasikan.');
  }
}

function normalizeImportRow_(o, rowNumber) {
  if (!o || typeof o !== 'object') throw new Error('Format baris ' + rowNumber + ' tidak valid.');
  const year = Number(o.y);
  const amount = Number(o.v);
  const month = String(o.m || '').trim().toUpperCase();
  const salesType = String(o.t || '').trim().toUpperCase();
  const validMonths = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  if (!Number.isInteger(year) || year < 2000 || year > 2100) {
    throw new Error('YEAR tidak valid pada baris data ' + rowNumber + '.');
  }
  if (validMonths.indexOf(month) < 0) {
    throw new Error('MONTH tidak valid pada baris data ' + rowNumber + ': ' + month);
  }
  if (salesType !== 'SO' && salesType !== 'SOO') {
    throw new Error('JENIS SALES harus SO atau SOO pada baris data ' + rowNumber + '.');
  }
  if (!Number.isFinite(amount)) {
    throw new Error('VAL bukan angka pada baris data ' + rowNumber + '.');
  }

  return [
    year, month, safeText_(o.d, 200), safeText_(o.s, 300),
    safeText_(o.b, 200), safeText_(o.c, 200), safeText_(o.u, 100),
    safeText_(o.p, 200), safeText_(o.fm, 200), safeText_(o.fq, 200),
    amount, salesType, nullableNumber_(o.tg, 'TARGET', rowNumber),
    nullableNumber_(o.be, 'BE', rowNumber)
  ];
}

function nullableNumber_(value, label, rowNumber) {
  if (value === null || value === '' || value === undefined) return '';
  const number = Number(value);
  if (!Number.isFinite(number)) throw new Error(label + ' bukan angka pada baris data ' + rowNumber + '.');
  return number;
}

function safeText_(value, maxLength) {
  let text = value === null || value === undefined ? '' : String(value).trim();
  if (text.length > maxLength) text = text.slice(0, maxLength);
  // Prevent imported text from being interpreted as a spreadsheet formula.
  if (/^[=+\-@]/.test(text)) text = "'" + text;
  return text;
}

function validateDatabaseHeaders_(headers) {
  // Support both the compact import schema and the existing human-readable
  // DATA_SALES schema already used in the user's Google Spreadsheet.
  const normalize = value => String(value == null ? '' : value)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');

  const actual = (headers || []).slice(0, SALES_HEADERS.length).map(normalize);
  const compact = SALES_HEADERS.map(normalize);
  const existingSheet = [
    'YEAR', 'MONTH', 'SUBDIST', 'Subbrand', 'Brand', 'CATEGORY',
    'SBU', 'PIC', 'PF Bulanan', 'PF Triwulan', 'VAL',
    'JENIS SALES', 'TARGET', 'BE'
  ].map(normalize);

  const matches = expected => actual.length === expected.length &&
    actual.every((value, index) => value === expected[index]);

  if (!matches(compact) && !matches(existingSheet)) {
    throw new Error(
      'Header DATA_SALES tidak dikenali. Format yang didukung: ' +
      'YEAR, MONTH, SUBDIST, Subbrand, Brand, CATEGORY, SBU, PIC, ' +
      'PF Bulanan, PF Triwulan, VAL, JENIS SALES, TARGET, BE.'
    );
  }
}

function assertImportSession_(props, token) {
  const savedToken = props.getProperty('SALES_IMPORT_TOKEN');
  const startedAt = Number(props.getProperty('SALES_IMPORT_STARTED') || 0);
  if (!savedToken || !token || savedToken !== token) {
    throw new Error('Sesi import tidak valid atau sudah diganti. Silakan mulai ulang upload.');
  }
  if (!startedAt || Date.now() - startedAt >= 2 * 60 * 60 * 1000) {
    clearImportSession_(props);
    throw new Error('Sesi import kedaluwarsa. Silakan mulai ulang upload.');
  }
}

function clearImportSession_(props) {
  [
    'SALES_IMPORT_TOKEN', 'SALES_IMPORT_EXPECTED', 'SALES_IMPORT_WRITTEN',
    'SALES_IMPORT_STARTED', 'SALES_IMPORT_FILENAME',
    'RAW_IMPORT_EXPECTED', 'RAW_IMPORT_WRITTEN'
  ].forEach(key => props.deleteProperty(key));
}

function ensureSheetGrid_(sheet, requiredRows, requiredColumns) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < requiredRows) sheet.insertRowsAfter(maxRows, requiredRows - maxRows);
  const maxColumns = sheet.getMaxColumns();
  if (maxColumns < requiredColumns) sheet.insertColumnsAfter(maxColumns, requiredColumns - maxColumns);
}

function cleanupOldBackups_(ss, keepCount) {
  const backups = ss.getSheets().filter(sh => /^(DATA|RAW)_BACKUP_\d{8}_\d{6}(?:_[a-f0-9]{6})?$/.test(sh.getName()));
  const stamps = Array.from(new Set(backups.map(sh => sh.getName().replace(/^(DATA|RAW)_BACKUP_/, '')))).sort().reverse();
  const keep = new Set(stamps.slice(0, keepCount));
  backups.forEach(sh => {
    const stamp = sh.getName().replace(/^(DATA|RAW)_BACKUP_/, '');
    if (!keep.has(stamp)) ss.deleteSheet(sh);
  });
}

function cacheDashboardData_() {
  const cache = CacheService.getScriptCache();
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = String(props.getProperty('SPREADSHEET_ID') || DEFAULT_SPREADSHEET_ID || 'unconfigured').trim();

  // Invalidate the exact current Page 1 cache key. The old v2/v3 keys are
  // retained for cleanup, but getDashboardData() now reads dashboard_rows_v4.
  cache.remove('dashboard_rows_v4_' + spreadsheetId);
  cache.remove('dashboard_rows_v2');
  cache.remove('dashboard_rows_v3');
}

function assertUploadPermission_() {
  const email = (Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const configured = PropertiesService.getScriptProperties().getProperty('UPLOAD_ALLOWED_EMAILS') || '';
  const allowed = configured.split(',').map(x => String(x).trim().toLowerCase()).filter(Boolean);
  if (!email || allowed.indexOf(email) === -1) {
    throw new Error('Akun Anda belum diizinkan mengunggah data. Atur UPLOAD_ALLOWED_EMAILS pada Apps Script > Project Settings > Script Properties (pisahkan beberapa email dengan koma).');
  }
}
