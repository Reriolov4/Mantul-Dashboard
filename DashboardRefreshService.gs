/**
 * MANTUL — SAFE SOURCE VALIDATION
 *
 * Deliberately read-only until the derived-index audit and Page 3 / Page 4
 * services are reviewed. This function does NOT call analyticsBuildDashboardIndexes()
 * and does NOT rewrite, clear, rename, or append to any spreadsheet sheet.
 */
function refreshMantulDashboard() {
  const started = Date.now();
  const validation = validateMantulDashboardSources();
  const indexStatus = typeof analyticsGetDashboardIndexStatus === 'function'
    ? analyticsGetDashboardIndexStatus()
    : { ready: false, status: 'STATUS_SERVICE_UNAVAILABLE' };

  const result = {
    ok: true,
    status: 'VALIDATED_ONLY',
    buildExecuted: false,
    message: 'Sumber berhasil divalidasi. BUILD indeks sengaja dinonaktifkan sampai audit selesai.',
    spreadsheetId: validation.spreadsheetId,
    sources: validation.sources,
    rawColumns: validation.rawColumns,
    derivedIndex: indexStatus,
    elapsedMs: Date.now() - started
  };
  Logger.log('MANTUL SAFE VALIDATION RESULT\\n' + JSON.stringify(result, null, 2));
  console.log('MANTUL SAFE VALIDATION RESULT', result);
  return result;
}

/** Read-only validation of maintained source sheets and their required headers. */
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
