/**
 * MANTUL — READ-ONLY DERIVED INDEX AUDIT
 *
 * Safe to run during the audit phase. This function only reads sheet names,
 * headers, row counts, and existing script-property metadata. It does not
 * call BUILD functions and does not modify spreadsheet contents or properties.
 */
function analyticsAuditDerivedIndexesReadOnly() {
  const ss = getDatabase_();
  const names = [
    'RAW_DATA_SALES',
    'P2_EVENT_INDEX',
    'P2_DASHBOARD_FACT',
    'P2_SUBDIST_FACT',
    'P3_OUTLET_MONTHLY',
    'SKU',
    'MAPPING TEAM'
  ];
  const sheets = {};
  names.forEach(function(name) {
    const sh = ss.getSheetByName(name);
    sheets[name] = sh ? {
      exists: true,
      rows: Math.max(0, sh.getLastRow() - 1),
      columns: sh.getLastColumn(),
      headers: sh.getLastColumn()
        ? sh.getRange(1, 1, 1, Math.min(sh.getLastColumn(), 30)).getDisplayValues()[0]
        : []
    } : { exists: false, rows: 0, columns: 0, headers: [] };
  });

  const props = PropertiesService.getScriptProperties();
  let stats = null;
  let ranges = null;
  try { stats = JSON.parse(props.getProperty('MANTUL_DERIVED_INDEX_STATS') || 'null'); } catch (e) {}
  try { ranges = JSON.parse(props.getProperty('MANTUL_P2_SUBDIST_RANGES') || 'null'); } catch (e) {}

  const sourceRows = sheets.P2_EVENT_INDEX.rows;
  const checks = {
    sourceIndexExists: sheets.P2_EVENT_INDEX.exists,
    dashboardFactExists: sheets.P2_DASHBOARD_FACT.exists,
    subdistFactExists: sheets.P2_SUBDIST_FACT.exists,
    outletMonthlyExists: sheets.P3_OUTLET_MONTHLY.exists,
    dashboardAndSubdistRowCountsEqual:
      sheets.P2_DASHBOARD_FACT.exists &&
      sheets.P2_SUBDIST_FACT.exists &&
      sheets.P2_DASHBOARD_FACT.rows === sheets.P2_SUBDIST_FACT.rows,
    sourceStatsMatch:
      !!stats && Number(stats.sourceLastRow) === sourceRows + 1,
    subdistRangeIndexMatches:
      !!ranges && Number(ranges.rowCount) === sheets.P2_SUBDIST_FACT.rows
  };

  const result = {
    ok: true,
    readOnly: true,
    buildExecuted: false,
    spreadsheetId: ss.getId(),
    buildStatus: props.getProperty('MANTUL_DERIVED_INDEX_STATUS') || 'NOT_BUILT',
    checks: checks,
    sheets: sheets,
    storedStats: stats ? {
      sourceLastRow: stats.sourceLastRow,
      p2SummaryRows: stats.p2SummaryRows,
      p2SubdistRows: stats.p2SubdistRows,
      p3SummaryRows: stats.p3SummaryRows,
      builtAt: stats.builtAt
    } : null,
    subdistRangeIndex: ranges ? {
      rowCount: ranges.rowCount,
      subdistCount: Object.keys(ranges.ranges || {}).length,
      builtAt: ranges.builtAt
    } : null
  };
  Logger.log('MANTUL READ-ONLY INDEX AUDIT\\n' + JSON.stringify(result, null, 2));
  console.log('MANTUL READ-ONLY INDEX AUDIT', result);
  return result;
}
