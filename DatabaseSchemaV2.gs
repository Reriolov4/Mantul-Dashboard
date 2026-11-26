/**
 * MANTUL DATABASE V2 — CANONICAL SCHEMAS
 *
 * Source of truth: RAW_DATA_SALES (27 columns; never modified).
 * This file defines the V2 contracts. Builder and services migrate to these
 * contracts in separate, validated steps; do not run legacy rebuild functions.
 */
const MANTUL_V2_SCHEMA_ = Object.freeze({
  RAW: Object.freeze([
    'TAHUN','BULAN','TANGGAL','KODE ITEM','NAMA ITEM','SUB BRAND','BRAND',
    'KARTON','SUBDIST','KODE CUSTOMER','CUSTOMER NAME','KODE SALUR','VAL',
    'QTY','CLASS PRODUCT','UMBRAND','DOT','CATEGORY','JENIS SALES','PIC',
    'SUB_DIST','TIM','SALESMAN','SBU','PF Bulanan','PF Triwulan','PRODUCT_DESC'
  ]),
  P2_EVENT_INDEX: Object.freeze([
    'YEAR','MONTH','DD','CUSTOMER_CODE','CUSTOMER_NAME','SUBDIST','PIC',
    'BRAND','SUBBRAND','CATEGORY','SBU','SALES_TYPE','SALESMAN','TEAM',
    'VAL','CARTON','SKU','SUB_DIST','SOURCE_ROW'
  ]),
  P2_DASHBOARD_FACT: Object.freeze([
    'YEAR','MONTH','DD','CUSTOMER_CODE','CUSTOMER_NAME','SUBDIST','PIC',
    'BRAND','SUBBRAND','CATEGORY','SBU','SALES_TYPE','SALESMAN','TEAM',
    'VAL','CARTON','SOURCE_ROWS'
  ]),
  P3_OUTLET_MONTHLY: Object.freeze([
    'YEAR','MONTH','DD','CUSTOMER_CODE','CUSTOMER_NAME','SUBDIST','SUB_DIST',
    'SALES_TYPE','PIC','BRAND','SUBBRAND','CATEGORY','SBU','ITEM_CODE','SKU',
    'VAL','CARTON','SOURCE_ROWS'
  ])
});

/** Zero-based mapping from RAW_DATA_SALES A:AA to canonical P2_EVENT_INDEX. */
function mantulV2MapRawToP2Event_(r, sourceRow) {
  const c = RAW_SALES_COL;
  return [
    r[c.YEAR], r[c.MONTH], r[c.DAY], r[c.CUSTOMER_CODE], r[c.CUSTOMER_NAME],
    r[c.SUBDIST], r[c.PIC], r[c.BRAND], r[c.SUBBRAND], r[c.CATEGORY],
    r[c.SBU], r[c.SALES_TYPE], r[c.SALESMAN], r[c.TEAM], r[c.VALUE],
    r[c.CARTON], r[c.PRODUCT_DESC], r[c.SUB_DIST], sourceRow
  ];
}

/** Canonical P3 daily outlet × brand × SKU grain, retaining SOO and SO. */
function mantulV2MapRawToP3Outlet_(r) {
  const c = RAW_SALES_COL;
  return [
    r[c.YEAR], r[c.MONTH], r[c.DAY], r[c.CUSTOMER_CODE], r[c.CUSTOMER_NAME],
    r[c.SUBDIST], r[c.SUB_DIST], r[c.SALES_TYPE], r[c.PIC], r[c.BRAND],
    r[c.SUBBRAND], r[c.CATEGORY], r[c.SBU], r[c.ITEM_CODE],
    r[c.PRODUCT_DESC], r[c.VALUE], r[c.CARTON], 1
  ];
}

/** Fail fast if RAW_DATA_SALES headers drift from the agreed 27-column contract. */
function mantulV2AssertRawHeaders_(sheet) {
  if (!sheet || sheet.getLastColumn() < MANTUL_V2_SCHEMA_.RAW.length) {
    throw new Error('RAW_DATA_SALES tidak memiliki 27 kolom sesuai schema V2.');
  }
  const actual = sheet.getRange(1, 1, 1, MANTUL_V2_SCHEMA_.RAW.length)
    .getDisplayValues()[0].map(v => String(v == null ? '' : v).trim().toUpperCase());
  const expected = MANTUL_V2_SCHEMA_.RAW.map(v => v.toUpperCase());
  const mismatch = expected.findIndex((v, i) => actual[i] !== v);
  if (mismatch >= 0) {
    throw new Error('Header RAW_DATA_SALES berubah pada kolom ' + (mismatch + 1) +
      '. Diharapkan "' + expected[mismatch] + '", terbaca "' + actual[mismatch] + '".');
  }
  return {columns: MANTUL_V2_SCHEMA_.RAW.length, headers: actual};
}
