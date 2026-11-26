/**
 * MANTUL DATABASE V2 — CANONICAL SCHEMAS
 * RAW_DATA_SALES adalah source of truth 25 kolom A:Y.
 * Schema output V2 dipertahankan agar kontrak layanan Page 2–4 tetap kompatibel.
 */
const MANTUL_V2_SCHEMA_ = Object.freeze({
  RAW: Object.freeze([
    'YEAR','MONTH','DD','KODE ITEM','NAMA ITEM','SUB BRAND','BRAND',
    'CARTON','SUBDIST','CUSTOMER_CODE','CUSTOMER_NAME','KODE SALUR',
    'VAL','QTY','SALES_TYPE','SUB_DIST','PIC','TEAM','SALESMAN',
    'CATEGORY','DOT','SBU','PF Bulanan','PF Triwulan','SKU'
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

/** Zero-based mapping dari RAW_DATA_SALES A:Y ke event index. */
function mantulV2MapRawToP2Event_(r, sourceRow) {
  const c = RAW_SALES_COL;
  return [
    r[c.YEAR], r[c.MONTH], r[c.DAY], r[c.CUSTOMER_CODE], r[c.CUSTOMER_NAME],
    r[c.SUBDIST], r[c.PIC], r[c.BRAND], r[c.SUBBRAND], r[c.CATEGORY],
    r[c.SBU], r[c.SALES_TYPE], r[c.SALESMAN], r[c.TEAM], r[c.VALUE],
    r[c.CARTON], r[c.SKU], r[c.SUB_DIST], sourceRow
  ];
}

/** P3 mempertahankan SUBDIST dan SUB_DIST untuk pemisahan SOO/SO. */
function mantulV2MapRawToP3Outlet_(r) {
  const c = RAW_SALES_COL;
  return [
    r[c.YEAR], r[c.MONTH], r[c.DAY], r[c.CUSTOMER_CODE], r[c.CUSTOMER_NAME],
    r[c.SUBDIST], r[c.SUB_DIST], r[c.SALES_TYPE], r[c.PIC], r[c.BRAND],
    r[c.SUBBRAND], r[c.CATEGORY], r[c.SBU], r[c.ITEM_CODE],
    r[c.SKU], r[c.VALUE], r[c.CARTON], 1
  ];
}

/** Gagal cepat jika header RAW_DATA_SALES tidak sesuai kontrak final 25 kolom. */
function mantulV2AssertRawHeaders_(sheet) {
  if (!sheet || sheet.getLastColumn() < MANTUL_V2_SCHEMA_.RAW.length) {
    throw new Error('RAW_DATA_SALES tidak memiliki 25 kolom sesuai schema V2.');
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
