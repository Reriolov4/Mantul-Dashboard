// ============================================================
// MANTUL DASHBOARD — CONFIGURATION / SOURCE OF TRUTH
// Database aktif: P1_FACT (dipertahankan untuk Page 1), RAW_DATA_SALES
// sebagai source of truth, serta indeks turunan Database V2 untuk Page 2–4.
// ============================================================

// ID spreadsheet utama. Script Property SPREADSHEET_ID boleh override.
const DEFAULT_SPREADSHEET_ID = '1I8gESqZUkgpn-9yvQpWOmTFGWkW-3E06SG94a7UiZjI';

// Script Properties opsional:
// SPREADSHEET_ID
// UPLOAD_ALLOWED_EMAILS (daftar email dipisahkan koma)

const DATA_SHEET_NAME = 'DATA_SALES';
const P1_FACT_SHEET_NAME = 'P1_FACT';
const STAGING_SHEET_NAME = 'IMPORT_STAGING';
const RAW_STAGING_SHEET_NAME = 'RAW_IMPORT_STAGING';
const RAW_DATA_SHEET_NAME = 'RAW_DATA_SALES';
const IMPORT_LOG_SHEET_NAME = 'IMPORT_LOG';
const SKU_SHEET_NAME = 'SKU';

// ------------------------------------------------------------
// P1_FACT — 14 kolom fakta analitik (A:N) pada workbook terbaru.
// SALES_HEADERS tetap merupakan format internal ringkas yang
// dipakai layanan dashboard; format ini mengikuti header P1_FACT.
// ------------------------------------------------------------
const DATA_SALES_SOURCE_HEADERS = [
  'YEAR', 'MONTH', 'SUBDIST', 'SUBBRAND', 'BRAND', 'CATEGORY',
  'SBU', 'PIC', 'PF_MONTHLY', 'PF_QUARTERLY', 'VAL',
  'SALES_TYPE', 'TARGET', 'BE'
];

const SALES_HEADERS = [
  'y', 'm', 'd', 's', 'b', 'c', 'u',
  'p', 'fm', 'fq', 'v', 't', 'tg', 'be'
];

// ------------------------------------------------------------
// RAW_DATA_SALES — kontrak final 25 kolom (A:Y).
const RAW_SALES_HEADERS = [
  'YEAR', 'MONTH', 'DD', 'KODE ITEM', 'NAMA ITEM', 'SUB BRAND', 'BRAND',
  'CARTON', 'SUBDIST', 'CUSTOMER_CODE', 'CUSTOMER_NAME', 'KODE SALUR',
  'VAL', 'QTY', 'SALES_TYPE', 'SUB_DIST', 'PIC', 'TEAM', 'SALESMAN',
  'CATEGORY', 'DOT', 'SBU', 'PF Bulanan', 'PF Triwulan', 'SKU'
];
const RAW_SALES_COLUMN_COUNT = 25;

const RAW_SALES_COL = Object.freeze({
  YEAR: 0, MONTH: 1, DAY: 2, ITEM_CODE: 3, ITEM_NAME: 4,
  SUBBRAND: 5, BRAND: 6, CARTON: 7, SUBDIST: 8,
  CUSTOMER_CODE: 9, CUSTOMER_NAME: 10, CHANNEL_CODE: 11,
  VALUE: 12, QTY: 13, SALES_TYPE: 14, SUB_DIST: 15,
  PIC: 16, TEAM: 17, SALESMAN: 18, CATEGORY: 19,
  DOT: 20, SBU: 21, PF_MONTHLY: 22, PF_QUARTERLY: 23,
  SKU: 24
});

// ------------------------------------------------------------
// SKU — 16 kolom sumber pada workbook terbaru.
// ------------------------------------------------------------
const SKU_HEADERS = [
  'ITEM_CODE', 'ITEM_DESC', 'CLASS_PRODUCT_DESC', 'CLASS PRODUCT',
  'PRODUCT_DESC', 'UMBRAND', 'SUB_BRAND', 'BRAND', 'UNBRAND',
  'DOT', 'CATEGORY', 'UOM', 'HARGA', 'ENTITAS',
  'PF Bulanan', 'PF Triwulan'
];
