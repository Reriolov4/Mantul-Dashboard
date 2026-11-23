// ============================================================
// MANTUL DASHBOARD — CONFIGURATION / SOURCE OF TRUTH
// Database aktif: P1_FACT, P2_EVENT_INDEX, RAW_DATA_SALES, SKU,
// dan MAPPING TEAM. P1_FACT/P2_EVENT_INDEX adalah sumber analitik utama.
// ============================================================

// ID spreadsheet utama. Script Property SPREADSHEET_ID boleh override.
const DEFAULT_SPREADSHEET_ID = '1I8gESqZUkgpn-9yvQpWOmTFGWkW-3E06SG94a7UiZjI';

// Script Properties opsional:
// SPREADSHEET_ID
// UPLOAD_ALLOWED_EMAILS (daftar email dipisahkan koma)

const DATA_SHEET_NAME = 'DATA_SALES';
const STAGING_SHEET_NAME = 'IMPORT_STAGING';
const RAW_STAGING_SHEET_NAME = 'RAW_IMPORT_STAGING';
const RAW_DATA_SHEET_NAME = 'RAW_DATA_SALES';
const IMPORT_LOG_SHEET_NAME = 'IMPORT_LOG';
const SKU_SHEET_NAME = 'SKU';

// ------------------------------------------------------------
// DATA_SALES — 14 kolom sumber (A:N) pada workbook terbaru.
// SALES_HEADERS tetap merupakan format internal ringkas yang
// dipakai proses import lama; jangan ubah tanpa mengubah importer.
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
// RAW_DATA_SALES — transaksi utama tepat 27 kolom (A:AA).
// Kolom transaksi yang dipakai dashboard hanya A:AA (27 kolom).
// Kolom AH:AI adalah tabel pembantu/formula pada workbook sumber untuk
// membantu pengisian array A:AA; bukan sumber transaksi dan tidak diimpor.
// ------------------------------------------------------------
const RAW_SALES_HEADERS = [
  'TAHUN',          // A
  'BULAN',          // B
  'TANGGAL',        // C — angka/hari dalam bulan pada workbook
  'KODE ITEM',      // D
  'NAMA ITEM',      // E
  'SUB BRAND',      // F
  'BRAND',          // G
  'KARTON',         // H
  'SUBDIST',        // I
  'KODE CUSTOMER',  // J
  'CUSTOMER NAME',  // K
  'KODE SALUR',     // L
  'VAL',            // M
  'QTY',            // N
  'CLASS PRODUCT',  // O
  'UMBRAND',        // P
  'DOT',            // Q
  'CATEGORY',       // R
  'JENIS SALES',    // S
  'PIC',            // T
  'SUB_DIST',       // U
  'TIM',            // V
  'SALESMAN',       // W
  'SBU',            // X
  'PF Bulanan',     // Y
  'PF Triwulan',    // Z
  'PRODUCT_DESC'    // AA
];

const RAW_SALES_COLUMN_COUNT = RAW_SALES_HEADERS.length; // 27

// Indeks kolom RAW_DATA_SALES berbasis 0 untuk JavaScript.
// Gunakan konstanta ini agar kode tidak memakai angka indeks tersebar.
const RAW_SALES_COL = Object.freeze({
  YEAR: 0,
  MONTH: 1,
  DAY: 2,
  ITEM_CODE: 3,
  ITEM_NAME: 4,
  SUBBRAND: 5,
  BRAND: 6,
  CARTON: 7,
  SUBDIST: 8,
  CUSTOMER_CODE: 9,
  CUSTOMER_NAME: 10,
  CHANNEL_CODE: 11,
  VALUE: 12,
  QTY: 13,
  CLASS_PRODUCT: 14,
  UMBRAND: 15,
  DOT: 16,
  CATEGORY: 17,
  SALES_TYPE: 18,
  PIC: 19,
  SUB_DIST: 20,
  TEAM: 21,
  SALESMAN: 22,
  SBU: 23,
  PF_MONTHLY: 24,
  PF_QUARTERLY: 25,
  PRODUCT_DESC: 26
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
