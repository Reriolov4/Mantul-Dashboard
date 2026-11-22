// ID Spreadsheet utama. Script Property SPREADSHEET_ID boleh override nilai ini.
// Jika Script Properties kosong, Web App otomatis memakai ID cadangan ini.
const DEFAULT_SPREADSHEET_ID = '1I8gESqZUkgpn-9yvQpWOmTFGWkW-3E06SG94a7UiZjI';

// Konfigurasi sensitif lain disimpan di Apps Script > Project Settings > Script Properties.
// Key opsional: SPREADSHEET_ID (override) dan UPLOAD_ALLOWED_EMAILS (email dipisahkan koma).

const DATA_SHEET_NAME = 'DATA_SALES';
const STAGING_SHEET_NAME = 'IMPORT_STAGING';
const RAW_STAGING_SHEET_NAME = 'RAW_IMPORT_STAGING';
const RAW_DATA_SHEET_NAME = 'RAW_DATA_SALES';
const IMPORT_LOG_SHEET_NAME = 'IMPORT_LOG';

const SALES_HEADERS = ['y','m','d','s','b','c','u','p','fm','fq','v','t','tg','be'];

// Skema transaksi RAW_DATA_SALES terbaru: A:AA (27 kolom).
// Tabel mapping/catatan di AH:AL bukan bagian dari baris transaksi dan tidak diimpor.
const RAW_SALES_HEADERS = [
  'TAHUN','BULAN','TANGGAL','KODE ITEM','NAMA ITEM','SUB BRAND','BRAND',
  'KARTON','SUBDIST','KODE CUSTOMER','CUSTOMER NAME','KODE SALUR','VAL','QTY',
  'CLASS PRODUCT','UMBRAND','DOT','CATEGORY','JENIS SALES','PIC','SUB_DIST',
  'TIM','SALESMAN','SBU','PF Bulanan','PF Triwulan','PRODUCT_DESC'
];
