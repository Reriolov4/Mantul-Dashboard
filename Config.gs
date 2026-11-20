// Ganti dengan ID Google Spreadsheet database Anda.
const SPREADSHEET_ID = '1I8gESqZUkgpn-9yvQpWOmTFGWkW-3E06SG94a7UiZjI';

const DATA_SHEET_NAME = 'DATA_SALES';
const STAGING_SHEET_NAME = 'IMPORT_STAGING';
const RAW_STAGING_SHEET_NAME = 'RAW_IMPORT_STAGING';
const RAW_DATA_SHEET_NAME = 'RAW_DATA_SALES';
const IMPORT_LOG_SHEET_NAME = 'IMPORT_LOG';

// Batasi akun yang boleh mengunggah data.
// Hanya akun berikut yang diberi hak upload.
// Pastikan akun ini sama dengan akun Google yang dipakai saat membuka Web App.
// Daftar kosong berarti seluruh upload ditolak (aman secara default).
const UPLOAD_ALLOWED_EMAILS = ['Reriolov4@gmail.com'];

const SALES_HEADERS = ['y','m','d','s','b','c','u','p','fm','fq','v','t','tg','be'];

// Struktur sumber RAW_DATA_SALES: A:AH (34 kolom).
const RAW_SALES_HEADERS = [
  'TAHUN','BULAN','DATE','KODE ITEM','NAMA ITEM','SUB BRAND','BRAND','ENTITAS',
  'JENIS FAKTUR','KARTON','DISTRIBUTORS','KODE CUSTOMER','CUSTOMER NAME','ALAMAT',
  'SALESMAN','KODE SALUR','VAL','QTY','CLASS PRODUCT','DOT','CATEGORY','JENIS SALES',
  'OUTLET SALUR','CABANG','SUBDIST','SISO','TIM','NAMA AKUN','CATEGORY','DOT','SBU',
  'PF Bulanan','PF Triwulan','PRODUCT_DESC'
];
