# Mantul Dashboard — Database V2

## Keputusan schema

- `RAW_DATA_SALES`: sumber transaksi asli, 27 kolom, tidak pernah ditulis oleh builder.
- `P2_EVENT_INDEX`: 19 kolom; 18 kolom transaksi kanonis ditambah `SOURCE_ROW`.
- `P2_DASHBOARD_FACT`: 17 kolom; agregat Page 2 dengan `SOURCE_ROWS`.
- `P3_OUTLET_MONTHLY`: 18 kolom; grain harian outlet–brand–SKU, mencakup SOO dan SO, dipakai bersama oleh Page 3 dan Page 4.
- `P1_FACT`: dipertahankan sementara karena masih menjadi sumber Page 1. Tidak dihapus sebelum Page 1 dimigrasikan dan divalidasi.
- Tidak membuat `P4_SUBDIST_INVENTORY`.

## Aturan grain dan mapping

### P2_EVENT_INDEX
Satu baris untuk setiap baris data RAW. `SOURCE_ROW` adalah nomor baris spreadsheet asal (baris pertama data = 2). `SKU` memakai `PRODUCT_DESC` agar konsisten dengan kontrak SKU master yang sudah dipakai Page 4.

### P2_DASHBOARD_FACT
Agregasi Page 2 dari event index; hanya SOO sesuai perilaku agregasi Page 2 yang berjalan sekarang. Dimensi dan urutan kolom mengikuti 16 kolom pertama event index; `SOURCE_ROWS` menghitung jumlah transaksi yang digabung.

### P3_OUTLET_MONTHLY
Satu baris untuk kombinasi YEAR, MONTH, DD, CUSTOMER_CODE, CUSTOMER_NAME, SUBDIST, SUB_DIST, SALES_TYPE, PIC, BRAND, SUBBRAND, CATEGORY, SBU, ITEM_CODE, dan SKU. VAL/CARTON dijumlahkan dan SOURCE_ROWS dihitung. Penyimpanan detail harian memungkinkan UI membentuk tren bulanan serta drilldown outlet → brand → SKU.

## Kebijakan migrasi

1. Bangun output ke sheet staging; jangan menimpa sheet aktif selama build.
2. Proses RAW dalam chunk yang dapat dilanjutkan (resumable).
3. Validasi jumlah baris dan total VAL/CARTON; P2 harus cocok dengan transaksi SOO sesuai filter bisnis Page 2, P3 harus cocok dengan cakupan SOO+SO yang disepakati.
4. Aktivasi tabel dilakukan setelah semua validasi berhasil.
5. Pertahankan backup sampai Page 2/3/4 dan import bulanan lolos uji.
6. Jangan menjalankan `runMantulRebuildOnce()` atau builder monolitik lama.

## Audit pembersihan script

File lama belum dihapus pada tahap schema ini karena beberapa di antaranya masih menyediakan konstanta/fungsi yang dipakai runtime:
- `AnalyticsBuilder.gs`: konstanta header P1/P2 masih dipakai service lain.
- `AnalyticsIndexAudit.gs`: helper registrasi masih dipanggil Page 3 / Page 2.
- `AnalyticsSummaryService.gs`: helper summary/update period masih dipakai import bulanan.
- `MonthlyImportService.gs`: endpoint import aktif.
- `DashboardRefreshService.gs`: pemeriksaan readiness aktif.

Pembersihan dilakukan setelah pengganti tersedia dan semua referensi runtime dipindahkan. Fungsi monolitik lama akan dinonaktifkan/dihapus dalam tahap migrasi builder, bukan dengan menghapus file yang masih memiliki dependency.
