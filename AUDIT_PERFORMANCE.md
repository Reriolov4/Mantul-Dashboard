# Audit & Update — Mantul Dashboard

Tanggal audit: 20 November 2026
Sumber: `Mantul Updates.zip` (9 file source; tidak termasuk isi spreadsheet).

## Ringkasan perubahan yang sudah diterapkan

1. **CTR Matrix Halaman 1**
   - Ditemukan ketidaksamaan periode: total pembagi (`grandTotal`) memakai bulan yang dipilih, sementara sebagian nilai baris dihitung sebagai YTD dari Januari sampai bulan terakhir.
   - `renderMatrix()` kini memakai `periodMonths` yang sama dengan total terpilih.
   - `monthlyMetrics()` kini menghitung actual, target, LY, dan average dari bulan yang benar-benar dipilih.
   - Ini menghindari CTR/kontribusi yang membandingkan numerator dan denominator dari rentang bulan berbeda.

2. **Optimasi aset CSS**
   - `Styles.html` memuat gambar dekoratif header PNG base64 yang sangat besar.
   - Gambar dikompres menjadi JPEG progressive kualitas 82, dengan tampilan visual yang dipertahankan.
   - Ukuran `Styles.html`: 624,800 byte menjadi 157,465 byte (turun sekitar 74.8%).
   - Total 9 file source: 773,666 byte menjadi 308,266 byte (turun sekitar 60.1%).

3. **Halaman 2**
   - Normalisasi nama bulan bahasa Inggris untuk `APRIL` diperbaiki menjadi `APR`, agar konsisten dengan kode bulan grafik.
   - Pembacaan RAW_DATA_SALES dikurangi dari 34 kolom menjadi 31 kolom (A:AE), sesuai kolom terjauh yang dipakai fungsi Page 2.
   - Ditambahkan cache untuk metadata filter (300 detik) dan hasil dashboard dengan filter yang sama (180 detik). Cache memakai versi `PAGE2_RAW_UPDATED`, sehingga import baru otomatis menggunakan namespace cache baru.
   - Ditambahkan request sequence guard agar hasil permintaan filter lama tidak menimpa hasil filter terbaru saat pengguna mengubah filter dengan cepat.

## Audit struktur file

Semua 9 file di ZIP digunakan oleh aplikasi saat ini:
- `Code.gs`: entry point Web App (`doGet`) dan fungsi include HTML.
- `Config.gs`: ID spreadsheet, nama sheet, header dan izin upload.
- `DataService.gs`: pembacaan DATA_SALES, import/staging, validasi, backup, cache.
- `Page2Service.gs`: metadata dan kalkulasi Halaman 2 dari RAW_DATA_SALES.
- `Index.html`: kerangka Halaman 1, navigasi dan include halaman.
- `Page2.html`: markup dan style lokal Halaman 2.
- `Scripts.html`: kalkulasi, render KPI/chart/matrix, upload dan export Halaman 1.
- `Page2Scripts.html`: filter, chart dan navigasi Halaman 2.
- `Styles.html`: style utama Halaman 1 dan header.

Tidak ada file source yang dihapus karena semua file tersebut dipanggil atau dibutuhkan oleh struktur aplikasi. Library Chart.js, XLSX dan html2canvas juga tidak dihapus: masing-masing dipakai untuk grafik, upload Excel, dan screenshot/export.

## Temuan penting yang belum selesai

### A. Pembacaan data belum benar-benar selektif
- `getDashboardData()` di `DataService.gs` masih membaca seluruh baris `DATA_SALES`, membentuk array `rows`, lalu mengirim seluruh array ke browser.
- `page2Read_()` masih membaca seluruh baris `RAW_DATA_SALES` pada cache miss atau ketika kombinasi filter berubah.
- Cache mempercepat pemanggilan berulang dengan filter yang sama, tetapi **bukan** pengganti server-side query.

### B. Pagination belum tersedia
- Halaman 1 merender seluruh baris matrix ke DOM.
- Halaman 2 saat ini berupa KPI dan grafik, belum memiliki tabel detail dengan pagination.
- Pagination yang benar perlu memindahkan filter, agregasi dan pemotongan halaman ke server. Menambahkan tombol Next/Previous di browser saja tidak akan mengurangi pembacaan seluruh spreadsheet.

### C. Target waktu di bawah 3 detik
- Perubahan di paket ini mengurangi aset dan pekerjaan berulang, tetapi waktu di bawah 3 detik belum dapat dijamin tanpa pengujian terhadap jumlah baris aktual, ukuran payload, dan latensi Google Sheets.
- Untuk memenuhi target secara konsisten, tahap berikutnya perlu membuat API ringkasan per filter dan API matrix per halaman, dengan agregasi server-side serta indeks/cache hasil agregasi.

## Rekomendasi arsitektur tahap berikutnya

1. Saat halaman dibuka, muat hanya metadata filter dan KPI ringkas.
2. Saat filter berubah, kirim parameter filter ke server; server mengembalikan ringkasan yang diperlukan saja.
3. Matrix mengambil 25–50 baris per halaman melalui `getMatrixPage(filters, page, pageSize)`, bukan seluruh baris.
4. Chart meminta agregat per dimensi (category/brand/subdist) tanpa mengirim seluruh transaksi ke browser.
5. Bangun cache agregat yang dibersihkan/di-version setiap import selesai.
6. Halaman 2 menggunakan agregat AO/EC per bulan/tanggal/dimensi agar tidak mengulang scan mentah untuk setiap grafik.
7. Uji cold load dan warm load dengan data aktual. Target: p95 di bawah 3 detik, bukan hanya sekali berhasil.

## Pemasangan

1. Buka project Apps Script dengan ID yang diberikan.
2. Backup project aktif terlebih dahulu.
3. Ganti isi file source dengan file bernama sama dari ZIP ini.
4. Simpan seluruh file, lalu deploy versi Web App baru.
5. Uji minimal: filter 1 bulan, 3 bulan non-berurutan, seluruh bulan, SO/SOO, dan kombinasi filter Halaman 2.
6. Periksa khusus CTR matrix Category, Brand, Subbrand, Subdist, dan baris B7/SAKA.

Catatan: ZIP ini berisi source yang diperbarui, bukan hasil deploy ke project Apps Script. Saya tidak melakukan perubahan langsung ke project Google Apps Script atau spreadsheet.
