# Mantul Dashboard — Handoff / Status Kerja

Terakhir diperbarui: 26 November 2026
Repository: https://github.com/Reriolov4/Mantul-Dashboard
Branch kerja tunggal: `perf/page2-single-pass`
PR: https://github.com/Reriolov4/Mantul-Dashboard/pull/1

## Sumber data dan kontrak yang sudah dipastikan
- Spreadsheet aktif: `1I8gESqZUkgpn-9yvQpWOmTFGWkW-3E06SG94a7UiZjI`.
- `RAW_DATA_SALES` adalah sumber transaksi, tepat 25 kolom A:Y.
- Workbook yang diperiksa berisi sekitar 130 ribu baris `RAW_DATA_SALES`; `P1_FACT` sekitar 14.7 ribu baris.
- Header A:Y didefinisikan di `Config.gs` / `RAW_SALES_HEADERS`. Jangan memakai schema 26/27 kolom dari cabang lama.
- `P1_FACT` dan seluruh logika Page 1 harus dipertahankan, tidak diubah dalam migrasi Page 2–4.
- `P2_EVENT_INDEX` saat ini adalah snapshot A:Y yang tervalidasi. Builder saat ini menyalin ulang semua baris; belum incremental dan belum menjadi aggregate fact.

## Aturan bisnis yang wajib dipertahankan
### Page 3
- Mode SOO: hanya `SALES_TYPE = SOO`; entitas berasal dari `SUBDIST`.
- Mode SO: `SALES_TYPE = Semua`; entitas berasal dari `SUB_DIST`; kosong dan `0` dikecualikan.
- Mapping khusus SO: `AP. GEUBRINA RIZKI` → EPM LSE; `PT RESTU INDO MEDICA` → EPM BAC.
- Tujuh entitas: EPM BAC, EPM LSE, GMP BAC, GMP LSE, GMP LGS, DB MBO, DB BAC.
### Page 2
- AO/EC mengikuti definisi bisnis terakhir yang disepakati dan wajib diuji terhadap angka pembanding pengguna sebelum mengganti formula.
- Klasifikasi Motoris/Canvaser harus memakai tabel `MAPPING TEAM`; jangan mengganti dengan tebakan dari nama.
### Page 4
- Dataset SOO dan master SKU; filter tahun, bulan, customer, dan entitas/SUBDIST harus benar-benar memengaruhi agregasi.

## Yang sudah dikerjakan
- Commit sebelumnya: `65437f0b93049212fdf4ab256d9908b4623a8249` — routing Page 2–4 ke read index A:Y yang tervalidasi.
- Commit: `5aeec398a3685474a8bb084ea125b11875d25470` — perbaikan backend Page 4 agar filter year/month diterapkan pada perhitungan.
- Pada sesi ini, Page 4 UI ditambah kontrol tahun dan bulan yang mengirim pilihan tersebut ke backend.
- Workflow GitHub diubah menjadi hanya `clasp push --force`. Ia tidak lagi menjalankan `clasp deploy`; pengelolaan versi/deployment Web App dilakukan pemilik Apps Script.
- Workflow tetap memerlukan GitHub Secrets `CLASP_CREDENTIALS` dan `SCRIPT_ID`. Jika keduanya valid, push ke branch aktif mengunggah source ke Apps Script. Keberhasilan harus diverifikasi pada Actions run terbaru dan project Apps Script yang benar.

## Yang belum dikerjakan / langkah berikut
1. Bangun tabel fakta materialized untuk Page 2–4 dengan builder chunk/resumable yang benar-benar sesuai 25 kolom A:Y. Jangan port builder dari `rebuild/database-v2` secara mentah: implementasi cabang tersebut mengasumsikan 27 kolom dan pemetaan kolom yang berbeda.
2. Rancang grain agregasi yang mempertahankan dimensi filter dan definisi AO/EC. Validasi total VAL/CARTON, jumlah source rows, serta hasil AO/EC terhadap angka kontrol.
3. Ubah Page 2 agar membaca aggregate/read model yang tepat, bukan memindai seluruh snapshot berulang kali.
4. Ubah Page 3 agar initial load dan Apply memakai fact terbangun; jangan membaca ulang 130 ribu event dan membangun metadata pada setiap klik.
5. Ubah Page 4 agar menggunakan fact SKU teragregasi; hindari pembacaan ulang seluruh event dan master SKU setiap Apply.
6. Implementasikan refresh incremental untuk baris baru yang ditambahkan di bawah data lama; sediakan full rebuild terkontrol untuk koreksi histori.
7. Uji halaman 2, 3 (SOO dan SO), dan 4; pastikan Page 1 tidak berubah.
8. Setelah source berhasil di-push, pemilik Apps Script melakukan Manage deployments / memilih versi baru sendiri.

## Catatan deployment
Push source ke Apps Script bukan berarti Web App live otomatis memakai versi baru. Workflow sekarang hanya mengunggah source. Pemilik project mengatur deployment/version dari Apps Script. Jangan mengklaim dashboard live telah diperbarui sebelum pemilik menyelesaikan langkah itu dan melakukan uji halaman.
