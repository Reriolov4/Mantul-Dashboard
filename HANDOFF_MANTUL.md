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
- `P2_EVENT_INDEX` adalah snapshot A:Y tervalidasi. Refresh kini menggunakan append-only incremental: bila hanya ada baris baru di bawah, hanya baris baru yang disalin dan diverifikasi; bila index tidak valid atau sumber menyusut, fallback ke full rebuild. Asumsi: histori lama tidak diedit. Ini belum menggantikan kebutuhan aggregate fact untuk mempercepat filter/Apply.

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
- Page 2: ditambahkan `buildPage2OutletFact()` untuk membangun `P2_OUTLET_FACT` dari `P2_EVENT_INDEX`. Grain menggabungkan baris yang identik pada tanggal/outlet/dimensi filter, menjumlahkan VAL, dan menyimpan SOURCE ROWS; builder memvalidasi source row count dan total VAL sebelum publish. Page 2 otomatis membaca fact jika status/source row count/header cocok, fallback ke event index bila fact belum tersedia.
- Page 2: identitas AO/EC pada agregator diarahkan ke `SUBDIST + CUSTOMER_CODE` yang valid (blank, `OTHERSB`, `#VALUE!`, `0` dikecualikan); EC menambahkan tanggal. Baris tanpa kode valid atau tanpa SUBDIST tetap dapat berkontribusi pada nilai penjualan, tetapi tidak pada AO/EC. Set AO/EC hanya menerima key yang valid (tidak menyisipkan key kosong). Perubahan definisi ini wajib diuji dengan angka kontrol sebelum dianggap final.
- `Code.gs`: `refreshMantulDashboard()` kini memanggil `refreshMantulReadIndexIncremental_()`. Refresh tanpa perubahan melewati penyalinan; penambahan data menyalin batch baru saja dan memvalidasi row count, 25 kolom, serta total VAL; kondisi index tidak valid fallback ke full rebuild.
- Commit sebelumnya: `65437f0b93049212fdf4ab256d9908b4623a8249` — routing Page 2–4 ke read index A:Y yang tervalidasi.
- Commit: `5aeec398a3685474a8bb084ea125b11875d25470` — perbaikan backend Page 4 agar filter year/month diterapkan pada perhitungan.
- Pada sesi ini, Page 4 UI ditambah kontrol tahun dan bulan yang mengirim pilihan tersebut ke backend.
- Workflow GitHub diubah menjadi hanya `clasp push --force`. Ia tidak lagi menjalankan `clasp deploy`; pengelolaan versi/deployment Web App dilakukan pemilik Apps Script.
- Workflow tetap memerlukan GitHub Secrets `CLASP_CREDENTIALS` dan `SCRIPT_ID`. Jika keduanya valid, push ke branch aktif mengunggah source ke Apps Script. Keberhasilan harus diverifikasi pada Actions run terbaru dan project Apps Script yang benar.


- Koreksi pasca-audit Page 2: daily AO/EC kini hanya membuat bucket saat tanggal valid dan hanya menambahkan EC key yang tidak kosong. Ini mencegah tanggal kosong dihitung sebagai satu EC pada agregasi harian.
- Verifikasi GitHub Actions commit `8eeed68b6b5eb2f256196d8733876e87bd432bd5`: run `45848042791` selesai `success`, termasuk step `Push source to Apps Script`. Workflow hanya `clasp push --force`, bukan `clasp deploy`. Pastikan secret `SCRIPT_ID` menunjuk project `MANTUL DASHBOARD V2`; workflow membangkitkan `.clasp.json` dari secret itu.

## Yang belum dikerjakan / langkah berikut
1. Jalankan `buildPage2OutletFact()` sekali dari Apps Script setelah source push dan `P2_EVENT_INDEX` siap; periksa log hasil `sourceRows`, `factRows`, `compressedRows`, `sourceValue`, `factValue`. Builder Page 2 saat ini satu kali jalan dan belum resumable; jika melewati batas runtime, pecah menjadi chunk/resumable sebelum pemakaian rutin.
2. Bandingkan AO/EC Page 2 dengan angka kontrol pengguna, terutama Jan–Mar AO (1121/1065/1050), AO total 3566, dan EC total valid customer code yang sebelumnya dibahas. Jangan lanjutkan optimasi KPI sebelum angka cocok. Jangan port builder dari `rebuild/database-v2` secara mentah: implementasi cabang tersebut mengasumsikan 27 kolom dan pemetaan kolom yang berbeda.
2. Rancang grain agregasi yang mempertahankan dimensi filter dan definisi AO/EC. Validasi total VAL/CARTON, jumlah source rows, serta hasil AO/EC terhadap angka kontrol.
3. Ubah Page 2 agar membaca aggregate/read model yang tepat, bukan memindai seluruh snapshot berulang kali.
4. Ubah Page 3 agar initial load dan Apply memakai fact terbangun; jangan membaca ulang 130 ribu event dan membangun metadata pada setiap klik.
5. Ubah Page 4 agar menggunakan fact SKU teragregasi; hindari pembacaan ulang seluruh event dan master SKU setiap Apply.
6. Implementasikan refresh incremental untuk baris baru yang ditambahkan di bawah data lama; sediakan full rebuild terkontrol untuk koreksi histori.
7. Uji halaman 2, 3 (SOO dan SO), dan 4; pastikan Page 1 tidak berubah.
8. Setelah source berhasil di-push, pemilik Apps Script melakukan Manage deployments / memilih versi baru sendiri.

## Catatan deployment
Push source ke Apps Script bukan berarti Web App live otomatis memakai versi baru. Workflow sekarang hanya mengunggah source. Pemilik project mengatur deployment/version dari Apps Script. Jangan mengklaim dashboard live telah diperbarui sebelum pemilik menyelesaikan langkah itu dan melakukan uji halaman.


## 26 Nov 2026 — RAW_DATA_SALES extra columns compatibility
- `RAW_DATA_SALES` runtime reported `getLastColumn() = 29` although the dashboard schema is A:Y (25 fields).
- `Code.gs` validation now requires at least 25 columns and reads/copies only A:Y; columns Z onward are ignored and never overwritten by the read-index refresh.
- Header validation remains limited to A:Y. This protects helper/formula columns Z:AC while keeping the dashboard's 25-column schema.
- Next: confirm the GitHub Actions `Push source to Apps Script` step succeeds, then run `refreshMantulDashboard()` and verify `P2_EVENT_INDEX` status COMPLETE before running `buildPage2OutletFact()`.


## 26 Nov 2026 — Page 2 fact built; shared Page 3/4 fact implementation
- User executed `buildPage2OutletFact()`: COMPLETE; sourceRows 130,367; factRows 90,872; compressedRows 39,495; sourceValue 81,777,256,858.05042; factValue 81,777,256,858.05019; validated true. Tiny floating-point delta is within validation tolerance.
- Added shared `P34_EVENT_FACT` materialized read model with dimensions needed by Page 3 (SOO/SO, SUBDIST/SUB_DIST mappings, customer, date, PIC, brand, subbrand, category) and Page 4 (SKU, SBU, carton/value). It groups exact same dimension rows and stores SOURCE ROWS.
- Page 3/4 readers prefer P34_EVENT_FACT only when its completion properties, spreadsheet ID, source row count, written count and header all match the current P2_EVENT_INDEX. Otherwise they safely fall back to P2_EVENT_INDEX.
- Added `buildPage34EventFact()` full rebuild. It validates source row count, VAL and CARTON before publishing. It is not incremental/resumable yet.
- Next: run Actions push; after source push succeeds, run `buildPage34EventFact()` once in Apps Script. Then test Page 3 SOO, Page 3 SO (including AP. GEUBRINA RIZKI → EPM LSE and PT RESTU INDO MEDICA → EPM BAC), and Page 4 filters. Compare outputs with existing results. Page 1/P1_FACT untouched.
- Page 2 AO/EC must still be checked against control values: Jan–Mar AO 1121/1065/1050; AO total 3566; EC valid-code control 19,328 (19,339 only if OTHERSB is included). Do not change KPI formulas before this reconciliation.
- After appending new source rows, P34 fact becomes stale by row-count mismatch and readers fall back safely to P2_EVENT_INDEX until the fact is rebuilt. Incremental P34 refresh remains future work.


## 26 Nov 2026 — Correct Page 2 fact column K mapping
- User identified that column K in `P2_OUTLET_FACT` was labeled `ACCOUNT NAME`, but the required field is `SALESMAN`.
- Root cause: Page2 raw reader populated fact position 11 (zero-based index 10) with a duplicate of CUSTOMER NAME instead of reading the actual `SALESMAN` field from `RAW_DATA_SALES` column S.
- Corrected `PAGE2_FACT_HEADERS_` column K to `SALESMAN`; raw reader now locates the `SALESMAN` header and reads that source field. Existing Page2 consumers already interpret index 10 as salesperson for the salesman picker, so the correction aligns stored data with intended semantics.
- Important: Existing `P2_OUTLET_FACT` is stale and still has old K values/header. After source push succeeds, run `buildPage2OutletFact()` again to rebuild and replace it. Do not manually edit the fact sheet. Check that K1 says SALESMAN and sample K values are actual salesperson names. Then verify salesman filter/options and AO/EC totals.
