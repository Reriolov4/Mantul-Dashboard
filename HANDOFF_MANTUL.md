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


## 27 Nov 2026 — Source audit and Page 3/4 diagnostics
- Confirmed Page 2 runtime source selection: `page2Source_()` prefers validated `P2_OUTLET_FACT`; it falls back to `P2_EVENT_INDEX` if fact status/source ID/source row count/written count/header validation fails. Page 2 does not directly read `RAW_DATA_SALES` during normal dashboard display/filter requests. `RAW_DATA_SALES` is upstream source used by the index refresh/build pipeline. Note: Page 2 still reads the complete selected fact/index into memory on a cache miss; `page2ReadFiltered_()` currently ignores its filters, so selective range/partition reading remains the next major performance task.
- Corrected stale Page 2 and Page 3/4 service comments that incorrectly claimed direct `RAW_DATA_SALES` reads. Page 2 fallback/read-model comments now reflect `P2_OUTLET_FACT` / `P2_EVENT_INDEX`.
- Page 3 and Page 4 now return split server diagnostics: `readMs`, `metaMs`, `aggregateMs`, `totalMs`, source name and source row count. Their browser status lines display server breakdown and client response duration on initial load / Apply.
- Page 3 and Page 4 loading status now names `P34_EVENT_FACT` rather than `RAW_DATA_SALES`. Existing `aria-busy` toggles are retained; visual overlay behavior still needs browser verification against CSS.
- No data model, KPI formula, mapping rule, or Page 1 / `P1_FACT` logic was changed. No fact rebuild is required for these diagnostics-only changes.
- Next: verify Actions push, then user manages Apps Script deployment. Test Page 3 SOO, Page 3 SO, and Page 4 filters; capture the new read/meta/aggregate timings. After timings, implement compact page-specific read models or partitioned reads to stop rescanning the entire `P34_EVENT_FACT` on every Apply. For Page 2, design a filter-aware index/range strategy while preserving outlet/date grain and AO/EC definitions. Do not claim sub-3-second performance until measured.

- Source push verification: commit `7f6791e7d7b8f73e9fe592417912d30eac898139` triggered Actions run `45865170823`; job `deploy` completed `success`, including `Push source to Apps Script`. This confirms source upload only; Web App deployment/version remains owner-managed.


## 27 Nov 2026 — Navigation and Page 3/4 click regression
- User test: Page 2 first load 3.314s; filter GMP LSE 15.0s (server 12.960s, read 10.203s, aggregate 153ms); adding Komix Herbal 23.1s (server 21.407s, read 18.995s, aggregate 234ms). Read time dominates; aggregation is not the bottleneck. Next optimization must target selective/partitioned reads and avoid full fact scans on each filter. Do not change KPI formulas to chase this timing.
- Root cause of Page 3 and Page 4 navigation buttons being inert: the initial-load success-handler status text in both `Page3Scripts.html` and `Page4Scripts.html` was missing the closing quote before the callback close (`...+' ms)}`). This is a JavaScript syntax error, so each entire script block fails to parse and functions `page3Show` / `page4Show` are never registered. Corrected both callback strings to close before `}`).
- Navigation jump root cause: `mantulShowPage_()` called `window.scrollTo(0,0)` on every page switch. Since the global banner/header is above the sticky nav, this jumped back to the top and made the nav appear below the banner. Removed forced scroll reset; current scroll position is preserved, allowing the sticky navigation to remain pinned at the top when the user has already scrolled there.
- Updated Page 2 initial loading status text to identify `P2_OUTLET_FACT` rather than imply a direct `RAW_DATA_SALES` read.
- No Page 1/P1_FACT business logic or Page 2 KPI formulas changed.
- Next: confirm source push succeeds, then user publishes via Apps Script Manage deployments. Test clicking Page 3 and Page 4 from the sticky nav; verify navigation remains pinned when switching from a scrolled position. Re-run Page 2 timing tests. After confirmation, implement filter-aware read model/partition strategy to reduce the 10–19s read stage. Do not rebuild existing facts for this UI/script syntax fix.


## 27 Nov 2026 — Page 2–4 follow-up: loading, filters, navigation
- User timing sample: Page 2 initial 2.885s; Komix Herbal filter server 9.418s (read 7.139s, aggregate 231ms); adding GMP LSE server 13.170s (read 11.071s, aggregate 239ms). Page 3 initial server 12.706s (read 12.418s); Komix Herbal server 17.996s (read 17.408s); month filter server 41.058s (read 39.953s). Page 4 initial server 41.609s (read 41.105s); month filter server 25.164s (read 24.814s). The read stage dominates; aggregate times remain very small.
- Added Page 3 and Page 4 blur/loading overlays matching Page 2's loading treatment. Busy callbacks now update their own loading message.
- Standardized Page 3 and Page 4 multi-select behavior: Select All/Reset/Apply, outside-click dismissal, a single scrollable checklist inside a non-scrolling dropdown, and explicit `display:none!important` for search-hidden items.
- Page 4 now exposes Month, Date, PIC, Brand, Subbrand, Category, Customer and SUBDIST filters in addition to Year. Its metadata and aggregation apply each selected dimension; SKU/value/carton aggregation definitions remain unchanged.
- Navigation reserves enough main-content height when switching to a shorter Page 3/4 so the browser does not clamp scroll position and reveal the banner above the sticky navigation. Page 1 removes this temporary minimum height.
- Page 1/P1_FACT calculations and business logic were not changed. No fact rebuild is required for these UI/filter changes.
- Performance work remains: Page 2/3/4 still read substantial fact ranges on cache misses. Next inspect selective-column reads and then design page-specific materialized/partitioned read models. Preserve outlet/date grain and AO/EC rules; validate totals before changing any data model. Test all Page 3/4 filters, search match/no-match, outside-click dismissal, one-scroll dropdown, loading blur and sticky navigation while scrolled.


## 27 Nov 2026 — Repair CSS leak and refresh after RAW_DATA append
- User reported raw Page 3/4 CSS declarations rendered as visible text above the banner. Root cause: the appended Page 3/4 CSS rules were placed after the closing `</style>` of `Styles.html`. Wrapped that entire block in a dedicated `<style id="mantul-page34-filter-loading">` element.
- User appended new sales rows through 26 Nov to `RAW_DATA_SALES`; previous verified read index/facts represented data through 25 Nov. The incremental refresh function `refreshMantulDashboard()` calls `refreshMantulReadIndexIncremental_()`: it appends only rows after the last indexed source row, validates row count and VAL, and falls back to full index build if the index metadata is invalid or source shrank.
- After running `refreshMantulDashboard()`, rebuild dependent facts so all pages reflect the new source row count: `buildPage2OutletFact()` and `buildPage34EventFact()`. Do not run Page 1/P1_FACT builders. Verify the index refresh reports `APPENDED` (or `UNCHANGED` if there truly are no new rows) with `validated:true`; verify both fact builders report `COMPLETE` and `validated:true`.
- Until the index is refreshed, Page 2 intentionally errors when no ready `P2_EVENT_INDEX` exists; after index refresh but before Page 2 fact rebuild, Page 2 can fall back to the raw event index, but will be slower. Page 3/4 likewise need their shared P34 fact rebuilt to reflect the new source.
- Current update procedure for appended source rows: (1) append rows in RAW_DATA_SALES A:Y; (2) run `refreshMantulDashboard()`; (3) run `buildPage2OutletFact()`; (4) run `buildPage34EventFact()`; (5) validate row counts and source/fact totals. Page 1 and P1_FACT are explicitly out of scope.
