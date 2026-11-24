# Mantul Dashboard — Upload Bulanan

## Format workbook

Gunakan workbook yang memuat tiga sheet data berikut dengan header sesuai sumber:

- `P1_FACT` — agregat Page 1 (14 kolom)
- `P2_EVENT_INDEX` — detail event Page 2 (15 kolom)
- `RAW_DATA_SALES` — detail transaksi (25 kolom)

Sheet `SKU` dan `MAPPING TEAM` tidak diganti oleh proses upload bulanan. Perubahan pada dua referensi tersebut perlu proses pembaruan referensi tersendiri.

## Cara update bulan terbaru

1. Sinkronkan source terbaru dari branch `perf/page2-single-pass` ke project Google Apps Script.
2. Deploy Web App sebagai **New version**.
3. Buka dashboard, pilih **Upload Data Terbaru**, lalu pilih workbook lengkap.
4. Masukkan periode yang akan diperbarui dengan format `YYYY-MM`, misalnya `2026-11`. Default yang ditawarkan adalah periode terakhir yang memiliki nilai aktual di `P1_FACT`.
5. Periksa ringkasan jumlah baris dan nilai SOO, lalu konfirmasi.
6. Tunggu sampai status import berhasil. Setelah itu refresh dashboard.

## Apa yang dilakukan sistem

- Membaca hanya baris periode yang dipilih dari ketiga sheet.
- Memastikan `P2_EVENT_INDEX` dan `RAW_DATA_SALES` memiliki jumlah baris, jumlah transaksi per distributor/jenis sales, dan total nilai yang sama.
- Menyimpan salinan baris periode lama pada `P1_MONTH_BACKUP`, `P2_MONTH_BACKUP`, dan `RAW_MONTH_BACKUP`.
- Menghapus partisi tahun-bulan lama, lalu memasukkan data periode terbaru. Jadi upload ulang bulan yang sama **mengganti**, bukan menambahkan duplikat.
- Memperbarui status index analitik, timestamp, cache, dan `MONTHLY_IMPORT_LOG`.
- Jika publikasi gagal, sistem mencoba memulihkan data periode sebelumnya.

## Catatan

- Untuk memperbaiki bulan yang sudah pernah diunggah, pilih kembali periode yang sama. Sistem akan menggantinya secara idempotent.
- Untuk memperbarui beberapa bulan, ulangi upload satu periode setiap kali.
- Jangan gunakan workbook dengan header atau struktur kolom berbeda. Validasi akan menghentikan import sebelum data live diubah.
- Proses ini mengganti `P1_FACT`, `P2_EVENT_INDEX`, dan `RAW_DATA_SALES`; tidak mengganti `SKU` atau `MAPPING TEAM`.
