# Bot Telegram Penomoran Surat Keluar Bidang — v8

Bot Telegram untuk menerbitkan nomor surat keluar dengan pola:

`KODE_KLASIFIKASI/NOMOR/118.4/TAHUN`

Contoh:

`500.13.3.4/123/118.4/2026`


## Klasifikasi tambahan

Versi v4 menambahkan klasifikasi:

```text
000.3.2 — Nota Penyampaian Rencana Pengadaan
```

Contoh nomor:

```text
000.3.2/123/118.4/2026
```

Nomor urut maksimal 3 digit tetap mengikuti sistem penomoran bot. Dengan `SEQUENCE_MODE=global`, urutan tetap melanjutkan nomor global seluruh surat.


### Klasifikasi v5

```text
000.1.2.3 — Perjalanan Dinas Pegawai
```

Contoh nomor:

```text
000.1.2.3/123/118.4/2026
```



## NPKND otomatis

Setelah nomor surat berhasil diterbitkan, bot menampilkan tombol:

```text
📄 Buat NPKND
```

Bot kemudian meminta:

1. **Kepada** - satu penerima per baris. Jika jumlah penerima lebih dari 2, dokumen otomatis menampilkan `Terlampir`.
2. **Tentang** - misalnya `Permohonan Tanda Tangan Surat Tugas`.
3. **Catatan** - uraian singkat. Bot otomatis menambahkan kalimat `sebagaimana berkas terlampir` pada bagian akhir.
4. **Lampiran** - pilihan `-`, `1 (satu) berkas`, atau `2 (dua) berkas`.

Tanggal dan nomor nota diambil otomatis dari register surat yang baru diterbitkan.

Bot mengirim dua file:

- PDF siap digunakan dengan tata letak mengikuti contoh;
- DOCX versi Word yang dapat diedit, menggunakan Arial ukuran 11.

Bagian kop, tujuan Kepala Dinas, asal Bidang, kalimat permohonan tanda tangan, disposisi, serta identitas Kepala Bidang dibuat tetap mengikuti template.


## Format nomor surat v7

Nomor surat sekarang menggunakan format:

```text
KODE_KLASIFIKASI/NOMOR_URUT/118.4/TAHUN
```

Contoh:

```text
500.13.3.4/123/118.4/2026
```

Ketentuannya:

- nomor urut tidak memakai nol di depan;
- `00123` ditampilkan sebagai `123`;
- nomor urut dibatasi maksimal tiga digit, yaitu `1` sampai `999`;
- angka `118.4` ditambahkan otomatis;
- tahun diambil otomatis dari **tanggal surat keluar**, bukan dari tanggal input bot;
- nomor lengkap otomatis dipakai pada NPKND, riwayat, pencarian, CSV, dan laporan Excel.


## Jenis surat

Setelah pengguna menulis perihal, bot menampilkan tiga pilihan:

- `NPKND`
- `Nota Dinas`
- `Surat Keluar`

Jenis surat disimpan ke database dan ikut tampil pada:

- hasil penerbitan nomor;
- riwayat;
- pencarian;
- export CSV;
- laporan bulanan Excel.

Setelah nomor berhasil diterbitkan, tombol dokumen sekarang bernama:

```text
Buat NPKND
```

## Fitur utama

- Pilih kode klasifikasi melalui tombol Telegram.
- Nomor urut otomatis maksimal 3 digit tanpa nol di depan.
- Tanggal surat keluar dicatat untuk setiap nomor.
- Pilihan tanggal hari ini atau input manual `DD-MM-YYYY`.
- Mode nomor global atau per klasifikasi.
- Database SQLite.
- Transaksi database untuk mencegah dua pengguna memperoleh nomor yang sama.
- Riwayat 10 nomor terakhir.
- Pencarian register surat.
- Pembatasan akses berdasarkan Telegram User ID.
- Admin dapat menentukan nomor berikutnya.
- Export seluruh register ke CSV.
- Laporan bulanan dalam format Excel `.xlsx`.
- Laporan Excel terdiri dari:
  - `Register Surat`
  - `Rekap Klasifikasi`
  - grafik jumlah surat per klasifikasi
- Migrasi otomatis database versi lama.
- Cocok dipasang di Railway dengan persistent volume.

## Alur pembuatan nomor

1. Pilih klasifikasi.
2. Pilih tanggal surat:
   - Hari ini
   - Masukkan tanggal lain
3. Isi perihal.
4. Isi tujuan/penerima.
5. Konfirmasi.
6. Nomor urut maksimal 3 digit diterbitkan dan disimpan.

## Laporan bulanan

Melalui tombol:

`📊 Laporan Bulanan`

Bot menampilkan pilihan 6 bulan terakhir.

Atau gunakan perintah:

```text
/laporan
```

Untuk bulan berjalan.

Bulan tertentu:

```text
/laporan 2026-07
```

Spreadsheet dibuat berdasarkan **tanggal surat keluar**, bukan tanggal input ke bot.

Kolom pada sheet `Register Surat`:

- No.
- Tanggal Surat
- Nomor Surat
- Kode Klasifikasi
- Klasifikasi
- Perihal
- Tujuan
- Dibuat Oleh
- Waktu Input
- Status

Sheet `Rekap Klasifikasi` berisi jumlah surat per kode klasifikasi dan grafik.

## Migrasi dari v1

Jika Anda sebelumnya memakai v1 dan database sudah tersimpan di Railway Volume, cukup update kode.

Saat v2 dijalankan pertama kali:

- kolom `letter_date` ditambahkan otomatis;
- data lama diisi tanggal berdasarkan tanggal input sebelumnya;
- nomor yang sudah ada tetap tersimpan;
- counter nomor tetap melanjutkan database lama.

Jangan hapus Railway Volume `/data`.


## Reset nomor setelah testing

Admin dapat mengembalikan bot ke kondisi awal dengan:

```text
/resetnomor
```

Bot akan menampilkan konfirmasi sebelum melakukan reset.

Saat dikonfirmasi, sistem akan:

- menghapus seluruh register surat percobaan;
- menghapus seluruh counter nomor;
- membuat nomor berikutnya kembali mulai dari `START_NUMBER`;
- default `START_NUMBER=1`, sehingga nomor urut berikutnya menjadi `1`.

Reset ini sengaja ikut menghapus data uji. Hanya mereset counter tanpa menghapus register lama dapat menimbulkan benturan nomor ketika nomor yang sama diterbitkan kembali.

Perintah ini hanya dapat dijalankan oleh Telegram ID yang terdaftar di:

```env
ADMIN_IDS=123456789
```

## Environment Variables Railway

```env
BOT_TOKEN=TOKEN_BOT_TELEGRAM
ADMIN_IDS=123456789
ALLOWED_USER_IDS=123456789,987654321
DB_PATH=/data/surat_keluar.db
SEQUENCE_MODE=global
START_NUMBER=1
APP_TIMEZONE=Asia/Jakarta
```

## Railway Volume

Mount path:

```text
/data
```

Database:

```text
/data/surat_keluar.db
```

## Perintah bot

- `/start` — menu utama
- `/baru` — buat nomor baru
- `/riwayat` — 10 nomor terakhir
- `/terakhir` — nomor terakhir
- `/cari kata` — cari register
- `/laporan` — laporan Excel bulan berjalan
- `/laporan 2026-07` — laporan Excel bulan tertentu
- `/klasifikasi` — daftar klasifikasi
- `/id` — lihat Telegram ID
- `/batal` — batalkan proses

Admin:

- `/setnomor 123` — nomor urut berikutnya menjadi `123`
- `/resetnomor` — hapus data uji dan mulai kembali dari nomor awal
- `/export` — export seluruh register ke CSV

## Deploy/update ke Railway

Jika bot v1 sudah berjalan:

1. Ganti/update file repository GitHub dengan versi v2.
2. Pastikan Railway Volume tetap terpasang di `/data`.
3. Tambahkan variable:

```env
APP_TIMEZONE=Asia/Jakarta
```

4. Railway akan redeploy.
5. Cek log sampai muncul:

```text
Bot started. DB=/data/surat_keluar.db | SEQUENCE_MODE=global | TIMEZONE=Asia/Jakarta
```

Data lama tidak perlu dihapus.
