# Bot Telegram Penomoran Surat Keluar Bidang — v2

Bot Telegram untuk menerbitkan nomor surat keluar dengan pola:

`KODE_KLASIFIKASI/00001`

Contoh:

`500.13.3.1/00001`

## Fitur utama

- Pilih kode klasifikasi melalui tombol Telegram.
- Nomor urut otomatis 5 digit.
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
6. Nomor 5 digit diterbitkan dan disimpan.

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

- `/setnomor 123` — nomor berikutnya menjadi `00123`
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
