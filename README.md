# Bot Telegram Penomoran Surat Keluar Bidang

Bot Telegram untuk menerbitkan nomor surat keluar dengan pola:

`KODE_KLASIFIKASI/00001`

Contoh:

`500.13.3.1/00001`

## Fitur

- Pilih kode klasifikasi melalui tombol Telegram.
- Nomor urut 5 digit otomatis.
- Mode nomor global atau per klasifikasi.
- Database SQLite.
- Transaksi database untuk mencegah dua pengguna memperoleh nomor yang sama.
- Riwayat 10 nomor terakhir.
- Pencarian register surat.
- Pembatasan akses berdasarkan Telegram User ID.
- Admin dapat menentukan nomor berikutnya.
- Export register ke CSV.
- Cocok dipasang di Railway dengan persistent volume.

## Struktur nomor

Default:

`500.13.3.1/00001`

Urutan berikutnya:

`500.13.3.2/00002`

Karena default `SEQUENCE_MODE=global`, nomor 5 digit berjalan terus untuk seluruh surat keluar bidang, meskipun klasifikasinya berbeda.

Jika ingin setiap klasifikasi punya urutan sendiri:

`SEQUENCE_MODE=per_classification`

## Menjalankan secara lokal

1. Install Python 3.12 atau versi Python modern yang kompatibel.
2. Buat virtual environment.
3. Install dependency:

```bash
pip install -r requirements.txt
```

4. Salin `.env.example` menjadi konfigurasi environment Anda.
5. Set minimal:

```env
BOT_TOKEN=token_bot_anda
ADMIN_IDS=id_telegram_admin
```

6. Jalankan:

```bash
python main.py
```

> Program membaca environment variable dari sistem/deployment platform. File `.env` tidak dibaca otomatis agar deployment tetap sederhana dan aman.

## Perintah bot

- `/start` — menu utama
- `/baru` — buat nomor baru
- `/riwayat` — 10 nomor terakhir
- `/terakhir` — nomor terakhir
- `/cari kata` — cari register
- `/klasifikasi` — daftar klasifikasi
- `/id` — lihat Telegram ID
- `/batal` — batalkan proses

Admin:

- `/setnomor 123` — nomor berikutnya menjadi `00123`
- `/export` — export register ke CSV

## Deploy ke Railway

### 1. Upload project ke GitHub

Upload seluruh file project ini ke satu repository.

### 2. Buat service Railway dari GitHub

Railway akan menjalankan project menggunakan `Dockerfile`.

### 3. Tambahkan Variables

```env
BOT_TOKEN=...
ADMIN_IDS=...
ALLOWED_USER_IDS=...
DB_PATH=/data/surat_keluar.db
SEQUENCE_MODE=global
START_NUMBER=1
```

### 4. Pasang Volume

Mount persistent volume ke:

`/data`

Dengan begitu database tetap tersimpan ketika aplikasi redeploy/restart.

## Penting tentang nomor surat

Bot tidak menghapus atau memakai ulang nomor yang sudah diterbitkan. Ini sengaja dibuat agar register tetap konsisten.

Jika ingin melanjutkan nomor manual yang sudah berjalan, misalnya nomor terakhir sebelumnya `00122`, gunakan:

`/setnomor 123`

Maka nomor berikutnya yang diterbitkan menjadi `00123`.

## Klasifikasi

Data kode klasifikasi sudah dimasukkan sesuai tabel pada foto yang diberikan, mulai dari `500.13` sampai `500.13.6.4`.

## Pengembangan lanjutan yang mudah ditambahkan

- Format nomor lengkap dengan kode perangkat daerah/tahun.
- Reset nomor otomatis setiap awal tahun.
- Persetujuan admin sebelum nomor diterbitkan.
- Edit metadata register tanpa mengubah nomor.
- Status BATAL/VOID untuk surat yang batal digunakan.
- Export PDF atau Excel.
- Dashboard web.
- Integrasi dengan bot Surat Tugas yang sudah ada.
