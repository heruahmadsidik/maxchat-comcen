# Maxchat Command Center — PLN UP2D Banten

Aplikasi internal untuk broadcast WhatsApp dan pengelolaan percakapan dua arah, dibangun di atas API **Maxchat**. Digunakan oleh ComCen PLN UP2D Banten untuk mengirim informasi gangguan/pemadaman ke pelanggan dan memantau balasannya dalam satu tampilan seperti WhatsApp Web.

## Fitur Utama

- 🔐 **Login & sesi** — seluruh aplikasi (kecuali `/login` dan `/webhook`) dilindungi login berbasis session (cookie, idle timeout 12 jam). Username/password login dikonfigurasi dari menu Pengaturan.
- 🛡️ **Password konfirmasi untuk pengaturan sensitif** — perubahan API Token, Webhook URL, dan kredensial login memerlukan "password konfirmasi" tambahan agar operator biasa tidak bisa iseng mengubahnya.
- 📢 **Broadcast** — kirim template WhatsApp ke banyak nomor sekaligus, dengan parameter dinamis dan progress bar live (polling status kirim per nomor).
- 📝 **Manajemen Template** — CRUD template pesan + urutkan ulang, kategori pekerjaan (WO Category, mis. Marking/Individu/Umum) per template.
- 💬 **Chat dua arah** — riwayat pesan masuk & keluar per kontak, balas teks/media langsung dari aplikasi, mendukung pesan lokasi (koordinat & link Google Maps), dan **quick reply** (balasan cepat siap pakai, bisa dikelola & diurutkan).
- 🔎 **Pencarian cepat** — pencarian kontak & riwayat sesi memakai index full-text search (SQLite FTS5), dengan fallback otomatis ke pencarian `LIKE` bila FTS5 tidak tersedia di build `better-sqlite3` yang terpasang.
- 🏷️ **Kategori kontak persisten** — badge kategori (WO Category) tidak hilang meskipun pelanggan membalas; hanya berubah saat broadcast baru dikirim ke nomor tersebut.
- 🏷️ **Label kontak (master label)** — label harus dibuat dulu lewat menu "Kelola Label" (nama + warna), baru bisa ditempelkan/dilepas ke kontak mana pun — tidak lagi bebas ketik teks di kontak.
- 👤 **Manajemen kontak** — daftar semua nomor yang pernah dikirimi pesan, digabung dengan kontak yang ditambahkan manual (belum pernah dikirimi pesan); nama kontak bisa diedit manual.
- 📊 **Dashboard & laporan** — rekap broadcast harian/bulanan, tren WO, daftar "pelapor sering" (frequent reporters), serta info database (ukuran, jumlah sesi, status backup).
- 📤 **Export data** — ekspor rekap broadcast maupun data mentah per nomor ke **XLSX atau CSV**.
- 🗄️ **Backup database otomatis** — backup penuh SQLite dijalankan otomatis tiap hari jam 04:00 WIB memakai Online Backup API bawaan `better-sqlite3` (tanpa mematikan aplikasi), backup lama dibersihkan otomatis (menyisakan sejumlah backup terbaru), plus tombol backup manual dari Pengaturan.
- 🗑️ **Manajemen riwayat** — hapus satu sesi, semua riwayat, atau riwayat berdasarkan rentang tanggal tertentu.
- 🔄 **Sinkronisasi via webhook** — status pengiriman (delivered/read), pesan keluar dari WhatsApp Web Maxchat, dan pesan masuk dari pelanggan.

## Tech Stack

| Komponen | Teknologi |
|---|---|
| Backend | Node.js + Express |
| Sesi & login | `express-session` (cookie-based, in-memory store) |
| Database | SQLite (`better-sqlite3`, mode WAL, dengan index FTS5 untuk pencarian) |
| Upload file | `multer` (in-memory, maks 16MB) |
| HTTP client ke Maxchat | `axios` |
| Export laporan | `xlsx` (SheetJS) — output `.xlsx` atau `.csv` |
| Kompresi respons | `compression` |
| Konfigurasi environment | `dotenv` |
| Frontend | Vanilla JS + HTML/CSS (satu file, tanpa build step), ikon `lucide` |
| Integrasi eksternal | Maxchat API (`https://app.maxchat.id`) |

Arsitektur: **single-file** — seluruh server (`app.js`) dan frontend (HTML/CSS/JS) berada dalam satu file yang sama, frontend disajikan langsung melalui route `GET /` (setelah login).

## Struktur Data

Disimpan di `storage/history.db` (SQLite), dibuat otomatis saat pertama kali dijalankan:

- `sessions` — satu baris per batch broadcast/balasan/sinkronisasi (termasuk `woCategory` per sesi).
- `contacts` — detail pesan per nomor dalam satu sesi.
- `read_state` — waktu terakhir kontak "dibaca" (untuk badge belum dibaca).
- `contact_category` — sumber kebenaran kategori WO (mis. Marking/Individu) per nomor (persisten, tidak ditimpa webhook).
- `contact_summary` — tabel ringkasan 1 baris per nomor (kondisi chat terbaru), dijaga tetap sinkron agar menu Chat/filter kategori/hitung unread tidak perlu scan seluruh riwayat.
- `contact_tags` — label yang ditempelkan ke tiap nomor kontak.
- `labels` — daftar master label (nama, warna, urutan).
- `contact_meta` — nama kustom kontak & penanda kontak yang ditambahkan manual.
- `quick_replies` — daftar balasan cepat beserta urutannya.
- `contacts_fts` — virtual table FTS5 (index pencarian nama/pesan/preview), disinkronkan otomatis lewat trigger pada tabel `contacts`.

Backup harian SQLite disimpan di `storage/backups/`.

Konfigurasi (`config.json`) dan daftar template (`templates.json`) disimpan sebagai file JSON terpisah di root project — **bukan** di database.

## Instalasi

### Prasyarat
- Node.js (disarankan versi LTS terbaru)
- Akun & API Token Maxchat aktif

### Langkah

```bash
# 1. Install dependency
npm install express axios better-sqlite3 multer compression express-session xlsx dotenv

# 2. (Opsional) buat file .env untuk konfigurasi awal — lihat bagian Environment Variables
cp .env.example .env   # jika tersedia, atau buat manual

# 3. Jalankan server
node app.js
```

Server berjalan di `http://localhost:<PORT>` (sesuaikan port di kode jika diperlukan). Folder `storage/` beserta database `history.db` dan folder `storage/backups/` akan dibuat otomatis pada saat pertama kali dijalankan.

## Environment Variables (`.env`)

Dibaca lewat `dotenv` saat konfigurasi awal (`config.json`) pertama kali dibuat — nilai default ini bisa diubah kapan saja lewat menu Pengaturan setelah aplikasi berjalan:

| Variabel | Fungsi | Default |
|---|---|---|
| `SESSION_SECRET` | Secret untuk menandatangani cookie session login | nilai bawaan di kode (⚠️ sebaiknya diganti lewat `.env`) |
| `DEFAULT_API_TOKEN` | API Token Maxchat awal | kosong |
| `DEFAULT_WEBHOOK_URL` | Webhook URL awal | kosong |
| `DEFAULT_LOGIN_USERNAME` | Username login awal | `admin` |
| `DEFAULT_LOGIN_PASSWORD` | Password login awal | `admin123` |
| `DEFAULT_SETTINGS_PASSWORD` | Password konfirmasi untuk pengaturan sensitif | kosong |

⚠️ **Wajib diganti sebelum dipakai di production**: `SESSION_SECRET`, `DEFAULT_LOGIN_USERNAME`/`DEFAULT_LOGIN_PASSWORD`, dan isi `DEFAULT_SETTINGS_PASSWORD`.

## Konfigurasi Awal

1. Buka aplikasi di browser → login dengan kredensial awal (lihat Environment Variables) → menu **Pengaturan**.
2. Isi **API Token Maxchat** (didapat dari dashboard Maxchat) dan **Webhook URL** — arahkan pengaturan webhook di dashboard Maxchat ke `https://<domain-server-anda>/webhook` agar pesan masuk, status pengiriman, dan sinkronisasi pesan keluar dapat diterima aplikasi. Perubahan ini butuh **password konfirmasi**.
3. Ganti username/password login serta password konfirmasi sesegera mungkin dari menu Pengaturan.

## Endpoint API (ringkas)

| Method | Endpoint | Fungsi |
|---|---|---|
| GET | `/login` · POST `/login` · POST `/logout` | Halaman & proses login/logout |
| GET | `/api/data` | Ambil data template |
| GET / POST | `/api/config` | Baca/simpan konfigurasi (API token, webhook URL) — POST butuh password konfirmasi |
| POST | `/api/login-credentials` | Ubah username/password login — butuh password konfirmasi |
| GET | `/api/system/db-info` | Info database (ukuran, jumlah sesi, status backup) |
| POST | `/api/system/backup-now` | Jalankan backup database manual |
| GET | `/api/dashboard/broadcast-history` | Rekap broadcast per hari/bulan |
| GET | `/api/dashboard/broadcast-history-export` | Export rekap broadcast (XLSX/CSV) |
| GET | `/api/dashboard/broadcast-export` | Export data broadcast mentah per rentang tanggal (XLSX/CSV) |
| GET | `/api/dashboard/frequent-reporters` | Daftar nomor yang paling sering melapor |
| GET | `/api/dashboard/wo-trend` | Tren jumlah WO per kategori dari waktu ke waktu |
| GET | `/api/history` | Ambil seluruh riwayat sesi |
| GET | `/api/history/count` | Jumlah total sesi |
| GET | `/api/history/latest-time` | Waktu sesi terbaru & jumlah unread |
| DELETE | `/api/history/:id` | Hapus satu sesi |
| DELETE | `/api/history` | Hapus semua riwayat |
| DELETE | `/api/history/range?start=&end=` | Hapus riwayat pada rentang tanggal tertentu |
| GET | `/api/contacts` | Daftar kontak chat (paginated, filter kategori/unread) |
| POST | `/api/contacts/:phone/read` | Tandai kontak sudah dibaca |
| GET | `/api/contacts/:phone/messages` | Riwayat pesan per kontak (paginated) |
| GET | `/api/contacts/search` | Cari kontak (FTS5, fallback `LIKE`) |
| GET | `/api/history/search` | Cari riwayat sesi (FTS5, fallback `LIKE`) |
| POST | `/api/blast` | Kirim broadcast template ke banyak nomor |
| GET | `/api/blast/progress/:id` | Polling progress broadcast yang sedang berjalan |
| POST | `/api/reply` | Kirim balasan teks bebas ke satu nomor |
| POST | `/api/reply-media-upload` | Upload file media untuk dilampirkan ke balasan |
| POST | `/api/reply-media` | Kirim balasan berupa media |
| POST / PUT / DELETE | `/api/templates`, `/api/templates/:id` | CRUD template pesan |
| POST | `/api/templates/reorder` | Urutkan ulang daftar template |
| GET / POST / PUT / DELETE | `/api/quick-replies`, `/api/quick-replies/:id` | CRUD balasan cepat |
| POST | `/api/quick-replies/reorder` | Urutkan ulang daftar balasan cepat |
| GET | `/api/tags` | Semua teks label yang pernah dipakai |
| GET / POST / PUT / DELETE | `/api/labels`, `/api/labels/:id` | CRUD master label (nama & warna) |
| GET / POST / DELETE | `/api/kontak/:phone/tags`, `/api/kontak/:phone/tags/:tag` | Tempel/lepas label pada satu kontak |
| PUT | `/api/kontak/:phone/name` | Ubah nama kontak |
| GET | `/api/kontak` | Daftar kontak (gabungan dari riwayat chat + kontak yang ditambah manual) |
| POST | `/api/kontak` | Tambah kontak manual |
| DELETE | `/api/kontak/:phone` | Hapus kontak manual (hanya jika belum punya riwayat) |
| POST | `/webhook` | Endpoint penerima event dari Maxchat (status, pesan keluar sync, pesan masuk) |

## Format Nomor Telepon

Semua nomor dinormalisasi otomatis ke format `628xxxxxxxxxx`:

| Input | Hasil normalisasi |
|---|---|
| `08123456789` | `628123456789` |
| `628123456789` | `628123456789` |
| `+628123456789` | `628123456789` |

## Catatan Operasional

- Semua route dilindungi login (session cookie), **kecuali** `/login` dan `/webhook` (webhook dipanggil server Maxchat, bukan browser).
- Perubahan pengaturan sensitif (API token, webhook URL, kredensial login) selalu meminta **password konfirmasi** terpisah dari password login — mencegah operator biasa iseng mengubah pengaturan walau sudah login.
- Pengiriman broadcast dilakukan **sekuensial** dengan jeda antar nomor (`DELAY_MS`) untuk menghindari pemblokiran oleh WhatsApp/Maxchat — jangan mengubah ini ke pengiriman paralel tanpa mempertimbangkan risiko rate-limit.
- Perubahan kategori kontak (WO Category) hanya terjadi saat broadcast baru dikirim ke nomor tersebut — webhook pesan masuk tidak pernah mengubah kategori.
- Label kontak sekarang berbasis **master label**: harus dibuat dulu lewat "Kelola Label" sebelum bisa ditempelkan ke kontak — tidak bisa lagi ketik teks bebas langsung di kontak.
- Kontak manual (ditambahkan lewat menu Kontak) hanya bisa dihapus selama **belum pernah** punya riwayat chat/broadcast; setelah punya riwayat, penghapusan hanya lewat Pengaturan → Hapus Riwayat.
- Penghapusan riwayat (semua atau berdasarkan rentang tanggal) bersifat **permanen** — tidak ada fitur undo/arsip di aplikasi (mitigasi: backup otomatis harian).
- Backup database otomatis berjalan tiap hari jam 04:00 WIB, hanya menyimpan sejumlah backup terbaru (yang lama otomatis dihapus) — cek jumlah & jadwalnya lewat menu Pengaturan atau endpoint `/api/system/db-info`.
- Pencarian kontak & riwayat memakai index FTS5 bila tersedia di build `better-sqlite3`; jika modul FTS5 tidak tersedia, aplikasi otomatis fallback ke pencarian `LIKE` biasa tanpa perlu konfigurasi tambahan.
- Deduplikasi pesan keluar menggunakan window waktu 15 detik berbasis kombinasi nomor + isi pesan (in-memory, akan reset jika server di-restart).

## Lisensi & Kepemilikan

Proyek dikembangkan oleh Command Center PLN UP2D Banten