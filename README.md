# Konverter & Rekap Laporan Insan Amanah

Aplikasi web dengan 2 menu:

1. **Konversi R-5401** — mengubah file laporan transaksi harian bank (R-5401, format
   teks lebar-tetap) menjadi file Excel rapi berisi sheet **Data Transaksi** dan
   **Ringkasan** (KPI, rekap per lokasi/channel, rekap per jam). Mendukung upload
   banyak file sekaligus, validasi otomatis terhadap total footer laporan, dan
   unduh gabungan.
2. **Rekap vs Master Siswa** — meng-join **laporan harian R-5401** (`.txt`) dengan
   **master siswa** (`.xlsx`: NO VA, NAMA, BPP, KEGIATAN, TABUNGAN) per No. Pelanggan
   = NO VA, menghitung Total Tagihan/Nilai Bayar/Selisih/Status per transaksi, lalu
   menghasilkan Excel yang **hanya berisi transaksi berstatus Sesuai**.

Dibangun **ringan** dengan Flask + openpyxl (tanpa numpy/pandas/pyarrow), memakai
HTTP request/response biasa — hemat memori dan stabil di container kecil.

## Fitur

**Konversi R-5401**
- Upload 1 atau banyak file `.txt` sekaligus.
- Parsing fixed-width + ekstraksi metadata (kode PT, nama, cabang, tanggal) otomatis.
- **Validasi**: total & jumlah transaksi hasil parsing dicocokkan dengan footer laporan.
- Deteksi duplikat (kode + tanggal sama) agar tidak dobel di file gabungan.
- Unduh per-file, unduh 1 Excel gabungan (+ sheet Rekap Harian), atau unduh `.zip` semua file.

**Rekap vs Master Siswa**
- Upload 1 file master siswa (`.xlsx`) + 1 file laporan harian R-5401 (`.txt`).
- Join per transaksi: No. Pelanggan (laporan) = NO VA (master); label kolom output
  Kegiatan/Tabungan/Lainnya mengikuti BPP/KEGIATAN/TABUNGAN dari master.
- Status per transaksi: `Sesuai` (Nilai Bayar = Total Tagihan), `Kurang`, atau `Lebih`.
- Unduh Excel (sheet Ringkasan + Transaksi) — sheet Transaksi **hanya memuat baris Sesuai**;
  Ringkasan tetap merangkum seluruh transaksi (termasuk Kurang/Lebih/tanpa data master).

## Menjalankan lokal
```bash
pip install -r requirements.txt

# cara cepat (server bawaan Flask, untuk development):
python app.py
# atau persis seperti produksi:
gunicorn app:app --workers 1 --threads 4 --bind 0.0.0.0:8501
```
Buka http://localhost:8501

## Struktur file
```
app.py            # aplikasi Flask (routing, upload, render hasil, endpoint unduh, 2 menu)
parser.py         # parsing R-5401 + pembuatan Excel (openpyxl)
validator.py      # parsing master siswa (xlsx), join ke laporan R-5401, rekap Excel (Sesuai saja)
requirements.txt  # Flask, openpyxl, gunicorn
Procfile          # start command untuk platform berbasis Procfile
railway.json      # start command untuk Railway
runtime.txt       # versi Python
```

## Deploy ke Railway

### Metode A — via GitHub (paling mudah)
1. Push folder ini ke sebuah repo GitHub.
2. Buka https://railway.app → **New Project** → **Deploy from GitHub repo** → pilih repo.
3. Railway mendeteksi Python (Nixpacks) dan menjalankan start command dari
   `railway.json` / `Procfile`:
   `gunicorn app:app --workers 1 --threads 4 --timeout 120 --bind 0.0.0.0:$PORT`
4. Setelah build selesai, buka **Settings → Networking → Generate Domain**
   untuk mendapatkan URL publik.

### Metode B — via Railway CLI
```bash
npm i -g @railway/cli
railway login
railway init
railway up
railway domain
```

### Catatan penting
- **`$PORT`**: Railway memberikan port lewat env var `$PORT`; start command sudah memakainya.
- **`--workers 1 --threads 4`**: 1 worker menjaga pemakaian memori rendah; beberapa
  thread cukup untuk melayani permintaan bersamaan. Hasil parsing disimpan sementara
  di memori proses (kedaluwarsa 30 menit), jadi tetap 1 worker agar link unduh selalu
  ketemu datanya.
- **Batas upload**: total 40 MB per sekali upload (diatur di `app.py` via
  `MAX_CONTENT_LENGTH`).
- Tidak perlu env var tambahan untuk fungsi dasar.

## Catatan data
Nama pelanggan & keterangan pada laporan sumber terpotong (field lebar-tetap
±16 karakter), jadi sebagian teks tidak lengkap. Nilai transaksi, tanggal, waktu,
dan lokasi akurat 100%.
