# Konverter Laporan R-5401 → Excel

Aplikasi Streamlit untuk mengubah file laporan transaksi harian bank (R-5401,
format teks lebar-tetap) menjadi file Excel rapi berisi sheet **Data Transaksi**
dan **Ringkasan** (KPI, rekap per lokasi/channel, rekap per jam). Mendukung
upload banyak file sekaligus, validasi otomatis terhadap total footer laporan,
dan unduh gabungan.

## Fitur
- Upload 1 atau banyak file `.txt` sekaligus.
- Parsing fixed-width + ekstraksi metadata (kode PT, nama, cabang, tanggal) otomatis.
- **Validasi**: total & jumlah transaksi hasil parsing dicocokkan dengan footer laporan.
- Deteksi duplikat (kode + tanggal sama) agar tidak dobel di file gabungan.
- Unduh per-file, unduh 1 Excel gabungan (+ sheet Rekap Harian), atau unduh `.zip` semua file.

## Menjalankan lokal
```bash
pip install -r requirements.txt
streamlit run app.py
```
Buka http://localhost:8501

## Deploy ke Railway

### Struktur file (wajib ada semua)
```
app.py
parser.py
requirements.txt
Procfile
railway.json
runtime.txt
.streamlit/config.toml
```

### Metode A — via GitHub (paling mudah)
1. Push folder ini ke sebuah repo GitHub.
2. Buka https://railway.app → **New Project** → **Deploy from GitHub repo** → pilih repo.
3. Railway otomatis mendeteksi Python (Nixpacks) dan menjalankan start command dari
   `Procfile` / `railway.json`:
   `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
4. Setelah build selesai, buka tab **Settings → Networking → Generate Domain**
   untuk mendapatkan URL publik.

### Metode B — via Railway CLI
```bash
npm i -g @railway/cli
railway login
railway init          # buat project baru
railway up            # deploy folder ini
railway domain        # generate URL publik
```

### Catatan penting
- **`$PORT`**: Railway memberikan port lewat env var `$PORT`. Start command sudah
  memakainya — jangan hard-code port.
- **`--server.address 0.0.0.0`**: wajib agar app bisa diakses dari luar container.
- Tidak perlu env var tambahan untuk fungsi dasar.
- Batas ukuran upload diatur di `.streamlit/config.toml` (`maxUploadSize`, default 20 MB).

## Catatan data
Nama pelanggan & keterangan pada laporan sumber terpotong (field lebar-tetap
±16 karakter), jadi sebagian teks tidak lengkap. Nilai transaksi, tanggal, waktu,
dan lokasi akurat 100%.
