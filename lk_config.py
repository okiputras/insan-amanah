"""
Konfigurasi & helper Laporan Keuangan SD Insan Amanah (dipakai menu Laporan
Keuangan di app.py).

Tata letak tiap tab Google Sheet ("<BULAN> <TAHUN>") dibuat MENGIKUTI PERSIS
file contoh "LAPORAN KEUANGAN BULAN DESEMBER 2023.xls" (sheet 'laporan
keuangan'), kecuali kolom Sumber Anggaran (RENCANA/SPP/BSM/BOSDA/BOSNAS) yang
sengaja tidak dipakai.

  Baris 1-3 : Judul (judul besar / tahun pelajaran / bulan), merge B..Q
  Baris 4   : kosong (pemisah, seperti di file asli)
  Baris 5-6 : Header 2 tingkat — grup "RINCIAN" menaungi VOLUME..TTL/SUB
  Baris 7+  : baris outline + baris subtotal "Jumlah Biaya" (otomatis)

Kolom (1-based):
  A  spacer sempit          B  NO (nomor Program)
  C  no Sub Program / nama Program
  D  nama Sub Program       E  no Kegiatan
  F  huruf Item / nama Kegiatan / teks Rincian (di bawah Kegiatan)
  G  nama Item / teks Rincian (di bawah Item)
  H  kolom lebar tempat teks F/G meluber (kosong, seperti aslinya)
  I  TANGGAL   J  KODE/KW   K  VOLUME   L  SATUAN   M  FK
  N  KEBUTUHAN O  UNIT COST P  TOTAL    Q  TTL/SUB  R  LEVEL (disembunyikan)

INI PERBEDAAN UTAMA vs versi sebelumnya: tidak ada satu kolom "LABEL". Teks
nama berpindah kolom sesuai LEVEL-nya (Program di C, Sub Program di D,
Kegiatan di F, Item di G) — itulah yang membentuk tampilan bertingkat seperti
di file asli. Nomor pun tersebar: Program di B, Sub Program di C, Kegiatan di
E, huruf Item di F.

LEVEL tetap disimpan di kolom R (disembunyikan) sebagai penanda internal.
Sebetulnya level bisa ditebak dari kolom mana yang terisi, tapi tebakan itu
rapuh (mis. teks Rincian yang kebetulan cuma satu huruf akan terbaca sebagai
huruf Item), jadi levelnya disimpan eksplisit. Kolom R tersembunyi sehingga
tampilan tetap identik dengan aslinya.

Nomor (B/C/E/F), teks "Jumlah Biaya", dan TTL/SUB semuanya DIHITUNG OTOMATIS
oleh resync() di lk_sheet.py tiap kali ada baris ditambah/diedit/dihapus —
tidak pernah diketik manual.
"""
from tab_config import MONTHS_ID  # reuse, jangan duplikat

SPREADSHEET_TITLE = "LAPORAN KEUANGAN SD INSAN AMANAH"
SPREADSHEET_ID = "1tVrKbIfxuVFzesZwh2p9a1xj3TTj872YwYpd6-uG_OU"
SERVICE_ACCOUNT_FILE = "sa-sheet.json"          # reuse via tab_sheet
SERVICE_ACCOUNT_EMAIL = "oki-gsheet@iconic-woods-355603.iam.gserviceaccount.com"

LEVEL_LABELS = ["Program", "Sub Program", "Kegiatan", "Item", "Rincian"]
LEVEL_SUBTOTAL = 6       # internal: baris "Jumlah Biaya", dikelola resync()
SUBTOTAL_LABEL = "Jumlah Biaya"

# ---- posisi baris (1-based) ----
TITLE_ROW1, TITLE_ROW2, TITLE_ROW3 = 1, 2, 3
BLANK_ROW = 4
HEADER_ROW1, HEADER_ROW2 = 5, 6
FIRST_DATA_ROW = 7

# ---- posisi kolom (1-based) ----
COL_SPACER = 1     # A
COL_PROG_NO = 2    # B  nomor Program
COL_PROG_NAME = 3  # C  nama Program  (kolom yang sama dipakai nomor Sub Program)
COL_SUB_NO = 3     # C
COL_SUB_NAME = 4   # D  (juga dipakai teks "Jumlah Biaya")
COL_KEG_NO = 5     # E
COL_KEG_NAME = 6   # F  (kolom yang sama dipakai huruf Item & teks Rincian)
COL_ITEM_LETTER = 6  # F
COL_ITEM_NAME = 7  # G
COL_WIDE = 8       # H  sengaja dikosongkan, tempat teks meluber
COL_TANGGAL = 9    # I
COL_KODE = 10      # J
COL_VOLUME = 11    # K
COL_SATUAN = 12    # L
COL_FK = 13        # M
COL_KEBUTUHAN = 14 # N
COL_UNIT_COST = 15 # O
COL_TOTAL = 16     # P
COL_TTL_SUB = 17   # Q
COL_LEVEL = 18     # R  (disembunyikan)

N_COLS = COL_LEVEL

# Kolom finansial yang diisi dari form tambah/edit -> nomor kolomnya.
FIN_COLS = {
    "tanggal": COL_TANGGAL, "kode": COL_KODE, "volume": COL_VOLUME,
    "satuan": COL_SATUAN, "fk": COL_FK, "kebutuhan": COL_KEBUTUHAN,
    "unit_cost": COL_UNIT_COST, "total": COL_TOTAL,
}
FIELD_KEYS = ["level", "label"] + list(FIN_COLS)

# Kolom yang diperlakukan sebagai angka (untuk format sel)
NUMERIC_COLS = [COL_VOLUME, COL_UNIT_COST, COL_TOTAL, COL_TTL_SUB]

# Header 2 tingkat. Tiap entri: (teks, kolom_awal, kolom_akhir, baris_awal, baris_akhir)
HEADER_CELLS = [
    ("NO",                       COL_PROG_NO,  COL_PROG_NO,  HEADER_ROW1, HEADER_ROW2),
    ("PROGRAM PENINGKATAN MUTU", COL_SUB_NO,   COL_SUB_NAME, HEADER_ROW1, HEADER_ROW2),
    ("RINCIAN KEGIATAN",         COL_KEG_NO,   COL_WIDE,     HEADER_ROW1, HEADER_ROW2),
    ("TANGGAL",                  COL_TANGGAL,  COL_TANGGAL,  HEADER_ROW1, HEADER_ROW2),
    ("KODE",                     COL_KODE,     COL_KODE,     HEADER_ROW1, HEADER_ROW1),
    ("KW",                       COL_KODE,     COL_KODE,     HEADER_ROW2, HEADER_ROW2),
    ("RINCIAN",                  COL_VOLUME,   COL_TTL_SUB,  HEADER_ROW1, HEADER_ROW1),
    ("VOLUME",                   COL_VOLUME,   COL_VOLUME,   HEADER_ROW2, HEADER_ROW2),
    ("SATUAN",                   COL_SATUAN,   COL_SATUAN,   HEADER_ROW2, HEADER_ROW2),
    ("FK",                       COL_FK,       COL_FK,       HEADER_ROW2, HEADER_ROW2),
    ("KEBUTUHAN",                COL_KEBUTUHAN, COL_KEBUTUHAN, HEADER_ROW2, HEADER_ROW2),
    ("UNIT COST",                COL_UNIT_COST, COL_UNIT_COST, HEADER_ROW2, HEADER_ROW2),
    ("TOTAL",                    COL_TOTAL,    COL_TOTAL,    HEADER_ROW2, HEADER_ROW2),
    ("TTL/SUB",                  COL_TTL_SUB,  COL_TTL_SUB,  HEADER_ROW2, HEADER_ROW2),
]

# Lebar kolom (px) — dikonversi dari lebar karakter di file asli (px ≈ char*7+5)
COL_WIDTHS_PX = {
    COL_SPACER: 9, COL_PROG_NO: 41, COL_SUB_NO: 45, COL_SUB_NAME: 172,
    COL_KEG_NO: 34, COL_KEG_NAME: 37, COL_ITEM_NAME: 37, COL_WIDE: 312,
    COL_TANGGAL: 95, COL_KODE: 69, COL_VOLUME: 65, COL_SATUAN: 84,
    COL_FK: 52, COL_KEBUTUHAN: 93, COL_UNIT_COST: 103, COL_TOTAL: 120,
    COL_TTL_SUB: 113,
}

# Font per peran — diambil persis dari file contoh asli.
FONT_TITLE = ("Algerian", 20)
FONT_HEADER = ("Bodoni MT", 12)
FONT_PROGRAM = ("Arial Black", 12)
FONT_BODY = ("Calibri", 12)


def tab_name(bulan_num, tahun):
    return f"{MONTHS_ID[bulan_num - 1]} {tahun}"


def academic_label(bulan_num, tahun):
    start = tahun if bulan_num >= 7 else tahun - 1
    return f"{start}/{start + 1}"


def label_col(level, parent_level=None):
    """Kolom tempat TEKS nama ditulis untuk sebuah level.
    Rincian (5) mengikuti kolom teks induknya: di bawah Item -> G, selain itu -> F
    (persis seperti di file asli)."""
    if level == 1:
        return COL_PROG_NAME
    if level == 2:
        return COL_SUB_NAME
    if level == 3:
        return COL_KEG_NAME
    if level == 4:
        return COL_ITEM_NAME
    if level == LEVEL_SUBTOTAL:
        return COL_SUB_NAME
    return COL_ITEM_NAME if parent_level == 4 else COL_KEG_NAME


# Semua kolom yang bisa memuat teks nama (dipakai saat membaca baris).
LABEL_COLS = [COL_PROG_NAME, COL_SUB_NAME, COL_KEG_NAME, COL_ITEM_NAME]


def number_col(level):
    """Kolom tempat NOMOR/HURUF ditulis; None kalau level itu tidak bernomor."""
    return {1: COL_PROG_NO, 2: COL_SUB_NO, 3: COL_KEG_NO, 4: COL_ITEM_LETTER}.get(level)


def indent_px(level):
    return max(0, (int(level) - 1)) * 22


# Template default (Program/Sub Program/Kegiatan/Item, tanpa tanggal & nominal)
# — diambil dari struktur "LAPORAN KEUANGAN BULAN DESEMBER 2023.xls" (SD Insan
# Amanah, 8 Standar Nasional Pendidikan). Dipakai untuk mengisi tab bulan
# pertama saat belum ada bulan sebelumnya untuk di-clone.

DEFAULT_TEMPLATE = [
    (1, 'Pemenuhan dan Peningkatan Mutu Standar Isi'),
    (2, 'Rapat Kerja (Raker)'),
    (3, 'Rapat Kerja SD Insan Amanah'),
    (2, 'Pengembangan Kurikulum'),
    (3, 'Buku Siswa'),
    (3, 'Konsultan Pendidikan SDIA'),
    (1, 'Pemenuhan dan Peningkatan Mutu Standar Proses'),
    (2, 'Sumbangan'),
    (4, 'Menjenguk siswa sakit'),
    (4, 'Sumbangan Sosial'),
    (4, 'Bantuan pengobatan kecelakaan siswa'),
    (4, 'Uang Duka'),
    (2, 'Kegiatan UKS'),
    (4, 'Obat - obatan & peralatan'),
    (4, 'Kebutuhan UKS'),
    (4, 'Insentif & Konsumsi'),
    (4, 'Tindakan'),
    (4, 'Iuran UKS'),
    (2, 'Pembentukan Karakter'),
    (4, 'Pelatihan & Learning Motivation Kelas 5 dan 6'),
    (4, 'Pondok UN'),
    (4, 'Motivasi menjelang US kelas 6'),
    (4, 'Learning by parenting (siswa)'),
    (4, 'Parenting (orang tua siswa)'),
    (2, 'Aplikasi Pembelajaran Siswa'),
    (3, 'Try Out'),
    (4, 'Try Out kelas 6'),
    (4, 'Try Out kelas 5'),
    (4, 'Try Out kelas 4'),
    (3, 'Fieldtrip'),
    (3, 'Implementasi Pembelajaran Tematik'),
    (3, 'Kegiatan Outing'),
    (3, 'Implementasi Kurikulum Khas UMMI'),
    (4, 'Dana operaional munaqosah'),
    (4, 'Konsumsi makanan untuk persiapan munaqosah'),
    (4, 'Biaya sewa kursi, mic dan sound'),
    (2, 'Penunjang Pembelajaran'),
    (3, 'Pembuatan Modul & Kumpulan Soal'),
    (4, 'Kelas 6'),
    (4, 'Kelas 5'),
    (4, 'Kelas 4'),
    (3, 'Pembuatan Bank Soal Kelas 1 - 6'),
    (4, 'Biaya pembuatan bank soal Semester 1'),
    (4, 'Biaya pembuatan bank soal Semester 2'),
    (3, 'Pembuatan Poin Prestasi'),
    (4, 'Stiker poin'),
    (4, 'Buku poin prestasi'),
    (4, 'Lembar poin'),
    (4, 'Hadiah poin terbanyak'),
    (3, 'Buku Pegangan untuk siswa dan guru'),
    (4, 'Buku Paket Pegangan Guru'),
    (4, 'Buku belajar bina sholat dan Alquran'),
    (4, 'Buku pegangan bina sholat dan Al Quran'),
    (4, 'Buku LK Kurikulum 2013 kls 1-6  (tematik)'),
    (3, 'Pengelolaan Perpustakaan'),
    (4, 'Pembelian Buku dan Inv. Perpus'),
    (4, 'Insentif Pustakawan Perpus Kota'),
    (4, 'Konsumsi'),
    (4, 'Majalah'),
    (4, 'Koran'),
    (3, 'Media Pembelajaran'),
    (4, 'Olah raga'),
    (4, 'Perkemahan Senin Minggu (PERSAMI)'),
    (4, 'Media Pembelajaran'),
    (4, 'Hosting website SD Insan Amanah'),
    (3, 'Administrasi Sekolah'),
    (4, 'Fotokopi administrasi sekolah'),
    (4, 'Jilid administrasi sekolah'),
    (4, 'Print administrasi sekolah'),
    (3, 'Studi Banding'),
    (1, 'Pemenuhan dan Peningkatan Mutu Standar Kompetensi Lulusan'),
    (2, 'Penambahan Jam Belajar Siswa'),
    (3, 'Jam Sore Kelas 4, 5 dan 6'),
    (4, 'Insentif pengajar jam sore kelas 4'),
    (4, 'Insentif pengajar jam sore kelas 5'),
    (4, 'Insentif pengajar jam sore kelas 6'),
    (4, 'Insentif keamanan jam sore'),
    (3, 'Remidi'),
    (3, 'Penambahan Jam Belajar Siswa'),
    (2, 'Hari Besar Islam'),
    (3, 'Idul Adha'),
    (4, 'Operasional Idul Adha'),
    (3, 'Manasik Haji'),
    (4, 'Operasional Manasik haji'),
    (3, "Isro' Mi'roj"),
    (4, "Operasional Isro' Mi'roj"),
    (3, '1 Muharam'),
    (4, 'Insentif pendongeng peringatan 1 Muharram'),
    (3, 'Maulid Nabi Muhammad'),
    (4, 'Operasional Maulid Nabi Muhammad'),
    (3, 'Wisata Hati'),
    (4, 'Insentif penceramah'),
    (3, 'Kegiatan Ramadhan dan pasca'),
    (4, 'Buka bersama'),
    (3, 'Halal Bihalal'),
    (4, 'Sewa perlengkapan acara HBH'),
    (2, 'Hari Besar Nasional'),
    (3, 'Hari Anak Nasional'),
    (4, 'Cetak banner Hari Anak Nasional'),
    (4, 'Biaya pembuatan sertifikat HAN'),
    (3, 'Hari Kemerdekaan'),
    (4, 'Operasional HUT RI ke 78 tahun'),
    (2, 'Uji Kompetensi Siswa'),
    (3, 'Ujian Siswa Kelas 1 sampai 6'),
    (4, 'Penilaian Akhir Semester (PAS)'),
    (4, 'Sumatif Tengah Semester (PTS)'),
    (3, 'Ujian Kelas 6'),
    (4, 'Insentif pengawas UN'),
    (4, 'Ujian Nasional Kelas 6 gugus'),
    (4, 'Ujian Praktek kelas 6'),
    (3, 'Kegiatan wisuda kelas 6'),
    (4, 'Publikasi acara pisah kenang'),
    (4, 'Publikasi SD Insan Amanah'),
    (4, 'Foto ijasah'),
    (3, 'Lain-lain'),
    (4, 'Cetak kalender SDIA'),
    (4, 'Cetak majalah insaniyah'),
    (1, 'Pemenuhan dan Peningkatan Mutu Standar Pendidik dan Tenaga Pendidikan'),
    (2, 'Pengembangan Diri Pegawai'),
    (3, 'Pelatihan Kurikulum & Pembelajaran'),
    (4, 'Guru & Karyawan'),
    (4, 'Insentif pelatihan HOTS'),
    (4, 'In On In'),
    (4, 'Pelatihan pembuatan kisi-kisi'),
    (3, 'Pengembangan Individu'),
    (4, 'Workshop'),
    (4, 'Kursus'),
    (4, 'Seminar'),
    (4, 'Magang'),
    (3, 'Pembinaan KKG/K3S'),
    (4, 'KKG Agama'),
    (4, 'Iuran KKG PAI'),
    (3, 'Gugus'),
    (4, 'Kegiatan Iuran Gugus V bulan Juli 2023'),
    (4, 'Kegiatan Raker Kepala Sekolah Gugus'),
    (4, 'Kegiatan Workshop Kurikulum 13'),
    (4, 'Gugus KKG'),
    (4, 'BIMTEK BOS'),
    (4, 'Iuran PGRI bulan januari 2023 - juni 2023'),
    (3, 'Kecamatan'),
    (4, 'Iuran kegiatan HUT PGRi Kecamatan'),
    (4, 'Kegiatan K3S, KKG Kecamatan Lowokwaru'),
    (4, 'Kegiatan sosialisasi PPK Kec. Lowokwaru'),
    (4, 'PHBI Kecamatan (Halal Bi Halal)'),
    (3, 'Insentif pengawas'),
    (3, 'Supervisi'),
    (3, 'Operasional pembelajaran UMMI'),
    (4, 'Insentif guru UMMI'),
    (4, 'Insetif supervisi YDSF'),
    (4, 'Peraga UMMI'),
    (4, 'Seragam guru UMMI'),
    (4, 'Konsumsi untuk kegiatan'),
    (3, 'Opreasional kegiatan Lomba Eksternal'),
    (4, 'Lomba siswa'),
    (4, 'Lomba guru'),
    (1, 'Pemenuhan dan Peningkatan Mutu Standar Sarana Prasarana'),
    (2, 'ATK habis pakai'),
    (3, 'Pembelian ATK Habis Pakai (HP)'),
    (4, 'ATK HP plastik label'),
    (4, 'ATK HP tinta refill'),
    (4, 'ATK HP kertas'),
    (4, 'ATK HP kertas'),
    (4, 'ATK HP Drum Kit fotokopi, Blade, cip mesin fotokopi'),
    (4, 'ATK HP dvd RW'),
    (4, 'ATK HP karet'),
    (4, 'ATK HP Kater, penggaris, gunting, preparator'),
    (4, 'ATK HP amplop'),
    (4, 'ATK HP kwitansi'),
    (4, 'ATK HP map L'),
    (4, 'ATK HP plastik mika'),
    (4, 'ATK HP spidol'),
    (4, 'ATK HP solasi'),
    (4, 'ATK HP baterai AAA'),
    (4, 'ATK HP buku besar'),
    (4, 'ATK cutter besar'),
    (3, 'Pembelian Inventaris Habis Pakai Bagian IT'),
    (2, 'Pembelian Inventaris'),
    (3, 'Pembelian Inventaris HP Administrasi'),
    (4, 'ATK Inventaris Admin Keuangan'),
    (4, 'ATK Inventaris Admin Akademik'),
    (3, 'Pembelian Inventaris HP Sarpras perlengkapan sekolah'),
    (4, 'Inventaris Sarpras Perlengkapan Sekolah'),
    (2, 'Kebersihan'),
    (3, 'Perlengkapan kebersihan'),
    (4, 'Pembersih lantai'),
    (4, 'Sabun'),
    (4, 'Tissue'),
    (4, 'Sapu'),
    (4, 'Sabun cuci piring'),
    (4, 'Sabun cuci tangan'),
    (4, 'Obat - obatan serangga'),
    (4, 'Pengharum ruangan'),
    (4, 'Lain-lain'),
    (3, 'Peralatan kebersihan'),
    (2, 'Laundry'),
    (3, 'Laundry'),
    (2, 'Biaya Jasa Bulanan'),
    (4, 'Biaya Listrik'),
    (4, 'Telepon'),
    (4, 'Speedy'),
    (4, 'PDAM'),
    (4, 'Kebersihan'),
    (4, 'Air Mineral dan Galon'),
    (4, 'Gas LPG'),
    (4, 'Gula, Teh, dan Kopi'),
    (4, 'Biaya jasa, parkir, tol dan pengiriman'),
    (4, 'Bahan Bakar Minyak'),
    (4, 'Biaya servis'),
    (4, 'Pajak Kendaraan'),
    (4, 'Operasional taman,kolam dan penghijauan'),
    (4, 'Konsumsi'),
    (4, 'Perlengkapan Elektronik'),
    (4, 'Perlengkapan Kebutuhan Sekolah'),
    (4, 'Perlengkapan Pertukangan'),
    (4, 'Perlengkapan lain-lain'),
    (2, 'Iventaris Sekolah'),
    (3, 'Pembelian Inventaris IT'),
    (3, 'Pembelian inventaris kelas dan gedung'),
    (1, 'Pemenuhan dan Peningkatan Mutu Standar Pengelolaan'),
    (2, 'Kegiatan Ramadhan'),
    (3, 'Parcel Kolega'),
    (4, 'Bahan parcel'),
    (4, 'Publikasi acara pentasyarufan zakat'),
    (3, 'b.THR Pegawai,Pengurus,Pembina UMMI & Ekstra'),
    (4, 'Tali Asih Pegawai Tetap dan kontrak'),
    (4, 'Tali Asih Pembina ekstra'),
    (4, 'Tali Asih Guru UMMI'),
    (4, 'Tali Asih Pengurus'),
    (1, 'Pemenuhan dan Peningkatan Mutu Standar Biaya'),
    (2, 'Biaya Operasi Sekolah'),
    (3, 'Gaji pegawai tetap & kontrak lembaga'),
    (4, 'Gaji pegawai tetap & kontrak'),
    (3, 'Insentif Pegawai Kontrak'),
    (4, 'Insentif Guru UMMI'),
    (4, 'Insentif Up grading UMMI'),
    (4, 'Insentif supervisi dan penilaian try out'),
    (4, 'Tali Asih untuk guru keluar'),
    (4, 'Tali Asih Lebaran'),
    (3, 'Insentif Pegawai Lepas'),
    (3, 'Insentif Pegawai Percobaan'),
    (4, 'Insentif guru /karyawan Percobaan'),
    (3, 'Insentif Pembina Ekstra'),
    (4, 'Insentif Pembina Ekstra'),
    (3, 'Insentif Proses Pembelajaran'),
    (4, 'Kelebihan jam mengajar'),
    (4, 'Penulisan rapor'),
    (3, 'Insentif Kegiatan diluar pembelajaran'),
    (4, 'Insentif Membuat Minum'),
    (4, 'Insentif membersihkan gedung'),
    (4, 'Insentif membuang sampah Juli'),
    (4, 'Insentif membersihkan masjid'),
    (4, 'Insentif merawat tanaman'),
    (4, 'Insentif membersihkan rumput belakang sekolah'),
    (4, 'Insentif pegawai kebersihan outsourcing'),
    (4, 'Insentif pemotong rumput'),
    (3, 'Lembur Administrasi dan lembur kegiatan'),
    (4, 'Keg. Adiwiyata, GSF, Promosi, Administrasi'),
    (4, 'Keg. Pembelajaran'),
    (4, 'Keg. Pendukung Pembelajaran dll'),
    (3, 'Insentif Jam Sore'),
    (4, 'Kelas 4'),
    (4, 'Kelas 5'),
    (4, 'Kelas 6'),
    (4, 'Petugas keamanan'),
    (3, 'Insentif persiapan kegiatan kelas 6'),
    (4, 'Persiapan kegiatan Kelas 6'),
    (4, 'Operasional Penulisan Ijazah'),
    (4, 'Pemberkasan Nilai Kelas 6'),
    (4, 'Pengumuman Kelulusan'),
    (4, 'Piket Sekolah dan Legalisir'),
    (3, 'Insentif transportasi kegiatan'),
    (4, 'Bank, belanja, pesan spanduk, antar surat dll'),
    (4, 'Lomba'),
    (3, 'Publikasi dan Kehumasan'),
    (4, 'Media Cetak Koran'),
    (4, 'Media elektronik'),
    (4, 'Baliho / spanduk'),
    (4, 'Backdrop acara di sekolah'),
    (4, 'Spanduk , Banner, Poster'),
    (4, 'Map Sekolah'),
    (4, 'Transport wartawan'),
    (3, 'Insentif Piket'),
    (4, 'Piket Sekolah'),
    (4, 'Piket IED'),
    (2, 'Operasional Lembaga'),
    (3, 'Biaya operasional lembaga rutinitas'),
    (4, 'Insentif Pengurus LPI bulan Agustus 2023'),
    (4, 'Iuran RW dan RT'),
    (4, 'Iuran FORTUSIA'),
    (4, 'BPJS Pak Hamdan'),
    (3, 'Cadangan'),
    (3, 'Biaya operasional lembaga berkala'),
    (4, 'Insentif tali asih idul fitri pengurus'),
    (4, 'Bonus PPDB pimpinan'),
    (4, 'Hadiah karyawan berprestasi'),
    (4, 'Peningkatan pendidikan'),
    (4, 'Perbaikan gedung sekolah'),
    (4, 'Biaya Operasional Teacher Gathering'),
    (4, 'Pembayaran PBB'),
    (4, 'Biaya Seragam'),
    (1, 'Pemenuhan dan Peningkatan Mutu Standar Penilaian'),
    (2, 'Remidi /Pengayaan'),
    (3, 'Remidi'),
    (4, 'Remidi Kelas 1'),
    (4, 'Remidi Kelas 2'),
    (4, 'Remidi Kelas 3'),
    (4, 'Remidi Kelas 4'),
    (4, 'Remidi Kelas 5'),
    (2, 'Pemberitahuan Hasil belajar'),
    (3, 'Penerimaan Rapor'),
    (4, 'Kertas rapor sisipan dan UAS I'),
    (4, 'Buletin I'),
    (4, 'Kertas rapor sisipan dan UAS II'),
    (4, 'Buletin II'),
    (4, 'Lembar Kerja'),
    (4, 'Rapor Rekam Jejak dari Gugus'),
    (2, 'Immersion Program'),
    (3, 'Operasional Kegiatan Immersion Study'),
    (2, 'Perbaikan Bangunan Sekolah'),
    (2, 'Lomba Nasional'),
    (2, 'lain - lain'),
    (3, 'Pemasangan instalasi hidroponik'),
]
