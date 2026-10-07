"""
Konfigurasi & logika tagihan menu Pembayaran SD (dipakai app.py & bayar_import.py).

Murni — tidak menyentuh Google Sheets — supaya aturan tagihan bisa diuji tanpa jaringan.

MODEL: buku besar + alokasi FIFO dua kantong.
Setiap pembayaran disimpan sebagai satu baris (tab PEMBAYARAN, hanya ditambah). Status
lunas TIDAK disimpan; dihitung ulang tiap kali dicek: total yang pernah dibayar seorang
siswa dialokasikan ke bulan tertua lebih dulu, terpisah untuk dua kantong —
  "bk"  = BPP + katering (digabung)
  "keg" = kegiatan
Sebuah bulan lunas bila kedua kantong tertutup, dengan kekurangan <= TOLERANSI diabaikan.

Kenapa saldo, bukan "di blok bulan mana pembayaran tercatat": banyak orang tua membayar
beberapa bulan sekaligus (2x, 6x, 12x), membayar tunggakan, atau membayar di muka.
Label bulan di Excel sumber pun tidak bisa dipercaya (ada blok tanggal 14 Agustus
berlabel "JUNI"). Dengan saldo, semua kasus itu benar tanpa membaca label apa pun.

Kenapa BPP dan katering digabung: di format Excel 9 kolom, "BPP murni" adalah rumus
(BPP+katering) - katering. Saat orang tua membayar 2 bulan tapi katering diketik sekali,
pembagiannya keliru (1.800.000 terbagi 1.600.000 + 200.000, seharusnya 1.400.000 +
400.000), sedangkan jumlah gabungannya selalu benar. Syarat lunas tetap mencakup
BPP + katering + kegiatan.
"""
import datetime as dt

from tab_config import MONTHS_ID, months_for_year

SPREADSHEET_TITLE = "ALARM PEMBAYARAN SD INSAN AMANAH"
SPREADSHEET_ID = "1ez1s2UvFSj5liyBq9V-adN7AbBGyZwQVq1XNU8s5emI"
TAHUN_AJARAN = 2026           # T.A. 2026/2027: Juli 2026 .. Juni 2027
TOLERANSI = 10_000            # kekurangan sekecil ini dianggap lunas (potongan bayar-kegiatan-setahun)

TAB_SISWA = f"SISWA {TAHUN_AJARAN}"
TAB_BAYAR = "PEMBAYARAN"
TAB_REVIEW = "PERLU DICEK"

HDR_SISWA = ["NO", "INDUK", "NAMA", "ROMBEL", "BPP", "KATERING", "KEGIATAN", "STATUS", "CATATAN"]
HDR_BAYAR = ["TANGGAL", "INDUK", "NAMA", "BPP", "KATERING", "KEGIATAN", "TABUNGAN",
             "SUMBER", "KUNCI", "CATATAN"]
HDR_REVIEW = ["TANGGAL", "NAMA TERCATAT", "INDUK TERCATAT", "BPP", "KATERING", "KEGIATAN",
              "TABUNGAN", "MASALAH", "SARAN INDUK", "KUNCI"]

STATUS_AKTIF, STATUS_PINDAH = "AKTIF", "PINDAH"
KANTONG = ("bk", "keg")
LABEL_KANTONG = {"bk": "BPP+katering", "keg": "kegiatan"}

# Aturan khusus T.A. 2026/2027 untuk siswa baru (kelas 1) di bulan Juli: BPP & katering
# Juli sudah dibayar saat daftar ulang/PPDB (di luar buku ini), sedangkan kegiatan Juli
# memakai tarif khusus. Dikonfirmasi pihak sekolah.
KELAS1_JULI_KEGIATAN = 114_500


def bulan_list():
    """Urutan bulan tahun ajaran: [7, 8, ..., 12, 1, ..., 6]."""
    return months_for_year(TAHUN_AJARAN)


def label_bulan(bulan, pendek=False):
    nama = MONTHS_ID[bulan - 1]
    if pendek:
        return nama[:3].title()
    tahun = TAHUN_AJARAN if bulan >= 7 else TAHUN_AJARAN + 1
    return f"{nama.title()} {tahun}"


def bulan_ini(now=None):
    """Bulan berjalan menurut WIB (UTC+7), dijepit ke rentang tahun ajaran.

    Sengaja tidak memakai jam server: server Railway berjalan dalam UTC, jadi tanggal 1
    pukul 00.00-06.59 WIB masih terbaca bulan sebelumnya."""
    now = now or (dt.datetime.utcnow() + dt.timedelta(hours=7))
    awal = dt.date(TAHUN_AJARAN, 7, 1)
    akhir = dt.date(TAHUN_AJARAN + 1, 6, 30)
    hari = now.date() if isinstance(now, dt.datetime) else now
    if hari < awal:
        return 7
    if hari > akhir:
        return 6
    return hari.month


def tagihan(siswa, bulan):
    """Tagihan satu siswa untuk satu bulan -> {"bk": int, "keg": int}."""
    bk = int(siswa.get("bpp") or 0) + int(siswa.get("katering") or 0)
    keg = int(siswa.get("kegiatan") or 0)
    if str(siswa.get("rombel", ""))[:1] == "1" and bulan == 7:
        bk, keg = 0, KELAS1_JULI_KEGIATAN
    return {"bk": bk, "keg": keg}


def status_siswa(siswa, bayar, sampai):
    """Alokasi FIFO per kantong dari Juli sampai bulan `sampai` (inklusif).

    bayar: {"bk": total dibayar, "keg": total dibayar}
    -> list[(bulan, {kantong: kekurangan})]; dict kosong = bulan itu lunas.

    Toleransi diterapkan per bulan per kantong, tapi sisa dana tetap mengalir FIFO, jadi
    kekurangan kecil yang berulang tetap menumpuk dan akhirnya muncul — toleransi hanya
    menyerap selisih kecil total, bukan kurang-bayar tiap bulan."""
    urutan = bulan_list()
    sisa = {k: int(bayar.get(k) or 0) for k in KANTONG}
    out = []
    for b in urutan[: urutan.index(sampai) + 1]:
        tg = tagihan(siswa, b)
        kurang = {}
        for k in KANTONG:
            pakai = min(sisa[k], tg[k])
            sisa[k] -= pakai
            if tg[k] - pakai > TOLERANSI:
                kurang[k] = tg[k] - pakai
        out.append((b, kurang))
    return out


def total_per_siswa(pembayaran):
    """Daftar baris pembayaran -> {induk: {"bk": total, "keg": total}}."""
    tot = {}
    for p in pembayaran:
        t = tot.setdefault(p["induk"], {"bk": 0, "keg": 0})
        t["bk"] += int(p.get("bpp") or 0) + int(p.get("katering") or 0)
        t["keg"] += int(p.get("kegiatan") or 0)
    return tot


def ringkas_kurang(st):
    """[(bulan, {kantong: kurang})] -> 'Agu: kegiatan · Sep: BPP+katering, kegiatan'."""
    bagian = []
    for b, kur in st:
        if kur:
            bagian.append(f"{label_bulan(b, pendek=True)}: "
                          + ", ".join(LABEL_KANTONG[k] for k in KANTONG if k in kur))
    return " · ".join(bagian)


def cek(siswa, totals, per_bulan, rombel=None):
    """Hasil tombol "Cek siapa yang belum bayar" per bulan `per_bulan`.

    siswa : {induk: dict siswa}     totals: hasil total_per_siswa()
    -> dict:
      menunggak     : ada bulan SEBELUM per_bulan yang belum lunas (baris memuat juga
                      per_bulan bila belum dibayar)
      bulan_ini     : bulan-bulan sebelumnya lunas, hanya per_bulan yang belum
      jumlah_lunas, jumlah_pindah
    Dua daftar sengaja tidak tumpang tindih, supaya yang benar-benar telat tidak
    tenggelam di antara siswa yang sekadar belum bayar bulan berjalan."""
    urutan = bulan_list()
    idx = urutan.index(per_bulan)
    menunggak, bulan_ini_, lunas, pindah = [], [], 0, 0
    for induk, s in siswa.items():
        if rombel and s.get("rombel") != rombel:
            continue
        if s.get("status") == STATUS_PINDAH:
            pindah += 1
            continue
        bayar = totals.get(induk, {"bk": 0, "keg": 0})
        st = status_siswa(s, bayar, per_bulan)
        lalu = [x for x in st[:idx] if x[1]]
        sekarang = st[idx][1]
        baris = {
            "induk": induk, "nama": s.get("nama", ""), "rombel": s.get("rombel", ""),
            "kurang": sum(sum(k.values()) for _, k in st),
            "rincian": ringkas_kurang(st),
            "n_bulan": sum(1 for _, k in st if k),
            "pernah_bayar": bool(bayar["bk"] or bayar["keg"]),
            "tagihan_bulan_ini": sum(tagihan(s, per_bulan).values()),
        }
        if lalu:
            menunggak.append(baris)
        elif sekarang:
            bulan_ini_.append(baris)
        else:
            lunas += 1
    menunggak.sort(key=lambda r: (-r["kurang"], r["rombel"], r["nama"]))
    bulan_ini_.sort(key=lambda r: (r["rombel"], r["nama"]))
    return {"menunggak": menunggak, "bulan_ini": bulan_ini_,
            "jumlah_lunas": lunas, "jumlah_pindah": pindah}


def bulan_lunas(siswa, bayar):
    """Sampai bulan apa siswa ini lunas berturut-turut sejak Juli (None bila Juli pun belum).
    Dipakai pesan setelah mencatat pembayaran."""
    terakhir = None
    for b, kur in status_siswa(siswa, bayar, bulan_list()[-1]):
        if kur:
            break
        terakhir = b
    return terakhir
