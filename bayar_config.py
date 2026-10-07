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
import re

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

# Pembayaran VA dicatat dari laporan harian bank (R-5401) yang diupload di menu Data
# Validasi SD; No. Pelanggan di laporan SD adalah No Induk. Laporan bertanggal sebelum
# LAPORAN_VA_MULAI sudah tercakup impor Excel PEMASUKAN BPP (blok harian di Excel itu
# salinan laporan VA), jadi dilewati. Batasnya TANGGAL LAPORAN, bukan tanggal transaksi:
# laporan tanggal X memuat transaksi sejak +-21.00 hari sebelumnya.
KODE_VA = "64219"
LAPORAN_VA_MULAI = dt.date(2026, 10, 1)

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


# ---------------------------------------------------------------- laporan VA bank
def _rb(n):
    return f"{int(n):,}".replace(",", ".")


def pecah_nominal(siswa, nilai):
    """Nominal VA -> (bpp, katering, kegiatan, keterangan, banyak_kemungkinan) atau None.

    Laporan bank hanya memuat total. Dicari a x (BPP+katering) + b x kegiatan = nilai
    (a, b = 0..12 bulan); bila ada beberapa, dipilih yang jumlah bulannya paling seimbang.
    Cara ini mereproduksi pembagian BPP+katering / kegiatan yang ditulis sekolah di Excel
    untuk 106 dari 106 transaksi VA 6-9 Agustus 2026 (bayar rangkap, BPP saja, kegiatan
    kelas 1 dirangkap), dan dengan tarif T.A. 2026/2027 tidak ada nominal lazim (bulan
    penuh / BPP saja / kegiatan saja) yang punya dua pecahan."""
    bpp, kat, keg = (int(siswa.get(k) or 0) for k in ("bpp", "katering", "kegiatan"))
    bk = bpp + kat
    sol = [(a, b) for a in range(13 if bk else 1) for b in range(13 if keg else 1)
           if (a or b) and a * bk + b * keg == nilai]
    if not sol:
        return None
    a, b = min(sol, key=lambda x: (abs(x[0] - x[1]), -x[0]))
    ket = " + ".join(t for t in (f"{a}× BPP+katering" if a else "", f"{b}× kegiatan" if b else "") if t)
    return a * bpp, a * kat, b * keg, ket, len(sol)


def induk_dari_va(cust):
    """No. Pelanggan laporan -> No Induk. R-5401 SD sudah memuat No Induk ('2551'); format
    BCA VA memuat VA penuh (kode biller + induk), jadi kode billernya dibuang."""
    c = re.sub(r"\D", "", str(cust or ""))
    if len(c) > len(KODE_VA) and c.startswith(KODE_VA):
        c = c[len(KODE_VA):]
    return str(int(c)) if c else ""


def tanggal_laporan(meta, rows):
    """Tanggal di kepala laporan ('06/08/2026'); cadangan: tanggal transaksi terakhir."""
    try:
        d, m, y = (int(x) for x in str(meta.get("tanggal_label", "")).split("/"))
        return dt.date(y, m, d)
    except ValueError:
        return max((r[4] for r in rows), default=None)


def alasan_tolak_va(meta, rows):
    """Alasan sebuah laporan tidak boleh dicatat sama sekali, atau None. Tanpa akses Sheet."""
    if str(meta.get("kode") or "").strip() != KODE_VA:
        return (f"Laporan ini milik kode biller {meta.get('kode') or '?'}, bukan SD ({KODE_VA}), "
                "jadi tidak dicatat ke Pembayaran SD.")
    tgl = tanggal_laporan(meta, rows)
    if tgl is None:
        return "Tanggal laporan tidak terbaca."
    if tgl < LAPORAN_VA_MULAI:
        akhir = LAPORAN_VA_MULAI - dt.timedelta(days=1)
        return (f"Laporan {tgl:%d/%m/%Y} sudah tercakup impor Excel PEMASUKAN BPP (laporan s/d "
                f"{akhir:%d/%m/%Y}); tidak dicatat lagi supaya tidak dobel.")
    return None


def rencana_va(meta, rows, siswa, kunci_ada):
    """Laporan VA harian (hasil parser.parse_laporan) -> rencana pencatatan ke buku besar.

    -> {"tanggal_laporan", "alasan": teks bila laporan tidak boleh dicatat sama sekali,
        "baris": [dict per transaksi dengan status "baru" | "sudah" | "cek"]}
    KUNCI per transaksi = tanggal-jam + induk + nominal, jadi laporan yang sama (atau yang
    tumpang tindih) diupload ulang tidak mendobelkan uang. Dipanggil lagi saat tombol Catat
    ditekan, terhadap isi Sheet terbaru, supaya klik ganda pun aman."""
    out = {"tanggal_laporan": tanggal_laporan(meta, rows), "alasan": alasan_tolak_va(meta, rows),
           "baris": []}
    if out["alasan"]:
        return out
    dilihat = set()
    for _no, cust, nama_lap, nilai, tgl, waktu, _jam, lok, k1, k2 in rows:
        nilai = int(round(nilai))
        induk = induk_dari_va(cust)
        kunci = f"VA:{tgl:%Y%m%d}-{waktu:%H%M%S}:{induk}:{nilai}"
        s = siswa.get(induk)
        b = {"kunci": kunci, "tanggal": tgl.isoformat(), "waktu": waktu.strftime("%H:%M:%S"),
             "induk": induk, "no_pelanggan": str(cust).strip(), "nama_laporan": str(nama_lap).strip(),
             "nama": s["nama"] if s else str(nama_lap).strip(), "rombel": s["rombel"] if s else "",
             "nilai": nilai, "lokasi": str(lok or "").strip(),
             "berita": " / ".join(x.strip() for x in (k1, k2) if x and x.strip(" -")),
             "bpp": 0, "katering": 0, "kegiatan": 0, "pecahan": "", "masalah": "", "status": "baru"}
        if kunci in kunci_ada or kunci in dilihat:
            b["status"] = "sudah"
        elif s is None:
            b.update(status="cek", bpp=nilai, pecahan="sementara: semua ke BPP",
                     masalah=f"No. Pelanggan {cust} tidak ada di tab {TAB_SISWA}")
        else:
            p = pecah_nominal(s, nilai)
            if p:
                b.update(bpp=p[0], katering=p[1], kegiatan=p[2], pecahan=p[3])
            else:
                # Pecahan sementara untuk ditinjau: bulan penuh BPP+katering dulu, sisanya kegiatan.
                bk = s["bpp"] + s["katering"]
                a = min(nilai // bk, 12) if bk else 0
                b.update(bpp=a * s["bpp"], katering=a * s["katering"], kegiatan=nilai - a * bk,
                         pecahan=(f"sementara: {a}× BPP+katering, sisa ke kegiatan" if a
                                  else "sementara: semua ke kegiatan"))
            if s["status"] == STATUS_PINDAH:
                b.update(status="cek", masalah=f"{s['nama']} berstatus PINDAH di tab {TAB_SISWA}")
            elif not p:
                b.update(status="cek", masalah=(
                    f"nominal {_rb(nilai)} bukan kelipatan tagihan (BPP+katering {_rb(s['bpp'] + s['katering'])}, "
                    f"kegiatan {_rb(s['kegiatan'])}) — mungkin termasuk tabungan atau cicilan; "
                    "ubah angkanya di tab PERLU DICEK bila perlu, lalu Simpan"))
            elif p[4] > 1:
                b.update(status="cek", masalah="nominal bisa dipecah lebih dari satu cara — periksa pecahannya")
        dilihat.add(kunci)
        out["baris"].append(b)
    return out


def laporan_va_terakhir(pembayaran):
    """Tanggal laporan VA terbaru yang sudah dicatat (dari SUMBER "VA yyyy-mm-dd"), atau None.
    Ditampilkan di hasil Cek: pembayaran VA dari laporan yang belum diupload belum terhitung."""
    tgl = [p["sumber"][3:13] for p in pembayaran if re.match(r"VA \d{4}-\d{2}-\d{2}", p.get("sumber", ""))]
    return dt.date.fromisoformat(max(tgl)) if tgl else None
