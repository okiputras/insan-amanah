"""
Helper Google Sheets (gspread) untuk menu Pembayaran SD.

Kredensial & koneksi di-reuse dari tab_sheet.py (service account yang sama).

Tab PEMBAYARAN hanya pernah DITAMBAH (append), tidak pernah ditulis ulang, jadi dua orang
yang mencatat bersamaan tidak saling menimpa — beda dengan Laporan Keuangan yang menulis
ulang seluruh grid. Baris PERLU DICEK dirujuk lewat KUNCI, bukan nomor baris, karena
nomor baris bergeser begitu ada baris lain yang dihapus.
"""
import gspread

import bayar_config as C
import tab_sheet as TS_BASE  # reuse kredensial & open_book

get_client = TS_BASE.get_client

_UANG = {C.TAB_SISWA: ("BPP", "KATERING", "KEGIATAN"),
         C.TAB_BAYAR: ("BPP", "KATERING", "KEGIATAN", "TABUNGAN"),
         C.TAB_REVIEW: ("BPP", "KATERING", "KEGIATAN", "TABUNGAN")}
_HEADER = {C.TAB_SISWA: C.HDR_SISWA, C.TAB_BAYAR: C.HDR_BAYAR, C.TAB_REVIEW: C.HDR_REVIEW}


def open_book(spreadsheet_id=None, client=None):
    return TS_BASE.open_book(spreadsheet_id or C.SPREADSHEET_ID, client=client)


# ---------------------------------------------------------------- tab
def siapkan_tab(book):
    """Pastikan ketiga tab ada dengan header & format dasar. Aman dipanggil berulang."""
    reqs = []
    for title, header in _HEADER.items():
        try:
            book.worksheet(title)
            continue
        except gspread.exceptions.WorksheetNotFound:
            pass
        ws = book.add_worksheet(title=title, rows=2000, cols=len(header))
        ws.update([header], "A1", value_input_option="RAW")
        sid = ws.id
        reqs += [
            {"repeatCell": {
                "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1},
                "cell": {"userEnteredFormat": {
                    "backgroundColor": {"red": 31 / 255, "green": 78 / 255, "blue": 95 / 255},
                    "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}}},
                "fields": "userEnteredFormat(backgroundColor,textFormat)"}},
            {"updateSheetProperties": {
                "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 1}},
                "fields": "gridProperties.frozenRowCount"}},
        ]
        for nama in _UANG[title]:
            c = header.index(nama)
            reqs.append({"repeatCell": {
                "range": {"sheetId": sid, "startRowIndex": 1, "startColumnIndex": c, "endColumnIndex": c + 1},
                "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}},
                "fields": "userEnteredFormat.numberFormat"}})
    if reqs:
        book.batch_update({"requests": reqs})
    for junk in ("Sheet1", "Sheet"):
        try:
            if len(book.worksheets()) > 1:
                book.del_worksheet(book.worksheet(junk))
        except gspread.exceptions.WorksheetNotFound:
            pass


def _baca(book, title):
    """-> list[dict] per baris data, berkunci nama header + '_row' (nomor baris sheet).
    Dibaca UNFORMATTED supaya angka yang diketik user (mis. "700.000" di locale in_ID)
    tetap terbaca sebagai angka."""
    ws = book.worksheet(title)
    grid = ws.get_all_values(value_render_option="UNFORMATTED_VALUE")
    if not grid:
        return ws, []
    header = [str(h).strip().upper() for h in grid[0]]
    out = []
    for i, row in enumerate(grid[1:], start=2):
        row = list(row) + [""] * (len(header) - len(row))
        if not any(str(v).strip() for v in row):
            continue
        d = {h: row[j] for j, h in enumerate(header)}
        d["_row"] = i
        out.append(d)
    return ws, out


def _uang(v):
    if isinstance(v, (int, float)):
        return int(round(v))
    s = str(v or "").strip().replace(".", "").replace(",", "")
    try:
        return int(s) if s else 0
    except ValueError:
        return 0


def _induk(v):
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


# ---------------------------------------------------------------- baca
def baca_siswa(book):
    """-> {induk: {induk, nama, rombel, bpp, katering, kegiatan, status, catatan}}."""
    _, rows = _baca(book, C.TAB_SISWA)
    out = {}
    for d in rows:
        induk = _induk(d.get("INDUK", ""))
        if not induk:
            continue
        status = str(d.get("STATUS", "")).strip().upper() or C.STATUS_AKTIF
        out[induk] = {
            "induk": induk, "nama": str(d.get("NAMA", "")).strip(),
            "rombel": str(d.get("ROMBEL", "")).strip().upper(),
            "bpp": _uang(d.get("BPP")), "katering": _uang(d.get("KATERING")),
            "kegiatan": _uang(d.get("KEGIATAN")),
            "status": C.STATUS_PINDAH if status.startswith(C.STATUS_PINDAH) else C.STATUS_AKTIF,
            "catatan": str(d.get("CATATAN", "")).strip(),
        }
    return out


def baca_pembayaran(book):
    _, rows = _baca(book, C.TAB_BAYAR)
    return [{"tanggal": str(d.get("TANGGAL", "")), "induk": _induk(d.get("INDUK", "")),
             "nama": str(d.get("NAMA", "")), "bpp": _uang(d.get("BPP")),
             "katering": _uang(d.get("KATERING")), "kegiatan": _uang(d.get("KEGIATAN")),
             "tabungan": _uang(d.get("TABUNGAN")), "sumber": str(d.get("SUMBER", "")),
             "kunci": str(d.get("KUNCI", "")), "catatan": str(d.get("CATATAN", ""))}
            for d in rows]


def baca_review(book):
    _, rows = _baca(book, C.TAB_REVIEW)
    return [{"tanggal": str(d.get("TANGGAL", "")), "nama": str(d.get("NAMA TERCATAT", "")),
             "induk_tercatat": _induk(d.get("INDUK TERCATAT", "")),
             "bpp": _uang(d.get("BPP")), "katering": _uang(d.get("KATERING")),
             "kegiatan": _uang(d.get("KEGIATAN")), "tabungan": _uang(d.get("TABUNGAN")),
             "masalah": str(d.get("MASALAH", "")), "saran": _induk(d.get("SARAN INDUK", "")),
             "kunci": str(d.get("KUNCI", ""))}
            for d in rows]


def kunci_terpakai(book):
    """Semua KUNCI di PEMBAYARAN & PERLU DICEK — dipakai impor supaya bisa diulang
    tanpa menggandakan uang."""
    return ({p["kunci"] for p in baca_pembayaran(book) if p["kunci"]}
            | {r["kunci"] for r in baca_review(book) if r["kunci"]})


# ---------------------------------------------------------------- tulis
def _baris_bayar(p):
    return [p["tanggal"], str(p["induk"]), p.get("nama", ""), int(p.get("bpp") or 0),
            int(p.get("katering") or 0), int(p.get("kegiatan") or 0), int(p.get("tabungan") or 0),
            p.get("sumber", ""), p.get("kunci", ""), p.get("catatan", "")]


def catat_pembayaran(book, p):
    """Tambahkan satu baris ke buku besar. RAW: INDUK tetap teks, angka tetap angka,
    tanggal ISO tidak ditafsirkan ulang oleh locale in_ID."""
    book.worksheet(C.TAB_BAYAR).append_row(_baris_bayar(p), value_input_option="RAW")


def catat_banyak(book, daftar):
    if daftar:
        book.worksheet(C.TAB_BAYAR).append_rows([_baris_bayar(p) for p in daftar],
                                                value_input_option="RAW")


def tambah_review(book, daftar):
    if daftar:
        book.worksheet(C.TAB_REVIEW).append_rows(
            [[r["tanggal"], r["nama"], str(r.get("induk_tercatat") or ""), int(r["bpp"]),
              int(r["katering"]), int(r["kegiatan"]), int(r["tabungan"]), r["masalah"],
              str(r.get("saran") or ""), r["kunci"]] for r in daftar],
            value_input_option="RAW")


def isi_siswa(book, siswa):
    """Tulis ulang tab SISWA dari daftar siswa (dipakai impor awal)."""
    ws = book.worksheet(C.TAB_SISWA)
    rows = [[i + 1, s["induk"], s["nama"], s["rombel"], s["bpp"], s["katering"], s["kegiatan"],
             s["status"], s.get("catatan", "")] for i, s in enumerate(siswa)]
    ws.batch_clear([f"A2:I{max(ws.row_count, 2)}"])
    if rows:
        ws.update(rows, f"A2:I{len(rows) + 1}", value_input_option="RAW")


def _hapus_review(book, kunci):
    ws, rows = _baca(book, C.TAB_REVIEW)
    target = next((d for d in rows if str(d.get("KUNCI", "")) == kunci), None)
    if target is None:
        return None
    ws.delete_rows(target["_row"])
    return target


def selesaikan_review(book, kunci, induk, nama):
    """Pindahkan baris PERLU DICEK ke buku besar atas nama siswa `induk`."""
    item = next((r for r in baca_review(book) if r["kunci"] == kunci), None)
    if item is None:
        raise ValueError("Baris tidak ditemukan (mungkin sudah diproses).")
    if kunci in {p["kunci"] for p in baca_pembayaran(book)}:
        _hapus_review(book, kunci)
        raise ValueError("Baris ini sudah pernah dicatat; dihapus dari daftar Perlu Dicek.")
    catat_pembayaran(book, {
        "tanggal": item["tanggal"], "induk": induk, "nama": nama,
        "bpp": item["bpp"], "katering": item["katering"], "kegiatan": item["kegiatan"],
        "tabungan": item["tabungan"], "sumber": "impor (ditinjau)", "kunci": kunci,
        "catatan": f"tercatat sbg '{item['nama']}' / induk '{item['induk_tercatat']}' — {item['masalah']}",
    })
    _hapus_review(book, kunci)
    return item


def abaikan_review(book, kunci):
    if _hapus_review(book, kunci) is None:
        raise ValueError("Baris tidak ditemukan (mungkin sudah diproses).")
