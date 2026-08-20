"""
Helper Google Sheets (gspread) untuk menu Laporan Keuangan SD.

Kredensial & koneksi spreadsheet di-reuse langsung dari tab_sheet.py (sama
service account, sama pola pencarian kredensial) supaya tidak duplikasi.

Beda dengan Tabungan: di sini barisnya dinamis (outline pohon Program/Sub
Program/Kegiatan/Item/Rincian), jadi pakai insert_row/delete_rows, bukan
tulis-ke-sel-tetap. Tiga hal dikelola otomatis lewat resync() setiap kali ada
baris ditambah/diedit/dihapus:
  1. Kolom NO & ITEM  — nomor/huruf berjenjang (lihat lk_config.py).
  2. Baris "Jumlah Biaya" — subtotal SUM per Program, di-hapus & dibuat ulang
     dari nol tiap resync supaya posisinya selalu tepat (lebih aman daripada
     mengandalkan auto-expand range formula Sheets saat insert/delete baris).
  3. Warna latar bergantian per Program & bold per LEVEL — ini TIDAK perlu
     disentuh ulang tiap resync karena dikerjakan lewat conditional formatting
     bawaan Sheets (rule COUNTIF ganjil/genap utk warna, rule per-LEVEL utk
     bold) yang otomatis mengikuti baris baru — cukup dipasang sekali di
     _format_tab() saat tab dibuat.
"""
import gspread
from gspread.utils import rowcol_to_a1

import lk_config as C
import tab_sheet as TS_BASE  # reuse kredensial & open_book

get_client = TS_BASE.get_client

# Kolom (0-based): 0=NO 1=ITEM 2=LEVEL 3=LABEL 4=TANGGAL 5=KODE 6=VOLUME
#                  7=SATUAN 8=FK 9=KEBUTUHAN 10=UNIT_COST 11=TOTAL
NO_COL0, ITEM_COL0, LEVEL_COL0, LABEL_COL0 = 0, 1, 2, 3
TOTAL_COL0 = 11
MONEY_COLS0 = [6, 10, 11]   # VOLUME, UNIT_COST, TOTAL
DATE_COL0 = 4                # TANGGAL


def open_book(spreadsheet_id=None, client=None):
    return TS_BASE.open_book(spreadsheet_id or C.SPREADSHEET_ID, client=client)


# ---------------------------------------------------------------- tab bulan
def _parse_tab_title(title):
    parts = title.split()
    if len(parts) != 2:
        return None
    bulan_name, tahun_s = parts
    bulan_name = bulan_name.upper()
    if bulan_name not in C.MONTHS_ID or not tahun_s.isdigit():
        return None
    return C.MONTHS_ID.index(bulan_name) + 1, int(tahun_s)


def list_bulan_tabs(book):
    """Return list of (bulan_num, tahun, title), urut kronologis naik."""
    out = []
    for ws in book.worksheets():
        parsed = _parse_tab_title(ws.title)
        if parsed:
            out.append((parsed[0], parsed[1], ws.title))
    out.sort(key=lambda t: (t[1], t[0]))
    return out


def ensure_bulan_tab(book, bulan_num, tahun, clone_from=None):
    """Buat tab '<BULAN> <TAHUN>' kalau belum ada.
    clone_from: worksheet bulan sebelumnya (opsional) — kalau diisi, salin
    kolom LEVEL+LABEL saja (baris outline, bukan subtotal), kolom finansial
    dikosongkan. Kalau tidak diisi, pakai C.DEFAULT_TEMPLATE."""
    title = C.tab_name(bulan_num, tahun)
    try:
        return book.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        pass

    if clone_from is not None:
        seed_rows = [(r["level"], r["label"]) for r in read_rows(clone_from)
                     if r["level"] != C.LEVEL_SUBTOTAL]
    else:
        seed_rows = C.DEFAULT_TEMPLATE
    n_rows = max(C.FIRST_DATA_ROW + len(seed_rows) + 20, FORMAT_ROW_BOUND)
    ws = book.add_worksheet(title=title, rows=n_rows, cols=C.N_COLS + 1)

    grid = [["" for _ in range(C.N_COLS)] for _ in range(C.FIRST_DATA_ROW - 1 + len(seed_rows))]

    def setv(r, c, v):
        grid[r - 1][c - 1] = v

    setv(C.TITLE_ROW1, 1, "LAPORAN PERTANGGUNGJAWABAN OPERASIONAL SD INSAN AMANAH")
    setv(C.TITLE_ROW2, 1, f"TAHUN PELAJARAN {C.academic_label(bulan_num, tahun)}")
    setv(C.TITLE_ROW3, 1, f"BULAN {title}")
    for i, h in enumerate(C.HEADERS, start=1):
        setv(C.HEADER_ROW, i, h)
    for idx, (level, label) in enumerate(seed_rows):
        r = C.FIRST_DATA_ROW + idx
        setv(r, LEVEL_COL0 + 1, level)
        setv(r, LABEL_COL0 + 1, label)

    ws.update(grid, "A1", value_input_option="USER_ENTERED")
    resync(ws)
    _format_tab(book, ws)
    return ws


# ---------------------------------------------------------------- baca
def read_rows(ws):
    values = ws.get_all_values()
    out = []
    for r in range(C.FIRST_DATA_ROW, len(values) + 1):
        row = values[r - 1] + [""] * C.N_COLS
        level_s = (row[LEVEL_COL0] or "").strip()
        label = (row[LABEL_COL0] or "").strip()
        if not level_s and not label:
            continue
        d = {"row": r, "no": (row[NO_COL0] or "").strip(), "item": (row[ITEM_COL0] or "").strip()}
        for i, key in enumerate(C.FIELD_KEYS):
            v = row[LEVEL_COL0 + i]
            d[key] = v.strip() if isinstance(v, str) else v
        try:
            d["level"] = int(float(d["level"])) if d["level"] not in ("", None) else 1
        except (ValueError, TypeError):
            d["level"] = 1
        out.append(d)
    return out


def last_data_row(ws):
    rows = read_rows(ws)
    return rows[-1]["row"] if rows else C.FIRST_DATA_ROW - 1


# ---------------------------------------------------------------- tulis (baris outline)
def _row_values(data):
    out = []
    for key in C.FIELD_KEYS:
        v = data.get(key, "")
        out.append(v if v not in (None,) else "")
    return out


def insert_row(ws, after_row_idx, data):
    """after_row_idx=None → tambah di akhir. Return nomor baris baru (perkiraan
    sebelum resync; posisi final bisa bergeser sedikit kalau ada subtotal)."""
    idx = (after_row_idx + 1) if after_row_idx else (last_data_row(ws) + 1)
    ws.insert_row(["", ""] + _row_values(data), idx, value_input_option="USER_ENTERED")
    resync(ws)
    return idx


def update_row(ws, row_idx, data):
    first_col = rowcol_to_a1(1, LEVEL_COL0 + 1)[:-1]   # "C"
    last_col = rowcol_to_a1(1, C.N_COLS)[:-1]           # "L"
    ws.update(f"{first_col}{row_idx}:{last_col}{row_idx}", [_row_values(data)],
              value_input_option="USER_ENTERED")
    resync(ws)


def delete_row(ws, row_idx):
    ws.delete_rows(row_idx)
    resync(ws)


# ---------------------------------------------------------------- resync (NO/ITEM + subtotal)
def _letter(n):
    """1->a, 2->b, ..., 26->z, 27->aa, ..."""
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(97 + rem) + s
    return s


def _compute_no_item(levels):
    """levels: list int (baris outline, LEVEL_SUBTOTAL diperbolehkan ikut lewat
    sebagai lvl=6 → dapat NO/ITEM kosong tanpa memengaruhi counter). Return
    list (no, item) sejajar."""
    counters = [0, 0, 0, 0, 0, 0]   # index 1..5 dipakai
    out = []
    for lvl in levels:
        if lvl == C.LEVEL_SUBTOTAL:
            out.append(("", ""))
            continue
        lvl = max(1, min(5, int(lvl)))
        counters[lvl] += 1
        for d in range(lvl + 1, 6):
            counters[d] = 0
        if lvl == 1:
            out.append((str(counters[1]), ""))
        elif lvl == 2:
            out.append((f"{counters[1]}.{counters[2]}", ""))
        elif lvl == 3:
            out.append((str(counters[3]), ""))
        elif lvl == 4:
            out.append(("", _letter(counters[4])))
        else:
            out.append(("", ""))
    return out


def _remove_subtotal_rows(ws):
    rows = read_rows(ws)
    sub_rows = [r["row"] for r in rows if r["level"] == C.LEVEL_SUBTOTAL]
    for r in sorted(sub_rows, reverse=True):
        ws.delete_rows(r)


def _insert_subtotal_rows(ws):
    """Kelompokkan baris outline per Program (level 1), sisipkan 1 baris
    'Jumlah Biaya' (formula SUM) di akhir tiap kelompok. Diproses dari
    kelompok terakhir ke pertama supaya nomor baris kelompok sebelumnya
    tidak perlu dihitung ulang di tengah proses."""
    rows = read_rows(ws)
    if not rows:
        return
    groups = []
    cur = None
    for r in rows:
        if r["level"] == 1:
            cur = [r]
            groups.append(cur)
        elif cur is not None:
            cur.append(r)
        # baris sebelum Program pertama (seharusnya tidak terjadi) diabaikan
    total_col = rowcol_to_a1(1, TOTAL_COL0 + 1)[:-1]   # "L"
    for grp in reversed(groups):
        first_row, last_row = grp[0]["row"], grp[-1]["row"]
        formula = f"=SUM({total_col}{first_row}:{total_col}{last_row})"
        values = [""] * C.N_COLS
        values[LEVEL_COL0] = C.LEVEL_SUBTOTAL
        values[LABEL_COL0] = C.SUBTOTAL_LABEL
        values[TOTAL_COL0] = formula
        ws.insert_row(values, last_row + 1, value_input_option="USER_ENTERED")


def resync(ws):
    """Panggil setelah insert/update/delete baris outline: bangun ulang baris
    subtotal per Program, lalu hitung ulang kolom NO & ITEM untuk semua baris."""
    _remove_subtotal_rows(ws)
    _insert_subtotal_rows(ws)

    rows = read_rows(ws)
    if not rows:
        return
    no_item = _compute_no_item([r["level"] for r in rows])
    first_row, last_row = rows[0]["row"], rows[-1]["row"]
    values = [[no, item] for no, item in no_item]
    ws.update(f"A{first_row}:B{last_row}", values, value_input_option="USER_ENTERED")


# ---------------------------------------------------------------- format (sekali saat tab dibuat)
def _rgb(r, g, b):
    return {"red": r / 255, "green": g / 255, "blue": b / 255}


TITLE_BG = _rgb(31, 78, 95)     # selaras tema teal app (--teal)
HEADER_BG = _rgb(43, 107, 128)  # --teal2
WHITE = _rgb(255, 255, 255)
BAND_A = _rgb(252, 231, 248)    # pink pastel (mirip Program ganjil di contoh)
BAND_B = _rgb(214, 247, 247)    # tosca pastel (mirip Program genap di contoh)
SUBTOTAL_BG = _rgb(224, 219, 245)
FORMAT_ROW_BOUND = 1000   # headroom baris untuk pertumbuhan data ke depan


def _format_tab(book, ws):
    sid = ws.id
    # Batas baris yang lega (bukan cuma sepanjang data saat ini) supaya
    # conditional formatting & format angka/tanggal otomatis ikut berlaku
    # untuk baris-baris yang ditambahkan user di masa depan.
    last_row = max(ws.row_count, FORMAT_ROW_BOUND)
    n_cols = C.N_COLS
    level_col_a1 = rowcol_to_a1(1, LEVEL_COL0 + 1)[:-1]   # "C"

    def rng(r1, c1, r2, c2):
        return {"sheetId": sid, "startRowIndex": r1, "endRowIndex": r2,
                "startColumnIndex": c1, "endColumnIndex": c2}

    data_first_row0 = C.FIRST_DATA_ROW - 1

    reqs = [
        {"mergeCells": {"range": rng(0, 0, 3, n_cols), "mergeType": "MERGE_ALL"}},
    ]
    for r0 in range(3):
        reqs.append({"repeatCell": {
            "range": rng(r0, 0, r0 + 1, n_cols),
            "cell": {"userEnteredFormat": {"backgroundColor": TITLE_BG, "horizontalAlignment": "CENTER",
                     "textFormat": {"foregroundColor": WHITE, "bold": True,
                                    "fontSize": 14 if r0 == 0 else 11}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)"}})
    reqs += [
        {"repeatCell": {
            "range": rng(3, 0, 4, n_cols),
            "cell": {"userEnteredFormat": {"backgroundColor": HEADER_BG, "horizontalAlignment": "CENTER",
                     "wrapStrategy": "WRAP", "textFormat": {"foregroundColor": WHITE, "bold": True, "fontSize": 9}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,wrapStrategy,textFormat)"}},
        {"repeatCell": {
            "range": rng(data_first_row0, 0, last_row, 2),
            "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat.horizontalAlignment"}},
        {"updateSheetProperties": {
            "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": C.HEADER_ROW}},
            "fields": "gridProperties.frozenRowCount"}},
    ]
    for c in MONEY_COLS0:
        reqs.append({"repeatCell": {
            "range": rng(data_first_row0, c, last_row, c + 1),
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}},
            "fields": "userEnteredFormat.numberFormat"}})
    reqs.append({"repeatCell": {
        "range": rng(data_first_row0, DATE_COL0, last_row, DATE_COL0 + 1),
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "DATE", "pattern": "dd/mm/yyyy"},
                 "horizontalAlignment": "CENTER"}},
        "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"}})

    data_range = rng(data_first_row0, 0, last_row, n_cols)
    anchor_row = C.FIRST_DATA_ROW   # baris pertama data, dipakai sbg anchor formula relatif

    def cond(formula, fmt, ranges=None):
        return {"addConditionalFormatRule": {"rule": {
            "ranges": ranges or [data_range],
            "booleanRule": {
                "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": formula}]},
                "format": fmt}}, "index": 0}}

    # Warna latar bergantian per Program: hitung berapa banyak baris LEVEL=1
    # dari baris pertama sampai baris ini (COUNTIF), ganjil/genap → 2 warna.
    # Pakai range absolut $col$first:col{row} supaya ikut membesar otomatis.
    count_formula = f"COUNTIF(${level_col_a1}${anchor_row}:{level_col_a1}{anchor_row},1)"
    reqs.append(cond(f"=ISODD({count_formula})", {"backgroundColor": BAND_A}))
    reqs.append(cond(f"=ISEVEN({count_formula})", {"backgroundColor": BAND_B}))

    # Bold per LEVEL (Program/Sub Program/Kegiatan/Item) & subtotal
    def lvl_cond(level, fmt):
        return cond(f"=${level_col_a1}{anchor_row}={level}", fmt)

    reqs.append(lvl_cond(1, {"textFormat": {"bold": True, "foregroundColor": TITLE_BG}}))
    reqs.append(lvl_cond(2, {"textFormat": {"bold": True, "italic": True}}))
    reqs.append(lvl_cond(3, {"textFormat": {"bold": True}}))
    reqs.append(lvl_cond(4, {"textFormat": {"bold": True}}))
    reqs.append(lvl_cond(C.LEVEL_SUBTOTAL,
                          {"backgroundColor": SUBTOTAL_BG, "textFormat": {"bold": True, "italic": True}}))

    def w(c0, c1, px):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": c0, "endIndex": c1},
            "properties": {"pixelSize": px}, "fields": "pixelSize"}})
    w(0, 1, 46); w(1, 2, 34)   # NO, ITEM
    w(3, 4, 260)                # LABEL
    # Kolom LEVEL (C) disembunyikan — tetap ada datanya (dipakai form edit &
    # conditional formatting), tapi tidak perlu terlihat karena sudah terwakili NO/ITEM.
    reqs.append({"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3},
        "properties": {"hiddenByUser": True}, "fields": "hiddenByUser"}})

    book.batch_update({"requests": reqs})
