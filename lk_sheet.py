"""
Helper Google Sheets (gspread) untuk menu Laporan Keuangan SD.

Kredensial & koneksi spreadsheet di-reuse langsung dari tab_sheet.py (sama
service account, sama pola pencarian kredensial) supaya tidak duplikasi.
Beda dengan Tabungan: di sini barisnya dinamis (outline pohon Program/Sub
Program/Kegiatan/Item/Rincian), jadi pakai insert_row/delete_rows, bukan
tulis-ke-sel-tetap.
"""
import gspread
from gspread.utils import rowcol_to_a1

import lk_config as C
import tab_sheet as TS_BASE  # reuse kredensial & open_book

get_client = TS_BASE.get_client


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
    kolom LEVEL+LABEL saja, kolom finansial dikosongkan. Kalau tidak diisi
    (mis. tab pertama / tidak ada bulan sebelumnya), pakai C.DEFAULT_TEMPLATE
    sebagai struktur awal."""
    title = C.tab_name(bulan_num, tahun)
    try:
        return book.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        pass

    if clone_from is not None:
        seed_rows = [(r["level"], r["label"]) for r in read_rows(clone_from)]
    else:
        seed_rows = C.DEFAULT_TEMPLATE
    n_rows = max(C.FIRST_DATA_ROW + len(seed_rows) + 5, 20)
    ws = book.add_worksheet(title=title, rows=n_rows, cols=C.N_COLS + 1)

    grid = [["" for _ in range(C.N_COLS)] for _ in range(C.FIRST_DATA_ROW - 1 + len(seed_rows))]

    def setv(r, c, v):
        grid[r - 1][c - 1] = v

    setv(C.TITLE_ROW, 1, f"LAPORAN KEUANGAN SD INSAN AMANAH — {title}")
    for i, h in enumerate(C.HEADERS, start=1):
        setv(C.HEADER_ROW, i, h)
    for idx, (level, label) in enumerate(seed_rows):
        r = C.FIRST_DATA_ROW + idx
        setv(r, 1, level)
        setv(r, 2, label)

    ws.update(grid, "A1", value_input_option="USER_ENTERED")
    _format_tab(book, ws)
    return ws


# ---------------------------------------------------------------- format
def _rgb(r, g, b):
    return {"red": r / 255, "green": g / 255, "blue": b / 255}


TITLE_BG = _rgb(31, 78, 95)     # selaras tema teal app (--teal)
HEADER_BG = _rgb(43, 107, 128)  # --teal2
WHITE = _rgb(255, 255, 255)
LVL_BG = {1: _rgb(210, 230, 236), 2: _rgb(230, 242, 247)}

MONEY_COLS0 = [4, 8, 9]   # VOLUME, UNIT_COST, TOTAL (0-based)
DATE_COL0 = 2              # TANGGAL


def _format_tab(book, ws):
    sid = ws.id
    last_row = max(ws.row_count, C.FIRST_DATA_ROW)
    n_cols = C.N_COLS

    def rng(r1, c1, r2, c2):
        return {"sheetId": sid, "startRowIndex": r1, "endRowIndex": r2,
                "startColumnIndex": c1, "endColumnIndex": c2}

    reqs = [
        {"mergeCells": {"range": rng(0, 0, 1, n_cols), "mergeType": "MERGE_ALL"}},
        {"repeatCell": {
            "range": rng(0, 0, 1, n_cols),
            "cell": {"userEnteredFormat": {"backgroundColor": TITLE_BG, "horizontalAlignment": "CENTER",
                     "textFormat": {"foregroundColor": WHITE, "bold": True, "fontSize": 12}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)"}},
        {"repeatCell": {
            "range": rng(1, 0, 2, n_cols),
            "cell": {"userEnteredFormat": {"backgroundColor": HEADER_BG, "horizontalAlignment": "CENTER",
                     "wrapStrategy": "WRAP", "textFormat": {"foregroundColor": WHITE, "bold": True, "fontSize": 9}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,wrapStrategy,textFormat)"}},
        {"updateSheetProperties": {
            "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 2}},
            "fields": "gridProperties.frozenRowCount"}},
    ]
    for c in MONEY_COLS0:
        reqs.append({"repeatCell": {
            "range": rng(2, c, last_row, c + 1),
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}},
            "fields": "userEnteredFormat.numberFormat"}})
    reqs.append({"repeatCell": {
        "range": rng(2, DATE_COL0, last_row, DATE_COL0 + 1),
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "DATE", "pattern": "dd/mm/yyyy"},
                 "horizontalAlignment": "CENTER"}},
        "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"}})

    # Bold/warna per LEVEL — pakai conditional formatting supaya tetap benar
    # walau baris ditambah/dihapus/edit lewat app nanti (bukan format statis).
    def cond(level, fmt):
        return {"addConditionalFormatRule": {"rule": {
            "ranges": [rng(2, 0, last_row, n_cols)],
            "booleanRule": {
                "condition": {"type": "CUSTOM_FORMULA",
                              "values": [{"userEnteredValue": f"=$A3={level}"}]},
                "format": fmt}}, "index": 0}}

    reqs.append(cond(1, {"backgroundColor": LVL_BG[1],
                          "textFormat": {"bold": True, "foregroundColor": TITLE_BG}}))
    reqs.append(cond(2, {"backgroundColor": LVL_BG[2], "textFormat": {"bold": True}}))
    reqs.append(cond(3, {"textFormat": {"bold": True}}))
    reqs.append(cond(4, {"textFormat": {"bold": True}}))

    def w(c0, c1, px):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": c0, "endIndex": c1},
            "properties": {"pixelSize": px}, "fields": "pixelSize"}})
    w(0, 1, 46); w(1, 2, 260)

    book.batch_update({"requests": reqs})


# ---------------------------------------------------------------- baca
def read_rows(ws):
    values = ws.get_all_values()
    out = []
    for r in range(C.FIRST_DATA_ROW, len(values) + 1):
        row = values[r - 1] + [""] * C.N_COLS
        level = (row[0] or "").strip()
        label = (row[1] or "").strip()
        if not level and not label:
            continue
        d = {"row": r}
        for i, key in enumerate(C.FIELD_KEYS):
            d[key] = row[i].strip() if isinstance(row[i], str) else row[i]
        try:
            d["level"] = int(float(d["level"])) if d["level"] not in ("", None) else 1
        except (ValueError, TypeError):
            d["level"] = 1
        out.append(d)
    return out


def last_data_row(ws):
    rows = read_rows(ws)
    return rows[-1]["row"] if rows else C.FIRST_DATA_ROW - 1


# ---------------------------------------------------------------- tulis
def _row_values(data):
    out = []
    for key in C.FIELD_KEYS:
        v = data.get(key, "")
        out.append(v if v not in (None,) else "")
    return out


def insert_row(ws, after_row_idx, data):
    """after_row_idx=None → tambah di akhir. Return nomor baris baru."""
    idx = (after_row_idx + 1) if after_row_idx else (last_data_row(ws) + 1)
    ws.insert_row(_row_values(data), idx, value_input_option="USER_ENTERED")
    return idx


def update_row(ws, row_idx, data):
    last_col = rowcol_to_a1(1, C.N_COLS)[:-1]   # mis. "J" (dari "J1")
    ws.update(f"A{row_idx}:{last_col}{row_idx}", [_row_values(data)], value_input_option="USER_ENTERED")


def delete_row(ws, row_idx):
    ws.delete_rows(row_idx)
