"""
Helper Google Sheets (gspread) untuk menu Laporan Keuangan SD.

Kredensial & koneksi spreadsheet di-reuse langsung dari tab_sheet.py (sama
service account, sama pola pencarian kredensial) supaya tidak duplikasi.
Beda dengan Tabungan: di sini barisnya dinamis (outline pohon Program/Sub
Program/Kegiatan/Item/Rincian), jadi pakai insert_row/delete_rows, bukan
tulis-ke-sel-tetap.
"""
import gspread

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
    """Buat tab '<BULAN> <TAHUN>' kalau belum ada. clone_from: worksheet bulan
    sebelumnya (opsional) — kalau diisi, salin kolom LEVEL+LABEL saja, kolom
    finansial dikosongkan."""
    title = C.tab_name(bulan_num, tahun)
    try:
        return book.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        pass

    clone_rows = read_rows(clone_from) if clone_from is not None else []
    n_rows = max(C.FIRST_DATA_ROW + len(clone_rows) + 5, 20)
    ws = book.add_worksheet(title=title, rows=n_rows, cols=C.N_COLS + 1)

    grid = [["" for _ in range(C.N_COLS)] for _ in range(C.FIRST_DATA_ROW - 1 + len(clone_rows))]

    def setv(r, c, v):
        grid[r - 1][c - 1] = v

    setv(C.TITLE_ROW, 1, f"LAPORAN KEUANGAN SD INSAN AMANAH — {C.tab_name(bulan_num, tahun)}")
    for i, h in enumerate(C.HEADERS, start=1):
        setv(C.HEADER_ROW, i, h)
    for idx, row in enumerate(clone_rows):
        r = C.FIRST_DATA_ROW + idx
        setv(r, 1, row["level"])
        setv(r, 2, row["label"])

    if grid:
        ws.update(grid, "A1", value_input_option="USER_ENTERED")
    else:
        ws.update([[f"LAPORAN KEUANGAN SD INSAN AMANAH — {C.tab_name(bulan_num, tahun)}"], C.HEADERS],
                  "A1", value_input_option="USER_ENTERED")
    return ws


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
    ws.update(f"A{row_idx}:O{row_idx}", [_row_values(data)], value_input_option="USER_ENTERED")


def delete_row(ws, row_idx):
    ws.delete_rows(row_idx)
