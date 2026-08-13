"""
Helper Google Sheets (gspread) untuk menu Tabungan SMP.

Kredensial service account, urutan pencarian:
  1. env GOOGLE_SERVICE_ACCOUNT_JSON  (isi JSON penuh — dipakai di Railway)
  2. env SA_FILE                        (path ke file json)
  3. file sa-sheet.json                 (dicari beberapa level di atas — lokal)

Spreadsheet harus di-share ke SERVICE_ACCOUNT_EMAIL sebagai Editor.
"""
import os
import json
import datetime as dt

import gspread
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

import tab_config as C

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

_HERE = os.path.dirname(__file__)
_FILE_CANDIDATES = [
    os.environ.get("SA_FILE", ""),
    C.SERVICE_ACCOUNT_FILE,
    os.path.join(_HERE, C.SERVICE_ACCOUNT_FILE),
    os.path.join(_HERE, "..", C.SERVICE_ACCOUNT_FILE),
    os.path.join(_HERE, "..", "..", C.SERVICE_ACCOUNT_FILE),
    os.path.join(_HERE, "..", "..", "..", C.SERVICE_ACCOUNT_FILE),
]


def _load_creds():
    env_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if env_json:
        return Credentials.from_service_account_info(json.loads(env_json), scopes=SCOPES)
    for p in _FILE_CANDIDATES:
        if p and os.path.exists(p):
            return Credentials.from_service_account_file(os.path.abspath(p), scopes=SCOPES)
    raise FileNotFoundError(
        "Kredensial service account tidak ditemukan. Set env GOOGLE_SERVICE_ACCOUNT_JSON "
        "(di Railway) atau taruh sa-sheet.json di folder aplikasi."
    )


def get_client():
    return gspread.authorize(_load_creds())


def open_book(spreadsheet_id=None, client=None):
    client = client or get_client()
    sid = spreadsheet_id or C.SPREADSHEET_ID
    msg = ("Service account belum punya akses ke spreadsheet.\n"
           "Share spreadsheet ke (sebagai Editor): " + C.SERVICE_ACCOUNT_EMAIL)
    try:
        return client.open_by_key(sid)
    except PermissionError as e:
        raise PermissionError(msg) from e
    except gspread.exceptions.APIError as e:
        if "403" in str(e) or "PERMISSION" in str(e).upper():
            raise PermissionError(msg) from e
        raise


# ---------------------------------------------------------------- warna
def _rgb(r, g, b):
    return {"red": r / 255, "green": g / 255, "blue": b / 255}


TITLE_BG = _rgb(31, 78, 95)     # selaras tema teal app (--teal #1F4E5F)
MONTH_BG = _rgb(43, 107, 128)   # --teal2
SUB_BG = _rgb(230, 242, 247)
FIXED_BG = _rgb(210, 230, 236)
WHITE = _rgb(255, 255, 255)


# ---------------------------------------------------------------- build
def build_tab(book, kelas, year, roster, jenjang="SMP"):
    """Buat/timpa tab '<jenjang> <kelas> <year>'. roster: list of (induk, nama, saldo_awal)."""
    title = C.tab_name(kelas, year, jenjang)
    months = C.months_for_year(year)
    n_cols = C.total_cols(year)
    n_rows = C.FIRST_DATA_ROW + len(roster) + 5

    try:
        ws = book.worksheet(title)
        ws.clear()
        ws.resize(rows=max(n_rows, 10), cols=max(n_cols, 5))
    except gspread.exceptions.WorksheetNotFound:
        ws = book.add_worksheet(title=title, rows=n_rows, cols=n_cols)

    grid = [["" for _ in range(n_cols)] for _ in range(C.FIRST_DATA_ROW - 1 + len(roster))]

    def setv(r, c, v):
        grid[r - 1][c - 1] = v

    setv(1, 1, f"TABUNGAN {jenjang} INSAN AMANAH  —  KELAS {kelas}  —  T.A. {C.academic_label(year)}")
    for i, h in enumerate(C.FIXED_HEADERS, start=1):
        setv(C.HEADER_MONTH_ROW, i, h)
    for pos, m in enumerate(months):
        start = C.block_start_col(pos)
        setv(C.HEADER_MONTH_ROW, start, C.month_label(m, year))
        for i, sub in enumerate(C.SUB_HEADERS):
            setv(C.HEADER_SUB_ROW, start + i, sub)

    for idx, (induk, nama, saldo_awal) in enumerate(roster):
        r = C.FIRST_DATA_ROW + idx
        setv(r, 1, idx + 1)
        setv(r, 2, str(induk))
        setv(r, 3, nama)
        setv(r, 4, saldo_awal if saldo_awal not in (None, "") else 0)
        for pos in range(len(months)):
            setv(r, C.block_start_col(pos) + 4, C.saldo_formula(r, pos))

    ws.update(grid, "A1", value_input_option="USER_ENTERED")
    induk_vals = [[str(ind)] for ind, _n, _s in roster]
    if induk_vals:
        last = C.FIRST_DATA_ROW + len(roster) - 1
        ws.update(induk_vals, f"B{C.FIRST_DATA_ROW}:B{last}", value_input_option="RAW")
    _format_tab(book, ws, kelas, year)
    return ws


def _format_tab(book, ws, kelas, year):
    months = C.months_for_year(year)
    n_cols = C.total_cols(year)
    sid = ws.id
    last_row = ws.row_count
    reqs = []

    def rng(r1, c1, r2, c2):
        return {"sheetId": sid, "startRowIndex": r1, "endRowIndex": r2,
                "startColumnIndex": c1, "endColumnIndex": c2}

    reqs.append({"repeatCell": {
        "range": rng(0, 0, 1, n_cols),
        "cell": {"userEnteredFormat": {"backgroundColor": TITLE_BG, "horizontalAlignment": "LEFT",
                 "textFormat": {"foregroundColor": WHITE, "bold": True, "fontSize": 12}}},
        "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)"}})

    for i in range(C.N_FIXED):
        reqs.append({"mergeCells": {"range": rng(1, i, 3, i + 1), "mergeType": "MERGE_ALL"}})
    reqs.append({"repeatCell": {
        "range": rng(1, 0, 3, C.N_FIXED),
        "cell": {"userEnteredFormat": {"backgroundColor": FIXED_BG, "horizontalAlignment": "CENTER",
                 "verticalAlignment": "MIDDLE", "wrapStrategy": "WRAP", "textFormat": {"bold": True}}},
        "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,wrapStrategy,textFormat)"}})

    for pos in range(len(months)):
        start = C.block_start_col(pos) - 1
        reqs.append({"mergeCells": {"range": rng(1, start, 2, start + C.COLS_PER_MONTH),
                                    "mergeType": "MERGE_ALL"}})
    reqs.append({"repeatCell": {
        "range": rng(1, C.N_FIXED, 2, n_cols),
        "cell": {"userEnteredFormat": {"backgroundColor": MONTH_BG, "horizontalAlignment": "CENTER",
                 "textFormat": {"foregroundColor": WHITE, "bold": True}}},
        "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)"}})
    reqs.append({"repeatCell": {
        "range": rng(2, C.N_FIXED, 3, n_cols),
        "cell": {"userEnteredFormat": {"backgroundColor": SUB_BG, "horizontalAlignment": "CENTER",
                 "wrapStrategy": "WRAP", "textFormat": {"bold": True, "fontSize": 9}}},
        "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,wrapStrategy,textFormat)"}})

    reqs.append({"updateSheetProperties": {
        "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 3, "frozenColumnCount": C.N_FIXED}},
        "fields": "gridProperties(frozenRowCount,frozenColumnCount)"}})

    money_cols = [3]
    date_cols = []
    for pos in range(len(months)):
        start = C.block_start_col(pos) - 1
        date_cols += [start + 0, start + 2]
        money_cols += [start + 1, start + 3, start + 4]
    for c in money_cols:
        reqs.append({"repeatCell": {
            "range": rng(3, c, last_row, c + 1),
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}},
            "fields": "userEnteredFormat.numberFormat"}})
    for c in date_cols:
        reqs.append({"repeatCell": {
            "range": rng(3, c, last_row, c + 1),
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "DATE", "pattern": "dd/mm/yyyy"},
                     "horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"}})
    reqs.append({"repeatCell": {
        "range": rng(3, 1, last_row, 2),
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "TEXT"}, "horizontalAlignment": "CENTER"}},
        "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"}})

    def w(c0, c1, px):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": c0, "endIndex": c1},
            "properties": {"pixelSize": px}, "fields": "pixelSize"}})
    w(0, 1, 34); w(1, 2, 58); w(2, 3, 230); w(3, 4, 100); w(C.N_FIXED, n_cols, 80)

    book.batch_update({"requests": reqs})


# ---------------------------------------------------------------- baca
def read_roster_from_tab(ws):
    values = ws.get_all_values()
    out = []
    for r in range(C.FIRST_DATA_ROW, len(values) + 1):
        row = values[r - 1]
        induk = row[1].strip() if len(row) > 1 else ""
        nama = row[2].strip() if len(row) > 2 else ""
        if not induk and not nama:
            continue
        out.append({"row": r, "induk": induk, "nama": nama})
    return out


def list_years(book, jenjang="SMP"):
    years = set()
    for ws in book.worksheets():
        parts = ws.title.split()
        if len(parts) == 3 and parts[0] == jenjang:
            try:
                years.add(int(parts[2]))
            except ValueError:
                pass
    return sorted(years)


def last_saldo_of_year(ws, year):
    months = C.months_for_year(year)
    saldo_col = C.block_start_col(len(months) - 1) + 4
    values = ws.get(f"A{C.FIRST_DATA_ROW}:{rowcol_to_a1(ws.row_count, saldo_col)}",
                    value_render_option="UNFORMATTED_VALUE")
    out = []
    for row in values:
        if len(row) < 3:
            continue
        induk = str(row[1]).strip() if len(row) > 1 else ""
        nama = str(row[2]).strip() if len(row) > 2 else ""
        if not induk and not nama:
            continue
        saldo = row[saldo_col - 1] if len(row) >= saldo_col else 0
        try:
            saldo = float(saldo) if saldo not in ("", None) else 0
        except (ValueError, TypeError):
            saldo = 0
        out.append((induk, nama, saldo))
    return out


def read_cell_month(ws, row, year, month_num):
    months = C.months_for_year(year)
    pos = months.index(month_num)
    start = C.block_start_col(pos)
    a1 = f"{rowcol_to_a1(row, start)}:{rowcol_to_a1(row, start + 4)}"
    vals = ws.get(a1, value_render_option="FORMATTED_VALUE")
    row_vals = (vals[0] if vals else []) + [""] * 5
    return dict(zip(["tgl_setor", "setor", "tgl_tarik", "tarik", "saldo"], row_vals[:5]))


# ---------------------------------------------------------------- tulis
def write_transaction(ws, row, year, month_num, jenis, tanggal, jumlah):
    """Tulis transaksi setor/tarik. SALDO otomatis via formula. Return saldo baru."""
    months = C.months_for_year(year)
    if month_num not in months:
        raise ValueError(f"Bulan {month_num} tidak ada di T.A. {year}.")
    pos = months.index(month_num)
    cols = C.block_cols(pos)
    if isinstance(tanggal, (dt.date, dt.datetime)):
        tanggal = tanggal.strftime("%d/%m/%Y")

    if str(jenis).upper().startswith("PENYETORAN"):
        tgl_col, amt_col = cols["tgl_setor"], cols["setor"]
    else:
        tgl_col, amt_col = cols["tgl_tarik"], cols["tarik"]

    ws.batch_update([
        {"range": f"{tgl_col}{row}", "values": [[tanggal]]},
        {"range": f"{amt_col}{row}", "values": [[jumlah]]},
    ], value_input_option="USER_ENTERED")

    val = ws.acell(f"{cols['saldo']}{row}", value_render_option="UNFORMATTED_VALUE").value
    try:
        return float(val)
    except (ValueError, TypeError):
        return val
