"""
Helper Google Sheets (gspread) untuk menu Laporan Keuangan SD.

Kredensial & koneksi spreadsheet di-reuse langsung dari tab_sheet.py (sama
service account, sama pola pencarian kredensial) supaya tidak duplikasi.

ALUR KERJA (dirombak agar tata letaknya persis file contoh .xls):

  _load_outline(ws)  -> baca sheet jadi daftar baris murni (level, label,
                        kolom finansial). Baris "Jumlah Biaya" DIBUANG saat
                        baca karena selalu dibuat ulang.
  _save_outline(ws, outline)
                     -> susun ulang SELURUH grid di memori (nomor bertingkat,
                        penempatan teks per level, baris subtotal, TTL/SUB),
                        lalu tulis sekali jalan.

insert/update/delete semuanya = load → ubah daftar di memori → save. Jadi
tidak ada lagi insert_row/delete_rows per baris ke Google (versi lama butuh
~16 panggilan API tiap edit sampai sering timeout); sekarang tiap operasi
tetap 1 baca + 2 tulis + 1 format, berapa pun jumlah barisnya.

Kenapa 2 tulis, bukan 1: kolom nomor (B/C/E/F) harus ditulis RAW supaya "1.10"
tidak dipelintir Sheets jadi angka 1.1, sedangkan kolom formula & tanggal
butuh USER_ENTERED. Satu panggilan values.batchUpdate cuma menerima satu
valueInputOption, jadi dipisah dua.
"""
import gspread
from gspread.utils import rowcol_to_a1

import lk_config as C
import tab_sheet as TS_BASE  # reuse kredensial & open_book

get_client = TS_BASE.get_client


def open_book(spreadsheet_id=None, client=None):
    return TS_BASE.open_book(spreadsheet_id or C.SPREADSHEET_ID, client=client)


def _col_a1(col):
    return rowcol_to_a1(1, col)[:-1]


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
    struktur (level + nama) saja, kolom finansial dikosongkan. Kalau tidak
    diisi, pakai C.DEFAULT_TEMPLATE."""
    title = C.tab_name(bulan_num, tahun)
    try:
        return book.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        pass

    if clone_from is not None:
        prev, _extent = _load_outline(clone_from)
        outline = [{"level": r["level"], "label": r["label"]} for r in prev]
    else:
        outline = [{"level": lv, "label": lb} for lv, lb in C.DEFAULT_TEMPLATE]

    n_rows = max(C.FIRST_DATA_ROW + len(outline) + 20, FORMAT_ROW_BOUND)
    ws = book.add_worksheet(title=title, rows=n_rows, cols=C.N_COLS)

    _write_heading(ws, bulan_num, tahun)
    _save_outline(ws, outline, prev_extent=0)
    _format_static(ws)
    return ws


def _write_heading(ws, bulan_num, tahun):
    """Tulis 3 baris judul + header 2 tingkat (baris 1..6)."""
    grid = [["" for _ in range(C.N_COLS)] for _ in range(C.HEADER_ROW2)]
    grid[C.TITLE_ROW1 - 1][C.COL_PROG_NO - 1] = "LAPORAN PERTANGGUNGJAWABAN OPERASIONAL SD INSAN AMANAH"
    grid[C.TITLE_ROW2 - 1][C.COL_PROG_NO - 1] = f"TAHUN PELAJARAN {C.academic_label(bulan_num, tahun)}"
    grid[C.TITLE_ROW3 - 1][C.COL_PROG_NO - 1] = f"BULAN {C.tab_name(bulan_num, tahun)}"
    for text, col1, _col2, row1, _row2 in C.HEADER_CELLS:
        grid[row1 - 1][col1 - 1] = text
    last = _col_a1(C.N_COLS)
    ws.update(grid, f"A1:{last}{C.HEADER_ROW2}", value_input_option="RAW")


# ---------------------------------------------------------------- baca
def _load_outline(ws):
    """Baca sheet -> (daftar dict {row, level, label, tanggal, ...}, extent).
    `extent` = nomor baris terakhir yang masih berisi data; dipakai
    _save_outline untuk mengosongkan sisa baris kalau grid menyusut.
    Baris subtotal dibuang (selalu dibuat ulang oleh _save_outline).
    Dibaca UNFORMATTED supaya angka tetap angka & tanggal tetap serial —
    aman untuk ditulis balik tanpa kehilangan presisi."""
    last = _col_a1(C.N_COLS)
    values = ws.get(f"A{C.FIRST_DATA_ROW}:{last}{ws.row_count}",
                    value_render_option="UNFORMATTED_VALUE")
    extent = C.FIRST_DATA_ROW + len(values) - 1
    out = []
    prev_shallow = None
    for i, raw in enumerate(values):
        row = list(raw) + [""] * (C.N_COLS - len(raw))
        level = _row_level(row)
        if level is None or level == C.LEVEL_SUBTOTAL:
            continue
        label = _read_label(row, level, prev_shallow)
        if level < 5:
            prev_shallow = level
        entry = {"row": C.FIRST_DATA_ROW + i, "level": level, "label": label}
        for key, col in C.FIN_COLS.items():
            entry[key] = row[col - 1]
        out.append(entry)
    return out, extent


def _as_level(v):
    if v in ("", None):
        return None
    try:
        return int(float(v))
    except (ValueError, TypeError):
        return None


def _row_level(row):
    """Level sebuah baris. None kalau baris itu memang kosong.

    Baris yang diketik LANGSUNG di Google Sheet tidak punya penanda LEVEL di
    kolom R. Baris seperti itu TIDAK BOLEH diabaikan: _save_outline menulis
    ulang seluruh area data, jadi baris yang tidak terbaca di sini akan hilang
    permanen. Selama masih ada teksnya, baris itu diadopsi sebagai Rincian
    (level 5) mengikuti kolom teks tempat ia diketik."""
    level = _as_level(row[C.COL_LEVEL - 1])
    if level is not None:
        return level
    for col in C.LABEL_COLS:
        if str(row[col - 1]).strip():
            return 5
    return None


def _read_label(row, level, parent_level):
    if level == 5:
        # Rincian bisa di kolom F atau G tergantung induknya — ambil yang terisi.
        for col in (C.COL_ITEM_NAME, C.COL_KEG_NAME):
            v = row[col - 1]
            if str(v).strip():
                return str(v).strip()
        return ""
    return str(row[C.label_col(level, parent_level) - 1]).strip()


def read_rows(ws):
    """Versi untuk ditampilkan di web app: nilai apa adanya (terformat),
    termasuk baris subtotal, plus nomor/huruf & TTL/SUB hasil hitungan."""
    values = ws.get_all_values()
    out = []
    prev_shallow = None
    for r in range(C.FIRST_DATA_ROW, len(values) + 1):
        row = values[r - 1] + [""] * C.N_COLS
        level = _row_level(row)
        if level is None:
            continue
        if level == C.LEVEL_SUBTOTAL:
            label = C.SUBTOTAL_LABEL
        else:
            label = _read_label(row, level, prev_shallow)
            if level < 5:
                prev_shallow = level
        d = {"row": r, "level": level, "label": label,
             "no": _first_filled(row, [C.COL_PROG_NO, C.COL_SUB_NO, C.COL_KEG_NO]),
             "item": str(row[C.COL_ITEM_LETTER - 1]).strip() if level == 4 else "",
             "ttl_sub": str(row[C.COL_TTL_SUB - 1]).strip()}
        for key, col in C.FIN_COLS.items():
            d[key] = str(row[col - 1]).strip()
        out.append(d)
    return out


def _first_filled(row, cols):
    for col in cols:
        v = str(row[col - 1]).strip()
        if v:
            return v
    return ""


def last_data_row(ws):
    rows = read_rows(ws)
    return rows[-1]["row"] if rows else C.FIRST_DATA_ROW - 1


# ---------------------------------------------------------------- susun grid
def _letter(n):
    """1->a, 2->b, ..., 26->z, 27->aa, ..."""
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(97 + rem) + s
    return s


def _plan(outline):
    """outline (tanpa subtotal) -> daftar baris final (sudah termasuk baris
    'Jumlah Biaya' di akhir tiap Program). Tiap elemen: dict siap tulis."""
    plan = []
    for entry in outline:
        if entry["level"] == 1 and plan:
            plan.append({"level": C.LEVEL_SUBTOTAL})
        plan.append(dict(entry))
    if plan:
        plan.append({"level": C.LEVEL_SUBTOTAL})
    return plan


def _build_grid(plan):
    """plan -> (grid_utama, grid_nomor).
    grid_utama : nilai untuk ditulis USER_ENTERED (finansial + formula),
                 kolom nomor & teks dikosongkan.
    grid_nomor : nilai kolom B..G (nomor + teks) untuk ditulis RAW.
    """
    total_a1 = _col_a1(C.COL_TOTAL)
    counters = [0] * 6
    prev_shallow = None
    rows_main, rows_num = [], []
    # nomor baris sheet untuk tiap elemen plan
    sheet_rows = [C.FIRST_DATA_ROW + i for i in range(len(plan))]
    prog_start = None

    for i, item in enumerate(plan):
        level = item["level"]
        main = ["" for _ in range(C.N_COLS)]
        num = ["" for _ in range(C.COL_ITEM_NAME - C.COL_PROG_NO + 1)]  # B..G

        def set_num(col, val):
            num[col - C.COL_PROG_NO] = val

        main[C.COL_LEVEL - 1] = level

        if level == C.LEVEL_SUBTOTAL:
            set_num(C.COL_SUB_NAME, C.SUBTOTAL_LABEL)
            if prog_start is not None and sheet_rows[i] > prog_start:
                main[C.COL_TOTAL - 1] = (f"=SUM({total_a1}{prog_start}:"
                                         f"{total_a1}{sheet_rows[i] - 1})")
            else:
                main[C.COL_TOTAL - 1] = 0
            rows_main.append(main)
            rows_num.append(num)
            continue

        if level == 1:
            prog_start = sheet_rows[i]

        counters[level] += 1
        for d in range(level + 1, 6):
            counters[d] = 0

        ncol = C.number_col(level)
        if ncol:
            if level == 1:
                set_num(ncol, str(counters[1]))
            elif level == 2:
                set_num(ncol, f"{counters[1]}.{counters[2]}")
            elif level == 3:
                set_num(ncol, str(counters[3]))
            elif level == 4:
                set_num(ncol, _letter(counters[4]))

        set_num(C.label_col(level, prev_shallow), item.get("label", ""))
        if level < 5:
            prev_shallow = level

        for key, col in C.FIN_COLS.items():
            v = item.get(key, "")
            main[col - 1] = "" if v is None else v

        rows_main.append(main)
        rows_num.append(num)

    # TTL/SUB: SUM atas TOTAL milik baris-baris anak langsung di bawahnya
    for i, item in enumerate(plan):
        if item["level"] == C.LEVEL_SUBTOTAL:
            continue
        j = i + 1
        while (j < len(plan) and plan[j]["level"] != C.LEVEL_SUBTOTAL
               and plan[j]["level"] > item["level"]):
            j += 1
        if j > i + 1:
            rows_main[i][C.COL_TTL_SUB - 1] = (f"=SUM({total_a1}{sheet_rows[i + 1]}:"
                                               f"{total_a1}{sheet_rows[j - 1]})")
    return rows_main, rows_num


# ---------------------------------------------------------------- tulis
def _save_outline(ws, outline, prev_extent):
    """Susun ulang seluruh area data lalu tulis. Lihat catatan 2-tulis di atas.
    prev_extent = baris terakhir yang tadinya berisi data, dipakai untuk
    mengosongkan sisa baris kalau jumlah baris menyusut."""
    plan = _plan(outline)
    rows_main, rows_num = _build_grid(plan)

    need = C.FIRST_DATA_ROW + len(plan) - 1
    if need > ws.row_count:
        ws.add_rows(need - ws.row_count)

    # padding supaya sisa baris lama (kalau grid menyusut) ikut dikosongkan
    pad = max(0, min(prev_extent, ws.row_count) - need)
    rows_main = rows_main + [["" for _ in range(C.N_COLS)] for _ in range(pad)]
    rows_num = rows_num + [["" for _ in range(C.COL_ITEM_NAME - C.COL_PROG_NO + 1)]
                           for _ in range(pad)]

    end_row = C.FIRST_DATA_ROW + len(rows_main) - 1
    last_col = _col_a1(C.N_COLS)
    ws.update(rows_main, f"A{C.FIRST_DATA_ROW}:{last_col}{end_row}",
              value_input_option="USER_ENTERED")
    ws.update(rows_num,
              f"{_col_a1(C.COL_PROG_NO)}{C.FIRST_DATA_ROW}:"
              f"{_col_a1(C.COL_ITEM_NAME)}{end_row}",
              value_input_option="RAW")
    _format_dynamic(ws, plan)


def _find_index(outline, row_idx):
    for i, e in enumerate(outline):
        if e["row"] == row_idx:
            return i
    return None


def _entry_from(data):
    entry = {"level": data.get("level", 1), "label": data.get("label", "")}
    for key in C.FIN_COLS:
        entry[key] = data.get(key, "")
    return entry


def insert_row(ws, after_row_idx, data):
    """Sisipkan baris outline setelah baris sheet `after_row_idx`
    (None/0 = tambah di akhir)."""
    outline, extent = _load_outline(ws)
    pos = len(outline)
    if after_row_idx:
        i = _find_index(outline, after_row_idx)
        if i is not None:
            pos = i + 1
        else:
            # baris acuan adalah baris subtotal: sisipkan sesudah baris
            # outline terakhir yang berada di atasnya.
            pos = sum(1 for e in outline if e["row"] <= after_row_idx)
    outline.insert(pos, _entry_from(data))
    _save_outline(ws, outline, extent)
    return C.FIRST_DATA_ROW + pos


def update_row(ws, row_idx, data):
    outline, extent = _load_outline(ws)
    i = _find_index(outline, row_idx)
    if i is None:
        raise ValueError("Baris tidak ditemukan / bukan baris yang bisa diedit.")
    outline[i] = _entry_from(data)
    _save_outline(ws, outline, extent)


def delete_row(ws, row_idx):
    outline, extent = _load_outline(ws)
    i = _find_index(outline, row_idx)
    if i is None:
        raise ValueError("Baris tidak ditemukan / bukan baris yang bisa dihapus.")
    outline.pop(i)
    _save_outline(ws, outline, extent)


def resync(ws):
    outline, extent = _load_outline(ws)
    _save_outline(ws, outline, extent)


# ---------------------------------------------------------------- format
def _rgb(r, g, b):
    return {"red": r / 255, "green": g / 255, "blue": b / 255}


# Warna diambil PERSIS dari file contoh asli lewat xlrd formatting_info.
TITLE_BG = _rgb(255, 153, 0)      # oranye — judul
HEADER_BG = _rgb(0, 255, 0)       # hijau terang — header kolom
BLACK = _rgb(0, 0, 0)
BAND_A = _rgb(255, 153, 204)      # pink   — Program ke-1, 4, 7 ...
BAND_B = _rgb(204, 255, 255)      # tosca  — Program ke-2, 5, 8 ...
BAND_C = _rgb(255, 255, 204)      # kuning — Program ke-3, 6 ...
SUBTOTAL_BG = _rgb(153, 204, 255)  # biru muda — baris "Jumlah Biaya"
FORMAT_ROW_BOUND = 1000            # headroom baris untuk pertumbuhan data


def _font(name_size, bold=False, italic=False, color=None):
    name, size = name_size
    tf = {"fontFamily": name, "fontSize": size, "bold": bold, "italic": italic}
    if color:
        tf["foregroundColor"] = color
    return tf


def _format_dynamic(ws, plan):
    """Format yang harus mengikuti posisi baris: font Arial Black untuk baris
    Program. Conditional formatting hanya bisa mengatur tebal/miring/warna —
    tidak bisa ganti jenis font — jadi jenis font dipasang statis di sini dan
    disegarkan tiap kali susunan baris berubah."""
    sid = ws.id
    n_cols = C.N_COLS
    first = C.FIRST_DATA_ROW - 1
    last = max(ws.row_count, FORMAT_ROW_BOUND)

    def rng(r1, c1, r2, c2):
        return {"sheetId": sid, "startRowIndex": r1, "endRowIndex": r2,
                "startColumnIndex": c1, "endColumnIndex": c2}

    reqs = [{"repeatCell": {
        "range": rng(first, 0, last, n_cols),
        "cell": {"userEnteredFormat": {"textFormat": _font(C.FONT_BODY)}},
        "fields": "userEnteredFormat.textFormat(fontFamily,fontSize,bold,italic)"}}]
    for i, item in enumerate(plan):
        if item["level"] != 1:
            continue
        r0 = C.FIRST_DATA_ROW - 1 + i
        reqs.append({"repeatCell": {
            "range": rng(r0, 0, r0 + 1, n_cols),
            "cell": {"userEnteredFormat": {"textFormat": _font(C.FONT_PROGRAM, bold=True)}},
            "fields": "userEnteredFormat.textFormat(fontFamily,fontSize,bold,italic)"}})
    ws.spreadsheet.batch_update({"requests": reqs})


def _format_static(ws):
    """Format yang cukup dipasang sekali saat tab dibuat: judul, header 2
    tingkat, lebar kolom, format angka/tanggal, dan conditional formatting
    (warna per Program + tebal/miring per level) yang otomatis ikut baris baru."""
    sid = ws.id
    n_cols = C.N_COLS
    last_row = max(ws.row_count, FORMAT_ROW_BOUND)
    data_first0 = C.FIRST_DATA_ROW - 1
    lvl_a1 = _col_a1(C.COL_LEVEL)

    def rng(r1, c1, r2, c2):
        return {"sheetId": sid, "startRowIndex": r1, "endRowIndex": r2,
                "startColumnIndex": c1, "endColumnIndex": c2}

    reqs = []

    # --- judul: 3 merge horizontal TERPISAH (B..Q), seperti di file asli.
    # Bukan 1 merge 3-baris: merge vertikal membuang isi baris ke-2 & ke-3.
    for row in (C.TITLE_ROW1, C.TITLE_ROW2, C.TITLE_ROW3):
        reqs.append({"mergeCells": {
            "range": rng(row - 1, C.COL_PROG_NO - 1, row, C.COL_TTL_SUB),
            "mergeType": "MERGE_ALL"}})
        reqs.append({"repeatCell": {
            "range": rng(row - 1, C.COL_PROG_NO - 1, row, C.COL_TTL_SUB),
            "cell": {"userEnteredFormat": {
                "backgroundColor": TITLE_BG, "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
                "textFormat": _font(C.FONT_TITLE if row == C.TITLE_ROW1 else (C.FONT_TITLE[0], 14),
                                    bold=True, color=BLACK)}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)"}})

    # --- header 2 tingkat + merge grup
    reqs.append({"repeatCell": {
        "range": rng(C.HEADER_ROW1 - 1, 0, C.HEADER_ROW2, n_cols),
        "cell": {"userEnteredFormat": {
            "backgroundColor": HEADER_BG, "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE", "wrapStrategy": "WRAP",
            "textFormat": _font(C.FONT_HEADER, bold=True, color=BLACK)}},
        "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,wrapStrategy,textFormat)"}})
    for _text, col1, col2, row1, row2 in C.HEADER_CELLS:
        if col1 == col2 and row1 == row2:
            continue
        reqs.append({"mergeCells": {
            "range": rng(row1 - 1, col1 - 1, row2, col2), "mergeType": "MERGE_ALL"}})

    reqs.append({"updateSheetProperties": {
        "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": C.HEADER_ROW2}},
        "fields": "gridProperties.frozenRowCount"}})

    # --- rata tengah kolom nomor, format angka & tanggal
    for col in (C.COL_PROG_NO, C.COL_SUB_NO, C.COL_KEG_NO, C.COL_ITEM_LETTER):
        reqs.append({"repeatCell": {
            "range": rng(data_first0, col - 1, last_row, col),
            "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat.horizontalAlignment"}})
    for col in C.NUMERIC_COLS:
        reqs.append({"repeatCell": {
            "range": rng(data_first0, col - 1, last_row, col),
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}},
            "fields": "userEnteredFormat.numberFormat"}})
    reqs.append({"repeatCell": {
        "range": rng(data_first0, C.COL_TANGGAL - 1, last_row, C.COL_TANGGAL),
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "DATE", "pattern": "dd/mm/yyyy"},
                 "horizontalAlignment": "CENTER"}},
        "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"}})

    # --- conditional formatting: warna per Program & tebal/miring per level.
    # Dipasang dengan range absolut yang lega supaya otomatis berlaku juga
    # untuk baris yang ditambahkan user nanti.
    data_range = rng(data_first0, 0, last_row, n_cols)
    anchor = C.FIRST_DATA_ROW

    def cond(formula, fmt):
        return {"addConditionalFormatRule": {"rule": {
            "ranges": [data_range],
            "booleanRule": {
                "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": formula}]},
                "format": fmt}}, "index": 0}}

    count = f"COUNTIF(${lvl_a1}${anchor}:{lvl_a1}{anchor},1)"
    band = f"MOD({count}-1,3)"
    reqs.append(cond(f"={band}=0", {"backgroundColor": BAND_A}))
    reqs.append(cond(f"={band}=1", {"backgroundColor": BAND_B}))
    reqs.append(cond(f"={band}=2", {"backgroundColor": BAND_C}))
    # Level 1 tidak diberi rule tebal di sini: fontnya (Arial Black bold)
    # sudah dipasang statis oleh _format_dynamic().
    reqs.append(cond(f"=${lvl_a1}{anchor}=2", {"textFormat": {"bold": True, "italic": True}}))
    reqs.append(cond(f"=${lvl_a1}{anchor}=3", {"textFormat": {"bold": True}}))
    reqs.append(cond(f"=${lvl_a1}{anchor}=4", {"textFormat": {"bold": True}}))
    reqs.append(cond(f"=${lvl_a1}{anchor}={C.LEVEL_SUBTOTAL}",
                     {"backgroundColor": SUBTOTAL_BG, "textFormat": {"bold": True, "italic": True}}))

    # --- lebar kolom + sembunyikan kolom LEVEL
    for col, px in C.COL_WIDTHS_PX.items():
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS",
                      "startIndex": col - 1, "endIndex": col},
            "properties": {"pixelSize": px}, "fields": "pixelSize"}})
    reqs.append({"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "COLUMNS",
                  "startIndex": C.COL_LEVEL - 1, "endIndex": C.COL_LEVEL},
        "properties": {"hiddenByUser": True}, "fields": "hiddenByUser"}})

    ws.spreadsheet.batch_update({"requests": reqs})
