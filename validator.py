"""Validasi silang: master siswa (xlsx) vs master bank UPLDREQ (txt, fixed-width).

Tanpa pandas/numpy — baca xlsx langsung dengan openpyxl (read_only) supaya
tetap ringan, konsisten dengan gaya parser.py.
"""
import re
from io import BytesIO

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

_KELAS = re.compile(r"\b(I|II|III|IV|V|VI)\s+([A-F])\b")
_UPLDREQ_HEADER = re.compile(r"^0(\d{5})C(\d{2})(\d{2})(\d{4})")


def _kelas(nama):
    m = _KELAS.search(str(nama).upper())
    return f"{m.group(1)} {m.group(2)}" if m else ""


def _strip_kelas(nama):
    return re.sub(r"\s*\b(I|II|III|IV|V|VI)\s+[A-F]\b.*$", "", str(nama).strip()).strip()


def _to_int(v):
    try:
        return int(round(float(v)))
    except (TypeError, ValueError):
        return 0


def parse_master_siswa(file_bytes):
    """Bytes xlsx (skema NO VA | NAMA | BPP | KEGIATAN | TABUNGAN) -> dict {no_va: row}."""
    wb = load_workbook(BytesIO(file_bytes), data_only=True, read_only=True)
    ws = wb.active
    out = {}
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue  # header
        if not row or row[0] is None or row[1] is None:
            continue
        try:
            no_va = int(row[0])
        except (TypeError, ValueError):
            continue
        nama = str(row[1]).strip()
        bpp = _to_int(row[2]) if len(row) > 2 else 0
        kegiatan = _to_int(row[3]) if len(row) > 3 else 0
        tabungan = _to_int(row[4]) if len(row) > 4 else 0
        out[no_va] = {
            "no_va": no_va, "nama": nama, "nama_bersih": _strip_kelas(nama),
            "kelas": _kelas(nama), "BPP": bpp, "KEGIATAN": kegiatan, "TABUNGAN": tabungan,
        }
    wb.close()
    return out


def parse_uploadreq(text, scale=100):
    """Isi teks UPLDREQ_ddmmyyyyhhmmss.txt -> (meta, dict {no_va: row}).

    Layout (lihat CLAUDE.md project referensi, hasil reverse-engineering):
      pos 0        : tipe record ('0' header, '1' detail, '9' trailer)
      pos 6..14    : VA penuh (9) = biller(5) + NO VA(4) -> no_va = va[5:]
      pos 29..idr  : nama (dipotong ~30 char)
      setelah IDR  : blok nominal, tiap field 15 digit, nilai = angka // scale.
                     Urutan: TOTAL, BPP, KEGIATAN, TABUNGAN.
    """
    meta = {"biller": "", "tanggal_label": ""}
    out = {}
    for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if not line:
            continue
        if line[0] == "0":
            if not meta["biller"]:
                m = _UPLDREQ_HEADER.match(line)
                if m:
                    meta["biller"] = m.group(1)
                    meta["tanggal_label"] = f"{m.group(2)}/{m.group(3)}/{m.group(4)}"
            continue
        if line[0] != "1":
            continue
        idr = line.find("IDR")
        if idr < 0:
            continue
        try:
            va_full = line[6:15].strip()
            no_va = int(va_full[5:]) if len(va_full) > 5 else int(va_full)
            nama = line[29:idr - 8].strip()
            amt = line[idr + 3:]
            fields = [int(amt[i:i + 15]) // scale for i in range(0, 15 * 4, 15)]
        except (ValueError, IndexError):
            continue
        out[no_va] = {
            "no_va": no_va, "nama_bank": nama,
            "total": fields[0], "BPP": fields[1], "KEGIATAN": fields[2], "TABUNGAN": fields[3],
        }
    return meta, out


_KOMPONEN = ("BPP", "KEGIATAN", "TABUNGAN")

STATUS_SESUAI = "Sesuai"
STATUS_BEDA = "Beda Nominal"
STATUS_HANYA_MASTER = "Hanya di Master Siswa"
STATUS_HANYA_BANK = "Hanya di UPLDREQ"


def cross_validate(master, bank):
    """Gabungkan master siswa & master bank by NO VA -> list of dict siap-lapor."""
    rows = []
    for no_va in sorted(set(master) | set(bank)):
        m, b = master.get(no_va), bank.get(no_va)
        if m and b:
            beda = [k for k in _KOMPONEN if m[k] != b[k]]
            status = STATUS_SESUAI if not beda else STATUS_BEDA
            ket = "" if not beda else "Beda: " + ", ".join(beda)
            nama, kelas = m["nama_bersih"], m["kelas"]
        elif m:
            status, ket, nama, kelas = STATUS_HANYA_MASTER, "Tidak ditemukan di UPLDREQ", m["nama_bersih"], m["kelas"]
        else:
            status, ket, nama, kelas = STATUS_HANYA_BANK, "Tidak ditemukan di Master Siswa", b["nama_bank"], ""
        rows.append({
            "no_va": no_va, "nama": nama, "kelas": kelas,
            "bpp_m": m["BPP"] if m else None, "keg_m": m["KEGIATAN"] if m else None, "tab_m": m["TABUNGAN"] if m else None,
            "bpp_b": b["BPP"] if b else None, "keg_b": b["KEGIATAN"] if b else None, "tab_b": b["TABUNGAN"] if b else None,
            "status": status, "keterangan": ket,
        })
    return rows


# ---------- styling (selaras dengan parser.py) ----------
_A = "Arial"
_HFILL = PatternFill("solid", fgColor="1F4E5F")
_HFONT = Font(name=_A, bold=True, color="FFFFFF", size=10)
_TITLE = Font(name=_A, bold=True, size=14, color="1F4E5F")
_SUB = Font(name=_A, italic=True, size=9, color="595959")
_BASE = Font(name=_A, size=10)
_BOLD = Font(name=_A, bold=True, size=10)
_TFILL = PatternFill("solid", fgColor="D9E7EC")
_BAND = PatternFill("solid", fgColor="F2F7F9")
_WARN = PatternFill("solid", fgColor="FFF2CC")
_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_C = Alignment("center", vertical="center")
_L = Alignment("left", vertical="center")
_R = Alignment("right", vertical="center")
_HEADERS = ["No.", "NO VA", "Nama Siswa", "Kelas", "BPP (Master)", "Kegiatan (Master)",
            "Tabungan (Master)", "BPP (UPLDREQ)", "Kegiatan (UPLDREQ)", "Tabungan (UPLDREQ)",
            "Status", "Keterangan"]
_MONEY_COLS = {5, 6, 7, 8, 9, 10}


def build_validation_workbook(rows, bank_meta):
    wb = Workbook()
    ws = wb.active
    ws.title = "Validasi"
    ws.sheet_view.showGridLines = False
    ws["A1"] = "VALIDASI MASTER SISWA vs MASTER BANK (UPLDREQ)"
    ws["A1"].font = _TITLE
    ws["A2"] = (f"Biller {bank_meta.get('biller') or '—'}  •  "
                f"Tanggal jatuh tempo: {bank_meta.get('tanggal_label') or '—'}")
    ws["A2"].font = _SUB
    HROW = 4
    for c, h in enumerate(_HEADERS, 1):
        cell = ws.cell(HROW, c, h)
        cell.fill = _HFILL; cell.font = _HFONT; cell.alignment = _C; cell.border = _BORDER
    r0 = HROW + 1
    for i, row in enumerate(rows):
        r = r0 + i
        vals = [i + 1, row["no_va"], row["nama"], row["kelas"],
                row["bpp_m"], row["keg_m"], row["tab_m"],
                row["bpp_b"], row["keg_b"], row["tab_b"],
                row["status"], row["keterangan"]]
        warn = row["status"] != STATUS_SESUAI
        fill = _WARN if warn else (_BAND if i % 2 == 1 else None)
        for j, v in enumerate(vals, 1):
            cell = ws.cell(r, j, v)
            cell.font = _BASE; cell.border = _BORDER
            if fill:
                cell.fill = fill
            if j in _MONEY_COLS:
                cell.number_format = '#,##0;-#,##0;"-"'; cell.alignment = _R
            elif j in (1, 2, 4, 11):
                cell.alignment = _C
            else:
                cell.alignment = _L
    last = r0 + len(rows) - 1

    gt = last + 1
    ws.cell(gt, 3, "TOTAL").alignment = _L
    ws.cell(gt, 3).font = _BOLD
    for j in _MONEY_COLS:
        col = get_column_letter(j)
        cell = ws.cell(gt, j, f"=SUM({col}{r0}:{col}{last})")
        cell.number_format = '#,##0;-#,##0;"-"'; cell.alignment = _R
    for j in range(1, len(_HEADERS) + 1):
        cell = ws.cell(gt, j); cell.fill = _TFILL; cell.border = _BORDER; cell.font = _BOLD

    widths = [6, 10, 28, 8, 13, 15, 15, 14, 15, 15, 20, 26]
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A5"

    _build_ringkasan(wb, rows, bank_meta)
    return wb


def _build_ringkasan(wb, rows, bank_meta):
    rs = wb.create_sheet("Ringkasan", 0)
    rs.sheet_view.showGridLines = False
    rs["A1"] = "RINGKASAN VALIDASI"; rs["A1"].font = _TITLE
    from collections import Counter
    n = Counter(r["status"] for r in rows)
    info = [
        ("Biller", bank_meta.get("biller") or "—"),
        ("Tanggal jatuh tempo (UPLDREQ)", bank_meta.get("tanggal_label") or "—"),
        ("Total NO VA (gabungan)", len(rows)),
        (STATUS_SESUAI, n.get(STATUS_SESUAI, 0)),
        (STATUS_BEDA, n.get(STATUS_BEDA, 0)),
        (STATUS_HANYA_MASTER, n.get(STATUS_HANYA_MASTER, 0)),
        (STATUS_HANYA_BANK, n.get(STATUS_HANYA_BANK, 0)),
    ]
    r = 3
    for k, v in info:
        a = rs.cell(r, 1, k); b = rs.cell(r, 2, v)
        a.font = _BASE; b.font = _BOLD if isinstance(v, int) else _BASE
        a.fill = _BAND; b.fill = _BAND; a.border = _BORDER; b.border = _BORDER
        a.alignment = _L; b.alignment = _R if isinstance(v, int) else _L
        r += 1
    rs.column_dimensions["A"].width = 32
    rs.column_dimensions["B"].width = 24
