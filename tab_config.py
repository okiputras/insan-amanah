"""
Konfigurasi & helper Tabungan SMP Insan Amanah (dipakai menu Tabungan di app.py).

Struktur tiap tab Google Sheet ("SMP <kelas> <tahun-ajaran>"), per TAHUN AJARAN:
  Baris 1 : Judul     Baris 2 : Nama bulan (merge 5 kolom)     Baris 3 : Sub-header
  Baris 4+: Data siswa
  Kolom tetap:  A=NO  B=INDUK  C=NAMA LENGKAP  D=SALDO AWAL
  Per bulan (5 kolom): TGL SETOR | SETOR | TGL TARIK | TARIK | SALDO
  Bulan: Juli (start_year) .. Juni (start_year+1)
"""
from openpyxl.utils import get_column_letter

START_YEAR = 2026
START_MONTH = 7            # Juli = bulan pertama tahun ajaran

MONTHS_ID = [
    "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI",
    "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER",
]
KELAS_LIST = [7, 8, 9]          # SMP (default, kompatibilitas lama)
SD_KELAS_LIST = [1, 2, 3, 4, 5, 6]

FIXED_HEADERS = ["NO", "INDUK", "NAMA LENGKAP", "SALDO AWAL"]
N_FIXED = len(FIXED_HEADERS)
SUB_HEADERS = ["TGL SETOR", "SETOR", "TGL TARIK", "TARIK", "SALDO"]
COLS_PER_MONTH = len(SUB_HEADERS)
FIRST_MONTH_COL = N_FIXED + 1

HEADER_MONTH_ROW = 2
HEADER_SUB_ROW = 3
FIRST_DATA_ROW = 4

SPREADSHEET_TITLE = "TABUNGAN SMP INSAN AMANAH"
SPREADSHEET_ID = "1SlO-nKeX31lgPMkHsZiUKcOxrmBYitUPsnX7mEDU-JM"       # SMP
SD_SPREADSHEET_ID = "1LsXzb4Ynsh3-Dg22wVAt-8eg0Fh26hk4zWjEQXKrZXg"   # SD
SERVICE_ACCOUNT_FILE = "sa-sheet.json"
SERVICE_ACCOUNT_EMAIL = "oki-gsheet@iconic-woods-355603.iam.gserviceaccount.com"

# Profil per jenjang (dipakai dashboard & skrip build)
PROFILES = {
    "smp": {"jenjang": "SMP", "kelas": KELAS_LIST, "spreadsheet_id": SPREADSHEET_ID,
            "title": "TABUNGAN SMP INSAN AMANAH"},
    "sd": {"jenjang": "SD", "kelas": SD_KELAS_LIST, "spreadsheet_id": SD_SPREADSHEET_ID,
           "title": "TABUNGAN SD INSAN AMANAH"},
}


def tab_name(kelas, start_year, jenjang="SMP"):
    return f"{jenjang} {kelas} {start_year}"


def academic_label(start_year):
    return f"{start_year}/{start_year + 1}"


def months_for_year(start_year):
    if start_year < START_YEAR:
        return []
    return list(range(START_MONTH, 13)) + list(range(1, START_MONTH))  # 7..12, 1..6


def month_cal_year(month_num, start_year):
    return start_year if month_num >= START_MONTH else start_year + 1


def month_label(month_num, start_year):
    return f"{MONTHS_ID[month_num - 1]} {month_cal_year(month_num, start_year)}"


def block_start_col(month_pos):
    return FIRST_MONTH_COL + month_pos * COLS_PER_MONTH


def block_cols(month_pos):
    c = block_start_col(month_pos)
    keys = ["tgl_setor", "setor", "tgl_tarik", "tarik", "saldo"]
    return {k: get_column_letter(c + i) for i, k in enumerate(keys)}


def saldo_formula(row, month_pos):
    cur = block_cols(month_pos)
    prev = f"D{row}" if month_pos == 0 else f"{block_cols(month_pos - 1)['saldo']}{row}"
    return f"=N({prev})+N({cur['setor']}{row})-N({cur['tarik']}{row})"


def total_cols(start_year):
    return N_FIXED + len(months_for_year(start_year)) * COLS_PER_MONTH
