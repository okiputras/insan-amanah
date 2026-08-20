"""
Konfigurasi & helper Laporan Keuangan SD Insan Amanah (dipakai menu Laporan
Keuangan di app.py).

Struktur tiap tab Google Sheet ("<BULAN> <TAHUN>"), per BULAN LAPORAN:
  Baris 1 : Judul     Baris 2 : Header kolom     Baris 3+: baris outline
  Kolom   : A=LEVEL B=LABEL C=TANGGAL D=KODE E=VOLUME F=SATUAN G=FK
            H=KEBUTUHAN I=UNIT_COST J=TOTAL K=RENCANA L=SPP M=BSM N=BOSDA O=BOSNAS

LEVEL menentukan indentasi tampilan (1=Program .. 5=Rincian) tapi kolom
finansial (TANGGAL..BOSNAS) tersedia bebas di semua level — tidak divalidasi
per level, supaya user bisa isi sesuai kebutuhan nyata di lapangan.
"""
from tab_config import MONTHS_ID  # reuse, jangan duplikat

SPREADSHEET_TITLE = "LAPORAN KEUANGAN SD INSAN AMANAH"
SPREADSHEET_ID = "1tVrKbIfxuVFzesZwh2p9a1xj3TTj872YwYpd6-uG_OU"
SERVICE_ACCOUNT_FILE = "sa-sheet.json"          # reuse via tab_sheet
SERVICE_ACCOUNT_EMAIL = "oki-gsheet@iconic-woods-355603.iam.gserviceaccount.com"

LEVEL_LABELS = ["Program", "Sub Program", "Kegiatan", "Item", "Rincian"]

FIELD_KEYS = ["level", "label", "tanggal", "kode", "volume", "satuan", "fk",
              "kebutuhan", "unit_cost", "total", "rencana", "spp", "bsm",
              "bosda", "bosnas"]
HEADERS = ["LEVEL", "LABEL", "TANGGAL", "KODE", "VOLUME", "SATUAN", "FK",
           "KEBUTUHAN", "UNIT COST", "TOTAL", "RENCANA", "SPP", "BSM",
           "BOSDA", "BOSNAS"]
N_COLS = len(FIELD_KEYS)

TITLE_ROW = 1
HEADER_ROW = 2
FIRST_DATA_ROW = 3

# Kolom yang diperlakukan sebagai angka (untuk format & parsing)
NUMERIC_KEYS = {"volume", "unit_cost", "total", "rencana", "spp", "bsm",
                 "bosda", "bosnas"}


def tab_name(bulan_num, tahun):
    return f"{MONTHS_ID[bulan_num - 1]} {tahun}"


def indent_px(level):
    return max(0, (int(level) - 1)) * 22
