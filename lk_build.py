"""
Skrip admin (sekali jalan): isi tab bulan pertama di Google Sheet Laporan
Keuangan SD. Spreadsheet-nya harus SUDAH dibuat manual di Google Drive dan
di-share ke service account (Editor) — sama seperti Tabungan SMP/SD, service
account tidak punya kuota Drive sendiri jadi tidak bisa membuat spreadsheet
baru lewat API.

Jalankan:
    python3 lk_build.py <SPREADSHEET_ID>
"""
import sys
from datetime import datetime

import lk_sheet as LS
import tab_config as TC


def main():
    if len(sys.argv) < 2:
        print("Pakai: python3 lk_build.py <SPREADSHEET_ID>")
        sys.exit(1)
    sid = sys.argv[1]

    book = LS.open_book(sid)
    print("Terhubung:", book.title)

    now = datetime.now()
    bulan_num, tahun = now.month, now.year
    LS.ensure_bulan_tab(book, bulan_num, tahun)
    print(f"  ✓ Tab {TC.MONTHS_ID[bulan_num - 1]} {tahun} dibuat")

    for junk in ("Sheet1", "Sheet"):
        try:
            book.del_worksheet(book.worksheet(junk))
        except Exception:
            pass

    print()
    print("Selesai:", f"https://docs.google.com/spreadsheets/d/{sid}")
    print(f"Tempel SPREADSHEET_ID = \"{sid}\" ke lk_config.SPREADSHEET_ID lalu restart app.")


if __name__ == "__main__":
    main()
