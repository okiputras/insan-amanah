"""
Skrip admin: backfill tab TAHUN AJARAN LAMA pada spreadsheet Tabungan.

    python3 tab_backfill.py smp 2024 2025      # buat T.A. 2024/2025 & 2025/2026
    python3 tab_backfill.py sd 2024 2025

Tab dibuat KOSONG (struktur bulan Juli..Juni + roster siswa), siap diisi lewat
menu Tabungan di aplikasi. Roster disalin dari tahun yang sudah ada di
spreadsheet (tahun terlama), jadi daftar siswanya sama persis.

SALDO AWAL antar tahun BARU dirantai memakai formula VLOOKUP ke tab tahun
sebelumnya — bukan angka hasil snapshot. Ini penting untuk backfill: tabnya
dibuat kosong sekarang, transaksinya diisi belakangan, dan saldo awal tahun
berikutnya harus ikut menyesuaikan sendiri. Tahun terlama dapat SALDO AWAL 0
karena tidak ada tahun sebelumnya.

Tab tahun yang SUDAH ADA sama sekali tidak disentuh secara default.

    --link-next   sambungkan juga SALDO AWAL tahun lama pertama yang sudah ada
                  ke tab baru terakhir.

BAHAYA — baca sebelum memakai --link-next: kolom SALDO AWAL tahun berjalan
biasanya berisi saldo akumulasi nyata dari sebelum sistem ini dipakai. Opsi
ini MENIMPA angka itu dengan formula ke tab tahun baru yang (kalau belum
diisi) bernilai nol, sehingga saldo nyata tersebut hilang. Pakai hanya kalau
transaksi tahun-tahun lama sudah selesai diinput sehingga rantainya
menghasilkan angka yang sama.
"""
import sys

import tab_config as C
import tab_sheet as SL


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    link_next = "--link-next" in sys.argv
    pkey = args[0].lower()
    if pkey not in C.PROFILES:
        print("Jenjang harus salah satu dari:", ", ".join(C.PROFILES))
        sys.exit(1)
    years = sorted({int(a) for a in args[1:]})

    prof = C.PROFILES[pkey]
    jenjang, kelas_list = prof["jenjang"], prof["kelas"]
    book = SL.open_book(prof["spreadsheet_id"])
    print("Terhubung:", book.title)

    existing = SL.list_years(book, jenjang)
    if not existing:
        print("Belum ada tab sama sekali — jalankan skrip build dulu.")
        sys.exit(1)
    print("Tahun yang sudah ada:", existing)

    bentrok = [y for y in years if y in existing]
    if bentrok:
        print("Tahun berikut sudah ada, dibatalkan agar datanya tidak tertimpa:", bentrok)
        sys.exit(1)

    src_year = existing[0]
    sep = SL.formula_arg_sep(book)
    print(f"Roster disalin dari T.A. {src_year}/{src_year + 1}")
    print(f"Pemisah argumen formula: '{sep}' (mengikuti locale spreadsheet)")

    for kelas in kelas_list:
        try:
            src = book.worksheet(C.tab_name(kelas, src_year, jenjang))
        except Exception:
            print(f"  ! tab sumber kelas {kelas} tidak ada, dilewati")
            continue
        names = [(r["induk"], r["nama"]) for r in SL.read_roster_from_tab(src)]

        prev_year = None
        for year in years:
            title = C.tab_name(kelas, year, jenjang)
            if prev_year is None:
                roster = [(induk, nama, 0) for induk, nama in names]
            else:
                prev_title = C.tab_name(kelas, prev_year, jenjang)
                roster = [
                    (induk, nama,
                     SL.saldo_awal_formula(prev_title, prev_year,
                                           C.FIRST_DATA_ROW + i, sep))
                    for i, (induk, nama) in enumerate(names)
                ]
            SL.build_tab(book, kelas, year, roster, jenjang)
            print(f"  ✓ {title}: {len(roster)} siswa")
            prev_year = year

        # Menyambungkan tahun yang sudah ada = MENIMPA saldo awalnya. Hanya
        # dilakukan kalau diminta eksplisit (lihat peringatan di docstring).
        after = [y for y in existing if y > years[-1]]
        if link_next and after and prev_year is not None:
            nxt = after[0]
            try:
                ws_next = book.worksheet(C.tab_name(kelas, nxt, jenjang))
            except Exception:
                continue
            n = SL.link_saldo_awal(ws_next, C.tab_name(kelas, prev_year, jenjang),
                                   prev_year, sep)
            print(f"    ↳ SALDO AWAL {C.tab_name(kelas, nxt, jenjang)} disambungkan "
                  f"ke {C.tab_name(kelas, prev_year, jenjang)} ({n} siswa)")

    # Pastikan formula benar-benar ter-parse. Formula yang gagal di-parse tidak
    # menggagalkan penulisan — sel-nya cuma berisi #ERROR! dan menular ke kolom
    # SALDO, jadi harus dicek eksplisit di sini.
    bad = []
    for kelas in kelas_list:
        for year in years[1:]:
            try:
                ws = book.worksheet(C.tab_name(kelas, year, jenjang))
            except Exception:
                continue
            vals = ws.get(f"D{C.FIRST_DATA_ROW}:D{C.FIRST_DATA_ROW + 4}",
                          value_render_option="FORMATTED_VALUE")
            if any(str(v).startswith("#") for row in vals for v in row):
                bad.append(C.tab_name(kelas, year, jenjang))
    if bad:
        print()
        print("!! SALDO AWAL bermasalah (#ERROR) di:", ", ".join(bad))
        print("   Formula gagal di-parse — cek pemisah argumen vs locale spreadsheet.")
        sys.exit(2)
    print("Cek SALDO AWAL: semua formula ter-parse dengan benar.")

    sisa = [y for y in existing if y > years[-1]]
    if sisa and not link_next:
        print()
        print(f"Catatan: SALDO AWAL T.A. {sisa[0]}/{sisa[0] + 1} dibiarkan apa adanya "
              f"(tidak disambungkan ke {years[-1]}).")
        print("Jadi saldo tahun berjalan TIDAK ikut berubah oleh backfill ini. Setelah")
        print("transaksi tahun lama selesai diinput dan angkanya sudah cocok, sambungkan")
        print(f"dengan: python3 tab_backfill.py {pkey} --link-next")

    print("Selesai:", f"https://docs.google.com/spreadsheets/d/{prof['spreadsheet_id']}")


if __name__ == "__main__":
    main()
