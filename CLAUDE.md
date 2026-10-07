# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running

```bash
pip install -r requirements.txt

python app.py                    # dev server, port 8501, debug on
gunicorn app:app --workers 1 --threads 4 --bind 0.0.0.0:8501   # as in production
```

Every page is behind HTTP Basic Auth (`AUTH_USER` / `AUTH_PASS` env, defaults in
`app.py`), so manual requests need credentials:

```bash
curl -u "$AUTH_USER:$AUTH_PASS" http://localhost:8501/tabungan
```

`/health` is the only open route (Railway healthcheck).

There is no test suite, no linter config, and no build step. Verification is done by
running the app and exercising a route, or by calling the module functions directly
against a scratch Google Sheet tab.

## Google Sheets credentials

The Tabungan and Laporan Keuangan menus talk to Google Sheets through a service
account. `tab_sheet._load_creds()` resolves it in this order:

1. `GOOGLE_SERVICE_ACCOUNT_JSON` — full JSON in the env var (this is what Railway uses)
2. `SA_FILE` — path to a json file
3. `sa-sheet.json`, searched in the app folder and up to three levels above it

Locally the key lives outside the repo, so run these menus with:

```bash
SA_FILE="/path/to/sa-sheet.json" python app.py
```

Target spreadsheets are hardcoded (`tab_config.SPREADSHEET_ID`,
`SD_SPREADSHEET_ID`, `lk_config.SPREADSHEET_ID`) and must be shared as Editor with
`tab_config.SERVICE_ACCOUNT_EMAIL`. The service account has no Drive storage quota of
its own, so it **cannot create a spreadsheet** — new ones are made by hand in Drive,
shared, then their id pasted into the config.

## Architecture

Flask app, single process, no database. Two unrelated halves:

**File-conversion menus** (`/`, `/validasi/<level>`) are stateless. Upload → parse in
memory → build an `.xlsx` with openpyxl → hand back a download link. The generated
bytes sit in `app._STORE`, an in-memory dict keyed by a random token with a 30-minute
TTL and a 20-entry cap. This is why the deploy is pinned to `--workers 1`: a second
worker would not see the token, and the links break on every restart.

**Google-Sheets menus** (`/tabungan`, `/tabungan-sd`, `/laporan-keuangan`) keep no
local state — the spreadsheet is the database, and every request reads it fresh.
`tab_*` and `lk_*` are imported with a `try/except` in `app.py`, so if gspread or the
credentials are missing the other menus still work and only these pages show an error.

### Menu layout

| Route | Module | Storage |
|---|---|---|
| `/` Konversi R-5401 | `parser.py` | in-memory token |
| `/validasi/sd`, `/validasi/smp` | `validator.py` + `parser.py` | in-memory token |
| `/tabungan`, `/tabungan-sd` | `tab_config.py`, `tab_sheet.py` | Google Sheets |
| `/laporan-keuangan` | `lk_config.py`, `lk_sheet.py` | Google Sheets |
| `/pembayaran` | `bayar_config.py`, `bayar_sheet.py` | Google Sheets |

All HTML lives in `app.py` as module-level strings rendered with
`render_template_string` — `PAGE`, `REKAP_PAGE`, `TABUNGAN_PAGE`,
`LAPORAN_KEUANGAN_PAGE`, `BAYAR_PAGE`, sharing `STYLE` plus per-page `TAB_EXTRA` /
`LK_EXTRA` / `BAYAR_EXTRA`. The nav bar is duplicated verbatim in each; adding a menu
means editing all five.

### Report parsing

`parser.parse_laporan()` sniffs the first 8 lines and dispatches to `parse_report()`
(R-5401, fixed-width — column offsets are the `SL_*` constants) or `parse_bca_va()`
(BCA Virtual Account, regex). Both return the same `(meta, rows)` shape so
`validator.reconcile_pembayaran()` consumes either.

`validator` joins those rows to the master student xlsx. **Master columns are read by
position, not by header name** (col 1 student no, 2 name, 3 BPP, 4 kegiatan, 5
tabungan), which is why real files with headers like `NO INDUK`/`KEG`/`TAB` parse
fine. The join key differs per level and is resolved by trying candidates against the
master rather than being fixed up front (`_no_va_candidates` → `_match_master`):

- SD: the report's No. Pelanggan already is the full NO VA.
- SMP: the report carries a short 4-digit code, so the school code is prefixed
  (`63713` + `0318` = `637130318`). Master files exist in both shapes — full NO VA and
  short student number — so both keys are tried. Committing to one made every
  transaction miss, which surfaced as "all transactions have a discrepancy".

### Tabungan (`tab_*`)

One spreadsheet per level, one tab per class per academic year, named
`"<SMP|SD> <kelas> <tahun>"` — `tab_sheet.list_years()` discovers years by parsing tab
titles, so a tab's name *is* its identity. Layout is a fixed grid: columns A–D are
NO / INDUK / NAMA / SALDO AWAL, then 12 five-column month blocks (TGL SETOR, SETOR,
TGL TARIK, TARIK, SALDO) for July→June. `tab_config` owns the column arithmetic
(`block_start_col`, `saldo_formula`, `last_saldo_col`); do not hand-compute offsets.

SALDO runs as a live formula per month, and **SALDO AWAL is a cell reference into the
previous year's June SALDO**, not a copied number. Prior-year tabs get backfilled
empty and filled in later, so a snapshot taken at creation time would leave every
following year silently wrong. The reference is a plain `='SMP 9 2025'!BL27` on
purpose: student numbers are not unique in this data (0229 was shared by two
students), so a VLOOKUP on INDUK returns the wrong row, and the spreadsheet locale is
`in_ID` where comma argument separators fail to parse into `#ERROR!`. Row pairing
between years is resolved in Python (`match_rows`), never in the sheet.

Admin scripts, all one-shot and run by hand:

```bash
python3 tab_backfill.py smp 2024 2025      # create empty prior-year tabs
python3 tab_import.py smp "file.xlsx"      # load a recap workbook; --dry-run first
```

`tab_import` pairs source rows to sheet rows **by order**, verifying INDUK at each
position and aborting on a mismatch, because keying on INDUK is unsafe here. It reads
deposits from the monthly cells rather than the TOTAL column, since a TOTAL formula in
the source had been overwritten with a literal 0 and hid real deposits.

`tab_build.py` / `tab_build_sd.py` are dead: their source workbooks no longer exist,
and they still write snapshot balances.

### Laporan Keuangan (`lk_*`)

Deliberately unlike Tabungan: one tab per calendar month, and rows are a
variable-length outline (Program → Sub Program → Kegiatan → Item → Rincian) that the
user can insert and delete anywhere.

The sheet mirrors a school-issued .xls, so the layout is unusual and `lk_config` is
the single source of truth for it. There is no one LABEL column — the text column
shifts with the level (Program C, Sub Program D, Kegiatan F, Item G) and that stepping
is what makes it look like the original. Numbering is spread across separate columns
too (`COL_PROG_NO`, `COL_SUB_NO`, `COL_KEG_NO`, `COL_ITEM_LETTER`). Column R holds the
level as a hidden machine-readable marker; it is what the reader keys on.

Every mutation is **load → edit the list in memory → rewrite the whole data area**
(`_load_outline` → `_plan` → `_build_grid` → `_save_outline`). The earlier per-row
insert/delete approach issued ~16 API calls per edit and timed out; this is 1 read +
2 writes + 1 format regardless of size. Two writes because the numbering columns need
`RAW` (so `"1.10"` is not mangled into `1.1`) while formulas and dates need
`USER_ENTERED`.

Because the writer replaces everything, anything the reader drops is destroyed. Rows
typed straight into Google Sheets carry no level marker, so `_row_level()` adopts them
as Rincian instead of skipping them. Keep that invariant when touching the reader.

`resync()` regenerates the "Jumlah Biaya" subtotal row per Program and the TTL/SUB
rollups; those rows are system-managed and the app refuses to edit or delete them.
Per-level styling and the alternating Program colours are conditional-formatting rules
installed once, so they follow rows added later without a reformat pass.

### Pembayaran SD (`bayar_*`)

Answers "who hasn't paid BPP + katering + kegiatan yet". Storage is a ledger, not a grid:
tab `PEMBAYARAN` holds one row per payment and is **only ever appended to** (so two
people recording at once can't clobber each other), tab `SISWA 2026` holds what each
student owes per month, tab `PERLU DICEK` holds imported rows whose student couldn't be
pinned down. Paid/unpaid status is never stored — `bayar_config.cek()` recomputes it from
the ledger on every request.

The computation is FIFO from July: a student's total paid is allocated to the oldest
month first. It has to be balance-based because parents routinely pay 2, 6 or 12 months
at once, pay arrears, or prepay; the month label on a source block is not trustworthy
either (an August block is labelled "JUNI"). Allocation runs on **two buckets** —
BPP+katering combined, and kegiatan — because in the source's 9-column layout the "pure
BPP" column is the formula `(BPP+katering) − katering`, which splits multi-month payments
wrongly while the combined figure is always right. A month is paid only when both buckets
cover it; shortfalls ≤ `TOLERANSI` (Rp 10.000) are ignored to absorb the discount on
paying kegiatan for the whole year. School-confirmed billing exceptions live in
`tagihan()`: kelas 1 owes no BPP/katering for July (paid at enrolment) and a special
July kegiatan of 114.500. The master's 220.000 katering for 33 kelas-4 students was wrong
and is corrected to 200.000 in the data at import, not in code.

`bayar_import.py data-alarm` (run `--dry-run` first) reads the school's monthly
workbooks, whose `PEMASUKAN BPP` tab is a row of side-by-side blocks, one per receipt
date, in an 8- or 9-column variant. The variant is told apart by where the TOTAL header
sits (`_lebar9`), not by the doubled "BPP" header — that header was blank in one block,
which shifted every amount one column over. In the 9-column variant BPP is derived as
combined − katering rather than read, because the "BPP" column is a formula that
sometimes hardcodes the wrong constant. Each run also compares every already-imported
row with its source cell by `KUNCI` and lists differences; `--perbaiki` overwrites them
(`bayar_sheet.koreksi_nominal`, the only in-place edit the ledger ever gets).
July rows mostly lack INDUK, so students are matched
by name (`cocokkan_nama`: initials, truncated words, near-spellings; refuses to guess
when ambiguous — measured 99.5% correct on 1,239 rows whose INDUK was known). It also
distrusts a written INDUK whose name plainly belongs to someone else, which is how it
caught the source typos 2037→2307 and 2315→2316. Anything uncertain goes to `PERLU
DICEK` for a human to resolve in the app. Every row carries a `KUNCI`
(`<month file>:<block column>:<row>`), so re-running the import never double-counts —
but it can't detect the same payment entered both in the app and in a later Excel
import, so don't import a month that is already being recorded in the app.

## Conventions

Code comments and all user-facing strings are Indonesian; commit messages are English.

Money is integer rupiah throughout — `rupiah()` / `ribuan()` in `app.py` format it;
never introduce floats for amounts.

Writes to Google Sheets are unguarded against concurrent edits: two simultaneous saves
on the same tab will clobber each other, and the Laporan Keuangan full-grid rewrite
makes that worse. Assume a single operator.

When changing anything that writes to a spreadsheet, verify against the live sheet
rather than trusting the code path — several bugs here (locale separators, duplicate
student numbers, a header parsed as a date) only appeared in the actual result.
