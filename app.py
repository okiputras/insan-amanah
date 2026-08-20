"""
Konverter Laporan R-5401 (Bank) -> Excel — versi ringan (Flask).

Upload file .txt laporan harian, validasi terhadap footer, unduh Excel rapi.
Tanpa Streamlit / numpy / pandas / pyarrow — hanya Flask + openpyxl, memakai
HTTP request/response biasa (tanpa WebSocket) supaya ringan & stabil di Railway.
"""
import io
import os
import hmac
import time
import zipfile
import secrets
import threading
from datetime import datetime

from flask import (
    Flask, request, render_template_string, send_file, abort, redirect, url_for,
    Response,
)

import json

from parser import parse_report, parse_laporan, build_workbook, build_combined_workbook
from validator import parse_master_siswa, reconcile_pembayaran, build_recon_workbook

# Menu Tabungan SMP (Google Sheets via gspread). Soft-import supaya menu lain
# tetap jalan meski gspread belum terpasang.
try:
    import tab_config as TC
    import tab_sheet as TS
    _TAB_IMPORT_ERR = None
except Exception as _e:            # noqa
    TC = TS = None
    _TAB_IMPORT_ERR = str(_e)

# Menu Laporan Keuangan SD (Google Sheets via gspread). Sama pola soft-import.
try:
    import lk_config as LK
    import lk_sheet as LSheet
    _LK_IMPORT_ERR = None
except Exception as _e:            # noqa
    LK = LSheet = None
    _LK_IMPORT_ERR = str(_e)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 40 * 1024 * 1024   # batas total 1x upload: 40 MB
PREVIEW_LIMIT = 100                                    # baris maksimum di tabel preview

# ---- login sederhana (HTTP Basic Auth) ----
# Kredensial hardcoded; bisa ditimpa via env AUTH_USER / AUTH_PASS tanpa edit kode.
AUTH_USER = os.environ.get("AUTH_USER", "evit")
AUTH_PASS = os.environ.get("AUTH_PASS", "evitcantik")


@app.before_request
def _require_login():
    if request.path == "/health":          # dibiarkan terbuka untuk healthcheck Railway
        return None
    auth = request.authorization
    if (auth and hmac.compare_digest(auth.username or "", AUTH_USER)
            and hmac.compare_digest(auth.password or "", AUTH_PASS)):
        return None
    return Response(
        "Perlu login untuk mengakses dashboard.", 401,
        {"WWW-Authenticate": 'Basic realm="Dashboard Insan Amanah"'},
    )

# ---- penyimpanan hasil sementara (in-memory, dibatasi TTL & jumlah) ----
# Menyimpan byte Excel hasil parsing supaya link unduh tidak perlu mem-parsing ulang.
# Dibatasi agar memori tidak menumpuk di container kecil.
_STORE = {}
_LOCK = threading.Lock()
_TTL = 1800          # hasil kedaluwarsa setelah 30 menit
_MAX_TOKENS = 20     # maksimum sesi hasil yang disimpan bersamaan


def _evict(now):
    dead = [k for k, v in _STORE.items() if now - v["ts"] > _TTL]
    for k in dead:
        _STORE.pop(k, None)
    if len(_STORE) > _MAX_TOKENS:
        oldest = sorted(_STORE, key=lambda k: _STORE[k]["ts"])[: len(_STORE) - _MAX_TOKENS]
        for k in oldest:
            _STORE.pop(k, None)


def _store(payload):
    token = secrets.token_urlsafe(16)
    with _LOCK:
        _evict(time.time())
        _STORE[token] = {"ts": time.time(), **payload}
    return token


def _get(token):
    with _LOCK:
        entry = _STORE.get(token)
        if entry and time.time() - entry["ts"] <= _TTL:
            return entry
    return None


# ---------- helper ----------
def rupiah(n):
    return "Rp " + f"{int(round(n)):,}".replace(",", ".")


def ribuan(n):
    return f"{int(n):,}".replace(",", ".")


def wb_to_bytes(wb):
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def xlsx_name(meta):
    kode = meta["kode"] or "NA"
    tgl = (meta["tanggal_label"] or "").replace("/", "")
    return f"Transaksi_R-5401_{kode}_{tgl}.xlsx"


XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _decode(raw):
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("latin-1", errors="replace")


# ---------- routes ----------
@app.route("/", methods=["GET"])
def index():
    return render_template_string(PAGE, results=None, token=None, combined=None,
                                  has_zip=False, summary=None, all_dup=False, active="convert")


@app.route("/analyze", methods=["POST"])
def analyze():
    files = [f for f in request.files.getlist("files") if f and f.filename]
    if not files:
        return redirect(url_for("index"))

    results = []       # data untuk ditampilkan per file
    per_file = []      # (download_name, xlsx_bytes) atau None, sejajar index dengan results
    built = []         # (download_name, xlsx_bytes) hanya untuk file unik -> isi .zip
    datasets = []      # (meta, rows) hanya untuk file unik -> Excel gabungan
    seen = set()       # (kode, tanggal) untuk deteksi duplikat

    for f in files:
        meta, rows = parse_report(_decode(f.read()))

        if not rows:
            results.append({"name": f.filename, "ok_parse": False})
            per_file.append(None)
            continue

        parsed_total = sum(r[3] for r in rows)
        parsed_count = len(rows)
        ft, fc = meta["footer_total"], meta["footer_count"]

        key = (meta["kode"], meta["tanggal_label"])
        is_dup = key in seen
        seen.add(key)

        total_ok = ft is not None and abs(parsed_total - ft) < 0.01
        count_ok = fc is not None and parsed_count == fc
        if total_ok and count_ok:
            status = "valid"
        elif ft is None and fc is None:
            status = "nofooter"
        else:
            status = "mismatch"

        dname = xlsx_name(meta)
        xbytes = wb_to_bytes(build_workbook(rows, meta))
        per_file.append((dname, xbytes))

        preview = [
            [r[0], r[1], r[2], ribuan(round(r[3])),
             r[4].strftime("%d/%m/%Y"), r[5].strftime("%H:%M:%S"),
             f"{r[6]:02d}", r[7], r[8], r[9]]
            for r in rows[:PREVIEW_LIMIT]
        ]

        results.append({
            "name": f.filename, "ok_parse": True, "status": status, "is_dup": is_dup,
            "meta": meta, "parsed_count": parsed_count,
            "total_str": rupiah(parsed_total), "count_str": ribuan(parsed_count),
            "ft_str": rupiah(ft) if ft is not None else None,
            "fc": fc,
            "selisih_str": rupiah(abs(parsed_total - ft)) if ft is not None else None,
            "total_ok": total_ok, "count_ok": count_ok,
            "preview": preview, "preview_more": max(0, parsed_count - len(preview)),
            "idx": len(per_file) - 1,
        })

        if not is_dup:
            datasets.append((meta, rows))
            built.append((dname, xbytes))

    # ---- gabungan ----
    combined = None
    zip_item = None
    summary = None
    if len(datasets) > 1:
        combined = (f"Transaksi_R-5401_GABUNGAN_{datetime.now():%Y%m%d}.xlsx",
                    wb_to_bytes(build_combined_workbook(datasets)))
        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, data in built:
                zf.writestr(name, data)
        zip_item = (f"Transaksi_R-5401_semua_{datetime.now():%Y%m%d}.zip", zbuf.getvalue())
        grand = sum(sum(r[3] for r in rows) for _, rows in datasets)
        total_txn = sum(len(rows) for _, rows in datasets)
        summary = {"n": len(datasets), "txn": ribuan(total_txn), "total": rupiah(grand)}

    all_dup = len(files) > 1 and len(datasets) <= 1

    token = _store({"per_file": per_file, "combined": combined, "zip": zip_item})
    return render_template_string(
        PAGE, results=results, token=token,
        combined=combined[0] if combined else None,
        has_zip=zip_item is not None, summary=summary, all_dup=all_dup, active="convert",
    )


def _send(item):
    if not item:
        abort(404)
    name, data = item
    mime = XLSX_MIME if name.endswith(".xlsx") else "application/zip"
    return send_file(io.BytesIO(data), as_attachment=True, download_name=name, mimetype=mime)


@app.route("/dl/<token>/f/<int:idx>")
def dl_file(token, idx):
    entry = _get(token)
    if not entry:
        abort(404)
    per_file = entry["per_file"]
    return _send(per_file[idx] if 0 <= idx < len(per_file) else None)


@app.route("/dl/<token>/combined")
def dl_combined(token):
    entry = _get(token)
    return _send(entry["combined"] if entry else None)


@app.route("/dl/<token>/zip")
def dl_zip(token):
    entry = _get(token)
    return _send(entry["zip"] if entry else None)


# ---------- menu: validasi pembayaran R-5401 vs master siswa (SD & SMP) ----------
# Beda SD vs SMP hanya pada cara mencocokkan No. Pelanggan (laporan) -> NO VA (master):
#   SD  : No. Pelanggan sudah = NO VA penuh.
#   SMP : NO VA = kode sekolah + No. Pelanggan (mis. 63713 + 0318 = 637130318).
LEVELS = {
    "sd": {
        "nama": "SD",
        "desc": ("Pencocokan tiap transaksi: <strong>No. Pelanggan (laporan) = NO VA (master)</strong> "
                 "secara langsung."),
    },
    "smp": {
        "nama": "SMP",
        "desc": ("Pencocokan tiap transaksi: <strong>NO VA = kode sekolah + No. Pelanggan</strong> "
                 "(mis. 63713 + 0318 = 637130318), karena laporan SMP memakai kode pelanggan pendek."),
    },
}


@app.route("/validasi/<level>", methods=["GET"])
def validasi(level):
    if level not in LEVELS:
        abort(404)
    return render_template_string(REKAP_PAGE, level=level, cfg=LEVELS[level],
                                  result=None, token=None, error=None, active="validasi_" + level)


@app.route("/validasi/<level>/proses", methods=["POST"])
def validasi_proses(level):
    if level not in LEVELS:
        abort(404)
    cfg = LEVELS[level]
    f_master = request.files.get("master")
    f_laporan = request.files.get("laporan")

    def _err(msg):
        return render_template_string(REKAP_PAGE, level=level, cfg=cfg,
                                      result=None, token=None, error=msg, active="validasi_" + level)

    if not f_master or not f_master.filename or not f_laporan or not f_laporan.filename:
        return _err("Mohon upload kedua file: master siswa (.xlsx) dan laporan harian R-5401 (.txt).")
    if not f_master.filename.lower().endswith((".xlsx", ".xlsm")):
        return _err(f"File master siswa ({f_master.filename}) harus berformat .xlsx.")
    if not f_laporan.filename.lower().endswith(".txt"):
        return _err(f"File laporan ({f_laporan.filename}) harus berformat .txt.")

    try:
        master = parse_master_siswa(f_master.read())
    except Exception:
        return _err("Gagal membaca file master siswa. Pastikan file Excel valid dengan kolom "
                    "NO VA | NAMA | BPP | KEGIATAN | TABUNGAN.")
    if not master:
        return _err("Tidak ada baris siswa valid yang terbaca dari file master siswa.")

    meta, report_rows = parse_laporan(_decode(f_laporan.read()))
    if not report_rows:
        return _err("Tidak ada baris transaksi terbaca dari file laporan. Pastikan ini file laporan "
                    "R-5401 (format lebar-tetap) atau Laporan BCA Virtual Account yang benar.")

    rows = reconcile_pembayaran(report_rows, master, meta.get("kode"), level=level)
    shown = [r for r in rows if r["status"] == "Sesuai"]

    preview = [{
        "no_pelanggan": r["no_pelanggan"], "nama": r["nama"], "lokasi": r["lokasi"],
        "tgl": r["tgl"].strftime("%d/%m/%Y"), "waktu": r["waktu"].strftime("%H:%M:%S"),
        "bpp": ribuan(r["bpp"]), "kegiatan": ribuan(r["kegiatan"]), "tabungan": ribuan(r["tabungan"]),
        "total_tagihan": ribuan(r["total_tagihan"]), "nilai_bayar": ribuan(r["nilai_bayar"]),
        "ket1": r["ket1"], "ket2": r["ket2"],
    } for r in shown[:PREVIEW_LIMIT]]

    tgl = (meta.get("tanggal_label") or "").replace("/", "")
    dname = f"Validasi_{cfg['nama']}_Sesuai_R-5401_{meta.get('kode') or 'NA'}_{tgl or 'NA'}.xlsx"
    xbytes = wb_to_bytes(build_recon_workbook(rows, meta, only_sesuai=True))
    token = _store({"rekap": (dname, xbytes)})

    result = {
        "master_name": f_master.filename, "laporan_name": f_laporan.filename, "meta": meta,
        "n_total": len(rows), "n_sesuai": len(shown),
        "n_kurang": sum(1 for r in rows if r["status"] == "Kurang"),
        "n_lebih": sum(1 for r in rows if r["status"] == "Lebih"),
        "n_unmatched": sum(1 for r in rows if not r["matched"]),
        "preview": preview, "preview_more": max(0, len(shown) - len(preview)),
    }
    return render_template_string(REKAP_PAGE, level=level, cfg=cfg,
                                  result=result, token=token, error=None, active="validasi_" + level)


@app.route("/rekap")
def rekap_redirect():
    # kompat lama: menu 2 dulu bernama /rekap, kini "Data Validasi SD"
    return redirect(url_for("validasi", level="sd"))


@app.route("/dl/<token>/rekap")
def dl_rekap(token):
    entry = _get(token)
    return _send(entry["rekap"] if entry else None)


# ---------- menu: Tabungan SMP (Google Sheets) ----------
def _tab_table(ws, year):
    """Susun tabel tampilan dari isi tab: header gabungan bulan·sub + baris data."""
    grid = ws.get_all_values()
    if len(grid) < TC.FIRST_DATA_ROW:
        return None
    month_row = grid[TC.HEADER_MONTH_ROW - 1]
    sub_row = grid[TC.HEADER_SUB_ROW - 1]
    headers = ["NO", "INDUK", "NAMA", "SALDO AWAL"]
    cur = ""
    for c in range(TC.N_FIXED, TC.total_cols(year)):
        if c < len(month_row) and month_row[c].strip():
            cur = month_row[c].strip()
        sub = sub_row[c].strip() if c < len(sub_row) else ""
        short = cur.split()[0][:3].title() if cur else ""
        headers.append(f"{short}·{sub}")
    rows = []
    for r in grid[TC.FIRST_DATA_ROW - 1:]:
        r = list(r[:len(headers)]) + [""] * (len(headers) - len(r))
        if not str(r[1]).strip():
            continue
        rows.append(r)
    # total saldo (kolom SALDO terakhir)
    last_saldo_idx = max((i for i, h in enumerate(headers) if h.endswith("SALDO")), default=None)
    total = 0
    if last_saldo_idx is not None:
        for r in rows:
            digits = "".join(ch for ch in str(r[last_saldo_idx]) if ch.isdigit() or ch == "-")
            try:
                total += int(digits) if digits not in ("", "-") else 0
            except ValueError:
                pass
    return {"headers": headers, "rows": rows, "n": len(rows),
            "total_saldo": rupiah(total),
            "last_month": headers[last_saldo_idx].split("·")[0] if last_saldo_idx else ""}


# Profil menu tabungan per jenjang (endpoint & label untuk template bersama)
_TAB_MENU = {
    "smp": {"jenjang": "SMP", "label": "Tabungan SMP", "active": "tab_smp",
            "ep_self": "tabungan", "ep_simpan": "tabungan_simpan", "ep_tahun": "tabungan_tahun_baru"},
    "sd": {"jenjang": "SD", "label": "Tabungan SD", "active": "tab_sd",
           "ep_self": "tabungan_sd", "ep_simpan": "tabungan_sd_simpan", "ep_tahun": "tabungan_sd_tahun_baru"},
}


def _profile(pkey):
    p = dict(_TAB_MENU[pkey])
    prof = TC.PROFILES[pkey] if TC else {}
    p["kelas"] = prof.get("kelas", [7, 8, 9] if pkey == "smp" else [1, 2, 3, 4, 5, 6])
    p["sid"] = prof.get("spreadsheet_id", "")
    return p


def _tab_ctx(pkey, kelas, year, msg=None, msgtype="info"):
    p = _profile(pkey)
    kelas = kelas or p["kelas"][0]
    ctx = {"active": p["active"], "jenjang": p["jenjang"], "label": p["label"],
           "ep_self": p["ep_self"], "ep_simpan": p["ep_simpan"], "ep_tahun": p["ep_tahun"],
           "kelas": kelas or p["kelas"][0], "kelas_list": p["kelas"], "years": [], "year": year,
           "tab_title": None, "months": [], "roster": [], "roster_json": "{}",
           "table": None, "error": None, "msg": msg, "msgtype": msgtype,
           "academic": None, "next_year": None,
           "sa_email": TC.SERVICE_ACCOUNT_EMAIL if TC else "",
           "sheet_url": f"https://docs.google.com/spreadsheets/d/{p['sid']}" if p["sid"] else "#",
           "today": datetime.now().strftime("%Y-%m-%d")}
    if TS is None:
        ctx["error"] = "Modul Tabungan belum siap: " + (_TAB_IMPORT_ERR or "gspread belum terpasang.")
        return ctx
    try:
        book = TS.open_book(p["sid"])
    except Exception as e:  # noqa (PermissionError / kredensial / dll.)
        ctx["error"] = str(e)
        return ctx
    years = TS.list_years(book, p["jenjang"])
    if not years:
        ctx["error"] = f"Belum ada tab data {p['jenjang']} di spreadsheet. Jalankan skrip build dulu."
        return ctx
    if year not in years:
        year = years[-1]
    ctx.update(year=year, years=years, next_year=max(years) + 1,
               academic=TC.academic_label(year), tab_title=TC.tab_name(kelas, year, p["jenjang"]))
    ws = book.worksheet(TC.tab_name(kelas, year, p["jenjang"]))
    roster = TS.read_roster_from_tab(ws)
    ctx["roster"] = roster
    ctx["roster_json"] = json.dumps({r["induk"]: r["nama"] for r in roster})
    ctx["months"] = [(m, TC.month_label(m, year)) for m in TC.months_for_year(year)]
    ctx["table"] = _tab_table(ws, year)
    return ctx


def _tab_render(pkey):
    return render_template_string(TABUNGAN_PAGE, **_tab_ctx(
        pkey, request.args.get("kelas", type=int), request.args.get("year", type=int),
        msg=request.args.get("msg"), msgtype=request.args.get("t", "info")))


def _tab_simpan(pkey):
    p = _profile(pkey)
    kelas = request.form.get("kelas", type=int)
    year = request.form.get("year", type=int)
    induk = (request.form.get("induk") or "").strip()
    jenis = request.form.get("jenis", "PENYETORAN")
    month_num = request.form.get("bulan", type=int)
    tanggal = request.form.get("tanggal")
    jumlah = request.form.get("jumlah", type=int)
    try:
        if TS is None:
            raise RuntimeError("Modul Tabungan belum siap.")
        book = TS.open_book(p["sid"])
        ws = book.worksheet(TC.tab_name(kelas, year, p["jenjang"]))
        roster = TS.read_roster_from_tab(ws)
        info = next((r for r in roster if r["induk"] == induk), None)
        if info is None:
            raise ValueError(f"No Induk '{induk}' tidak ditemukan di {TC.tab_name(kelas, year, p['jenjang'])}.")
        if not jumlah or jumlah <= 0:
            raise ValueError("Jumlah harus lebih dari 0.")
        tgl = (datetime.strptime(tanggal, "%Y-%m-%d").strftime("%d/%m/%Y")
               if tanggal else datetime.now().strftime("%d/%m/%Y"))
        new_saldo = TS.write_transaction(ws, info["row"], year, month_num, jenis, tgl, jumlah)
        msg = (f"✓ {jenis.title()} {rupiah(jumlah)} — {info['nama']} "
               f"({TC.month_label(month_num, year)}). Saldo baru: {rupiah(new_saldo)}.")
        return redirect(url_for(p["ep_self"], kelas=kelas, year=year, msg=msg, t="ok"), code=303)
    except Exception as e:  # noqa
        return redirect(url_for(p["ep_self"], kelas=kelas, year=year, msg=str(e), t="err"), code=303)


def _tab_tahun_baru(pkey):
    p = _profile(pkey)
    try:
        if TS is None:
            raise RuntimeError("Modul Tabungan belum siap.")
        book = TS.open_book(p["sid"])
        years = TS.list_years(book, p["jenjang"])
        ny = max(years) + 1
        for k in p["kelas"]:
            prev = book.worksheet(TC.tab_name(k, ny - 1, p["jenjang"]))
            carried = TS.last_saldo_of_year(prev, ny - 1)
            TS.build_tab(book, k, ny, carried, p["jenjang"])
        return redirect(url_for(p["ep_self"], year=ny,
                                msg=f"✓ Tab T.A. {TC.academic_label(ny)} dibuat.", t="ok"), code=303)
    except Exception as e:  # noqa
        return redirect(url_for(p["ep_self"], msg=str(e), t="err"), code=303)


@app.route("/tabungan")
def tabungan():
    return _tab_render("smp")


@app.route("/tabungan/simpan", methods=["POST"])
def tabungan_simpan():
    return _tab_simpan("smp")


@app.route("/tabungan/tahun-baru", methods=["POST"])
def tabungan_tahun_baru():
    return _tab_tahun_baru("smp")


@app.route("/tabungan-sd")
def tabungan_sd():
    return _tab_render("sd")


@app.route("/tabungan-sd/simpan", methods=["POST"])
def tabungan_sd_simpan():
    return _tab_simpan("sd")


@app.route("/tabungan-sd/tahun-baru", methods=["POST"])
def tabungan_sd_tahun_baru():
    return _tab_tahun_baru("sd")


# ---------------------------------------------------------------- Laporan Keuangan SD
def _lk_next_bulan_tahun(bulan, tahun):
    return (1, tahun + 1) if bulan == 12 else (bulan + 1, tahun)


def _lk_ctx(bulan, tahun, msg=None, msgtype="info"):
    ctx = {"active": "lk_sd", "bulan": bulan, "tahun": tahun, "bulan_tahun_list": [],
           "rows": [], "rows_json": "{}", "error": None, "msg": msg, "msgtype": msgtype,
           "level_labels": LK.LEVEL_LABELS if LK else [],
           "sa_email": LK.SERVICE_ACCOUNT_EMAIL if LK else "",
           "sheet_url": (f"https://docs.google.com/spreadsheets/d/{LK.SPREADSHEET_ID}"
                         if LK and LK.SPREADSHEET_ID else "#"),
           "tab_title": None, "next_bulan": None, "next_tahun": None}
    if LSheet is None:
        ctx["error"] = "Modul Laporan Keuangan belum siap: " + (_LK_IMPORT_ERR or "gspread belum terpasang.")
        return ctx
    if not LK.SPREADSHEET_ID:
        ctx["error"] = "SPREADSHEET_ID belum diisi di lk_config.py. Jalankan lk_build.py dulu."
        return ctx
    try:
        book = LSheet.open_book()
    except Exception as e:  # noqa
        ctx["error"] = str(e)
        return ctx
    tabs = LSheet.list_bulan_tabs(book)
    if not tabs:
        ctx["error"] = "Belum ada tab bulan di spreadsheet Laporan Keuangan. Jalankan lk_build.py dulu."
        return ctx
    if not any(b == bulan and t == tahun for b, t, _ in tabs):
        bulan, tahun, _ = tabs[-1]
    ctx["bulan"], ctx["tahun"], ctx["bulan_tahun_list"] = bulan, tahun, tabs
    ctx["next_bulan"], ctx["next_tahun"] = _lk_next_bulan_tahun(bulan, tahun)
    title = LK.tab_name(bulan, tahun)
    ctx["tab_title"] = title
    ws = book.worksheet(title)
    rows = LSheet.read_rows(ws)
    ctx["rows"] = rows
    ctx["rows_json"] = json.dumps({str(r["row"]): r for r in rows})
    return ctx


def _lk_render():
    return render_template_string(LAPORAN_KEUANGAN_PAGE, **_lk_ctx(
        request.args.get("bulan", type=int), request.args.get("tahun", type=int),
        msg=request.args.get("msg"), msgtype=request.args.get("t", "info")))


def _lk_form_data():
    keys = ["tanggal", "kode", "volume", "satuan", "fk", "kebutuhan", "unit_cost", "total"]
    data = {"level": request.form.get("level", type=int) or 1,
            "label": (request.form.get("label") or "").strip()}
    for k in keys:
        data[k] = (request.form.get(k) or "").strip()
    return data


def _lk_simpan():
    bulan = request.form.get("bulan", type=int)
    tahun = request.form.get("tahun", type=int)
    row_idx = request.form.get("row_idx", type=int)
    after_row_idx = request.form.get("after_row_idx", type=int)
    try:
        if LSheet is None:
            raise RuntimeError("Modul Laporan Keuangan belum siap.")
        data = _lk_form_data()
        if not data["label"]:
            raise ValueError("Label tidak boleh kosong.")
        book = LSheet.open_book()
        ws = book.worksheet(LK.tab_name(bulan, tahun))
        if row_idx:
            LSheet.update_row(ws, row_idx, data)
            msg = f"✓ Baris diperbarui: {data['label']}"
        else:
            LSheet.insert_row(ws, after_row_idx, data)
            msg = f"✓ Baris ditambahkan: {data['label']}"
        return redirect(url_for("laporan_keuangan", bulan=bulan, tahun=tahun, msg=msg, t="ok"), code=303)
    except Exception as e:  # noqa
        return redirect(url_for("laporan_keuangan", bulan=bulan, tahun=tahun, msg=str(e), t="err"), code=303)


def _lk_hapus():
    bulan = request.form.get("bulan", type=int)
    tahun = request.form.get("tahun", type=int)
    row_idx = request.form.get("row_idx", type=int)
    try:
        if LSheet is None:
            raise RuntimeError("Modul Laporan Keuangan belum siap.")
        book = LSheet.open_book()
        ws = book.worksheet(LK.tab_name(bulan, tahun))
        LSheet.delete_row(ws, row_idx)
        return redirect(url_for("laporan_keuangan", bulan=bulan, tahun=tahun,
                                msg="✓ Baris dihapus.", t="ok"), code=303)
    except Exception as e:  # noqa
        return redirect(url_for("laporan_keuangan", bulan=bulan, tahun=tahun, msg=str(e), t="err"), code=303)


def _lk_bulan_baru():
    bulan = request.form.get("bulan", type=int)
    tahun = request.form.get("tahun", type=int)
    new_bulan = request.form.get("new_bulan", type=int)
    new_tahun = request.form.get("new_tahun", type=int)
    try:
        if LSheet is None:
            raise RuntimeError("Modul Laporan Keuangan belum siap.")
        book = LSheet.open_book()
        clone_from = None
        try:
            clone_from = book.worksheet(LK.tab_name(bulan, tahun))
        except Exception:
            pass
        LSheet.ensure_bulan_tab(book, new_bulan, new_tahun, clone_from=clone_from)
        msg = f"✓ Tab {LK.tab_name(new_bulan, new_tahun)} dibuat."
        return redirect(url_for("laporan_keuangan", bulan=new_bulan, tahun=new_tahun, msg=msg, t="ok"), code=303)
    except Exception as e:  # noqa
        return redirect(url_for("laporan_keuangan", msg=str(e), t="err"), code=303)


@app.route("/laporan-keuangan")
def laporan_keuangan():
    return _lk_render()


@app.route("/laporan-keuangan/simpan", methods=["POST"])
def laporan_keuangan_simpan():
    return _lk_simpan()


@app.route("/laporan-keuangan/hapus", methods=["POST"])
def laporan_keuangan_hapus():
    return _lk_hapus()


@app.route("/laporan-keuangan/bulan-baru", methods=["POST"])
def laporan_keuangan_bulan_baru():
    return _lk_bulan_baru()


@app.route("/health")
def health():
    return "ok", 200


@app.errorhandler(413)
def too_large(_e):
    return ("File terlalu besar. Batas total 1x upload adalah 40 MB. "
            "Coba upload lebih sedikit file sekaligus."), 413


# ---------- template ----------
STYLE = """
  :root { --teal:#1F4E5F; --teal2:#2b6b80; --line:#e3e8ea; --band:#f2f7f9;
          --ok:#1a7f37; --okbg:#e8f5ec; --warn:#8a6d00; --warnbg:#fff7e0;
          --err:#b3261e; --errbg:#fdecea; --info:#0b5a75; --infobg:#e6f2f7; }
  * { box-sizing: border-box; }
  body { margin:0; font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
         color:#1a2327; background:#f7f9fa; line-height:1.5; }
  .wrap { max-width:960px; margin:0 auto; padding:28px 18px 60px; }
  h1 { font-size:1.7rem; margin:0 0 6px; color:var(--teal); }
  .sub { color:#5c6b70; margin:0 0 22px; font-size:.95rem; }
  .nav { display:flex; gap:6px; margin-bottom:20px; border-bottom:1px solid var(--line); }
  .nav a { padding:10px 16px; font-size:.9rem; font-weight:600; text-decoration:none;
           color:#5c6b70; border-bottom:2px solid transparent; }
  .nav a.active { color:var(--teal); border-bottom-color:var(--teal); }
  .nav a:hover { color:var(--teal2); }
  form.up { background:#fff; border:1px solid var(--line); border-radius:12px; padding:20px; }
  label.lbl { font-weight:600; display:block; margin-bottom:10px; }
  input[type=file] { display:block; width:100%; padding:14px; border:2px dashed #c4d2d7;
                     border-radius:10px; background:#fafcfd; cursor:pointer; }
  .field { margin-bottom:16px; }
  .field:last-of-type { margin-bottom:0; }
  .btn { display:inline-block; border:0; border-radius:9px; padding:11px 18px; font-size:.95rem;
         font-weight:600; cursor:pointer; text-decoration:none; }
  .btn.primary { background:var(--teal); color:#fff; margin-top:14px; }
  .btn.primary:hover { background:var(--teal2); }
  .btn.dl { background:var(--teal); color:#fff; margin-top:12px; }
  .btn.dl:hover { background:var(--teal2); }
  .btn.ghost { background:#eef3f5; color:var(--teal); }
  .card { background:#fff; border:1px solid var(--line); border-left:5px solid #ccc;
          border-radius:12px; padding:18px 20px; margin-top:18px; }
  .card.valid { border-left-color:var(--ok); }
  .card.mismatch, .card.fail { border-left-color:var(--err); }
  .card.nofooter { border-left-color:var(--warn); }
  .chead { display:flex; justify-content:space-between; align-items:flex-start; gap:12px; }
  .chead h2 { font-size:1.1rem; margin:0; word-break:break-all; }
  .badge { flex:none; font-size:.8rem; font-weight:700; padding:5px 11px; border-radius:999px; white-space:nowrap; }
  .badge.valid { background:var(--okbg); color:var(--ok); }
  .badge.mismatch, .badge.fail { background:var(--errbg); color:var(--err); }
  .badge.nofooter { background:var(--warnbg); color:var(--warn); }
  .badge.dup { background:var(--warnbg); color:var(--warn); }
  .metrics { display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin:16px 0 4px; }
  @media (max-width:640px){ .metrics{ grid-template-columns:repeat(2,1fr); } }
  .metric .k { font-size:.72rem; text-transform:uppercase; letter-spacing:.03em; color:#6b7a80; }
  .metric .v { font-size:1.15rem; font-weight:700; color:#16232a; }
  .pt { color:#5c6b70; font-size:.9rem; margin:6px 0 0; }
  .alert { border-radius:9px; padding:10px 13px; margin-top:12px; font-size:.9rem; }
  .alert.err { background:var(--errbg); color:var(--err); }
  .alert.warn { background:var(--warnbg); color:var(--warn); }
  .alert.info { background:var(--infobg); color:var(--info); }
  details { margin-top:14px; }
  summary { cursor:pointer; font-weight:600; color:var(--teal); }
  .tblwrap { overflow-x:auto; margin-top:12px; border:1px solid var(--line); border-radius:8px; }
  table { border-collapse:collapse; width:100%; font-size:.82rem; white-space:nowrap; }
  th,td { padding:7px 10px; border-bottom:1px solid var(--line); text-align:left; }
  th { background:var(--teal); color:#fff; position:sticky; top:0; }
  tbody tr:nth-child(even){ background:var(--band); }
  td.num { text-align:right; font-variant-numeric:tabular-nums; }
  tr.warn td { background:var(--warnbg); }
  .more { color:#6b7a80; font-size:.8rem; padding:8px 10px; }
  .combined { background:#fff; border:1px solid var(--line); border-radius:12px; padding:20px; margin-top:22px; }
  .combined h2 { margin:0 0 6px; color:var(--teal); font-size:1.2rem; }
  .cbtns { display:flex; flex-wrap:wrap; gap:12px; margin-top:6px; }
  .foot { color:#7c8a90; font-size:.82rem; margin-top:30px; }
  hr.sep { border:0; border-top:1px solid var(--line); margin:26px 0; }
"""

PAGE = """<!doctype html>
<html lang="id">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Konverter Laporan R-5401 &rarr; Excel</title>
<style>""" + STYLE + """</style>
</head>
<body>
<div class="wrap">
  <div class="nav">
    <a href="{{ url_for('index') }}" class="{{ 'active' if active=='convert' else '' }}">Konversi R-5401</a>
    <a href="{{ url_for('validasi', level='sd') }}" class="{{ 'active' if active=='validasi_sd' else '' }}">Data Validasi SD</a>
    <a href="{{ url_for('validasi', level='smp') }}" class="{{ 'active' if active=='validasi_smp' else '' }}">Data Validasi SMP</a>
    <a href="{{ url_for('tabungan') }}" class="{{ 'active' if active=='tab_smp' else '' }}">Tabungan SMP</a>
    <a href="{{ url_for('tabungan_sd') }}" class="{{ 'active' if active=='tab_sd' else '' }}">Tabungan SD</a>
    <a href="{{ url_for('laporan_keuangan') }}" class="{{ 'active' if active=='lk_sd' else '' }}">Laporan Keuangan SD</a>
  </div>
  <h1>&#128202; Konverter Laporan R-5401 &rarr; Excel</h1>
  <p class="sub">Upload file laporan transaksi harian (.txt format lebar-tetap dari bank).
     Aplikasi akan mem-parsing, memvalidasi terhadap total footer, dan menyiapkan file
     Excel rapi (Data + Ringkasan) untuk diunduh.</p>

  <form class="up" method="post" action="{{ url_for('analyze') }}" enctype="multipart/form-data">
    <label class="lbl" for="files">Pilih satu atau beberapa file laporan (.txt)</label>
    <input id="files" type="file" name="files" accept=".txt" multiple required>
    <button class="btn primary" type="submit">&#9889; Proses &amp; Validasi</button>
  </form>

  {% if results is not none %}
    {% if results|length == 0 %}
      <div class="alert info" style="margin-top:22px;">Tidak ada file yang bisa diproses.</div>
    {% endif %}

    {% for r in results %}
      {% if not r.ok_parse %}
        <div class="card fail">
          <div class="chead"><h2>&#128196; {{ r.name }}</h2><span class="badge fail">Gagal</span></div>
          <div class="alert warn">Tidak ada baris transaksi terbaca. Pastikan ini file laporan
             R-5401 dengan format lebar-tetap yang benar.</div>
        </div>
      {% else %}
        <div class="card {{ r.status }}">
          <div class="chead">
            <h2>&#128196; {{ r.name }}</h2>
            {% if r.status == 'valid' %}<span class="badge valid">&#10003; Valid</span>
            {% elif r.status == 'nofooter' %}<span class="badge nofooter">Tanpa footer</span>
            {% else %}<span class="badge mismatch">Selisih!</span>{% endif %}
          </div>

          <div class="metrics">
            <div class="metric"><div class="k">Kode Perusahaan</div><div class="v">{{ r.meta.kode or '—' }}</div></div>
            <div class="metric"><div class="k">Tanggal</div><div class="v">{{ r.meta.tanggal_label or '—' }}</div></div>
            <div class="metric"><div class="k">Jumlah Transaksi</div><div class="v">{{ r.count_str }}</div></div>
            <div class="metric"><div class="k">Total Nilai</div><div class="v">{{ r.total_str }}</div></div>
          </div>

          {% if r.meta.nama_pt %}<p class="pt"><strong>{{ r.meta.nama_pt }}</strong> &bull; {{ r.meta.cabang }}</p>{% endif %}

          {% if not r.total_ok and r.ft_str %}
            <div class="alert err">&#9888; Total hasil parsing ({{ r.total_str }}) TIDAK cocok dengan
               total footer laporan ({{ r.ft_str }}). Selisih {{ r.selisih_str }}.</div>
          {% endif %}
          {% if not r.count_ok and r.fc is not none %}
            <div class="alert err">&#9888; Jumlah transaksi hasil parsing ({{ r.parsed_count }}) tidak
               cocok dengan footer ({{ r.fc }}).</div>
          {% endif %}
          {% if r.is_dup %}
            <div class="alert warn">&#128257; File ini duplikat (kode &amp; tanggal sama dengan file lain).
               Tetap bisa diunduh, tapi <strong>tidak</strong> ikut dalam file gabungan agar total tidak dobel.</div>
          {% endif %}

          <details>
            <summary>&#128065; Lihat data ({{ r.count_str }} baris)</summary>
            <div class="tblwrap">
              <table>
                <thead><tr>
                  <th>No.</th><th>No. Pelanggan/TXN</th><th>Nama</th><th>Nilai</th><th>Tgl</th>
                  <th>Waktu</th><th>Jam</th><th>Lokasi</th><th>Keterangan 1</th><th>Keterangan 2</th>
                </tr></thead>
                <tbody>
                  {% for row in r.preview %}
                    <tr>
                      <td>{{ row[0] }}</td><td>{{ row[1] }}</td><td>{{ row[2] }}</td>
                      <td class="num">{{ row[3] }}</td><td>{{ row[4] }}</td><td>{{ row[5] }}</td>
                      <td>{{ row[6] }}</td><td>{{ row[7] }}</td><td>{{ row[8] }}</td><td>{{ row[9] }}</td>
                    </tr>
                  {% endfor %}
                </tbody>
              </table>
              {% if r.preview_more %}<div class="more">… {{ r.preview_more }} baris lainnya (lengkap di file Excel).</div>{% endif %}
            </div>
          </details>

          <a class="btn dl" href="{{ url_for('dl_file', token=token, idx=r.idx) }}">&#11015; Unduh Excel file ini</a>
        </div>
      {% endif %}
    {% endfor %}

    {% if summary %}
      <div class="combined">
        <h2>&#128230; Unduh Gabungan</h2>
        <p class="sub" style="margin:0 0 10px;"><strong>{{ summary.n }} laporan</strong> unik &bull;
           <strong>{{ summary.txn }}</strong> transaksi &bull; total <strong>{{ summary.total }}</strong></p>
        <div class="cbtns">
          <a class="btn dl" href="{{ url_for('dl_combined', token=token) }}">&#11015; Unduh 1 Excel gabungan (Data + Rekap Harian)</a>
          {% if has_zip %}<a class="btn ghost" href="{{ url_for('dl_zip', token=token) }}">&#11015; Unduh semua (.zip berisi file terpisah)</a>{% endif %}
        </div>
      </div>
    {% elif all_dup %}
      <div class="alert info" style="margin-top:22px;">&#128230; Tidak ada file gabungan untuk diunduh —
         semua file yang diupload adalah laporan duplikat (kode &amp; tanggal sama), jadi hanya dianggap
         1 laporan unik. Unduh lewat tombol per-file di atas.</div>
    {% endif %}
  {% endif %}

  <hr class="sep">
  <p class="foot">Catatan: Nama pelanggan &amp; keterangan pada laporan sumber terpotong
     (field lebar-tetap &plusmn;16 karakter). Nilai, tanggal, waktu, dan lokasi akurat 100%.</p>
</div>
</body>
</html>"""


REKAP_PAGE = """<!doctype html>
<html lang="id">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Data Validasi {{ cfg.nama }} &mdash; R-5401 vs Master Siswa</title>
<style>""" + STYLE + """</style>
</head>
<body>
<div class="wrap">
  <div class="nav">
    <a href="{{ url_for('index') }}" class="{{ 'active' if active=='convert' else '' }}">Konversi R-5401</a>
    <a href="{{ url_for('validasi', level='sd') }}" class="{{ 'active' if active=='validasi_sd' else '' }}">Data Validasi SD</a>
    <a href="{{ url_for('validasi', level='smp') }}" class="{{ 'active' if active=='validasi_smp' else '' }}">Data Validasi SMP</a>
    <a href="{{ url_for('tabungan') }}" class="{{ 'active' if active=='tab_smp' else '' }}">Tabungan SMP</a>
    <a href="{{ url_for('tabungan_sd') }}" class="{{ 'active' if active=='tab_sd' else '' }}">Tabungan SD</a>
    <a href="{{ url_for('laporan_keuangan') }}" class="{{ 'active' if active=='lk_sd' else '' }}">Laporan Keuangan SD</a>
  </div>
  <h1>&#128203; Data Validasi {{ cfg.nama }}</h1>
  <p class="sub">Upload <strong>master siswa {{ cfg.nama }}</strong> (.xlsx: NO VA, NAMA, BPP, KEGIATAN,
     TABUNGAN) dan <strong>laporan harian R-5401</strong> (.txt lebar-tetap). {{ cfg.desc|safe }}
     Hasil Excel hanya berisi transaksi berstatus <strong>Sesuai</strong> (Nilai Bayar = Total Tagihan).</p>

  <form class="up" method="post" action="{{ url_for('validasi_proses', level=level) }}" enctype="multipart/form-data">
    <div class="field">
      <label class="lbl" for="master">Master siswa (.xlsx)</label>
      <input id="master" type="file" name="master" accept=".xlsx" required>
    </div>
    <div class="field">
      <label class="lbl" for="laporan">Laporan harian R-5401 (.txt)</label>
      <input id="laporan" type="file" name="laporan" accept=".txt" required>
    </div>
    <button class="btn primary" type="submit">&#9889; Proses</button>
  </form>

  {% if error %}
    <div class="alert err" style="margin-top:22px;">&#9888; {{ error }}</div>
  {% endif %}

  {% if result %}
    <div class="card {{ 'valid' if result.n_sesuai == result.n_total else 'mismatch' }}">
      <div class="chead">
        <h2>&#128196; {{ result.master_name }} &harr; {{ result.laporan_name }}</h2>
        {% if result.n_sesuai == result.n_total %}<span class="badge valid">&#10003; Semua Sesuai</span>
        {% else %}<span class="badge mismatch">Ada Selisih</span>{% endif %}
      </div>

      <div class="metrics">
        <div class="metric"><div class="k">Total Transaksi</div><div class="v">{{ result.n_total }}</div></div>
        <div class="metric"><div class="k">Sesuai</div><div class="v">{{ result.n_sesuai }}</div></div>
        <div class="metric"><div class="k">Kurang</div><div class="v">{{ result.n_kurang }}</div></div>
        <div class="metric"><div class="k">Lebih</div><div class="v">{{ result.n_lebih }}</div></div>
      </div>
      {% if result.n_unmatched %}
        <div class="alert warn">&#9888; {{ result.n_unmatched }} transaksi tidak ditemukan No. Pelanggannya
           di master siswa (dianggap tagihan 0, otomatis tidak masuk status Sesuai).</div>
      {% endif %}

      {% if result.meta.nama_pt %}<p class="pt"><strong>{{ result.meta.nama_pt }}</strong> &bull; {{ result.meta.cabang }} &bull; {{ result.meta.tanggal_label }}</p>{% endif %}

      <details open>
        <summary>&#128065; Lihat transaksi Sesuai ({{ result.n_sesuai }} baris)</summary>
        <div class="tblwrap">
          <table>
            <thead><tr>
              <th>No. Pelanggan</th><th>Nama</th><th>Tgl</th><th>Waktu</th><th>Lokasi</th>
              <th>BPP</th><th>Kegiatan</th><th>Tabungan</th>
              <th>Total Tagihan</th><th>Nilai Bayar</th><th>Ket 1</th><th>Ket 2</th>
            </tr></thead>
            <tbody>
              {% for row in result.preview %}
                <tr>
                  <td>{{ row.no_pelanggan }}</td><td>{{ row.nama }}</td><td>{{ row.tgl }}</td>
                  <td>{{ row.waktu }}</td><td>{{ row.lokasi }}</td>
                  <td class="num">{{ row.bpp }}</td><td class="num">{{ row.kegiatan }}</td><td class="num">{{ row.tabungan }}</td>
                  <td class="num">{{ row.total_tagihan }}</td><td class="num">{{ row.nilai_bayar }}</td>
                  <td>{{ row.ket1 }}</td><td>{{ row.ket2 }}</td>
                </tr>
              {% endfor %}
            </tbody>
          </table>
          {% if result.preview_more %}<div class="more">… {{ result.preview_more }} baris lainnya (lengkap di file Excel).</div>{% endif %}
        </div>
      </details>

      <a class="btn dl" href="{{ url_for('dl_rekap', token=token) }}">&#11015; Unduh Excel (hanya Sesuai)</a>
    </div>
  {% endif %}

  <hr class="sep">
  <p class="foot">Catatan: hanya transaksi berstatus Sesuai yang masuk file Excel; sheet Ringkasan
     tetap merangkum seluruh transaksi termasuk Kurang/Lebih/tanpa data master.</p>
</div>
</body>
</html>"""


TAB_EXTRA = """
  .tab-grid { display:grid; grid-template-columns:1fr 1fr; gap:18px; }
  @media (max-width:640px){ .tab-grid{ grid-template-columns:1fr; } }
  .in { width:100%; padding:11px 12px; border:1px solid #c4d2d7; border-radius:9px;
        background:#fff; font-size:.95rem; font-family:inherit; }
  .in:focus { outline:2px solid var(--teal2); border-color:var(--teal2); }
  .in[readonly]{ background:#eef3f5; color:#5c6b70; }
  .radios { display:flex; gap:20px; align-items:center; padding:6px 0; }
  .radios label{ font-weight:600; display:flex; gap:6px; align-items:center; cursor:pointer; }
  .filterbar { display:flex; flex-wrap:wrap; gap:14px; align-items:flex-end; background:#fff;
               border:1px solid var(--line); border-radius:12px; padding:16px 18px; margin-bottom:18px; }
  .filterbar .fld{ display:flex; flex-direction:column; gap:5px; }
  .filterbar .fld .lbl{ margin:0; font-size:.78rem; text-transform:uppercase; letter-spacing:.03em; }
  .search { padding:9px 12px; border:1px solid #c4d2d7; border-radius:9px; font-size:.9rem;
            width:280px; max-width:100%; margin-bottom:10px; }
  table.data th:nth-child(3), table.data td:nth-child(3){ text-align:left; white-space:nowrap; }
  .alert.ok { background:var(--okbg); color:var(--ok); }
  .hint{ color:#6b7a80; font-size:.82rem; margin:4px 0 0; }
"""

TABUNGAN_PAGE = """<!doctype html>
<html lang="id">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{{ label }} Insan Amanah</title>
<style>""" + STYLE + TAB_EXTRA + """</style>
</head>
<body>
<div class="wrap">
  <div class="nav">
    <a href="{{ url_for('index') }}" class="{{ 'active' if active=='convert' else '' }}">Konversi R-5401</a>
    <a href="{{ url_for('validasi', level='sd') }}" class="{{ 'active' if active=='validasi_sd' else '' }}">Data Validasi SD</a>
    <a href="{{ url_for('validasi', level='smp') }}" class="{{ 'active' if active=='validasi_smp' else '' }}">Data Validasi SMP</a>
    <a href="{{ url_for('tabungan') }}" class="{{ 'active' if active=='tab_smp' else '' }}">Tabungan SMP</a>
    <a href="{{ url_for('tabungan_sd') }}" class="{{ 'active' if active=='tab_sd' else '' }}">Tabungan SD</a>
    <a href="{{ url_for('laporan_keuangan') }}" class="{{ 'active' if active=='lk_sd' else '' }}">Laporan Keuangan SD</a>
  </div>
  <h1>&#127974; {{ label }} Insan Amanah</h1>
  <p class="sub">Catat penyetoran/penarikan tabungan siswa langsung ke Google Sheet.
     Saldo dihitung otomatis (saldo bulan lalu + setor &minus; tarik).</p>

  {% if msg %}<div class="alert {{ msgtype }}" style="margin-bottom:16px;">{{ msg }}</div>{% endif %}

  {% if error %}
    <div class="alert err">&#9888; {{ error }}</div>
    {% if sa_email and 'share' in error|lower %}
      <p class="hint">Bagikan spreadsheet ke <strong>{{ sa_email }}</strong> sebagai <strong>Editor</strong>,
         lalu <a href="{{ url_for(ep_self) }}">muat ulang</a>.</p>
    {% endif %}
  {% else %}

    <form class="filterbar" method="get" action="{{ url_for(ep_self) }}">
      <div class="fld">
        <label class="lbl">Kelas</label>
        <select class="in" name="kelas" onchange="this.form.submit()">
          {% for k in kelas_list %}<option value="{{ k }}" {{ 'selected' if k==kelas else '' }}>Kelas {{ k }}</option>{% endfor %}
        </select>
      </div>
      <div class="fld">
        <label class="lbl">Tahun Ajaran</label>
        <select class="in" name="year" onchange="this.form.submit()">
          {% for y in years %}<option value="{{ y }}" {{ 'selected' if y==year else '' }}>{{ y }}/{{ y+1 }}</option>{% endfor %}
        </select>
      </div>
      <div class="fld"><a class="btn ghost" href="{{ sheet_url }}" target="_blank" rel="noopener">&#128196; Buka Google Sheet</a></div>
    </form>

    <form class="up" method="post" action="{{ url_for(ep_simpan) }}">
      <input type="hidden" name="kelas" value="{{ kelas }}">
      <input type="hidden" name="year" value="{{ year }}">
      <div class="tab-grid">
        <div>
          <div class="field">
            <label class="lbl" for="induk">1 &middot; No Induk (ketik / pilih)</label>
            <input class="in" id="induk" name="induk" list="siswa" autocomplete="off"
                   placeholder="mis. 0344" oninput="isiNama()" required>
            <datalist id="siswa">
              {% for r in roster %}<option value="{{ r.induk }}">{{ r.nama }}</option>{% endfor %}
            </datalist>
            <p class="hint">{{ roster|length }} siswa di {{ tab_title }}</p>
          </div>
          <div class="field">
            <label class="lbl" for="nama">Nama Lengkap</label>
            <input class="in" id="nama" readonly placeholder="otomatis dari No Induk">
          </div>
        </div>
        <div>
          <div class="field">
            <label class="lbl">2 &middot; Jenis</label>
            <div class="radios">
              <label><input type="radio" name="jenis" value="PENYETORAN" checked onchange="setTgl()"> Penyetoran</label>
              <label><input type="radio" name="jenis" value="PENARIKAN" onchange="setTgl()"> Penarikan</label>
            </div>
          </div>
          <div class="field">
            <label class="lbl" for="bulan">3 &middot; Bulan (boleh acak)</label>
            <select class="in" id="bulan" name="bulan">
              {% for m,lab in months %}<option value="{{ m }}">{{ lab }}</option>{% endfor %}
            </select>
          </div>
          <div class="field">
            <label class="lbl" for="tanggal" id="tglLbl">Tanggal Penyetoran</label>
            <input class="in" type="date" id="tanggal" name="tanggal" value="{{ today }}">
          </div>
          <div class="field">
            <label class="lbl" for="jumlah">Jumlah (Rp)</label>
            <input class="in" type="number" id="jumlah" name="jumlah" min="0" step="1" placeholder="0" required>
          </div>
        </div>
      </div>
      <button class="btn primary" type="submit">&#128190; Simpan (saldo terisi otomatis)</button>
    </form>

    {% if table %}
    <div class="combined" style="margin-top:22px;">
      <h2>&#128202; {{ tab_title }}</h2>
      <p class="sub" style="margin:0 0 10px;">{{ table.n }} siswa &bull;
         Total Saldo ({{ table.last_month }}): <strong>{{ table.total_saldo }}</strong></p>
      <input class="search" id="cari" placeholder="&#128269; Cari induk / nama…" onkeyup="filterTabel()">
      <div class="tblwrap" style="max-height:520px; overflow:auto;">
        <table class="data" id="dataTabungan">
          <thead><tr>{% for h in table.headers %}<th>{{ h }}</th>{% endfor %}</tr></thead>
          <tbody>
            {% for row in table.rows %}<tr>{% for cell in row %}<td>{{ cell }}</td>{% endfor %}</tr>{% endfor %}
          </tbody>
        </table>
      </div>
    </div>
    {% endif %}

    <form method="post" action="{{ url_for(ep_tahun) }}" style="margin-top:18px;"
          onsubmit="return confirm('Buat tab T.A. {{ next_year }}/{{ next_year+1 }} untuk semua kelas? Saldo awal = saldo Juni tahun ini.');">
      <button class="btn ghost" type="submit">&#10133; Buat Tahun Ajaran {{ next_year }}/{{ next_year+1 }}</button>
    </form>

  {% endif %}

  <hr class="sep">
  <p class="foot">Data tersimpan di Google Sheet (sumber tunggal). Kolom SALDO memakai formula
     berjalan; simpan ulang pada siswa &amp; bulan yang sama untuk mengedit.</p>
</div>

<script>
  const NAMA = {{ roster_json|safe }};
  function isiNama(){
    var v = document.getElementById('induk').value.trim();
    document.getElementById('nama').value = NAMA[v] || '';
  }
  function setTgl(){
    var p = document.querySelector('input[name=jenis]:checked').value;
    document.getElementById('tglLbl').textContent =
      (p === 'PENYETORAN') ? 'Tanggal Penyetoran' : 'Tanggal Penarikan';
  }
  function filterTabel(){
    var q = document.getElementById('cari').value.toLowerCase();
    var rows = document.querySelectorAll('#dataTabungan tbody tr');
    rows.forEach(function(tr){
      var t = (tr.cells.length > 2)
        ? (tr.cells[1].textContent + ' ' + tr.cells[2].textContent).toLowerCase()
        : tr.textContent.toLowerCase();
      tr.style.display = t.indexOf(q) > -1 ? '' : 'none';
    });
  }
</script>
</body>
</html>"""


LK_EXTRA = """
  .lk-row td.lbl-cell { white-space:nowrap; }
  .lk-row .indent { display:inline-block; }
  .lk-lvl1 .lbl-cell { font-weight:700; color:var(--teal); }
  .lk-lvl2 .lbl-cell { font-weight:700; }
  .lk-lvl3 .lbl-cell { font-weight:600; }
  .lk-actions a { margin-right:10px; font-size:.85rem; cursor:pointer; }
  .lk-actions .del { color:#b3261e; }
  table.data td.num, table.data th.num { text-align:right; white-space:nowrap; }
  .form-title { color:var(--teal); font-weight:700; margin:0 0 10px; }
  .fin-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:14px; }
  @media (max-width:900px){ .fin-grid{ grid-template-columns:repeat(2,1fr); } }
"""

LAPORAN_KEUANGAN_PAGE = """<!doctype html>
<html lang="id">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Laporan Keuangan SD Insan Amanah</title>
<style>""" + STYLE + TAB_EXTRA + LK_EXTRA + """</style>
</head>
<body>
<div class="wrap">
  <div class="nav">
    <a href="{{ url_for('index') }}" class="{{ 'active' if active=='convert' else '' }}">Konversi R-5401</a>
    <a href="{{ url_for('validasi', level='sd') }}" class="{{ 'active' if active=='validasi_sd' else '' }}">Data Validasi SD</a>
    <a href="{{ url_for('validasi', level='smp') }}" class="{{ 'active' if active=='validasi_smp' else '' }}">Data Validasi SMP</a>
    <a href="{{ url_for('tabungan') }}" class="{{ 'active' if active=='tab_smp' else '' }}">Tabungan SMP</a>
    <a href="{{ url_for('tabungan_sd') }}" class="{{ 'active' if active=='tab_sd' else '' }}">Tabungan SD</a>
    <a href="{{ url_for('laporan_keuangan') }}" class="{{ 'active' if active=='lk_sd' else '' }}">Laporan Keuangan SD</a>
  </div>
  <h1>&#128176; Laporan Keuangan SD Insan Amanah</h1>
  <p class="sub">Rincian Program &rarr; Sub Program &rarr; Kegiatan &rarr; Rincian, tersimpan
     langsung di Google Sheet. Tiap baris bisa ditambah, diedit, atau dihapus bebas.</p>

  {% if msg %}<div class="alert {{ msgtype }}" style="margin-bottom:16px;">{{ msg }}</div>{% endif %}

  {% if error %}
    <div class="alert err">&#9888; {{ error }}</div>
    {% if sa_email and 'share' in error|lower %}
      <p class="hint">Bagikan spreadsheet ke <strong>{{ sa_email }}</strong> sebagai <strong>Editor</strong>,
         lalu <a href="{{ url_for('laporan_keuangan') }}">muat ulang</a>.</p>
    {% endif %}
  {% else %}

    <form class="filterbar" method="get" action="{{ url_for('laporan_keuangan') }}">
      <div class="fld">
        <label class="lbl">Bulan</label>
        <select class="in" name="bt" onchange="var p=this.value.split('|');
                document.getElementById('selBulan').value=p[0];
                document.getElementById('selTahun').value=p[1]; this.form.submit();">
          {% for b,t,title in bulan_tahun_list %}
            <option value="{{ b }}|{{ t }}" {{ 'selected' if b==bulan and t==tahun else '' }}>{{ title }}</option>
          {% endfor %}
        </select>
        <input type="hidden" id="selBulan" name="bulan" value="{{ bulan }}">
        <input type="hidden" id="selTahun" name="tahun" value="{{ tahun }}">
      </div>
      <div class="fld"><a class="btn ghost" href="{{ sheet_url }}" target="_blank" rel="noopener">&#128196; Buka Google Sheet</a></div>
    </form>

    <div class="combined">
      <p class="form-title" id="formTitle">&#10133; Tambah Baris</p>
      <form method="post" action="{{ url_for('laporan_keuangan_simpan') }}" id="lkForm">
        <input type="hidden" name="bulan" value="{{ bulan }}">
        <input type="hidden" name="tahun" value="{{ tahun }}">
        <input type="hidden" name="row_idx" id="f_row_idx" value="">
        <input type="hidden" name="after_row_idx" id="f_after_row_idx" value="">
        <div class="tab-grid">
          <div>
            <div class="field">
              <label class="lbl" for="f_level">Level</label>
              <select class="in" id="f_level" name="level">
                {% for lab in level_labels %}
                  <option value="{{ loop.index }}">{{ loop.index }} &middot; {{ lab }}</option>
                {% endfor %}
              </select>
            </div>
            <div class="field">
              <label class="lbl" for="f_label">Label / Rincian</label>
              <input class="in" id="f_label" name="label" placeholder="mis. Rapat Kerja (Raker)" required>
            </div>
          </div>
          <div class="field">
            <label class="lbl">Kolom Keuangan (isi jika perlu)</label>
            <div class="fin-grid">
              <div><label class="lbl" for="f_tanggal">Tanggal</label>
                <input class="in" type="date" id="f_tanggal" name="tanggal"></div>
              <div><label class="lbl" for="f_kode">Kode</label>
                <input class="in" id="f_kode" name="kode"></div>
              <div><label class="lbl" for="f_volume">Volume</label>
                <input class="in" type="number" step="any" id="f_volume" name="volume" oninput="autoTotal()"></div>
              <div><label class="lbl" for="f_satuan">Satuan</label>
                <input class="in" id="f_satuan" name="satuan"></div>
              <div><label class="lbl" for="f_fk">FK</label>
                <input class="in" id="f_fk" name="fk"></div>
              <div><label class="lbl" for="f_kebutuhan">Kebutuhan</label>
                <input class="in" id="f_kebutuhan" name="kebutuhan"></div>
              <div><label class="lbl" for="f_unit_cost">Unit Cost</label>
                <input class="in" type="number" step="any" id="f_unit_cost" name="unit_cost" oninput="autoTotal()"></div>
              <div><label class="lbl" for="f_total">Total</label>
                <input class="in" type="number" step="any" id="f_total" name="total"></div>
            </div>
          </div>
        </div>
        <button class="btn primary" type="submit">&#128190; Simpan</button>
        <a class="btn ghost" onclick="resetForm()" style="margin-left:10px;">Batal / Baris Baru</a>
      </form>
    </div>

    <div class="combined" style="margin-top:22px;">
      <h2>&#128202; {{ tab_title }}</h2>
      <div class="tblwrap" style="max-height:560px; overflow:auto;">
        <table class="data" id="dataLK">
          <thead><tr>
            <th>Label</th><th>Tanggal</th><th>Kode</th><th class="num">Volume</th><th>Satuan</th>
            <th>FK</th><th>Kebutuhan</th><th class="num">Unit Cost</th><th class="num">Total</th><th>Aksi</th>
          </tr></thead>
          <tbody>
            {% for r in rows %}
            <tr class="lk-row lk-lvl{{ r.level }}">
              <td class="lbl-cell"><span class="indent" style="width:{{ (r.level-1)*22 }}px;"></span>{{ r.label }}</td>
              <td>{{ r.tanggal }}</td><td>{{ r.kode }}</td><td class="num">{{ r.volume }}</td><td>{{ r.satuan }}</td>
              <td>{{ r.fk }}</td><td>{{ r.kebutuhan }}</td><td class="num">{{ r.unit_cost }}</td><td class="num">{{ r.total }}</td>
              <td class="lk-actions">
                <a onclick="editRow({{ r.row }})">Edit</a>
                <a onclick="addAfter({{ r.row }})">+ Baris</a>
                <form method="post" action="{{ url_for('laporan_keuangan_hapus') }}" style="display:inline;"
                      onsubmit="return confirm('Hapus baris &quot;{{ r.label }}&quot;?');">
                  <input type="hidden" name="bulan" value="{{ bulan }}">
                  <input type="hidden" name="tahun" value="{{ tahun }}">
                  <input type="hidden" name="row_idx" value="{{ r.row }}">
                  <button type="submit" class="del" style="background:none;border:0;padding:0;
                          font-size:.85rem;color:#b3261e;cursor:pointer;">Hapus</button>
                </form>
              </td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
    </div>

    <form method="post" action="{{ url_for('laporan_keuangan_bulan_baru') }}" style="margin-top:18px;"
          onsubmit="return confirm('Buat tab bulan baru? Struktur Program/Sub Program/Kegiatan bulan ini akan disalin, kolom keuangan dikosongkan.');">
      <input type="hidden" name="bulan" value="{{ bulan }}">
      <input type="hidden" name="tahun" value="{{ tahun }}">
      <input type="hidden" name="new_bulan" value="{{ next_bulan }}">
      <input type="hidden" name="new_tahun" value="{{ next_tahun }}">
      <button class="btn ghost" type="submit">&#10133; Buat Bulan Berikutnya</button>
    </form>

  {% endif %}

  <hr class="sep">
  <p class="foot">Data tersimpan di Google Sheet (sumber tunggal). Kolom keuangan bebas diisi
     di level manapun sesuai kebutuhan.</p>
</div>

<script>
  const LK_ROWS = {{ rows_json|safe }};
  function resetForm(){
    document.getElementById('lkForm').reset();
    document.getElementById('f_row_idx').value = '';
    document.getElementById('f_after_row_idx').value = '';
    document.getElementById('formTitle').innerHTML = '&#10133; Tambah Baris';
  }
  function fillForm(r){
    document.getElementById('f_level').value = r.level;
    document.getElementById('f_label').value = r.label || '';
    document.getElementById('f_tanggal').value = r.tanggal || '';
    document.getElementById('f_kode').value = r.kode || '';
    document.getElementById('f_volume').value = r.volume || '';
    document.getElementById('f_satuan').value = r.satuan || '';
    document.getElementById('f_fk').value = r.fk || '';
    document.getElementById('f_kebutuhan').value = r.kebutuhan || '';
    document.getElementById('f_unit_cost').value = r.unit_cost || '';
    document.getElementById('f_total').value = r.total || '';
  }
  function editRow(rowIdx){
    var r = LK_ROWS[rowIdx];
    if (!r) return;
    fillForm(r);
    document.getElementById('f_row_idx').value = rowIdx;
    document.getElementById('f_after_row_idx').value = '';
    document.getElementById('formTitle').innerHTML = '&#9998; Edit Baris #' + rowIdx;
    document.getElementById('lkForm').scrollIntoView({behavior:'smooth'});
  }
  function addAfter(rowIdx){
    resetForm();
    document.getElementById('f_after_row_idx').value = rowIdx;
    document.getElementById('formTitle').innerHTML = '&#10133; Tambah Baris (setelah baris ini)';
    document.getElementById('lkForm').scrollIntoView({behavior:'smooth'});
    document.getElementById('f_label').focus();
  }
  function autoTotal(){
    var vol = parseFloat(document.getElementById('f_volume').value) || 0;
    var uc = parseFloat(document.getElementById('f_unit_cost').value) || 0;
    var totalEl = document.getElementById('f_total');
    if (!totalEl.value && vol && uc) totalEl.value = Math.round(vol*uc);
  }
</script>
</body>
</html>"""


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8501, debug=True)
