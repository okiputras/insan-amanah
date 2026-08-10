"""
Konverter Laporan R-5401 (Bank) -> Excel — versi ringan (Flask).

Upload file .txt laporan harian, validasi terhadap footer, unduh Excel rapi.
Tanpa Streamlit / numpy / pandas / pyarrow — hanya Flask + openpyxl, memakai
HTTP request/response biasa (tanpa WebSocket) supaya ringan & stabil di Railway.
"""
import io
import time
import zipfile
import secrets
import threading
from datetime import datetime

from flask import (
    Flask, request, render_template_string, send_file, abort, redirect, url_for,
)

from parser import parse_report, build_workbook, build_combined_workbook
from validator import parse_master_siswa, reconcile_pembayaran, build_recon_workbook

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 40 * 1024 * 1024   # batas total 1x upload: 40 MB
PREVIEW_LIMIT = 100                                    # baris maksimum di tabel preview

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


# ---------- menu: rekap pembayaran R-5401 vs master siswa (hanya Sesuai) ----------
@app.route("/rekap", methods=["GET"])
def rekap():
    return render_template_string(REKAP_PAGE, result=None, token=None, error=None, active="rekap")


@app.route("/rekap/proses", methods=["POST"])
def rekap_proses():
    f_master = request.files.get("master")
    f_laporan = request.files.get("laporan")

    def _err(msg):
        return render_template_string(REKAP_PAGE, result=None, token=None, error=msg, active="rekap")

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

    meta, report_rows = parse_report(_decode(f_laporan.read()))
    if not report_rows:
        return _err("Tidak ada baris transaksi terbaca dari file laporan. Pastikan ini file laporan "
                    "R-5401 dengan format lebar-tetap yang benar.")

    rows = reconcile_pembayaran(report_rows, master)
    shown = [r for r in rows if r["status"] == "Sesuai"]

    preview = [{
        "no_pelanggan": r["no_pelanggan"], "nama": r["nama"], "lokasi": r["lokasi"],
        "tgl": r["tgl"].strftime("%d/%m/%Y"), "waktu": r["waktu"].strftime("%H:%M:%S"),
        "kegiatan": ribuan(r["kegiatan"]), "tabungan": ribuan(r["tabungan"]), "lainnya": ribuan(r["lainnya"]),
        "total_tagihan": ribuan(r["total_tagihan"]), "nilai_bayar": ribuan(r["nilai_bayar"]),
        "ket1": r["ket1"], "ket2": r["ket2"],
    } for r in shown[:PREVIEW_LIMIT]]

    tgl = (meta.get("tanggal_label") or "").replace("/", "")
    dname = f"Rekap_Sesuai_R-5401_{meta.get('kode') or 'NA'}_{tgl or 'NA'}.xlsx"
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
    return render_template_string(REKAP_PAGE, result=result, token=token, error=None, active="rekap")


@app.route("/dl/<token>/rekap")
def dl_rekap(token):
    entry = _get(token)
    return _send(entry["rekap"] if entry else None)


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
    <a href="{{ url_for('rekap') }}" class="{{ 'active' if active=='rekap' else '' }}">Rekap vs Master Siswa</a>
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
<title>Rekap Pembayaran vs Master Siswa</title>
<style>""" + STYLE + """</style>
</head>
<body>
<div class="wrap">
  <div class="nav">
    <a href="{{ url_for('index') }}" class="{{ 'active' if active=='convert' else '' }}">Konversi R-5401</a>
    <a href="{{ url_for('rekap') }}" class="{{ 'active' if active=='rekap' else '' }}">Rekap vs Master Siswa</a>
  </div>
  <h1>&#128203; Rekap Pembayaran vs Master Siswa</h1>
  <p class="sub">Upload <strong>master siswa</strong> (.xlsx: NO VA, NAMA, BPP, KEGIATAN, TABUNGAN) dan
     <strong>laporan harian R-5401</strong> (.txt lebar-tetap) untuk mencocokkan tiap transaksi
     (No. Pelanggan = NO VA) ke tagihannya. Hasil Excel hanya berisi transaksi berstatus
     <strong>Sesuai</strong> (Nilai Bayar = Total Tagihan).</p>

  <form class="up" method="post" action="{{ url_for('rekap_proses') }}" enctype="multipart/form-data">
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
              <th>Kegiatan</th><th>Tabungan</th><th>Lainnya</th>
              <th>Total Tagihan</th><th>Nilai Bayar</th><th>Ket 1</th><th>Ket 2</th>
            </tr></thead>
            <tbody>
              {% for row in result.preview %}
                <tr>
                  <td>{{ row.no_pelanggan }}</td><td>{{ row.nama }}</td><td>{{ row.tgl }}</td>
                  <td>{{ row.waktu }}</td><td>{{ row.lokasi }}</td>
                  <td class="num">{{ row.kegiatan }}</td><td class="num">{{ row.tabungan }}</td><td class="num">{{ row.lainnya }}</td>
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
  <p class="foot">Catatan: pencocokan berdasarkan No. Pelanggan (laporan) = NO VA (master siswa).
     Hanya transaksi berstatus Sesuai yang masuk file Excel; sheet Ringkasan tetap merangkum
     seluruh transaksi termasuk Kurang/Lebih/tanpa data master.</p>
</div>
</body>
</html>"""


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8501, debug=True)
