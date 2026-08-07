"""
Konverter Laporan R-5401 (Bank) -> Excel
Aplikasi Streamlit: upload file .txt laporan harian, validasi, unduh Excel.
"""
import io
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st

from parser import parse_report, build_workbook, build_combined_workbook

st.set_page_config(page_title="Konverter R-5401 → Excel", page_icon="📊", layout="wide")

# ---------- header ----------
st.title("📊 Konverter Laporan R-5401 → Excel")
st.caption(
    "Upload file laporan transaksi harian (.txt format lebar-tetap dari bank), "
    "aplikasi akan mem-parsing, memvalidasi terhadap total footer, dan menyiapkan "
    "file Excel rapi (Data + Ringkasan) untuk diunduh."
)


def rupiah(n):
    return "Rp " + f"{int(round(n)):,}".replace(",", ".")


def wb_to_bytes(wb):
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


uploaded = st.file_uploader(
    "Pilih satu atau beberapa file laporan (.txt)",
    type=["txt"],
    accept_multiple_files=True,
    help="Bisa upload banyak file sekaligus — tiap tanggal / kode perusahaan.",
)

if not uploaded:
    st.info("⬆️ Silakan upload minimal satu file .txt untuk mulai.")
    st.stop()

datasets = []          # (meta, rows) untuk file valid & unik
seen = set()           # (kode, tanggal) untuk deteksi duplikat

st.divider()

for uf in uploaded:
    raw = uf.getvalue()
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError:
        text = raw.decode("latin-1", errors="replace")

    meta, rows = parse_report(text)

    with st.container(border=True):
        c_title, c_badge = st.columns([4, 1])
        c_title.subheader(f"📄 {uf.name}")

        if not rows:
            c_badge.error("Gagal")
            st.warning(
                "Tidak ada baris transaksi terbaca. Pastikan ini file laporan R-5401 "
                "dengan format lebar-tetap yang benar."
            )
            continue

        parsed_total = sum(r[3] for r in rows)
        parsed_count = len(rows)
        ft, fc = meta["footer_total"], meta["footer_count"]

        key = (meta["kode"], meta["tanggal_label"])
        is_dup = key in seen
        seen.add(key)

        # --- validasi ---
        total_ok = ft is not None and abs(parsed_total - ft) < 0.01
        count_ok = fc is not None and parsed_count == fc
        if total_ok and count_ok:
            c_badge.success("✓ Valid")
        elif ft is None and fc is None:
            c_badge.warning("Tanpa footer")
        else:
            c_badge.error("Selisih!")

        # --- metadata + metrics ---
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Kode Perusahaan", meta["kode"] or "—")
        m2.metric("Tanggal", meta["tanggal_label"] or "—")
        m3.metric("Jumlah Transaksi", f"{parsed_count:,}".replace(",", "."))
        m4.metric("Total Nilai", rupiah(parsed_total))

        if meta["nama_pt"]:
            st.caption(f"**{meta['nama_pt']}**  •  {meta['cabang']}")

        # --- validasi detail ---
        if not total_ok and ft is not None:
            st.error(
                f"⚠️ Total hasil parsing ({rupiah(parsed_total)}) TIDAK cocok dengan "
                f"total footer laporan ({rupiah(ft)}). Selisih {rupiah(abs(parsed_total-ft))}."
            )
        if not count_ok and fc is not None:
            st.error(
                f"⚠️ Jumlah transaksi hasil parsing ({parsed_count}) tidak cocok dengan "
                f"footer ({fc})."
            )
        if is_dup:
            st.warning(
                "🔁 File ini duplikat (kode & tanggal sama dengan file lain). "
                "Tetap bisa diunduh, tapi **tidak** ikut dalam file gabungan agar total tidak dobel."
            )

        # --- preview tabel ---
        df = pd.DataFrame(
            rows,
            columns=["No.", "No. Pelanggan/TXN", "Nama", "Nilai", "Tgl",
                     "Waktu", "Jam", "Lokasi", "Keterangan 1", "Keterangan 2"],
        )
        with st.expander(f"👁️ Lihat data ({parsed_count} baris)"):
            st.dataframe(df, use_container_width=True, hide_index=True)

        # --- tombol unduh per file ---
        wb = build_workbook(rows, meta)
        fname = f"Transaksi_R-5401_{meta['kode'] or 'NA'}_{(meta['tanggal_label'] or '').replace('/','')}.xlsx"
        st.download_button(
            "⬇️ Unduh Excel file ini",
            data=wb_to_bytes(wb),
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{uf.name}",
        )

        if not is_dup:
            datasets.append((meta, rows))

# ---------- gabungan ----------
if len(datasets) > 1:
    st.divider()
    st.subheader("📦 Unduh Gabungan")
    grand = sum(sum(r[3] for r in rows) for _, rows in datasets)
    total_txn = sum(len(rows) for _, rows in datasets)
    st.write(
        f"**{len(datasets)} laporan** unik  •  **{total_txn:,}** transaksi  •  "
        f"total **{rupiah(grand)}**".replace(",", ".")
    )

    cga, cgb = st.columns(2)

    # 1 workbook gabungan
    wb_comb = build_combined_workbook(datasets)
    cga.download_button(
        "⬇️ Unduh 1 Excel gabungan (Data + Rekap Harian)",
        data=wb_to_bytes(wb_comb),
        file_name=f"Transaksi_R-5401_GABUNGAN_{datetime.now():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # zip semua file terpisah
    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
        for meta, rows in datasets:
            wb = build_workbook(rows, meta)
            name = f"Transaksi_R-5401_{meta['kode']}_{meta['tanggal_label'].replace('/','')}.xlsx"
            zf.writestr(name, wb_to_bytes(wb))
    zbuf.seek(0)
    cgb.download_button(
        "⬇️ Unduh semua (.zip berisi file terpisah)",
        data=zbuf.getvalue(),
        file_name=f"Transaksi_R-5401_semua_{datetime.now():%Y%m%d}.zip",
        mime="application/zip",
    )

st.divider()
st.caption(
    "Catatan: Nama pelanggan & keterangan pada laporan sumber terpotong "
    "(field lebar-tetap ±16 karakter). Nilai, tanggal, waktu, dan lokasi akurat 100%."
)
