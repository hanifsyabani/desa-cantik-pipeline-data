"""Streamlit entrypoint for the Desa Cantik data pipeline."""

import sys
from pathlib import Path

import streamlit as st

PROJECT_ROOT = Path(__file__).resolve().parent
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from desa_cantik.config import (  # noqa: E402
    DASHBOARD_CLEANED_SHEET,
    DASHBOARD_RAW_SHEET,
    QUALITY_SHEET,
    SRC_SHEET,
    WILAYAH_SHEET,
)
from desa_cantik.pipeline import generate_bytes  # noqa: E402


st.set_page_config(page_title="Pipeline Data Harga", page_icon=":bar_chart:", layout="centered")

st.title("Pipeline Data Harga Kebutuhan Pokok")
st.caption(
    f"File input harus berisi sheet '{SRC_SHEET}', '{WILAYAH_SHEET}', dan '{QUALITY_SHEET}'."
)

uploaded_file = st.file_uploader("Pilih file Excel (.xlsx)", type=["xlsx"])

min_harga = st.number_input(
    "Harga minimum",
    min_value=0,
    value=1000,
    step=100,
    help="Baris dengan harga di bawah angka ini tidak masuk ke sheet cleaned.",
)

if uploaded_file is None:
    st.info("Upload file .xlsx untuk mulai.")
else:
    if st.button("Proses", type="primary"):
        progress_bar = st.progress(0.0, text="Memulai...")

        def progress_cb(frac, nama_komoditas):
            progress_bar.progress(frac, text=f"Memproses: {nama_komoditas}")

        try:
            output_bytes, stats = generate_bytes(
                uploaded_file.getvalue(),
                min_harga=int(min_harga),
                progress_cb=progress_cb,
            )
        except ValueError as exc:
            st.error(str(exc))
        except KeyError as exc:
            st.error(f"Sheet tidak ditemukan atau format referensi tidak sesuai: {exc}")
        else:
            progress_bar.progress(1.0, text="Selesai")
            st.success("Pipeline selesai dijalankan.")

            col1, col2 = st.columns(2)
            with col1:
                st.metric("Total responden", stats["n_rows"])
                st.metric(f"Baris {DASHBOARD_RAW_SHEET}", stats["n_raw"])
            with col2:
                st.metric(f"Baris {DASHBOARD_CLEANED_SHEET}", stats["n_clean"])
                st.metric(
                    "Total dibuang",
                    stats["n_blank"] + stats["n_dropped_wilayah"] + stats["n_dropped_price"],
                )

            with st.expander("Rincian baris yang dibuang"):
                st.write(f"- Tanpa jawaban: **{stats['n_blank']}**")
                st.write(f"- Wilayah gagal lookup (#N/A): **{stats['n_dropped_wilayah']}**")
                st.write(f"- Harga di bawah {int(min_harga)}: **{stats['n_dropped_price']}**")

            out_name = uploaded_file.name.rsplit(".", 1)[0] + "_hasil.xlsx"
            st.download_button(
                "Download hasil (.xlsx)",
                data=output_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

st.divider()
st.caption(
    f"Output berisi sheet '{DASHBOARD_RAW_SHEET}' dan '{DASHBOARD_CLEANED_SHEET}'."
)
