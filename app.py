import io
import zipfile
import pandas as pd
import streamlit as st
from process_report import process_csv

st.set_page_config(
    page_title="Report Processor",
    page_icon="📊",
    layout="centered",
)

st.title("📊 Omnichannel Report Processor")
st.markdown("Upload the client config and all CSV exports for the month. Reports download as a ZIP.")

# ── Step 1: config ────────────────────────────────────────────────────────────
st.subheader("1. Client config")
config_file = st.file_uploader("Upload client_config.xlsx", type=["xlsx"])

# ── Step 2: CSVs ──────────────────────────────────────────────────────────────
st.subheader("2. CSV reports")
csv_files = st.file_uploader(
    "Upload one or more Trade Desk CSV exports",
    type=["csv"],
    accept_multiple_files=True,
)

# ── Process ───────────────────────────────────────────────────────────────────
if config_file and csv_files:
    st.markdown(f"**{len(csv_files)} file(s) ready.**")

    if st.button("Process all reports", type="primary", use_container_width=True):
        config_df = pd.read_excel(config_file)

        results, errors = [], []
        zip_buf = io.BytesIO()

        progress_bar = st.progress(0, text="Starting…")

        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, csv_file in enumerate(csv_files):
                progress_bar.progress(
                    (i + 1) / len(csv_files),
                    text=f"Processing {csv_file.name}…",
                )
                try:
                    client_name, excel_bytes, html_str = process_csv(
                        csv_file.read(), csv_file.name, config_df
                    )
                    zf.writestr(f"{client_name}_summary.xlsx", excel_bytes)
                    zf.writestr(f"{client_name}_report.html", html_str.encode("utf-8"))
                    results.append(client_name)
                except Exception as e:
                    errors.append((csv_file.name, str(e)))

        progress_bar.empty()

        # ── Results ───────────────────────────────────────────────────────────
        if results:
            st.success(f"✓ {len(results)} report(s) processed: {', '.join(results)}")

        for fname, err in errors:
            st.error(f"✗ {fname}: {err}")

        if results:
            zip_buf.seek(0)
            st.download_button(
                label=f"⬇ Download all reports (.zip)",
                data=zip_buf,
                file_name="reports.zip",
                mime="application/zip",
                use_container_width=True,
            )

else:
    st.info("Upload client_config.xlsx and at least one CSV to get started.")
