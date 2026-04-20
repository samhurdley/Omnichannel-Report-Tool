import io
import zipfile
from datetime import datetime
import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from process_report import process_csv, extract_report_month, find_client_config

st.set_page_config(
    page_title="Report Processor",
    page_icon="📊",
    layout="centered",
)

st.title("📊 Omnichannel Report Processor")
st.markdown("Upload the client config and all CSV exports for the month. Reports download as a ZIP.")

HISTORY_WS   = "History"
HISTORY_COLS = ['Client', 'Month', 'Impressions', 'Clicks', 'Spend', 'Conversions', 'Revenue']
_SCOPES      = ["https://spreadsheets.google.com/feeds",
                "https://www.googleapis.com/auth/drive"]


def _get_gsheet():
    try:
        cfg = dict(st.secrets["connections"]["gsheets"])
        url = cfg.pop("spreadsheet")
        cfg.pop("worksheet", None)
        creds = Credentials.from_service_account_info(cfg, scopes=_SCOPES)
        gc = gspread.authorize(creds)
        sh = gc.open_by_url(url)
        try:
            return sh.worksheet(HISTORY_WS)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(HISTORY_WS, rows=1000, cols=len(HISTORY_COLS))
            ws.append_row(HISTORY_COLS)
            return ws
    except Exception:
        return None


def _load_history(ws):
    if ws is None:
        return pd.DataFrame(columns=HISTORY_COLS)
    try:
        records = ws.get_all_records()
        if not records:
            return pd.DataFrame(columns=HISTORY_COLS)
        df = pd.DataFrame(records)
        for col in ['Impressions', 'Clicks', 'Spend', 'Conversions', 'Revenue']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        return df
    except Exception:
        return pd.DataFrame(columns=HISTORY_COLS)


def _save_history(ws, history_df):
    if ws is None:
        return
    try:
        rows = [history_df.columns.tolist()] + history_df.fillna('').values.tolist()
        ws.clear()
        ws.update('A1', rows)
    except Exception as e:
        st.warning(f"History not saved: {e}")


def _prev_month_str(report_month):
    try:
        dt = datetime.strptime(report_month, '%B %Y')
        prev = dt.replace(year=dt.year - 1, month=12) if dt.month == 1 \
               else dt.replace(month=dt.month - 1)
        return prev.strftime('%B %Y')
    except Exception:
        return None


def _get_prev_data(history_df, client_name, report_month):
    if history_df.empty or 'Client' not in history_df.columns:
        return None
    prev_month = _prev_month_str(report_month)
    if not prev_month:
        return None
    rows = history_df[
        (history_df['Client'] == client_name) & (history_df['Month'] == prev_month)
    ]
    if rows.empty:
        return None
    r = rows.iloc[0]
    return {k: float(r.get(k, 0) or 0)
            for k in ['Impressions', 'Clicks', 'Spend', 'Conversions', 'Revenue']}


def _upsert_history(history_df, totals_row):
    if not history_df.empty and 'Client' in history_df.columns:
        history_df = history_df[
            ~((history_df['Client'] == totals_row['Client']) &
              (history_df['Month']  == totals_row['Month']))
        ]
    return pd.concat([history_df, pd.DataFrame([totals_row])], ignore_index=True)


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
        ws = _get_gsheet()
        history_df = _load_history(ws)

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
                    csv_bytes = csv_file.read()
                    config_row = find_client_config(csv_file.name, config_df)
                    client_candidate = str(config_row['Client Name'])
                    report_month = extract_report_month(csv_file.name)
                    prev_data = _get_prev_data(history_df, client_candidate, report_month)

                    client_name, html_str, totals = process_csv(
                        csv_bytes, csv_file.name, config_df, prev_data
                    )
                    zf.writestr(f"{client_name}_report.html", html_str.encode("utf-8"))
                    results.append(client_name)

                    history_df = _upsert_history(history_df, totals)
                    _save_history(ws, history_df)

                except Exception as e:
                    errors.append((csv_file.name, str(e)))

        progress_bar.empty()

        if results:
            st.success(f"✓ {len(results)} report(s) processed: {', '.join(results)}")

        for fname, err in errors:
            st.error(f"✗ {fname}: {err}")

        if results:
            zip_buf.seek(0)
            st.download_button(
                label="⬇ Download all reports (.zip)",
                data=zip_buf,
                file_name="reports.zip",
                mime="application/zip",
                use_container_width=True,
            )

else:
    st.info("Upload client_config.xlsx and at least one CSV to get started.")
