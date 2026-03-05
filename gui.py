"""Streamlit GUI for Bell Canada BCECN yield spread automater."""

import os
import tempfile

import streamlit as st

from main import detect_bank, load_env_file, BANK_PARSERS, MASTER_FILE
from excel_writer import append_row, update_charts

# Load .env defaults on startup
load_env_file()

st.set_page_config(page_title="Bell Automater", layout="centered")
st.title("Bell Canada — BCECN Automater")

tab_pdf, tab_outlook = st.tabs(["Process PDFs", "Fetch from Outlook"])

# ── Shared settings in sidebar ──────────────────────────────────────
master_file = st.sidebar.text_input("Master workbook path", value=os.environ.get("MASTER_FILE", MASTER_FILE))
dry_run = st.sidebar.checkbox("Dry run (preview only)")


def _process_pdf_file(pdf_path: str, label: str) -> None:
    """Parse one PDF and either preview or write to Excel."""
    try:
        bank_key = detect_bank(pdf_path)
        parser = BANK_PARSERS[bank_key]
        data = parser(pdf_path)

        if dry_run:
            st.success(f"{label} — {bank_key.upper()} — {data.get('date')}")
            display = {}
            for k, v in data.items():
                if k in ("date", "bank"):
                    continue
                if isinstance(v, float) and ("yield" in k or "coupon" in k):
                    display[k] = f"{v:.3%}"
                elif isinstance(v, (int, float)):
                    display[k] = v
                else:
                    display[k] = str(v) if v is not None else "-"
            st.table(display)
        else:
            append_row(master_file, data)
            update_charts(master_file)
            st.success(f"{label} — {bank_key.upper()} — written to Excel")
    except Exception as exc:
        st.error(f"{label} — {exc}")


# ── Tab 1: Local PDFs ───────────────────────────────────────────────
with tab_pdf:
    uploaded_files = st.file_uploader(
        "Drop PDF files here", type=["pdf"], accept_multiple_files=True
    )

    if st.button("Process", key="btn_pdf") and uploaded_files:
        with tempfile.TemporaryDirectory() as tmp_dir:
            for uf in uploaded_files:
                tmp_path = os.path.join(tmp_dir, uf.name)
                with open(tmp_path, "wb") as f:
                    f.write(uf.getbuffer())
                _process_pdf_file(tmp_path, uf.name)

# ── Tab 2: Outlook fetch ────────────────────────────────────────────
with tab_outlook:
    col1, col2 = st.columns(2)
    with col1:
        email = st.text_input("Email", value=os.environ.get("OUTLOOK_EMAIL", ""))
        server = st.text_input("Server", value=os.environ.get("OUTLOOK_SERVER", "outlook.office365.com"))
    with col2:
        password = st.text_input("Password", type="password", value=os.environ.get("OUTLOOK_PASSWORD", ""))
        days = st.number_input("Days back", min_value=1, value=int(os.environ.get("OUTLOOK_DAYS", "7")))

    sender = st.text_input("Sender filter (optional)", value=os.environ.get("BCECN_SENDER", ""))

    if st.button("Fetch & Process", key="btn_outlook"):
        if not email or not password:
            st.error("Email and password are required.")
        else:
            try:
                from email_fetcher import connect_outlook, fetch_bcecn_pdfs

                with st.spinner("Connecting to Outlook..."):
                    account = connect_outlook(email, password, server=server)
                    pdfs = fetch_bcecn_pdfs(
                        account,
                        sender_filter=sender or None,
                        days_back=days,
                    )

                if not pdfs:
                    st.warning("No PDFs found.")
                else:
                    for pdf_path in pdfs:
                        _process_pdf_file(pdf_path, os.path.basename(pdf_path))
            except Exception as exc:
                st.error(f"Outlook fetch failed: {exc}")
