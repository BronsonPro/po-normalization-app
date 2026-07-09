"""
Add this as a new tab/page in your existing Streamlit PO normalization app.
It gives a manual "Run now" button (for on-demand checks) alongside the
automatic once-a-day GitHub Actions run.

Usage in your existing app.py:
    from po_email_fetcher.streamlit_tab import render_po_fetcher_tab
    ...
    with tab_po_fetcher:
        render_po_fetcher_tab()
"""

import streamlit as st
import pandas as pd

from po_email_fetcher.main import run
from po_email_fetcher.sheets_writer import _get_worksheet


def render_po_fetcher_tab():
    st.subheader("PO email fetcher")
    st.caption("Runs automatically every day at 9:00 AM. Use the button below to check on demand.")

    col1, col2 = st.columns([1, 3])
    with col1:
        if st.button("Run now", type="primary"):
            with st.spinner("Fetching PO emails..."):
                summary = run()
            st.success(
                f"Done - {summary['fetched']} PDFs processed, "
                f"{summary['success']} successful, "
                f"{summary['failed']} need review, "
                f"{summary['no_pdf']} had no PDF, "
                f"{summary['unmatched_party']} unmatched sender."
            )

    st.markdown("---")
    st.markdown("**Status log** (most recent first)")

    try:
        ws = _get_worksheet()
        records = ws.get_all_records()
        if records:
            df = pd.DataFrame(records)
            df = df.iloc[::-1]  # most recent first

            def highlight_status(row):
                status = str(row.get("Status", ""))
                if status == "SUCCESS":
                    return ["background-color: #eaf3de"] * len(row)
                elif "FAILED" in status or "UNKNOWN" in str(row.get("Party", "")):
                    return ["background-color: #fcebeb"] * len(row)
                elif "NEEDS REVIEW" in status:
                    return ["background-color: #faeeda"] * len(row)
                return [""] * len(row)

            st.dataframe(df.style.apply(highlight_status, axis=1), use_container_width=True)
        else:
            st.info("No runs logged yet.")
    except Exception as e:
        st.error(f"Could not load status log: {e}")
