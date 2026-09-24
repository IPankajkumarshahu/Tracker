"""Web page for converting bank statement PDFs to Excel.

Run:  streamlit run bank_statement_app.py
Files are processed on this computer only; nothing is uploaded anywhere else.
"""

import io
import tempfile

import streamlit as st

from bank_statement import StatementError, parse_statement, write_workbook

st.set_page_config(page_title="Bank Statement to Excel", page_icon="🏦", layout="wide")
st.title("🏦 Bank Statement → Excel")
st.caption("Upload one or more bank statement PDFs (downloaded from net banking). "
           "You get one Excel file with a Summary sheet and one sheet per statement.")

files = st.file_uploader("Statement PDFs", type=["pdf"], accept_multiple_files=True)
password = st.text_input("PDF password (only if the statements are locked)", type="password")

if files and st.button("Convert to Excel", type="primary"):
    statements = []
    for f in files:
        with tempfile.NamedTemporaryFile(suffix=".pdf") as tmp:
            tmp.write(f.getvalue())
            tmp.flush()
            try:
                stmt = parse_statement(tmp.name, password=password or None, source=f.name)
            except StatementError as e:
                st.error(f"**{f.name}**: {e}")
                continue
        statements.append(stmt)
        with st.expander(f"✅ {f.name} — {stmt.bank or 'bank not detected'}, "
                         f"{len(stmt.transactions)} transactions", expanded=len(files) == 1):
            c = st.columns(4)
            c[0].metric("Opening balance", f"{stmt.opening_balance or 0:,.2f}")
            c[1].metric("Total debit", f"{stmt.total_debit:,.2f}")
            c[2].metric("Total credit", f"{stmt.total_credit:,.2f}")
            c[3].metric("Closing balance", f"{stmt.closing_balance or 0:,.2f}")
            for w in stmt.warnings:
                st.warning(w)
            amt = lambda v: "" if v is None else f"{v:,.2f}"  # noqa: E731
            st.dataframe([{
                "Date": t.date.strftime("%d-%m-%Y"), "Description": t.description, "Ref": t.ref,
                "Debit": amt(t.debit), "Credit": amt(t.credit), "Balance": amt(t.balance),
                "Check": t.check,
            } for t in stmt.transactions], use_container_width=True, hide_index=True)

    if statements:
        buf = io.BytesIO()
        write_workbook(statements, buf)
        name = (files[0].name.rsplit(".", 1)[0] if len(statements) == 1 else "bank_statements") + ".xlsx"
        st.download_button("⬇️ Download Excel", buf.getvalue(), file_name=name, type="primary",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
