import streamlit as st
import pandas as pd
from datetime import datetime

from capex_module import generate_capex_report_bytes
from budget_module import (
    convert_budget_to_df,
    salary_analysis_with_ledger,
    build_output_excel_bytes
)

st.set_page_config(page_title="BSNL Finance Tools", layout="wide")

st.title("BSNL Finance Tools")

tool = st.radio(
    "Select Tool:",
    ["CAPEX Report Generator", "Budget Report"],
    horizontal=True
)

st.divider()

# ---------------- CAPEX ----------------
if tool == "CAPEX Report Generator":
    st.subheader("CAPEX Report")

    capex_file = st.file_uploader("Upload CAPEX Excel", type=["xlsx"])

    if capex_file:
        if st.button("Generate CAPEX"):
            out_bytes, out_name = generate_capex_report_bytes(capex_file.read())

            st.success("CAPEX report ready")
            st.download_button("Download", out_bytes, file_name=out_name)


# ---------------- BUDGET ----------------
elif tool == "Budget Report":
    st.subheader("Budget Report (Salary Optional)")

    budget_file = st.file_uploader("Upload Budget Excel (Required)", type=["xlsx"])
    salary_file = st.file_uploader("Upload Salary Excel (Optional)", type=["xlsx"])

    budget_sheet = st.text_input("Budget Sheet (optional)", "")

    salary_sheet = None
    if salary_file:
        xls = pd.ExcelFile(salary_file)
        salary_sheet = st.selectbox("Select Salary Sheet", xls.sheet_names)

    if budget_file:
        if st.button("Run Budget Report"):

            budget_df = convert_budget_to_df(
                budget_file,
                sheet_name=budget_sheet.strip() or None
            )

            if budget_df.empty:
                st.error("Budget not extracted. Check format.")
                st.stop()

            today = datetime.today().strftime("%Y-%m-%d")

            if salary_file and salary_sheet:
                salary_df, ledger_df = salary_analysis_with_ledger(
                    salary_file,
                    budget_df,
                    salary_sheet
                )

                out_bytes = build_output_excel_bytes(
                    budget_df, salary_df, ledger_df
                )

                filename = f"BUDGET_WITH_SALARY_{today}.xlsx"
            else:
                out_bytes = build_output_excel_bytes(budget_df)
                filename = f"BUDGET_ONLY_{today}.xlsx"

            st.success("Report generated")
            st.download_button("Download Output", out_bytes, file_name=filename)


# ---------------- FOOTER ----------------
st.markdown(
    """
    <hr>
    <div style='text-align:center; font-size:14px;'>
        <b>Created by Hrushikesh Kesale | MH Circle</b><br>
        🚴 Follow on Instagram:
        <a href='https://www.instagram.com/cycle_stories4' target='_blank'>
            @cycle_stories4
        </a>
    </div>
    """,
    unsafe_allow_html=True
)
