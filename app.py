import streamlit as st
from datetime import datetime

from capex_module import generate_capex_report_bytes
from budget_module import (
    convert_budget_to_df,
    salary_analysis_with_ledger,
    build_output_excel_bytes
)

st.set_page_config(page_title="BSNL Tools", layout="wide")

st.title("BSNL Finance Tools")
tool = st.radio(
    "Select which tool to run:",
    ["CAPEX Report Generator", "Budget + Salary Diversion Analysis"],
    horizontal=True
)

st.divider()

if tool == "CAPEX Report Generator":
    st.subheader("CAPEX Report Generator")
    st.caption("Upload ONE Excel file that contains Sheet1 (raw) + Sheet2 (template).")

    capex_file = st.file_uploader("Upload CAPEX Excel", type=["xlsx"])

    if capex_file is not None:
        if st.button("Generate CAPEX Report", type="primary"):
            out_bytes, out_name = generate_capex_report_bytes(capex_file.read())
            st.success("CAPEX report generated.")
            st.download_button(
                "Download CAPEX Output",
                data=out_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

elif tool == "Budget + Salary Diversion Analysis":
    st.subheader("Budget + Salary Diversion Analysis")
    st.caption("Upload Budget Excel + Salary Excel. Output includes BUDGET_ALL, FundCenter sheets, SALARY_ANALYSIS, DONOR_LEDGER.")

    c1, c2 = st.columns(2)
    with c1:
        budget_file = st.file_uploader("Upload Budget Excel", type=["xlsx"], key="budget")
    with c2:
        salary_file = st.file_uploader("Upload Salary Excel", type=["xlsx"], key="salary")

    sheet_name = st.text_input("Budget sheet name (optional, leave blank for ALL sheets)", value="")

    if budget_file and salary_file:
        if st.button("Run Budget Analysis", type="primary"):
            # Budget parsing
            budget_df = convert_budget_to_df(budget_file, sheet_name=sheet_name.strip() or None)

            if budget_df.empty:
                st.error("No budget data extracted. Please check Budget file format.")
                st.stop()

            # Salary diversion + ledger
            salary_df, ledger_df = salary_analysis_with_ledger(salary_file, budget_df)

            # Build output excel bytes
            out_bytes = build_output_excel_bytes(budget_df, salary_df, ledger_df)

            today = datetime.today().strftime("%Y-%m-%d")
            out_name = f"BUDGET_ANALYSIS_{today}.xlsx"

            st.success("Budget analysis generated.")
            st.download_button(
                "Download Budget Analysis Output",
                data=out_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
