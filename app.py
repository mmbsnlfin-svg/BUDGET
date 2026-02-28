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

st.title("BSNL Finance Tools (Single Page)")

tool = st.radio(
    "Select which tool to run:",
    ["CAPEX Report Generator", "Budget Report (Salary Optional)"],
    horizontal=True
)

st.divider()

# -------------------- CAPEX TOOL --------------------
if tool == "CAPEX Report Generator":
    st.subheader("CAPEX Report Generator")
    st.caption("Upload ONE Excel file that contains Sheet1 (Raw) + Sheet2 (Template).")

    capex_file = st.file_uploader("Upload CAPEX Excel", type=["xlsx"], key="capex")

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

# -------------------- BUDGET TOOL --------------------
elif tool == "Budget Report (Salary Optional)":
    st.subheader("Budget Report")
    st.caption("Upload Budget Excel (Required). Salary Excel is Optional.")

    budget_file = st.file_uploader("Upload Budget Excel (Required)", type=["xlsx"], key="budget")
    salary_file = st.file_uploader("Upload Salary Excel (Optional)", type=["xlsx"], key="salary")

    budget_sheet_name = st.text_input(
        "Budget sheet name (Optional) - leave blank for ALL sheets",
        value=""
    )

    salary_sheet_name = None
    if salary_file is not None:
        try:
            xls_sal = pd.ExcelFile(salary_file, engine="openpyxl")
            salary_sheet_name = st.selectbox(
                "Select Salary Sheet (only if Salary uploaded)",
                options=xls_sal.sheet_names,
                index=0
            )
        except Exception as e:
            st.error(f"Unable to read Salary file sheets: {e}")
            st.stop()

    if budget_file is not None:
        if st.button("Run Budget Report", type="primary"):
            # Build budget df always
            budget_df = convert_budget_to_df(budget_file, sheet_name=budget_sheet_name.strip() or None)

            if budget_df.empty:
                st.error("No budget data extracted. Please check Budget file format.")
                st.stop()

            # If salary uploaded -> include salary analysis, else only budget output
            today = datetime.today().strftime("%Y-%m-%d")

            if salary_file is not None and salary_sheet_name is not None:
                salary_df, ledger_df = salary_analysis_with_ledger(
                    salary_file,
                    budget_df,
                    salary_sheet=salary_sheet_name
                )
                out_bytes = build_output_excel_bytes(budget_df, salary_df, ledger_df)
                out_name = f"BUDGET_WITH_SALARY_{today}.xlsx"
                st.success("Budget + Salary report generated.")
            else:
                out_bytes = build_output_excel_bytes(budget_df)
                out_name = f"BUDGET_ONLY_{today}.xlsx"
                st.success("Budget-only report generated (Salary not uploaded).")

            st.download_button(
                "Download Output Excel",
                data=out_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
