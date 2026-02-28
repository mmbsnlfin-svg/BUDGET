import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO

# ✅ Same constants as your code
SHEET_RAW = "Sheet1"
SHEET_TEMPLATE = "Sheet2"

MAIN_TYPES = {"RE", "KR"}       # RE and KR are same nature → main invoice
DEDUCT_TYPES = {"RT", "LD"}     # deductions: Batch admin RT + Retention/LD etc.
BATCH_KEYWORD = "BATCH_ADMIN"


def to_float(x):
    try:
        if pd.isna(x):
            return np.nan
        return float(x)
    except Exception:
        return np.nan


def generate_capex_report_bytes(excel_file_bytes: bytes) -> tuple[bytes, str]:
    """
    Input: excel file bytes that contains Sheet1 (raw) + Sheet2 (template)
    Output: (output_xlsx_bytes, output_filename)
    """
    bio = BytesIO(excel_file_bytes)

    # Read data
    df_raw = pd.read_excel(bio, sheet_name=SHEET_RAW)
    bio.seek(0)
    df_template = pd.read_excel(bio, sheet_name=SHEET_TEMPLATE)
    template_cols = list(df_template.columns)

    # Standardize
    df_raw["Document Type"] = df_raw["Document Type"].astype(str).str.strip().str.upper()
    df_raw["User Name"] = df_raw["User Name"].astype(str).str.strip().str.upper()
    df_raw["Vendor Code"] = df_raw["Vendor Code"].astype(str).str.strip()
    df_raw["Document Number"] = df_raw["Document Number"].astype(str).str.strip()
    df_raw["Invoice reference"] = df_raw["Invoice reference"].astype(str).str.strip()

    df_raw["Posting Date"] = pd.to_datetime(df_raw["Posting Date"], errors="coerce")
    df_raw["Amount in local currency"] = df_raw["Amount in local currency"].apply(to_float)

    mains = df_raw[df_raw["Document Type"].isin(MAIN_TYPES)].copy()
    ded = df_raw[df_raw["Document Type"].isin(DEDUCT_TYPES)].copy()

    results = []
    exceptions = []

    for _, m in mains.iterrows():
        main_doc = m["Document Number"]
        vendor_code = m["Vendor Code"]

        linked = ded[ded["Invoice reference"] == main_doc].copy()

        # Batch Admin by User Name
        batch_rows = linked[linked["User Name"].str.contains(BATCH_KEYWORD, na=False)].copy()

        batch_doc = ""
        batch_amt = np.nan

        if len(batch_rows) > 0:
            batch_rows = batch_rows.sort_values(["Posting Date", "Document Number"])
            br = batch_rows.iloc[0]
            batch_doc = str(br["Document Number"])
            batch_amt = br["Amount in local currency"]
            linked = linked[linked["Document Number"].astype(str) != batch_doc].copy()
        else:
            exceptions.append({"Invoice Doc": main_doc, "Vendor Code": vendor_code, "Issue": "BATCH_ADMIN deduction not found"})

        # Remaining deductions include RT/LD sorted
        linked = linked.sort_values(["Posting Date", "Document Number"])

        other_docs = linked["Document Number"].astype(str).tolist()
        other_amts = linked["Amount in local currency"].tolist()

        d2_doc = other_docs[0] if len(other_docs) > 0 else ""
        d2_amt = other_amts[0] if len(other_amts) > 0 else np.nan

        d3_doc = other_docs[1] if len(other_docs) > 1 else ""
        d3_amt = other_amts[1] if len(other_amts) > 1 else np.nan

        d4_doc = other_docs[2] if len(other_docs) > 2 else ""
        d4_amt = other_amts[2] if len(other_amts) > 2 else np.nan

        if len(other_docs) > 3:
            exceptions.append({"Invoice Doc": main_doc, "Vendor Code": vendor_code, "Issue": f"More than 4 deductions found (Batch + {len(other_docs)} others). Extra ignored."})

        gross = abs(m["Amount in local currency"]) if not pd.isna(m["Amount in local currency"]) else np.nan
        total_ded = sum([x for x in [batch_amt, d2_amt, d3_amt, d4_amt] if not pd.isna(x)])
        net = gross - total_ded if not pd.isna(gross) else np.nan

        out = {c: "" for c in template_cols}

        if "Business Area Code" in out:
            out["Business Area Code"] = m.get("Business Area", "")
        if "Vendor code" in out:
            out["Vendor code"] = vendor_code
        if "Vendor Name" in out:
            out["Vendor Name"] = m.get("Vendor Name", "")
        if "Document Type" in out:
            out["Document Type"] = m.get("Document Type", "")
        if "Invoice Document No" in out:
            out["Invoice Document No"] = main_doc
        if "PO No" in out:
            out["PO No"] = m.get("Purchasing Document", "")
        if "Posting date" in out:
            out["Posting date"] = m.get("Posting Date", "")

        if "Gross Invoice Value (Basic amount + Input tax credit)" in out:
            out["Gross Invoice Value (Basic amount + Input tax credit)"] = gross

        if "Net Invoice Value payable to vendor" in out:
            out["Net Invoice Value payable to vendor"] = net

        if "Net payable To Vendor" in out:
            out["Net payable To Vendor"] = net

        if "Retention Document No/LD Document /Any other Deduction -1" in out:
            out["Retention Document No/LD Document /Any other Deduction -1"] = batch_doc
        if "Retention/LD/Any other Deduction -1 Amount" in out:
            out["Retention/LD/Any other Deduction -1 Amount"] = batch_amt

        if "Retention Document No/LD Document /Any other Deduction -2" in out:
            out["Retention Document No/LD Document /Any other Deduction -2"] = d2_doc
        if "Retention/LD/Any other Deduction -2 Amount" in out:
            out["Retention/LD/Any other Deduction -2 Amount"] = d2_amt

        if "Retention Document No/LD Document /Any other Deduction -3" in out:
            out["Retention Document No/LD Document /Any other Deduction -3"] = d3_doc
        if "Retention/LD/Any other Deduction -3 Amount" in out:
            out["Retention/LD/Any other Deduction -3 Amount"] = d3_amt

        if "Retention Document No/LD Document /Any other Deduction -4" in out:
            out["Retention Document No/LD Document /Any other Deduction -4"] = d4_doc
        if "Retention/LD/Any other Deduction -4 Amount" in out:
            out["Retention/LD/Any other Deduction -4 Amount"] = d4_amt

        if "Vendor codeInvoice Document No" in out:
            out["Vendor codeInvoice Document No"] = str(vendor_code) + str(main_doc)

        results.append(out)

    report_df = pd.DataFrame(results, columns=template_cols)

    today_str = datetime.today().strftime("%Y-%m-%d")
    output_filename = f"CAPEX_{today_str}.xlsx"

    out_bio = BytesIO()
    with pd.ExcelWriter(out_bio, engine="openpyxl") as writer:
        df_raw.to_excel(writer, sheet_name=SHEET_RAW, index=False)
        report_df.to_excel(writer, sheet_name=SHEET_TEMPLATE, index=False)
        if len(exceptions) > 0:
            pd.DataFrame(exceptions).to_excel(writer, sheet_name="Exceptions", index=False)

    return out_bio.getvalue(), output_filename
