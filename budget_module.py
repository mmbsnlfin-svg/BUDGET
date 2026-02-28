import pandas as pd
import re

re_fc_header_2star = re.compile(r"^\s*\*\*\s*([FG]\d{4})\b")
re_fc_star_line    = re.compile(r"^\s*\*\s*([FG]\d{4})\b")
re_item            = re.compile(r"^\s*([A-Z]\d{5})\s+(.+)$")
re_3star_bsnl      = re.compile(r"^\s*\*\*\*\s*BSNL\b", re.I)

def _to_number(x) -> float:
    if pd.isna(x):
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if not s:
        return 0.0
    s = s.replace(",", "")
    try:
        return float(s)
    except Exception:
        return 0.0


def convert_budget_to_df(budget_file, sheet_name: str | None = None) -> pd.DataFrame:
    xls = pd.ExcelFile(budget_file, engine="openpyxl")
    sheets = [sheet_name] if sheet_name else xls.sheet_names

    rows = []
    for sheet in sheets:
        df = pd.read_excel(budget_file, sheet_name=sheet, header=None, engine="openpyxl")
        if df.empty:
            continue

        particulars_col = 1

        # find "Available Budge" column
        available_col = df.shape[1] - 1
        for r in range(min(30, len(df))):
            row_vals = df.iloc[r].astype(str).str.strip().tolist()
            for c, v in enumerate(row_vals):
                if v.lower() == "available budge":
                    available_col = c
                    break

        current_fc = None
        inside_section = False

        for i in range(len(df)):
            cell = df.iat[i, particulars_col]
            if not isinstance(cell, str):
                continue
            txt = cell.strip()
            if not txt:
                continue

            if re_3star_bsnl.match(txt):
                current_fc = None
                inside_section = False
                continue

            if txt.startswith("**"):
                m = re_fc_header_2star.match(txt)
                if m:
                    current_fc = m.group(1)
                    inside_section = False
                continue

            if txt.startswith("*") and not txt.startswith("**") and not txt.startswith("***"):
                m = re_fc_star_line.match(txt)
                if m:
                    current_fc = m.group(1)
                    inside_section = True
                continue

            if inside_section and current_fc:
                m = re_item.match(txt)
                if m:
                    comm_code = m.group(1).strip()
                    text_val = m.group(2).strip()
                    budget_val = _to_number(df.iat[i, available_col])
                    rows.append({
                        "Fund Center": current_fc,
                        "Comm. Code": comm_code,
                        "TEXT": text_val,
                        "Fund CenterComm. Code": f"{current_fc}{comm_code}",
                        "Budget Available": budget_val
                    })

    out = pd.DataFrame(rows)
    if out.empty:
        return out

    out = (out.groupby(["Fund Center", "Comm. Code"], as_index=False)
              .agg({"TEXT": "first", "Fund CenterComm. Code": "first", "Budget Available": "sum"}))
    return out


def salary_analysis_with_ledger(salary_file, budget_df: pd.DataFrame,
                               salary_sheet: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    sal1 = pd.read_excel(salary_file, sheet_name=salary_sheet, engine="openpyxl")

    budget_map = {(r["Fund Center"], r["Comm. Code"]): float(r["Budget Available"])
                  for _, r in budget_df.iterrows()}
    remaining = budget_map.copy()
    donor_ledger = []

    def pick_donors(comm_code: str, receiver_fc: str, prefix: str):
        donors = []
        for (fc, cc), rem_amt in remaining.items():
            if cc == comm_code and fc != receiver_fc and str(fc).startswith(prefix) and rem_amt > 0:
                donors.append((fc, rem_amt))
        donors.sort(key=lambda x: x[1], reverse=True)
        return donors

    records = []

    for _, row in sal1.iterrows():
        ba = str(row.get("BA CODE", "")).strip()
        comm = str(row.get("Commitment Code", "")).strip()
        amt = _to_number(row.get("AMOUNT", 0))

        receiver_fc = f"F{ba}" if ba and not ba.startswith(("F", "G")) else ba
        original_available = budget_map.get((receiver_fc, comm), None)

        if original_available is None:
            records.append({
                "Particular": row.get("Particular", ""),
                "BA CODE": row.get("BA CODE", ""),
                "Commitment Code": comm,
                "GL CODE": row.get("GL CODE", ""),
                "AMOUNT (Required Budget)": amt,
                "BUDGET SHEET VALUE (Available Budget)": "",
                "DIFF (Budget - Required)": "",
                "Diversion From Fund Center": "",
                "Diversion Amount": "",
                "Donor Balance After Donation": "",
                "DIFF AFTER DIVERSION": ""
            })
            continue

        diff = float(original_available) - float(amt)
        diversion_notes = []
        donor_balance_notes = []
        diversion_total = 0.0

        if diff < 0:
            shortage = abs(diff)

            def donate_from(donors):
                nonlocal shortage, diversion_total
                for donor_fc, donor_budget in donors:
                    if shortage <= 0:
                        break
                    take = min(donor_budget, shortage)
                    new_donor_balance = donor_budget - take
                    remaining[(donor_fc, comm)] = new_donor_balance

                    shortage -= take
                    diversion_total += take

                    diversion_notes.append(f"{donor_fc} ({take:,.2f})")
                    donor_balance_notes.append(f"{donor_fc} Bal ({new_donor_balance:,.2f})")

                    donor_ledger.append({
                        "Receiver Fund Center": receiver_fc,
                        "Donor Fund Center": donor_fc,
                        "Comm. Code": comm,
                        "Required Amount": amt,
                        "Receiver Original Budget": float(original_available),
                        "Diverted Amount": float(take),
                        "Donor Budget Before": float(donor_budget),
                        "Donor Budget After": float(new_donor_balance),
                    })

            donate_from(pick_donors(comm, receiver_fc, "F"))
            if shortage > 0:
                donate_from(pick_donors(comm, receiver_fc, "G"))

        diff_after = diff + diversion_total

        records.append({
            "Particular": row.get("Particular", ""),
            "BA CODE": row.get("BA CODE", ""),
            "Commitment Code": comm,
            "GL CODE": row.get("GL CODE", ""),
            "AMOUNT (Required Budget)": amt,
            "BUDGET SHEET VALUE (Available Budget)": float(original_available),
            "DIFF (Budget - Required)": float(diff),
            "Diversion From Fund Center": "; ".join(diversion_notes),
            "Diversion Amount": float(diversion_total) if diversion_total else 0.0,
            "Donor Balance After Donation": "; ".join(donor_balance_notes),
            "DIFF AFTER DIVERSION": float(diff_after)
        })

    salary_out = pd.DataFrame(records)
    ledger_out = pd.DataFrame(donor_ledger)

    if ledger_out.empty:
        ledger_out = pd.DataFrame(columns=[
            "Receiver Fund Center", "Donor Fund Center", "Comm. Code",
            "Required Amount", "Receiver Original Budget",
            "Diverted Amount", "Donor Budget Before", "Donor Budget After"
        ])

    return salary_out, ledger_out


def build_output_excel_bytes(budget_df: pd.DataFrame,
                             salary_df: pd.DataFrame | None = None,
                             ledger_df: pd.DataFrame | None = None) -> bytes:
    """
    Always writes Budget sheets.
    Writes Salary outputs only if salary_df and ledger_df provided.
    """
    from io import BytesIO
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        budget_df.to_excel(writer, sheet_name="BUDGET_ALL", index=False)

        for fc, grp in budget_df.groupby("Fund Center"):
            grp.to_excel(writer, sheet_name=str(fc)[:31], index=False)

        if salary_df is not None and ledger_df is not None:
            salary_df.to_excel(writer, sheet_name="SALARY_ANALYSIS", index=False)
            ledger_df.to_excel(writer, sheet_name="DONOR_LEDGER", index=False)

    output.seek(0)
    return output.read()
