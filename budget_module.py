import pandas as pd
import re

re_fc_header_2star = re.compile(r"^\s*\*\*\s*([FG]\d{4})")
re_fc_star_line = re.compile(r"^\s*\*\s*([FG]\d{4})")
re_item = re.compile(r"^\s*([A-Z]\d{5})\s+(.+)$")


def _to_number(x):
    try:
        return float(str(x).replace(",", ""))
    except:
        return 0.0


def convert_budget_to_df(budget_file, sheet_name=None):
    xls = pd.ExcelFile(budget_file)
    sheets = [sheet_name] if sheet_name else xls.sheet_names

    rows = []

    for sheet in sheets:
        df = pd.read_excel(budget_file, sheet_name=sheet, header=None)

        available_col = df.shape[1] - 1

        # ✅ FIXED PART
        for r in range(min(30, len(df))):
            row_vals = df.iloc[r].tolist()
            for c, v in enumerate(row_vals):
                try:
                    v_clean = str(v).strip().lower()
                except:
                    v_clean = ""

                if v_clean in ["available budge", "available budget"]:
                    available_col = c
                    break

        current_fc = None
        inside = False

        for i in range(len(df)):
            cell = str(df.iat[i, 1]).strip()

            if cell.startswith("**"):
                m = re_fc_header_2star.match(cell)
                if m:
                    current_fc = m.group(1)
                    inside = False
                continue

            if cell.startswith("*") and not cell.startswith("**"):
                m = re_fc_star_line.match(cell)
                if m:
                    current_fc = m.group(1)
                    inside = True
                continue

            if inside and current_fc:
                m = re_item.match(cell)
                if m:
                    comm = m.group(1)
                    text = m.group(2)
                    val = _to_number(df.iat[i, available_col])

                    rows.append({
                        "Fund Center": current_fc,
                        "Comm. Code": comm,
                        "TEXT": text,
                        "Budget Available": val
                    })

    return pd.DataFrame(rows)


def salary_analysis_with_ledger(salary_file, budget_df, salary_sheet):
    sal = pd.read_excel(salary_file, sheet_name=salary_sheet)

    budget_map = {
        (r["Fund Center"], r["Comm. Code"]): r["Budget Available"]
        for _, r in budget_df.iterrows()
    }

    results = []

    for _, row in sal.iterrows():
        fc = "F" + str(row.get("BA CODE", ""))
        comm = str(row.get("Commitment Code", ""))
        amt = _to_number(row.get("AMOUNT", 0))

        budget = budget_map.get((fc, comm), 0)

        results.append({
            "BA": fc,
            "Comm": comm,
            "Required": amt,
            "Available": budget,
            "Diff": budget - amt
        })

    return pd.DataFrame(results), pd.DataFrame()


def build_output_excel_bytes(budget_df, salary_df=None, ledger_df=None):
    from io import BytesIO

    output = BytesIO()

    with pd.ExcelWriter(output) as writer:
        budget_df.to_excel(writer, sheet_name="BUDGET_ALL", index=False)

        for fc, grp in budget_df.groupby("Fund Center"):
            grp.to_excel(writer, sheet_name=str(fc)[:31], index=False)

        if salary_df is not None:
            salary_df.to_excel(writer, sheet_name="SALARY", index=False)

    output.seek(0)
    return output.read()
