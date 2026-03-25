"""
Run once to generate data/mapping.xlsx with sample data.
Adapt the compartment names and clauses to match your real document.
"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

OUTPUT = "data/mapping.xlsx"


def main():
    wb = openpyxl.Workbook()

    # ------------------------------------------------------------------ #
    # Sheet 1 — "clauses"                                                 #
    # Each row defines a clause: its ID, its title, and the anchor text   #
    # (the exact paragraph after which the clause will be inserted).      #
    # ------------------------------------------------------------------ #
    ws1 = wb.active
    ws1.title = "clauses"

    headers = ["ClauseID", "ClauseTitre", "InsererApres"]
    ws1.append(headers)

    sample_clauses = [
        ["CL01", "Sustainable Finance Disclosure Clause", "Risk Management:"],
        ["CL02", "Stewardship Policy Clause",             "Conflicts of interest"],
        ["CL03", "German Investment Tax Act Clause",      "Reference Currency:"],
    ]
    for row in sample_clauses:
        ws1.append(row)

    # Style header row
    for cell in ws1[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
    ws1.column_dimensions["A"].width = 10
    ws1.column_dimensions["B"].width = 45
    ws1.column_dimensions["C"].width = 55

    # ------------------------------------------------------------------ #
    # Sheet 2 — "mapping"                                                 #
    # Rows = compartments, columns = clause IDs.                          #
    # Put an "X" in a cell to insert that clause in that compartment.     #
    # ------------------------------------------------------------------ #
    ws2 = wb.create_sheet("mapping")

    clause_ids = [r[0] for r in sample_clauses]
    ws2.append(["Compartiment"] + clause_ids)

    sample_mapping = [
        ["CPR Invest \u2013 Silver Age",  "X", "X", ""],
        ["CPR Invest \u2013 Reactive",    "X", "X", "X"],
    ]
    for row in sample_mapping:
        ws2.append(row)

    # Style header row
    for cell in ws2[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
    ws2.column_dimensions["A"].width = 45
    for col in ["B", "C", "D"]:
        ws2.column_dimensions[col].width = 8

    # Center the X cells
    for row in ws2.iter_rows(min_row=2):
        for cell in row[1:]:
            cell.alignment = Alignment(horizontal="center")

    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    wb.save(OUTPUT)
    print(f"Created: {OUTPUT}")


if __name__ == "__main__":
    main()
