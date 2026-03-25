import openpyxl


def load_mapping(excel_path):
    """
    Reads the Excel file and returns:
        clauses : dict { clause_id -> {title, anchor} }
        mapping : dict { compartment_name -> [clause_id, ...] }
    """
    wb = openpyxl.load_workbook(excel_path)

    # --- Sheet "clauses" ---
    ws_clauses = wb["clauses"]
    clauses = {}
    for row in ws_clauses.iter_rows(min_row=2, values_only=True):
        clause_id, title, anchor = row[0], row[1], row[2]
        if clause_id:
            clauses[str(clause_id).strip()] = {
                "title": str(title).strip() if title else "",
                "anchor": str(anchor).strip() if anchor else "",
            }

    # --- Sheet "mapping" (matrix: rows = compartments, cols = clause IDs) ---
    ws_mapping = wb["mapping"]
    header = [cell.value for cell in ws_mapping[1]]
    clause_ids = [str(v).strip() for v in header[1:] if v]

    mapping = {}
    for row in ws_mapping.iter_rows(min_row=2, values_only=True):
        comp_name = row[0]
        if not comp_name:
            continue
        active_clauses = []
        for j, val in enumerate(row[1: len(clause_ids) + 1]):
            if val and str(val).strip().upper() in ("X", "YES", "OUI", "1", "TRUE"):
                active_clauses.append(clause_ids[j])
        mapping[str(comp_name).strip()] = active_clauses

    return clauses, mapping
