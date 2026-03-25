from docx import Document

from mapping import load_mapping
from parser import find_compartments, find_anchor, match_compartment
from inserter import insert_clause_after, find_body_ref, reset_counter


def _collect_insertions(input_docx, input_excel, log):
    """
    Loads the document + mapping, resolves all anchor positions.
    Returns (doc, paragraphs, insertions, warnings) without touching the file.
    insertions : list of (anchor_idx, clause_title), sorted descending.
    """
    log("Loading document…")
    doc = Document(input_docx)
    paragraphs = doc.paragraphs

    log("Loading mapping…")
    clauses, mapping = load_mapping(input_excel)
    log(f"  {len(clauses)} clauses defined, {len(mapping)} compartments in mapping")

    log("Parsing compartments…")
    compartments = find_compartments(doc)
    log(f"  {len(compartments)} compartments found in document")

    insertions = []
    warnings   = []

    for comp_name_excel, clause_ids in mapping.items():
        if not clause_ids:
            continue

        comp = match_compartment(comp_name_excel, compartments)
        if comp is None:
            warnings.append(f"Compartment not found in document: '{comp_name_excel}'")
            continue

        log(f"\n[{comp['name']}]  (paras {comp['start']}–{comp['end']})")

        for clause_id in clause_ids:
            if clause_id not in clauses:
                warnings.append(f"Clause ID '{clause_id}' not defined in 'clauses' sheet")
                continue

            clause     = clauses[clause_id]
            anchor     = clause["anchor"]
            anchor_idx = find_anchor(paragraphs, anchor, comp["start"], comp["end"])

            if anchor_idx is None:
                warnings.append(
                    f"  Anchor '{anchor}' not found in '{comp['name']}' "
                    f"— clause '{clause_id}' skipped"
                )
                continue

            body_ref = find_body_ref(paragraphs, comp["start"], comp["end"])
            log(f"  ✓ {clause_id}: insert after para #{anchor_idx} "
                f"'{paragraphs[anchor_idx].text[:60]}'")
            insertions.append((anchor_idx, clause["title"], body_ref))

    insertions.sort(key=lambda x: x[0], reverse=True)
    return doc, paragraphs, insertions, warnings


def run(input_docx, input_excel, output_docx, output_review=None, log=print):
    """
    Produces up to two output files from the same source document:
      - output_docx   : clean insertion (no markup)
      - output_review : same insertions as Word track changes (w:ins),
                        so reviewers can Accept / Reject in Word.
                        Omitted if output_review is None.
    """
    # ── Normal output ─────────────────────────────────────────────────────────
    doc, paragraphs, insertions, warnings = _collect_insertions(
        input_docx, input_excel, log)

    reset_counter()
    log(f"\nInserting {len(insertions)} clause(s) [normal]…")
    for anchor_idx, clause_title, body_ref in insertions:
        insert_clause_after(paragraphs[anchor_idx], clause_title, body_ref, review=False)

    log(f"Saving normal  → {output_docx}")
    doc.save(output_docx)

    # ── Review output ─────────────────────────────────────────────────────────
    if output_review:
        log(f"\nBuilding review version…")
        doc_r, paras_r, insertions_r, _ = _collect_insertions(
            input_docx, input_excel, lambda _: None)

        reset_counter()
        for anchor_idx, clause_title, body_ref in insertions_r:
            insert_clause_after(paras_r[anchor_idx], clause_title, body_ref, review=True)

        log(f"Saving review  → {output_review}")
        doc_r.save(output_review)

    if warnings:
        log("\n⚠  Warnings:")
        for w in warnings:
            log(f"   {w}")

    log(f"\nDone. ({len(insertions)} insertion(s), {len(warnings)} warning(s))")
    return {"insertions": len(insertions), "warnings": warnings}


if __name__ == "__main__":
    run(
        input_docx    = "data/prospectus.docx",
        input_excel   = "data/mapping.xlsx",
        output_docx   = "data/prospectus_updated.docx",
        output_review = "data/prospectus_review.docx",
    )
