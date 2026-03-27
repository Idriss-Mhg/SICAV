from docx import Document

from mapping import load_mapping
from parser import find_compartments, find_anchor, find_insert_idx, match_compartment
from inserter import insert_clause_after, find_body_ref, find_bullet_ref, reset_counter


def _collect_insertions(input_docx, input_excel, log):
    """
    Loads the document + mapping, resolves all anchor positions.
    Returns (doc, paragraphs, insertions, warnings) without touching the file.
    insertions : list of (insert_idx, anchor_idx, title, typ, content, body_ref, is_exact),
                 sorted by insert_idx descending so later positions are processed first.
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

            # ── Insertion index ───────────────────────────────────────────────
            if clause["exact_pos"]:
                # User specified an exact paragraph to insert after (overrides
                # automatic apres_titre / apres_section detection).
                insert_idx = find_anchor(
                    paragraphs, clause["exact_pos"], comp["start"], comp["end"])
                if insert_idx is None:
                    warnings.append(
                        f"  PositionExacte '{clause['exact_pos']}' not found "
                        f"in '{comp['name']}' — clause '{clause_id}' skipped"
                    )
                    continue
                log(f"  ✓ {clause_id} [{clause['type']}]: exact insert BEFORE para "
                    f"#{insert_idx} '{paragraphs[insert_idx].text[:60]}'")
            else:
                insert_idx = find_insert_idx(
                    paragraphs, anchor_idx, comp["end"], clause["position"])
                log(f"  ✓ {clause_id} [{clause['type']}]: insert after para #{insert_idx} "
                    f"'{paragraphs[insert_idx].text[:60]}'"
                    + (f"  [{clause['position']}]" if clause["position"] != "apres_titre" else ""))

            body_ref  = find_body_ref(paragraphs, comp["start"], comp["end"])
            insertions.append((
                insert_idx,                    # where to insert (last para of section)
                anchor_idx,                    # style reference (colored section title)
                clause["title"],
                clause["type"],
                clause["content"],
                body_ref,
                bool(clause["exact_pos"]),     # is_exact: bypass sectPr logic
            ))

    insertions.sort(key=lambda x: x[0], reverse=True)
    return doc, paragraphs, insertions, warnings


def _convert_to_pdf(docx_path, pdf_path, log):
    """
    Converts docx_path to PDF using docx2pdf.
    - Windows / macOS : requires Microsoft Word to be installed.
    - Linux           : requires LibreOffice to be installed.
    Logs a warning instead of crashing if conversion is unavailable.
    """
    try:
        from docx2pdf import convert
        log(f"Converting to PDF → {pdf_path}")
        convert(docx_path, pdf_path)
    except ImportError:
        log("⚠  docx2pdf not installed — pip install docx2pdf")
    except Exception as exc:
        log(f"⚠  PDF conversion failed: {exc}")
        log("   (Word required on Windows/macOS, LibreOffice on Linux)")


def run(input_docx, input_excel, output_docx, output_review=None,
        output_pdf=None, log=print):
    """
    Produces up to three output files:
      - output_docx   : clean insertion
      - output_review : same insertions as Word track changes (w:ins)
      - output_pdf    : PDF export of the normal output
    """
    # ── Normal output ─────────────────────────────────────────────────────────
    doc, paragraphs, insertions, warnings = _collect_insertions(
        input_docx, input_excel, log)

    bullet_ref = find_bullet_ref(doc)

    reset_counter()
    log(f"\nInserting {len(insertions)} clause(s) [normal]…")
    for insert_idx, anchor_idx, title, typ, content, body_ref, is_exact in insertions:
        insert_clause_after(
            paragraphs[insert_idx], title, typ, content,
            body_ref, bullet_ref, review=False,
            title_style_para=paragraphs[anchor_idx],
            exact=is_exact)

    log(f"Saving normal  → {output_docx}")
    doc.save(output_docx)

    # ── Review output ─────────────────────────────────────────────────────────
    if output_review:
        log("\nBuilding review version…")
        doc_r, paras_r, insertions_r, _ = _collect_insertions(
            input_docx, input_excel, lambda _: None)
        bullet_ref_r = find_bullet_ref(doc_r)

        reset_counter()
        for insert_idx, anchor_idx, title, typ, content, body_ref, is_exact in insertions_r:
            insert_clause_after(
                paras_r[insert_idx], title, typ, content,
                body_ref, bullet_ref_r, review=True,
                title_style_para=paras_r[anchor_idx],
                exact=is_exact)

        log(f"Saving review  → {output_review}")
        doc_r.save(output_review)

    # ── PDF output ────────────────────────────────────────────────────────────
    if output_pdf:
        log("\nGenerating PDF…")
        _convert_to_pdf(output_docx, output_pdf, log)

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
        output_pdf    = "data/prospectus_updated.pdf",
    )
