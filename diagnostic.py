"""
diagnostic.py — affiche la structure de paragraphes autour des points d'insertion.

Usage:
    python diagnostic.py data/prospectus.docx data/mapping.xlsx
    python diagnostic.py data/prospectus.docx data/mapping.xlsx "defensive"
    python diagnostic.py data/prospectus.docx data/mapping.xlsx "global equity"
"""
import sys
from docx import Document
from docx.oxml.ns import qn
from mapping import load_mapping
from parser import find_compartments, find_anchor, find_insert_idx, _is_in_table


def dump_body_structure(doc, paragraphs, comp_start, comp_end):
    """
    Shows the raw sequence of body-level paragraphs and tables for the
    compartment range.  Uses element identity rather than index counting
    so large tables outside the compartment cannot throw off the display.
    """
    # Index all body-level paragraph elements in [comp_start, comp_end]
    comp_elems = {
        paragraphs[i]._element: i
        for i in range(comp_start, comp_end + 1)
        if not _is_in_table(paragraphs[i])
    }

    print("\n  --- Body-level structure (paragraphs + tables) ---")
    in_range = False
    for child in doc.element.body:
        if child.tag == qn("w:p"):
            if child in comp_elems:
                in_range = True
                idx = comp_elems[child]
                si  = _sect_info(paragraphs[idx])
                text = repr(paragraphs[idx].text.strip()[:60] or "(blank)")
                print(f"  P[{idx:5d}] {text}{si}")
            elif in_range:
                break
        elif child.tag == qn("w:tbl") and in_range:
            tc_paras   = list(child.iter(qn("w:p")))
            first_text = "".join(t.text or "" for t in tc_paras[0].iter(qn("w:t"))) if tc_paras else ""
            print(f"  {'':7}[TABLE  {len(tc_paras)} cell-paras, first: {repr(first_text[:40])}]")


def _sect_info(para):
    pPr = para._element.pPr
    if pPr is None:
        return ""
    sp = pPr.find(qn("w:sectPr"))
    if sp is None:
        return ""
    te = sp.find(qn("w:type"))
    val = te.get(qn("w:val")) if te is not None else "nextPage(default)"
    cols_elem = sp.find(qn("w:cols"))
    num = cols_elem.get(qn("w:num")) if cols_elem is not None else "1"
    return f"  [sectPr type={val} cols={num}]"


def _para_line(i, para, markers=None):
    text = repr(para.text.strip()[:70] or "(blank)")
    tbl  = " [TABLE_CELL]" if _is_in_table(para) else ""
    sect = _sect_info(para)
    tag  = f"  <<<{markers}>>>" if markers else ""
    return f"  {i:5d}: {text}{tbl}{sect}{tag}"


def dump_compartment(doc, paragraphs, clauses, mapping, comp):
    print(f"\n{'='*70}")
    print(f"  {comp['name']}  (paras {comp['start']}–{comp['end']})")
    print(f"{'='*70}")

    # Match compartment in mapping (case-insensitive substring)
    short = comp["short_name"].lower()
    comp_key = next(
        (k for k in mapping if k.lower() == short or
         short in k.lower() or k.lower() in short),
        None,
    )
    clause_ids = mapping.get(comp_key, []) if comp_key else []

    if not clause_ids:
        print("  (no clauses mapped to this compartment)")

    for clause_id in clause_ids:
        if clause_id not in clauses:
            print(f"\n  Clause {clause_id}: NOT DEFINED in clauses sheet")
            continue
        clause = clauses[clause_id]
        anchor_idx = find_anchor(
            paragraphs, clause["anchor"], comp["start"], comp["end"])
        if anchor_idx is None:
            print(f"\n  Clause {clause_id}: anchor '{clause['anchor']}' NOT FOUND")
            continue

        insert_idx = find_insert_idx(
            paragraphs, anchor_idx, comp["end"], clause["position"])

        print(f"\n  Clause {clause_id} [{clause['type']}]"
              f"  anchor='{clause['anchor']}'  position={clause['position']}")
        print(f"  anchor_idx={anchor_idx}  →  insert_idx={insert_idx}")
        print()

        # Print context: 5 paras before anchor, anchor → insert+5
        start = max(comp["start"], anchor_idx - 5)
        end   = min(comp["end"],   insert_idx + 5)
        for i in range(start, end + 1):
            markers = ""
            if i == anchor_idx:
                markers = "ANCHOR"
            elif i == insert_idx:
                markers = "INSERT"
            print(_para_line(i, paragraphs[i], markers or None))

    # Always list every sectPr in the compartment
    print(f"\n  --- All sectPr paragraphs in this compartment ---")
    found = False
    for i in range(comp["start"], comp["end"] + 1):
        si = _sect_info(paragraphs[i])
        if si:
            found = True
            # Print 3 lines of context around it
            for j in range(max(comp["start"], i - 2), min(comp["end"] + 1, i + 3)):
                m = "sectPr" if j == i else None
                print(_para_line(j, paragraphs[j], m))
            print()
    if not found:
        print("  (none found)")

    # Show body-level structure (paragraphs + tables)
    dump_body_structure(doc, paragraphs, comp["start"], comp["end"])


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    docx_path  = sys.argv[1]
    xlsx_path  = sys.argv[2]
    comp_filter = sys.argv[3].lower() if len(sys.argv) > 3 else None

    doc = Document(docx_path)
    paragraphs = doc.paragraphs
    clauses, mapping = load_mapping(xlsx_path)
    compartments = find_compartments(doc)

    if comp_filter is None:
        print(f"Found {len(compartments)} compartments:\n")
        for c in compartments:
            short = c["short_name"]
            print(f"  {c['start']:5d}–{c['end']:5d}  {c['name']}")
        print("\nRe-run with a compartment name fragment to inspect it.")
        return

    matches = [c for c in compartments if comp_filter in c["name"].lower()]
    if not matches:
        print(f"No compartment matching '{comp_filter}'")
        return

    for comp in matches:
        dump_compartment(doc, paragraphs, clauses, mapping, comp)


if __name__ == "__main__":
    main()
