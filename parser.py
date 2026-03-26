import re

SUPPLEMENT_RE = re.compile(r"^SUPPLEMENT\s+\d+\.", re.IGNORECASE)


def _extract_short_name(full_text):
    """
    'SUPPLEMENT 3. CPR Invest – Global' → 'CPR Invest – Global'
    Returns the part after 'SUPPLEMENT N. ', stripped.
    """
    return SUPPLEMENT_RE.sub("", full_text).strip()


def find_compartments(doc):
    """
    Scans doc.paragraphs for lines matching 'SUPPLEMENT N. ...'
    Returns a list of dicts: {name, short_name, start, end}.
      name       : full text  ('SUPPLEMENT 1. CPR Invest – Silver Age')
      short_name : name only  ('CPR Invest – Silver Age')
    """
    paragraphs = doc.paragraphs
    compartments = []

    for i, para in enumerate(paragraphs):
        text = para.text.strip()
        if SUPPLEMENT_RE.match(text):
            compartments.append({
                "name":       text,
                "short_name": _extract_short_name(text),
                "start":      i,
                "end":        None,
            })

    for j in range(len(compartments) - 1):
        compartments[j]["end"] = compartments[j + 1]["start"] - 1
    if compartments:
        compartments[-1]["end"] = len(paragraphs) - 1

    return compartments


def _is_title_like(para, anchor_para):
    """
    Heuristic: a paragraph looks like a section title if it shares
    the same bold formatting as the anchor paragraph.
    Also catches short lines ending with ':'.
    """
    text = para.text.strip()
    if not text:
        return False
    if text.endswith(":") and len(text) < 80:
        return True
    # Same bold state as anchor
    anchor_bold = anchor_para.runs[0].bold if anchor_para.runs else False
    para_bold   = para.runs[0].bold        if para.runs        else False
    return bool(para_bold) == bool(anchor_bold) and bool(para_bold)


def find_insert_idx(paragraphs, anchor_idx, comp_end, position):
    """
    Returns the paragraph index AFTER WHICH the clause should be inserted.

    position='apres_titre'   → anchor_idx  (right after the anchor title)
    position='apres_section' → last paragraph of the anchor's section,
                               i.e. the paragraph just before the next
                               title-like paragraph within the compartment.
    """
    if position != "apres_section":
        return anchor_idx

    anchor_para = paragraphs[anchor_idx]
    last_content = anchor_idx

    for i in range(anchor_idx + 1, comp_end + 1):
        para = paragraphs[i]
        if not para.text.strip():
            continue
        if _is_title_like(para, anchor_para):
            break
        last_content = i

    return last_content


def find_anchor(paragraphs, anchor_text, start, end):
    """
    Searches paragraphs[start..end] for the one that best matches anchor_text.
    Strategy (in order of preference):
      1. Exact match (case-insensitive, stripped)
      2. Paragraph starts with anchor text
      3. Anchor text is a substring of paragraph text
    Returns the paragraph index, or None if not found.
    """
    anchor_norm = anchor_text.lower().strip()

    for i in range(start, end + 1):
        para_norm = paragraphs[i].text.lower().strip()
        if para_norm == anchor_norm:
            return i

    for i in range(start, end + 1):
        para_norm = paragraphs[i].text.lower().strip()
        if para_norm.startswith(anchor_norm):
            return i

    for i in range(start, end + 1):
        para_norm = paragraphs[i].text.lower().strip()
        if anchor_norm in para_norm:
            return i

    return None


def match_compartment(comp_name_excel, compartments):
    """
    Matches the short name from the Excel ('CPR Invest – Silver Age')
    against compartments' short_name extracted from the document.
    Strategy: exact → startswith → substring (all case-insensitive).
    """
    name_norm = comp_name_excel.lower().strip()

    for comp in compartments:
        if comp["short_name"].lower() == name_norm:
            return comp

    for comp in compartments:
        short = comp["short_name"].lower()
        if short.startswith(name_norm) or name_norm.startswith(short):
            return comp

    for comp in compartments:
        short = comp["short_name"].lower()
        if name_norm in short or short in name_norm:
            return comp

    return None
