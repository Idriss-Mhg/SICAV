import copy
from datetime import datetime, timezone

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ── Track-changes metadata ────────────────────────────────────────────────────
REVIEW_AUTHOR = "Clause Inserter"
REVIEW_DATE   = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
_ins_counter  = 0


def _next_id():
    global _ins_counter
    _ins_counter += 1
    return str(_ins_counter)


def reset_counter():
    global _ins_counter
    _ins_counter = 0


# ── Reference paragraph finders ───────────────────────────────────────────────

def _is_bold(para):
    if not para.runs:
        return False
    run = para.runs[0]
    return bool(run.bold) or (
        run._element.rPr is not None
        and run._element.rPr.find(qn("w:b")) is not None
    )


def find_body_ref(paragraphs, comp_start, comp_end):
    """First non-empty, non-bold Normal paragraph in the compartment."""
    for i in range(comp_start + 1, comp_end + 1):
        p = paragraphs[i]
        if not p.text.strip():
            continue
        if p.style.name not in ("Normal", "Body Text", "Default Paragraph Font"):
            continue
        if _is_bold(p):
            continue
        return p
    return paragraphs[comp_start]


def find_bullet_ref(doc):
    """First non-empty List Paragraph in the whole document (bullet reference)."""
    for para in doc.paragraphs:
        if para.style.name == "List Paragraph" and para.text.strip():
            return para
    return None


# ── Low-level XML helpers ─────────────────────────────────────────────────────

def _is_multicolumn_sectPr(pPr_elem):
    """True if the sectPr ending at this paragraph defines a multi-column section."""
    sect_pr = pPr_elem.find(qn("w:sectPr"))
    if sect_pr is None:
        return False
    cols = sect_pr.find(qn("w:cols"))
    if cols is None:
        return False
    try:
        return int(cols.get(qn("w:num"), "1")) > 1
    except (ValueError, TypeError):
        return False


def _body_level_elem(para):
    """
    Returns the direct-child-of-body XML element that contains para.
    If para is at body level, returns para._element.
    If para is inside a table cell, walks up and returns the w:tbl ancestor
    that is a direct child of w:body — so insertion happens after the whole
    table rather than inside a cell.
    """
    elem = para._element
    while True:
        parent = elem.getparent()
        if parent is None or parent.tag == qn("w:body"):
            return elem
        elem = parent


def _copy_pPr(ref_para):
    if ref_para._element.pPr is not None:
        return copy.deepcopy(ref_para._element.pPr)
    return None


def _copy_rPr(ref_para, bold=None):
    """
    Copy run properties from ref_para's first run.
    bold=True  → force bold on
    bold=False → force bold off
    bold=None  → keep reference as-is
    """
    if ref_para.runs and ref_para.runs[0]._element.rPr is not None:
        rPr = copy.deepcopy(ref_para.runs[0]._element.rPr)
    else:
        rPr = OxmlElement("w:rPr")

    for b in rPr.findall(qn("w:b")):
        rPr.remove(b)
    for b in rPr.findall(qn("w:bCs")):
        rPr.remove(b)

    if bold is True:
        rPr.insert(0, OxmlElement("w:b"))
    elif bold is False:
        b_off = OxmlElement("w:b")
        b_off.set(qn("w:val"), "0")
        rPr.insert(0, b_off)

    return rPr


def _make_run(ref_para, text, bold=None):
    r = OxmlElement("w:r")
    r.append(_copy_rPr(ref_para, bold))
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    r.append(t)
    return r


def _mark_pPr_ins(pPr):
    """Add w:ins marker inside pPr/rPr (marks the paragraph mark as inserted)."""
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    ins = OxmlElement("w:ins")
    ins.set(qn("w:id"),     _next_id())
    ins.set(qn("w:author"), REVIEW_AUTHOR)
    ins.set(qn("w:date"),   REVIEW_DATE)
    rPr.insert(0, ins)


def _wrap_ins(run_elem):
    """Wrap a w:r element inside a w:ins element."""
    w_ins = OxmlElement("w:ins")
    w_ins.set(qn("w:id"),     _next_id())
    w_ins.set(qn("w:author"), REVIEW_AUTHOR)
    w_ins.set(qn("w:date"),   REVIEW_DATE)
    w_ins.append(run_elem)
    return w_ins


def _ensure_pPr(p_elem):
    """Get or create w:pPr for a paragraph element."""
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_elem.insert(0, pPr)
    return pPr


def _add_keep_together(elements):
    """
    Add w:keepNext to every paragraph except the last so Word keeps the
    whole clause block in the same column (no mid-clause column break).
    w:keepNext prevents both page and column breaks between paragraphs.
    """
    for elem in elements[:-1]:
        pPr = _ensure_pPr(elem)
        if pPr.find(qn("w:keepNext")) is None:
            pPr.insert(0, OxmlElement("w:keepNext"))


# ── Paragraph factories ───────────────────────────────────────────────────────

def _make_para(ref_para, text, bold=None):
    """Plain paragraph styled like ref_para."""
    new_p = OxmlElement("w:p")
    pPr = _copy_pPr(ref_para)
    if pPr is not None:
        new_p.append(pPr)
    if text:
        new_p.append(_make_run(ref_para, text, bold))
    return new_p


def _make_para_review(ref_para, text, bold=None):
    """Same as _make_para but with w:ins track-changes markup."""
    new_p = OxmlElement("w:p")
    pPr = _copy_pPr(ref_para)
    if pPr is None:
        pPr = OxmlElement("w:pPr")
    _mark_pPr_ins(pPr)
    new_p.append(pPr)
    if text:
        new_p.append(_wrap_ins(_make_run(ref_para, text, bold)))
    return new_p


def _make_bullet(ref_bullet, ref_body, text, review=False):
    """
    Bullet paragraph: pPr (incl. numPr) from ref_bullet, rPr from ref_body.
    ref_bullet must be a List Paragraph paragraph to carry correct bullet style.
    """
    new_p = OxmlElement("w:p")
    pPr = _copy_pPr(ref_bullet)
    if pPr is None:
        pPr = OxmlElement("w:pPr")
    if review:
        _mark_pPr_ins(pPr)
    new_p.append(pPr)
    if text:
        run = _make_run(ref_body, text, bold=False)
        new_p.append(_wrap_ins(run) if review else run)
    return new_p


# ── Public API ────────────────────────────────────────────────────────────────

def insert_clause_after(anchor_para, clause_title, clause_type, content_items,
                         body_ref_para, bullet_ref_para=None, review=False,
                         title_style_para=None, exact=False):
    """
    Inserts a clause block immediately after anchor_para.

    clause_type     : 'texte' | 'liste' | 'sous_titres'
    content_items   : list of dicts —
        texte       : [{"texte": "..."}]
        liste       : [{"texte": "bullet 1"}, {"texte": "bullet 2"}, ...]
        sous_titres : [{"texte": "Sub A", "sous_texte": "Body A"}, ...]

    Formatting references:
        title / blank        ← title_style_para (colored section title anchor)
                               falls back to anchor_para if not provided
        body text            ← body_ref_para (first normal body para)
        bullet pPr           ← bullet_ref_para (List Paragraph)

    review=True  → w:ins track-changes markup
    """
    make        = _make_para_review if review else _make_para
    bullet_ref  = bullet_ref_para or body_ref_para
    title_ref   = title_style_para if title_style_para is not None else anchor_para

    # Build paragraphs in final display order
    elements = []
    elements.append(make(title_ref, "", None))           # leading blank
    elements.append(make(title_ref, clause_title, True)) # clause title (bold, colored)

    if clause_type == "texte":
        text = (
            content_items[0].get("texte") if content_items and content_items[0].get("texte")
            else "[TEXTE À COMPLÉTER]"
        )
        elements.append(make(body_ref_para, text, False))

    elif clause_type == "liste":
        items = content_items or [{"texte": "[PUCE À COMPLÉTER]"}]
        for item in items:
            t = item.get("texte") or "[PUCE À COMPLÉTER]"
            elements.append(_make_bullet(bullet_ref, body_ref_para, t, review=review))

    elif clause_type == "sous_titres":
        items = content_items or [{"texte": "[SOUS-TITRE]", "sous_texte": "[TEXTE À COMPLÉTER]"}]
        for item in items:
            st = (item.get("texte") or "[SOUS-TITRE]").rstrip()
            if not st.endswith(":"):
                st += ":"
            body_text = item.get("sous_texte") or "[TEXTE À COMPLÉTER]"
            # Subtitle: body_ref formatting + bold (same size/colour as body, just bold)
            elements.append(make(body_ref_para, st, True))
            elements.append(make(body_ref_para, body_text, False))

    elements.append(make(title_ref, "", None))             # trailing blank

    # Always work at document-body level (avoids inserting inside table cells).
    ref_elem = _body_level_elem(anchor_para)

    pPr_elem   = ref_elem.find(qn("w:pPr")) if ref_elem.tag == qn("w:p") else None
    has_sectPr = pPr_elem is not None and pPr_elem.find(qn("w:sectPr")) is not None
    is_2col    = has_sectPr and _is_multicolumn_sectPr(pPr_elem)

    if exact:
        # PositionExacte: insert BEFORE the named paragraph.
        for elem in elements:
            ref_elem.addprevious(elem)

    elif is_2col:
        if not anchor_para.text.strip():
            # Blank paragraph carries the sectPr.
            # Walk backwards from the sectPr paragraph to find the last non-blank
            # body-level element (paragraph or table), then insert after it.
            # This avoids placing the clause after pre-existing trailing blanks.
            # Exclude the leading blank from keepNext so Word can use it to
            # balance the two columns (avoids the large empty-space problem).
            _add_keep_together(elements[1:])

            pivot = ref_elem.getprevious()
            while pivot is not None:
                if pivot.tag == qn("w:p"):
                    if "".join(t.text or "" for t in pivot.iter(qn("w:t"))).strip():
                        break
                # Skip tables and blank paragraphs; keep walking back.
                pivot = pivot.getprevious()

            if pivot is not None:
                for elem in reversed(elements):
                    pivot.addnext(elem)
            else:
                # Fallback: no non-blank predecessor found, insert before sectPr.
                for elem in elements:
                    ref_elem.addprevious(elem)

        else:
            # Content paragraph carries the sectPr (e.g. "T3 USD — Acc §").
            # Move the sectPr to the trailing blank so the clause stays in the
            # 2-column section and "Main Share Classes" is in 1-column after it.
            _add_keep_together(elements[1:])
            sect_pr = pPr_elem.find(qn("w:sectPr"))
            pPr_elem.remove(sect_pr)
            trailing_pPr = elements[-1].find(qn("w:pPr"))
            if trailing_pPr is None:
                trailing_pPr = OxmlElement("w:pPr")
                elements[-1].insert(0, trailing_pPr)
            trailing_pPr.append(sect_pr)
            for elem in reversed(elements):
                ref_elem.addnext(elem)

    else:
        # Normal case (apres_titre, or apres_section without a 2-col sectPr).
        for elem in reversed(elements):
            ref_elem.addnext(elem)
