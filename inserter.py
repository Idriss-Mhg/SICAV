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


# ── Reference paragraph selection ─────────────────────────────────────────────

def _is_bold(para):
    """True if the paragraph's first run has explicit bold formatting."""
    if not para.runs:
        return False
    run = para.runs[0]
    # Direct bold OR inherited bold from rPr
    return bool(run.bold) or (
        run._element.rPr is not None
        and run._element.rPr.find(qn("w:b")) is not None
    )


def find_body_ref(paragraphs, comp_start, comp_end):
    """
    Returns the first non-empty, non-bold Normal paragraph inside the
    compartment — used as formatting reference for inserted body text.
    Falls back to paragraphs[comp_start] if nothing qualifies.
    """
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


# ── XML helpers ───────────────────────────────────────────────────────────────

def _copy_pPr(ref_para):
    if ref_para._element.pPr is not None:
        return copy.deepcopy(ref_para._element.pPr)
    return None


def _copy_rPr(ref_para, bold=None):
    """
    Copy run properties from ref_para's first run.
    bold=True  → force bold on
    bold=False → force bold off (remove w:b / add w:b with w:val="0")
    bold=None  → keep whatever the reference has
    """
    if ref_para.runs and ref_para.runs[0]._element.rPr is not None:
        rPr = copy.deepcopy(ref_para.runs[0]._element.rPr)
    else:
        rPr = OxmlElement("w:rPr")

    # Remove any existing bold elements before applying our override
    for b in rPr.findall(qn("w:b")):
        rPr.remove(b)
    for b in rPr.findall(qn("w:bCs")):
        rPr.remove(b)

    if bold is True:
        rPr.insert(0, OxmlElement("w:b"))
    elif bold is False:
        # Explicitly cancel bold (needed when ref has bold inherited from style)
        b_off = OxmlElement("w:b")
        b_off.set(qn("w:val"), "0")
        rPr.insert(0, b_off)
    # bold=None → leave as-is (reference formatting preserved)

    return rPr


# ── Paragraph factories ───────────────────────────────────────────────────────

def _make_para(ref_para, text, bold=None):
    """Plain paragraph, formatting copied from ref_para."""
    new_p = OxmlElement("w:p")
    pPr = _copy_pPr(ref_para)
    if pPr is not None:
        new_p.append(pPr)
    if not text:
        return new_p
    new_r = OxmlElement("w:r")
    new_r.append(_copy_rPr(ref_para, bold))
    new_t = OxmlElement("w:t")
    new_t.text = text
    new_t.set(qn("xml:space"), "preserve")
    new_r.append(new_t)
    new_p.append(new_r)
    return new_p


def _make_para_review(ref_para, text, bold=None):
    """
    Same as _make_para but wrapped in Word track-changes (w:ins).
    Appears green + underlined; reviewers can Accept / Reject in Word.
    """
    new_p = OxmlElement("w:p")

    # Mark the paragraph mark itself as inserted (needed for correct track-change display)
    pPr = _copy_pPr(ref_para)
    if pPr is None:
        pPr = OxmlElement("w:pPr")
    rPr_in_pPr = pPr.find(qn("w:rPr"))
    if rPr_in_pPr is None:
        rPr_in_pPr = OxmlElement("w:rPr")
        pPr.append(rPr_in_pPr)
    ins_mark = OxmlElement("w:ins")
    ins_mark.set(qn("w:id"),     _next_id())
    ins_mark.set(qn("w:author"), REVIEW_AUTHOR)
    ins_mark.set(qn("w:date"),   REVIEW_DATE)
    rPr_in_pPr.insert(0, ins_mark)
    new_p.append(pPr)

    if not text:
        return new_p

    w_ins = OxmlElement("w:ins")
    w_ins.set(qn("w:id"),     _next_id())
    w_ins.set(qn("w:author"), REVIEW_AUTHOR)
    w_ins.set(qn("w:date"),   REVIEW_DATE)

    new_r = OxmlElement("w:r")
    new_r.append(_copy_rPr(ref_para, bold))
    new_t = OxmlElement("w:t")
    new_t.text = text
    new_t.set(qn("xml:space"), "preserve")
    new_r.append(new_t)
    w_ins.append(new_r)
    new_p.append(w_ins)

    return new_p


# ── Public API ────────────────────────────────────────────────────────────────

def insert_clause_after(anchor_para, clause_title, body_ref_para, review=False):
    """
    Inserts a clause block immediately after anchor_para.

    Formatting references:
      - blank lines   → anchor_para  (neutral)
      - clause title  → anchor_para  (section-title style)
      - body text     → body_ref_para (first normal body para of the compartment)

    review=True  → w:ins track-changes markup (green underline in Word)
    review=False → clean insertion
    """
    make = _make_para_review if review else _make_para

    # (ref_para, text, bold)
    blocks = [
        (anchor_para,   "",                                    None),   # trailing blank
        (body_ref_para, "[CLAUSE CONTENT TO BE COMPLETED]",   False),  # body placeholder
        (anchor_para,   clause_title,                         True),   # title
        (anchor_para,   "",                                    None),   # leading blank
    ]

    ref = anchor_para._element
    for ref_para, text, bold in blocks:
        ref.addnext(make(ref_para, text, bold=bold))
