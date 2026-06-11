# Source Trace:
# File: formatter.py
# Knowledge Files: CodeSourceDB v3.4, SyntaxBiasDB v2.3, HumanSyntaxDB v1.2, mainProtocol v5.5
# REF_IDs: RM_DO178_001, RM_ISO_001, RM_HCI_002, WEB_PY_001, RM_HCI_001
# Logic: Style application and document structure formatting using explicit typing,
# lxml deep copying, XML relationship sanitization, and dynamic placeholder injection.

"""
Rhema -- Auto Formatter
core/formatter.py

Single responsibility: Applies the selected template's declared Word styles
to the confirmed document structure. Maps heading levels, body text, and
other styles from the source document to the template's style names.

TEMPLATE STRUCTURE (4 sections via inline sectPr):
  Section 1 (cover)      : paragraphs [00-07], inline sectPr at [08]
  Section 2 (TOC)        : CONTENTS heading + TOC field, inline sectPr at [12]
  Section 3 (body)       : [[BODY_CONTENT_START]] at [13], inline sectPr at [14]
  Section 4 (references) : REFERENCES heading + [[REFERENCES_START]], final sectPr

Body content is inserted AT the [[BODY_CONTENT_START]] paragraph (replacing it).
Reference content is inserted AT the [[REFERENCES_START]] paragraph (replacing it).
All other template structure is preserved exactly.

Governing standards:
  DO-178C source traceability    -- REF_ID: RM_DO178_001
  ISO 31000 risk treatment       -- REF_ID: RM_ISO_001
  ISO 9241-110:2020 transparency -- REF_ID: RM_HCI_002
"""

import copy
import re
from pathlib import Path
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

from docx import Document as DocxDocument
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from core.detector import ImplicitElement
from config.constants import TEMPLATE_EXTENSION


# ---------------------------------------------------------------------------
# Result data structure
# ---------------------------------------------------------------------------

@dataclass
class FormatterResult:
    success:         bool = False
    formatted_doc:   Optional[object] = None
    error:           str = ""
    unmapped_styles: List[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Style constants
# ---------------------------------------------------------------------------

STYLE_NORMAL  = "Normal"
STYLE_H1      = "Heading 1"
STYLE_H2      = "Heading 2"
STYLE_H3      = "Heading 3"
STYLE_CAPTION = "Caption"

_HEADING_LABEL_TO_STYLE: Dict[str, str] = {
    "heading 1":  STYLE_H1,
    "heading 2":  STYLE_H2,
    "heading 3":  STYLE_H3,
    "subheading": STYLE_H2,   # subheading maps to H2, not H3
    "title":      STYLE_H1,
    "subtitle":   STYLE_H2,
    "list item":  STYLE_NORMAL,
}

_WORD_HEADING_STYLES = {
    "heading 1", "heading 2", "heading 3",
    "heading 4", "heading 5", "heading 6",
    "title", "subtitle",
}

# Known section names whose heading ends immediately after the first word
_KNOWN_SECTION_NAMES = {
    "references", "bibliography", "works cited", "appendix",
    "appendices", "abstract", "conclusion", "introduction",
    "methodology", "results", "discussion",
}

# Words that signal body text has started
_BODY_STARTERS = {
    "its", "their", "this", "these", "those",
    "through", "via", "by", "as", "since", "because",
    "university", "florida", "college",
}

_TITLE_CONNECTORS = {
    "and", "of", "for", "the", "a", "an", "in", "to",
    "with", "vs", "vs.", "or", "&",
}

_MAX_TITLE_WORDS = 6

# Reference entry patterns -- lines that are citations, not headings
_REF_ENTRY_RE = re.compile(
    r'^[A-Z][^.]+\.\s+\(\d{4}\)\.'  # Author. (Year).
    r'|^https?://'                   # URL
    r'|^\[\d+\]'                     # [1] numbered ref
    r'|^\d+\.\s+[A-Z][^.]+\.\s+\(\d{4}\)',  # 1. Author. (Year).
)


# ---------------------------------------------------------------------------
# Heading/body split helper (deterministic, no AI)
# REF_ID: RM_DO178_001
# ---------------------------------------------------------------------------

def _split_heading_from_body(text: str):
    """
    Deterministically split a numbered heading from concatenated body text.
    Returns (heading_text, body_text) or (text, None) if no split needed.
    """
    from core.detector import _NUMBERED_HEADING_RE

    text = text.strip()
    if not _NUMBERED_HEADING_RE.match(text):
        return text, None

    prefix_match = re.match(r'^(\d+\.)+\s*', text)
    if not prefix_match:
        return text, None

    after_prefix = text[prefix_match.end():]
    words = after_prefix.split()
    num_prefix = prefix_match.group(0).rstrip()

    if not words:
        return text, None

    # Known section names: heading is just the section word, rest is body
    first_clean = words[0].rstrip(".,;:()").lower()
    if first_clean in _KNOWN_SECTION_NAMES and len(words) > 1:
        heading = (num_prefix + " " + words[0]).strip()
        body    = " ".join(words[1:]).strip()
        return heading, body if body else None

    # Short title -- no split
    if len(words) <= _MAX_TITLE_WORDS:
        return text, None

    title_words = []
    body_start_idx = None

    for i, word in enumerate(words):
        clean       = word.rstrip(".,;:()")
        clean_lower = clean.lower()

        if i == 0:
            title_words.append(word)
            continue

        if clean_lower in _BODY_STARTERS and len(title_words) >= 1:
            body_start_idx = i; break
        if "-" in clean:
            parts = clean.split("-")
            if len(parts) >= 2 and parts[1] and parts[1][0].islower():
                body_start_idx = i; break
        if clean and clean[0].islower() and clean_lower not in _TITLE_CONNECTORS:
            body_start_idx = i; break
        if word.startswith("(") and len(title_words) >= 1:
            body_start_idx = i; break
        if len(title_words) >= _MAX_TITLE_WORDS:
            body_start_idx = i; break
        if (len(title_words) >= 2 and title_words[-1].endswith(".") and
                clean and clean[0].isupper()):
            body_start_idx = i; break
        title_words.append(word)

    if body_start_idx is None or not title_words:
        return text, None

    heading = (num_prefix + " " + " ".join(title_words)).strip()
    body    = " ".join(words[body_start_idx:]).strip()
    return heading, body if body else None


# ---------------------------------------------------------------------------
# Style helpers
# ---------------------------------------------------------------------------

def _get_template_style_names(template_doc: Any) -> set:
    return {style.name for style in template_doc.styles}


def _build_style_id_map(template_doc: Any) -> Dict[str, str]:
    """Map style display name -> XML style ID (e.g. 'Heading 1' -> 'Heading1')."""
    return {
        style.name: style.element.get(qn("w:styleId"))
        for style in template_doc.styles
        if style.element.get(qn("w:styleId"))
    }


def _set_paragraph_style(para_xml, style_id: str) -> None:
    """Set w:pStyle on a paragraph XML element using the XML style ID."""
    pPr = para_xml.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        para_xml.insert(0, pPr)
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is None:
        pStyle = OxmlElement("w:pStyle")
        pPr.insert(0, pStyle)
    pStyle.set(qn("w:val"), style_id)


# ---------------------------------------------------------------------------
# Paragraph builder helpers
# ---------------------------------------------------------------------------

def _strip_run_formatting(para_xml) -> None:
    """
    Remove explicit run-level and paragraph-level formatting overrides
    from a copied paragraph so the template style governs appearance.

    Paragraph-level: removes w:spacing override so Normal style spacing applies.
    Run-level: removes color, font, size overrides. Preserves bold/italic/underline.
    Does NOT touch runs containing fldChar or instrText elements.
    REF_ID: RM_DO178_001
    """
    # Strip paragraph-level spacing override (source documents carry their own
    # spacing values which override the template's Normal style declaration).
    pPr = para_xml.find(qn("w:pPr"))
    if pPr is not None:
        for spacing in pPr.findall(qn("w:spacing")):
            pPr.remove(spacing)

    # Strip run-level formatting overrides
    for r in para_xml.findall(".//" + qn("w:r")):
        if (r.find(qn("w:fldChar")) is not None or
                r.find(qn("w:instrText")) is not None):
            continue
        rPr = r.find(qn("w:rPr"))
        if rPr is None:
            continue
        for color in rPr.findall(qn("w:color")):
            rPr.remove(color)
        for fonts in rPr.findall(qn("w:rFonts")):
            rPr.remove(fonts)
        for sz in rPr.findall(qn("w:sz")):
            rPr.remove(sz)
        for szCs in rPr.findall(qn("w:szCs")):
            rPr.remove(szCs)


def _find_source_body_start(
    body_children: list,
    implicit_map: dict,
    para_by_elem: dict,
) -> int:
    """
    Find the index in body_children where actual body content starts.
    Skips TOC fields and leading cover page content (placeholders, short
    labels, empty paragraphs) so the template cover/TOC are not duplicated.
    REF_ID: RM_DO178_001
    """
    import re as _re
    para_idx = 0

    for i, child in enumerate(body_children):
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag != "p":
            if tag == "tbl":
                para_idx += 1
            continue

        text = "".join(t.text or "" for t in child.iter() if t.tag.endswith("}t"))

        # Skip TOC field paragraphs
        instr = [t.text or "" for t in child.iter() if t.tag.endswith("}instrText")]
        if any("TOC" in t for t in instr):
            para_idx += 1
            continue

        # First detected heading = body starts here
        if para_idx in implicit_map:
            return i

        # Substantial non-placeholder, non-label text = body starts here
        clean = text.strip()
        is_placeholder = bool(_re.search(r'\[\[[\w_]+\]\]', clean))
        is_short_label = len(clean.split()) <= 6 and not any(c in clean for c in ".,:;")
        if clean and not is_placeholder and not is_short_label:
            return i

        para_idx += 1

    return 0


def _sample_template_styles(template_doc: Any) -> tuple:
    """
    Sample table and list formatting from the template dummy section,
    then strip the dummy section so it never appears in output documents.

    Returns (ref_tbl_xml, ref_list_numPr) where:
      ref_tbl_xml    -- w:tbl element of the reference table, or None
      ref_list_numPr -- w:numPr element of the reference list item, or None

    Dummy section is all body children between [[BODY_CONTENT_START]]
    and the body section inline sectPr. REF_ID: RM_DO178_001
    """
    body     = template_doc.element.body
    children = list(body)

    body_start_idx = None
    body_sect_idx  = None

    for i, child in enumerate(children):
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag != "p":
            continue
        text = "".join(t.text or "" for t in child.iter() if t.tag.endswith("}t"))
        if "[[BODY_CONTENT_START]]" in text:
            body_start_idx = i
            continue
        if body_start_idx is not None:
            pPr = child.find(qn("w:pPr"))
            if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
                body_sect_idx = i
                break

    if body_start_idx is None or body_sect_idx is None:
        return None, None

    dummy_children = children[body_start_idx + 1 : body_sect_idx]
    ref_tbl_xml  = None
    ref_list_numPr = None

    for child in dummy_children:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "tbl":
            tblPr = child.find(qn("w:tblPr"))
            if tblPr is not None:
                cap = tblPr.find(qn("w:tblCaption"))
                if cap is not None and "RHEMA_TABLE_STYLE" in (cap.get(qn("w:val")) or ""):
                    ref_tbl_xml = copy.deepcopy(child)

        if tag == "p" and ref_list_numPr is None:
            text = "".join(t.text or "" for t in child.iter() if t.tag.endswith("}t"))
            if "RHEMA_LIST_STYLE" in text:
                pPr = child.find(qn("w:pPr"))
                if pPr is not None:
                    numPr = pPr.find(qn("w:numPr"))
                    if numPr is not None:
                        ref_list_numPr = copy.deepcopy(numPr)

    # Strip dummy section from template
    for child in dummy_children:
        body.remove(child)

    return ref_tbl_xml, ref_list_numPr


def _apply_table_style(source_tbl_xml, ref_tbl_xml) -> None:
    """
    Apply reference table formatting to a source table.
    Applies tblPr, header row shading and border, body row run properties.
    REF_ID: RM_DO178_001
    """
    if ref_tbl_xml is None:
        return

    ref_tblPr = ref_tbl_xml.find(qn("w:tblPr"))
    if ref_tblPr is not None:
        new_tblPr = copy.deepcopy(ref_tblPr)
        for cap in new_tblPr.findall(qn("w:tblCaption")):
            new_tblPr.remove(cap)
        for w in new_tblPr.findall(qn("w:tblW")):
            new_tblPr.remove(w)
        auto_w = OxmlElement("w:tblW")
        auto_w.set(qn("w:w"), "0")
        auto_w.set(qn("w:type"), "auto")
        new_tblPr.insert(0, auto_w)
        src_tblPr = source_tbl_xml.find(qn("w:tblPr"))
        if src_tblPr is not None:
            source_tbl_xml.replace(src_tblPr, new_tblPr)
        else:
            source_tbl_xml.insert(0, new_tblPr)

    ref_rows = ref_tbl_xml.findall(qn("w:tr"))
    ref_hdr_tc_shd = ref_hdr_para_bdr = ref_hdr_rPr = None
    ref_body_rPr = ref_body_spacing = None

    if ref_rows:
        for ref_tc in ref_rows[0].findall(qn("w:tc")):
            tcPr = ref_tc.find(qn("w:tcPr"))
            if tcPr is not None:
                ref_hdr_tc_shd = tcPr.find(qn("w:shd"))
            for ref_p in ref_tc.findall(qn("w:p")):
                pPr = ref_p.find(qn("w:pPr"))
                if pPr is not None:
                    ref_hdr_para_bdr = pPr.find(qn("w:pBdr"))
                for ref_r in ref_p.findall(qn("w:r")):
                    rPr = ref_r.find(qn("w:rPr"))
                    if rPr is not None:
                        ref_hdr_rPr = rPr
                        break
            if ref_hdr_rPr is not None:
                break

    for ref_row in ref_rows[1:]:
        for ref_tc in ref_row.findall(qn("w:tc")):
            for ref_p in ref_tc.findall(qn("w:p")):
                pPr = ref_p.find(qn("w:pPr"))
                if pPr is not None and ref_body_spacing is None:
                    ref_body_spacing = pPr.find(qn("w:spacing"))
                for ref_r in ref_p.findall(qn("w:r")):
                    rPr = ref_r.find(qn("w:rPr"))
                    if rPr is not None and ref_body_rPr is None:
                        ref_body_rPr = rPr
        if ref_body_rPr is not None:
            break

    if ref_body_rPr is None:
        return

    for row_idx, src_tr in enumerate(source_tbl_xml.findall(qn("w:tr"))):
        is_hdr   = (row_idx == 0)
        tgt_rPr  = ref_hdr_rPr if (is_hdr and ref_hdr_rPr is not None) else ref_body_rPr

        for src_tc in src_tr.findall(qn("w:tc")):
            tcPr = src_tc.find(qn("w:tcPr"))
            if tcPr is None:
                tcPr = OxmlElement("w:tcPr")
                src_tc.insert(0, tcPr)

            # Remove any source cell shading from all rows
            for s in tcPr.findall(qn("w:shd")):
                tcPr.remove(s)

            if is_hdr and ref_hdr_tc_shd is not None:
                # Apply template header shading
                tcPr.append(copy.deepcopy(ref_hdr_tc_shd))
            else:
                # Body cells: set clear shading so page background shows through
                clear_shd = OxmlElement("w:shd")
                clear_shd.set(qn("w:val"),   "clear")
                clear_shd.set(qn("w:color"), "auto")
                clear_shd.set(qn("w:fill"),  "auto")
                tcPr.append(clear_shd)

            for src_p in src_tc.findall(qn("w:p")):
                pPr = src_p.find(qn("w:pPr"))
                if pPr is None:
                    pPr = OxmlElement("w:pPr")
                    src_p.insert(0, pPr)

                if is_hdr and ref_hdr_para_bdr is not None:
                    for b in pPr.findall(qn("w:pBdr")):
                        pPr.remove(b)
                    pPr.append(copy.deepcopy(ref_hdr_para_bdr))

                if not is_hdr and ref_body_spacing is not None:
                    for s in pPr.findall(qn("w:spacing")):
                        pPr.remove(s)
                    pPr.append(copy.deepcopy(ref_body_spacing))

                for src_r in src_p.findall(qn("w:r")):
                    if (src_r.find(qn("w:fldChar")) is not None or
                            src_r.find(qn("w:instrText")) is not None):
                        continue
                    old = src_r.find(qn("w:rPr"))
                    if old is not None:
                        src_r.remove(old)
                    src_r.insert(0, copy.deepcopy(tgt_rPr))


def _apply_list_style(para_xml, ref_list_numPr) -> None:
    """
    Apply the reference list numPr to a source list paragraph.
    REF_ID: RM_DO178_001
    """
    if ref_list_numPr is None:
        return
    pPr = para_xml.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        para_xml.insert(0, pPr)
    for old in pPr.findall(qn("w:numPr")):
        pPr.remove(old)
    pPr.append(copy.deepcopy(ref_list_numPr))


def _make_para(text: str, style_id: str) -> object:
    """Build a new paragraph XML element with given text and style ID."""
    new_p = OxmlElement("w:p")
    pPr   = OxmlElement("w:pPr")
    pSty  = OxmlElement("w:pStyle")
    pSty.set(qn("w:val"), style_id)
    pPr.append(pSty)
    new_p.append(pPr)
    if text:
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r.append(t)
        new_p.append(r)
    return new_p


def _copy_para_with_style(para, style_id: str) -> object:
    """Deep-copy a paragraph and set its style to style_id."""
    new_p = copy.deepcopy(para._p)
    _set_paragraph_style(new_p, style_id)
    return new_p


# ---------------------------------------------------------------------------
# Source classification helpers
# ---------------------------------------------------------------------------

def _is_reference_entry(text: str) -> bool:
    """Return True if this line looks like a citation entry."""
    return bool(_REF_ENTRY_RE.match(text.strip()))


def _is_references_heading(text: str) -> bool:
    """Return True if this paragraph is the References section heading."""
    clean = text.strip().lower()
    return clean in {"references", "bibliography", "works cited"}


# ---------------------------------------------------------------------------
# Main formatter
# ---------------------------------------------------------------------------

def _format_docx(
    source_path: Path,
    template_doc: Any,
    confirmed_elements: List[ImplicitElement],
    ref_tbl_xml=None,
    ref_list_numPr=None,
) -> FormatterResult:
    """
    Insert source document content into the correct template sections.

    Algorithm:
    1. Find the [[BODY_CONTENT_START]] placeholder -- this is the insertion
       point for all body content (section 3).
    2. Find the [[REFERENCES_START]] placeholder -- insertion point for
       reference entries (section 4).
    3. Walk source paragraphs:
       - If the paragraph is the References heading or comes after it,
         route to the references section.
       - Otherwise route to the body section.
    4. Replace each placeholder with the routed content.
    5. Apply correct styles using the template's XML style IDs.
    6. Split heading+body concatenations deterministically.

    REF_ID: RM_DO178_001
    """
    try:
        source_doc = DocxDocument(str(source_path))
    except Exception:
        return FormatterResult(error="Rhema could not open the source document.")

    template_styles = _get_template_style_names(template_doc)
    style_id_map    = _build_style_id_map(template_doc)
    unmapped: List[str] = []

    implicit_map: Dict[int, str] = {
        el.paragraph_index: el.inferred_type
        for el in confirmed_elements
    }

    body = template_doc.element.body
    children = list(body)

    # Find placeholder elements by their text content
    body_placeholder   = None
    refs_placeholder   = None

    for child in children:
        if not child.tag.endswith("}p"):
            continue
        text = "".join(t.text or "" for t in child.iter() if t.tag.endswith("}t"))
        if "[[BODY_CONTENT_START]]" in text:
            body_placeholder = child
        elif "[[REFERENCES_START]]" in text:
            refs_placeholder = child

    if body_placeholder is None:
        return FormatterResult(error="Template is missing the body content marker.")

    # Build body and references paragraph lists from source.
    # Walk the source body XML directly to preserve tables (w:tbl) in document order.
    # source_doc.paragraphs only returns w:p elements -- tables would be skipped.
    body_paras   = []   # XML elements to insert at body placeholder
    refs_paras   = []   # XML elements to insert at refs placeholder

    in_references = False

    # Build a paragraph-index map for the implicit_map lookup.
    # paragraph_index in confirmed_elements refers to position in source_doc.paragraphs.
    # When walking XML children we need to track the paragraph counter separately.
    # Pre-build a list mapping XML element id to paragraph object
    # so we can look up style/text without reconstructing Paragraph objects.
    source_para_list = source_doc.paragraphs
    source_para_by_elem = {id(p._p): p for p in source_para_list}

    source_body_children = list(source_doc.element.body)
    para_index = 0  # tracks position in source_doc.paragraphs

    # Detect where the actual body content starts in the source document.
    # Skip leading cover page content and TOC fields so the template's
    # cover page and TOC are not duplicated. REF_ID: RM_DO178_001
    _source_body_start = _find_source_body_start(
        source_body_children, implicit_map, source_para_by_elem
    )

    for child_idx, child in enumerate(source_body_children):
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        # Skip everything before the detected body start (cover page, TOC)
        if child_idx < _source_body_start:
            if tag == "p":
                para_index += 1
            continue

        # --- Table: deepcopy, apply template style, insert in order ---
        if tag == "tbl":
            new_tbl = copy.deepcopy(child)
            for p_el in new_tbl.findall(".//" + qn("w:p")):
                _strip_run_formatting(p_el)
            _apply_table_style(new_tbl, ref_tbl_xml)
            if in_references:
                refs_paras.append(new_tbl)
            else:
                body_paras.append(new_tbl)
            continue

        # --- Skip non-paragraph, non-table elements (sectPr, bookmarks etc.) ---
        if tag != "p":
            continue

        # --- Paragraph: look up the python-docx Paragraph object by element id ---
        para = source_para_by_elem.get(id(child))
        if para is None:
            para_index += 1
            continue

        i = para_index
        para_index += 1

        text         = para.text.strip()
        source_style: str = (para.style.name or STYLE_NORMAL) if para.style else STYLE_NORMAL
        source_lower = source_style.lower()

        # Detect transition into references section.
        # Handles both plain "References" and numbered "5. References".
        clean_text = text.lower().strip()
        is_refs = _is_references_heading(text)
        # Also catch numbered: "5. References", "5. Bibliography" etc.
        if not is_refs:
            m = re.match(r"^(\d+\.)+\s*(\S+)", text.strip())
            if m:
                first_word = m.group(2).strip().rstrip(".,;:").lower()
                if first_word in {"references", "bibliography", "works cited"}:
                    is_refs = True
        if is_refs:
            in_references = True

        if in_references:
            # Route to references section.
            # The REFERENCES heading is already declared in the template -- skip it.
            # Any text content (citations, body text) goes to the refs section.
            if is_refs:
                # This paragraph IS the references heading (plain or numbered).
                # If it is numbered (e.g. "5. References Florida..."), extract
                # the body text portion and route it to refs as a citation.
                _, body_text = _split_heading_from_body(text) if text else (text, None)
                if body_text:
                    new_p = _make_para(body_text, style_id_map.get(STYLE_NORMAL, STYLE_NORMAL))
                    refs_paras.append(new_p)
                continue
            if text:
                new_p = _make_para(text, style_id_map.get(STYLE_NORMAL, STYLE_NORMAL))
                refs_paras.append(new_p)
            continue

        # --- Body section routing ---

        if not text:
            # Preserve empty paragraphs for spacing
            body_paras.append(OxmlElement("w:p"))
            continue

        # Determine target style
        target_style = STYLE_NORMAL

        if i in implicit_map:
            label        = implicit_map[i].lower()
            target_style = _HEADING_LABEL_TO_STYLE.get(label, STYLE_NORMAL)
        elif source_lower in _WORD_HEADING_STYLES:
            mapped = source_style.title()
            if mapped in template_styles:
                target_style = mapped
            else:
                for candidate in [STYLE_H1, STYLE_H2, STYLE_H3]:
                    if candidate in template_styles:
                        target_style = candidate
                        break

        if target_style not in template_styles:
            if target_style not in unmapped:
                unmapped.append(target_style)
            target_style = STYLE_NORMAL

        style_id = style_id_map.get(target_style, target_style)

        # Check for heading+body concatenation and split if needed
        if i in implicit_map and text:
            heading_text, body_text = _split_heading_from_body(text)
        else:
            heading_text, body_text = text, None

        if body_text is not None:
            # Insert heading paragraph
            h_p = copy.deepcopy(para._p)
            # Strip source run formatting so template style governs
            _strip_run_formatting(h_p)
            # Clear runs and set heading text only
            for r in h_p.findall(qn("w:r")):
                h_p.remove(r)
            r_el = OxmlElement("w:r")
            t_el = OxmlElement("w:t")
            t_el.text = heading_text
            t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            r_el.append(t_el)
            h_p.append(r_el)
            _set_paragraph_style(h_p, style_id)
            body_paras.append(h_p)

            # Insert body text as Normal
            normal_id = style_id_map.get(STYLE_NORMAL, STYLE_NORMAL)
            body_paras.append(_make_para(body_text, normal_id))
        else:
            # Normal copy with style applied
            # Strip source run formatting so template style governs color/font
            new_p = copy.deepcopy(para._p)
            _strip_run_formatting(new_p)
            _set_paragraph_style(new_p, style_id)
            # Apply list style if this paragraph is a list item
            src_pPr = para._p.find(qn("w:pPr"))
            if src_pPr is not None and src_pPr.find(qn("w:numPr")) is not None:
                _apply_list_style(new_p, ref_list_numPr)
            body_paras.append(new_p)

    # Insert body content -- replace placeholder with collected paragraphs
    body_idx = list(body).index(body_placeholder)
    body.remove(body_placeholder)
    for offset, p_el in enumerate(body_paras):
        body.insert(body_idx + offset, p_el)

    # Insert references content -- replace placeholder with collected paragraphs
    if refs_placeholder is not None:
        # Refresh children list after body insertion
        refs_idx = list(body).index(refs_placeholder)
        body.remove(refs_placeholder)
        for offset, p_el in enumerate(refs_paras):
            body.insert(refs_idx + offset, p_el)

    return FormatterResult(
        success=True,
        formatted_doc=template_doc,
        unmapped_styles=unmapped,
    )


def _format_txt(
    source_path: Path,
    template_doc: Any,
    confirmed_elements: List[ImplicitElement],
) -> FormatterResult:
    """Apply template styles to a plain text source file."""
    try:
        lines = source_path.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError:
        return FormatterResult(error="Rhema could not open the source document.")

    template_styles = _get_template_style_names(template_doc)
    style_id_map    = _build_style_id_map(template_doc)
    unmapped: List[str] = []

    implicit_map: Dict[int, str] = {
        el.paragraph_index: el.inferred_type
        for el in confirmed_elements
    }

    body = template_doc.element.body
    children = list(body)

    body_placeholder = None
    refs_placeholder = None

    for child in children:
        if not child.tag.endswith("}p"):
            continue
        text = "".join(t.text or "" for t in child.iter() if t.tag.endswith("}t"))
        if "[[BODY_CONTENT_START]]" in text:
            body_placeholder = child
        elif "[[REFERENCES_START]]" in text:
            refs_placeholder = child

    if body_placeholder is None:
        return FormatterResult(error="Template is missing the body content marker.")

    body_paras = []
    refs_paras = []
    in_references = False

    for i, line in enumerate(lines):
        stripped = line.strip()

        # Detect references section -- same logic as _format_docx
        is_refs = _is_references_heading(stripped)
        if not is_refs:
            m = re.match(r"^(\d+\.)+\s*(\S+)", stripped)
            if m:
                first_word = m.group(2).strip().rstrip(".,;:").lower()
                if first_word in {"references", "bibliography", "works cited"}:
                    is_refs = True
        if is_refs:
            in_references = True

        if in_references:
            if is_refs:
                # References section heading -- already in template, skip it.
                # Extract any concatenated body text and add as first citation.
                _, body_text = _split_heading_from_body(stripped) if stripped else (stripped, None)
                if body_text:
                    refs_paras.append(_make_para(body_text, style_id_map.get(STYLE_NORMAL, STYLE_NORMAL)))
                continue
            if stripped:
                refs_paras.append(_make_para(stripped, style_id_map.get(STYLE_NORMAL, STYLE_NORMAL)))
            continue

        target_style = STYLE_NORMAL
        if i in implicit_map:
            label        = implicit_map[i].lower()
            target_style = _HEADING_LABEL_TO_STYLE.get(label, STYLE_NORMAL)

        if target_style not in template_styles:
            if target_style not in unmapped:
                unmapped.append(target_style)
            target_style = STYLE_NORMAL

        style_id = style_id_map.get(target_style, target_style)

        if i in implicit_map and stripped:
            heading_text, body_text = _split_heading_from_body(stripped)
        else:
            heading_text, body_text = stripped, None

        if body_text is not None:
            body_paras.append(_make_para(heading_text, style_id))
            body_paras.append(_make_para(body_text, style_id_map.get(STYLE_NORMAL, STYLE_NORMAL)))
        else:
            body_paras.append(_make_para(stripped, style_id))

    body_idx = list(body).index(body_placeholder)
    body.remove(body_placeholder)
    for offset, p_el in enumerate(body_paras):
        body.insert(body_idx + offset, p_el)

    if refs_placeholder is not None:
        refs_idx = list(body).index(refs_placeholder)
        body.remove(refs_placeholder)
        for offset, p_el in enumerate(refs_paras):
            body.insert(refs_idx + offset, p_el)

    return FormatterResult(
        success=True,
        formatted_doc=template_doc,
        unmapped_styles=unmapped,
    )


# ---------------------------------------------------------------------------
# Template opener (handles both .docx and .dotx content types)
# python-docx refuses to open files with wordprocessingml.template content type.
# This helper patches the content type before opening if needed.
# REF_ID: RM_DO178_001
# ---------------------------------------------------------------------------

def _open_template(template_path: Path):
    """
    Open a template file as a python-docx Document.
    Handles .dotx files (Word Template format) by patching the content type.
    Returns the Document object or None on failure.
    """
    import io as _io
    import zipfile as _zf

    # Try direct open first (.tmpt and .docx work without patching)
    try:
        return DocxDocument(str(template_path))
    except Exception:
        pass

    # Patch content type for .dotx files and retry
    try:
        with _zf.ZipFile(str(template_path), "r") as z:
            files = {n: z.read(n) for n in z.namelist()}

        ct = files["[Content_Types].xml"].decode("utf-8")
        ct = ct.replace(
            "wordprocessingml.template",
            "wordprocessingml.document"
        )
        files["[Content_Types].xml"] = ct.encode("utf-8")

        buf = _io.BytesIO()
        with _zf.ZipFile(buf, "w", _zf.ZIP_DEFLATED) as z:
            for name, data in files.items():
                z.writestr(name, data)
        buf.seek(0)
        return DocxDocument(buf)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Normal style spacing enforcer
# Ensures the template's Normal style has single-line, one-blank-line-between
# paragraph spacing regardless of what the template declares.
# This is enforced by the CLI, not by the template itself.
# REF_ID: RM_DO178_001
# ---------------------------------------------------------------------------

_TARGET_SPACING_BEFORE   = "0"
_TARGET_SPACING_AFTER    = "240"   # one blank line gap between paragraphs
_TARGET_SPACING_LINE     = "240"   # single line height
_TARGET_SPACING_LINERULE = "auto"


def _enforce_normal_spacing(template_doc: Any) -> None:
    """
    Check the template's Normal style paragraph spacing and correct it
    if it does not match the CLI's declared APA-compliant spacing standard.

    Target: before=0, after=240 (one line gap), line=240 (single), lineRule=auto.
    If any value differs, the entire spacing declaration is replaced.

    This runs on the template before any content is inserted so all body
    paragraphs inherit the correct spacing via the Normal style.
    REF_ID: RM_DO178_001
    """
    import re as _re

    # Locate the Normal style element in the styles part
    styles_part = template_doc.part.styles
    if styles_part is None:
        return

    normal_style = None
    for style in template_doc.styles:
        if style.style_id == "Normal":
            normal_style = style
            break

    if normal_style is None:
        return

    # Get or create pPr on the Normal style element
    elem = normal_style.element
    pPr = elem.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        elem.insert(0, pPr)

    # Check existing spacing
    spacing = pPr.find(qn("w:spacing"))
    needs_fix = False

    if spacing is None:
        needs_fix = True
    else:
        before    = spacing.get(qn("w:before"),    "0")
        after     = spacing.get(qn("w:after"),     "0")
        line      = spacing.get(qn("w:line"),      "240")
        line_rule = spacing.get(qn("w:lineRule"),  "auto")
        if (before    != _TARGET_SPACING_BEFORE   or
                after     != _TARGET_SPACING_AFTER    or
                line      != _TARGET_SPACING_LINE     or
                line_rule != _TARGET_SPACING_LINERULE):
            needs_fix = True
            pPr.remove(spacing)

    if needs_fix:
        new_spacing = OxmlElement("w:spacing")
        new_spacing.set(qn("w:before"),    _TARGET_SPACING_BEFORE)
        new_spacing.set(qn("w:after"),     _TARGET_SPACING_AFTER)
        new_spacing.set(qn("w:line"),      _TARGET_SPACING_LINE)
        new_spacing.set(qn("w:lineRule"),  _TARGET_SPACING_LINERULE)
        pPr.insert(0, new_spacing)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def format_document(
    source_path: Path,
    template_path: Path,
    confirmed_elements: List[ImplicitElement],
) -> FormatterResult:
    """Apply the template's styles to the source document. REF_ID: RM_DO178_001"""
    if not source_path.exists():
        return FormatterResult(error="Rhema could not find the source document.")
    if not template_path.exists():
        return FormatterResult(error="Rhema could not find the template.")
    template_doc = _open_template(template_path)
    if template_doc is None:
        return FormatterResult(error="Rhema could not open the template.")

    # Enforce correct paragraph spacing on the Normal style
    _enforce_normal_spacing(template_doc)

    # Sample table/list styles from the dummy section, then strip it
    ref_tbl_xml, ref_list_numPr = _sample_template_styles(template_doc)

    suffix = source_path.suffix.lower()
    if suffix == ".docx":
        return _format_docx(source_path, template_doc, confirmed_elements,
                            ref_tbl_xml, ref_list_numPr)
    if suffix == ".txt":
        return _format_txt(source_path, template_doc, confirmed_elements)
    return FormatterResult(error="Unsupported format. Rhema works with .txt and .docx files.")
