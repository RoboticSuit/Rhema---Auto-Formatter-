# Source Trace:
# File: detector_v0.2.py
# Knowledge Files: CodeSourceDB v3.6, SyntaxBiasDB v2.3, HumanSyntaxDB v1.2
# REF_IDs: RM_DO178_001, RM_ISO_001, WEB_PY_001, WEB_PY_009, WEB_PY_010
# Logic: Implicit structure detection using explicitly typed variables, strict
#        truthiness checks, and named constants.
#        v0.2 -- _NUMBERED_HEADING_RE renamed to NUMBERED_HEADING_RE (public) so
#                formatter.py can import it without violating WEB_PY_004.
#                Font size conversion uses OOXML_EMU_PER_HALF_POINT constant.

"""
Rhema -- Auto Formatter
core/detector.py

Single responsibility: Detects implicit structure in an input document.
Reads the document and identifies paragraphs that visually resemble
headings or list items but are not declared as Word heading or list styles.

Returns a DetectionResult containing:
  - The list of confirmed Word-styled paragraphs (no action needed)
  - The list of implicitly structured paragraphs (need Preview Mode review)
  - A flag indicating whether Preview Mode must be triggered

This module reads documents only. It never modifies them.

Supported input formats: .docx, .txt
.doc files must be converted to .docx before reaching this module.

Governing standards:
  DO-178C source traceability    -- REF_ID: RM_DO178_001
  ISO 31000 risk identification  -- REF_ID: RM_ISO_001
"""

import re
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Optional, Any

# python-docx for .docx parsing
try:
    from docx import Document as DocxDocument
    from docx.oxml.ns import qn
    DOCX_AVAILABLE: bool = True
except ImportError:
    DOCX_AVAILABLE = False
    DocxDocument: Any = None
    def qn(tag: str) -> str:
        return tag

from config.constants import SUPPORTED_INPUT_EXTENSIONS, OOXML_EMU_PER_HALF_POINT


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class ImplicitElement:
    """
    A paragraph that appears to have structural meaning (heading, list item)
    but is not declared as a Word heading or list style.
    """
    paragraph_index: int
    text_excerpt:    str
    inferred_type:   str
    confidence:      str = "high"


@dataclass
class DetectionResult:
    """
    Output of the detector. Consumed by preview.py and single.py/batch.py.
    """
    needs_preview:    bool = False
    implicit:         List[ImplicitElement] = field(default_factory=list)
    declared_count:   int = 0
    total_paragraphs: int = 0
    source_format:    str = ""
    error:            str = ""


# ---------------------------------------------------------------------------
# Constants -- Word heading and list style names
# ---------------------------------------------------------------------------

_WORD_HEADING_STYLES: set[str] = {
    "heading 1", "heading 2", "heading 3",
    "heading 4", "heading 5", "heading 6",
    "title", "subtitle",
}

_WORD_LIST_STYLES: set[str] = {
    "list paragraph", "list bullet", "list number",
    "list bullet 2", "list bullet 3",
    "list number 2", "list number 3",
}

# Heuristics for implicit structure detection
_MAX_HEADING_WORDS: int       = 12
_SHORT_HEADING_MIN_WORDS: int = 2
_UNNUMBERED_SHORT_MAX_WORDS: int = 5
_BODY_PARAGRAPH_MIN_WORDS: int = 8

_MIN_HEADING_FONT_SIZE_HP: int = 28   # 14pt in half-points
_HEADING_1_FONT_SIZE_HP: int   = 40   # 20pt in half-points

# ---------------------------------------------------------------------------
# Compiled regex patterns -- module-level public constants.
# Public (no leading underscore) so formatter.py can import them without
# violating the WEB_PY_004 private name rule. REF_ID: WEB_PY_009
# ---------------------------------------------------------------------------

NUMBERED_HEADING_RE = re.compile(r'^\d+(\.(\d+))*\.?\s+\S')
ROMAN_HEADING_RE    = re.compile(r'^(I{1,3}|IV|V|VI{0,3}|IX|X)\.\s+\S', re.IGNORECASE)
LETTERED_HEADING_RE = re.compile(r'^[A-Z]\.\s+\S')


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _is_all_caps_short(text: str) -> bool:
    """
    True if the text is short, all-uppercase, and contains at least
    two words -- a common implicit heading pattern.
    """
    stripped: str = text.strip()
    words: List[str] = stripped.split()

    if len(words) < _SHORT_HEADING_MIN_WORDS or len(words) > _MAX_HEADING_WORDS:
        return False

    alpha_chars: List[str] = [c for c in stripped if c.isalpha()]
    if len(alpha_chars) == 0:
        return False

    return all(c.isupper() for c in alpha_chars)


def _get_run_font_size(paragraph: Any) -> Optional[int]:
    """
    Return the font size in half-points of the first run in a paragraph,
    or None if not explicitly set.
    Uses OOXML_EMU_PER_HALF_POINT constant -- no magic number. REF_ID: RM_DO178_001
    """
    for run in paragraph.runs:
        if run.font.size is not None:
            return int(run.font.size / OOXML_EMU_PER_HALF_POINT)
    return None


def _infer_heading_level(text: str, font_size_hp: Optional[int]) -> str:
    """
    Infer a human-readable heading level label from text and font size.
    REF_ID: RM_DO178_001
    """
    m = re.match(r"^(\d+(\.\d+)*)", text.strip())
    if m is not None:
        depth: int = len(m.group(1).split("."))
        if depth == 1:
            return "Heading 1"
        elif depth == 2:
            return "Heading 2"
        else:
            return "Heading 3"

    if (ROMAN_HEADING_RE.match(text.strip()) is not None or
            LETTERED_HEADING_RE.match(text.strip()) is not None):
        return "Heading 1"

    if font_size_hp is not None:
        if font_size_hp >= _HEADING_1_FONT_SIZE_HP:
            return "Heading 1"
        elif font_size_hp >= _MIN_HEADING_FONT_SIZE_HP:
            return "Heading 2"

    if _is_all_caps_short(text) is True:
        return "Heading 1"

    return "Heading 1"


def _excerpt(text: str, max_chars: int = 60) -> str:
    """Return a trimmed excerpt of the paragraph text for display."""
    text = text.strip()
    if len(text) <= max_chars:
        return text
    return text[:max_chars].rstrip() + "..."


# ---------------------------------------------------------------------------
# .docx detector
# ---------------------------------------------------------------------------

def _detect_docx(path: Path) -> DetectionResult:
    """
    Parse a .docx file and detect implicit structure.
    REF_ID: RM_DO178_001
    """
    if DOCX_AVAILABLE is False:
        return DetectionResult(
            error="python-docx is not installed. Run: pip install python-docx"
        )

    try:
        doc = DocxDocument(str(path))
    except Exception:
        return DetectionResult(error="Rhema could not open the file.")

    result = DetectionResult(source_format="docx")
    paragraphs: List[Any] = doc.paragraphs
    result.total_paragraphs = len(paragraphs)

    for i, para in enumerate(paragraphs):
        text: str = para.text.strip()
        if len(text) == 0:
            continue

        style_name: str = para.style.name.lower() if para.style is not None else ""

        if style_name in _WORD_HEADING_STYLES:
            result.declared_count += 1
            continue

        if style_name in _WORD_LIST_STYLES:
            result.declared_count += 1
            continue

        pPr = para._p.find(qn("w:pPr"))
        if pPr is not None and pPr.find(qn("w:numPr")) is not None:
            if style_name in _WORD_LIST_STYLES:
                result.declared_count += 1
                continue

        word_count: int = len(text.split())
        font_size_hp: Optional[int] = _get_run_font_size(para)
        is_bold: bool = any(
            run.bold is True for run in para.runs
            if len(run.text.strip()) > 0
        )

        signals: List[str] = []

        if (NUMBERED_HEADING_RE.match(text) is not None or
                ROMAN_HEADING_RE.match(text) is not None or
                LETTERED_HEADING_RE.match(text) is not None):
            signals.append("numbered_pattern")

        if _is_all_caps_short(text) is True:
            signals.append("all_caps_short")

        if font_size_hp is not None and font_size_hp >= _MIN_HEADING_FONT_SIZE_HP:
            signals.append("large_font")

        if (is_bold is True and word_count <= _MAX_HEADING_WORDS and
                text.endswith(".") is False):
            signals.append("bold_short")

        # Unnumbered section heading check
        if (word_count <= _UNNUMBERED_SHORT_MAX_WORDS and
                len(text) > 0 and
                text[-1] not in ".,:;?!)" and
                i + 1 < len(paragraphs) and
                len(paragraphs[i + 1].text.split()) > _BODY_PARAGRAPH_MIN_WORDS):
            signals.append("short_no_punct_before_body")

        if len(signals) == 0:
            continue

        confidence: str = "high" if len(signals) >= 2 else "low"
        inferred: str = _infer_heading_level(text, font_size_hp)

        result.implicit.append(ImplicitElement(
            paragraph_index=i,
            text_excerpt=_excerpt(text),
            inferred_type=inferred,
            confidence=confidence,
        ))

    result.needs_preview = len(result.implicit) > 0
    return result


# ---------------------------------------------------------------------------
# .txt detector
# ---------------------------------------------------------------------------

def _detect_txt(path: Path) -> DetectionResult:
    """
    Parse a plain text file and detect implicit structure.
    All headings in a .txt file are implicit -- there are no Word styles.
    REF_ID: RM_DO178_001
    """
    try:
        text: str = path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return DetectionResult(error="Rhema could not open the file.")

    lines: List[str] = text.splitlines()
    result = DetectionResult(source_format="txt", total_paragraphs=len(lines))

    for i, line in enumerate(lines):
        stripped: str = line.strip()
        if len(stripped) == 0:
            continue

        word_count: int = len(stripped.split())
        signals: List[str] = []

        if NUMBERED_HEADING_RE.match(stripped) is not None:
            signals.append("numbered_pattern")
        if ROMAN_HEADING_RE.match(stripped) is not None:
            signals.append("roman_pattern")
        if LETTERED_HEADING_RE.match(stripped) is not None:
            signals.append("lettered_pattern")
        if _is_all_caps_short(stripped) is True:
            signals.append("all_caps_short")

        if (word_count <= _MAX_HEADING_WORDS and
                len(stripped) > 0 and
                stripped[-1] not in ".,:;?!)"):
            signals.append("short_no_punct")

        if len(signals) == 0:
            continue

        confidence: str = "high" if len(signals) >= 2 else "low"
        inferred: str = _infer_heading_level(stripped, None)

        result.implicit.append(ImplicitElement(
            paragraph_index=i,
            text_excerpt=_excerpt(stripped),
            inferred_type=inferred,
            confidence=confidence,
        ))

    result.needs_preview = len(result.implicit) > 0
    return result


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def detect(path: Path) -> DetectionResult:
    """
    Analyse the document at the given path and return a DetectionResult.
    REF_ID: RM_DO178_001
    """
    suffix: str = path.suffix.lower()

    if suffix not in SUPPORTED_INPUT_EXTENSIONS:
        return DetectionResult(
            error="Unsupported file format. Rhema works with .txt, .docx, and .doc files only."
        )

    if path.exists() is False:
        return DetectionResult(
            error="Rhema could not find the file: " + str(path)
        )

    if suffix == ".docx":
        return _detect_docx(path)

    if suffix == ".txt":
        return _detect_txt(path)

    return DetectionResult(
        error="This file needs to be converted to .docx before Rhema can process it."
    )
