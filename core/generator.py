# Source Trace:
# File: generator_v0.4.py
# Knowledge Files: CodeSourceDB v3.6, SyntaxBiasDB v2.3, HumanSyntaxDB v1.2
# REF_IDs: RM_DO178_001, RM_NIST_001, RM_HCI_002, WEB_PY_001, WEB_PY_003
# Logic: Fully dynamic placeholder replacement. CoverMetadata reduced to two
#        required fields (title, author). All other cover fields discovered by
#        scanning the template and prompted generically by token name.
#        v0.4 -- CoverMetadata removed. prompt_cover_metadata returns Dict[str, str].
#                generate() accepts Dict[str, str]. Unknown custom placeholders
#                are discovered and prompted automatically.

"""
Rhema -- Auto Formatter
core/generator.py

Single responsibility: Injects cover page metadata into a formatted document
by replacing [[PLACEHOLDER]] tokens with user-supplied values.

Two tiers of fields:

  Tier 1 -- Always required (hardcoded):
    [[DOCUMENT_TITLE]]  -- defaults to source filename stem
    [[AUTHOR_NAME]]     -- always prompted, no default

  Tier 2 -- Fully dynamic (template-driven):
    Every other [[TOKEN]] found in the template is discovered by scanning
    the document and prompted generically. The user may press Enter to skip
    any Tier 2 field. Unknown custom tokens (e.g. [[EMPLOYEE_ID]]) are
    handled identically to known ones -- Rhema prompts for them by name.

Structural markers ([[BODY_CONTENT_START]], [[REFERENCES_START]]) are
silently replaced with empty string and never prompted.

Governing standards:
  DO-178C source traceability    -- REF_ID: RM_DO178_001
  NIST RMF Step 6 HITL gate     -- REF_ID: RM_NIST_001
  ISO 9241-110:2020 interaction -- REF_ID: RM_HCI_002
"""

import re
from pathlib import Path
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional, Set

from docx import Document as DocxDocument
from docx.oxml.ns import qn


# ---------------------------------------------------------------------------
# Structural markers -- replaced silently, never prompted.
# REF_ID: RM_DO178_001
# ---------------------------------------------------------------------------

_STRUCTURAL_PLACEHOLDERS: Set[str] = {
    "[[BODY_CONTENT_START]]",
    "[[REFERENCES_START]]",
}

# ---------------------------------------------------------------------------
# Tier 1 -- always required fields.
# REF_ID: RM_DO178_001, RM_NIST_001
# ---------------------------------------------------------------------------

_REQUIRED_PLACEHOLDER_TITLE:  str = "[[DOCUMENT_TITLE]]"
_REQUIRED_PLACEHOLDER_AUTHOR: str = "[[AUTHOR_NAME]]"

# ---------------------------------------------------------------------------
# Known smart defaults for common Tier 2 tokens.
# Any token not listed here gets an empty default (press Enter to skip).
# Values that require runtime computation are declared as empty here and
# filled by _get_default() at prompt time. REF_ID: RM_DO178_001
# ---------------------------------------------------------------------------

_DATE_TOKENS: Set[str] = {"[[MONTH_YEAR]]", "[[SUBMISSION_DATE]]"}

_KNOWN_DEFAULTS: Dict[str, str] = {
    "[[DOCUMENT_TYPE]]":          "Report",
    "[[CONFIDENTIALITY_LABEL]]":  "INTERNAL",
}

# ---------------------------------------------------------------------------
# Result data structure
# ---------------------------------------------------------------------------

@dataclass
class GeneratorResult:
    """
    Output of the generator.

    success        -- True if replacement ran without error
    final_doc      -- python-docx Document ready for writer.py (None on error)
    missing_fields -- tokens that remained unfilled after replacement
    error          -- plain-language error message (empty on success)
    """
    success:        bool = False
    final_doc:      Optional[object] = None
    missing_fields: List[str] = field(default_factory=list)
    error:          str = ""


# ---------------------------------------------------------------------------
# Template scanner
# Discovers all [[TOKEN]] placeholders present in the document.
# REF_ID: RM_DO178_001
# ---------------------------------------------------------------------------

_PH_PATTERN = re.compile(r'\[\[[A-Z][A-Z0-9_]*\]\]')


def _scan_placeholders(doc: Any) -> Set[str]:
    """
    Scan all zones of the document (body, tables, headers, footers) for
    [[PLACEHOLDER]] tokens. Returns the set of all unique tokens found.
    REF_ID: RM_DO178_001
    """
    found: Set[str] = set()

    def _scan_paragraphs(paragraphs: Any) -> None:
        for para in paragraphs:
            for match in _PH_PATTERN.finditer(para.text):
                found.add(match.group(0))

    _scan_paragraphs(doc.paragraphs)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                _scan_paragraphs(cell.paragraphs)

    for section in doc.sections:
        for hdr in [section.header, section.first_page_header]:
            if hdr is not None:
                _scan_paragraphs(hdr.paragraphs)
        for ftr in [section.footer, section.first_page_footer]:
            if ftr is not None:
                _scan_paragraphs(ftr.paragraphs)

    return found


# ---------------------------------------------------------------------------
# Display name helper
# Converts [[EMPLOYEE_ID]] -> "Employee Id" for terminal display.
# REF_ID: RM_HCI_002
# ---------------------------------------------------------------------------

def _display_name(token: str) -> str:
    """Strip [[ and ]] and convert SNAKE_CASE to Title Case for display."""
    inner: str = token.strip("[]")
    return inner.replace("_", " ").title()


# ---------------------------------------------------------------------------
# Default value helper
# Returns the appropriate default for a given token at runtime.
# REF_ID: RM_DO178_001
# ---------------------------------------------------------------------------

def _get_default(token: str, source_path: Path) -> str:
    """Return the smart default for a known token, or empty string."""
    if token == _REQUIRED_PLACEHOLDER_TITLE:
        return source_path.stem.replace("_", " ").replace("-", " ").title()

    if token in _DATE_TOKENS:
        import platform as _platform
        _day_fmt: str = "%#d" if _platform.system() == "Windows" else "%-d"
        return datetime.now().strftime(_day_fmt + " %B %Y")

    return _KNOWN_DEFAULTS.get(token, "")


# ---------------------------------------------------------------------------
# Placeholder replacement engine -- unchanged from v0.3
# REF_ID: RM_DO178_001
# ---------------------------------------------------------------------------

def _replace_in_run(run: Any, replacements: Dict[str, str]) -> None:
    """Replace all placeholder tokens in a single run's text."""
    text: str = run.text
    for placeholder, value in replacements.items():
        if placeholder in text:
            text = text.replace(placeholder, value)
    run.text = text


def _replace_placeholders_in_doc(
    doc: Any,
    replacements: Dict[str, str],
) -> List[str]:
    """
    Walk all paragraphs (body, tables, headers, footers) and replace
    placeholder tokens with their values.

    Returns a list of tokens that remained in the document after replacement
    (tokens present but whose value was empty string).
    REF_ID: RM_DO178_001
    """
    _qn = qn

    def _process_paragraph(para: Any) -> None:
        full_text: str = para.text
        has_placeholder: bool = any(ph in full_text for ph in replacements)
        if has_placeholder is False:
            return

        has_any_field: bool = any(
            r._r.find(_qn("w:fldChar")) is not None or
            r._r.find(_qn("w:instrText")) is not None
            for r in para.runs
        )

        if has_any_field is True:
            runs: List[Any] = para.runs
            i: int = 0
            while i < len(runs):
                run: Any = runs[i]
                is_field: bool = (
                    run._r.find(_qn("w:fldChar")) is not None or
                    run._r.find(_qn("w:instrText")) is not None
                )
                if is_field is True:
                    i += 1
                    continue
                j: int = i + 1
                while j < len(runs):
                    r_j: Any = runs[j]
                    if (r_j._r.find(_qn("w:fldChar")) is not None or
                            r_j._r.find(_qn("w:instrText")) is not None):
                        break
                    j += 1
                if j > i + 1:
                    combined: str = "".join(runs[k].text for k in range(i, j))
                    for ph, val in replacements.items():
                        combined = combined.replace(ph, val)
                    runs[i].text = combined
                    for k in range(i + 1, j):
                        runs[k].text = ""
                else:
                    for ph, val in replacements.items():
                        if ph in run.text:
                            run.text = run.text.replace(ph, val)
                i = j
        else:
            if len(para.runs) > 0:
                combined_all: str = "".join(r.text for r in para.runs)
                for ph, val in replacements.items():
                    combined_all = combined_all.replace(ph, val)
                para.runs[0].text = combined_all
                for run in para.runs[1:]:
                    run.text = ""

    for para in doc.paragraphs:
        _process_paragraph(para)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _process_paragraph(para)

    for section in doc.sections:
        for hdr in [section.header, section.first_page_header]:
            if hdr is not None:
                for para in hdr.paragraphs:
                    _process_paragraph(para)
        for ftr in [section.footer, section.first_page_footer]:
            if ftr is not None:
                for para in ftr.paragraphs:
                    _process_paragraph(para)

    # Safety check: report any token that survived replacement with empty value
    unfilled: List[str] = []
    all_text: str = "\n".join(
        p.text for p in doc.paragraphs
    )
    for ph, val in replacements.items():
        if ph in all_text and len(val) == 0:
            unfilled.append(ph)
    return unfilled


# ---------------------------------------------------------------------------
# Public API -- generate()
# REF_ID: RM_DO178_001, RM_NIST_001
# ---------------------------------------------------------------------------

def generate(
    formatted_doc: Any,
    metadata: Dict[str, str],
    source_path: Path,
) -> GeneratorResult:
    """
    Inject cover page values into the formatted document.

    metadata is a Dict[str, str] mapping [[TOKEN]] -> value, produced by
    prompt_cover_metadata(). Structural markers are added here and always
    replaced with empty string.

    REF_ID: RM_DO178_001, RM_NIST_001
    """
    if formatted_doc is None:
        return GeneratorResult(error="No formatted document provided.")

    # Build the full replacement dict.
    # Structural markers are always cleared -- they must not appear in output.
    replacements: Dict[str, str] = {ph: "" for ph in _STRUCTURAL_PLACEHOLDERS}
    replacements.update(metadata)

    unfilled: List[str] = _replace_placeholders_in_doc(formatted_doc, replacements)

    return GeneratorResult(
        success=True,
        final_doc=formatted_doc,
        missing_fields=unfilled,
    )


# ---------------------------------------------------------------------------
# Public API -- prompt_cover_metadata()
# REF_ID: RM_NIST_001, RM_HCI_002
# ---------------------------------------------------------------------------

def prompt_cover_metadata(
    source_path: Path,
    formatted_doc: Any = None,
) -> Dict[str, str]:
    """
    Prompt the user for cover page values and return a Dict[str, str]
    mapping [[TOKEN]] -> user-supplied value.

    Tier 1 (always prompted):
      [[DOCUMENT_TITLE]] -- defaults to source filename stem
      [[AUTHOR_NAME]]    -- no default, user must supply or press Enter

    Tier 2 (template-driven):
      Every other [[TOKEN]] found in formatted_doc is prompted by its
      display name. Known tokens get smart defaults. Unknown custom
      tokens (e.g. [[EMPLOYEE_ID]]) are prompted the same way.
      User may press Enter to leave any Tier 2 field blank.

    Structural markers are excluded from prompting entirely.
    REF_ID: RM_NIST_001, RM_DO178_001
    """
    from ui.messages import (
        MSG_COVER_PAGE_INTRO,
        fmt, LABEL_CONFIRM,
    )

    print(MSG_COVER_PAGE_INTRO)

    def ask(label: str, default: str = "") -> str:
        """Prompt for one field. Returns user input or default on Enter."""
        if len(default) > 0:
            print(fmt(LABEL_CONFIRM, label + " [" + default + "]"))
        else:
            print(fmt(LABEL_CONFIRM, label + " (or press Enter to skip):"))
        try:
            val: str = input("> ").strip()
        except (KeyboardInterrupt, EOFError):
            val = ""
        return val if len(val) > 0 else default

    result: Dict[str, str] = {}

    # --- Tier 1: always required ---
    title_default: str = _get_default(_REQUIRED_PLACEHOLDER_TITLE, source_path)
    result[_REQUIRED_PLACEHOLDER_TITLE] = ask("Document title", title_default)
    result[_REQUIRED_PLACEHOLDER_AUTHOR] = ask("Author name", "")

    if formatted_doc is None:
        return result

    # --- Discover all tokens in the template ---
    all_tokens: Set[str] = _scan_placeholders(formatted_doc)

    # Tier 2: everything except structural markers and already-prompted Tier 1
    skip: Set[str] = _STRUCTURAL_PLACEHOLDERS | {
        _REQUIRED_PLACEHOLDER_TITLE,
        _REQUIRED_PLACEHOLDER_AUTHOR,
    }
    tier2: List[str] = sorted(t for t in all_tokens if t not in skip)

    if len(tier2) > 0:
        print(fmt(LABEL_CONFIRM,
            "Additional fields found in your template. "
            "Press Enter to skip any field."))

    for token in tier2:
        default: str = _get_default(token, source_path)
        label: str   = _display_name(token)
        result[token] = ask(label, default)

    return result
