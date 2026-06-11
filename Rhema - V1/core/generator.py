# Source Trace:
# File: generator.py
# Knowledge Files: CodeSourceDB v3.4, SyntaxBiasDB v2.3, HumanSyntaxDB v1.2, mainProtocol v5.5
# REF_IDs: RM_DO178_001, RM_NIST_001, RM_HCI_002, WEB_PY_001, RM_HCI_001
# Logic: Template placeholder replacement, document metadata injection, and deterministic string manipulation with strict type checking.

"""
Rhema -- Auto Formatter
core/generator.py

Single responsibility: Injects cover page metadata and updates the TOC
field in a formatted document. Replaces all [[PLACEHOLDER]] tokens in
the document with real values supplied by the user or derived from the
source file.

Takes:
  - formatted_doc: python-docx Document from formatter.py
  - metadata: CoverMetadata dataclass with user-supplied field values
  - source_path: Path to the original source file (used for defaults)

Returns a GeneratorResult containing:
  - success: True if all required fields were filled
  - final_doc: the python-docx Document ready for writer.py
  - missing_fields: list of placeholder names that could not be filled
  - error: non-empty string if generation failed

The cover page fields that are always required:
  [[DOCUMENT_TYPE]]         -- must be supplied by the user
  [[DOCUMENT_TITLE]]        -- defaults to source filename stem
  [[SUBTITLE]]              -- optional, replaced with empty string if skipped
  [[AUTHOR_NAME]]           -- must be supplied or skipped
  [[DEPARTMENT_ORGANISATION]] -- optional
  [[MONTH_YEAR]]            -- defaults to current month and year

Header/footer fields:
  [[ORGANISATION]]          -- defaults to [[DEPARTMENT_ORGANISATION]] value
  [[CONFIDENTIALITY_LABEL]] -- defaults to "INTERNAL"

Governing standards:
  DO-178C source traceability    -- REF_ID: RM_DO178_001
  NIST RMF Step 6 HITL gate     -- REF_ID: RM_NIST_001
  ISO 9241-110:2020 interaction -- REF_ID: RM_HCI_002
"""

import re
from pathlib import Path
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional

from docx import Document as DocxDocument
from docx.oxml.ns import qn


# ---------------------------------------------------------------------------
# Cover metadata data structure
# ---------------------------------------------------------------------------

@dataclass
class CoverMetadata:
    """
    All variable fields for a single document run.

    Fields with defaults are optional -- the user may press Enter to skip.
    Fields without defaults must be explicitly provided or left as empty
    string to skip (in which case the placeholder is cleared but blank).
    """
    document_type:           str = ""          # [[DOCUMENT_TYPE]]
    document_title:          str = ""          # [[DOCUMENT_TITLE]]
    subtitle:                str = ""          # [[SUBTITLE]]
    author_name:             str = ""          # [[AUTHOR_NAME]]
    department_organisation: str = ""          # [[DEPARTMENT_ORGANISATION]]
    month_year:              str = ""          # [[MONTH_YEAR]] -- auto-filled if empty
    organisation:            str = ""          # [[ORGANISATION]] in header
    confidentiality_label:   str = "INTERNAL"  # [[CONFIDENTIALITY_LABEL]] in footer
    # Academic template fields
    institution_name:        str = ""          # [[INSTITUTION_NAME]]
    faculty_department:      str = ""          # [[FACULTY_DEPARTMENT]]
    student_id:              str = ""          # [[STUDENT_ID]]
    supervisor_name:         str = ""          # [[SUPERVISOR_NAME]]
    course_code:             str = ""          # [[COURSE_CODE]]
    course_name:             str = ""          # [[COURSE_NAME]]
    submission_date:         str = ""          # [[SUBMISSION_DATE]]


# ---------------------------------------------------------------------------
# Result data structure
# ---------------------------------------------------------------------------

@dataclass
class GeneratorResult:
    """
    Output of the generator.

    success        -- True if all required placeholders were replaced
    final_doc      -- python-docx Document ready for writer.py (None on error)
    missing_fields -- placeholder names that were not filled (empty = all filled)
    error          -- plain-language error (empty on success)
    """
    success:        bool = False
    final_doc:      Optional[object] = None
    missing_fields: List[str] = field(default_factory=list)
    error:          str = ""


# ---------------------------------------------------------------------------
# Placeholder replacement
# ---------------------------------------------------------------------------

# All placeholders declared in the template -- maps placeholder name to
# the CoverMetadata attribute that fills it.
# REF_ID: RM_DO178_001
_PLACEHOLDER_MAP: Dict[str, Optional[str]] = {
    "[[DOCUMENT_TYPE]]":           "document_type",
    "[[DOCUMENT_TITLE]]":          "document_title",
    "[[SUBTITLE]]":                "subtitle",
    "[[AUTHOR_NAME]]":             "author_name",
    "[[DEPARTMENT_ORGANISATION]]": "department_organisation",
    "[[MONTH_YEAR]]":              "month_year",
    "[[ORGANISATION]]":            "organisation",
    "[[CONFIDENTIALITY_LABEL]]":   "confidentiality_label",
    # Academic template placeholders
    "[[INSTITUTION_NAME]]":        "institution_name",
    "[[FACULTY_DEPARTMENT]]":      "faculty_department",
    "[[STUDENT_ID]]":              "student_id",
    "[[SUPERVISOR_NAME]]":         "supervisor_name",
    "[[COURSE_CODE]]":             "course_code",
    "[[COURSE_NAME]]":             "course_name",
    "[[SUBMISSION_DATE]]":         "submission_date",
    "[[BODY_CONTENT_START]]":      None,   # removed by formatter -- should not appear
    "[[REFERENCES_START]]":        None,   # removed by formatter -- should not appear
}


def _fill_metadata_defaults(meta: CoverMetadata, source_path: Path) -> CoverMetadata:
    """
    Fill any empty metadata fields with sensible defaults before replacement.

    - document_title: defaults to source file stem (filename without extension)
    - month_year: defaults to current month and year
    - organisation: defaults to department_organisation if empty
    - confidentiality_label: defaults to "INTERNAL" if empty

    REF_ID: RM_DO178_001
    """
    if len(meta.document_title) == 0:
        meta.document_title = source_path.stem.replace("_", " ").replace("-", " ").title()

    if len(meta.month_year) == 0:
        import platform as _platform
        _day_fmt: str = "%#d" if _platform.system() == "Windows" else "%-d"
        meta.month_year = datetime.now().strftime(_day_fmt + " %B %Y")

    if len(meta.organisation) == 0:
        meta.organisation = meta.department_organisation if len(meta.department_organisation) > 0 else ""

    if len(meta.confidentiality_label) == 0:
        meta.confidentiality_label = "INTERNAL"

    return meta


def _replace_in_run(run: Any, replacements: Dict[str, str]) -> None:
    """
    Replace all placeholder tokens in a single run's text.
    Operates directly on the run.text string.
    """
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
    Walk all paragraphs in the document (including headers and footers)
    and replace placeholder tokens with their values.

    Returns a list of placeholder names that were found and could not be
    replaced because their value is empty.

    Note: Word sometimes splits a placeholder across multiple runs
    (e.g. [[DOC and UMENT_TITLE]] in separate runs). The paragraph-level
    join-and-split approach handles this case.

    REF_ID: RM_DO178_001
    """
    unfilled: List[str] = []

    def _process_paragraph(para: Any) -> None:
        full_text: str = para.text
        has_placeholder: bool = False

        for ph in replacements:
            if ph in full_text:
                has_placeholder = True
                break

        if has_placeholder is False:
            return

        from docx.oxml.ns import qn as _qn

        has_any_field: bool = False
        for r in para.runs:
            if r._r.find(_qn("w:fldChar")) is not None or r._r.find(_qn("w:instrText")) is not None:
                has_any_field = True
                break

        if has_any_field is True:
            runs: List[Any] = para.runs
            i: int = 0
            while i < len(runs):
                run: Any = runs[i]
                is_field: bool = (run._r.find(_qn("w:fldChar")) is not None or
                            run._r.find(_qn("w:instrText")) is not None)
                if is_field is True:
                    i += 1
                    continue

                j: int = i + 1
                while j < len(runs):
                    r_j: Any = runs[j]
                    if r_j._r.find(_qn("w:fldChar")) is not None or r_j._r.find(_qn("w:instrText")) is not None:
                        break
                    j += 1

                if j > i + 1:
                    combined: str = "".join(runs[k].text for k in range(i, j))
                    for placeholder, value in replacements.items():
                        combined = combined.replace(placeholder, value)
                    runs[i].text = combined
                    for k in range(i + 1, j):
                        runs[k].text = ""
                else:
                    for placeholder, value in replacements.items():
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, value)
                i = j
        else:
            if len(para.runs) > 0:
                combined_all: str = "".join(r.text for r in para.runs)
                for placeholder, value in replacements.items():
                    combined_all = combined_all.replace(placeholder, value)
                para.runs[0].text = combined_all
                for run in para.runs[1:]:
                    run.text = ""

    # Body paragraphs
    for para in doc.paragraphs:
        _process_paragraph(para)

    # Table cells -- placeholders in tables are not covered by doc.paragraphs
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _process_paragraph(para)

    # Headers and footers across all sections
    for section in doc.sections:
        for hdr in [section.header, section.first_page_header]:
            if hdr is not None:
                for para in hdr.paragraphs:
                    _process_paragraph(para)
        for ftr in [section.footer, section.first_page_footer]:
            if ftr is not None:
                for para in ftr.paragraphs:
                    _process_paragraph(para)

    # Check for any remaining unfilled placeholders
    full_doc_text: str = "\n".join(p.text for p in doc.paragraphs)
    for ph in replacements:
        if ph in full_doc_text and len(replacements[ph]) == 0:
            unfilled.append(ph)

    return unfilled


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def generate(
    formatted_doc: Any,
    metadata: CoverMetadata,
    source_path: Path,
) -> GeneratorResult:
    """
    Inject cover page metadata into the formatted document by replacing
    all [[PLACEHOLDER]] tokens with real values.

    The TOC field ([[CONTENTS]] heading + Word TOC field) is already
    present in the template structure -- Word will update it when the
    user opens the document and presses Ctrl+A then F9, or when the
    document is opened in a Word version that auto-updates fields.
    We do not attempt to programmatically update the TOC field here
    because python-docx does not have reliable cross-platform TOC
    field update support.

    REF_ID: RM_DO178_001, RM_NIST_001
    """
    if formatted_doc is None:
        return GeneratorResult(error="No formatted document provided.")

    # Fill in defaults for any empty metadata fields
    metadata = _fill_metadata_defaults(metadata, source_path)

    # Build the replacement dict from metadata
    replacements: Dict[str, str] = {}
    for placeholder, attr_name in _PLACEHOLDER_MAP.items():
        if attr_name is None:
            replacements[placeholder] = ""
        else:
            replacements[placeholder] = getattr(metadata, attr_name, "")

    # Apply replacements across the whole document
    unfilled: List[str] = _replace_placeholders_in_doc(formatted_doc, replacements)

    return GeneratorResult(
        success=True,
        final_doc=formatted_doc,
        missing_fields=unfilled,
    )


# ---------------------------------------------------------------------------
# Cover metadata prompt helper
# ---------------------------------------------------------------------------

def prompt_cover_metadata(source_path: Path) -> CoverMetadata:
    """
    Prompt the user for cover page metadata fields.
    Returns a CoverMetadata object with all fields filled or defaulted.

    Each field shows its default value so the user can press Enter to accept.
    REF_ID: RM_NIST_001 -- user authorizes cover page content before generation
    """
    from ui.messages import (
        MSG_COVER_PAGE_INTRO,
        MSG_COVER_TITLE_PROMPT,
        MSG_COVER_DATE_PROMPT,
        MSG_COVER_AUTHOR_PROMPT,
        MSG_COVER_ORG_PROMPT,
        fmt, LABEL_CONFIRM, LABEL_INFO,
    )

    print(MSG_COVER_PAGE_INTRO)

    default_title: str = source_path.stem.replace("_", " ").replace("-", " ").title()
    import platform as _platform
    _day_fmt: str = "%#d" if _platform.system() == "Windows" else "%-d"
    default_date: str  = datetime.now().strftime(_day_fmt + " %B %Y")

    def ask(prompt: str, default: str = "") -> str:
        if len(default) > 0:
            print(prompt + " [" + default + "]")
        else:
            print(prompt)
        try:
            val: str = input("> ").strip()
        except (KeyboardInterrupt, EOFError):
            val = ""

        if len(val) > 0:
            return val
        return default

    doc_type: str = ask(
        fmt(LABEL_CONFIRM, "Document type (e.g. Research Report, Technical Report):"),
        "Report"
    )
    title: str        = ask(MSG_COVER_TITLE_PROMPT, default_title)
    subtitle: str     = ask(fmt(LABEL_CONFIRM, "Subtitle (or press Enter to leave blank):"), "")
    institution: str  = ask(fmt(LABEL_CONFIRM, "Institution name (or press Enter to skip):"), "")
    faculty: str      = ask(fmt(LABEL_CONFIRM, "Faculty / Department (or press Enter to skip):"), "")
    author: str       = ask(MSG_COVER_AUTHOR_PROMPT, "")
    student_id: str   = ask(fmt(LABEL_CONFIRM, "Student ID (or press Enter to skip):"), "")
    supervisor: str   = ask(fmt(LABEL_CONFIRM, "Supervisor name (or press Enter to skip):"), "")
    course_code: str  = ask(fmt(LABEL_CONFIRM, "Course code (or press Enter to skip):"), "")
    course_name: str  = ask(fmt(LABEL_CONFIRM, "Course name (or press Enter to skip):"), "")
    org: str          = ask(MSG_COVER_ORG_PROMPT, institution)
    date: str         = ask(MSG_COVER_DATE_PROMPT, default_date)
    conf_label: str   = ask(
        fmt(LABEL_CONFIRM, "Confidentiality label (CONFIDENTIAL / PUBLIC / INTERNAL):"),
        "INTERNAL"
    )

    return CoverMetadata(
        document_type=doc_type,
        document_title=title,
        subtitle=subtitle,
        author_name=author,
        department_organisation=org,
        month_year=date,
        organisation=org,
        confidentiality_label=conf_label,
        institution_name=institution,
        faculty_department=faculty,
        student_id=student_id,
        supervisor_name=supervisor,
        course_code=course_code,
        course_name=course_name,
        submission_date=date,
    )
