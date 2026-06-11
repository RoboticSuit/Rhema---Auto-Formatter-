# Source Trace:
# File: single.py
# Knowledge Files: CodeSourceDB v3.4, SyntaxBiasDB v2.3, HumanSyntaxDB v1.2, mainProtocol v5.5
# REF_IDs: RM_DO178_001, RM_NIST_001, RM_HCI_002, WEB_PY_001, RM_HCI_004
# Logic: Orchestrates the single file formatting pipeline with explicitly typed variables, strict checks, and path sanitization.

"""
Rhema -- Auto Formatter
modes/single.py

Single responsibility: Orchestrates the full formatting pipeline for a
single input document. Coordinates detector, preview, formatter, generator,
and writer in the correct sequence. Handles user interaction for file path
input, template selection, output path, and cover metadata collection.

This module contains no business logic -- it sequences calls to the core
modules and surfaces results to the user via messages.py.

Governing standards:
  NIST RMF Step 6 HITL gate     -- REF_ID: RM_NIST_001
  DO-178C state traceability    -- REF_ID: RM_DO178_001
  ISO 9241-110:2020 interaction -- REF_ID: RM_HCI_002
  POSIX.1-2017 path sanitation  -- REF_ID: RM_HCI_004
"""

import sys
from pathlib import Path
from typing import Optional, List, Dict, Any

from core.detector import detect, DetectionResult, ImplicitElement
from ui.preview import run_preview, PreviewResult
from core.formatter import format_document, FormatterResult
from core.generator import generate, prompt_cover_metadata, GeneratorResult, CoverMetadata
from core.writer import write_document, WriterResult
from storage.template_store import (
    ensure_library_exists,
    list_templates,
    get_template_path,
    save_template,
    template_exists,
)
from config.constants import (
    SUPPORTED_INPUT_EXTENSIONS,
    DEFAULT_OUTPUT_DIR,
    PREVIEW_VALID_CONFIRM_SIGNALS,
    MENU_OPTION_EXIT,
    TEMPLATE_EXTENSION,
    ACCEPTED_TEMPLATE_EXTENSIONS
)
from ui.messages import (
    MSG_INPUT_PATH_SINGLE_PROMPT,
    MSG_OUTPUT_PATH_PROMPT,
    MSG_NO_TEMPLATES,
    MSG_SAVE_AFTER_FORMAT,
    MSG_SAVE_TEMPLATE_NAME_PROMPT,
    MSG_NEW_TEMPLATE_PATH_PROMPT,
    MSG_WRONG_TEMPLATE_EXTENSION,
    MSG_LIBRARY_SELECT_PROMPT,
    MSG_LIBRARY_SELECT_INVALID,
    LIBRARY_HEADER,
    LIBRARY_FOOTER_SEPARATOR,
    msg_library_entry,
    msg_library_count,
    msg_starting_single,
    msg_reading,
    msg_applying_template,
    MSG_ADDING_COVER,
    MSG_BUILDING_TOC,
    MSG_SAVING_DOC,
    msg_done_single,
    msg_file_not_found,
    msg_file_unreadable,
    msg_unsupported_format,
    msg_template_not_found,
    msg_template_missing_styles,
    msg_output_not_writable,
    msg_template_saved,
    msg_template_path_not_found,
    fmt,
    LABEL_INFO,
    LABEL_ERROR,
    LABEL_CONFIRM,
    SEPARATOR,
)

# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _prompt(message: str) -> str:
    """Print a prompt and return stripped input. Handles Ctrl+C."""
    print(message)
    try:
        return input("> ").strip()
    except (KeyboardInterrupt, EOFError):
        print()
        return ""


def _clean_path(raw: str) -> str:
    """
    Remove PowerShell and terminal path artifacts.
    Handles: & 'path', & "path", 'path', "path", plain path.
    Uses single-pair removal to avoid stripping apostrophes inside paths.
    REF_ID: RM_HCI_004
    """
    raw = raw.strip()

    # Remove PowerShell invocation prefix
    if raw.startswith("& ") is True:
        raw = raw[2:].strip()
    elif raw.startswith("&") is True:
        raw = raw[1:].strip()

    # Remove ONE matching outer quote pair only
    if len(raw) >= 2:
        if (raw[0] == '"' and raw[-1] == '"') or (raw[0] == "'" and raw[-1] == "'"):
            raw = raw[1:-1]

    # Un-escape PowerShell single-quote escaping: '' -> '
    raw = raw.replace("''", "'")
    return raw.strip()


def _get_input_path() -> Optional[Path]:
    """
    Prompt the user for the input document path.
    Returns a validated Path or None if the user cancels.
    REF_ID: RM_NIST_001
    """
    while True:
        raw: str = _prompt(MSG_INPUT_PATH_SINGLE_PROMPT)
        if len(raw) == 0:
            return None

        raw = _clean_path(raw)
        path: Path = Path(raw)

        if path.exists() is False:
            print(msg_file_not_found(str(path)))
            continue

        if path.suffix.lower() not in SUPPORTED_INPUT_EXTENSIONS:
            print(msg_unsupported_format(path.name))
            continue

        return path


def _get_output_dir() -> Path:
    """
    Prompt the user for the output directory.
    Returns DEFAULT_OUTPUT_DIR if the user presses Enter.
    """
    raw: str = _prompt(MSG_OUTPUT_PATH_PROMPT)
    if len(raw) == 0:
        return DEFAULT_OUTPUT_DIR

    raw = _clean_path(raw)
    out: Path = Path(raw)

    # Verify we can write to it
    try:
        out.mkdir(parents=True, exist_ok=True)
        return out
    except OSError:
        print(msg_output_not_writable(str(out)))
        print(fmt(LABEL_INFO, "Using default output folder instead."))
        return DEFAULT_OUTPUT_DIR


def _select_template() -> Optional[Path]:
    """
    Ask the user to select a template from the library or provide one.
    Returns the Path to the selected .tmpt file, or None if cancelled.
    REF_ID: RM_NIST_001
    """
    ensure_library_exists()
    entries: List[Dict[str, str]] = list_templates()

    print("\n" + SEPARATOR)
    print(fmt(LABEL_INFO, "Select a template."))

    if len(entries) > 0:
        print(LIBRARY_HEADER)
        for i, entry in enumerate(entries, start=1):
            display_name: str = entry.get("display_name", "Unknown")
            saved_date: str = entry.get("saved", "Unknown")
            print(msg_library_entry(i, display_name, saved_date))

        print(LIBRARY_FOOTER_SEPARATOR)
        print(msg_library_count(len(entries)))
        print(fmt(LABEL_CONFIRM, "Type the number to select, 'new' to use a different file, or press Enter to cancel."))
    else:
        print(MSG_NO_TEMPLATES)
        print(fmt(LABEL_CONFIRM, "Type 'new' to use a template file, or press Enter to cancel."))

    while True:
        raw: str = _prompt("").lower()

        if len(raw) == 0:
            return None

        if raw == "new":
            return _prompt_new_template()

        if raw.isdigit() is True and len(entries) > 0:
            idx: int = int(raw) - 1
            if 0 <= idx and idx < len(entries):
                display_name = entries[idx].get("display_name", "")
                path: Optional[Path] = get_template_path(display_name)

                if path is None:
                    print(msg_template_not_found(display_name))
                    continue
                return path

            print(MSG_LIBRARY_SELECT_INVALID)
            continue

        print(MSG_LIBRARY_SELECT_INVALID)


def _prompt_new_template() -> Optional[Path]:
    """
    Ask the user to provide a .tmpt file path directly.
    Returns the validated Path or None if cancelled.
    """
    raw: str = _prompt(MSG_NEW_TEMPLATE_PATH_PROMPT)
    if len(raw) == 0:
        return None

    raw = _clean_path(raw)
    path: Path = Path(raw)

    if path.exists() is False:
        print(msg_template_path_not_found(raw))
        return None

    if path.suffix.lower() not in ACCEPTED_TEMPLATE_EXTENSIONS:
        print(MSG_WRONG_TEMPLATE_EXTENSION)
        return None

    # Offer to save to library
    _offer_save_template(path)
    return path


def _offer_save_template(template_path: Path) -> None:
    """
    After using a drag-dropped template, offer to save it to the library.
    REF_ID: RM_NIST_001
    """
    print(MSG_SAVE_AFTER_FORMAT)
    raw: str = _prompt("").lower()

    if raw in PREVIEW_VALID_CONFIRM_SIGNALS:
        name_raw: str = _prompt(MSG_SAVE_TEMPLATE_NAME_PROMPT)
        if len(name_raw) > 0:
            ok: bool = save_template(template_path, name_raw)
            if ok is True:
                print(msg_template_saved(name_raw))
            else:
                print(fmt(LABEL_ERROR, "Rhema could not save the template. You can try again from Manage Templates."))


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def run_single_mode() -> None:
    """
    Run the single file format mode.
    Orchestrates the full pipeline: input -> detect -> preview -> format
    -> generate -> write -> success message.

    Called by main.py when the user selects option 1.
    REF_ID: RM_DO178_001 -- state machine with defined transitions
    """
    # -----------------------------------------------------------------------
    # STATE 1: Get input file
    # -----------------------------------------------------------------------
    source_path: Optional[Path] = _get_input_path()
    if source_path is None:
        print(fmt(LABEL_INFO, "No file selected. Returning to main menu."))
        return

    print(msg_starting_single(source_path.name))

    # -----------------------------------------------------------------------
    # STATE 2: Select template
    # -----------------------------------------------------------------------
    template_path: Optional[Path] = _select_template()
    if template_path is None:
        print(fmt(LABEL_INFO, "No template selected. Returning to main menu."))
        return

    # -----------------------------------------------------------------------
    # STATE 3: Get output directory
    # -----------------------------------------------------------------------
    output_dir: Path = _get_output_dir()

    # -----------------------------------------------------------------------
    # STATE 4: Detect structure
    # REF_ID: RM_DO178_001
    # -----------------------------------------------------------------------
    print(msg_reading(source_path.name))
    detection: DetectionResult = detect(source_path)

    if len(detection.error) > 0:
        print(fmt(LABEL_ERROR, detection.error))
        return

    # -----------------------------------------------------------------------
    # STATE 5: Preview Mode (HITL gate -- only if implicit structure found)
    # REF_ID: RM_NIST_001
    # -----------------------------------------------------------------------
    confirmed_elements: List[ImplicitElement] = []

    if detection.needs_preview is True:
        preview_result: PreviewResult = run_preview(source_path.name, detection.implicit)
        if preview_result.cancelled is True:
            # msg_preview_cancelled or msg_preview_exhausted already printed
            # by preview.py -- just return cleanly
            return
        confirmed_elements = preview_result.confirmed_elements
    else:
        print(fmt(LABEL_INFO, "Document structure is fully declared. No review needed."))

    # -----------------------------------------------------------------------
    # STATE 6: Collect cover page metadata
    # REF_ID: RM_NIST_001
    # -----------------------------------------------------------------------
    metadata: CoverMetadata = prompt_cover_metadata(source_path)

    # -----------------------------------------------------------------------
    # STATE 7: Apply template formatting
    # -----------------------------------------------------------------------
    print(msg_applying_template(template_path.stem))
    fmt_result: FormatterResult = format_document(source_path, template_path, confirmed_elements)

    if fmt_result.success is False:
        print(fmt(LABEL_ERROR, fmt_result.error))
        return

    # Diagnostic check: ensure formatting actually occurred
    if getattr(fmt_result, "formatted_doc", None) is None:
        print(fmt(LABEL_ERROR, "Formatter error: No document content generated."))
        return

    # Surface any unmapped styles to the user before writing
    unmapped_styles: List[str] = getattr(fmt_result, "unmapped_styles", [])
    if len(unmapped_styles) > 0:
        print(fmt(LABEL_INFO,
            "Some styles in your document were not found in the template "
            "and have been set to Normal: " + ", ".join(unmapped_styles)))

    # -----------------------------------------------------------------------
    # STATE 8: Generate cover page and TOC
    # -----------------------------------------------------------------------
    print(MSG_ADDING_COVER)
    print(MSG_BUILDING_TOC)
    gen_result: GeneratorResult = generate(fmt_result.formatted_doc, metadata, source_path)

    if gen_result.success is False:
        print(fmt(LABEL_ERROR, gen_result.error))
        return

    if len(gen_result.missing_fields) > 0:
        print(fmt(LABEL_INFO, "Some cover page fields were left blank: " + ", ".join(gen_result.missing_fields)))

    # -----------------------------------------------------------------------
    # STATE 9: Write output file
    # -----------------------------------------------------------------------
    print(MSG_SAVING_DOC)
    # Use the document title as the output filename so the file is named
    # after the document, not the source file. REF_ID: RM_DO178_001
    output_filename: Optional[str] = metadata.document_title if len(metadata.document_title) > 0 else None
    write_result: WriterResult = write_document(gen_result.final_doc, source_path, output_dir, output_filename)

    if write_result.success is False:
        print(fmt(LABEL_ERROR, write_result.error))
        return

    if write_result.toc_updated is True:
        print(fmt(LABEL_INFO, "Table of contents updated automatically."))
    else:
        print(fmt(LABEL_INFO, "Open the output file in Word and press Ctrl+A then F9 to update the table of contents."))

    if write_result.output_path is not None:
        print(msg_done_single(str(write_result.output_path)))
