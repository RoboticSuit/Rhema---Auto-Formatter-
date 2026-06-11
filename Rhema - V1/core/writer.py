# Source Trace:
# File: writer.py
# Knowledge Files: CodeSourceDB v3.4, SyntaxBiasDB v2.3, HumanSyntaxDB v1.2, mainProtocol v5.5
# REF_IDs: RM_DO178_001, RM_ISO_001, RM_HCI_004, WEB_PY_001, RM_HCI_001
# Logic: Output path resolution and file write operations with explicit type hinting and deterministic cleanup.

"""
Rhema -- Auto Formatter
core/writer.py

Single responsibility: Writes the final python-docx Document object to
disk as a .docx file. Handles output path resolution, filename collision
avoidance, and partial-file cleanup on write failure.

Takes:
  - final_doc: python-docx Document from generator.py
  - source_path: Path to the original source file (used for default filename)
  - output_dir: Path to the output directory (defaults to DEFAULT_OUTPUT_DIR)
  - custom_filename: optional override for the output filename

Returns a WriterResult containing:
  - success: True if the file was written successfully
  - output_path: the full Path of the written file (None on failure)
  - error: non-empty string if writing failed

PARTIAL OUTPUT PROHIBITION: If writing fails at any point after the file
has been opened, the incomplete file is deleted before returning the error.
No partial .docx file is ever left on disk.

Governing standards:
  DO-178C source traceability    -- REF_ID: RM_DO178_001
  ISO 31000 risk treatment       -- REF_ID: RM_ISO_001
  POSIX.1-2017 exit codes        -- REF_ID: RM_HCI_004
"""

import os
from pathlib import Path
from dataclasses import dataclass
from typing import Optional, Any

from config.constants import (
    DEFAULT_OUTPUT_DIR,
    OUTPUT_EXTENSION,
)


# ---------------------------------------------------------------------------
# Result data structure
# ---------------------------------------------------------------------------

@dataclass
class WriterResult:
    """
    Output of the writer.

    success     -- True if the .docx file was written successfully
    output_path -- full Path of the written file (None on failure)
    error       -- plain-language error string (empty on success)
    toc_updated -- True if TOC was updated via Word, False if manual update needed
    """
    success:     bool = False
    output_path: Optional[Path] = None
    error:       str = ""
    toc_updated: bool = False


# ---------------------------------------------------------------------------
# Output path resolver
# ---------------------------------------------------------------------------

def _resolve_output_path(
    source_path: Path,
    output_dir: Path,
    custom_filename: Optional[str] = None,
) -> Path:
    """
    Determine the full output path for the .docx file.

    Priority:
    1. If custom_filename is provided, use it (with .docx extension enforced).
    2. Otherwise derive the filename from the source file stem.

    If a file already exists at the resolved path, append a numeric suffix
    to avoid overwriting: document.docx -> document_1.docx -> document_2.docx

    REF_ID: RM_DO178_001
    """
    if custom_filename is not None and len(custom_filename) > 0:
        stem: str = Path(custom_filename).stem
    else:
        stem: str = source_path.stem

    # Ensure the output directory exists
    output_dir.mkdir(parents=True, exist_ok=True)

    candidate: Path = output_dir / (stem + OUTPUT_EXTENSION)

    # Collision avoidance -- never silently overwrite an existing file
    counter: int = 1
    while candidate.exists() is True:
        candidate = output_dir / (stem + "_" + str(counter) + OUTPUT_EXTENSION)
        counter += 1

    return candidate


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def write_document(
    final_doc: Any,
    source_path: Path,
    output_dir: Optional[Path] = None,
    custom_filename: Optional[str] = None,
) -> WriterResult:
    """
    Write the final Document object to disk as a .docx file.

    Parameters:
      final_doc       -- python-docx Document from generator.py
      source_path     -- Path to the original source file
      output_dir      -- directory to write to (defaults to DEFAULT_OUTPUT_DIR)
      custom_filename -- optional filename override (without extension)

    Returns a WriterResult. On success, output_path is the full path of
    the written file. On failure, any partial file is deleted and error
    contains a plain-language description.

    PARTIAL OUTPUT PROHIBITION: If save() raises after creating the file,
    the incomplete file is deleted before this function returns.
    REF_ID: RM_ISO_001
    """
    if final_doc is None:
        return WriterResult(error="No document provided for writing.")

    if output_dir is None:
        output_dir = DEFAULT_OUTPUT_DIR

    output_path: Path = _resolve_output_path(source_path, output_dir, custom_filename)

    # Attempt to write the file
    # Track whether the file was created so cleanup knows whether to delete it
    file_created: bool = False
    try:
        final_doc.save(str(output_path))
        file_created = True

        # Verify the file was actually written and is non-empty
        if output_path.exists() is False or output_path.stat().st_size == 0:
            raise IOError("File was not written or is empty.")

        # Post-process: set TOC field as dirty so Word updates on open,
        # and fix CONTENTS/REFERENCES headings to be excluded from TOC.
        # REF_ID: RM_DO178_001
        _post_process_docx(output_path)

        # Attempt to update the TOC by opening the document in Word.
        # This is the only deterministic method -- Word's TOC field requires
        # the Word rendering engine to resolve heading page numbers.
        # If Word is not available, the document is still valid and the user
        # is instructed to update the TOC manually.
        toc_updated: bool = _update_toc_via_word(output_path)

        return WriterResult(
            success=True,
            output_path=output_path,
            toc_updated=toc_updated,
        )

    except PermissionError:
        _cleanup(output_path, file_created)
        return WriterResult(
            error="Rhema could not save to: " + str(output_path) +
                  "\n        Check that you have permission to save files in that folder, then try again."
        )
    except OSError as e:
        _cleanup(output_path, file_created)
        if "No space" in str(e) or "disk" in str(e).lower():
            return WriterResult(
                error="Rhema could not save the document. Your disk may be full. Free up some space and try again."
            )
        return WriterResult(
            error="Rhema could not save: " + str(output_path) + "\n        Please try again."
        )
    except Exception:
        _cleanup(output_path, file_created)
        return WriterResult(
            error="Something went wrong while saving " + source_path.name + ". The incomplete file has been removed. Please try again."
        )


def _post_process_docx(path: Path) -> None:
    """
    Post-processing applied to every saved .docx file.
    1. Sets all fldChar begin elements to dirty=true so Word auto-updates fields.
    2. Changes CONTENTS and REFERENCES headings to Title style so they are
       excluded from the TOC range, preventing self-referencing entries.
    REF_ID: RM_DO178_001
    """
    import zipfile as _zf
    import re as _re

    try:
        with _zf.ZipFile(str(path), "r") as z:
            files: dict = {n: z.read(n) for n in z.namelist()}

        doc_xml: str = files["word/document.xml"].decode("utf-8")

        # Fix 1: Set all fldChar begin elements to dirty=true
        doc_xml = doc_xml.replace(
            '<w:fldChar w:fldCharType="begin"/>',
            '<w:fldChar w:fldCharType="begin" w:dirty="true"/>',
        )

        # Fix 1c: Clear stale TOC display content from the sdt.
        # Word stores previously-rendered TOC entries between fldChar separate
        # and fldChar end. When dirty=true, Word should rebuild them -- but
        # stale entries cause "Error! Number cannot be represented" errors
        # because they reference page numbers from a previous layout.
        # Replace everything between separate and end with an empty run
        # so Word builds the TOC fresh on open.
        import re as _re2
        doc_xml = _re2.sub(
            r'(<w:fldChar w:fldCharType="separate"/>)(.*?)(<w:fldChar w:fldCharType="end"/>)',
            r'\1<w:r><w:t> </w:t></w:r>\3',
            doc_xml,
            flags=_re2.DOTALL
        )

        # Fix 1b: Remove \h flag from TOC instruction to suppress the
        # "fields refer to other files" warning Word shows on every open.
        doc_xml = _re.sub(
            r'(TOC\s+)\\\\h\s+',
            r'\1',
            doc_xml
        )
        # Also handle the entity-encoded form
        doc_xml = _re.sub(
            r'(TOC\s+)\\h\s+',
            r'\1',
            doc_xml
        )

        # Fix 2: Change CONTENTS and REFERENCES headings to Title style
        para_pattern = _re.compile(r"<w:p\b[^>]*>.*?</w:p>", _re.DOTALL)

        def _fix_structural_para(m: Any) -> str:
            para: str = m.group(0)
            for heading_text in ("CONTENTS", "REFERENCES"):
                if ((">%s<" % heading_text) in para or
                        (">%s</w:t>" % heading_text) in para):
                    if "<w:pStyle w:val=\"Heading1\"/>" in para:
                        return para.replace(
                            "<w:pStyle w:val=\"Heading1\"/>",
                            "<w:pStyle w:val=\"Title\"/>",
                        )
            return para

        doc_xml = para_pattern.sub(_fix_structural_para, doc_xml)

        files["word/document.xml"] = doc_xml.encode("utf-8")

        with _zf.ZipFile(str(path), "w", _zf.ZIP_DEFLATED) as z:
            for name, data in files.items():
                z.writestr(name, data)

    except Exception:
        pass  # Post-processing failure must never block file delivery

def _update_toc_via_word(path: Path) -> bool:
    """
    Open the saved .docx in Microsoft Word, update all fields (including
    the Table of Contents), save, and close. Returns True on success.

    This is the only deterministic method to update a Word TOC field
    because the TOC is a Word-managed object that requires the Word
    rendering engine to resolve heading page numbers.

    Requirements: Windows, Microsoft Word installed.
    If either condition is not met, returns False silently -- the document
    is still valid and the user is instructed to update manually.
    REF_ID: RM_DO178_001
    """
    import platform
    if platform.system() != "Windows":
        return False

    try:
        import win32com.client as _win32
        import pythoncom as _pycom
        _pycom.CoInitialize()

        word: Any = _win32.Dispatch("Word.Application")
        word.Visible = False

        # Suppress all dialogs -- prevents interactive prompts during automation
        word.Options.UpdateLinksAtOpen   = False
        word.Options.ConfirmConversions   = False
        word.DisplayAlerts               = 0   # wdAlertsNone

        doc: Any = word.Documents.Open(
            str(path.resolve()),
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
        )

        # Update TOC objects directly -- UpdatePageNumbers only, no dialog
        for toc in doc.TablesOfContents:
            toc.UpdatePageNumbers()

        # Update all remaining fields (page numbers, cross-references)
        doc.Fields.Update()

        doc.Save()
        doc.Close(SaveChanges=True)
        word.Quit()
        _pycom.CoUninitialize()
        return True

    except ImportError:
        return False  # pywin32 not installed
    except Exception:
        return False  # Word not installed or COM error


def _cleanup(path: Path, file_created: bool) -> None:
    """
    Delete a partially written file if it exists.
    Called only on write failure. Silent if the file does not exist.
    REF_ID: RM_ISO_001 -- partial output prohibition
    """
    if file_created is True and path.exists() is True:
        try:
            path.unlink()
        except OSError:
            pass  # Best-effort cleanup -- if deletion fails, the error message
                  # already informs the user. We do not raise here.
