# Source Trace:
# File: writer_v0.7.py
# Knowledge Files: CodeSourceDB v3.6, SyntaxBiasDB v2.3, HumanSyntaxDB v1.2
# REF_IDs: RM_DO178_001, RM_ISO_001, RM_HCI_004, WEB_PY_001, WEB_PY_006,
#          WEB_PY_011, WEB_PY_015
# Logic: v0.7 -- para.Style.Name -> para.Style.NameLocal throughout
#        _apply_professional_formatting. Fixes stale win32com gen_py
#        cache issue where .Name raises AttributeError on every paragraph.
#        v0.6 -- Debug logging added to _finalize_document_via_word and
#        _apply_professional_formatting. Writes rhema_debug.log to project
#        root on every format run. No logic changes from v0.5.

"""
Rhema -- Auto Formatter
core/writer.py

Single responsibility: Writes the final python-docx Document object to
disk as a .docx file. Handles output path resolution, filename collision
avoidance, partial-file cleanup on write failure, OOXML post-processing,
and document finalisation via the Word COM API on Windows.

v0.6 adds debug logging to the COM session. Every significant step writes
a timestamped line to rhema_debug.log in the project root. Check this file
after a format run to see exactly where the COM session succeeded or failed.

Governing standards:
  DO-178C source traceability -- REF_ID: RM_DO178_001
  ISO 31000 risk treatment    -- REF_ID: RM_ISO_001
  POSIX.1-2017 exit codes     -- REF_ID: RM_HCI_004
  Python COM automation       -- REF_ID: WEB_PY_011
"""

import os
import logging
from pathlib import Path
from dataclasses import dataclass
from typing import Optional, Any

from config.constants import (
    DEFAULT_OUTPUT_DIR,
    OUTPUT_EXTENSION,
)

# ---------------------------------------------------------------------------
# Debug logger -- writes to rhema_debug.log in the project root.
# Logs are appended so multiple runs accumulate. REF_ID: WEB_PY_013
# ---------------------------------------------------------------------------

_LOG_PATH: Path = Path(__file__).resolve().parent.parent / "rhema_debug.log"

def _get_logger() -> logging.Logger:
    logger = logging.getLogger("rhema.writer")
    if not logger.handlers:
        handler = logging.FileHandler(str(_LOG_PATH), encoding="utf-8")
        handler.setFormatter(logging.Formatter(
            "%(asctime)s  %(levelname)-7s  %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        ))
        logger.addHandler(handler)
        logger.setLevel(logging.DEBUG)
    return logger


# ---------------------------------------------------------------------------
# Result data structure
# ---------------------------------------------------------------------------

@dataclass
class WriterResult:
    success:     bool           = False
    output_path: Optional[Path] = None
    error:       str            = ""
    toc_updated: bool           = False


# ---------------------------------------------------------------------------
# Output path resolver
# ---------------------------------------------------------------------------

def _resolve_output_path(
    source_path: Path,
    output_dir: Path,
    custom_filename: Optional[str] = None,
) -> Path:
    if custom_filename is not None and len(custom_filename) > 0:
        stem: str = Path(custom_filename).stem
    else:
        stem = source_path.stem

    output_dir.mkdir(parents=True, exist_ok=True)

    candidate: Path = output_dir / (stem + OUTPUT_EXTENSION)
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
    if final_doc is None:
        return WriterResult(error="No document provided for writing.")

    if output_dir is None:
        output_dir = DEFAULT_OUTPUT_DIR

    output_path: Path = _resolve_output_path(source_path, output_dir, custom_filename)

    file_created: bool = False
    try:
        final_doc.save(str(output_path))
        file_created = True

        if output_path.exists() is False or output_path.stat().st_size == 0:
            raise IOError("File was not written or is empty.")

        _post_process_docx(output_path)
        toc_updated: bool = _finalize_document_via_word(output_path)

        return WriterResult(
            success=True,
            output_path=output_path,
            toc_updated=toc_updated,
        )

    except PermissionError:
        _cleanup(output_path, file_created)
        return WriterResult(
            error="Rhema could not save to: " + str(output_path) +
                  "\n        Check that you have permission to save files in "
                  "that folder, then try again."
        )
    except OSError as e:
        _cleanup(output_path, file_created)
        if "No space" in str(e) or "disk" in str(e).lower():
            return WriterResult(
                error="Rhema could not save the document. Your disk may be "
                      "full. Free up some space and try again."
            )
        return WriterResult(
            error="Rhema could not save: " + str(output_path) +
                  "\n        Please try again."
        )
    except Exception:
        _cleanup(output_path, file_created)
        return WriterResult(
            error="Something went wrong while saving " + source_path.name +
                  ". The incomplete file has been removed. Please try again."
        )


# ---------------------------------------------------------------------------
# OOXML post-processing
# ---------------------------------------------------------------------------

def _post_process_docx(path: Path) -> None:
    import zipfile as _zf
    import re as _re

    try:
        with _zf.ZipFile(str(path), "r") as z:
            files: dict = {n: z.read(n) for n in z.namelist()}

        doc_xml: str = files["word/document.xml"].decode("utf-8")

        doc_xml = doc_xml.replace(
            '<w:fldChar w:fldCharType="begin"/>',
            '<w:fldChar w:fldCharType="begin" w:dirty="true"/>',
        )

        # Fix TOC field instruction.
        # \h adds hyperlink navigation so Ctrl+Click works on every TOC entry.
        # \o "1-3" uses outline-level attributes to detect headings -- reliable
        # regardless of style display name spacing or casing in the template.
        # The previous \t switch referenced "Heading1,1,Heading2,2" (no spaces)
        # which never matched Word's actual style names "Heading 1", "Heading 2",
        # so it silently did nothing. Dropped in favour of \o alone.
        # Lambda replacement used instead of r-string to avoid Python 3.12
        # re.sub treating \h as an invalid backreference. REF_ID: RM_DO178_001
        _TOC_INSTR: str = ' TOC \\h \\o "1-3" '
        doc_xml = _re.sub(
            r'(<w:instrText[^>]*>)[^<]*TOC[^<]*(</w:instrText>)',
            lambda m: m.group(1) + _TOC_INSTR + m.group(2),
            doc_xml,
        )

        import re as _re2
        doc_xml = _re2.sub(
            r'(<w:fldChar w:fldCharType="separate"/>)(.*?)'
            r'(<w:fldChar w:fldCharType="end"/>)',
            r'\1<w:r><w:t> </w:t></w:r>\3',
            doc_xml,
            flags=_re2.DOTALL,
        )

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
        pass


# ---------------------------------------------------------------------------
# Professional formatting rules -- with full debug logging
# ---------------------------------------------------------------------------

def _apply_professional_formatting(doc: Any, log: logging.Logger) -> None:
    """
    Apply all four professional formatting rules via paragraph properties.
    Every significant action is logged so failures are immediately visible.
    REF_ID: WEB_PY_011, RM_DO178_001
    """
    WD_LIST_NONE: int = 0
    MAX_SCAN:     int = 600

    total: int = doc.Paragraphs.Count
    log.debug(f"  Total paragraphs in document: {total}")

    # -----------------------------------------------------------------------
    # Locate body section
    # -----------------------------------------------------------------------
    body_start: int = 1
    body_end:   int = total

    for i in range(1, min(total + 1, MAX_SCAN)):
        try:
            sname = doc.Paragraphs(i).Style.NameLocal
            txt   = doc.Paragraphs(i).Range.Text.strip()
            if sname.startswith("Heading 1") and body_start == 1:
                body_start = i
                log.debug(f"  Body start: para {i} ({sname}: '{txt[:40]}')")
            if sname == "Title" and txt in ("REFERENCES", "CONTENTS"):
                body_end = i - 1
                log.debug(f"  Body end:   para {body_end} (before '{txt}' at {i})")
                if txt == "REFERENCES":
                    break
        except Exception as e:
            log.debug(f"  Body scan error at para {i}: {e}")

    log.debug(f"  Body section: paras {body_start} to {body_end}")

    # -----------------------------------------------------------------------
    # Forward pass: Rules 1, 2, 4
    # -----------------------------------------------------------------------
    kn_ok = kn_fail = kt_ok = kt_fail = list_groups = 0

    i: int = body_start
    while i <= body_end:
        try:
            para      = doc.Paragraphs(i)
            sname     = para.Style.NameLocal
            is_heading: bool = sname.startswith("Heading")

            try:
                list_id: int  = para.Range.ListFormat.ListId
                is_list: bool = (list_id != WD_LIST_NONE)
            except Exception as e:
                log.debug(f"  Para {i}: ListId error: {e}")
                is_list  = False
                list_id  = 0

            # Rule 4: KeepWithNext on headings
            if is_heading:
                try:
                    before = para.Format.KeepWithNext
                    para.Format.KeepWithNext = True
                    after  = para.Format.KeepWithNext
                    log.debug(
                        f"  Para {i} [{sname}]: KeepWithNext "
                        f"{before} -> {after}  '{para.Range.Text.strip()[:40]}'"
                    )
                    if after:
                        kn_ok += 1
                    else:
                        log.warning(
                            f"  Para {i}: KeepWithNext SET but read back False!"
                        )
                        kn_fail += 1
                except Exception as e:
                    log.warning(f"  Para {i} [{sname}]: KeepWithNext FAILED: {e}")
                    kn_fail += 1
                i += 1
                continue

            # Rule 2: List group
            if is_list:
                group_start: int = i
                j: int = i + 1
                while j <= body_end:
                    try:
                        np_lid = doc.Paragraphs(j).Range.ListFormat.ListId
                        if np_lid == WD_LIST_NONE:
                            break
                        j += 1
                    except Exception:
                        break

                group_end: int = j - 1
                list_groups += 1
                log.debug(
                    f"  List group {list_groups}: paras {group_start}-{group_end} "
                    f"({group_end - group_start + 1} items)"
                )

                for k in range(group_start, group_end):
                    try:
                        doc.Paragraphs(k).Format.KeepWithNext = True
                        kn_ok += 1
                    except Exception as e:
                        log.warning(f"    Para {k}: list KeepWithNext FAILED: {e}")
                        kn_fail += 1

                try:
                    doc.Paragraphs(group_end).Format.KeepWithNext = False
                except Exception as e:
                    log.warning(f"    Para {group_end}: list last-item clear FAILED: {e}")

                i = j
                continue

            # Rule 1: KeepTogether on body paragraphs
            try:
                txt = para.Range.Text.strip()
                if len(txt) > 0:
                    before = para.Format.KeepTogether
                    para.Format.KeepTogether = True
                    after  = para.Format.KeepTogether
                    if after:
                        kt_ok += 1
                    else:
                        log.warning(
                            f"  Para {i}: KeepTogether SET but read back False!"
                        )
                        kt_fail += 1
            except Exception as e:
                log.warning(f"  Para {i}: KeepTogether FAILED: {e}")
                kt_fail += 1

        except Exception as e:
            log.warning(f"  Para {i}: outer loop error: {e}")

        i += 1

    log.debug(
        f"  Pass complete -- KeepWithNext OK:{kn_ok} FAIL:{kn_fail} | "
        f"KeepTogether OK:{kt_ok} FAIL:{kt_fail} | "
        f"ListGroups:{list_groups}"
    )

    # -----------------------------------------------------------------------
    # Rule 3: Table row integrity
    # -----------------------------------------------------------------------
    try:
        body_start_char: int = doc.Paragraphs(body_start).Range.Start
        body_end_char:   int = doc.Paragraphs(body_end).Range.End
        rows_ok = rows_fail = 0

        for t in range(1, doc.Tables.Count + 1):
            try:
                tbl = doc.Tables(t)
                if (tbl.Range.Start < body_start_char or
                        tbl.Range.Start > body_end_char):
                    continue
                log.debug(
                    f"  Table {t}: {tbl.Rows.Count} rows "
                    f"(char {tbl.Range.Start})"
                )
                for r in range(1, tbl.Rows.Count + 1):
                    try:
                        tbl.Rows(r).AllowBreakAcrossPages = False
                        rows_ok += 1
                    except Exception as e:
                        log.warning(
                            f"    Table {t} row {r}: AllowBreakAcrossPages FAILED: {e}"
                        )
                        rows_fail += 1
            except Exception as e:
                log.warning(f"  Table {t}: error: {e}")

        log.debug(f"  Tables done -- rows OK:{rows_ok} FAIL:{rows_fail}")

    except Exception as e:
        log.warning(f"  Table pass error: {e}")


# ---------------------------------------------------------------------------
# Word COM finalisation -- with full debug logging
# ---------------------------------------------------------------------------

def _finalize_document_via_word(path: Path) -> bool:
    log = _get_logger()
    log.info("=" * 60)
    log.info(f"COM session START: {path.name}")

    import platform
    log.debug(f"  Platform: {platform.system()}")

    if platform.system() != "Windows":
        log.info("  Non-Windows platform -- COM session skipped")
        return False

    try:
        import win32com.client as _win32
        import pythoncom as _pycom
        log.debug("  pywin32 imported OK")
    except ImportError as e:
        log.error(f"  pywin32 import FAILED: {e}")
        return False

    try:
        _pycom.CoInitialize()
        log.debug("  CoInitialize OK")

        word: Any = _win32.Dispatch("Word.Application")
        word.Visible = False
        word.Options.UpdateLinksAtOpen  = False
        word.Options.ConfirmConversions = False
        word.DisplayAlerts              = 0
        log.debug("  Word.Application dispatched OK")

        doc: Any = word.Documents.Open(
            str(path.resolve()),
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
        )
        log.debug(f"  Document opened OK: {path.resolve()}")

        # Step 1: Apply professional formatting rules
        log.info("  Applying professional formatting rules ...")
        _apply_professional_formatting(doc, log)
        log.info("  Formatting rules applied")

        # Step 2: Full TOC rebuild
        toc_count: int = doc.TablesOfContents.Count
        log.debug(f"  TOC tables found: {toc_count}")
        for toc in doc.TablesOfContents:
            toc.Update()
            log.debug("  toc.Update() called")

        # Step 3: Update remaining fields
        doc.Fields.Update()
        log.debug("  Fields.Update() called")

        doc.Save()
        log.debug("  doc.Save() OK")

        doc.Close(SaveChanges=True)
        word.Quit()
        _pycom.CoUninitialize()

        log.info("  COM session COMPLETE -- toc_updated=True")
        return True

    except Exception as e:
        log.error(f"  COM session FAILED with exception: {type(e).__name__}: {e}")
        try:
            _pycom.CoUninitialize()
        except Exception:
            pass
        return False


# ---------------------------------------------------------------------------
# Cleanup helper
# ---------------------------------------------------------------------------

def _cleanup(path: Path, file_created: bool) -> None:
    if file_created is True and path.exists() is True:
        try:
            path.unlink()
        except OSError:
            pass
