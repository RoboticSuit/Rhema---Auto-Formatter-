# Source Trace:
# File: constants_v0.3.py
# Knowledge Files: CodeSourceDB v3.6, SyntaxBiasDB v2.3, HumanSyntaxDB v1.2
# REF_IDs: RM_DO178_001, RM_HCI_004, RM_NIST_001, WEB_PY_010
# Logic: Centralized declaration of application constants, POSIX exit codes, explicit UI
#        signals, and OOXML unit conversion values. No logic lives here.
#        v0.3 -- Added MENU_OPTION_SAMPLE (option 4). Exit renumbered to 5.
#                Added SAMPLE_OUTPUT_FILENAME for the generated format sample.

"""
Rhema -- Auto Formatter
config/constants.py

Single responsibility: Declares all named constants used across the application.
No logic lives here. Every numeric value, default path, and fixed string that
appears more than once in the codebase is declared here and imported by name.

Governing standards:
  POSIX.1-2017 exit code semantics -- REF_ID: RM_HCI_004
  DO-178C no magic numbers         -- REF_ID: RM_DO178_001
  SyntaxBiasDB v2.3 named constants rule
"""

import os
from pathlib import Path

# ---------------------------------------------------------------------------
# Application identity
# ---------------------------------------------------------------------------

APP_NAME: str    = "Rhema"
APP_VERSION: str = "1.0.0"
APP_TAGLINE: str = "Document Formatter"

# ---------------------------------------------------------------------------
# Exit codes -- all declared as named constants per POSIX.1-2017.
# Exit code 0: success. Non-zero: failure. Each code has a unique meaning.
# REF_ID: RM_HCI_004
# ---------------------------------------------------------------------------

EXIT_SUCCESS: int       = 0   # All operations completed successfully
EXIT_ERROR: int         = 1   # General error -- operation could not complete
EXIT_INPUT_ERROR: int   = 2   # Input file missing, unreadable, or unsupported format
EXIT_TEMPLATE_ERROR: int = 3  # Template not found, unreadable, or missing required styles
EXIT_OUTPUT_ERROR: int  = 4   # Cannot write to declared output path
EXIT_USER_CANCEL: int   = 5   # User cancelled the operation in Preview Mode

# ---------------------------------------------------------------------------
# File extensions
# ---------------------------------------------------------------------------

# Supported input document formats
SUPPORTED_INPUT_EXTENSIONS: set[str] = {".txt", ".docx", ".doc"}

# Template file extension -- used when saving to library
TEMPLATE_EXTENSION: str = ".tmpt"

# Accepted input extensions for template files
# .tmpt = Rhema template, .dotx = Word template, .docx = Word document
ACCEPTED_TEMPLATE_EXTENSIONS: set[str] = {".tmpt", ".dotx", ".docx"}

# Output document extension
OUTPUT_EXTENSION: str = ".docx"

# Default filename for the generated format sample document.
# Collision avoidance is applied at write time if this name already exists.
# REF_ID: RM_DO178_001
SAMPLE_OUTPUT_FILENAME: str = "Rhema_Format_Sample"

# ---------------------------------------------------------------------------
# Default paths -- resolved relative to the user's home directory
# so the application is portable across machines.
# REF_ID: RM_DO178_001, WEB_PY_005
# ---------------------------------------------------------------------------

# Base directory for all Rhema data stored on the user's machine
_RHEMA_HOME: Path = Path.home() / ".rhema"

# Persistent template library -- where saved .tmpt files are stored
DEFAULT_TEMPLATE_DIR: Path = _RHEMA_HOME / "templates"

# Default output directory -- where formatted .docx files are written
DEFAULT_OUTPUT_DIR: Path = Path.cwd() / "output"

# ---------------------------------------------------------------------------
# Preview Mode limits
# Maximum number of adjustment cycles before escalation to skip.
# Three cycles gives the user two correction attempts after the initial
# detection display. Beyond this, automatic detection cannot resolve the
# document structure. REF_ID: RM_NIST_001, RM_MISRA_001
# ---------------------------------------------------------------------------

PREVIEW_MAX_ADJUSTMENT_CYCLES: int = 3

# ---------------------------------------------------------------------------
# Batch mode
# ---------------------------------------------------------------------------

# Minimum number of files to trigger batch mode display
BATCH_MIN_FILES: int = 2

# ---------------------------------------------------------------------------
# Terminal display -- REF_ID: RM_HCI_004, RM_VIS_001
# ---------------------------------------------------------------------------

# Column width for the template library listing separator line
LIBRARY_DISPLAY_WIDTH: int = 52

# Name column width in the template library listing table.
# Declared here to avoid a magic number in messages.py.
# REF_ID: RM_DO178_001
LIBRARY_NAME_COLUMN_WIDTH: int = 30

# Maximum line width for terminal output -- POSIX.1-2017 80-column convention
# REF_ID: RM_HCI_004
TERMINAL_MAX_WIDTH: int = 80

# ---------------------------------------------------------------------------
# OOXML unit conversion constants.
# OOXML uses English Metric Units (EMU) for all dimensional measurements.
# All OOXML calculations must use these named constants -- never bare integers.
# REF_ID: RM_DO178_001, WEB_PY_010
# ---------------------------------------------------------------------------

OOXML_EMU_PER_HALF_POINT: int = 6350    # Half-points to EMU -- used for font sizes
OOXML_EMU_PER_POINT: int      = 12700   # Points to EMU -- used for spacing values
OOXML_EMU_PER_INCH: int       = 914400  # Inches to EMU -- used for page dimensions

# ---------------------------------------------------------------------------
# Page layout constants for the page-break algorithm (writer.py).
# All values are in points. REF_ID: RM_DO178_001
# ---------------------------------------------------------------------------

# Usable body height per A4 page (297mm - top margin 25mm - bottom margin 25mm
# - header 12.5mm - footer 12.5mm = approximately 222mm = 630pt).
# Declared conservatively at 630pt to account for template margin variation.
OOXML_USABLE_HEIGHT_PT: float = 630.0

# Estimated characters per line at 11pt body font on A4 with standard margins
OOXML_CHARS_PER_LINE: int = 95

# Estimated line height in points for a single-spaced 11pt body font
OOXML_LINE_HEIGHT_PT: float = 13.2

# ---------------------------------------------------------------------------
# Valid authorization signals for Preview Mode confirmation.
# Lowercase only -- input is lowercased before comparison.
# REF_ID: RM_NIST_001
# ---------------------------------------------------------------------------

PREVIEW_VALID_CONFIRM_SIGNALS: set[str] = {
    "yes", "confirm", "ok", "correct", "proceed", "y",
}
PREVIEW_CANCEL_SIGNALS: set[str] = {"cancel", "stop", "no", "n"}

# ---------------------------------------------------------------------------
# Valid signals for template delete confirmation gate.
# Only the exact word "delete" triggers deletion -- all others cancel.
# REF_ID: RM_NIST_001
# ---------------------------------------------------------------------------

TEMPLATE_DELETE_CONFIRM_SIGNAL: str = "delete"

# ---------------------------------------------------------------------------
# Valid signals for template overwrite confirmation gate.
# REF_ID: RM_NIST_001
# ---------------------------------------------------------------------------

TEMPLATE_OVERWRITE_CONFIRM_SIGNAL: str = "replace"

# ---------------------------------------------------------------------------
# Main menu option numbers -- declared as constants to avoid magic numbers
# in menu routing logic.
# ---------------------------------------------------------------------------

MENU_OPTION_SINGLE: str  = "1"
MENU_OPTION_BATCH: str   = "2"
MENU_OPTION_LIBRARY: str = "3"
MENU_OPTION_SAMPLE: str  = "4"   # Generate a format sample template
MENU_OPTION_EXIT: str    = "5"

MENU_VALID_OPTIONS: set[str] = {
    MENU_OPTION_SINGLE,
    MENU_OPTION_BATCH,
    MENU_OPTION_LIBRARY,
    MENU_OPTION_SAMPLE,
    MENU_OPTION_EXIT,
}

# ---------------------------------------------------------------------------
# Template library sub-menu option numbers
# ---------------------------------------------------------------------------

LIBRARY_OPTION_LIST: str   = "1"
LIBRARY_OPTION_SAVE: str   = "2"
LIBRARY_OPTION_DELETE: str = "3"
LIBRARY_OPTION_RENAME: str = "4"
LIBRARY_OPTION_BACK: str   = "5"

LIBRARY_VALID_OPTIONS: set[str] = {
    LIBRARY_OPTION_LIST,
    LIBRARY_OPTION_SAVE,
    LIBRARY_OPTION_DELETE,
    LIBRARY_OPTION_RENAME,
    LIBRARY_OPTION_BACK,
}
