# Source Trace:
# File: constants.py
# Knowledge Files: CodeSourceDB v3.4, SyntaxBiasDB v2.3, HumanSyntaxDB v1.2, mainProtocol v5.5
# REF_IDs: RM_DO178_001, RM_HCI_004, RM_NIST_001
# Logic: Centralized declaration of application constants, POSIX exit codes, and explicit UI signals with strict type hinting.

"""
Rhema - Auto Formatter
config/constants.py

Single responsibility: Declares all named constants used across the application.
No logic lives here. Every numeric value, default path, and fixed string that
appears more than once in the codebase is declared here and imported by name.

Governing standards:
  POSIX.1-2017 exit code semantics - REF_ID: RM_HCI_004
  DO-178C no magic numbers - REF_ID: RM_DO178_001
  SyntaxBiasDB v2.3 named constants rule
"""

import os
from pathlib import Path

# ---------------------------------------------------------------------------
# Application identity
# ---------------------------------------------------------------------------

APP_NAME: str = "Rhema"
APP_VERSION: str = "1.0.0"
APP_TAGLINE: str = "Document Formatter"

# ---------------------------------------------------------------------------
# Exit codes - all declared as named constants per POSIX.1-2017
# Exit code 0: success. Non-zero: failure. Each code has a unique meaning.
# REF_ID: RM_HCI_004
# ---------------------------------------------------------------------------

EXIT_SUCCESS: int = 0          # All operations completed successfully
EXIT_ERROR: int = 1            # General error - operation could not complete
EXIT_INPUT_ERROR: int = 2      # Input file missing, unreadable, or unsupported format
EXIT_TEMPLATE_ERROR: int = 3   # Template not found, unreadable, or missing required styles
EXIT_OUTPUT_ERROR: int = 4     # Cannot write to declared output path
EXIT_USER_CANCEL: int = 5      # User cancelled the operation in Preview Mode

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

# ---------------------------------------------------------------------------
# Default paths - resolved relative to the user's home directory
# so the application is portable across machines.
# REF_ID: RM_DO178_001
# ---------------------------------------------------------------------------

# Base directory for all Rhema data stored on the user's machine
_RHEMA_HOME: Path = Path.home() / ".rhema"

# Persistent template library - where saved .tmpt files are stored
DEFAULT_TEMPLATE_DIR: Path = _RHEMA_HOME / "templates"

# Default output directory - where formatted .docx files are written
DEFAULT_OUTPUT_DIR: Path = Path.cwd() / "output"

# ---------------------------------------------------------------------------
# Preview Mode limits
# Maximum number of adjustment cycles before escalation to skip.
# Three cycles gives the user two correction attempts after the initial
# detection display. Beyond this, automatic detection cannot resolve the
# document. REF_ID: RM_NIST_001
# ---------------------------------------------------------------------------

PREVIEW_MAX_ADJUSTMENT_CYCLES: int = 3

# ---------------------------------------------------------------------------
# Batch mode
# ---------------------------------------------------------------------------

# Minimum number of files to trigger batch mode display (vs single file mode)
BATCH_MIN_FILES: int = 2

# ---------------------------------------------------------------------------
# Terminal display
# ---------------------------------------------------------------------------

# Column width for the template library listing separator line
LIBRARY_DISPLAY_WIDTH: int = 52

# Maximum line width for terminal output - POSIX.1-2017 80-column convention
# REF_ID: RM_HCI_004
TERMINAL_MAX_WIDTH: int = 80

# ---------------------------------------------------------------------------
# Valid authorization signals for Preview Mode confirmation
# Lowercase only - input is lowercased before comparison
# REF_ID: RM_NIST_001
# ---------------------------------------------------------------------------

PREVIEW_VALID_CONFIRM_SIGNALS: set[str] = {"yes", "confirm", "ok", "correct", "proceed", "y"}
PREVIEW_CANCEL_SIGNALS: set[str] = {"cancel", "stop", "no", "n"}

# ---------------------------------------------------------------------------
# Valid signals for template delete confirmation gate
# Only the exact word "delete" triggers deletion - all others cancel
# REF_ID: RM_NIST_001
# ---------------------------------------------------------------------------

TEMPLATE_DELETE_CONFIRM_SIGNAL: str = "delete"

# ---------------------------------------------------------------------------
# Valid signals for template overwrite confirmation gate
# REF_ID: RM_NIST_001
# ---------------------------------------------------------------------------

TEMPLATE_OVERWRITE_CONFIRM_SIGNAL: str = "replace"

# ---------------------------------------------------------------------------
# Main menu option numbers - declared as constants to avoid magic numbers
# in menu routing logic
# ---------------------------------------------------------------------------

MENU_OPTION_SINGLE: str = "1"
MENU_OPTION_BATCH: str = "2"
MENU_OPTION_LIBRARY: str = "3"
MENU_OPTION_EXIT: str = "4"

MENU_VALID_OPTIONS: set[str] = {
    MENU_OPTION_SINGLE,
    MENU_OPTION_BATCH,
    MENU_OPTION_LIBRARY,
    MENU_OPTION_EXIT,
}

# ---------------------------------------------------------------------------
# Template library sub-menu option numbers
# ---------------------------------------------------------------------------

LIBRARY_OPTION_LIST: str = "1"
LIBRARY_OPTION_SAVE: str = "2"
LIBRARY_OPTION_DELETE: str = "3"
LIBRARY_OPTION_RENAME: str = "4"
LIBRARY_OPTION_BACK: str = "5"

LIBRARY_VALID_OPTIONS: set[str] = {
    LIBRARY_OPTION_LIST,
    LIBRARY_OPTION_SAVE,
    LIBRARY_OPTION_DELETE,
    LIBRARY_OPTION_RENAME,
    LIBRARY_OPTION_BACK,
}
