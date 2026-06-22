# Source Trace:
# File: messages_v0.4.py
# Knowledge Files: CodeSourceDB v3.6, SyntaxBiasDB v2.3, HumanSyntaxDB v1.2
# REF_IDs: RM_DO178_001, RM_ACC_001, RM_VIS_001, RM_HCI_002, WEB_PY_003
# Logic: Centralized message registry with strictly typed strings and helper functions.
#        v0.4 -- PREVIEW_PROMPT corrected to accurately describe correction input.
#                LIBRARY_NAME_COLUMN_WIDTH constant replaces bare magic number 30.

"""
Rhema -- Auto Formatter
ui/messages.py

Single responsibility: Holds every string Rhema displays to the user in the
terminal. No logic lives here -- only string definitions and one formatting
helper. All other modules import from here and never hardcode display text.

All messages follow the OutputPresentationDB v0.1 rules:
  - Plain language, no technical jargon
  - Active voice, sentences under 20 words
  - Every message prefixed with a label: [OK], [ERROR], [INFO], [CONFIRM]
  - Every error message ends with an action the user can take
  - No blame language

Governing standards:
  ISO 9241-110:2020 interaction principles -- REF_ID: RM_HCI_002
  WCAG 2.2 SC 1.4.1 use of color           -- REF_ID: RM_ACC_001
  ISO 9241-112:2017 visual presentation    -- REF_ID: RM_VIS_001
  DO-178C source traceability              -- REF_ID: RM_DO178_001
"""

from config.constants import (
    APP_NAME,
    APP_VERSION,
    APP_TAGLINE,
    LIBRARY_DISPLAY_WIDTH,
    LIBRARY_NAME_COLUMN_WIDTH,
)

# ---------------------------------------------------------------------------
# Label prefixes -- used on every message for screen reader compatibility.
# Color alone is never used to convey meaning. REF_ID: RM_ACC_001
# ---------------------------------------------------------------------------

LABEL_OK: str      = "[OK]"
LABEL_ERROR: str   = "[ERROR]"
LABEL_INFO: str    = "[INFO]"
LABEL_CONFIRM: str = "[CONFIRM]"

# ---------------------------------------------------------------------------
# Separator line -- used in menus and library listings.
# Plain ASCII dashes only. REF_ID: RM_VIS_001
# ---------------------------------------------------------------------------

SEPARATOR: str = "-" * LIBRARY_DISPLAY_WIDTH


# ---------------------------------------------------------------------------
# Formatting helper
# ---------------------------------------------------------------------------

def fmt(label: str, text: str) -> str:
    """Prepend a label to a message string for consistent terminal output."""
    return label + " " + text


# ---------------------------------------------------------------------------
# Application header -- displayed at launch
# ---------------------------------------------------------------------------

HEADER: str = (
    "\n" + SEPARATOR + "\n"
    + APP_NAME + " -- " + APP_TAGLINE + "  v" + APP_VERSION + "\n"
    + SEPARATOR
)

# ---------------------------------------------------------------------------
# Main menu
# ---------------------------------------------------------------------------

MAIN_MENU: str = (
    "\n" + fmt(LABEL_INFO, "What would you like to do?") + "\n\n"
    "  1. Format a document\n"
    "  2. Format multiple documents (batch)\n"
    "  3. Manage my templates\n"
    "  4. Exit\n\n"
    "Type a number and press Enter."
)

MAIN_MENU_INVALID: str = fmt(
    LABEL_ERROR,
    "That is not a valid option. Please type 1, 2, 3, or 4."
)

# ---------------------------------------------------------------------------
# Exit
# ---------------------------------------------------------------------------

MSG_EXIT: str = fmt(LABEL_INFO, "Closing Rhema. Goodbye.")

# ---------------------------------------------------------------------------
# General progress messages
# ---------------------------------------------------------------------------

def msg_reading(filename: str) -> str:
    return fmt(LABEL_INFO, "Reading " + filename + " ...")

def msg_applying_template(template_name: str) -> str:
    return fmt(LABEL_INFO, "Applying template: " + template_name)

MSG_ADDING_COVER: str = fmt(LABEL_INFO, "Adding cover page ...")
MSG_BUILDING_TOC: str = fmt(LABEL_INFO, "Building table of contents ...")
MSG_SAVING_DOC: str   = fmt(LABEL_INFO, "Saving your document ...")

# ---------------------------------------------------------------------------
# Success messages
# ---------------------------------------------------------------------------

def msg_done_single(output_path: str) -> str:
    return (
        fmt(LABEL_OK, "Done. Your formatted document is saved here:") + "\n"
        "         " + output_path
    )

def msg_done_batch_file(filename: str, output_path: str) -> str:
    return fmt(LABEL_OK, filename + " -- done. Saved to " + output_path)

def msg_template_saved(display_name: str) -> str:
    return fmt(LABEL_OK, "Template '" + display_name + "' saved to your library.")

def msg_template_deleted(display_name: str) -> str:
    return fmt(LABEL_OK, "Template '" + display_name + "' has been deleted.")

def msg_template_renamed(new_name: str) -> str:
    return fmt(LABEL_OK, "Template renamed to '" + new_name + "'.")

# ---------------------------------------------------------------------------
# Error messages -- each ends with a plain-language user action.
# REF_ID: RM_HCI_002, RM_ACC_002
# ---------------------------------------------------------------------------

def msg_file_not_found(filename: str) -> str:
    return (
        fmt(LABEL_ERROR, "Rhema could not find the file: " + filename) + "\n"
        "        Check that the file exists in the location you selected, then try again."
    )

def msg_file_unreadable(filename: str) -> str:
    return (
        fmt(LABEL_ERROR, "Rhema could not open: " + filename) + "\n"
        "        The file may be open in another program. Close it and try again."
    )

def msg_unsupported_format(filename: str) -> str:
    return (
        fmt(LABEL_ERROR, filename + " is not a supported file type.") + "\n"
        "        Rhema works with .txt, .docx, and .doc files only."
    )

def msg_template_not_found(name: str) -> str:
    return (
        fmt(LABEL_ERROR, "The template '" + name + "' could not be found.") + "\n"
        "        Go to Manage Templates to check your saved templates."
    )

def msg_template_missing_styles(name: str) -> str:
    return (
        fmt(LABEL_ERROR, "The template '" + name + "' is missing some required sections.") + "\n"
        "        Rhema cannot format your document without a complete template.\n"
        "        Try a different template or check that the template file is complete."
    )

def msg_output_not_writable(output_path: str) -> str:
    return (
        fmt(LABEL_ERROR, "Rhema could not save to: " + output_path) + "\n"
        "        Check that you have permission to save files in that folder, then try again."
    )

def msg_partial_output_removed(filename: str) -> str:
    return (
        fmt(LABEL_ERROR, "Something went wrong while saving " + filename + ".") + "\n"
        "        The incomplete file has been removed. Please try again."
    )

def msg_preview_exhausted(filename: str) -> str:
    return (
        fmt(LABEL_ERROR, "Rhema could not automatically identify the structure of: " + filename) + "\n"
        "        This document needs to be formatted manually before Rhema can process it.\n"
        "        The file has been skipped."
    )

MSG_NO_TEMPLATES: str = (
    fmt(LABEL_INFO, "You have no saved templates yet.") + "\n"
    "        You can save a template when formatting a document, or add one from Manage Templates."
)

MSG_TEMPLATE_NAME_TAKEN: str = (
    fmt(LABEL_ERROR, "A template with that name already exists.") + "\n"
    "        Please choose a different name."
)

MSG_TEMPLATE_NAME_EMPTY: str = fmt(
    LABEL_ERROR,
    "Template name cannot be empty. Please enter a name."
)

MSG_NO_FILE_PROVIDED: str = fmt(
    LABEL_ERROR,
    "No file was provided. Please drag a file onto the window or type its path."
)

MSG_GENERAL_ERROR: str = fmt(
    LABEL_ERROR,
    "Something went wrong. Please try again."
)

# ---------------------------------------------------------------------------
# Preview Mode messages -- HITL gate
# REF_ID: RM_NIST_001
# ---------------------------------------------------------------------------

def msg_preview_header(filename: str) -> str:
    return (
        "\n" + SEPARATOR + "\n"
        + fmt(LABEL_CONFIRM, "Before formatting " + filename + ", Rhema needs your help.") + "\n"
        "\nRhema found text that looks like headings and sections, but they are not\n"
        "formally marked as such. Here is what Rhema detected:\n"
    )

def msg_preview_element(index: int, text_excerpt: str, inferred_type: str) -> str:
    """
    Display one detected element. index is the 1-based sequential display number
    shown to the user -- NOT the raw paragraph_index from the source document.
    The user types this number to remove the element. REF_ID: RM_HCI_002
    """
    return "  " + str(index) + ". \"" + text_excerpt + "\" -- looks like: " + inferred_type


# PREVIEW_PROMPT: corrected from "describe what needs to change" to accurately
# state that the user types a number to remove an item. REF_ID: RM_HCI_002 (fix 4)
PREVIEW_PROMPT: str = (
    "\nIs this correct?\n"
    "Type 'yes' to continue, or type the number of any item to remove it.\n"
    "Type 'cancel' to skip this file."
)

def msg_preview_adjusted(cycle: int, max_cycles: int) -> str:
    remaining: int = max_cycles - cycle
    suffix: str = "s" if remaining != 1 else ""
    return (
        "\n" + fmt(LABEL_INFO, "Got it. Here is the updated structure:") + "\n"
        "(" + str(remaining) + " adjustment" + suffix + " remaining)\n"
    )

# PREVIEW_RECONFIRM_PROMPT: also corrected to match actual input handling.
PREVIEW_RECONFIRM_PROMPT: str = (
    "Is this correct now? Type 'yes' to continue, or type a number to remove another item.\n"
    "Type 'cancel' to skip this file."
)

def msg_preview_cancelled(filename: str) -> str:
    return fmt(LABEL_INFO, filename + " has been skipped. No changes were made.")

# ---------------------------------------------------------------------------
# Confirmation prompts -- irreversible actions
# REF_ID: RM_NIST_001
# ---------------------------------------------------------------------------

def msg_delete_gate(template_name: str) -> str:
    return (
        "\n" + fmt(LABEL_CONFIRM, "Are you sure you want to delete '" + template_name + "'?") + "\n"
        "        This cannot be undone.\n"
        "        Type 'delete' to confirm, or press Enter to cancel."
    )

def msg_overwrite_gate(name: str) -> str:
    return (
        "\n" + fmt(LABEL_CONFIRM, "A template named '" + name + "' already exists.") + "\n"
        "        Do you want to replace it?\n"
        "        Type 'replace' to confirm, or press Enter to cancel."
    )

MSG_SAVE_AFTER_FORMAT: str = (
    "\n" + fmt(LABEL_CONFIRM, "Would you like to save this template to your library for future use?") + "\n"
    "        Type 'yes' to save it, or press Enter to skip."
)

MSG_SAVE_TEMPLATE_NAME_PROMPT: str = fmt(
    LABEL_CONFIRM,
    "What would you like to name this template?"
)

MSG_RENAME_PROMPT: str = fmt(
    LABEL_CONFIRM,
    "Enter the new name for this template."
)

# ---------------------------------------------------------------------------
# Cover page field prompts -- triggered when required metadata is missing
# REF_ID: RM_NIST_001
# ---------------------------------------------------------------------------

MSG_COVER_PAGE_INTRO: str = fmt(
    LABEL_CONFIRM,
    "Rhema needs a few details for your cover page."
)

MSG_COVER_TITLE_PROMPT: str  = "  Document title (or press Enter to use the filename):"
MSG_COVER_DATE_PROMPT: str   = "  Date (or press Enter to use today's date):"
MSG_COVER_AUTHOR_PROMPT: str = "  Author name (or press Enter to skip):"
MSG_COVER_ORG_PROMPT: str    = "  Organisation name (or press Enter to skip):"

# ---------------------------------------------------------------------------
# Template library listing
# REF_ID: RM_VIS_001
# ---------------------------------------------------------------------------

LIBRARY_HEADER: str = "\n" + fmt(LABEL_INFO, "Your saved templates:") + "\n" + SEPARATOR
LIBRARY_FOOTER_SEPARATOR: str = SEPARATOR

def msg_library_entry(index: int, display_name: str, saved_date: str) -> str:
    # Uses LIBRARY_NAME_COLUMN_WIDTH constant -- no magic number. REF_ID: RM_DO178_001
    name_col: str = display_name.ljust(LIBRARY_NAME_COLUMN_WIDTH)
    return "  " + str(index) + ". " + name_col + "  Saved: " + saved_date

def msg_library_count(count: int) -> str:
    suffix: str = "s" if count != 1 else ""
    return "\n" + str(count) + " template" + suffix + " saved."

LIBRARY_MENU: str = (
    "\n" + fmt(LABEL_INFO, "What would you like to do?") + "\n\n"
    "  1. List saved templates\n"
    "  2. Save a new template\n"
    "  3. Delete a template\n"
    "  4. Rename a template\n"
    "  5. Back to main menu\n\n"
    "Type a number and press Enter."
)

LIBRARY_MENU_INVALID: str = fmt(
    LABEL_ERROR,
    "That is not a valid option. Please type 1, 2, 3, 4, or 5."
)

MSG_LIBRARY_SELECT_PROMPT: str = fmt(
    LABEL_CONFIRM,
    "Type the number of the template you want to select, or press Enter to cancel."
)

MSG_LIBRARY_SELECT_INVALID: str = fmt(
    LABEL_ERROR,
    "That number is not in the list. Please try again."
)

MSG_NEW_TEMPLATE_PATH_PROMPT: str = fmt(
    LABEL_CONFIRM,
    "Drag the .tmpt file here or type its full path, then press Enter."
)

def msg_template_path_not_found(path: str) -> str:
    return (
        fmt(LABEL_ERROR, "Rhema could not find a file at: " + path) + "\n"
        "        Check the path and try again."
    )

MSG_WRONG_TEMPLATE_EXTENSION: str = (
    fmt(LABEL_ERROR, "That file does not appear to be a .tmpt template.") + "\n"
    "        Please select a file that ends in .tmpt."
)

# ---------------------------------------------------------------------------
# Batch summary report
# REF_ID: RM_HCI_002
# ---------------------------------------------------------------------------

def msg_batch_header(total: int) -> str:
    return (
        "\n" + SEPARATOR + "\n"
        + fmt(LABEL_INFO, "Batch complete. Here is a summary:") + "\n"
        + SEPARATOR
    )

def msg_batch_success_line(filename: str, output_path: str) -> str:
    return "  " + LABEL_OK + "      " + filename + "  ->  " + output_path

def msg_batch_skip_line(filename: str, reason: str) -> str:
    return "  " + LABEL_ERROR + "   " + filename + " -- skipped. " + reason

def msg_batch_footer(success_count: int, total: int) -> str:
    skipped: int = total - success_count
    result: str = "\n" + str(success_count) + " of " + str(total) + " documents formatted successfully."
    if skipped > 0:
        suffix: str = "s" if skipped != 1 else ""
        result += "\n" + str(skipped) + " document" + suffix + " skipped. See above for details."
    return result

# ---------------------------------------------------------------------------
# Input path prompts -- used in single and batch mode
# ---------------------------------------------------------------------------

MSG_INPUT_PATH_SINGLE_PROMPT: str = fmt(
    LABEL_CONFIRM,
    "Drag your document here or type its full path, then press Enter."
)

MSG_INPUT_PATH_BATCH_PROMPT: str = (
    fmt(LABEL_CONFIRM, "Drag your documents here or type their full paths, one per line.") + "\n"
    "        Press Enter twice when you are done."
)

MSG_OUTPUT_PATH_PROMPT: str = fmt(
    LABEL_CONFIRM,
    "Where would you like to save the output? Press Enter to use the default output folder."
)

def msg_starting_single(filename: str) -> str:
    return fmt(LABEL_INFO, "Starting to format: " + filename)

def msg_starting_batch(count: int) -> str:
    return fmt(LABEL_INFO, "Starting batch. " + str(count) + " documents to process.")
