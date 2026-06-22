# Source Trace:
# File: library_v0.5.py
# Knowledge Files: CodeSourceDB v3.6, SyntaxBiasDB v2.3, HumanSyntaxDB v1.2, mainProtocol v5.5
# REF_IDs: RM_DO178_001, RM_NIST_001, RM_HCI_002, WEB_PY_001, WEB_PY_004, RM_HCI_004
# Logic: Orchestrates template library operations with explicit types, strictly bounded
#        loops, and path sanitization.
#        v0.5 -- clean_path imported from utils.py (WEB_PY_004 DRY fix 7).

"""
Rhema -- Auto Formatter
modes/library.py

Single responsibility: Orchestrates the template library management mode.
Routes user selections to template_store operations and handles all
user interaction for save, list, select, delete, and rename flows.
No storage logic lives here -- only orchestration and UI interaction.

Governing standards:
  NIST RMF Step 6 HITL controls     -- REF_ID: RM_NIST_001
  ISO 9241-110:2020 interaction     -- REF_ID: RM_HCI_002
  DO-178C state traceability        -- REF_ID: RM_DO178_001
  POSIX.1-2017 path sanitation      -- REF_ID: RM_HCI_004
"""

import sys
from pathlib import Path
from typing import List, Dict, Any, Optional

from storage.template_store import (
    ensure_library_exists,
    list_templates,
    template_exists,
    get_template_path,
    save_template,
    delete_template,
    rename_template,
)
from config.constants import (
    TEMPLATE_EXTENSION,
    ACCEPTED_TEMPLATE_EXTENSIONS,
    TEMPLATE_DELETE_CONFIRM_SIGNAL,
    TEMPLATE_OVERWRITE_CONFIRM_SIGNAL,
    LIBRARY_OPTION_LIST,
    LIBRARY_OPTION_SAVE,
    LIBRARY_OPTION_DELETE,
    LIBRARY_OPTION_RENAME,
    LIBRARY_OPTION_BACK,
    LIBRARY_VALID_OPTIONS,
    PREVIEW_VALID_CONFIRM_SIGNALS,
)

# Import shared path utility -- replaces locally duplicated _clean_path.
# REF_ID: WEB_PY_004 (DRY rule)
from utils import clean_path
from ui.messages import (
    LIBRARY_MENU,
    LIBRARY_MENU_INVALID,
    LIBRARY_HEADER,
    LIBRARY_FOOTER_SEPARATOR,
    MSG_NO_TEMPLATES,
    MSG_TEMPLATE_NAME_TAKEN,
    MSG_TEMPLATE_NAME_EMPTY,
    MSG_NEW_TEMPLATE_PATH_PROMPT,
    MSG_WRONG_TEMPLATE_EXTENSION,
    MSG_LIBRARY_SELECT_PROMPT,
    MSG_LIBRARY_SELECT_INVALID,
    MSG_RENAME_PROMPT,
    MSG_SAVE_TEMPLATE_NAME_PROMPT,
    msg_library_entry,
    msg_library_count,
    msg_template_saved,
    msg_template_deleted,
    msg_template_renamed,
    msg_delete_gate,
    msg_overwrite_gate,
    msg_template_not_found,
    msg_template_path_not_found,
    fmt,
    LABEL_ERROR,
    LABEL_INFO,
)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _prompt(message: str) -> str:
    """Print a prompt and return stripped user input. Handles Ctrl+C."""
    print(message)
    try:
        return input("> ").strip()
    except (KeyboardInterrupt, EOFError):
        print()
        return ""




def _display_library() -> List[Dict[str, str]]:
    """
    Print the full template library listing.
    Returns the list of template entries for use by callers.
    """
    entries: List[Dict[str, str]] = list_templates()
    if len(entries) == 0:
        print(MSG_NO_TEMPLATES)
        return []

    print(LIBRARY_HEADER)
    for i, entry in enumerate(entries, start=1):
        display_name: str = entry.get("display_name", "Unknown")
        saved_date: str = entry.get("saved", "Unknown")
        print(msg_library_entry(i, display_name, saved_date))

    print(LIBRARY_FOOTER_SEPARATOR)
    print(msg_library_count(len(entries)))
    return entries


def _pick_template_by_number(entries: List[Dict[str, str]]) -> Optional[Dict[str, str]]:
    """
    Ask the user to pick a template by number from a displayed list.
    Returns the selected entry dict, or None if the user cancels.
    REF_ID: RM_NIST_001
    """
    if len(entries) == 0:
        return None

    while True:
        raw: str = _prompt(MSG_LIBRARY_SELECT_PROMPT)
        if len(raw) == 0:
            return None

        if raw.isdigit() is True:
            index: int = int(raw) - 1
            if 0 <= index and index < len(entries):
                return entries[index]

        print(MSG_LIBRARY_SELECT_INVALID)


# ---------------------------------------------------------------------------
# Sub-operation handlers
# ---------------------------------------------------------------------------

def _do_list() -> None:
    """List all saved templates."""
    _display_library()


def _do_save() -> None:
    """
    Save a new .tmpt file to the library.
    Prompts for file path and display name.
    Applies overwrite gate if a name conflict exists.
    REF_ID: RM_NIST_001
    """
    # Step 1 -- get file path
    raw_path: str = _prompt(MSG_NEW_TEMPLATE_PATH_PROMPT)
    if len(raw_path) == 0:
        return

    # Clean path artifacts introduced by different terminals.
    raw_path = clean_path(raw_path)
    src: Path = Path(raw_path)

    if src.exists() is False:
        print(msg_template_path_not_found(raw_path))
        return

    if src.suffix.lower() not in ACCEPTED_TEMPLATE_EXTENSIONS:
        print(MSG_WRONG_TEMPLATE_EXTENSION)
        return

    # Step 2 -- get display name
    while True:
        raw_name: str = _prompt(MSG_SAVE_TEMPLATE_NAME_PROMPT)
        if len(raw_name) == 0:
            print(MSG_TEMPLATE_NAME_EMPTY)
            continue
        break

    # Step 3 -- check for name conflict and apply overwrite gate
    # REF_ID: RM_NIST_001
    if template_exists(raw_name) is True:
        confirm: str = _prompt(msg_overwrite_gate(raw_name))
        if confirm.lower() != TEMPLATE_OVERWRITE_CONFIRM_SIGNAL:
            print(fmt(LABEL_INFO, "Save cancelled."))
            return
        # Delete existing entry before saving new one
        delete_template(raw_name)

    # Step 4 -- save (src is explicitly passed as a Path object to satisfy Pylance)
    ok: bool = save_template(src, raw_name)
    if ok is True:
        print(msg_template_saved(raw_name))
    else:
        print(fmt(LABEL_ERROR, "Rhema could not save the template. Check the file and try again."))


def _do_delete() -> None:
    """
    Delete a saved template.
    Displays the library, asks user to pick by number, then applies
    the delete gate before executing. REF_ID: RM_NIST_001
    """
    entries: List[Dict[str, str]] = _display_library()
    if len(entries) == 0:
        return

    selected: Optional[Dict[str, str]] = _pick_template_by_number(entries)
    if selected is None:
        print(fmt(LABEL_INFO, "No template selected. Returning to library menu."))
        return

    display_name: str = selected.get("display_name", "")

    # Delete gate -- user must type the exact confirmation word
    # REF_ID: RM_NIST_001
    confirm: str = _prompt(msg_delete_gate(display_name))
    if confirm.lower() != TEMPLATE_DELETE_CONFIRM_SIGNAL:
        print(fmt(LABEL_INFO, "Delete cancelled."))
        return

    ok: bool = delete_template(display_name)
    if ok is True:
        print(msg_template_deleted(display_name))
    else:
        print(msg_template_not_found(display_name))


def _do_rename() -> None:
    """
    Rename a saved template's display name.
    Displays the library, asks user to pick, then prompts for new name.
    REF_ID: RM_NIST_001
    """
    entries: List[Dict[str, str]] = _display_library()
    if len(entries) == 0:
        return

    selected: Optional[Dict[str, str]] = _pick_template_by_number(entries)
    if selected is None:
        print(fmt(LABEL_INFO, "No template selected. Returning to library menu."))
        return

    display_name: str = selected.get("display_name", "")

    # Get new name
    while True:
        new_name: str = _prompt(MSG_RENAME_PROMPT)
        if len(new_name) == 0:
            print(MSG_TEMPLATE_NAME_EMPTY)
            continue

        if new_name == display_name:
            print(fmt(LABEL_INFO, "That is already the template name. Returning to library menu."))
            return

        if template_exists(new_name) is True:
            print(MSG_TEMPLATE_NAME_TAKEN)
            continue

        break

    ok: bool = rename_template(display_name, new_name)
    if ok is True:
        print(msg_template_renamed(new_name))
    else:
        print(fmt(LABEL_ERROR, "Rhema could not rename the template. Please try again."))


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def run_library_mode() -> None:
    """
    Run the template library management mode.
    Displays a sub-menu and loops until the user selects Back.
    Called by main.py when the user selects option 3 from the main menu.
    REF_ID: RM_DO178_001 -- loop has a defined exit condition
    """
    ensure_library_exists()

    while True:
        print(LIBRARY_MENU)
        try:
            raw: str = input("> ").strip()
        except (KeyboardInterrupt, EOFError):
            print()
            return

        if raw not in LIBRARY_VALID_OPTIONS:
            print(LIBRARY_MENU_INVALID)
            continue

        if raw == LIBRARY_OPTION_LIST:
            _do_list()
        elif raw == LIBRARY_OPTION_SAVE:
            _do_save()
        elif raw == LIBRARY_OPTION_DELETE:
            _do_delete()
        elif raw == LIBRARY_OPTION_RENAME:
            _do_rename()
        elif raw == LIBRARY_OPTION_BACK:
            return
