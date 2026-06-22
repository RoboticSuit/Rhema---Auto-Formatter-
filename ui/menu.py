# Source Trace:
# File: menu_v0.1.py
# Knowledge Files: CodeSourceDB v3.6, SyntaxBiasDB v2.3, HumanSyntaxDB v1.2, mainProtocol v5.5
# REF_IDs: RM_DO178_001, RM_HCI_002, RM_HCI_004, WEB_PY_001
# Logic: Interactive menu display and strictly validated user input routing with explicit typing.

"""
Rhema -- Auto Formatter
ui/menu.py

Single responsibility: Renders the main menu and reads the user's mode
selection. Returns the selected option as a constant string for main.py
to route. No business logic lives here.

Governing standards:
  ISO 9241-110:2020 interaction principles -- REF_ID: RM_HCI_002
  POSIX.1-2017 standard streams            -- REF_ID: RM_HCI_004
  DO-178C state machine traceability       -- REF_ID: RM_DO178_001
"""

import sys
from typing import Optional

from ui.messages import (
    HEADER,
    MAIN_MENU,
    MAIN_MENU_INVALID,
    MSG_EXIT,
)
from config.constants import MENU_VALID_OPTIONS, MENU_OPTION_EXIT


def show_header() -> None:
    """Print the application header to stdout."""
    print(HEADER)


def prompt_main_menu() -> str:
    """
    Display the main menu and block until the user enters a valid option.

    Returns the selected option string (one of MENU_VALID_OPTIONS).
    Loops on invalid input, printing an error message each time.
    Handles KeyboardInterrupt (Ctrl+C) and EOF (Ctrl+D) gracefully by
    returning MENU_OPTION_EXIT so the caller can shut down cleanly.

    Output goes to stdout. Input is read from stdin.
    REF_ID: RM_HCI_004, RM_HCI_002
    """
    while True:
        print(MAIN_MENU)
        try:
            raw: str = input("> ").strip()
        except (KeyboardInterrupt, EOFError):
            # Ctrl+C or Ctrl+D -- treat as exit request
            print()  # newline after the interrupt character
            return MENU_OPTION_EXIT

        if raw in MENU_VALID_OPTIONS:
            return raw

        # Invalid input -- inform the user and loop
        print(MAIN_MENU_INVALID)
