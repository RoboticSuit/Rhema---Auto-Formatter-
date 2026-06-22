# Source Trace:
# File: main_v0.2.py
# Knowledge Files: CodeSourceDB v3.6, SyntaxBiasDB v2.3, HumanSyntaxDB v1.2
# REF_IDs: RM_DO178_001, RM_HCI_004, RM_HCI_002, WEB_PY_002, WEB_PY_007
# Logic: Entry point. CLI argument parsing, mode routing, interactive menu loop,
#        and top-level exception handler that routes diagnostics to stderr.
#        v0.2 -- Added MENU_OPTION_SAMPLE routing and _run_sample_mode dispatcher.

"""
Rhema -- Auto Formatter
main.py

Single responsibility: Entry point. Parses CLI arguments, initialises the
application environment, and routes to single, batch, or library mode.
No business logic lives here -- routing only.

Run from the rhema/ root folder:
    python main.py

Governing standards:
  POSIX.1-2017 argument syntax and exit codes -- REF_ID: RM_HCI_004
  ISO 9241-110:2020 interaction principles   -- REF_ID: RM_HCI_002
  DO-178C state machine traceability         -- REF_ID: RM_DO178_001
"""

import sys
import argparse

from config.constants import (
    APP_NAME,
    APP_VERSION,
    MENU_OPTION_SINGLE,
    MENU_OPTION_BATCH,
    MENU_OPTION_LIBRARY,
    MENU_OPTION_SAMPLE,
    MENU_OPTION_EXIT,
    EXIT_SUCCESS,
    EXIT_ERROR,
)
from ui.menu import show_header, prompt_main_menu
from ui.messages import MSG_EXIT, MSG_GENERAL_ERROR, fmt, LABEL_INFO


# ---------------------------------------------------------------------------
# Mode dispatchers
# REF_ID: RM_HCI_002 -- each mode is independently navigable
# ---------------------------------------------------------------------------

def _run_single_mode() -> None:
    """Runs the single file format mode."""
    from modes.single import run_single_mode
    run_single_mode()


def _run_batch_mode() -> None:
    """Placeholder -- replaced by modes/batch.py integration."""
    print(fmt(LABEL_INFO, "Batch formatting is coming soon."))


def _run_library_mode() -> None:
    """Runs the template library management mode."""
    from modes.library import run_library_mode
    run_library_mode()


def _run_sample_mode() -> None:
    """Runs the format template generator mode."""
    from modes.sample import run_sample_mode
    run_sample_mode()


# ---------------------------------------------------------------------------
# Argument parser -- POSIX.1-2017 compliant flag syntax
# REF_ID: RM_HCI_004
# ---------------------------------------------------------------------------

def _build_arg_parser() -> argparse.ArgumentParser:
    """
    Build and return the argument parser.

    Supported flags:
      -s / --single   : launch directly into single file mode
      -b / --batch    : launch directly into batch mode
      -l / --library  : launch directly into template library mode
      -v / --version  : print version string and exit

    When no flags are provided the interactive menu is shown.
    """
    parser = argparse.ArgumentParser(
        prog=APP_NAME.lower(),
        description=f"{APP_NAME} -- Auto Formatter v{APP_VERSION}",
        add_help=True,
    )
    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument(
        "-s", "--single",
        action="store_true",
        help="Format a single document",
    )
    mode_group.add_argument(
        "-b", "--batch",
        action="store_true",
        help="Format multiple documents in batch",
    )
    mode_group.add_argument(
        "-l", "--library",
        action="store_true",
        help="Manage saved templates",
    )
    parser.add_argument(
        "-v", "--version",
        action="version",
        version=f"{APP_NAME} v{APP_VERSION}",
    )
    return parser


# ---------------------------------------------------------------------------
# Mode router
# ---------------------------------------------------------------------------

def _route(option: str) -> None:
    """
    Route a menu option string to the correct mode function.
    Unrecognised options are a no-op -- the menu loop guards against
    invalid input before reaching here.
    REF_ID: RM_DO178_001
    """
    if option == MENU_OPTION_SINGLE:
        _run_single_mode()
    elif option == MENU_OPTION_BATCH:
        _run_batch_mode()
    elif option == MENU_OPTION_LIBRARY:
        _run_library_mode()
    elif option == MENU_OPTION_SAMPLE:
        _run_sample_mode()
    # MENU_OPTION_EXIT is handled by the caller -- not routed here


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def main() -> int:
    """
    Application entry point.

    Flow:
      1. Parse CLI arguments.
      2. If a mode flag is provided, run that mode directly and exit.
      3. Otherwise show the interactive menu and loop until the user exits.

    Returns an exit code per POSIX.1-2017. REF_ID: RM_HCI_004
    """
    parser = _build_arg_parser()
    args = parser.parse_args()

    show_header()

    # Direct launch via CLI flag -- skip menu, run mode once, exit
    if args.single:
        _run_single_mode()
        return EXIT_SUCCESS
    if args.batch:
        _run_batch_mode()
        return EXIT_SUCCESS
    if args.library:
        _run_library_mode()
        return EXIT_SUCCESS

    # Interactive menu loop -- runs until the user selects Exit.
    # Loop has a defined exit condition (MENU_OPTION_EXIT). REF_ID: RM_DO178_001
    while True:
        option = prompt_main_menu()

        if option == MENU_OPTION_EXIT:
            print(MSG_EXIT)
            return EXIT_SUCCESS

        _route(option)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        # Entry-point only catch-all -- routes diagnostic detail to stderr so
        # the stack trace is never visible to the non-technical user on stdout.
        # Per POSIX.1-2017 and WEB_PY_002 boundary exception rule.
        # REF_ID: RM_HCI_004, WEB_PY_002
        import traceback
        traceback.print_exc(file=sys.stderr)
        print(MSG_GENERAL_ERROR)
        sys.exit(EXIT_ERROR)
