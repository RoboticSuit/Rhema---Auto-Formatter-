# Source Trace:
# File: preview.py
# Knowledge Files: CodeSourceDB v3.4, SyntaxBiasDB v2.3, HumanSyntaxDB v1.2, mainProtocol v5.5
# REF_IDs: RM_DO178_001, RM_NIST_001, RM_HCI_002, WEB_PY_001, RM_HCI_001
# Logic: HITL validation loop with strict bounds and explicitly typed control flow.

"""
Rhema -- Auto Formatter
ui/preview.py

Single responsibility: Runs the Preview Mode HITL interaction loop.
Displays detected implicit structure to the user, collects confirmation
or correction, and returns the final confirmed structure for processing.

Preview Mode is the primary HITL gate. No transformation is applied
to any document until the user explicitly authorizes the detected
structure from this module.

Returns a PreviewResult containing:
  - authorized: True if the user confirmed and processing may proceed
  - confirmed_elements: the final list of ImplicitElement objects
    after any user adjustments
  - cancelled: True if the user cancelled or cycles were exhausted

Governing standards:
  NIST RMF Step 6 HITL gate     -- REF_ID: RM_NIST_001
  ISO 9241-110:2020 interaction -- REF_ID: RM_HCI_002
  DO-178C state traceability    -- REF_ID: RM_DO178_001
"""

from dataclasses import dataclass, field
from typing import List

from core.detector import ImplicitElement, DetectionResult
from config.constants import (
    PREVIEW_MAX_ADJUSTMENT_CYCLES,
    PREVIEW_VALID_CONFIRM_SIGNALS,
    PREVIEW_CANCEL_SIGNALS,
)
from ui.messages import (
    msg_preview_header,
    msg_preview_element,
    PREVIEW_PROMPT,
    msg_preview_adjusted,
    PREVIEW_RECONFIRM_PROMPT,
    msg_preview_cancelled,
    msg_preview_exhausted,
    fmt,
    LABEL_INFO,
    LABEL_ERROR,
)


# ---------------------------------------------------------------------------
# Result data structure
# ---------------------------------------------------------------------------

@dataclass
class PreviewResult:
    """
    Output of the preview loop.
    """
    authorized: bool = False
    cancelled: bool = False
    confirmed_elements: List[ImplicitElement] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _display_elements(elements: List[ImplicitElement]) -> None:
    """Print the current list of implicit elements."""
    for elem in elements:
        print(msg_preview_element(elem.paragraph_index, elem.inferred_type, elem.text_excerpt))
    print()


def _prompt_user(prompt_text: str) -> str:
    """Read a line from stdin securely."""
    try:
        return input(prompt_text + "> ").strip()
    except (KeyboardInterrupt, EOFError):
        print()
        return "cancel"


def _apply_correction(elements: List[ImplicitElement], raw_input: str) -> List[ImplicitElement]:
    """
    Attempt to parse a correction command (e.g., a paragraph number)
    and remove it from the list.
    Returns the updated list.
    """
    try:
        target_index: int = int(raw_input)
        updated: List[ImplicitElement] = []
        for elem in elements:
            if elem.paragraph_index != target_index:
                updated.append(elem)
        return updated
    except ValueError:
        print(fmt(LABEL_ERROR, "Invalid input. Please type 'yes', 'cancel', or a paragraph number."))
        return elements


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def run_preview(filename: str, elements: List[ImplicitElement]) -> PreviewResult:
    """
    Run the HITL preview loop for the given elements.
    REF_ID: RM_NIST_001
    """
    if len(elements) == 0:
        return PreviewResult(authorized=True, confirmed_elements=[])

    print(msg_preview_header(filename))
    _display_elements(elements)
    print(PREVIEW_PROMPT)

    cycle: int = 0
    while cycle < PREVIEW_MAX_ADJUSTMENT_CYCLES:
        raw: str = _prompt_user("")

        # -- CANCEL path --
        if raw.lower() in PREVIEW_CANCEL_SIGNALS or len(raw) == 0:
            print(msg_preview_cancelled(filename))
            return PreviewResult(cancelled=True)

        # -- CONFIRM path --
        if raw.lower() in PREVIEW_VALID_CONFIRM_SIGNALS:
            return PreviewResult(authorized=True, confirmed_elements=elements)

        # -- ADJUST path --
        updated: List[ImplicitElement] = _apply_correction(elements, raw)
        cycle += 1
        remaining: int = PREVIEW_MAX_ADJUSTMENT_CYCLES - cycle

        if len(updated) == 0:
            print(fmt(LABEL_INFO, "All elements cleared. Proceeding with no structural changes."))
            return PreviewResult(authorized=True, confirmed_elements=[])

        elements = updated
        print(msg_preview_adjusted(cycle, PREVIEW_MAX_ADJUSTMENT_CYCLES))
        _display_elements(elements)

        if remaining > 0:
            print(PREVIEW_RECONFIRM_PROMPT)

    # Cycle limit reached -- escalate
    print(msg_preview_exhausted(filename))
    return PreviewResult(cancelled=True)
