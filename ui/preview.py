# Source Trace:
# File: preview_v0.2.py
# Knowledge Files: CodeSourceDB v3.6, SyntaxBiasDB v2.3, HumanSyntaxDB v1.2
# REF_IDs: RM_DO178_001, RM_NIST_001, RM_HCI_002, WEB_PY_001, WEB_PY_007
# Logic: HITL validation loop with strict bounds and explicitly typed control flow.
#        v0.2 -- _display_elements now uses sequential 1-based display numbers.
#                _apply_correction matches against sequential display number, not
#                paragraph_index. Consistent with corrected PREVIEW_PROMPT text.

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
    """Output of the preview loop."""
    authorized: bool = False
    cancelled: bool = False
    confirmed_elements: List[ImplicitElement] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _display_elements(elements: List[ImplicitElement]) -> None:
    """
    Print the current list of implicit elements with sequential 1-based numbers.

    The number shown to the user is a sequential display index (1, 2, 3...),
    NOT the raw paragraph_index from the source document. The user types
    these sequential numbers to remove items. REF_ID: RM_HCI_002 (fix 4)
    """
    for display_num, elem in enumerate(elements, start=1):
        print(msg_preview_element(display_num, elem.text_excerpt, elem.inferred_type))
    print()


def _prompt_user(prompt_text: str) -> str:
    """Read a line from stdin. Handles Ctrl+C and EOF gracefully."""
    try:
        return input(prompt_text + "> ").strip()
    except (KeyboardInterrupt, EOFError):
        print()
        return "cancel"


def _apply_correction(
    elements: List[ImplicitElement],
    raw_input: str,
) -> List[ImplicitElement]:
    """
    Remove an element by its 1-based sequential display number.

    Accepts the display number the user saw on screen (1, 2, 3...),
    NOT the paragraph_index. Returns the updated list unchanged if the
    input is not a valid display number. REF_ID: RM_HCI_002 (fix 4)
    """
    try:
        display_num: int = int(raw_input)
    except ValueError:
        print(fmt(LABEL_ERROR,
            "Please type 'yes', 'cancel', or a number from the list above."))
        return elements

    # Display numbers are 1-based; convert to 0-based list index
    target_idx: int = display_num - 1

    if target_idx < 0 or target_idx >= len(elements):
        print(fmt(LABEL_ERROR,
            "That number is not in the list. Please try again."))
        return elements

    # Remove the element at the target position
    updated: List[ImplicitElement] = [
        elem for i, elem in enumerate(elements) if i != target_idx
    ]
    return updated


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def run_preview(filename: str, elements: List[ImplicitElement]) -> PreviewResult:
    """
    Run the HITL preview loop for the given elements.

    Displays elements with sequential 1-based numbers. Accepts:
      - Confirm signal ('yes', 'y', etc.) -> authorized=True
      - Cancel signal ('cancel', 'no', empty) -> cancelled=True
      - Integer -> removes the element at that display position

    Bounded by PREVIEW_MAX_ADJUSTMENT_CYCLES. Exhaustion returns cancelled=True.
    REF_ID: RM_NIST_001, WEB_PY_007
    """
    if len(elements) == 0:
        return PreviewResult(authorized=True, confirmed_elements=[])

    print(msg_preview_header(filename))
    _display_elements(elements)
    print(PREVIEW_PROMPT)

    cycle: int = 0
    # Loop is bounded by PREVIEW_MAX_ADJUSTMENT_CYCLES -- REF_ID: RM_MISRA_001
    while cycle < PREVIEW_MAX_ADJUSTMENT_CYCLES:
        raw: str = _prompt_user("")

        # -- CANCEL path --
        if raw.lower() in PREVIEW_CANCEL_SIGNALS or len(raw) == 0:
            print(msg_preview_cancelled(filename))
            return PreviewResult(cancelled=True)

        # -- CONFIRM path --
        if raw.lower() in PREVIEW_VALID_CONFIRM_SIGNALS:
            return PreviewResult(authorized=True, confirmed_elements=elements)

        # -- ADJUST path -- try to remove an element by display number
        updated: List[ImplicitElement] = _apply_correction(elements, raw)

        # If _apply_correction returned the same list (invalid input), do not
        # increment the cycle counter -- the user should try again without
        # consuming an adjustment slot. REF_ID: RM_HCI_002
        if updated is elements:
            continue

        cycle += 1
        remaining: int = PREVIEW_MAX_ADJUSTMENT_CYCLES - cycle

        if len(updated) == 0:
            print(fmt(LABEL_INFO,
                "All items cleared. Proceeding with no structural changes."))
            return PreviewResult(authorized=True, confirmed_elements=[])

        elements = updated
        print(msg_preview_adjusted(cycle, PREVIEW_MAX_ADJUSTMENT_CYCLES))
        _display_elements(elements)

        if remaining > 0:
            print(PREVIEW_RECONFIRM_PROMPT)

    # Cycle limit reached -- escalate, do not proceed
    # REF_ID: RM_DO178_001 (defined escalation path on loop exhaustion)
    print(msg_preview_exhausted(filename))
    return PreviewResult(cancelled=True)
