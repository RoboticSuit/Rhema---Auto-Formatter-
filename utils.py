# Source Trace:
# File: utils_v0.1.py
# Knowledge Files: CodeSourceDB v3.6, SyntaxBiasDB v2.3, HumanSyntaxDB v1.2
# REF_IDs: RM_DO178_001, RM_HCI_004, WEB_PY_004, WEB_PY_012
# Logic: Shared utility functions used by multiple modules. Extracted here per
#        WEB_PY_004 DRY rule -- clean_path was duplicated in single.py and library.py.

"""
Rhema -- Auto Formatter
utils.py

Single responsibility: Shared utility functions used across multiple modules.
No business logic, no UI output, no file I/O lives here -- only pure utilities.

Currently provides:
  clean_path(raw) -- strips terminal path artifacts from drag-dropped paths

Governing standards:
  POSIX.1-2017 path sanitation    -- REF_ID: RM_HCI_004
  DO-178C source traceability     -- REF_ID: RM_DO178_001
  WEB_PY_004 DRY -- shared utility -- REF_ID: WEB_PY_004
  WEB_PY_012 input sanitization   -- REF_ID: WEB_PY_012
"""


def clean_path(raw: str) -> str:
    """
    Remove terminal path artifacts from drag-dropped or copy-pasted paths.

    Handles paths pasted from:
      - Windows Explorer drag-and-drop: plain path
      - PowerShell with & invocation prefix: & 'C:\\path' or & "C:\\path"
      - PowerShell quote-wrapped paths: 'C:\\path' or "C:\\path"
      - PowerShell escaped apostrophes: C:\\Iason''s PC\\...

    Processing order (WEB_PY_012 -- apply each step once, in declared sequence):
      1. strip() -- remove surrounding whitespace
      2. Strip PowerShell & prefix ("& " prefix or bare "&")
      3. Remove ONE matching outer quote pair (single or double)
      4. Replace '' with ' (PowerShell escaped single-quote)
      5. Final strip()

    Returns the sanitized path string. Never raises.
    REF_ID: RM_HCI_004, WEB_PY_012
    """
    raw = raw.strip()

    # Step 2 -- Remove PowerShell invocation prefix (& 'path' or &'path')
    if raw.startswith("& ") is True:
        raw = raw[2:].strip()
    elif raw.startswith("&") is True:
        raw = raw[1:].strip()

    # Step 3 -- Remove ONE matching outer quote pair only.
    # Single-pair removal avoids stripping apostrophes inside paths.
    if len(raw) >= 2:
        if (raw[0] == '"' and raw[-1] == '"') or (raw[0] == "'" and raw[-1] == "'"):
            raw = raw[1:-1]

    # Step 4 -- Un-escape PowerShell single-quote escaping: '' -> '
    raw = raw.replace("''", "'")

    # Step 5 -- Final whitespace trim
    return raw.strip()
