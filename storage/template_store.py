# Source Trace:
# File: template_store_v0.1.py
# Knowledge Files: CodeSourceDB v3.6, SyntaxBiasDB v2.3, HumanSyntaxDB v1.2, mainProtocol v5.5
# REF_IDs: RM_DO178_001, RM_ISO_001, RM_HCI_004, WEB_PY_001
# Logic: File system operations for template library with explicit typing and bounded checks.

"""
Rhema -- Auto Formatter
storage/template_store.py

Single responsibility: Reads and writes the persistent template library.
Handles all file system operations for saving, listing, selecting,
deleting, and renaming .tmpt files in the templates/ folder.
No UI logic lives here -- only storage operations.

Governing standards:
  DO-178C source traceability       -- REF_ID: RM_DO178_001
  ISO 31000 risk treatment          -- REF_ID: RM_ISO_001
  POSIX.1-2017 file operations      -- REF_ID: RM_HCI_004
"""

import os
import shutil
import json
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Any

from config.constants import (
    DEFAULT_TEMPLATE_DIR,
    TEMPLATE_EXTENSION,
)

# ---------------------------------------------------------------------------
# Library index file
# Stores display names mapped to filenames so the user sees friendly names
# rather than raw filenames. Stored as JSON in the templates/ folder.
# REF_ID: RM_DO178_001
# ---------------------------------------------------------------------------

_INDEX_FILENAME: str = "library_index.json"


def _index_path() -> Path:
    """Return the full path to the library index file."""
    return DEFAULT_TEMPLATE_DIR / _INDEX_FILENAME


def _load_index() -> Dict[str, Dict[str, str]]:
    """
    Load the library index from disk.
    Returns an empty dict if the index does not exist yet.
    Index format: { "display_name": { "filename": str, "saved": str } }
    """
    path: Path = _index_path()

    if path.exists() is False:
        return {}

    try:
        content: str = path.read_text(encoding="utf-8")
        return json.loads(content)
    except Exception:
        return {}


def _save_index(index: Dict[str, Dict[str, str]]) -> None:
    """
    Save the library index to disk.
    """
    path: Path = _index_path()
    path.write_text(json.dumps(index, indent=2), encoding="utf-8")


def ensure_library_exists() -> None:
    """
    Ensure the templates directory exists on disk.
    """
    if DEFAULT_TEMPLATE_DIR.exists() is False:
        DEFAULT_TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)


def list_templates() -> List[Dict[str, str]]:
    """
    Return a list of all saved templates, sorted alphabetically by display name.
    Each item is a dict with keys: display_name, filename, saved
    """
    index: Dict[str, Dict[str, str]] = _load_index()
    result: List[Dict[str, str]] = []

    for display_name, data in index.items():
        item: Dict[str, str] = {
            "display_name": display_name,
            "filename": data.get("filename", ""),
            "saved": data.get("saved", ""),
        }
        result.append(item)

    return sorted(result, key=lambda x: x["display_name"].lower())


def template_exists(display_name: str) -> bool:
    """
    Check if a template with the given display name exists in the library.
    """
    index: Dict[str, Dict[str, str]] = _load_index()
    return display_name in index


def get_template_path(display_name: str) -> Path:
    """
    Return the full file path for the given template display name.
    Raises ValueError if the template does not exist.
    """
    index: Dict[str, Dict[str, str]] = _load_index()

    if display_name not in index:
        raise ValueError("Template not found in index.")

    filename: str = index[display_name].get("filename", "")
    return DEFAULT_TEMPLATE_DIR / filename


def save_template(source_path: Path, display_name: str) -> bool:
    """
    Copy a document into the library and add it to the index.
    Generates a safe filename to avoid file system issues.

    Returns True on success, False if the file could not be copied.
    """
    ensure_library_exists()
    index: Dict[str, Dict[str, str]] = _load_index()

    # Generate safe filename by removing non-alphanumeric characters
    safe_name: str = "".join(c if c.isalnum() else "_" for c in display_name)
    filename: str = safe_name + TEMPLATE_EXTENSION
    target_path: Path = DEFAULT_TEMPLATE_DIR / filename

    # Collision avoidance: append a number if the file already exists
    counter: int = 1
    while target_path.exists() is True:
        filename = safe_name + "_" + str(counter) + TEMPLATE_EXTENSION
        target_path = DEFAULT_TEMPLATE_DIR / filename
        counter += 1

    try:
        shutil.copy2(source_path, target_path)
    except Exception:
        return False

    index[display_name] = {
        "filename": filename,
        "saved": datetime.now().strftime("%Y-%m-%d"),
    }
    _save_index(index)
    return True


def delete_template(display_name: str) -> bool:
    """
    Remove a template from the library and delete its file from disk.

    Returns True on success.
    Returns False if the display name is not in the library.

    The delete gate (user confirmation) is enforced by the caller before
    this function is called. REF_ID: RM_NIST_001
    """
    index: Dict[str, Dict[str, str]] = _load_index()

    if display_name not in index:
        return False

    filename: str = index[display_name].get("filename", "")
    file_path: Path = DEFAULT_TEMPLATE_DIR / filename

    # Delete the file if it exists -- if already missing, proceed with
    # index cleanup rather than blocking the user.
    if file_path.exists() is True:
        try:
            file_path.unlink()
        except OSError:
            return False

    del index[display_name]
    _save_index(index)
    return True


def rename_template(old_name: str, new_name: str) -> bool:
    """
    Rename a template's display name in the library index.
    The underlying file on disk is not renamed -- only the index entry.

    Returns True on success.
    Returns False if old_name is not found or new_name already exists.
    REF_ID: RM_NIST_001 (rename confirmation enforced by caller)
    """
    index: Dict[str, Dict[str, str]] = _load_index()

    if old_name not in index:
        return False

    if new_name in index:
        return False

    data: Dict[str, str] = index[old_name]
    del index[old_name]
    index[new_name] = data
    _save_index(index)

    return True
