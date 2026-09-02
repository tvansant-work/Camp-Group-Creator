# camp_sync.py
"""
Backend module for Manual Group Alignment.
Handles: Google Drive path discovery, JSON state load/save,
and optimistic-lock concurrency checks.
"""

import json
import os
import time
import getpass
import platform
from pathlib import Path
from datetime import datetime

STATE_FILENAME = "camp_state.json"
APP_DATA_CHAIN = ["Outdoor Education Master Folder", "operations folder", "App Data"]
LOCAL_CONFIG_PATH = Path.home() / "Library" / "Application Support" / \
                    "Camp_Group_Creator" / "drive_location.json"


# ============================================================================
# 1. GOOGLE DRIVE PATH RESOLUTION
# ============================================================================

def find_google_drive_roots():
    """
    Scan ~/Library/CloudStorage/ for mounted Google Drive for Desktop folders.
    Returns a list of Path objects, one per mounted Google account
    (e.g. GoogleDrive-tvansant@friends.tas.edu.au).
    """
    cloud_storage = Path.home() / "Library" / "CloudStorage"
    if not cloud_storage.exists():
        return []

    roots = []
    for entry in cloud_storage.iterdir():
        if entry.is_dir() and entry.name.startswith("GoogleDrive-"):
            roots.append(entry)
    return sorted(roots)


def _load_local_config():
    """Read the previously-resolved App Data path, if any, from local (non-synced) config."""
    if LOCAL_CONFIG_PATH.exists():
        try:
            return json.loads(LOCAL_CONFIG_PATH.read_text())
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def _save_local_config(config: dict):
    LOCAL_CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    LOCAL_CONFIG_PATH.write_text(json.dumps(config, indent=2))


def _find_folder_by_name(start: Path, name: str, max_depth: int = 4) -> Path | None:
    """Depth-limited search under `start` for a directory named `name` (case-insensitive)."""
    if not start.exists():
        return None
    name_lower = name.strip().lower()
    start_depth = len(start.parts)
    for dirpath, dirnames, _filenames in os.walk(start):
        depth = len(Path(dirpath).parts) - start_depth
        if depth >= max_depth:
            dirnames[:] = []  # don't recurse further
            continue
        for d in dirnames:
            if d.strip().lower() == name_lower:
                return Path(dirpath) / d
    return None


def _locate_app_data_folder(drive_root: Path) -> Path | None:
    """
    Search one mounted Google Drive account for the App Data folder, checking
    every place a shared item can live: My Drive, shortcut targets (shared
    folders added via 'Add shortcut to Drive'), and Shared drives.
    """
    search_bases = [drive_root]  # fallback: scan the whole account root
    for sub in ("My Drive", ".shortcut-targets-by-id", "Shared drives"):
        p = drive_root / sub
        if p.exists():
            if sub == "My Drive":
                search_bases.insert(0, p)
            else:
                # these contain multiple sub-items, each a possible root
                search_bases[0:0] = [c for c in p.iterdir() if c.is_dir()]

    for base in search_bases:
        master = _find_folder_by_name(base, APP_DATA_CHAIN[0], max_depth=3)
        if not master:
            continue
        ops = _find_folder_by_name(master, APP_DATA_CHAIN[1], max_depth=2)
        if not ops:
            continue
        app_data = _find_folder_by_name(ops, APP_DATA_CHAIN[2], max_depth=2)
        if app_data:
            return app_data
        # Chain found down to "operations folder" but "App Data" is missing -- create it.
        app_data = ops / "App Data"
        app_data.mkdir(exist_ok=True)
        return app_data
    return None


def resolve_state_file_path(chosen_path: str | None = None) -> Path | None:
    """
    Determine the full path to camp_state.json inside
    'Outdoor Education Master Folder/operations folder/App Data'.

    Resolution order:
      1. `chosen_path` passed explicitly (user manually located it in the UI).
      2. Cached path from a prior successful resolution (LOCAL_CONFIG_PATH),
         if it still exists on disk.
      3. Auto-search every mounted Google Drive account for the App Data folder.
      4. None -- caller (UI) should show a manual folder picker / path entry.
    """
    if chosen_path:
        app_data = Path(chosen_path)
    else:
        config = _load_local_config()
        cached = config.get("app_data_path")
        if cached and Path(cached).exists():
            app_data = Path(cached)
        else:
            app_data = None
            for root in find_google_drive_roots():
                found = _locate_app_data_folder(root)
                if found:
                    app_data = found
                    break
            if app_data is None:
                return None

    state_path = app_data / STATE_FILENAME

    _save_local_config({"app_data_path": str(app_data)})

    if not state_path.exists():
        _write_json_atomic(state_path, _default_state())

    return state_path


# ============================================================================
# 2. JSON STATE SCHEMA + DEFAULTS
# ============================================================================

def _default_state() -> dict:
    return {
        "version": 1,                     # bumped on every successful save
        "last_modified": _now_iso(),
        "last_modified_by": _current_user_label(),
        "groups": {},                      # { "Bay of Fires": ["S001", "S002"], ... }
        "students": {}                     # { "S001": {name, friend_requests, form_data, medical_flags}, ... }
    }


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _current_user_label() -> str:
    """A human-readable 'who saved this' label -- OS username, since we don't have real auth."""
    try:
        return f"{getpass.getuser()} ({platform.node()})"
    except Exception:
        return "unknown"


# ============================================================================
# 3. ATOMIC READ / WRITE HELPERS
# ============================================================================

def _write_json_atomic(path: Path, data: dict):
    """
    Write via a temp file + rename, so a half-written file is never visible
    to Google Drive's sync watcher (which would otherwise sync a corrupt file).
    """
    tmp_path = path.with_suffix(".json.tmp")
    tmp_path.write_text(json.dumps(data, indent=2, ensure_ascii=False))
    os.replace(tmp_path, path)  # atomic on the same filesystem


def _read_json_with_retry(path: Path, retries: int = 3, delay: float = 0.3) -> dict:
    """
    Google Drive for Desktop can briefly lock/partially-sync a file.
    Retry a few times with a short delay before giving up.
    """
    last_err = None
    for _ in range(retries):
        try:
            text = path.read_text()
            return json.loads(text)
        except (json.JSONDecodeError, OSError) as e:
            last_err = e
            time.sleep(delay)
    raise RuntimeError(
        f"Could not read state file after {retries} attempts "
        f"(it may still be syncing from Google Drive): {last_err}"
    )


# ============================================================================
# 4. LOAD / SAVE WITH OPTIMISTIC LOCKING
# ============================================================================

class ConcurrencyError(Exception):
    """Raised when a save is blocked because the file changed since it was loaded."""
    pass


def load_state(path: Path) -> dict:
    """
    Load the current state from disk.
    The returned dict includes '_loaded_version' and '_loaded_mtime' --
    stash these (e.g. in st.session_state) and pass them back into save_state().
    """
    data = _read_json_with_retry(path)
    data["_loaded_version"] = data.get("version", 1)
    data["_loaded_mtime"] = path.stat().st_mtime
    return data


def save_state(path: Path, new_state: dict, loaded_version: int, loaded_mtime: float) -> dict:
    """
    Attempt to save new_state to disk, guarded by optimistic locking.

    Two independent checks are used:
      - version: the integer counter stored inside the JSON itself.
      - mtime: the filesystem modification time at load time.

    If either check indicates the file was changed by someone else since this
    session loaded it, raise ConcurrencyError and DO NOT write -- the caller
    (UI) should show an error telling the user to sync first.
    """
    if not path.exists():
        raise FileNotFoundError(f"State file missing at {path}")

    current_on_disk = _read_json_with_retry(path)
    current_version = current_on_disk.get("version", 1)
    current_mtime = path.stat().st_mtime

    if current_version != loaded_version or current_mtime > loaded_mtime + 0.5:
        raise ConcurrencyError(
            "The shared file has changed since you last loaded it. "
            "Click 'Sync Latest Changes' to pull the newest version, "
            "reapply your moves, and try saving again."
        )

    # Strip internal bookkeeping keys before persisting, bump version + metadata.
    to_write = {k: v for k, v in new_state.items() if not k.startswith("_")}
    to_write["version"] = current_version + 1
    to_write["last_modified"] = _now_iso()
    to_write["last_modified_by"] = _current_user_label()

    _write_json_atomic(path, to_write)

    to_write["_loaded_version"] = to_write["version"]
    to_write["_loaded_mtime"] = path.stat().st_mtime
    return to_write


def sync_latest(path: Path) -> dict:
    """
    Explicit 'Sync Latest Changes' action: simply re-reads the file fresh.
    (Actual Drive-to-local file sync is handled by the Google Drive for
    Desktop client in the background -- this just re-reads what's currently
    on disk and refreshes our loaded_version/mtime markers.)
    """
    return load_state(path)