# camp_sync.py
"""
Backend module for Manual Group Alignment.
Handles: Google Drive path discovery, JSON state load/save,
multi-"camp" (independent group sets, e.g. Y8 2026 / Y9 2026) support,
and per-camp optimistic-lock concurrency checks.
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
DEFAULT_CAMP_NAME = "Default"


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
        _write_json_atomic(state_path, _default_file())

    return state_path


# ============================================================================
# 2. JSON SCHEMA + DEFAULTS
# ============================================================================
#
# The shared file holds MULTIPLE independent "camps" (e.g. "Y8 2026",
# "Y9 2026") in one JSON file, so different staff can work on different
# camps without blocking each other's saves:
#
# {
#   "camps": {
#     "Y8 2026": {
#       "version": 3, "last_modified": "...", "last_modified_by": "...",
#       "groups":   { "Bay of Fires": ["S001", "S002"], ... },
#       "students": { "S001": {name, friend_requests, form_data, medical_flags}, ... }
#     },
#     "Y9 2026": { ... }
#   }
# }

def _default_camp() -> dict:
    return {
        "version": 1,                    # bumped on every successful save OF THIS CAMP
        "last_modified": _now_iso(),
        "last_modified_by": _current_user_label(),
        "groups": {},
        "students": {},
    }


def _default_file() -> dict:
    return {"camps": {DEFAULT_CAMP_NAME: _default_camp()}}


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _current_user_label() -> str:
    """A human-readable 'who saved this' label -- OS username, since we don't have real auth."""
    try:
        return f"{getpass.getuser()} ({platform.node()})"
    except Exception:
        return "unknown"


def _migrate_legacy_schema(raw: dict) -> dict:
    """
    Older versions of this app wrote a flat {version, groups, students, ...}
    file with no "camps" wrapper. If we detect that shape, wrap it into a
    single camp so nothing is lost, and mark it dirty so the caller
    persists the migrated shape back to disk.
    """
    if "camps" in raw:
        return raw
    if "groups" in raw or "students" in raw:
        migrated_camp = {
            "version": raw.get("version", 1),
            "last_modified": raw.get("last_modified", _now_iso()),
            "last_modified_by": raw.get("last_modified_by", _current_user_label()),
            "groups": raw.get("groups", {}),
            "students": raw.get("students", {}),
        }
        return {"camps": {DEFAULT_CAMP_NAME: migrated_camp}}
    # Genuinely empty/unrecognised file -- start fresh.
    return _default_file()


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


def _load_full_file(path: Path) -> dict:
    """
    Read the whole shared file, migrating legacy schema if needed.
    If migration happened, the fixed-up shape is written straight back to
    disk so every other client also sees the new "camps" structure.
    """
    raw = _read_json_with_retry(path)
    migrated = _migrate_legacy_schema(raw)
    if migrated is not raw:
        _write_json_atomic(path, migrated)
    return migrated


# ============================================================================
# 4. CAMP DISCOVERY / CREATION
# ============================================================================

def list_camp_names(path: Path) -> list[str]:
    """Return the names of all camps currently in the shared file, sorted."""
    full = _load_full_file(path)
    return sorted(full["camps"].keys())


def create_camp(path: Path, camp_name: str) -> None:
    """Add a new, empty camp to the shared file (no-op if it already exists)."""
    camp_name = camp_name.strip()
    if not camp_name:
        raise ValueError("Camp name can't be empty.")
    full = _load_full_file(path)
    if camp_name not in full["camps"]:
        full["camps"][camp_name] = _default_camp()
        _write_json_atomic(path, full)


# ============================================================================
# 5. LOAD / SAVE ONE CAMP, WITH PER-CAMP OPTIMISTIC LOCKING
# ============================================================================

class ConcurrencyError(Exception):
    """Raised when a save is blocked because this camp changed since it was loaded."""
    pass


def load_camp(path: Path, camp_name: str) -> dict:
    """
    Load a single camp's data from the shared file.
    The returned dict includes '_loaded_version' and '_loaded_mtime' --
    stash these (e.g. in st.session_state) and pass them back into save_camp().
    """
    full = _load_full_file(path)
    if camp_name not in full["camps"]:
        create_camp(path, camp_name)
        full = _load_full_file(path)

    camp = full["camps"][camp_name]
    camp["_loaded_version"] = camp.get("version", 1)
    camp["_loaded_mtime"] = path.stat().st_mtime
    return camp


def peek_camp_version(path: Path, camp_name: str) -> int | None:
    """
    Cheap check: read the file and return this camp's CURRENT version on disk,
    without touching session state. Used to show a staleness banner
    ("a newer version is available") without forcing a full reload.
    Returns None if the camp doesn't exist (e.g. was never created, or
    deleted by someone else).
    """
    full = _load_full_file(path)
    camp = full["camps"].get(camp_name)
    return camp.get("version") if camp else None


def save_camp(path: Path, camp_name: str, new_camp_data: dict,
              loaded_version: int, loaded_mtime: float) -> dict:
    """
    Attempt to save one camp's data to disk, guarded by optimistic locking
    scoped to THIS CAMP ONLY -- edits to other camps by other staff, made
    since this camp was loaded, are preserved untouched.

    Re-reads the full file fresh right before writing, so:
      - the concurrency check is against the true current state, and
      - any other camp's concurrent changes are carried forward rather
        than overwritten by a stale in-memory copy.
    """
    if not path.exists():
        raise FileNotFoundError(f"State file missing at {path}")

    full = _load_full_file(path)
    current_camp = full["camps"].get(camp_name)
    current_version = current_camp.get("version", 1) if current_camp else 0

    if current_camp is None or current_version != loaded_version:
        raise ConcurrencyError(
            f"'{camp_name}' has changed since you last loaded it "
            "(someone else saved, or it was renamed/removed). "
            "Click 'Sync Latest Changes' to pull the newest version, "
            "reapply your moves, and try saving again."
        )

    # Strip internal bookkeeping keys before persisting, bump version + metadata.
    to_write_camp = {k: v for k, v in new_camp_data.items() if not k.startswith("_")}
    to_write_camp["version"] = current_version + 1
    to_write_camp["last_modified"] = _now_iso()
    to_write_camp["last_modified_by"] = _current_user_label()

    full["camps"][camp_name] = to_write_camp
    _write_json_atomic(path, full)

    to_write_camp["_loaded_version"] = to_write_camp["version"]
    to_write_camp["_loaded_mtime"] = path.stat().st_mtime
    return to_write_camp


def sync_camp(path: Path, camp_name: str) -> dict:
    """
    Explicit 'Sync Latest Changes' action for the current camp: re-reads
    the file fresh and refreshes loaded_version/mtime markers.
    (Actual Drive-to-local file sync is handled by the Google Drive for
    Desktop client in the background -- this just re-reads what's
    currently on disk.)
    """
    return load_camp(path, camp_name)