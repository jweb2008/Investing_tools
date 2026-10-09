"""
Weekly backup to OneDrive.

  py backup.py

Copies the master fill record and your newest tracker version into a dated
folder, e.g. OneDrive\\TradingBackups\\2026-10-10\\, and keeps the most recent
KEEP_WEEKS folders. Runs automatically at the end of auth_setup.py.

Never copies .env or schwab_token.json; those stay on this computer only.
"""
import datetime as dt
import os
import re
import shutil
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

HERE = Path(__file__).parent
TRACKER_DIR = Path(os.getenv("TRACKER_DIR", str(HERE)))
PREFIX = os.getenv("TRACKER_PREFIX", "2025-26_Trading_Tracker_")
KEEP_WEEKS = int(os.getenv("BACKUP_KEEP_WEEKS", "12"))
NEVER_COPY = {".env", "schwab_token.json"}


def backup_root():
    if os.getenv("BACKUP_DIR"):
        return Path(os.environ["BACKUP_DIR"])
    onedrive = os.getenv("OneDrive") or str(Path.home() / "OneDrive")
    return Path(onedrive) / "TradingBackups"


def main():
    root = backup_root()
    if not root.parent.exists():
        print(f"Backup skipped: OneDrive folder not found at {root.parent}. "
              "Set BACKUP_DIR in .env to the right folder.")
        return

    trackers = [p for p in TRACKER_DIR.glob(f"{PREFIX}*.xlsx") if not p.name.startswith("~$")]
    files = [HERE / "fills_cache.csv"]
    if trackers:
        files.append(max(trackers, key=lambda p: p.stat().st_mtime))
    files = [f for f in files if f.exists() and f.name not in NEVER_COPY]
    if not files:
        print("Backup skipped: nothing to back up yet.")
        return

    dest = root / dt.date.today().isoformat()
    dest.mkdir(parents=True, exist_ok=True)
    for f in files:
        shutil.copy2(f, dest / f.name)

    # keep only the most recent KEEP_WEEKS dated folders
    dated = sorted(p for p in root.iterdir() if p.is_dir() and re.fullmatch(r"\d{4}-\d{2}-\d{2}", p.name))
    for old in dated[:-KEEP_WEEKS]:
        shutil.rmtree(old, ignore_errors=True)

    print(f"Backed up {', '.join(f.name for f in files)} to {dest}")


if __name__ == "__main__":
    main()
