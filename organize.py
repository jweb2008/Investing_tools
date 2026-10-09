"""
Tidies the Investing folder. Safe to run any time; nothing is deleted.

  py organize.py           preview: lists what would move, changes nothing
  py organize.py --apply   do it

Main folder keeps what you use day to day:
  newest tracker version, newest Review file, scanner_log.xlsx,
  trade_log.xlsx, watchlist.txt, the scripts, .env and the Schwab token

Moves
  older tracker versions     -> Tracker Archive\\
  older Review_*.xlsx files  -> Reviews\\
  *.zip                      -> Archive\\
  scanner data files         -> data\\   (universe list, scan log CSV, saved pages)

Stays put on purpose: fills_cache.csv and update_log.txt. The trade feed and
its scheduled task expect them in the main folder.

Skips any file that is open in Excel (~$ lock file present) and never
overwrites a file that already exists at the destination.
"""
import argparse
import os
import shutil
import sys
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

HERE = Path(__file__).parent
TRACKER_DIR = Path(os.getenv("TRACKER_DIR", str(HERE)))
PREFIX = os.getenv("TRACKER_PREFIX", "2025-26_Trading_Tracker_")

SCANNER_DATA = ["universe.txt", "universe_info.csv", "scan_log.csv", "universe_pages"]


def newest(paths):
    return max(paths, key=lambda p: p.stat().st_mtime) if paths else None


def is_open(p: Path):
    return (p.parent / f"~${p.name}").exists() or (p.parent / f"~${p.name[2:]}").exists()


def plan():
    moves = []

    trackers = [p for p in TRACKER_DIR.glob(f"{PREFIX}*.xlsx") if not p.name.startswith("~$")]
    keep = newest(trackers)
    moves += [(p, TRACKER_DIR / "Tracker Archive") for p in trackers if p != keep]

    reviews = [p for p in HERE.glob("Review_*.xlsx") if not p.name.startswith("~$")]
    keep_r = newest(reviews)
    moves += [(p, HERE / "Reviews") for p in reviews if p != keep_r]

    moves += [(p, HERE / "Archive") for p in HERE.glob("*.zip")]

    moves += [(HERE / n, HERE / "data") for n in SCANNER_DATA if (HERE / n).exists()]
    return moves, keep, keep_r


def main():
    ap = argparse.ArgumentParser(description="Tidy the Investing folder")
    ap.add_argument("--apply", action="store_true", help="move the files (default: preview only)")
    args = ap.parse_args()

    moves, keep, keep_r = plan()
    print(f"Keeping in main folder: {keep.name if keep else '(no tracker found)'}"
          + (f", {keep_r.name}" if keep_r else ""))
    if not moves:
        print("Nothing to move. Folder is already tidy.")
        return

    print(f"\n{'Moving' if args.apply else 'Would move'} {len(moves)} item(s):")
    moved = skipped = 0
    for src, dest_dir in moves:
        dest = dest_dir / src.name
        label = f"  {src.name}  ->  {dest_dir.name}\\"
        if src.is_file() and is_open(src):
            print(f"{label}   SKIPPED: open in Excel")
            skipped += 1
            continue
        if dest.exists():
            print(f"{label}   SKIPPED: already exists there")
            skipped += 1
            continue
        print(label)
        if args.apply:
            dest_dir.mkdir(exist_ok=True)
            shutil.move(str(src), str(dest))
            moved += 1

    if args.apply:
        print(f"\nDone: {moved} moved, {skipped} skipped.")
    else:
        print("\nPreview only. Run  py organize.py --apply  to move them.")


if __name__ == "__main__":
    sys.exit(main())
