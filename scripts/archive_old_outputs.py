"""Move legacy root-level outputs into outputs/legacy without deleting files."""

from __future__ import annotations

import argparse
import shutil
from datetime import datetime
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUTS = PROJECT_ROOT / "outputs"
LEGACY = OUTPUTS / "legacy"
ARCHIVE_SUFFIXES = {
    ".xlsx",
    ".xlsm",
    ".xls",
    ".csv",
    ".json",
    ".md",
    ".html",
    ".png",
    ".pdf",
}
RESERVED_NAMES = {"manifest.json", ".gitkeep"}


def unique_destination(source: Path) -> Path:
    candidate = LEGACY / source.name
    if not candidate.exists():
        return candidate
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    counter = 1
    while True:
        candidate = LEGACY / f"{source.stem}_{stamp}_{counter}{source.suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def archive_old_outputs(dry_run: bool = False) -> list[tuple[Path, Path]]:
    LEGACY.mkdir(parents=True, exist_ok=True)
    moved: list[tuple[Path, Path]] = []
    for source in sorted(OUTPUTS.iterdir()):
        if not source.is_file():
            continue
        if source.name in RESERVED_NAMES or source.suffix.lower() not in ARCHIVE_SUFFIXES:
            continue
        destination = unique_destination(source)
        moved.append((source, destination))
        if not dry_run:
            shutil.move(str(source), str(destination))
    return moved


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dry-run", action="store_true", help="Only print planned moves")
    args = parser.parse_args()
    moved = archive_old_outputs(args.dry_run)
    action = "WOULD MOVE" if args.dry_run else "MOVED"
    for source, destination in moved:
        print(f"{action}: {source} -> {destination}")
    print(f"{action} COUNT: {len(moved)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
