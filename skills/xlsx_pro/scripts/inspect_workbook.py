from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT / "src"))

from excel_agent.validators import inspect_workbook  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Print workbook summary.")
    parser.add_argument("input", help="Workbook path")
    args = parser.parse_args()
    print(json.dumps(inspect_workbook(args.input), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

