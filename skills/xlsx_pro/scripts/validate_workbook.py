from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT / "src"))

from excel_agent.validators import validate_workbook  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate an Excel workbook and print JSON.")
    parser.add_argument("input", help="Workbook path")
    parser.add_argument("--json", default=None, help="Optional JSON output path")
    args = parser.parse_args()
    report = validate_workbook(args.input, args.json)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 1 if report["status"] == "fail" else 0


if __name__ == "__main__":
    raise SystemExit(main())

