from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT / "src"))

from excel_agent.validators import inspect_workbook  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Render workbook preview when OfficeCLI is available; otherwise print summary.")
    parser.add_argument("input", help="Workbook path")
    parser.add_argument("--output", default=None, help="Optional preview output path")
    args = parser.parse_args()
    officecli = shutil.which("officecli")
    if officecli:
        cmd = [officecli, "preview", args.input]
        if args.output:
            cmd.extend(["--output", args.output])
        subprocess.run(cmd, check=True)
        return 0
    print("未检测到 OfficeCLI，输出 workbook summary：")
    print(inspect_workbook(args.input))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

