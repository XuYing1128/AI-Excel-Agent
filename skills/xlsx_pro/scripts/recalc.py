from __future__ import annotations

import argparse
import platform
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def find_soffice() -> str | None:
    discovered = shutil.which("soffice") or shutil.which("libreoffice")
    if discovered:
        return discovered
    if platform.system().lower() != "windows":
        return None
    candidates = (
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    )
    return str(next((item for item in candidates if item.exists()), "")) or None


def recalc_with_libreoffice(path: Path) -> bool:
    soffice = find_soffice()
    if not soffice:
        return False
    with tempfile.TemporaryDirectory(prefix="ai_excel_recalc_") as temporary:
        output_dir = Path(temporary)
        completed = subprocess.run(
            [
                soffice,
                "--headless",
                "--convert-to",
                "xlsx",
                "--outdir",
                str(output_dir),
                str(path.resolve()),
            ],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        converted = output_dir / path.name
        if not converted.exists() or converted.stat().st_size == 0:
            detail = (completed.stdout or completed.stderr).strip()
            raise RuntimeError(f"LibreOffice 未产生重算文件：{detail}")
        local_temporary = path.with_name(f".{path.stem}.recalc{path.suffix}")
        shutil.copy2(converted, local_temporary)
        local_temporary.replace(path)
    return True


def recalc_with_excel_com(path: Path) -> bool:
    if platform.system().lower() != "windows":
        return False
    try:
        import win32com.client  # type: ignore
    except Exception:
        return False
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        workbook = excel.Workbooks.Open(str(path.resolve()))
        workbook.RefreshAll()
        excel.CalculateFullRebuild()
        workbook.Save()
        workbook.Close(SaveChanges=True)
    finally:
        excel.Quit()
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description="Recalculate formulas using LibreOffice or Excel COM when available.")
    parser.add_argument("input", help="Workbook path")
    args = parser.parse_args()
    path = Path(args.input)
    if not path.exists():
        print(f"文件不存在: {path}", file=sys.stderr)
        return 1
    try:
        if recalc_with_libreoffice(path):
            print(f"已使用 LibreOffice headless 重算并保存: {path}")
            return 0
    except Exception as exc:
        print(f"LibreOffice 重算失败: {exc}", file=sys.stderr)
    try:
        if recalc_with_excel_com(path):
            print(f"已使用 Windows Excel COM 重算并保存: {path}")
            return 0
    except Exception as exc:
        print(f"Excel COM 重算失败: {exc}", file=sys.stderr)
    print("已写入公式但未完成真实 Excel 引擎重算；将仅依赖静态校验。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
