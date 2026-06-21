"""Command line entrypoint."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .data_cleaner import clean_table_file
from .intent_classifier import classify_intent, normalize_table_type
from .planner import plan_task
from .task_runner import run_task_file
from .validators import inspect_workbook, validate_workbook
from .workbook_builder import analyze_sales_file, create_workbook


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="excel-agent", description="Local AI Excel automation workbench")
    subparsers = parser.add_subparsers(dest="command", required=True)

    create = subparsers.add_parser("create", help="Create a workbook template")
    create.add_argument("--type", default="generic_table", help="Template type or alias")
    create.add_argument("--prompt", default=None, help="Optional natural language prompt for classification")
    create.add_argument("--output", required=True, help="Output xlsx path")

    analyze = subparsers.add_parser("analyze", help="Analyze sales CSV/XLSX data into a report")
    analyze.add_argument("--input", required=True, help="Input CSV/XLSX path")
    analyze.add_argument("--output", required=True, help="Output xlsx path")

    clean = subparsers.add_parser("clean", help="Clean CSV/XLSX data")
    clean.add_argument("--input", required=True, help="Input CSV/XLSX path")
    clean.add_argument("--output", required=True, help="Output xlsx path")

    validate = subparsers.add_parser("validate", help="Validate a workbook")
    validate.add_argument("--input", required=True, help="Workbook path")
    validate.add_argument("--json", default=None, help="Optional JSON report path")

    run_task = subparsers.add_parser("run-task", help="Run a natural-language task file")
    run_task.add_argument("--task", required=True, help="Task text file path")
    run_task.add_argument("--output", required=True, help="Output xlsx path")

    classify = subparsers.add_parser("classify", help="Classify a natural language request")
    classify.add_argument("text", nargs="+")

    inspect = subparsers.add_parser("inspect", help="Print workbook summary")
    inspect.add_argument("--input", required=True)

    args = parser.parse_args(argv)

    if args.command == "classify":
        result = classify_intent(" ".join(args.text))
        print(json.dumps(result.__dict__, ensure_ascii=False, indent=2))
        return 0

    if args.command == "inspect":
        print(json.dumps(inspect_workbook(args.input), ensure_ascii=False, indent=2))
        return 0

    if args.command == "validate":
        report = validate_workbook(args.input, args.json)
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 1 if report["status"] == "fail" else 0

    if args.command == "run-task":
        result = run_task_file(args.task, args.output)
        summary = result["validation"]["summary"]
        concise = {
            "task": result["task"],
            "validation_summary": {
                "status": result["validation"]["status"],
                "sheet_count": summary.get("sheet_count"),
                "formula_cell_count": summary.get("formula_cell_count"),
                "error_count": summary.get("error_count"),
                "warning_count": summary.get("warning_count"),
            },
            "errors": result["validation"]["errors"],
            "warnings": result["validation"]["warnings"],
        }
        print(json.dumps(concise, ensure_ascii=False, indent=2))
        return 1 if result["validation"]["status"] == "fail" else 0

    if args.command == "create":
        table_type = normalize_table_type(args.type)
        if table_type == "generic_table" and args.prompt:
            table_type = classify_intent(args.prompt).table_type
        plan = plan_task(args.prompt, table_type)
        output = create_workbook(plan.table_type, args.output)
        report = validate_workbook(output)
        payload = {"output": str(Path(output)), "plan": plan.steps, "risk_note": plan.risk_note, "validation": report}
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 1 if report["status"] == "fail" else 0

    if args.command == "analyze":
        output = analyze_sales_file(args.input, args.output)
        report = validate_workbook(output)
        print(json.dumps({"output": str(Path(output)), "validation": report}, ensure_ascii=False, indent=2))
        return 1 if report["status"] == "fail" else 0

    if args.command == "clean":
        output = clean_table_file(args.input, args.output)
        report = validate_workbook(output)
        print(json.dumps({"output": str(Path(output)), "validation": report}, ensure_ascii=False, indent=2))
        return 1 if report["status"] == "fail" else 0

    parser.print_help()
    return 2


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main(sys.argv[1:]))
