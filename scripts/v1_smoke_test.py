"""End-to-end V1 smoke test using the existing sales analyzer."""

from __future__ import annotations

import json
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from excel_agent.manifest import append_manifest_record, build_manifest_record  # noqa: E402
from excel_agent.services.generation_service import generate_from_task_spec  # noqa: E402
from excel_agent.services.subjective_review_service import run_subjective_review  # noqa: E402
from excel_agent.services.validation_service import validate_generated_workbook  # noqa: E402
from excel_agent.task_paths import create_task_paths, stage_input_files  # noqa: E402
from excel_agent.task_spec import save_task_spec  # noqa: E402
from excel_agent.task_spec_builder import (  # noqa: E402
    build_task_spec_draft,
    merge_user_answers_into_task_spec,
)
from excel_agent.validators import inspect_workbook  # noqa: E402


def main() -> int:
    prompt = "根据订单数据做销售月报和图表"
    input_file = PROJECT_ROOT / "examples" / "input" / "sales.csv"
    task_paths = None
    try:
        draft = build_task_spec_draft(prompt, [str(input_file)])
        task_spec = draft.task_spec
        if draft.needs_clarification:
            task_spec = merge_user_answers_into_task_spec(
                task_spec,
                {
                    "task_type": "sales_report",
                    "goal_detail": "按销售月报处理，并包含图表。",
                },
            )
        task_paths = create_task_paths(task_spec.task_type)
        task_spec.input_files = stage_input_files([input_file], task_paths)
        save_task_spec(task_spec, task_paths.task_spec_file)

        generation = generate_from_task_spec(task_spec, task_paths)
        validation = validate_generated_workbook(task_paths.output_file, task_spec, task_paths)
        workbook_summary = inspect_workbook(task_paths.output_file) if task_paths.output_file.exists() else {}
        workbook_summary["sheet_count"] = len(workbook_summary.get("sheets", []))
        review = run_subjective_review(
            task_spec,
            {"status": validation.status, **validation.summary},
            workbook_summary,
            generation.to_dict(),
            task_paths,
        )
        status = validation.status if generation.success else "error"
        append_manifest_record(
            build_manifest_record(
                task_id=task_paths.task_id,
                task_type=task_spec.task_type,
                user_prompt=prompt,
                input_files=task_spec.input_files,
                output_file=str(task_paths.output_file) if task_paths.output_file.exists() else None,
                validation_report=str(task_paths.validation_report),
                status=status,
                error=generation.error,
            )
        )
        result = {
            "success": generation.success and task_paths.output_file.exists(),
            "task_id": task_paths.task_id,
            "task_dir": str(task_paths.task_dir),
            "output_file": generation.output_file,
            "validation_status": validation.status,
            "validation_report": validation.report_file,
            "subjective_review_enabled": review["enabled"],
            "error": generation.error,
        }
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["success"] and validation.status in {"pass", "warn"} else 1
    except Exception as exc:
        print(
            json.dumps(
                {
                    "success": False,
                    "task_dir": str(task_paths.task_dir) if task_paths else None,
                    "error": f"{type(exc).__name__}: {exc}",
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
