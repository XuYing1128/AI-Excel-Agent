from excel_agent.manifest import (
    append_manifest_record,
    build_manifest_record,
    load_manifest,
    recent_tasks,
    update_manifest_record,
)


def test_manifest_append_update_and_recent(tmp_path):
    path = tmp_path / "manifest.json"
    record = build_manifest_record(
        task_id="20260621_153012_sales_report",
        task_type="sales_report",
        user_prompt="做销售月报",
        input_files=["orders.csv"],
        output_file="result.xlsx",
        validation_report="validation.json",
        status="pass",
        error=None,
    )
    append_manifest_record(record, path)
    loaded = load_manifest(path)
    assert len(loaded["tasks"]) == 1
    assert loaded["tasks"][0]["task_id"] == record["task_id"]

    update_manifest_record(record["task_id"], {"status": "warn"}, path)
    assert recent_tasks(1, path)[0]["status"] == "warn"
