from datetime import datetime

from excel_agent.task_paths import create_task_paths, save_uploaded_bytes


def test_create_unique_task_paths(tmp_path):
    now = datetime(2026, 6, 21, 15, 30, 12)
    first = create_task_paths("sales_report", tmp_path, now)
    second = create_task_paths("sales_report", tmp_path, now)

    assert first.task_id == "20260621_153012_sales_report"
    assert second.task_id != first.task_id
    assert first.input_dir.is_dir()
    assert first.output_dir.is_dir()
    assert first.reports_dir.is_dir()
    assert first.output_file.name == "result.xlsx"

    uploaded = save_uploaded_bytes("orders.csv", b"a,b\n1,2\n", first)
    assert uploaded.exists()
    assert uploaded.parent == first.input_dir.resolve()
