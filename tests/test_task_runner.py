from excel_agent.task_runner import run_task_file


def test_run_task_file_sales_with_input(tmp_path):
    output = tmp_path / "task_sales.xlsx"
    result = run_task_file("examples/tasks/task_sales.txt", output)
    assert result["task"]["table_type"] == "sales_report"
    assert result["task"]["used_input_file"] is True
    assert result["task"]["validation_status"] == "pass"
    assert output.exists()

