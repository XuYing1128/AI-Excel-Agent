from skills.registry import match_skills

from excel_agent.task_spec import TaskSpec


def test_match_schedule_skill_for_exam_prompt():
    spec = TaskSpec(
        task_type="schedule",
        user_goal="三个教室两天排考，保证同一考生同一时间不能冲突。",
    )
    names = [skill.name for skill in match_skills(spec)]
    assert "schedule" in names


def test_match_perf_review_skill_for_salary_prompt():
    spec = TaskSpec(
        task_type="finance_model",
        user_goal="生成员工绩效评估及薪酬调整明细表，包含缺勤扣分和调薪比例。",
    )
    names = [skill.name for skill in match_skills(spec)]
    assert names[0] == "perf_review"


def test_match_data_clean_skill_for_deduplicate_prompt():
    spec = TaskSpec(
        task_type="generic_table",
        user_goal="合并两个名单，按姓名去重，并输出缺失值和异常报告。",
    )
    names = [skill.name for skill in match_skills(spec)]
    assert "data_clean" in names

