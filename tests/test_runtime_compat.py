import inspect

from excel_agent.services.runtime_compat import load_generation_service


def test_generation_runtime_contract_is_current():
    module = load_generation_service()
    signature = inspect.signature(module.generate_from_task_spec)

    assert module.GENERATION_API_VERSION >= 2
    assert "api_settings" in signature.parameters
    assert "progress" in signature.parameters
