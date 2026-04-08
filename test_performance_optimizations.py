from performance_optimizations import (
    get_employee_stats,
    get_excel_file_info,
    initialize_performance_optimizations,
)


def test_context_dependent_performance_helpers_return_safe_defaults_without_request_context():
    employee_stats = get_employee_stats()
    excel_info = get_excel_file_info()

    assert employee_stats == {
        "total_employees": 0,
        "selected_employees": 0,
        "selection_percentage": 0,
    }
    assert excel_info["exists"] is False
    assert excel_info["filename"] == "Unknown"
    assert "last_checked" in excel_info


def test_initialize_performance_optimizations_is_safe_without_context():
    initialize_performance_optimizations()
