from datetime import date

from employee_management import EmployeeManager
from services.export_service import build_csv, build_pdf
from services.phase2_services import (
    build_statistics,
    create_backup,
    find_data_issues,
    list_backups,
    read_audit,
    restore_backup,
    write_audit,
)


def test_exports_have_expected_formats():
    report = {"Jan Test": {"total_hours": 8.5, "free_days": 1}}
    assert b"Jan Test" in build_csv(report, 9, 2026)
    assert build_pdf(report, 9, 2026).startswith(b"%PDF")


def test_statistics_and_issue_detection():
    daily = {"2026-09-01": {"hours": 8.0, "has_entry": True}}
    report = {"Jan Test": {"total_hours": 8.0, "free_days": 0}}
    stats = build_statistics(daily, report, 9, 2026)
    assert stats["total_hours"] == 8.0
    issues = find_data_issues(daily, ["Jan Test"], date(2026, 9, 1), date(2026, 9, 2))
    assert any(issue["code"] == "MISSING_ENTRY" for issue in issues)


def test_backup_restore_and_audit(tmp_path):
    source = tmp_path / "active.xlsx"
    source.write_bytes(b"original")
    backup_dir = tmp_path / "backups"
    backup = create_backup(source, backup_dir, "test")
    source.write_bytes(b"changed")
    restore_backup(backup_dir, backup.name, source)
    assert source.read_bytes() == b"original"
    assert len(list_backups(backup_dir)) == 1
    audit_path = tmp_path / "audit.jsonl"
    write_audit(audit_path, "TEST", {"ok": True})
    assert read_audit(audit_path)[0]["action"] == "TEST"


def test_employee_can_be_deactivated_without_deletion(tmp_path):
    manager = EmployeeManager(tmp_path)
    assert manager.pridat_zamestnance("Jan Test") is True
    assert manager.deactivate_employee("Jan Test") is True
    assert manager.get_employee_names() == ["Jan Test"]
    assert manager.get_active_employee_names() == []
    assert manager.get_all_employees()[0]["active"] is False
    assert manager.activate_employee("Jan Test") is True
    assert manager.get_active_employee_names() == ["Jan Test"]
