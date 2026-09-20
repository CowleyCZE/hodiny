"""Služby fáze 2: statistiky, kontroly, zálohy a audit."""

import json
import shutil
from datetime import datetime, timedelta
from pathlib import Path


def build_statistics(calendar_data, report_data, month, year):
    """Sestaví souhrnné statistiky z denních a zaměstnaneckých agregací."""
    total_hours = round(sum(float(item.get("total_hours", 0)) for item in report_data.values()), 2)
    free_days = sum(int(item.get("free_days", 0)) for item in report_data.values())
    working_days = sum(1 for item in calendar_data.values() if item.get("hours", 0) > 0)
    average = round(total_hours / working_days, 2) if working_days else 0.0
    return {
        "month": month,
        "year": year,
        "total_hours": total_hours,
        "working_days": working_days,
        "free_days": free_days,
        "average_hours_per_day": average,
        "days_with_entries": len(calendar_data),
        "employees": report_data,
        "daily": calendar_data,
    }


def find_data_issues(calendar_data, employees, period_start, period_end):
    """Najde chybějící pracovní dny a podezřele dlouhé souhrny."""
    issues = []
    employee_count = len(employees)
    current = period_start
    while current <= period_end:
        if current.weekday() < 5:
            key = current.isoformat()
            item = calendar_data.get(key, {})
            if employee_count and not item.get("has_entry"):
                issues.append(
                    {
                        "severity": "warning",
                        "code": "MISSING_ENTRY",
                        "date": key,
                        "message": f"Chybí záznam pro {current:%d.%m.%Y}.",
                    }
                )
            if item.get("hours", 0) > 24:
                issues.append(
                    {
                        "severity": "error",
                        "code": "EXCESSIVE_HOURS",
                        "date": key,
                        "message": f"Součet za den {current:%d.%m.%Y} přesahuje 24 hodin.",
                    }
                )
        current += timedelta(days=1)
    return issues


def write_audit(audit_path, action, details=None, actor="local"):
    """Zapíše auditní událost do JSONL souboru."""
    audit_path = Path(audit_path)
    audit_path.parent.mkdir(parents=True, exist_ok=True)
    event = {
        "timestamp": datetime.now().astimezone().isoformat(),
        "action": action,
        "actor": actor,
        "details": details or {},
    }
    with audit_path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(event, ensure_ascii=False) + "\n")
    try:
        from flask import g, has_request_context

        if has_request_context() and hasattr(g, "database"):
            g.database.record_audit(action, actor, json.dumps(details or {}, ensure_ascii=False))
    except (RuntimeError, AttributeError, OSError):
        pass
    return event


def read_audit(audit_path, limit=100):
    """Načte poslední auditní události."""
    path = Path(audit_path)
    if not path.exists():
        return []
    lines = path.read_text(encoding="utf-8").splitlines()[-limit:]
    events = []
    for line in reversed(lines):
        try:
            events.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    return events


def create_backup(source_path, backup_dir, reason="manual"):
    """Vytvoří kopii XLSX souboru s časovým razítkem."""
    source = Path(source_path)
    if not source.exists():
        raise FileNotFoundError(source)
    target_dir = Path(backup_dir)
    target_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target = target_dir / f"{source.stem}_{stamp}_{reason}{source.suffix}"
    shutil.copy2(source, target)
    return target


def list_backups(backup_dir):
    """Vrátí zálohy seřazené od nejnovější."""
    directory = Path(backup_dir)
    if not directory.exists():
        return []
    return sorted(
        [
            {
                "name": path.name,
                "size": path.stat().st_size,
                "created": datetime.fromtimestamp(path.stat().st_mtime).astimezone().isoformat(),
            }
            for path in directory.glob("*.xlsx")
        ],
        key=lambda item: item["created"],
        reverse=True,
    )


def restore_backup(backup_dir, backup_name, target_path):
    """Bezpečně obnoví povolenou zálohu do aktivního souboru."""
    backup_root = Path(backup_dir).resolve()
    source = (backup_root / backup_name).resolve()
    if source.parent != backup_root or source.suffix.lower() != ".xlsx" or not source.exists():
        raise ValueError("Neplatná záloha.")
    target = Path(target_path)
    temp = target.with_suffix(target.suffix + ".restore.tmp")
    shutil.copy2(source, temp)
    temp.replace(target)
    return target
