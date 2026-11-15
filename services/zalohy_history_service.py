"""Správa historie záznamů zálohy/výdaje/příjem z audit logu."""

import json
from pathlib import Path

ACTIONS = {
    "zaloha": "ADVANCE_ADDED",
    "vydaj": "VYDAJ_ADDED",
    "prijem": "PRIJEM_ADDED",
}


def save_record(audit_path: Path, section: str, data: dict) -> None:
    """Uloží záznam sekce do audit logu jako JSON řádek."""
    action = ACTIONS.get(section, f"{section.upper()}_ADDED")
    event = {"action": action, "section": section, "data": data}
    audit_path.parent.mkdir(parents=True, exist_ok=True)
    with audit_path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(event, ensure_ascii=False) + "\n")


def load_records(audit_path: Path, section: str, limit: int = 50) -> list[dict]:
    """Načte záznamy dané sekce z audit logu (nejnovější první)."""
    path = Path(audit_path)
    if not path.exists():
        return []
    action = ACTIONS.get(section, f"{section.upper()}_ADDED")
    records = []
    for line in path.read_text(encoding="utf-8").splitlines():
        try:
            event = json.loads(line)
            if event.get("action") == action:
                records.append(event.get("data", {}))
        except (json.JSONDecodeError, KeyError):
            continue
    return list(reversed(records[-limit:]))


def label_for_record(section: str, record: dict) -> str:
    """Vytvoří čitelný popisek záznamu pro zobrazení v selectu."""
    if section == "zaloha":
        return f"{record.get('date', '')} | {record.get('employee_name', '')} | {record.get('amount', '')} {record.get('currency', '')} | {record.get('option', '')}"
    if section == "vydaj":
        return f"{record.get('date', '')} | {record.get('category', '')} | {record.get('amount', '')} {record.get('currency', '')}"
    if section == "prijem":
        source = record.get('source', '')
        colleague = record.get('colleague', '')
        label = f"{record.get('date', '')} | {source}"
        if colleague:
            label += f" ({colleague})"
        label += f" | {record.get('amount', '')} {record.get('currency', '')}"
        return label
    return str(record)
