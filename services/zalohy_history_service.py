"""Správa historie záznamů zálohy/výdaje/příjem z audit logu.

Každý záznam dostane unikátní record_id (timestamp-based).
Při aktualizaci se uloží nový řádek se stejným record_id –
load_records pak pro každé record_id vrátí jen nejnovější verzi.
"""

import json
import uuid
from pathlib import Path

ACTIONS = {
    "zaloha": "ADVANCE_ADDED",
    "vydaj": "VYDAJ_ADDED",
    "prijem": "PRIJEM_ADDED",
}


def _new_id() -> str:
    return uuid.uuid4().hex[:12]


def save_record(audit_path: Path, section: str, data: dict, record_id: str | None = None) -> str:
    """Uloží nebo aktualizuje záznam sekce v audit logu.

    Pokud je record_id zadáno, přepíše existující záznam (nový řádek
    se stejným ID – load_records vrátí vždy nejnovější verzi).
    Vrátí record_id záznamu.
    """
    if not record_id:
        record_id = _new_id()
    action = ACTIONS.get(section, f"{section.upper()}_ADDED")
    data_with_id = dict(data, record_id=record_id)
    event = {"action": action, "section": section, "data": data_with_id}
    audit_path.parent.mkdir(parents=True, exist_ok=True)
    with audit_path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(event, ensure_ascii=False) + "\n")
    return record_id


def load_records(audit_path: Path, section: str, limit: int = 50) -> list[dict]:
    """Načte záznamy dané sekce (nejnovější verze každého record_id, nejnovější první)."""
    path = Path(audit_path)
    if not path.exists():
        return []
    action = ACTIONS.get(section, f"{section.upper()}_ADDED")
    # Uchovej jen poslední výskyt každého record_id (přepisuje starší)
    seen: dict[str, dict] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        try:
            event = json.loads(line)
            if event.get("action") != action:
                continue
            d = event.get("data", {})
            rid = d.get("record_id", _new_id())
            seen[rid] = d  # novější řádek vždy přepíše starší
        except (json.JSONDecodeError, KeyError):
            continue
    # Vrať nejnovějších `limit` záznamů (pořadí: nejnovější první = reversed insertion order)
    ordered = list(reversed(list(seen.values())))
    return ordered[:limit]


def get_record_by_id(audit_path: Path, section: str, record_id: str) -> dict | None:
    """Najde poslední verzi záznamu podle record_id."""
    if not record_id:
        return None
    records = load_records(audit_path, section, limit=1000)
    for r in records:
        if r.get("record_id") == record_id:
            return r
    return None


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
