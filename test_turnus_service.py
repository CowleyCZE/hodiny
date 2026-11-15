"""Jednotkové testy pro turnus_service."""

import json
from pathlib import Path
import pytest

from services.turnus_service import (
    close_turnus,
    create_turnus,
    get_active_turnus,
    get_all_turnusy,
    get_turnus_by_id,
    load_turnus_data,
    save_turnus_data,
    validate_date_for_entry,
)


@pytest.fixture
def turnus_file(tmp_path):
    file_path = tmp_path / "turnusy.json"
    initial_data = {
        "active_turnus_id": None,
        "turnusy": [],
    }
    file_path.write_text(json.dumps(initial_data), encoding="utf-8")
    return file_path


def test_create_and_get_active_turnus(turnus_file):
    new_t = create_turnus(
        name="Montaz Rotterdam",
        start_date="2026-11-01",
        end_date=None,
        firma="Stow",
        mesto="Rotterdam",
        filepath=turnus_file,
    )
    assert new_t["name"] == "Montaz Rotterdam"
    assert new_t["status"] == "active"
    assert new_t["start_date"] == "2026-11-01"

    active = get_active_turnus(turnus_file)
    assert active is not None
    assert active["id"] == new_t["id"]


def test_create_turnus_validations(turnus_file):
    with pytest.raises(ValueError, match="Název turnusu je povinný"):
        create_turnus("", "2026-11-01", filepath=turnus_file)

    with pytest.raises(ValueError, match="Datum začátku turnusu je povinné"):
        create_turnus("Test", "", filepath=turnus_file)

    with pytest.raises(ValueError, match="Datum začátku musí mít formát"):
        create_turnus("Test", "neplatne-datum", filepath=turnus_file)

    with pytest.raises(ValueError, match="Datum ukončení nesmí být před datem začátku"):
        create_turnus("Test", "2026-11-10", end_date="2026-11-05", filepath=turnus_file)


def test_cannot_create_when_already_active(turnus_file):
    create_turnus("Turnus 1", "2026-11-01", filepath=turnus_file)
    with pytest.raises(ValueError, match="Stále běží aktivní turnus 'Turnus 1'"):
        create_turnus("Turnus 2", "2026-11-15", filepath=turnus_file)


def test_close_turnus(turnus_file, tmp_path):
    t = create_turnus("Turnus Blankendaal", "2026-11-01", filepath=turnus_file)

    excel_dummy = tmp_path / "Hodiny_Cap.xlsx"
    excel_dummy.write_text("dummy", encoding="utf-8")
    backup_dir = tmp_path / "backups"

    closed = close_turnus(
        turnus_id=t["id"],
        end_date="2026-11-20",
        filepath=turnus_file,
        excel_path=excel_dummy,
        backup_dir=backup_dir,
    )
    assert closed["status"] == "closed"
    assert closed["end_date"] == "2026-11-20"
    assert get_active_turnus(turnus_file) is None

    # Ověření vytvoření zálohy v backup_dir
    backups = list(backup_dir.glob("Turnus_*.xlsx"))
    assert len(backups) == 1
    assert "Turnus_Blankendaal_2026-11-01_2026-11-20.xlsx" in backups[0].name


def test_validate_date_for_entry_active_turnus(turnus_file):
    create_turnus("Turnus A", "2026-11-10", filepath=turnus_file)

    # Datum v rámci aktivního turnusu
    ok, msg, t = validate_date_for_entry("2026-11-10", filepath=turnus_file)
    assert ok is True
    assert msg == ""
    assert t["name"] == "Turnus A"

    ok, msg, t = validate_date_for_entry("2026-11-15", filepath=turnus_file)
    assert ok is True

    # Datum před začátkem aktivního turnusu
    ok, msg, t = validate_date_for_entry("2026-11-09", filepath=turnus_file)
    assert ok is False
    assert "Datum předchází začátku aktivního turnusu 'Turnus A'" in msg


def test_validate_date_for_entry_closed_turnus_allows_historical_edits(turnus_file):
    # Založíme a uzavřeme turnus: 01.11.2026 až 15.11.2026
    t = create_turnus("Turnus Minuly", "2026-11-01", filepath=turnus_file)
    close_turnus(turnus_id=t["id"], end_date="2026-11-15", filepath=turnus_file)

    # 1. Datum v rámci ukončeného turnusu je POVOLENO (uživatel si přeje možnost doplnit/upravit)
    ok, msg, t_found = validate_date_for_entry("2026-11-05", filepath=turnus_file)
    assert ok is True
    assert t_found["id"] == t["id"]

    ok, msg, t_found = validate_date_for_entry("2026-11-15", filepath=turnus_file)
    assert ok is True

    # 2. Datum PO ukončení turnusu je STRIKTNĚ ZAKÁZÁNO
    ok, msg, t_found = validate_date_for_entry("2026-11-16", filepath=turnus_file)
    assert ok is False
    assert "byl ukončen k 15.11.2026. Po tomto datu nelze zapisovat údaje" in msg

    # 3. Datum před zahájením turnusu je zakázáno
    ok, msg, t_found = validate_date_for_entry("2026-10-31", filepath=turnus_file)
    assert ok is False
