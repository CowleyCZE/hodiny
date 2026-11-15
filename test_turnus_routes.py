"""Integrační testy pro routy turnusů a blokování zápisů mimo turnus."""

import json
import pytest
from openpyxl import Workbook

from app import app
from config import Config


@pytest.fixture
def turnus_client(tmp_path, monkeypatch):
    data_path = tmp_path / "data"
    excel_path = tmp_path / "excel"
    backup_path = tmp_path / "backups"
    settings_path = data_path / "settings.json"
    turnus_path = data_path / "turnusy.json"
    config_path = tmp_path / "config.json"

    data_path.mkdir(parents=True, exist_ok=True)
    excel_path.mkdir(parents=True, exist_ok=True)
    backup_path.mkdir(parents=True, exist_ok=True)

    monkeypatch.setattr(Config, "DATA_PATH", data_path)
    monkeypatch.setattr(Config, "EXCEL_BASE_PATH", excel_path)
    monkeypatch.setattr(Config, "BACKUP_PATH", backup_path)
    monkeypatch.setattr(Config, "SETTINGS_FILE_PATH", settings_path)
    monkeypatch.setattr(Config, "TURNUSY_FILE_PATH", turnus_path)
    monkeypatch.setattr(Config, "CONFIG_FILE_PATH", config_path)

    # Initial settings
    settings = Config.get_default_settings()
    settings["preferred_employee_name"] = "Jan Novak"
    settings_path.write_text(json.dumps(settings), encoding="utf-8")

    (data_path / "employee_config.json").write_text(
        json.dumps(
            {
                "zamestnanci": ["Jan Novak"],
                "vybrani_zamestnanci": ["Jan Novak"],
            }
        ),
        encoding="utf-8",
    )

    # Template workbook
    template_path = excel_path / Config.EXCEL_TEMPLATE_NAME
    workbook = Workbook()
    weekly_sheet = workbook.active
    weekly_sheet.title = Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME
    advances_sheet = workbook.create_sheet(Config.EXCEL_ADVANCES_SHEET_NAME)
    workbook.save(template_path)

    # Turnus file with an active turnus starting 2026-11-10
    turnus_data = {
        "active_turnus_id": "turnus_test",
        "turnusy": [
            {
                "id": "turnus_test",
                "name": "Test Turnus",
                "firma": "Test Firma",
                "mesto": "Test Mesto",
                "start_date": "2026-11-10",
                "end_date": None,
                "status": "active",
                "created_at": "2026-11-10T00:00:00",
                "closed_at": None,
                "archive_file": None,
            }
        ],
    }
    turnus_path.write_text(json.dumps(turnus_data), encoding="utf-8")

    app.config["TESTING"] = True
    with app.test_client() as client:
        yield client


def test_turnusy_view_renders_active_turnus(turnus_client):
    response = turnus_client.get("/turnusy")
    assert response.status_code == 200
    html = response.get_data(as_text=True)
    assert "Pracovní turnusy a výjezdy" in html
    assert "Test Turnus" in html
    assert "2026-11-10" in html


def test_quick_entry_blocked_before_turnus_start(turnus_client):
    # Pokus o zápis na 2026-11-09 (před začátkem 2026-11-10)
    response = turnus_client.post(
        "/api/quick_time_entry",
        json={
            "date": "2026-11-09",
            "start_time": "08:00",
            "end_time": "16:00",
            "lunch_duration": "1.0",
            "is_free_day": False,
        },
    )
    assert response.status_code == 400
    data = response.get_json()
    assert data["success"] is False
    assert "Datum předchází začátku aktivního turnusu" in data["error"]


def test_quick_entry_allowed_during_active_turnus(turnus_client):
    # Zápis na 2026-11-12 (v rámci turnusu)
    response = turnus_client.post(
        "/api/quick_time_entry",
        json={
            "date": "2026-11-12",
            "start_time": "08:00",
            "end_time": "16:00",
            "lunch_duration": "1.0",
            "is_free_day": False,
        },
    )
    assert response.status_code == 200
    data = response.get_json()
    assert data["success"] is True


def test_close_turnus_and_block_subsequent_entries(turnus_client):
    # Ukončíme turnus k datu 2026-11-15
    res = turnus_client.post(
        "/turnusy/ukoncit",
        data={"turnus_id": "turnus_test", "end_date": "2026-11-15"},
        follow_redirects=True,
    )
    assert res.status_code == 200

    # 1. Pokus o zápis na 2026-11-16 (po ukončení) MUSÍ BÝT ZAMÍTNUT
    res_after = turnus_client.post(
        "/api/quick_time_entry",
        json={
            "date": "2026-11-16",
            "start_time": "08:00",
            "end_time": "16:00",
            "lunch_duration": "1.0",
            "is_free_day": False,
        },
    )
    assert res_after.status_code == 400
    data_after = res_after.get_json()
    assert data_after["success"] is False
    assert "byl ukončen k 15.11.2026. Po tomto datu nelze zapisovat údaje" in data_after["error"]

    # 2. Zápis na 2026-11-14 (v rámci již uzavřeného turnusu) MUSÍ BÝT POVOLEN (doplnění)
    res_during = turnus_client.post(
        "/api/quick_time_entry",
        json={
            "date": "2026-11-14",
            "start_time": "08:00",
            "end_time": "16:00",
            "lunch_duration": "1.0",
            "is_free_day": False,
        },
    )
    assert res_during.status_code == 200
    assert res_during.get_json()["success"] is True


def test_turnus_api_endpoints(turnus_client):
    # GET /api/v1/turnus
    res = turnus_client.get("/api/v1/turnus")
    assert res.status_code == 200
    body = res.get_json()
    assert body["success"] is True
    assert body["data"]["active_turnus"]["name"] == "Test Turnus"

    # POST /api/v1/turnus/close
    res_close = turnus_client.post(
        "/api/v1/turnus/close",
        json={"end_date": "2026-11-20"},
    )
    assert res_close.status_code == 200
    assert res_close.get_json()["data"]["status"] == "closed"

    # POST /api/v1/turnus/start (nový turnus)
    res_start = turnus_client.post(
        "/api/v1/turnus/start",
        json={
            "name": "Novy Turnus Belgie",
            "start_date": "2026-11-25",
            "firma": "Dematic",
            "mesto": "Antverpy",
        },
    )
    assert res_start.status_code == 200
    assert res_start.get_json()["data"]["name"] == "Novy Turnus Belgie"
    assert res_start.get_json()["data"]["status"] == "active"
