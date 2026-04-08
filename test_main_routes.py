import json

import pytest
from openpyxl import Workbook

from app import app
from config import Config


@pytest.fixture
def main_client(tmp_path, monkeypatch):
    data_path = tmp_path / "data"
    excel_path = tmp_path / "excel"
    settings_path = data_path / "settings.json"
    config_path = tmp_path / "config.json"

    data_path.mkdir(parents=True, exist_ok=True)
    excel_path.mkdir(parents=True, exist_ok=True)

    monkeypatch.setattr(Config, "DATA_PATH", data_path)
    monkeypatch.setattr(Config, "EXCEL_BASE_PATH", excel_path)
    monkeypatch.setattr(Config, "SETTINGS_FILE_PATH", settings_path)
    monkeypatch.setattr(Config, "CONFIG_FILE_PATH", config_path)

    settings_path.write_text(json.dumps(Config.get_default_settings()), encoding="utf-8")
    (data_path / "employee_config.json").write_text(
        json.dumps({"zamestnanci": ["Test Zamestnanec"], "vybrani_zamestnanci": ["Test Zamestnanec"]}),
        encoding="utf-8",
    )

    template_path = excel_path / Config.EXCEL_TEMPLATE_NAME
    workbook = Workbook()
    weekly_sheet = workbook.active
    weekly_sheet.title = Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME
    weekly_sheet["A1"] = "Ukazka"
    advances_sheet = workbook.create_sheet(Config.EXCEL_ADVANCES_SHEET_NAME)
    advances_sheet["B80"] = Config.DEFAULT_ADVANCE_OPTION_1
    advances_sheet["D80"] = Config.DEFAULT_ADVANCE_OPTION_2
    advances_sheet["F80"] = Config.DEFAULT_ADVANCE_OPTION_3
    advances_sheet["H80"] = Config.DEFAULT_ADVANCE_OPTION_4
    workbook.save(template_path)

    app.config["TESTING"] = True
    with app.test_client() as client:
        yield client


def test_dashboard_route_is_available(main_client):
    response = main_client.get("/")

    assert response.status_code == 200
    assert "Evidence pracovní doby" in response.get_data(as_text=True)


def test_record_time_route_is_available(main_client):
    response = main_client.get("/zaznam")

    assert response.status_code == 200
    assert "Záznam pracovní doby" in response.get_data(as_text=True)


def test_quick_time_entry_api_accepts_valid_payload(main_client):
    response = main_client.post(
        "/api/quick_time_entry",
        json={
            "date": "2026-04-08",
            "start_time": "07:00",
            "end_time": "15:30",
            "lunch_duration": "0.5",
            "is_free_day": False,
        },
    )

    assert response.status_code == 200
    data = response.get_json()
    assert data["success"] is True
    assert "Pracovní doba" in data["message"]


def test_voice_command_route_processes_record_time_command(main_client):
    response = main_client.post("/voice-command", json={"command": "zapiš práci dnes od 7:00 do 15:00"})

    assert response.status_code == 200
    data = response.get_json()
    assert data["success"] is True
    assert data["entities"]["action"] == "record_time"
    assert "operation_result" in data
