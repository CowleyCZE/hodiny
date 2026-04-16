import json

import pytest
from openpyxl import Workbook

from app import app
from config import Config


@pytest.fixture
def api_client(tmp_path, monkeypatch):
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


def test_health_endpoint_returns_success(api_client):
    response = api_client.get("/api/v1/health")

    assert response.status_code == 200
    assert response.get_json()["success"] is True


def test_employees_endpoint_returns_serialized_employees(api_client):
    response = api_client.get("/api/v1/employees")

    assert response.status_code == 200
    data = response.get_json()["data"]
    assert data == [{"name": "Test Zamestnanec", "selected": True}]


def test_selected_employees_endpoint_updates_selection(api_client):
    response = api_client.post("/api/v1/employees/selected", json={"employees": ["Test Zamestnanec"]})

    assert response.status_code == 200
    assert response.get_json()["data"] == ["Test Zamestnanec"]


def test_time_entry_endpoint_creates_entry(api_client):
    response = api_client.post(
        "/api/v1/time-entry",
        json={
            "date": "2026-04-08",
            "start_time": "07:00",
            "end_time": "15:00",
            "lunch_duration": "0.5",
            "is_free_day": False,
        },
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["success"] is True
    assert payload["data"]["employees_count"] == 1


def test_settings_endpoint_persists_updates(api_client):
    response = api_client.post("/api/v1/settings", json={"start_time": "08:00", "end_time": "17:00"})

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["data"]["start_time"] == "08:00"
    assert payload["data"]["end_time"] == "17:00"


def test_excel_status_endpoint_returns_payload(api_client):
    response = api_client.get("/api/v1/excel/status")

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["success"] is True
    assert payload["data"]["filename"]


def test_time_entries_endpoint_returns_current_week(api_client):
    response = api_client.get("/api/v1/time-entries")

    assert response.status_code == 200
    assert response.get_json()["success"] is True


def test_time_entries_endpoint_supports_employee_filter(api_client):
    api_client.post(
        "/api/v1/time-entry",
        json={
            "date": "2026-04-08",
            "start_time": "07:00",
            "end_time": "15:00",
            "lunch_duration": "0.5",
            "is_free_day": False,
        },
    )

    response = api_client.get("/api/v1/time-entries?employee=Test%20Zamestnanec&week=15")

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["success"] is True
    assert any(row and row[0] == "Test Zamestnanec" for row in payload["data"]["data"][1:])
