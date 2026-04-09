import json

import pytest
from openpyxl import Workbook

from app import app
from config import Config


@pytest.fixture
def route_client(tmp_path, monkeypatch):
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

    settings = Config.get_default_settings()
    settings["preferred_employee_name"] = "Jan Test"
    settings_path.write_text(json.dumps(settings), encoding="utf-8")
    (data_path / "employee_config.json").write_text(
        json.dumps(
            {
                "zamestnanci": ["Alpha Worker", "Jan Test"],
                "vybrani_zamestnanci": ["Alpha Worker"],
            }
        ),
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


def test_employee_management_route_is_available(route_client):
    response = route_client.get("/zamestnanci")

    assert response.status_code == 200
    assert "Správa zaměstnanců" in response.get_data(as_text=True)


def test_excel_viewer_route_is_available(route_client):
    response = route_client.get("/excel_viewer")

    assert response.status_code == 200
    assert "Prohlížeč Excel tabulek" in response.get_data(as_text=True)


def test_excel_editor_route_is_available(route_client):
    response = route_client.get("/excel_editor")

    assert response.status_code == 200
    assert "Editor Excel souborů" in response.get_data(as_text=True)


def test_advances_route_is_available(route_client):
    response = route_client.get("/zalohy")

    assert response.status_code == 200
    body = response.get_data(as_text=True)
    assert "Správa záloh" in body
    assert body.index('<option value="Jan Test">Jan Test</option>') < body.index(
        '<option value="Alpha Worker">Alpha Worker</option>'
    )


def test_monthly_report_route_is_available(route_client):
    response = route_client.get("/monthly_report")

    assert response.status_code == 200
    assert "Měsíční Report" in response.get_data(as_text=True)
