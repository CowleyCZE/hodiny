from pathlib import Path

from services.settings_service import (
    load_app_settings,
    load_dynamic_config,
    save_app_settings,
    save_dynamic_config,
)


def test_load_app_settings_returns_defaults_for_missing_file(tmp_path):
    settings_path = tmp_path / "settings.json"

    loaded_settings = load_app_settings(settings_path)

    assert loaded_settings["start_time"] == "07:00"
    assert loaded_settings["project_info"]["name"] == ""
    assert loaded_settings["last_archived_week"] == 0


def test_load_app_settings_merges_with_defaults(tmp_path):
    settings_path = tmp_path / "settings.json"
    settings_path.write_text('{"start_time":"08:30","project_info":{"name":"Test projekt"}}', encoding="utf-8")

    loaded_settings = load_app_settings(settings_path)

    assert loaded_settings["start_time"] == "08:30"
    assert loaded_settings["end_time"] == "18:00"
    assert loaded_settings["project_info"]["name"] == "Test projekt"
    assert loaded_settings["project_info"]["start_date"] == ""


def test_save_app_settings_normalizes_structure(tmp_path):
    settings_path = tmp_path / "settings.json"

    saved = save_app_settings({"start_time": "09:00", "project_info": {"name": "Projekt X"}}, settings_path)

    assert saved is True
    reloaded_settings = load_app_settings(settings_path)
    assert reloaded_settings["start_time"] == "09:00"
    assert reloaded_settings["end_time"] == "18:00"
    assert reloaded_settings["project_info"]["name"] == "Projekt X"


def test_dynamic_config_roundtrip(tmp_path):
    config_path = tmp_path / "config.json"
    payload = {"weekly_time": {"date": [{"file": "Hodiny_Cap.xlsx", "sheet": "Týden", "cell": "B6"}]}}

    saved = save_dynamic_config(payload, config_path)

    assert saved is True
    assert load_dynamic_config(config_path) == payload


def test_invalid_dynamic_config_payload_is_rejected(tmp_path):
    config_path = Path(tmp_path) / "config.json"

    saved = save_dynamic_config(["invalid"], config_path)

    assert saved is False
    assert load_dynamic_config(config_path) == {}
