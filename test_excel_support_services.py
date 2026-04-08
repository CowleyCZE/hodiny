import json

from services.excel_config_service import get_configured_cells, load_dynamic_excel_config
from services.excel_metadata_service import load_metadata, save_metadata, set_file_category


def test_load_dynamic_excel_config_returns_dict(tmp_path):
    config_path = tmp_path / "config.json"
    config_path.write_text(
        json.dumps({"weekly_time": {"date": [{"file": "Hodiny_Cap.xlsx", "sheet": "Týden", "cell": "B6"}]}}),
        encoding="utf-8",
    )

    loaded_config = load_dynamic_excel_config(config_path)

    assert "weekly_time" in loaded_config


def test_get_configured_cells_filters_by_file_and_sheet(tmp_path):
    config_path = tmp_path / "config.json"
    config_path.write_text(
        json.dumps(
            {
                "advances": {
                    "employee_name": [
                        {"file": "Hodiny_Cap.xlsx", "sheet": "Zálohy", "cell": "A9"},
                        {"file": "Jiny.xlsx", "sheet": "Zálohy", "cell": "A10"},
                    ]
                }
            }
        ),
        encoding="utf-8",
    )

    coordinates = get_configured_cells(
        "advances",
        "employee_name",
        "Hodiny_Cap.xlsx",
        sheet_name="Zálohy",
        config_path=config_path,
    )

    assert coordinates == [(9, 1)]


def test_metadata_roundtrip_and_category_update(tmp_path):
    metadata_path = tmp_path / "metadata.json"

    assert save_metadata(metadata_path, {"report.xlsx": {"category": "Ostatní"}}) is True
    assert load_metadata(metadata_path)["report.xlsx"]["category"] == "Ostatní"

    assert set_file_category(metadata_path, "report.xlsx", "Rozpočet") is True
    assert load_metadata(metadata_path)["report.xlsx"]["category"] == "Rozpočet"
