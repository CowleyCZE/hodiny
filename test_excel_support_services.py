import json

from services.excel_config_service import get_configured_cells, load_dynamic_excel_config
from services.excel_metadata_service import load_metadata, save_metadata, set_file_category


def test_load_dynamic_excel_config_returns_dict(tmp_path):
    config_path = tmp_path / "config.json"
    config_path.write_text(
        json.dumps({"weekly_time": {"date": [{"file": "Hodiny_Cap.xlsx", "sheet": "Týden", "cell": "B80"}]}}),
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
                        {"file": "Hodiny_Cap.xlsx", "sheet": "Zálohy", "cell": "A8"},
                        {"file": "Jiny.xlsx", "sheet": "Zálohy", "cell": "A9"},
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

    assert coordinates == [(8, 1)]


def test_get_configured_cells_accepts_week_template_for_dynamic_week_sheet(tmp_path):
    config_path = tmp_path / "config.json"
    config_path.write_text(
        json.dumps(
            {
                "weekly_time": {
                    "start_time": [
                        {"file": "Hodiny_Cap.xlsx", "sheet": "Týden", "cell": "B7"},
                    ]
                }
            }
        ),
        encoding="utf-8",
    )

    coordinates = get_configured_cells(
        "weekly_time",
        "start_time",
        "Hodiny_Cap.xlsx",
        sheet_name="Týden 42",
        config_path=config_path,
    )

    assert coordinates == [(7, 2)]


def test_metadata_roundtrip_and_category_update(tmp_path):
    metadata_path = tmp_path / "metadata.json"

    assert save_metadata(metadata_path, {"report.xlsx": {"category": "Ostatní"}}) is True
    assert load_metadata(metadata_path)["report.xlsx"]["category"] == "Ostatní"

    assert set_file_category(metadata_path, "report.xlsx", "Rozpočet") is True
    assert load_metadata(metadata_path)["report.xlsx"]["category"] == "Rozpočet"


def test_copy_worksheet_deep_preserves_formatting_and_dimensions():
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from services.excel_week_service import copy_worksheet_deep, create_week_sheet_from_template

    src_wb = Workbook()
    src_ws = src_wb.active
    src_ws.title = "Týden"

    # Buňka se stylem
    src_ws["A1"] = "Hlavička"
    src_ws["A1"].font = Font(bold=True, size=14, color="FF0000")
    src_ws["A1"].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    src_ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    src_ws["A1"].border = Border(top=Side(style="thin", color="000000"))

    # Sloučení
    src_ws.merge_cells("A1:C2")

    # Rozměry
    src_ws.column_dimensions["A"].width = 30.5
    src_ws.column_dimensions["B"].width = 12.0
    src_ws.row_dimensions[1].height = 45.0

    # Freeze panes a gridlines
    src_ws.freeze_panes = "B3"
    src_ws.sheet_view.showGridLines = True

    # 1. Test cross-workbook deep copy
    dst_wb = Workbook()
    dst_ws = dst_wb.create_sheet("Týden 10")
    copy_worksheet_deep(src_ws, dst_ws)

    assert dst_ws["A1"].value == "Hlavička"
    assert dst_ws["A1"].font.bold is True
    assert dst_ws["A1"].font.color.rgb == "00FF0000" or dst_ws["A1"].font.color.rgb == "FF0000"
    assert dst_ws["A1"].alignment.horizontal == "center"
    assert "A1:C2" in [str(r) for r in dst_ws.merged_cells.ranges]
    assert dst_ws.column_dimensions["A"].width == 30.5
    assert dst_ws.column_dimensions["B"].width == 12.0
    assert dst_ws.row_dimensions[1].height == 45.0
    assert dst_ws.freeze_panes == "B3"
    assert dst_ws.sheet_view.showGridLines is True

    # 2. Test create_week_sheet_from_template within same workbook
    create_week_sheet_from_template(src_wb, "Týden 11")
    assert "Týden 11" in src_wb.sheetnames
    cloned_ws = src_wb["Týden 11"]
    assert cloned_ws["A1"].value == "Hlavička"
    assert "A1:C2" in [str(r) for r in cloned_ws.merged_cells.ranges]
    assert cloned_ws.column_dimensions["A"].width == 30.5
    assert cloned_ws.freeze_panes == "B3"

