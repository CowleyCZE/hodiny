"""Helpery pro agregaci reportů z Excel týdenních listů."""

from datetime import datetime

from config import Config


def generate_monthly_report_from_workbook(workbook, month, year, employees=None):
    """Agreguje hodiny a volné dny z workbooku za daný měsíc."""
    report_data = {}
    for sheet in get_monthly_sheets(workbook, month, year):
        process_sheet_for_report(sheet, employees, report_data, month, year)

    return {
        employee_name: data
        for employee_name, data in report_data.items()
        if data["total_hours"] > 0 or data["free_days"] > 0
    }


def get_monthly_sheets(workbook, month, year):
    """Generátor listů, které spadají do daného měsíce a roku."""
    for sheet_name in workbook.sheetnames:
        if not sheet_name.startswith(Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME):
            continue

        sheet = workbook[sheet_name]
        for column_index in range(2, 15, 2):
            date_value = sheet.cell(row=80, column=column_index).value
            if isinstance(date_value, datetime) and date_value.month == month and date_value.year == year:
                yield sheet
                break


def process_sheet_for_report(sheet, employees, report_data, month, year):
    """Zpracuje jeden list a přičte data do agregace."""
    for row_index in range(Config.EXCEL_EMPLOYEE_START_ROW, sheet.max_row + 1):
        employee_name = sheet.cell(row=row_index, column=1).value
        if not employee_name or (employees and employee_name not in employees):
            continue

        if employee_name not in report_data:
            report_data[employee_name] = {"total_hours": 0.0, "free_days": 0}

        for column_index in range(2, 15, 2):
            date_value = sheet.cell(row=80, column=column_index).value
            if not (isinstance(date_value, datetime) and date_value.month == month and date_value.year == year):
                continue

            hours = sheet.cell(row=row_index, column=column_index).value
            if not isinstance(hours, (int, float)):
                continue

            if hours > 0:
                report_data[employee_name]["total_hours"] += hours
            else:
                report_data[employee_name]["free_days"] += 1
