"""Sdílené služby pro REST API vrstvu."""

from performance_optimizations import get_excel_file_info
from services.main_service import save_time_entry
from services.settings_service import load_app_settings, save_app_settings


def serialize_employees(employee_manager):
    """Převede zaměstnance do stabilního API formátu."""
    selected_employees = set(employee_manager.get_vybrani_zamestnanci())
    return [
        {"name": employee["name"], "selected": employee["name"] in selected_employees}
        for employee in employee_manager.get_all_employees()
    ]


def update_selected_employees(employee_manager, employees):
    """Aktualizuje výběr zaměstnanců a vrátí novou sadu."""
    employee_manager.set_vybrani_zamestnanci(employees)
    return employee_manager.get_vybrani_zamestnanci()


def create_time_entry(data, employee_manager, excel_manager, hodiny2025_manager):
    """Vytvoří záznam pracovní doby přes jednotnou doménovou službu."""
    selected_employees = employee_manager.get_vybrani_zamestnanci()
    if not selected_employees:
        raise ValueError("No employees selected")

    date = data["date"]
    start_time = data.get("start_time")
    end_time = data.get("end_time")
    lunch_duration = data.get("lunch_duration", "1.0")
    is_free_day = data.get("is_free_day", False)
    notes = data.get("notes", "")

    message = save_time_entry(
        excel_manager,
        hodiny2025_manager,
        date,
        start_time,
        end_time,
        lunch_duration,
        selected_employees,
        is_free_day,
    )

    return {
        "date": date,
        "start_time": start_time if not is_free_day else None,
        "end_time": end_time if not is_free_day else None,
        "lunch_duration": lunch_duration if not is_free_day else None,
        "is_free_day": is_free_day,
        "employees_count": len(selected_employees),
        "notes": notes,
        "message": message,
    }


def get_time_entries(excel_manager, week_number=None):
    """Vrátí data aktuálního nebo konkrétního týdne."""
    return excel_manager.get_current_week_data(week_number)


def filter_time_entries_by_employee(week_data, employee_name):
    """Zúží týdenní tabulku na konkrétního zaměstnance, pokud existuje."""
    if not week_data or not employee_name:
        return week_data

    rows = week_data.get("data", [])
    if not rows:
        return week_data

    header = rows[0]
    filtered_rows = [row for row in rows[1:] if row and row[0] == employee_name]
    filtered_data = [header, *filtered_rows] if filtered_rows else [header]

    return {
        **week_data,
        "data": filtered_data,
        "rows": len(filtered_data),
        "cols": len(filtered_data[0]) if filtered_data else 0,
    }


def get_excel_status():
    """Vrátí stav aktivního Excel souboru přes sdílený helper."""
    return get_excel_file_info()


def get_settings():
    """Načte normalizované runtime nastavení aplikace."""
    return load_app_settings()


def update_settings(current_settings, updates):
    """Sloučí a uloží runtime nastavení aplikace."""
    merged_settings = current_settings.copy()
    merged_settings.update(updates)
    if not save_app_settings(merged_settings):
        raise IOError("Failed to persist settings")
    return load_app_settings()
