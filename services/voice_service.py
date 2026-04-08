"""Služby pro zpracování hlasových a textových příkazů."""

from performance_optimizations import get_employee_stats
from utils.voice_processor import VoiceProcessor


def process_voice_command(command_text, employee_manager, excel_manager, hodiny2025_manager, save_time_entry):
    """Zpracuje textový příkaz a vrátí payload připravený pro frontend."""
    result = VoiceProcessor().process_command(text=command_text)
    if not result.get("success"):
        return result, 400

    entities = {
        "action": result.get("action"),
        "date": result.get("date"),
        "start_time": result.get("start_time"),
        "end_time": result.get("end_time"),
        "lunch_duration": result.get("lunch_duration"),
        "is_free_day": result.get("is_free_day", False),
        "employee": result.get("employee"),
        "time_period": result.get("time_period"),
    }

    payload = {
        "success": True,
        "confidence": 1.0,
        "entities": entities,
        "original_text": result.get("original_text", command_text),
    }

    if entities["action"] in {"record_time", "record_free_day"}:
        selected_employees = employee_manager.get_vybrani_zamestnanci()
        if not selected_employees:
            return {"success": False, "error": "Nejsou vybráni žádní zaměstnanci"}, 400

        operation_message = save_time_entry(
            excel_manager,
            hodiny2025_manager,
            entities["date"],
            entities["start_time"],
            entities["end_time"],
            str(entities["lunch_duration"]),
            selected_employees,
            entities["is_free_day"],
        )
        payload["operation_result"] = {"message": operation_message}
        return payload, 200

    if entities["action"] == "get_stats":
        payload["stats"] = get_employee_stats()
        return payload, 200

    return {"success": False, "error": "Nepodporovaná akce"}, 400
