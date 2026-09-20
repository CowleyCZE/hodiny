"""Synchronizační stav mezi kompatibilním Excel výkazem a SQLite."""

from datetime import datetime


def sync_excel_state(database, excel_manager, employees=None):
    """Aktualizuje stav synchronizace po ověření dostupnosti Excel souboru."""
    status = excel_manager.get_excel_status() if hasattr(excel_manager, "get_excel_status") else {}
    database.set_sync_state(
        "excel_filename", str(status.get("filename", getattr(excel_manager, "active_filename", "")))
    )
    database.set_sync_state("last_sync", datetime.now().astimezone().isoformat())
    database.set_sync_state("employees", ", ".join(employees or []))
    return {"status": "ok", "last_sync": database.get_sync_state().get("last_sync")}


def sync_status(database):
    state = database.get_sync_state()
    return {"status": "ok" if state.get("last_sync") else "never", "state": state}
