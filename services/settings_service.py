"""Služby pro práci s runtime nastavením aplikace a dynamickou konfigurací."""

import json

from config import Config
from utils.logger import setup_logger

logger = setup_logger("settings_service")


def _merge_app_settings(raw_settings):
    """Sloučí načtená data s výchozí strukturou nastavení."""
    merged_settings = Config.get_default_settings()

    if not isinstance(raw_settings, dict):
        return merged_settings

    for key in ("start_time", "end_time", "lunch_duration", "last_archived_week", "preferred_employee_name"):
        if key in raw_settings:
            merged_settings[key] = raw_settings[key]

    raw_project_info = raw_settings.get("project_info", {})
    if isinstance(raw_project_info, dict):
        for key in ("name", "start_date", "end_date", "firma", "mesto"):
            if key in raw_project_info:
                merged_settings["project_info"][key] = raw_project_info[key]

    return merged_settings


def load_app_settings(settings_path=None):
    """Načte aplikační nastavení a doplní chybějící klíče defaulty."""
    target_path = settings_path or Config.SETTINGS_FILE_PATH
    if not target_path.exists():
        return Config.get_default_settings()

    try:
        with open(target_path, "r", encoding="utf-8") as settings_file:
            raw_settings = json.load(settings_file)
    except (json.JSONDecodeError, IOError) as exc:
        logger.error("Chyba při načítání nastavení: %s", exc, exc_info=True)
        return Config.get_default_settings()

    return _merge_app_settings(raw_settings)


def save_app_settings(settings_data, settings_path=None):
    """Uloží aplikační nastavení v normalizované podobě."""
    target_path = settings_path or Config.SETTINGS_FILE_PATH
    normalized_settings = _merge_app_settings(settings_data)

    try:
        target_path.parent.mkdir(parents=True, exist_ok=True)
        with open(target_path, "w", encoding="utf-8") as settings_file:
            json.dump(normalized_settings, settings_file, indent=4, ensure_ascii=False)
        return True
    except (IOError, TypeError) as exc:
        logger.error("Chyba při ukládání nastavení: %s", exc, exc_info=True)
        return False


def load_dynamic_config(config_path=None):
    """Načte dynamickou Excel konfiguraci z JSON."""
    target_path = config_path or Config.CONFIG_FILE_PATH
    if not target_path.exists():
        return {}

    try:
        with open(target_path, "r", encoding="utf-8") as config_file:
            loaded_config = json.load(config_file)
            return loaded_config if isinstance(loaded_config, dict) else {}
    except (json.JSONDecodeError, IOError) as exc:
        logger.error("Chyba při načítání dynamické konfigurace: %s", exc, exc_info=True)
        return {}


def save_dynamic_config(config_data, config_path=None):
    """Uloží dynamickou Excel konfiguraci do JSON."""
    target_path = config_path or Config.CONFIG_FILE_PATH

    if not isinstance(config_data, dict):
        logger.error("Dynamická konfigurace musí být slovník.")
        return False

    try:
        target_path.parent.mkdir(parents=True, exist_ok=True)
        with open(target_path, "w", encoding="utf-8") as config_file:
            json.dump(config_data, config_file, indent=4, ensure_ascii=False)
        return True
    except (IOError, TypeError) as exc:
        logger.error("Chyba při ukládání dynamické konfigurace: %s", exc, exc_info=True)
        return False


def get_mapping_schema():
    """Vrací schéma kategorií, polí, figurek (tvarů) a barev pro vizuální mapovač."""
    return {
        "weekly_time": {
            "label": "Týdenní evidence",
            "icon": "🕒",
            "shape": "circle",
            "description": "Konfigurace pro ukládání týdenních časových záznamů",
            "fields": {
                "employee_name": {
                    "label": "Jméno zaměstnance",
                    "color": "#6366f1",
                    "badge": "EMP",
                    "description": "Místo, kam se ukládá jméno zaměstnance",
                },
                "date": {
                    "label": "Datum",
                    "color": "#06b6d4",
                    "badge": "DATE",
                    "description": "Místo, kam se ukládá datum záznamu",
                },
                "start_time": {
                    "label": "Čas začátku (Od)",
                    "color": "#10b981",
                    "badge": "OD",
                    "description": "Místo, kam se ukládá čas začátku práce",
                },
                "end_time": {
                    "label": "Čas konce (Do)",
                    "color": "#f59e0b",
                    "badge": "DO",
                    "description": "Místo, kam se ukládá čas konce práce",
                },
                "lunch_duration": {
                    "label": "Doba oběda",
                    "color": "#8b5cf6",
                    "badge": "OBĚD",
                    "description": "Místo, kam se ukládá doba oběda",
                },
                "total_hours": {
                    "label": "Celkové hodiny",
                    "color": "#ec4899",
                    "badge": "HOD",
                    "description": "Místo, kam se ukládají celkové odpracované hodiny",
                },
            },
        },
        "advances": {
            "label": "Zálohy a půjčky",
            "icon": "💰",
            "shape": "square",
            "description": "Konfigurace pro ukládání záloh a půjček zaměstnanců",
            "fields": {
                "employee_name": {
                    "label": "Jméno zaměstnance (Zálohy)",
                    "color": "#3b82f6",
                    "badge": "Z-EMP",
                    "description": "Místo pro jméno zaměstnance na listu Zálohy",
                },
                "amount_eur": {
                    "label": "Částka EUR",
                    "color": "#059669",
                    "badge": "EUR",
                    "description": "Místo pro částky záloh v EUR",
                },
                "amount_czk": {
                    "label": "Částka CZK",
                    "color": "#0d9488",
                    "badge": "CZK",
                    "description": "Místo pro částky záloh v CZK",
                },
                "date": {
                    "label": "Datum zálohy",
                    "color": "#d97706",
                    "badge": "Z-DAT",
                    "description": "Místo pro datum zápisu zálohy",
                },
                "option_type": {
                    "label": "Kategorie zálohy (Hlavička sloupce)",
                    "color": "#7c3aed",
                    "badge": "TYP",
                    "description": "Názvy kategorií/možností záloh",
                },
            },
        },
        "monthly_time": {
            "label": "Měsíční evidence (Hodiny2025)",
            "icon": "📅",
            "shape": "diamond",
            "description": "Konfigurace pro ukládání měsíčních časových záznamů",
            "fields": {
                "employee_name": {
                    "label": "Jméno zaměstnance",
                    "color": "#4f46e5",
                    "badge": "M-EMP",
                    "description": "Jméno zaměstnance v měsíční evidenci",
                },
                "date": {
                    "label": "Datum",
                    "color": "#0891b2",
                    "badge": "M-DAT",
                    "description": "Datum v měsíční evidenci",
                },
                "start_time": {
                    "label": "Čas začátku",
                    "color": "#16a34a",
                    "badge": "M-OD",
                    "description": "Čas začátku v měsíční evidenci",
                },
                "end_time": {
                    "label": "Čas konce",
                    "color": "#ea580c",
                    "badge": "M-DO",
                    "description": "Čas konce v měsíční evidenci",
                },
                "lunch_hours": {
                    "label": "Hodiny oběda",
                    "color": "#9333ea",
                    "badge": "M-OB",
                    "description": "Doba oběda v hodinách",
                },
                "total_hours": {
                    "label": "Celkové hodiny",
                    "color": "#db2777",
                    "badge": "M-HOD",
                    "description": "Celkové odpracované hodiny",
                },
                "overtime": {
                    "label": "Přesčasy",
                    "color": "#dc2626",
                    "badge": "PŘES",
                    "description": "Přesčasové hodiny",
                },
                "num_employees": {
                    "label": "Počet zaměstnanců",
                    "color": "#475569",
                    "badge": "POČET",
                    "description": "Počet zaměstnanců pro daný den",
                },
                "total_all_employees": {
                    "label": "Celkem všichni zaměstnanci",
                    "color": "#ca8a04",
                    "badge": "SUM-VŠI",
                    "description": "Součet hodin všech zaměstnanců",
                },
            },
        },
        "projects": {
            "label": "Projekty",
            "icon": "📁",
            "shape": "hexagon",
            "description": "Konfigurace pro ukládání informací o projektech",
            "fields": {
                "project_name": {
                    "label": "Název projektu",
                    "color": "#eab308",
                    "badge": "PROJ",
                    "description": "Místo, kam se ukládá název projektu",
                },
                "start_date": {
                    "label": "Datum začátku projektu",
                    "color": "#84cc16",
                    "badge": "P-OD",
                    "description": "Místo pro počáteční datum projektu",
                },
                "end_date": {
                    "label": "Datum konce projektu",
                    "color": "#64748b",
                    "badge": "P-DO",
                    "description": "Místo pro konečné datum projektu",
                },
            },
        },
    }


def get_default_dynamic_config():
    """Vrátí výchozí dynamické mapování odpovídající standardním šablonám."""
    return {
        "weekly_time": {
            "employee_name": [{"file": "Hodiny_Cap.xlsx", "sheet": "Týden", "cell": "A8"}],
            "date": [{"file": "Hodiny_Cap.xlsx", "sheet": "Týden", "cell": "B80"}],
            "start_time": [{"file": "Hodiny_Cap.xlsx", "sheet": "Týden", "cell": "B7"}],
            "end_time": [{"file": "Hodiny_Cap.xlsx", "sheet": "Týden", "cell": "C7"}],
            "lunch_duration": [],
            "total_hours": [{"file": "Hodiny_Cap.xlsx", "sheet": "Týden", "cell": "B8"}],
        },
        "advances": {
            "employee_name": [{"file": "Hodiny_Cap.xlsx", "sheet": "Zálohy", "cell": "A8"}],
            "amount_eur": [
                {"file": "Hodiny_Cap.xlsx", "sheet": "Zálohy", "cell": "B8"},
                {"file": "Hodiny_Cap.xlsx", "sheet": "Zálohy", "cell": "D8"},
                {"file": "Hodiny_Cap.xlsx", "sheet": "Zálohy", "cell": "F8"},
                {"file": "Hodiny_Cap.xlsx", "sheet": "Zálohy", "cell": "H8"},
            ],
            "amount_czk": [
                {"file": "Hodiny_Cap.xlsx", "sheet": "Zálohy", "cell": "C8"},
                {"file": "Hodiny_Cap.xlsx", "sheet": "Zálohy", "cell": "E8"},
                {"file": "Hodiny_Cap.xlsx", "sheet": "Zálohy", "cell": "G8"},
                {"file": "Hodiny_Cap.xlsx", "sheet": "Zálohy", "cell": "I8"},
            ],
            "date": [{"file": "Hodiny_Cap.xlsx", "sheet": "Zálohy", "cell": "Z8"}],
            "option_type": [
                {"file": "Hodiny_Cap.xlsx", "sheet": "Zálohy", "cell": "B80"},
                {"file": "Hodiny_Cap.xlsx", "sheet": "Zálohy", "cell": "D80"},
                {"file": "Hodiny_Cap.xlsx", "sheet": "Zálohy", "cell": "F80"},
                {"file": "Hodiny_Cap.xlsx", "sheet": "Zálohy", "cell": "H80"},
            ],
        },
        "monthly_time": {
            "employee_name": [],
            "date": [],
            "start_time": [],
            "end_time": [],
            "lunch_hours": [],
            "total_hours": [],
            "overtime": [],
            "num_employees": [],
            "total_all_employees": [],
        },
        "projects": {
            "project_name": [{"file": "Hodiny_Cap.xlsx", "sheet": "Týden", "cell": "B4"}],
            "start_date": [],
            "end_date": [],
        },
    }
