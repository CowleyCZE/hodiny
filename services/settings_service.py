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
        merged_settings["project_info"].update(
            {
                key: value
                for key, value in raw_project_info.items()
                if key in merged_settings["project_info"]
            }
        )

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
