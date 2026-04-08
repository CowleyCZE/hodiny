"""Sdílené helpery pro čtení dynamické Excel konfigurace."""

import json

from openpyxl.utils import coordinate_to_tuple

from config import Config
from utils.logger import setup_logger

logger = setup_logger("excel_config_service")


def load_dynamic_excel_config(config_path=None):
    """Načte dynamickou Excel konfiguraci z JSON souboru."""
    target_path = config_path or Config.CONFIG_FILE_PATH
    if not target_path.exists():
        return {}

    try:
        with open(target_path, "r", encoding="utf-8") as config_file:
            loaded_config = json.load(config_file)
            return loaded_config if isinstance(loaded_config, dict) else {}
    except (json.JSONDecodeError, IOError) as exc:
        logger.error("Chyba při načítání dynamické Excel konfigurace: %s", exc, exc_info=True)
        return {}


def get_configured_cells(config_section, field_key, active_filename, sheet_name=None, config_path=None):
    """Vrátí seznam souřadnic z dynamické konfigurace pro konkrétní field."""
    config = load_dynamic_excel_config(config_path)
    field_configs = config.get(config_section, {}).get(field_key, [])

    coordinates = []
    for field_config in field_configs:
        if field_config.get("file") != active_filename:
            logger.warning(
                "Konfigurace pro %s/%s odkazuje na jiný soubor: %s",
                config_section,
                field_key,
                field_config.get("file"),
            )
            continue

        if sheet_name and field_config.get("sheet") != sheet_name:
            logger.warning(
                "Konfigurace pro %s/%s odkazuje na jiný list: %s (očekáván %s)",
                config_section,
                field_key,
                field_config.get("sheet"),
                sheet_name,
            )
            continue

        cell = field_config.get("cell")
        if not cell:
            continue

        try:
            coordinates.append(coordinate_to_tuple(cell))
        except ValueError as exc:
            logger.error("Neplatná buňka v konfiguraci pro %s/%s: %s - %s", config_section, field_key, cell, exc)

    return coordinates
