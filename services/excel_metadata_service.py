"""Perzistence metadat Excel souborů."""

import json

from utils.logger import setup_logger

logger = setup_logger("excel_metadata_service")


def load_metadata(metadata_path):
    """Načte metadata souborů z JSON souboru."""
    if not metadata_path.exists():
        return {}

    try:
        with open(metadata_path, "r", encoding="utf-8") as metadata_file:
            loaded_metadata = json.load(metadata_file)
            return loaded_metadata if isinstance(loaded_metadata, dict) else {}
    except (IOError, json.JSONDecodeError) as exc:
        logger.error("Chyba při načítání metadat: %s", exc, exc_info=True)
        return {}


def save_metadata(metadata_path, metadata):
    """Uloží metadata souborů do JSON souboru."""
    try:
        with open(metadata_path, "w", encoding="utf-8") as metadata_file:
            json.dump(metadata, metadata_file, indent=4, ensure_ascii=False)
        return True
    except IOError as exc:
        logger.error("Chyba při ukládání metadat: %s", exc, exc_info=True)
        return False


def set_file_category(metadata_path, filename, category):
    """Nastaví kategorii pro daný soubor v metadatech."""
    metadata = load_metadata(metadata_path)
    if filename not in metadata:
        metadata[filename] = {}
    metadata[filename]["category"] = category
    return save_metadata(metadata_path, metadata)
