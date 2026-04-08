"""Routy pro advanced technickou konfiguraci projektu."""

from flask import Blueprint, jsonify, render_template, request

from services.excel_file_service import (
    get_sheet_content,
    get_sheet_names,
    list_excel_files,
    rename_excel_file,
)
from services.settings_service import load_dynamic_config, save_dynamic_config
from utils.logger import setup_logger

logger = setup_logger("configuration_routes")

configuration_bp = Blueprint("configuration", __name__)


@configuration_bp.route("/api/settings", methods=["GET"])
def api_get_settings():
    """Vrací aktuální obsah technické konfigurace."""
    try:
        return jsonify(load_dynamic_config())
    except Exception as exc:
        logger.error("Chyba při načítání API nastavení: %s", exc, exc_info=True)
        return jsonify({"error": "Chyba při načítání nastavení"}), 500


@configuration_bp.route("/api/settings", methods=["POST"])
def api_save_settings():
    """Přijímá technickou konfiguraci ve formátu JSON a ukládá ji."""
    try:
        config_data = request.get_json()
        if not config_data:
            return jsonify({"error": "Žádná data nebyla odeslána"}), 400

        if not save_dynamic_config(config_data):
            return jsonify({"error": "Nepodařilo se uložit nastavení"}), 500

        return jsonify({"success": True, "message": "Nastavení bylo úspěšně uloženo"})
    except Exception as exc:
        logger.error("Chyba při ukládání API nastavení: %s", exc, exc_info=True)
        return jsonify({"error": "Chyba při ukládání nastavení"}), 500


@configuration_bp.route("/api/files", methods=["GET"])
def api_get_files():
    """Vrátí seznam Excel souborů pro technickou konfiguraci."""
    try:
        return jsonify({"files": list_excel_files()})
    except Exception as exc:
        logger.error("Chyba při načítání Excel souborů: %s", exc, exc_info=True)
        return jsonify({"error": "Chyba při načítání souborů"}), 500


@configuration_bp.route("/api/sheets/<filename>", methods=["GET"])
def api_get_sheets(filename):
    """Vrátí názvy listů pro zadaný soubor."""
    try:
        return jsonify({"sheets": get_sheet_names(filename)})
    except FileNotFoundError:
        return jsonify({"error": "Soubor nenalezen"}), 404
    except Exception as exc:
        logger.error("Chyba při načítání listů ze souboru %s: %s", filename, exc, exc_info=True)
        return jsonify({"error": "Chyba při načítání listů"}), 500


@configuration_bp.route("/api/sheet_content/<filename>/<sheetname>", methods=["GET"])
def api_get_sheet_content(filename, sheetname):
    """Vrátí obsah zadaného listu ve formátu JSON."""
    try:
        return jsonify(get_sheet_content(filename, sheetname))
    except FileNotFoundError:
        return jsonify({"error": "Soubor nenalezen"}), 404
    except ValueError:
        return jsonify({"error": "List nenalezen"}), 404
    except Exception as exc:
        logger.error("Chyba při načítání obsahu listu %s/%s: %s", filename, sheetname, exc, exc_info=True)
        return jsonify({"error": "Chyba při načítání obsahu listu"}), 500


@configuration_bp.route("/nastaveni")
def advanced_settings_page():
    """Stránka pro dynamické technické nastavení ukládání do XLSX."""
    return render_template("nastaveni.html")


@configuration_bp.route("/api/files/rename", methods=["POST"])
def rename_file():
    """API endpoint pro přejmenování XLSX souborů."""
    try:
        data = request.get_json() or {}
        renamed_filename = rename_excel_file(data.get("old_filename"), data.get("new_filename"))
        return jsonify({"success": True, "message": f"Soubor byl úspěšně přejmenován na {renamed_filename}"})
    except ValueError as exc:
        return jsonify({"success": False, "error": str(exc)}), 400
    except FileNotFoundError as exc:
        return jsonify({"success": False, "error": str(exc)}), 404
    except FileExistsError as exc:
        return jsonify({"success": False, "error": str(exc)}), 409
    except Exception as exc:
        logger.error("Chyba při přejmenování souboru: %s", exc, exc_info=True)
        return jsonify({"success": False, "error": str(exc)}), 500
