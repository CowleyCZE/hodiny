"""Routy pro správu turnusů (pracovních výjezdů a akcí)."""

import datetime as dt
from pathlib import Path

from flask import Blueprint, flash, redirect, render_template, request, send_from_directory, url_for

from config import Config
from services.turnus_service import (
    close_turnus,
    create_turnus,
    get_active_turnus,
    get_all_turnusy,
)
from utils.logger import setup_logger

logger = setup_logger("turnus_routes")

turnus_bp = Blueprint("turnus", __name__)


@turnus_bp.route("/turnusy", methods=["GET"])
def turnusy_view():
    """Zobrazí přehled aktivního turnusu a historii všech předchozích akcí."""
    active_turnus = get_active_turnus()
    all_turnusy = get_all_turnusy()
    today_str = dt.date.today().isoformat()

    return render_template(
        "turnusy.html",
        active_turnus=active_turnus,
        all_turnusy=all_turnusy,
        today_str=today_str,
    )


@turnus_bp.route("/turnusy/novy", methods=["POST"])
def novy_turnus():
    """Vytvoří nový aktivní turnus."""
    try:
        name = request.form.get("name", "").strip()
        start_date = request.form.get("start_date", "").strip()
        end_date = request.form.get("end_date", "").strip() or None
        firma = request.form.get("firma", "").strip()
        mesto = request.form.get("mesto", "").strip()

        created = create_turnus(name, start_date, end_date=end_date, firma=firma, mesto=mesto)
        flash(f"Turnus '{created['name']}' byl úspěšně založen jako aktivní.", "success")
    except ValueError as exc:
        flash(str(exc), "error")
    except Exception as exc:
        logger.error("Chyba při zakládání turnusu: %s", exc, exc_info=True)
        flash(f"Neočekávaná chyba při zakládání turnusu: {exc}", "error")

    return redirect(url_for("turnus.turnusy_view"))


@turnus_bp.route("/turnusy/ukoncit", methods=["POST"])
def ukoncit_turnus():
    """Ukončí běžící turnus k zadanému datu a vytvoří archivní kopii."""
    try:
        turnus_id = request.form.get("turnus_id")
        end_date = request.form.get("end_date", "").strip() or dt.date.today().isoformat()

        closed = close_turnus(turnus_id=turnus_id, end_date=end_date)
        archive_info = f" Archiv byl uložen: {closed['archive_file']}." if closed.get("archive_file") else ""
        flash(f"Turnus '{closed['name']}' byl úspěšně ukončen k {closed['end_date']}.{archive_info}", "success")
    except ValueError as exc:
        flash(str(exc), "error")
    except Exception as exc:
        logger.error("Chyba při ukončování turnusu: %s", exc, exc_info=True)
        flash(f"Neočekávaná chyba při ukončování turnusu: {exc}", "error")

    return redirect(url_for("turnus.turnusy_view"))


@turnus_bp.route("/turnusy/stahnout-archiv/<path:filename>")
def stahnout_archiv(filename):
    """Umožní stáhnout archivní kopii ukončeného turnusu."""
    try:
        target_path = Path(filename)
        # Pokud je v cestě "backups/Turnus_...", extrahujeme pouze jméno souboru
        filename_only = target_path.name
        return send_from_directory(Config.BACKUP_PATH, filename_only, as_attachment=True)
    except FileNotFoundError:
        flash("Archivní soubor turnusu nebyl nalezen.", "error")
        return redirect(url_for("turnus.turnusy_view"))
