"""Hlavní dashboard, ruční záznam a rychlé zadání pracovní doby."""

import datetime as dt
import smtplib

from flask import Blueprint, flash, g, jsonify, redirect, render_template, request, session, url_for

from performance_optimizations import timing_decorator
from services.main_service import (
    build_dashboard_context,
    get_next_workday,
    save_time_entry,
    send_active_excel_email,
)
from services.voice_service import process_voice_command
from utils.logger import setup_logger

logger = setup_logger("main_routes")

main_bp = Blueprint("main", __name__)


@main_bp.route("/")
@timing_decorator
def index():
    """Úvodní stránka s rychlými informacemi a rychlým zadáním času."""
    context = build_dashboard_context(g.excel_manager, session.get("settings", {}))
    return render_template("index.html", **context)


@main_bp.route("/send_email", methods=["POST"])
def send_email():
    """Odešle aktivní Excel jako přílohu na konfigurovaný e-mail."""
    try:
        send_active_excel_email(g.excel_manager)
        flash("Email byl úspěšně odeslán.", "success")
    except (ValueError, smtplib.SMTPException, FileNotFoundError) as exc:
        logger.error("Chyba při odesílání emailu: %s", exc, exc_info=True)
        flash("Chyba při odesílání emailu.", "error")
    return redirect(url_for("main.index"))


@main_bp.route("/zaznam", methods=["GET", "POST"])
def record_time():
    """Formulář pro zápis pracovní doby nebo volného dne."""
    selected_employees = g.employee_manager.get_vybrani_zamestnanci()
    if not selected_employees:
        flash("Nejsou vybráni žádní zaměstnanci.", "warning")
        return redirect(url_for("employees.manage_employees"))

    current_date = request.args.get("next_date", dt.datetime.now().strftime("%Y-%m-%d"))
    start_time = session["settings"].get("start_time", "07:00")
    end_time = session["settings"].get("end_time", "18:00")
    lunch_duration = str(session["settings"].get("lunch_duration", 1.0))
    is_free_day = False

    if request.method == "POST":
        current_date = request.form.get("date", current_date)
        start_time = request.form.get("start_time", start_time)
        end_time = request.form.get("end_time", end_time)
        lunch_duration = request.form.get("lunch_duration", lunch_duration)
        is_free_day = request.form.get("is_free_day") == "on"

        try:
            entry_date = dt.datetime.strptime(current_date, "%Y-%m-%d").date()
            save_time_entry(
                g.excel_manager,
                g.hodiny2025_manager,
                current_date,
                start_time,
                end_time,
                lunch_duration,
                selected_employees,
                is_free_day,
            )
            flash("Záznam byl úspěšně uložen.", "success")
            return redirect(url_for("main.record_time", next_date=get_next_workday(entry_date).strftime("%Y-%m-%d")))
        except (ValueError, IOError, FileNotFoundError) as exc:
            flash(str(exc), "error")

    return render_template(
        "record_time.html",
        selected_employees=selected_employees,
        current_date=current_date,
        start_time=start_time,
        end_time=end_time,
        lunch_duration=lunch_duration,
        is_free_day=is_free_day,
    )


@main_bp.route("/api/quick_time_entry", methods=["POST"])
def quick_time_entry():
    """API endpoint pro rychlé zadání pracovní doby z hlavní stránky."""
    try:
        data = request.get_json() or {}
        date = data.get("date")
        start_time = data.get("start_time")
        end_time = data.get("end_time")
        lunch_duration = data.get("lunch_duration", "1.0")
        is_free_day = data.get("is_free_day", False)

        if not date:
            return jsonify({"success": False, "error": "Chybí datum"}), 400

        selected_employees = g.employee_manager.get_vybrani_zamestnanci()
        if not selected_employees:
            return jsonify({"success": False, "error": "Nejsou vybráni žádní zaměstnanci"}), 400

        try:
            dt.datetime.strptime(date, "%Y-%m-%d")
            message = save_time_entry(
                g.excel_manager,
                g.hodiny2025_manager,
                date,
                start_time,
                end_time,
                lunch_duration,
                selected_employees,
                is_free_day,
            )
            return jsonify({"success": True, "message": message})
        except ValueError as exc:
            return jsonify({"success": False, "error": f"Neplatné datum nebo čas: {exc}"}), 400

    except Exception as exc:
        logger.error("Chyba při rychlém zadání času: %s", exc, exc_info=True)
        return jsonify({"success": False, "error": str(exc)}), 500


@main_bp.route("/voice-command", methods=["POST"])
def voice_command():
    """Zpracuje textový hlasový příkaz z homepage voice handleru."""
    try:
        data = request.get_json() or {}
        command_text = (data.get("command") or "").strip()
        if not command_text:
            return jsonify({"success": False, "error": "Chybí text příkazu"}), 400

        payload, status_code = process_voice_command(
            command_text,
            g.employee_manager,
            g.excel_manager,
            g.hodiny2025_manager,
            save_time_entry,
        )
        return jsonify(payload), status_code
    except Exception as exc:
        logger.error("Chyba při zpracování hlasového příkazu: %s", exc, exc_info=True)
        return jsonify({"success": False, "error": "Interní chyba při zpracování hlasového příkazu"}), 500
