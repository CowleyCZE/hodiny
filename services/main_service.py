"""Služby pro hlavní dashboard, záznam pracovní doby a odesílání e-mailu."""

import datetime as dt
import logging
import smtplib
import time
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

from config import Config
from performance_optimizations import invalidate_excel_status_cache, optimize_excel_operations, perf_monitor
from services.phase2_services import create_backup, find_data_issues, write_audit
from services.sync_service import sync_excel_state

logger = logging.getLogger(__name__)


def cleanup_temp_files(base_path):
    """Odstraní dočasné upload soubory starší než jednu hodinu."""
    current_time = time.time()
    for temp_file in base_path.glob("temp_*.xlsx"):
        if current_time - temp_file.stat().st_mtime > 3600:
            temp_file.unlink()


def build_dashboard_context(excel_manager, settings, employees=None):
    """Sestaví data pro úvodní dashboard."""
    cleanup_temp_files(Config.EXCEL_BASE_PATH)

    request_start_time = time.time()
    current_datetime = dt.datetime.now()
    excel_exists = False
    current_week_data = None
    period_summary = {"today_hours": 0.0, "week_hours": 0.0, "month_hours": 0.0, "today_free": False}
    data_issues = []

    try:
        excel_exists = optimize_excel_operations()
        if excel_exists:
            current_week_data = excel_manager.get_current_week_data()
            month_data = excel_manager.generate_calendar_data(current_datetime.month, current_datetime.year, employees)
            month_report = excel_manager.generate_monthly_report(
                current_datetime.month, current_datetime.year, employees
            )
            today = month_data.get(current_datetime.date().isoformat(), {})
            period_summary["today_hours"] = today.get("hours", 0.0)
            period_summary["today_free"] = today.get("free_days", 0) > 0 and today.get("hours", 0.0) == 0
            period_summary["month_hours"] = round(sum(item["total_hours"] for item in month_report.values()), 2)
            week_data = excel_manager.generate_calendar_data(
                current_datetime.month, current_datetime.year, employees
            )
            period_summary["week_hours"] = round(
                sum(
                    item["hours"]
                    for date_key, item in week_data.items()
                    if dt.date.fromisoformat(date_key).isocalendar().week == current_datetime.isocalendar().week
                ),
                2,
            )
            month_start = current_datetime.date().replace(day=1)
            data_issues = find_data_issues(month_data, employees or [], month_start, current_datetime.date())
    except Exception:
        excel_exists = False

    context = {
        "active_filename": Config.EXCEL_TEMPLATE_NAME,
        "week_number": current_datetime.isocalendar().week,
        "current_date": current_datetime.strftime("%Y-%m-%d"),
        "current_date_formatted": current_datetime.strftime("%d.%m.%Y"),
        "excel_exists": excel_exists,
        "project_name": settings.get("project_info", {}).get("name", "Nepojmenovaný projekt"),
        "current_week_data": current_week_data,
        "start_time": settings.get("start_time", "07:00"),
        "end_time": settings.get("end_time", "18:00"),
        "lunch_duration": settings.get("lunch_duration", 1.0),
        "period_summary": period_summary,
        "selected_employees": employees or [],
        "data_issues": data_issues,
    }

    perf_monitor.record_request("index", time.time() - request_start_time)
    return context


def send_active_excel_email(excel_manager):
    """Odešle aktivní Excel soubor na konfigurovaný příjemce."""
    recipient = Config.RECIPIENT_EMAIL or ""
    sender = Config.SMTP_USERNAME or ""

    if not all([recipient, sender, Config.SMTP_PASSWORD, Config.SMTP_SERVER, Config.SMTP_PORT]):
        raise ValueError("SMTP údaje nejsou kompletní.")

    message = MIMEMultipart()
    message["Subject"] = f'Výkaz práce - {dt.datetime.now().strftime("%Y-%m-%d")}'
    message["From"] = sender
    message["To"] = recipient
    message.attach(MIMEText("V příloze zasílám výkaz práce.", "plain", "utf-8"))

    with open(excel_manager.get_active_file_path(), "rb") as excel_file:
        attachment = MIMEApplication(
            excel_file.read(),
            _subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        attachment.add_header("Content-Disposition", "attachment", filename=excel_manager.active_filename)
        message.attach(attachment)

    with smtplib.SMTP_SSL(Config.SMTP_SERVER, Config.SMTP_PORT, timeout=Config.SMTP_TIMEOUT) as smtp:
        smtp.login(sender, Config.SMTP_PASSWORD if Config.SMTP_PASSWORD is not None else "")
        smtp.send_message(message)


def save_time_entry(
    excel_manager,
    hodiny2025_manager,
    date,
    start_time,
    end_time,
    lunch_duration,
    employees,
    is_free_day,
):
    """Zapíše pracovní dobu nebo volný den do všech relevantních workbooků."""
    del hodiny2025_manager

    active_path = excel_manager.get_active_file_path()
    if isinstance(active_path, (str, Path)):
        try:
            create_backup(active_path, Config.BACKUP_PATH, "before_entry")
        except (FileNotFoundError, OSError) as error:
            # Záloha nesmí zabránit běžnému zápisu, pokud soubor ještě neexistuje.
            logger.warning("Automatickou zálohu se nepodařilo vytvořit: %s", error)

    if is_free_day:
        success = excel_manager.ulozit_pracovni_dobu(date, "00:00", "00:00", "0", employees)
        if not success:
            raise IOError("Nepodařilo se uložit volný den do Excel souboru.")
        invalidate_excel_status_cache()
        write_audit(Config.AUDIT_LOG_PATH, "FREE_DAY_CREATED", {"date": date, "employees": employees})
        _mirror_entry(date, "00:00", "00:00", 0, True, employees, excel_manager)
        return f"Volný den pro {date} byl zaznamenán pro {len(employees)} zaměstnanců"

    if not start_time or not end_time:
        raise ValueError("Chybí čas začátku nebo konce")

    try:
        start = dt.datetime.strptime(start_time, "%H:%M")
        end = dt.datetime.strptime(end_time, "%H:%M")
        lunch = float(str(lunch_duration).replace(",", "."))
    except (TypeError, ValueError) as error:
        raise ValueError("Čas musí být ve formátu HH:MM a přestávka číslo.") from error
    if end <= start:
        raise ValueError("Čas odchodu musí být později než čas příchodu.")
    if not 0 <= lunch <= 4:
        raise ValueError("Přestávka musí být v rozmezí 0 až 4 hodin.")
    if (end - start).total_seconds() / 3600 - lunch > 24:
        raise ValueError("Pracovní doba je neobvykle dlouhá.")

    success = excel_manager.ulozit_pracovni_dobu(date, start_time, end_time, lunch_duration, employees)
    if not success:
        raise IOError("Nepodařilo se uložit pracovní dobu do Excel souboru.")
    invalidate_excel_status_cache()
    write_audit(
        Config.AUDIT_LOG_PATH,
        "TIME_ENTRY_CREATED",
        {"date": date, "start": start_time, "end": end_time, "lunch": lunch_duration, "employees": employees},
    )
    _mirror_entry(date, start_time, end_time, lunch_duration, False, employees, excel_manager)
    return f"Pracovní doba pro {date} byla zaznamenána pro {len(employees)} zaměstnanců"


def _mirror_entry(date, start_time, end_time, lunch_duration, is_free_day, employees, excel_manager):
    """Best-effort zrcadlení do SQLite; Excel zůstává zpětně kompatibilním výkazem."""
    try:
        from flask import g, has_request_context

        if has_request_context() and hasattr(g, "database"):
            g.database.record_time_entry(
                date, start_time, end_time, lunch_duration, is_free_day, employees, source="excel"
            )
            sync_excel_state(g.database, excel_manager, employees)
    except (OSError, RuntimeError, AttributeError, TypeError) as error:
        logger.warning("SQLite zrcadlení se nepodařilo dokončit: %s", error)


def get_next_workday(current_date):
    """Vrátí nejbližší další pracovní den."""
    next_day = current_date + dt.timedelta(days=1)
    while next_day.weekday() >= 5:
        next_day += dt.timedelta(days=1)
    return next_day
