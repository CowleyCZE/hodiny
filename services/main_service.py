"""Služby pro hlavní dashboard, záznam pracovní doby a odesílání e-mailu."""

import datetime as dt
import smtplib
import time
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from config import Config
from performance_optimizations import invalidate_excel_status_cache, optimize_excel_operations, perf_monitor


def cleanup_temp_files(base_path):
    """Odstraní dočasné upload soubory starší než jednu hodinu."""
    current_time = time.time()
    for temp_file in base_path.glob("temp_*.xlsx"):
        if current_time - temp_file.stat().st_mtime > 3600:
            temp_file.unlink()


def build_dashboard_context(excel_manager, settings):
    """Sestaví data pro úvodní dashboard."""
    cleanup_temp_files(Config.EXCEL_BASE_PATH)

    request_start_time = time.time()
    current_datetime = dt.datetime.now()
    excel_exists = False
    current_week_data = None

    try:
        excel_exists = optimize_excel_operations()
        if excel_exists:
            current_week_data = excel_manager.get_current_week_data()
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
    if is_free_day:
        excel_manager.ulozit_pracovni_dobu(date, "00:00", "00:00", "0", employees)
        hodiny2025_manager.zapis_pracovni_doby(date, "00:00", "00:00", "0", len(employees))
        invalidate_excel_status_cache()
        return f"Volný den pro {date} byl zaznamenán pro {len(employees)} zaměstnanců"

    if not start_time or not end_time:
        raise ValueError("Chybí čas začátku nebo konce")

    excel_manager.ulozit_pracovni_dobu(date, start_time, end_time, lunch_duration, employees)
    hodiny2025_manager.zapis_pracovni_doby(date, start_time, end_time, lunch_duration, len(employees))
    invalidate_excel_status_cache()
    return f"Pracovní doba pro {date} byla zaznamenána pro {len(employees)} zaměstnanců"


def get_next_workday(current_date):
    """Vrátí nejbližší další pracovní den."""
    next_day = current_date + dt.timedelta(days=1)
    while next_day.weekday() >= 5:
        next_day += dt.timedelta(days=1)
    return next_day
