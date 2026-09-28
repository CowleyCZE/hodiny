"""Routy pro zálohy, reporty, statistiky a exporty."""

import calendar
import datetime as dt
import io

from flask import Blueprint, flash, g, redirect, render_template, request, send_file, url_for

from config import Config
from services.export_service import build_csv, build_pdf
from services.phase2_services import (
    build_statistics,
    create_backup,
    find_data_issues,
    list_backups,
    read_audit,
    restore_backup,
    write_audit,
)

reports_bp = Blueprint("reports", __name__)
CZECH_MONTHS = (
    "",
    "leden",
    "únor",
    "březen",
    "duben",
    "květen",
    "červen",
    "červenec",
    "srpen",
    "září",
    "říjen",
    "listopad",
    "prosinec",
)


@reports_bp.route("/zalohy", methods=["GET", "POST"])
def zalohy():
    """Správa záloh, výdajů a výběrů z bankomatu (příjmů)."""
    if request.method == "POST":
        try:
            form = request.form
            form_type = form.get("form_type", "zaloha")

            if form_type == "vydaj":
                category = form["expense_category"]
                amount = float(form["expense_amount"].replace(",", "."))
                currency = form["expense_currency"]
                payment_method = form["expense_payment_method"]
                description = form.get("expense_description", "")
                date_str = form["expense_date"]

                from hodiny2025_manager import Hodiny2025Manager
                hodiny_mgr = Hodiny2025Manager(g.zalohy_manager.base_path)
                hodiny_mgr.zapis_vydaje(
                    category=category,
                    amount=amount,
                    currency=currency,
                    payment_method=payment_method,
                    date_str=date_str,
                    description=description,
                )
                flash("Výdaj byl úspěšně uložen.", "success")

            elif form_type == "bankomat":
                amount = float(form["atm_amount"].replace(",", "."))
                currency = form["atm_currency"]
                date_str = form["atm_date"]

                from hodiny2025_manager import Hodiny2025Manager
                hodiny_mgr = Hodiny2025Manager(g.zalohy_manager.base_path)
                hodiny_mgr.zapis_vydaje(
                    category="Bankomat",
                    amount=amount,
                    currency=currency,
                    payment_method="Hotově",
                    date_str=date_str,
                    description="Výběr z bankomatu",
                )
                flash("Výběr z bankomatu byl úspěšně uložen.", "success")

            else:
                amount = float(form["amount"].replace(",", "."))
                g.zalohy_manager.add_or_update_employee_advance(
                    form["employee_name"], amount, form["currency"], form["option"], form["date"]
                )
                flash("Záloha byla úspěšně uložena.", "success")
        except (ValueError, IOError, KeyError) as exc:
            flash(str(exc), "error")

    return render_template(
        "zalohy.html",
        employees=g.employee_manager.get_employee_names(),
        options=g.zalohy_manager.get_option_names(),
        current_date=dt.datetime.now().strftime("%Y-%m-%d"),
    )


@reports_bp.route("/monthly_report", methods=["GET", "POST"])
def monthly_report_route():
    """Generuje měsíční agregace z týdenních listů podle zvolených zaměstnanců."""
    report_data = None
    selected_employees_post = []
    if request.method == "POST":
        try:
            month = int(request.form["month"])
            year = int(request.form["year"])
            selected_employees_post = request.form.getlist("employees")
            report_data = g.excel_manager.generate_monthly_report(month, year, selected_employees_post or None)
            if not report_data:
                flash("Nebyly nalezeny žádné záznamy.", "info")
        except (ValueError, KeyError, FileNotFoundError) as exc:
            flash(str(exc), "error")

    employee_names = [employee["name"] for employee in g.employee_manager.get_all_employees()]
    return render_template(
        "monthly_report.html",
        employee_names=employee_names,
        report_data=report_data,
        selected_employees_post=selected_employees_post,
        current_month=dt.datetime.now().month,
        current_year=dt.datetime.now().year,
    )


def _report_filters():
    today = dt.date.today()
    try:
        month = int(request.args.get("month", request.form.get("month", today.month)))
        year = int(request.args.get("year", request.form.get("year", today.year)))
    except (TypeError, ValueError) as error:
        raise ValueError("Neplatný měsíc nebo rok.") from error
    employees = request.args.getlist("employee") or request.form.getlist("employees") or None
    return month, year, employees


@reports_bp.route("/reports/export/csv")
def export_csv():
    """Stáhne měsíční report jako UTF-8 CSV."""
    try:
        month, year, employees = _report_filters()
        data = g.excel_manager.generate_monthly_report(month, year, employees)
        payload = build_csv(data, month, year, g.excel_manager.current_project_name or "")
        write_audit(Config.AUDIT_LOG_PATH, "REPORT_EXPORTED_CSV", {"month": month, "year": year})
        return send_file(io.BytesIO(payload), mimetype="text/csv; charset=utf-8", as_attachment=True,
                         download_name=f"report_{year}_{month:02d}.csv")
    except (ValueError, FileNotFoundError) as error:
        flash(str(error), "error")
        return redirect(url_for("reports.monthly_report_route"))


@reports_bp.route("/reports/export/pdf")
def export_pdf():
    """Stáhne měsíční report jako PDF."""
    try:
        month, year, employees = _report_filters()
        data = g.excel_manager.generate_monthly_report(month, year, employees)
        payload = build_pdf(data, month, year, g.excel_manager.current_project_name or "")
        write_audit(Config.AUDIT_LOG_PATH, "REPORT_EXPORTED_PDF", {"month": month, "year": year})
        return send_file(io.BytesIO(payload), mimetype="application/pdf", as_attachment=True,
                         download_name=f"report_{year}_{month:02d}.pdf")
    except (ValueError, FileNotFoundError) as error:
        flash(str(error), "error")
        return redirect(url_for("reports.monthly_report_route"))


@reports_bp.route("/statistics")
def statistics():
    """Zobrazí detailní statistiky vybraného měsíce."""
    month, year, employees = _report_filters()
    report = g.excel_manager.generate_monthly_report(month, year, employees)
    daily = g.excel_manager.generate_calendar_data(month, year, employees)
    return render_template(
        "statistics.html", statistics=build_statistics(daily, report, month, year), month=month, year=year
    )


@reports_bp.route("/issues")
def issues():
    """Zobrazí kontrolu chybějících a podezřelých záznamů."""
    today = dt.date.today()
    employees = g.employee_manager.get_vybrani_zamestnanci()
    daily = g.excel_manager.generate_calendar_data(today.month, today.year, employees)
    return render_template(
        "issues.html",
        issues=find_data_issues(daily, employees, today.replace(day=1), today),
        month=today.month,
        year=today.year,
    )


@reports_bp.route("/calendar")
def calendar_view():
    """Zobrazí denní souhrny v přehledném měsíčním kalendáři."""
    today = dt.date.today()
    try:
        month = int(request.args.get("month", today.month))
        year = int(request.args.get("year", today.year))
        selected_month = dt.date(year, month, 1)
    except (TypeError, ValueError):
        flash("Neplatný měsíc nebo rok.", "error")
        return redirect(url_for("reports.calendar_view"))

    employees = g.employee_manager.get_vybrani_zamestnanci()
    return render_template(
        "calendar.html",
        month=month,
        year=year,
        month_name=CZECH_MONTHS[month],
        weeks=calendar.monthcalendar(year, month),
        daily_data=g.excel_manager.generate_calendar_data(month, year, employees),
        previous_month=selected_month - dt.timedelta(days=1),
        next_month=(selected_month + dt.timedelta(days=32)).replace(day=1),
        today=today.isoformat(),
    )


@reports_bp.route("/backups", methods=["GET", "POST"])
def backups():
    """Správa vytvoření a bezpečné obnovy XLSX záloh."""
    if request.method == "POST":
        action = request.form.get("action")
        try:
            if action == "create":
                backup = create_backup(g.excel_manager.get_active_file_path(), Config.BACKUP_PATH, "manual")
                write_audit(Config.AUDIT_LOG_PATH, "BACKUP_CREATED", {"name": backup.name})
                flash("Záloha byla vytvořena.", "success")
            elif action == "restore":
                backup_name = request.form.get("backup_name", "")
                create_backup(g.excel_manager.get_active_file_path(), Config.BACKUP_PATH, "before_restore")
                restore_backup(Config.BACKUP_PATH, backup_name, g.excel_manager.get_active_file_path())
                write_audit(Config.AUDIT_LOG_PATH, "BACKUP_RESTORED", {"name": backup_name})
                flash("Záloha byla obnovena.", "success")
        except (ValueError, FileNotFoundError, OSError) as error:
            flash(str(error), "error")
    return render_template("backups.html", backups=list_backups(Config.BACKUP_PATH))


@reports_bp.route("/audit")
def audit():
    """Zobrazí poslední auditní události."""
    events = g.database.list_audit() or read_audit(Config.AUDIT_LOG_PATH)
    return render_template("audit.html", events=events)
