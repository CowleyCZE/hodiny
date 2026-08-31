"""Routy pro zálohy a reporty."""

import datetime as dt

from flask import Blueprint, flash, g, render_template, request

reports_bp = Blueprint("reports", __name__)


@reports_bp.route("/zalohy", methods=["GET", "POST"])
def zalohy():
    """Správa záloh, výdajů a výběrů z bankomatu (příjmů)."""
    if request.method == "POST":
        try:
            form = request.form
            form_type = form.get("form_type", "zaloha")

            if form_type == "vydaj":
                category = form.get("expense_category", "")
                amount_str = form.get("expense_amount", "")
                currency = form.get("expense_currency", "EUR")
                payment_method = form.get("expense_payment_method", "Hotově")
                date_str = form.get("expense_date", "")
                description = form.get("expense_description", "")

                if not amount_str:
                    raise ValueError("Částka je povinný údaj.")
                amount = float(amount_str.replace(",", "."))

                g.hodiny2025_manager.zapis_vydaje(
                    category=category,
                    amount=amount,
                    currency=currency,
                    payment_method=payment_method,
                    date_str=date_str,
                    description=description,
                )
                flash("Výdaj byl úspěšně uložen do Excelu.", "success")

            elif form_type == "bankomat":
                amount_str = form.get("atm_amount", "")
                currency = form.get("atm_currency", "EUR")
                date_str = form.get("atm_date", "")

                if not amount_str:
                    raise ValueError("Částka je povinný údaj.")
                amount = float(amount_str.replace(",", "."))

                g.hodiny2025_manager.zapis_vydaje(
                    category="Bankomat",
                    amount=amount,
                    currency=currency,
                    payment_method="Kartou",
                    date_str=date_str,
                    description="Výběr z bankomatu",
                )
                flash("Výběr z bankomatu byl úspěšně uložen do Excelu.", "success")

            else:
                amount_str = form.get("amount", "")
                if not amount_str:
                    raise ValueError("Částka je povinný údaj.")
                amount = float(amount_str.replace(",", "."))
                g.zalohy_manager.add_or_update_employee_advance(
                    form.get("employee_name", ""),
                    amount,
                    form.get("currency", ""),
                    form.get("option", ""),
                    form.get("date", ""),
                )
                flash("Záloha byla úspěšně uložena.", "success")
        except (ValueError, KeyError, IOError) as exc:
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
    if request.method == "POST":
        try:
            month_val = request.form.get("month")
            year_val = request.form.get("year")
            if not month_val or not year_val:
                raise ValueError("Měsíc a rok jsou povinné údaje.")
            month = int(month_val)
            year = int(year_val)
            employees = request.form.getlist("employees") or None
            report_data = g.excel_manager.generate_monthly_report(month, year, employees)
            if not report_data:
                flash("Nebyly nalezeny žádné záznamy.", "info")
        except (ValueError, KeyError, FileNotFoundError) as exc:
            flash(str(exc), "error")

    employee_names = [employee["name"] for employee in g.employee_manager.get_all_employees()]
    return render_template(
        "monthly_report.html",
        employee_names=employee_names,
        report_data=report_data,
        current_month=dt.datetime.now().month,
        current_year=dt.datetime.now().year,
    )
