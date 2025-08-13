# app.py
import json
import smtplib
from datetime import datetime, timedelta
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from flask import (Flask, flash, redirect, render_template, request,
                   send_file, session, url_for, g)
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from config import Config
from employee_management import EmployeeManager
from excel_manager import ExcelManager
from hodiny2025_manager import Hodiny2025Manager
from utils.logger import setup_logger
from zalohy_manager import ZalohyManager

logger = setup_logger("app")

app = Flask(__name__)
app.secret_key = Config.SECRET_KEY
Config.init_app(app)


def save_settings_to_file(settings_data):
    try:
        with open(Config.SETTINGS_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(settings_data, f, indent=4, ensure_ascii=False)
        return True
    except (IOError, Exception) as e:
        logger.error(f"Chyba při ukládání nastavení: {e}", exc_info=True)
        return False


def load_settings_from_file():
    if not Config.SETTINGS_FILE_PATH.exists():
        return Config.get_default_settings()
    try:
        with open(Config.SETTINGS_FILE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, Exception) as e:
        logger.error(f"Chyba při načítání nastavení: {e}", exc_info=True)
        return Config.get_default_settings()


@app.before_request
def before_request():
    session['settings'] = load_settings_from_file()
    g.employee_manager = EmployeeManager(Config.DATA_PATH)
    g.excel_manager = ExcelManager(Config.EXCEL_BASE_PATH)
    g.zalohy_manager = ZalohyManager(Config.EXCEL_BASE_PATH)
    g.hodiny2025_manager = Hodiny2025Manager(Config.EXCEL_BASE_PATH)

    # Archivace na začátku týdne
    current_week = datetime.now().isocalendar().week
    if g.excel_manager.archive_if_needed(current_week, session['settings']):
        save_settings_to_file(session['settings'])
        flash(f"Týden {session['settings']['last_archived_week'] - 1} byl archivován.", "info")


@app.teardown_request
def teardown_request(exception=None):
    if hasattr(g, 'excel_manager') and g.excel_manager:
        g.excel_manager.close_cached_workbooks()


@app.route("/")
def index():
    active_filename = Config.EXCEL_TEMPLATE_NAME
    week_num_int = datetime.now().isocalendar().week
    current_date = datetime.now().strftime("%Y-%m-%d")
    return render_template("index.html",
                           active_filename=active_filename,
                           week_number=week_num_int,
                           current_date=current_date)


@app.route("/download")
def download_file():
    try:
        return send_file(g.excel_manager.get_active_file_path(), as_attachment=True)
    except Exception as e:
        logger.error(f"Chyba při stahování souboru: {e}", exc_info=True)
        flash("Chyba při stahování souboru.", "error")
        return redirect(url_for("index"))


@app.route("/send_email", methods=["POST"])
def send_email():
    try:
        recipient = Config.RECIPIENT_EMAIL or ""
        sender = Config.SMTP_USERNAME or ""
        if not all([recipient, sender, Config.SMTP_PASSWORD, Config.SMTP_SERVER, Config.SMTP_PORT]):
            raise ValueError("SMTP údaje nejsou kompletní.")

        msg = MIMEMultipart()
        msg["Subject"] = f'Výkaz práce - {datetime.now().strftime("%Y-%m-%d")}'
        msg["From"] = sender
        msg["To"] = recipient
        msg.attach(MIMEText("V příloze zasílám výkaz práce.", "plain", "utf-8"))

        with open(g.excel_manager.get_active_file_path(), "rb") as f:
            attachment = MIMEApplication(f.read(), _subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            attachment.add_header("Content-Disposition", "attachment", filename=g.excel_manager.active_filename)
            msg.attach(attachment)

        with smtplib.SMTP_SSL(Config.SMTP_SERVER, Config.SMTP_PORT, timeout=Config.SMTP_TIMEOUT) as smtp:
            smtp.login(sender, Config.SMTP_PASSWORD if Config.SMTP_PASSWORD is not None else "")
            smtp.send_message(msg)
        flash("Email byl úspěšně odeslán.", "success")
    except (ValueError, smtplib.SMTPException, Exception) as e:
        logger.error(f"Chyba při odesílání emailu: {e}", exc_info=True)
        flash("Chyba při odesílání emailu.", "error")
    return redirect(url_for("index"))


@app.route("/zamestnanci", methods=["GET", "POST"])
def manage_employees():
    if request.method == "POST":
        action = request.form.get("action")
        try:
            if action == "add":
                g.employee_manager.pridat_zamestnance(request.form.get("name", "").strip())
            elif action == "select":
                name = request.form.get("employee_name", "")
                if name in g.employee_manager.vybrani_zamestnanci:
                    g.employee_manager.odebrat_vybraneho_zamestnance(name)
                else:
                    g.employee_manager.pridat_vybraneho_zamestnance(name)
            elif action == "edit":
                old_name = request.form.get("old_name", "").strip()
                new_name = request.form.get("new_name", "").strip()
                g.employee_manager.upravit_zamestnance_podle_jmena(old_name, new_name)
            elif action == "delete":
                g.employee_manager.smazat_zamestnance_podle_jmena(request.form.get("employee_name", ""))
        except (ValueError, Exception) as e:
            flash(str(e), "error")
    return render_template("employees.html", employees=g.employee_manager.get_all_employees())


@app.route("/zaznam", methods=["GET", "POST"])
def record_time():
    selected_employees = g.employee_manager.get_vybrani_zamestnanci()
    if not selected_employees:
        flash("Nejsou vybráni žádní zaměstnanci.", "warning")
        return redirect(url_for("manage_employees"))

    current_date = request.args.get('next_date', datetime.now().strftime("%Y-%m-%d"))
    start_time = session['settings'].get("start_time", "07:00")
    end_time = session['settings'].get("end_time", "18:00")
    lunch_duration = str(session['settings'].get("lunch_duration", 1.0))
    is_free_day = False

    if request.method == "POST":
        current_date = request.form.get("date", current_date)
        start_time = request.form.get("start_time", start_time)
        end_time = request.form.get("end_time", end_time)
        lunch_duration = request.form.get("lunch_duration", lunch_duration)
        is_free_day = request.form.get("is_free_day") == "on"

        try:
            date = datetime.strptime(current_date, "%Y-%m-%d").date()
            if is_free_day:
                # Pro volný den nastavíme 0 hodin
                g.excel_manager.ulozit_pracovni_dobu(
                    current_date, "00:00", "00:00", "0", selected_employees)
                g.hodiny2025_manager.zapis_pracovni_doby(
                    current_date, "00:00", "00:00", "0", len(selected_employees))
            else:
                g.excel_manager.ulozit_pracovni_dobu(
                    current_date, start_time, end_time, lunch_duration, selected_employees)
                g.hodiny2025_manager.zapis_pracovni_doby(
                    current_date, start_time, end_time, lunch_duration, len(selected_employees))

            flash("Záznam byl úspěšně uložen.", "success")
            next_day = (date + timedelta(days=1))
            while next_day.weekday() >= 5:
                next_day += timedelta(days=1)
            return redirect(url_for('record_time', next_date=next_day.strftime("%Y-%m-%d")))
        except (ValueError, IOError, Exception) as e:
            flash(str(e), "error")

    return render_template("record_time.html",
                           selected_employees=selected_employees,
                           current_date=current_date,
                           start_time=start_time,
                           end_time=end_time,
                           lunch_duration=lunch_duration,
                           is_free_day=is_free_day)


@app.route("/excel_viewer", methods=["GET"])
def excel_viewer():
    active_sheet_name = request.args.get("sheet")
    data, sheet_names = [], []
    try:
        wb = load_workbook(g.excel_manager.get_active_file_path(), read_only=True, data_only=True)
        sheet_names = wb.sheetnames
        active_sheet_name = active_sheet_name if active_sheet_name in sheet_names else sheet_names[0]
        sheet = wb[active_sheet_name]
        for i, row in enumerate(sheet.iter_rows(values_only=True)):
            if i >= Config.MAX_ROWS_TO_DISPLAY_EXCEL_VIEWER:
                break
            data.append([str(c) if c is not None else "" for c in row])
        wb.close()
    except (FileNotFoundError, InvalidFileException, Exception) as e:
        flash(f"Chyba při zobrazení souboru: {e}", "error")
        return redirect(url_for("index"))

    return render_template("excel_viewer.html", sheet_names=sheet_names,
                           active_sheet=active_sheet_name, data=data)


@app.route("/settings", methods=["GET", "POST"])
def settings_page():
    if request.method == "POST":
        try:
            settings_to_save = session['settings'].copy()
            settings_to_save.update({
                "start_time": request.form["start_time"],
                "end_time": request.form["end_time"],
                "lunch_duration": float(request.form["lunch_duration"].replace(",", ".")),
            })
            save_settings_to_file(settings_to_save)
            session['settings'] = settings_to_save
            flash("Nastavení bylo úspěšně uloženo.", "success")
        except (ValueError, Exception) as e:
            flash(str(e), "error")

    return render_template("settings_page.html", settings=session.get('settings', {}))


@app.route("/zalohy", methods=["GET", "POST"])
def zalohy():
    if request.method == "POST":
        try:
            form = request.form
            amount = float(form["amount"].replace(",", "."))
            g.zalohy_manager.add_or_update_employee_advance(
                form["employee_name"], amount, form["currency"],
                form["option"], form["date"])
            flash("Záloha byla úspěšně uložena.", "success")
        except (ValueError, Exception) as e:
            flash(str(e), "error")

    return render_template("zalohy.html",
                           employees=g.employee_manager.zamestnanci,
                           options=g.zalohy_manager.get_option_names(),
                           current_date=datetime.now().strftime("%Y-%m-%d"))


@app.route('/monthly_report', methods=['GET', 'POST'])
def monthly_report_route():
    report_data = None
    if request.method == 'POST':
        try:
            month = int(request.form['month'])
            year = int(request.form['year'])
            employees = request.form.getlist('employees') or None
            report_data = g.excel_manager.generate_monthly_report(month, year, employees)
            if not report_data:
                flash('Nebyly nalezeny žádné záznamy.', 'info')
        except (ValueError, Exception) as e:
            flash(str(e), 'error')

    employee_names = [emp['name'] for emp in g.employee_manager.get_all_employees()]
    return render_template("monthly_report.html",
                           employee_names=employee_names,
                           report_data=report_data,
                           current_month=datetime.now().month,
                           current_year=datetime.now().year)


if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=5000)
