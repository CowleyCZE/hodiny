# app.py
import json
import logging
import os
import re
import smtplib
import ssl
from datetime import datetime, timedelta
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from functools import wraps

from flask import (Flask, flash, jsonify, redirect, render_template, request,
                   send_file, session, url_for, g, get_flashed_messages)
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook import Workbook

from config import Config
from employee_management import EmployeeManager
from excel_manager import ExcelManager
from utils.logger import setup_logger
from zalohy_manager import ZalohyManager
from utils.voice_processor import VoiceProcessor

logger = setup_logger("app")

app = Flask(__name__)
app.secret_key = Config.SECRET_KEY
Config.init_app(app)


def save_settings_to_file(settings_data):
    try:
        Config.SETTINGS_FILE_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(Config.SETTINGS_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(settings_data, f, indent=4, ensure_ascii=False)
        logger.info(f"Nastavení uložena do souboru: {Config.SETTINGS_FILE_PATH}")
        return True
    except (IOError, Exception) as e:
        logger.error(f"Chyba při ukládání nastavení: {e}", exc_info=True)
        return False


def load_settings_from_file():
    default_settings = Config.get_default_settings()
    if not Config.SETTINGS_FILE_PATH.exists():
        logger.warning(f"Soubor s nastavením '{Config.SETTINGS_FILE_PATH}' nenalezen, použijí se výchozí.")
        save_settings_to_file(default_settings)
        return default_settings

    try:
        with open(Config.SETTINGS_FILE_PATH, "r", encoding="utf-8") as f:
            loaded_settings = json.load(f)
        if not isinstance(loaded_settings, dict):
            raise ValueError("Neplatný formát JSON.")
        
        settings = default_settings.copy()
        settings.update(loaded_settings)
        
        if not isinstance(settings.get("project_info"), dict):
            settings["project_info"] = default_settings["project_info"]
        if not isinstance(settings.get("active_excel_file"), (str, type(None))):
            settings["active_excel_file"] = None
        
        logger.info(f"Nastavení načtena ze souboru: {Config.SETTINGS_FILE_PATH}")
        return settings
    except (json.JSONDecodeError, ValueError, Exception) as e:
        logger.error(f"Chyba při načítání nastavení: {e}", exc_info=True)
        save_settings_to_file(default_settings)
        return default_settings.copy()


def ensure_active_excel_file(settings):
    fixed_filename = Config.EXCEL_TEMPLATE_NAME
    fixed_file_path = Config.EXCEL_BASE_PATH / fixed_filename

    if not fixed_file_path.exists():
        try:
            Config.EXCEL_BASE_PATH.mkdir(parents=True, exist_ok=True)
            wb = Workbook()
            wb.create_sheet(Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME)
            wb.create_sheet(Config.EXCEL_ADVANCES_SHEET_NAME)
            if "Sheet" in wb.sheetnames:
                wb.remove(wb["Sheet"])
            wb.save(fixed_file_path)
            logger.info(f"Vytvořen nový Excel soubor: {fixed_file_path}")
        except Exception as e:
            logger.error(f"Nepodařilo se vytvořit Excel soubor '{fixed_file_path}': {e}", exc_info=True)
            flash("Chyba při vytváření Excel souboru.", "error")
            return settings

    if settings.get("active_excel_file") != fixed_filename:
        settings["active_excel_file"] = fixed_filename
        if not save_settings_to_file(settings):
            flash("Nepodařilo se uložit nastavení aktivního souboru.", "error")

    return settings


@app.before_request
def before_request():
    settings = load_settings_from_file()
    settings = ensure_active_excel_file(settings)
    session['settings'] = settings

    g.employee_manager = EmployeeManager(Config.DATA_PATH)
    active_filename = settings.get("active_excel_file")
    if active_filename:
        try:
            g.excel_manager = ExcelManager(Config.EXCEL_BASE_PATH, active_filename, Config.EXCEL_TEMPLATE_NAME)
            project_name = settings.get("project_info", {}).get("name")
            if project_name:
                g.excel_manager.set_project_name(project_name)
            g.zalohy_manager = ZalohyManager(Config.EXCEL_BASE_PATH, active_filename)
        except (ValueError, FileNotFoundError, Exception) as e:
            logger.error(f"Chyba při inicializaci manažerů pro soubor '{active_filename}': {e}", exc_info=True)
            g.excel_manager = None
            g.zalohy_manager = None
            flash(f"Chyba při inicializaci pracovního souboru '{active_filename}'.", "error")
    else:
        g.excel_manager = None
        g.zalohy_manager = None


@app.teardown_request
def teardown_request(exception=None):
    if hasattr(g, 'excel_manager') and g.excel_manager:
        g.excel_manager.close_cached_workbooks()


def require_excel_managers(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not g.get('excel_manager') or not g.get('zalohy_manager'):
            flash("Chyba: Není definován aktivní Excel soubor. Zkuste archivovat a začít nový.", "error")
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function


@app.route("/")
def index():
    settings = session.get('settings', {})
    active_filename = settings.get('active_excel_file')
    excel_exists = False
    week_num_int = 0
    current_date = datetime.now().strftime("%Y-%m-%d")

    if active_filename and g.get('excel_manager'):
        excel_exists = (Config.EXCEL_BASE_PATH / active_filename).exists()
        week_calendar_data = g.excel_manager.ziskej_cislo_tydne(current_date)
        week_num_int = week_calendar_data.week if week_calendar_data else 0
    
    return render_template("index.html", excel_exists=excel_exists, active_filename=active_filename,
                           week_number=week_num_int, current_date=current_date)


@app.route("/download")
@require_excel_managers
def download_file():
    try:
        active_file_path = g.excel_manager.get_active_file_path()
        max_week_number = 0
        with load_workbook(active_file_path, read_only=True) as wb:
            week_pattern = re.compile(r"Týden (\d+)")
            for sheet_name in wb.sheetnames:
                match = week_pattern.match(sheet_name)
                if match:
                    max_week_number = max(max_week_number, int(match.group(1)))
        
        template_stem = Path(Config.EXCEL_TEMPLATE_NAME).stem
        download_filename = f"{template_stem}_Tyden_{max_week_number}.xlsx" if max_week_number > 0 else Config.EXCEL_TEMPLATE_NAME
        
        return send_file(str(active_file_path), as_attachment=True, download_name=download_filename)

    except (FileNotFoundError, ValueError, IOError, Exception) as e:
        logger.error(f"Chyba při stahování souboru: {e}", exc_info=True)
        flash("Chyba při stahování souboru.", "error")
        return redirect(url_for("index"))


@app.route("/send_email", methods=["POST"])
@require_excel_managers
def send_email():
    try:
        active_file_path = g.excel_manager.get_active_file_path()
        recipient = Config.RECIPIENT_EMAIL
        sender = Config.SMTP_USERNAME

        if not all([recipient, sender, Config.SMTP_PASSWORD, Config.SMTP_SERVER, Config.SMTP_PORT]):
            raise ValueError("SMTP údaje nebo e-mail příjemce nejsou kompletní.")
        if not validate_email(sender) or not validate_email(recipient):
            raise ValueError("Neplatná e-mailová adresa.")

        msg = MIMEMultipart()
        msg["Subject"] = f'{active_file_path.name} - {datetime.now().strftime("%Y-%m-%d %H:%M")}'
        msg["From"] = sender
        msg["To"] = recipient
        msg.attach(MIMEText(f"Dobrý den,\n\nv příloze zasílám aktuální výkaz ({active_file_path.name}).\n\nS pozdravem,\n{Config.DEFAULT_APP_NAME}", "plain", "utf-8"))

        with open(active_file_path, "rb") as f:
            attachment = MIMEApplication(f.read(), _subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            attachment.add_header("Content-Disposition", "attachment", filename=active_file_path.name)
            msg.attach(attachment)

        with smtplib.SMTP_SSL(Config.SMTP_SERVER, Config.SMTP_PORT, context=ssl.create_default_context(), timeout=Config.SMTP_TIMEOUT) as smtp:
            smtp.login(sender, Config.SMTP_PASSWORD)
            smtp.send_message(msg)

        flash("Email byl úspěšně odeslán.", "success")
    except (ValueError, FileNotFoundError, smtplib.SMTPException, Exception) as e:
        logger.error(f"Chyba při odesílání emailu: {e}", exc_info=True)
        flash(f"Chyba při odesílání emailu: {e}", "error")

    return redirect(url_for("index"))


@app.route("/zamestnanci", methods=["GET", "POST"])
def manage_employees():
    if not g.get('employee_manager'):
        flash("Správce zaměstnanců není k dispozici.", "error")
        return redirect(url_for('index'))

    if request.method == "POST":
        action = request.form.get("action")
        try:
            if action == "add":
                name = request.form.get("name", "").strip()
                if not name or len(name) > Config.EMPLOYEE_NAME_MAX_LENGTH or not re.match(Config.EMPLOYEE_NAME_VALIDATION_REGEX, name):
                    raise ValueError("Neplatné jméno zaměstnance.")
                if g.employee_manager.pridat_zamestnance(name):
                    flash(f'Zaměstnanec "{name}" byl přidán.', "success")
                else:
                    flash(f'Zaměstnanec "{name}" již existuje.', "error")

            elif action == "select":
                name = request.form.get("employee_name", "")
                if name in g.employee_manager.vybrani_zamestnanci:
                    g.employee_manager.odebrat_vybraneho_zamestnance(name)
                    flash(f'"{name}" odebrán z výběru.', "info")
                else:
                    g.employee_manager.pridat_vybraneho_zamestnance(name)
                    flash(f'"{name}" přidán do výběru.', "success")

            elif action == "edit":
                old_name = request.form.get("old_name", "").strip()
                new_name = request.form.get("new_name", "").strip()
                if not new_name or len(new_name) > Config.EMPLOYEE_NAME_MAX_LENGTH or not re.match(Config.EMPLOYEE_NAME_VALIDATION_REGEX, new_name):
                    raise ValueError("Neplatné nové jméno.")
                if g.employee_manager.upravit_zamestnance_podle_jmena(old_name, new_name):
                    flash(f'"{old_name}" upraven na "{new_name}".', "success")
                else:
                    flash(f'Nepodařilo se upravit "{old_name}".', "error")

            elif action == "delete":
                name = request.form.get("employee_name", "")
                if g.employee_manager.smazat_zamestnance_podle_jmena(name):
                    flash(f'Zaměstnanec "{name}" byl smazán.', "success")
                else:
                    flash(f'Nepodařilo se smazat "{name}".', "error")
            
            return redirect(url_for('manage_employees'))

        except (ValueError, Exception) as e:
            flash(str(e), "error")
            logger.error(f"Chyba při správě zaměstnanců (akce: {action}): {e}", exc_info=True)

    return render_template("employees.html", employees=g.employee_manager.get_all_employees())


@app.route("/zaznam", methods=["GET", "POST"])
@require_excel_managers
def record_time():
    selected_employees = g.employee_manager.get_vybrani_zamestnanci()
    if not selected_employees:
        flash("Nejsou vybráni žádní zaměstnanci.", "warning")
        return redirect(url_for("manage_employees"))

    settings = session.get('settings', {})
    form_data = {
        "date": request.args.get('next_date', datetime.now().strftime("%Y-%m-%d")),
        "start_time": settings.get("start_time", "07:00"),
        "end_time": settings.get("end_time", "18:00"),
        "lunch_duration": str(settings.get("lunch_duration", 1.0)),
        "is_free_day": False
    }

    if request.method == "POST":
        form_data.update(request.form.to_dict())
        form_data["is_free_day"] = "is_free_day" in request.form
        try:
            date = datetime.strptime(form_data["date"], "%Y-%m-%d").date()
            if date > datetime.now().date():
                raise ValueError("Nelze zadat budoucí datum.")
            
            if form_data["is_free_day"]:
                start_time, end_time, lunch = "00:00", "00:00", 0.0
            else:
                start = datetime.strptime(form_data["start_time"], "%H:%M")
                end = datetime.strptime(form_data["end_time"], "%H:%M")
                if end <= start:
                    raise ValueError("Konec musí být po začátku.")
                lunch = float(form_data["lunch_duration"].replace(",", "."))
                if not (0 <= lunch <= 4 and lunch < (end - start).total_seconds() / 3600):
                    raise ValueError("Neplatná délka pauzy.")
                start_time, end_time = form_data["start_time"], form_data["end_time"]

            if g.excel_manager.ulozit_pracovni_dobu(form_data["date"], start_time, end_time, lunch, selected_employees):
                flash("Záznam byl úspěšně uložen.", "success")
                next_day = (date + timedelta(days=1))
                while next_day.weekday() >= 5:
                    next_day += timedelta(days=1)
                return redirect(url_for('record_time', next_date=next_day.strftime("%Y-%m-%d")))
            else:
                raise IOError("Nepodařilo se uložit záznam do Excelu.")

        except (ValueError, IOError, Exception) as e:
            flash(str(e), "error")
            logger.error(f"Chyba při záznamu času: {e}", exc_info=True)

    return render_template("record_time.html", selected_employees=selected_employees, form_data=form_data,
                           excel_files=sorted([f.name for f in Config.EXCEL_BASE_PATH.glob('*.xlsx')], reverse=True),
                           active_excel_file=session.get('settings', {}).get("active_excel_file"))


@app.route("/set_active_file", methods=["POST"])
def set_active_file():
    selected_file = request.form.get("excel_file")
    if selected_file and (Config.EXCEL_BASE_PATH / selected_file).exists():
        settings = load_settings_from_file()
        settings["active_excel_file"] = selected_file
        if save_settings_to_file(settings):
            session['settings'] = settings
            flash(f"Aktivní soubor nastaven na '{selected_file}'.", "success")
    else:
        flash("Vybraný soubor neexistuje.", "error")
    return redirect(url_for('record_time'))


@app.route("/rename_project", methods=["POST"])
def rename_project():
    old_filename = request.form.get("old_excel_file")
    new_filename = request.form.get("new_excel_file", "").strip()
    if not new_filename.endswith(".xlsx"): new_filename += ".xlsx"

    old_path = Config.EXCEL_BASE_PATH / old_filename
    new_path = Config.EXCEL_BASE_PATH / new_filename

    if old_path.exists() and not new_path.exists():
        try:
            os.rename(old_path, new_path)
            settings = load_settings_from_file()
            if settings.get("active_excel_file") == old_filename:
                settings["active_excel_file"] = new_filename
                save_settings_to_file(settings)
                session['settings'] = settings
            flash("Soubor přejmenován.", "success")
        except Exception as e:
            flash(f"Chyba při přejmenování: {e}", "error")
    else:
        flash("Starý soubor neexistuje nebo nový již existuje.", "error")
    return redirect(url_for('settings_page'))


@app.route("/delete_project", methods=["POST"])
def delete_project():
    filename = request.form.get("excel_file_to_delete")
    settings = load_settings_from_file()
    if filename == settings.get("active_excel_file"):
        flash("Nelze smazat aktivní soubor.", "error")
    elif filename and (Config.EXCEL_BASE_PATH / filename).exists():
        try:
            os.remove(Config.EXCEL_BASE_PATH / filename)
            flash(f"Soubor '{filename}' byl smazán.", "success")
        except Exception as e:
            flash(f"Chyba při mazání souboru: {e}", "error")
    else:
        flash("Soubor neexistuje.", "error")
    return redirect(url_for('settings_page'))


@app.route("/excel_viewer", methods=["GET"])
@require_excel_managers
def excel_viewer():
    active_filename = g.excel_manager.get_active_file_path().name
    active_sheet_name = request.args.get("sheet")
    data, sheet_names = [], []
    try:
        with load_workbook(g.excel_manager.get_active_file_path(), read_only=True, data_only=True) as wb:
            sheet_names = wb.sheetnames
            if not active_sheet_name or active_sheet_name not in sheet_names:
                active_sheet_name = sheet_names[0]
            sheet = wb[active_sheet_name]
            for i, row in enumerate(sheet.iter_rows(values_only=True)):
                if i >= Config.MAX_ROWS_TO_DISPLAY_EXCEL_VIEWER:
                    flash(f"Zobrazeno prvních {Config.MAX_ROWS_TO_DISPLAY_EXCEL_VIEWER} řádků.", "warning")
                    break
                data.append([str(c) if c is not None else "" for c in row])
    except (FileNotFoundError, InvalidFileException, Exception) as e:
        flash(f"Chyba při zobrazení souboru: {e}", "error")
        return redirect(url_for("index"))
    
    return render_template("excel_viewer.html", excel_files=[active_filename], selected_file=active_filename,
                           sheet_names=sheet_names, active_sheet=active_sheet_name, data=data)


@app.route("/settings", methods=["GET", "POST"])
@require_excel_managers
def settings_page():
    if request.method == "POST":
        try:
            settings_to_save = session.get('settings', Config.get_default_settings()).copy()
            form = request.form
            lunch_duration = float(form["lunch_duration"].replace(",", "."))
            if not (0 <= lunch_duration <= 4): raise ValueError("Neplatná délka pauzy.")
            
            settings_to_save.update({
                "start_time": form["start_time"], "end_time": form["end_time"],
                "lunch_duration": lunch_duration,
                "project_info": {
                    "name": form["project_name"].strip(),
                    "start_date": form["start_date"],
                    "end_date": form["end_date"],
                },
            })

            if not save_settings_to_file(settings_to_save):
                raise RuntimeError("Nepodařilo se uložit nastavení.")
            session['settings'] = settings_to_save

            if g.excel_manager.update_project_info(form["project_name"], form["start_date"], form["end_date"]):
                flash("Nastavení bylo úspěšně uloženo.", "success")
            else:
                flash("Nastavení uloženo, ale nepodařilo se aktualizovat Excel.", "warning")
            
            return redirect(url_for("settings_page"))

        except (ValueError, RuntimeError, Exception) as e:
            flash(str(e), "error")
    
    return render_template("settings_page.html", settings=session.get('settings', {}))


@app.route("/zalohy", methods=["GET", "POST"])
@require_excel_managers
def zalohy():
    employees_list = g.employee_manager.zamestnanci
    advance_options = g.zalohy_manager.get_option_names()
    
    if request.method == "POST":
        try:
            form = request.form
            amount = float(form["amount"].replace(",", "."))
            if not g.zalohy_manager.add_or_update_employee_advance(form["employee_name"], amount, form["currency"], form["option"], form["date"]):
                raise RuntimeError("Nepodařilo se uložit zálohu.")
            flash("Záloha byla úspěšně uložena.", "success")
            return redirect(url_for('zalohy'))
        except (ValueError, RuntimeError, Exception) as e:
            flash(str(e), "error")
    
    return render_template("zalohy.html", employees=employees_list, options=advance_options,
                           current_date=datetime.now().strftime("%Y-%m-%d"), advance_history=[])


@app.route("/start_new_file", methods=["POST"])
def start_new_file():
    settings = load_settings_from_file()
    current_active_file = settings.get("active_excel_file")
    project_end_str = settings.get("project_info", {}).get("end_date")

    if not current_active_file:
        flash("Není nastaven žádný aktivní soubor.", "info")
    elif not project_end_str:
        flash("Před archivací musí být zadáno datum konce projektu.", "error")
    else:
        settings["active_excel_file"] = None
        if save_settings_to_file(settings):
            session['settings'] = settings
            flash(f"Soubor '{current_active_file}' byl archivován.", "success")
        else:
            flash("Nepodařilo se archivovat soubor.", "error")
            
    return redirect(url_for('settings_page'))


@app.route('/voice-command', methods=['POST'])
def voice_command():
    try:
        data = request.get_json()
        if not data or 'command' not in data:
            return jsonify({'success': False, 'error': 'Chybí hlasový příkaz'})

        result = VoiceProcessor().process_voice_text(data['command'])
        if not result['success']: return jsonify(result)

        entities = result['entities']
        if entities['action'] == 'record_time':
            selected_employees = g.employee_manager.get_vybrani_zamestnanci()
            if not selected_employees:
                return jsonify({'success': False, 'error': 'Nejsou vybráni zaměstnanci'})
            
            if entities.get('is_free_day'):
                entities.update({'start_time': "00:00", 'end_time': "00:00", 'lunch_duration': 0.0})
            
            success, message = g.excel_manager.ulozit_pracovni_dobu(
                entities['date'], entities['start_time'], entities['end_time'],
                entities.get('lunch_duration', 1.0), selected_employees
            )
            result['operation_result'] = {'success': success, 'message': message}

        elif entities['action'] == 'get_stats':
            # Implementace statistik
            pass

        return jsonify(result)

    except Exception as e:
        logger.error(f"Chyba při zpracování hlasového příkazu: {e}", exc_info=True)
        return jsonify({'success': False, 'error': 'Interní chyba serveru'})


@app.route('/monthly_report', methods=['GET', 'POST'])
@require_excel_managers
def monthly_report_route():
    employee_names = [emp['name'] for emp in g.employee_manager.get_all_employees()]
    report_data = None
    selected_employees_post = []
    
    current_month = request.form.get('month', datetime.now().month, type=int)
    current_year = request.form.get('year', datetime.now().year, type=int)

    if request.method == 'POST':
        selected_employees_post = request.form.getlist('employees')
        try:
            if not (1 <= current_month <= 12 and 2000 <= current_year <= 2100):
                raise ValueError("Neplatný měsíc nebo rok.")
            
            report_data = g.excel_manager.generate_monthly_report(
                month=current_month, year=current_year,
                employees=selected_employees_post or None
            )
            if not report_data:
                flash('Nebyly nalezeny žádné záznamy.', 'info')
        except (ValueError, IOError, Exception) as e:
            flash(str(e), 'error')
            logger.error(f"Chyba při generování reportu: {e}", exc_info=True)

    return render_template("monthly_report.html", employee_names=employee_names, current_month=current_month,
                           current_year=current_year, report_data=report_data,
                           selected_employees_post=selected_employees_post)


def validate_email(email):
    return re.match(r'^[\w\.-]+@[\w\.-]+\.\w+$', email) is not None


if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=5000)
