# app.py
import json
import logging
import os
import re
import shutil
import smtplib
import ssl
from datetime import datetime, timedelta
from functools import wraps
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import parseaddr

from flask import Flask, flash, jsonify, redirect, render_template, request, send_file, session, url_for, g
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook import Workbook

from config import Config
from employee_management import EmployeeManager
from excel_manager import ExcelManager
from utils.logger import setup_logger
from zalohy_manager import ZalohyManager
from utils.voice_processor import VoiceProcessor

# Inicializace aplikace
app = Flask(__name__)
app.secret_key = Config.SECRET_KEY
Config.init_app(app)

# Nastavení loggeru
logger = setup_logger("app")

# Globální proměnné
employee_manager = None
excel_manager = None
zalohy_manager = None

# Pomocné funkce
def require_excel_managers(f):
    """Dekorátor pro routy, které vyžadují inicializované manažery"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not hasattr(g, 'excel_manager') or not g.excel_manager:
            flash("Správce Excel souboru není k dispozici.", "error")
            return redirect(url_for('index'))
        if not hasattr(g, 'employee_manager') or not g.employee_manager:
            flash("Správce zaměstnanců není k dispozici.", "error")
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function

def validate_email(email):
    """Validace emailové adresy"""
    if not email or '@' not in email:
        return False
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return re.match(pattern, email) is not None

@app.before_request
def before_request():
    """Inicializace manažerů před každým requestem"""
    global employee_manager, excel_manager, zalohy_manager
    
    try:
        # Inicializace manažerů pouze jednou na request
        if not hasattr(g, 'employee_manager'):
            g.employee_manager = EmployeeManager(Config.DATA_PATH)
        
        if not hasattr(g, 'excel_manager'):
            g.excel_manager = ExcelManager(Config.EXCEL_BASE_PATH, Config.EXCEL_TEMPLATE_NAME)
        
        if not hasattr(g, 'zalohy_manager'):
            g.zalohy_manager = ZalohyManager(Config.EXCEL_BASE_PATH, Config.EXCEL_TEMPLATE_NAME)
            
    except Exception as e:
        logger.error(f"Neočekávaná chyba při inicializaci manažerů: {e}", exc_info=True)
        g.excel_manager = None
        g.zalohy_manager = None
        flash("Neočekávaná chyba při přípravě aplikace.", "error")

# Routy
@app.route('/')
def index():
    """Hlavní stránka"""
    try:
        settings = session.get('settings', {})
        active_filename = settings.get('active_excel_file')
        
        if not active_filename:
            flash("Chyba: Není definován aktivní Excel soubor pro práci.", "error")
            return redirect(url_for('settings_page'))
            
        excel_exists = os.path.exists(os.path.join(Config.EXCEL_BASE_PATH, active_filename))
        
        week_num_int = 0  # Implementace pro získání čísla týdne
        current_date = datetime.now().strftime("%Y-%m-%d")
        
        return render_template('index.html', 
                              week_number=week_num_int,
                              current_date=current_date,
                              current_year=datetime.now().year)
                              
    except Exception as e:
        logger.error(f"Chyba na hlavní stránce: {e}", exc_info=True)
        flash("Došlo k chybě při načítání hlavní stránky.", "error")
        return redirect(url_for('settings_page'))

@app.route('/zamestnanci', methods=['GET', 'POST'])
@require_excel_managers
def manage_employees():
    """Správa zaměstnanců"""
    employee_manager_instance = g.employee_manager
    if not employee_manager_instance:
        flash("Správce zaměstnanců není k dispozici.", "error")
        return redirect(url_for('index'))
    
    if request.method == 'POST':
        action = request.form.get('action')
        try:
            if not action:
                raise ValueError("Nebyla specifikována akce")
                
            if action == 'add':
                employee_name = request.form.get('name', '').strip()
                if not employee_name:
                    raise ValueError("Jméno zaměstnance nemůže být prázdné")
                if len(employee_name) > 100:
                    raise ValueError("Jméno zaměstnance je příliš dlouhé")
                if not re.match(r'^[\w\s\-\.ěščřžýáíéúůďťňĚŠČŘŽÝÁÍÉÚŮĎŤŇ]+$', employee_name):
                    raise ValueError("Jméno zaměstnance obsahuje nepovolené znaky.")
                    
                if employee_manager_instance.pridat_zamestnance(employee_name):
                    flash(f"Zaměstnanec {employee_name} byl úspěšně přidán", "success")
                else:
                    flash("Nepodařilo se přidat zaměstnance", "error")
                    
            elif action == 'toggle':
                employee_name = request.form.get('name')
                if employee_name:
                    if employee_manager_instance.prepinat_zamestnance(employee_name):
                        status = "vybrán" if employee_name in employee_manager_instance.get_vybrani_zamestnanci() else "odebrán"
                        flash(f"Zaměstnanec {employee_name} byl {status}", "success")
                    else:
                        flash("Nepodařilo se upravit zaměstnance", "error")
                        
            elif action == 'edit':
                old_name = request.form.get('old_name')
                new_name = request.form.get('new_name', '').strip()
                if not old_name or not new_name:
                    raise ValueError("Musíte zadat nové jméno zaměstnance")
                if old_name == new_name:
                    flash("Nové jméno je stejné jako původní", "info")
                elif employee_manager_instance.upravit_zamestnance(old_name, new_name):
                    flash(f"Zaměstnanec {old_name} byl přejmenován na {new_name}", "success")
                else:
                    flash("Nepodařilo se upravit zaměstnance", "error")
                    
            elif action == 'delete':
                index = request.form.get('index')
                if index:
                    if employee_manager_instance.smazat_zamestnance(int(index)):
                        flash("Zaměstnanec byl úspěšně smazán", "success")
                    else:
                        flash("Nepodařilo se smazat zaměstnance", "error")
                        
        except ValueError as e:
            flash(str(e), "error")
        except Exception as e:
            logger.error(f"Neočekávaná chyba při správě zaměstnanců: {e}", exc_info=True)
            flash("Došlo k neočekávané chybě", "error")
    
    employees = employee_manager_instance.get_all_employees()
    selected_employees = employee_manager_instance.get_vybrani_zamestnanci()
    return render_template('employees.html', 
                          employees=employees, 
                          selected_employees=selected_employees)

@app.route('/process-audio', methods=['POST'])
@require_excel_managers
def process_audio():
    """Zpracování hlasového souboru přes Gemini API"""
    try:
        if 'audio' not in request.files:
            return jsonify({'success': False, 'error': 'Chybí audio soubor'})
            
        # Uložení dočasného souboru
        audio_file = request.files['audio']
        temp_dir = Path("temp_audio")
        temp_dir.mkdir(exist_ok=True)
        temp_path = temp_dir / f"{datetime.now().timestamp()}.webm"
        
        audio_file.save(temp_path)
        logger.info(f"Dočasný soubor uložen: {temp_path}")
        
        # Zpracování hlasu přes Gemini
        voice_processor = VoiceProcessor()
        result = voice_processor.process_voice_audio(temp_path)
        
        # Odstranění dočasného souboru
        os.remove(temp_path)
        
        if not result['success']:
            return jsonify(result)
            
        # Zpracování extrahovaných dat
        entities = result['entities']
        excel_manager = g.excel_manager
        employee_manager = g.employee_manager
        
        # Podle typu akce vykonáme odpovídající operaci
        if entities['action'] == 'record_time':
            success, message = excel_manager.record_time(
                entities['employee'], 
                entities['date'], 
                entities['start_time'], 
                entities['end_time']
            )
            result['operation_result'] = {'success': success, 'message': message}
            
        elif entities['action'] == 'add_advance':
            success, message = zalohy_manager.add_or_update_employee_advance(
                entities['employee'], 
                entities['amount'], 
                entities['currency'], 
                entities['option'], 
                entities['date']
            )
            result['operation_result'] = {'success': success, 'message': message}
            
        elif entities['action'] == 'get_stats':
            stats = {}
            if 'time_period' in entities:
                if entities['time_period'] == 'week':
                    stats = excel_manager.get_week_stats(entities.get('employee'))
                elif entities['time_period'] == 'month':
                    stats = excel_manager.get_month_stats(entities.get('employee'))
                elif entities['time_period'] == 'year':
                    stats = excel_manager.get_year_stats(entities.get('employee'))
            else:
                stats = excel_manager.get_total_stats(entities.get('employee'))
                
            result['stats'] = stats
            
        return jsonify(result)
        
    except Exception as e:
        logger.error(f"Chyba při zpracování hlasového souboru: {e}", exc_info=True)
        return jsonify({'success': False, 'error': 'Interní chyba serveru'})

@app.route('/zaznam', methods=['GET', 'POST'])
@require_excel_managers
def record_time():
    """Záznam pracovní doby"""
    employee_manager_instance = g.employee_manager
    excel_manager_instance = g.excel_manager
    
    if not employee_manager_instance or not excel_manager_instance:
        flash("Není možné zaznamenat pracovní dobu bez inicializovaných manažerů", "error")
        return redirect(url_for('index'))
    
    selected_employees = employee_manager_instance.get_vybrani_zamestnanci()
    if not selected_employees:
        flash("Nejsou vybráni žádní zaměstnanci pro záznam.", "warning")
        return redirect(url_for('manage_employees'))
    
    settings = session.get('settings', {})
    default_start_time = settings.get("start_time", "07:00")
    default_end_time = settings.get("end_time", "18:00")
    default_lunch_duration = settings.get("lunch_duration", "1.0")
    
    if request.method == 'POST':
        try:
            date = request.form.get('date')
            start_time = request.form.get('start_time')
            end_time = request.form.get('end_time')
            lunch_duration = request.form.get('lunch_duration', '1.0')
            
            # Záznam pro všechny vybrané zaměstnance
            success_count = 0
            error_messages = []
            
            for employee_name in selected_employees:
                success, message = excel_manager_instance.record_time(
                    employee_name, date, start_time, end_time, lunch_duration
                )
                if success:
                    success_count += 1
                else:
                    error_messages.append(f"{employee_name}: {message}")
            
            if success_count > 0:
                flash(f"Záznam byl úspěšně uložen pro {success_count} zaměstnanců", "success")
            
            if error_messages:
                for msg in error_messages:
                    flash(msg, "error")
                    
        except Exception as e:
            logger.error(f"Chyba při ukládání záznamu: {e}", exc_info=True)
            flash("Došlo k chybě při ukládání záznamu", "error")
    
    return render_template('record_time.html',
                          selected_employees=selected_employees,
                          default_start_time=default_start_time,
                          default_end_time=default_end_time,
                          default_lunch_duration=default_lunch_duration)

@app.route('/voice-command', methods=['POST'])
@require_excel_managers
def voice_command():
    """Zpracování hlasového příkazu"""
    try:
        data = request.get_json()
        if not data or 'command' not in data:
            return jsonify({'success': False, 'error': 'Chybí hlasový příkaz'})
        
        voice_processor = VoiceProcessor()
        result = voice_processor.process_voice_command(data['command'])
        
        if not result['success']:
            return jsonify(result)
            
        # Podle typu akce vykonáme odpovídající operaci
        entities = result['entities']
        excel_manager = g.excel_manager
        employee_manager = g.employee_manager
        
        if entities['action'] == 'record_time':
            # Záznam pracovní doby
            success, message = excel_manager.record_time(
                entities['employee'], 
                entities['date'], 
                entities['start_time'], 
                entities['end_time']
            )
            result['operation_result'] = {'success': success, 'message': message}
            
        elif entities['action'] == 'add_advance':
            # Přidání zálohy
            success, message = zalohy_manager.add_or_update_employee_advance(
                entities['employee'], 
                entities['amount'], 
                entities['currency'], 
                entities['option'], 
                entities['date']
            )
            result['operation_result'] = {'success': success, 'message': message}
            
        elif entities['action'] == 'get_stats':
            # Získání statistik
            stats = {}
            if 'time_period' in entities:
                if entities['time_period'] == 'week':
                    stats = excel_manager.get_week_stats(entities.get('employee'))
                elif entities['time_period'] == 'month':
                    stats = excel_manager.get_month_stats(entities.get('employee'))
                elif entities['time_period'] == 'year':
                    stats = excel_manager.get_year_stats(entities.get('employee'))
            else:
                stats = excel_manager.get_total_stats(entities.get('employee'))
                
            result['stats'] = stats
            
        return jsonify(result)
        
    except Exception as e:
        logger.error(f"Chyba při zpracování hlasového příkazu: {e}", exc_info=True)
        return jsonify({'success': False, 'error': 'Interní chyba serveru'})

@app.route('/statistiky')
@require_excel_managers
def statistics():
    """Zobrazení statistik"""
    try:
        employee_manager = g.employee_manager
        excel_manager = g.excel_manager
        
        employee = request.args.get('employee')
        period = request.args.get('period', 'week')
        
        if employee:
            if period == 'week':
                stats = excel_manager.get_week_stats(employee)
            elif period == 'month':
                stats = excel_manager.get_month_stats(employee)
            elif period == 'year':
                stats = excel_manager.get_year_stats(employee)
            else:
                stats = excel_manager.get_total_stats(employee)
        else:
            stats = {
                'total_hours': excel_manager.get_total_hours(),
                'employees': {emp: excel_manager.get_total_stats(emp) for emp in employee_manager.get_all_employees()}
            }
            
        return render_template('statistics.html', 
                              stats=stats, 
                              employee=employee, 
                              period=period,
                              employees=employee_manager.get_all_employees())
                              
    except Exception as e:
        logger.error(f"Chyba při načítání statistik: {e}", exc_info=True)
        flash("Došlo k chybě při načítání statistik", "error")
        return redirect(url_for('index'))

@app.route('/nastaveni', methods=['GET', 'POST'])
def settings_page():
    """Nastavení aplikace"""
    if request.method == 'POST':
        try:
            # Zpracování nastavení
            start_time = request.form.get('start_time')
            end_time = request.form.get('end_time')
            lunch_duration = request.form.get('lunch_duration')
            project_name = request.form.get('project_name')
            start_date = request.form.get('start_date')
            end_date = request.form.get('end_date')
            active_excel_file = request.form.get('active_excel_file')
            
            # Aktualizace nastavení ve session
            session['settings'] = {
                'start_time': start_time,
                'end_time': end_time,
                'lunch_duration': lunch_duration,
                'project_info': {
                    'name': project_name,
                    'start_date': start_date,
                    'end_date': end_date
                },
                'active_excel_file': active_excel_file
            }
            
            flash("Nastavení bylo úspěšně uloženo", "success")
            return redirect(url_for('index'))
            
        except Exception as e:
            logger.error(f"Chyba při ukládání nastavení: {e}", exc_info=True)
            flash("Došlo k chybě při ukládání nastavení", "error")
    
    # Načtení aktuálních nastavení
    settings = session.get('settings', {})
    return render_template('settings.html', settings=settings)

@app.route('/stahnout')
@require_excel_managers
def download_excel():
    """Stažení aktuálního Excel souboru"""
    try:
        active_file_path = g.excel_manager.get_active_file_path()
        download_filename = os.path.basename(active_file_path)
        
        return send_file(
            str(active_file_path),
            as_attachment=True,
            download_name=download_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        logger.error(f"Chyba při stahování souboru: {e}", exc_info=True)
        flash("Došlo k chybě při stahování souboru", "error")
        return redirect(url_for('index'))

@app.route('/odeslat-email', methods=['POST'])
@require_excel_managers
def send_email():
    """Odeslání Excel souboru emailem"""
    try:
        # Získání informací o souboru
        active_file_path = g.excel_manager.get_active_file_path()
        active_filename = os.path.basename(active_file_path)
        
        # Získání parametrů z formuláře
        recipient = request.form.get('recipient')
        subject = request.form.get('subject', f'Výkaz pracovní doby - {active_filename}')
        message = request.form.get('message', 'Dobrý den,\n\nv příloze zasílám aktuální výkaz pracovní doby.\n\nS pozdravem')
        
        # Validace emailu
        if not recipient or '@' not in recipient:
            flash("Zadejte platnou emailovou adresu příjemce", "error")
            return redirect(url_for('index'))
        
        # Vytvoření emailu
        sender = Config.SMTP_USERNAME
        msg = MIMEMultipart()
        
        # Přidání obsahu
        msg['Subject'] = subject
        msg['From'] = sender
        msg['To'] = recipient
        msg.attach(MIMEText(message, 'plain', 'utf-8'))
        
        # Přidání přílohy
        with open(active_file_path, 'rb') as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(active_file_path))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(active_file_path)}"'
            msg.attach(part)
        
        # Odeslání emailu
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as server:
            server.login(sender, Config.SMTP_PASSWORD)
            server.sendmail(sender, recipient, msg.as_string())
        
        flash("Email byl úspěšně odeslán", "success")
        return redirect(url_for('index'))
        
    except smtplib.SMTPAuthenticationError:
        flash("Chyba při autentifikaci na SMTP server", "error")
    except Exception as e:
        logger.error(f"Chyba při odesílání emailu: {e}", exc_info=True)
        flash("Došlo k chybě při odesílání emailu", "error")
    
    return redirect(url_for('index'))

@app.route('/zalohy', methods=['GET', 'POST'])
@require_excel_managers
def advances():
    """Správa záloh"""
    zalohy_manager_instance = g.zalohy_manager
    if not zalohy_manager_instance:
        flash("Správce záloh není k dispozici.", "error")
        return redirect(url_for('index'))
    
    if request.method == 'POST':
        try:
            employee = request.form.get('employee')
            amount = request.form.get('amount')
            currency = request.form.get('currency')
            option = request.form.get('option')
            date = request.form.get('date')
            
            if not all([employee, amount, currency, option, date]):
                raise ValueError("Musíte vyplnit všechna pole")
                
            success, message = zalohy_manager_instance.add_or_update_employee_advance(
                employee, amount, currency, option, date
            )
            
            if success:
                flash(f"Záloha pro {employee} byla úspěšně uložena", "success")
            else:
                flash(message, "error")
                
        except ValueError as e:
            flash(str(e), "error")
        except Exception as e:
            logger.error(f"Chyba při ukládání zálohy: {e}", exc_info=True)
            flash("Došlo k chybě při ukládání zálohy", "error")
    
    employees = g.employee_manager.get_all_employees()
    option_names = zalohy_manager_instance.get_option_names()
    return render_template('advances.html',
                          employees=employees,
                          option_names=option_names)

@app.route('/start_new_file', methods=['POST'])
def start_new_file():
    """Začne nový soubor"""
    try:
        settings = session.get('settings', {})
        project_info = settings.get('project_info', {})
        project_name = project_info.get('name')
        project_start = project_info.get('start_date')
        project_end = project_info.get('end_date')
        
        if not all([project_name, project_start, project_end]):
            raise ValueError("Musíte zadat všechny informace o projektu")
            
        # Archivace starého souboru
        old_file = settings.get('active_excel_file')
        if old_file:
            archive_path = os.path.join(Config.EXCEL_ARCHIVE_PATH, old_file)
            shutil.move(os.path.join(Config.EXCEL_BASE_PATH, old_file), archive_path)
        
        # Vytvoření nového souboru
        new_file = f"{project_name}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        shutil.copy(
            os.path.join(Config.EXCEL_TEMPLATE_PATH, Config.EXCEL_TEMPLATE_NAME),
            os.path.join(Config.EXCEL_BASE_PATH, new_file)
        )
        
        # Aktualizace nastavení
        settings['active_excel_file'] = new_file
        session['settings'] = settings
        
        flash(f"Nový soubor {new_file} byl vytvořen", "success")
        return redirect(url_for('index'))
        
    except ValueError as e:
        flash(str(e), "error")
    except Exception as e:
        logger.error(f"Chyba při vytváření nového souboru: {e}", exc_info=True)
        flash("Došlo k chybě při vytváření nového souboru", "error")
    
    return redirect(url_for('settings_page'))

# Spuštění aplikace
if __name__ == '__main__':
    app.run(debug=True)
