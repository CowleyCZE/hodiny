from flask import Flask, jsonify, render_template, request, redirect, url_for, flash, send_file, session
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import os
import re
import json
from utils.logger import setup_logger
import logging
import shutil
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib
import ssl
from email.utils import parseaddr
from pathlib import Path

# Local imports
from config import Config
from employee_management import EmployeeManager
from excel_manager import ExcelManager
from zalohy_manager import ZalohyManager

# Nahrazení základního loggeru naším vlastním
logger = setup_logger('app')

# Initialize Flask app
app = Flask(__name__)
app.secret_key = Config.SECRET_KEY

# Constants - upravené na používání Path pro konzistenci
DATA_PATH = Path(Config.DATA_PATH)
EXCEL_BASE_PATH = Path(Config.EXCEL_BASE_PATH)
EXCEL_FILE_NAME = Config.EXCEL_FILE_NAME
EXCEL_FILE_NAME_2025 = Config.EXCEL_FILE_NAME_2025
SETTINGS_FILE_PATH = Path(Config.SETTINGS_FILE_PATH)
RECIPIENT_EMAIL = Config.RECIPIENT_EMAIL

# Ensure required directories exist
for path in [DATA_PATH, EXCEL_BASE_PATH]:
    path.mkdir(parents=True, exist_ok=True)

# Initialize managers and load settings
employee_manager = EmployeeManager(DATA_PATH)
excel_manager = ExcelManager(EXCEL_BASE_PATH, EXCEL_FILE_NAME)

def get_settings():
    """Get settings from session or load them if not present"""
    if 'settings' not in session:
        default_settings = Config.get_default_settings()
        try:
            if os.path.exists(SETTINGS_FILE_PATH):
                with open(SETTINGS_FILE_PATH, 'r', encoding='utf-8') as f:
                    saved_settings = json.load(f)
                    default_settings.update(saved_settings)
            session['settings'] = default_settings
        except Exception as e:
            logger.error(f"Error loading settings: {e}")
            session['settings'] = default_settings
    return session['settings']

@app.route('/')
def index():
    excel_exists = os.path.exists(excel_manager.file_path)
    current_date = datetime.now().strftime('%Y-%m-%d')
    week_number = excel_manager.ziskej_cislo_tydne(current_date)
    return render_template('index.html', excel_exists=excel_exists, week_number=week_number, current_date=current_date)

@app.route('/download')
def download_file():
    try:
        return send_file(excel_manager.file_path, as_attachment=True)
    except Exception as e:
        logger.error(f"Chyba při stahování souboru: {e}")
        flash('Chyba při stahování souboru.', 'error')
        return redirect(url_for('index'))

def validate_email(email):
    """Validace emailové adresy"""
    if not email or '@' not in parseaddr(email)[1]:
        raise ValueError("Neplatná emailová adresa")
    return True

@app.route('/send_email', methods=['POST'])
def send_email():
    try:
        msg = MIMEMultipart()
        msg['Subject'] = 'Hodiny_Cap.xlsx'
        
        # Validace emailových adres
        sender = Config.SMTP_USERNAME
        recipient = RECIPIENT_EMAIL
        
        if not all([validate_email(addr) for addr in [sender, recipient]]):
            raise ValueError("Neplatné emailové adresy")
            
        msg['From'] = sender
        msg['To'] = recipient

        # Nastavení SSL/TLS kontextu
        ssl_context = ssl.create_default_context()
        
        with open(excel_manager.file_path, 'rb') as f:
            attachment = MIMEApplication(f.read(), _subtype="xlsx")
            attachment.add_header('Content-Disposition', 'attachment', filename=EXCEL_FILE_NAME)
            msg.attach(attachment)

        with smtplib.SMTP_SSL(Config.SMTP_SERVER, Config.SMTP_PORT, context=ssl_context) as smtp:
            smtp.login(Config.SMTP_USERNAME, Config.SMTP_PASSWORD)
            smtp.send_message(msg)

        flash('Email byl odeslán.', 'success')
    except Exception as e:
        logger.error(f"Chyba při odesílání emailu: {e}")
        flash('Chyba při odesílání emailu.', 'error')
    return redirect(url_for('index'))

@app.route('/zamestnanci', methods=['GET', 'POST'])
def manage_employees():
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add':
            employee_name = request.form['name']
            if employee_manager.pridat_zamestnance(employee_name):
                flash(f'Zaměstnanec "{employee_name}" byl přidán.', 'success')
            else:
                flash(f'Zaměstnanec "{employee_name}" už existuje.', 'error')
        elif action == 'select':
            employee_name = request.form['employee_name']
            if employee_name in employee_manager.zamestnanci:
                if employee_name in employee_manager.vybrani_zamestnanci:
                    employee_manager.odebrat_vybraneho_zamestnance(employee_name)
                else:
                    employee_manager.pridat_vybraneho_zamestnance(employee_name)
                employee_manager.save_config()
        elif action == 'edit':
            old_name = request.form['old_name']
            new_name = request.form['new_name']
            try:
                idx = employee_manager.zamestnanci.index(old_name) + 1
                if employee_manager.upravit_zamestnance(idx, new_name):
                    flash(f'Zaměstnanec "{old_name}" byl upraven na "{new_name}".', 'success')
                else:
                    flash(f'Nepodařilo se upravit zaměstnance "{old_name}".', 'error')
            except ValueError:
                flash(f'Zaměstnanec "{old_name}" nebyl nalezen.', 'error')
        elif action == 'delete':
            employee_name = request.form['employee_name']
            try:
                idx = employee_manager.zamestnanci.index(employee_name) + 1
                if employee_manager.smazat_zamestnance(idx):
                    flash(f'Zaměstnanec "{employee_name}" byl smazán.', 'success')
                else:
                    flash(f'Nepodařilo se smazat zaměstnance "{employee_name}".', 'error')
            except ValueError:
                flash(f'Zaměstnanec "{employee_name}" nebyl nalezen.', 'error')
        else:
            flash('Neplatná akce.', 'error')

    # Převedení seznamů na formát očekávaný šablonou
    employees = [
        {'name': name, 'selected': name in employee_manager.vybrani_zamestnanci}
        for name in employee_manager.zamestnanci
    ]
    return render_template('employees.html', employees=employees)

@app.route('/zaznam', methods=['GET', 'POST'])
def record_time():
    selected_employees = employee_manager.vybrani_zamestnanci
    current_date = datetime.now().strftime('%Y-%m-%d')
    settings = get_settings()
    start_time = settings.get('start_time', '07:00')
    end_time = settings.get('end_time', '18:00')
    lunch_duration = settings.get('lunch_duration', 1)

    if request.method == 'POST':
        date = request.form['date']
        start_time = request.form['start_time']
        end_time = request.form['end_time']
        lunch_duration = float(request.form['lunch_duration'])

        # Uložení do Hodiny_Cap.xlsx
        excel_manager.ulozit_pracovni_dobu(date, start_time, end_time, lunch_duration, selected_employees)

        flash('Pracovní doba byla zaznamenána.', 'success')

    return render_template(
        'record_time.html',
        selected_employees=selected_employees,
        current_date=current_date,
        start_time=start_time,
        end_time=end_time,
        lunch_duration=lunch_duration
    )

@app.route('/excel_viewer', methods=['GET'])
def excel_viewer():
    excel_files = ['Hodiny_Cap.xlsx', 'Hodiny2025.xlsx']
    selected_file = request.args.get('file', excel_files[0])
    active_sheet = request.args.get('sheet', None)

    try:
        if selected_file == 'Hodiny_Cap.xlsx':
            workbook = load_workbook(excel_manager.file_path, read_only=True)
        elif selected_file == 'Hodiny2025.xlsx':
            workbook = load_workbook(os.path.join(EXCEL_BASE_PATH, EXCEL_FILE_NAME_2025), read_only=True)
        else:
            raise ValueError("Neplatný název souboru.")

        sheet_names = workbook.sheetnames
        if active_sheet is None:
            active_sheet = sheet_names[0]

        if active_sheet not in sheet_names:
            raise ValueError("Neplatný název listu.")

        sheet = workbook[active_sheet]
        data = [[cell.value for cell in row] for row in sheet.iter_rows()]

        return render_template(
            'excel_viewer.html',
            excel_files=excel_files,
            selected_file=selected_file,
            sheet_names=sheet_names,
            active_sheet=active_sheet,
            data=data
        )

    except InvalidFileException:
        flash('Soubor nebyl nalezen nebo je poškozen.', 'error')
        return redirect(url_for('index'))
    except Exception as e:
        logger.error(f"Chyba při zobrazení Excel souboru: {e}")
        flash('Chyba při zobrazení Excel souboru.', 'error')
        return redirect(url_for('index'))

@app.route('/settings', methods=['GET', 'POST'])
def settings_page():
    """Handle settings page"""
    if request.method == 'POST':
        try:
            new_settings = get_settings()  # Get current settings as base
            new_settings.update({
                'start_time': request.form['start_time'],
                'end_time': request.form['end_time'],
                'lunch_duration': float(request.form['lunch_duration']),
                'project_info': {
                    'name': request.form.get('project_name', ''),
                    'start_date': request.form.get('start_date', ''),
                    'end_date': request.form.get('end_date', '')
                }
            })

            # Ensure settings directory exists
            os.makedirs(os.path.dirname(SETTINGS_FILE_PATH), exist_ok=True)
            
            # Save new settings
            with open(SETTINGS_FILE_PATH, 'w', encoding='utf-8') as f:
                json.dump(new_settings, f, indent=4)
            
            session['settings'] = new_settings
            excel_manager.update_project_info(
                new_settings['project_info']['name'],
                new_settings['project_info']['start_date'],
                new_settings['project_info']['end_date']
            )
            flash('Nastavení bylo úspěšně uloženo.', 'success')
            
        except Exception as e:
            logger.error(f"Error saving settings: {e}")
            flash('Chyba při ukládání nastavení.', 'error')
            
    return render_template('settings_page.html', settings=get_settings())

@app.route('/zalohy', methods=['GET', 'POST'])
def zalohy():
    if request.method == 'POST':
        try:
            employee_name = request.form['employee_name']
            amount = float(request.form['amount'])
            currency = request.form['currency']
            option = request.form['option']
            date = request.form['date']

            # Vytvoření instance ZalohyManager
            zalohy_manager = ZalohyManager(EXCEL_BASE_PATH)

            # Validace a uložení zálohy
            zalohy_manager.add_or_update_employee_advance(employee_name, amount, currency, option, date)
            flash('Záloha byla úspěšně uložena.', 'success')

        except ValueError as e:
            flash(f'Chyba validace: {str(e)}', 'error')
            logger.error(f"Chyba validace zálohy: {e}")
        except Exception as e:
            flash(f'Chyba při ukládání zálohy: {str(e)}', 'error')
            logger.error(f"Chyba při ukládání zálohy: {e}")

    employees = employee_manager.zamestnanci
    options = excel_manager.get_advance_options()
    current_date = datetime.now().strftime('%Y-%m-%d')

    # Načtení historie záloh z Hodiny2025.xlsx
    try:
        hodiny2025_path = os.path.join(EXCEL_BASE_PATH, EXCEL_FILE_NAME_2025)
        if not os.path.exists(hodiny2025_path):
            logger.error(f"Soubor {EXCEL_FILE_NAME_2025} nebyl nalezen")
            flash(f'Soubor {EXCEL_FILE_NAME_2025} nebyl nalezen', 'error')
            advance_history = []
        else:
            workbook_2025 = load_workbook(hodiny2025_path)
            if 'Zalohy25' in workbook_2025.sheetnames:
                sheet = workbook_2025['Zalohy25']
                data = list(sheet.values)
                keys = data[0]  # Předpokládáme, že první řádek obsahuje hlavičky
                advance_history = [dict(zip(keys, row)) for row in data[1:]]  # Vytvoření seznamu slovníků
            else:
                advance_history = []
    except Exception as e:
        logger.error(f"Chyba při načítání historie záloh: {str(e)}")
        advance_history = []

    return render_template(
        'zalohy.html',
        employees=employees,
        options=options,
        current_date=current_date,
        advance_history=advance_history
    )

if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO, filename='app.log', filemode='a')
    app.run(debug=True)