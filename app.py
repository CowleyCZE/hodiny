import pandas as pd
from flask import Flask, jsonify, render_template, request, redirect, url_for, flash, send_file
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import os
import re
import json
import logging
import shutil
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib
import openpyxl

from employee_management import EmployeeManager
from excel_manager import ExcelManager
from excel_manager2024 import ExcelManager2024

app = Flask(__name__)
app.secret_key = 'your_secret_key'

# Konstanty
DATA_PATH = '/home/Cowley/hodiny/data'
EXCEL_BASE_PATH = '/home/Cowley/hodiny/excel'
EXCEL_FILE_NAME = 'Hodiny_Cap.xlsx'  # Původní název souboru
EXCEL_FILE_NAME_2024 = 'Hodiny2024.xlsx'  # Nový název souboru pro rok 2024
SETTINGS_FILE_PATH = 'settings.json'
RECIPIENT_EMAIL = 'cowleyy@gmail.com'

# Inicializace manažerů
employee_manager = EmployeeManager(DATA_PATH)
excel_manager = ExcelManager(EXCEL_BASE_PATH, EXCEL_FILE_NAME)  # Předání názvu souboru
excel_manager2024 = ExcelManager2024(os.path.join(EXCEL_BASE_PATH, EXCEL_FILE_NAME_2024))
# Načtení nastavení
def load_settings():
    if os.path.exists(SETTINGS_FILE_PATH):
        with open(SETTINGS_FILE_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    else:
        return {
            'start_time': '07:00',
            'end_time': '18:00',
            'lunch_duration': 1
        }

# Uložení nastavení
def save_settings(settings):
    with open(SETTINGS_FILE_PATH, 'w', encoding='utf-8') as f:
        json.dump(settings, f, indent=4, ensure_ascii=False)

# Globální proměnná pro nastavení
settings = load_settings()

def get_last_sheet_name(file_path):
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        week_sheets = [sheet for sheet in wb.sheetnames if sheet.startswith('Week')]

        if not week_sheets:
            wb.close()
            return None

        last_sheet = week_sheets[-1]
        wb.close()
        return last_sheet
    except Exception as e:
        logging.error(f"Chyba při čtení Excel souboru: {str(e)}")
        return None

def get_excel_with_week(base_path, original_name):
    original_path = os.path.join(base_path, original_name)
    if not os.path.exists(original_path):
        raise FileNotFoundError(f"Soubor {original_name} nebyl nalezen v {base_path}")

    last_sheet = get_last_sheet_name(original_path)
    if last_sheet is None:
        raise ValueError("Nelze určit název posledního listu z Excel souboru")

    new_name = f"Hodiny_Cap_{last_sheet}.xlsx"
    new_path = os.path.join(base_path, new_name)

    shutil.copy2(original_path, new_path)
    return new_path, new_name

@app.route('/')
def index():
    logging.info("Přístup na hlavní stránku")
    excel_exists = os.path.exists(os.path.join(EXCEL_BASE_PATH, EXCEL_FILE_NAME))
    return render_template('index.html', excel_exists=excel_exists)

@app.route('/download')
def download_file():
    try:
        # Cesta k původnímu souboru
        original_file_path = os.path.join(EXCEL_BASE_PATH, 'Hodiny_Cap.xlsx')

        logging.info(f"Načítám Excel soubor: {original_file_path}")
        # Načtení Excel souboru
        workbook = load_workbook(original_file_path)

        # Nalezení listu s nejvyšším číslem týdne
        week_sheets = [sheet for sheet in workbook.sheetnames if sheet.startswith('Týden')]
        logging.info(f"Nalezené listy týdnů: {week_sheets}")

        if not week_sheets:
            raise ValueError("V souboru nejsou žádné listy začínající 'Týden'")

        def safe_week_number(sheet_name):
            try:
                # Nalezení čísla v názvu listu pomocí regulárního výrazu
                match = re.search(r'\d+', sheet_name)
                if match:
                    return int(match.group())
                else:
                    return 0  # Nebo jiná výchozí hodnota, pokud číslo není nalezeno
            except ValueError:
                return -1  # Vrátí -1 pro neplatné názvy

        highest_week_sheet = max(week_sheets, key=safe_week_number)
        logging.info(f"Nejvyšší týden: {highest_week_sheet}")

        # Vytvoření nového názvu souboru
        new_file_name = f"Hodiny_Cap_{highest_week_sheet}.xlsx"
        new_file_path = os.path.join(EXCEL_BASE_PATH, new_file_name)

        logging.info(f"Vytvářím kopii souboru: {new_file_path}")
        # Uložení kopie souboru s novým názvem
        original_path = os.path.join(EXCEL_BASE_PATH, EXCEL_FILE_NAME)  # Definování original_path
        shutil.copy2(original_path, new_file_path)
        # Odeslání souboru uživateli ke stažení
        logging.info(f"Odesílám soubor ke stažení: {new_file_name}")
        return send_file(new_file_path, as_attachment=True, download_name=new_file_name)

    except Exception as e:
        logging.error(f"Chyba při stahování souboru: {str(e)}", exc_info=True)
        flash(f'Chyba při stahování souboru: {str(e)}', 'error')

    return redirect(url_for('index'))

@app.route('/send_email', methods=['POST'])
def send_email():
    try:
        file_name = excel_manager.get_file_name_with_week()
        weekly_copy_path = excel_manager.create_weekly_copy()

        if not weekly_copy_path:
            raise ValueError("Nepodařilo se vytvořit týdenní kopii souboru.")
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        sender_email = "cowleyy@gmail.com"
        password = os.getenv("EMAIL_PASSWORD")

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = RECIPIENT_EMAIL
        msg['Subject'] = f'{file_name} - Export'

        body = "V příloze najdete aktuální výkaz hodin."
        msg.attach(MIMEText(body, 'plain'))

        with app.open_resource(excel_manager.file_path) as fp:
            msg.attach(file_name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fp.read())

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, password)
            server.send_message(msg)

        return jsonify({"message": "E-mail byl úspěšně odeslán"}), 200
    except Exception as e:
        logging.error(f"Chyba při odesílání e-mailu: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/manage_employees', methods=['GET', 'POST'])
def manage_employees():
    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'add':
            name = request.form.get('name')
            if name:
                if employee_manager.pridat_zamestnance(name):
                    flash('Zaměstnanec byl úspěšně přidán.', 'success')
                else:
                    flash('Zaměstnanec již existuje.', 'error')

        elif action == 'select':
            employee_name = request.form.get('employee_name')
            for employee in employee_manager.zamestnanci:
                if employee == employee_name:
                    if employee in employee_manager.vybrani_zamestnanci:
                        employee_manager.odebrat_vybraneho_zamestnance(employee)
                        flash('Zaměstnanec byl odznačen.', 'success')
                    else:
                        employee_manager.pridat_vybraneho_zamestnance(employee)
                        flash('Zaměstnanec byl označen.', 'success')

        elif action == 'edit':
            old_name = request.form.get('old_name')
            new_name = request.form.get('new_name')
            if old_name and new_name:
                # Najít index starého jména
                try:
                    index = employee_manager.zamestnanci.index(old_name) + 1
                    if employee_manager.upravit_zamestnance(index, new_name):
                        # Aktualizovat také ve vybraných zaměstnancích
                        if old_name in employee_manager.vybrani_zamestnanci:
                            employee_manager.vybrani_zamestnanci.remove(old_name)
                            employee_manager.vybrani_zamestnanci.append(new_name)
                        flash('Jméno zaměstnance bylo úspěšně upraveno.', 'success')
                    else:
                        flash('Nepodařilo se upravit jméno zaměstnance.', 'error')
                except ValueError:
                    flash('Zaměstnanec nebyl nalezen.', 'error')

        elif action == 'delete':
            employee_name = request.form.get('employee_name')
            if employee_name:
                try:
                    index = employee_manager.zamestnanci.index(employee_name) + 1
                    if employee_manager.smazat_zamestnance(index):
                        flash('Zaměstnanec byl úspěšně smazán.', 'success')
                    else:
                        flash('Nepodařilo se smazat zaměstnance.', 'error')
                except ValueError:
                    flash('Zaměstnanec nebyl nalezen.', 'error')

        return redirect(url_for('manage_employees'))

    # Příprava dat pro šablonu
    employees_data = [{'name': employee, 'selected': employee in employee_manager.vybrani_zamestnanci}
                      for employee in sorted(employee_manager.zamestnanci)]

    return render_template('employees.html', employees=employees_data)



@app.route('/record_time', methods=['GET', 'POST'])
def record_time():
    settings = load_settings()
    if request.method == 'POST':
        date = request.form['date']
        start_time = request.form['start_time']
        end_time = request.form['end_time']
        lunch_duration = float(request.form['lunch_duration'])
        employees = employee_manager.get_vybrani_zamestnanci()
        try:
            excel_manager.ulozit_pracovni_dobu(date, start_time, end_time, lunch_duration, employees)
            flash('Pracovní doba byla úspěšně zaznamenána.', 'success')
        except Exception as e:
            flash(f'Chyba při ukládání pracovní doby: {str(e)}', 'error')
        return redirect(url_for('record_time'))

    return render_template('record_time.html',
                            current_date=datetime.now().strftime('%Y-%m-%d'),
                            start_time=settings['start_time'],
                            end_time=settings['end_time'],
                            lunch_duration=settings['lunch_duration'])

@app.route('/save_time', methods=['POST'])
def save_time():
    try:
        date = request.form['date']
        start_time = request.form['start_time']
        end_time = request.form['end_time']
        lunch_duration = float(request.form['lunch_duration'])
        employees = employee_manager.get_vybrani_zamestnanci()

        # Uložení do původního souboru
        excel_manager.ulozit_pracovni_dobu(date, start_time, end_time, lunch_duration, employees)

        # Uložení do nového souboru 2024
        excel_manager2024.ulozit_pracovni_dobu(date, start_time, end_time, lunch_duration)

        flash('Data byla úspěšně uložena do obou souborů.', 'success')
        return jsonify({"status": "success", "message": "Data byla úspěšně uložena."})
    except Exception as e:
        logging.error(f"Chyba při ukládání dat: {str(e)}", exc_info=True)
        flash(f'Chyba při ukládání dat: {str(e)}', 'error')
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/excel_viewer', methods=['GET'])
def excel_viewer():
    try:
        excel_dir = "/home/Cowley/hodiny/excel/"
        excel_files = [f for f in os.listdir(excel_dir) if f.endswith('.xlsx')]

        # Upřednostnění "Hodiny_Cap.xlsx"
        if 'Hodiny_Cap.xlsx' in excel_files:
            selected_file = 'Hodiny_Cap.xlsx'
        else:
            selected_file = request.args.get('file')
            if not selected_file or selected_file not in excel_files:
                selected_file = excel_files[0] if excel_files else None

        if not selected_file:
            return "Žádné Excel soubory nebyly nalezeny."

        file_path = os.path.join(excel_dir, selected_file)
        workbook = load_workbook(file_path, data_only=True)
        sheet_names = workbook.sheetnames

        active_sheet = request.args.get('sheet', sheet_names[0])

        if active_sheet not in sheet_names:
            raise ValueError("Neplatný název listu")

        sheet = workbook[active_sheet]
        data = []
        for row in sheet.iter_rows():
            row_data = []
            for cell in row:
                value = cell.value
                if value is None:
                    value = ""
                elif isinstance(value, str) and value.startswith("="):
                    try:
                        value = sheet.evaluate(value)
                    except Exception as e:
                        print(f"Chyba při vyhodnocení vzorce v buňce {cell.coordinate}: {e}")
                        value = "Chyba vzorce"
                row_data.append(value)
            data.append(row_data)

        workbook.close()
        return render_template('excel_viewer.html',
                               excel_files=excel_files,
                               selected_file=selected_file,
                               sheet_names=sheet_names,
                               active_sheet=active_sheet,
                               data=data)
    except Exception as e:
        return f"Chyba při načítání Excel souboru: {e}"

@app.route('/settings', methods=['GET', 'POST'])
def settings_page():
    global settings
    if request.method == 'POST':
        settings['start_time'] = request.form['start_time']
        settings['end_time'] = request.form['end_time']
        settings['lunch_duration'] = float(request.form['lunch_duration'])

        project_name = request.form.get('project_name')
        start_date = request.form.get('start_date')
        end_date = request.form.get('end_date')

        if project_name and start_date and end_date:
            settings['project_name'] = project_name
            settings['start_date'] = start_date
            settings['end_date'] = end_date

        save_settings(settings)
        flash('Nastavení bylo úspěšně uloženo.', 'success')

        excel_manager.update_project_info(project_name, start_date, end_date)

    return render_template('settings_page.html', settings=settings)

@app.route('/zalohy', methods=['GET', 'POST'])
def zalohy():
    if request.method == 'POST':
        employee_name = request.form['employee_name']
        amount = request.form['amount']
        currency = request.form['currency']
        option = request.form['option']
        date = request.form['date']

        try:
            excel_manager.save_advance(employee_name, amount, currency, option, date)
            flash('Záloha byla úspěšně uložena.', 'success')
        except Exception as e:
            flash(f'Chyba při ukládání zálohy: {str(e)}', 'error')

    employees = employee_manager.zamestnanci
    options = excel_manager.get_advance_options()
    current_date = datetime.now().strftime('%Y-%m-%d')

    return render_template('zalohy.html', employees=employees, options=options, current_date=current_date)

if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO, filename='app.log', filemode='a',
                        format='%(asctime)s - %(levelname)s - %(message)s')
    app.run(debug=True, host='0.0.0.0')