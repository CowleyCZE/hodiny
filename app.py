from flask import Flask, render_template, request, jsonify, redirect, url_for, flash, send_file
from datetime import datetime
from employee_management import EmployeeManagement
from excel_manager import ExcelManager
from zalohy_manager import ZalohyManager
import logging
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

app = Flask(__name__)
app.secret_key = 'tajny_klic_pro_flash_zpravy'

employee_manager = EmployeeManagement()
excel_manager = ExcelManager()
zalohy_manager = ZalohyManager()

EXCEL_FILE_PATH = 'Hodiny_Cap.xlsx'
RECIPIENT_EMAIL = 'czechmontagesro@gmail.com'

@app.route('/')
def index():
    logging.info("Přístup na hlavní stránku")
    excel_exists = os.path.exists(EXCEL_FILE_PATH)
    return render_template('index.html', excel_exists=excel_exists)

@app.route('/download-excel')
def download_excel():
    try:
        return send_file(
            EXCEL_FILE_PATH,
            as_attachment=True,
            download_name='Hodiny_Cap.xlsx'
        )
    except Exception as e:
        flash(f'Chyba při stahování souboru: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/open-excel')
def open_excel():
    try:
        return send_file(
            EXCEL_FILE_PATH,
            as_attachment=False,
            download_name='Hodiny_Cap.xlsx'
        )
    except Exception as e:
        flash(f'Chyba při otevírání souboru: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/send-excel', methods=['POST'])
def send_excel():
    try:
        # Email configuration
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        sender_email = "your-email@gmail.com"  # Nahraďte vlastním emailem
        password = "your-app-password"  # Nahraďte vlastním heslem pro aplikaci

        # Create message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = RECIPIENT_EMAIL
        msg['Subject'] = 'Hodiny_Cap.xlsx - Export'

        # Add body
        body = "V příloze najdete aktuální výkaz hodin."
        msg.attach(MIMEText(body, 'plain'))

        # Add attachment
        with open(EXCEL_FILE_PATH, 'rb') as f:
            attachment = MIMEApplication(f.read(), _subtype='xlsx')
            attachment.add_header('Content-Disposition', 'attachment', filename='Hodiny_Cap.xlsx')
            msg.attach(attachment)

        # Send email
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, password)
            server.send_message(msg)

        flash('Soubor byl úspěšně odeslán emailem.', 'success')
    except Exception as e:
        flash(f'Chyba při odesílání emailu: {str(e)}', 'error')
    
    return redirect(url_for('index'))


@app.route('/employees', methods=['GET', 'POST'])
def manage_employees():
    logging.info(f"manage_employees called, method: {request.method}")
    if request.method == 'POST':
        action = request.form.get('action')
        logging.info(f"Action: {action}")
        
        if action == 'add':
            name = request.form.get('name')
            if name and name not in employee_manager.zamestnanci:
                success = employee_manager.pridat_zamestnance(name)
                if success:
                    flash(f'Zaměstnanec {name} byl úspěšně přidán.', 'success')
                else:
                    flash('Chyba při přidávání zaměstnance.', 'error')
            else:
                flash('Jméno zaměstnance je prázdné nebo již existuje.', 'error')
        
        elif action == 'update':
            index = int(request.form.get('index')) - 1
            new_name = request.form.get('name')
            if 0 <= index < len(employee_manager.zamestnanci) and new_name:
                old_name = employee_manager.zamestnanci[index]
                employee_manager.zamestnanci[index] = new_name
                employee_manager.save_config()
                flash(f'Jméno zaměstnance bylo změněno z {old_name} na {new_name}.', 'success')
            else:
                flash('Neplatný index zaměstnance nebo prázdné jméno.', 'error')
        
        elif action == 'delete':
            index = int(request.form.get('index')) - 1
            if 0 <= index < len(employee_manager.zamestnanci):
                deleted_name = employee_manager.zamestnanci.pop(index)
                if deleted_name in employee_manager.vybrani_zamestnanci:
                    employee_manager.vybrani_zamestnanci.remove(deleted_name)
                employee_manager.save_config()
                flash(f'Zaměstnanec {deleted_name} byl smazán.', 'success')
            else:
                flash('Neplatný index zaměstnance.', 'error')
        
        elif action == 'toggle':
            index = int(request.form.get('index')) - 1
            if 0 <= index < len(employee_manager.zamestnanci):
                employee = employee_manager.zamestnanci[index]
                if employee in employee_manager.vybrani_zamestnanci:
                    employee_manager.odebrat_vybraneho_zamestnance(employee)
                    flash(f'Zaměstnanec {employee} byl odznačen.', 'info')
                else:
                    employee_manager.pridat_vybraneho_zamestnance(employee)
                    flash(f'Zaměstnanec {employee} byl označen.', 'info')
            else:
                flash('Neplatný index zaměstnance.', 'error')
        
        return redirect(url_for('manage_employees'))
    
    employees = [{"name": e, "selected": e in employee_manager.vybrani_zamestnanci} for e in employee_manager.zamestnanci]
    return render_template('employees.html', employees=employees)

@app.route('/record_time', methods=['GET', 'POST'])
def record_time():
    if request.method == 'POST':
        date = request.form.get('date')
        start_time = request.form.get('start_time')
        end_time = request.form.get('end_time')
        lunch_duration = float(request.form.get('lunch_duration', 0))
        
        try:
            date_obj = datetime.strptime(date, '%Y-%m-%d')
            
            excel_manager.ulozit_pracovni_dobu(date_obj, start_time, end_time, lunch_duration, employee_manager.vybrani_zamestnanci)
            
            logging.info(f"Záznam pracovní doby uložen: datum={date}, začátek={start_time}, konec={end_time}, oběd={lunch_duration}")
            return jsonify({"message": "Záznam pracovní doby byl úspěšně uložen do Excel souboru."})
        except Exception as e:
            logging.error(f"Chyba při ukládání pracovní doby: {str(e)}")
            return jsonify({"error": str(e)}), 400
    
    return render_template('record_time.html')

@app.route('/zalohy', methods=['GET', 'POST'])
def zalohy():
    if request.method == 'POST':
        employee_name = request.form.get('employee_name')
        amount = float(request.form.get('amount'))
        currency = request.form.get('currency')
        option = request.form.get('option')
        date = request.form.get('date')

        try:
            success = zalohy_manager.add_or_update_employee_advance(employee_name, amount, currency, option, date)
            if success:
                flash(f'Záloha pro zaměstnance {employee_name} byla úspěšně zapsána.', 'success')
            else:
                flash('Chyba při zápisu zálohy do Excel souboru.', 'error')
        except Exception as e:
            flash(f'Chyba při zpracování zálohy: {str(e)}', 'error')

        return redirect(url_for('zalohy'))

    employees = employee_manager.zamestnanci
    option1_name, option2_name = zalohy_manager.get_option_names()
    return render_template('zalohy.html', employees=employees, current_date=datetime.today().strftime('%Y-%m-%d'),
                           option1_name=option1_name, option2_name=option2_name)

if __name__ == '__main__':
    app.run(debug=True)