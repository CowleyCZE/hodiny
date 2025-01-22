from flask import Flask, request, render_template, redirect, url_for, flash
import json
import logging
import os

from excel_manager import ExcelManager

# Flask aplikace
app = Flask(__name__)
app.secret_key = 'your_secret_key'

# Konstanty
BASE_DIR = '/home/Cowley/hodiny'
DATA_PATH = os.path.join(BASE_DIR, 'data')
SETTINGS_FILE_PATH = os.path.join(DATA_PATH, 'settings.json')
EXCEL_BASE_PATH = os.path.join(BASE_DIR, 'excel')

def load_settings():
    """Načtení nastavení ze souboru JSON."""
    default_settings = {
        'start_time': '07:00',
        'end_time': '18:00',
        'lunch_duration': 1,
        'project_info': {
            'name': '',
            'start_date': '',
            'end_date': ''
        }
    }

    try:
        if not os.path.exists(SETTINGS_FILE_PATH):
            return default_settings

        with open(SETTINGS_FILE_PATH, 'r', encoding='utf-8') as f:
            saved_settings = json.load(f)
            # Sloučení uložených nastavení s výchozími
            default_settings.update(saved_settings)
            return default_settings
    except Exception as e:
        logging.warning(f"Chyba při načítání nastavení: {e}")
        return default_settings

def save_settings(settings_data):
    """Uložení nastavení do JSON a aktualizace Excel souboru."""
    try:
        # Kontrola existence adresáře
        os.makedirs(os.path.dirname(SETTINGS_FILE_PATH), exist_ok=True)

        # Načtení existujících nastavení (pokud existují)
        current_settings = load_settings()

        # Aktualizace nastavení novými hodnotami
        current_settings.update(settings_data)

        # Uložení nastavení do JSON souboru
        with open(SETTINGS_FILE_PATH, 'w', encoding='utf-8') as f:
            json.dump(current_settings, f, indent=4, ensure_ascii=False)

        # Aktualizace projektových dat v Excel souboru
        excel_manager = ExcelManager(EXCEL_BASE_PATH)
        project_name = settings_data['project_info']['name']
        start_date = settings_data['project_info']['start_date']
        end_date = settings_data['project_info']['end_date']

        excel_manager.update_project_info_in_advances(project_name, start_date, end_date)

        logging.info("Nastavení byla úspěšně uložena.")
        return True
    except Exception as e:
        logging.error(f"Chyba při ukládání nastavení: {e}")
        return False

@app.route('/settings', methods=['GET', 'POST'])
def settings_page():
    """Zobrazení a zpracování stránky pro nastavení."""
    if request.method == 'POST':
        logging.info("Přijat POST požadavek na stránce nastavení")
        settings_data = {
            'start_time': request.form['start_time'],
            'end_time': request.form['end_time'],
            'lunch_duration': float(request.form['lunch_duration']),
            'project_info': {
                'name': request.form['project_name'],
                'start_date': request.form['start_date'],
                'end_date': request.form['end_date']
            }
        }
        logging.debug(f"Přijatá data nastavení: {settings_data}")

        # Zápis nastavení
        if save_settings(settings_data):
            logging.info(f"Nastavení pro projekt '{settings_data['project_info']['name']}' úspěšně uložena")
            flash(f'Nastavení pro projekt "{settings_data["project_info"]["name"]}" byla úspěšně uložena.', 'success')
        else:
            logging.error("Chyba při ukládání nastavení")
            flash('Chyba při ukládání nastavení.', 'error')

        return redirect(url_for('settings_page'))

    # Načtení aktuálního nastavení
    current_settings = load_settings()
    logging.debug(f"Načtená aktuální nastavení: {current_settings}")
    return render_template('settings_page.html', settings=current_settings)

if __name__ == '__main__':
    app.run(debug=True)