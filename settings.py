from flask import Flask, request, render_template, redirect, url_for, flash
import json
import logging
import os
from config import Config
from utils.logger import setup_logger
from excel_manager import ExcelManager

# Flask aplikace
app = Flask(__name__)
app.secret_key = Config.SECRET_KEY

# Konstanty - použití konfigurací z Config
DATA_PATH = Config.DATA_PATH
SETTINGS_FILE_PATH = Config.SETTINGS_FILE_PATH
EXCEL_BASE_PATH = Config.EXCEL_BASE_PATH

logger = setup_logger('settings')

def load_settings():
    """Načtení nastavení ze souboru JSON."""
    default_settings = Config.get_default_settings()
    
    try:
        if not os.path.exists(Config.SETTINGS_FILE_PATH):
            logger.info("Soubor s nastavením neexistuje, používám výchozí hodnoty")
            return default_settings

        with open(Config.SETTINGS_FILE_PATH, 'r', encoding='utf-8') as f:
            saved_settings = json.load(f)
            
            # Validace struktury nastavení
            if not isinstance(saved_settings, dict):
                raise ValueError("Neplatný formát nastavení")
                
            # Sloučení časového nastavení
            for key in vars(Config.DEFAULT_TIME_CONFIG):
                if key in saved_settings:
                    default_settings[key] = saved_settings[key]
                    
            # Sloučení projektového nastavení
            if 'project_info' in saved_settings:
                project_info = saved_settings['project_info']
                if isinstance(project_info, dict):
                    for key in vars(Config.DEFAULT_PROJECT_CONFIG):
                        if key in project_info:
                            default_settings['project_info'][key] = project_info[key]
            
            logger.info("Nastavení úspěšně načteno")
            return default_settings
            
    except Exception as e:
        logger.error(f"Chyba při načítání nastavení: {str(e)}")
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

        excel_manager.update_project_info(project_name, start_date, end_date)

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
