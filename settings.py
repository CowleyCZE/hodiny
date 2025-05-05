# settings.py
import json
import logging
import os
from datetime import datetime, timedelta
from functools import wraps
import re
import shutil

from flask import Flask, flash, redirect, render_template, request, session, url_for, g
from config import Config
from excel_manager import ExcelManager
from employee_management import EmployeeManager
from utils.logger import setup_logger

logger = setup_logger("settings")

# Inicializace Flask aplikace
app = Flask(__name__)
app.secret_key = Config.SECRET_KEY
Config.init_app(app)

# Pomocné funkce
def login_required(f):
    """Dekorátor pro routy vyžadující přihlášení"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        # V této verzi není přihlášení vyžadováno
        return f(*args, **kwargs)
    return decorated_function

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

@app.before_request
def before_request():
    """Inicializace manažerů před každým requestem"""
    global employee_manager, excel_manager
    
    try:
        # Inicializace manažerů pouze jednou na request
        if not hasattr(g, 'employee_manager'):
            g.employee_manager = EmployeeManager(Config.DATA_PATH)
        
        if not hasattr(g, 'excel_manager'):
            g.excel_manager = ExcelManager(Config.EXCEL_BASE_PATH, Config.EXCEL_TEMPLATE_NAME)
            
    except Exception as e:
        logger.error(f"Neočekávaná chyba při inicializaci manažerů: {e}", exc_info=True)
        g.excel_manager = None
        flash("Neočekávaná chyba při přípravě aplikace.", "error")

# Funkce pro správu nastavení
def save_settings_to_file(settings_data):
    """Uloží slovník nastavení do JSON souboru"""
    try:
        # Ujistíme se, že adresář pro nastavení existuje
        Config.SETTINGS_FILE_PATH.parent.mkdir(parents=True, exist_ok=True)
        
        with open(Config.SETTINGS_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(settings_data, f, indent=4, ensure_ascii=False)
            
        logger.info(f"Nastavení uložena do souboru: {Config.SETTINGS_FILE_PATH}")
        return True
        
    except IOError as e:
        logger.error(f"Chyba při zápisu do souboru nastavení '{Config.SETTINGS_FILE_PATH}': {e}", exc_info=True)
        return False
    except Exception as e:
        logger.error(f"Neočekávaná chyba při ukládání nastavení do souboru: {e}", exc_info=True)
        return False

def load_settings_from_file():
    """Načte nastavení ze souboru JSON, vrátí výchozí při chybě"""
    default_settings = Config.get_default_settings()
    
    if not Config.SETTINGS_FILE_PATH.exists():
        logger.warning(f"Soubor s nastavením '{Config.SETTINGS_FILE_PATH}' nenalezen, použijí se výchozí.")
        save_settings_to_file(default_settings)
        return default_settings
        
    try:
        with open(Config.SETTINGS_FILE_PATH, "r", encoding="utf-8") as f:
            loaded_settings = json.load(f)
            
        if not isinstance(loaded_settings, dict):
            raise ValueError("Neplatný formát JSON (očekáván slovník)")
            
        # Sloučení výchozího nastavení s načteným
        settings = default_settings.copy()
        settings.update(loaded_settings)
        
        # Validace struktury
        if not isinstance(settings.get("project_info"), dict):
            logger.warning("Klíč 'project_info' v nastavení není slovník, resetuji na výchozí.")
            settings["project_info"] = default_settings["project_info"]
            
        # Zajistíme, že active_excel_file je None nebo string
        if not isinstance(settings.get("active_excel_file"), (str, type(None))):
            logger.warning("Klíč 'active_excel_file' má neplatný typ, resetuji na None.")
            settings["active_excel_file"] = None
            
        logger.info(f"Nastavení úspěšně načtena ze souboru: {Config.SETTINGS_FILE_PATH}")
        return settings
        
    except (json.JSONDecodeError, ValueError) as e:
        logger.error(f"Chyba při čtení nebo parsování souboru nastavení '{Config.SETTINGS_FILE_PATH}': {e}. Použijí se výchozí.", exc_info=True)
        save_settings_to_file(default_settings)
        return default_settings.copy()
    except Exception as e:
        logger.error(f"Neočekávaná chyba při načítání nastavení ze souboru: {e}", exc_info=True)
        return default_settings.copy()

def ensure_active_excel_file(settings):
    """Zkontroluje a případně vytvoří aktivní Excel soubor"""
    active_filename = settings.get("active_excel_file")
    template_path = Config.EXCEL_BASE_PATH / Config.EXCEL_TEMPLATE_NAME
    
    if active_filename:
        active_file_path = Config.EXCEL_BASE_PATH / active_filename
        
        if active_file_path.exists():
            return settings
            
        logger.warning(f"Aktivní soubor '{active_filename}' neexistuje, vytvořím nový.")
    
    # Pokud neexistuje žádný aktivní soubor, vytvoříme nový
    try:
        # Získání informací o projektu
        project_info = settings.get("project_info", {})
        project_name = project_info.get("name", "Nový_projekt")
        project_start_str = project_info.get("start_date", "")
        project_end_str = project_info.get("end_date", "")
        
        # Validace začátku projektu
        if not project_start_str:
            logger.warning("Není zadán začátek projektu, použiji dnešní datum.")
            project_start_str = datetime.now().strftime("%Y-%m-%d")
            
        # Validace konce projektu
        if project_end_str:
            try:
                end_date = datetime.strptime(project_end_str, "%Y-%m-%d").date()
                if project_start_str:
                    start_date = datetime.strptime(project_start_str, "%Y-%m-%d").date()
                    if end_date < start_date:
                        logger.warning("Konec projektu je před začátkem, použiji datum začátku + 1 měsíc")
                        end_date = start_date + timedelta(days=30)
                        project_end_str = end_date.strftime("%Y-%m-%d")
            except ValueError as e:
                logger.warning(f"Neplatné datum konce: {e}, resetuji.")
                project_end_str = ""
                
        # Upravení názvu projektu pro použití v názvu souboru
        safe_project_name = re.sub(r"[\\/:*?\"<>|']", "", project_name).replace(" ", "_")
        if not safe_project_name:
            safe_project_name = "Projekt"
            
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        new_active_filename = f"{safe_project_name}_{timestamp}.xlsx"
        new_active_file_path = Config.EXCEL_BASE_PATH / new_active_filename
        
        # Kontrola existence šablony
        if not template_path.exists():
            logger.error(f"Šablona '{Config.EXCEL_TEMPLATE_NAME}' nebyla nalezena v '{Config.EXCEL_BASE_PATH}'.")
            settings["active_excel_file"] = None
            return settings
            
        # Vytvoření nového souboru
        shutil.copy2(template_path, new_active_file_path)
        logger.info(f"Nový aktivní soubor '{new_active_filename}' vytvořen z šablony.")
        
        # Aktualizace nastavení
        settings["active_excel_file"] = new_active_filename
        
        # Uložení aktualizovaného nastavení
        if not save_settings_to_file(settings):
            logger.error("Kritická chyba: Nepodařilo se uložit informaci o novém souboru.")
            settings["active_excel_file"] = None
            
    except Exception as e:
        logger.error(f"Nepodařilo se vytvořit nový Excel soubor: {e}", exc_info=True)
        settings["active_excel_file"] = None
        
    return settings

# Routy
@app.route('/nastaveni', methods=['GET', 'POST'])
@login_required
def settings_page():
    """Zobrazení a zpracování nastavení aplikace"""
    try:
        if request.method == 'POST':
            current_settings = session.get('settings', Config.get_default_settings())
            try:
                # Načtení hodnot z formuláře
                start_time_str = request.form.get("start_time", "").strip()
                end_time_str = request.form.get("end_time", "").strip()
                lunch_duration_str = request.form.get("lunch_duration", "").strip()
                project_name = request.form.get("project_name", "").strip()
                project_start_str = request.form.get("start_date", "").strip()
                project_end_str = request.form.get("end_date", "").strip()
                
                # Validace času
                if not start_time_str or not re.match(r"^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$", start_time_str):
                    raise ValueError("Neplatný formát počátečního času. Použijte HH:MM (00:00-23:59)")
                    
                if not end_time_str or not re.match(r"^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$", end_time_str):
                    raise ValueError("Neplatný formát koncového času. Použijte HH:MM (00:00-23:59)")
                    
                # Validace pauzy
                try:
                    lunch_duration = float(lunch_duration_str.replace(",", "."))
                    if lunch_duration <= 0:
                        raise ValueError("Délka oběda musí být kladné číslo")
                except ValueError:
                    raise ValueError("Neplatná délka oběda")
                    
                # Validace projektu
                if not project_name:
                    raise ValueError("Musíte zadat název projektu")
                    
                if project_start_str:
                    try:
                        datetime.strptime(project_start_str, "%Y-%m-%d")
                    except ValueError:
                        raise ValueError("Neplatný formát počátečního data. Použijte YYYY-MM-DD")
                        
                if project_end_str:
                    try:
                        end_date = datetime.strptime(project_end_str, "%Y-%m-%d")
                        if project_start_str:
                            start_date = datetime.strptime(project_start_str, "%Y-%m-%d")
                            if end_date < start_date:
                                raise ValueError("Datum konce projektu nemůže být dřívější než datum začátku")
                    except ValueError:
                        raise ValueError("Neplatný formát koncového data. Použijte YYYY-MM-DD")
                        
                # Uložení nastavení
                settings_to_save = current_settings.copy()
                settings_to_save.update({
                    "start_time": start_time_str,
                    "end_time": end_time_str,
                    "lunch_duration": lunch_duration,
                    "project_info": {
                        "name": project_name,
                        "start_date": project_start_str,
                        "end_date": project_end_str
                    }
                })
                
                if not save_settings_to_file(settings_to_save):
                    raise RuntimeError("Nepodařilo se uložit nastavení do souboru")
                    
                # Aktualizace Excel souboru
                active_filename = settings_to_save.get("active_excel_file")
                if active_filename:
                    excel_manager = ExcelManager(Config.EXCEL_BASE_PATH, active_filename, Config.EXCEL_TEMPLATE_NAME)
                    excel_update_success = excel_manager.update_project_info(project_name, project_start_str, project_end_str if project_end_str else None)
                    
                    if excel_update_success:
                        flash("Nastavení bylo úspěšně uloženo.", "success")
                    else:
                        flash("Nastavení bylo uloženo, ale nepodařilo se aktualizovat Excel.", "warning")
                else:
                    flash("Nastavení uloženo, ale není definován aktivní Excel soubor.", "info")
                    
                session['settings'] = settings_to_save
                logger.info("Nastavení úspěšně uložena do souboru a session.")
                
                return redirect(url_for('settings_page'))
                
            except (ValueError, RuntimeError) as e:
                flash(str(e), "error")
                logger.warning(f"Chyba při ukládání nastavení: {e}")
            except Exception as e:
                flash("Neočekávaná chyba při ukládání nastavení.", "error")
                logger.error(f"Neočekávaná chyba při ukládání nastavení: {e}", exc_info=True)
                
        # Načtení aktuálních nastavení
        settings = session.get('settings', {})
        return render_template('settings.html', settings=settings)
        
    except Exception as e:
        logger.error(f"Neočekávaná chyba na stránce nastavení: {e}", exc_info=True)
        flash("Došlo k chybě při načítání nastavení.", "error")
        return redirect(url_for('index'))

@app.route('/start_new_file', methods=['POST'])
@login_required
def start_new_file():
    """Začne nový soubor"""
    try:
        settings = session.get('settings', {})
        project_info = settings.get('project_info', {})
        project_name = project_info.get('name')
        project_start_str = project_info.get('start_date')
        project_end_str = project_info.get('end_date')
        
        # Validace všech informací o projektu
        if not all([project_name, project_start_str]):
            raise ValueError("Musíte zadat všechny informace o projektu")
            
        # Validace dat
        try:
            start_date = datetime.strptime(project_start_str, "%Y-%m-%d").date()
            if project_end_str:
                end_date = datetime.strptime(project_end_str, "%Y-%m-%d").date()
                if end_date < start_date:
                    raise ValueError("Konec projektu nemůže být před začátkem")
        except ValueError as e:
            raise ValueError(f"Neplatné datum: {e}")
            
        # Archivace starého souboru
        old_file = settings.get('active_excel_file')
        if old_file:
            archive_path = Config.EXCEL_ARCHIVE_PATH / old_file
            shutil.move(Config.EXCEL_BASE_PATH / old_file, archive_path)
            
        # Vytvoření nového souboru
        safe_project_name = re.sub(r"[\\/:*?\"<>|']", "", project_name).replace(" ", "_")
        if not safe_project_name:
            raise ValueError("Neplatný název projektu pro vytvoření souboru")
            
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        new_active_filename = f"{safe_project_name}_{timestamp}.xlsx"
        new_active_file_path = Config.EXCEL_BASE_PATH / new_active_filename
        
        # Kopírování šablony
        template_path = Config.EXCEL_BASE_PATH / Config.EXCEL_TEMPLATE_NAME
        if not template_path.exists():
            raise FileNotFoundError(f"Šablona '{Config.EXCEL_TEMPLATE_NAME}' nebyla nalezena")
            
        shutil.copy2(template_path, new_active_file_path)
        logger.info(f"Vytvořen nový soubor '{new_active_filename}' z šablony")
        
        # Aktualizace nastavení
        settings['active_excel_file'] = new_active_filename
        settings['project_info'] = {
            "name": project_name,
            "start_date": project_start_str,
            "end_date": project_end_str
        }
        
        if not save_settings_to_file(settings):
            logger.error("Kritická chyba: Nepodařilo se uložit informaci o novém souboru!")
            flash("Kritická chyba: Nepodařilo se uložit informaci o novém souboru. Kontaktujte administrátora.", "error")
            settings["active_excel_file"] = None
        else:
            flash(f"Nový soubor '{new_active_filename}' byl vytvořen", "success")
            
        session['settings'] = settings
        return redirect(url_for('index'))
        
    except ValueError as e:
        flash(str(e), "error")
    except FileNotFoundError as e:
        flash(f"Chyba: {str(e)}", "error")
    except Exception as e:
        logger.error(f"Neočekávaná chyba při vytváření nového souboru: {e}", exc_info=True)
        flash("Došlo k neočekávané chybě při vytváření nového souboru.", "error")
        
    return redirect(url_for('settings_page'))

@app.route('/archivace', methods=['POST'])
@login_required
def archive_file():
    """Archivuje aktuální soubor a vytvoří nový"""
    try:
        settings = session.get('settings', {})
        project_info = settings.get('project_info', {})
        project_name = project_info.get('name')
        project_start_str = project_info.get('start_date')
        project_end_str = project_info.get('end_date')
        
        if not project_end_str:
            raise ValueError("Před archivací souboru musí být zadáno datum konce projektu")
            
        # Validace dat
        try:
            end_date = datetime.strptime(project_end_str, "%Y-%m-%d").date()
            if project_start_str:
                start_date = datetime.strptime(project_start_str, "%Y-%m-%d").date()
                if end_date < start_date:
                    raise ValueError("Datum konce projektu nemůže být dřívější než datum začátku")
        except ValueError as e:
            raise ValueError(f"Neplatné datum konce: {e}")
            
        # Archivace souboru
        current_active_file = settings.get('active_excel_file')
        if not current_active_file:
            raise ValueError("Není definován aktivní Excel soubor pro archivaci")
            
        active_file_path = Config.EXCEL_BASE_PATH / current_active_file
        if active_file_path.exists():
            archive_path = Config.EXCEL_ARCHIVE_PATH / current_active_file
            shutil.move(active_file_path, archive_path)
            logger.info(f"Aktivní soubor '{current_active_file}' archivován")
            
        # Vytvoření nového souboru
        template_path = Config.EXCEL_BASE_PATH / Config.EXCEL_TEMPLATE_NAME
        if not template_path.exists():
            raise FileNotFoundError(f"Šablona '{Config.EXCEL_TEMPLATE_NAME}' nebyla nalezena")
            
        shutil.copy2(template_path, active_file_path)
        logger.info(f"Vytvořen nový soubor pro projekt '{project_name}'")
        
        # Aktualizace nastavení
        settings['project_info']['start_date'] = project_end_str
        settings['project_info']['end_date'] = ""
        settings['project_info']['name'] = project_name
        settings['active_excel_file'] = current_active_file
        
        if not save_settings_to_file(settings):
            logger.error("Kritická chyba: Nepodařilo se uložit informaci o novém souboru!")
            flash("Kritická chyba: Nepodařilo se uložit informaci o novém souboru.", "error")
            settings["active_excel_file"] = None
        else:
            flash("Soubor byl úspěšně archivován a vytvořen nový soubor pro projekt.", "success")
            
        session['settings'] = settings
        return redirect(url_for('index'))
        
    except ValueError as e:
        flash(str(e), "error")
    except FileNotFoundError as e:
        flash(f"Chyba: {str(e)}", "error")
    except Exception as e:
        logger.error(f"Neočekávaná chyba při archivaci souboru: {e}", exc_info=True)
        flash("Došlo k neočekávané chybě při archivaci souboru.", "error")
        
    return redirect(url_for('settings_page'))

# Spuštění aplikace
if __name__ == '__main__':
    # Nastavení logování pro vývoj
    if not app.debug:
        log_handler = logging.FileHandler('app_dev.log', encoding='utf-8')
        log_handler.setLevel(logging.INFO)
        log_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        log_handler.setFormatter(log_formatter)
        app.logger.addHandler(log_handler)
    else:
        app.logger.setLevel(logging.DEBUG)
        
    app.run(debug=True, host='0.0.0.0', port=5000)
