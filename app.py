# app.py
import json
import logging
import os
import re
import shutil
import smtplib
import ssl
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import parseaddr
from pathlib import Path
from functools import wraps # Pro dekorátor

import openpyxl
import pandas as pd
from flask import (Flask, flash, jsonify, redirect, render_template, request,
                   send_file, session, url_for, g) # Přidáno g pro ukládání manageru
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook import Workbook

# Local imports
from config import Config
from employee_management import EmployeeManager
from excel_manager import ExcelManager
from utils.logger import setup_logger
from zalohy_manager import ZalohyManager

# Nahrazení základního loggeru naším vlastním
logger = setup_logger("app")

# Initialize Flask app
app = Flask(__name__)
app.secret_key = Config.SECRET_KEY
Config.init_app(app) # Inicializace konfigurace a adresářů

# --- Globální instance manažerů (budou inicializovány v before_request) ---
# Tyto proměnné budou obsahovat instance pro aktuální request
# employee_manager = None # Bude v g.employee_manager
# excel_manager = None    # Bude v g.excel_manager
# zalohy_manager = None   # Bude v g.zalohy_manager

# --- Funkce pro správu nastavení a aktivního souboru ---

def save_settings_to_file(settings_data):
    """Uloží slovník nastavení do JSON souboru."""
    try:
        # Zajistíme existenci adresáře
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
    """Načte nastavení z JSON souboru, vrátí výchozí při chybě."""
    default_settings = Config.get_default_settings()
    if not Config.SETTINGS_FILE_PATH.exists():
        logger.warning(f"Soubor s nastavením '{Config.SETTINGS_FILE_PATH}' nenalezen, použijí se výchozí.")
        # Uložíme výchozí nastavení pro příští spuštění
        save_settings_to_file(default_settings)
        return default_settings

    try:
        with open(Config.SETTINGS_FILE_PATH, "r", encoding="utf-8") as f:
            try:
                loaded_settings = json.load(f)
                if not isinstance(loaded_settings, dict):
                     raise ValueError("Neplatný formát JSON (očekáván slovník)")
                # Sloučíme načtené s výchozími, abychom zajistili všechny klíče
                # Načtené hodnoty přepíší výchozí
                settings = default_settings.copy()
                settings.update(loaded_settings)
                # Zkontrolujeme, zda project_info je stále slovník
                if not isinstance(settings.get("project_info"), dict):
                     logger.warning("Klíč 'project_info' v nastavení není slovník, resetuji na výchozí.")
                     settings["project_info"] = default_settings["project_info"]

                logger.info(f"Nastavení úspěšně načtena ze souboru: {Config.SETTINGS_FILE_PATH}")
                return settings
            except (json.JSONDecodeError, ValueError) as e:
                logger.error(f"Chyba při čtení nebo parsování souboru nastavení '{Config.SETTINGS_FILE_PATH}': {e}. Použijí se výchozí.", exc_info=True)
                # V případě chyby vrátíme výchozí a zkusíme je uložit
                save_settings_to_file(default_settings)
                return default_settings.copy()
    except Exception as e:
        logger.error(f"Neočekávaná chyba při načítání nastavení ze souboru: {e}", exc_info=True)
        return default_settings.copy()


def ensure_active_excel_file(settings):
    """
    Zkontroluje, zda existuje aktivní Excel soubor definovaný v nastavení.
    Pokud neexistuje nebo není definován, vytvoří nový z šablony.
    Vrátí aktualizovaný slovník nastavení.
    """
    active_filename = settings.get("active_excel_file")
    template_path = Config.EXCEL_BASE_PATH / Config.EXCEL_TEMPLATE_NAME
    file_created = False

    if active_filename:
        active_file_path = Config.EXCEL_BASE_PATH / active_filename
        if active_file_path.exists():
            logger.debug(f"Aktivní soubor '{active_filename}' nalezen.")
            return settings # Vše je v pořádku, vracíme původní nastavení
        else:
            logger.warning(f"Aktivní soubor '{active_filename}' definovaný v nastavení nebyl nalezen.")
            # Budeme pokračovat a vytvoříme nový

    # Pokud aktivní soubor chybí nebo nebyl nalezen, vytvoříme nový
    logger.info("Vytváření nového aktivního Excel souboru...")
    project_name = settings.get("project_info", {}).get("name", "NeznamyProjekt")
    # Odstraníme nebezpečné znaky z názvu projektu pro název souboru
    safe_project_name = re.sub(r'[\\/*?:"<>|]', "", project_name).replace(" ", "_")
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    new_active_filename = f"{safe_project_name}_{timestamp}.xlsx"
    new_active_file_path = Config.EXCEL_BASE_PATH / new_active_filename

    if not template_path.exists():
        logger.error(f"Šablona '{Config.EXCEL_TEMPLATE_NAME}' nebyla nalezena v '{Config.EXCEL_BASE_PATH}'. Nelze vytvořit nový aktivní soubor.")
        # V tomto případě nemůžeme pokračovat, vrátíme původní nastavení s chybou
        # Nebo bychom mohli vyvolat výjimku
        settings["active_excel_file"] = None # Zajistíme, že není nastaven žádný aktivní soubor
        flash(f"Chyba: Šablona '{Config.EXCEL_TEMPLATE_NAME}' nebyla nalezena. Kontaktujte administrátora.", "error")
        return settings


    try:
        shutil.copy2(template_path, new_active_file_path)
        logger.info(f"Nový aktivní soubor '{new_active_filename}' vytvořen z šablony.")
        settings["active_excel_file"] = new_active_filename
        file_created = True
    except Exception as e:
        logger.error(f"Nepodařilo se zkopírovat šablonu '{template_path}' do '{new_active_file_path}': {e}", exc_info=True)
        settings["active_excel_file"] = None # Resetujeme aktivní soubor v nastavení
        flash("Chyba při vytváření nového Excel souboru.", "error")

    # Pokud byl soubor úspěšně vytvořen, uložíme aktualizované nastavení
    if file_created:
        if not save_settings_to_file(settings):
             # Pokud se nepodaří uložit nastavení, měli bychom ideálně vrátit změny zpět
             # (např. smazat nově vytvořený soubor), ale pro jednoduchost jen zalogujeme
             logger.error("Kritická chyba: Nový aktivní soubor byl vytvořen, ale nepodařilo se uložit jeho název do nastavení!")
             flash("Kritická chyba: Nepodařilo se uložit informaci o novém souboru. Kontaktujte administrátora.", "error")
             settings["active_excel_file"] = None # Vrátíme stav bez aktivního souboru

    return settings


# --- Request Handlers ---

@app.before_request
def before_request():
    """Spustí se před každým requestem."""
    # 1. Načteme nastavení ze souboru (ne ze session, aby byly změny aktuální)
    settings = load_settings_from_file()

    # 2. Zajistíme existenci aktivního souboru (vytvoří nový, pokud je potřeba)
    #    Tato funkce také uloží nastavení, pokud vytvoří nový soubor.
    settings = ensure_active_excel_file(settings)

    # 3. Uložíme aktuální (možná aktualizovaná) nastavení do session pro použití v šablonách
    session['settings'] = settings

    # 4. Inicializujeme manažery s aktuálními daty a uložíme je do 'g' (globální objekt pro request)
    g.employee_manager = EmployeeManager(Config.DATA_PATH)
    active_filename = settings.get("active_excel_file")
    if active_filename:
         # Předáme base_path, název aktivního souboru a název šablony
         g.excel_manager = ExcelManager(Config.EXCEL_BASE_PATH, active_filename, Config.EXCEL_TEMPLATE_NAME)
         # Nastavíme aktuální název projektu do ExcelManageru
         project_name = settings.get("project_info", {}).get("name")
         if project_name:
              g.excel_manager.set_project_name(project_name)
         # ZalohyManager nyní také potřebuje vědět, se kterým souborem pracovat
         g.zalohy_manager = ZalohyManager(Config.EXCEL_BASE_PATH, active_filename)
    else:
         # Pokud není aktivní soubor (např. kvůli chybě při vytváření),
         # nastavíme manažery na None, aby operace selhaly kontrolovaně.
         g.excel_manager = None
         g.zalohy_manager = None
         logger.error("ExcelManager a ZalohyManager nebyly inicializovány, protože není definován aktivní Excel soubor.")
         # Můžeme přidat flash zprávu, pokud už nebyla přidána v ensure_active_excel_file
         if not any(msg[1] == 'error' for msg in get_flashed_messages(with_categories=True)):
              flash("Chyba: Není definován aktivní Excel soubor pro práci. Zkuste archivovat a začít nový nebo kontaktujte administrátora.", "error")


@app.teardown_request
def teardown_request(exception=None):
    """Spustí se po každém requestu, i když dojde k výjimce."""
    # Vyčistíme cache workbooků v ExcelManageru, pokud existuje
    excel_manager_instance = getattr(g, 'excel_manager', None)
    if excel_manager_instance:
        try:
            excel_manager_instance.close_cached_workbooks()
            logger.debug("Cache workbooků vyčištěna na konci requestu.")
        except Exception as e:
            logger.error(f"Chyba při čištění cache workbooků na konci requestu: {e}", exc_info=True)

    # Odstraníme manažery z 'g', aby se nepřenášely mezi requesty (pro jistotu)
    g.pop('employee_manager', None)
    g.pop('excel_manager', None)
    g.pop('zalohy_manager', None)


# --- Dekorátor pro kontrolu existence manažerů ---
def require_excel_managers(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not getattr(g, 'excel_manager', None) or not getattr(g, 'zalohy_manager', None):
            # Flash zpráva by měla být přidána v before_request
            logger.error(f"Přístup k route '{request.path}' zamítnut, protože Excel manažeři nejsou inicializováni.")
            return redirect(url_for('index')) # Nebo na stránku s chybou
        return f(*args, **kwargs)
    return decorated_function

# --- Routes ---

@app.route("/")
def index():
    # Získáme nastavení ze session (aktualizovaná v before_request)
    settings = session.get('settings', {})
    active_filename = settings.get('active_excel_file')
    excel_exists = False
    week_num_int = 0

    if active_filename:
         active_file_path = Config.EXCEL_BASE_PATH / active_filename
         excel_exists = active_file_path.exists()
         # Získáme číslo týdne z ExcelManageru, pokud je inicializován
         excel_manager_instance = getattr(g, 'excel_manager', None)
         if excel_manager_instance:
              current_date = datetime.now().strftime("%Y-%m-%d")
              week_calendar_data = excel_manager_instance.ziskej_cislo_tydne(current_date)
              week_num_int = week_calendar_data.week if week_calendar_data else 0
         else:
              logger.warning("ExcelManager není k dispozici pro získání čísla týdne v index route.")
    else:
         # Pokud není aktivní soubor, zobrazíme varování
         flash("Není aktivní žádný Excel soubor pro záznamy. Můžete začít nový v Nastavení.", "warning")


    return render_template(
        "index.html",
        excel_exists=excel_exists, # Zda existuje aktivní soubor
        active_filename=active_filename, # Předáme název aktivního souboru šabloně
        week_number=week_num_int,
        current_date=datetime.now().strftime("%Y-%m-%d")
    )


@app.route("/download")
@require_excel_managers # Zajistí, že g.excel_manager existuje
def download_file():
    """Stáhne aktivní Excel soubor s názvem obsahujícím nejvyšší číslo týdne."""
    workbook = None
    try:
        active_file_path = g.excel_manager.get_active_file_path() # Získá cestu k aktivnímu souboru

        # Načtení workbooku pro zjištění názvů listů (read-only)
        try:
            workbook = load_workbook(active_file_path, read_only=True)
            sheet_names = workbook.sheetnames
        except Exception as e:
            logger.error(f"Nepodařilo se načíst aktivní soubor '{active_file_path.name}' pro čtení listů: {e}", exc_info=True)
            raise ValueError(f"Nepodařilo se otevřít soubor '{active_file_path.name}' pro analýzu.")
        finally:
             if workbook:
                workbook.close()

        # Nalezení nejvyššího čísla týdne
        max_week_number = 0
        week_pattern = re.compile(r"Týden (\d+)")
        for sheet_name in sheet_names:
            match = week_pattern.match(sheet_name)
            if match:
                try:
                    week_num = int(match.group(1))
                    if week_num > max_week_number:
                        max_week_number = week_num
                except ValueError: continue

        # Vytvoření názvu pro stažení
        base_name = active_file_path.stem # Název souboru bez přípony
        if max_week_number > 0:
            # Použijeme název aktivního souboru a přidáme týden
            # Např. ProjektX_20250503_Tyden18.xlsx
            download_filename = f"{base_name}_Tyden{max_week_number}.xlsx"
        else:
            # Pokud nejsou týdny, stáhne se pod původním názvem aktivního souboru
            download_filename = active_file_path.name
            logger.warning(f"V souboru '{active_file_path.name}' nenalezen žádný list 'Týden X', stahuje se pod původním názvem.")

        # Odeslání aktivního souboru ke stažení s novým názvem
        # Není potřeba vytvářet kopii, posíláme přímo aktivní soubor
        return send_file(
            str(active_file_path),
            as_attachment=True,
            download_name=download_filename
        )

    except (FileNotFoundError, ValueError, IOError) as e:
        logger.error(f"Chyba při přípravě souboru ke stažení: {e}")
        flash(str(e), "error")
        return redirect(url_for("index"))
    except Exception as e:
        logger.error(f"Neočekávaná chyba při stahování souboru: {e}", exc_info=True)
        flash("Neočekávaná chyba při stahování souboru.", "error")
        return redirect(url_for("index"))


@app.route("/send_email", methods=["POST"])
@require_excel_managers # Zajistí, že g.excel_manager existuje
def send_email():
    """Odešle aktivní Excel soubor emailem."""
    try:
        active_file_path = g.excel_manager.get_active_file_path()
        active_filename = active_file_path.name # Získáme název aktivního souboru

        # Kontrola emailové konfigurace (stejná jako dříve)
        # ... (kód pro kontrolu konfigurace) ...
        if not Config.RECIPIENT_EMAIL: raise ValueError("E-mail příjemce není nastaven.")
        if not Config.SMTP_USERNAME or not Config.SMTP_PASSWORD: raise ValueError("SMTP údaje nejsou nastaveny.")
        if not Config.SMTP_SERVER or not Config.SMTP_PORT: raise ValueError("SMTP server/port není nastaven.")

        sender = Config.SMTP_USERNAME
        recipient = Config.RECIPIENT_EMAIL
        validate_email(sender)
        validate_email(recipient)

        # Vytvoření zprávy
        msg = MIMEMultipart()
        subject = f'{active_filename} - {datetime.now().strftime("%Y-%m-%d %H:%M")}'
        msg["Subject"] = subject
        msg["From"] = sender
        msg["To"] = recipient

        app_name = getattr(Config, 'APP_NAME', 'Evidence pracovní doby')
        body = f"""Dobrý den,\n\nv příloze zasílám aktuální výkaz pracovní doby ({active_filename}).\n\nS pozdravem,\n{app_name}"""
        msg.attach(MIMEText(body, "plain", "utf-8"))

        # Přidání aktivního souboru jako přílohy
        try:
            with open(active_file_path, "rb") as f:
                attachment = MIMEApplication(f.read(), _subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                attachment.add_header("Content-Disposition", "attachment", filename=active_filename)
                msg.attach(attachment)
        except IOError as e:
             logger.error(f"Chyba při čtení aktivního souboru '{active_filename}' pro email: {e}", exc_info=True)
             raise ValueError(f"Nepodařilo se připojit soubor '{active_filename}' k emailu.")

        # Odeslání emailu (stejný kód jako dříve)
        # ... (kód pro nastavení SSL a odeslání přes smtplib) ...
        ssl_context = ssl.create_default_context()
        with smtplib.SMTP_SSL(Config.SMTP_SERVER, Config.SMTP_PORT, context=ssl_context, timeout=60) as smtp:
            smtp.login(Config.SMTP_USERNAME, Config.SMTP_PASSWORD)
            smtp.send_message(msg)

        flash("Email byl úspěšně odeslán.", "success")
        logger.info(f"Email s výkazem '{active_filename}' byl úspěšně odeslán na adresu {recipient}")

    except FileNotFoundError as e:
        logger.error(f"Aktivní soubor pro odeslání emailem nebyl nalezen: {e}")
        flash(str(e), "error")
    except ValueError as e:
        logger.error(f"Chyba konfigurace nebo dat pro odeslání emailu: {e}")
        flash(str(e), "error")
    except (ConnectionError, IOError, smtplib.SMTPException, ssl.SSLError, TimeoutError) as e:
         logger.error(f"Chyba připojení, souboru nebo SMTP při odesílání emailu: {e}", exc_info=True)
         flash(f"Chyba při odesílání emailu: {e}", "error")
    except Exception as e:
        logger.error(f"Neočekávaná chyba v procesu odesílání emailu: {e}", exc_info=True)
        flash("Neočekávaná chyba při odesílání emailu.", "error")

    return redirect(url_for("index"))


@app.route("/zamestnanci", methods=["GET", "POST"])
def manage_employees():
    # Použijeme g.employee_manager inicializovaný v before_request
    employee_manager_instance = getattr(g, 'employee_manager', None)
    if not employee_manager_instance:
         # Toto by nemělo nastat, ale pro jistotu
         flash("Správce zaměstnanců není k dispozici.", "error")
         return redirect(url_for('index'))

    if request.method == "POST":
        action = request.form.get("action")
        try:
            if not action: raise ValueError("Nebyla specifikována akce")

            if action == "add":
                employee_name = request.form.get("name", "").strip()
                if not employee_name: raise ValueError("Jméno zaměstnance nemůže být prázdné")
                if len(employee_name) > 100: raise ValueError("Jméno zaměstnance je příliš dlouhé")
                if not re.match(r"^[\w\s\-\.ěščřžýáíéúůďťňĚŠČŘŽÝÁÍÉÚŮĎŤŇ]+$", employee_name):
                    raise ValueError("Jméno zaměstnance obsahuje nepovolené znaky.")

                if employee_manager_instance.pridat_zamestnance(employee_name):
                    flash(f'Zaměstnanec "{employee_name}" byl přidán.', "success")
                    return redirect(url_for('manage_employees'))
                else:
                    flash(f'Zaměstnanec "{employee_name}" již existuje nebo došlo k chybě.', "error")

            elif action == "select":
                employee_name = request.form.get("employee_name", "")
                if not employee_name: raise ValueError("Nebyl vybrán zaměstnanec")
                if employee_name not in employee_manager_instance.zamestnanci:
                     raise ValueError(f'Zaměstnanec "{employee_name}" neexistuje')

                if employee_name in employee_manager_instance.vybrani_zamestnanci:
                    if employee_manager_instance.odebrat_vybraneho_zamestnance(employee_name):
                         flash(f'Zaměstnanec "{employee_name}" byl odebrán z výběru.', "success")
                    else: flash(f'Nepodařilo se odebrat "{employee_name}" z výběru.', "error")
                else:
                    if employee_manager_instance.pridat_vybraneho_zamestnance(employee_name):
                         flash(f'Zaměstnanec "{employee_name}" byl přidán do výběru.', "success")
                    else: flash(f'Nepodařilo se přidat "{employee_name}" do výběru.', "error")
                return redirect(url_for('manage_employees'))

            elif action == "edit":
                old_name = request.form.get("old_name", "").strip()
                new_name = request.form.get("new_name", "").strip()
                if not old_name or not new_name: raise ValueError("Původní i nové jméno musí být vyplněno")
                if len(new_name) > 100: raise ValueError("Nové jméno je příliš dlouhé")
                if old_name == new_name:
                     flash("Nové jméno je stejné jako původní.", "info")
                     return redirect(url_for('manage_employees'))
                if not re.match(r"^[\w\s\-\.ěščřžýáíéúůďťňĚŠČŘŽÝÁÍÉÚŮĎŤŇ]+$", new_name):
                    raise ValueError("Nové jméno obsahuje nepovolené znaky.")

                if employee_manager_instance.upravit_zamestnance_podle_jmena(old_name, new_name):
                    flash(f'Zaměstnanec "{old_name}" byl upraven na "{new_name}".', "success")
                    return redirect(url_for('manage_employees'))
                else:
                    flash(f'Nepodařilo se upravit "{old_name}". Jméno neexistuje nebo nové jméno již existuje.', "error")

            elif action == "delete":
                employee_name = request.form.get("employee_name", "")
                if not employee_name: raise ValueError("Nebyl vybrán zaměstnanec k odstranění")

                if employee_manager_instance.smazat_zamestnance_podle_jmena(employee_name):
                    flash(f'Zaměstnanec "{employee_name}" byl smazán.', "success")
                    return redirect(url_for('manage_employees'))
                else:
                    flash(f'Nepodařilo se smazat "{employee_name}". Zaměstnanec neexistuje.', "error")

            else: raise ValueError(f"Neznámá akce: {action}")

        except ValueError as e:
            flash(str(e), "error")
            logger.warning(f"Chyba při správě zaměstnanců (akce: {action}): {e}")
        except Exception as e:
            flash("Došlo k neočekávané chybě při správě zaměstnanců.", "error")
            logger.error(f"Neočekávaná chyba při správě zaměstnanců (akce: {action}): {e}", exc_info=True)

    # Pro GET nebo po chybě v POST
    employees = employee_manager_instance.get_all_employees()
    return render_template("employees.html", employees=employees)


@app.route("/zaznam", methods=["GET", "POST"])
@require_excel_managers # Zajistí, že g.excel_manager existuje
def record_time():
    employee_manager_instance = g.employee_manager # Získáme z g
    excel_manager_instance = g.excel_manager       # Získáme z g

    selected_employees = employee_manager_instance.get_vybrani_zamestnanci()
    if not selected_employees:
        flash("Nejsou vybráni žádní zaměstnanci pro záznam.", "warning")
        return redirect(url_for("manage_employees"))

    settings = session.get('settings', {}) # Získáme aktuální nastavení ze session
    default_start_time = settings.get("start_time", "07:00")
    default_end_time = settings.get("end_time", "18:00")
    default_lunch_duration = settings.get("lunch_duration", 1.0)

    # Získáme hodnoty z formuláře nebo použijeme výchozí
    current_date = request.form.get("date", datetime.now().strftime("%Y-%m-%d"))
    start_time = request.form.get("start_time", default_start_time)
    end_time = request.form.get("end_time", default_end_time)
    lunch_duration_input = request.form.get("lunch_duration", str(default_lunch_duration))

    if request.method == "POST":
        try:
            # Validace (stejná jako dříve)
            # ... (kód pro validaci data, časů, pauzy) ...
            date_str = request.form.get("date", "")
            start_time_str = request.form.get("start_time", "")
            end_time_str = request.form.get("end_time", "")
            lunch_duration_str = request.form.get("lunch_duration", "")

            # Datum
            try:
                selected_date = datetime.strptime(date_str, "%Y-%m-%d").date()
                if selected_date > datetime.now().date(): raise ValueError("Nelze zadat budoucí datum")
                if selected_date.weekday() >= 5: raise ValueError("Nelze zadat víkend")
            except ValueError as e: raise ValueError(f"Neplatné datum: {e}")

            # Časy
            try:
                start = datetime.strptime(start_time_str, "%H:%M")
                end = datetime.strptime(end_time_str, "%H:%M")
                if end <= start: raise ValueError("Konec musí být po začátku")
            except ValueError as e: raise ValueError(f"Neplatný formát času: {e}")

            # Pauza
            try:
                lunch_duration = float(lunch_duration_str.replace(",", "."))
                if lunch_duration < 0: raise ValueError("Pauza nemůže být záporná")
                work_duration_hours = (end - start).total_seconds() / 3600
                if work_duration_hours > 0 and lunch_duration >= work_duration_hours:
                    raise ValueError("Pauza nemůže být delší než pracovní doba")
                if lunch_duration > 4: raise ValueError("Pauza nesmí být delší než 4 hodiny")
            except ValueError as e: raise ValueError(f"Neplatná délka pauzy: {e}")


            # Uložení pomocí excel_manager instance z 'g'
            success = excel_manager_instance.ulozit_pracovni_dobu(
                date_str, start_time_str, end_time_str, lunch_duration, selected_employees
            )

            if success:
                flash("Pracovní doba byla úspěšně zaznamenána.", "success")
                # Po úspěchu přesměrujeme, aby se formulář znovu neodeslal
                return redirect(url_for('index'))
            else:
                # Chyba byla zalogována v excel_manageru
                raise IOError("Nepodařilo se uložit záznam do Excel souboru.")

        except (ValueError, IOError) as e:
            flash(str(e), "error")
            logger.warning(f"Chyba při záznamu pracovní doby: {e}")
            # Hodnoty pro formulář zůstanou ty, které uživatel zadal
            current_date = request.form.get("date", current_date)
            start_time = request.form.get("start_time", start_time)
            end_time = request.form.get("end_time", end_time)
            lunch_duration_input = request.form.get("lunch_duration", lunch_duration_input)
        except Exception as e:
            flash("Došlo k neočekávané chybě při zpracování záznamu.", "error")
            logger.error(f"Neočekávaná chyba při záznamu pracovní doby: {e}", exc_info=True)
            # Vrátíme výchozí hodnoty

    # Formátování délky pauzy pro zobrazení
    try: lunch_duration_formatted = str(float(lunch_duration_input.replace(",", ".")))
    except ValueError: lunch_duration_formatted = str(default_lunch_duration)

    return render_template(
        "record_time.html",
        selected_employees=selected_employees,
        current_date=current_date,
        start_time=start_time,
        end_time=end_time,
        lunch_duration=lunch_duration_formatted,
    )


@app.route("/excel_viewer", methods=["GET"])
@require_excel_managers # Zajistí, že g.excel_manager existuje
def excel_viewer():
    """Zobrazí obsah aktivního Excel souboru."""
    excel_manager_instance = g.excel_manager
    active_file_path = excel_manager_instance.get_active_file_path()
    active_filename = active_file_path.name

    # Zobrazujeme vždy jen aktivní soubor
    excel_files = [active_filename]
    selected_file = active_filename # Neměnný

    active_sheet_name = request.args.get("sheet", None)
    workbook = None
    data = []
    sheet_names = []

    try:
        # Načtení workbooku (read-only)
        workbook = load_workbook(active_file_path, read_only=True, data_only=True)
        sheet_names = workbook.sheetnames

        if not sheet_names: raise ValueError("Aktivní soubor neobsahuje žádné listy.")

        # Výběr aktivního listu
        if active_sheet_name not in sheet_names:
             active_sheet_name = sheet_names[0] # Default na první list
        sheet = workbook[active_sheet_name]

        # Načtení dat (s limitem)
        MAX_ROWS_TO_DISPLAY = 500
        header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if header_row: data.append([str(c) if c is not None else "" for c in header_row])

        rows_loaded = 0
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if rows_loaded >= MAX_ROWS_TO_DISPLAY:
                flash(f"Zobrazeno prvních {MAX_ROWS_TO_DISPLAY} řádků dat.", "warning")
                break
            data.append([str(c) if c is not None else "" for c in row])
            rows_loaded += 1

        if not data: data.append([]) # Pro případ prázdného listu

    except (FileNotFoundError, ValueError, InvalidFileException, PermissionError) as e:
        logger.error(f"Chyba při zobrazování souboru '{active_filename}': {e}", exc_info=True)
        flash(f"Chyba při zobrazování souboru '{active_filename}': {e}", "error")
        return redirect(url_for("index"))
    except Exception as e:
        logger.error(f"Neočekávaná chyba při zobrazení Excel souboru '{active_filename}': {e}", exc_info=True)
        flash("Neočekávaná chyba při zobrazení Excel souboru.", "error")
        return redirect(url_for("index"))
    finally:
        if workbook: workbook.close()

    return render_template(
        "excel_viewer.html",
        excel_files=excel_files, # Obsahuje jen aktivní soubor
        selected_file=selected_file,
        sheet_names=sheet_names,
        active_sheet=active_sheet_name,
        data=data,
    )


@app.route("/settings", methods=["GET", "POST"])
@require_excel_managers # Zajistí g.excel_manager pro update_project_info
def settings_page():
    """Zobrazí a zpracuje nastavení aplikace."""
    excel_manager_instance = g.excel_manager # Získáme z g

    if request.method == "POST":
        # Získáme aktuální nastavení ze session pro případnou aktualizaci
        current_settings = session.get('settings', Config.get_default_settings())
        original_active_file = current_settings.get("active_excel_file")

        try:
            # Validace vstupů (stejná jako dříve)
            # ... (kód pro validaci časů, pauzy, názvu projektu, dat) ...
            start_time_str = request.form.get("start_time", "")
            end_time_str = request.form.get("end_time", "")
            lunch_duration_str = request.form.get("lunch_duration", "")
            project_name = request.form.get("project_name", "").strip()
            project_start_str = request.form.get("start_date", "")
            project_end_str = request.form.get("end_date", "") # Může být prázdné

            # Validace časů
            try:
                datetime.strptime(start_time_str, "%H:%M"); datetime.strptime(end_time_str, "%H:%M")
            except ValueError: raise ValueError("Neplatný formát času (HH:MM)")
            # Validace pauzy
            try:
                lunch_duration = float(lunch_duration_str.replace(",", "."))
                if lunch_duration < 0 or lunch_duration > 4: raise ValueError()
            except ValueError: raise ValueError("Neplatná délka pauzy (0-4)")
            # Validace projektu
            if not project_name: raise ValueError("Název projektu je povinný")
            if not project_start_str: raise ValueError("Datum začátku projektu je povinné")
            try: start_date = datetime.strptime(project_start_str, "%Y-%m-%d").date()
            except ValueError: raise ValueError("Neplatný formát data začátku (YYYY-MM-DD)")
            if project_end_str:
                try:
                    end_date = datetime.strptime(project_end_str, "%Y-%m-%d").date()
                    if end_date < start_date: raise ValueError("Konec projektu nemůže být před začátkem")
                except ValueError as e: raise ValueError(f"Neplatné datum konce: {e}")


            # Připravíme data pro uložení
            settings_to_save = current_settings.copy() # Pracujeme s kopií
            settings_to_save.update({
                "start_time": start_time_str,
                "end_time": end_time_str,
                "lunch_duration": lunch_duration,
                "project_info": {
                    "name": project_name,
                    "start_date": project_start_str,
                    "end_date": project_end_str,
                },
                # active_excel_file se zde nemění, mění se jen při archivaci
            })

            # Uložíme nastavení do souboru
            if not save_settings_to_file(settings_to_save):
                 raise RuntimeError("Nepodařilo se uložit nastavení do konfiguračního souboru.")

            # Aktualizujeme session
            session['settings'] = settings_to_save
            logger.info("Nastavení uložena do souboru a session.")

            # Aktualizujeme informace v aktivním Excel souboru
            excel_update_success = excel_manager_instance.update_project_info(
                project_name,
                project_start_str,
                project_end_str if project_end_str else None,
            )

            if excel_update_success:
                flash("Nastavení bylo úspěšně uloženo a informace v Excelu aktualizovány.", "success")
            else:
                # Chyba byla zalogována v excel_manageru
                flash("Nastavení bylo uloženo, ale nepodařilo se aktualizovat informace v Excel souboru.", "warning")

            return redirect(url_for("settings_page"))

        except (ValueError, RuntimeError) as e:
            flash(str(e), "error")
            logger.warning(f"Chyba při ukládání nastavení: {e}")
            # Zůstaneme na stránce, šablona zobrazí data z request.form nebo session
        except Exception as e:
            flash("Došlo k neočekávané chybě při ukládání nastavení.", "error")
            logger.error(f"Neočekávaná chyba při ukládání nastavení: {e}", exc_info=True)
            # Zůstaneme na stránce

    # Pro GET nebo po chybě v POST zobrazíme stránku s aktuálním nastavením ze session
    return render_template("settings_page.html", settings=session.get('settings', {}))


@app.route("/zalohy", methods=["GET", "POST"])
@require_excel_managers # Zajistí g.zalohy_manager a g.excel_manager
def zalohy():
    """Zpracuje přidání zálohy a zobrazí formulář."""
    employee_manager_instance = g.employee_manager
    zalohy_manager_instance = g.zalohy_manager
    excel_manager_instance = g.excel_manager

    employees_list = employee_manager_instance.zamestnanci
    # Získáme možnosti záloh z aktivního souboru
    advance_options = excel_manager_instance.get_advance_options()
    # Historie se již nenačítá
    advance_history = []

    if request.method == "POST":
        try:
            # Validace vstupů (stejná jako dříve)
            # ... (kód pro validaci jména, částky, měny, možnosti, data) ...
            employee_name = request.form.get("employee_name")
            amount_str = request.form.get("amount")
            currency = request.form.get("currency")
            option = request.form.get("option")
            date_str = request.form.get("date")

            if not employee_name or employee_name not in employees_list: raise ValueError("Vyberte platného zaměstnance")
            try: amount = float(amount_str.replace(",", ".")); zalohy_manager_instance.validate_amount(amount)
            except Exception as e: raise ValueError(f"Neplatná částka: {e}")
            zalohy_manager_instance.validate_currency(currency)
            if not option or option not in advance_options: raise ValueError("Vyberte platnou možnost")
            zalohy_manager_instance.validate_date(date_str)


            # Uložení zálohy pomocí zalohy_manager instance z 'g'
            # Ta již ví, do kterého (aktivního) souboru ukládat
            success = zalohy_manager_instance.add_or_update_employee_advance(
                employee_name=employee_name, amount=amount, currency=currency, option=option, date=date_str
            )

            if success:
                flash("Záloha byla úspěšně uložena.", "success")
                return redirect(url_for('zalohy'))
            else:
                raise RuntimeError("Nepodařilo se uložit zálohu. Zkontrolujte logy.")

        except (ValueError, RuntimeError) as e:
            flash(str(e), "error")
            logger.warning(f"Chyba při ukládání zálohy: {e}")
        except Exception as e:
            flash("Došlo k neočekávané chybě při ukládání zálohy.", "error")
            logger.error(f"Neočekávaná chyba při ukládání zálohy: {e}", exc_info=True)

    # Pro GET nebo po chybě v POST
    return render_template(
        "zalohy.html",
        employees=employees_list,
        options=advance_options,
        current_date=datetime.now().strftime("%Y-%m-%d"),
        advance_history=advance_history, # Historie je prázdná
    )

# --- Nová route pro archivaci a start nového souboru ---
@app.route("/start_new_file", methods=["POST"])
def start_new_file():
    """Vymaže název aktivního souboru v nastavení, čímž vynutí vytvoření nového."""
    try:
        # 1. Načteme aktuální nastavení
        settings = load_settings_from_file()
        current_active_file = settings.get("active_excel_file")

        if not current_active_file:
             flash("Již není nastaven žádný aktivní soubor. Nový bude vytvořen při příští akci.", "info")
             return redirect(url_for('settings_page'))

        # 2. Vymažeme název aktivního souboru
        settings["active_excel_file"] = None
        logger.info(f"Archivace souboru '{current_active_file}'. Aktivní soubor bude resetován.")

        # 3. Uložíme změněná nastavení
        if save_settings_to_file(settings):
             # 4. Aktualizujeme i session
             session['settings'] = settings
             flash(f"Soubor '{current_active_file}' byl archivován. Při příští akci (např. záznamu času) bude vytvořen nový soubor.", "success")
             # Můžeme zde případně vyčistit cache pro starý soubor, i když by se měla vyčistit na konci requestu
             excel_manager_instance = getattr(g, 'excel_manager', None)
             if excel_manager_instance and excel_manager_instance.active_filename == current_active_file:
                  excel_manager_instance.close_cached_workbooks() # Vyčistí celou cache

        else:
             # Pokud se nepodařilo uložit, vrátíme chybu
             flash("Chyba: Nepodařilo se uložit změnu nastavení pro archivaci.", "error")
             # Vrátíme původní aktivní soubor do session, aby nedošlo k nekonzistenci
             settings["active_excel_file"] = current_active_file
             session['settings'] = settings


    except Exception as e:
        logger.error(f"Neočekávaná chyba při archivaci a startu nového souboru: {e}", exc_info=True)
        flash("Došlo k neočekávané chybě při archivaci souboru.", "error")

    # Přesměrujeme zpět na stránku nastavení
    return redirect(url_for('settings_page'))


if __name__ == "__main__":
    # Nastavení logování pro vývoj
    if not app.debug:
         log_handler = logging.FileHandler('app_prod.log', encoding='utf-8')
         log_handler.setLevel(logging.WARNING)
         log_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
         log_handler.setFormatter(log_formatter)
         app.logger.addHandler(log_handler)
    else:
         app.logger.setLevel(logging.DEBUG) # Logujeme více v debug módu

    # Spuštění aplikace
    app.run(debug=True, host='0.0.0.0', port=5000)

