# app.py
import json
import logging
import os
import re
import smtplib
import ssl
from datetime import datetime, timedelta  # Přidán timedelta
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
# from email.utils import parseaddr  # nepoužito
from pathlib import Path
from functools import wraps

# Nepoužívané importy odstraněny
from flask import (Flask, flash, jsonify, redirect, render_template, request,
                   send_file, session, url_for, g, get_flashed_messages)
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook import Workbook

# Local imports
from config import Config
from employee_management import EmployeeManager
from excel_manager import ExcelManager
from utils.logger import setup_logger
from zalohy_manager import ZalohyManager
from utils.voice_processor import VoiceProcessor

# Nahrazení základního loggeru naším vlastním
logger = setup_logger("app")

# Initialize Flask app
app = Flask(__name__)
app.secret_key = Config.SECRET_KEY
Config.init_app(app)

# --- Funkce pro správu nastavení a aktivního souboru ---


def save_settings_to_file(settings_data):
    """Uloží slovník nastavení do JSON souboru."""
    try:
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
        save_settings_to_file(default_settings)
        return default_settings

    try:
        with open(Config.SETTINGS_FILE_PATH, "r", encoding="utf-8") as f:
            try:
                loaded_settings = json.load(f)
                if not isinstance(loaded_settings, dict):
                    raise ValueError("Neplatný formát JSON (očekáván slovník)")
                settings = default_settings.copy()
                settings.update(loaded_settings)
                if not isinstance(settings.get("project_info"), dict):
                    logger.warning(
                        "Klíč 'project_info' v nastavení není slovník, resetuji na výchozí."
                    )
                    settings["project_info"] = default_settings["project_info"]
                # Zajistíme, že active_excel_file je None nebo string
                if not isinstance(settings.get("active_excel_file"), (str, type(None))):
                    logger.warning(
                        "Klíč 'active_excel_file' má neplatný typ, resetuji na None."
                    )
                    settings["active_excel_file"] = None

                logger.info(
                    f"Nastavení úspěšně načtena ze souboru: {Config.SETTINGS_FILE_PATH}"
                )
                return settings
            except (json.JSONDecodeError, ValueError) as e:
                logger.error(
                    (
                        "Chyba při čtení nebo parsování souboru nastavení "
                        f"'{Config.SETTINGS_FILE_PATH}': {e}. Použijí se výchozí."
                    ),
                    exc_info=True,
                )
                save_settings_to_file(default_settings)
                return default_settings.copy()
    except Exception as e:
        logger.error(f"Neočekávaná chyba při načítání nastavení ze souboru: {e}", exc_info=True)
        return default_settings.copy()


def ensure_active_excel_file(settings):
    """
    Zkontroluje a případně vytvoří aktivní Excel soubor.
    Vrátí aktualizovaný slovník nastavení.
    """
    # Cíl: Ujistit se, že aktivní soubor je vždy pevně "Hodiny_Cap.xlsx"
    fixed_filename = Config.EXCEL_TEMPLATE_NAME  # Očekává se "Hodiny_Cap.xlsx"
    fixed_file_path = Config.EXCEL_BASE_PATH / fixed_filename

    # Pokud soubor neexistuje, pokusíme se ho vytvořit (init_app jej běžně vytváří)
    if not fixed_file_path.exists():
        try:
            Config.EXCEL_BASE_PATH.mkdir(parents=True, exist_ok=True)
            # Vytvoříme jednoduchý workbook s požadovanými listy
            wb = Workbook()
            if "Sheet" in wb.sheetnames:
                wb["Sheet"].title = Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME
            else:
                wb.create_sheet(Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME)
            if Config.EXCEL_ADVANCES_SHEET_NAME not in wb.sheetnames:
                wb.create_sheet(Config.EXCEL_ADVANCES_SHEET_NAME)
            wb.save(fixed_file_path)
            wb.close()
            logger.info(f"Vytvořen chybějící aktivní Excel soubor: {fixed_file_path}")
        except Exception as e:
            logger.error(f"Nepodařilo se vytvořit aktivní Excel soubor '{fixed_file_path}': {e}", exc_info=True)
            flash("Chyba při vytváření aktivního Excel souboru.", "error")
            # Necháme původní nastavení a vrátíme
            return settings

    # Pokud v nastavení není správný název, uložíme jej
    if settings.get("active_excel_file") != fixed_filename:
        settings["active_excel_file"] = fixed_filename
        if not save_settings_to_file(settings):
            logger.error("Nepodařilo se uložit nastavení s pevným názvem aktivního souboru.")
            flash("Nepodařilo se uložit nastavení aktivního souboru.", "error")

    return settings


# --- Request Handlers ---

@app.before_request
def before_request():
    """Spustí se před každým requestem."""
    settings = load_settings_from_file()
    settings = ensure_active_excel_file(settings)
    session['settings'] = settings  # Uložíme aktuální stav do session

    # Inicializace manažerů
    g.employee_manager = EmployeeManager(Config.DATA_PATH)
    active_filename = settings.get("active_excel_file")
    if active_filename:
        try:
            g.excel_manager = ExcelManager(
                Config.EXCEL_BASE_PATH, active_filename, Config.EXCEL_TEMPLATE_NAME
            )
            project_name = settings.get("project_info", {}).get("name")
            if project_name:
                g.excel_manager.set_project_name(project_name)
            g.zalohy_manager = ZalohyManager(Config.EXCEL_BASE_PATH, active_filename)
        except (ValueError, FileNotFoundError) as e:  # Chyby z __init__ manažerů
            logger.error(
                f"Chyba při inicializaci manažerů pro soubor '{active_filename}': {e}"
            )
            g.excel_manager = None
            g.zalohy_manager = None
            flash(
                f"Chyba při inicializaci pracovního souboru '{active_filename}'. Kontaktujte administrátora.",
                "error",
            )
        except Exception as e:
            logger.error(
                f"Neočekávaná chyba při inicializaci manažerů: {e}", exc_info=True
            )
            g.excel_manager = None
            g.zalohy_manager = None
            flash("Neočekávaná chyba při přípravě aplikace.", "error")
    else:
        g.excel_manager = None
        g.zalohy_manager = None
        # Flash zpráva by měla být přidána v ensure_active_excel_file


@app.teardown_request
def teardown_request(exception=None):
    """Spustí se po každém requestu."""
    excel_manager_instance = getattr(g, 'excel_manager', None)
    if excel_manager_instance:
        try:
            excel_manager_instance.close_cached_workbooks()
            logger.debug("Cache workbooků vyčištěna na konci requestu.")
        except Exception as e:
            logger.error(f"Chyba při čištění cache workbooků na konci requestu: {e}", exc_info=True)

    g.pop('employee_manager', None)
    g.pop('excel_manager', None)
    g.pop('zalohy_manager', None)


# --- Dekorátor pro kontrolu existence manažerů ---
def require_excel_managers(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not getattr(g, 'excel_manager', None) or not getattr(g, 'zalohy_manager', None):
            logger.error(f"Přístup k route '{request.path}' zamítnut, Excel manažeři nejsou inicializováni.")
            # Přidáme flash zprávu, pokud ještě není
            if not any(
                'Chyba: Není definován aktivní Excel soubor' in msg[1]
                for msg in get_flashed_messages(with_categories=True)
            ):
                flash(
                    (
                        "Chyba: Není definován aktivní Excel soubor pro práci. "
                        "Zkuste archivovat a začít nový nebo kontaktujte administrátora."
                    ),
                    "error",
                )
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function


# --- Routes ---

@app.route("/")
def index():
    settings = session.get('settings', {})
    active_filename = settings.get('active_excel_file')
    excel_exists = False
    week_num_int = 0
    current_date = datetime.now().strftime("%Y-%m-%d")  # Datum vždy zobrazíme

    if active_filename:
        active_file_path = Config.EXCEL_BASE_PATH / active_filename
        excel_exists = active_file_path.exists()
        excel_manager_instance = getattr(g, 'excel_manager', None)
        if excel_manager_instance:
            week_calendar_data = excel_manager_instance.ziskej_cislo_tydne(current_date)
            week_num_int = week_calendar_data.week if week_calendar_data else 0
        else:
            logger.warning(
                "ExcelManager není k dispozici pro získání čísla týdne v index route."
            )
            # Můžeme zobrazit 0 nebo se pokusit vypočítat týden přímo zde
            try:
                week_num_int = (
                    datetime.strptime(current_date, "%Y-%m-%d").isocalendar().week
                )
            except Exception:
                pass  # Ignorujeme chybu, zůstane 0
    # else: # Flash zpráva je přidána v before_request nebo ensure_active_excel_file

    return render_template(
        "index.html",
        excel_exists=excel_exists,
        active_filename=active_filename,
        week_number=week_num_int,
        current_date=current_date
    )


@app.route("/download")
@require_excel_managers
def download_file():
    """
    Stáhne aktivní Excel soubor s názvem Hodiny_Cap_Tyden<MaxWeekNumber>.xlsx.
    """
    workbook = None
    try:
        active_file_path = g.excel_manager.get_active_file_path()

        # Načtení workbooku (read-only)
        try:
            workbook = load_workbook(active_file_path, read_only=True)
            sheet_names = workbook.sheetnames
        except Exception as e:
            logger.error(f"Nepodařilo se načíst '{active_file_path.name}' pro čtení listů: {e}", exc_info=True)
            raise ValueError(f"Nepodařilo se otevřít soubor '{active_file_path.name}'.")
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
                    max_week_number = max(
                        max_week_number, int(match.group(1))
                    )
                except ValueError:
                    continue

        # Vytvoření názvu pro stažení - Použijeme název šablony jako základ
        template_stem = Path(Config.EXCEL_TEMPLATE_NAME).stem  # Např. "Hodiny_Cap"
        if max_week_number > 0:
            download_filename = f"{template_stem}_Tyden_{max_week_number}.xlsx"
        else:
            # Pokud nejsou týdny, stáhne se jako "Hodiny_Cap.xlsx"
            download_filename = Config.EXCEL_TEMPLATE_NAME
            logger.warning(
                (
                    f"V souboru '{active_file_path.name}' nenalezen žádný list "
                    f"'Týden X', stahuje se jako '{download_filename}'."
                )
            )

        # Odeslání aktivního souboru ke stažení s novým názvem
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
@require_excel_managers
def send_email():
    """Odešle aktivní Excel soubor emailem."""
    try:
        active_file_path = g.excel_manager.get_active_file_path()
        active_filename = active_file_path.name

        # Kontrola konfigurace a validace emailů (stejná jako dříve)
        recipient = Config.RECIPIENT_EMAIL or ""
        if not recipient:
            raise ValueError("E-mail příjemce není nastaven.")
        sender = Config.SMTP_USERNAME or ""
        if not sender or not Config.SMTP_PASSWORD:
            raise ValueError("SMTP údaje nejsou nastaveny.")
        if not Config.SMTP_SERVER or not Config.SMTP_PORT:
            raise ValueError("SMTP server/port není nastaven.")
        if not validate_email(sender) or not validate_email(recipient):
            raise ValueError("Neplatná e-mailová adresa odesílatele nebo příjemce.")

        # Vytvoření zprávy (stejné jako dříve)
        msg = MIMEMultipart()
        subject = f'{active_filename} - {datetime.now().strftime("%Y-%m-%d %H:%M")}'
        msg["Subject"] = subject
        msg["From"] = sender
        msg["To"] = recipient
        app_name = Config.DEFAULT_APP_NAME
        body = (
            f"Dobrý den,\n\n"
            f"v příloze zasílám aktuální výkaz pracovní doby ({active_filename}).\n\n"
            f"S pozdravem,\n{app_name}"
        )
        msg.attach(MIMEText(body, "plain", "utf-8"))

        # Přidání přílohy (stejné jako dříve)
        try:
            with open(active_file_path, "rb") as f:
                attachment = MIMEApplication(
                    f.read(), _subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                attachment.add_header(
                    "Content-Disposition", "attachment", filename=active_filename
                )
                msg.attach(attachment)
        except IOError as e:
            logger.error(
                f"Chyba při čtení souboru '{active_filename}' pro email: {e}",
                exc_info=True,
            )
            raise ValueError(
                f"Nepodařilo se připojit soubor '{active_filename}' k emailu."
            )

        # Odeslání emailu (stejné jako dříve)
        ssl_context = ssl.create_default_context()
        with smtplib.SMTP_SSL(
            Config.SMTP_SERVER, Config.SMTP_PORT, context=ssl_context, timeout=Config.SMTP_TIMEOUT
        ) as smtp:
            smtp.login(sender, Config.SMTP_PASSWORD)
            smtp.send_message(msg)

        flash("Email byl úspěšně odeslán.", "success")
        logger.info(f"Email s výkazem '{active_filename}' odeslán na {recipient}")

    except (FileNotFoundError, ValueError) as e:
        logger.error(f"Chyba dat nebo souboru pro odeslání emailu: {e}")
        flash(str(e), "error")
    except (ConnectionError, IOError, smtplib.SMTPException, ssl.SSLError, TimeoutError) as e:
        logger.error(f"Chyba připojení/SMTP při odesílání emailu: {e}", exc_info=True)
        flash(f"Chyba při odesílání emailu: {e}", "error")
    except Exception as e:
        logger.error(f"Neočekávaná chyba v procesu odesílání emailu: {e}", exc_info=True)
        flash("Neočekávaná chyba při odesílání emailu.", "error")

    return redirect(url_for("index"))


@app.route("/zamestnanci", methods=["GET", "POST"])
def manage_employees():
    # Použije g.employee_manager (kontrola není nutná, pokud neprovádí kritické operace)
    employee_manager_instance = getattr(g, 'employee_manager', None)
    if not employee_manager_instance:
        flash("Správce zaměstnanců není k dispozici.", "error")
        return redirect(url_for('index'))

    if request.method == "POST":
        action = request.form.get("action")
        try:
            # Zpracování akcí (stejné jako dříve)
            # ... (kód pro add, select, edit, delete) ...
            if not action:
                raise ValueError("Nebyla specifikována akce")

            if action == "add":
                employee_name = request.form.get("name", "").strip()
                if not employee_name:
                    raise ValueError("Jméno zaměstnance nemůže být prázdné")
                if len(employee_name) > Config.EMPLOYEE_NAME_MAX_LENGTH:
                    raise ValueError("Jméno zaměstnance je příliš dlouhé")
                if not re.match(Config.EMPLOYEE_NAME_VALIDATION_REGEX, employee_name):
                    raise ValueError("Jméno zaměstnance obsahuje nepovolené znaky.")
                if employee_manager_instance.pridat_zamestnance(employee_name):
                    flash(f'Zaměstnanec "{employee_name}" byl přidán.', "success")
                    return redirect(url_for('manage_employees'))
                else:
                    flash(
                        f'Zaměstnanec "{employee_name}" již existuje nebo došlo k chybě.',
                        "error",
                    )

            elif action == "select":
                employee_name = request.form.get("employee_name", "")
                if not employee_name:
                    raise ValueError("Nebyl vybrán zaměstnanec")
                if employee_name not in employee_manager_instance.zamestnanci:
                    raise ValueError(f'Zaměstnanec "{employee_name}" neexistuje')
                if employee_name in employee_manager_instance.vybrani_zamestnanci:
                    if employee_manager_instance.odebrat_vybraneho_zamestnance(employee_name):
                        flash(f'"{employee_name}" odebrán z výběru.', "success")
                    else:
                        flash(
                            f'Nepodařilo se odebrat "{employee_name}" z výběru.',
                            "error",
                        )
                else:
                    if employee_manager_instance.pridat_vybraneho_zamestnance(employee_name):
                        flash(f'"{employee_name}" přidán do výběru.', "success")
                    else:
                        flash(
                            f'Nepodařilo se přidat "{employee_name}" do výběru.',
                            "error",
                        )
                return redirect(url_for('manage_employees'))

            elif action == "edit":
                old_name = request.form.get("old_name", "").strip()
                new_name = request.form.get("new_name", "").strip()
                if not old_name or not new_name:
                    raise ValueError("Původní i nové jméno musí být vyplněno")
                if len(new_name) > Config.EMPLOYEE_NAME_MAX_LENGTH:
                    raise ValueError("Nové jméno je příliš dlouhé")
                if old_name == new_name:
                    flash("Jména jsou stejná.", "info")
                    return redirect(url_for('manage_employees'))
                if not re.match(Config.EMPLOYEE_NAME_VALIDATION_REGEX, new_name):
                    raise ValueError("Nové jméno obsahuje nepovolené znaky.")
                if employee_manager_instance.upravit_zamestnance_podle_jmena(old_name, new_name):
                    flash(f'"{old_name}" upraven na "{new_name}".', "success")
                    return redirect(url_for('manage_employees'))
                else:
                    flash(f'Nepodařilo se upravit "{old_name}".', "error")

            elif action == "delete":
                employee_name = request.form.get("employee_name", "")
                if not employee_name:
                    raise ValueError("Nebyl vybrán zaměstnanec k odstranění")
                if employee_manager_instance.smazat_zamestnance_podle_jmena(employee_name):
                    flash(f'Zaměstnanec "{employee_name}" byl smazán.', "success")
                    return redirect(url_for('manage_employees'))
                else:
                    flash(f'Nepodařilo se smazat "{employee_name}".', "error")

            else:
                raise ValueError(f"Neznámá akce: {action}")

        except ValueError as e:
            flash(str(e), "error")
            logger.warning(f"Chyba při správě zaměstnanců (akce: {action}): {e}")
        except Exception as e:
            flash("Neočekávaná chyba při správě zaměstnanců.", "error")
            logger.error(f"Neočekávaná chyba při správě zaměstnanců (akce: {action}): {e}", exc_info=True)

    employees = employee_manager_instance.get_all_employees()
    return render_template("employees.html", employees=employees)


@app.route("/zaznam", methods=["GET", "POST"])
@require_excel_managers
def record_time():
    employee_manager_instance = g.employee_manager
    excel_manager_instance = g.excel_manager

    selected_employees = employee_manager_instance.get_vybrani_zamestnanci()
    if not selected_employees:
        flash("Nejsou vybráni žádní zaměstnanci pro záznam.", "warning")
        return redirect(url_for("manage_employees"))

    settings = session.get('settings', {})
    default_start_time = settings.get("start_time", "07:00")
    default_end_time = settings.get("end_time", "18:00")
    default_lunch_duration = settings.get("lunch_duration", 1.0)

    # Získáme datum z URL parametru 'next_date' nebo použijeme dnešek
    default_date_str = request.args.get('next_date', datetime.now().strftime("%Y-%m-%d"))

    # Získáme hodnoty z formuláře (pro případ chyby) nebo použijeme výchozí/předané
    current_date = request.form.get("date", default_date_str)
    start_time = request.form.get("start_time", default_start_time)
    end_time = request.form.get("end_time", default_end_time)
    lunch_duration_input = request.form.get("lunch_duration", str(default_lunch_duration))
    is_free_day = request.form.get("is_free_day") == "on"  # Získáme stav checkboxu

    if request.method == "POST":
        # Zpracování formuláře pro záznam času
        try:
            date_str = request.form.get("date", "")
            is_free_day_submitted = request.form.get("is_free_day") == "on"

            # Validace data (stejná jako dříve)
            try:
                selected_date = datetime.strptime(date_str, "%Y-%m-%d").date()
                if selected_date > datetime.now().date():
                    raise ValueError("Nelze zadat budoucí datum")
            except ValueError as e:
                raise ValueError(f"Neplatné datum: {e}")

            if is_free_day_submitted:
                # Záznam volného dne
                start_time_str = "00:00"
                end_time_str = "00:00"
                lunch_duration = 0.0
                logger.info(f"Zaznamenává se volný den pro datum {date_str}")
            else:
                # Záznam pracovní doby - validace časů a pauzy
                start_time_str = request.form.get("start_time", "")
                end_time_str = request.form.get("end_time", "")
                lunch_duration_str = request.form.get("lunch_duration", "")
                try:
                    start = datetime.strptime(start_time_str, "%H:%M")
                    end = datetime.strptime(end_time_str, "%H:%M")
                    if end <= start:
                        raise ValueError("Konec musí být po začátku")
                except ValueError as e:
                    raise ValueError(f"Neplatný formát času: {e}")
                try:
                    lunch_duration = float(lunch_duration_str.replace(",", "."))
                    if lunch_duration < 0:
                        raise ValueError("Pauza nemůže být záporná")
                    work_duration_hours = (end - start).total_seconds() / 3600
                    if work_duration_hours > 0 and lunch_duration >= work_duration_hours:
                        raise ValueError("Pauza nemůže být delší než pracovní doba")
                    if lunch_duration > 4:
                        raise ValueError("Pauza nesmí být delší než 4 hodiny")
                except ValueError as e:
                    raise ValueError(f"Neplatná délka pauzy: {e}")

            # Uložení do Excelu
            success = excel_manager_instance.ulozit_pracovni_dobu(
                date_str, start_time_str, end_time_str, lunch_duration, selected_employees
            )

            if success:
                flash("Záznam byl úspěšně uložen.", "success")
                next_day = selected_date + timedelta(days=1)
                while next_day.weekday() >= 5:
                    next_day += timedelta(days=1)
                next_date_str = next_day.strftime("%Y-%m-%d")
                return redirect(url_for('record_time', next_date=next_date_str))
            else:
                raise IOError("Nepodařilo se uložit záznam do Excel souboru.")

        except (ValueError, IOError) as e:
            flash(str(e), "error")
            logger.warning(f"Chyba při záznamu pracovní doby/volna: {e}")
            # Hodnoty pro formulář zůstanou ty, které uživatel zadal
            current_date = request.form.get("date", current_date)
            start_time = request.form.get("start_time", start_time)
            end_time = request.form.get("end_time", end_time)
            lunch_duration_input = request.form.get("lunch_duration", lunch_duration_input)
            is_free_day = request.form.get("is_free_day") == "on"
        except Exception as e:
            flash("Došlo k neočekávané chybě při zpracování záznamu.", "error")
            logger.error(f"Neočekávaná chyba při záznamu pracovní doby/volna: {e}", exc_info=True)

    # Získání seznamu Excel souborů
    excel_files = []
    try:
        excel_files = sorted([f.name for f in Config.EXCEL_BASE_PATH.glob('*.xlsx')], reverse=True)
    except Exception as e:
        flash("Nepodařilo se načíst seznam Excel souborů.", "error")
        logger.error(f"Chyba při načítání seznamu souborů z {Config.EXCEL_BASE_PATH}: {e}")

    active_excel_file = settings.get("active_excel_file")

    # Formátování délky pauzy pro zobrazení
    try:
        lunch_duration_formatted = str(float(lunch_duration_input.replace(",", ".")))
    except ValueError:
        lunch_duration_formatted = str(default_lunch_duration)

    return render_template(
        "record_time.html",
        selected_employees=selected_employees,
        current_date=current_date,
        start_time=start_time,
        end_time=end_time,
        lunch_duration=lunch_duration_formatted,
        is_free_day=is_free_day,
        excel_files=excel_files,
        active_excel_file=active_excel_file
    )


@app.route("/set_active_file", methods=["POST"])
def set_active_file():
    """Nastaví aktivní Excel soubor pro zápis."""
    selected_file = request.form.get("excel_file")
    if not selected_file:
        flash("Nebyl vybrán žádný soubor.", "error")
        return redirect(url_for('record_time'))

    file_path = Config.EXCEL_BASE_PATH / selected_file
    if not file_path.exists():
        flash(f"Vybraný soubor '{selected_file}' neexistuje.", "error")
        return redirect(url_for('record_time'))

    try:
        settings = load_settings_from_file()
        settings["active_excel_file"] = selected_file
        if save_settings_to_file(settings):
            session['settings'] = settings  # Aktualizujeme session
            flash(f"Aktivní soubor byl nastaven na '{selected_file}'.", "success")
            logger.info(f"Aktivní soubor změněn na: {selected_file}")
        else:
            flash("Nepodařilo se uložit nastavení aktivního souboru.", "error")
    except Exception as e:
        flash("Došlo k chybě při nastavování aktivního souboru.", "error")
        logger.error(f"Chyba při nastavování aktivního souboru na '{selected_file}': {e}", exc_info=True)

    return redirect(url_for('record_time'))


@app.route("/rename_project", methods=["POST"])
def rename_project():
    """Přejmenuje existující Excel soubor (projekt)."""
    old_filename = request.form.get("old_excel_file")
    new_filename = request.form.get("new_excel_file")

    if not old_filename or not new_filename:
        flash("Starý a nový název souboru musí být vyplněn.", "error")
        return redirect(url_for('settings_page'))

    if not new_filename.endswith(".xlsx"):
        new_filename += ".xlsx"

    old_path = Config.EXCEL_BASE_PATH / old_filename
    new_path = Config.EXCEL_BASE_PATH / new_filename

    if not old_path.exists():
        flash(f"Soubor '{old_filename}' neexistuje.", "error")
        return redirect(url_for('settings_page'))

    if new_path.exists():
        flash(f"Soubor s názvem '{new_filename}' již existuje. Zvolte jiný název.", "error")
        return redirect(url_for('settings_page'))

    try:
        os.rename(old_path, new_path)
        logger.info(f"Soubor '{old_filename}' přejmenován na '{new_filename}'.")

        settings = load_settings_from_file()
        if settings.get("active_excel_file") == old_filename:
            settings["active_excel_file"] = new_filename
            save_settings_to_file(settings)
            session['settings'] = settings
            logger.info(f"Aktivní soubor v nastavení aktualizován na '{new_filename}'.")

        flash(f"Soubor '{old_filename}' byl úspěšně přejmenován na '{new_filename}'.", "success")
    except Exception as e:
        flash(f"Chyba při přejmenování souboru: {e}", "error")
        logger.error(f"Chyba při přejmenování souboru '{old_filename}' na '{new_filename}': {e}", exc_info=True)

    return redirect(url_for('settings_page'))


@app.route("/delete_project", methods=["POST"])
def delete_project():
    """Smaže existující Excel soubor (projekt)."""
    filename_to_delete = request.form.get("excel_file_to_delete")

    if not filename_to_delete:
        flash("Nebyl vybrán žádný soubor ke smazání.", "error")
        return redirect(url_for('settings_page'))

    file_path = Config.EXCEL_BASE_PATH / filename_to_delete

    if not file_path.exists():
        flash(f"Soubor '{filename_to_delete}' neexistuje.", "error")
        return redirect(url_for('settings_page'))

    try:
        settings = load_settings_from_file()
        if settings.get("active_excel_file") == filename_to_delete:
            flash(
                (
                    f"Nelze smazat aktivní soubor '{filename_to_delete}'. "
                    "Nejprve archivujte nebo vyberte jiný aktivní soubor."
                ),
                "error"
            )
            return redirect(url_for('settings_page'))

        os.remove(file_path)
        logger.info(f"Soubor '{filename_to_delete}' byl smazán.")
        flash(f"Soubor '{filename_to_delete}' byl úspěšně smazán.", "success")
    except Exception as e:
        flash(f"Chyba při mazání souboru: {e}", "error")
        logger.error(f"Chyba při mazání souboru '{filename_to_delete}': {e}", exc_info=True)

    return redirect(url_for('settings_page'))


@app.route("/excel_viewer", methods=["GET"])
@require_excel_managers
def excel_viewer():
    """Zobrazí obsah aktivního Excel souboru."""
    excel_manager_instance = g.excel_manager
    active_file_path = excel_manager_instance.get_active_file_path()
    active_filename = active_file_path.name

    excel_files = [active_filename]  # Jen aktivní soubor
    selected_file = active_filename

    active_sheet_name = request.args.get("sheet", None)
    workbook = None
    data = []
    sheet_names = []

    try:
        # Načtení workbooku a dat (stejné jako dříve)
        workbook = load_workbook(active_file_path, read_only=True, data_only=True)
        sheet_names = workbook.sheetnames
        if not sheet_names:
            raise ValueError("Soubor neobsahuje listy.")
        if active_sheet_name not in sheet_names:
            active_sheet_name = sheet_names[0]
        sheet = workbook[active_sheet_name]

        MAX_ROWS_TO_DISPLAY = Config.MAX_ROWS_TO_DISPLAY_EXCEL_VIEWER
        header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if header_row:
            data.append([str(c) if c is not None else "" for c in header_row])
        rows_loaded = 0
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if rows_loaded >= MAX_ROWS_TO_DISPLAY:
                flash(
                    f"Zobrazeno prvních {MAX_ROWS_TO_DISPLAY} řádků dat.", "warning"
                )
                break
            data.append([str(c) if c is not None else "" for c in row])
            rows_loaded += 1
        if not data:
            data.append([])

    except (FileNotFoundError, ValueError, InvalidFileException, PermissionError) as e:
        logger.error(f"Chyba při zobrazování '{active_filename}': {e}", exc_info=True)
        flash(f"Chyba při zobrazování souboru '{active_filename}': {e}", "error")
        return redirect(url_for("index"))
    except Exception as e:
        logger.error(f"Neočekávaná chyba při zobrazení '{active_filename}': {e}", exc_info=True)
        flash("Neočekávaná chyba při zobrazení Excel souboru.", "error")
        return redirect(url_for("index"))
    finally:
        if workbook:
            workbook.close()

    return render_template(
        "excel_viewer.html",
        excel_files=excel_files,
        selected_file=selected_file,
        sheet_names=sheet_names,
        active_sheet=active_sheet_name,
        data=data,
    )


@app.route("/settings", methods=["GET", "POST"])
@require_excel_managers  # Potřebujeme pro update_project_info
def settings_page():
    """Zobrazí a zpracuje nastavení aplikace."""
    excel_manager_instance = g.excel_manager

    if request.method == "POST":
        current_settings = session.get('settings', Config.get_default_settings())
        try:
            # Validace vstupů
            start_time_str = request.form.get("start_time", "")
            end_time_str = request.form.get("end_time", "")
            lunch_duration_str = request.form.get("lunch_duration", "")
            project_name = request.form.get("project_name", "").strip()
            project_start_str = request.form.get("start_date", "")
            project_end_str = request.form.get("end_date", "")  # Nepovinné zde

            # Validace (stejná jako dříve, ale end_date není required)
            try:
                datetime.strptime(start_time_str, "%H:%M")
                datetime.strptime(end_time_str, "%H:%M")
            except ValueError:
                raise ValueError("Neplatný formát času (HH:MM)")
            try:
                lunch_duration = float(lunch_duration_str.replace(",", "."))
                if not (0 <= lunch_duration <= 4):
                    raise ValueError()
            except ValueError:
                raise ValueError("Neplatná délka pauzy (0-4)")
            if not project_name:
                raise ValueError("Název projektu je povinný")
            if not project_start_str:
                raise ValueError("Datum začátku projektu je povinné")
            try:
                start_date = datetime.strptime(project_start_str, "%Y-%m-%d").date()
            except ValueError:
                raise ValueError("Neplatný formát data začátku (YYYY-MM-DD)")
            # Validujeme end_date pouze pokud je zadáno
            if project_end_str:
                try:
                    end_date = datetime.strptime(project_end_str, "%Y-%m-%d").date()
                    if end_date < start_date:
                        raise ValueError("Konec projektu nemůže být před začátkem")
                except ValueError as e:
                    raise ValueError(f"Neplatné datum konce: {e}")

            # Uložení nastavení
            settings_to_save = current_settings.copy()
            settings_to_save.update({
                "start_time": start_time_str, "end_time": end_time_str,
                "lunch_duration": lunch_duration,
                "project_info": {
                    "name": project_name, "start_date": project_start_str,
                    "end_date": project_end_str,  # Uložíme i prázdný string
                },
            })

            if not save_settings_to_file(settings_to_save):
                raise RuntimeError("Nepodařilo se uložit nastavení do souboru.")
            session['settings'] = settings_to_save  # Aktualizujeme session
            logger.info("Nastavení uložena do souboru a session.")

            # Aktualizace Excelu
            excel_update_success = excel_manager_instance.update_project_info(
                project_name, project_start_str,
                project_end_str if project_end_str else None,
            )
            if excel_update_success:
                flash("Nastavení bylo úspěšně uloženo.", "success")
            else:
                flash("Nastavení uloženo, ale nepodařilo se aktualizovat Excel.", "warning")

            return redirect(url_for("settings_page"))

        except (ValueError, RuntimeError) as e:
            flash(str(e), "error")
            logger.warning(f"Chyba při ukládání nastavení: {e}")
        except Exception as e:
            flash("Neočekávaná chyba při ukládání nastavení.", "error")
            logger.error(f"Neočekávaná chyba při ukládání nastavení: {e}", exc_info=True)

    return render_template("settings_page.html", settings=session.get('settings', {}))


@app.route("/zalohy", methods=["GET", "POST"])
@require_excel_managers
def zalohy():
    """Zpracuje přidání zálohy."""
    employee_manager_instance = g.employee_manager
    zalohy_manager_instance = g.zalohy_manager
    excel_manager_instance = g.excel_manager

    employees_list = employee_manager_instance.zamestnanci
    advance_options = excel_manager_instance.get_advance_options()
    advance_history = []  # Historie se nenačítá

    if request.method == "POST":
        try:
            # Validace vstupů (stejná jako dříve)
            employee_name = request.form.get("employee_name")
            amount_str = request.form.get("amount")
            currency = request.form.get("currency")
            option = request.form.get("option")
            date_str = request.form.get("date")

            if not employee_name or employee_name not in employees_list:
                raise ValueError("Vyberte platného zaměstnance")
            amount_str = request.form.get("amount", "")
            try:
                amount = float(amount_str.replace(",", "."))
                zalohy_manager_instance.validate_amount(amount)
            except Exception as e:
                raise ValueError(f"Neplatná částka: {e}")
            zalohy_manager_instance.validate_currency(currency)
            if not option or option not in advance_options:
                raise ValueError("Vyberte platnou možnost")
            zalohy_manager_instance.validate_date(date_str)

            # Uložení zálohy
            success = zalohy_manager_instance.add_or_update_employee_advance(
                employee_name=employee_name,
                amount=amount,
                currency=currency,
                option=option,
                date=date_str
            )
            if success:
                flash("Záloha byla úspěšně uložena.", "success")
                return redirect(url_for('zalohy'))
            else:
                raise RuntimeError("Nepodařilo se uložit zálohu.")

        except (ValueError, RuntimeError) as e:
            flash(str(e), "error")
            logger.warning(f"Chyba při ukládání zálohy: {e}")
        except Exception as e:
            flash("Neočekávaná chyba při ukládání zálohy.", "error")
            logger.error(f"Neočekávaná chyba při ukládání zálohy: {e}", exc_info=True)

    return render_template(
        "zalohy.html",
        employees=employees_list, options=advance_options,
        current_date=datetime.now().strftime("%Y-%m-%d"),
        advance_history=advance_history,
    )


@app.route("/start_new_file", methods=["POST"])
def start_new_file():
    """Archivuje aktuální soubor (resetuje active_excel_file v nastavení)."""
    try:
        settings = load_settings_from_file()
        current_active_file = settings.get("active_excel_file")
        project_info = settings.get("project_info", {})
        project_end_str = project_info.get("end_date")
        project_start_str = project_info.get("start_date")

        if not current_active_file:
            flash("Již není nastaven žádný aktivní soubor.", "info")
            return redirect(url_for('settings_page'))

        # Validace: Konec projektu musí být zadán a platný PŘED archivací
        if not project_end_str:
            raise ValueError("Před archivací souboru musí být zadáno datum konce projektu v nastavení.")
        try:
            end_date = datetime.strptime(project_end_str, "%Y-%m-%d").date()
            if project_start_str:  # Pokud máme i start date, zkontrolujeme pořadí
                start_date = datetime.strptime(project_start_str, "%Y-%m-%d").date()
                if end_date < start_date:
                    raise ValueError(
                        "Datum konce projektu nemůže být dřívější než datum začátku."
                    )
        except ValueError as e:
            # Zobrazíme specifickou chybu z validace data
            raise ValueError(f"Neplatné datum konce projektu pro archivaci: {e}")

        # Resetujeme aktivní soubor v nastavení
        settings["active_excel_file"] = None
        logger.info(f"Archivace souboru '{current_active_file}'. Aktivní soubor bude resetován.")

        # Uložíme změněná nastavení
        if save_settings_to_file(settings):
            session['settings'] = settings  # Aktualizujeme session
            flash(
                f"Soubor '{current_active_file}' byl archivován. Při příští akci bude vytvořen nový.",
                "success",
            )
            # Vyčistíme cache pro starý soubor
            excel_manager_instance = getattr(g, 'excel_manager', None)
            if (
                excel_manager_instance
                and excel_manager_instance.active_filename == current_active_file
            ):
                excel_manager_instance.close_cached_workbooks()
        else:
            flash("Chyba: Nepodařilo se uložit změnu nastavení pro archivaci.", "error")
            # Vrátíme původní aktivní soubor do session
            settings["active_excel_file"] = current_active_file
            session['settings'] = settings

    except ValueError as e:  # Zachytíme validační chyby (např. chybějící end_date)
        flash(str(e), "error")
        logger.warning(f"Chyba validace při archivaci: {e}")
    except Exception as e:
        logger.error(f"Neočekávaná chyba při archivaci souboru: {e}", exc_info=True)
        flash("Došlo k neočekávané chybě při archivaci souboru.", "error")

    return redirect(url_for('settings_page'))


@app.route('/voice-command', methods=['POST'])
def voice_command():
    """Zpracování hlasového příkazu"""
    try:
        data = request.get_json()
        if not data or 'command' not in data:
            return jsonify({'success': False, 'error': 'Chybí hlasový příkaz'})

        voice_processor = VoiceProcessor()
        result = voice_processor.process_voice_text(data['command'])

        if not result['success']:
            return jsonify(result)

        # Podle typu akce vykonáme odpovídající operaci
        entities = result['entities']
        excel_manager = g.excel_manager
        employee_manager = g.employee_manager

        if entities['action'] == 'record_time':
            # Získáme všechny vybrané zaměstnance
            selected_employees = employee_manager.get_vybrani_zamestnanci()
            if not selected_employees:
                return jsonify({
                    'success': False,
                    'error': 'Nejsou vybráni žádní zaměstnanci'
                })

            # Pokud je to volný den, použijeme speciální hodnoty
            if entities.get('is_free_day'):
                entities['start_time'] = "00:00"
                entities['end_time'] = "00:00"
                entities['lunch_duration'] = 0.0

            # Záznam pracovní doby nebo volna pro všechny vybrané zaměstnance
            success, message = excel_manager.record_time(
                employee=selected_employees,
                date=entities['date'],
                start_time=entities['start_time'],
                end_time=entities['end_time'],
                lunch_duration=entities.get('lunch_duration', 1.0)
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


@app.route('/monthly_report', methods=['GET', 'POST'])
@require_excel_managers
def monthly_report_route():
    employee_manager_instance = g.employee_manager
    excel_manager_instance = g.excel_manager

    all_employees_data = employee_manager_instance.get_all_employees()
    employee_names = [emp['name'] for emp in all_employees_data]

    report_data = None

    if request.method == 'POST':
        try:
            selected_month = request.form.get('month', type=int)
            selected_year = request.form.get('year', type=int)
            # getlist pro případ, že by frontend posílal vícekrát stejný název parametru
            # Pokud chceme, aby 'employees' byl vždy seznam, i když je vybrán jen jeden nebo žádný,
            # getlist je správná volba. Pokud by 'employees' mohlo chybět, zvážili bychom default.
            selected_employees = request.form.getlist('employees')
            if not selected_employees:  # Pokud je seznam prázdný (žádný zaměstnanec nebyl vybrán)
                selected_employees = None  # Nastavíme na None pro metodu generate_monthly_report

            # Validace vstupů
            if selected_month is None or not (1 <= selected_month <= 12):
                flash('Neplatný měsíc. Zadejte hodnotu od 1 do 12.', 'error')
                # Znovu vykreslíme s původními GET hodnotami nebo aktuálními, pokud POST selhal na začátku
                return render_template("monthly_report.html",
                                       employee_names=employee_names,
                                       current_month=datetime.now().month,
                                       current_year=datetime.now().year,
                                       report_data=None,
                                       # Předáme zpět vybrané
                                       selected_employees_post=request.form.getlist('employees'))

            if selected_year is None or not (2000 <= selected_year <= 2100):  # Rozumný rozsah pro rok
                flash('Neplatný rok. Zadejte hodnotu např. mezi 2000 a 2100.', 'error')
                return render_template("monthly_report.html",
                                       employee_names=employee_names,
                                       current_month=selected_month,  # Použijeme již zadaný měsíc
                                       current_year=datetime.now().year,  # Rok můžeme resetovat nebo ponechat
                                       report_data=None,
                                       selected_employees_post=request.form.getlist('employees'))

            logger.info(f"Generování měsíčního reportu pro {selected_month}/{selected_year}. "
                        f"Zaměstnanci: {selected_employees}")

            report_data = excel_manager_instance.generate_monthly_report(
                month=selected_month,
                year=selected_year,
                employees=selected_employees
            )

            if not report_data:
                flash('Nebyly nalezeny žádné záznamy pro zadané období a zaměstnance.', 'info')

            # Vykreslení šablony s výsledky (nebo prázdnými daty, pokud nic nebylo nalezeno)
            return render_template("monthly_report.html",
                                   employee_names=employee_names,
                                   current_month=selected_month,
                                   current_year=selected_year,
                                   report_data=report_data,
                                   # Předáme vybrané pro zachování stavu
                                   selected_employees_post=selected_employees if selected_employees else [])

        except ValueError as e:  # Chyby z generate_monthly_report nebo validace typů
            flash(str(e), 'error')
            logger.warning(f"Chyba hodnoty při generování měsíčního reportu: {e}")
            report_data = None
        except IOError as e:  # Chyby souboru z generate_monthly_report
            flash(str(e), 'error')
            logger.error(f"Chyba souboru při generování měsíčního reportu: {e}", exc_info=True)
            report_data = None
        except Exception as e:
            logger.error(f"Neočekávaná chyba při generování měsíčního reportu: {e}", exc_info=True)
            flash('Došlo k neočekávané chybě při generování reportu.', 'error')
            report_data = None

        # Pokud došlo k chybě, znovu vykreslíme formulář s chybovou hláškou
        # a zachováme co nejvíce zadaných hodnot
        return render_template("monthly_report.html",
                               employee_names=employee_names,
                               current_month=request.form.get('month', datetime.now().month, type=int),
                               current_year=request.form.get('year', datetime.now().year, type=int),
                               report_data=report_data,  # Bude None
                               selected_employees_post=request.form.getlist('employees'))

    else:  # GET request
        current_month = datetime.now().month
        current_year = datetime.now().year
        return render_template("monthly_report.html",
                               employee_names=employee_names,
                               current_month=current_month,
                               current_year=current_year,
                               report_data=None,
                               selected_employees_post=[])  # Pro GET je seznam vybraných prázdný


def validate_email(email):
    """Validuje e-mailovou adresu."""
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return re.match(pattern, email) is not None


if __name__ == "__main__":
    if not app.debug:
        log_handler = logging.FileHandler('app_prod.log', encoding='utf-8')
        log_handler.setLevel(logging.WARNING)
        log_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        log_handler.setFormatter(log_formatter)
        app.logger.addHandler(log_handler)
    else:
        app.logger.setLevel(logging.DEBUG)
    app.run(debug=True, host='0.0.0.0', port=5000)
