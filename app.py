import json
import logging
import os
import re # Import pro regulární výrazy
import shutil # Import pro kopírování souborů
import smtplib
import ssl
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import parseaddr
from pathlib import Path

import openpyxl # Import pro práci s Excel soubory
import pandas as pd
from flask import Flask, flash, jsonify, redirect, render_template, request, send_file, session, url_for
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
    if "settings" not in session:
        default_settings = Config.get_default_settings()
        try:
            if SETTINGS_FILE_PATH.exists():
                with open(SETTINGS_FILE_PATH, "r", encoding="utf-8") as f:
                    try:
                        saved_settings = json.load(f)
                        # Validace načtených nastavení
                        if not isinstance(saved_settings, dict):
                            raise ValueError("Neplatný formát nastavení")
                        default_settings.update(saved_settings)
                    except json.JSONDecodeError:
                        logger.error("Poškozený soubor s nastaveními")
                        flash("Soubor s nastaveními je poškozen, používám výchozí nastavení.", "warning")
            session["settings"] = default_settings
        except Exception as e:
            logger.error(f"Chyba při načítání nastavení: {e}")
            session["settings"] = default_settings
    return session["settings"]


@app.route("/")
def index():
    try:
        # Použijeme Path pro konzistentní práci s cestami
        excel_path = Path(excel_manager.file_path)
        excel_exists = excel_path.exists()

        # Pokud soubor neexistuje, pokusíme se ho vytvořit prázdný
        if not excel_exists:
            try:
                # Ujistíme se, že adresář existuje
                excel_path.parent.mkdir(parents=True, exist_ok=True)
                # Vytvoříme nový workbook
                wb = Workbook()
                # Přidáme výchozí list
                if "Sheet" in wb.sheetnames:
                    sheet = wb["Sheet"]
                    sheet.title = "Týden"
                else:
                    sheet = wb.create_sheet("Týden")
                # Uložíme workbook
                wb.save(str(excel_path))
                excel_exists = True
                logger.info(f"Vytvořen nový Excel soubor: {excel_path}")
            except Exception as e:
                logger.error(f"Nepodařilo se vytvořit Excel soubor: {e}")

        # Kontrola existence druhého Excel souboru
        excel_path_2025 = Path(excel_manager.file_path_2025)
        if not excel_path_2025.exists():
            try:
                excel_path_2025.parent.mkdir(parents=True, exist_ok=True)
                wb = Workbook()
                # Přidáme výchozí listy
                if "Sheet" in wb.sheetnames:
                    sheet = wb["Sheet"]
                    sheet.title = "Zalohy25"
                wb.create_sheet("(pp)cash25")
                wb.save(str(excel_path_2025))
                logger.info(f"Vytvořen nový Excel soubor: {excel_path_2025}")
            except Exception as e:
                logger.error(f"Nepodařilo se vytvořit Excel soubor Hodiny2025.xlsx: {e}")

        current_date = datetime.now().strftime("%Y-%m-%d")
        week_number = excel_manager.ziskej_cislo_tydne(current_date)
        # Získání čísla týdne z objektu vráceného isocalendar()
        week_num_int = week_number.week if hasattr(week_number, 'week') else 0

        return render_template(
            "index.html", excel_exists=excel_exists, week_number=week_num_int, current_date=current_date
        )
    except Exception as e:
        logger.error(f"Chyba při načítání hlavní stránky: {e}")
        flash("Došlo k neočekávané chybě při načítání stránky.", "error")
        return render_template(
            "index.html", excel_exists=False, week_number=0, current_date=datetime.now().strftime("%Y-%m-%d")
        )


@app.route("/download")
def download_file():
    """
    Zpracuje požadavek na stažení souboru.
    Najde nejvyšší číslo týdne v názvech listů souboru Hodiny_Cap.xlsx,
    vytvoří kopii souboru s názvem obsahujícím toto číslo týdne,
    uloží kopii na server a nabídne ji ke stažení uživateli.
    """
    workbook = None # Inicializace workbook proměnné
    try:
        original_file_path = Path(excel_manager.file_path)
        if not original_file_path.exists():
            raise FileNotFoundError("Původní Excel soubor (Hodiny_Cap.xlsx) nebyl nalezen")

        # Načtení workbooku pro zjištění názvů listů
        try:
            workbook = load_workbook(original_file_path, read_only=True)
            sheet_names = workbook.sheetnames
        except Exception as e:
            logger.error(f"Nepodařilo se načíst Excel soubor pro čtení listů: {e}")
            raise ValueError("Nepodařilo se otevřít Excel soubor pro analýzu.")
        finally:
             if workbook:
                workbook.close() # Zajistí uzavření souboru

        # Nalezení nejvyššího čísla týdne
        max_week_number = 0
        week_pattern = re.compile(r"Týden (\d+)") # Regulární výraz pro "Týden X"

        for sheet_name in sheet_names:
            match = week_pattern.match(sheet_name)
            if match:
                week_num = int(match.group(1))
                if week_num > max_week_number:
                    max_week_number = week_num

        # Vytvoření názvu pro kopii
        if max_week_number > 0:
            new_filename = f"Hodiny_Cap_Tyden{max_week_number}.xlsx"
        else:
            # Pokud se nenašel žádný list "Týden X", použije se původní název
            # nebo můžete nastavit jinou výchozí logiku
            new_filename = EXCEL_FILE_NAME
            logger.warning("Nenalezen žádný list ve formátu 'Týden X', stahuje se původní soubor.")
            # V tomto případě bychom mohli rovnou poslat originál, ale pro konzistenci
            # vytvoříme kopii i s původním názvem.
            # Alternativně: return send_file(str(original_file_path), as_attachment=True)

        # Cesta pro novou kopii
        new_file_path = EXCEL_BASE_PATH / new_filename

        # Vytvoření kopie souboru na serveru
        try:
            shutil.copy2(original_file_path, new_file_path) # copy2 zachovává metadata
            logger.info(f"Vytvořena kopie souboru na serveru: {new_file_path}")
        except Exception as e:
            logger.error(f"Nepodařilo se zkopírovat soubor: {e}")
            raise IOError("Chyba při vytváření kopie souboru na serveru.")

        # Odeslání kopie souboru ke stažení
        return send_file(
            str(new_file_path),
            as_attachment=True,
            download_name=new_filename # Zajistí správný název při stahování
        )

    except FileNotFoundError as e:
        logger.error(f"Soubor nebyl nalezen: {e}")
        flash(str(e), "error") # Zobrazí specifickou chybovou hlášku
        return redirect(url_for("index"))
    except (ValueError, IOError) as e: # Zachytí chyby při načítání nebo kopírování
        logger.error(f"Chyba při zpracování souboru: {e}")
        flash(str(e), "error")
        return redirect(url_for("index"))
    except Exception as e:
        logger.error(f"Neočekávaná chyba při stahování souboru: {e}")
        flash("Chyba při stahování souboru.", "error")
        return redirect(url_for("index"))


def validate_email(email):
    """Validace emailové adresy"""
    if not email or "@" not in parseaddr(email)[1]:
        raise ValueError("Neplatná emailová adresa")
    return True


@app.route("/send_email", methods=["POST"])
def send_email():
    try:
        # Kontrola existence souboru
        file_path = Path(excel_manager.file_path)
        if not file_path.exists():
            raise FileNotFoundError("Excel soubor nebyl nalezen")

        # Kontrola konfigurace
        if not Config.RECIPIENT_EMAIL:
            raise ValueError("E-mailová adresa příjemce není nastavena v konfiguraci")

        if not Config.SMTP_USERNAME or not Config.SMTP_PASSWORD:
            raise ValueError("SMTP přihlašovací údaje nejsou nastaveny v konfiguraci")

        if not Config.SMTP_SERVER or not Config.SMTP_PORT:
            raise ValueError("SMTP server nebo port není nastaven v konfiguraci")

        # Validace emailových adres
        sender = Config.SMTP_USERNAME
        recipient = Config.RECIPIENT_EMAIL

        if not all([validate_email(addr) for addr in [sender, recipient]]):
            raise ValueError("Neplatné emailové adresy")

        # Vytvoření zprávy
        msg = MIMEMultipart()
        msg["Subject"] = f'Hodiny_Cap.xlsx - {datetime.now().strftime("%Y-%m-%d")}'
        msg["From"] = sender
        msg["To"] = recipient

        # Přidání textu zprávy
        body = f"""Dobrý den,

v příloze zasílám aktuální výkaz pracovní doby ({EXCEL_FILE_NAME}).

S pozdravem
{Config.APP_NAME if hasattr(Config, 'APP_NAME') else 'Evidence pracovní doby'}
"""
        msg.attach(MIMEText(body, "plain", "utf-8"))

        # Přidání přílohy
        with open(file_path, "rb") as f:
            attachment = MIMEApplication(f.read(), _subtype="xlsx")
            attachment.add_header("Content-Disposition", "attachment", filename=("utf-8", "", EXCEL_FILE_NAME))
            msg.attach(attachment)

        # Nastavení SSL/TLS kontextu s vyšším zabezpečením
        ssl_context = ssl.create_default_context()
        ssl_context.minimum_version = ssl.TLSVersion.TLSv1_2
        ssl_context.verify_mode = ssl.CERT_REQUIRED

        # Odeslání emailu
        with smtplib.SMTP_SSL(Config.SMTP_SERVER, Config.SMTP_PORT, context=ssl_context, timeout=30) as smtp:
            smtp.login(Config.SMTP_USERNAME, Config.SMTP_PASSWORD)
            smtp.send_message(msg)

        flash("Email byl úspěšně odeslán.", "success")
        logger.info(f"Email s výkazem byl úspěšně odeslán na adresu {recipient}")

    except FileNotFoundError as e:
        logger.error(f"Soubor nebyl nalezen: {e}")
        flash("Excel soubor nebyl nalezen.", "error")
    except ValueError as e:
        logger.error(f"Chyba konfigurace: {e}")
        flash(str(e), "error")
    except smtplib.SMTPAuthenticationError:
        logger.error("Chyba při přihlášení k SMTP serveru")
        flash("Nesprávné přihlašovací údaje k e-mailovému serveru.", "error")
    except smtplib.SMTPException as e:
        logger.error(f"SMTP chyba: {e}")
        flash("Chyba při komunikaci s e-mailovým serverem.", "error")
    except ssl.SSLError as e:
        logger.error(f"SSL chyba: {e}")
        flash("Chyba zabezpečeného spojení s e-mailovým serverem.", "error")
    except Exception as e:
        logger.error(f"Neočekávaná chyba při odesílání emailu: {e}")
        flash("Chyba při odesílání emailu.", "error")

    return redirect(url_for("index"))


@app.route("/zamestnanci", methods=["GET", "POST"])
def manage_employees():
    if request.method == "POST":
        try:
            action = request.form.get("action")
            if not action:
                raise ValueError("Nebyla specifikována akce")

            if action == "add":
                employee_name = request.form.get("name", "").strip()
                if not employee_name:
                    raise ValueError("Jméno zaměstnance nemůže být prázdné")
                if len(employee_name) > 100:
                    raise ValueError("Jméno zaměstnance je příliš dlouhé")
                # Povolíme i jiné znaky než alfanumerické, např. diakritiku
                if not re.match(r"^[\w\s\-\.ěščřžýáíéúůďťňĚŠČŘŽÝÁÍÉÚŮĎŤŇ]+$", employee_name):
                     raise ValueError("Jméno zaměstnance obsahuje nepovolené znaky")


                if employee_manager.pridat_zamestnance(employee_name):
                    flash(f'Zaměstnanec "{employee_name}" byl přidán.', "success")
                else:
                    flash(f'Zaměstnanec "{employee_name}" už existuje.', "error")

            elif action == "select":
                employee_name = request.form.get("employee_name", "")
                if not employee_name:
                    raise ValueError("Nebyl vybrán zaměstnanec")

                if employee_name in employee_manager.zamestnanci:
                    if employee_name in employee_manager.vybrani_zamestnanci:
                        employee_manager.odebrat_vybraneho_zamestnance(employee_name)
                        flash(f'Zaměstnanec "{employee_name}" byl odebrán z výběru.', "success")
                    else:
                        employee_manager.pridat_vybraneho_zamestnance(employee_name)
                        flash(f'Zaměstnanec "{employee_name}" byl přidán do výběru.', "success")
                    employee_manager.save_config()
                else:
                    raise ValueError(f'Zaměstnanec "{employee_name}" neexistuje')

            elif action == "edit":
                old_name = request.form.get("old_name", "").strip()
                new_name = request.form.get("new_name", "").strip()

                if not old_name or not new_name:
                    raise ValueError("Původní i nové jméno musí být vyplněno")
                if len(new_name) > 100:
                    raise ValueError("Nové jméno je příliš dlouhé")
                # Povolíme i jiné znaky než alfanumerické
                if not re.match(r"^[\w\s\-\.ěščřžýáíéúůďťňĚŠČŘŽÝÁÍÉÚŮĎŤŇ]+$", new_name):
                    raise ValueError("Nové jméno obsahuje nepovolené znaky")


                try:
                    # Hledání indexu podle jména (bezpečnější než předpokládat pořadí)
                    # idx = employee_manager.zamestnanci.index(old_name) + 1 # Starý způsob
                    if employee_manager.upravit_zamestnance_podle_jmena(old_name, new_name): # Nový způsob
                        flash(f'Zaměstnanec "{old_name}" byl upraven na "{new_name}".', "success")
                    else:
                        # Chyba je již logována v metodě
                        raise ValueError(f'Nepodařilo se upravit zaměstnance "{old_name}" (možná neexistuje nebo nové jméno již existuje)')
                except ValueError as e:
                    flash(str(e), "error") # Zobrazíme chybu uživateli
                    # Není potřeba znovu logovat, loguje se v EmployeeManager

            elif action == "delete":
                employee_name = request.form.get("employee_name", "")
                if not employee_name:
                    raise ValueError("Nebyl vybrán zaměstnanec k odstranění")

                try:
                    # Mazání podle jména
                    if employee_manager.smazat_zamestnance_podle_jmena(employee_name): # Nový způsob
                        flash(f'Zaměstnanec "{employee_name}" byl smazán.', "success")
                    else:
                        # Chyba je již logována v metodě
                         raise ValueError(f'Nepodařilo se smazat zaměstnance "{employee_name}" (možná neexistuje)')
                except ValueError as e:
                    flash(str(e), "error") # Zobrazíme chybu uživateli
                    # Není potřeba znovu logovat, loguje se v EmployeeManager

            else:
                raise ValueError("Neplatná akce")

        except ValueError as e:
            flash(str(e), "error")
            logger.error(f"Chyba při správě zaměstnanců: {e}")
        except Exception as e:
            flash("Došlo k neočekávané chybě při správě zaměstnanců.", "error")
            logger.error(f"Neočekávaná chyba při správě zaměstnanců: {e}", exc_info=True) # Přidáno exc_info pro detailnější logování

    # Převedení seznamů na formát očekávaný šablonou
    employees = employee_manager.get_all_employees() # Použijeme metodu get_all_employees
    return render_template("employees.html", employees=employees)


@app.route("/zaznam", methods=["GET", "POST"])
def record_time():
    selected_employees = employee_manager.vybrani_zamestnanci
    if not selected_employees:
        flash("Nejsou vybráni žádní zaměstnanci.", "warning")
        return redirect(url_for("manage_employees"))

    current_date = datetime.now().strftime("%Y-%m-%d")
    settings = get_settings()
    start_time = settings.get("start_time", "07:00")
    end_time = settings.get("end_time", "18:00")
    lunch_duration = settings.get("lunch_duration", 1.0) # Zajistíme float

    if request.method == "POST":
        try:
            # Validace data
            date = request.form.get("date", "")
            try:
                selected_date = datetime.strptime(date, "%Y-%m-%d").date()
                # Přidána kontrola, zda není víkend
                if selected_date.weekday() >= 5: # 5 = Sobota, 6 = Neděle
                    raise ValueError("Nelze zaznamenat pracovní dobu na víkend")
            except ValueError as e:
                 if "víkend" in str(e):
                     raise
                 raise ValueError("Neplatný formát data (použijte YYYY-MM-DD)")


            # Validace časů
            start_time_str = request.form.get("start_time", "")
            end_time_str = request.form.get("end_time", "")
            try:
                start = datetime.strptime(start_time_str, "%H:%M")
                end = datetime.strptime(end_time_str, "%H:%M")
                if end <= start:
                    raise ValueError("Čas konce musí být později než čas začátku")
            except ValueError as e:
                if "musí být později" in str(e):
                    raise
                raise ValueError("Neplatný formát času (použijte HH:MM)")

            # Validace délky pauzy
            lunch_duration_str = request.form.get("lunch_duration", "")
            try:
                lunch_duration = float(lunch_duration_str.replace(",", "."))
                if lunch_duration < 0:
                    raise ValueError("Délka pauzy nemůže být záporná")
                work_hours = (end - start).total_seconds() / 3600 # Použijeme total_seconds pro přesnější výpočet
                if lunch_duration >= work_hours and work_hours > 0: # Povolíme nulovou pauzu, pokud je pracovní doba kladná
                    raise ValueError("Délka pauzy nemůže být delší nebo rovna pracovní době")
            except ValueError as e:
                if any(msg in str(e) for msg in ["nemůže být záporná", "nemůže být delší"]):
                    raise
                raise ValueError("Délka pauzy musí být číslo")

            # Uložení do Hodiny_Cap.xlsx
            try:
                # Předáváme validované časy a délku pauzy
                excel_manager.ulozit_pracovni_dobu(date, start_time_str, end_time_str, lunch_duration, selected_employees)
                flash("Pracovní doba byla úspěšně zaznamenána.", "success")
            except Exception as e:
                logger.error(f"Chyba při ukládání do Excel souboru: {e}", exc_info=True)
                # Zobrazíme obecnější chybu, detaily jsou v logu
                raise ValueError("Nepodařilo se uložit pracovní dobu do Excel souboru. Zkontrolujte logy pro více informací.")

        except ValueError as e:
            flash(str(e), "error")
            logger.error(f"Chyba validace pracovní doby: {e}")
        except Exception as e:
            flash("Došlo k neočekávané chybě při ukládání pracovní doby.", "error")
            logger.error(f"Neočekávaná chyba při ukládání pracovní doby: {e}", exc_info=True)

    # Předáme formátovanou délku pauzy do šablony
    lunch_duration_formatted = str(lunch_duration).replace('.', ',')

    return render_template(
        "record_time.html",
        selected_employees=selected_employees,
        current_date=current_date,
        start_time=start_time,
        end_time=end_time,
        lunch_duration=lunch_duration_formatted, # Použijeme formátovanou hodnotu
    )


@app.route("/excel_viewer", methods=["GET"])
def excel_viewer():
    excel_files = ["Hodiny_Cap.xlsx", "Hodiny2025.xlsx"]
    selected_file = request.args.get("file", excel_files[0])
    active_sheet = request.args.get("sheet", None)
    workbook = None
    data = [] # Inicializace dat

    try:
        # Určení cesty k souboru
        if selected_file == "Hodiny_Cap.xlsx":
            file_path = excel_manager.file_path
        elif selected_file == "Hodiny2025.xlsx":
            file_path = EXCEL_BASE_PATH / EXCEL_FILE_NAME_2025
        else:
            # Pokud by byl přidán další soubor, je třeba ho zde ošetřit
            raise ValueError(f"Neznámý nebo nepodporovaný soubor: {selected_file}")

        # Kontrola existence souboru - konverze na Path objekt pro jednotnou práci
        file_path = Path(file_path)
        if not file_path.exists():
            raise FileNotFoundError(f"Soubor {selected_file} nebyl nalezen")

        # Načtení workbooku v read-only módu pro úsporu paměti
        workbook = load_workbook(file_path, read_only=True, data_only=True)

        if not workbook.sheetnames:
            raise ValueError("Excel soubor neobsahuje žádné listy")

        # Výběr aktivního listu
        # Pokud active_sheet není specifikován nebo neexistuje, vybereme první list
        if active_sheet not in workbook.sheetnames:
             active_sheet = workbook.sheetnames[0]
        sheet = workbook[active_sheet]

        # Načtení dat s omezením počtu řádků pro prevenci přetížení paměti
        MAX_ROWS = 1000
        # data = [] # Přesunuto na začátek try bloku
        # Načtení hlavičky (první řádek)
        header = [str(cell.value) if cell.value is not None else "" for cell in sheet[1]]
        data.append(header)

        # Načtení zbytku dat
        for i, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1): # Začínáme od druhého řádku
            if i >= MAX_ROWS:
                flash(f"Zobrazeno prvních {MAX_ROWS} řádků dat (včetně hlavičky).", "warning")
                break
            # Převod všech buněk na string, None hodnoty na prázdný řetězec
            data.append([str(cell) if cell is not None else "" for cell in row])

        # Pokud nejsou žádná data (ani hlavička), přidáme prázdný řádek, aby šablona neselhala
        if not data:
             data.append([])


    except FileNotFoundError as e:
        logger.error(f"Soubor nebyl nalezen: {e}")
        flash(f"Požadovaný Excel soubor '{selected_file}' nebyl nalezen.", "error")
        # V případě nenalezení souboru přesměrujeme nebo zobrazíme chybovou stránku
        return redirect(url_for("index")) # Nebo render_template s chybovou zprávou
    except InvalidFileException:
        logger.error(f"Soubor {selected_file} je poškozen nebo má neplatný formát")
        flash(f"Soubor '{selected_file}' je poškozen nebo má neplatný formát.", "error")
        return redirect(url_for("index"))
    except ValueError as e:
        logger.error(f"Chyba při práci s Excel souborem: {e}")
        flash(str(e), "error")
        return redirect(url_for("index"))
    except PermissionError:
        logger.error(f"Nedostatečná oprávnění pro čtení souboru {selected_file}")
        flash(f"Nedostatečná oprávnění pro čtení souboru '{selected_file}'.", "error")
        return redirect(url_for("index"))
    except Exception as e:
        logger.error(f"Neočekávaná chyba při zobrazení Excel souboru: {e}", exc_info=True)
        flash("Chyba při zobrazení Excel souboru.", "error")
        return redirect(url_for("index"))
    finally:
        if workbook:
            workbook.close() # Zajistíme uzavření workbooku

    # Předání dat do šablony
    return render_template(
        "excel_viewer.html",
        excel_files=excel_files,
        selected_file=selected_file,
        sheet_names=workbook.sheetnames if workbook else [], # Předáme sheetnames jen pokud workbook existuje
        active_sheet=active_sheet,
        data=data, # Předáme načtená data
    )


@app.route("/settings", methods=["GET", "POST"])
def settings_page():
    """Handle settings page"""
    if request.method == "POST":
        try:
            # Validace vstupních dat
            start_time_str = request.form.get("start_time", "")
            end_time_str = request.form.get("end_time", "")
            lunch_duration_str = request.form.get("lunch_duration", "")

            # Validace času
            try:
                datetime.strptime(start_time_str, "%H:%M")
                datetime.strptime(end_time_str, "%H:%M")
            except ValueError:
                raise ValueError("Neplatný formát času (použijte HH:MM)")

            # Validace délky pauzy
            try:
                lunch_duration = float(lunch_duration_str.replace(",", "."))
                if lunch_duration < 0:
                    # Původní chyba byla ValueError, ale pro srozumitelnost ji specifikujeme
                    raise ValueError("Délka pauzy nemůže být záporná")
            except ValueError as e:
                 if "záporná" in str(e):
                     raise
                 raise ValueError("Délka pauzy musí být nezáporné číslo")


            # Validace dat projektu
            project_name = request.form.get("project_name", "").strip()
            project_start_str = request.form.get("start_date", "")
            project_end_str = request.form.get("end_date", "")

            # Název projektu je povinný
            if not project_name:
                 raise ValueError("Název projektu je povinný")


            start_date = None
            end_date = None

            # Validace datumu začátku (povinný)
            if not project_start_str:
                 raise ValueError("Datum začátku projektu je povinné")
            try:
                start_date = datetime.strptime(project_start_str, "%Y-%m-%d").date()
            except ValueError:
                raise ValueError("Neplatný formát data začátku projektu (použijte YYYY-MM-DD)")

            # Validace datumu konce (nepovinný, ale pokud je zadán, musí být platný a ne dřívější než začátek)
            if project_end_str:
                try:
                    end_date = datetime.strptime(project_end_str, "%Y-%m-%d").date()
                    if end_date < start_date:
                        raise ValueError("Datum konce projektu nemůže být dřívější než datum začátku")
                except ValueError as e:
                    if "dřívější" in str(e):
                        raise
                    raise ValueError("Neplatný formát data konce projektu (použijte YYYY-MM-DD)")

            # Aktualizace nastavení v session a souboru
            current_settings = get_settings() # Získáme aktuální nastavení (ze session nebo souboru)
            current_settings.update(
                {
                    "start_time": start_time_str,
                    "end_time": end_time_str,
                    "lunch_duration": lunch_duration,
                    "project_info": {
                        "name": project_name,
                        "start_date": project_start_str, # Ukládáme string
                        "end_date": project_end_str,     # Ukládáme string
                    },
                }
            )

            # Vytvoření adresáře pro nastavení, pokud neexistuje
            SETTINGS_FILE_PATH.parent.mkdir(parents=True, exist_ok=True)

            # Uložení nastavení do JSON souboru
            try:
                with open(SETTINGS_FILE_PATH, "w", encoding="utf-8") as f:
                    json.dump(current_settings, f, indent=4, ensure_ascii=False)
                session["settings"] = current_settings # Aktualizujeme i session
            except IOError as e:
                 logger.error(f"Chyba při zápisu do souboru nastavení: {e}", exc_info=True)
                 raise ValueError("Nepodařilo se uložit nastavení do souboru.")


            # Aktualizace informací o projektu v Excel souboru
            try:
                # Předáváme stringy datumu
                excel_manager.update_project_info(
                    project_name,
                    project_start_str,
                    project_end_str if project_end_str else None, # Předáme None, pokud není konec zadán
                )
            except Exception as e:
                 logger.error(f"Chyba při aktualizaci Excel souboru s informacemi o projektu: {e}", exc_info=True)
                 # I když Excel selže, nastavení v JSONu a session jsou uložena
                 flash("Nastavení bylo uloženo, ale nepodařilo se aktualizovat informace v Excel souboru.", "warning")
                 return redirect(url_for("settings_page")) # Zůstaneme na stránce


            flash("Nastavení bylo úspěšně uloženo.", "success")
            return redirect(url_for("settings_page")) # Přesměrujeme po úspěšném uložení

        except ValueError as e:
            flash(str(e), "error")
            logger.error(f"Chyba validace nastavení: {e}")
            # Necháme uživatele na stránce, aby mohl opravit chybu
        except Exception as e:
            flash("Došlo k neočekávané chybě při ukládání nastavení.", "error")
            logger.error(f"Neočekávaná chyba při ukládání nastavení: {e}", exc_info=True)
            # Necháme uživatele na stránce

    # Zobrazíme stránku s aktuálním nastavením (buď původním nebo neúspěšně upraveným)
    return render_template("settings_page.html", settings=get_settings())


@app.route("/zalohy", methods=["GET", "POST"])
def zalohy():
    # Inicializace ZalohyManager s cestou k adresáři excel souborů
    zalohy_manager = ZalohyManager(EXCEL_BASE_PATH)
    advance_history = []
    # Získání jmen zaměstnanců pro formulář
    employees_list = employee_manager.zamestnanci
    # Získání možností záloh pro formulář
    advance_options = excel_manager.get_advance_options()

    if request.method == "POST":
        try:
            # Validace vstupních dat z formuláře
            employee_name = request.form.get("employee_name")
            if not employee_name or employee_name not in employees_list: # Kontrola, zda je zaměstnanec platný
                raise ValueError("Vyberte platného zaměstnance")

            amount_str = request.form.get("amount")
            try:
                amount = float(amount_str.replace(",", "."))
                # Použití validační metody z ZalohyManager
                zalohy_manager.validate_amount(amount)
            except (ValueError, TypeError, AttributeError) as e:
                 # Poskytneme srozumitelnější chybu
                 raise ValueError(f"Neplatná částka: {e}. Zadejte kladné číslo.")


            currency = request.form.get("currency")
            # Použití validační metody z ZalohyManager
            zalohy_manager.validate_currency(currency)


            option = request.form.get("option")
            # Validace, zda je option jednou z načtených možností
            if not option or option not in advance_options:
                raise ValueError("Vyberte platnou možnost zálohy")
            # Použití validační metody z ZalohyManager (pro jistotu, i když by mělo být pokryto výše)
            # zalohy_manager.validate_option(option) # Tato validace kontroluje proti ['Option 1', 'Option 2'], což nemusí odpovídat načteným možnostem


            date_str = request.form.get("date")
            # Použití validační metody z ZalohyManager
            zalohy_manager.validate_date(date_str)


            # Uložení zálohy pomocí ZalohyManager
            zalohy_manager.add_or_update_employee_advance(
                employee_name=employee_name, amount=amount, currency=currency, option=option, date=date_str
            )
            flash("Záloha byla úspěšně uložena.", "success")
            # Přesměrování po úspěšném uložení, aby se zabránilo opětovnému odeslání formuláře
            return redirect(url_for('zalohy'))


        except ValueError as e:
            flash(str(e), "error")
            logger.error(f"Chyba validace zálohy: {e}")
        except Exception as e:
            flash("Došlo k neočekávané chybě při ukládání zálohy.", "error")
            logger.error(f"Neočekávaná chyba při ukládání zálohy: {e}", exc_info=True)

    # Načtení historie záloh (pro GET požadavek nebo po neúspěšném POST)
    workbook_2025 = None # Inicializace
    try:
        hodiny2025_path = EXCEL_BASE_PATH / EXCEL_FILE_NAME_2025
        if not hodiny2025_path.exists():
            logger.warning(f"Soubor {EXCEL_FILE_NAME_2025} pro historii záloh nebyl nalezen")
            # Můžeme zobrazit varování, ale stránka se stále načte
            flash(f"Soubor '{EXCEL_FILE_NAME_2025}' se záznamy historie záloh nebyl nalezen.", "warning")
        else:
            try:
                 workbook_2025 = load_workbook(hodiny2025_path, read_only=True, data_only=True)
                 if "Zalohy25" in workbook_2025.sheetnames:
                     sheet = workbook_2025["Zalohy25"]
                     # Načteme data bezpečněji
                     data_iter = sheet.values
                     try:
                         # Předpokládáme, že první řádek je hlavička
                         header = next(data_iter)
                         keys = [str(k) if k is not None else f"col_{i}" for i, k in enumerate(header)]

                         # Zpracujeme zbývající řádky
                         for row in data_iter:
                              # Přeskočíme prázdné řádky
                              if not any(cell is not None for cell in row):
                                   continue

                              record = {}
                              for key, value in zip(keys, row):
                                   # Můžeme přidat konverzi datumu, pokud je potřeba
                                   # if key == 'Datum' and isinstance(value, datetime):
                                   #     record[key] = value.strftime('%Y-%m-%d')
                                   # else:
                                   record[key] = str(value) if value is not None else ""
                              advance_history.append(record)

                     except StopIteration:
                          # Soubor obsahuje pouze hlavičku nebo je prázdný
                          logger.info(f"List 'Zalohy25' v souboru {EXCEL_FILE_NAME_2025} neobsahuje data.")
                 else:
                      logger.warning(f"List 'Zalohy25' nebyl nalezen v souboru {EXCEL_FILE_NAME_2025}")
                      flash(f"List 'Zalohy25' pro historii záloh nebyl nalezen v souboru '{EXCEL_FILE_NAME_2025}'.", "warning")

            except Exception as e:
                logger.error(f"Chyba při čtení souboru {EXCEL_FILE_NAME_2025}: {e}", exc_info=True)
                flash(f"Chyba při načítání historie záloh ze souboru '{EXCEL_FILE_NAME_2025}'.", "error")
            finally:
                 if workbook_2025:
                      workbook_2025.close()


    except Exception as e:
        # Obecná chyba při zpracování historie
        logger.error(f"Neočekávaná chyba při zpracování historie záloh: {e}", exc_info=True)
        flash("Chyba při načítání historie záloh.", "error")

    # Předání dat do šablony
    return render_template(
        "zalohy.html",
        employees=employees_list,
        options=advance_options,
        current_date=datetime.now().strftime("%Y-%m-%d"),
        advance_history=advance_history, # Předáme načtenou nebo prázdnou historii
    )


if __name__ == "__main__":
    # Nastavení logování pro vývojový server Flask
    # V produkci (např. Gunicorn) se logování obvykle konfiguruje jinak
    if not app.debug:
         # V produkci můžeme chtít logovat do souboru
         log_handler = logging.FileHandler('app_prod.log')
         log_handler.setLevel(logging.WARNING) # Logovat jen warning a vyšší
         app.logger.addHandler(log_handler)
    else:
         # V debug módu stačí výchozí Flask logger (na konzoli)
         app.logger.setLevel(logging.INFO)

    # Spuštění aplikace
    # Host='0.0.0.0' zpřístupní aplikaci v lokální síti
    app.run(debug=True, host='0.0.0.0')
