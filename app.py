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
# EXCEL_FILE_NAME_2025 = Config.EXCEL_FILE_NAME_2025 # Odstraněno
SETTINGS_FILE_PATH = Path(Config.SETTINGS_FILE_PATH)
RECIPIENT_EMAIL = Config.RECIPIENT_EMAIL

# Ensure required directories exist
for path in [DATA_PATH, EXCEL_BASE_PATH]:
    path.mkdir(parents=True, exist_ok=True)

# Initialize managers and load settings
employee_manager = EmployeeManager(DATA_PATH)
# Předáme pouze název hlavního souboru
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
                    # Pokud existuje výchozí 'Sheet', přejmenujeme ho
                    sheet = wb["Sheet"]
                    sheet.title = "Týden" # Název šablonového listu
                else:
                    # Pokud neexistuje, vytvoříme ho
                    sheet = wb.create_sheet("Týden")
                # Můžeme přidat i list Zálohy
                if "Zálohy" not in wb.sheetnames:
                     wb.create_sheet("Zálohy")
                     logger.info("Přidán list 'Zálohy' do nového souboru.")

                # Uložíme workbook
                wb.save(str(excel_path))
                wb.close() # Zavřeme workbook po uložení
                excel_exists = True
                logger.info(f"Vytvořen nový Excel soubor: {excel_path}")
            except Exception as e:
                logger.error(f"Nepodařilo se vytvořit Excel soubor: {e}", exc_info=True)
                # Pokud se nepodaří vytvořit, flash zpráva není nutná,
                # protože šablona zobrazí varování, že soubor neexistuje.

        # --- Odstraněna kontrola a vytváření Hodiny2025.xlsx ---

        current_date = datetime.now().strftime("%Y-%m-%d")
        # Získáme objekt isocalendar
        week_calendar_data = excel_manager.ziskej_cislo_tydne(current_date)
        # Získáme číslo týdne z objektu
        week_num_int = week_calendar_data.week if week_calendar_data else 0

        return render_template(
            "index.html", excel_exists=excel_exists, week_number=week_num_int, current_date=current_date
        )
    except Exception as e:
        logger.error(f"Chyba při načítání hlavní stránky: {e}", exc_info=True)
        flash("Došlo k neočekávané chybě při načítání stránky.", "error")
        # V případě chyby vrátíme šablonu s výchozími hodnotami
        return render_template(
            "index.html", excel_exists=False, week_number=0, current_date=datetime.now().strftime("%Y-%m-%d")
        )


@app.route("/download")
def download_file():
    """
    Zpracuje požadavek na stažení souboru Hodiny_Cap.xlsx.
    Najde nejvyšší číslo týdne v názvech listů,
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
            # Použijeme read_only=True pro rychlejší načtení jen pro čtení názvů
            workbook = load_workbook(original_file_path, read_only=True)
            sheet_names = workbook.sheetnames
        except Exception as e:
            logger.error(f"Nepodařilo se načíst Excel soubor pro čtení listů: {e}", exc_info=True)
            # Pokud selže načtení, nemůžeme pokračovat
            raise ValueError("Nepodařilo se otevřít Excel soubor pro analýzu.")
        finally:
             if workbook:
                workbook.close() # Zajistí uzavření souboru i v read-only módu

        # Nalezení nejvyššího čísla týdne
        max_week_number = 0
        week_pattern = re.compile(r"Týden (\d+)") # Regulární výraz pro "Týden X"

        for sheet_name in sheet_names:
            match = week_pattern.match(sheet_name)
            if match:
                try:
                    week_num = int(match.group(1))
                    if week_num > max_week_number:
                        max_week_number = week_num
                except ValueError:
                    # Ignorujeme listy, kde za "Týden " není číslo
                    logger.warning(f"Nalezen list '{sheet_name}', ale neobsahuje platné číslo týdne.")
                    continue


        # Vytvoření názvu pro kopii
        if max_week_number > 0:
            new_filename = f"Hodiny_Cap_Tyden{max_week_number}.xlsx"
        else:
            # Pokud se nenašel žádný list "Týden X", použije se původní název
            new_filename = EXCEL_FILE_NAME
            logger.warning("Nenalezen žádný list ve formátu 'Týden X', stahuje se soubor s původním názvem.")
            # V tomto případě bychom mohli rovnou poslat originál, ale pro konzistenci
            # vytvoříme kopii i s původním názvem.

        # Cesta pro novou kopii
        # Ukládáme kopii do stejného adresáře jako originál
        new_file_path = EXCEL_BASE_PATH / new_filename

        # Vytvoření kopie souboru na serveru
        try:
            shutil.copy2(original_file_path, new_file_path) # copy2 zachovává metadata
            logger.info(f"Vytvořena kopie souboru na serveru: {new_file_path}")
        except Exception as e:
            logger.error(f"Nepodařilo se zkopírovat soubor z '{original_file_path}' do '{new_file_path}': {e}", exc_info=True)
            raise IOError("Chyba při vytváření kopie souboru na serveru.")

        # Odeslání kopie souboru ke stažení
        # Použijeme try-with-resources pro odeslání, aby byl soubor správně uzavřen
        try:
            return send_file(
                str(new_file_path),
                as_attachment=True,
                download_name=new_filename # Zajistí správný název při stahování
            )
        except Exception as send_error:
             logger.error(f"Chyba při odesílání souboru '{new_file_path}': {send_error}", exc_info=True)
             # I když odeslání selže, kopie na serveru zůstane
             raise IOError("Chyba při odesílání souboru uživateli.")


    except FileNotFoundError as e:
        logger.error(f"Soubor nebyl nalezen: {e}")
        flash(str(e), "error") # Zobrazí specifickou chybovou hlášku
        return redirect(url_for("index"))
    except (ValueError, IOError) as e: # Zachytí chyby při načítání, kopírování nebo odesílání
        logger.error(f"Chyba při zpracování souboru pro stažení: {e}")
        flash(str(e), "error")
        return redirect(url_for("index"))
    except Exception as e:
        # Zachytí jakékoli jiné neočekávané chyby
        logger.error(f"Neočekávaná chyba při stahování souboru: {e}", exc_info=True)
        flash("Neočekávaná chyba při stahování souboru.", "error")
        return redirect(url_for("index"))


def validate_email(email):
    """Validace emailové adresy"""
    # Jednoduchá kontrola, lze vylepšit regulárním výrazem
    if not email or "@" not in parseaddr(email)[1] or "." not in email.split('@')[-1]:
        raise ValueError("Neplatná emailová adresa")
    return True


@app.route("/send_email", methods=["POST"])
def send_email():
    """Odešle hlavní Excel soubor emailem."""
    try:
        # Kontrola existence hlavního souboru
        file_path = Path(excel_manager.file_path)
        if not file_path.exists():
            raise FileNotFoundError(f"Excel soubor '{EXCEL_FILE_NAME}' nebyl nalezen pro odeslání emailem")

        # Kontrola emailové konfigurace
        if not Config.RECIPIENT_EMAIL:
            raise ValueError("E-mailová adresa příjemce není nastavena v konfiguraci")
        if not Config.SMTP_USERNAME or not Config.SMTP_PASSWORD:
            raise ValueError("SMTP přihlašovací údaje nejsou nastaveny v konfiguraci")
        if not Config.SMTP_SERVER or not Config.SMTP_PORT:
            raise ValueError("SMTP server nebo port není nastaven v konfiguraci")

        # Validace emailových adres
        sender = Config.SMTP_USERNAME
        recipient = Config.RECIPIENT_EMAIL
        try:
             validate_email(sender)
             validate_email(recipient)
        except ValueError as e:
             # Přidáme kontext k chybě
             raise ValueError(f"Neplatná emailová adresa v konfiguraci: {e}")


        # Vytvoření zprávy
        msg = MIMEMultipart()
        # Předmět emailu
        subject = f'{EXCEL_FILE_NAME} - {datetime.now().strftime("%Y-%m-%d %H:%M")}'
        msg["Subject"] = subject
        msg["From"] = sender
        msg["To"] = recipient

        # Přidání textu zprávy
        # Můžeme přidat název aplikace z konfigurace, pokud existuje
        app_name = getattr(Config, 'APP_NAME', 'Evidence pracovní doby')
        body = f"""Dobrý den,

v příloze zasílám aktuální výkaz pracovní doby ({EXCEL_FILE_NAME}).

S pozdravem,
{app_name}
"""
        # Použijeme UTF-8 kódování pro tělo emailu
        msg.attach(MIMEText(body, "plain", "utf-8"))

        # Přidání přílohy
        try:
            with open(file_path, "rb") as f:
                # Použijeme application/vnd.openxmlformats-officedocument.spreadsheetml.sheet pro moderní Excel
                attachment = MIMEApplication(f.read(), _subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                # Správné kódování názvu přílohy pro různé emailové klienty
                attachment.add_header("Content-Disposition", "attachment", filename=EXCEL_FILE_NAME)
                msg.attach(attachment)
        except IOError as e:
             logger.error(f"Chyba při čtení souboru přílohy '{file_path}': {e}", exc_info=True)
             raise ValueError("Nepodařilo se připojit soubor k emailu.")


        # Nastavení SSL/TLS kontextu s vyšším zabezpečením
        ssl_context = ssl.create_default_context()
        # Můžeme vynutit TLSv1.2 nebo vyšší
        # ssl_context.minimum_version = ssl.TLSVersion.TLSv1_2
        ssl_context.check_hostname = True # Důležité pro bezpečnost
        ssl_context.verify_mode = ssl.CERT_REQUIRED

        # Odeslání emailu pomocí SMTP_SSL
        try:
            # Zvýšený timeout pro případ pomalejšího spojení
            with smtplib.SMTP_SSL(Config.SMTP_SERVER, Config.SMTP_PORT, context=ssl_context, timeout=60) as smtp:
                smtp.login(Config.SMTP_USERNAME, Config.SMTP_PASSWORD)
                smtp.send_message(msg)
            flash("Email byl úspěšně odeslán.", "success")
            logger.info(f"Email s výkazem '{EXCEL_FILE_NAME}' byl úspěšně odeslán na adresu {recipient}")
        except smtplib.SMTPAuthenticationError:
            logger.error("Chyba při přihlášení k SMTP serveru (nesprávné jméno nebo heslo)")
            raise ValueError("Nesprávné přihlašovací údaje k e-mailovému serveru.")
        except smtplib.SMTPException as e:
            logger.error(f"Obecná SMTP chyba při odesílání emailu: {e}", exc_info=True)
            raise ConnectionError(f"Chyba při komunikaci s e-mailovým serverem: {e}")
        except ssl.SSLError as e:
            logger.error(f"SSL chyba při připojování k SMTP serveru: {e}", exc_info=True)
            raise ConnectionError("Chyba zabezpečeného spojení s e-mailovým serverem.")
        except TimeoutError:
             logger.error("Vypršel časový limit při připojování nebo odesílání emailu.")
             raise ConnectionError("Časový limit pro spojení s e-mailovým serverem vypršel.")
        except Exception as e:
             # Zachytí jiné možné chyby (např. síťové)
             logger.error(f"Neočekávaná chyba při odesílání emailu: {e}", exc_info=True)
             raise ConnectionError(f"Neočekávaná chyba při odesílání emailu: {e}")


    except FileNotFoundError as e:
        logger.error(f"Soubor pro odeslání emailem nebyl nalezen: {e}")
        flash(str(e), "error")
    except ValueError as e:
        # Chyby z validace konfigurace nebo emailových adres
        logger.error(f"Chyba konfigurace nebo dat pro odeslání emailu: {e}")
        flash(str(e), "error")
    except (ConnectionError, IOError) as e:
         # Chyby při připojování k SMTP nebo práci se souborem
         logger.error(f"Chyba připojení nebo souboru při odesílání emailu: {e}")
         flash(str(e), "error")
    except Exception as e:
        # Obecná neočekávaná chyba
        logger.error(f"Neočekávaná chyba v procesu odesílání emailu: {e}", exc_info=True)
        flash("Neočekávaná chyba při odesílání emailu.", "error")

    # Vždy přesměrujeme zpět na index, ať už došlo k chybě nebo ne
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
                    raise ValueError("Jméno zaměstnance je příliš dlouhé (max 100 znaků)")
                # Upřesnění povolených znaků (písmena, čísla, mezery, pomlčka, tečka, česká diakritika)
                if not re.match(r"^[\w\s\-\.ěščřžýáíéúůďťňĚŠČŘŽÝÁÍÉÚŮĎŤŇ]+$", employee_name):
                     raise ValueError("Jméno zaměstnance obsahuje nepovolené znaky.")


                if employee_manager.pridat_zamestnance(employee_name):
                    flash(f'Zaměstnanec "{employee_name}" byl přidán.', "success")
                    # Přesměrování po úspěšném přidání pro čisté URL
                    return redirect(url_for('manage_employees'))
                else:
                    # Chyba (např. zaměstnanec již existuje) je logována v EmployeeManager
                    flash(f'Zaměstnanec "{employee_name}" již existuje nebo došlo k chybě při přidávání.', "error")

            elif action == "select":
                employee_name = request.form.get("employee_name", "")
                if not employee_name:
                    raise ValueError("Nebyl vybrán zaměstnanec pro označení/odznačení")

                # Ověření existence zaměstnance před akcí
                if employee_name not in employee_manager.zamestnanci:
                     raise ValueError(f'Zaměstnanec "{employee_name}" neexistuje')

                if employee_name in employee_manager.vybrani_zamestnanci:
                    if employee_manager.odebrat_vybraneho_zamestnance(employee_name):
                         flash(f'Zaměstnanec "{employee_name}" byl odebrán z výběru.', "success")
                    else:
                         flash(f'Nepodařilo se odebrat zaměstnance "{employee_name}" z výběru.', "error")
                else:
                    if employee_manager.pridat_vybraneho_zamestnance(employee_name):
                         flash(f'Zaměstnanec "{employee_name}" byl přidán do výběru.', "success")
                    else:
                         flash(f'Nepodařilo se přidat zaměstnance "{employee_name}" do výběru.', "error")

                # Není potřeba volat save_config() zde, volá se uvnitř metod EmployeeManager
                # Přesměrování po úspěšné akci
                return redirect(url_for('manage_employees'))


            elif action == "edit":
                old_name = request.form.get("old_name", "").strip()
                new_name = request.form.get("new_name", "").strip()

                if not old_name or not new_name:
                    raise ValueError("Původní i nové jméno musí být vyplněno")
                if len(new_name) > 100:
                    raise ValueError("Nové jméno je příliš dlouhé (max 100 znaků)")
                if old_name == new_name:
                     flash("Nové jméno je stejné jako původní. Nebyly provedeny žádné změny.", "info")
                     return redirect(url_for('manage_employees')) # Není co měnit

                # Validace nového jména
                if not re.match(r"^[\w\s\-\.ěščřžýáíéúůďťňĚŠČŘŽÝÁÍÉÚŮĎŤŇ]+$", new_name):
                    raise ValueError("Nové jméno obsahuje nepovolené znaky.")

                # Použití metody pro úpravu podle jména
                if employee_manager.upravit_zamestnance_podle_jmena(old_name, new_name):
                    flash(f'Zaměstnanec "{old_name}" byl úspěšně upraven na "{new_name}".', "success")
                    return redirect(url_for('manage_employees'))
                else:
                    # Chyba (např. staré jméno neexistuje, nové jméno už existuje) je logována v EmployeeManager
                    flash(f'Nepodařilo se upravit zaměstnance "{old_name}". Zkontrolujte, zda původní jméno existuje a nové jméno není již použito.', "error")


            elif action == "delete":
                employee_name = request.form.get("employee_name", "")
                if not employee_name:
                    raise ValueError("Nebyl vybrán zaměstnanec k odstranění")

                # Použití metody pro smazání podle jména
                if employee_manager.smazat_zamestnance_podle_jmena(employee_name):
                    flash(f'Zaměstnanec "{employee_name}" byl úspěšně smazán.', "success")
                    return redirect(url_for('manage_employees'))
                else:
                    # Chyba (např. zaměstnanec neexistuje) je logována v EmployeeManager
                    flash(f'Nepodařilo se smazat zaměstnance "{employee_name}". Zkontrolujte, zda zaměstnanec existuje.', "error")

            else:
                raise ValueError(f"Neznámá akce: {action}")

        except ValueError as e:
            # Zobrazíme validační chyby uživateli
            flash(str(e), "error")
            logger.error(f"Chyba při správě zaměstnanců (akce: {request.form.get('action', 'N/A')}): {e}")
            # Necháme uživatele na stránce, aby viděl chybu a mohl ji opravit
        except Exception as e:
            # Obecná chyba
            flash("Došlo k neočekávané chybě při správě zaměstnanců.", "error")
            logger.error(f"Neočekávaná chyba při správě zaměstnanců (akce: {request.form.get('action', 'N/A')}): {e}", exc_info=True)
            # Necháme uživatele na stránce

    # Pro GET požadavek nebo po chybě v POST
    # Získáme aktuální seznam zaměstnanců pro zobrazení
    employees = employee_manager.get_all_employees()
    return render_template("employees.html", employees=employees)


@app.route("/zaznam", methods=["GET", "POST"])
def record_time():
    selected_employees = employee_manager.get_vybrani_zamestnanci() # Získáme aktuální vybrané
    if not selected_employees:
        flash("Nejsou vybráni žádní zaměstnanci pro záznam pracovní doby.", "warning")
        return redirect(url_for("manage_employees"))

    # Získáme výchozí hodnoty z nastavení
    settings = get_settings()
    default_start_time = settings.get("start_time", "07:00")
    default_end_time = settings.get("end_time", "18:00")
    default_lunch_duration = settings.get("lunch_duration", 1.0)

    # Pro GET požadavek použijeme výchozí hodnoty nebo hodnoty z formuláře, pokud byly odeslány
    current_date = request.form.get("date", datetime.now().strftime("%Y-%m-%d"))
    start_time = request.form.get("start_time", default_start_time)
    end_time = request.form.get("end_time", default_end_time)
    lunch_duration_input = request.form.get("lunch_duration", str(default_lunch_duration))

    if request.method == "POST":
        try:
            # Validace data
            date_str = request.form.get("date", "")
            try:
                selected_date = datetime.strptime(date_str, "%Y-%m-%d").date()
                today = datetime.now().date()
                if selected_date > today:
                     raise ValueError("Nelze zaznamenat pracovní dobu pro budoucí datum")
                if selected_date.weekday() >= 5: # 5 = Sobota, 6 = Neděle
                    raise ValueError("Nelze zaznamenat pracovní dobu na víkend")
                # Můžeme přidat i limit, jak daleko do minulosti lze zaznamenávat
                # if (today - selected_date).days > 30:
                #    raise ValueError("Lze zaznamenávat maximálně 30 dní zpětně")
            except ValueError as e:
                 if "budoucí" in str(e) or "víkend" in str(e): # Propagujeme specifické chyby
                     raise
                 raise ValueError("Neplatný formát data (použijte YYYY-MM-DD) nebo neplatné datum")


            # Validace časů
            start_time_str = request.form.get("start_time", "")
            end_time_str = request.form.get("end_time", "")
            try:
                start = datetime.strptime(start_time_str, "%H:%M")
                end = datetime.strptime(end_time_str, "%H:%M")
                if end <= start:
                    raise ValueError("Čas konce práce musí být pozdější než čas začátku")
            except ValueError as e:
                if "pozdější" in str(e):
                    raise
                raise ValueError("Neplatný formát času (použijte HH:MM)")

            # Validace délky pauzy
            lunch_duration_str = request.form.get("lunch_duration", "")
            try:
                # Povolíme desetinnou čárku i tečku
                lunch_duration = float(lunch_duration_str.replace(",", "."))
                if lunch_duration < 0:
                    raise ValueError("Délka pauzy na oběd nemůže být záporná")
                # Výpočet délky pracovní doby v hodinách
                work_duration_hours = (end - start).total_seconds() / 3600
                # Pauza nemůže být delší než pracovní doba (pokud je pracovní doba kladná)
                if work_duration_hours > 0 and lunch_duration >= work_duration_hours:
                    raise ValueError("Délka pauzy nemůže být stejná nebo delší než celková pracovní doba")
                # Můžeme přidat i horní limit pro pauzu, např. 4 hodiny
                if lunch_duration > 4:
                     raise ValueError("Délka pauzy nesmí přesáhnout 4 hodiny")
            except ValueError as e:
                # Propagujeme specifické chyby
                if any(msg in str(e) for msg in ["záporná", "delší", "přesáhnout"]):
                    raise
                # Obecná chyba pro nečíselný vstup
                raise ValueError("Délka pauzy na oběd musí být platné číslo (např. 1 nebo 0,5)")


            # Uložení do Hodiny_Cap.xlsx
            try:
                # Předáváme validované stringy a float
                success = excel_manager.ulozit_pracovni_dobu(
                    date_str, start_time_str, end_time_str, lunch_duration, selected_employees
                )
                if success:
                    flash("Pracovní doba byla úspěšně zaznamenána.", "success")
                    # Po úspěšném uložení můžeme přesměrovat nebo zobrazit potvrzení
                    # Přesměrování na index nebo jinou stránku je často lepší než zůstat na formuláři
                    return redirect(url_for('index'))
                else:
                    # Chyba byla zalogována v excel_manager
                    raise IOError("Nepodařilo se uložit záznam do Excel souboru. Zkuste to prosím znovu.")

            except (IOError, Exception) as e:
                 # Zobrazíme chybu uživateli
                 logger.error(f"Chyba při komunikaci s Excel souborem během ukládání pracovní doby: {e}", exc_info=True)
                 # Použijeme e přímo, pokud je to IOError, jinak obecnou zprávu
                 error_message = str(e) if isinstance(e, IOError) else "Nastala chyba při ukládání do Excelu."
                 raise RuntimeError(error_message) # Změníme na RuntimeError pro odlišení od ValueError


        except ValueError as e:
            # Chyby z validace vstupů
            flash(str(e), "error")
            logger.warning(f"Chyba validace při záznamu pracovní doby: {e}")
            # Hodnoty pro formulář zůstanou ty, které uživatel zadal (jsou v request.form)
            current_date = request.form.get("date", current_date)
            start_time = request.form.get("start_time", start_time)
            end_time = request.form.get("end_time", end_time)
            lunch_duration_input = request.form.get("lunch_duration", lunch_duration_input)
        except RuntimeError as e:
             # Chyby při ukládání do Excelu
             flash(str(e), "error")
             # Hodnoty pro formulář zůstanou ty, které uživatel zadal
             current_date = request.form.get("date", current_date)
             start_time = request.form.get("start_time", start_time)
             end_time = request.form.get("end_time", end_time)
             lunch_duration_input = request.form.get("lunch_duration", lunch_duration_input)
        except Exception as e:
            # Obecné neočekávané chyby
            flash("Došlo k neočekávané chybě při zpracování záznamu.", "error")
            logger.error(f"Neočekávaná chyba při záznamu pracovní doby: {e}", exc_info=True)
            # Vrátíme výchozí hodnoty pro formulář v případě vážné chyby
            current_date = datetime.now().strftime("%Y-%m-%d")
            start_time = default_start_time
            end_time = default_end_time
            lunch_duration_input = str(default_lunch_duration)


    # Pro GET nebo po chybě v POST zobrazíme formulář
    # Zajistíme, aby se délka pauzy zobrazovala s desetinnou tečkou pro input type="number"
    try:
         lunch_duration_formatted = str(float(lunch_duration_input.replace(",", ".")))
    except ValueError:
         lunch_duration_formatted = str(default_lunch_duration) # Fallback na výchozí


    return render_template(
        "record_time.html",
        selected_employees=selected_employees,
        current_date=current_date,
        start_time=start_time,
        end_time=end_time,
        lunch_duration=lunch_duration_formatted, # Použijeme formátovanou hodnotu pro input
    )


@app.route("/excel_viewer", methods=["GET"])
def excel_viewer():
    # Zobrazuje pouze hlavní soubor
    excel_files = [EXCEL_FILE_NAME]
    # Pokud není soubor specifikován, použije se první (a jediný) v seznamu
    selected_file = request.args.get("file", excel_files[0])

    # Pokud byl z nějakého důvodu poslán jiný název souboru, vrátíme chybu
    if selected_file != EXCEL_FILE_NAME:
         flash(f"Zobrazení souboru '{selected_file}' není podporováno.", "error")
         return redirect(url_for("index"))


    active_sheet = request.args.get("sheet", None)
    workbook = None
    data = [] # Inicializace dat
    sheet_names = [] # Inicializace seznamu listů

    try:
        file_path = Path(excel_manager.file_path)
        if not file_path.exists():
            # Pokud hlavní soubor neexistuje, nemá smysl pokračovat
            raise FileNotFoundError(f"Hlavní Excel soubor '{EXCEL_FILE_NAME}' nebyl nalezen.")

        # Načtení workbooku v read-only módu
        workbook = load_workbook(file_path, read_only=True, data_only=True)
        sheet_names = workbook.sheetnames # Získáme názvy listů

        if not sheet_names:
            # Měl by existovat alespoň list "Týden" nebo "Zálohy"
            raise ValueError("Excel soubor neobsahuje žádné listy.")

        # Výběr aktivního listu
        # Pokud active_sheet není specifikován nebo neexistuje, vybereme první list
        if active_sheet not in sheet_names:
             active_sheet = sheet_names[0]
        sheet = workbook[active_sheet]

        # Načtení dat s omezením počtu řádků
        MAX_ROWS_TO_DISPLAY = 500 # Snížení limitu pro rychlejší načítání a menší zátěž
        # Načtení hlavičky (první řádek), pokud existuje
        header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if header_row:
             header = [str(cell) if cell is not None else "" for cell in header_row]
             data.append(header)

        # Načtení zbytku dat (od druhého řádku)
        rows_loaded = 0
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if rows_loaded >= MAX_ROWS_TO_DISPLAY:
                flash(f"Zobrazeno prvních {MAX_ROWS_TO_DISPLAY} řádků dat (bez hlavičky).", "warning")
                break
            # Převod všech buněk na string, None hodnoty na prázdný řetězec
            data.append([str(cell) if cell is not None else "" for cell in row])
            rows_loaded += 1

        # Pokud nejsou žádná data (ani hlavička), přidáme prázdný řádek pro šablonu
        if not data:
             data.append([])


    except FileNotFoundError as e:
        logger.error(f"Soubor pro zobrazení nebyl nalezen: {e}")
        flash(str(e), "error")
        # V případě nenalezení souboru přesměrujeme na index
        return redirect(url_for("index"))
    except InvalidFileException:
        logger.error(f"Soubor {selected_file} je poškozen nebo má neplatný formát")
        flash(f"Soubor '{selected_file}' je poškozen nebo má neplatný formát.", "error")
        return redirect(url_for("index"))
    except ValueError as e:
        logger.error(f"Chyba při práci s Excel souborem '{selected_file}': {e}")
        flash(str(e), "error")
        return redirect(url_for("index"))
    except PermissionError:
        logger.error(f"Nedostatečná oprávnění pro čtení souboru {selected_file}")
        flash(f"Nedostatečná oprávnění pro čtení souboru '{selected_file}'.", "error")
        return redirect(url_for("index"))
    except Exception as e:
        logger.error(f"Neočekávaná chyba při zobrazení Excel souboru '{selected_file}': {e}", exc_info=True)
        flash("Chyba při zobrazení Excel souboru.", "error")
        return redirect(url_for("index"))
    finally:
        if workbook:
            workbook.close() # Zajistíme uzavření workbooku

    # Předání dat do šablony
    return render_template(
        "excel_viewer.html",
        excel_files=excel_files, # Seznam obsahuje jen hlavní soubor
        selected_file=selected_file,
        sheet_names=sheet_names, # Předáme načtené názvy listů
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
                    raise ValueError("Délka pauzy nemůže být záporná")
                if lunch_duration > 4: # Přidán horní limit
                     raise ValueError("Délka pauzy nesmí být větší než 4 hodiny")
            except ValueError as e:
                 if "záporná" in str(e) or "větší než 4" in str(e):
                     raise
                 raise ValueError("Délka pauzy musí být nezáporné číslo (0 až 4)")


            # Validace dat projektu
            project_name = request.form.get("project_name", "").strip()
            project_start_str = request.form.get("start_date", "")
            project_end_str = request.form.get("end_date", "") # Může být prázdné

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
            current_settings = get_settings() # Získáme aktuální nastavení
            current_settings.update(
                {
                    "start_time": start_time_str,
                    "end_time": end_time_str,
                    "lunch_duration": lunch_duration,
                    "project_info": {
                        "name": project_name,
                        "start_date": project_start_str, # Ukládáme string
                        "end_date": project_end_str,     # Ukládáme string (může být prázdný)
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
                logger.info("Nastavení byla úspěšně uložena do souboru a session.")
            except IOError as e:
                 logger.error(f"Chyba při zápisu do souboru nastavení '{SETTINGS_FILE_PATH}': {e}", exc_info=True)
                 # Pokud selže uložení JSON, nebudeme pokračovat s Excelem
                 raise RuntimeError("Nepodařilo se uložit nastavení do konfiguračního souboru.")


            # Aktualizace informací o projektu v Excel souboru
            excel_update_success = False
            try:
                # Předáme stringy datumu, metoda update_project_info si je zpracuje
                excel_update_success = excel_manager.update_project_info(
                    project_name,
                    project_start_str,
                    project_end_str if project_end_str else None, # Předáme None, pokud není konec zadán
                )
                if not excel_update_success:
                     # Chyba byla zalogována v excel_manager
                     raise RuntimeError("Metoda update_project_info v ExcelManager selhala.")

            except Exception as e:
                 logger.error(f"Chyba při aktualizaci Excel souboru s informacemi o projektu: {e}", exc_info=True)
                 # I když Excel selže, nastavení v JSONu a session jsou uložena
                 flash("Nastavení bylo uloženo, ale nepodařilo se aktualizovat informace v Excel souboru.", "warning")
                 # Přesměrujeme, aby se formulář znovu neodeslal
                 return redirect(url_for("settings_page"))


            flash("Nastavení bylo úspěšně uloženo a informace v Excelu aktualizovány.", "success")
            return redirect(url_for("settings_page")) # Přesměrujeme po úspěšném uložení

        except ValueError as e:
            # Chyby z validace vstupů
            flash(str(e), "error")
            logger.warning(f"Chyba validace při ukládání nastavení: {e}")
            # Necháme uživatele na stránce, aby mohl opravit chybu
            # Hodnoty pro formulář se vezmou z request.form v šabloně nebo z get_settings()
        except RuntimeError as e:
             # Chyby při ukládání (JSON nebo Excel)
             flash(str(e), "error")
             # Necháme uživatele na stránce
        except Exception as e:
            # Obecná neočekávaná chyba
            flash("Došlo k neočekávané chybě při ukládání nastavení.", "error")
            logger.error(f"Neočekávaná chyba při ukládání nastavení: {e}", exc_info=True)
            # Necháme uživatele na stránce

    # Pro GET požadavek nebo po chybě v POST zobrazíme stránku s aktuálním nastavením
    return render_template("settings_page.html", settings=get_settings())


@app.route("/zalohy", methods=["GET", "POST"])
def zalohy():
    # Inicializace ZalohyManager s cestou k adresáři excel souborů
    # ZalohyManager pracuje pouze s Hodiny_Cap.xlsx
    zalohy_manager = ZalohyManager(EXCEL_BASE_PATH)
    # Historie záloh se již nenačítá z Hodiny2025.xlsx
    advance_history = []
    # Získání jmen zaměstnanců pro formulář
    employees_list = employee_manager.zamestnanci
    # Získání možností záloh pro formulář z Hodiny_Cap.xlsx
    advance_options = excel_manager.get_advance_options()

    if request.method == "POST":
        try:
            # Validace vstupních dat z formuláře
            employee_name = request.form.get("employee_name")
            if not employee_name or employee_name not in employees_list:
                raise ValueError("Vyberte platného zaměstnance ze seznamu")

            amount_str = request.form.get("amount")
            try:
                amount = float(amount_str.replace(",", "."))
                # Použití validační metody z ZalohyManager
                zalohy_manager.validate_amount(amount)
            except (ValueError, TypeError, AttributeError) as e:
                 raise ValueError(f"Neplatná částka zálohy: {e}. Zadejte kladné číslo.")


            currency = request.form.get("currency")
            # Použití validační metody z ZalohyManager
            zalohy_manager.validate_currency(currency)


            option = request.form.get("option")
            # Validace, zda je option jednou z načtených/platných možností
            if not option or option not in advance_options:
                raise ValueError(f"Vyberte platnou možnost zálohy ({', '.join(advance_options)})")
            # Validace proti pevně daným možnostem v ZalohyManager (pokud je potřeba)
            # zalohy_manager.validate_option(option)


            date_str = request.form.get("date")
            # Použití validační metody z ZalohyManager
            zalohy_manager.validate_date(date_str)


            # Uložení zálohy pomocí ZalohyManager (ukládá do Hodiny_Cap.xlsx)
            success = zalohy_manager.add_or_update_employee_advance(
                employee_name=employee_name, amount=amount, currency=currency, option=option, date=date_str
            )

            if success:
                flash("Záloha byla úspěšně uložena.", "success")
                # Přesměrování po úspěšném uložení
                return redirect(url_for('zalohy'))
            else:
                 # Chyba byla zalogována v ZalohyManager
                 raise RuntimeError("Nepodařilo se uložit zálohu. Zkontrolujte logy.")


        except ValueError as e:
            # Chyby z validace
            flash(str(e), "error")
            logger.warning(f"Chyba validace při ukládání zálohy: {e}")
        except RuntimeError as e:
             # Chyby při ukládání
             flash(str(e), "error")
        except Exception as e:
            # Obecné chyby
            flash("Došlo k neočekávané chybě při ukládání zálohy.", "error")
            logger.error(f"Neočekávaná chyba při ukládání zálohy: {e}", exc_info=True)

    # --- Odstraněno načítání historie z Hodiny2025.xlsx ---
    # Místo toho můžeme implementovat načítání historie z Hodiny_Cap.xlsx,
    # pokud je to požadováno, ale prozatím necháme historii prázdnou.
    # Pokud byste chtěli načítat historii z listu "Zálohy" v Hodiny_Cap.xlsx,
    # kód by vypadal podobně jako původní kód pro Hodiny2025, ale četl by z excel_manager.file_path
    # a listu "Zálohy". Je třeba dát pozor na formát dat v tomto listu.

    # Pro GET požadavek nebo po chybě v POST zobrazíme stránku
    return render_template(
        "zalohy.html",
        employees=employees_list,
        options=advance_options,
        current_date=datetime.now().strftime("%Y-%m-%d"),
        advance_history=advance_history, # Předáme prázdnou historii
    )


if __name__ == "__main__":
    # Nastavení logování pro vývojový server Flask
    if not app.debug:
         log_handler = logging.FileHandler('app_prod.log', encoding='utf-8')
         log_handler.setLevel(logging.WARNING)
         log_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
         log_handler.setFormatter(log_formatter)
         app.logger.addHandler(log_handler)
    else:
         app.logger.setLevel(logging.INFO) # V debug módu logujeme více informací

    # Spuštění aplikace
    # Host='0.0.0.0' zpřístupní aplikaci v lokální síti
    # Použití portu 5000 je standardní pro Flask
    app.run(debug=True, host='0.0.0.0', port=5000)
