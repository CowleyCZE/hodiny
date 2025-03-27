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

import openpyxl
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
        return render_template(
            "index.html", excel_exists=excel_exists, week_number=week_number, current_date=current_date
        )
    except Exception as e:
        logger.error(f"Chyba při načítání hlavní stránky: {e}")
        flash("Došlo k neočekávané chybě při načítání stránky.", "error")
        return render_template(
            "index.html", excel_exists=False, week_number=0, current_date=datetime.now().strftime("%Y-%m-%d")
        )


@app.route("/download")
def download_file():
    try:
        file_path = Path(excel_manager.file_path)
        if not file_path.exists():
            raise FileNotFoundError("Excel soubor nebyl nalezen")
        return send_file(str(file_path), as_attachment=True)
    except FileNotFoundError as e:
        logger.error(f"Soubor nebyl nalezen: {e}")
        flash("Excel soubor nebyl nalezen.", "error")
        return redirect(url_for("index"))
    except Exception as e:
        logger.error(f"Chyba při stahování souboru: {e}")
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
                if not employee_name.replace(" ", "").isalnum():
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
                if not new_name.replace(" ", "").isalnum():
                    raise ValueError("Nové jméno obsahuje nepovolené znaky")

                try:
                    idx = employee_manager.zamestnanci.index(old_name) + 1
                    if employee_manager.upravit_zamestnance(idx, new_name):
                        flash(f'Zaměstnanec "{old_name}" byl upraven na "{new_name}".', "success")
                    else:
                        raise ValueError(f'Nepodařilo se upravit zaměstnance "{old_name}"')
                except ValueError as e:
                    if "list.index" in str(e):
                        raise ValueError(f'Zaměstnanec "{old_name}" nebyl nalezen')
                    raise

            elif action == "delete":
                employee_name = request.form.get("employee_name", "")
                if not employee_name:
                    raise ValueError("Nebyl vybrán zaměstnanec k odstranění")

                try:
                    idx = employee_manager.zamestnanci.index(employee_name) + 1
                    if employee_manager.smazat_zamestnance(idx):
                        flash(f'Zaměstnanec "{employee_name}" byl smazán.', "success")
                    else:
                        raise ValueError(f'Nepodařilo se smazat zaměstnance "{employee_name}"')
                except ValueError as e:
                    if "list.index" in str(e):
                        raise ValueError(f'Zaměstnanec "{employee_name}" nebyl nalezen')
                    raise

            else:
                raise ValueError("Neplatná akce")

        except ValueError as e:
            flash(str(e), "error")
            logger.error(f"Chyba při správě zaměstnanců: {e}")
        except Exception as e:
            flash("Došlo k neočekávané chybě při správě zaměstnanců.", "error")
            logger.error(f"Neočekávaná chyba při správě zaměstnanců: {e}")

    # Převedení seznamů na formát očekávaný šablonou
    employees = [
        {"name": name, "selected": name in employee_manager.vybrani_zamestnanci}
        for name in employee_manager.zamestnanci
    ]
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
    lunch_duration = settings.get("lunch_duration", 1)

    if request.method == "POST":
        try:
            # Validace data
            date = request.form.get("date", "")
            try:
                datetime.strptime(date, "%Y-%m-%d")
            except ValueError:
                raise ValueError("Neplatný formát data (použijte YYYY-MM-DD)")

            # Validace časů
            start_time = request.form.get("start_time", "")
            end_time = request.form.get("end_time", "")
            try:
                start = datetime.strptime(start_time, "%H:%M")
                end = datetime.strptime(end_time, "%H:%M")
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
                work_hours = (end - start).seconds / 3600
                if lunch_duration >= work_hours:
                    raise ValueError("Délka pauzy nemůže být delší než pracovní doba")
            except ValueError as e:
                if any(msg in str(e) for msg in ["nemůže být záporná", "nemůže být delší"]):
                    raise
                raise ValueError("Délka pauzy musí být číslo")

            # Uložení do Hodiny_Cap.xlsx
            try:
                excel_manager.ulozit_pracovni_dobu(date, start_time, end_time, lunch_duration, selected_employees)
                flash("Pracovní doba byla úspěšně zaznamenána.", "success")
            except Exception as e:
                logger.error(f"Chyba při ukládání do Excel souboru: {e}")
                raise ValueError("Nepodařilo se uložit pracovní dobu do Excel souboru")

        except ValueError as e:
            flash(str(e), "error")
            logger.error(f"Chyba validace pracovní doby: {e}")
        except Exception as e:
            flash("Došlo k neočekávané chybě při ukládání pracovní doby.", "error")
            logger.error(f"Neočekávaná chyba při ukládání pracovní doby: {e}")

    return render_template(
        "record_time.html",
        selected_employees=selected_employees,
        current_date=current_date,
        start_time=start_time,
        end_time=end_time,
        lunch_duration=lunch_duration,
    )


@app.route("/excel_viewer", methods=["GET"])
def excel_viewer():
    excel_files = ["Hodiny_Cap.xlsx", "Hodiny2025.xlsx"]
    selected_file = request.args.get("file", excel_files[0])
    active_sheet = request.args.get("sheet", None)
    workbook = None

    try:
        # Určení cesty k souboru
        if selected_file == "Hodiny_Cap.xlsx":
            file_path = excel_manager.file_path
        elif selected_file == "Hodiny2025.xlsx":
            file_path = EXCEL_BASE_PATH / EXCEL_FILE_NAME_2025
        else:
            raise ValueError("Neplatný název souboru")

        # Kontrola existence souboru - konverze na Path objekt pro jednotnou práci
        file_path = Path(file_path)
        if not file_path.exists():
            raise FileNotFoundError(f"Soubor {selected_file} nebyl nalezen")

        # Načtení workbooku v read-only módu pro úsporu paměti
        workbook = load_workbook(file_path, read_only=True, data_only=True)

        if not workbook.sheetnames:
            raise ValueError("Excel soubor neobsahuje žádné listy")

        # Výběr aktivního listu
        active_sheet = active_sheet if active_sheet in workbook.sheetnames else workbook.sheetnames[0]
        sheet = workbook[active_sheet]

        # Načtení dat s omezením počtu řádků pro prevenci přetížení paměti
        MAX_ROWS = 1000
        data = []
        for i, row in enumerate(sheet.iter_rows(values_only=True)):
            if i >= MAX_ROWS:
                flash(f"Zobrazeno prvních {MAX_ROWS} řádků.", "warning")
                break
            data.append([str(cell) if cell is not None else "" for cell in row])

        return render_template(
            "excel_viewer.html",
            excel_files=excel_files,
            selected_file=selected_file,
            sheet_names=workbook.sheetnames,
            active_sheet=active_sheet,
            data=data,
        )

    except FileNotFoundError as e:
        logger.error(f"Soubor nebyl nalezen: {e}")
        flash("Požadovaný Excel soubor nebyl nalezen.", "error")
    except InvalidFileException:
        logger.error(f"Soubor {selected_file} je poškozen")
        flash("Soubor je poškozen nebo má neplatný formát.", "error")
    except ValueError as e:
        logger.error(f"Chyba při práci s Excel souborem: {e}")
        flash(str(e), "error")
    except PermissionError:
        logger.error(f"Nedostatečná oprávnění pro čtení souboru {selected_file}")
        flash("Nedostatečná oprávnění pro čtení souboru.", "error")
    except Exception as e:
        logger.error(f"Neočekávaná chyba při zobrazení Excel souboru: {e}")
        flash("Chyba při zobrazení Excel souboru.", "error")
    finally:
        if workbook:
            workbook.close()

    return redirect(url_for("index"))


@app.route("/settings", methods=["GET", "POST"])
def settings_page():
    """Handle settings page"""
    if request.method == "POST":
        try:
            # Validace vstupních dat
            start_time = request.form.get("start_time", "")
            end_time = request.form.get("end_time", "")
            lunch_duration_str = request.form.get("lunch_duration", "")

            # Validace času
            try:
                datetime.strptime(start_time, "%H:%M")
                datetime.strptime(end_time, "%H:%M")
            except ValueError:
                raise ValueError("Neplatný formát času (použijte HH:MM)")

            # Validace délky pauzy
            try:
                lunch_duration = float(lunch_duration_str.replace(",", "."))
                if lunch_duration < 0:
                    raise ValueError
            except ValueError:
                raise ValueError("Délka pauzy musí být nezáporné číslo")

            # Validace dat projektu
            project_start = request.form.get("start_date", "")
            project_end = request.form.get("end_date", "")

            if project_start and project_end:
                try:
                    start_date = datetime.strptime(project_start, "%Y-%m-%d")
                    end_date = datetime.strptime(project_end, "%Y-%m-%d")
                    if end_date < start_date:
                        raise ValueError("Datum konce projektu nemůže být dřívější než datum začátku")
                except ValueError as e:
                    if "nemůže být dřívější" in str(e):
                        raise
                    raise ValueError("Neplatný formát data (použijte YYYY-MM-DD)")

            # Aktualizace nastavení
            new_settings = get_settings()
            new_settings.update(
                {
                    "start_time": start_time,
                    "end_time": end_time,
                    "lunch_duration": lunch_duration,
                    "project_info": {
                        "name": request.form.get("project_name", "").strip(),
                        "start_date": project_start,
                        "end_date": project_end,
                    },
                }
            )

            # Vytvoření adresáře pro nastavení, pokud neexistuje
            SETTINGS_FILE_PATH.parent.mkdir(parents=True, exist_ok=True)

            # Uložení nastavení
            with open(SETTINGS_FILE_PATH, "w", encoding="utf-8") as f:
                json.dump(new_settings, f, indent=4, ensure_ascii=False)

            session["settings"] = new_settings

            # Aktualizace informací o projektu v Excel souboru
            if new_settings["project_info"]["name"]:
                excel_manager.update_project_info(
                    new_settings["project_info"]["name"],
                    new_settings["project_info"]["start_date"],
                    new_settings["project_info"]["end_date"],
                )

            flash("Nastavení bylo úspěšně uloženo.", "success")

        except ValueError as e:
            flash(str(e), "error")
            logger.error(f"Chyba validace nastavení: {e}")
        except Exception as e:
            flash("Došlo k neočekávané chybě při ukládání nastavení.", "error")
            logger.error(f"Neočekávaná chyba při ukládání nastavení: {e}")

    return render_template("settings_page.html", settings=get_settings())


@app.route("/zalohy", methods=["GET", "POST"])
def zalohy():
    zalohy_manager = ZalohyManager(EXCEL_BASE_PATH)
    advance_history = []

    if request.method == "POST":
        try:
            # Validace vstupních dat
            employee_name = request.form.get("employee_name")
            if not employee_name:
                raise ValueError("Jméno zaměstnance je povinné")

            amount_str = request.form.get("amount")
            try:
                amount = float(amount_str.replace(",", "."))
                if amount <= 0:
                    raise ValueError
            except (ValueError, AttributeError):
                raise ValueError("Částka musí být kladné číslo")

            currency = request.form.get("currency")
            if not currency in ["CZK", "EUR"]:
                raise ValueError("Neplatná měna")

            option = request.form.get("option")
            if not option:
                raise ValueError("Typ zálohy je povinný")

            date = request.form.get("date")
            try:
                datetime.strptime(date, "%Y-%m-%d")
            except ValueError:
                raise ValueError("Neplatný formát data")

            # Uložení zálohy
            zalohy_manager.add_or_update_employee_advance(
                employee_name=employee_name, amount=amount, currency=currency, option=option, date=date
            )
            flash("Záloha byla úspěšně uložena.", "success")

        except ValueError as e:
            flash(str(e), "error")
            logger.error(f"Chyba validace zálohy: {e}")
        except Exception as e:
            flash("Došlo k neočekávané chybě při ukládání zálohy.", "error")
            logger.error(f"Neočekávaná chyba při ukládání zálohy: {e}")

    # Načtení historie záloh
    try:
        hodiny2025_path = EXCEL_BASE_PATH / EXCEL_FILE_NAME_2025
        if not hodiny2025_path.exists():
            logger.warning(f"Soubor {EXCEL_FILE_NAME_2025} nebyl nalezen")
            flash(f"Soubor se záznamy záloh nebyl nalezen.", "warning")
        else:
            workbook_2025 = load_workbook(hodiny2025_path, read_only=True)
            if "Zalohy25" in workbook_2025.sheetnames:
                sheet = workbook_2025["Zalohy25"]
                data = list(sheet.values)
                if data:
                    keys = [str(k) for k in data[0] if k is not None]  # Konverze na string a odstranění None
                    advance_history = [
                        {k: v for k, v in zip(keys, row) if k is not None}
                        for row in data[1:]
                        if any(cell is not None for cell in row)
                    ]
            workbook_2025.close()

    except Exception as e:
        logger.error(f"Chyba při načítání historie záloh: {e}")
        flash("Chyba při načítání historie záloh.", "error")

    return render_template(
        "zalohy.html",
        employees=employee_manager.zamestnanci,
        options=excel_manager.get_advance_options(),
        current_date=datetime.now().strftime("%Y-%m-%d"),
        advance_history=advance_history,
    )


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, filename="app.log", filemode="a")
    app.run(debug=True)
