# config.py
import os
import secrets
from dataclasses import dataclass
from pathlib import Path
from typing import Optional
import logging
from openpyxl import Workbook


@dataclass
class ProjectConfig:
    name: str = ""
    start_date: str = ""
    end_date: str = ""


@dataclass
class TimeConfig:
    start_time: str = "07:00"
    end_time: str = "18:00"
    lunch_duration: float = 1.0


class Config:
    # Detekce prostředí
    IS_PYTHONANYWHERE = "PYTHONANYWHERE_SITE" in os.environ

    # Bezpečnostní nastavení
    SECRET_KEY = os.environ.get("SECRET_KEY") or secrets.token_hex(32)

    # Základní adresář - používá Path pro lepší přenositelnost
    if IS_PYTHONANYWHERE:
        BASE_DIR = Path(os.path.expanduser("~/hodiny"))
    else:
        BASE_DIR = Path(os.environ.get("HODINY_BASE_DIR", ".")).resolve()

    # Cesty - používají Path pro konzistentní práci s cestami
    DATA_PATH = Path(os.environ.get("HODINY_DATA_PATH", BASE_DIR / "data"))
    EXCEL_BASE_PATH = Path(os.environ.get("HODINY_EXCEL_PATH", BASE_DIR / "excel"))
    # Přejmenováno: Toto je nyní název šablony
    EXCEL_TEMPLATE_NAME = "Hodiny_Cap.xlsx"
    SETTINGS_FILE_PATH = Path(os.environ.get("HODINY_SETTINGS_PATH", DATA_PATH / "settings.json"))

    # Email konfigurace
    SMTP_SERVER = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
    SMTP_PORT = int(os.environ.get("SMTP_PORT", 465))
    SMTP_USERNAME = os.environ.get("SMTP_USERNAME")
    SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD")
    RECIPIENT_EMAIL = os.environ.get("RECIPIENT_EMAIL")

    # Výchozí konfigurace
    DEFAULT_PROJECT_CONFIG = ProjectConfig()
    DEFAULT_TIME_CONFIG = TimeConfig()

    GEMINI_API_KEY = "AIzaSyBvfpvviIHxJgOxkQeVyZCT2rnhyzI7bMo"  # Replace with your actual Gemini API key

    # Gemini API konfigurace
    GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "AIzaSyBvfpvviIHxJgOxkQeVyZCT2rnhyzI7bMo")
    GEMINI_API_URL = os.environ.get("GEMINI_API_URL", "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent")
    GEMINI_REQUEST_TIMEOUT = int(os.environ.get("GEMINI_REQUEST_TIMEOUT", 10))
    GEMINI_MAX_RETRIES = int(os.environ.get("GEMINI_MAX_RETRIES", 3))
    GEMINI_CACHE_TTL = int(os.environ.get("GEMINI_CACHE_TTL", 3600))  # 1 hodina v sekundách
    
    # Rate limiting
    RATE_LIMIT_REQUESTS = int(os.environ.get("RATE_LIMIT_REQUESTS", 100))  # počet požadavků
    RATE_LIMIT_WINDOW = int(os.environ.get("RATE_LIMIT_WINDOW", 3600))    # časové okno v sekundách

    @classmethod
    def get_default_settings(cls):
        """Vrátí kompletní výchozí nastavení, včetně klíče pro aktivní soubor"""
        return {
            "start_time": cls.DEFAULT_TIME_CONFIG.start_time,
            "end_time": cls.DEFAULT_TIME_CONFIG.end_time,
            "lunch_duration": cls.DEFAULT_TIME_CONFIG.lunch_duration,
            "project_info": vars(cls.DEFAULT_PROJECT_CONFIG),
            # Přidán klíč pro sledování aktivního souboru, defaultně None
            "active_excel_file": None,
        }

    @classmethod
    def init_app(cls, app):
        """Inicializace aplikace s konfigurací"""
        # Vytvoření potřebných adresářů
        cls.DATA_PATH.mkdir(parents=True, exist_ok=True)
        cls.EXCEL_BASE_PATH.mkdir(parents=True, exist_ok=True)

        # Zajistíme existenci šablony při inicializaci (volitelné)
        template_path = cls.EXCEL_BASE_PATH / cls.EXCEL_TEMPLATE_NAME
        if not template_path.exists():
             try:
                  wb = Workbook()
                  # Přidáme základní listy do šablony, pokud neexistuje
                  if "Sheet" in wb.sheetnames:
                       sheet = wb["Sheet"]
                       sheet.title = "Týden"
                  else:
                       wb.create_sheet("Týden")
                  if "Zálohy" not in wb.sheetnames:
                      wb.create_sheet("Zálohy")
                      # Můžeme přidat i výchozí hlavičky nebo hodnoty do šablony zde
                      zalohy_sheet = wb["Zálohy"]
                      zalohy_sheet["B80"] = "Option 1"
                      zalohy_sheet["D80"] = "Option 2"
                      # Případně další buňky A79, C81, D81 atd.

                  wb.save(template_path)
                  wb.close()
                  logging.info(f"Vytvořena chybějící šablona Excel souboru: {template_path}")
             except Exception as e:
                  logging.error(f"Nepodařilo se vytvořit chybějící šablonu {template_path}: {e}", exc_info=True)


        # Nastavení logovacího adresáře
        log_dir = cls.BASE_DIR / "logs"
        log_dir.mkdir(exist_ok=True)

        # Nastavení Flask aplikace
        app.config["SECRET_KEY"] = cls.SECRET_KEY
        # UPLOAD_FOLDER se typicky používá pro nahrávání souborů uživatelem,
        # zde ho necháváme, ale pro ukládání dat používáme EXCEL_BASE_PATH
        app.config["UPLOAD_FOLDER"] = str(cls.EXCEL_BASE_PATH)

        if cls.IS_PYTHONANYWHERE:
            app.config["PROPAGATE_EXCEPTIONS"] = True
            app.config["PREFERRED_URL_SCHEME"] = "https"

        # Nastavení pro produkci
        if not app.debug:
            app.config["SESSION_COOKIE_SECURE"] = True
            app.config["SESSION_COOKIE_HTTPONLY"] = True
            app.config["PERMANENT_SESSION_LIFETIME"] = 3600  # 1 hodina

        return app
