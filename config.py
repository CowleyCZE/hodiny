# config.py
import os
import secrets
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

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
    
    # Gemini API konfigurace
    GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "AIzaSyBvfpvviIHxJgOxkQeVyZCT2rnhyzI7bMo")
    
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
    
    @classmethod
    def get_default_settings(cls):
        """Vrátí kompletní výchozí nastavení, včetně klíče pro aktivní soubor"""
        return {
            "start_time": cls.DEFAULT_TIME_CONFIG.start_time,
            "end_time": cls.DEFAULT_TIME_CONFIG.end_time,
            "lunch_duration": cls.DEFAULT_TIME_CONFIG.lunch_duration,
            "project_info": {
                "name": cls.DEFAULT_PROJECT_CONFIG.name,
                "start_date": cls.DEFAULT_PROJECT_CONFIG.start_date,
                "end_date": cls.DEFAULT_PROJECT_CONFIG.end_date
            }
        }
    
    @classmethod
    def init_app(cls, app):
        """Inicializace aplikace s konfigurací"""
        app.config["ENV"] = os.getenv("FLASK_ENV", "production")
        app.config["DEBUG"] = os.getenv("FLASK_DEBUG", "False").lower() in ["true", "1", "t"]
        
        # Nastavení cest do Flask konfigurace
        app.config["DATA_PATH"] = cls.DATA_PATH
        app.config["EXCEL_BASE_PATH"] = cls.EXCEL_BASE_PATH
        app.config["SETTINGS_FILE_PATH"] = cls.SETTINGS_FILE_PATH
        
        # Nastavení pro produkci
        if not app.config["DEBUG"]:
            app.config["SESSION_COOKIE_SECURE"] = True
            app.config["SESSION_COOKIE_HTTPONLY"] = True
            app.config["PERMANENT_SESSION_LIFETIME"] = 3600  # 1 hodina

# Nastavení pro WSGI (v produkčním prostředí)
if "PYTHONANYWHERE_SITE" in os.environ:
    Config.IS_PYTHONANYWHERE = True
    Config.BASE_DIR = Path(os.path.expanduser("~/hodiny"))
    Config.DATA_PATH = Config.BASE_DIR / "data"
    Config.EXCEL_BASE_PATH = Config.BASE_DIR / "excel"
    Config.SETTINGS_FILE_PATH = Config.DATA_PATH / "settings.json"

# Nastavení vývojového prostředí
else:
    Config.IS_PYTHONANYWHERE = False
    Config.BASE_DIR = Path(os.getenv("HODINY_BASE_DIR", ".")).resolve()
    Config.DATA_PATH = Config.BASE_DIR / "data"
    Config.EXCEL_BASE_PATH = Config.BASE_DIR / "excel"
    Config.SETTINGS_FILE_PATH = Config.DATA_PATH / "settings.json"
