import os
import secrets
from dataclasses import dataclass
from typing import Optional
from pathlib import Path

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
    IS_PYTHONANYWHERE = 'PYTHONANYWHERE_SITE' in os.environ
    
    # Bezpečnostní nastavení
    SECRET_KEY = os.environ.get('SECRET_KEY') or secrets.token_hex(32)
    
    # Základní adresář - používá Path pro lepší přenositelnost
    if IS_PYTHONANYWHERE:
        BASE_DIR = Path(os.path.expanduser('~/hodiny'))
    else:
        BASE_DIR = Path(os.environ.get('HODINY_BASE_DIR', '.')).resolve()
    
    # Cesty - používají Path pro konzistentní práci s cestami
    DATA_PATH = Path(os.environ.get('HODINY_DATA_PATH', BASE_DIR / 'data'))
    EXCEL_BASE_PATH = Path(os.environ.get('HODINY_EXCEL_PATH', BASE_DIR / 'excel'))
    EXCEL_FILE_NAME = 'Hodiny_Cap.xlsx'
    EXCEL_FILE_NAME_2025 = 'Hodiny2025.xlsx'
    SETTINGS_FILE_PATH = Path(os.environ.get('HODINY_SETTINGS_PATH', DATA_PATH / 'settings.json'))
    
    # Email konfigurace
    SMTP_SERVER = os.environ.get('SMTP_SERVER', 'smtp.gmail.com')
    SMTP_PORT = int(os.environ.get('SMTP_PORT', 465))
    SMTP_USERNAME = os.environ.get('SMTP_USERNAME')
    SMTP_PASSWORD = os.environ.get('SMTP_PASSWORD')
    RECIPIENT_EMAIL = os.environ.get('RECIPIENT_EMAIL')

    # Výchozí konfigurace
    DEFAULT_PROJECT_CONFIG = ProjectConfig()
    DEFAULT_TIME_CONFIG = TimeConfig()
    
    @classmethod
    def get_default_settings(cls):
        """Vrátí kompletní výchozí nastavení"""
        return {
            'start_time': cls.DEFAULT_TIME_CONFIG.start_time,
            'end_time': cls.DEFAULT_TIME_CONFIG.end_time,
            'lunch_duration': cls.DEFAULT_TIME_CONFIG.lunch_duration,
            'project_info': vars(cls.DEFAULT_PROJECT_CONFIG)
        }

    @classmethod
    def init_app(cls, app):
        """Inicializace aplikace s konfigurací"""
        # Vytvoření potřebných adresářů
        cls.DATA_PATH.mkdir(parents=True, exist_ok=True)
        cls.EXCEL_BASE_PATH.mkdir(parents=True, exist_ok=True)
        
        # Nastavení logovacího adresáře
        log_dir = cls.BASE_DIR / 'logs'
        log_dir.mkdir(exist_ok=True)
        
        # Nastavení Flask aplikace
        app.config['SECRET_KEY'] = cls.SECRET_KEY
        app.config['UPLOAD_FOLDER'] = str(cls.EXCEL_BASE_PATH)
        
        if cls.IS_PYTHONANYWHERE:
            app.config['PROPAGATE_EXCEPTIONS'] = True
            app.config['PREFERRED_URL_SCHEME'] = 'https'
            
        # Nastavení pro produkci
        if not app.debug:
            app.config['SESSION_COOKIE_SECURE'] = True
            app.config['SESSION_COOKIE_HTTPONLY'] = True
            app.config['PERMANENT_SESSION_LIFETIME'] = 3600  # 1 hodina
            
        return app
