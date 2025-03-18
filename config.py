import os

class Config:
    # Bezpečnostní nastavení
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'vygeneruj-bezpecny-klic'
    
    # Cesty
    DATA_PATH = '/home/Cowley/hodiny/data'
    EXCEL_BASE_PATH = '/home/Cowley/hodiny/excel'
    EXCEL_FILE_NAME = 'Hodiny_Cap.xlsx'
    EXCEL_FILE_NAME_2025 = 'Hodiny2025.xlsx'
    SETTINGS_FILE_PATH = '/home/Cowley/hodiny/data/settings.json'
    
    # Email konfigurace
    SMTP_SERVER = 'smtp.gmail.com'
    SMTP_PORT = 465
    SMTP_USERNAME = os.environ.get('SMTP_USERNAME')
    SMTP_PASSWORD = os.environ.get('SMTP_PASSWORD')
    RECIPIENT_EMAIL = 'cowleyy@gmail.com'
