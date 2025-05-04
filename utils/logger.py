# utils/logger.py
import logging
import os
from logging.handlers import RotatingFileHandler
from pathlib import Path

from config import Config


def setup_logger(name: str) -> logging.Logger:
    """
    Vytvoří a nakonfiguruje logger s konzistentním formátem
    
    Args:
        name (str): Název loggeru (typicky __name__)
        
    Returns:
        logging.Logger: Nakonfigurovaný logger
    """
    logger = logging.getLogger(name)
    
    # Ujistíme se, že nejsou přidány duplicitní handlery
    if logger.hasHandlers():
        return logger
    
    # Nastavení úrovně logování
    logger.setLevel(logging.INFO)
    
    # Určení log adresáře podle prostředí
    try:
        if Config.IS_PYTHONANYWHERE:
            log_dir = Path(os.path.expanduser("~/hodiny/logs"))
        else:
            log_dir = Path("logs").resolve()
            
        # Vytvoření adresáře pro logy pokud neexistuje
        log_dir.mkdir(parents=True, exist_ok=True)
        
    except Exception as e:
        # Fallback - použijeme aktuální adresář pokud selže vytvoření log adresáře
        log_dir = Path.cwd()
        logger.warning(f"Nepodařilo se vytvořit cílový log adresář: {e}. Používám {log_dir} jako náhradní.")

    # Konfigurace handlerů
    try:
        # File handler s rotací (max 1MB, 5 záloh)
        file_handler = RotatingFileHandler(
            log_dir / f"{name}.log",
            maxBytes=1024 * 1024,  # 1MB
            backupCount=5,
            encoding="utf-8"
        )
        file_handler.setLevel(logging.INFO)
        
        # Formát pro logování do souboru
        file_formatter = logging.Formatter(
            "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )
        file_handler.setFormatter(file_formatter)
        logger.addHandler(file_handler)

        # Console handler pouze pro lokální prostředí
        if not Config.IS_PYTHONANYWHERE:
            console_handler = logging.StreamHandler()
            console_handler.setLevel(logging.DEBUG)
            
            # Barevné logování pro konzoli (pokud je dostupné)
            try:
                from colorlog import ColoredFormatter
                
                console_formatter = ColoredFormatter(
                    "%(asctime)s - %(name)s - %(log_color)s%(levelname)s%(reset)s - %(message)s",
                    datefmt="%Y-%m-%d %H:%M:%S",
                    log_colors={
                        'DEBUG': 'cyan',
                        'INFO': 'green',
                        'WARNING': 'yellow',
                        'ERROR': 'red',
                        'CRITICAL': 'red,bg_white'
                    }
                )
            except ImportError:
                # Záložní formatter bez barev
                console_formatter = logging.Formatter(
                    "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
                    datefmt="%Y-%m-%d %H:%M:%S"
                )
            
            console_handler.setFormatter(console_formatter)
            console_handler.setLevel(logging.DEBUG)
            logger.addHandler(console_handler)
            
    except Exception as e:
        # Fallback - minimální logování do konzole při kritické chybě
        logger.addHandler(logging.StreamHandler())
        logger.setLevel(logging.DEBUG)
        logger.warning(f"Chyba při plné konfiguraci loggeru: {e}. Používám základní nastavení.")
    
    return logger
