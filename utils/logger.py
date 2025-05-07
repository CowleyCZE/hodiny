import logging
import os
from logging.handlers import RotatingFileHandler
from pathlib import Path

from config import Config


def setup_logger(name):
    """
    Vytvoří a nakonfiguruje logger s konzistentním formátem
    """
    logger = logging.getLogger(name)

    if not logger.handlers:
        logger.setLevel(logging.DEBUG)

        # Vytvoření log adresáře v závislosti na prostředí
        if Config.IS_PYTHONANYWHERE:
            log_dir = Path(os.path.expanduser("~/hodiny/logs"))
        else:
            log_dir = Path("logs")
        log_dir.mkdir(exist_ok=True)

        # File handler s rotací - použití Path pro spojování cest
        file_handler = RotatingFileHandler(
            log_dir / f"{name}.log", maxBytes=1024 * 1024, backupCount=5, encoding="utf-8"  # 1MB
        )
        file_handler.setLevel(logging.INFO)

        # Console handler pouze pro lokální prostředí
        if not Config.IS_PYTHONANYWHERE:
            console_handler = logging.StreamHandler()
            console_handler.setLevel(logging.INFO)
            console_formatter = logging.Formatter(
                "%(asctime)s - %(name)s - %(levelname)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
            )
            console_handler.setFormatter(console_formatter)
            logger.addHandler(console_handler)

        # Formát pro file handler
        file_formatter = logging.Formatter(
            "%(asctime)s - %(name)s - %(levelname)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
        )
        file_handler.setFormatter(file_formatter)
        logger.addHandler(file_handler)

    return logger
