import logging
import os
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path


# Nastavení základního loggeru pro WSGI
def setup_wsgi_logger():
    logger = logging.getLogger("wsgi")
    if not logger.handlers:
        logger.setLevel(logging.INFO)

        # Vytvoření log adresáře
        if "PYTHONANYWHERE_SITE" in os.environ:
            log_dir = Path(os.path.expanduser("~/hodiny/logs"))
        else:
            log_dir = Path("logs")
        log_dir.mkdir(parents=True, exist_ok=True)

        # Nastavení handleru s rotací
        handler = RotatingFileHandler(log_dir / "wsgi.log", maxBytes=1024 * 1024, backupCount=5, encoding="utf-8")
        formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    return logger


logger = setup_wsgi_logger()

# Nastavení cesty k aplikaci
try:
    # Pokud jsme na PythonAnywhere
    if "PYTHONANYWHERE_SITE" in os.environ:
        path = os.path.expanduser("~/hodiny")
        logger.info(f"Běžíme na PythonAnywhere, cesta: {path}")
    else:
        # Lokální vývoj - použijeme aktuální adresář
        path = os.path.dirname(os.path.abspath(__file__))
        logger.info(f"Lokální vývoj, cesta: {path}")

    if path not in sys.path:
        sys.path.insert(0, path)
        logger.info(f"Přidána cesta do sys.path: {path}")
except Exception as e:
    logger.error(f"Chyba při nastavování cesty: {e}")
    sys.exit(1)

try:
    from app import app as application
    from config import Config

    # Nastavení produkčního prostředí
    application.config["ENV"] = "production"
    application.config["DEBUG"] = False

    # Inicializace aplikace
    Config.init_app(application)
    logger.info("Aplikace úspěšně inicializována")

    # Vytvoření potřebných adresářů
    base_dir = Path(path)
    (base_dir / "data").mkdir(parents=True, exist_ok=True)
    (base_dir / "excel").mkdir(parents=True, exist_ok=True)
    (base_dir / "logs").mkdir(parents=True, exist_ok=True)
    logger.info("Vytvořeny potřebné adresáře")

except Exception as e:
    logger.error(f"Chyba při inicializaci aplikace: {e}")
    sys.exit(1)

# Pouze pro lokální vývoj
if __name__ == "__main__":
    application.run()
