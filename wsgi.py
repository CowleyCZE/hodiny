import os
import sys
from pathlib import Path

# Nastavení cesty k aplikaci
try:
    # Pokud jsme na PythonAnywhere
    if 'PYTHONANYWHERE_SITE' in os.environ:
        path = os.path.expanduser('~/hodiny')
    else:
        # Lokální vývoj - použijeme aktuální adresář
        path = os.path.dirname(os.path.abspath(__file__))
    
    if path not in sys.path:
        sys.path.insert(0, path)
except Exception as e:
    print(f"Chyba při nastavování cesty: {e}")
    sys.exit(1)

try:
    from app import app as application
    from config import Config
    
    # Nastavení produkčního prostředí
    application.config['ENV'] = 'production'
    application.config['DEBUG'] = False
    
    # Inicializace aplikace
    Config.init_app(application)
    
    # Vytvoření potřebných adresářů
    base_dir = Path(path)
    (base_dir / 'data').mkdir(parents=True, exist_ok=True)
    (base_dir / 'excel').mkdir(parents=True, exist_ok=True)
    (base_dir / 'logs').mkdir(parents=True, exist_ok=True)
    
except Exception as e:
    print(f"Chyba při inicializaci aplikace: {e}")
    sys.exit(1)

# Pouze pro lokální vývoj
if __name__ == '__main__':
    application.run() 