# Status a Inicializace Projektu `hodiny`

**Datum inicializace:** 28. září 2026  
**Účel souboru:** Přehled stavu projektu, architektury a testů po provedené inicializaci.

---

## 1. Přehled Projektu

Aplikace **Hodiny** je Flask webová aplikace určená k evidenci docházky, odpracovaných hodin, záloh/plateb zaměstnanců a generování výkazů přímo v Excel souborech (`.xlsx`), například `Hodiny2026.xlsx` nebo `Hodiny_Cap.xlsx`.

---

## 2. Technologický Stack a Architektura

- **Backend:** Python 3, Flask (strukturováno přes Flask Blueprints)
- **Excel Engine:** `openpyxl` s vláknově bezpečným zamykáním souborů v [excel_manager.py](file:///data/data/com.termux/files/home/projects/hodiny/excel_manager.py)
- **Frontend:** HTML5, Jinja2 šablony, Vanilla JavaScript & CSS
- **Testování:** `pytest`
- **Konfigurace a data:** `data/employee_config.json`, `config.json`, `data/settings.json`

---

## 3. Hlavní Vstupní Body a Moduly

- **Vstupní bod:** [app.py](file:///data/data/com.termux/files/home/projects/hodiny/app.py) & [wsgi.py](file:///data/data/com.termux/files/home/projects/hodiny/wsgi.py)
- **Blueprinty (`blueprints/`):**
  - `main.py`: Ruční zápis, docházka, hlasové/textové NLP příkazy, e-maily
  - `employees.py`: Správa zaměstnanců (CRUD)
  - `settings.py` & `configuration.py`: Správa nastavení a Excel mapování
  - `reports.py`: Výkazy, měsíční přehledy, zálohy
- **Jádro Excel logiky:** [excel_manager.py](file:///data/data/com.termux/files/home/projects/hodiny/excel_manager.py) & [hodiny2025_manager.py](file:///data/data/com.termux/files/home/projects/hodiny/hodiny2025_manager.py)
- **NLP / Hlasový procesor:** [utils/voice_processor.py](file:///data/data/com.termux/files/home/projects/hodiny/utils/voice_processor.py)

---

## 4. Výsledek Inicializace a Kontroly

- **Struktura adresářů:** Ověřena a kompletní.
- **Git stav:** Všechny soubory jsou sledovány a sync s `origin/main`.
- **Testovací sada:** Spuštěna testovací sada `pytest`.
