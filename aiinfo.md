# aiinfo.md – Technický souhrn projektu Hodiny pro AI agenty

Tento dokument slouží jako rychlý, vysoce koncentrovaný technický briefing pro AI asistenty. Šetří tokeny i čas potřebný pro analýzu repozitáře.

---

## 1. Základní přehled a účel
* **Projekt:** `hodiny` (Git repozitář: `CowleyCZE/hodiny`, větev: `main`)
* **Prostředí:**
  * Lokální vývoj: Android / Termux (`/data/data/com.termux/files/home/projekty/hodiny`)
  * Produkční server: PythonAnywhere (`https://cowley.pythonanywhere.com/`)
* **Architektura:** Server-rendered Flask webová aplikace s REST API endpointy (`/api/v1/*`), SQLite databází pro audit a synchronizaci a integrací na Excel (`openpyxl`).

---

## 2. Klíčové soubory a jejich role
* `app.py`: Inicializace Flask aplikace, middleware, inicializace kontextu `g` (databáze, managerů) a webhook.
* `wsgi.py`: WSGI entrypoint pro PythonAnywhere / Gunicorn.
* `config.py`: Globální konfigurace, cesty (`EXCEL_PATH`, `BACKUP_PATH`, `DATA_PATH` atd.).
* `config.json`: Mapování pojmenovaných buněk v Excel listech (např. souřadnice projektů, buněk pro dny).
* `data/settings.json`: Runtime konfigurace – název projektu, začátek akce, výchozí pracovní doba, preferovaný zaměstnanec (v `.gitignore`).
* `data/employee_config.json`: Seznam všech zaměstnanců a vybraných zaměstnanců (v `.gitignore`).
* `data/hodiny.sqlite3`: Lokální SQLite databáze (uživatelé, audit, stavy docházky).

### Blueprinty (`blueprints/`):
* `main.py`: Přehled (dashboard), rychlý zápis hodin, zápis volných dnů, hlasové/textové příkazy docházky.
* `reports.py`: Měsíční reporty, kalendář, výdaje/zálohy (`/zalohy`), zálohy souborů (`/backups`), kontroly a statistiky.
* `employees.py`: Správa zaměstnanců (přidání, aktivace/deaktivace, pořadí).
* `excel.py`: Prohlížeč, upload/download a editace tabulek.
* `settings.py`: Nastavení projektu, časů a preferencí.
* `auth.py`: Autentizace, role (admin/manager/viewer), správa uživatelů.
* `api_endpoints.py`: REST API `/api/v1/` (`/health`, `/employees`, `/settings`, `/time-entries`, `/excel/status`, `/sync/*`).

### Servisní vrstva (`services/`):
* `excel_week_service.py`: Generování týdenních kopií a zápis docházky do týdenních listů. Obsahuje `copy_worksheet_deep()` pro plné zachování stylů, rozměrů, sloučení, freeze panes a pohledů.
* `phase2_services.py`: Správa XLSX záloh (`create_backup`, `restore_backup`, `delete_backup`, `delete_all_backups`), audit log (`audit.jsonl`), statistiky a detekce chyb v datech.
* `employee_management.py`: Správa a řazení zaměstnanců (`EmployeeManager`), preferovaný zaměstnanec je vždy první.
* `zalohy_manager.py`: Správa zápisu záloh/půjček do listu `Zálohy` (EUR/CZK).
* `sync_service.py`: Synchronizace stavu mezi SQLite databází a Excelem.

---

## 3. Autoritativní Excel model a mapování
* **Aktivní runtime soubor:** `excel/Hodiny_Cap.xlsx`
* **Šablona pro nové týdny:** List s názvem `Týden` (případně šablona definovaná v `Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME`).
* **Mapování buněk v týdenním listu:**
  * **Číslo týdne:** Buňka `B3` (sloučená oblast `B3:D3`, formát `Týden XX`). Buňka `C3` je vnitřek sloučení, nesmí se zapisovat přímo do ní!
  * **Popisek projektu:** `B4` (`NÁZEV PROJEKTU :`)
  * **Název projektu / akce:** Buňka `I4` (sloučená oblast `I4:P4`)
  * **Datumy dnů:** Řádek `6`, sloupce po dvou: Po (`B6`), Út (`D6`), St (`F6`), Čt (`H6`), Pá (`J6`), So (`L6`), Ne (`N6`). Formát text `DD.MM.YYYY`.
  * **Časy od/do:** Řádek `7`, Po (`B7`/`C7`), Út (`D7`/`E7`), St (`F7`/`G7`), Čt (`H7`/`I7`), Pá (`J7`/`K7`), So (`L7`/`M7`), Ne (`N7`/`O7`). Formát `h:mm`.
  * **Zaměstnanci a hodiny:** Řádky `8` až `23`. Jméno ve sloupci `A`. Hodiny v odpovídajících dnech: `B`, `D`, `F`, `H`, `J`, `L`, `N`. Celkový součet týdne ve sloupci `P`.

---

## 4. Důležitá pravidla a konvence
1. **Zákaz rozbíjení Excel formátování:**
   * Při kopírování listů vždy používat `copy_worksheet_deep()` nebo rozšířenou logiku, která přenáší nejen hodnoty a styly buněk, ale i rozměry řádků/sloupců, sloučení (`merged_cells`), ukotvení příček (`freeze_panes`) a mřížku (`showGridLines`).
2. **Git a citlivá data:**
   * Soubory v `data/settings.json`, `data/employee_config.json` a `data/*.sqlite3` jsou v `.gitignore` (obsahují lokální / produkční runtime stavy).
   * Produkční Excel soubory (`.xlsx`) jsou standardně ignorovány, přidávají se pouze úmyslně přes `git add -f` (např. šablona `Hodiny_Cap.xlsx`).
   * Konvence commit zpráv: **Conventional Commits 1.0.0** (`feat:`, `fix:`, `style:`, `refactor:`, `test:`).
3. **Synchronizace s PythonAnywhere:**
   * API na `https://cowley.pythonanywhere.com/api/v1/` umožňuje bezpečně číst aktuální `settings` i `employees` přímo z běžící instance.
