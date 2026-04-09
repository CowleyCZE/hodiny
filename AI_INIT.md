# AI Init: projekt `hodiny`

Tento soubor slouží jako rychlý technický briefing pro AI model nebo nového vývojáře. Cílem je, aby po jeho přečtení bylo jasné:

- k čemu aplikace slouží
- jak je napsaná
- kde jsou hlavní vstupní body
- jak funguje zápis do Excelu
- které soubory jsou důležité
- na co si dát pozor při změnách

## 1. Účel aplikace

Projekt `hodiny` je Flask webová aplikace pro evidenci pracovní doby do Excelu.

Hlavní use-cases:

- zapisovat pracovní dobu vybraným zaměstnancům
- zapisovat volné dny
- spravovat zaměstnance a výběr zaměstnanců pro zápis
- zapisovat zálohy / půjčky / platby do listu `Zálohy`
- generovat měsíční report
- prohlížet a editovat Excel soubory v browseru
- poslat aktivní Excel e-mailem
- zpracovat textový hlasový příkaz typu `zapiš práci dnes od 7:00 do 15:00`

## 2. Technologický stack

- Python
- Flask
- Jinja2 šablony
- openpyxl pro práci s Excel `.xlsx`
- pytest pro testy
- flake8 pro statickou kontrolu

Nejde o SPA ani o REST-only backend. Je to server-rendered Flask aplikace s několika JSON endpointy.

## 3. Jak aplikaci spustit

Z kořene projektu:

```bash
cd /home/cowley/Dokumenty/projekty/hodiny
.venv/bin/python app.py
```

Otevřít:

```text
http://127.0.0.1:5000/
```

Testy:

```bash
.venv/bin/pytest -q
```

Lint:

```bash
.venv/bin/flake8 app.py blueprints services utils *.py
```

## 4. Hlavní vstupní body

### Bootstrap

- `app.py`
  - vytváří Flask app
  - registruje blueprinty
  - v `before_request` připravuje managery
  - načítá runtime nastavení ze `data/settings.json`

- `wsgi.py`
  - produkční vstupní bod

### Blueprinty

- `blueprints/main.py`
  - homepage `/`
  - ruční zápis `/zaznam`
  - `POST /api/quick_time_entry`
  - `POST /voice-command`
  - `POST /send_email`

- `blueprints/settings.py`
  - `/settings`
  - běžné aplikační nastavení

- `blueprints/configuration.py`
  - `/nastaveni`
  - technická Excel konfigurace
  - `/api/settings`

- `blueprints/employees.py`
  - `/zamestnanci`

- `blueprints/reports.py`
  - `/zalohy`
  - `/monthly_report`

- `blueprints/excel.py`
  - upload
  - download
  - viewer
  - editor

### API vrstva

- `api_endpoints.py`
  - `/api/v1/*`
  - health, employees, settings, time-entry, excel status

## 5. Důležité service a manager vrstvy

### Runtime nastavení

- `services/settings_service.py`
  - načítá a ukládá `data/settings.json`
  - drží normalizovanou strukturu aplikačního nastavení

Aktuálně runtime nastavení obsahuje mimo jiné:

- `start_time`
- `end_time`
- `lunch_duration`
- `preferred_employee_name`
- `project_info.name`
- `project_info.start_date`
- `project_info.end_date`
- `last_archived_week`

### Zaměstnanci

- `employee_management.py`
  - spravuje `data/employee_config.json`
  - drží seznam všech zaměstnanců
  - drží seznam vybraných zaměstnanců
  - respektuje `preferred_employee_name`

Důležité pravidlo:

- pokud `preferred_employee_name` existuje v seznamu zaměstnanců, je vždy řazen první
- to platí pro:
  - seznam vybraných zaměstnanců pro docházku
  - seznam zaměstnanců v zálohách
  - přehledy založené na `EmployeeManager`

### Docházka a dashboard

- `services/main_service.py`
  - staví data pro homepage
  - centralizuje zápis pracovní doby
  - posílá aktivní Excel e-mailem

Poznámka:

- `save_time_entry()` je autoritativní cesta pro zapisování času z webu i z hlasového flow
- musí zůstat konzistentní s `ExcelManagerem`

### Excel doména

- `excel_manager.py`
  - hlavní orchestrátor zápisu do aktivního Excelu
  - pracuje s hlavním souborem i týdenními kopiemi
  - synchronizuje data i do `Hodiny2025.xlsx`

- `services/excel_week_service.py`
  - zápis docházky do týdenních listů

- `services/excel_report_service.py`
  - měsíční agregace

- `services/excel_config_service.py`
  - čtení technického mappingu z `config.json`

- `services/excel_metadata_service.py`
  - metadata Excel souborů

### Zálohy

- `zalohy_manager.py`
  - zapisuje zálohy do listu `Zálohy`
  - podporuje EUR/CZK
  - vybírá sloupec podle option + měny

## 6. Autoritativní Excel model

Nejdůležitější soubory:

- autoritativní šablona: `excel/Hodiny_Cap:vzor.xlsx`
- runtime aktivní soubor: `excel/Hodiny_Cap.xlsx`

Pokud `Hodiny_Cap.xlsx` chybí, vytváří se kopií autoritativní šablony.

### Týdenní evidence

Autoritativní layout je popsán v:

- `AUTORITATIVNI_SABLONA_HODINY_CAP.md`

Zjednodušeně:

- zaměstnanec: `A8` a níže
- pondělí: `B/C`
- úterý: `D/E`
- středa: `F/G`
- čtvrtek: `H/I`
- pátek: `J/K`
- sobota: `L/M`
- neděle: `N/O`

Význam:

- datum: řádek `80`
- začátek směny: řádek `7`, první sloupec páru
- konec směny: řádek `7`, druhý sloupec páru
- čisté hodiny zaměstnance: první sloupec páru v řádku zaměstnance

Příklad pro čtvrtek:

- datum: `H80`
- start: `H7`
- konec: `I7`
- hodiny zaměstnance: `H8`, `H9`, ...

### Zálohy

List:

- `Zálohy`

Mapování:

- zaměstnanec: `A8` a níže
- option popisky: `B80`, `D80`, `F80`, `H80`
- EUR: `B/D/F/H`
- CZK: `C/E/G/I`
- datum posledního zápisu: `Z`

## 7. Důležité datové soubory

### Verzované nebo očekávané v repu

- `config.json`
  - technické mapování buněk
- `excel/Hodiny_Cap:vzor.xlsx`
  - autoritativní šablona
- `README.md`
- `AUTORITATIVNI_SABLONA_HODINY_CAP.md`

### Runtime data

- `data/settings.json`
  - běžné nastavení aplikace
- `data/employee_config.json`
  - zaměstnanci a vybraní zaměstnanci
- `excel/Hodiny_Cap.xlsx`
  - aktivní soubor
- `excel/Hodiny_Cap_Tyden*.xlsx`
  - týdenní kopie / archiv
- `excel/Hodiny2025.xlsx`
  - měsíční souhrnný workbook
- `logs/*.log`

## 8. Důležité funkce aplikace

### Homepage `/`

Obsahuje:

- přehled projektu
- rychlý zápis času
- týdenní preview
- upload
- e-mail
- hlasové ovládání

Quick-entry jde přes:

- frontend JS v `templates/index.html`
- backend endpoint `POST /api/quick_time_entry`

### Ruční zápis `/zaznam`

- zapisuje čas pro všechny vybrané zaměstnance
- používá pořadí z `EmployeeManager.get_vybrani_zamestnanci()`

### Zálohy `/zalohy`

- formulář vybírá zaměstnance z `EmployeeManager.get_employee_names()`
- preferované jméno je tedy první i zde

### Hlasové ovládání `/voice-command`

- přijímá JSON s textovým příkazem
- zpracování entity logiky je v `utils/voice_processor.py`
- finální zápis jde přes stejnou service logiku jako běžný zápis času

## 9. Známé architektonické zásady

### 1. Nepřepisovat Excel logiku bokem

Pokud je potřeba zapisovat pracovní dobu:

- používat `save_time_entry()` a `ExcelManager`
- nepsat paralelní zápis do `Hodiny2025.xlsx` mimo hlavní flow

To už v minulosti způsobilo nekonzistenci a chyby.

### 2. `settings.json` a `config.json` jsou dvě různé věci

- `settings.json` = běžné runtime nastavení aplikace
- `config.json` = technický mapping Excel buněk

Tyto dvě vrstvy nesmí být směšovány.

### 3. Autoritativní šablona je závazná

- Excel layout se má řídit podle `Hodiny_Cap:vzor.xlsx`
- fallbacky v kódu mají být jen opatrná kompatibilita, ne nová pravda

### 4. Preferované jméno je uživatelské nastavení, ne hardcoded pravidlo

Historicky bylo v projektu natvrdo preferováno jedno konkrétní jméno.

Aktuální správné chování:

- priorita jména je určena přes `preferred_employee_name` v nastavení
- žádné jméno nesmí být znovu natvrdo zakódováno v business logice

## 10. Kde začít při další práci

Pokud AI model přichází k projektu poprvé, doporučené pořadí čtení je:

1. `AI_INIT.md`
2. `README.md`
3. `AUTORITATIVNI_SABLONA_HODINY_CAP.md`
4. `app.py`
5. relevantní blueprint
6. relevantní service/manager
7. související test

Příklad:

- problém v zálohách:
  - `blueprints/reports.py`
  - `zalohy_manager.py`
  - `templates/zalohy.html`
  - `test_zalohy_manager.py`
  - `test_domain_routes.py`

- problém v docházce:
  - `blueprints/main.py`
  - `services/main_service.py`
  - `excel_manager.py`
  - `services/excel_week_service.py`
  - `test_main_routes.py`
  - `test_weekly_files.py`

## 11. Jak dělat bezpečné změny

- nejdřív ověřit, jestli už neexistuje service vrstva pro daný problém
- nepsat novou business logiku přímo do route handleru, pokud už existuje service
- zachovat konzistenci mezi frontendem, backendem a testy
- po změně pustit:

```bash
.venv/bin/pytest -q
.venv/bin/flake8 app.py blueprints services utils *.py
```

## 12. Aktuální kvalitativní stav

V době vytvoření tohoto souboru:

- testy procházejí
- `flake8` prochází
- blueprint architektura je zavedená
- hlavní Excel zápisy jsou sladěné s autoritativní šablonou
- quick-entry i voice-command používají stejnou zápisovou logiku
- preferované uživatelské jméno se propisuje do pořadí seznamů zaměstnanců

## 13. Související dokumenty

- `README.md`
- `AUTORITATIVNI_SABLONA_HODINY_CAP.md`
- `PROJEKT_POTREBY_A_REFAKTOR.md`

