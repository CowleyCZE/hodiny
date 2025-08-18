# Hodiny – evidence pracovní doby

Kompletní webová aplikace ve Flasku pro evidenci pracovní doby do Excelu, správu zaměstnanců, správu záloh (zálohy/výplaty), generování měsíčních přehledů, náhled Excel souborů a základní „hlasové“ ovládání přes textové příkazy. Aplikace pracuje s jedním „aktivním“ Excel souborem (kopie šablony), který se průběžně doplňuje a lze ho stáhnout nebo poslat e‑mailem.

## Hlavní schopnosti

- Aktivní Excel soubor na bázi šablony Hodiny_Cap.xlsx (automaticky vytvořen, pokud chybí)
- Záznam pracovní doby/volna pro vybrané zaměstnance po dnech (počítání čistých hodin − pauza)
- Správa zaměstnanců (přidání, úprava, smazání, výběr pro záznam)
- Zálohy (příspěvky) se sumací po „možnostech“ a měnách (EUR/CZK)
- Měsíční report souhrnů hodin a počtu volných dní
- Náhled obsahu aktivního Excel souboru (read‑only, s omezením řádků)
- Stažení a odeslání aktivního souboru e‑mailem (SMTP)
- Přejmenování/smazání projektových souborů, archivace a založení nového aktivního souboru
- Textové „hlasové“ příkazy pro záznam času a rychlé statistiky
- Pružné logování do souborů s rotací
- **🆕 Editor Excel souborů – úprava buněk přímo v prohlížeči s automatickým uložením**
- **🆕 Responzivní design pro všechny velikosti obrazovek**

## Architektura a soubory

- app.py – Flask aplikace, routy, práce se session, předzpracování/úklid, integrace Excel/zalohy managerů a hlasových příkazů
- config.py – konfigurace cest, šablon, validací, SMTP a „Gemini“ voleb; vytváří chybějící šablonu a složky
- excel_manager.py – zápis/čtení do aktivního Excelu (listy „Týden N“, časy a hodiny, projektové info, měsíční report)
- employee_management.py – správa zaměstnanců + „vybraných“; trvalé uložení v data/employee_config.json
- zalohy_manager.py – práce s listem „Zálohy“ v Excelu (částky, měna, opce, datum)
- utils/logger.py – jednotné logování (RotatingFileHandler do logs/ + konzole lokálně)
- utils/voice_processor.py – zpracování textových příkazů (regex entity), skeleton pro volání Gemini API, rate‑limit a cache
- templates/* – Jinja2 šablony (index, záznam, zaměstnanci, zálohy, nastavení, náhled Excelu, měsíční report)
- static/* – CSS/JS pro UI a podporu akcí (hlasové ovládání, potvrzení, správa zaměstnanců)
- wsgi.py – WSGI vstupní bod (např. pro Gunicorn / PythonAnywhere)
- data/ – konfigurace a uživatelská nastavení (settings.json, employee_config.json)
- excel/ – šablona a generované/aktivní Excel soubory
- logs/ – aplikační logy (rotace)

## Datový model v Excelu (zjednodušeně)

- Šablona: excel/Hodiny_Cap.xlsx
- Týdenní listy: „Týden {cislo}“; pokud chybí, kopíruje se z listu „Týden“ nebo se vytvoří prázdný
- Jména zaměstnanců: od řádku 9 ve sloupci A
- Pro každý pracovní den se zapisují:
  - začátek/konec do řádku 7, do odpovídajících sloupců dne
  - datum do řádku 80 (B80/D80/F80/H80/J80)
  - čisté hodiny (po odečtení pauzy) k řádku zaměstnance do sloupce daného dne
- Zálohy: list „Zálohy“, volby v buňkách B80/D80/F80/H80 (popisky možností), hodnoty po zaměstnancích ve sloupcích B..I, datum do sloupce Z

## Webové rozhraní (routy)

- GET / – přehled týdne/dne, akce pro stažení a e‑mail (pokud existuje aktivní soubor), tlačítko „hlasové“ ovládání
- POST /send_email – odešle aktivní Excel e‑mailem (SMTP_SSL)
- GET+POST /zamestnanci – správa zaměstnanců a výběru
- GET+POST /zaznam – formulář pro záznam pracovní doby/volna pro vybrané zaměstnance, výběr aktivního souboru
- POST /set_active_file – přepnutí aktivního souboru
- POST /rename_project – přejmenování existujícího excel souboru
- POST /delete_project – smazání excel souboru (neaktivního)
- GET /excel_viewer – read‑only náhled aktivního souboru (výběr listu)
- **🆕 GET+POST /excel_editor – interaktivní editor Excel souborů s možností úprav buněk přímo v prohlížeči**
- GET+POST /settings – uložení časových defaultů a projektových informací, propsání do Excelu
- GET+POST /zalohy – přidání/aktualizace zálohy pro zaměstnance
- POST /start_new_file – „archivace“: zneaktivnění aktuálního souboru (po nastavení data konce v Nastavení)
- POST /voice-command – zpracování textového příkazu JSON; provede záznam času nebo vrátí statistiky
- GET+POST /monthly_report – souhrn hodin a volných dní za měsíc (volitelně filtrováno na vybrané zaměstnance)

Pozn.: Některé akce vyžadují inicializované Excel/Zálohy managery; ochrana je přes dekorátor require_excel_managers.

## Hlasové/textové příkazy

Endpoint POST /voice-command očekává JSON:

```json
{ "command": "zapiš práci dnes od 7 do 16 oběd 0.5" }
```

Detekuje:

- action: record_time / get_stats (+ record_free_day přes příznak is_free_day)
- date: „dnes“, „včera“, „zítra“ nebo 2025‑08‑11 / 11.08.2025 / 11/08/2025
- start_time, end_time: „od 7 do 17“, „7‑17“, „7:00‑17:00“, slovně „od sedmi do šestnácti“
- lunch_duration: „oběd 0.5“ (0..4 h)
- get_stats: time_period „week|month|year“ a volitelně jméno zaměstnance

V této aplikaci se používá textový vstup (neposílá se audio). Soubor utils/voice_processor.py obsahuje i skeleton pro volání externího API (Gemini) – pro něj nastavte GEMINI_* proměnné prostředí, pokud chcete rozšířit na reálný STT/LLM.

## Excel Editor - Interaktivní úpravy

Nová funkce **Excel Editor** umožňuje editaci buněk přímo v prohlížeči bez nutnosti stahování souborů:

### Hlavní funkce:
- **Interaktivní tabulka**: Každá buňka je editovatelná textové pole
- **Automatické ukládání**: Změny se okamžitě ukládají do Excel souboru při opuštění buňky nebo stisknutí Enter
- **Vizuální feedback**: Indikátor ukládání zobrazuje stav operace (⏳ Ukládá se... → ✅ Uloženo)
- **Výběr souborů a listů**: Stejně jako u Excel Viewer lze vybrat konkrétní soubor a list
- **Responzivní design**: Plně optimalizováno pro mobilní zařízení a tablety

### Použití:
1. Klikněte na "Editor Tabulek" v navigaci
2. Vyberte soubor a list, který chcete editovat
3. Kliknutím na libovolnou buňku začněte editaci
4. Potvrzení změn: stiskněte Enter nebo klikněte mimo buňku
5. Změny se automaticky uloží a zobrazí se potvrzovací zpráva

### Technické detaily:
- Bezpečná validace vstupů na backend straně
- Ochrana proti souběžným úpravám
- Podpora pro všechny typy dat (text, čísla, formule)
- Kompatibilní se stávající Excel infrastrukturou aplikace

## Požadavky a prostředí

- Python: viz runtime.txt (doporučeno 3.12.7)
- Závislosti: requirements.txt
- OS: Linux/Windows; cesty a logování jsou ošetřené pro běh lokálně i na PythonAnywhere

## Instalace a spuštění

1. Vytvoření a aktivace prostředí

   Doporučujeme virtuální prostředí (venv/conda).

2. Instalace závislostí

    ```bash
    pip install -r requirements.txt
    ```

3. Lokální vývojový běh

    ```bash
    python app.py
    ```

4. Produkční běh (např. Gunicorn)

    ```bash
    gunicorn wsgi:application
    ```

Adresáře data/, excel/ a logs/ se vytvoří automaticky. Při prvním běhu se do excel/ vygeneruje šablona Hodiny_Cap.xlsx, pokud chybí.

## Konfigurace (proměnné prostředí)

- SECRET_KEY – tajný klíč Flasku (jinak se vygeneruje)
- HODINY_BASE_DIR – kořen projektu (default: aktuální složka)
- HODINY_DATA_PATH – cesta k data/ (default: BASE_DIR/data)
- HODINY_EXCEL_PATH – cesta k excel/ (default: BASE_DIR/excel)
- HODINY_SETTINGS_PATH – cesta k settings.json (default: data/settings.json)
- SMTP_SERVER, SMTP_PORT – SMTP konfigurace (default: smtp.gmail.com, 465)
- SMTP_USERNAME, SMTP_PASSWORD – přihlašovací údaje pro SMTP
- RECIPIENT_EMAIL – cílový e‑mail pro odeslání aktivního souboru
- GEMINI_API_KEY, GEMINI_API_URL – přístup k externímu API pro STT/NLP (volitelné)
- GEMINI_REQUEST_TIMEOUT, GEMINI_MAX_RETRIES, GEMINI_CACHE_TTL – síťové/caching parametry
- RATE_LIMIT_REQUESTS, RATE_LIMIT_WINDOW – omezení počtu požadavků pro hlasové API

Poznámka: Tajné hodnoty nesdílejte v repozitáři; nastavte je přes prostředí (např. .env / CI / hosting).

## Práce s aktivním souborem a projekty

- Aktivní soubor je trackován v data/settings.json (klíč active_excel_file)
- Při absenci nebo chybě se vytvoří nový soubor z šablony: {Projekt}_{YYYYMMDD_HHMMSS}.xlsx
- V Nastavení vyplňte název projektu a datum začátku; datum konce je povinné před archivací
- Archivace (POST /start_new_file) pouze resetuje active_excel_file – soubory zůstávají v excel/

## Logování

- logs/app.log, excel_manager.log, employee_management.log, verify_excel_data.log apod.
- Rotace po ~1 MB, až 5 záloh
- V lokálním vývoji také výstup na konzoli

## Testy

- K dispozici jsou testovací soubory (pytest). Spusťte: `pytest`
- Doporučení: oddělit integrační testy Excelu (práce se soubory) a jednotkové testy

## Nasazení (příklad)

- Gunicorn: `gunicorn wsgi:application`
- PythonAnywhere: wsgi.py nastaví cesty, vytvoří složky a inicializuje aplikaci; logy v `~/hodiny/logs`

## Známá omezení a tipy

- Excel formáty: při ručním zásahu do šablony zachovejte očekávané listy a buňky
- Zálohy: názvy možností čerpá aplikace z buněk B80/D80/F80/H80 listu „Zálohy“
- Hlasové příkazy: v základu textové; pro audio a LLM transkripci doplňte reálné API a bezpečné zacházení s klíči
- SMTP: některé služby vyžadují specifická hesla pro aplikace (Gmail – App Passwords)

## Rychlá reference rout

- /, /zaznam, /zamestnanci, /zalohy, /excel_viewer, **/excel_editor**, /settings, /monthly_report
- POST: /send_email, /set_active_file, /rename_project, /delete_project, /start_new_file, /voice-command

## Licence a autorství

Tento repozitář je určen pro interní použití. Ujistěte se, že neukládáte citlivé údaje (hesla/API klíče) do verzovacího systému.
