# Projekt: potřeby a cílený refaktor

Tento dokument shrnuje, co projekt aktuálně potřebuje, v jakém pořadí to dává smysl řešit a které oblasti jsou vhodné pro cílený refaktor.

## Aktuální stav

- Aplikace je provozně funkční.
- Testy aktuálně procházejí.
- Zjevný mrtvý kód a vývojářské artefakty byly odstraněny.
- Statická kontrola `flake8` aktuálně prochází.
- Běžné nastavení aplikace a technická Excel konfigurace jsou nově oddělené i na route/service vrstvě.
- HTML routy, hlavní flow i hlasové ovládání už jsou rozdělené do samostatných blueprintů a service vrstev.
- REST API `/api/v1/*` bylo konsolidováno na sdílené doménové služby místo paralelní logiky.
- Performance/cache vrstva byla zredukována na helpery, které jsou skutečně používané aplikací.
- V projektu ale zůstává několik strukturálních slabin, které zpomalují další vývoj.

## Hlavní potřeby projektu

### 1. Konsolidace nastavení a konfigurace

Projekt má dvě paralelní vrstvy nastavení:

- standardní flow přes `/settings`
- rozšířené flow přes `/nastaveni` a `/api/settings`

To vytváří několik problémů:

- nejasné, která vrstva je autoritativní
- část logiky je v HTML formulářích, část v JSON API
- `config.json` funguje jako runtime mapping, ale není jasně oddělen od běžného nastavení aplikace

Co projekt potřebuje:

- rozhodnout, zda zůstane jedna sjednocená stránka nastavení, nebo dvě jasně oddělené oblasti
- oddělit:
  - uživatelské nastavení aplikace
  - technické mapování do Excelu
  - deployment/env konfiguraci
- popsat a stabilizovat formát `config.json`

### 2. Rozdělení přerostlé aplikace

Soubor `app.py` je příliš velký a obsahuje:

- web routy
- pomocné utility
- upload logiku
- settings logiku
- API endpointy
- část validační logiky

Co projekt potřebuje:

- rozdělit `app.py` do menších doménových modulů nebo blueprintů
- oddělit HTML routy od JSON API
- oddělit pomocné služby od request handlerů

Aktuální stav:

- hotovo pro hlavní HTML flow, Excel flow, reporty, zaměstnance i nastavení
- hotovo i pro `/api/v1/*`, kde už nezůstává hlavní business logika rozlitá přímo v route handleru
- hotovo i pro základní cleanup `performance_optimizations.py`, kde byly odstraněny mrtvé helpery a doplněna invalidace cache po zápisu do Excelu

Doporučené cílové rozdělení:

- `routes/main.py`
- `routes/settings.py`
- `routes/excel.py`
- `routes/api.py`
- `services/upload_service.py`
- `services/settings_service.py`

### 3. Stabilizace Excel vrstvy

Excel logika je klíčová část systému, ale nese historické vrstvy:

- fallbacky na staré rozložení buněk
- dynamickou konfiguraci přes `config.json`
- práci s aktivním souborem
- generování týdenních kopií

Co projekt potřebuje:

- jasně definovat veřejné API `ExcelManager`
- oddělit:
  - zápis pracovní doby
  - čtení/reporting
  - správu souborů
  - metadata
- zmenšit počet implicitních fallbacků, které maskují chybnou konfiguraci

### 4. Stabilizace záloh a mapování

`ZalohyManager` aktuálně kombinuje:

- dynamickou konfiguraci
- historický fallback podle pevných sloupců

Co projekt potřebuje:

- rozhodnout, zda je primární:
  - pevná šablona
  - nebo dynamické mapování
- odstranit zbytečné dvojkolejnosti
- doplnit explicitní validaci konfigurace při startu aplikace

### 5. Lepší práce s runtime daty

Projekt používá runtime data v:

- `data/`
- `excel/`
- `logs/`

To je v pořádku, ale je potřeba jasněji oddělit:

- verzované šablony
- lokální data
- dočasné soubory
- generované metadata soubory

Co projekt potřebuje:

- mít jasně definované, co patří do repa a co ne
- doplnit dokumentaci k tomu, jak inicializovat projekt na čistém stroji
- případně zavést bootstrap skript pro vytvoření očekávané struktury

### 6. Přesnější testovací strategie

Testy teď pokrývají funkční minimum, ale ne architekturu.

Co projekt potřebuje:

- rozdělit testy na:
  - unit testy managerů
  - integrační testy Flask rout
  - testy upload flow
  - testy konfigurace a mappingu
- doplnit testy pro:
  - `/settings`
  - `/nastaveni`
  - `/api/*` a `/api/v1/*`
  - chování při neplatném `config.json`

### 7. Dokumentace pro další vývoj

Projekt potřebuje mít stručně, ale jasně popsané:

- co je hlavní workflow aplikace
- co je zdroj pravdy pro nastavení
- jak funguje Excel mapování
- jaké soubory jsou runtime data
- jak spustit projekt a testy

## Priority

### Priorita A: udělat co nejdřív

- sjednotit nebo jasně oddělit `/settings` a `/nastaveni`
- rozdělit `app.py`
- stabilizovat veřejné API `ExcelManager`
- rozhodnout autoritativní model pro Excel mapping

### Priorita B: velmi vhodné

- rozdělit služby do samostatných modulů
- zpřesnit validaci konfigurace při startu
- doplnit testy pro API a konfigurační flow
- zpřehlednit práci s runtime daty

### Priorita C: následně

- zlepšit dokumentaci
- zvážit migraci na konzistentnější doménový model
- případně zlepšit deployment a bootstrap workflow

## Navržený postup refaktoru

### Fáze 1: Nastavení a konfigurace

- zmapovat veškeré použití `settings.json` a `config.json`
- navrhnout cílový model konfigurace
- sjednotit UI a API pro nastavení

### Fáze 2: Rozdělení `app.py`

- vyseparovat HTML routy
- vyseparovat API routy
- přesunout pomocnou logiku do service vrstev

Stav:

- dokončeno pro hlavní webové routy
- dokončeno pro REST API vrstvu
- zbývá hlavně další zmenšení legacy utilit a konsolidace performance helperů

### Fáze 3: Excel doména

- refaktor `ExcelManager`
- refaktor `ZalohyManager`
- sjednotit kontrakty managerů

### Fáze 4: Testy a dokumentace

- doplnit chybějící testy
- aktualizovat README
- dopsat technickou dokumentaci k mappingu a runtime datům

## Rozhodovací body

Před větším refaktorem je potřeba potvrdit tyto body:

- Má zůstat advanced konfigurace přes `/nastaveni`, nebo se má odstranit?
- Má být `config.json` dlouhodobě podporovaný mechanismus, nebo jen přechodné řešení?
- Má aplikace dál pracovat s kopiemi týdenních souborů, nebo jen s jedním aktivním souborem?
- Má být cílem jen interní nástroj, nebo stabilní dlouhodobě udržovaná aplikace?

## Doporučení

Nejvhodnější další krok:

- začít refaktor konsolidací nastavení a konfigurace

Důvod:

- právě tady je největší architektonická nejasnost
- tato vrstva zasahuje `app.py`, frontend i Excel mapping
- bez jejího srovnání budou další změny zbytečně drahé
