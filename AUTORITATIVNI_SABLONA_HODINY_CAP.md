# Autoritativní šablona `Hodiny_Cap:vzor.xlsx`

Tento dokument popisuje autoritativní Excel layout, podle kterého aplikace zapisuje data do runtime souboru `Hodiny_Cap.xlsx`.

## Princip

- Zdrojová autoritativní šablona je `excel/Hodiny_Cap:vzor.xlsx`.
- Aktivní runtime soubor aplikace je `excel/Hodiny_Cap.xlsx`.
- Pokud `Hodiny_Cap.xlsx` chybí, aplikace ho vytvoří kopií z autoritativní šablony.
- `config.json` proto mapuje buňky vůči runtime souboru `Hodiny_Cap.xlsx`, ale layout je převzatý z `Hodiny_Cap:vzor.xlsx`.

## Týdenní evidence

Listy:
- šablona `Týden`
- konkrétní týdny `Týden {cislo}`

Zápis používá vždy stejný sloupcový vzor po dnech:
- pondělí: `B/C`
- úterý: `D/E`
- středa: `F/G`
- čtvrtek: `H/I`
- pátek: `J/K`
- sobota: `L/M`
- neděle: `N/O`

Mapování:
- název projektu: `B4`
  - zapisuje se text `NÁZEV PROJEKTU : {project_name}`
  - oblast `B4:P4` je merged, zapisuje se do levé horní buňky `B4`
- zaměstnanec: `A8` a níže
- datum dne: `B80`, `D80`, `F80`, `H80`, `J80`, `L80`, `N80`
  - řádek `6` typicky obsahuje vzorce odkazující na řádek `80`
  - autoritativní zápis jde proto do řádku `80`
- začátek směny: `B7`, `D7`, `F7`, `H7`, `J7`, `L7`, `N7`
- konec směny: `C7`, `E7`, `G7`, `I7`, `K7`, `M7`, `O7`
- čisté hodiny zaměstnance: `B8`, `D8`, `F8`, `H8`, `J8`, `L8`, `N8` a níže po řádcích zaměstnanců

Pravidla zápisu:
- zaměstnanci začínají na řádku `8`
- datum se zapisuje vždy do řádku `80`
- pracovní časy se zapisují jako skutečné hodnoty `HH:MM`
- pro volný den (`00:00` až `00:00`) se do časových buněk nic nezapisuje
- hodiny se zapisují jako čistý čas po odečtení pauzy
- konfigurace z `sheet: "Týden"` platí i pro dynamické listy `Týden 1`, `Týden 2`, ...

## Zálohy

List:
- `Zálohy`

Mapování:
- zaměstnanec: `A8` a níže
- názvy možností: `B80`, `D80`, `F80`, `H80`
- EUR částky:
  - možnost 1: `B8`
  - možnost 2: `D8`
  - možnost 3: `F8`
  - možnost 4: `H8`
- CZK částky:
  - možnost 1: `C8`
  - možnost 2: `E8`
  - možnost 3: `G8`
  - možnost 4: `I8`
- datum posledního zápisu: `Z8` a níže po řádcích zaměstnanců

Pravidla zápisu:
- zaměstnanec se hledá nebo vytváří od řádku `8`
- sloupec částky se vybírá podle pořadí zvolené možnosti a měny
- datum se zapisuje do stejného řádku zaměstnance ve sloupci `Z`

## Runtime konfigurace v `config.json`

`config.json` nese technické mapování do runtime souboru `Hodiny_Cap.xlsx`:
- `weekly_time.employee_name` -> `A8`
- `weekly_time.date` -> `B80`
- `weekly_time.start_time` -> `B7`
- `weekly_time.end_time` -> `C7`
- `weekly_time.total_hours` -> `B8`
- `projects.project_name` -> `B4`
- `advances.employee_name` -> `A8`
- `advances.amount_eur` -> `B8`, `D8`, `F8`, `H8`
- `advances.amount_czk` -> `C8`, `E8`, `G8`, `I8`
- `advances.option_type` -> `B80`, `D80`, `F80`, `H80`
- `advances.date` -> `Z8`

## Důsledky pro kód

Aktuální implementace je s touto šablonou sladěná takto:
- `Config.EXCEL_EMPLOYEE_START_ROW = 8`
- `Config.init_app()` preferuje kopii z `Hodiny_Cap:vzor.xlsx`
- týdenní zápis respektuje oddělené buňky pro start a konec směny
- reporty čtou hodiny z denních sloupců `B/D/F/H/J/L/N`
- manager záloh vybírá částkové sloupce podle pořadí option mappingu, ne zápisem do všech kandidátů
