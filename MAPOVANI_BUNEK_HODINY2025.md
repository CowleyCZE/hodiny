# Dokumentace mapování buněk - Hodiny2025.xlsx

## 📊 Struktura Excel souboru Hodiny2025.xlsx

### 🗂️ Organizace listů

- **Template list**: `MMhod25` - Šablona pro vytváření nových měsíčních listů
- **Měsíční listy**: `01hod25`, `02hod25`, ..., `12hod25` (leden až prosinec 2025)

### 📋 Mapování buněk pro každý měsíční list

#### 🏷️ ZÁHLAVÍ SOUBORU

```text
A1: "Měsíční výkaz práce - [Název měsíce] 2025"
B1: "Hodiny Evidence System"
```

#### 📝 HLAVIČKY TABULKY (řádek 2)

```text
A2: "Den"                 - Den v měsíci (1-31)
B2: "Datum"               - Datum ve formátu DD.MM.YYYY
C2: "Den v týdnu"         - Po, Út, St, Čt, Pá, So, Ne
D2: "Svátek"              - Označení svátku (pokud je)
E2: "Začátek"             - Čas začátku práce (HH:MM)
F2: "Oběd (h)"            - Délka oběda v hodinách (desetinné číslo)
G2: "Konec"               - Čas konce práce (HH:MM)
H2: "Celkem hodin"        - Celkově odpracované hodiny (vzorec)
I2: "Přesčasy"            - Přesčasové hodiny (vzorec)
J2: "Noční práce"         - Noční hodiny (rezervováno)
K2: "Víkend"              - Víkendové hodiny (rezervováno)
L2: "Svátky"              - Sváteční hodiny (rezervováno)
M2: "Zaměstnanci"         - Počet zaměstnanců daný den
N2: "Celkem odpracováno"  - Celkově odpracováno všemi zaměstnanci (vzorec)
```

#### 💾 DATOVÁ OBLAST (řádky 3-33)

**Indexování řádků**: `řádek = den_v_měsíci + 2`

- 1. den měsíce → řádek 3
- další den měsíce → řádek 4
- ...
- poslední den měsíce → řádek 33

**Mapování sloupců**:

```text
A[řádek]:  Den v měsíci (1, 2, 3, ..., 31)
B[řádek]:  Datum (01.01.2025, 02.01.2025, ...)
C[řádek]:  Den v týdnu (Po, Út, St, ...)
D[řádek]:  Svátek (prázdné nebo název svátku)
E[řádek]:  Začátek práce (07:00, 08:00, ...)      ← VSTUP
F[řádek]:  Oběd v hodinách (0.5, 1.0, 1.5, ...)  ← VSTUP
G[řádek]:  Konec práce (15:00, 16:00, ...)        ← VSTUP
H[řádek]:  Celkem hodin = vzorec                  ← VYPOČÍTÁNO
I[řádek]:  Přesčasy = vzorec                      ← VYPOČÍTÁNO
J[řádek]:  Noční práce (rezervováno)
K[řádek]:  Víkend (rezervováno)
L[řádek]:  Svátky (rezervováno)
M[řádek]:  Počet zaměstnanců (1, 2, 3, ...)       ← VSTUP
N[řádek]:  Celkem za všechny = vzorec             ← VYPOČÍTÁNO
```

#### 🧮 VZORCE

**Celkové hodiny (sloupec H)**:

```excel
=IF(AND(E[řádek]<>"",G[řádek]<>""),(G[řádek]-E[řádek])*24-F[řádek],0)
```

- Počítá rozdíl mezi koncem a začátkem práce v hodinách
- Odečte dobu oběda
- Pokud nejsou zadány časy, vrátí 0

**Přesčasy (sloupec I)**:

```excel
=MAX(0,H[řádek]-8)
```

- Vše nad 8 hodin je považováno za přesčas
- Minimálně 0 (záporné přesčasy nejsou)

**Celkem za všechny zaměstnance (sloupec N)**:

```excel
=H[řádek]*M[řádek]
```

- Vynásobí odpracované hodiny počtem zaměstnanců

#### 📊 SOUHRNY (řádek 34)

```text
A34: "SOUHRN:"
H34: =SUM(H3:H33)  - Celkem hodin za měsíc
I34: =SUM(I3:I33)  - Celkem přesčasů za měsíc  
N34: =SUM(N3:N33)  - Celkem odpracováno všemi za měsíc
```

### 🎨 FORMÁTOVÁNÍ

#### Víkendové dny (sobota, neděle)

- Světle červené pozadí (#FFE6E6)
- Aplikuje se na celý řádek

#### Souhrny (řádek 34)

- Tučné písmo
- Šedé pozadí (#CCCCCC)

#### Hlavičky (řádek 2)

- Tučné písmo
- Středové zarovnání

### 🔧 KONSTANTY V KÓDU

```python
# Řádky
HEADER_ROW = 2        # Řádek s hlavičkami
DATA_START_ROW = 3    # První řádek s daty (1. den)
DATA_END_ROW = 33     # Poslední řádek s daty (31. den)
SUMMARY_ROW = 34      # Řádek se souhrny

# Sloupce (1-based indexování)
COL_DAY = 1           # A - Den v měsíci
COL_DATE = 2          # B - Datum
COL_WEEKDAY = 3       # C - Den v týdnu
COL_HOLIDAY = 4       # D - Svátek
COL_START = 5         # E - Začátek práce      ← HLAVNÍ VSTUP
COL_LUNCH = 6         # F - Oběd (hodiny)      ← HLAVNÍ VSTUP
COL_END = 7           # G - Konec práce        ← HLAVNÍ VSTUP
COL_TOTAL_HOURS = 8   # H - Celkem hodin
COL_OVERTIME = 9      # I - Přesčasy
COL_NIGHT = 10        # J - Noční práce
COL_WEEKEND = 11      # K - Víkend
COL_HOLIDAY_HOURS = 12 # L - Sváteční hodiny
COL_EMPLOYEES = 13    # M - Počet zaměstnanců  ← HLAVNÍ VSTUP
COL_TOTAL_ALL = 14    # N - Celkem za všechny
```

### 📅 NÁZVY LISTŮ

**Template**: `MMhod25`

**Měsíční listy**:

- Leden: `01hod25`
- Únor: `02hod25`
- Březen: `03hod25`
- Duben: `04hod25`
- Květen: `05hod25`
- Červen: `06hod25`
- Červenec: `07hod25`
- Srpen: `08hod25`
- Září: `09hod25`
- Říjen: `10hod25`
- Listopad: `11hod25`
- Prosinec: `12hod25`

### 🚀 HLAVNÍ API METODY

#### `zapis_pracovni_doby(date, start_time, end_time, lunch_duration, num_employees)`

**Vstupní parametry**:

- `date`: "2025-01-15" (YYYY-MM-DD)
- `start_time`: "07:00" (HH:MM)
- `end_time`: "15:30" (HH:MM)  
- `lunch_duration`: "1.0" (hodiny jako string)
- `num_employees`: 3 (integer)

**Zapisuje do buněk**:

- E[řádek]: start_time jako time objekt
- F[řádek]: lunch_duration jako float
- G[řádek]: end_time jako time objekt  
- M[řádek]: num_employees jako integer

#### `get_daily_record(date)`

**Vrací dictionary s údaji o dni**:

```python
{
    'date': '2025-01-15',
    'start_time': '07:00',
    'end_time': '15:30', 
    'lunch_hours': 1.0,
    'total_hours': 7.5,
    'overtime': 0.0,
    'num_employees': 3,
    'total_all_employees': 22.5
}
```

#### `get_monthly_summary(month, year)`

**Vrací měsíční souhrn**:

```python
{
    'month': 1,
    'year': 2025,
    'month_name': 'Leden',
    'total_hours': 168.5,
    'total_overtime': 12.0,
    'total_all_employees': 487.5
}
```

### ✅ TESTOVACÍ SCÉNÁŘE

#### Běžný pracovní den

- Začátek: 07:00, Konec: 15:30, Oběd: 0.5h, Zaměstnanci: 3
- Výsledek: 8h práce, 0h přesčasů, 24h celkem za všechny

#### Přesčasový den

- Začátek: 07:00, Konec: 17:00, Oběd: 1.0h, Zaměstnanci: 2  
- Výsledek: 9h práce, 1h přesčasů, 18h celkem za všechny

#### Volný den

- Začátek: 00:00, Konec: 00:00, Oběd: 0h, Zaměstnanci: 0
- Výsledek: 0h práce, 0h přesčasů, 0h celkem za všechny
