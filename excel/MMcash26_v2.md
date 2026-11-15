# MMcash26 – přesná specifikace a aplikační protokol

## Účel
`MMcash26` je **master šablona pro jedno vyúčtování firemních peněz / jednu akci**.

Do této šablony se běžná akce nezapisuje. Aplikace:
1. zkopíruje list `MMcash26`,
2. pojmenuje kopii podle čísla akce ve formátu `ČAcash26`, kde `ČA` je číslo akce,
3. zapíše datum začátku do `Q2`,
4. během akce zapisuje finanční pohyby do řádků `3:302`,
5. na konci akce zapíše datum konce do `Q3`.

Vyplnění `Q3` znamená, že akce je z pohledu systému uzavřená. Pro další akci aplikace znovu použije čistý master `MMcash26`.

## 1. Hlavička transakcí

Řádek `2`:

| Sloupec | Název | Typ | Co se zadává |
|---|---|---|---|
| A | Datum / název | VSTUP | Pro TANKOVÁNÍ, PEAGE a UBYTOVÁNÍ datum. Pro OSTATNÍ se zde podle pravidla používá název kupované věci; lze použít i text. |
| B | Kategorie | VSTUP | Kategorie finančního pohybu. |
| C | Popis / detail | VSTUP | Volitelný bližší popis transakce. |
| D | Měna | VSTUP | `EUR` nebo `CZK`. |
| E | Částka | VSTUP | Částka v měně uvedené ve sloupci D. |
| F | Úhrada | VSTUP | `HOTOVĚ` nebo `KARTOU`. |
| G | Zaměstnanec | VSTUP | Jméno zaměstnance u zálohy zaměstnanci; jinak lze ponechat prázdné. |
| H | Typ pohybu | VSTUP | Určuje, zda jde o příjem, výdaj, zálohu zaměstnanci nebo vrácení zálohy. |
| I | Poznámka | VSTUP | Volitelná doplňující poznámka. |
| J | ID | VÝPOČET | Pořadové číslo záznamu. |
| K | Částka Kč | VÝPOČET | Přepočet na Kč podle kurzu Eura v `Q5`. |
| L | Výdaj Kč | VÝPOČET | Hodnota z K, pokud je typ `VÝDAJ` nebo `ZÁLOHA_ZAMĚSTNANCI`. |
| M | Příjem Kč | VÝPOČET | Hodnota z K, pokud je typ `PŘÍJEM` nebo `VRÁCENÍ_ZÁLOHY`. |
| N | Záloha zaměstnance Kč | VÝPOČET | Hodnota z K, pokud je typ `ZÁLOHA_ZAMĚSTNANCI`. |

## 2. Povolené hodnoty

### Kategorie – `B3:B302`
- `TANKOVÁNÍ`
- `PEAGE`
- `UBYTOVÁNÍ`
- `OSTATNÍ`
- `ZÁLOHA_ZAMĚSTNANCI`
- `VRÁCENÍ_ZÁLOHY`
- `PŘÍJEM`
- `PŘEVOD`

### Měna – `D3:D302`
- `EUR`
- `CZK`

### Úhrada – `F3:F302`
- `HOTOVĚ`
- `KARTOU`

### Typ pohybu – `H3:H302`
- `PŘÍJEM`
- `VÝDAJ`
- `ZÁLOHA_ZAMĚSTNANCI`
- `VRÁCENÍ_ZÁLOHY`
- `PŘEVOD`

## 3. Přesné vzorce řádku

Pro každý řádek od 3 do 302 platí stejná logika.

### `J3` – ID
`=IF(E3="","",ROW()-2)`

### `K3` – částka převedená na Kč
`=IF(E3="","",IF(D3="EUR",E3*$Q$5,E3))`

### `L3` – výdaj
`=IF(OR(H3="VÝDAJ",H3="ZÁLOHA_ZAMĚSTNANCI"),K3,0)`

### `M3` – příjem
`=IF(OR(H3="PŘÍJEM",H3="VRÁCENÍ_ZÁLOHY"),K3,0)`

### `N3` – záloha zaměstnanci
`=IF(H3="ZÁLOHA_ZAMĚSTNANCI",K3,0)`

Tyto vzorce jsou vyplněny přes `J3:N302`.

## 4. Panel akce a vyúčtování

Panel je v `P1:Q15`.

| Buňka | Význam | Typ | Funkce |
|---|---|---|---|
| `P1:Q1` | Stav a vyúčtování akce | POPIS | Nadpis panelu. |
| `P2` | Začátek akce | POPIS | Popisek vstupu. |
| `Q2` | Datum začátku akce | VSTUP | Aplikace sem zapíše start akce. |
| `P3` | Konec akce | POPIS | Popisek vstupu. |
| `Q3` | Datum konce akce | VSTUP | Aplikace sem zapíše konec akce. Jeho vyplnění znamená UZAVŘENÍ listu. |
| `P4` | Stav listu | POPIS | Popis stavu akce. |
| `Q4` | Stav listu | VÝPOČET | `NEZAHÁJENO`, `OTEVŘENO`, nebo `UZAVŘENO`. |
| `P5` | Kurz EUR (Kč/EUR) | POPIS | Popis vstupu kurzu. |
| `Q5` | Kurz EUR | VSTUP | Kurzem se přepočítají všechny částky EUR v K. |
| `P6` | Celkem přijaté firemní peníze Kč | POPIS | Součet všech příjmů. |
| `Q6` | Celkem přijaté firemní peníze | VÝPOČET | `SUM(M3:M302)`. |
| `P7` | Celkem výdaje Kč | POPIS | Součet všech výdajů včetně záloh zaměstnancům. |
| `Q7` | Celkem výdaje | VÝPOČET | `SUM(L3:L302)`. |
| `P8` | Zálohy zaměstnancům Kč | POPIS | Součet samostatných záloh zaměstnancům. |
| `Q8` | Zálohy zaměstnancům | VÝPOČET | `SUM(N3:N302)`. |
| `P9` | Vrácené zálohy Kč | POPIS | Peníze vrácené zaměstnanci. |
| `Q9` | Vrácené zálohy | VÝPOČET | `SUMIF(H3:H302,"VRÁCENÍ_ZÁLOHY",K3:K302)`. |
| `P10` | Očekávaná hotovost u sebe Kč | POPIS | Co by podle hotovostních pohybů mělo fyzicky zůstat. |
| `Q10` | Očekávaná hotovost | VÝPOČET | `SUMIFS(M3:M302,F3:F302,"HOTOVĚ")-SUMIFS(L3:L302,F3:F302,"HOTOVĚ")`. |
| `P11` | Skutečná hotovost u sebe Kč | POPIS | Skutečně spočítané peníze u sebe. |
| `Q11` | Skutečná hotovost | VSTUP | Zadává uživatel / aplikace při vyúčtování. |
| `P12` | Rozdíl hotovosti Kč | POPIS | Kontrola fyzické hotovosti. |
| `Q12` | Rozdíl hotovosti | VÝPOČET | `Q11-Q10`, pokud je Q11 vyplněno. |
| `P13` | Vyúčtováno celkem Kč | POPIS | Celková hodnota, kterou musí akce finančně pokrýt. |
| `Q13` | Vyúčtováno celkem | VÝPOČET | `Q7-Q9+Q11` (pokud Q11 není zadáno, bere se 0). |
| `P14` | Rozdíl proti příjmům Kč | POPIS | Hlavní kontrola, zda vyúčtování sedí na přijaté firemní peníze. |
| `Q14` | Rozdíl proti příjmům | VÝPOČET | `Q13-Q6`. |
| `P15` | Stav vyúčtování | POPIS | Výsledek uzavření akce. |
| `Q15` | Stav vyúčtování | VÝPOČET | Při chybějícím Q3 je `OTEVŘENO`. Po zadání konce: `UZAVŘENO – SEDÍ` nebo `UZAVŘENO – ROZDÍL`. |

## 5. Význam hlavního výsledku

Logika vyúčtování je:

`PŘIJATÉ FIREMNÍ PENÍZE`
=
`VÝDAJE − VRÁCENÉ ZÁLOHY + PENÍZE, KTERÉ ZŮSTALY UŽIVATELI`

V tabulce:

`Q6` se porovnává s `Q13`.

- `Q14 = 0` → celkově peníze sedí.
- `Q14 > 0` → ve vyúčtování je více peněz, než odpovídá evidovaným příjmům.
- `Q14 < 0` → ve vyúčtování chybí část peněz nebo je některý příjem evidován nesprávně.

Současně:
- `Q10` = očekávaná fyzická hotovost.
- `Q11` = skutečně spočítaná fyzická hotovost.
- `Q12` = rozdíl fyzické hotovosti.

Pro uzavření s výsledkem `UZAVŘENO – SEDÍ` musí být současně:
`Q3` vyplněno,
`Q14 ≈ 0`,
`Q12 ≈ 0`.

## 6. Přehled podle kategorií

Rozsah `S1:T9`.

| Rozsah | Kategorie | Výpočet |
|---|---|---|
| `S2:T2` | TANKOVÁNÍ | `SUMIF(B3:B302,"TANKOVÁNÍ",K3:K302)` |
| `S3:T3` | PEAGE | `SUMIF(B3:B302,"PEAGE",K3:K302)` |
| `S4:T4` | UBYTOVÁNÍ | `SUMIF(B3:B302,"UBYTOVÁNÍ",K3:K302)` |
| `S5:T5` | OSTATNÍ | `SUMIF(B3:B302,"OSTATNÍ",K3:K302)` |
| `S6:T6` | ZÁLOHA_ZAMĚSTNANCI | `SUMIF(B3:B302,"ZÁLOHA_ZAMĚSTNANCI",K3:K302)` |
| `S7:T7` | VRÁCENÍ_ZÁLOHY | `SUMIF(B3:B302,"VRÁCENÍ_ZÁLOHY",K3:K302)` |
| `S8:T8` | PŘÍJEM | `SUMIF(B3:B302,"PŘÍJEM",K3:K302)` |
| `S9:T9` | PŘEVOD | `SUMIF(B3:B302,"PŘEVOD",K3:K302)` |

Všechny součty jsou v Kč.

## 7. Pravidla pro aplikaci

### Vytvoření akce
1. Zdroj je vždy čistý `MMcash26`.
2. Aplikace vytvoří kopii.
3. Kopii přejmenuje na `ČAcash26`.
4. Do `Q2` zapíše datum začátku akce.
5. Během akce zapisuje záznamy do `A3:I302`.

### Běžný výdaj
Záznam má:
- kategorii `TANKOVÁNÍ`, `PEAGE`, `UBYTOVÁNÍ` nebo `OSTATNÍ`,
- měnu `EUR`/`CZK`,
- částku,
- `HOTOVĚ`/`KARTOU`,
- typ `VÝDAJ`.

### Záloha zaměstnanci
- kategorie `ZÁLOHA_ZAMĚSTNANCI`,
- jméno v `G`,
- typ `ZÁLOHA_ZAMĚSTNANCI`.

### Příjem firemních peněz
- kategorie `PŘÍJEM`,
- typ `PŘÍJEM`,
- částka se automaticky zahrne do `Q6`.

### Vrácení zálohy
- kategorie `VRÁCENÍ_ZÁLOHY`,
- typ `VRÁCENÍ_ZÁLOHY`,
- částka zvýší příjem a současně se odečte v čistém vyúčtování.

### OSTATNÍ
U položky `OSTATNÍ` může být v `A` místo klasického data přímo název kupované věci. Detail lze doplnit do `C`.

## 8. Uzavření akce

Po skončení akce aplikace vyplní `Q3`.

Tím vznikne stav:
`Q4 = UZAVŘENO`

a `Q15` vyhodnotí finanční shodu.

Master `MMcash26` se nepoužívá jako archiv konkrétní akce. Archivní data patří do kopie `ČAcash26`.

## 9. Co aplikace nesmí přepisovat

Výpočtové buňky:
`J3:N302`,
`Q4`,
`Q6:Q10`,
`Q12:Q15`,
`T2:T9`.

Vstupy pro aplikaci:
`Q2`, `Q3`, `Q5`, `Q11`
a v transakčních řádcích `A3:I302`.

Poznámka: `J:N` a `Q4/Q6:Q10/Q12:Q15/T2:T9` jsou pouze výpočty a mají zůstat nedotčené aplikací.
