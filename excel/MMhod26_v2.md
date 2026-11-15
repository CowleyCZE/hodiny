# MMhod26 / měsíční listy – přesná specifikace 2026

## Účel
Sešit obsahuje hlavní roční list `MMhod26` a 12 měsíčních listů:

`Leden`, `Únor`, `Březen`, `Duben`, `Květen`, `Červen`, `Červenec`, `Srpen`, `Září`, `Říjen`, `Listopad`, `Prosinec`.

Všechny měsíční listy používají stejnou strukturu jako `MMhod26`, ale obsahují pouze datumy příslušného měsíce. `MMhod26` obsahuje všech 365 dnů roku 2026.

Žluté buňky jsou vstupy. Zelené buňky jsou výpočty. Aplikace smí zapisovat pouze do vstupních buněk.

## 1. Společná struktura `MMhod26` a měsíčních listů

### Hlavička
- `A1:T1` – název listu.
- `A2:Q2` – hlavní tabulka.
- `S1:T1` – panel `NASTAVENÍ A SOUHRN`.

### Hlavní tabulka – sloupce

| Buňka / rozsah | Název | Typ | Co se zadává / co zobrazuje |
|---|---|---|---|
| `A3:A367` | Datum | VSTUP | Datum konkrétního dne. V `MMhod26` je vyplněn celý rok 1.1.–31.12.2026. V měsíčních listech jen datumy daného měsíce. |
| `B3:B367` | Den | VÝPOČET | Den v týdnu podle data ve sloupci A: Po, Út, St, Čt, Pá, So, Ne. |
| `C3:C367` | Od | VSTUP | Začátek práce, čas `hh:mm`. |
| `D3:D367` | Pauza (h) | VSTUP | Délka pauzy v hodinách, např. `0,5` nebo `1`. |
| `E3:E367` | Do | VSTUP | Konec práce, čas `hh:mm`. |
| `F3:F367` | Čisté hodiny | VÝPOČET | Čistý pracovní čas: `Do − Od − Pauza`. Excel jej zobrazuje jako `[h]:mm`. |
| `G3:G367` | Hodinová mzda (Kč/h) | VÝPOČET | Přebírá hodinovou mzdu z `T2` stejného listu. |
| `H3:H367` | Celkem Kč | VÝPOČET | Hodnota práce: `čisté hodiny × hodinová mzda × počet osob`. Pokud je počet osob prázdný, výpočet použije 1 osobu. |
| `I3:I367` | Přesčas h | VÝPOČET | Denní přesčas nad 8 hodin: `MAX(0; čisté hodiny − 8)`. |
| `J3:J367` | Přesčas Kč | VÝPOČET | Přesčasové hodiny × hodinová mzda × počet osob. |
| `K3:K367` | Město | VSTUP | Město / místo montáže. |
| `L3:L367` | Firma | VSTUP | Firma / zákazník. |
| `M3:M367` | Auto | VSTUP | Údaj o použitém vozidle. |
| `N3:N367` | Počet osob | VSTUP | Počet osob, pro které se má řádek násobit. Výchozí hodnota v šabloně je `1`. |
| `O3:O367` | Mrazák (h) | VSTUP | Počet hodin práce v mrazáku pro samostatný měsíční součet. |
| `P3:P367` | Poznámka | VSTUP | Volná poznámka k danému dni. |
| `Q3:Q367` | Status | VÝPOČET | `VOLNO`, pokud není zadán začátek ani konec; `NEKOMPLETNÍ`, pokud chybí jeden z časů; `OK`, pokud je začátek i konec zadán. |

### Přesné základní vzorce
Vzorce jsou vždy lokální pro daný list.

- `B3`:
  `=IF(A3="","",CHOOSE(WEEKDAY(A3,2),"Po","Út","St","Čt","Pá","So","Ne"))`
- `F3`:
  `=IF(OR(C3="",E3=""),"",(E3-C3)-D3/24)`
- `G3`:
  `=IF(A3="","",$T$2)`
- `H3`:
  `=IF(F3="","",F3*G3*24*IF(N3="",1,N3))`
- `I3`:
  `=IF(F3="","",MAX(0,F3*24-8))`
- `J3`:
  `=IF(I3="","",I3*G3*IF(N3="",1,N3))`
- `Q3`:
  `=IF(A3="","",IF(AND(C3="",E3=""),"VOLNO",IF(OR(C3="",E3=""),"NEKOMPLETNÍ","OK")))`

Tyto vzorce se kopírují dolů pouze přes aktivní řádky daného listu.

## 2. Panel nastavení a měsíční souhrn

| Buňka | Význam | Typ |
|---|---|---|
| `S2` | Popisek `Hodinová mzda (Kč/h)` | POPIS |
| `T2` | Hodinová mzda | VSTUP |
| `S3` | Popisek `Kurz Eura (Kč/EUR)` | POPIS |
| `T3` | Kurz Eura | VSTUP |
| `S4` | Popisek `Odpracováno hodin` | POPIS |
| `T4` | Celkové odpracované hodiny | VÝPOČET |
| `S5` | Popisek `Celkem práce Kč` | POPIS |
| `T5` | Celková hodnota práce podle sloupce H | VÝPOČET |
| `S6` | Popisek `Přesčas Kč` | POPIS |
| `T6` | Celkový součet přesčasu podle sloupce J | VÝPOČET |
| `S7` | Popisek `Počet pracovních dnů` | POPIS |
| `T7` | Počet řádků, kde je zadán začátek i konec práce | VÝPOČET |
| `S8` | Popisek `Mrazák hodin` | POPIS |
| `T8` | Celkem hodin mrazáku | VÝPOČET |
| `S9` | Popisek `Rok` | POPIS |
| `T9` | Rok `2026` | KONSTANTA |
| `S10` | Popisek `Poznámka` | POPIS |
| `T10` | Legenda `Žluté = vstup; zelené = výpočet` | POPIS |

### Souhrnné vzorce
Pro měsíční list jsou rozsahy nastaveny na skutečný počet dnů měsíce.

- `T4` = `SUM(F3:Fxx)*24`
- `T5` = `SUM(H3:Hxx)`
- `T6` = `SUM(J3:Jxx)`
- `T7` = `COUNTIFS(C3:Cxx,"<>",E3:Exx,"<>")`
- `T8` = `SUM(O3:Oxx)`

Pro `MMhod26` je `xx=367`.
Pro měsíc s 31 dny je `xx=33`, se 30 dny `xx=32`, s 28 dny `xx=30`.

## 3. Kalendářní rozsahy pro 2026

| List | Datum od | Datum do | Počet dnů |
|---|---:|---:|---:|
| Leden | 1.1.2026 | 31.1.2026 | 31 |
| Únor | 1.2.2026 | 28.2.2026 | 28 |
| Březen | 1.3.2026 | 31.3.2026 | 31 |
| Duben | 1.4.2026 | 30.4.2026 | 30 |
| Květen | 1.5.2026 | 31.5.2026 | 31 |
| Červen | 1.6.2026 | 30.6.2026 | 30 |
| Červenec | 1.7.2026 | 31.7.2026 | 31 |
| Srpen | 1.8.2026 | 31.8.2026 | 31 |
| Září | 1.9.2026 | 30.9.2026 | 30 |
| Říjen | 1.10.2026 | 31.10.2026 | 31 |
| Listopad | 1.11.2026 | 30.11.2026 | 30 |
| Prosinec | 1.12.2026 | 31.12.2026 | 31 |

## 4. Aplikační zápis

Aplikace zapisující docházku má používat pouze:
`A, C, D, E, K, L, M, N, O, P` v příslušném řádku.

Aplikace nikdy nepřepisuje:
`B, F, G, H, I, J, Q, T4:T8`.

Pro běžnou práci jednoho člověka stačí `N=1`. Pokud je evidována skupina, lze počet osob změnit.

## 5. Rozdíl mezi `MMhod26` a měsíčními listy

- `MMhod26` = kompletní roční výkaz se všemi 365 daty.
- `Leden`–`Prosinec` = stejné výpočty a nastavení, ale vždy jen jeden konkrétní měsíc.
- Měsíční listy nejsou napojeny na sebe a každý má vlastní hodnotu hodinové mzdy a kurzu Eura v `T2:T3`.
