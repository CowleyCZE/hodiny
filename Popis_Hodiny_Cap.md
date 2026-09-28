# Kompletní popis buněk – `Hodiny_Cap.xlsx`

Tento dokument je datový slovník aktuální verze tabulky. Obsahuje popis všech neprázdných buněk, všech vzorců, struktury listů, formátů, sloučení a dalších prvků důležitých pro další práci s tabulkou.

## Legenda

| Typ | Význam |
|---|---|
| `VSTUP` | Údaj, který běžně zadává uživatel. |
| `VÝPOČET` | Buňka obsahující vzorec; hodnotu dopočítává Excel. |
| `KONSTANTA` | Pevná číselná nebo jiná hodnota používaná tabulkou. |
| `POPIS / HODNOTA` | Nadpis, popis, jméno, text nebo pevně uložený údaj. |
| `PRÁZDNÁ` | Buňka bez obsahu. Samotná prázdnota neznamená, že jde o vstupní pole. |

# Přehled listů

| List | Řádky | Sloupce | Vzorce | Neprázdné buňky |
|---|---:|---:|---:|---:|
| `ZÁLOHY` | 308 | 26 | 32 | 55 |
| `Týden` | 280 | 18 | 14 | 64 |

# List `ZÁLOHY`

## Základní informace

- Použitý rozsah: **308 × 26**.
- Zamrznuté panely: nejsou nastaveny.
- Automatický filtr není nastaven.

### Sloučené oblasti

- `I5:I7`
- `D6:E6`
- `B6:C6`
- `D4:J4`
- `F25:K25`
- `A30:E30`
- `F29:L29`
- `A6:A7`
- `A1:A5`
- `J5:J7`
- `B4:C4`

## Struktura záhlaví

| Sloupec | 1. řádek | 2. řádek | 3. řádek | 4. řádek |
|---|---|---|---|---|
| `B` | `` | `` | `Zálohy ` | `NÁZEV PROJEKTU      :` |

## Kompletní inventář neprázdných buněk

| Buňka | Typ | Obsah | Formát |
|---|---|---|---|
| `B3` | **POPIS / HODNOTA** | `Zálohy ` | `General` |
| `B4` | **POPIS / HODNOTA** | `NÁZEV PROJEKTU      :` | `General` |
| `I5` | **POPIS / HODNOTA** | `€ CELKEM €` | `General` |
| `J5` | **POPIS / HODNOTA** | `Kč CELKEM Kč` | `General` |
| `A6` | **POPIS / HODNOTA** | `Jméno pracovníka:` | `General` |
| `B6` | **POPIS / HODNOTA** | `Akt.mesic` | `General` |
| `D6` | **POPIS / HODNOTA** | `Příští měsíc` | `General` |
| `B7` | **POPIS / HODNOTA** | `Eura` | `General` |
| `C7` | **POPIS / HODNOTA** | `CZK` | `General` |
| `D7` | **POPIS / HODNOTA** | `Eura` | `General` |
| `E7` | **POPIS / HODNOTA** | `CZK` | `General` |
| `F7` | **POPIS / HODNOTA** | `Eura` | `General` |
| `G7` | **POPIS / HODNOTA** | `CZK` | `General` |
| `A8` | **POPIS / HODNOTA** | `Čáp Jakub` | `General` |
| `I8` | **VÝPOČET** | `=SUM(B8+D8+F8)` | `#,##0\ [$€-1];[RED]\-#,##0\ [$€-1]` |
| `J8` | **VÝPOČET** | `=SUM(C8+E8+G8)` | `General` |
| `I9` | **VÝPOČET** | `=SUM(B9+D9+F9)` | `#,##0\ [$€-1];[RED]\-#,##0\ [$€-1]` |
| `J9` | **VÝPOČET** | `=SUM(C9+E9+G9)` | `General` |
| `Z9` | **POPIS / HODNOTA** | `2025-02-27 00:00:00` | `yyyy\-mm\-dd` |
| `I10` | **VÝPOČET** | `=SUM(B10+D10+F10)` | `#,##0\ [$€-1];[RED]\-#,##0\ [$€-1]` |
| `J10` | **VÝPOČET** | `=SUM(C10+E10+G10)` | `General` |
| `Z10` | **POPIS / HODNOTA** | `2025-02-27 00:00:00` | `yyyy\-mm\-dd` |
| `I11` | **VÝPOČET** | `=SUM(B11+D11+F11)` | `#,##0\ [$€-1];[RED]\-#,##0\ [$€-1]` |
| `J11` | **VÝPOČET** | `=SUM(C11+E11+G11)` | `General` |
| `Z11` | **POPIS / HODNOTA** | `2025-02-27 00:00:00` | `yyyy\-mm\-dd` |
| `I12` | **VÝPOČET** | `=SUM(B12+D12+F12)` | `#,##0\ [$€-1];[RED]\-#,##0\ [$€-1]` |
| `J12` | **VÝPOČET** | `=SUM(C12+E12+G12)` | `General` |
| `Z12` | **POPIS / HODNOTA** | `2025-02-27 00:00:00` | `yyyy\-mm\-dd` |
| `I13` | **VÝPOČET** | `=SUM(B13+D13+F13)` | `#,##0\ [$€-1];[RED]\-#,##0\ [$€-1]` |
| `J13` | **VÝPOČET** | `=SUM(C13+E13+G13)` | `General` |
| `Z13` | **POPIS / HODNOTA** | `2026-09-02 00:00:00` | `dd/mm/yyyy` |
| `I14` | **VÝPOČET** | `=SUM(B14+D14+F14)` | `#,##0\ [$€-1];[RED]\-#,##0\ [$€-1]` |
| `J14` | **VÝPOČET** | `=SUM(C14+E14+G14)` | `General` |
| `Z14` | **POPIS / HODNOTA** | `2025-02-01 00:00:00` | `yyyy\-mm\-dd` |
| `I15` | **VÝPOČET** | `=SUM(B15+D15+F15)` | `#,##0\ [$€-1];[RED]\-#,##0\ [$€-1]` |
| `J15` | **VÝPOČET** | `=SUM(C15+E15+G15)` | `General` |
| `I16` | **VÝPOČET** | `=SUM(B16+D16+F16)` | `#,##0\ [$€-1];[RED]\-#,##0\ [$€-1]` |
| `J16` | **VÝPOČET** | `=SUM(C16+E16+G16)` | `General` |
| `I17` | **VÝPOČET** | `=SUM(B17+D17+F17)` | `#,##0\ [$€-1];[RED]\-#,##0\ [$€-1]` |
| `J17` | **VÝPOČET** | `=SUM(C17+E17+G17)` | `General` |
| `I18` | **VÝPOČET** | `=SUM(B18+D18+F18)` | `#,##0\ [$€-1];[RED]\-#,##0\ [$€-1]` |
| `J18` | **VÝPOČET** | `=SUM(C18+E18+G18)` | `General` |
| `I19` | **VÝPOČET** | `=SUM(B19+D19+F19)` | `#,##0\ [$€-1];[RED]\-#,##0\ [$€-1]` |
| `J19` | **VÝPOČET** | `=SUM(C19+E19+G19)` | `General` |
| `I20` | **VÝPOČET** | `=SUM(B20+D20+F20)` | `#,##0\ [$€-1];[RED]\-#,##0\ [$€-1]` |
| `J20` | **VÝPOČET** | `=SUM(C20+E20+G20)` | `General` |
| `I21` | **VÝPOČET** | `=SUM(B21+D21+F21)` | `#,##0\ [$€-1];[RED]\-#,##0\ [$€-1]` |
| `J21` | **VÝPOČET** | `=SUM(C21+E21+G21)` | `General` |
| `I22` | **VÝPOČET** | `=SUM(B22+D22+F22)` | `#,##0\ [$€-1];[RED]\-#,##0\ [$€-1]` |
| `J22` | **VÝPOČET** | `=SUM(C22+E22+G22)` | `General` |
| `I23` | **VÝPOČET** | `=SUM(I8:I22)` | `#,##0\ [$€-1]` |
| `J23` | **VÝPOČET** | `=SUM(J8:J22)` | `#,##0.00\ [$Kč-405];\-#,##0.00\ [$Kč-405]` |
| `Z90` | **POPIS / HODNOTA** | `2025-02-28 00:00:00` | `yyyy\-mm\-dd` |
| `Z99` | **POPIS / HODNOTA** | `2025-02-28 00:00:00` | `yyyy\-mm\-dd` |
| `Z108` | **POPIS / HODNOTA** | `2025-02-28 00:00:00` | `yyyy\-mm\-dd` |

## Všechny automatické výpočty

| Buňka | Vzorec |
|---|---|
| `I8` | `=SUM(B8+D8+F8)` |
| `J8` | `=SUM(C8+E8+G8)` |
| `I9` | `=SUM(B9+D9+F9)` |
| `J9` | `=SUM(C9+E9+G9)` |
| `I10` | `=SUM(B10+D10+F10)` |
| `J10` | `=SUM(C10+E10+G10)` |
| `I11` | `=SUM(B11+D11+F11)` |
| `J11` | `=SUM(C11+E11+G11)` |
| `I12` | `=SUM(B12+D12+F12)` |
| `J12` | `=SUM(C12+E12+G12)` |
| `I13` | `=SUM(B13+D13+F13)` |
| `J13` | `=SUM(C13+E13+G13)` |
| `I14` | `=SUM(B14+D14+F14)` |
| `J14` | `=SUM(C14+E14+G14)` |
| `I15` | `=SUM(B15+D15+F15)` |
| `J15` | `=SUM(C15+E15+G15)` |
| `I16` | `=SUM(B16+D16+F16)` |
| `J16` | `=SUM(C16+E16+G16)` |
| `I17` | `=SUM(B17+D17+F17)` |
| `J17` | `=SUM(C17+E17+G17)` |
| `I18` | `=SUM(B18+D18+F18)` |
| `J18` | `=SUM(C18+E18+G18)` |
| `I19` | `=SUM(B19+D19+F19)` |
| `J19` | `=SUM(C19+E19+G19)` |
| `I20` | `=SUM(B20+D20+F20)` |
| `J20` | `=SUM(C20+E20+G20)` |
| `I21` | `=SUM(B21+D21+F21)` |
| `J21` | `=SUM(C21+E21+G21)` |
| `I22` | `=SUM(B22+D22+F22)` |
| `J22` | `=SUM(C22+E22+G22)` |
| `I23` | `=SUM(I8:I22)` |
| `J23` | `=SUM(J8:J22)` |

## Praktický popis práce s listem

### Uživatelské vstupy

Za uživatelský vstup lze považovat buňku, která je podle struktury listu určena k zadání údajů a zároveň není výpočtová. Přesný obsah a formát každé neprázdné buňky je uveden v inventáři výše.

### Automatické údaje

Všechny buňky obsahující vzorec jsou vypsány v části **Všechny automatické výpočty**. Tyto buňky se běžně ručně nepřepisují.

### Důležité pravidlo

Prázdná buňka není automaticky považována za vstupní. U šablon je nutné rozlišovat mezi skutečným vstupním polem, volným místem, formátovanou oblastí a pomocnou buňkou.

# List `Týden`

## Základní informace

- Použitý rozsah: **280 × 18**.
- Zamrznuté panely: nejsou nastaveny.
- Automatický filtr není nastaven.

### Sloučené oblasti

- `L17:M17`
- `D20:E20`
- `N20:O20`
- `N11:O11`
- `F16:G16`
- `H10:I10`
- `J10:K10`
- `J19:K19`
- `B22:C22`
- `L19:M19`
- `D22:E22`
- `N22:O22`
- `H9:I9`
- `B12:C12`
- `J9:K9`
- `D6:E6`
- `B21:C21`
- `L12:M12`
- `N12:O12`
- `F15:G15`
- `D13:E13`
- `B14:C14`
- `L5:M5`
- `P5:P7`
- `N5:O5`
- `L14:M14`
- `N14:O14`
- `B13:C13`
- `F13:G13`
- `H13:I13`
- `J21:K21`
- `I4:P4`
- `L21:M21`
- `N15:O15`
- `H15:I15`
- `B16:C16`
- `D10:E10`
- `F10:G10`
- `L16:M16`
- `D19:E19`
- `N16:O16`
- `F19:G19`
- `F6:G6`
- `B18:C18`
- `B11:C11`
- `D11:E11`
- `F5:G5`
- `F20:G20`
- `F18:G18`
- `F17:G17`
- `H17:I17`
- `H11:I11`
- `J11:K11`
- `J20:K20`
- `L20:M20`
- `H19:I19`
- `J22:K22`
- `D9:E9`
- `N13:O13`
- `L22:M22`
- `F9:G9`
- `B6:C6`
- `H12:I12`
- `J12:K12`
- `B15:C15`
- `L6:M6`
- `D15:E15`
- `N6:O6`
- `H5:I5`
- `H14:I14`
- `A30:H30`
- `A6:A7`
- `H31:I31`
- `F21:G21`
- `H21:I21`
- `J25:R25`
- `B4:H4`
- `L18:M18`
- `N18:O18`
- `B17:C17`
- `L8:M8`
- `D17:E17`
- `N8:O8`
- `F8:G8`
- `N17:O17`
- `H16:I16`
- `J16:K16`
- `B19:C19`
- `L10:M10`
- `N10:O10`
- `N19:O19`
- `F22:G22`
- `H22:I22`
- `B9:C9`
- `A1:A5`
- `L9:M9`
- `D12:E12`
- `N9:O9`
- `J29:R29`
- `B5:C5`
- `B20:C20`
- `D5:E5`
- `L11:M11`
- `D14:E14`
- `F14:G14`
- `F11:G11`
- `H20:I20`
- `J5:K5`
- `J14:K14`
- `J13:K13`
- `L13:M13`
- `D21:E21`
- `F12:G12`
- `N21:O21`
- `H6:I6`
- `J6:K6`
- `J15:K15`
- `L15:M15`
- `B8:C8`
- `D16:E16`
- `B10:C10`
- `D18:E18`
- `D8:E8`
- `H18:I18`
- `J18:K18`
- `H8:I8`
- `J8:K8`
- `J17:K17`

## Struktura záhlaví

| Sloupec | 1. řádek | 2. řádek | 3. řádek | 4. řádek |
|---|---|---|---|---|
| `B` | `` | `` | `Týden` | `NÁZEV PROJEKTU :` |
| `C` | `` | `` | `číslo týdne ` | `` |
| `I` | `` | `` | `` | `název` |

## Kompletní inventář neprázdných buněk

| Buňka | Typ | Obsah | Formát |
|---|---|---|---|
| `B3` | **POPIS / HODNOTA** | `Týden` | `@` |
| `C3` | **POPIS / HODNOTA** | `číslo týdne ` | `@` |
| `B4` | **POPIS / HODNOTA** | `NÁZEV PROJEKTU :` | `General` |
| `I4` | **POPIS / HODNOTA** | `název` | `General` |
| `B5` | **POPIS / HODNOTA** | `Pondělí` | `General` |
| `D5` | **POPIS / HODNOTA** | `Úterý` | `General` |
| `F5` | **POPIS / HODNOTA** | `Středa` | `General` |
| `H5` | **POPIS / HODNOTA** | `Čtvrtek` | `General` |
| `J5` | **POPIS / HODNOTA** | `Pátek` | `General` |
| `L5` | **POPIS / HODNOTA** | `Sobota` | `General` |
| `N5` | **POPIS / HODNOTA** | `Neděle` | `General` |
| `P5` | **POPIS / HODNOTA** | `HODINY CELKEM` | `General` |
| `A6` | **POPIS / HODNOTA** | `Jméno pracovníka:` | `General` |
| `B6` | **POPIS / HODNOTA** | `datum` | `dd/mm/yyyy` |
| `D6` | **POPIS / HODNOTA** | `datum` | `dd/mm/yyyy` |
| `F6` | **POPIS / HODNOTA** | `datum` | `dd/mm/yyyy` |
| `H6` | **POPIS / HODNOTA** | `datum` | `dd/mm/yyyy` |
| `J6` | **POPIS / HODNOTA** | `datum` | `dd/mm/yyyy` |
| `L6` | **POPIS / HODNOTA** | `datum` | `dd/mm/yyyy` |
| `N6` | **POPIS / HODNOTA** | `datum` | `dd/mm/yyyy` |
| `B7` | **POPIS / HODNOTA** | `07:00:00` | `h:mm` |
| `C7` | **POPIS / HODNOTA** | `17:00:00` | `h:mm` |
| `D7` | **POPIS / HODNOTA** | `Od` | `h:mm` |
| `E7` | **POPIS / HODNOTA** | `Do` | `h:mm` |
| `F7` | **POPIS / HODNOTA** | `Od` | `h:mm` |
| `G7` | **POPIS / HODNOTA** | `Do` | `h:mm` |
| `H7` | **POPIS / HODNOTA** | `Od` | `h:mm` |
| `I7` | **POPIS / HODNOTA** | `Do` | `h:mm` |
| `J7` | **POPIS / HODNOTA** | `Od` | `h:mm` |
| `K7` | **POPIS / HODNOTA** | `Do` | `h:mm` |
| `L7` | **POPIS / HODNOTA** | `Od` | `h:mm` |
| `M7` | **POPIS / HODNOTA** | `Do` | `h:mm` |
| `N7` | **POPIS / HODNOTA** | `Od` | `h:mm` |
| `O7` | **POPIS / HODNOTA** | `Do` | `h:mm` |
| `A8` | **POPIS / HODNOTA** | `Čáp Jakub` | `General` |
| `P8` | **POPIS / HODNOTA** | `celkem hodin zaměstnance ` | `General` |
| `P9` | **VÝPOČET** | `=IF(COUNTA(B9:O9)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B9),ISNUMBER(C9)),MOD(C9-B9,1),0),IF(AND(ISNUMBER(D9),ISNUMBER(E9)),MOD(E9-D9,1),0),IF(AND(ISNUMBER(F9),ISNUMBER(G9)),MOD(G9-F9,1),0),IF(AND(ISNUMBER(H9),ISNUMBER(I9)),MOD(I9-H9,1),0),IF(AND(ISNUMBER(J9),ISNUMBER(K9)),MOD(K9-J9,1),0),IF(AND(ISNUMBER(L9),ISNUMBER(M9)),MOD(M9-L9,1),0),IF(AND(ISNUMBER(N9),ISNUMBER(O9)),MOD(O9-N9,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B9<=$D$25,C9>=$F$25,D9<=$D$25,E9>=$F$25,F9<=$D$25,G9>=$F$25,H9<=$D$25,I9>=$F$25,J9<=$D$25,K9>=$F$25,L9<=$D$25,M9>=$F$25,N9<=$D$25,O9>=$F$25),),$F$25-$D$25,0))*24)` | `General` |
| `P10` | **VÝPOČET** | `=IF(COUNTA(B10:O10)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B10),ISNUMBER(C10)),MOD(C10-B10,1),0),IF(AND(ISNUMBER(D10),ISNUMBER(E10)),MOD(E10-D10,1),0),IF(AND(ISNUMBER(F10),ISNUMBER(G10)),MOD(G10-F10,1),0),IF(AND(ISNUMBER(H10),ISNUMBER(I10)),MOD(I10-H10,1),0),IF(AND(ISNUMBER(J10),ISNUMBER(K10)),MOD(K10-J10,1),0),IF(AND(ISNUMBER(L10),ISNUMBER(M10)),MOD(M10-L10,1),0),IF(AND(ISNUMBER(N10),ISNUMBER(O10)),MOD(O10-N10,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B10<=$D$25,C10>=$F$25,D10<=$D$25,E10>=$F$25,F10<=$D$25,G10>=$F$25,H10<=$D$25,I10>=$F$25,J10<=$D$25,K10>=$F$25,L10<=$D$25,M10>=$F$25,N10<=$D$25,O10>=$F$25),),$F$25-$D$25,0))*24)` | `General` |
| `P11` | **VÝPOČET** | `=IF(COUNTA(B11:O11)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B11),ISNUMBER(C11)),MOD(C11-B11,1),0),IF(AND(ISNUMBER(D11),ISNUMBER(E11)),MOD(E11-D11,1),0),IF(AND(ISNUMBER(F11),ISNUMBER(G11)),MOD(G11-F11,1),0),IF(AND(ISNUMBER(H11),ISNUMBER(I11)),MOD(I11-H11,1),0),IF(AND(ISNUMBER(J11),ISNUMBER(K11)),MOD(K11-J11,1),0),IF(AND(ISNUMBER(L11),ISNUMBER(M11)),MOD(M11-L11,1),0),IF(AND(ISNUMBER(N11),ISNUMBER(O11)),MOD(O11-N11,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B11<=$D$25,C11>=$F$25,D11<=$D$25,E11>=$F$25,F11<=$D$25,G11>=$F$25,H11<=$D$25,I11>=$F$25,J11<=$D$25,K11>=$F$25,L11<=$D$25,M11>=$F$25,N11<=$D$25,O11>=$F$25),),$F$25-$D$25,0))*24)` | `General` |
| `P12` | **VÝPOČET** | `=IF(COUNTA(B12:O12)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B12),ISNUMBER(C12)),MOD(C12-B12,1),0),IF(AND(ISNUMBER(D12),ISNUMBER(E12)),MOD(E12-D12,1),0),IF(AND(ISNUMBER(F12),ISNUMBER(G12)),MOD(G12-F12,1),0),IF(AND(ISNUMBER(H12),ISNUMBER(I12)),MOD(I12-H12,1),0),IF(AND(ISNUMBER(J12),ISNUMBER(K12)),MOD(K12-J12,1),0),IF(AND(ISNUMBER(L12),ISNUMBER(M12)),MOD(M12-L12,1),0),IF(AND(ISNUMBER(N12),ISNUMBER(O12)),MOD(O12-N12,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B12<=$D$25,C12>=$F$25,D12<=$D$25,E12>=$F$25,F12<=$D$25,G12>=$F$25,H12<=$D$25,I12>=$F$25,J12<=$D$25,K12>=$F$25,L12<=$D$25,M12>=$F$25,N12<=$D$25,O12>=$F$25),),$F$25-$D$25,0))*24)` | `General` |
| `P13` | **VÝPOČET** | `=IF(COUNTA(B13:O13)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B13),ISNUMBER(C13)),MOD(C13-B13,1),0),IF(AND(ISNUMBER(D13),ISNUMBER(E13)),MOD(E13-D13,1),0),IF(AND(ISNUMBER(F13),ISNUMBER(G13)),MOD(G13-F13,1),0),IF(AND(ISNUMBER(H13),ISNUMBER(I13)),MOD(I13-H13,1),0),IF(AND(ISNUMBER(J13),ISNUMBER(K13)),MOD(K13-J13,1),0),IF(AND(ISNUMBER(L13),ISNUMBER(M13)),MOD(M13-L13,1),0),IF(AND(ISNUMBER(N13),ISNUMBER(O13)),MOD(O13-N13,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B13<=$D$25,C13>=$F$25,D13<=$D$25,E13>=$F$25,F13<=$D$25,G13>=$F$25,H13<=$D$25,I13>=$F$25,J13<=$D$25,K13>=$F$25,L13<=$D$25,M13>=$F$25,N13<=$D$25,O13>=$F$25),),$F$25-$D$25,0))*24)` | `General` |
| `P14` | **VÝPOČET** | `=IF(COUNTA(B14:O14)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B14),ISNUMBER(C14)),MOD(C14-B14,1),0),IF(AND(ISNUMBER(D14),ISNUMBER(E14)),MOD(E14-D14,1),0),IF(AND(ISNUMBER(F14),ISNUMBER(G14)),MOD(G14-F14,1),0),IF(AND(ISNUMBER(H14),ISNUMBER(I14)),MOD(I14-H14,1),0),IF(AND(ISNUMBER(J14),ISNUMBER(K14)),MOD(K14-J14,1),0),IF(AND(ISNUMBER(L14),ISNUMBER(M14)),MOD(M14-L14,1),0),IF(AND(ISNUMBER(N14),ISNUMBER(O14)),MOD(O14-N14,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B14<=$D$25,C14>=$F$25,D14<=$D$25,E14>=$F$25,F14<=$D$25,G14>=$F$25,H14<=$D$25,I14>=$F$25,J14<=$D$25,K14>=$F$25,L14<=$D$25,M14>=$F$25,N14<=$D$25,O14>=$F$25),),$F$25-$D$25,0))*24)` | `General` |
| `P15` | **VÝPOČET** | `=IF(COUNTA(B15:O15)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B15),ISNUMBER(C15)),MOD(C15-B15,1),0),IF(AND(ISNUMBER(D15),ISNUMBER(E15)),MOD(E15-D15,1),0),IF(AND(ISNUMBER(F15),ISNUMBER(G15)),MOD(G15-F15,1),0),IF(AND(ISNUMBER(H15),ISNUMBER(I15)),MOD(I15-H15,1),0),IF(AND(ISNUMBER(J15),ISNUMBER(K15)),MOD(K15-J15,1),0),IF(AND(ISNUMBER(L15),ISNUMBER(M15)),MOD(M15-L15,1),0),IF(AND(ISNUMBER(N15),ISNUMBER(O15)),MOD(O15-N15,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B15<=$D$25,C15>=$F$25,D15<=$D$25,E15>=$F$25,F15<=$D$25,G15>=$F$25,H15<=$D$25,I15>=$F$25,J15<=$D$25,K15>=$F$25,L15<=$D$25,M15>=$F$25,N15<=$D$25,O15>=$F$25),),$F$25-$D$25,0))*24)` | `General` |
| `P16` | **VÝPOČET** | `=IF(COUNTA(B16:O16)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B16),ISNUMBER(C16)),MOD(C16-B16,1),0),IF(AND(ISNUMBER(D16),ISNUMBER(E16)),MOD(E16-D16,1),0),IF(AND(ISNUMBER(F16),ISNUMBER(G16)),MOD(G16-F16,1),0),IF(AND(ISNUMBER(H16),ISNUMBER(I16)),MOD(I16-H16,1),0),IF(AND(ISNUMBER(J16),ISNUMBER(K16)),MOD(K16-J16,1),0),IF(AND(ISNUMBER(L16),ISNUMBER(M16)),MOD(M16-L16,1),0),IF(AND(ISNUMBER(N16),ISNUMBER(O16)),MOD(O16-N16,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B16<=$D$25,C16>=$F$25,D16<=$D$25,E16>=$F$25,F16<=$D$25,G16>=$F$25,H16<=$D$25,I16>=$F$25,J16<=$D$25,K16>=$F$25,L16<=$D$25,M16>=$F$25,N16<=$D$25,O16>=$F$25),),$F$25-$D$25,0))*24)` | `General` |
| `P17` | **VÝPOČET** | `=IF(COUNTA(B17:O17)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B17),ISNUMBER(C17)),MOD(C17-B17,1),0),IF(AND(ISNUMBER(D17),ISNUMBER(E17)),MOD(E17-D17,1),0),IF(AND(ISNUMBER(F17),ISNUMBER(G17)),MOD(G17-F17,1),0),IF(AND(ISNUMBER(H17),ISNUMBER(I17)),MOD(I17-H17,1),0),IF(AND(ISNUMBER(J17),ISNUMBER(K17)),MOD(K17-J17,1),0),IF(AND(ISNUMBER(L17),ISNUMBER(M17)),MOD(M17-L17,1),0),IF(AND(ISNUMBER(N17),ISNUMBER(O17)),MOD(O17-N17,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B17<=$D$25,C17>=$F$25,D17<=$D$25,E17>=$F$25,F17<=$D$25,G17>=$F$25,H17<=$D$25,I17>=$F$25,J17<=$D$25,K17>=$F$25,L17<=$D$25,M17>=$F$25,N17<=$D$25,O17>=$F$25),),$F$25-$D$25,0))*24)` | `General` |
| `P18` | **VÝPOČET** | `=IF(COUNTA(B18:O18)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B18),ISNUMBER(C18)),MOD(C18-B18,1),0),IF(AND(ISNUMBER(D18),ISNUMBER(E18)),MOD(E18-D18,1),0),IF(AND(ISNUMBER(F18),ISNUMBER(G18)),MOD(G18-F18,1),0),IF(AND(ISNUMBER(H18),ISNUMBER(I18)),MOD(I18-H18,1),0),IF(AND(ISNUMBER(J18),ISNUMBER(K18)),MOD(K18-J18,1),0),IF(AND(ISNUMBER(L18),ISNUMBER(M18)),MOD(M18-L18,1),0),IF(AND(ISNUMBER(N18),ISNUMBER(O18)),MOD(O18-N18,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B18<=$D$25,C18>=$F$25,D18<=$D$25,E18>=$F$25,F18<=$D$25,G18>=$F$25,H18<=$D$25,I18>=$F$25,J18<=$D$25,K18>=$F$25,L18<=$D$25,M18>=$F$25,N18<=$D$25,O18>=$F$25),),$F$25-$D$25,0))*24)` | `General` |
| `P19` | **VÝPOČET** | `=IF(COUNTA(B19:O19)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B19),ISNUMBER(C19)),MOD(C19-B19,1),0),IF(AND(ISNUMBER(D19),ISNUMBER(E19)),MOD(E19-D19,1),0),IF(AND(ISNUMBER(F19),ISNUMBER(G19)),MOD(G19-F19,1),0),IF(AND(ISNUMBER(H19),ISNUMBER(I19)),MOD(I19-H19,1),0),IF(AND(ISNUMBER(J19),ISNUMBER(K19)),MOD(K19-J19,1),0),IF(AND(ISNUMBER(L19),ISNUMBER(M19)),MOD(M19-L19,1),0),IF(AND(ISNUMBER(N19),ISNUMBER(O19)),MOD(O19-N19,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B19<=$D$25,C19>=$F$25,D19<=$D$25,E19>=$F$25,F19<=$D$25,G19>=$F$25,H19<=$D$25,I19>=$F$25,J19<=$D$25,K19>=$F$25,L19<=$D$25,M19>=$F$25,N19<=$D$25,O19>=$F$25),),$F$25-$D$25,0))*24)` | `General` |
| `P20` | **VÝPOČET** | `=IF(COUNTA(B20:O20)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B20),ISNUMBER(C20)),MOD(C20-B20,1),0),IF(AND(ISNUMBER(D20),ISNUMBER(E20)),MOD(E20-D20,1),0),IF(AND(ISNUMBER(F20),ISNUMBER(G20)),MOD(G20-F20,1),0),IF(AND(ISNUMBER(H20),ISNUMBER(I20)),MOD(I20-H20,1),0),IF(AND(ISNUMBER(J20),ISNUMBER(K20)),MOD(K20-J20,1),0),IF(AND(ISNUMBER(L20),ISNUMBER(M20)),MOD(M20-L20,1),0),IF(AND(ISNUMBER(N20),ISNUMBER(O20)),MOD(O20-N20,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B20<=$D$25,C20>=$F$25,D20<=$D$25,E20>=$F$25,F20<=$D$25,G20>=$F$25,H20<=$D$25,I20>=$F$25,J20<=$D$25,K20>=$F$25,L20<=$D$25,M20>=$F$25,N20<=$D$25,O20>=$F$25),),$F$25-$D$25,0))*24)` | `General` |
| `P21` | **VÝPOČET** | `=IF(COUNTA(B21:O21)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B21),ISNUMBER(C21)),MOD(C21-B21,1),0),IF(AND(ISNUMBER(D21),ISNUMBER(E21)),MOD(E21-D21,1),0),IF(AND(ISNUMBER(F21),ISNUMBER(G21)),MOD(G21-F21,1),0),IF(AND(ISNUMBER(H21),ISNUMBER(I21)),MOD(I21-H21,1),0),IF(AND(ISNUMBER(J21),ISNUMBER(K21)),MOD(K21-J21,1),0),IF(AND(ISNUMBER(L21),ISNUMBER(M21)),MOD(M21-L21,1),0),IF(AND(ISNUMBER(N21),ISNUMBER(O21)),MOD(O21-N21,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B21<=$D$25,C21>=$F$25,D21<=$D$25,E21>=$F$25,F21<=$D$25,G21>=$F$25,H21<=$D$25,I21>=$F$25,J21<=$D$25,K21>=$F$25,L21<=$D$25,M21>=$F$25,N21<=$D$25,O21>=$F$25),),$F$25-$D$25,0))*24)` | `General` |
| `P22` | **POPIS / HODNOTA** | `CELKEM` | `General` |
| `P23` | **VÝPOČET** | `=SUM(P9:P21)` | `0.00` |
| `A24` | **POPIS / HODNOTA** | `Poznámky / Zálohy :` | `General` |
| `D24` | **POPIS / HODNOTA** | `Přestávka od - do` | `General` |
| `D25` | **POPIS / HODNOTA** | `12:00:00` | `hh:mm` |
| `F25` | **POPIS / HODNOTA** | `12:30:00` | `hh:mm` |
| `J25` | **POPIS / HODNOTA** | `HODINY PROSÍM VYPLŇOVAT V ČISTÉM ČASE, BEZ PŘESTÁVEK.` | `General` |
| `J29` | **POPIS / HODNOTA** | `HODINY PROSÍM POSÍLAT TÝDNĚ NA EMAIL: hodiny@czechmontage.cz` | `General` |
| `A30` | **POPIS / HODNOTA** | `Přejezdy z projektu na projekt - pouze pracovní cesty z jednoho projektu na druhý!` | `General` |
| `A31` | **POPIS / HODNOTA** | `Jméno pracovníka` | `General` |
| `B31` | **POPIS / HODNOTA** | `Odkud` | `General` |
| `D31` | **POPIS / HODNOTA** | `Kam ` | `General` |
| `F31` | **POPIS / HODNOTA** | `Kdy ` | `General` |
| `H31` | **POPIS / HODNOTA** | `Délka cesty` | `General` |
| `J31` | **POPIS / HODNOTA** | `Km` | `General` |

## Všechny automatické výpočty

| Buňka | Vzorec |
|---|---|
| `P9` | `=IF(COUNTA(B9:O9)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B9),ISNUMBER(C9)),MOD(C9-B9,1),0),IF(AND(ISNUMBER(D9),ISNUMBER(E9)),MOD(E9-D9,1),0),IF(AND(ISNUMBER(F9),ISNUMBER(G9)),MOD(G9-F9,1),0),IF(AND(ISNUMBER(H9),ISNUMBER(I9)),MOD(I9-H9,1),0),IF(AND(ISNUMBER(J9),ISNUMBER(K9)),MOD(K9-J9,1),0),IF(AND(ISNUMBER(L9),ISNUMBER(M9)),MOD(M9-L9,1),0),IF(AND(ISNUMBER(N9),ISNUMBER(O9)),MOD(O9-N9,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B9<=$D$25,C9>=$F$25,D9<=$D$25,E9>=$F$25,F9<=$D$25,G9>=$F$25,H9<=$D$25,I9>=$F$25,J9<=$D$25,K9>=$F$25,L9<=$D$25,M9>=$F$25,N9<=$D$25,O9>=$F$25),),$F$25-$D$25,0))*24)` |
| `P10` | `=IF(COUNTA(B10:O10)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B10),ISNUMBER(C10)),MOD(C10-B10,1),0),IF(AND(ISNUMBER(D10),ISNUMBER(E10)),MOD(E10-D10,1),0),IF(AND(ISNUMBER(F10),ISNUMBER(G10)),MOD(G10-F10,1),0),IF(AND(ISNUMBER(H10),ISNUMBER(I10)),MOD(I10-H10,1),0),IF(AND(ISNUMBER(J10),ISNUMBER(K10)),MOD(K10-J10,1),0),IF(AND(ISNUMBER(L10),ISNUMBER(M10)),MOD(M10-L10,1),0),IF(AND(ISNUMBER(N10),ISNUMBER(O10)),MOD(O10-N10,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B10<=$D$25,C10>=$F$25,D10<=$D$25,E10>=$F$25,F10<=$D$25,G10>=$F$25,H10<=$D$25,I10>=$F$25,J10<=$D$25,K10>=$F$25,L10<=$D$25,M10>=$F$25,N10<=$D$25,O10>=$F$25),),$F$25-$D$25,0))*24)` |
| `P11` | `=IF(COUNTA(B11:O11)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B11),ISNUMBER(C11)),MOD(C11-B11,1),0),IF(AND(ISNUMBER(D11),ISNUMBER(E11)),MOD(E11-D11,1),0),IF(AND(ISNUMBER(F11),ISNUMBER(G11)),MOD(G11-F11,1),0),IF(AND(ISNUMBER(H11),ISNUMBER(I11)),MOD(I11-H11,1),0),IF(AND(ISNUMBER(J11),ISNUMBER(K11)),MOD(K11-J11,1),0),IF(AND(ISNUMBER(L11),ISNUMBER(M11)),MOD(M11-L11,1),0),IF(AND(ISNUMBER(N11),ISNUMBER(O11)),MOD(O11-N11,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B11<=$D$25,C11>=$F$25,D11<=$D$25,E11>=$F$25,F11<=$D$25,G11>=$F$25,H11<=$D$25,I11>=$F$25,J11<=$D$25,K11>=$F$25,L11<=$D$25,M11>=$F$25,N11<=$D$25,O11>=$F$25),),$F$25-$D$25,0))*24)` |
| `P12` | `=IF(COUNTA(B12:O12)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B12),ISNUMBER(C12)),MOD(C12-B12,1),0),IF(AND(ISNUMBER(D12),ISNUMBER(E12)),MOD(E12-D12,1),0),IF(AND(ISNUMBER(F12),ISNUMBER(G12)),MOD(G12-F12,1),0),IF(AND(ISNUMBER(H12),ISNUMBER(I12)),MOD(I12-H12,1),0),IF(AND(ISNUMBER(J12),ISNUMBER(K12)),MOD(K12-J12,1),0),IF(AND(ISNUMBER(L12),ISNUMBER(M12)),MOD(M12-L12,1),0),IF(AND(ISNUMBER(N12),ISNUMBER(O12)),MOD(O12-N12,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B12<=$D$25,C12>=$F$25,D12<=$D$25,E12>=$F$25,F12<=$D$25,G12>=$F$25,H12<=$D$25,I12>=$F$25,J12<=$D$25,K12>=$F$25,L12<=$D$25,M12>=$F$25,N12<=$D$25,O12>=$F$25),),$F$25-$D$25,0))*24)` |
| `P13` | `=IF(COUNTA(B13:O13)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B13),ISNUMBER(C13)),MOD(C13-B13,1),0),IF(AND(ISNUMBER(D13),ISNUMBER(E13)),MOD(E13-D13,1),0),IF(AND(ISNUMBER(F13),ISNUMBER(G13)),MOD(G13-F13,1),0),IF(AND(ISNUMBER(H13),ISNUMBER(I13)),MOD(I13-H13,1),0),IF(AND(ISNUMBER(J13),ISNUMBER(K13)),MOD(K13-J13,1),0),IF(AND(ISNUMBER(L13),ISNUMBER(M13)),MOD(M13-L13,1),0),IF(AND(ISNUMBER(N13),ISNUMBER(O13)),MOD(O13-N13,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B13<=$D$25,C13>=$F$25,D13<=$D$25,E13>=$F$25,F13<=$D$25,G13>=$F$25,H13<=$D$25,I13>=$F$25,J13<=$D$25,K13>=$F$25,L13<=$D$25,M13>=$F$25,N13<=$D$25,O13>=$F$25),),$F$25-$D$25,0))*24)` |
| `P14` | `=IF(COUNTA(B14:O14)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B14),ISNUMBER(C14)),MOD(C14-B14,1),0),IF(AND(ISNUMBER(D14),ISNUMBER(E14)),MOD(E14-D14,1),0),IF(AND(ISNUMBER(F14),ISNUMBER(G14)),MOD(G14-F14,1),0),IF(AND(ISNUMBER(H14),ISNUMBER(I14)),MOD(I14-H14,1),0),IF(AND(ISNUMBER(J14),ISNUMBER(K14)),MOD(K14-J14,1),0),IF(AND(ISNUMBER(L14),ISNUMBER(M14)),MOD(M14-L14,1),0),IF(AND(ISNUMBER(N14),ISNUMBER(O14)),MOD(O14-N14,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B14<=$D$25,C14>=$F$25,D14<=$D$25,E14>=$F$25,F14<=$D$25,G14>=$F$25,H14<=$D$25,I14>=$F$25,J14<=$D$25,K14>=$F$25,L14<=$D$25,M14>=$F$25,N14<=$D$25,O14>=$F$25),),$F$25-$D$25,0))*24)` |
| `P15` | `=IF(COUNTA(B15:O15)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B15),ISNUMBER(C15)),MOD(C15-B15,1),0),IF(AND(ISNUMBER(D15),ISNUMBER(E15)),MOD(E15-D15,1),0),IF(AND(ISNUMBER(F15),ISNUMBER(G15)),MOD(G15-F15,1),0),IF(AND(ISNUMBER(H15),ISNUMBER(I15)),MOD(I15-H15,1),0),IF(AND(ISNUMBER(J15),ISNUMBER(K15)),MOD(K15-J15,1),0),IF(AND(ISNUMBER(L15),ISNUMBER(M15)),MOD(M15-L15,1),0),IF(AND(ISNUMBER(N15),ISNUMBER(O15)),MOD(O15-N15,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B15<=$D$25,C15>=$F$25,D15<=$D$25,E15>=$F$25,F15<=$D$25,G15>=$F$25,H15<=$D$25,I15>=$F$25,J15<=$D$25,K15>=$F$25,L15<=$D$25,M15>=$F$25,N15<=$D$25,O15>=$F$25),),$F$25-$D$25,0))*24)` |
| `P16` | `=IF(COUNTA(B16:O16)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B16),ISNUMBER(C16)),MOD(C16-B16,1),0),IF(AND(ISNUMBER(D16),ISNUMBER(E16)),MOD(E16-D16,1),0),IF(AND(ISNUMBER(F16),ISNUMBER(G16)),MOD(G16-F16,1),0),IF(AND(ISNUMBER(H16),ISNUMBER(I16)),MOD(I16-H16,1),0),IF(AND(ISNUMBER(J16),ISNUMBER(K16)),MOD(K16-J16,1),0),IF(AND(ISNUMBER(L16),ISNUMBER(M16)),MOD(M16-L16,1),0),IF(AND(ISNUMBER(N16),ISNUMBER(O16)),MOD(O16-N16,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B16<=$D$25,C16>=$F$25,D16<=$D$25,E16>=$F$25,F16<=$D$25,G16>=$F$25,H16<=$D$25,I16>=$F$25,J16<=$D$25,K16>=$F$25,L16<=$D$25,M16>=$F$25,N16<=$D$25,O16>=$F$25),),$F$25-$D$25,0))*24)` |
| `P17` | `=IF(COUNTA(B17:O17)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B17),ISNUMBER(C17)),MOD(C17-B17,1),0),IF(AND(ISNUMBER(D17),ISNUMBER(E17)),MOD(E17-D17,1),0),IF(AND(ISNUMBER(F17),ISNUMBER(G17)),MOD(G17-F17,1),0),IF(AND(ISNUMBER(H17),ISNUMBER(I17)),MOD(I17-H17,1),0),IF(AND(ISNUMBER(J17),ISNUMBER(K17)),MOD(K17-J17,1),0),IF(AND(ISNUMBER(L17),ISNUMBER(M17)),MOD(M17-L17,1),0),IF(AND(ISNUMBER(N17),ISNUMBER(O17)),MOD(O17-N17,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B17<=$D$25,C17>=$F$25,D17<=$D$25,E17>=$F$25,F17<=$D$25,G17>=$F$25,H17<=$D$25,I17>=$F$25,J17<=$D$25,K17>=$F$25,L17<=$D$25,M17>=$F$25,N17<=$D$25,O17>=$F$25),),$F$25-$D$25,0))*24)` |
| `P18` | `=IF(COUNTA(B18:O18)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B18),ISNUMBER(C18)),MOD(C18-B18,1),0),IF(AND(ISNUMBER(D18),ISNUMBER(E18)),MOD(E18-D18,1),0),IF(AND(ISNUMBER(F18),ISNUMBER(G18)),MOD(G18-F18,1),0),IF(AND(ISNUMBER(H18),ISNUMBER(I18)),MOD(I18-H18,1),0),IF(AND(ISNUMBER(J18),ISNUMBER(K18)),MOD(K18-J18,1),0),IF(AND(ISNUMBER(L18),ISNUMBER(M18)),MOD(M18-L18,1),0),IF(AND(ISNUMBER(N18),ISNUMBER(O18)),MOD(O18-N18,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B18<=$D$25,C18>=$F$25,D18<=$D$25,E18>=$F$25,F18<=$D$25,G18>=$F$25,H18<=$D$25,I18>=$F$25,J18<=$D$25,K18>=$F$25,L18<=$D$25,M18>=$F$25,N18<=$D$25,O18>=$F$25),),$F$25-$D$25,0))*24)` |
| `P19` | `=IF(COUNTA(B19:O19)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B19),ISNUMBER(C19)),MOD(C19-B19,1),0),IF(AND(ISNUMBER(D19),ISNUMBER(E19)),MOD(E19-D19,1),0),IF(AND(ISNUMBER(F19),ISNUMBER(G19)),MOD(G19-F19,1),0),IF(AND(ISNUMBER(H19),ISNUMBER(I19)),MOD(I19-H19,1),0),IF(AND(ISNUMBER(J19),ISNUMBER(K19)),MOD(K19-J19,1),0),IF(AND(ISNUMBER(L19),ISNUMBER(M19)),MOD(M19-L19,1),0),IF(AND(ISNUMBER(N19),ISNUMBER(O19)),MOD(O19-N19,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B19<=$D$25,C19>=$F$25,D19<=$D$25,E19>=$F$25,F19<=$D$25,G19>=$F$25,H19<=$D$25,I19>=$F$25,J19<=$D$25,K19>=$F$25,L19<=$D$25,M19>=$F$25,N19<=$D$25,O19>=$F$25),),$F$25-$D$25,0))*24)` |
| `P20` | `=IF(COUNTA(B20:O20)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B20),ISNUMBER(C20)),MOD(C20-B20,1),0),IF(AND(ISNUMBER(D20),ISNUMBER(E20)),MOD(E20-D20,1),0),IF(AND(ISNUMBER(F20),ISNUMBER(G20)),MOD(G20-F20,1),0),IF(AND(ISNUMBER(H20),ISNUMBER(I20)),MOD(I20-H20,1),0),IF(AND(ISNUMBER(J20),ISNUMBER(K20)),MOD(K20-J20,1),0),IF(AND(ISNUMBER(L20),ISNUMBER(M20)),MOD(M20-L20,1),0),IF(AND(ISNUMBER(N20),ISNUMBER(O20)),MOD(O20-N20,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B20<=$D$25,C20>=$F$25,D20<=$D$25,E20>=$F$25,F20<=$D$25,G20>=$F$25,H20<=$D$25,I20>=$F$25,J20<=$D$25,K20>=$F$25,L20<=$D$25,M20>=$F$25,N20<=$D$25,O20>=$F$25),),$F$25-$D$25,0))*24)` |
| `P21` | `=IF(COUNTA(B21:O21)=0,0,MAX(0,SUM(IF(AND(ISNUMBER(B21),ISNUMBER(C21)),MOD(C21-B21,1),0),IF(AND(ISNUMBER(D21),ISNUMBER(E21)),MOD(E21-D21,1),0),IF(AND(ISNUMBER(F21),ISNUMBER(G21)),MOD(G21-F21,1),0),IF(AND(ISNUMBER(H21),ISNUMBER(I21)),MOD(I21-H21,1),0),IF(AND(ISNUMBER(J21),ISNUMBER(K21)),MOD(K21-J21,1),0),IF(AND(ISNUMBER(L21),ISNUMBER(M21)),MOD(M21-L21,1),0),IF(AND(ISNUMBER(N21),ISNUMBER(O21)),MOD(O21-N21,1),0))-IF(AND(ISNUMBER($D$25),ISNUMBER($F$25),OR(B21<=$D$25,C21>=$F$25,D21<=$D$25,E21>=$F$25,F21<=$D$25,G21>=$F$25,H21<=$D$25,I21>=$F$25,J21<=$D$25,K21>=$F$25,L21<=$D$25,M21>=$F$25,N21<=$D$25,O21>=$F$25),),$F$25-$D$25,0))*24)` |
| `P23` | `=SUM(P9:P21)` |

## Ověření vstupních údajů

| Oblast | Typ | Podmínka 1 | Podmínka 2 |
|---|---|---|---|
| `B9:O21` | `time` | `TIME(0,0,0)` | `TIME(23,59,59)` |
| `B6 D6 F6 H6 J6 L6 N6` | `date` | `DATE(2020,1,1)` | `DATE(2035,12,31)` |

## Praktický popis práce s listem

### Uživatelské vstupy

Za uživatelský vstup lze považovat buňku, která je podle struktury listu určena k zadání údajů a zároveň není výpočtová. Přesný obsah a formát každé neprázdné buňky je uveden v inventáři výše.

### Automatické údaje

Všechny buňky obsahující vzorec jsou vypsány v části **Všechny automatické výpočty**. Tyto buňky se běžně ručně nepřepisují.

### Důležité pravidlo

Prázdná buňka není automaticky považována za vstupní. U šablon je nutné rozlišovat mezi skutečným vstupním polem, volným místem, formátovanou oblastí a pomocnou buňkou.

# Souhrnná kontrola

## Vzorce

- `ZÁLOHY`: **32** vzorců.
- `Týden`: **14** vzorců.
- Celkem v sešitu: **46** vzorců.

## Poznámka k přesnosti dokumentace

Tento soubor popisuje aktuální stav nahrané verze `Hodiny_Cap.xlsx`. Pokud se později změní vzorce, názvy listů, rozložení nebo pomocné buňky, je potřeba dokumentaci znovu vygenerovat.