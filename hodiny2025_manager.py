# hodiny2025_manager.py
"""
Hodiny2025Manager - Správa Excel souboru pro evidenci pracovních hodin v roce 2025

MAPOVÁNÍ BUNĚK V SOUBORU Hodiny2025.xlsx:
=========================================

STRUKTURA LISTU (např. "01hod25" pro leden 2025):
-------------------------------------------------
A1: "Měsíční výkaz práce - [Měsíc] 2025"
B1: Název společnosti/oddělení

ZÁHLAVÍ TABULKY (řádek 2):
--------------------------
A2: "Den"         B2: "Datum"       C2: "Den v týdnu"  D2: "Svátek"
E2: "Začátek"     F2: "Oběd (h)"    G2: "Konec"       H2: "Celkem hodin"
I2: "Přesčasy"    J2: "Noční práce" K2: "Víkend"      L2: "Svátky"
M2: "Zaměstnanci" N2: "Celkem odpracováno"

DATOVÁ OBLAST (řádky 3-33):
---------------------------
- Řádek = den v měsíci + 2 (tzn. 1. den = řádek 3, 2. den = řádek 4, atd.)
- A[řádek]: Den v měsíci (1-31)
- B[řádek]: Datum (DD.MM.YYYY)
- C[řádek]: Den v týdnu (Po, Út, St, Čt, Pá, So, Ne)
- D[řádek]: Označení svátku (pokud je)
- E[řádek]: Čas začátku práce (HH:MM)
- F[řádek]: Délka oběda v hodinách (desetinné číslo)
- G[řádek]: Čas konce práce (HH:MM)
- H[řádek]: Celkem odpracovaných hodin (vzorec)
- I[řádek]: Přesčasové hodiny (vzorec)
- J[řádek]: Noční práce (vzorec)
- K[řádek]: Víkendové hodiny (vzorec)
- L[řádek]: Sváteční hodiny (vzorec)
- M[řádek]: Počet zaměstnanců
- N[řádek]: Celkem odpracováno všemi zaměstnanci (vzorec)

SOUHRNY (řádky 34-40):
----------------------
A34: "SOUHRN:"
H34: =SUM(H3:H33) - Celkem hodin za měsíc
I34: =SUM(I3:I33) - Celkem přesčasů za měsíc
N34: =SUM(N3:N33) - Celkem odpracováno všemi za měsíc

VZORCE:
-------
H[řádek]: =(G[řádek]-E[řádek])*24-F[řádek]  # Celkem hodin = (konec-začátek)*24-oběd
I[řádek]: =MAX(0,H[řádek]-8)                # Přesčasy = max(0, celkem-8)
N[řádek]: =H[řádek]*M[řádek]                # Celkem za všechny = hodiny*počet_zaměstnanců
"""

from datetime import datetime, time
from pathlib import Path
import logging
import calendar
from openpyxl import load_workbook, Workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

try:
    from utils.logger import setup_logger
    logger = setup_logger("hodiny2025_manager")
except ImportError:
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger("hodiny2025_manager")


class Hodiny2025Manager:
    """
    Manager pro správu Excel souboru s evidencí pracovních hodin pro rok 2025.
    
    Struktura souboru:
    - Jeden list pro každý měsíc (formát: MMhod25, např. 01hod25, 02hod25)
    - Template list: MMhod25 (slouží jako šablona pro nové měsíce)
    """
    
    # Konstanty pro mapování buněk
    HEADER_ROW = 2
    DATA_START_ROW = 3
    DATA_END_ROW = 33
    SUMMARY_ROW = 34
    
    # Sloupce (1-based indexování pro Excel)
    COL_DAY = 1          # A - Den v měsíci
    COL_DATE = 2         # B - Datum
    COL_WEEKDAY = 3      # C - Den v týdnu
    COL_HOLIDAY = 4      # D - Svátek
    COL_START = 5        # E - Začátek práce
    COL_LUNCH = 6        # F - Oběd (hodiny)
    COL_END = 7          # G - Konec práce
    COL_TOTAL_HOURS = 8  # H - Celkem hodin
    COL_OVERTIME = 9     # I - Přesčasy
    COL_NIGHT = 10       # J - Noční práce
    COL_WEEKEND = 11     # K - Víkend
    COL_HOLIDAY_HOURS = 12 # L - Sváteční hodiny
    COL_EMPLOYEES = 13   # M - Počet zaměstnanců
    COL_TOTAL_ALL = 14   # N - Celkem za všechny
    
    # Názvy měsíců v češtině
    CZECH_MONTHS = {
        1: "Leden", 2: "Únor", 3: "Březen", 4: "Duben",
        5: "Květen", 6: "Červen", 7: "Červenec", 8: "Srpen",
        9: "Září", 10: "Říjen", 11: "Listopad", 12: "Prosinec"
    }
    
    # Názvy dnů v týdnu
    CZECH_WEEKDAYS = ["Po", "Út", "St", "Čt", "Pá", "So", "Ne"]
    
    def __init__(self, excel_path):
        """
        Inicializace manageru.
        
        Args:
            excel_path (str|Path): Cesta k adresáři s Excel soubory
        """
        self.excel_path = Path(excel_path)
        self.workbook_name = "Hodiny2025.xlsx"
        self.template_sheet_name = "MMhod25"
        self.file_path = self.excel_path / self.workbook_name
        
        # Vytvoř Excel soubor, pokud neexistuje
        self._ensure_excel_file_exists()
        
        logger.info(f"Hodiny2025Manager inicializován pro soubor: {self.file_path}")
    
    def _ensure_excel_file_exists(self):
        """Zajistí, že Excel soubor existuje. Pokud ne, vytvoří ho s template."""
        if not self.file_path.exists():
            logger.info(f"Vytváří se nový soubor: {self.file_path}")
            self._create_new_workbook()
    
    def _create_new_workbook(self):
        """Vytvoří nový Excel soubor s template listem."""
        workbook = Workbook()
        
        # Odstraň výchozí list
        default_sheet = workbook.active
        workbook.remove(default_sheet)
        
        # Vytvoř template list
        template_sheet = workbook.create_sheet(title=self.template_sheet_name)
        self._setup_template_sheet(template_sheet)
        
        # Vytvoř list pro aktuální měsíc
        current_month = datetime.now().month
        current_sheet_name = f"{current_month:02d}hod25"
        current_sheet = workbook.copy_worksheet(template_sheet)
        current_sheet.title = current_sheet_name
        self._setup_month_sheet(current_sheet, current_month, 2025)
        
        workbook.save(self.file_path)
        logger.info(f"Vytvořen nový Excel soubor: {self.file_path}")
    
    def _setup_template_sheet(self, sheet: Worksheet):
        """
        Nastaví template list s formátováním a vzorci.
        
        Args:
            sheet: Worksheet objekt
        """
        # Záhlaví
        sheet['A1'] = "Měsíční výkaz práce - [Měsíc] 2025"
        sheet['B1'] = "Hodiny Evidence System"
        
        # Hlavičky sloupců
        headers = [
            "Den", "Datum", "Den v týdnu", "Svátek",
            "Začátek", "Oběd (h)", "Konec", "Celkem hodin",
            "Přesčasy", "Noční práce", "Víkend", "Svátky",
            "Zaměstnanci", "Celkem odpracováno"
        ]
        
        for col, header in enumerate(headers, 1):
            cell = sheet.cell(row=self.HEADER_ROW, column=col)
            cell.value = header
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
        
        # Formátování a vzorce pro datovou oblast
        for day in range(1, 32):  # 1-31 dní
            row = self.DATA_START_ROW + day - 1
            
            # Den v měsíci
            sheet.cell(row=row, column=self.COL_DAY).value = day
            
            # Vzorce
            # Celkem hodin = (konec - začátek) * 24 - oběd
            sheet.cell(row=row, column=self.COL_TOTAL_HOURS).value = \
                f"=IF(AND(E{row}<>\"\",G{row}<>\"\"),(G{row}-E{row})*24-F{row},0)"
            
            # Přesčasy = max(0, celkem - 8)
            sheet.cell(row=row, column=self.COL_OVERTIME).value = \
                f"=MAX(0,H{row}-8)"
            
            # Celkem za všechny = hodiny * počet zaměstnanců
            sheet.cell(row=row, column=self.COL_TOTAL_ALL).value = \
                f"=H{row}*M{row}"
        
        # Souhrny
        sheet.cell(row=self.SUMMARY_ROW, column=1).value = "SOUHRN:"
        sheet.cell(row=self.SUMMARY_ROW, column=self.COL_TOTAL_HOURS).value = \
            f"=SUM(H{self.DATA_START_ROW}:H{self.DATA_END_ROW})"
        sheet.cell(row=self.SUMMARY_ROW, column=self.COL_OVERTIME).value = \
            f"=SUM(I{self.DATA_START_ROW}:I{self.DATA_END_ROW})"
        sheet.cell(row=self.SUMMARY_ROW, column=self.COL_TOTAL_ALL).value = \
            f"=SUM(N{self.DATA_START_ROW}:N{self.DATA_END_ROW})"
        
        # Styling pro souhrny
        for col in [1, self.COL_TOTAL_HOURS, self.COL_OVERTIME, self.COL_TOTAL_ALL]:
            cell = sheet.cell(row=self.SUMMARY_ROW, column=col)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
    
    def _setup_month_sheet(self, sheet: Worksheet, month: int, year: int):
        """
        Nastaví specifický měsíční list s daty a formátováním.
        
        Args:
            sheet: Worksheet objekt
            month: Číslo měsíce (1-12)
            year: Rok
        """
        # Aktualizuj název
        month_name = self.CZECH_MONTHS[month]
        sheet['A1'] = f"Měsíční výkaz práce - {month_name} {year}"
        
        # Naplň data pro dny v měsíci
        days_in_month = calendar.monthrange(year, month)[1]
        
        for day in range(1, days_in_month + 1):
            row = self.DATA_START_ROW + day - 1
            date_obj = datetime(year, month, day)
            
            # Datum
            sheet.cell(row=row, column=self.COL_DATE).value = date_obj.strftime("%d.%m.%Y")
            
            # Den v týdnu
            weekday = self.CZECH_WEEKDAYS[date_obj.weekday()]
            sheet.cell(row=row, column=self.COL_WEEKDAY).value = weekday
            
            # Víkendové označení
            if date_obj.weekday() >= 5:  # Sobota, Neděle
                for col in range(1, 15):
                    cell = sheet.cell(row=row, column=col)
                    cell.fill = PatternFill(start_color="FFE6E6", end_color="FFE6E6", fill_type="solid")
        
        # Vymaž řádky pro dny, které v měsíci nejsou
        for day in range(days_in_month + 1, 32):
            row = self.DATA_START_ROW + day - 1
            for col in range(1, 15):
                sheet.cell(row=row, column=col).value = ""
    
    def get_or_create_month_sheet(self, month: int, year: int = 2025) -> Worksheet:
        """
        Získá nebo vytvoří list pro zadaný měsíc.
        
        Args:
            month: Číslo měsíce (1-12)
            year: Rok (default 2025)
            
        Returns:
            Worksheet objekt
        """
        sheet_name = f"{month:02d}hod{str(year)[2:]}"
        
        try:
            workbook = load_workbook(self.file_path)
        except (FileNotFoundError, InvalidFileException):
            self._create_new_workbook()
            workbook = load_workbook(self.file_path)
        
        if sheet_name not in workbook.sheetnames:
            # Zkopíruj template
            if self.template_sheet_name in workbook.sheetnames:
                template_sheet = workbook[self.template_sheet_name]
                new_sheet = workbook.copy_worksheet(template_sheet)
                new_sheet.title = sheet_name
                self._setup_month_sheet(new_sheet, month, year)
                workbook.save(self.file_path)
                logger.info(f"Vytvořen nový list: {sheet_name}")
            else:
                raise ValueError(f"Template list '{self.template_sheet_name}' nebyl nalezen")
        
        return workbook[sheet_name]
    
    def zapis_pracovni_doby(self, date: str, start_time: str, end_time: str, 
                           lunch_duration: str, num_employees: int):
        """
        Zapíše pracovní dobu do Excel souboru.
        
        Args:
            date: Datum ve formátu YYYY-MM-DD
            start_time: Čas začátku ve formátu HH:MM
            end_time: Čas konce ve formátu HH:MM
            lunch_duration: Délka oběda v hodinách (string)
            num_employees: Počet zaměstnanců
        """
        try:
            date_obj = datetime.strptime(date, "%Y-%m-%d")
            month = date_obj.month
            year = date_obj.year
            day = date_obj.day
            
            # Získej nebo vytvoř list pro měsíc (tato metoda už spravuje workbook)
            sheet = self.get_or_create_month_sheet(month, year)
            
            # Načti aktuální workbook
            workbook = load_workbook(self.file_path)
            sheet = workbook[sheet.title]  # Získej aktuální verzi sheetu
            
            # Vypočti řádek pro daný den
            row = self.DATA_START_ROW + day - 1
            
            # Zápis dat do buněk
            if start_time and start_time != "00:00":
                sheet.cell(row=row, column=self.COL_START).value = \
                    datetime.strptime(start_time, "%H:%M").time()
            
            if end_time and end_time != "00:00":
                sheet.cell(row=row, column=self.COL_END).value = \
                    datetime.strptime(end_time, "%H:%M").time()
            
            # Oběd jako číslo (ne čas!) - vždy nastavíme hodnotu
            lunch_hours = float(lunch_duration) if lunch_duration else 0.0
            lunch_cell = sheet.cell(row=row, column=self.COL_LUNCH)
            lunch_cell.value = lunch_hours
            lunch_cell.number_format = '0.0'  # Explicitně nastavit jako číselný formát
            
            # Počet zaměstnanců - vždy nastavíme hodnotu
            employee_cell = sheet.cell(row=row, column=self.COL_EMPLOYEES)
            employee_cell.value = num_employees if num_employees > 0 else 0
            
            # Ujisti se, že formule jsou správně nastavené
            if not sheet.cell(row=row, column=self.COL_TOTAL_HOURS).value or \
               not str(sheet.cell(row=row, column=self.COL_TOTAL_HOURS).value).startswith("="):
                sheet.cell(row=row, column=self.COL_TOTAL_HOURS).value = \
                    f"=IF(AND(E{row}<>\"\",G{row}<>\"\"),(G{row}-E{row})*24-F{row},0)"
            
            if not sheet.cell(row=row, column=self.COL_OVERTIME).value or \
               not str(sheet.cell(row=row, column=self.COL_OVERTIME).value).startswith("="):
                sheet.cell(row=row, column=self.COL_OVERTIME).value = \
                    f"=MAX(0,H{row}-8)"
            
            if not sheet.cell(row=row, column=self.COL_TOTAL_ALL).value or \
               not str(sheet.cell(row=row, column=self.COL_TOTAL_ALL).value).startswith("="):
                sheet.cell(row=row, column=self.COL_TOTAL_ALL).value = \
                    f"=H{row}*M{row}"
            
            # Uložení a vynucení automatického přepočtu formulí
            try:
                workbook.calculation.calcMode = 'auto'
            except Exception:  # kompatibilita pokud atribut není dostupný
                pass
            workbook.save(self.file_path)
            logger.info(
                f"Pracovní doba pro {date} byla zapsána do listu {sheet.title}, řádek {row}, "
                f"sloupce E-F-G-M: {start_time}-{lunch_duration}-{end_time}-{num_employees}"
            )
            
        except Exception as e:
            logger.error(f"Chyba při zápisu pracovní doby pro {date}: {e}", exc_info=True)
            raise
    
    def get_monthly_summary(self, month: int, year: int = 2025) -> dict:
        """
        Získá měsíční souhrn z Excel souboru.
        
        Args:
            month: Číslo měsíce (1-12)
            year: Rok (default 2025)
            
        Returns:
            Dictionary se souhrnem dat
        """
        try:
            sheet = self.get_or_create_month_sheet(month, year)
            
            # Načti souhrny z řádku SUMMARY_ROW
            total_hours = sheet.cell(row=self.SUMMARY_ROW, column=self.COL_TOTAL_HOURS).value or 0
            total_overtime = sheet.cell(row=self.SUMMARY_ROW, column=self.COL_OVERTIME).value or 0
            total_all_employees = sheet.cell(row=self.SUMMARY_ROW, column=self.COL_TOTAL_ALL).value or 0
            
            return {
                'month': month,
                'year': year,
                'month_name': self.CZECH_MONTHS[month],
                'total_hours': float(total_hours) if isinstance(total_hours, (int, float)) else 0,
                'total_overtime': float(total_overtime) if isinstance(total_overtime, (int, float)) else 0,
                'total_all_employees': float(total_all_employees) if isinstance(total_all_employees, (int, float)) else 0,
                'sheet_name': f"{month:02d}hod{str(year)[2:]}"
            }
            
        except Exception as e:
            logger.error(f"Chyba při získávání měsíčního souhrnu pro {month}/{year}: {e}")
            return {
                'month': month, 'year': year, 'month_name': self.CZECH_MONTHS.get(month, 'Neznámý'),
                'total_hours': 0, 'total_overtime': 0, 'total_all_employees': 0,
                'sheet_name': f"{month:02d}hod{str(year)[2:]}", 'error': str(e)
            }
    
    def get_daily_record(self, date: str) -> dict:
        """
        Získá záznam pro konkrétní den.
        
        Args:
            date: Datum ve formátu YYYY-MM-DD
            
        Returns:
            Dictionary s údaji o dni
        """
        try:
            date_obj = datetime.strptime(date, "%Y-%m-%d")
            month = date_obj.month
            year = date_obj.year
            day = date_obj.day
            
            # Načti existující workbook s vypočítanými hodnotami
            try:
                workbook = load_workbook(self.file_path, data_only=True)
            except (FileNotFoundError, InvalidFileException):
                return {'date': date, 'error': 'Excel soubor nenalezen'}
            
            sheet_name = f"{month:02d}hod{str(year)[2:]}"
            
            if sheet_name not in workbook.sheetnames:
                return {'date': date, 'error': f'List {sheet_name} nenalezen'}
            
            sheet = workbook[sheet_name]
            row = self.DATA_START_ROW + day - 1
            
            # Načti data z buněk
            start_time = sheet.cell(row=row, column=self.COL_START).value
            end_time = sheet.cell(row=row, column=self.COL_END).value
            lunch_hours = sheet.cell(row=row, column=self.COL_LUNCH).value or 0
            total_hours = sheet.cell(row=row, column=self.COL_TOTAL_HOURS).value or 0
            overtime = sheet.cell(row=row, column=self.COL_OVERTIME).value or 0
            num_employees = sheet.cell(row=row, column=self.COL_EMPLOYEES).value or 0
            total_all = sheet.cell(row=row, column=self.COL_TOTAL_ALL).value or 0
            
            # Převod hodnot z buněk na správné typy
            def safe_time_format(value):
                """Bezpečně převede hodnotu na časový formát"""
                if value is None:
                    return None
                if isinstance(value, time):
                    return value.strftime("%H:%M")
                if isinstance(value, str):
                    return value
                if isinstance(value, (int, float)):
                    # Excel čas jako desetinné číslo (např. 0.5 = 12:00)
                    hours = int(value * 24)
                    minutes = int((value * 24 * 60) % 60)
                    return f"{hours:02d}:{minutes:02d}"
                return str(value)
            
            def safe_float_convert(value):
                """Bezpečně převede hodnotu na float"""
                if value is None:
                    return 0.0
                if isinstance(value, (int, float)):
                    return float(value)
                if isinstance(value, str):
                    try:
                        return float(value)
                    except ValueError:
                        return 0.0
                return 0.0
            
            def safe_int_convert(value):
                """Bezpečně převede hodnotu na int"""
                if value is None:
                    return 0
                if isinstance(value, (int, float)):
                    return int(value)
                if isinstance(value, str):
                    try:
                        return int(value)
                    except ValueError:
                        return 0
                return 0
            
            def calculate_hours_manually(start_val, end_val, lunch_val):
                """Ručně spočítá hodiny, pokud formule nevrací hodnoty"""
                if not start_val or not end_val:
                    return 0.0
                
                try:
                    if isinstance(start_val, time) and isinstance(end_val, time):
                        # Převod času na hodiny
                        start_hours = start_val.hour + start_val.minute / 60.0
                        end_hours = end_val.hour + end_val.minute / 60.0
                        
                        # Celková doba minus oběd
                        total_hours = end_hours - start_hours - safe_float_convert(lunch_val)
                        return max(0.0, total_hours)
                except Exception:
                    pass
                return 0.0
            
            # Pokud formule nevrací hodnoty, spočítáme ručně
            calculated_hours = 0.0
            if safe_float_convert(total_hours) == 0.0 and start_time and end_time:
                calculated_hours = calculate_hours_manually(start_time, end_time, lunch_hours)
                total_hours = calculated_hours
            
            calculated_overtime = max(0.0, safe_float_convert(total_hours) - 8.0)
            if safe_float_convert(overtime) == 0.0 and calculated_hours > 8.0:
                overtime = calculated_overtime
                
            calculated_total_all = safe_float_convert(total_hours) * safe_int_convert(num_employees)
            if safe_float_convert(total_all) == 0.0 and calculated_hours > 0:
                total_all = calculated_total_all
            
            return {
                'date': date,
                'day': day,
                'start_time': safe_time_format(start_time),
                'end_time': safe_time_format(end_time),
                'lunch_hours': safe_float_convert(lunch_hours),
                'total_hours': safe_float_convert(total_hours),
                'overtime': safe_float_convert(overtime),
                'num_employees': safe_int_convert(num_employees),
                'total_all_employees': safe_float_convert(total_all),
                'row': row,
                'sheet_name': sheet_name
            }
            
        except Exception as e:
            logger.error(f"Chyba při získávání záznamu pro {date}: {e}")
            return {'date': date, 'error': str(e)}
    
    def create_test_data(self):
        """Vytvoří testovací data pro ověření funkčnosti."""
        logger.info("Vytváří se testovací data pro Hodiny2025.xlsx")
        
        # Testovací data pro první týden ledna 2025
        test_dates = [
            ("2025-01-02", "07:00", "15:30", "0.5", 3),  # Čtvrtek
            ("2025-01-03", "07:00", "16:00", "1.0", 3),  # Pátek
            ("2025-01-06", "08:00", "16:30", "0.5", 2),  # Pondělí
            ("2025-01-07", "07:30", "15:30", "1.0", 4),  # Úterý
            ("2025-01-08", "07:00", "17:00", "1.0", 3),  # Středa (přesčas)
        ]
        
        for date, start, end, lunch, employees in test_dates:
            try:
                self.zapis_pracovni_doby(date, start, end, lunch, employees)
                logger.info(f"✅ Testovací záznam vytvořen: {date}")
            except Exception as e:
                logger.error(f"❌ Chyba při vytváření testovacího záznamu {date}: {e}")
        
        # Vytvoř několik měsíčních listů pro test
        for month in [1, 2, 3]:  # Leden, Únor, Březen
            try:
                sheet = self.get_or_create_month_sheet(month, 2025)
                logger.info(f"✅ Vytvořen list pro měsíc {month}: {sheet.title}")
            except Exception as e:
                logger.error(f"❌ Chyba při vytváření listu pro měsíc {month}: {e}")
        
        logger.info("Testovací data byla vytvořena!")
        
    def validate_data_integrity(self) -> dict:
        """
        Ověří integritu dat v Excel souboru.
        
        Returns:
            Dictionary s výsledky validace
        """
        results = {
            'valid': True,
            'errors': [],
            'warnings': [],
            'sheets_checked': [],
            'records_checked': 0
        }
        
        try:
            workbook = load_workbook(self.file_path)
            
            for sheet_name in workbook.sheetnames:
                if sheet_name == self.template_sheet_name:
                    continue
                    
                results['sheets_checked'].append(sheet_name)
                sheet = workbook[sheet_name]
                
                # Zkontroluj vzorce a data
                for row in range(self.DATA_START_ROW, self.DATA_END_ROW + 1):
                    results['records_checked'] += 1
                    
                    # Zkontroluj, zda jsou vzorce správné
                    total_formula = sheet.cell(row=row, column=self.COL_TOTAL_HOURS).value
                    if isinstance(total_formula, str) and not total_formula.startswith('='):
                        results['warnings'].append(f"List {sheet_name}, řádek {row}: Chybí vzorec pro celkové hodiny")
                    
                    # Zkontroluj konzistenci dat
                    start_time = sheet.cell(row=row, column=self.COL_START).value
                    end_time = sheet.cell(row=row, column=self.COL_END).value
                    
                    if (start_time and not end_time) or (not start_time and end_time):
                        results['warnings'].append(f"List {sheet_name}, řádek {row}: Nekonzistentní časy začátku/konce")
            
        except Exception as e:
            results['valid'] = False
            results['errors'].append(f"Chyba při validaci: {e}")
        
        return results
