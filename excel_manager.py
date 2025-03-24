import openpyxl
import os
import logging
import shutil
import re
from datetime import datetime 
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.copier import WorksheetCopy
import fcntl
import contextlib
from threading import Lock
from utils.logger import setup_logger
from contextlib import contextmanager

logger = setup_logger('excel_manager')

class ExcelManager:
    def __init__(self, base_path, excel_file_name):
        self.base_path = base_path
        self.file_path = os.path.join(self.base_path, excel_file_name)
        self.file_path_2025 = os.path.join(self.base_path, 'Hodiny2025.xlsx')
        self.current_project_name = None
        self._file_lock = Lock()
        self._workbook_cache = {}

    @contextmanager 
    def _get_workbook(self, file_path, read_only=False):
        """Vylepšený context manager pro práci s workbookem"""
        cache_key = os.path.abspath(file_path)
        
        try:
            if cache_key in self._workbook_cache:
                wb = self._workbook_cache[cache_key]
                if not wb.is_active:  # Kontrola, zda workbook není uzavřený
                    del self._workbook_cache[cache_key]
                    wb = None
            else:
                wb = None

            if wb is None:
                with self._file_lock:
                    if os.path.exists(file_path):
                        wb = load_workbook(file_path, read_only=read_only)
                    else:
                        os.makedirs(os.path.dirname(file_path), exist_ok=True)
                        wb = Workbook()
                        wb.save(file_path)
                    self._workbook_cache[cache_key] = wb

            yield wb
            
            if not read_only:
                wb.save(file_path)
                
        except Exception as e:
            logger.error(f"Chyba při práci s workbookem {file_path}: {e}")
            raise

    def _clear_workbook_cache(self):
        """Vylepšená metoda pro čištění cache"""
        for path, wb in list(self._workbook_cache.items()):
            try:
                if wb.is_active:  # Kontrola, zda workbook není již uzavřený
                    wb.save(path)
                    wb.close()
            except Exception as e:
                logger.error(f"Chyba při ukládání a zavírání workbooku {path}: {e}")
            finally:
                self._workbook_cache.pop(path, None)

    def __del__(self):
        """Destruktor - zajistí uvolnění všech prostředků"""
        self._clear_workbook_cache()

    def ulozit_pracovni_dobu(self, date, start_time, end_time, lunch_duration, employees):
        """Uloží pracovní dobu do Excel souboru"""
        try:
            with self._get_workbook(self.file_path) as workbook:
                # Získání čísla týdne z datumu
                week_number = self.ziskej_cislo_tydne(date)
                week_number = week_number.week

                sheet_name = f"Týden {week_number}"

                if sheet_name not in workbook.sheetnames:
                    if "Týden" in workbook.sheetnames:
                        source_sheet = workbook["Týden"]
                        sheet = workbook.copy_worksheet(source_sheet)
                        sheet.title = sheet_name
                    else:
                        sheet = workbook.create_sheet(sheet_name)
                    sheet['A80'] = sheet_name
                else:
                    sheet = workbook[sheet_name]

                # Výpočet odpracovaných hodin
                start = datetime.strptime(start_time, "%H:%M")
                end = datetime.strptime(end_time, "%H:%M")
                total_hours = (end - start).total_seconds() / 3600 - lunch_duration

                # Určení sloupce podle dne v týdnu (0 = pondělí, 1 = úterý, atd.)
                weekday = datetime.strptime(date, '%Y-%m-%d').weekday()
                # Pro každý den posuneme o 2 sloupce (B,D,F,H,J)
                day_column = chr(ord('B') + 2 * weekday)

                # Ukládání dat pro každého zaměstnance
                start_row = 9
                for employee in employees:
                    # Hledání řádku pro zaměstnance
                    current_row = start_row
                    row_found = False
                    
                    while not row_found:
                        cell_value = sheet[f'A{current_row}'].value
                        if cell_value is None or cell_value == employee:
                            if cell_value is None:
                                # Prázdný řádek - přidáme nového zaměstnance
                                sheet[f'A{current_row}'] = employee
                            sheet[f'{day_column}{current_row}'] = total_hours
                            row_found = True
                        else:
                            current_row += 1

                # Ukládání časů začátku a konce do řádku 7
                sheet[f"{day_column}7"] = start_time
                sheet[f"{chr(ord(day_column)+1)}7"] = end_time
                
                # Uložení data do buňky v řádku 80
                sheet[f"{day_column}80"] = date

                # Zápis názvu projektu do B79
                if self.current_project_name:
                    sheet['B79'] = self.current_project_name

                workbook.save(self.file_path)
                logger.info(f"Úspěšně uložena pracovní doba pro datum {date}")
                return True

        except Exception as e:
            logger.error(f"Chyba při ukládání pracovní doby: {e}")
            return False

    def update_project_info(self, project_name, start_date, end_date=None):
        """Aktualizuje informace o projektu"""
        try:
            with self._get_workbook(self.file_path) as workbook:
                # Nastavíme název projektu pro použití v ulozit_pracovni_dobu
                self.set_project_name(project_name)

                if 'Zálohy' not in workbook.sheetnames:
                    workbook.create_sheet('Zálohy')

                zalohy_sheet = workbook['Zálohy']
                zalohy_sheet['A79'] = project_name

                start_date_obj = datetime.strptime(start_date, '%Y-%m-%d')
                zalohy_sheet['C81'] = start_date_obj.strftime('%d.%m.%y')

                if end_date:
                    end_date_obj = datetime.strptime(end_date, '%Y-%m-%d')
                    zalohy_sheet['D81'] = end_date_obj.strftime('%d.%m.%y')

                workbook.save(self.file_path)
                logger.info(f"Aktualizovány informace o projektu: {project_name}")
                return True
        except Exception as e:
            logger.error(f"Chyba při aktualizaci informací o projektu: {e}")
            return False

    def get_advance_options(self):
        """Získá možnosti záloh z Excel souboru"""
        try:
            with self._get_workbook(self.file_path) as workbook:
                options = []

                if 'Zálohy' in workbook.sheetnames:
                    zalohy_sheet = workbook['Zálohy']
                    option1 = zalohy_sheet['B80'].value or 'Option 1'
                    option2 = zalohy_sheet['D80'].value or 'Option 2'
                    options = [option1, option2]
                    logger.info(f"Načteny možnosti záloh: {options}")
                else:
                    logger.warning("List 'Zálohy' nebyl nalezen v Excel souboru")
                    options = ['Option 1', 'Option 2']

                return options
        except Exception as e:
            logger.error(f"Chyba při načítání možností záloh: {str(e)}")
            return ['Option 1', 'Option 2']

    def save_advance(self, employee_name, amount, currency, option, date):
        """Vylepšená metoda pro ukládání zálohy"""
        try:
            # Použití dvou context managerů najednou
            with self._get_workbook(self.file_path) as wb1, \
                 self._get_workbook(self.file_path_2025) as wb2:
                
                # Uložení do Hodiny_Cap.xlsx
                self._save_advance_main(wb1, employee_name, amount, currency, option, date)
                
                # Uložení do Hodiny2025.xlsx
                self._save_advance_zalohy25(wb2, employee_name, amount, currency, date)
                self._save_advance_cash25(wb2, employee_name, amount, currency, date)
                
                return True

        except Exception as e:
            logger.error(f"Chyba při ukládání zálohy: {str(e)}")
            return False

    def _save_advance_main(self, workbook, employee_name, amount, currency, option, date):
        """Pomocná metoda pro ukládání do hlavního workbooku"""
        if 'Zálohy' not in workbook.sheetnames:
            sheet = workbook.create_sheet('Zálohy')
            sheet['B80'] = 'Option 1'
            sheet['D80'] = 'Option 2'
        else:
            sheet = workbook['Zálohy']

        row = 9
        while row < 1000:
            if not sheet[f'A{row}'].value:
                sheet[f'A{row}'] = employee_name
                break
            if sheet[f'A{row}'].value == employee_name:
                break
            row += 1

        option1_value = sheet['B80'].value or 'Option 1'
        option2_value = sheet['D80'].value or 'Option 2'

        if option == option1_value:
            column = 'B' if currency == 'EUR' else 'C'
        elif option == option2_value:
            column = 'D' if currency == 'EUR' else 'E'
        else:
            raise ValueError(f"Neplatná volba: {option}")

        current_value = sheet[f'{column}{row}'].value
        if current_value is None:
            current_value = 0

        sheet[f'{column}{row}'] = current_value + float(amount)

        # Přidání data zálohy
        date_column = 26  # Předpokládáme, že datum bude v sloupci Z
        sheet.cell(row=row, column=date_column, value=datetime.strptime(date, '%Y-%m-%d').date())

    def _save_advance_zalohy25(self, workbook, employee_name, amount, currency, date):
        if 'Zalohy25' not in workbook.sheetnames:
            workbook.create_sheet('Zalohy25')
        sheet = workbook['Zalohy25']

        # Hledání řádku pro zaměstnance
        row = 3  # Začínáme od řádku 3
        found = False
        while sheet.cell(row=row, column=1).value:
            if sheet.cell(row=row, column=1).value == employee_name:
                found = True
                break
            row += 1

        if not found:
            # Nový zaměstnanec
            sheet.cell(row=row, column=1).value = employee_name

        # Datum
        date_obj = datetime.strptime(date, '%Y-%m-%d').date()
        min_date = sheet.cell(row=row, column=2).value
        max_date = sheet.cell(row=row, column=3).value
        if min_date is None or date_obj < min_date:
            sheet.cell(row=row, column=2).value = date_obj
        if max_date is None or date_obj > max_date:
            sheet.cell(row=row, column=3).value = date_obj

        # Částka
        eur_column = 4
        czk_column = 5
        if currency == 'EUR':
            current_eur = sheet.cell(row=row, column=eur_column).value or 0
            sheet.cell(row=row, column=eur_column).value = current_eur + amount
        elif currency == 'CZK':
            current_czk = sheet.cell(row=row, column=czk_column).value or 0
            sheet.cell(row=row, column=czk_column).value = current_czk + amount

    def _save_advance_cash25(self, workbook, employee_name, amount, currency, date):
        if '(pp)cash25' not in workbook.sheetnames:
            workbook.create_sheet('(pp)cash25')
        sheet = workbook['(pp)cash25']

        # Hledání buňky se jménem zaměstnance nebo "SloupecN"
        row = 1
        col = 1
        found = False
        while col < sheet.max_column + 1:
            row = 1
            while row < sheet.max_row + 1:
                cell_value = sheet.cell(row=row, column=col).value
                if cell_value == employee_name or cell_value == f"Sloupec{col}":
                    found = True
                    break
                row += 1
            if found:
                break
            col += 1

        if not found:
            # Pokud není nalezen ani zaměstnanec ani "SloupecN", použijeme další sloupec
            col += 1
            sheet.cell(row=1, column=col).value = employee_name

        # Zápis dat
        row = self._get_next_empty_row_in_column(sheet, col)
        sheet.cell(row=row, column=col).value = amount
        sheet.cell(row=row, column=col + 1).value = datetime.strptime(date, '%Y-%m-%d').date()

    def _get_next_empty_row_in_column(self, sheet, col):
        row = 1
        while sheet.cell(row=row, column=col).value is not None:
            row += 1
        return row

    def ziskej_cislo_tydne(self, datum):
        """
        Získá číslo týdne pro zadané datum.
        
        Args:
            datum: Datum jako string ('YYYY-MM-DD') nebo datetime objekt
            
        Returns:
            int: Číslo týdne (1-53)
        """
        try:
            if isinstance(datum, str):
                datum = datetime.strptime(datum, '%Y-%m-%d')
            return datum.isocalendar()
        except (ValueError, TypeError) as e:
            logger.error(f"Chyba při zpracování data: {e}")
            current_date = datetime.now()
            return current_date.isocalendar()

if __name__ == "__main__":
    # Test ukládání zálohy
    logging.basicConfig(level=logging.INFO)
    manager = ExcelManager("./data")
    success = manager.save_advance("Test User", 100, "EUR", "Option 1")
    print(f"Test uložení zálohy: {'úspěšný' if success else 'neúspěšný'}")
