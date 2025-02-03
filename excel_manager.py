import openpyxl
import os
import logging
import shutil
import re
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.copier import WorksheetCopy

class ExcelManager:
    def __init__(self, base_path):
        self.base_path = base_path
        self.file_path = os.path.join(self.base_path, 'Hodiny_Cap.xlsx')
        self.current_project_name = None

    def _load_or_create_workbook(self):
        try:
            if not os.path.exists(self.base_path):
                os.makedirs(self.base_path)
            
            if os.path.exists(self.file_path):
                workbook = load_workbook(self.file_path)
            else:
                workbook = Workbook()
                workbook.save(self.file_path)
            return workbook
        except Exception as e:
            logging.error(f"Chyba při načítání nebo vytváření Excel souboru: {e}")
            raise

    def ziskej_cislo_tydne(self, datum):
        datum_objekt = datetime.strptime(datum, '%Y-%m-%d').date()
        return datum_objekt.isocalendar()[1]

    def ulozit_pracovni_dobu(self, date, start_time, end_time, lunch_duration, employees):
        workbook = self._load_or_create_workbook()
        week_number = self.ziskej_cislo_tydne(date)
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

        start = datetime.strptime(start_time, "%H:%M")
        end = datetime.strptime(end_time, "%H:%M")
        total_hours = (end - start).total_seconds() / 3600 - lunch_duration

        weekday = datetime.strptime(date, '%Y-%m-%d').weekday()
        day_column = chr(ord('B') + 2 * weekday)

        sheet[f"{day_column}7"] = start_time
        sheet[f"{chr(ord(day_column)+1)}7"] = end_time

        row = 9
        for employee in employees:
            sheet[f"A{row}"] = employee
            sheet[f"{day_column}{row}"] = total_hours
            row += 1

        sheet[f"{day_column}8"] = total_hours
        sheet[f"{day_column}80"] = date

        if self.current_project_name:
            current_project = sheet['B79'].value
            if current_project:
                projects = current_project.split(',')
                if self.current_project_name not in projects:
                    projects.append(self.current_project_name)
                    sheet['B79'] = ','.join(projects)
            else:
                sheet['B79'] = self.current_project_name

        workbook.save(self.file_path)

    def create_weekly_copy(self):
        try:
            last_week = self.get_last_week_number()
            if last_week == 0:
                logging.warning("Nebyl nalezen žádný týden v Excel souboru.")
                return None

            original_dir = os.path.dirname(self.file_path)
            original_filename = os.path.basename(self.file_path)
            name, ext = os.path.splitext(original_filename)
            new_filename = f"{name}_Tyden_{last_week}{ext}"
            new_filepath = os.path.join(original_dir, new_filename)
        
            if os.path.exists(self.file_path):
                shutil.copy(self.file_path, new_filepath)
                logging.info(f"Vytvořena týdenní kopie: {new_filepath}")
                return new_filepath
            else:
                logging.error(f"Zdrojový soubor neexistuje: {self.file_path}")
                return None
        except Exception as e:
            logging.error(f"Chyba při vytváření týdenní kopie: {str(e)}")
            return None

    def get_last_week_number(self):
        try:
            workbook = self._load_or_create_workbook()
            week_numbers = []
            for sheet_name in workbook.sheetnames:
                match = re.search(r'Týden (\d+)', sheet_name)
                if match:
                    week_numbers.append(int(match.group(1)))
            return max(week_numbers) if week_numbers else 0
        except Exception as e:
            logging.error(f"Chyba při získávání čísla posledního týdne: {str(e)}")
            return 0

    def get_file_name_with_week(self):
        last_week = self.get_last_week_number()
        base_name = os.path.basename(self.file_path)
        name, ext = os.path.splitext(base_name)
        return f"{name}_Tyden_{last_week}{ext}"

    def update_project_info(self, project_name, start_date, end_date=None):
        workbook = self._load_or_create_workbook()
        
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

    def get_advance_options(self):
        workbook = self._load_or_create_workbook()
        options = []

        if 'Zálohy' in workbook.sheetnames:
            zalohy_sheet = workbook['Zálohy']
            option1 = zalohy_sheet['B80'].value or 'Option 1'
            option2 = zalohy_sheet['D80'].value or 'Option 2'
            options = [option1, option2]

        return options

    def save_advance(self, employee_name, amount, currency, option):
        try:
            workbook = self._load_or_create_workbook()
            
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
            workbook.save(self.file_path)
            
            logging.info(f"Úspěšně uložena záloha pro {employee_name}: {amount} {currency}")
            return True
            
        except Exception as e:
            logging.error(f"Chyba při ukládání zálohy: {str(e)}")
            return False

import openpyxl
import os
import logging
import shutil
import re
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.copier import WorksheetCopy

class ExcelManager:
    def __init__(self, base_path):
        self.base_path = base_path
        self.file_path = os.path.join(self.base_path, 'Hodiny_Cap.xlsx')
        self.current_project_name = None

    def set_project_name(self, project_name):
        """Nastaví aktuální název projektu"""
        self.current_project_name = project_name
        logging.info(f"Nastaven název projektu: {project_name}")

    def _load_or_create_workbook(self):
        try:
            if not os.path.exists(self.base_path):
                os.makedirs(self.base_path)
            
            if os.path.exists(self.file_path):
                workbook = load_workbook(self.file_path)
            else:
                workbook = Workbook()
                workbook.save(self.file_path)
            return workbook
        except Exception as e:
            logging.error(f"Chyba při načítání nebo vytváření Excel souboru: {e}")
            raise

    def ziskej_cislo_tydne(self, datum):
        datum_objekt = datetime.strptime(datum, '%Y-%m-%d').date()
        return datum_objekt.isocalendar()[1]

    def ulozit_pracovni_dobu(self, date, start_time, end_time, lunch_duration, employees):
        try:
            workbook = self._load_or_create_workbook()
            week_number = self.ziskej_cislo_tydne(date)
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

            start = datetime.strptime(start_time, "%H:%M")
            end = datetime.strptime(end_time, "%H:%M")
            total_hours = (end - start).total_seconds() / 3600 - lunch_duration

            weekday = datetime.strptime(date, '%Y-%m-%d').weekday()
            day_column = chr(ord('B') + 2 * weekday)

            sheet[f"{day_column}7"] = start_time
            sheet[f"{chr(ord(day_column)+1)}7"] = end_time

            row = 9
            for employee in employees:
                sheet[f"A{row}"] = employee
                sheet[f"{day_column}{row}"] = total_hours
                row += 1

            sheet[f"{day_column}8"] = total_hours
            sheet[f"{day_column}80"] = date

            # Zápis názvu projektu
            if self.current_project_name:
                try:
                    # Přímý zápis názvu projektu do buňky B79
                    sheet['B79'] = self.current_project_name
                    logging.info(f"Zapsán název projektu '{self.current_project_name}' do listu {sheet_name}")
                except Exception as e:
                    logging.error(f"Chyba při zápisu názvu projektu: {e}")

            workbook.save(self.file_path)
            logging.info(f"Úspěšně uložena pracovní doba pro datum {date}")
            return True
        except Exception as e:
            logging.error(f"Chyba při ukládání pracovní doby: {e}")
            return False

    def create_weekly_copy(self):
        try:
            last_week = self.get_last_week_number()
            if last_week == 0:
                logging.warning("Nebyl nalezen žádný týden v Excel souboru.")
                return None

            original_dir = os.path.dirname(self.file_path)
            original_filename = os.path.basename(self.file_path)
            name, ext = os.path.splitext(original_filename)
            new_filename = f"{name}_Tyden_{last_week}{ext}"
            new_filepath = os.path.join(original_dir, new_filename)
        
            if os.path.exists(self.file_path):
                shutil.copy(self.file_path, new_filepath)
                logging.info(f"Vytvořena týdenní kopie: {new_filepath}")
                return new_filepath
            else:
                logging.error(f"Zdrojový soubor neexistuje: {self.file_path}")
                return None
        except Exception as e:
            logging.error(f"Chyba při vytváření týdenní kopie: {str(e)}")
            return None

    def get_last_week_number(self):
        try:
            workbook = self._load_or_create_workbook()
            week_numbers = []
            for sheet_name in workbook.sheetnames:
                match = re.search(r'Týden (\d+)', sheet_name)
                if match:
                    week_numbers.append(int(match.group(1)))
            return max(week_numbers) if week_numbers else 0
        except Exception as e:
            logging.error(f"Chyba při získávání čísla posledního týdne: {str(e)}")
            return 0

    def get_file_name_with_week(self):
        last_week = self.get_last_week_number()
        base_name = os.path.basename(self.file_path)
        name, ext = os.path.splitext(base_name)
        return f"{name}_Tyden_{last_week}{ext}"

    def update_project_info(self, project_name, start_date, end_date=None):
        try:
            workbook = self._load_or_create_workbook()
            
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
            logging.info(f"Aktualizovány informace o projektu: {project_name}")
            return True
        except Exception as e:
            logging.error(f"Chyba při aktualizaci informací o projektu: {e}")
            return False

    def get_advance_options(self):
        workbook = self._load_or_create_workbook()
        options = []

        if 'Zálohy' in workbook.sheetnames:
            zalohy_sheet = workbook['Zálohy']
            option1 = zalohy_sheet['B80'].value or 'Option 1'
            option2 = zalohy_sheet['D80'].value or 'Option 2'
            options = [option1, option2]

        return options

    def save_advance(self, employee_name, amount, currency, option):
        try:
            workbook = self._load_or_create_workbook()
            
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
            workbook.save(self.file_path)
            
            logging.info(f"Úspěšně uložena záloha pro {employee_name}: {amount} {currency}")
            return True
            
        except Exception as e:
            logging.error(f"Chyba při ukládání zálohy: {str(e)}")
            return False

if __name__ == "__main__":
    # Test ukládání zálohy
    logging.basicConfig(level=logging.INFO)
    manager = ExcelManager("./data")
    success = manager.save_advance("Test User", 100, "EUR", "Option 1")
    print(f"Test uložení zálohy: {'úspěšný' if success else 'neúspěšný'}")
