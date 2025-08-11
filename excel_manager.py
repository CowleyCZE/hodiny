# excel_manager.py
from config import Config
import logging
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from threading import Lock

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

try:
    from utils.logger import setup_logger
    logger = setup_logger("excel_manager")
except ImportError:
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger("excel_manager")


class ExcelManager:
    def __init__(self, base_path, active_filename, template_filename):
        self.base_path = Path(base_path)
        self.active_filename = active_filename
        self.template_filename = template_filename
        self.active_file_path = self.base_path / self.active_filename if self.active_filename else None
        self.template_file_path = self.base_path / self.template_filename
        self.current_project_name = None
        self._file_lock = Lock()
        self._workbook_cache = {}
        logger.info(f"ExcelManager inicializován pro: {self.active_file_path}")

    def get_active_file_path(self):
        if not self.active_file_path:
            raise ValueError("Aktivní Excel soubor není definován.")
        return self.active_file_path

    @contextmanager
    def _get_workbook(self, file_path_to_open, read_only=False):
        file_path = Path(file_path_to_open)
        cache_key = str(file_path.absolute())
        wb = None
        with self._file_lock:
            if cache_key in self._workbook_cache:
                try:
                    wb = self._workbook_cache[cache_key]
                    _ = wb.sheetnames  # Check if workbook is alive
                except Exception:
                    wb = None
            if wb is None:
                if not file_path.exists():
                    raise FileNotFoundError(f"Soubor '{file_path.name}' nenalezen.")
                try:
                    wb = load_workbook(filename=str(file_path), read_only=read_only, data_only=not read_only)
                    if not read_only:
                        self._workbook_cache[cache_key] = wb
                except Exception as e:
                    raise IOError(f"Chyba při otevírání souboru '{file_path.name}': {e}")
        try:
            yield wb
        finally:
            if read_only and wb:
                wb.close()

    def ulozit_pracovni_dobu(self, date, start_time, end_time, lunch_duration, employees, week_number=None):
        active_path = self.get_active_file_path()
        try:
            with self._get_workbook(active_path, read_only=False) as workbook:
                if week_number is None:
                    week_number = self.ziskej_cislo_tydne(date).week
                
                sheet_name = f"{Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME} {week_number}"
                if sheet_name not in workbook.sheetnames:
                    if Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME in workbook.sheetnames:
                        source_sheet = workbook[Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME]
                        sheet = workbook.copy_worksheet(source_sheet)
                        sheet.title = sheet_name
                    else:
                        sheet = workbook.create_sheet(sheet_name)
                    sheet["A80"] = sheet_name
                else:
                    sheet = workbook[sheet_name]

                weekday = datetime.strptime(date, "%Y-%m-%d").weekday()
                day_column_index = 2 + 2 * weekday
                
                if start_time == "00:00" and end_time == "00:00":
                    total_hours = 0
                else:
                    start = datetime.strptime(start_time, "%H:%M")
                    end = datetime.strptime(end_time, "%H:%M")
                    total_hours = round(((end - start).total_seconds() / 3600) - float(lunch_duration), 2)
                
                employee_rows = {sheet.cell(row=r, column=1).value: r for r in range(Config.EXCEL_EMPLOYEE_START_ROW, sheet.max_row + 1)}
                next_empty_row = sheet.max_row + 1

                for employee in employees:
                    row = employee_rows.get(employee)
                    if not row:
                        row = next_empty_row
                        sheet.cell(row=row, column=1, value=employee)
                        next_empty_row += 1
                    sheet.cell(row=row, column=day_column_index + 1, value=total_hours).number_format = '0.00'
                
                sheet.cell(row=7, column=day_column_index).value = datetime.strptime(start_time, "%H:%M").time()
                sheet.cell(row=7, column=day_column_index).number_format = 'HH:MM'
                sheet.cell(row=80, column=day_column_index).value = datetime.strptime(date, "%Y-%m-%d").date()
                sheet.cell(row=80, column=day_column_index).number_format = 'DD.MM.YYYY'
                
                if self.current_project_name:
                    project_cell = sheet["B79"]
                    if self.current_project_name not in (project_cell.value or ""):
                        project_cell.value = ", ".join(filter(None, [(project_cell.value or ""), self.current_project_name]))
                
                logger.info(f"Uložena pracovní doba pro {date} do listu {sheet_name}.")
                return True
        except (FileNotFoundError, ValueError, IOError, Exception) as e:
            logger.error(f"Chyba při ukládání pracovní doby: {e}", exc_info=True)
            return False

    def close_cached_workbooks(self):
        with self._file_lock:
            for path_str, wb in self._workbook_cache.items():
                try:
                    wb.save(path_str)
                    wb.close()
                except Exception as e:
                    logger.error(f"Chyba při ukládání/zavírání workbooku {path_str}: {e}", exc_info=True)
            self._workbook_cache.clear()

    def set_project_name(self, project_name):
        self.current_project_name = project_name if project_name else None

    def ziskej_cislo_tydne(self, datum):
        try:
            datum_obj = datetime.strptime(datum, "%Y-%m-%d") if isinstance(datum, str) else datum
            return datum_obj.isocalendar()
        except (ValueError, TypeError) as e:
            logger.error(f"Chyba při zpracování data '{datum}': {e}")
            return None

    def generate_monthly_report(self, month, year, employees=None):
        if not (1 <= month <= 12 and 2000 <= year <= 2100):
            raise ValueError("Neplatný měsíc nebo rok.")
        
        report_data = {}
        active_path = self.get_active_file_path()
        try:
            with self._get_workbook(active_path, read_only=True) as workbook:
                for sheet_name in workbook.sheetnames:
                    if sheet_name.startswith(Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME):
                        sheet = workbook[sheet_name]
                        for r_idx in range(Config.EXCEL_EMPLOYEE_START_ROW, sheet.max_row + 1):
                            employee_name = sheet.cell(row=r_idx, column=1).value
                            if not employee_name or (employees and employee_name not in employees):
                                continue
                            
                            if employee_name not in report_data:
                                report_data[employee_name] = {"total_hours": 0.0, "free_days": 0}

                            for c_idx in range(2, 12, 2):
                                date_val = sheet.cell(row=80, column=c_idx).value
                                if isinstance(date_val, datetime) and date_val.month == month and date_val.year == year:
                                    hours = sheet.cell(row=r_idx, column=c_idx + 1).value
                                    if isinstance(hours, (int, float)):
                                        if hours > 0:
                                            report_data[employee_name]["total_hours"] += hours
                                        else:
                                            report_data[employee_name]["free_days"] += 1
        except (FileNotFoundError, IOError, Exception) as e:
            logger.error(f"Chyba při generování měsíčního reportu: {e}", exc_info=True)
            return {}
        
        return {emp: data for emp, data in report_data.items() if data["total_hours"] > 0 or data["free_days"] > 0}
    
    def update_project_info(self, project_name, start_date, end_date):
        # Tato metoda by měla být implementována, pokud je potřeba.
        # Prozatím vrací True.
        return True
