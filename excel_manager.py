# excel_manager.py
import shutil
from config import Config
import logging
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from threading import Lock

from openpyxl import load_workbook

try:
    from utils.logger import setup_logger
    logger = setup_logger("excel_manager")
except ImportError:
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger("excel_manager")


class ExcelManager:
    def __init__(self, base_path):
        self.base_path = Path(base_path)
        self.active_filename = Config.EXCEL_TEMPLATE_NAME
        self.active_file_path = self.base_path / self.active_filename
        self.current_project_name = None
        self._file_lock = Lock()
        self._workbook_cache = {}
        logger.info(f"ExcelManager inicializován pro: {self.active_file_path}")

    def get_active_file_path(self):
        return self.active_file_path

    @contextmanager
    def _get_workbook(self, read_only=False):
        cache_key = str(self.active_file_path.absolute())
        wb = None
        with self._file_lock:
            if cache_key in self._workbook_cache:
                try:
                    wb = self._workbook_cache[cache_key]
                    _ = wb.sheetnames
                except Exception:
                    wb = None
            if wb is None:
                if not self.active_file_path.exists():
                    raise FileNotFoundError(f"Soubor '{self.active_filename}' nenalezen.")
                try:
                    wb = load_workbook(
                        filename=str(self.active_file_path), read_only=read_only, data_only=not read_only
                    )
                    if not read_only:
                        self._workbook_cache[cache_key] = wb
                except Exception as e:
                    raise IOError(f"Chyba při otevírání souboru '{self.active_filename}': {e}")
        try:
            yield wb
        finally:
            if read_only and wb:
                wb.close()

    def archive_if_needed(self, current_week_number, settings):
        last_archived_week = settings.get("last_archived_week", 0)
        if current_week_number > last_archived_week:
            archive_filename = f"Hodiny_Cap_Tyden_{last_archived_week}.xlsx"
            archive_path = self.base_path / archive_filename

            # 1. Vytvoření archivní kopie
            shutil.copy(self.active_file_path, archive_path)
            logger.info(f"Archivován soubor: {archive_path}")

            # 2. Vyčištění hlavního souboru
            with self._get_workbook() as wb:
                for sheet_name in wb.sheetnames:
                    if sheet_name.startswith(Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME):
                        wb.remove(wb[sheet_name])

                # Vytvoření nového čistého listu pro aktuální týden
                wb.create_sheet(Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME)

            # 3. Aktualizace nastavení
            settings["last_archived_week"] = current_week_number
            # Tuto funkci bude muset zavolat app.py, protože excel_manager nemá přístup k save_settings_to_file
            return True
        return False

# excel_manager.py
import logging
import shutil
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from threading import Lock

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

from config import Config

try:
    from utils.logger import setup_logger

    logger = setup_logger("excel_manager")
except ImportError:
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger("excel_manager")


class ExcelManager:
    def __init__(self, base_path):
        self.base_path = Path(base_path)
        self.active_filename = Config.EXCEL_TEMPLATE_NAME
        self.active_file_path = self.base_path / self.active_filename
        self.current_project_name = None
        self._file_lock = Lock()
        self._workbook_cache = {}
        logger.info(f"ExcelManager inicializován pro: {self.active_file_path}")

    def get_active_file_path(self):
        return self.active_file_path

    @contextmanager
    def _get_workbook(self, read_only=False):
        cache_key = str(self.active_file_path.absolute())
        wb = None
        with self._file_lock:
            if cache_key in self._workbook_cache:
                try:
                    wb = self._workbook_cache[cache_key]
                    _ = wb.sheetnames
                except Exception:
                    wb = None
            if wb is None:
                if not self.active_file_path.exists():
                    raise FileNotFoundError(f"Soubor '{self.active_filename}' nenalezen.")
                try:
                    wb = load_workbook(
                        filename=str(self.active_file_path), read_only=read_only, data_only=not read_only
                    )
                    if not read_only:
                        self._workbook_cache[cache_key] = wb
                except Exception as e:
                    raise IOError(f"Chyba při otevírání souboru '{self.active_filename}': {e}")
        try:
            yield wb
        finally:
            if read_only and wb:
                wb.close()

    def archive_if_needed(self, current_week_number, settings):
        last_archived_week = settings.get("last_archived_week", 0)
        if current_week_number > last_archived_week:
            archive_filename = f"Hodiny_Cap_Tyden_{last_archived_week}.xlsx"
            archive_path = self.base_path / archive_filename

            # 1. Vytvoření archivní kopie
            shutil.copy(self.active_file_path, archive_path)
            logger.info(f"Archivován soubor: {archive_path}")

            # 2. Vyčištění hlavního souboru
            with self._get_workbook() as wb:
                for sheet_name in wb.sheetnames:
                    if sheet_name.startswith(Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME):
                        wb.remove(wb[sheet_name])

                # Vytvoření nového čistého listu pro aktuální týden
                wb.create_sheet(Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME)

            # 3. Aktualizace nastavení
            settings["last_archived_week"] = current_week_number
            # Tuto funkci bude muset zavolat app.py, protože excel_manager nemá přístup k save_settings_to_file
            return True
        return False

    def ulozit_pracovni_dobu(self, date, start_time, end_time, lunch_duration, employees):
        try:
            week_calendar_data = self.ziskej_cislo_tydne(date)
            if not week_calendar_data:
                logger.error(f"Nepodařilo se získat číslo týdne pro datum: {date}")
                return False
            week_number = week_calendar_data.week
            sheet_name = f"{Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME} {week_number}"

            with self._get_workbook() as workbook:
                if sheet_name not in workbook.sheetnames:
                    if Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME in workbook.sheetnames:
                        template_sheet = workbook[Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME]
                        sheet = workbook.copy_worksheet(template_sheet)
                        sheet.title = sheet_name
                    else:
                        sheet = workbook.create_sheet(sheet_name)
                else:
                    sheet = workbook[sheet_name]

                weekday = datetime.strptime(date, "%Y-%m-%d").weekday()
                day_column_index = 2 + 2 * weekday

                total_hours = round(
                    (
                        (datetime.strptime(end_time, "%H:%M") - datetime.strptime(start_time, "%H:%M")).total_seconds()
                        / 3600
                    )
                    - float(lunch_duration),
                    2,
                )

                employee_rows = {
                    sheet.cell(row=r, column=1).value: r
                    for r in range(Config.EXCEL_EMPLOYEE_START_ROW, sheet.max_row + 1)
                }
                next_empty_row = (
                    sheet.max_row + 1
                    if sheet.max_row >= Config.EXCEL_EMPLOYEE_START_ROW
                    else Config.EXCEL_EMPLOYEE_START_ROW
                )

                for employee in employees:
                    row = employee_rows.get(employee)
                    if not row:
                        row = next_empty_row
                        sheet.cell(row=row, column=1, value=employee)
                        next_empty_row += 1

                    data_cell = sheet.cell(row=row, column=day_column_index + 1)
                    if not isinstance(data_cell, MergedCell):
                        data_cell.value = total_hours
                        data_cell.number_format = "0.00"

                start_time_cell = sheet.cell(row=7, column=day_column_index)
                if not isinstance(start_time_cell, MergedCell):
                    start_time_cell.value = datetime.strptime(start_time, "%H:%M").time()
                    start_time_cell.number_format = "HH:MM"

                date_cell = sheet.cell(row=80, column=day_column_index)
                if not isinstance(date_cell, MergedCell):
                    date_cell.value = datetime.strptime(date, "%Y-%m-%d").date()
                    date_cell.number_format = "DD.MM.YYYY"

                logger.info(f"Uložena pracovní doba pro {date} do listu {sheet_name}.")
                return True
        except (FileNotFoundError, ValueError, IOError) as e:
            logger.error(f"Chyba při ukládání pracovní doby: {e}", exc_info=True)
            return False
        except Exception as e:
            logger.error(f"Neočekávaná chyba při ukládání pracovní doby: {e}", exc_info=True)
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
        try:
            with self._get_workbook(read_only=True) as workbook:
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
        except (FileNotFoundError, IOError) as e:
            logger.error(f"Chyba při generování měsíčního reportu: {e}", exc_info=True)
            return {}
        except Exception as e:
            logger.error(f"Neočekávaná chyba při generování měsíčního reportu: {e}", exc_info=True)
            return {}
        return report_data
    def close_cached_workbooks(self):
        with self._file_lock:
            for path_str, wb in self._workbook_cache.items():
                try:
                    wb.save(path_str)
                    wb.close()
                except Exception as e:
                    logger.error(f"Chyba při ukládání/zavírání workbooku {path_str}: {e}", exc_info=True)
            self._workbook_cache.clear()

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
        try:
            with self._get_workbook(read_only=True) as workbook:
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
        return True
