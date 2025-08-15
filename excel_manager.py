"""Správa týdenního Excel souboru (šablona + dynamické listy týdnů, reporty).

Zodpovědnosti:
 - Lazy/cached otevření workbooku (thread‑safe)
 - Archivace starého týdne a vyčištění listů
 - Zápis denních hodin pro vybrané zaměstnance do listu konkrétního týdne
 - Generace měsíčního reportu agregací z týdenních listů
"""
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
    """Správa a manipulace s hlavním Excel souborem."""

    def __init__(self, base_path):
        """base_path: adresář kde je hlavní soubor (Hodiny_Cap.xlsx)."""
        self.base_path = Path(base_path)
        self.active_filename = Config.EXCEL_TEMPLATE_NAME
        self.active_file_path = self.base_path / self.active_filename
        self.current_project_name = None
        self._file_lock = Lock()
        self._workbook_cache = {}
        logger.info("ExcelManager inicializován pro: %s", self.active_file_path)

    def get_active_file_path(self):
        """Cesta k aktivnímu Excel souboru (šablona)."""
        return self.active_file_path

    @contextmanager
    def _get_workbook(self, read_only=False):
        """Context manager vracející workbook (cache pro read_write režim).

        read_only=True -> vždy otevře nový objekt (neukládá do cache),
        jinak recykluje instanci pro snížení IO.
        """
        cache_key = str(self.active_file_path.absolute())
        wb = None
        with self._file_lock:
            if cache_key in self._workbook_cache:
                try:
                    wb = self._workbook_cache[cache_key]
                    if wb:
                        _ = wb.sheetnames  # Test if workbook is alive
                except Exception:
                    wb = None
            if wb is None:
                if not self.active_file_path.exists():
                    raise FileNotFoundError(f"Soubor '{self.active_filename}' nenalezen.")
                try:
                    wb = load_workbook(
                        filename=str(self.active_file_path),
                        read_only=read_only,
                        data_only=not read_only,
                    )
                    if not read_only:
                        self._workbook_cache[cache_key] = wb
                except Exception as e:
                    raise IOError(
                        f"Chyba při otevírání souboru '{self.active_filename}': {e}"
                    ) from e
        try:
            yield wb
        finally:
            if read_only and wb:
                wb.close()

    def archive_if_needed(self, current_week_number, settings):
        """Archivuje minulé týdny při posunu čísla týdne."""
        last_archived_week = settings.get("last_archived_week", 0)
        if current_week_number <= last_archived_week:
            return False

        archive_filename = f"Hodiny_Cap_Tyden_{last_archived_week}.xlsx"
        archive_path = self.base_path / archive_filename

        try:
            shutil.copy(self.active_file_path, archive_path)
            logger.info("Archivován soubor: %s", archive_path)

            with self._get_workbook() as wb:
                if not wb:
                    raise IOError("Workbook se nepodařilo otevřít pro archivaci.")
                for sheet_name in wb.sheetnames:
                    if sheet_name.startswith(Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME):
                        wb.remove(wb[sheet_name])

                wb.create_sheet(Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME)

            settings["last_archived_week"] = current_week_number
            return True
        except (IOError, FileNotFoundError) as e:
            logger.error("Chyba při archivaci souboru: %s", e, exc_info=True)
            return False

    def ulozit_pracovni_dobu(
        self, date_str, start_time_str, end_time_str, lunch_duration_str, employees
    ):
        """Zapíše pracovní dobu do listu týdne; vytvoří list z template pokud chybí."""
        try:
            date_obj = datetime.strptime(date_str, "%Y-%m-%d")
            week_number = date_obj.isocalendar().week
            sheet_name = f"{Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME} {week_number}"

            with self._get_workbook() as workbook:
                if not workbook:
                    raise IOError("Workbook se nepodařilo otevřít pro uložení.")

                if sheet_name not in workbook.sheetnames:
                    if Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME in workbook.sheetnames:
                        template_sheet = workbook[Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME]
                        sheet = workbook.copy_worksheet(template_sheet)
                        sheet.title = sheet_name
                    else:
                        sheet = workbook.create_sheet(sheet_name)
                else:
                    sheet = workbook[sheet_name]

                self._zapsat_data_do_listu(
                    sheet, date_obj, start_time_str, end_time_str, lunch_duration_str, employees
                )

            logger.info("Uložena pracovní doba pro %s do listu %s.", date_str, sheet_name)
            return True
        except (FileNotFoundError, ValueError, IOError) as e:
            logger.error("Chyba při ukládání pracovní doby: %s", e, exc_info=True)
            return False
        except Exception as e:
            logger.error("Neočekávaná chyba při ukládání pracovní doby: %s", e, exc_info=True)
            return False

    def _zapsat_data_do_listu(
        self, sheet, date_obj, start_time_str, end_time_str, lunch_duration_str, employees
    ):
        """Pomocná metoda pro zápis dat do konkrétního listu."""
        day_column_index = 2 + 2 * date_obj.weekday()
        start_time = datetime.strptime(start_time_str, "%H:%M")
        end_time = datetime.strptime(end_time_str, "%H:%M")
        total_hours = round(
            (end_time - start_time).total_seconds() / 3600 - float(lunch_duration_str), 2
        )

        employee_rows = {
            sheet.cell(row=r, column=1).value: r
            for r in range(Config.EXCEL_EMPLOYEE_START_ROW, sheet.max_row + 1)
        }
        next_empty_row = (
            max(employee_rows.values() or [Config.EXCEL_EMPLOYEE_START_ROW - 1]) + 1
        )

        for employee in employees:
            row_idx = employee_rows.get(employee)
            if not row_idx:
                row_idx = next_empty_row
                sheet.cell(row=row_idx, column=1, value=employee)
                next_empty_row += 1

            data_cell = sheet.cell(row=row_idx, column=day_column_index + 1)
            if not isinstance(data_cell, MergedCell):
                data_cell.value = total_hours
                data_cell.number_format = "0.00"

        start_time_cell = sheet.cell(row=7, column=day_column_index)
        if not isinstance(start_time_cell, MergedCell):
            start_time_cell.value = start_time.time()
            start_time_cell.number_format = "HH:MM"

        date_cell = sheet.cell(row=80, column=day_column_index)
        if not isinstance(date_cell, MergedCell):
            date_cell.value = date_obj.date()
            date_cell.number_format = "DD.MM.YYYY"

    def close_cached_workbooks(self):
        """Flush + zavření všech workbooků v cache (volat při ukončení requestu)."""
        with self._file_lock:
            for path_str, wb in self._workbook_cache.items():
                try:
                    if wb:
                        wb.save(path_str)
                        wb.close()
                except Exception as e:
                    logger.error(
                        "Chyba při ukládání/zavírání workbooku %s: %s",
                        path_str,
                        e,
                        exc_info=True,
                    )
            self._workbook_cache.clear()

    def ziskej_cislo_tydne(self, datum):
        """Vrátí ISO kalendář (year, week, weekday) nebo None při chybě."""
        try:
            datum_obj = (
                datetime.strptime(datum, "%Y-%m-%d") if isinstance(datum, str) else datum
            )
            return datum_obj.isocalendar()
        except (ValueError, TypeError) as e:
            logger.error("Chyba při zpracování data '%s': %s", datum, e)
            return None

    def generate_monthly_report(self, month, year, employees=None):
        """Agreguje hodiny / volné dny z týdenních listů pro daný měsíc."""
        if not (1 <= month <= 12 and 2000 <= year <= 2100):
            raise ValueError("Neplatný měsíc nebo rok.")

        report_data = {}
        try:
            with self._get_workbook(read_only=True) as workbook:
                if not workbook:
                    raise IOError("Workbook se nepodařilo otevřít pro report.")
                for sheet in self._get_monthly_sheets(workbook, month, year):
                    self._process_sheet_for_report(sheet, employees, report_data)
        except (FileNotFoundError, IOError) as e:
            logger.error("Chyba při generování měsíčního reportu: %s", e, exc_info=True)
            return {}
        return {
            emp: data
            for emp, data in report_data.items()
            if data["total_hours"] > 0 or data["free_days"] > 0
        }

    def _get_monthly_sheets(self, workbook, month, year):
        """Generátor pro listy, které spadají do daného měsíce a roku."""
        for sheet_name in workbook.sheetnames:
            if not sheet_name.startswith(Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME):
                continue
            # Heuristika pro kontrolu datumu v listu, aby se zbytečně neprocházel
            # každý list. Předpokládáme, že alespoň jedna buňka s datem je vyplněná.
            sheet = workbook[sheet_name]
            for c_idx in range(2, 12, 2):
                date_val = sheet.cell(row=80, column=c_idx).value
                if isinstance(date_val, datetime) and date_val.month == month and date_val.year == year:
                    yield sheet
                    break

    def _process_sheet_for_report(self, sheet, employees, report_data):
        """Zpracuje jeden list a agreguje data do report_data."""
        for r_idx in range(Config.EXCEL_EMPLOYEE_START_ROW, sheet.max_row + 1):
            employee_name = sheet.cell(row=r_idx, column=1).value
            if not employee_name or (employees and employee_name not in employees):
                continue

            if employee_name not in report_data:
                report_data[employee_name] = {"total_hours": 0.0, "free_days": 0}

            for c_idx in range(3, 13, 2):  # Sloupce s hodinami
                hours = sheet.cell(row=r_idx, column=c_idx).value
                if not isinstance(hours, (int, float)):
                    continue
                # Ověření, že datum odpovídá, není zde nutné, protože
                # _get_monthly_sheets již zajistil relevantnost listu.
                if hours > 0:
                    report_data[employee_name]["total_hours"] += hours
                else:
                    report_data[employee_name]["free_days"] += 1

    def update_project_info(self, _project_name, _start_date, _end_date):
        """Placeholder pro budoucí implementaci údajů o projektu (aktuálně no-op)."""
        return True
