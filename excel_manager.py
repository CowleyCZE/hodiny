"""Správa týdenního Excel souboru (šablona + dynamické listy týdnů, reporty).

Zodpovědnosti:
 - Lazy/cached otevření workbooku (thread‑safe)
 - Archivace starého týdne a vyčištění listů
 - Zápis denních hodin pro vybrané zaměstnance do listu konkrétního týdne
 - Generace měsíčního reportu agregací z týdenních listů
"""

import json
import logging
import shutil
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from threading import Lock

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import coordinate_to_tuple

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
        self._metadata_path = self.base_path / "metadata.json"
        logger.info("ExcelManager inicializován pro: %s", self.active_file_path)

    def get_active_file_path(self):
        """Cesta k aktivnímu Excel souboru (šablona)."""
        return self.active_file_path

    def _load_dynamic_config(self):
        """Načte dynamickou konfiguraci pro ukládání do XLSX souborů."""
        if not Config.CONFIG_FILE_PATH.exists():
            return {}
        try:
            with open(Config.CONFIG_FILE_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError) as e:
            logger.error("Chyba při načítání dynamické konfigurace: %s", e, exc_info=True)
            return {}

    def _get_cell_coordinates(self, field_key, sheet_name=None, data_type="weekly_time"):
        """Vrátí seznam (row, col) souřadnic pro daný field z dynamické konfigurace.

        Args:
            field_key: Klíč pole z konfigurace (např. 'start_time', 'date')
            sheet_name: Název listu, pokud chceme ověřit shodu
            data_type: Typ dat ('weekly_time', 'advances', 'monthly_time', 'projects')

        Returns:
            list: Seznam (row, col) souřadnic nebo prázdný seznam pokud není nakonfigurováno
        """
        config = self._load_dynamic_config()
        data_config = config.get(data_type, {})
        field_configs = data_config.get(field_key, [])

        if not field_configs:
            return []

        coordinates = []
        for field_config in field_configs:
            # Ověř, že konfigurace je pro správný soubor a list
            if field_config.get("file") != self.active_filename:
                logger.warning(
                    "Konfigurace pro %s/%s odkazuje na jiný soubor: %s", data_type, field_key, field_config.get("file")
                )
                continue

            if sheet_name and field_config.get("sheet") != sheet_name:
                logger.warning(
                    "Konfigurace pro %s/%s odkazuje na jiný list: %s (očekáván %s)",
                    data_type,
                    field_key,
                    field_config.get("sheet"),
                    sheet_name,
                )
                continue

            cell = field_config.get("cell")
            if not cell:
                continue

            try:
                coordinates.append(coordinate_to_tuple(cell))  # Převede např. "A1" na (1, 1)
            except ValueError as e:
                logger.error("Neplatná buňka v konfiguraci pro %s/%s: %s - %s", data_type, field_key, cell, e)
                continue

        return coordinates

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
                    raise IOError(f"Chyba při otevírání souboru '{self.active_filename}': {e}") from e
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

    def ulozit_pracovni_dobu(self, date_str, start_time_str, end_time_str, lunch_duration_str, employees):
        """Zapíše pracovní dobu do týdenního souboru; vytvoří soubor a list z template pokud chybí."""
        try:
            date_obj = datetime.strptime(date_str, "%Y-%m-%d")
            week_number = date_obj.isocalendar().week

            # Získej nebo vytvoř týdenní soubor
            weekly_file_path = self._get_or_create_weekly_file(week_number)

            # Pracuj s týdenním souborem
            weekly_workbook = self._get_weekly_workbook(weekly_file_path)
            if not weekly_workbook:
                raise IOError(f"Týdenní workbook {weekly_file_path} se nepodařilo otevřít.")

            # Pracuj s listem "Týden" v týdenním souboru
            sheet_name = Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME
            if sheet_name not in weekly_workbook.sheetnames:
                # Vytvoř list zkopírováním z hlavní šablony
                self._create_week_sheet_from_template(weekly_workbook, sheet_name)

            sheet = weekly_workbook[sheet_name]
            self._zapsat_data_do_listu(
                sheet, sheet_name, date_obj, start_time_str, end_time_str, lunch_duration_str, employees
            )

            # Ulož týdenní soubor
            weekly_workbook.save(weekly_file_path)
            weekly_workbook.close()

            logger.info("Uložena pracovní doba pro %s do týdenního souboru %s.", date_str, weekly_file_path.name)
            return True
        except (FileNotFoundError, ValueError, IOError) as e:
            logger.error("Chyba při ukládání pracovní doby: %s", e, exc_info=True)
            return False
        except Exception as e:
            logger.error("Neočekávaná chyba při ukládání pracovní doby: %s", e, exc_info=True)
            return False

    def _get_or_create_weekly_file(self, week_number):
        """Získá cestu k týdennímu souboru, vytvoří ho pokud neexistuje."""
        weekly_filename = f"Hodiny_Cap_Tyden{week_number}.xlsx"
        weekly_file_path = self.base_path / weekly_filename

        if not weekly_file_path.exists():
            # Najdi předchozí týdenní soubor jako zdroj pro kopírování
            previous_week_file = self._find_previous_weekly_file(week_number)

            if previous_week_file and previous_week_file.exists():
                # Zkopíruj z předchozího týdenního souboru
                shutil.copy(previous_week_file, weekly_file_path)
                logger.info("Vytvořen týdenní soubor %s zkopírováním z %s", weekly_filename, previous_week_file.name)
            else:
                # Zkopíruj z hlavní šablony jako fallback
                shutil.copy(self.active_file_path, weekly_file_path)
                logger.info(
                    "Vytvořen týdenní soubor %s zkopírováním ze šablony %s", weekly_filename, self.active_filename
                )

        return weekly_file_path

    def _find_previous_weekly_file(self, current_week):
        """Najde nejnovější týdenní soubor před aktuálním týdnem."""
        for week in range(current_week - 1, 0, -1):
            potential_file = self.base_path / f"Hodiny_Cap_Tyden{week}.xlsx"
            if potential_file.exists():
                return potential_file
        return None

    def _get_weekly_workbook(self, weekly_file_path):
        """Otevře týdenní workbook pro práci."""
        try:
            from openpyxl import load_workbook

            return load_workbook(str(weekly_file_path))
        except Exception as e:
            logger.error("Chyba při otevírání týdenního souboru %s: %s", weekly_file_path, e)
            return None

    def _create_week_sheet_from_template(self, workbook, sheet_name):
        """Vytvoří list 'Týden' zkopírováním z hlavní šablony."""
        try:
            # Otevři hlavní šablonu a zkopíruj z ní list "Týden"
            with self._get_workbook(read_only=True) as template_wb:
                if not template_wb:
                    raise IOError("Šablonu se nepodařilo otevřít.")

                if Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME in template_wb.sheetnames:
                    # Zkopíruj list z šablony
                    template_sheet = template_wb[Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME]

                    # Vytvoř nový list v cílovém workbooku
                    new_sheet = workbook.create_sheet(sheet_name)

                    # Zkopíruj obsah buňka po buňce
                    for row in template_sheet.iter_rows():
                        for cell in row:
                            if cell.value is not None:
                                new_cell = new_sheet.cell(row=cell.row, column=cell.column, value=cell.value)
                                # Zkopíruj i formátování pokud je to možné
                                if hasattr(cell, "number_format"):
                                    new_cell.number_format = cell.number_format

                    logger.info("Vytvořen list %s zkopírováním ze šablony", sheet_name)
                else:
                    # Fallback - vytvoř prázdný list
                    workbook.create_sheet(sheet_name)
                    logger.warning(
                        "Šablona neobsahuje list %s, vytvořen prázdný list", Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME
                    )
        except Exception as e:
            logger.error("Chyba při vytváření listu ze šablony: %s", e)
            # Fallback - vytvoř prázdný list
            if sheet_name not in workbook.sheetnames:
                workbook.create_sheet(sheet_name)

    def _zapsat_data_do_listu(
        self, sheet, sheet_name, date_obj, start_time_str, end_time_str, lunch_duration_str, employees
    ):
        """Pomocná metoda pro zápis dat do konkrétního listu."""
        day_column_index = 2 + 2 * date_obj.weekday()
        start_time = datetime.strptime(start_time_str, "%H:%M")
        end_time = datetime.strptime(end_time_str, "%H:%M")
        total_hours = round((end_time - start_time).total_seconds() / 3600 - float(lunch_duration_str), 2)

        employee_rows = {
            sheet.cell(row=r, column=1).value: r for r in range(Config.EXCEL_EMPLOYEE_START_ROW, sheet.max_row + 1)
        }
        next_empty_row = max(employee_rows.values() or [Config.EXCEL_EMPLOYEE_START_ROW - 1]) + 1

        # Zapíše zaměstnance a jejich hodiny
        for employee in employees:
            row_idx = employee_rows.get(employee)
            if not row_idx:
                row_idx = next_empty_row

                # Zkus použít dynamickou konfiguraci pro jméno zaměstnance
                employee_name_coords = self._get_cell_coordinates("employee_name", sheet_name, "weekly_time")
                if employee_name_coords:
                    emp_row, emp_col = employee_name_coords[0]  # Použij první lokaci
                    # Pokud je nakonfigurovaná buňka pro jméno, přidej zaměstnance na nový řádek od této pozice
                    row_idx = max(next_empty_row, emp_row)

                sheet.cell(row=row_idx, column=1, value=employee)
                next_empty_row = max(next_empty_row, row_idx) + 1

            data_cell = sheet.cell(row=row_idx, column=day_column_index + 1)
            if not isinstance(data_cell, MergedCell):
                data_cell.value = total_hours
                data_cell.number_format = "0.00"

        # Zápis času začátku - použij dynamickou konfiguraci nebo fallback
        start_time_coords = self._get_cell_coordinates("start_time", sheet_name, "weekly_time")
        if start_time_coords:
            # Zapíše do všech nakonfigurovaných lokací
            for start_row, start_col in start_time_coords:
                start_time_cell = sheet.cell(row=start_row, column=start_col)
                if not isinstance(start_time_cell, MergedCell):
                    start_time_cell.value = start_time.time()
                    start_time_cell.number_format = "HH:MM"
                logger.info(
                    "Používám dynamickou konfiguraci pro čas začátku: buňka %s",
                    f"{chr(64 + start_col)}{start_row}",
                )
        else:
            # Fallback na původní logiku
            start_row, start_col = 7, day_column_index
            start_time_cell = sheet.cell(row=start_row, column=start_col)
            if not isinstance(start_time_cell, MergedCell):
                start_time_cell.value = start_time.time()
                start_time_cell.number_format = "HH:MM"
            logger.info("Používám původní logiku pro čas začátku: řádek 7, sloupec %d", start_col)

        # Zápis času konce - použij dynamickou konfiguraci pokud je nastavena
        end_time_coords = self._get_cell_coordinates("end_time", sheet_name, "weekly_time")
        if end_time_coords:
            for end_row, end_col in end_time_coords:
                end_time_cell = sheet.cell(row=end_row, column=end_col)
                if not isinstance(end_time_cell, MergedCell):
                    end_time_cell.value = end_time.time()
                    end_time_cell.number_format = "HH:MM"
                logger.info("Používám dynamickou konfiguraci pro čas konce: buňka %s", f"{chr(64 + end_col)}{end_row}")

        # Zápis doby oběda - použij dynamickou konfiguraci pokud je nastavena
        lunch_coords = self._get_cell_coordinates("lunch_duration", sheet_name, "weekly_time")
        if lunch_coords:
            for lunch_row, lunch_col in lunch_coords:
                lunch_cell = sheet.cell(row=lunch_row, column=lunch_col)
                if not isinstance(lunch_cell, MergedCell):
                    lunch_cell.value = float(lunch_duration_str)
                    lunch_cell.number_format = "0.00"
                logger.info(
                    "Používám dynamickou konfiguraci pro dobu oběda: buňka %s", f"{chr(64 + lunch_col)}{lunch_row}"
                )

        # Zápis celkových hodin - použij dynamickou konfiguraci pokud je nastavena
        total_hours_coords = self._get_cell_coordinates("total_hours", sheet_name, "weekly_time")
        if total_hours_coords:
            for total_row, total_col in total_hours_coords:
                total_cell = sheet.cell(row=total_row, column=total_col)
                if not isinstance(total_cell, MergedCell):
                    total_cell.value = total_hours
                    total_cell.number_format = "0.00"
                logger.info(
                    "Používám dynamickou konfiguraci pro celkové hodiny: buňka %s",
                    f"{chr(64 + total_col)}{total_row}",
                )

        # Zápis data - použij dynamickou konfiguraci nebo fallback
        date_coords = self._get_cell_coordinates("date", sheet_name, "weekly_time")
        if date_coords:
            for date_row, date_col in date_coords:
                date_cell = sheet.cell(row=date_row, column=date_col)
                if not isinstance(date_cell, MergedCell):
                    date_cell.value = date_obj.date()
                    date_cell.number_format = "DD.MM.YYYY"
                logger.info("Používám dynamickou konfiguraci pro datum: buňka %s", f"{chr(64 + date_col)}{date_row}")
        else:
            # Fallback na původní logiku
            date_row, date_col = 80, day_column_index
            date_cell = sheet.cell(row=date_row, column=date_col)
            if not isinstance(date_cell, MergedCell):
                date_cell.value = date_obj.date()
                date_cell.number_format = "DD.MM.YYYY"
            logger.info("Používám původní logiku pro datum: řádek 80, sloupec %d", date_col)

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
            datum_obj = datetime.strptime(datum, "%Y-%m-%d") if isinstance(datum, str) else datum
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
                    self._process_sheet_for_report(sheet, employees, report_data, month, year)
        except (FileNotFoundError, IOError) as e:
            logger.error("Chyba při generování měsíčního reportu: %s", e, exc_info=True)
            return {}
        return {emp: data for emp, data in report_data.items() if data["total_hours"] > 0 or data["free_days"] > 0}

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

    def _process_sheet_for_report(self, sheet, employees, report_data, month, year):
        """Zpracuje jeden list a agreguje data do report_data."""
        for r_idx in range(Config.EXCEL_EMPLOYEE_START_ROW, sheet.max_row + 1):
            employee_name = sheet.cell(row=r_idx, column=1).value
            if not employee_name or (employees and employee_name not in employees):
                continue

            if employee_name not in report_data:
                report_data[employee_name] = {"total_hours": 0.0, "free_days": 0}

            for c_idx in range(3, 13, 2):  # Sloupce s hodinami
                # Zkontroluj odpovídající datum pro tento sloupec
                date_column = c_idx - 1  # Datum je v předchozím sloupci
                date_val = sheet.cell(row=80, column=date_column).value

                # Pouze pokud datum odpovídá cílovému měsíci a roku
                if not (isinstance(date_val, datetime) and date_val.month == month and date_val.year == year):
                    continue

                hours = sheet.cell(row=r_idx, column=c_idx).value
                if not isinstance(hours, (int, float)):
                    continue

                if hours > 0:
                    report_data[employee_name]["total_hours"] += hours
                else:
                    report_data[employee_name]["free_days"] += 1

    def update_project_info(self, _project_name, _start_date, _end_date):
        """Placeholder pro budoucí implementaci údajů o projektu (aktuálně no-op)."""
        return True

    def get_current_week_data(self, week_number=None):
        """Získá data z týdenního souboru pro zobrazení na hlavní stránce.

        Args:
            week_number (int, optional): Číslo týdne. Pokud None, použije se aktuální týden.
        """
        try:
            if week_number is None:
                current_week = datetime.now().isocalendar().week
            else:
                current_week = week_number

            # Zkus najít týdenní soubor
            weekly_filename = f"Hodiny_Cap_Tyden{current_week}.xlsx"
            weekly_file_path = self.base_path / weekly_filename

            # Pokud týdenní soubor neexistuje, zkus hlavní šablonu
            if weekly_file_path.exists():
                wb = self._get_weekly_workbook(weekly_file_path)
                sheet_name = Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME  # "Týden"
            else:
                # Fallback na hlavní šablonu
                with self._get_workbook(read_only=True) as wb:
                    if not wb:
                        return None
                    sheet_name = Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME

            if not wb:
                return None

            # Pokusíme se najít list "Týden"
            if sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
            else:
                return None

            # Získáme data z listu (prvních několik řádků a sloupců pro přehled)
            data = []
            max_row = min(sheet.max_row, 20)  # Omezíme na prvních 20 řádků
            max_col = min(sheet.max_column, 10)  # Omezíme na prvních 10 sloupců

            for row in range(1, max_row + 1):
                row_data = []
                for col in range(1, max_col + 1):
                    cell = sheet.cell(row=row, column=col)
                    value = cell.value
                    if value is None:
                        value = ""
                    elif isinstance(value, (int, float)):
                        value = str(value)
                    row_data.append(str(value))
                data.append(row_data)

            # Zavři workbook pokud je to týdenní soubor
            if weekly_file_path.exists():
                wb.close()

            return {
                "sheet_name": f"{sheet_name} {current_week}",
                "data": data,
                "rows": len(data),
                "cols": len(data[0]) if data else 0,
            }

        except Exception as e:
            logger.error("Chyba při načítání dat aktuálního týdne: %s", e, exc_info=True)
            return None

    def _load_metadata(self):
        """Načte metadata souborů z JSON souboru."""
        if not self._metadata_path.exists():
            return {}
        try:
            with self._file_lock:
                with open(self._metadata_path, "r", encoding="utf-8") as f:
                    return json.load(f)
        except (IOError, json.JSONDecodeError) as e:
            logger.error("Chyba při načítání metadat: %s", e)
            return {}

    def _save_metadata(self, metadata):
        """Uloží metadata souborů do JSON souboru."""
        try:
            with self._file_lock:
                with open(self._metadata_path, "w", encoding="utf-8") as f:
                    json.dump(metadata, f, indent=4, ensure_ascii=False)
        except IOError as e:
            logger.error("Chyba při ukládání metadat: %s", e)

    def get_all_metadata(self):
        """Vrátí všechna metadata."""
        return self._load_metadata()

    def set_category(self, filename, category):
        """Nastaví kategorii pro daný soubor."""
        metadata = self._load_metadata()
        if filename not in metadata:
            metadata[filename] = {}
        metadata[filename]["category"] = category
        self._save_metadata(metadata)
