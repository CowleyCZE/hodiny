# excel_manager.py
from config import Config
import contextlib
import logging
import os
import platform
import re
import shutil
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from threading import Lock

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.copier import WorksheetCopy
from openpyxl.utils import get_column_letter # Import pro převod indexu na písmeno sloupce

# Předpokládá existenci utils.logger
try:
    from utils.logger import setup_logger # Opraveno na setup_logger
    logger = setup_logger("excel_manager")
except ImportError:
    # Fallback na základní logger, pokud utils.logger není dostupný
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger("excel_manager")
from datetime import datetime # Přidán import datetime


# Kontrola jestli je systém Windows
IS_WINDOWS = platform.system() == "Windows"


class ExcelManager:
    """
    Spravuje operace s aktivním Excel souborem, který je kopií šablony.
    """
    def __init__(self, base_path, active_filename, template_filename):
        """
        Inicializuje ExcelManager.

        Args:
            base_path (Path): Cesta k adresáři s Excel soubory.
            active_filename (str): Název aktuálně používaného souboru.
            template_filename (str): Název souboru šablony.
        """
        self.base_path = Path(base_path)
        self.active_filename = active_filename
        self.template_filename = template_filename
        # Sestavení plné cesty k aktivnímu souboru
        self.active_file_path = self.base_path / self.active_filename if self.active_filename else None
        self.template_file_path = self.base_path / self.template_filename

        self.current_project_name = None # Pro ukládání názvu projektu do listů
        self._file_lock = Lock()
        # Cache nyní ukládá workbooky podle jejich plné cesty
        self._workbook_cache = {}
        logger.info(f"ExcelManager inicializován pro aktivní soubor: {self.active_file_path}")

    def get_active_file_path(self):
        """Vrátí cestu k aktuálnímu aktivnímu souboru."""
        if not self.active_file_path:
             logger.error("Aktivní Excel soubor není nastaven.")
             # Můžeme zde vyvolat výjimku nebo vrátit None
             raise ValueError("Aktivní Excel soubor není definován.")
        return self.active_file_path

    @contextmanager
    def _get_workbook(self, file_path_to_open, read_only=False):
        """
        Context manager pro bezpečné otevření, cachování a zavření workbooku.
        Pracuje s konkrétní cestou k souboru.
        """
        file_path = Path(file_path_to_open)
        cache_key = str(file_path.absolute())
        wb = None
        is_from_cache = False

        with self._file_lock:
            try:
                if cache_key in self._workbook_cache:
                    try:
                        wb = self._workbook_cache[cache_key]
                        _ = wb.sheetnames
                        is_from_cache = True
                        logger.debug(f"Workbook načten z cache: {cache_key}")
                    except Exception as cache_err:
                        logger.warning(f"Chyba při použití workbooku z cache ({cache_key}): {cache_err}. Workbook bude znovu načten.")
                        if cache_key in self._workbook_cache:
                            with contextlib.suppress(Exception):
                                self._workbook_cache[cache_key].close()
                            del self._workbook_cache[cache_key]
                        wb = None
                        is_from_cache = False

                if wb is None:
                    file_path.parent.mkdir(parents=True, exist_ok=True)
                    if not file_path.exists():
                        raise FileNotFoundError(f"Požadovaný Excel soubor '{file_path.name}' nebyl nalezen na cestě '{file_path}'.")
                    
                    try:
                        wb = load_workbook(filename=str(file_path), read_only=read_only, data_only=read_only)
                        logger.debug(f"Workbook načten ze souboru: {file_path} (read_only={read_only}, data_only={read_only})")
                        
                        if not read_only:
                            self._workbook_cache[cache_key] = wb
                            logger.debug(f"Workbook přidán do cache: {cache_key}")
                    except Exception as load_err:
                        logger.error(f"Nelze načíst Excel soubor {file_path}: {load_err}", exc_info=True)
                        raise IOError(f"Chyba při otevírání souboru '{file_path.name}'.")

                yield wb

                # OPRAVA: Okamžité ukládání je odstraněno. O uložení se postará _clear_workbook_cache na konci requestu.

            except Exception as e:
                logger.error(f"Obecná chyba v _get_workbook pro {file_path}: {e}", exc_info=True)
                if cache_key in self._workbook_cache:
                    del self._workbook_cache[cache_key]
                    logger.info(f"Workbook odstraněn z cache kvůli obecné chybě: {cache_key}")
                raise

            finally:
                if read_only and wb is not None:
                    try:
                        wb.close()
                        logger.debug(f"Read-only workbook uzavřen v finally: {file_path}")
                        if cache_key in self._workbook_cache and self._workbook_cache[cache_key] is wb:
                            del self._workbook_cache[cache_key]
                    except Exception as close_err:
                        logger.warning(f"Chyba při zavírání read-only workbooku {file_path} v finally: {close_err}")
                elif not read_only and wb is not None and cache_key in self._workbook_cache:
                    logger.debug(f"Workbook pro zápis '{file_path}' je v cache a nebude zde uzavřen ani uložen. Správu přebírá _clear_workbook_cache.")

    def ulozit_pracovni_dobu(self, date, start_time, end_time, lunch_duration, employees, week_number=None):
        """Uloží pracovní dobu do aktivního Excel souboru"""
        active_path = self.get_active_file_path()
        try:
            # Zde potřebujeme zapisovat, takže read_only=False
            with self._get_workbook(active_path, read_only=False) as workbook:
                logger.debug(f"Ukládám pracovní dobu: Datum={date}, Start={start_time}, End={end_time}, Lunch={lunch_duration}, Zaměstnanci={employees}, Explicitní týden={week_number}")
                
                if week_number is None:
                    week_calendar_data = self.ziskej_cislo_tydne(date)
                    if not week_calendar_data:
                        raise ValueError("Nepodařilo se získat číslo týdne pro zadané datum.")
                    week_number = week_calendar_data.week

                sheet_name = f"{Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME} {week_number}"

                if sheet_name not in workbook.sheetnames:
                    if Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME in workbook.sheetnames:
                        source_sheet = workbook[Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME]
                        try:
                            # OPRAVA: Použití standardní a spolehlivé metody copy_worksheet
                            target_sheet = workbook.copy_worksheet(source_sheet)
                            target_sheet.title = sheet_name
                            sheet = target_sheet
                            logger.info(f"Zkopírován list '{Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME}' na '{sheet_name}' v souboru {self.active_filename}.")
                        except Exception as copy_err:
                            logger.error(f"Nepodařilo se zkopírovat list '{Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME}' v {self.active_filename}: {copy_err}", exc_info=True)
                            sheet = workbook.create_sheet(sheet_name)
                            logger.warning(f"Vytvořen nový prázdný list '{sheet_name}' v souboru {self.active_filename} kvůli chybě při kopírování.")
                    else:
                        sheet = workbook.create_sheet(sheet_name)
                        logger.warning(f"Vytvořen nový prázdný list '{sheet_name}' v souboru {self.active_filename} (šablona '{Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME}' nenalezena).")

                    sheet["A80"] = sheet_name
                else:
                    sheet = workbook[sheet_name]

                weekday = datetime.strptime(date, "%Y-%m-%d").weekday()
                day_column_index = 1 + 2 * weekday
                start_time_col_index = day_column_index
                end_time_col_index = day_column_index + 1
                date_col_index = day_column_index

                if start_time == "00:00" and end_time == "00:00" and lunch_duration == 0.0:
                    total_hours = 0
                else:
                    start = datetime.strptime(start_time, "%H:%M")
                    end = datetime.strptime(end_time, "%H:%M")
                    lunch_duration_float = float(lunch_duration)
                    total_hours = (end - start).total_seconds() / 3600 - lunch_duration_float
                    total_hours = round(total_hours, 2)

                start_row = Config.EXCEL_EMPLOYEE_START_ROW
                for employee in employees:
                    current_row = start_row
                    row_found = False
                    max_search_row = start_row + 1000

                    while current_row < max_search_row:
                        employee_cell = sheet.cell(row=current_row, column=1)
                        if employee_cell.value == employee:
                            sheet.cell(row=current_row, column=end_time_col_index + 2, value=total_hours)
                            row_found = True
                            break
                        elif employee_cell.value is None or str(employee_cell.value).strip() == "":
                            employee_cell.value = employee
                            sheet.cell(row=current_row, column=end_time_col_index + 2, value=total_hours)
                            row_found = True
                            break
                        current_row += 1

                    if not row_found:
                        logger.warning(f"Nepodařilo se najít ani vytvořit řádek pro zaměstnance '{employee}' v listu '{sheet_name}' souboru {self.active_filename}.")

                start_time_cell = sheet[f"{get_column_letter(start_time_col_index + 1)}7"]
                end_time_cell = sheet[f"{get_column_letter(end_time_col_index + 1)}7"]
                
                try:
                    start_time_obj = datetime.strptime(start_time, "%H:%M").time()
                    end_time_obj = datetime.strptime(end_time, "%H:%M").time()
                except ValueError:
                    logger.error(f"Neplatný formát času: Start='{start_time}', End='{end_time}'.")
                    return False

                start_time_cell.value = start_time_obj
                end_time_cell.value = end_time_obj
                start_time_cell.number_format = 'HH:MM'
                end_time_cell.number_format = 'HH:MM'
                logger.debug(f"Zapsáno do {start_time_cell.coordinate}: {start_time_cell.value}")
                logger.debug(f"Zapsáno do {end_time_cell.coordinate}: {end_time_cell.value}")

                date_col_letter = get_column_letter(date_col_index + 1)
                try:
                    date_obj = datetime.strptime(date, "%Y-%m-%d").date()
                    date_cell = sheet[f"{date_col_letter}80"]
                    date_cell.value = date_obj
                    date_cell.number_format = 'DD.MM.YYYY'
                    logger.debug(f"Zapsáno do {date_cell.coordinate}: {date_cell.value}")
                except ValueError:
                    logger.error(f"Neplatný formát data '{date}' při ukládání do buňky {date_col_letter}80.")
                    return False

                if self.current_project_name:
                    project_cell = sheet["B79"]
                    existing_projects = project_cell.value or ""
                    project_list = [p.strip() for p in existing_projects.split(',') if p.strip()]
                    if self.current_project_name not in project_list:
                        project_list.append(self.current_project_name)
                    project_cell.value = ", ".join(project_list)

                logger.info(f"Úspěšně uložena pracovní doba pro datum {date} do listu {sheet_name} v souboru {self.active_filename}")
                return True

        except (FileNotFoundError, ValueError, IOError) as e:
            logger.error(f"Chyba při ukládání pracovní doby do {self.active_filename}: {e}", exc_info=True)
            return False
        except Exception as e:
            logger.error(f"Neočekávaná chyba při ukládání pracovní doby do {self.active_filename}: {e}", exc_info=True)
            return False
    def _clear_workbook_cache(self):
        """Vyčistí cache workbooků a pokusí se uložit neuložené změny."""
        with self._file_lock: # Zajistíme atomicitu operace s cache
             logger.info(f"Čištění cache workbooků ({len(self._workbook_cache)} položek)...")
             for path_str, wb in list(self._workbook_cache.items()):
                 try:
                     # Workbooky v cache jsou vždy pro zápis (read-only se neukládají)
                     wb.save(path_str)
                     logger.info(f"Workbook uložen při čištění cache: {path_str}")
                     wb.close()
                 except Exception as e:
                     logger.error(f"Chyba při ukládání/zavírání workbooku {path_str} z cache: {e}", exc_info=True)
                     # I přes chybu odstraníme z cache
                 finally:
                     # Odstraníme z cache bez ohledu na úspěch uložení/zavření
                     self._workbook_cache.pop(path_str, None)
             logger.info("Cache workbooků vyčištěna.")

    def __del__(self):
        """Destruktor - zajistí uvolnění prostředků a uložení změn."""
        self._clear_workbook_cache()

    def close_cached_workbooks(self):
         """Metoda pro explicitní vyčištění cache (např. na konci requestu)."""
         self._clear_workbook_cache()

    def set_project_name(self, project_name):
         """Nastaví aktuální název projektu pro použití v jiných metodách."""
         self.current_project_name = project_name if project_name else None

    def ziskej_cislo_tydne(self, datum):
        """
        Získá ISO kalendářní data (rok, číslo týdne, den v týdnu) pro zadané datum.
        """
        try:
            logger.debug(f"Ziskej_cislo_tydne: Vstupní datum '{datum}', typ: {type(datum)}")
            if isinstance(datum, str):
                try:
                    datum_obj = datetime.strptime(datum, "%Y-%m-%d %H:%M:%S")
                except ValueError:
                    datum_obj = datetime.strptime(datum, "%Y-%m-%d")
            elif isinstance(datum, datetime):
                 datum_obj = datum
            else:
                 raise TypeError("Datum musí být string ve formátu YYYY-MM-DD nebo datetime objekt")

            iso_cal = datum_obj.isocalendar()
            logger.debug(f"Ziskej_cislo_tydne: Datum objekt: {datum_obj}, isocalendar: {iso_cal}")
            return iso_cal
        except (ValueError, TypeError) as e:
            logger.error(f"Chyba při zpracování data '{datum}' pro získání čísla týdne: {e}")
            # Vrátíme None nebo vyvoláme výjimku, aby volající věděl o chybě
            return None

    def get_advance_options(self):
        """Získá možnosti záloh z aktivního Excel souboru"""
        # Pokud aktivní soubor není nastaven, vrátíme výchozí
        if not self.active_file_path:
             logger.warning("Nelze načíst možnosti záloh, aktivní soubor není nastaven. Používají se výchozí.")
             return [Config.DEFAULT_ADVANCE_OPTION_1, Config.DEFAULT_ADVANCE_OPTION_2, Config.DEFAULT_ADVANCE_OPTION_3, Config.DEFAULT_ADVANCE_OPTION_4]

        try:
            # Použijeme read-only mód
            with self._get_workbook(self.active_file_path, read_only=True) as workbook:
                options = []
                default_options = [Config.DEFAULT_ADVANCE_OPTION_1, Config.DEFAULT_ADVANCE_OPTION_2, Config.DEFAULT_ADVANCE_OPTION_3, Config.DEFAULT_ADVANCE_OPTION_4]

                if Config.EXCEL_ADVANCES_SHEET_NAME in workbook.sheetnames:
                    zalohy_sheet = workbook[Config.EXCEL_ADVANCES_SHEET_NAME]
                    option1 = zalohy_sheet["B6"].value
                    option2 = zalohy_sheet["D6"].value
                    option3 = zalohy_sheet["F6"].value
                    option4 = zalohy_sheet["H6"].value
                    options = [
                        str(option1).strip() if option1 else default_options[0],
                        str(option2).strip() if option2 else default_options[1],
                        str(option3).strip() if option3 else default_options[2],
                        str(option4).strip() if option4 else default_options[3]
                    ]
                    # Filter out empty strings if any option is not set in Excel
                    options = [opt for opt in options if opt]
                    logger.info(f"Načteny možnosti záloh z {self.active_filename}: {options}")
                else:
                    logger.warning(f"List '{Config.EXCEL_ADVANCES_SHEET_NAME}' nebyl nalezen v souboru {self.active_filename}, použity výchozí možnosti.")
                    options = default_options

                return options
        except FileNotFoundError:
             logger.error(f"Aktivní soubor {self.active_filename} nebyl nalezen při načítání možností záloh.")
             return [Config.DEFAULT_ADVANCE_OPTION_1, Config.DEFAULT_ADVANCE_OPTION_2, Config.DEFAULT_ADVANCE_OPTION_3, Config.DEFAULT_ADVANCE_OPTION_4]
        except Exception as e:
            logger.error(f"Chyba při načítání možností záloh z {self.active_filename}: {str(e)}", exc_info=True)
            return [Config.DEFAULT_ADVANCE_OPTION_1, Config.DEFAULT_ADVANCE_OPTION_2, Config.DEFAULT_ADVANCE_OPTION_3, Config.DEFAULT_ADVANCE_OPTION_4]



    def generate_monthly_report(self, month, year, employees=None):
        """
        Generuje měsíční report pracovní doby a volných dnů pro zadaný měsíc a rok.
        
        Args:
            month (int): Měsíc (1-12).
            year (int): Rok.
            employees (list, optional): Seznam jmen zaměstnanců, pro které se má report generovat.
                                        Pokud None, generuje se pro všechny nalezené.
        
        Returns:
            dict: Slovník, kde klíčem je jméno zaměstnance a hodnotou je slovník
                  s 'total_hours' a 'free_days'.
        """
        if not (1 <= month <= 12):
            raise ValueError("Měsíc musí být v rozsahu 1-12.")
        if not (2000 <= year <= 2100): # Rozumný rozsah pro roky
            raise ValueError("Rok musí být v rozsahu 2000-2100.")

        report_data = {}
        active_path = self.get_active_file_path()
        workbook = None

        try:
            with self._get_workbook(active_path, read_only=True) as workbook:
                for sheet_name in workbook.sheetnames:
                    if sheet_name.startswith(Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME):
                        sheet = workbook[sheet_name]
                        
                        # Získání dat pro dny v týdnu z řádku 80
                        week_dates = []
                        for col_idx in range(2, 12, 2): # Sloupce B, D, F, H, J (indexy 2, 4, 6, 8, 10)
                            date_cell_coord = f"{get_column_letter(col_idx)}80"
                            date_cell_value = sheet[date_cell_coord].value
                            
                            current_date = None
                            if isinstance(date_cell_value, datetime):
                                current_date = date_cell_value.date()
                            elif isinstance(date_cell_value, str):
                                try:
                                    current_date = datetime.strptime(date_cell_value, "%d.%m.%Y").date()
                                except ValueError:
                                    try:
                                        current_date = datetime.strptime(date_cell_value, "%Y-%m-%d").date()
                                    except ValueError:
                                        logger.warning(f"Neplatný formát data v buňce {date_cell_coord} listu {sheet_name}: {date_cell_value}. Tento den bude ignorován.")
                            
                            week_dates.append(current_date)

                        # Procházení zaměstnanců
                        for row_idx in range(Config.EXCEL_EMPLOYEE_START_ROW, sheet.max_row + 1):
                            employee_name = sheet.cell(row=row_idx, column=1).value
                            if employee_name is None or str(employee_name).strip() == "":
                                break # Předpokládáme, že prázdný řádek znamená konec seznamu zaměstnanců
                            
                            employee_name = str(employee_name).strip()

                            if employees and employee_name not in employees:
                                continue # Přeskočíme, pokud není ve vybraných zaměstnancích

                            if employee_name not in report_data:
                                report_data[employee_name] = {"total_hours": 0.0, "free_days": 0}

                            # Procházení hodin pro daného zaměstnance
                            for i, date_obj in enumerate(week_dates):
                                if date_obj and date_obj.month == month and date_obj.year == year:
                                    # Sloupce pro hodiny jsou C, E, G, I, K (indexy 3, 5, 7, 9, 11)
                                    hours_col_idx = (i * 2) + 3 
                                    hours_cell_value = sheet.cell(row=row_idx, column=hours_col_idx).value
                                    
                                    if hours_cell_value is not None:
                                        try:
                                            hours = float(hours_cell_value)
                                            if hours > 0:
                                                report_data[employee_name]["total_hours"] += hours
                                            elif hours == 0:
                                                report_data[employee_name]["free_days"] += 1
                                        except ValueError:
                                            logger.warning(f"Neplatná hodnota hodin pro {employee_name} v buňce {get_column_letter(hours_col_idx)}{row_idx} listu {sheet_name}: {hours_cell_value}. Bude ignorováno.")
                                    # else: pokud je buňka prázdná, ignorujeme ji (není to ani práce, ani volno)
        except FileNotFoundError:
            logger.error(f"Aktivní soubor {self.active_filename} nebyl nalezen při generování reportu.")
            return {}
        except Exception as e:
            logger.error(f"Chyba při generování měsíčního reportu z {self.active_filename}: {str(e)}", exc_info=True)
            return {}
        
        # Odstraníme zaměstnance, kteří nemají žádné hodiny ani volné dny v daném měsíci
        final_report = {
            employee: data for employee, data in report_data.items()
            if data["total_hours"] > 0 or data["free_days"] > 0
        }
        
        return final_report
