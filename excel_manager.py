# excel_manager.py
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
    from utils.logger import setup_logger
    logger = setup_logger("excel_manager")
except ImportError:
    # Fallback na základní logger, pokud utils.logger není dostupný
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger("excel_manager")


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
        Context manager pro bezpečné otevření a zavření workbooku.
        Pracuje s konkrétní cestou k souboru.
        """
        # Použijeme předanou cestu
        file_path = Path(file_path_to_open)
        cache_key = str(file_path.absolute())
        wb = None
        is_from_cache = False

        # Zámek pro zajištění thread-safe operací se souborem a cache
        with self._file_lock:
            try:
                # Zkusíme získat z cache
                if cache_key in self._workbook_cache:
                    try:
                        wb = self._workbook_cache[cache_key]
                        # Jednoduchý test, zda je workbook stále použitelný
                        _ = wb.sheetnames
                        is_from_cache = True
                        logger.debug(f"Workbook načten z cache: {cache_key}")
                    except Exception as cache_err:
                        logger.warning(f"Chyba při použití workbooku z cache ({cache_key}): {cache_err}. Workbook bude znovu načten.")
                        # Odstraníme neplatný workbook z cache
                        try:
                             self._workbook_cache[cache_key].close()
                        except: pass
                        del self._workbook_cache[cache_key]
                        wb = None
                        is_from_cache = False

                # Pokud není v cache nebo byl neplatný, načteme ze souboru
                if wb is None:
                    # Ujistíme se, že adresář existuje (i když by měl z init_app)
                    file_path.parent.mkdir(parents=True, exist_ok=True)

                    if not file_path.exists():
                         # Toto by nemělo nastat, pokud app.py správně vytváří soubor
                         logger.error(f"Soubor {file_path} neexistuje!")
                         raise FileNotFoundError(f"Požadovaný soubor {file_path.name} nebyl nalezen.")

                    try:
                        # Načteme workbook
                        wb = load_workbook(filename=str(file_path), read_only=read_only, data_only=True)
                        logger.debug(f"Workbook načten ze souboru: {file_path} (read_only={read_only})")
                        # Pokud není read-only, uložíme ho do cache pro případné další úpravy v rámci stejného requestu
                        if not read_only:
                            self._workbook_cache[cache_key] = wb
                    except Exception as load_err:
                        logger.error(f"Nelze načíst Excel soubor {file_path}: {load_err}", exc_info=True)
                        raise IOError(f"Chyba při otevírání souboru {file_path.name}.")


                # Předáme workbook kódu uvnitř 'with' bloku
                yield wb

                # Po skončení 'with' bloku:
                # Pokud workbook nebyl read-only a *nebyl* z cache (tj. byl nově načten), uložíme ho.
                # Pokud byl z cache, neukládáme ho zde, uloží se až na konci requestu nebo při vyčištění cache.
                if not read_only and wb is not None and not is_from_cache:
                    try:
                        wb.save(str(file_path))
                        logger.debug(f"Workbook uložen (po 'with' bloku): {file_path}")
                    except Exception as save_err:
                        logger.error(f"Chyba při ukládání workbooku {file_path} po 'with' bloku: {save_err}", exc_info=True)
                        # Odstraníme z cache, pokud tam je, aby se neuložil poškozený
                        if cache_key in self._workbook_cache:
                             try: self._workbook_cache[cache_key].close()
                             except: pass
                             del self._workbook_cache[cache_key]
                        raise IOError(f"Chyba při ukládání změn do souboru {file_path.name}.")

            except Exception as e:
                logger.error(f"Obecná chyba v _get_workbook pro {file_path}: {e}", exc_info=True)
                # Pokud došlo k chybě, odstraníme workbook z cache, pokud tam je
                if cache_key in self._workbook_cache:
                    try: self._workbook_cache[cache_key].close()
                    except: pass
                    del self._workbook_cache[cache_key]
                # Znovu vyvoláme výjimku, aby byla zpracována výše
                raise
            finally:
                # Pokud byl workbook read-only a je v cache (což by nemělo nastat dle logiky výše,
                # ale pro jistotu), nebo pokud byl read-only a nebyl z cache, zavřeme ho.
                if read_only and wb is not None:
                     try:
                          # Read-only workbooky neukládáme do cache, takže je můžeme rovnou zavřít
                          wb.close()
                          logger.debug(f"Read-only workbook uzavřen: {file_path}")
                     except Exception as close_err:
                          logger.warning(f"Chyba při zavírání read-only workbooku {file_path}: {close_err}")
                # Pokud byl workbook pro zápis a je v cache, NEZAVÍRÁME ho zde,
                # zavře se až při čištění cache nebo na konci aplikace.


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


    def ulozit_pracovni_dobu(self, date, start_time, end_time, lunch_duration, employees):
        """Uloží pracovní dobu do aktivního Excel souboru"""
        active_path = self.get_active_file_path() # Získáme cestu k aktivnímu souboru
        try:
            # Použijeme context manager pro aktivní soubor
            with self._get_workbook(active_path) as workbook:
                # Získání čísla týdne z datumu
                week_calendar_data = self.ziskej_cislo_tydne(date)
                if not week_calendar_data: # Pokud ziskej_cislo_tydne selhalo
                     raise ValueError("Nepodařilo se získat číslo týdne pro zadané datum.")
                week_number = week_calendar_data.week

                sheet_name = f"Týden {week_number}"

                # Vytvoření nebo získání listu pro daný týden
                if sheet_name not in workbook.sheetnames:
                    # Zkusíme zkopírovat ze šablonového listu "Týden" v aktivním souboru
                    if "Týden" in workbook.sheetnames:
                        source_sheet = workbook["Týden"]
                        try:
                            sheet = workbook.copy_worksheet(source_sheet)
                            sheet.title = sheet_name
                            logger.info(f"Zkopírován list 'Týden' a přejmenován na '{sheet_name}' v souboru {self.active_filename}")
                        except Exception as copy_err:
                             logger.error(f"Nepodařilo se zkopírovat list 'Týden' v {self.active_filename}: {copy_err}", exc_info=True)
                             # Pokud kopírování selže, vytvoříme prázdný list
                             sheet = workbook.create_sheet(sheet_name)
                             logger.warning(f"Vytvořen nový prázdný list '{sheet_name}' v souboru {self.active_filename} kvůli chybě při kopírování.")
                    else:
                        # Pokud ani šablonový list "Týden" neexistuje, vytvoříme prázdný
                        sheet = workbook.create_sheet(sheet_name)
                        logger.warning(f"Vytvořen nový prázdný list '{sheet_name}' v souboru {self.active_filename} (šablona 'Týden' nenalezena).")

                    # Nastavení názvu týdne do buňky A80
                    sheet["A80"] = sheet_name
                else:
                    sheet = workbook[sheet_name]

                # Určení sloupce podle dne v týdnu (0 = Po, ..., 6 = Ne)
                weekday = datetime.strptime(date, "%Y-%m-%d").weekday()

                # Pondělí (0) -> sloupec B (index 1), Úterý (1) -> D (index 3), atd.
                # Použijeme +1 pro index sloupce (openpyxl je 1-based)
                day_column_index = 1 + 2 * weekday # Index pro hodiny (B, D, F, H, J)
                start_time_col_index = day_column_index
                end_time_col_index = day_column_index + 1
                date_col_index = day_column_index # Datum do stejného sloupce jako hodiny

                # Výpočet odpracovaných hodin - pro volný den vložíme 0
                if start_time == "00:00" and end_time == "00:00" and lunch_duration == 0.0:
                    total_hours = 0
                else:
                    # Normální výpočet pro pracovní den
                    start = datetime.strptime(start_time, "%H:%M")
                    end = datetime.strptime(end_time, "%H:%M")
                    lunch_duration_float = float(lunch_duration)
                    total_hours = (end - start).total_seconds() / 3600 - lunch_duration_float
                    # Zaokrouhlení na 2 desetinná místa pro konzistenci
                    total_hours = round(total_hours, 2)

                # Ukládání dat pro každého zaměstnance
                start_row = 9 # Řádek, kde začínají jména zaměstnanců
                for employee in employees:
                    current_row = start_row
                    row_found = False
                    # Hledáme existujícího zaměstnance nebo první volný řádek
                    # Omezíme hledání na rozumný počet řádků, např. 1000
                    max_search_row = start_row + 1000

                    while current_row < max_search_row:
                        employee_cell = sheet.cell(row=current_row, column=1) # Sloupec A
                        if employee_cell.value == employee:
                            # Našli jsme existujícího
                            sheet.cell(row=current_row, column=start_time_col_index + 1, value=total_hours)
                            row_found = True
                            break
                        elif employee_cell.value is None or str(employee_cell.value).strip() == "":
                            # Našli jsme prázdný řádek, přidáme nového
                            employee_cell.value = employee
                            sheet.cell(row=current_row, column=start_time_col_index + 1, value=total_hours)
                            row_found = True
                            break
                        current_row += 1

                    if not row_found:
                        logger.warning(f"Nepodařilo se najít ani vytvořit řádek pro zaměstnance '{employee}' v listu '{sheet_name}' souboru {self.active_filename}. Možná je dosažen limit řádků.")

                # Ukládání časů začátku a konce do řádku 7
                start_time_col_letter = get_column_letter(start_time_col_index + 1)
                end_time_col_letter = get_column_letter(end_time_col_index + 1)
                sheet[f"{start_time_col_letter}7"] = start_time
                sheet[f"{end_time_col_letter}7"] = end_time

                # Uložení data do buňky v řádku 80
                date_col_letter = get_column_letter(date_col_index + 1)
                try:
                    date_obj = datetime.strptime(date, "%Y-%m-%d").date()
                    date_cell = sheet[f"{date_col_letter}80"]
                    date_cell.value = date_obj
                    date_cell.number_format = 'DD.MM.YYYY' # Nastavení formátu
                except ValueError:
                     logger.error(f"Neplatný formát data '{date}' při ukládání do buňky {date_col_letter}80.")
                     sheet[f"{date_col_letter}80"] = date # Uložíme jako text, pokud selže konverze

                # Zápis názvu projektu do B79 (pokud je nastaven)
                if self.current_project_name:
                    sheet["B79"] = self.current_project_name

                # Workbook se uloží automaticky context managerem _get_workbook
                logger.info(f"Úspěšně uložena pracovní doba pro datum {date} do listu {sheet_name} v souboru {self.active_filename}")
                return True

        except (FileNotFoundError, ValueError, IOError) as e:
             logger.error(f"Chyba při ukládání pracovní doby do {self.active_filename}: {e}", exc_info=True)
             return False
        except Exception as e:
            # Zachytí neočekávané chyby (např. z openpyxl)
            logger.error(f"Neočekávaná chyba při ukládání pracovní doby do {self.active_filename}: {e}", exc_info=True)
            return False

    def update_project_info(self, project_name, start_date, end_date=None):
        """Aktualizuje informace o projektu v listu Zálohy aktivního souboru"""
        active_path = self.get_active_file_path()
        try:
            with self._get_workbook(active_path) as workbook:
                # Uložíme název projektu do instance pro použití v ulozit_pracovni_dobu
                self.set_project_name(project_name)

                # Zajistíme existenci listu Zálohy
                if "Zálohy" not in workbook.sheetnames:
                    workbook.create_sheet("Zálohy")
                    logger.info(f"Vytvořen list 'Zálohy' v souboru {self.active_filename}, protože neexistoval.")
                    # Můžeme přidat výchozí hodnoty do nového listu Zálohy
                    zalohy_sheet = workbook["Zálohy"]
                    zalohy_sheet["B80"] = "Option 1"
                    zalohy_sheet["D80"] = "Option 2"
                else:
                    zalohy_sheet = workbook["Zálohy"]

                # Aktualizace buněk A79, C81, D81
                zalohy_sheet["A79"] = project_name

                # Zpracování datumu začátku
                try:
                    start_date_obj = datetime.strptime(start_date, "%Y-%m-%d")
                    cell_c81 = zalohy_sheet["C81"]
                    cell_c81.value = start_date_obj
                    cell_c81.number_format = 'DD.MM.YY'
                except (ValueError, TypeError):
                    logger.warning(f"Neplatný formát data začátku '{start_date}' pro projekt '{project_name}', buňka C81 nebude aktualizována.")
                    zalohy_sheet["C81"] = None

                # Zpracování datumu konce
                if end_date:
                    try:
                        end_date_obj = datetime.strptime(end_date, "%Y-%m-%d")
                        cell_d81 = zalohy_sheet["D81"]
                        cell_d81.value = end_date_obj
                        cell_d81.number_format = 'DD.MM.YY'
                    except (ValueError, TypeError):
                         logger.warning(f"Neplatný formát data konce '{end_date}' pro projekt '{project_name}', buňka D81 nebude aktualizována.")
                         zalohy_sheet["D81"] = None
                else:
                    # Pokud end_date není zadáno, vymažeme buňku
                    zalohy_sheet["D81"] = None

                # Workbook se uloží automaticky
                logger.info(f"Aktualizovány informace o projektu '{project_name}' v souboru {self.active_filename}")
                return True
        except (FileNotFoundError, ValueError, IOError) as e:
             logger.error(f"Chyba při aktualizaci informací o projektu v {self.active_filename}: {e}", exc_info=True)
             return False
        except Exception as e:
            logger.error(f"Neočekávaná chyba při aktualizaci informací o projektu v {self.active_filename}: {e}", exc_info=True)
            return False

    def set_project_name(self, project_name):
         """Nastaví aktuální název projektu pro použití v jiných metodách."""
         self.current_project_name = project_name if project_name else None


    def get_advance_options(self):
        """Získá možnosti záloh z aktivního Excel souboru"""
        # Pokud aktivní soubor není nastaven, vrátíme výchozí
        if not self.active_file_path:
             logger.warning("Nelze načíst možnosti záloh, aktivní soubor není nastaven. Používají se výchozí.")
             return ["Option 1", "Option 2"]

        try:
            # Použijeme read-only mód
            with self._get_workbook(self.active_file_path, read_only=True) as workbook:
                options = []
                default_options = ["Option 1", "Option 2"]

                if "Zálohy" in workbook.sheetnames:
                    zalohy_sheet = workbook["Zálohy"]
                    option1 = zalohy_sheet["B80"].value
                    option2 = zalohy_sheet["D80"].value
                    options = [
                        str(option1).strip() if option1 else default_options[0],
                        str(option2).strip() if option2 else default_options[1]
                    ]
                    logger.info(f"Načteny možnosti záloh z {self.active_filename}: {options}")
                else:
                    logger.warning(f"List 'Zálohy' nebyl nalezen v souboru {self.active_filename}, použity výchozí možnosti.")
                    options = default_options

                return options
        except FileNotFoundError:
             logger.error(f"Aktivní soubor {self.active_filename} nebyl nalezen při načítání možností záloh.")
             return ["Option 1", "Option 2"]
        except Exception as e:
            logger.error(f"Chyba při načítání možností záloh z {self.active_filename}: {str(e)}", exc_info=True)
            return ["Option 1", "Option 2"]

    def save_advance(self, employee_name, amount, currency, option, date):
        """Uloží zálohu do aktivního Excel souboru."""
        active_path = self.get_active_file_path()
        try:
            # Použití context manageru pro aktivní workbook
            with self._get_workbook(active_path) as workbook:
                # Uložení do listu 'Zálohy' v aktivním souboru
                self._save_advance_main(workbook, employee_name, amount, currency, option, date)
                # Workbook se uloží automaticky

            logger.info(f"Záloha pro {employee_name} ({amount} {currency}, {option}, {date}) úspěšně uložena do {self.active_filename}")
            return True

        except (FileNotFoundError, ValueError, IOError) as e:
            logger.error(f"Chyba při ukládání zálohy do {self.active_filename}: {str(e)}", exc_info=True)
            return False
        except Exception as e:
            logger.error(f"Neočekávaná chyba při ukládání zálohy do {self.active_filename}: {str(e)}", exc_info=True)
            return False

    def _save_advance_main(self, workbook, employee_name, amount, currency, option, date):
        """Pomocná metoda pro ukládání zálohy do listu 'Zálohy' daného workbooku"""
        sheet_name = "Zálohy"
        if sheet_name not in workbook.sheetnames:
            sheet = workbook.create_sheet(sheet_name)
            sheet["B80"] = "Option 1"
            sheet["D80"] = "Option 2"
            logger.info(f"Vytvořen list '{sheet_name}' s výchozími názvy možností.")
        else:
            sheet = workbook[sheet_name]

        option1_value = sheet["B80"].value or "Option 1"
        option2_value = sheet["D80"].value or "Option 2"

        # Najdeme řádek pro zaměstnance nebo vytvoříme nový
        row_index = 9 # Startovní řádek
        found_row = None
        for r in range(row_index, sheet.max_row + 2): # +2 pro případ přidání na konec
             cell_a = sheet.cell(row=r, column=1)
             if cell_a.value == employee_name:
                  found_row = r
                  break
             if cell_a.value is None or str(cell_a.value).strip() == "":
                  found_row = r
                  cell_a.value = employee_name # Zapíšeme jméno nového
                  break
        if found_row is None:
             # Pokud jsme ani po prohledání nenašli/nevytvořili řádek (velmi nepravděpodobné)
             logger.error(f"Nepodařilo se najít ani vytvořit řádek pro zálohu zaměstnance '{employee_name}'")
             raise IOError(f"Nelze najít/vytvořit řádek pro zaměstnance '{employee_name}' v listu '{sheet_name}'")


        # Určíme sloupec podle možnosti a měny (indexy od 1)
        if option == option1_value:
            column_index = 2 if currency == "EUR" else 3 # B nebo C
        elif option == option2_value:
            column_index = 4 if currency == "EUR" else 5 # D nebo E
        else:
            logger.error(f"Neznámá možnost zálohy '{option}'. Používají se sloupce pro '{option1_value}' (B/C).")
            column_index = 2 if currency == "EUR" else 3

        # Aktualizujeme hodnotu zálohy
        target_cell = sheet.cell(row=found_row, column=column_index)
        current_value = target_cell.value or 0
        try:
             current_value_float = float(current_value)
             amount_float = float(amount)
             new_value = current_value_float + amount_float
             target_cell.value = new_value
             # Můžeme nastavit formát čísla, např. na dvě desetinná místa
             target_cell.number_format = '#,##0.00'
        except (ValueError, TypeError):
             logger.error(f"Neplatná hodnota v buňce {target_cell.coordinate} ('{current_value}') nebo neplatná částka '{amount}'. Záloha nebude přičtena.")
             # Můžeme zde vyvolat chybu nebo jen zalogovat a nepokračovat
             raise ValueError("Neplatná číselná hodnota pro zálohu.")


        # Přidání data zálohy do sloupce Z (index 25 -> sloupec 26 v openpyxl)
        date_column_index = 26
        date_cell = sheet.cell(row=found_row, column=date_column_index)
        try:
             date_obj = datetime.strptime(date, "%Y-%m-%d").date()
             date_cell.value = date_obj
             date_cell.number_format = 'DD.MM.YYYY'
        except ValueError:
             logger.error(f"Neplatný formát data '{date}' pro zálohu. Datum nebude uloženo.")
             date_cell.value = None


    # Metoda _get_next_empty_row_in_column není potřeba, pokud hledáme řádek výše uvedeným způsobem

    def ziskej_cislo_tydne(self, datum):
        """
        Získá ISO kalendářní data (rok, číslo týdne, den v týdnu) pro zadané datum.
        """
        try:
            if isinstance(datum, str):
                datum_obj = datetime.strptime(datum, "%Y-%m-%d")
            elif isinstance(datum, datetime):
                 datum_obj = datum
            else:
                 raise TypeError("Datum musí být string ve formátu YYYY-MM-DD nebo datetime objekt")

            return datum_obj.isocalendar()
        except (ValueError, TypeError) as e:
            logger.error(f"Chyba při zpracování data '{datum}' pro získání čísla týdne: {e}")
            # Vrátíme None nebo vyvoláme výjimku, aby volající věděl o chybě
            return None

    def record_time(self, employee, date, start_time, end_time, lunch_duration=1.0):
        """
        Zaznamená pracovní dobu pro jednoho nebo více zaměstnanců.
        
        Args:
            employee (Union[str, List[str]]): Jméno zaměstnance nebo seznam jmen zaměstnanců
            date (str): Datum ve formátu YYYY-MM-DD
            start_time (str): Čas začátku ve formátu HH:MM
            end_time (str): Čas konce ve formátu HH:MM
            lunch_duration (float): Délka oběda v hodinách, výchozí 1.0
            
        Returns:
            tuple: (success, message)
        """
        try:
            # Převedeme jeden string na seznam pro jednotné zpracování
            employees = employee if isinstance(employee, list) else [employee]
            
            # Zavoláme existující metodu ulozit_pracovni_dobu se seznamem zaměstnanců
            success = self.ulozit_pracovni_dobu(
                date=date,
                start_time=start_time,
                end_time=end_time,
                lunch_duration=lunch_duration,
                employees=employees
            )
            
            if success:
                message = "Záznam byl úspěšně uložen"
                logger.info(f"Úspěšně uložen záznam pro zaměstnance {', '.join(employees)}: Datum {date}, Začátek {start_time}, Konec {end_time}, Oběd {lunch_duration}")
            else:
                message = "Nepodařilo se uložit záznam"
                logger.error(f"Nepodařilo se uložit záznam pro zaměstnance {', '.join(employees)}")
                
            return success, message

        except Exception as e:
            logger.error(f"Chyba při ukládání záznamu: {e}", exc_info=True)
            return False, str(e)
