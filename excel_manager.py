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
import datetime # Přidán import datetime


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

        Životní cyklus sešitu a cachování:
        1. Otevření/Načtení:
           - Metoda nejprve zkontroluje, zda je platný workbook pro danou `file_path_to_open`
             již v interní cache (`self._workbook_cache`).
           - Cache klíč je absolutní cesta k souboru.
           - Pokud je v cache nalezen použitelný workbook (ověřeno jednoduchým testem `wb.sheetnames`),
             je tento workbook použit (`is_from_cache = True`).
           - Pokud workbook v cache není, nebo je neplatný (např. byl mezitím zavřen nebo poškozen),
             je načten ze souborového systému pomocí `load_workbook()`.
           - Při načítání ze souboru:
             - Pokud je `read_only=False`, nově načtený workbook je přidán do cache.
               To umožňuje následným operacím ve stejném kontextu (např. v rámci jednoho HTTP requestu)
               pracovat se stejnou instancí workbooku bez nutnosti opakovaného načítání a pro zápis změn.
             - Pokud je `read_only=True`, workbook se do cache neukládá.

        2. Použití workbooku:
           - Workbook (buď z cache nebo nově načtený) je poskytnut volajícímu kódu přes `yield wb`.

        3. Po použití (v `try` bloku za `yield`):
           - Ukládání na disk:
             - Pokud byl workbook otevřen pro zápis (`read_only=False`), *nebyl* načten z cache
               (tj. byl nově načten v tomto volání `_get_workbook`), a je stále platný (`wb is not None`),
               jeho změny jsou uloženy na disk pomocí `wb.save(str(file_path))`.
             - Toto okamžité uložení je pro případy, kdy se s workbookem pracuje mimo cache
               nebo je to první operace, která ho do cache přidá.
             - Workbooky načtené z cache (které jsou vždy pro zápis) se zde *neukládají*. Jejich finální
               uložení proběhne až při volání `_clear_workbook_cache()` (typicky na konci requestu
               nebo při destrukci instance `ExcelManager`).

        4. Zpracování chyb:
           - Pokud dojde k jakékoliv chybě během načítání nebo používání workbooku,
             a pokud je tento workbook v cache, je z cache odstraněn, aby se předešlo
             použití potenciálně poškozeného nebo nekonzistentního stavu.

        5. Uzavření (v `finally` bloku):
           - Read-only workbooky:
             - Pokud byl workbook otevřen jako `read_only=True` a je platný (`wb is not None`),
               je vždy uzavřen pomocí `wb.close()`, bez ohledu na to, zda byl (hypoteticky) z cache
               nebo nově načten. Read-only workbooky se totiž do cache standardně neukládají,
               takže je bezpečné je po použití ihned uzavřít.
           - Workbooky pro zápis (writeable):
             - Pokud byl workbook otevřen pro zápis (`read_only=False`) a je v cache,
               zde se *nezavírá* ani *neukládá*. Jeho životní cyklus (uložení a následné zavření)
               je plně spravován metodou `_clear_workbook_cache()`. Tím se zajišťuje,
               že všechny změny provedené během jeho "života" v cache jsou kumulativní
               a uloží se najednou.
             - Workbooky pro zápis, které *nebyly* z cache a byly uloženy v bodě 3,
               zůstávají otevřené a jsou v cache (pokud byly přidány). Jejich uzavření
               také proběhne až přes `_clear_workbook_cache()`.

        Args:
            file_path_to_open (str nebo Path): Cesta k Excel souboru, který má být otevřen.
            read_only (bool): Pokud True, soubor se otevře pouze pro čtení.
                              Výchozí je False (otevření pro čtení i zápis).

        Yields:
            openpyxl.Workbook: Otevřený workbook.

        Raises:
            FileNotFoundError: Pokud soubor na `file_path_to_open` neexistuje.
            IOError: Pokud dojde k chybě při otevírání nebo ukládání souboru.
            Exception: Jiné obecné chyby.
        """
        # Použijeme předanou cestu
        file_path = Path(file_path_to_open)
        cache_key = str(file_path.absolute()) # Cache klíč je absolutní cesta
        wb = None
        is_from_cache = False # Příznak, zda byl workbook načten z cache

        # Zámek pro zajištění thread-safe operací se souborem a cache
        with self._file_lock:
            try:
                # Fáze 1: Pokus o načtení z cache
                if cache_key in self._workbook_cache:
                    try:
                        wb = self._workbook_cache[cache_key]
                        # Jednoduchý test, zda je workbook stále použitelný (např. nebyl externě zavřen)
                        _ = wb.sheetnames  # Vyvolá chybu, pokud je wb zavřený nebo neplatný
                        is_from_cache = True
                        logger.debug(f"Workbook načten z cache: {cache_key}")
                    except Exception as cache_err:
                        logger.warning(f"Chyba při použití workbooku z cache ({cache_key}): {cache_err}. Workbook bude znovu načten.")
                        # Odstraníme neplatný workbook z cache
                        # Nejprve zkusíme zavřít, pokud by byl stále nějakým způsobem "aktivní"
                        try:
                             if self._workbook_cache[cache_key] is not None: # Kontrola pro jistotu
                                 self._workbook_cache[cache_key].close()
                        except Exception as close_ex:
                             logger.warning(f"Nepodařilo se zavřít neplatný workbook z cache ({cache_key}): {close_ex}")
                        finally:
                             del self._workbook_cache[cache_key] # Odstranění z cache
                        wb = None # Resetujeme wb, aby se načetl ze souboru
                        is_from_cache = False # Resetujeme příznak

                # Fáze 1 (pokračování): Pokud není v cache nebo byl neplatný, načteme ze souboru
                if wb is None:
                    # Ujistíme se, že adresář pro soubor existuje
                    file_path.parent.mkdir(parents=True, exist_ok=True)

                    if not file_path.exists():
                         logger.error(f"Soubor {file_path} neexistuje!")
                         raise FileNotFoundError(f"Požadovaný Excel soubor '{file_path.name}' nebyl nalezen na cestě '{file_path}'.")

                    try:
                        # Načteme workbook ze souboru
                        wb = load_workbook(filename=str(file_path), read_only=read_only, data_only=True)
                        logger.debug(f"Workbook načten ze souboru: {file_path} (read_only={read_only})")
                        
                        # Pokud je workbook otevřen pro zápis (read_only=False), přidáme ho do cache.
                        # To umožní dalším operacím v rámci tohoto kontextu (např. requestu)
                        # pracovat se stejnou instancí a kumulovat změny.
                        if not read_only:
                            self._workbook_cache[cache_key] = wb
                            logger.debug(f"Workbook přidán do cache: {cache_key}")
                    except Exception as load_err:
                        logger.error(f"Nelze načíst Excel soubor {file_path}: {load_err}", exc_info=True)
                        raise IOError(f"Chyba při otevírání souboru '{file_path.name}'.")

                # Fáze 2: Předáme workbook kódu uvnitř 'with' bloku
                yield wb

                # Fáze 3: Po skončení 'with' bloku (pokud nebyly výjimky v 'yield' části)
                # Ukládání na disk pro workbooky, které nebyly z cache a jsou pro zápis.
                # Workbooky, které byly načteny z cache (a jsou tedy vždy pro zápis),
                # se zde neukládají. Jejich uložení je řízeno metodou _clear_workbook_cache().
                if not read_only and wb is not None and not is_from_cache:
                    try:
                        wb.save(str(file_path))
                        logger.debug(f"Workbook (nově načtený, pro zápis) uložen na disk po 'with' bloku: {file_path}")
                    except Exception as save_err:
                        logger.error(f"Chyba při ukládání workbooku {file_path} po 'with' bloku: {save_err}", exc_info=True)
                        # Pokud uložení selže, a workbook je v cache (což by měl být, pokud not read_only),
                        # odstraníme ho z cache, aby se neuložil potenciálně poškozený stav později.
                        if cache_key in self._workbook_cache:
                             try:
                                 # Není třeba zavírat, protože _clear_workbook_cache se o to postará,
                                 # nebo pokud by byl workbook poškozen, close může selhat.
                                 # Stačí odstranit z cache.
                                 del self._workbook_cache[cache_key]
                                 logger.warning(f"Workbook odstraněn z cache kvůli chybě při ukládání: {cache_key}")
                             except KeyError:
                                 pass # Workbook už tam z nějakého důvodu není
                        raise IOError(f"Chyba při ukládání změn do souboru '{file_path.name}'.")

            except Exception as e: # Fáze 4: Zpracování chyb
                logger.error(f"Obecná chyba v _get_workbook pro {file_path}: {e}", exc_info=True)
                # Pokud došlo k chybě (buď při načítání, nebo v 'yield' bloku),
                # a workbook je v cache, odstraníme ho, aby se předešlo použití nekonzistentního stavu.
                if cache_key in self._workbook_cache:
                    try:
                        # Opět, není třeba explicitně zavírat zde, stačí odstranit z cache.
                        # _clear_workbook_cache se postará o případné zavření, pokud je to nutné.
                        del self._workbook_cache[cache_key]
                        logger.info(f"Workbook odstraněn z cache kvůli obecné chybě: {cache_key}")
                    except KeyError:
                        pass # Workbook už tam z nějakého důvodu není
                raise # Znovu vyvoláme výjimku, aby byla zpracována volajícím kódem

            finally: # Fáze 5: Uzavření
                # Tento blok se vykoná vždy, i po výjimkách.

                # Read-only workbooky:
                # Pokud byl workbook otevřen jako read_only=True, měl by být uzavřen.
                # Standardně se read-only workbooky do cache neukládají.
                # Pokud by se tam však nějakým způsobem dostal (což by byla chyba v logice),
                # i takový read-only workbook z cache by měl být zde uzavřen,
                # protože _clear_workbook_cache je primárně pro zapisovatelné workbooky.
                if read_only and wb is not None:
                     try:
                          wb.close()
                          logger.debug(f"Read-only workbook uzavřen v finally: {file_path}")
                          # Pokud byl read-only workbook (hypoteticky) v cache, odstraníme ho,
                          # protože read-only instance by v cache neměly přetrvávat.
                          if cache_key in self._workbook_cache and self._workbook_cache[cache_key] is wb:
                              del self._workbook_cache[cache_key]
                              logger.debug(f"Read-only workbook odstraněn z cache v finally: {cache_key}")
                     except Exception as close_err:
                          logger.warning(f"Chyba při zavírání read-only workbooku {file_path} v finally: {close_err}")
                
                # Workbooky pro zápis (writeable), které jsou v cache:
                # Tyto workbooky se zde *NEZAVÍRAJÍ* ani *NEUKLÁDAJÍ*.
                # Jejich životní cyklus (uložení všech změn a následné zavření) je plně
                # spravován metodou `_clear_workbook_cache()`. To zajišťuje, že všechny změny
                # provedené během "života" workbooku v cache jsou kumulativní a uloží se najednou
                # na konci (např. na konci HTTP requestu nebo při destrukci ExcelManageru).
                # Pokud by se zde zavřely, ztratily by se neuložené změny nebo by se cache stala neplatnou.
                # Tato logika je implicitní - pokud `read_only` je `False`, neděláme nic s `wb` zde,
                # protože jeho správa je přenechána `_clear_workbook_cache`.
                elif not read_only and wb is not None and cache_key in self._workbook_cache:
                    logger.debug(f"Workbook pro zápis '{file_path}' je v cache a nebude zde uzavřen ani uložen. Správu přebírá _clear_workbook_cache.")


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

                sheet_name = f"{Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME} {week_number}"

                # Vytvoření nebo získání listu pro daný týden
                if sheet_name not in workbook.sheetnames:
                    # Zkusíme zkopírovat ze šablonového listu Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME v aktivním souboru
                    if Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME in workbook.sheetnames:
                        source_sheet = workbook[Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME]
                        try:
                            sheet = workbook.copy_worksheet(source_sheet)
                            sheet.title = sheet_name
                            logger.info(f"Zkopírován list '{Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME}' a přejmenován na '{sheet_name}' v souboru {self.active_filename}")
                        except Exception as copy_err:
                             logger.error(f"Nepodařilo se zkopírovat list '{Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME}' v {self.active_filename}: {copy_err}", exc_info=True)
                             # Pokud kopírování selže, vytvoříme prázdný list
                             sheet = workbook.create_sheet(sheet_name)
                             logger.warning(f"Vytvořen nový prázdný list '{sheet_name}' v souboru {self.active_filename} kvůli chybě při kopírování.")
                    else:
                        # Pokud ani šablonový list Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME neexistuje, vytvoříme prázdný
                        sheet = workbook.create_sheet(sheet_name)
                        logger.warning(f"Vytvořen nový prázdný list '{sheet_name}' v souboru {self.active_filename} (šablona '{Config.EXCEL_WEEK_SHEET_TEMPLATE_NAME}' nenalezena).")

                    # Nastavení názvu týdne do buňky A80
                    sheet["A80"] = sheet_name # Název listu (např. "Týden 5") se ukládá do A80
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
                start_row = Config.EXCEL_EMPLOYEE_START_ROW
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
                            # Použijeme end_time_col_index + 2 pro sloupec "Celkem hodin"
                            # Po (weekday 0): day_column_index=1, start_time_col_index=1, end_time_col_index=2. Sloupec pro total_hours = 2+2=4 (D)
                            # Út (weekday 1): day_column_index=3, start_time_col_index=3, end_time_col_index=4. Sloupec pro total_hours = 4+2=6 (F)
                            sheet.cell(row=current_row, column=end_time_col_index + 2, value=total_hours)
                            row_found = True
                            break
                        elif employee_cell.value is None or str(employee_cell.value).strip() == "":
                            # Našli jsme prázdný řádek, přidáme nového
                            employee_cell.value = employee
                            sheet.cell(row=current_row, column=end_time_col_index + 2, value=total_hours)
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
                     logger.error(f"Neplatný formát data '{date}' při ukládání do buňky {date_col_letter}80. Ukládání pracovní doby selhalo.")
                     # sheet[f"{date_col_letter}80"] = date # Neukládáme nevalidní datum jako text, operace selže
                     return False

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
                if Config.EXCEL_ADVANCES_SHEET_NAME not in workbook.sheetnames:
                    workbook.create_sheet(Config.EXCEL_ADVANCES_SHEET_NAME)
                    logger.info(f"Vytvořen list '{Config.EXCEL_ADVANCES_SHEET_NAME}' v souboru {self.active_filename}, protože neexistoval.")
                    # Můžeme přidat výchozí hodnoty do nového listu Zálohy
                    zalohy_sheet = workbook[Config.EXCEL_ADVANCES_SHEET_NAME]
                    zalohy_sheet["B80"] = Config.DEFAULT_ADVANCE_OPTION_1
                    zalohy_sheet["D80"] = Config.DEFAULT_ADVANCE_OPTION_2
                else:
                    zalohy_sheet = workbook[Config.EXCEL_ADVANCES_SHEET_NAME]

                # Aktualizace buněk A79, C81, D81
                zalohy_sheet["A79"] = project_name

                # Zpracování datumu začátku
                try:
                    start_date_obj = datetime.strptime(start_date, "%Y-%m-%d")
                    cell_c81 = zalohy_sheet["C81"]
                    cell_c81.value = start_date_obj
                    cell_c81.number_format = 'DD.MM.YY'
                except (ValueError, TypeError):
                    logger.warning(f"Neplatný formát data začátku '{start_date}' pro projekt '{project_name}'. Aktualizace projektu selhala.")
                    # zalohy_sheet["C81"] = None # Nechceme měnit hodnotu, pokud je chyba
                    return False

                # Zpracování datumu konce
                if end_date:
                    try:
                        end_date_obj = datetime.strptime(end_date, "%Y-%m-%d")
                        cell_d81 = zalohy_sheet["D81"]
                        cell_d81.value = end_date_obj
                        cell_d81.number_format = 'DD.MM.YY'
                    except (ValueError, TypeError):
                         logger.warning(f"Neplatný formát data konce '{end_date}' pro projekt '{project_name}'. Aktualizace projektu selhala.")
                         # zalohy_sheet["D81"] = None # Nechceme měnit hodnotu, pokud je chyba
                         return False
                else:
                    # Pokud end_date není zadáno, vymažeme buňku (to je v pořádku)
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
             return [Config.DEFAULT_ADVANCE_OPTION_1, Config.DEFAULT_ADVANCE_OPTION_2]

        try:
            # Použijeme read-only mód
            with self._get_workbook(self.active_file_path, read_only=True) as workbook:
                options = []
                default_options = [Config.DEFAULT_ADVANCE_OPTION_1, Config.DEFAULT_ADVANCE_OPTION_2]

                if Config.EXCEL_ADVANCES_SHEET_NAME in workbook.sheetnames:
                    zalohy_sheet = workbook[Config.EXCEL_ADVANCES_SHEET_NAME]
                    option1 = zalohy_sheet["B80"].value
                    option2 = zalohy_sheet["D80"].value
                    options = [
                        str(option1).strip() if option1 else default_options[0],
                        str(option2).strip() if option2 else default_options[1]
                    ]
                    logger.info(f"Načteny možnosti záloh z {self.active_filename}: {options}")
                else:
                    logger.warning(f"List '{Config.EXCEL_ADVANCES_SHEET_NAME}' nebyl nalezen v souboru {self.active_filename}, použity výchozí možnosti.")
                    options = default_options

                return options
        except FileNotFoundError:
             logger.error(f"Aktivní soubor {self.active_filename} nebyl nalezen při načítání možností záloh.")
             return [Config.DEFAULT_ADVANCE_OPTION_1, Config.DEFAULT_ADVANCE_OPTION_2]
        except Exception as e:
            logger.error(f"Chyba při načítání možností záloh z {self.active_filename}: {str(e)}", exc_info=True)
            return [Config.DEFAULT_ADVANCE_OPTION_1, Config.DEFAULT_ADVANCE_OPTION_2]

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
        """Pomocná metoda pro ukládání zálohy do listu Config.EXCEL_ADVANCES_SHEET_NAME daného workbooku"""
        sheet_name = Config.EXCEL_ADVANCES_SHEET_NAME
        if sheet_name not in workbook.sheetnames:
            sheet = workbook.create_sheet(sheet_name)
            sheet["B80"] = Config.DEFAULT_ADVANCE_OPTION_1
            sheet["D80"] = Config.DEFAULT_ADVANCE_OPTION_2
            logger.info(f"Vytvořen list '{sheet_name}' s výchozími názvy možností.")
        else:
            sheet = workbook[sheet_name]

        option1_value = sheet["B80"].value or Config.DEFAULT_ADVANCE_OPTION_1
        option2_value = sheet["D80"].value or Config.DEFAULT_ADVANCE_OPTION_2

        # Najdeme řádek pro zaměstnance nebo vytvoříme nový
        start_row = Config.EXCEL_EMPLOYEE_START_ROW # Použití konstanty
        found_row = None
        # Hledáme od start_row, ne od pevně daného indexu 9
        for r in range(start_row, sheet.max_row + 2): # +2 pro případ přidání na konec
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
            # Nahrazeno logování a fallback vyvoláním ValueError
            raise ValueError(f"Neznámá možnost zálohy: '{option}'. Platné možnosti jsou: '{option1_value}', '{option2_value}'.")

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

    def generate_monthly_report(self, month, year, employees=None):
        """
        Generuje měsíční report odpracovaných hodin pro zaměstnance.

        Args:
            month (int): Měsíc (1-12).
            year (int): Rok.
            employees (list, optional): Seznam jmen zaměstnanců. Defaults to None (všichni).

        Returns:
            dict: Slovník s reportem. Klíče jsou jména zaměstnanců, hodnoty jsou slovníky
                  s "total_hours" a "free_days".
                  Příklad: {"Novák Josef": {"total_hours": 160, "free_days": 2}}

        Raises:
            ValueError: Pokud je neplatný měsíc nebo rok.
        """
        if not (1 <= month <= 12):
            logger.error(f"Neplatný měsíc: {month}. Musí být mezi 1 a 12.")
            raise ValueError("Měsíc musí být v rozsahu 1-12.")
        if not (2000 <= year <= 2100): # Přiměřený rozsah pro rok
            logger.error(f"Neplatný rok: {year}. Musí být mezi 2000 a 2100.")
            raise ValueError("Rok musí být v rozsahu 2000-2100.")

        logger.info(f"Generování měsíčního reportu pro {month}/{year}. Zaměstnanci: {employees if employees else 'Všichni'}")

        report_data = {}
        try:
            active_path = self.get_active_file_path()
            with self._get_workbook(active_path, read_only=True) as wb:
                if not wb:
                    logger.error("Nepodařilo se otevřít workbook pro generování reportu.")
                    return {}

                employee_data_template = {"total_hours": 0, "free_days": 0}
                # Inicializace report_data pro specifikované zaměstnance, pokud jsou zadáni
                if employees:
                    for emp_name in employees:
                        report_data[emp_name] = employee_data_template.copy()

                for sheet_name in wb.sheetnames:
                    if not sheet_name.startswith("Týden"):
                        continue # Přeskočíme listy, které neodpovídají vzoru

                    sheet = wb[sheet_name]
                    logger.info(f"Zpracovávám list: {sheet_name}")

                    # hour_columns_indices: 1-based indexy sloupců, které označují ZAČÁTEK bloku pro daný den.
                    # Pondělí začíná ve sloupci B (index 2), Úterý ve sloupci D (index 4), atd.
                    # Celkové odpracované hodiny pro daný den jsou uloženy ve sloupci o 2 vyšším než tento základní index
                    # (např. pro Pondělí: B (2) -> hodiny v D (4); pro Úterý: D (4) -> hodiny v F (6)).
                    hour_columns_indices = [2, 4, 6, 8, 10, 12, 14] 
                    date_cells_coords = ["B80", "D80", "F80", "H80", "J80", "L80", "N80"] # Buňky s daty pro dny v týdnu

                    # Načtení dat pro tento týden
                    week_dates = []
                    for col_idx, date_cell_coord in enumerate(date_cells_coords):
                        date_cell_value = sheet[date_cell_coord].value
                        if isinstance(date_cell_value, datetime.datetime): # Očekáváme datetime objekty
                            week_dates.append(date_cell_value.date())
                        elif isinstance(date_cell_value, str): # Pokus o převod, pokud je string
                            try:
                                week_dates.append(datetime.datetime.strptime(date_cell_value, "%d.%m.%Y").date())
                            except ValueError:
                                try:
                                    week_dates.append(datetime.datetime.strptime(date_cell_value, "%Y-%m-%d").date())
                                except ValueError:
                                    logger.warning(f"Neplatný formát data v buňce {date_cell_coord} listu {sheet_name}: {date_cell_value}. Tento den bude ignorován.")
                                    week_dates.append(None) # Chyba při čtení data
                        else:
                            week_dates.append(None) # Žádné datum nebo neznámý typ

                    # Procházení řádků se zaměstnanci (od Config.EXCEL_EMPLOYEE_START_ROW)
                    for row_num in range(Config.EXCEL_EMPLOYEE_START_ROW, sheet.max_row + 1):
                        employee_name_cell = sheet.cell(row=row_num, column=1) # Sloupec A
                        employee_name = employee_name_cell.value
                        if not employee_name or not str(employee_name).strip():
                            # Pokud je jméno prázdné, předpokládáme konec seznamu zaměstnanců pro tento list
                            break 
                        
                        employee_name = str(employee_name).strip()

                        # Pokud jsou specifikováni zaměstnanci a tento není v seznamu, přeskočíme ho
                        if employees and employee_name not in employees:
                            continue

                        # Pokud zaměstnanec ještě není v report_data, inicializujeme ho
                        if employee_name not in report_data:
                            report_data[employee_name] = employee_data_template.copy()

                        for day_idx, base_day_column_index in enumerate(hour_columns_indices):
                            # actual_col_idx je 1-based index sloupce, kde jsou uloženy celkové hodiny pro daný den.
                            # Vypočítá se jako základní index sloupce dne + 2.
                            # Příklad Pondělí: base_day_column_index = 2 (sloupec B). Celkem hodin je v 2+2=4 (sloupec D).
                            # Příklad Úterý: base_day_column_index = 4 (sloupec D). Celkem hodin je v 4+2=6 (sloupec F).
                            actual_col_idx = base_day_column_index + 2
                            
                            current_date = week_dates[day_idx]
                            if not current_date: # Pokud se nepodařilo přečíst datum pro tento den
                                continue

                            # Kontrola, zda datum spadá do požadovaného měsíce a roku
                            if current_date.month == month and current_date.year == year:
                                hours_cell = sheet.cell(row=row_num, column=actual_col_idx)
                                hours_value = hours_cell.value

                                if hours_value is None or str(hours_value).strip() == "":
                                    # Považujeme za 0 hodin, ale ne nutně volný den (může být jen nevyplněno)
                                    # Pro účely tohoto reportu to ale znamená 0 odpracovaných hodin.
                                    # Volný den se počítá jen pokud je explicitně 0.
                                    pass # Nepřidáváme hodiny
                                else:
                                    try:
                                        hours_worked = float(hours_value)
                                        report_data[employee_name]["total_hours"] += hours_worked
                                        if hours_worked == 0:
                                            report_data[employee_name]["free_days"] += 1
                                    except (ValueError, TypeError):
                                        logger.warning(f"Neplatná hodnota hodin '{hours_value}' pro zaměstnance {employee_name} v listu {sheet_name}, buňka {hours_cell.coordinate}. Ignorováno.")
                
                # Po zpracování všech listů, odfiltrujeme zaměstnance, kteří nemají žádné hodiny ani volné dny.
                # Toto je relevantní hlavně v případě, kdy `employees` nebylo specifikováno,
                # a my jsme mohli potenciálně inicializovat záznam pro zaměstnance,
                # kteří v daném měsíci nemají žádná data.
                # Pokud `employees` bylo specifikováno, odstraníme pouze ty, pro které se nic nenašlo.
                
                final_report_data = {}
                for emp_name, data in report_data.items():
                    if data["total_hours"] > 0 or data["free_days"] > 0:
                        final_report_data[emp_name] = data
                report_data = final_report_data

            logger.info(f"Měsíční report pro {month}/{year} úspěšně vygenerován. Nalezeno záznamů pro {len(report_data)} zaměstnanců.")
            return report_data

        except FileNotFoundError:
            logger.error(f"Aktivní soubor {self.active_filename} nebyl nalezen při generování reportu.", exc_info=True)
            return {} # Vrací prázdný report v případě chyby
        except Exception as e:
            logger.error(f"Chyba při generování měsíčního reportu pro {month}/{year}: {e}", exc_info=True)
            return {} # Vrací prázdný report v případě obecné chyby
