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

from utils.logger import setup_logger

logger = setup_logger("excel_manager")

# Kontrola jestli je systém Windows
IS_WINDOWS = platform.system() == "Windows"


class ExcelManager:
    def __init__(self, base_path, excel_file_name):
        from pathlib import Path

        self.base_path = Path(base_path)
        self.file_path = self.base_path / excel_file_name
        # self.file_path_2025 = self.base_path / "Hodiny2025.xlsx" # Odstraněno
        self.current_project_name = None
        self._file_lock = Lock()
        self._workbook_cache = {}

    @contextmanager
    def _get_workbook(self, file_path, read_only=False):
        """Vylepšený context manager pro práci s workbookem"""
        # Konverze na Path objekt pro jednotnou práci s cestami
        file_path = Path(file_path)
        cache_key = str(file_path.absolute())
        wb = None

        try:
            if cache_key in self._workbook_cache:
                try:
                    # Test if workbook is still usable by accessing a property
                    wb = self._workbook_cache[cache_key]
                    _ = wb.sheetnames
                except Exception:
                    # If accessing properties fails, remove from cache
                    del self._workbook_cache[cache_key]
                    wb = None

            if wb is None:
                with self._file_lock:
                    # Ujistíme se, že adresář existuje
                    file_path.parent.mkdir(parents=True, exist_ok=True)

                    if file_path.exists():
                        try:
                            wb = load_workbook(str(file_path), read_only=read_only)
                        except Exception as e:
                            logger.error(f"Nelze načíst Excel soubor {file_path}: {str(e)}")
                            if not read_only:  # Pokud nejsme v read-only módu, vytvoříme nový soubor
                                logger.info(f"Vytvářím nový Excel soubor {file_path}")
                                wb = Workbook()
                            else:
                                raise  # V read-only módu propagujeme chybu
                    else:
                        if not read_only:  # Nový soubor vytvoříme jen když nejsme v read-only módu
                            logger.info(f"Vytvářím nový Excel soubor {file_path}")
                            wb = Workbook()
                        else:
                            logger.error(f"Soubor {file_path} neexistuje a nelze ho vytvořit v read-only módu")
                            raise FileNotFoundError(f"Soubor {file_path} nebyl nalezen")

                    if wb is not None:
                        self._workbook_cache[cache_key] = wb

            yield wb

            if not read_only and wb is not None:
                try:
                    # Ukládáme pouze pokud se nejedná o read-only operaci
                    wb.save(str(file_path))
                except Exception as e:
                    logger.error(f"Chyba při ukládání workbooku {file_path}: {str(e)}")
                    raise

        except Exception as e:
            logger.error(f"Chyba při práci s workbookem {file_path}: {str(e)}")
            # V případě chyby se pokusíme odstranit workbook z cache, pokud existuje
            if cache_key in self._workbook_cache:
                try:
                    self._workbook_cache[cache_key].close()
                except Exception:
                    pass # Ignorujeme chyby při zavírání
                del self._workbook_cache[cache_key]
            raise # Znovu vyvoláme původní výjimku
        finally:
            # Pokud je workbook v read-only módu a je v cache, zavřeme ho a odstraníme z cache
            if read_only and cache_key in self._workbook_cache:
                try:
                    self._workbook_cache[cache_key].close()
                except Exception:
                    pass # Ignorujeme chyby při zavírání
                del self._workbook_cache[cache_key]


    def _clear_workbook_cache(self):
        """Vylepšená metoda pro čištění cache"""
        for path, wb in list(self._workbook_cache.items()):
            try:
                # Pokusíme se uložit pouze pokud workbook není read-only (což zde nemůžeme přímo zjistit,
                # ale předpokládáme, že read-only workooky byly odstraněny v _get_workbook)
                # Bezpečnější je prostě jen zavřít.
                wb.close()
            except Exception as e:
                logger.error(f"Chyba při zavírání workbooku {path}: {e}")
            finally:
                # Odstraníme z cache bez ohledu na úspěch zavření
                self._workbook_cache.pop(path, None)

    def __del__(self):
        """Destruktor - zajistí uvolnění všech prostředků"""
        self._clear_workbook_cache()

    def ulozit_pracovni_dobu(self, date, start_time, end_time, lunch_duration, employees):
        """Uloží pracovní dobu do Excel souboru"""
        try:
            with self._get_workbook(self.file_path) as workbook:
                # Získání čísla týdne z datumu
                week_calendar_data = self.ziskej_cislo_tydne(date)
                week_number = week_calendar_data.week # Získáme číslo týdne

                sheet_name = f"Týden {week_number}"

                if sheet_name not in workbook.sheetnames:
                    if "Týden" in workbook.sheetnames:
                        source_sheet = workbook["Týden"]
                        # Použijeme WorksheetCopy pro kopírování listu
                        sheet = workbook.copy_worksheet(source_sheet)
                        sheet.title = sheet_name
                        logger.info(f"Zkopírován list 'Týden' a přejmenován na '{sheet_name}'")
                    else:
                        sheet = workbook.create_sheet(sheet_name)
                        logger.info(f"Vytvořen nový list '{sheet_name}'")
                    # Nastavení názvu týdne do buňky A80 i pro nově vytvořené/zkopírované listy
                    sheet["A80"] = sheet_name
                else:
                    sheet = workbook[sheet_name]

                # Výpočet odpracovaných hodin
                start = datetime.strptime(start_time, "%H:%M")
                end = datetime.strptime(end_time, "%H:%M")
                # Zajistíme, že lunch_duration je float
                lunch_duration_float = float(lunch_duration)
                total_hours = (end - start).total_seconds() / 3600 - lunch_duration_float

                # Určení sloupce podle dne v týdnu (0 = pondělí, 1 = úterý, atd.)
                weekday = datetime.strptime(date, "%Y-%m-%d").weekday()
                # Pro každý den posuneme o 2 sloupce (B,D,F,H,J,L,N)
                # Pondělí (0) -> B (index 1), Úterý (1) -> D (index 3), ...
                day_column_index = 1 + 2 * weekday
                day_column = openpyxl.utils.get_column_letter(day_column_index + 1) # +1 protože indexy sloupců jsou od 1

                # Ukládání dat pro každého zaměstnance
                start_row = 9
                for employee in employees:
                    # Hledání řádku pro zaměstnance
                    current_row = start_row
                    row_found = False
                    # Projdeme maximálně rozumný počet řádků, abychom nezacyklili
                    max_search_row = sheet.max_row + len(employees) + 5

                    while current_row <= max_search_row:
                        cell_value = sheet.cell(row=current_row, column=1).value # Sloupec A
                        if cell_value == employee:
                            # Našli jsme existujícího zaměstnance
                            sheet.cell(row=current_row, column=day_column_index + 1, value=total_hours)
                            row_found = True
                            break
                        elif cell_value is None or cell_value == "":
                            # Našli jsme prázdný řádek, přidáme nového zaměstnance
                            sheet.cell(row=current_row, column=1, value=employee) # Jméno do sloupce A
                            sheet.cell(row=current_row, column=day_column_index + 1, value=total_hours)
                            row_found = True
                            break
                        else:
                            current_row += 1

                    if not row_found:
                        logger.warning(f"Nepodařilo se najít ani vytvořit řádek pro zaměstnance '{employee}' v listu '{sheet_name}'. Možná je soubor příliš velký.")


                # Ukládání časů začátku a konce do řádku 7
                start_time_col_letter = day_column
                end_time_col_letter = openpyxl.utils.get_column_letter(day_column_index + 2) # Sloupec vpravo
                sheet[f"{start_time_col_letter}7"] = start_time
                sheet[f"{end_time_col_letter}7"] = end_time

                # Uložení data do buňky v řádku 80
                date_col_letter = day_column # Stejný sloupec jako hodiny
                # Uložíme jako datumový objekt pro lepší práci v Excelu
                date_obj = datetime.strptime(date, "%Y-%m-%d").date()
                sheet[f"{date_col_letter}80"] = date_obj
                # Můžeme nastavit i formát buňky, pokud je potřeba
                # sheet[f"{date_col_letter}80"].number_format = 'DD.MM.YYYY'

                # Zápis názvu projektu do B79
                if self.current_project_name:
                    sheet["B79"] = self.current_project_name

                # Workbook se uloží automaticky context managerem _get_workbook
                logger.info(f"Úspěšně uložena pracovní doba pro datum {date} do listu {sheet_name}")
                return True

        except Exception as e:
            logger.error(f"Chyba při ukládání pracovní doby: {e}", exc_info=True)
            return False

    def update_project_info(self, project_name, start_date, end_date=None):
        """Aktualizuje informace o projektu v listu Zálohy"""
        try:
            with self._get_workbook(self.file_path) as workbook:
                # Nastavíme název projektu pro použití v ulozit_pracovni_dobu
                self.set_project_name(project_name) # Uložíme název do instance

                if "Zálohy" not in workbook.sheetnames:
                    # Pokud list neexistuje, vytvoříme ho
                    workbook.create_sheet("Zálohy")
                    logger.info("Vytvořen list 'Zálohy', protože neexistoval.")

                zalohy_sheet = workbook["Zálohy"]
                zalohy_sheet["A79"] = project_name

                # Zpracování datumu začátku
                try:
                    start_date_obj = datetime.strptime(start_date, "%Y-%m-%d")
                    zalohy_sheet["C81"] = start_date_obj # Uložíme jako datetime objekt
                    zalohy_sheet["C81"].number_format = 'DD.MM.YY' # Nastavíme formát
                except (ValueError, TypeError):
                    logger.warning(f"Neplatný formát data začátku '{start_date}', buňka C81 nebude aktualizována.")
                    zalohy_sheet["C81"] = None # Nebo ponechat původní hodnotu

                # Zpracování datumu konce
                if end_date:
                    try:
                        end_date_obj = datetime.strptime(end_date, "%Y-%m-%d")
                        zalohy_sheet["D81"] = end_date_obj # Uložíme jako datetime objekt
                        zalohy_sheet["D81"].number_format = 'DD.MM.YY' # Nastavíme formát
                    except (ValueError, TypeError):
                         logger.warning(f"Neplatný formát data konce '{end_date}', buňka D81 nebude aktualizována.")
                         zalohy_sheet["D81"] = None # Nebo ponechat původní hodnotu
                else:
                    # Pokud end_date není zadáno, vymažeme buňku
                    zalohy_sheet["D81"] = None

                # Workbook se uloží automaticky context managerem
                logger.info(f"Aktualizovány informace o projektu: {project_name}")
                return True
        except Exception as e:
            logger.error(f"Chyba při aktualizaci informací o projektu: {e}", exc_info=True)
            return False

    def set_project_name(self, project_name):
         """Nastaví aktuální název projektu pro použití v jiných metodách."""
         self.current_project_name = project_name


    def get_advance_options(self):
        """Získá možnosti záloh z Excel souboru (list Zálohy, buňky B80, D80)"""
        try:
            # Použijeme read-only mód, protože jen čteme
            with self._get_workbook(self.file_path, read_only=True) as workbook:
                options = []
                default_options = ["Option 1", "Option 2"] # Výchozí hodnoty

                if "Zálohy" in workbook.sheetnames:
                    zalohy_sheet = workbook["Zálohy"]
                    option1 = zalohy_sheet["B80"].value
                    option2 = zalohy_sheet["D80"].value
                    # Použijeme hodnoty z buněk, pokud existují a nejsou prázdné, jinak výchozí
                    options = [
                        str(option1).strip() if option1 else default_options[0],
                        str(option2).strip() if option2 else default_options[1]
                    ]
                    logger.info(f"Načteny možnosti záloh: {options}")
                else:
                    logger.warning("List 'Zálohy' nebyl nalezen v Excel souboru, použity výchozí možnosti.")
                    options = default_options

                return options
        except FileNotFoundError:
             logger.error(f"Soubor {self.file_path} nebyl nalezen při načítání možností záloh.")
             return ["Option 1", "Option 2"] # Vrátíme výchozí v případě chyby
        except Exception as e:
            logger.error(f"Chyba při načítání možností záloh: {str(e)}", exc_info=True)
            return ["Option 1", "Option 2"] # Vrátíme výchozí v případě chyby

    def save_advance(self, employee_name, amount, currency, option, date):
        """
        Uloží zálohu do hlavního Excel souboru (Hodiny_Cap.xlsx).
        """
        try:
            # Použití context manageru pro hlavní workbook
            with self._get_workbook(self.file_path) as wb1:
                # Uložení do Hodiny_Cap.xlsx
                self._save_advance_main(wb1, employee_name, amount, currency, option, date)
                # Workbook se uloží automaticky context managerem

            logger.info(f"Záloha pro {employee_name} ({amount} {currency}, {option}, {date}) úspěšně uložena do {self.file_path}")
            return True

        except Exception as e:
            logger.error(f"Chyba při ukládání zálohy do {self.file_path}: {str(e)}", exc_info=True)
            return False

    def _save_advance_main(self, workbook, employee_name, amount, currency, option, date):
        """Pomocná metoda pro ukládání zálohy do listu 'Zálohy' daného workbooku"""
        if "Zálohy" not in workbook.sheetnames:
            sheet = workbook.create_sheet("Zálohy")
            # Nastavíme výchozí názvy možností, pokud list vytváříme
            sheet["B80"] = "Option 1"
            sheet["D80"] = "Option 2"
            logger.info("Vytvořen list 'Zálohy' s výchozími názvy možností.")
        else:
            sheet = workbook["Zálohy"]

        # Získáme aktuální názvy možností z listu
        option1_value = sheet["B80"].value or "Option 1"
        option2_value = sheet["D80"].value or "Option 2"

        # Najdeme řádek pro zaměstnance nebo vytvoříme nový
        row = 9 # Startovní řádek pro zaměstnance
        found_row = None
        # Projdeme existující řádky
        for r in range(row, sheet.max_row + 1):
             if sheet.cell(row=r, column=1).value == employee_name:
                  found_row = r
                  break
        # Pokud nebyl nalezen, najdeme první prázdný řádek od start_row
        if found_row is None:
             current_row = row
             while sheet.cell(row=current_row, column=1).value is not None:
                  current_row += 1
             found_row = current_row
             sheet.cell(row=found_row, column=1, value=employee_name) # Zapíšeme jméno nového zaměstnance


        # Určíme sloupec podle možnosti a měny
        if option == option1_value:
            column = 2 if currency == "EUR" else 3 # B nebo C
        elif option == option2_value:
            column = 4 if currency == "EUR" else 5 # D nebo E
        else:
            # Pokud se option neshoduje, zalogujeme chybu a použijeme výchozí mapování
            logger.error(f"Neznámá možnost zálohy '{option}'. Používají se sloupce pro '{option1_value}' (B/C).")
            # Můžeme zde vyvolat i ValueError, pokud chceme být striktnější
            # raise ValueError(f"Neplatná volba zálohy: {option}. Dostupné: {option1_value}, {option2_value}")
            column = 2 if currency == "EUR" else 3 # Fallback na Option 1


        # Aktualizujeme hodnotu zálohy
        current_value_cell = sheet.cell(row=found_row, column=column)
        current_value = current_value_cell.value or 0
        # Zajistíme, že pracujeme s čísly
        try:
             current_value_float = float(current_value)
             amount_float = float(amount)
             new_value = current_value_float + amount_float
        except (ValueError, TypeError):
             logger.error(f"Neplatná hodnota v buňce {current_value_cell.coordinate} nebo neplatná částka {amount}. Záloha nebude přičtena.")
             # Můžeme zde vyvolat chybu nebo pokračovat bez přičtení
             return # Nepokračujeme v ukládání této zálohy


        current_value_cell.value = new_value

        # Přidání data zálohy do sloupce Z (index 25)
        date_column_index = 25 # Sloupec Z
        date_cell = sheet.cell(row=found_row, column=date_column_index + 1) # Indexy jsou od 1
        try:
             date_obj = datetime.strptime(date, "%Y-%m-%d").date()
             date_cell.value = date_obj
             date_cell.number_format = 'DD.MM.YYYY' # Nastavení formátu data
        except ValueError:
             logger.error(f"Neplatný formát data '{date}'. Datum zálohy nebude uloženo.")
             date_cell.value = None # Vymažeme případnou neplatnou hodnotu


    # Metody _save_advance_zalohy25 a _save_advance_cash25 byly odstraněny

    def _get_next_empty_row_in_column(self, sheet, col):
        """Najde další prázdný řádek v daném sloupci, začíná od řádku 2."""
        row = 2 # Předpokládáme, že řádek 1 je hlavička
        while sheet.cell(row=row, column=col).value is not None:
            row += 1
        return row

    def ziskej_cislo_tydne(self, datum):
        """
        Získá ISO kalendářní data (rok, číslo týdne, den v týdnu) pro zadané datum.

        Args:
            datum: Datum jako string ('YYYY-MM-DD') nebo datetime objekt

        Returns:
            datetime.IsoCalendarDate: Objekt obsahující rok, číslo týdne a den v týdnu.
                                      Přístup k číslu týdne: result.week
        """
        try:
            if isinstance(datum, str):
                # Přísnější validace formátu
                datum_obj = datetime.strptime(datum, "%Y-%m-%d")
            elif isinstance(datum, datetime):
                 datum_obj = datum
            else:
                 raise TypeError("Datum musí být string ve formátu YYYY-MM-DD nebo datetime objekt")

            return datum_obj.isocalendar()
        except (ValueError, TypeError) as e:
            logger.error(f"Chyba při zpracování data '{datum}': {e}")
            # V případě chyby vrátíme kalendářní data pro aktuální den
            current_date = datetime.now()
            logger.warning(f"Používají se kalendářní data pro aktuální datum: {current_date.strftime('%Y-%m-%d')}")
            return current_date.isocalendar()


# Blok if __name__ == "__main__": byl odstraněn, protože testoval odstraněnou funkcionalitu.
# Pokud potřebujete testovat stávající metody, vytvořte nový testovací blok.

