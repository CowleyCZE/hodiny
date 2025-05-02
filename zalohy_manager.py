# zalohy_manager.py
import logging
import os
from datetime import datetime
from pathlib import Path # Import Path

from openpyxl import Workbook, load_workbook

# Předpokládá existenci utils.logger
try:
    from utils.logger import setup_logger
    logger = setup_logger("zalohy_manager")
except ImportError:
    # Fallback na základní logger, pokud utils.logger není dostupný
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger("zalohy_manager")


class ZalohyManager:
    """
    Správce záloh pro zaměstnance v aktivním Excel souboru.
    """
    def __init__(self, base_path, active_filename):
        """
        Inicializuje ZalohyManager.

        Args:
            base_path (Path): Cesta k adresáři s Excel soubory.
            active_filename (str): Název aktuálně používaného souboru.
        """
        if not active_filename:
            logger.error("Nelze inicializovat ZalohyManager bez názvu aktivního souboru.")
            # Můžeme vyvolat chybu nebo nastavit cestu na None a kontrolovat později
            raise ValueError("Chybí název aktivního souboru pro ZalohyManager.")

        self.base_path = Path(base_path)
        self.active_filename = active_filename
        # Sestavení plné cesty k aktivnímu souboru
        self.active_file_path = self.base_path / self.active_filename

        self.ZALOHY_SHEET_NAME = "Zálohy"
        self.EMPLOYEE_START_ROW = 9 # Řádek, kde začínají jména
        self.VALID_CURRENCIES = ["EUR", "CZK"]
        # VALID_OPTIONS se nyní načítají dynamicky z Excelu, ale můžeme je nechat pro případnou validaci
        # self.VALID_OPTIONS = ["Option 1", "Option 2"]

        # Zajistíme existenci adresáře (i když by měl z init_app)
        self.base_path.mkdir(parents=True, exist_ok=True)
        logger.info(f"ZalohyManager inicializován pro soubor: {self.active_file_path}")


    def _get_active_workbook(self, read_only=False):
        """
        Načte aktivní workbook. Vrací workbook objekt.
        Je zodpovědností volajícího workbook uzavřít.
        Používá se interně, kde nepotřebujeme složitou cache jako v ExcelManageru.
        """
        if not self.active_file_path.exists():
             logger.error(f"Aktivní soubor '{self.active_filename}' nebyl nalezen.")
             raise FileNotFoundError(f"Aktivní soubor '{self.active_filename}' neexistuje.")
        try:
            # data_only=True načte hodnoty místo vzorců
            return load_workbook(filename=self.active_file_path, read_only=read_only, data_only=True)
        except Exception as e:
            logger.error(f"Chyba při načítání workbooku '{self.active_filename}': {e}", exc_info=True)
            raise IOError(f"Nepodařilo se otevřít soubor '{self.active_filename}'.")

    def _save_workbook(self, workbook):
         """Uloží workbook do aktivního souboru."""
         try:
              workbook.save(self.active_file_path)
              logger.info(f"Workbook '{self.active_filename}' úspěšně uložen.")
         except Exception as e:
              logger.error(f"Chyba při ukládání workbooku '{self.active_filename}': {e}", exc_info=True)
              raise IOError(f"Nepodařilo se uložit změny do souboru '{self.active_filename}'.")


    # --- Validační metody (zůstávají stejné) ---
    def validate_amount(self, amount):
        """Validuje částku zálohy"""
        if not isinstance(amount, (int, float)):
            raise ValueError("Částka musí být číslo")
        if amount <= 0:
            raise ValueError("Částka musí být větší než 0")
        if amount > 1000000:  # Limit
            raise ValueError("Částka je příliš vysoká")
        return True

    def validate_currency(self, currency):
        """Validuje měnu"""
        if not isinstance(currency, str):
            raise ValueError("Měna musí být textový řetězec")
        if currency not in self.VALID_CURRENCIES:
            raise ValueError(f"Neplatná měna. Povolené měny jsou: {', '.join(self.VALID_CURRENCIES)}")
        return True

    def validate_employee_name(self, employee_name):
        """Validuje jméno zaměstnance"""
        if not isinstance(employee_name, str) or not employee_name.strip():
            raise ValueError("Jméno zaměstnance nemůže být prázdné")
        if len(employee_name) > 100:
            raise ValueError("Jméno zaměstnance je příliš dlouhé")
        # Můžeme přidat kontrolu na znaky, pokud je potřeba
        return True

    # Validace option se nyní dělá v app.py proti dynamicky načteným možnostem
    # def validate_option(self, option): ...

    def validate_date(self, date_str):
        """Validuje formát data"""
        try:
            datetime.strptime(date_str, "%Y-%m-%d")
            return True
        except ValueError:
            raise ValueError("Neplatný formát data. Použijte formát YYYY-MM-DD")

    # --- Metody pro práci se zálohami ---

    def _ensure_zalohy_sheet(self, workbook):
         """Zajistí existenci listu 'Zálohy' a vrátí ho."""
         if self.ZALOHY_SHEET_NAME not in workbook.sheetnames:
             sheet = workbook.create_sheet(self.ZALOHY_SHEET_NAME)
             # Nastavíme výchozí názvy možností
             sheet["B80"] = "Option 1"
             sheet["D80"] = "Option 2"
             logger.info(f"Vytvořen list '{self.ZALOHY_SHEET_NAME}' v souboru {self.active_filename}.")
             return sheet
         else:
             return workbook[self.ZALOHY_SHEET_NAME]

    def _get_employee_row(self, sheet, employee_name):
        """Najde řádek pro daného zaměstnance v listu."""
        # Projdeme řádky od EMPLOYEE_START_ROW
        for row in range(self.EMPLOYEE_START_ROW, sheet.max_row + 1):
            if sheet.cell(row=row, column=1).value == employee_name:
                return row
        return None # Zaměstnanec nenalezen

    def _get_next_empty_row(self, sheet):
         """Najde první prázdný řádek pro nového zaměstnance."""
         row = self.EMPLOYEE_START_ROW
         while sheet.cell(row=row, column=1).value is not None:
              row += 1
         return row


    def add_or_update_employee_advance(self, employee_name, amount, currency, option, date):
        """
        Přidá nebo aktualizuje zálohu zaměstnance v aktivním souboru.
        """
        workbook = None
        try:
            # Validace vstupů (volitelně zde, nebo spoléhat na validaci v app.py)
            self.validate_employee_name(employee_name)
            self.validate_amount(amount)
            self.validate_currency(currency)
            # self.validate_option(option) # Validace proti dynamickým možnostem je v app.py
            self.validate_date(date)

            # Načteme workbook pro zápis
            workbook = self._get_active_workbook(read_only=False)
            sheet = self._ensure_zalohy_sheet(workbook)

            # Získáme aktuální názvy možností z listu
            option1_value = sheet["B80"].value or "Option 1"
            option2_value = sheet["D80"].value or "Option 2"

            # Najdeme řádek zaměstnance nebo vytvoříme nový
            row = self._get_employee_row(sheet, employee_name)
            if row is None:
                row = self._get_next_empty_row(sheet)
                sheet.cell(row=row, column=1, value=employee_name)
                logger.info(f"Přidán nový zaměstnanec '{employee_name}' do listu '{self.ZALOHY_SHEET_NAME}' na řádek {row}.")

            # Určíme sloupec podle možnosti a měny (indexy od 1)
            if option == option1_value:
                column_index = 2 if currency == "EUR" else 3 # B nebo C
            elif option == option2_value:
                column_index = 4 if currency == "EUR" else 5 # D nebo E
            else:
                logger.error(f"Neznámá možnost zálohy '{option}' při ukládání. Používají se sloupce pro '{option1_value}'.")
                # Můžeme vyvolat chybu nebo použít fallback
                # raise ValueError(f"Neplatná volba zálohy: {option}")
                column_index = 2 if currency == "EUR" else 3 # Fallback

            # Aktualizujeme hodnotu zálohy
            target_cell = sheet.cell(row=row, column=column_index)
            current_value = target_cell.value or 0
            try:
                current_value_float = float(current_value)
                amount_float = float(amount)
                new_value = current_value_float + amount_float
                target_cell.value = new_value
                target_cell.number_format = '#,##0.00' # Nastavení formátu
            except (ValueError, TypeError):
                logger.error(f"Neplatná číselná hodnota v buňce {target_cell.coordinate} ('{current_value}') nebo částka '{amount}'. Záloha nebude přičtena.")
                raise ValueError("Neplatná číselná hodnota pro zálohu.")

            # Přidání data zálohy do sloupce Z (index 25 -> sloupec 26)
            date_column_index = 26
            date_cell = sheet.cell(row=row, column=date_column_index)
            try:
                date_obj = datetime.strptime(date, "%Y-%m-%d").date()
                date_cell.value = date_obj
                date_cell.number_format = 'DD.MM.YYYY'
            except ValueError:
                logger.error(f"Neplatný formát data '{date}' pro zálohu. Datum nebude uloženo.")
                date_cell.value = None

            # Uložíme změny
            self._save_workbook(workbook)
            logger.info(f"Záloha pro {employee_name} ({amount} {currency}, {option}, {date}) úspěšně uložena do {self.active_filename}")
            return True

        except (FileNotFoundError, ValueError, IOError) as e:
            logger.error(f"Chyba při ukládání zálohy do {self.active_filename}: {e}", exc_info=False) # False pro stručnější log u validací
            # Propagujeme chybu výše, aby ji mohlo zachytit app.py
            raise e
        except Exception as e:
            logger.error(f"Neočekávaná chyba při ukládání zálohy do {self.active_filename}: {e}", exc_info=True)
            # Propagujeme jako obecnou chybu
            raise RuntimeError(f"Neočekávaná chyba při ukládání zálohy: {e}")

        finally:
            # Zajistíme uzavření workbooku, pokud byl otevřen
            if workbook:
                try:
                    workbook.close()
                except Exception as close_err:
                    logger.warning(f"Chyba při zavírání workbooku v add_or_update_employee_advance: {close_err}")

    # Metody get_employee_advances a get_option_names mohou být upraveny, pokud je potřeba číst data
    # Zde je příklad pro get_option_names (get_employee_advances by bylo podobné)
    def get_option_names(self):
         """Získá názvy možností záloh z aktivního souboru."""
         workbook = None
         try:
              workbook = self._get_active_workbook(read_only=True)
              if self.ZALOHY_SHEET_NAME in workbook.sheetnames:
                   sheet = workbook[self.ZALOHY_SHEET_NAME]
                   option1 = sheet["B80"].value or "Option 1"
                   option2 = sheet["D80"].value or "Option 2"
                   return str(option1).strip(), str(option2).strip()
              else:
                   logger.warning(f"List '{self.ZALOHY_SHEET_NAME}' nenalezen v {self.active_filename} při čtení názvů možností.")
                   return "Option 1", "Option 2" # Výchozí
         except (FileNotFoundError, IOError) as e:
              logger.error(f"Chyba při čtení názvů možností z {self.active_filename}: {e}")
              return "Option 1", "Option 2"
         except Exception as e:
              logger.error(f"Neočekávaná chyba při čtení názvů možností: {e}", exc_info=True)
              return "Option 1", "Option 2"
         finally:
              if workbook:
                   try: workbook.close()
                   except: pass


# Blok if __name__ == "__main__": není relevantní pro tuto třídu v kontextu aplikace
