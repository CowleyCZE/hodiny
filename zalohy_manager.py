# zalohy_manager.py
from config import Config
import logging
import os
from datetime import datetime
from pathlib import Path

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
            raise ValueError("Chybí název aktivního souboru pro ZalohyManager.")

        self.base_path = Path(base_path)
        self.active_filename = active_filename
        self.active_file_path = self.base_path / self.active_filename

        self.ZALOHY_SHEET_NAME = Config.EXCEL_ADVANCES_SHEET_NAME
        self.EMPLOYEE_START_ROW = Config.EXCEL_EMPLOYEE_START_ROW
        self.VALID_CURRENCIES = ["EUR", "CZK"]
        self.DATE_COLUMN_INDEX = 26  # Sloupec Z pro datum

        self.base_path.mkdir(parents=True, exist_ok=True)
        logger.info(f"ZalohyManager inicializován pro soubor: {self.active_file_path}")

    def _get_active_workbook(self, read_only=False):
        """Načte aktivní workbook (Excel soubor)."""
        if not self.active_file_path.exists():
            logger.error(f"Aktivní soubor '{self.active_filename}' nebyl nalezen na cestě '{self.active_file_path}'.")
            raise FileNotFoundError(f"Aktivní soubor '{self.active_filename}' neexistuje na cestě '{self.active_file_path}'.")
        try:
            logger.debug(f"Načítání workbooku: {self.active_file_path} (read_only={read_only})")
            return load_workbook(filename=self.active_file_path, read_only=read_only, data_only=True)
        except Exception as e:
            logger.error(f"Chyba při načítání workbooku '{self.active_filename}': {e}", exc_info=True)
            raise IOError(f"Nepodařilo se otevřít soubor '{self.active_filename}'.")

    def _save_workbook(self, workbook):
        """Uloží daný workbook (sešit) do aktivního souboru."""
        if workbook is None:
            logger.error("Pokus o uložení None workbooku. Operace byla přeskočena.")
            return
        try:
            logger.debug(f"Pokus o uložení změn do workbooku: {self.active_file_path}")
            workbook.save(self.active_file_path)
            logger.info(f"Workbook '{self.active_filename}' úspěšně uložen.")
        except Exception as e:
            logger.error(f"Chyba při ukládání workbooku '{self.active_filename}': {e}", exc_info=True)
            raise IOError(f"Nepodařilo se uložit změny do souboru '{self.active_filename}'.")

    def validate_amount(self, amount):
        if not isinstance(amount, (int, float)) or amount <= 0:
            raise ValueError("Částka musí být kladné číslo.")
        return True

    def validate_currency(self, currency):
        if currency not in self.VALID_CURRENCIES:
            raise ValueError(f"Neplatná měna. Povolené měny jsou: {', '.join(self.VALID_CURRENCIES)}")
        return True

    def validate_employee_name(self, employee_name):
        if not isinstance(employee_name, str) or not employee_name.strip():
            raise ValueError("Jméno zaměstnance nemůže být prázdné.")
        return True

    def validate_date(self, date_str):
        try:
            datetime.strptime(date_str, "%Y-%m-%d")
            return True
        except ValueError:
            raise ValueError("Neplatný formát data. Použijte formát YYYY-MM-DD.")

    def _ensure_zalohy_sheet(self, workbook):
        """Zajistí existenci listu 'Zálohy' a vrátí ho."""
        if self.ZALOHY_SHEET_NAME not in workbook.sheetnames:
            sheet = workbook.create_sheet(self.ZALOHY_SHEET_NAME)
            sheet["B80"] = Config.DEFAULT_ADVANCE_OPTION_1
            sheet["D80"] = Config.DEFAULT_ADVANCE_OPTION_2
            sheet["F80"] = Config.DEFAULT_ADVANCE_OPTION_3
            sheet["H80"] = Config.DEFAULT_ADVANCE_OPTION_4
            logger.info(f"Vytvořen list '{self.ZALOHY_SHEET_NAME}'.")
            return sheet
        return workbook[self.ZALOHY_SHEET_NAME]

    def _get_employee_row(self, sheet, employee_name):
        """Najde řádek pro daného zaměstnance v listu."""
        for row in range(self.EMPLOYEE_START_ROW, sheet.max_row + 1):
            if sheet.cell(row=row, column=1).value == employee_name:
                return row
        return None

    def _get_next_empty_row(self, sheet):
        """Najde první prázdný řádek pro nového zaměstnance."""
        row = self.EMPLOYEE_START_ROW
        while sheet.cell(row=row, column=1).value is not None:
            row += 1
        return row

    def add_or_update_employee_advance(self, employee_name, amount, currency, option, date):
        """Přidá nebo aktualizuje zálohu pro zaměstnance."""
        workbook = None
        try:
            self.validate_employee_name(employee_name)
            self.validate_amount(amount)
            self.validate_currency(currency)
            self.validate_date(date)

            workbook = self._get_active_workbook(read_only=False)
            sheet = self._ensure_zalohy_sheet(workbook)

            options = self.get_option_names()
            if option not in options:
                raise ValueError(f"Neplatná volba zálohy: {option}")

            row = self._get_employee_row(sheet, employee_name)
            if row is None:
                row = self._get_next_empty_row(sheet)
                sheet.cell(row=row, column=1, value=employee_name)
                logger.info(f"Přidán nový zaměstnanec '{employee_name}' na řádek {row}.")

            option_index = options.index(option)
            column_index = 2 + (option_index * 2) + (1 if currency == "CZK" else 0)

            target_cell = sheet.cell(row=row, column=column_index)
            current_value = float(target_cell.value or 0)
            new_value = current_value + float(amount)
            target_cell.value = new_value
            target_cell.number_format = '#,##0.00'

            date_cell = sheet.cell(row=row, column=self.DATE_COLUMN_INDEX)
            date_cell.value = datetime.strptime(date, "%Y-%m-%d").date()
            date_cell.number_format = 'DD.MM.YYYY'

            self._save_workbook(workbook)
            logger.info(f"Záloha pro {employee_name} ({amount} {currency}, {option}, {date}) úspěšně uložena.")
            return True

        except (FileNotFoundError, ValueError, IOError) as e:
            logger.error(f"Chyba při ukládání zálohy: {e}")
            raise e
        except Exception as e:
            logger.error(f"Neočekávaná chyba při ukládání zálohy: {e}", exc_info=True)
            raise RuntimeError(f"Neočekávaná chyba při ukládání zálohy: {e}")
        finally:
            if workbook:
                workbook.close()

    def get_option_names(self):
        """
        Získá názvy možností záloh z listu 'Zálohy'.
        Returns:
            tuple: Čtveřice stringů (option1, option2, option3, option4).
        """
        workbook = None
        default_options = (
            Config.DEFAULT_ADVANCE_OPTION_1,
            Config.DEFAULT_ADVANCE_OPTION_2,
            Config.DEFAULT_ADVANCE_OPTION_3,
            Config.DEFAULT_ADVANCE_OPTION_4
        )
        try:
            workbook = self._get_active_workbook(read_only=True)
            if self.ZALOHY_SHEET_NAME in workbook.sheetnames:
                sheet = workbook[self.ZALOHY_SHEET_NAME]
                option1 = sheet["B80"].value or default_options[0]
                option2 = sheet["D80"].value or default_options[1]
                option3 = sheet["F80"].value or default_options[2]
                option4 = sheet["H80"].value or default_options[3]
                return str(option1).strip(), str(option2).strip(), str(option3).strip(), str(option4).strip()
            else:
                logger.warning(f"List '{self.ZALOHY_SHEET_NAME}' nenalezen. Používají se výchozí názvy možností.")
                return default_options
        except (FileNotFoundError, IOError) as e:
            logger.error(f"Chyba souboru při čtení názvů možností: {e}")
            return default_options
        except Exception as e:
            logger.error(f"Neočekávaná chyba při čtení názvů možností: {e}", exc_info=True)
            return default_options
        finally:
            if workbook:
                workbook.close()
