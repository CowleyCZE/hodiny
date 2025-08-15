"""Správa listu záloh (Zálohy) v hlavním Excel souboru.

Funkce:
 - validace vstupů (částka, měna, datum, jméno)
 - kumulativní přičítání částek do správného měnového sloupce dle vybrané "option"
 - zápis data poslední transakce do dedikovaného sloupce
"""
import logging
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

from config import Config

try:
    from utils.logger import setup_logger

    logger = setup_logger("zalohy_manager")
except ImportError:
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger("zalohy_manager")


class ZalohyManager:
    def __init__(self, base_path):
        """base_path: adresář s hlavním souborem Hodiny_Cap.xlsx."""
        self.base_path = Path(base_path)
        self.active_filename = Config.EXCEL_TEMPLATE_NAME
        self.active_file_path = self.base_path / self.active_filename
        self.ZALOHY_SHEET_NAME = Config.EXCEL_ADVANCES_SHEET_NAME
        self.EMPLOYEE_START_ROW = Config.EXCEL_EMPLOYEE_START_ROW
        self.VALID_CURRENCIES = ["EUR", "CZK"]
        self.DATE_COLUMN_INDEX = 26
        logger.info(f"ZalohyManager inicializován pro soubor: {self.active_file_path}")

    def _get_active_workbook(self, read_only=False):
        if not self.active_file_path.exists():
            raise FileNotFoundError(f"Soubor '{self.active_filename}' neexistuje.")
        try:
            return load_workbook(filename=self.active_file_path, read_only=read_only, data_only=True)
        except Exception:
            raise IOError(f"Nepodařilo se otevřít soubor '{self.active_filename}'.")

    def _save_workbook(self, workbook):
        try:
            workbook.save(self.active_file_path)
        except Exception:
            raise IOError(f"Nepodařilo se uložit změny do souboru '{self.active_filename}'.")

    def validate_amount(self, amount):
        """Částka musí být kladná (int/float)."""
        if not isinstance(amount, (int, float)) or amount <= 0:
            raise ValueError("Částka musí být kladné číslo.")
        return True

    def validate_currency(self, currency):
        """Kontrola proti whitelistu měn."""
        if currency not in self.VALID_CURRENCIES:
            raise ValueError("Neplatná měna.")
        return True

    def validate_employee_name(self, employee_name):
        """Nesmí být prázdné; další validace (regex) lze doplnit později."""
        if not isinstance(employee_name, str) or not employee_name.strip():
            raise ValueError("Jméno zaměstnance nemůže být prázdné.")
        return True

    def validate_date(self, date_str):
        """YYYY-MM-DD datum; ValueError při neplatném formátu."""
        try:
            datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            raise ValueError("Neplatný formát data.")
        return True

    def add_or_update_employee_advance(self, employee_name, amount, currency, option, date):
        """Přičte zálohu zaměstnanci (vytvoří řádek pokud chybí)."""
        workbook = None
        try:
            self.validate_employee_name(employee_name)
            self.validate_amount(amount)
            self.validate_currency(currency)
            self.validate_date(date)

            workbook = self._get_active_workbook(read_only=False)
            sheet = workbook[self.ZALOHY_SHEET_NAME]

            options = self.get_option_names()
            if option not in options:
                raise ValueError(f"Neplatná volba zálohy: {option}")

            row = self._get_employee_row(sheet, employee_name)
            if row is None:
                row = self._get_next_empty_row(sheet)
                sheet.cell(row=row, column=1, value=employee_name)

            option_index = options.index(option)
            column_index = 2 + (option_index * 2) + (1 if currency == "CZK" else 0)

            target_cell = sheet.cell(row=row, column=column_index)
            # Bezpečný převod existující hodnoty (ignoruje nečíselné / vzorce)
            if not isinstance(target_cell, MergedCell):
                raw_val = target_cell.value
                try:
                    if isinstance(raw_val, (int, float, str)) and str(raw_val).strip():
                        current_value = float(raw_val)
                    else:
                        current_value = 0.0
                except (ValueError, TypeError):
                    current_value = 0.0
                target_cell.value = current_value + float(amount)
                target_cell.number_format = "#,##0.00"

            date_cell = sheet.cell(row=row, column=self.DATE_COLUMN_INDEX)
            if not isinstance(date_cell, MergedCell):
                date_cell.value = datetime.strptime(date, "%Y-%m-%d").date()
                date_cell.number_format = "DD.MM.YYYY"

            self._save_workbook(workbook)
            return True
        except (FileNotFoundError, ValueError, IOError, Exception) as e:
            logger.error(f"Chyba při ukládání zálohy: {e}", exc_info=True)
            raise e
        finally:
            if workbook:
                workbook.close()

    def _get_employee_row(self, sheet, employee_name):
        """Najde existující řádek zaměstnance nebo vrátí None."""
        for row in range(self.EMPLOYEE_START_ROW, sheet.max_row + 1):
            if sheet.cell(row=row, column=1).value == employee_name:
                return row
        return None

    def _get_next_empty_row(self, sheet):
        """První volný řádek od konfigurované start row."""
        row = self.EMPLOYEE_START_ROW
        while sheet.cell(row=row, column=1).value is not None:
            row += 1
        return row

    def get_option_names(self):
        """Názvy 4 možností čtené z buněk (fallback na default)."""
        default_options = (
            Config.DEFAULT_ADVANCE_OPTION_1,
            Config.DEFAULT_ADVANCE_OPTION_2,
            Config.DEFAULT_ADVANCE_OPTION_3,
            Config.DEFAULT_ADVANCE_OPTION_4,
        )
        workbook = None
        try:
            workbook = self._get_active_workbook(read_only=True)
            if self.ZALOHY_SHEET_NAME in workbook.sheetnames:
                sheet = workbook[self.ZALOHY_SHEET_NAME]
                return tuple(
                    str(sheet[cell].value or default).strip()
                    for cell, default in zip(["B80", "D80", "F80", "H80"], default_options)
                )
            return default_options
        except (FileNotFoundError, IOError, Exception) as e:
            logger.error(f"Chyba při čtení názvů možností: {e}", exc_info=True)
            return default_options
        finally:
            if workbook:
                workbook.close()
