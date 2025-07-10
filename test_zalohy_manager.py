import unittest
import tempfile
import os
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook, load_workbook

# Importujeme třídy, které budeme testovat
from zalohy_manager import ZalohyManager
from config import Config

class TestZalohyManager(unittest.TestCase):
    def setUp(self):
        # Vytvoříme dočasný adresář pro testovací soubory
        self.temp_dir = tempfile.TemporaryDirectory()
        self.excel_base_path = Path(self.temp_dir.name)
        self.active_filename = "test_zalohy.xlsx"
        self.active_file_path = self.excel_base_path / self.active_filename

        # Vytvoříme mock Excel soubor s listem "Zálohy" a výchozími hodnotami
        self.create_mock_excel_file()

        # Inicializujeme ZalohyManager
        self.zalohy_manager = ZalohyManager(
            base_path=str(self.excel_base_path),
            active_filename=self.active_filename
        )

    def tearDown(self):
        # Uklidíme dočasný adresář
        self.temp_dir.cleanup()

    def create_mock_excel_file(self):
        # Vytvoříme nový workbook
        wb = Workbook()
        # Vytvoříme list "Zálohy"
        ws = wb.create_sheet(Config.EXCEL_ADVANCES_SHEET_NAME)

        # Nastavíme výchozí názvy možností záloh v buňkách B80, D80, F80, H80
        ws["B80"] = Config.DEFAULT_ADVANCE_OPTION_1
        ws["D80"] = Config.DEFAULT_ADVANCE_OPTION_2
        ws["F80"] = Config.DEFAULT_ADVANCE_OPTION_3
        ws["H80"] = Config.DEFAULT_ADVANCE_OPTION_4

        # Uložíme soubor
        wb.save(self.active_file_path)
        wb.close()

    def _read_cell_value(self, sheet_name, cell_coordinate):
        # Pomocná metoda pro čtení hodnoty buňky z aktivního souboru
        wb = load_workbook(self.active_file_path, data_only=True)
        ws = wb[sheet_name]
        value = ws[cell_coordinate].value
        wb.close()
        return value

    def test_ensure_zalohy_sheet_creation(self):
        # Smažeme existující soubor a vytvoříme nový, abychom otestovali vytvoření listu
        os.remove(self.active_file_path)
        # Vytvoříme prázdný workbook, aby _ensure_zalohy_sheet mohl vytvořit list
        wb = Workbook()
        wb.save(self.active_file_path)
        wb.close()

        # Získáme workbook a zavoláme _ensure_zalohy_sheet
        wb_reloaded = load_workbook(self.active_file_path)
        sheet = self.zalohy_manager._ensure_zalohy_sheet(wb_reloaded)
        
        self.assertIn(Config.EXCEL_ADVANCES_SHEET_NAME, wb_reloaded.sheetnames)
        self.assertEqual(sheet["B80"].value, Config.DEFAULT_ADVANCE_OPTION_1)
        self.assertEqual(sheet["D80"].value, Config.DEFAULT_ADVANCE_OPTION_2)
        self.assertEqual(sheet["F80"].value, Config.DEFAULT_ADVANCE_OPTION_3)
        self.assertEqual(sheet["H80"].value, Config.DEFAULT_ADVANCE_OPTION_4)
        wb_reloaded.close()

    def test_get_option_names(self):
        option1, option2, option3, option4 = self.zalohy_manager.get_option_names()
        self.assertEqual(option1, Config.DEFAULT_ADVANCE_OPTION_1)
        self.assertEqual(option2, Config.DEFAULT_ADVANCE_OPTION_2)
        self.assertEqual(option3, Config.DEFAULT_ADVANCE_OPTION_3)
        self.assertEqual(option4, Config.DEFAULT_ADVANCE_OPTION_4)

    def test_add_or_update_employee_advance_new_employee(self):
        employee_name = "Nový Zaměstnanec"
        amount = 100.0
        currency = "EUR"
        option = Config.DEFAULT_ADVANCE_OPTION_1
        date = "2025-07-10"

        success = self.zalohy_manager.add_or_update_employee_advance(
            employee_name, amount, currency, option, date
        )
        self.assertTrue(success)

        # Ověříme, že se zaměstnanec přidal a částka zapsala
        wb = load_workbook(self.active_file_path, data_only=True)
        ws = wb[Config.EXCEL_ADVANCES_SHEET_NAME]
        
        # Najdeme řádek nového zaměstnance
        row = None
        for r in range(Config.EXCEL_EMPLOYEE_START_ROW, ws.max_row + 1):
            if ws.cell(row=r, column=1).value == employee_name:
                row = r
                break
        self.assertIsNotNone(row)
        self.assertEqual(ws.cell(row=row, column=2).value, amount) # EUR pro Option 1 je sloupec B (index 2)
        self.assertEqual(ws.cell(row=row, column=26).value.date(), datetime.strptime(date, "%Y-%m-%d").date()) # Datum v sloupci Z (index 26)
        wb.close()

    def test_add_or_update_employee_advance_option1_eur_czk(self):
        employee_name = "Test Zaměstnanec 1"
        date = "2025-07-10"

        # EUR
        self.zalohy_manager.add_or_update_employee_advance(employee_name, 50.0, "EUR", Config.DEFAULT_ADVANCE_OPTION_1, date)
        self.assertEqual(self._read_cell_value(Config.EXCEL_ADVANCES_SHEET_NAME, "B9"), 50.0) # Předpokládáme řádek 9 pro prvního zaměstnance

        # CZK
        self.zalohy_manager.add_or_update_employee_advance(employee_name, 1000.0, "CZK", Config.DEFAULT_ADVANCE_OPTION_1, date)
        self.assertEqual(self._read_cell_value(Config.EXCEL_ADVANCES_SHEET_NAME, "C9"), 1000.0) # CZK pro Option 1 je sloupec C (index 3)

        # Přičtení k existující hodnotě
        self.zalohy_manager.add_or_update_employee_advance(employee_name, 25.0, "EUR", Config.DEFAULT_ADVANCE_OPTION_1, date)
        self.assertEqual(self._read_cell_value(Config.EXCEL_ADVANCES_SHEET_NAME, "B9"), 75.0)

    def test_add_or_update_employee_advance_option2_eur_czk(self):
        employee_name = "Test Zaměstnanec 2"
        date = "2025-07-11"

        # EUR
        self.zalohy_manager.add_or_update_employee_advance(employee_name, 75.0, "EUR", Config.DEFAULT_ADVANCE_OPTION_2, date)
        self.assertEqual(self._read_cell_value(Config.EXCEL_ADVANCES_SHEET_NAME, "D9"), 75.0) # EUR pro Option 2 je sloupec D (index 4)

        # CZK
        self.zalohy_manager.add_or_update_employee_advance(employee_name, 2000.0, "CZK", Config.DEFAULT_ADVANCE_OPTION_2, date)
        self.assertEqual(self._read_cell_value(Config.EXCEL_ADVANCES_SHEET_NAME, "E9"), 2000.0) # CZK pro Option 2 je sloupec E (index 5)

    def test_add_or_update_employee_advance_option3_eur_czk(self):
        employee_name = "Test Zaměstnanec 3"
        date = "2025-07-12"

        # EUR
        self.zalohy_manager.add_or_update_employee_advance(employee_name, 120.0, "EUR", Config.DEFAULT_ADVANCE_OPTION_3, date)
        self.assertEqual(self._read_cell_value(Config.EXCEL_ADVANCES_SHEET_NAME, "F9"), 120.0) # EUR pro Option 3 je sloupec F (index 6)

        # CZK
        self.zalohy_manager.add_or_update_employee_advance(employee_name, 3000.0, "CZK", Config.DEFAULT_ADVANCE_OPTION_3, date)
        self.assertEqual(self._read_cell_value(Config.EXCEL_ADVANCES_SHEET_NAME, "G9"), 3000.0) # CZK pro Option 3 je sloupec G (index 7)

    def test_add_or_update_employee_advance_option4_eur_czk(self):
        employee_name = "Test Zaměstnanec 4"
        date = "2025-07-13"

        # EUR
        self.zalohy_manager.add_or_update_employee_advance(employee_name, 200.0, "EUR", Config.DEFAULT_ADVANCE_OPTION_4, date)
        self.assertEqual(self._read_cell_value(Config.EXCEL_ADVANCES_SHEET_NAME, "H9"), 200.0) # EUR pro Option 4 je sloupec H (index 8)

        # CZK
        self.zalohy_manager.add_or_update_employee_advance(employee_name, 4000.0, "CZK", Config.DEFAULT_ADVANCE_OPTION_4, date)
        self.assertEqual(self._read_cell_value(Config.EXCEL_ADVANCES_SHEET_NAME, "I9"), 4000.0) # CZK pro Option 4 je sloupec I (index 9)

    def test_add_or_update_employee_advance_invalid_amount(self):
        with self.assertRaises(ValueError):
            self.zalohy_manager.add_or_update_employee_advance("Test", -10.0, "EUR", Config.DEFAULT_ADVANCE_OPTION_1, "2025-07-10")
        with self.assertRaises(ValueError):
            self.zalohy_manager.add_or_update_employee_advance("Test", "abc", "EUR", Config.DEFAULT_ADVANCE_OPTION_1, "2025-07-10")

    def test_add_or_update_employee_advance_invalid_currency(self):
        with self.assertRaises(ValueError):
            self.zalohy_manager.add_or_update_employee_advance("Test", 100.0, "USD", Config.DEFAULT_ADVANCE_OPTION_1, "2025-07-10")

    def test_add_or_update_employee_advance_invalid_option(self):
        # Tato validace se děje v app.py, ale pro jistotu testujeme i zde, pokud by se logika změnila
        # V ZalohyManageru je fallback, takže to nevyhodí chybu, ale použije defaultní sloupec
        employee_name = "Test Invalid Option"
        amount = 50.0
        currency = "EUR"
        invalid_option = "Neplatná Možnost"
        date = "2025-07-10"

        success = self.zalohy_manager.add_or_update_employee_advance(
            employee_name, amount, currency, invalid_option, date
        )
        self.assertTrue(success) # Mělo by být True, protože je tam fallback

        # Ověříme, že se zapsalo do defaultního sloupce (B9)
        wb = load_workbook(self.active_file_path, data_only=True)
        ws = wb[Config.EXCEL_ADVANCES_SHEET_NAME]
        row = None
        for r in range(Config.EXCEL_EMPLOYEE_START_ROW, ws.max_row + 1):
            if ws.cell(row=r, column=1).value == employee_name:
                row = r
                break
        self.assertIsNotNone(row)
        self.assertEqual(ws.cell(row=row, column=2).value, amount) # Mělo by se zapsat do B (index 2)
        wb.close()

    def test_add_or_update_employee_advance_invalid_date(self):
        with self.assertRaises(ValueError):
            self.zalohy_manager.add_or_update_employee_advance("Test", 100.0, "EUR", Config.DEFAULT_ADVANCE_OPTION_1, "2025/07/10")
        with self.assertRaises(ValueError):
            self.zalohy_manager.add_or_update_employee_advance("Test", 100.0, "EUR", Config.DEFAULT_ADVANCE_OPTION_1, "neplatné datum")

if __name__ == '__main__':
    unittest.main()
