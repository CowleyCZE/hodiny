import os
from openpyxl import load_workbook, Workbook
import logging
from datetime import datetime
from utils.logger import setup_logger

logger = setup_logger('zalohy_manager')

class ZalohyManager:
    """
    Správce záloh pro zaměstnance.
    
    Attributes:
        excel_path (str): Cesta k Excel souborům
        VALID_CURRENCIES (list): Povolené měny
        VALID_OPTIONS (list): Povolené možnosti záloh
    """
    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.excel_cesta = os.path.join(self.excel_path, "Hodiny_Cap.xlsx")
        self.ZALOHY_SHEET_NAME = 'Zálohy'
        self.EMPLOYEE_START_ROW = 9
        self.VALID_CURRENCIES = ['EUR', 'CZK']
        self.VALID_OPTIONS = ['Option 1', 'Option 2']  # Změněno z 'option1', 'option2'
        os.makedirs(self.excel_path, exist_ok=True)

    def validate_amount(self, amount):
        """Validuje částku zálohy"""
        if not isinstance(amount, (int, float)):
            raise ValueError("Částka musí být číslo")
        if amount <= 0:
            raise ValueError("Částka musí být větší než 0")
        if amount > 1000000:  # Reasonably high limit
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
        if not isinstance(employee_name, str):
            raise ValueError("Jméno zaměstnance musí být textový řetězec")
        if not employee_name.strip():
            raise ValueError("Jméno zaměstnance nemůže být prázdné")
        if len(employee_name) > 100:
            raise ValueError("Jméno zaměstnance je příliš dlouhé")
        return True

    def validate_option(self, option):
        """Validuje možnost zálohy"""
        if not isinstance(option, str):
            raise ValueError("Možnost musí být textový řetězec")
        if option not in self.VALID_OPTIONS:
            raise ValueError(f"Neplatná možnost. Povolené možnosti jsou: {', '.join(self.VALID_OPTIONS)}")
        return True

    def validate_date(self, date_str):
        """Validuje formát data"""
        try:
            datetime.strptime(date_str, '%Y-%m-%d')
            return True
        except ValueError:
            raise ValueError("Neplatný formát data. Použijte formát YYYY-MM-DD")

    def nacti_nebo_vytvor_excel(self):
        try:
            os.makedirs(self.excel_path, exist_ok=True)
            
            if os.path.exists(self.excel_cesta):
                workbook = load_workbook(self.excel_cesta)
                logger.info(f"Načten existující Excel soubor: {self.excel_cesta}")
            else:
                workbook = Workbook()
                workbook.save(self.excel_cesta)
                logger.info(f"Vytvořen nový Excel soubor: {self.excel_cesta}")
            
            if self.ZALOHY_SHEET_NAME not in workbook.sheetnames:
                workbook.create_sheet(self.ZALOHY_SHEET_NAME)
                logger.info(f"Vytvořen nový list '{self.ZALOHY_SHEET_NAME}'")
            
            return workbook
        except Exception as e:
            logger.error(f"Chyba při načítání nebo vytváření Excel souboru: {e}")
            raise

    def get_employee_row(self, employee_name):
        workbook = self.nacti_nebo_vytvor_excel()
        sheet = workbook[self.ZALOHY_SHEET_NAME]
        for row in range(self.EMPLOYEE_START_ROW, sheet.max_row + 1):
            if sheet.cell(row=row, column=1).value == employee_name:
                return row
        return None

    def add_or_update_employee_advance(self, employee_name, amount, currency, option, date):
        """
        Přidá nebo aktualizuje zálohu zaměstnance.
        
        Args:
            employee_name (str): Jméno zaměstnance
            amount (float): Částka zálohy
            currency (str): Měna (EUR/CZK)
            option (str): Typ zálohy
            date (str): Datum ve formátu YYYY-MM-DD
            
        Returns:
            bool: True pokud byla záloha úspěšně uložena
        
        Raises:
            ValueError: Při neplatných vstupních hodnotách
        """
        try:
            # Validace všech vstupů
            self.validate_employee_name(employee_name)
            self.validate_amount(amount)
            self.validate_currency(currency)
            self.validate_option(option)
            self.validate_date(date)

            workbook = self.nacti_nebo_vytvor_excel()
            sheet = workbook[self.ZALOHY_SHEET_NAME]
            row = self.get_employee_row(employee_name)
            
            if row is None:
                row = self.get_next_empty_row(sheet)
                sheet.cell(row=row, column=1, value=employee_name)
            
            if option == 'Option 1':  # Změněno z 'option1'
                column = 2 if currency == 'EUR' else 3
            else:  # Option 2
                column = 4 if currency == 'EUR' else 5
            
            current_value = sheet.cell(row=row, column=column).value or 0
            new_value = current_value + amount

            # Dodatečná validace celkové částky
            if new_value > 1000000:
                raise ValueError("Celková částka záloh by překročila povolený limit")

            sheet.cell(row=row, column=column, value=new_value)
            
            # Přidání data zálohy
            date_column = 26
            sheet.cell(row=row, column=date_column, value=datetime.strptime(date, '%Y-%m-%d').date())
            
            workbook.save(self.excel_cesta)
            logger.info(f"Záloha pro {employee_name} aktualizována: {amount} {currency} ({option}) k datu {date}")
            return True

        except Exception as e:
            logger.error(f"Chyba při ukládání zálohy: {e}")
            raise
        finally:
            if 'workbook' in locals():
                workbook.close()

    def get_next_empty_row(self, sheet):
        for row in range(self.EMPLOYEE_START_ROW, sheet.max_row + 2):
            if sheet.cell(row=row, column=1).value is None:
                return row
        return sheet.max_row + 1

    def get_employee_advances(self, employee_name):
        workbook = None
        try:
            workbook = self.nacti_nebo_vytvor_excel()
            sheet = workbook[self.ZALOHY_SHEET_NAME]
            row = self.get_employee_row(employee_name)
            if row is None:
                return None
            advances = {
                'Option1_EUR': sheet.cell(row=row, column=2).value or 0,
                'Option1_CZK': sheet.cell(row=row, column=3).value or 0,
                'Option2_EUR': sheet.cell(row=row, column=4).value or 0,
                'Option2_CZK': sheet.cell(row=row, column=5).value or 0
            }
            return advances
        finally:
            if workbook:
                workbook.close()

    def get_option_names(self):
        workbook = self.nacti_nebo_vytvor_excel()
        sheet = workbook[self.ZALOHY_SHEET_NAME]
        option1_name = sheet['B80'].value or 'Option 1'
        option2_name = sheet['D80'].value or 'Option 2'
        return option1_name, option2_name

if __name__ == "__main__":
    # Test code
    excel_path = "/home/Cowley/excel"  # Nastavení správné cesty k Excel souborům
    manager = ZalohyManager(excel_path)
    manager.add_or_update_employee_advance("Jan Novák", 100, 'EUR', 'Option 1', '2023-05-01')
    manager.add_or_update_employee_advance("Jan Novák", 2000, 'CZK', 'Option 2', '2023-05-02')
    print(manager.get_employee_advances("Jan Novák"))
    print(manager.get_option_names())