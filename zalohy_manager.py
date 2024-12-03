import os
from openpyxl import load_workbook, Workbook
import logging
from datetime import datetime

class ZalohyManager:
    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.excel_cesta = os.path.join(self.excel_path, "Hodiny_Cap.xlsx")
        self.ZALOHY_SHEET_NAME = 'Zálohy'
        self.EMPLOYEE_START_ROW = 9
        os.makedirs(self.excel_path, exist_ok=True)

    def nacti_nebo_vytvor_excel(self):
        try:
            if os.path.exists(self.excel_cesta):
                workbook = load_workbook(self.excel_cesta)
                logging.info(f"Načten existující Excel soubor: {self.excel_cesta}")
            else:
                workbook = Workbook()
                workbook.save(self.excel_cesta)
                logging.info(f"Vytvořen nový Excel soubor: {self.excel_cesta}")
            
            if self.ZALOHY_SHEET_NAME not in workbook.sheetnames:
                workbook.create_sheet(self.ZALOHY_SHEET_NAME)
                logging.info(f"Vytvořen nový list '{self.ZALOHY_SHEET_NAME}'")
            
            return workbook
        except Exception as e:
            logging.error(f"Chyba při načítání nebo vytváření Excel souboru: {e}")
            raise

    def get_employee_row(self, employee_name):
        workbook = self.nacti_nebo_vytvor_excel()
        sheet = workbook[self.ZALOHY_SHEET_NAME]
        for row in range(self.EMPLOYEE_START_ROW, sheet.max_row + 1):
            if sheet.cell(row=row, column=1).value == employee_name:
                return row
        return None

    def add_or_update_employee_advance(self, employee_name, amount, currency, option, date):
        try:
            workbook = self.nacti_nebo_vytvor_excel()
            sheet = workbook[self.ZALOHY_SHEET_NAME]
            row = self.get_employee_row(employee_name)
            
            if row is None:
                row = self.get_next_empty_row(sheet)
                sheet.cell(row=row, column=1, value=employee_name)
            
            if option == 'option1':
                column = 2 if currency == 'EUR' else 3
            else:  # option2
                column = 4 if currency == 'EUR' else 5
            
            current_value = sheet.cell(row=row, column=column).value or 0
            sheet.cell(row=row, column=column, value=current_value + amount)
            
            # Přidání data zálohy
            date_column = 26  # Předpokládáme, že datum bude v sloupci Z
            sheet.cell(row=row, column=date_column, value=datetime.strptime(date, '%Y-%m-%d').date())
            
            workbook.save(self.excel_cesta)
            logging.info(f"Záloha pro {employee_name} aktualizována: {amount} {currency} ({option}) k datu {date}")
            return True
        except Exception as e:
            logging.error(f"Chyba při ukládání zálohy: {e}")
            return False

    def get_next_empty_row(self, sheet):
        for row in range(self.EMPLOYEE_START_ROW, sheet.max_row + 2):
            if sheet.cell(row=row, column=1).value is None:
                return row
        return sheet.max_row + 1

    def get_employee_advances(self, employee_name):
        workbook = self.nacti_nebo_vytvor_excel()
        sheet = workbook[self.ZALOHY_SHEET_NAME]
        row = self.get_employee_row(employee_name)
        if row is None:
            return None
        return {
            'Option1_EUR': sheet.cell(row=row, column=2).value or 0,
            'Option1_CZK': sheet.cell(row=row, column=3).value or 0,
            'Option2_EUR': sheet.cell(row=row, column=4).value or 0,
            'Option2_CZK': sheet.cell(row=row, column=5).value or 0
        }

    def get_option_names(self):
        workbook = self.nacti_nebo_vytvor_excel()
        sheet = workbook[self.ZALOHY_SHEET_NAME]
        option1_name = sheet['B80'].value or 'Option 1'
        option2_name = sheet['D80'].value or 'Option 2'
        return option1_name, option2_name

if __name__ == "__main__":
    # Test code
    manager = ZalohyManager()
    manager.add_or_update_employee_advance("Jan Novák", 100, 'EUR', 'option1', '2023-05-01')
    manager.add_or_update_employee_advance("Jan Novák", 2000, 'CZK', 'option2', '2023-05-02')
    print(manager.get_employee_advances("Jan Novák"))
    print(manager.get_option_names())