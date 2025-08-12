# hodiny2025_manager.py
from datetime import datetime
from pathlib import Path
import logging
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

try:
    from utils.logger import setup_logger
    logger = setup_logger("hodiny2025_manager")
except ImportError:
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger("hodiny2025_manager")

class Hodiny2025Manager:
    def __init__(self, excel_path):
        self.excel_path = Path(excel_path)
        self.workbook_name = "Hodiny2025.xlsx"
        self.template_sheet_name = "MMhod25"

    def zapis_pracovni_doby(self, date, start_time, end_time, lunch_duration, num_employees):
        try:
            date_obj = datetime.strptime(date, "%Y-%m-%d")
            sheet_name = date_obj.strftime("%mhod%y")
            
            try:
                workbook = load_workbook(self.excel_path / self.workbook_name)
            except (FileNotFoundError, InvalidFileException):
                logger.error(f"Soubor {self.workbook_name} nenalezen nebo je poškozen.")
                return

            if sheet_name not in workbook.sheetnames:
                if self.template_sheet_name in workbook.sheetnames:
                    template_sheet = workbook[self.template_sheet_name]
                    new_sheet = workbook.copy_worksheet(template_sheet)
                    new_sheet.title = sheet_name
                else:
                    logger.error(f"List '{self.template_sheet_name}' nebyl nalezen.")
                    return
            
            sheet = workbook[sheet_name]
            row = date_obj.day + 2  # 1. den v měsíci je na řádku 3

            sheet.cell(row=row, column=5).value = datetime.strptime(start_time, "%H:%M").time()
            sheet.cell(row=row, column=7).value = datetime.strptime(end_time, "%H:%M").time()
            sheet.cell(row=row, column=6).value = float(lunch_duration)
            sheet.cell(row=row, column=14).value = num_employees

            workbook.save(self.excel_path / self.workbook_name)
            logger.info(f"Pracovní doba pro {date} byla zapsána do {self.workbook_name}/{sheet_name}.")

        except Exception as e:
            logger.error(f"Chyba při zápisu do {self.workbook_name}: {e}", exc_info=True)
