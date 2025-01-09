from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

try:
    workbook = load_workbook(r'C:\Users\Cowley\Desktop\Programovani\Hodweb\excel\hodiny2024.xlsx', data_only=True)
    workbook.save(r'C:\Users\Cowley\Desktop\Programovani\Hodweb\excel\Hodiny2024_repaired.xlsx')
    print("Soubor byl úspěšně opraven a uložen.")
except InvalidFileException as e:
    print(f"Chyba při načítání Excelového souboru: {e}")
