"""Služby pro práci s Excel soubory používané advanced konfigurací."""

from openpyxl import load_workbook

from config import Config
from security import safe_excel_path
from utils.logger import setup_logger

logger = setup_logger("excel_file_service")


def list_excel_files():
    """Vrátí seřazený seznam všech dostupných XLSX souborů."""
    return sorted(path.name for path in Config.EXCEL_BASE_PATH.glob("*.xlsx"))


def get_sheet_names(filename):
    """Vrátí názvy listů v zadaném Excel souboru."""
    file_path = safe_excel_path(Config.EXCEL_BASE_PATH, filename)
    if not file_path.exists():
        raise FileNotFoundError("Soubor nenalezen")

    workbook = load_workbook(file_path, read_only=True)
    try:
        return workbook.sheetnames
    finally:
        workbook.close()


def get_sheet_content(filename, sheet_name, max_rows=None, max_cols=30):
    """Vrátí obsah listu ve formátu vhodném pro frontend."""
    file_path = safe_excel_path(Config.EXCEL_BASE_PATH, filename)
    if not file_path.exists():
        raise FileNotFoundError("Soubor nenalezen")

    workbook = load_workbook(file_path, read_only=True, data_only=True)
    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError("List nenalezen")

        sheet = workbook[sheet_name]
        # Omezíme řádky pro zobrazení
        row_limit = min(max_rows or Config.MAX_ROWS_TO_DISPLAY_EXCEL_VIEWER, max(sheet.max_row or 0, 40), 100)
        col_limit = min(max_cols, max(sheet.max_column or 0, 16))

        data = []
        # iter_rows je v read_only režimu dramaticky rychlejší než cell(row, col)
        for row in sheet.iter_rows(min_row=1, max_row=row_limit, min_col=1, max_col=col_limit, values_only=True):
            row_data = [str(val) if val is not None else "" for val in row]
            # Pokud je řádek kratší než col_limit, doplníme prázdné stringy
            if len(row_data) < col_limit:
                row_data.extend([""] * (col_limit - len(row_data)))
            data.append(row_data)

        # Pokud sešit nevrátil dostatek řádků, doplníme prázdné
        while len(data) < row_limit:
            data.append([""] * col_limit)

        return {"data": data, "rows": len(data), "cols": col_limit}
    finally:
        workbook.close()


def rename_excel_file(old_filename, new_filename):
    """Přejmenuje XLSX soubor po základní validaci vstupu."""
    if not old_filename or not new_filename:
        raise ValueError("Chybí název souboru")

    if not old_filename.endswith(".xlsx") or not new_filename.endswith(".xlsx"):
        raise ValueError("Pouze .xlsx soubory mohou být přejmenovány")

    old_path = safe_excel_path(Config.EXCEL_BASE_PATH, old_filename)
    new_path = safe_excel_path(Config.EXCEL_BASE_PATH, new_filename)

    if not old_path.exists():
        raise FileNotFoundError(f"Soubor {old_filename} neexistuje")

    if new_path.exists():
        raise FileExistsError(f"Soubor {new_filename} již existuje")

    old_path.rename(new_path)
    logger.info("Soubor %s byl přejmenován na %s", old_filename, new_filename)
    return new_filename
