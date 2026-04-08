"""Služby pro upload a potvrzení přepisu Excel souborů."""

import io

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from werkzeug.utils import secure_filename

from config import Config


def normalize_upload_filename(raw_filename):
    """Validuje a normalizuje název nahrávaného souboru."""
    if not raw_filename or not raw_filename.lower().endswith(".xlsx"):
        raise ValueError("Lze nahrávat pouze soubory s příponou .xlsx.")

    filename = secure_filename(raw_filename)
    if not filename:
        raise ValueError("Neplatný název souboru.")

    return filename


def read_and_validate_excel(file_storage):
    """Načte obsah uploadu do paměti a ověří, že jde o platný XLSX."""
    file_bytes = file_storage.read()
    try:
        load_workbook(io.BytesIO(file_bytes))
    except InvalidFileException as exc:
        raise ValueError("Soubor není platný Excel soubor (.xlsx).") from exc
    return file_bytes


def store_temp_upload(filename, file_bytes):
    """Uloží dočasný soubor pro potvrzení přepsání."""
    temp_path = Config.EXCEL_BASE_PATH / f"temp_{filename}"
    temp_path.write_bytes(file_bytes)
    return temp_path.name


def save_uploaded_file(filename, file_bytes):
    """Uloží nahraný soubor do cílové složky."""
    file_path = Config.EXCEL_BASE_PATH / filename
    file_path.write_bytes(file_bytes)
    return file_path


def confirm_overwrite(temp_filename, filename):
    """Dokončí přepsání přesunutím dočasného souboru na finální cestu."""
    if not temp_filename or not filename:
        raise ValueError("Chyba při zpracování potvrzení.")

    temp_path = Config.EXCEL_BASE_PATH / temp_filename
    final_path = Config.EXCEL_BASE_PATH / filename

    if not temp_path.exists():
        raise FileNotFoundError("Dočasný soubor nenalezen. Zkuste nahrání znovu.")

    temp_path.rename(final_path)
    return final_path
