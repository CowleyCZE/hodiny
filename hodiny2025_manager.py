"""Správa souboru Hodiny2025.xlsx – měsíční evidence (1 sheet = 1 měsíc).

Zjednodušené schema listu:
 A=Den, B=Datum, C=Den v týdnu, D=Svátek, E/F/G=Začátek/Oběd/Konec,
 H=Celkem hodin (vzorec), I=Přesčasy, M=Počet zaměstnanců, N=Celkem * M.

Třída zajišťuje:
 - lazy vytvoření pracovního sešitu + template list
 - generování / inicializaci listu pro měsíc (01hod25 ... 12hod25)
 - zápis denních údajů + udržení vzorců
 - načítání souhrnů, denních záznamů a validaci integrity
"""

import calendar
import json
import logging
from datetime import datetime, time
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import MergedCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import coordinate_to_tuple
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.worksheet import Worksheet

try:
    from utils.logger import setup_logger

    logger = setup_logger("hodiny2025_manager")
except ImportError:
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger("hodiny2025_manager")


class Hodiny2025Manager:
    """
    Manager pro správu Excel souboru s evidencí pracovních hodin pro rok 2025.
    """

    HEADER_ROW, DATA_START_ROW, DATA_END_ROW, SUMMARY_ROW = 2, 3, 33, 34
    COL_DAY, COL_DATE, COL_WEEKDAY, COL_HOLIDAY = 1, 2, 3, 4
    COL_START, COL_LUNCH, COL_END = 5, 6, 7
    COL_TOTAL_HOURS, COL_OVERTIME, COL_NIGHT, COL_WEEKEND = 8, 9, 10, 11
    COL_HOLIDAY_HOURS, COL_EMPLOYEES, COL_TOTAL_ALL = 12, 13, 14

    CZECH_MONTHS = {
        1: "Leden",
        2: "Únor",
        3: "Březen",
        4: "Duben",
        5: "Květen",
        6: "Červen",
        7: "Červenec",
        8: "Srpen",
        9: "Září",
        10: "Říjen",
        11: "Listopad",
        12: "Prosinec",
    }
    CZECH_WEEKDAYS = ["Po", "Út", "St", "Čt", "Pá", "So", "Ne"]

    def __init__(self, excel_path):
        self.excel_path = Path(excel_path)
        self.workbook_name = "Hodiny2026.xlsx"
        self.template_sheet_name = "MMhod26"
        self.cash_template_sheet_name = "MMcash26"
        self.file_path = self.excel_path / self.workbook_name
        self._ensure_excel_file_exists()
        logger.info("Hodiny2025Manager inicializován pro soubor: %s", self.file_path)

    def _load_dynamic_config(self):
        """Načte dynamickou konfiguraci pro ukládání do XLSX souborů."""
        from config import Config

        if not Config.CONFIG_FILE_PATH.exists():
            return {}
        try:
            with open(Config.CONFIG_FILE_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError) as e:
            logger.error("Chyba při načítání dynamické konfigurace: %s", e, exc_info=True)
            return {}

    def _sheet_matches(self, configured_sheet, requested_sheet):
        if not configured_sheet or not requested_sheet:
            return configured_sheet == requested_sheet
        if configured_sheet == requested_sheet:
            return True
        # Zvláštní pravidlo pro měsíční listy: MMhod26 nebo 01hod26 by mělo platit pro jakýkoliv XXhod26
        if "hod" in configured_sheet and "hod" in requested_sheet:
            return configured_sheet.split("hod")[-1] == requested_sheet.split("hod")[-1]
        return False

    def _get_cell_coordinates(self, field_key, sheet_name=None):
        """Vrátí seznam (row, col) souřadnic pro daný field z dynamické konfigurace.

        Args:
            field_key: Klíč pole z konfigurace (např. 'start_time', 'date')
            sheet_name: Název listu, pokud chceme ověřit shodu

        Returns:
            list: Seznam (row, col) souřadnic nebo prázdný seznam pokud není nakonfigurováno
        """
        config = self._load_dynamic_config()
        monthly_config = config.get("monthly_time", {})
        field_configs = monthly_config.get(field_key, [])

        if not field_configs:
            return []

        coordinates = []
        for field_config in field_configs:
            # Ověř, že konfigurace je pro správný soubor a list
            if field_config.get("file") != self.workbook_name:
                logger.warning(
                    "Konfigurace pro monthly_time/%s odkazuje na jiný soubor: %s", field_key, field_config.get("file")
                )
                continue

            if sheet_name and not self._sheet_matches(field_config.get("sheet"), sheet_name):
                logger.warning(
                    "Konfigurace pro monthly_time/%s odkazuje na jiný list: %s (očekáván %s)",
                    field_key,
                    field_config.get("sheet"),
                    sheet_name,
                )
                continue

            cell = field_config.get("cell")
            if not cell:
                continue

            try:
                coordinates.append(coordinate_to_tuple(cell))  # Převede např. "A1" na (1, 1)
            except ValueError as e:
                logger.error("Neplatná buňka v konfiguraci pro monthly_time/%s: %s - %s", field_key, cell, e)
                continue

        return coordinates

    def _ensure_excel_file_exists(self):
        if not self.file_path.exists():
            logger.info("Vytváří se nový soubor: %s", self.file_path)
            self._create_new_workbook()

    def _create_new_workbook(self):
        workbook = Workbook()
        if workbook.active:
            workbook.remove(workbook.active)
        template_sheet = workbook.create_sheet(title=self.template_sheet_name)
        self._setup_template_sheet(template_sheet)
        current_month = datetime.now().month
        current_sheet = workbook.copy_worksheet(template_sheet)
        current_sheet.title = f"{current_month:02d}hod25"
        self._setup_month_sheet(current_sheet, current_month, 2025)
        workbook.save(self.file_path)
        logger.info("Vytvořen nový Excel soubor: %s", self.file_path)

    def _setup_template_sheet(self, sheet: Worksheet):
        self._set_cell_value(sheet, 1, 1, "Měsíční výkaz práce - [Měsíc] 2025")
        headers = [
            "Den",
            "Datum",
            "Den v týdnu",
            "Svátek",
            "Začátek",
            "Oběd (h)",
            "Konec",
            "Celkem hodin",
            "Přesčasy",
            "Noční práce",
            "Víkend",
            "Svátky",
            "Zaměstnanci",
            "Celkem odpracováno",
        ]

        for col, header in enumerate(headers, 1):
            target = self._set_cell_value(sheet, self.HEADER_ROW, col, header)
            if target:
                target.font = Font(bold=True)
                target.alignment = Alignment(horizontal="center")

        for day in range(1, 32):
            row = self.DATA_START_ROW + day - 1
            # use helper to avoid writing into MergedCell objects
            self._set_cell_value(sheet, row, self.COL_DAY, day)
            formula = f'=IF(AND(E{row}<>"",G{row}<>""),(G{row}-E{row})*24-F{row},0)'
            self._set_cell_formula(sheet, row, self.COL_TOTAL_HOURS, formula)
            self._set_cell_formula(sheet, row, self.COL_OVERTIME, f"=MAX(0,H{row}-8)")
            self._set_cell_formula(sheet, row, self.COL_TOTAL_ALL, f"=H{row}*M{row}")

        self._set_summary_formulas(sheet)

    def _set_summary_formulas(self, sheet: Worksheet):
        self._set_cell_formula(sheet, self.SUMMARY_ROW, 1, "SOUHRN:")
        for col, formula_col in [(self.COL_TOTAL_HOURS, "H"), (self.COL_OVERTIME, "I"), (self.COL_TOTAL_ALL, "N")]:
            formula = f"=SUM({formula_col}{self.DATA_START_ROW}:{formula_col}{self.DATA_END_ROW})"
            self._set_cell_formula(sheet, self.SUMMARY_ROW, col, formula)
            cell = self._get_actual_cell(sheet, self.SUMMARY_ROW, col)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="CCCCCC")

    def _setup_month_sheet(self, sheet: Worksheet, month: int, year: int):
        month_name = self.CZECH_MONTHS[month]
        self._set_cell_value(sheet, 1, 1, f"Měsíční výkaz práce - {month_name} {year}")
        days_in_month = calendar.monthrange(year, month)[1]

        for day in range(1, days_in_month + 1):
            row = self.DATA_START_ROW + day - 1
            date_obj = datetime(year, month, day)
            self._set_cell_value(sheet, row, self.COL_DATE, date_obj.strftime("%d.%m.%Y"))
            self._set_cell_value(sheet, row, self.COL_WEEKDAY, self.CZECH_WEEKDAYS[date_obj.weekday()])
            if date_obj.weekday() >= 5:
                for col in range(1, 15):
                    cell = self._get_actual_cell(sheet, row, col)
                    cell.fill = PatternFill("solid", fgColor="FFE6E6")

        for day in range(days_in_month + 1, 32):
            row = self.DATA_START_ROW + day - 1
            for col in range(1, 15):
                self._set_cell_formula(sheet, row, col, "")

    def _get_actual_cell(self, sheet: Worksheet, row: int, col: int):
        """Return the real (top-left) cell for the given coordinates, handling merged cells.

        If the cell at (row, col) is part of a merged range, return the top-left anchor cell of that range.
        Otherwise return sheet.cell(row, col).
        """
        cell = sheet.cell(row=row, column=col)
        if isinstance(cell, MergedCell):
            for merged in sheet.merged_cells.ranges:
                if merged.min_row <= row <= merged.max_row and merged.min_col <= col <= merged.max_col:
                    return sheet.cell(row=merged.min_row, column=merged.min_col)
        return cell

    def get_or_create_year_sheet(self, year: int = 2026) -> tuple[Workbook, Worksheet]:
        sheet_name = f"{year}hod"
        try:
            workbook = load_workbook(self.file_path)
        except (FileNotFoundError, InvalidFileException):
            self._create_new_workbook()
            workbook = load_workbook(self.file_path)

        if sheet_name not in workbook.sheetnames:
            if self.template_sheet_name not in workbook.sheetnames:
                raise ValueError(f"Template list '{self.template_sheet_name}' nebyl nalezen")
            template_sheet = workbook[self.template_sheet_name]
            new_sheet = workbook.copy_worksheet(template_sheet)
            new_sheet.title = sheet_name
            # Vyplnit dny v roce
            start_date = datetime(year, 1, 1)
            for i in range(365 + (1 if calendar.isleap(year) else 0)):
                date_obj = start_date + __import__("datetime").timedelta(days=i)
                new_sheet.cell(row=3+i, column=1).value = date_obj.strftime("%d.%m.%Y")
            logger.info("Vytvořen nový roční list: %s", sheet_name)
            return workbook, new_sheet

        return workbook, workbook[sheet_name]

    def get_or_create_month_sheet(self, month: int, year: int = 2025):
        # Backward compatibility, redirect to year sheet
        return self.get_or_create_year_sheet(year)

    def get_or_create_cash_sheet(self, month: int, year: int = 2026) -> tuple[Workbook, Worksheet]:
        """Získá nebo vytvoří list výdajů ze šablony MMcash26. Vytváří se formát CAcashRR od CA=05."""
        year_suffix = str(year)[2:]
        try:
            workbook = load_workbook(self.file_path)
        except (FileNotFoundError, InvalidFileException):
            self._create_new_workbook()
            workbook = load_workbook(self.file_path)

        import re
        current_ca = 5
        for name in workbook.sheetnames:
            match = re.match(r"^(\d+)cash" + year_suffix + r"$", name)
            if match:
                ca = int(match.group(1))
                if ca > current_ca:
                    current_ca = ca

        sheet_name = f"{current_ca:02d}cash{year_suffix}"

        if sheet_name not in workbook.sheetnames:
            if self.cash_template_sheet_name not in workbook.sheetnames:
                raise ValueError(f"Template list výdajů '{self.cash_template_sheet_name}' nebyl nalezen")
            template_sheet = workbook[self.cash_template_sheet_name]
            new_sheet = workbook.copy_worksheet(template_sheet)
            new_sheet.title = sheet_name
            logger.info("Vytvořen nový list výdajů: %s", sheet_name)
            return workbook, new_sheet

        return workbook, workbook[sheet_name]

    def _get_project_start_date(self):
        """Načte datum začátku projektu ze settings.json. Vrátí datetime nebo None."""
        try:
            from services.settings_service import load_settings
            settings = load_settings()
            start_str = settings.get("project_info", {}).get("start_date")
            if start_str:
                return datetime.strptime(start_str, "%Y-%m-%d")
        except Exception:
            pass
        return None

    def _find_free_row(self, sheet, col: int, row_start: int = 3, row_end: int = 42) -> int:
        """Najde první volný (None) řádek v daném sloupci v rozsahu row_start..row_end."""
        for r in range(row_start, row_end + 1):
            if sheet.cell(row=r, column=col).value is None:
                return r
        logger.warning("Sloupec %d: všechny řádky %d–%d jsou obsazeny.", col, row_start, row_end)
        return row_end

    def zapis_vydaje(self, category: str, amount: float, currency: str, payment_method: str,
                     date_str: str, description: str = ""):
        """Zapíše výdaj do odpovídajícího listu CAcashRR podle specifikace MMcash26."""
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")

        project_start_date = self._get_project_start_date()
        if project_start_date and date_obj < project_start_date:
            logger.info("Zápis výdaje %s přeskočen: datum je před začátkem projektu", category)
            return

        workbook, sheet = self.get_or_create_cash_sheet(date_obj.month, date_obj.year)
        date_formatted = date_obj.strftime("%d.%m.%Y")
        
        category_upper = category.upper()
        
        if "ZÁLOHA" in category_upper or "ZALOHA" in category_upper:
            typ_pohybu = "ZÁLOHA_ZAMĚSTNANCI"
            zamesnanec = (description.replace("Záloha - ", "").split(" (")[0].strip()
                          if "Záloha - " in description else description)
            kategorie = "ZÁLOHA_ZAMĚSTNANCI"
        elif "BANKOMAT" in category_upper or "VÝBĚR" in category_upper or "VYBER" in category_upper:
            typ_pohybu = "PŘÍJEM"
            kategorie = "BANKOMAT"
            zamesnanec = ""
        else:
            typ_pohybu = "VÝDAJ"
            zamesnanec = ""
            if "NAFTA" in category_upper or "TANKOVÁNÍ" in category_upper or "TANKOVANI" in category_upper:
                kategorie = "TANKOVÁNÍ"
            elif "PEAGE" in category_upper or "MÝTO" in category_upper or "MYTO" in category_upper:
                kategorie = "PEAGE"
            elif "UBYTOVÁNÍ" in category_upper or "UBYTOVANI" in category_upper:
                kategorie = "UBYTOVÁNÍ"
            else:
                kategorie = "OSTATNÍ"

        uhrada = "KARTOU" if "KART" in payment_method.upper() else "HOTOVĚ"
        
        target_row = self._find_free_row(sheet, 1, 3, 302)
        
        self._set_cell_value(sheet, target_row, 1, date_formatted)
        self._set_cell_value(sheet, target_row, 2, kategorie)
        self._set_cell_value(sheet, target_row, 3, description if description else kategorie)
        self._set_cell_value(sheet, target_row, 4, currency.upper())
        self._set_cell_value(sheet, target_row, 5, float(amount))
        self._set_cell_value(sheet, target_row, 6, uhrada)
        self._set_cell_value(sheet, target_row, 7, zamesnanec)
        self._set_cell_value(sheet, target_row, 8, typ_pohybu)
        self._set_cell_value(sheet, target_row, 9, "")
        self._set_cell_value(sheet, target_row, 10, "API")
        
        workbook.save(self.file_path)
        logger.info(f"Zapsán pohyb do {sheet.title}: {kategorie} {amount} {currency} na řádek {target_row}")

    def zapis_pracovni_doby(self, date_str, start_time_str, end_time_str, lunch_duration_str, num_employees):
        try:
            date_obj = datetime.strptime(date_str, "%Y-%m-%d")
            workbook, sheet = self.get_or_create_year_sheet(date_obj.year)
            
            # Find row by date in column A
            date_formatted = date_obj.strftime("%d.%m.%Y")
            row = 3
            for r in range(3, 369):
                cell_val = sheet.cell(row=r, column=1).value
                if cell_val and str(cell_val).strip() == date_formatted:
                    row = r
                    break
                # If we encounter empty date, we can use this row
                if not cell_val:
                    row = r
                    sheet.cell(row=row, column=1).value = date_formatted
                    break

            self._update_day_record(sheet, row, start_time_str, end_time_str, lunch_duration_str, num_employees)

            workbook.save(self.file_path)
            logger.info("Pracovní doba pro %s byla zapsána do listu %s", date_str, sheet.title)
        except (ValueError, IOError, FileNotFoundError) as e:
            logger.error("Chyba při zápisu pracovní doby pro %s: %s", date_str, e, exc_info=True)
            raise

    def _update_day_record(self, sheet, row, start_time_str, end_time_str, lunch_duration_str, num_employees):
        # C (3): Začátek práce
        if start_time_str and start_time_str != "00:00":
            self._set_cell_value(sheet, row, 3, datetime.strptime(start_time_str, "%H:%M").time())
        # D (4): Pauza
        lunch_hours = float(lunch_duration_str) if lunch_duration_str else 0.0
        lunch_cell = self._set_cell_value(sheet, row, 4, lunch_hours)
        if lunch_cell:
            lunch_cell.number_format = "0.0"
        # E (5): Konec práce
        if end_time_str and end_time_str != "00:00":
            self._set_cell_value(sheet, row, 5, datetime.strptime(end_time_str, "%H:%M").time())
        # N (14): Počet osob
        self._set_cell_value(sheet, row, 14, num_employees if num_employees > 0 else 0)

    def _ensure_formulas_are_set(self, sheet, row):
        formulas = {
            self.COL_TOTAL_HOURS: f'=IF(AND(E{row}<>"",G{row}<>""),(G{row}-E{row})*24-F{row},0)',
            self.COL_OVERTIME: f"=MAX(0,H{row}-8)",
            self.COL_TOTAL_ALL: f"=H{row}*M{row}",
        }
        for col, formula in formulas.items():
            cell = sheet.cell(row=row, column=col)
            if not isinstance(cell, MergedCell) and (not cell.value or not str(cell.value).startswith("=")):
                cell.value = formula

    def get_monthly_summary(self, month: int, year: int = 2025) -> dict:
        summary = {
            "month": month,
            "year": year,
            "month_name": self.CZECH_MONTHS.get(month, "Neznámý"),
            "total_hours": 0,
            "total_overtime": 0,
            "total_all_employees": 0,
            "sheet_name": f"{month:02d}hod{str(year)[2:]}",
            "error": None,
        }
        try:
            _, sheet = self.get_or_create_month_sheet(month, year)
            summary.update(
                {
                    "total_hours": self._safe_float(sheet.cell(self.SUMMARY_ROW, self.COL_TOTAL_HOURS).value),
                    "total_overtime": self._safe_float(sheet.cell(self.SUMMARY_ROW, self.COL_OVERTIME).value),
                    "total_all_employees": self._safe_float(sheet.cell(self.SUMMARY_ROW, self.COL_TOTAL_ALL).value),
                }
            )
        except (ValueError, IOError, FileNotFoundError) as e:
            logger.error("Chyba při získávání měsíčního souhrnu pro %d/%d: %s", month, year, e)
            summary["error"] = str(e)
        return summary

    def get_daily_record(self, date_str: str) -> dict:
        try:
            date_obj = datetime.strptime(date_str, "%Y-%m-%d")
            _, sheet = self.get_or_create_month_sheet(date_obj.month, date_obj.year)
            row = self.DATA_START_ROW + date_obj.day - 1

            # Load with data_only=True to get calculated values
            data_workbook = load_workbook(self.file_path, data_only=True)
            data_sheet = data_workbook[sheet.title]

            record = self._extract_daily_data(data_sheet, row)
            record["date"] = date_str
            record["day"] = date_obj.day
            record["row"] = row
            record["sheet_name"] = sheet.title

            self._recalculate_if_needed(record)

            return record
        except (ValueError, IOError, FileNotFoundError) as e:
            logger.error("Chyba při získávání záznamu pro %s: %s", date_str, e)
            return {"date": date_str, "error": str(e)}

    def _extract_daily_data(self, sheet, row):
        return {
            "start_time": self._safe_time_format(sheet.cell(row, self.COL_START).value),
            "end_time": self._safe_time_format(sheet.cell(row, self.COL_END).value),
            "lunch_hours": self._safe_float(sheet.cell(row, self.COL_LUNCH).value),
            "total_hours": self._safe_float(sheet.cell(row, self.COL_TOTAL_HOURS).value),
            "overtime": self._safe_float(sheet.cell(row, self.COL_OVERTIME).value),
            "num_employees": self._safe_int(sheet.cell(row, self.COL_EMPLOYEES).value),
            "total_all_employees": self._safe_float(sheet.cell(row, self.COL_TOTAL_ALL).value),
        }

    def _recalculate_if_needed(self, record):
        if record["total_hours"] == 0.0 and record["start_time"] and record["end_time"]:
            try:
                start = datetime.strptime(record["start_time"], "%H:%M")
                end = datetime.strptime(record["end_time"], "%H:%M")
                delta_seconds = (end - start).total_seconds()
                if delta_seconds < 0:
                    delta_seconds += 24 * 3600
                hours = delta_seconds / 3600 - record["lunch_hours"]
                record["total_hours"] = max(0.0, hours)
            except (ValueError, TypeError):
                pass

        if record["overtime"] == 0.0 and record["total_hours"] > 8.0:
            record["overtime"] = record["total_hours"] - 8.0

        if record["total_all_employees"] == 0.0 and record["total_hours"] > 0:
            record["total_all_employees"] = record["total_hours"] * record["num_employees"]

    def _set_cell_value(self, sheet, row, col, value):
        cell = sheet.cell(row=row, column=col)
        if isinstance(cell, MergedCell):
            # find merged range and set value on the top-left anchor cell
            for merged in sheet.merged_cells.ranges:
                if merged.min_row <= row <= merged.max_row and merged.min_col <= col <= merged.max_col:
                    target = sheet.cell(row=merged.min_row, column=merged.min_col)
                    target.value = value
                    return target
            # fallback: cannot set on a MergedCell that doesn't match a known range
            return None
        else:
            cell.value = value
            return cell

    def _set_cell_formula(self, sheet, row, col, formula):
        return self._set_cell_value(sheet, row, col, formula)

    def _safe_time_format(self, value):
        if isinstance(value, time):
            return value.strftime("%H:%M")
        return value if isinstance(value, str) else None

    def _safe_float(self, value):
        try:
            return float(value) if value is not None else 0.0
        except (ValueError, TypeError):
            return 0.0

    def _safe_int(self, value):
        try:
            return int(value) if value is not None else 0
        except (ValueError, TypeError):
            return 0

    def create_test_data(self):
        logger.info("Vytváří se testovací data pro Hodiny2025.xlsx")
        test_dates = [
            ("2025-01-02", "07:00", "15:30", "0.5", 3),
            ("2025-01-03", "07:00", "16:00", "1.0", 3),
            ("2025-01-06", "08:00", "16:30", "0.5", 2),
        ]
        for data in test_dates:
            try:
                self.zapis_pracovni_doby(*data)
                logger.info("✅ Testovací záznam vytvořen: %s", data[0])
            except Exception as e:
                logger.error("❌ Chyba při vytváření testovacího záznamu %s: %s", data[0], e)

    def validate_data_integrity(self) -> dict:
        results = {"valid": True, "errors": [], "warnings": [], "sheets_checked": 0, "records_checked": 0}
        try:
            workbook = load_workbook(self.file_path)
            sheet_names = [s for s in workbook.sheetnames if s != self.template_sheet_name]
            results["sheets_checked"] = len(sheet_names)

            for sheet_name in sheet_names:
                sheet = workbook[sheet_name]
                for row in range(self.DATA_START_ROW, self.DATA_END_ROW + 1):
                    results["records_checked"] += 1
                    total_formula = sheet.cell(row=row, column=self.COL_TOTAL_HOURS).value
                    if isinstance(total_formula, str) and not total_formula.startswith("="):
                        results["valid"] = False
                        results["errors"].append(f"List {sheet_name}, řádek {row}: Chybí vzorec.")

                    start = sheet.cell(row=row, column=self.COL_START).value
                    end = sheet.cell(row=row, column=self.COL_END).value
                    if bool(start) != bool(end):
                        results["warnings"].append(f"List {sheet_name}, řádek {row}: Chybí čas začátku/konce.")
        except Exception as e:
            results["valid"] = False
            results["errors"].append(f"Chyba při validaci: {e}")
        return results
