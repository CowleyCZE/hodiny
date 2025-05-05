# excel_manager.py
import contextlib
import logging
import os
import platform
import re
import shutil
from contextlib import contextmanager
from datetime import datetime, timedelta
from pathlib import Path
from threading import Lock
from typing import Dict, List, Optional, Tuple, Union, Generator, Any

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.copier import WorksheetCopy

from config import Config
from utils.logger import setup_logger

logger = setup_logger("excel_manager")

# Kontrola operačního systému
IS_WINDOWS = platform.system() == "Windows"

class ExcelManager:
    """Spravuje operace s aktivním Excel souborem"""
    
    def __init__(self, base_path: Union[str, Path], active_filename: str, template_filename: str):
        """Inicializace správce Excel souboru"""
        if not base_path:
            raise ValueError("Base path nesmí být prázdný")
            
        if not active_filename:
            raise ValueError("Aktivní název souboru nesmí být prázdný")
            
        self.base_path = Path(base_path)
        self.active_filename = active_filename
        self.template_filename = template_filename
        self.active_file_path = self.base_path / active_filename
        
        # Cache pro workbooky
        self._workbook_cache = {}
        self._cache_lock = Lock()
        
        # Validace při inicializaci
        self._validate_excel_structure()
        logger.info(f"Inicializován ExcelManager pro {active_filename}")

    def _validate_excel_structure(self) -> None:
        """Validuje strukturu Excel souboru"""
        try:
            if not self.active_file_path.exists():
                logger.warning(f"Aktivní soubor {self.active_file_path} neexistuje")
                return
                
            with self._get_workbook(self.active_file_path) as wb:
                required_sheets = ["Pracovní doba", "Zálohy"]
                for sheet_name in required_sheets:
                    if sheet_name not in wb.sheetnames:
                        raise InvalidFileException(f"Chybí požadovaný list: {sheet_name}")
                
                # Validace sloupců na listu "Pracovní doba"
                ws = wb["Pracovní doba"]
                expected_columns = ["Datum", "Zaměstnanec", "Začátek", "Konec", "Oběd", "Čistý čas"]
                for i, col in enumerate(expected_columns):
                    if ws.cell(row=1, column=i+1).value != col:
                        raise InvalidFileException(f"Neplatná hlavička v Excel souboru: {ws.cell(row=1, column=i+1).value}")
                        
        except InvalidFileException as e:
            logger.error(f"Chyba při validaci struktury Excel souboru: {e}")
            raise
        except Exception as e:
            logger.error(f"Neočekávaná chyba při validaci struktury: {e}", exc_info=True)

    def _get_cache_key(self, file_path: Path, read_only: bool) -> str:
        """Vytvoří unikátní klíč pro cache"""
        return f"{file_path.resolve()}_{read_only}"

    @contextmanager
    def _get_workbook(self, file_path: Path, read_only: bool = False) -> Generator[Any, Any, Any]:
        """Získá workbook z cache nebo ho načte z disku"""
        try:
            # Získání workbooku z cache
            cache_key = self._get_cache_key(file_path, read_only)
            
            with self._cache_lock:
                if cache_key in self._workbook_cache:
                    wb = self._workbook_cache[cache_key]
                    logger.debug(f"Načten workbook z cache: {file_path}")
                else:
                    # Načtení nového workbooku
                    wb = load_workbook(file_path, read_only=read_only)
                    self._workbook_cache[cache_key] = wb
                    logger.debug(f"Načten nový workbook: {file_path}")
            
            # Ujistíme se, že adresář existuje
            file_path.parent.mkdir(parents=True, exist_ok=True)
            
            try:
                yield wb
                # Uložení změn, pokud není read-only
                if not read_only:
                    wb.save(file_path)
                    logger.debug(f"Uloženy změny v souboru: {file_path}")
                    
            finally:
                # Uvolnění zdrojů
                if cache_key not in self._get_cache_key(file_path, not read_only):
                    wb.close()
                    
        except Exception as e:
            logger.error(f"Chyba při práci s workbookem: {e}", exc_info=True)
            raise

    def close_cached_workbooks(self) -> None:
        """Zavře všechny cached workbooky"""
        with self._cache_lock:
            for wb in self._workbook_cache.values():
                try:
                    wb.close()
                except Exception as e:
                    logger.error(f"Chyba při zavírání workbooku: {e}")
            self._workbook_cache.clear()
        logger.debug("Všechny cached workbooky byly zavřeny")

    def get_active_file_path(self) -> Path:
        """Vrátí cestu k aktuálnímu souboru"""
        if not self.active_file_path:
            raise ValueError("Aktivní soubor není definován")
        return self.active_file_path

    def _find_empty_row(self, ws, column: int = 1) -> int:
        """Najde první prázdný řádek ve specifikovaném sloupci"""
        for row in ws.iter_rows(min_col=column, max_col=column):
            if not row[0].value:
                return row[0].row
        return ws.max_row + 1

    def _get_employee_row(self, ws, employee_name: str) -> Optional[int]:
        """Najde řádek zaměstnance v listu"""
        for row in ws.iter_rows(min_col=2, max_col=2):  # Sloupec B
            if row[0].value == employee_name:
                return row[0].row
        return None

    def _normalize_date(self, date_str: str) -> str:
        """Normalizuje datum do formátu YYYY-MM-DD"""
        try:
            return datetime.strptime(date_str, "%Y-%m-%d").strftime("%Y-%m-%d")
        except ValueError:
            try:
                return datetime.strptime(date_str, "%d.%m.%Y").strftime("%Y-%m-%d")
            except ValueError:
                raise ValueError(f"Neplatný formát data: {date_str}. Očekáváno YYYY-MM-DD nebo DD.MM.YYYY")

    def _calculate_work_hours(self, start_time: str, end_time: str, lunch_duration: float) -> float:
        """Vypočítá čistý čas práce"""
        try:
            start = datetime.strptime(start_time, "%H:%M")
            end = datetime.strptime(end_time, "%H:%M")
            
            total_seconds = (end - start).total_seconds()
            work_hours = total_seconds / 3600 - lunch_duration
            
            # Validace výsledku
            if work_hours < 0:
                raise ValueError(f"Čistý čas práce nesmí být záporný: {work_hours:.2f} hodin")
                
            return round(work_hours, 2)
            
        except ValueError as e:
            logger.error(f"Chyba při výpočtu pracovního času: {e}")
            raise

    def record_time(self, employee: str, date: str, start_time: str, end_time: str, lunch_duration: str) -> Tuple[bool, str]:
        """Zaznamená pracovní dobu do Excel souboru"""
        try:
            # Validace vstupů
            date = self._normalize_date(date)
            datetime.strptime(start_time, "%H:%M")
            datetime.strptime(end_time, "%H:%M")
            
            # Převod pauzy na float
            try:
                lunch_duration = float(lunch_duration.replace(",", "."))
            except ValueError:
                raise ValueError(f"Neplatná délka pauzy: {lunch_duration}")
                
            # Výpočet čistého času
            work_hours = self._calculate_work_hours(start_time, end_time, lunch_duration)
            
            # Zápis do Excel souboru
            with self._get_workbook(self.active_file_path) as wb:
                ws = wb["Pracovní doba"]
                empty_row = self._find_empty_row(ws)
                
                ws.cell(row=empty_row, column=1, value=date)
                ws.cell(row=empty_row, column=2, value=employee)
                ws.cell(row=empty_row, column=3, value=start_time)
                ws.cell(row=empty_row, column=4, value=end_time)
                ws.cell(row=empty_row, column=5, value=lunch_duration)
                ws.cell(row=empty_row, column=6, value=work_hours)
                
                # Aktualizace statistik
                self._update_employee_stats(ws, employee)
                
            logger.info(f"Zaznamenán čas pro {employee}: {date}, {start_time}-{end_time}")
            return True, "Pracovní doba byla úspěšně uložena"
            
        except Exception as e:
            logger.error(f"Chyba při ukládání pracovní doby: {e}", exc_info=True)
            return False, str(e)

    def _update_employee_stats(self, ws, employee: str) -> None:
        """Aktualizuje statistiky zaměstnance"""
        # Implementace aktualizace statistik zaměstnance
        pass

    def get_week_stats(self, employee: str = None) -> Dict:
        """Získá statistiky za týden"""
        try:
            with self._get_workbook(self.active_file_path, read_only=True) as wb:
                ws = wb["Pracovní doba"]
                
                stats = {
                    "total_hours": 0,
                    "daily_hours": {},
                    "total_days": 0,
                    "week_number": datetime.now().isocalendar()[1]
                }
                
                # Implementace logiky pro týdenní statistiky
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if employee and row[1] != employee:
                        continue
                        
                    try:
                        date = datetime.strptime(row[0], "%Y-%m-%d")
                        if date.isocalendar()[1] == stats["week_number"]:
                            stats["total_hours"] += float(row[5]) if row[5] else 0
                            stats["daily_hours"][row[0]] = float(row[5]) if row[5] else 0
                    except (ValueError, TypeError) as e:
                        logger.warning(f"Chyba při zpracování řádku: {row} - {e}")
                        continue
                
                stats["total_days"] = len(stats["daily_hours"])
                return stats
                
        except Exception as e:
            logger.error(f"Chyba při získávání týdenních statistik: {e}")
            return {"error": str(e)}

    def get_month_stats(self, employee: str = None) -> Dict:
        """Získá statistiky za měsíc"""
        try:
            with self._get_workbook(self.active_file_path, read_only=True) as wb:
                ws = wb["Pracovní doba"]
                
                current_month = datetime.now().month
                stats = {
                    "total_hours": 0,
                    "daily_hours": {},
                    "total_days": 0,
                    "month": current_month
                }
                
                # Implementace logiky pro měsíční statistiky
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if employee and row[1] != employee:
                        continue
                        
                    try:
                        date = datetime.strptime(row[0], "%Y-%m-%d")
                        if date.month == current_month:
                            stats["total_hours"] += float(row[5]) if row[5] else 0
                            stats["daily_hours"][row[0]] = float(row[5]) if row[5] else 0
                    except (ValueError, TypeError) as e:
                        logger.warning(f"Chyba při zpracování řádku: {row} - {e}")
                        continue
                
                stats["total_days"] = len(stats["daily_hours"])
                return stats
                
        except Exception as e:
            logger.error(f"Chyba při získávání měsíčních statistik: {e}")
            return {"error": str(e)}

    def get_year_stats(self, employee: str = None) -> Dict:
        """Získá statistiky za rok"""
        try:
            with self._get_workbook(self.active_file_path, read_only=True) as wb:
                ws = wb["Pracovní doba"]
                
                current_year = datetime.now().year
                stats = {
                    "total_hours": 0,
                    "monthly_hours": {str(m): 0 for m in range(1, 13)},
                    "total_days": 0,
                    "year": current_year
                }
                
                # Implementace logiky pro roční statistiky
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if employee and row[1] != employee:
                        continue
                        
                    try:
                        date = datetime.strptime(row[0], "%Y-%m-%d")
                        if date.year == current_year:
                            month = str(date.month)
                            stats["monthly_hours"][month] += float(row[5]) if row[5] else 0
                            stats["total_days"] += 1
                    except (ValueError, TypeError) as e:
                        logger.warning(f"Chyba při zpracování řádku: {row} - {e}")
                        continue
                
                stats["total_hours"] = sum(stats["monthly_hours"].values())
                return stats
                
        except Exception as e:
            logger.error(f"Chyba při získávání ročních statistik: {e}")
            return {"error": str(e)}

    def get_total_stats(self, employee: str = None) -> Dict:
        """Získá celkové statistiky"""
        try:
            with self._get_workbook(self.active_file_path, read_only=True) as wb:
                ws = wb["Pracovní doba"]
                
                stats = {
                    "total_hours": 0,
                    "total_days": 0,
                    "total_records": 0,
                    "first_record": None,
                    "last_record": None
                }
                
                # Implementace logiky pro celkové statistiky
                dates = []
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if employee and row[1] != employee:
                        continue
                        
                    try:
                        stats["total_hours"] += float(row[5]) if row[5] else 0
                        stats["total_days"] += 1
                        stats["total_records"] += 1
                        dates.append(datetime.strptime(row[0], "%Y-%m-%d"))
                    except (ValueError, TypeError) as e:
                        logger.warning(f"Chyba při zpracování řádku: {row} - {e}")
                        continue
                
                if dates:
                    stats["first_record"] = min(dates).strftime("%Y-%m-%d")
                    stats["last_record"] = max(dates).strftime("%Y-%m-%d")
                
                return stats
                
        except Exception as e:
            logger.error(f"Chyba při získávání celkových statistik: {e}")
            return {"error": str(e)}

    def update_project_info(self, project_name: str, start_date: str, end_date: str = None) -> bool:
        """Aktualizuje informace o projektu v Excel souboru"""
        try:
            with self._get_workbook(self.active_file_path) as wb:
                ws = wb.active  # Předpokládáme, že hlavní list je aktivní
                
                ws["B1"] = f"Název projektu: {project_name}"
                ws["B2"] = f"Zahájení projektu: {start_date}"
                if end_date:
                    ws["B3"] = f"Ukončení projektu: {end_date}"
                else:
                    ws["B3"] = "Ukončení projektu: "
                    
            logger.info(f"Informace o projektu aktualizovány: {project_name}")
            return True
            
        except Exception as e:
            logger.error(f"Chyba při aktualizaci informací projektu: {e}")
            return False

    def _ensure_zalohy_sheet(self, wb) -> None:
        """Zajistí existenci listu pro zálohy"""
        if "Zálohy" not in wb.sheetnames:
            wb.create_sheet("Zálohy")
            # Inicializace nového listu
            ws = wb["Zálohy"]
            ws["A1"] = "Zálohy zaměstnanců"
            ws["B1"] = "Částka"
            ws["C1"] = "Měna"
            ws["D1"] = "Datum"
            ws["E1"] = "Poznámka"
            logger.info("Vytvořen nový list pro zálohy")

    def add_or_update_employee_advance(self, employee: str, amount: float, currency: str, option: str, date: str) -> Tuple[bool, str]:
        """Přidá nebo aktualizuje zálohu zaměstnance"""
        try:
            # Validace vstupů
            if not employee:
                raise ValueError("Zaměstnanec nesmí být prázdný")
                
            if amount <= 0:
                raise ValueError("Částka musí být kladná")
                
            if currency not in ["CZK", "EUR"]:
                raise ValueError(f"Neplatná měna: {currency}")
                
            normalized_date = self._normalize_date(date)
            
            with self._get_workbook(self.active_file_path) as wb:
                self._ensure_zalohy_sheet(wb)
                ws = wb["Zálohy"]
                
                # Najdeme nebo vytvoříme řádek pro zaměstnance
                row = self._get_employee_row(ws, employee)
                if not row:
                    row = self._find_empty_row(ws)
                    ws.cell(row=row, column=1, value=employee)
                
                # Najdeme volné místo pro zálohu
                col = 2  # Sloupec B pro částky
                while ws.cell(row=row, column=col).value:
                    col += 4  # Každá záloha zabírá 4 sloupce (částka, měna, datum, poznamka)
                
                ws.cell(row=row, column=col, value=amount)
                ws.cell(row=row, column=col+1, value=currency)
                ws.cell(row=row, column=col+2, value=normalized_date)
                ws.cell(row=row, column=col+3, value=f"Záloha {option}")
                
            logger.info(f"Přidána záloha pro {employee}: {amount} {currency} ({normalized_date})")
            return True, "Záloha byla úspěšně uložena"
            
        except Exception as e:
            logger.error(f"Chyba při ukládání zálohy: {e}", exc_info=True)
            return False, str(e)

    def _get_advance_options(self, ws) -> List[str]:
        """Získá možnosti záloh z listu"""
        options = []
        for col in range(2, ws.max_column + 1, 4):  # Každá záloha zabírá 4 sloupce
            cell = ws.cell(row=1, column=col)
            if cell.value:
                options.append(cell.value)
        return options

    def get_advance_options(self) -> List[str]:
        """Získá možnosti záloh"""
        try:
            with self._get_workbook(self.active_file_path, read_only=True) as wb:
                ws = wb["Zálohy"]
                return self._get_advance_options(ws)
                
        except Exception as e:
            logger.error(f"Chyba při získávání možností záloh: {e}")
            return ["Záloha 1", "Záloha 2"]  # Výchozí hodnoty

    def _get_employee_row(self, ws, employee_name: str) -> Optional[int]:
        """Najde řádek zaměstnance v listu"""
        for row in ws.iter_rows(min_col=1, max_col=1):  # Sloupec A
            if row[0].value == employee_name:
                return row[0].row
        return None

    def _find_empty_row(self, ws, column: int = 1) -> int:
        """Najde první prázdný řádek ve specifikovaném sloupci"""
        for row in ws.iter_rows(min_col=column, max_col=column):
            if not row[0].value:
                return row[0].row
        return ws.max_row + 1

    def _get_cislo_tydne(self, date_str: str) -> Dict[str, int]:
        """Získá číslo týdne pro dané datum"""
        try:
            date = datetime.strptime(date_str, "%Y-%m-%d")
            return {"week": date.isocalendar()[1], "year": date.isocalendar()[0]}
        except Exception as e:
            logger.error(f"Chyba při získávání čísla týdne: {e}")
            return {"week": 0, "year": 0}

    def ziskej_cislo_tydne(self, date_str: str) -> Dict[str, int]:
        """Veřejná metoda pro získání čísla týdne"""
        return self._get_cislo_tydne(date_str)

    def _validate_date(self, date_str: str) -> None:
        """Validuje datum"""
        try:
            datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            try:
                datetime.strptime(date_str, "%d.%m.%Y")
            except ValueError:
                raise ValueError(f"Neplatný formát data: {date_str}. Očekáváno YYYY-MM-DD nebo DD.MM.YYYY")

    def _validate_time(self, time_str: str) -> None:
        """Validuje čas"""
        try:
            datetime.strptime(time_str, "%H:%M")
        except ValueError:
            raise ValueError(f"Neplatný formát času: {time_str}. Očekáváno HH:MM")

    def _validate_time_range(self, start_time: str, end_time: str) -> None:
        """Validuje časový rozsah"""
        start = datetime.strptime(start_time, "%H:%M")
        end = datetime.strptime(end_time, "%H:%M")
        
        if start >= end:
            raise ValueError("Začátek práce musí být před koncem")

    def _validate_lunch_duration(self, duration: float) -> None:
        """Validuje délku oběda"""
        if duration < 0:
            raise ValueError("Délka oběda nesmí být záporná")
        if duration > 4:
            raise ValueError("Oběd nesmí být delší než 4 hodiny")

    def _validate_employee_name(self, name: str) -> None:
        """Validuje jméno zaměstnance"""
        if not name or not isinstance(name, str):
            raise ValueError("Neplatné jméno zaměstnance")
            
        name = name.strip()
        
        if len(name) < 2:
            raise ValueError("Jméno zaměstnance musí mít alespoň 2 znaky")
        
        if any(char.isdigit() for char in name):
            raise ValueError("Jméno zaměstnance nesmí obsahovat čísla")
        
        # Povolené znaky: písmena, mezery, pomlčky a tečky
        if not re.match(r'^[a-zA-Zá-žÁ-Ž\s\-\.]+$', name):
            raise ValueError("Jméno může obsahovat pouze písmena, mezery, pomlčky a tečky")

    def _update_employee_stats(self, ws, employee: str) -> None:
        """Aktualizuje statistiky zaměstnance"""
        # Implementace aktualizace statistik
        pass

    def get_total_hours(self) -> float:
        """Získá celkové hodiny"""
        try:
            with self._get_workbook(self.active_file_path, read_only=True) as wb:
                ws = wb["Pracovní doba"]
                
                total = 0
                for row in ws.iter_rows(min_row=2, values_only=True):
                    try:
                        total += float(row[5]) if row[5] else 0
                    except (ValueError, TypeError):
                        continue
                        
                return round(total, 2)
                
        except Exception as e:
            logger.error(f"Chyba při získávání celkových hodin: {e}")
            return 0

    def get_week_stats(self, employee: str = None) -> Dict:
        """Získá statistiky za aktuální týden"""
        try:
            with self._get_workbook(self.active_file_path, read_only=True) as wb:
                ws = wb["Pracovní doba"]
                
                current_date = datetime.now()
                current_week = current_date.isocalendar()[1]
                current_year = current_date.year
                
                stats = {
                    "total_hours": 0,
                    "daily_hours": {},
                    "total_days": 0,
                    "week_number": current_week
                }
                
                # Implementace pro týdenní statistiky
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if employee and row[1] != employee:
                        continue
                        
                    try:
                        date = datetime.strptime(row[0], "%Y-%m-%d")
                        if date.isocalendar()[1] == current_week and date.year == current_year:
                            stats["total_hours"] += float(row[5]) if row[5] else 0
                            stats["daily_hours"][row[0]] = float(row[5]) if row[5] else 0
                    except (ValueError, TypeError) as e:
                        logger.warning(f"Chyba při zpracování řádku: {row} - {e}")
                        continue
                
                stats["total_days"] = len(stats["daily_hours"])
                return stats
                
        except Exception as e:
            logger.error(f"Chyba při získávání týdenních statistik: {e}")
            return {"error": str(e)}

    def get_month_stats(self, employee: str = None) -> Dict:
        """Získá statistiky za aktuální měsíc"""
        try:
            with self._get_workbook(self.active_file_path, read_only=True) as wb:
                ws = wb["Pracovní doba"]
                
                current_month = datetime.now().month
                current_year = datetime.now().year
                
                stats = {
                    "total_hours": 0,
                    "daily_hours": {},
                    "total_days": 0,
                    "month": current_month
                }
                
                # Implementace pro měsíční statistiky
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if employee and row[1] != employee:
                        continue
                        
                    try:
                        date = datetime.strptime(row[0], "%Y-%m-%d")
                        if date.month == current_month and date.year == current_year:
                            stats["total_hours"] += float(row[5]) if row[5] else 0
                            stats["daily_hours"][row[0]] = float(row[5]) if row[5] else 0
                    except (ValueError, TypeError) as e:
                        logger.warning(f"Chyba při zpracování řádku: {row} - {e}")
                        continue
                
                stats["total_days"] = len(stats["daily_hours"])
                return stats
                
        except Exception as e:
            logger.error(f"Chyba při získávání měsíčních statistik: {e}")
            return {"error": str(e)}
