import logging
import re
from datetime import datetime
from typing import Union

from config import Config
from excel_manager import ExcelManager
from settings import load_settings
from utils.logger import setup_logger

logger = setup_logger("time_record")


class TimeRecord:
    def __init__(self, employee_manager):
        self.employee_manager = employee_manager
        self.excel_manager = ExcelManager(Config.EXCEL_BASE_PATH, Config.EXCEL_FILE_NAME)
        self.settings = None
        self.pracovni_doba = 0
        self.time_pattern = re.compile(r"^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$")

    def get_current_settings(self):
        """Načte aktuální nastavení ze souboru"""
        try:
            self.settings = load_settings()
            logging.info("Nastavení úspěšně načteno")
            return self.settings
        except Exception as e:
            logging.error(f"Chyba při načítání nastavení: {e}")
            self.settings = Config.get_default_settings()
            return self.settings

    def validate_time(self, time_str: str) -> bool:
        """Validuje formát času"""
        if not time_str:
            raise ValueError("Čas nesmí být prázdný")

        if not self.time_pattern.match(time_str):
            raise ValueError("Neplatný formát času. Použijte formát HH:MM (00:00-23:59)")

        return True

    def validate_lunch_duration(self, duration: Union[int, float]) -> bool:
        """Validuje délku obědové pauzy"""
        if duration is None:
            raise ValueError("Délka oběda nesmí být prázdná")

        if not isinstance(duration, (int, float)):
            raise ValueError("Délka oběda musí být číslo")

        if duration < 0 or duration > 4:
            raise ValueError("Délka oběda musí být mezi 0 a 4 hodinami")

        return True

    def validate_date(self, date_str: str) -> bool:
        """Validuje vybrané datum"""
        try:
            date = datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            raise ValueError("Neplatný formát data. Použijte YYYY-MM-DD")

        today = datetime.now().date()
        selected_date = date.date()

        if selected_date > today:
            raise ValueError("Nelze vybrat datum v budoucnosti")

        if (today - selected_date).days > 365:
            raise ValueError("Nelze vybrat datum starší než jeden rok")

        return True

    def validate_time_range(self, start_time: str, end_time: str) -> bool:
        """Validuje rozsah časů"""
        if not all([start_time, end_time]):
            raise ValueError("Časy nesmí být prázdné")

        start = datetime.strptime(start_time, "%H:%M")
        end = datetime.strptime(end_time, "%H:%M")

        if end <= start:
            raise ValueError("Konec práce musí být později než začátek")

        duration = (end - start).total_seconds() / 3600
        if duration > 24:
            raise ValueError("Pracovní doba nesmí přesáhnout 24 hodin")

        if duration < 0.5:
            raise ValueError("Pracovní doba musí být alespoň 30 minut")

        return True

    def calculate_work_hours(self, start_time: str, end_time: str, lunch_duration: float) -> float:
        """Vypočítá odpracované hodiny"""
        start = datetime.strptime(start_time, "%H:%M")
        end = datetime.strptime(end_time, "%H:%M")
        total_hours = (end - start).total_seconds() / 3600 - lunch_duration

        if total_hours <= 0:
            raise ValueError("Celková pracovní doba musí být kladná")

        return total_hours

    def record_time(self, date: str, start_time: str, end_time: str, lunch_duration: float):
        """Zaznamená pracovní dobu"""
        try:
            # Validace všech vstupů
            self.validate_date(date)
            self.validate_time(start_time)
            self.validate_time(end_time)
            self.validate_time_range(start_time, end_time)
            self.validate_lunch_duration(lunch_duration)

            # Výpočet pracovní doby
            self.pracovni_doba = self.calculate_work_hours(start_time, end_time, lunch_duration)

            # Uložení záznamu
            success = self.excel_manager.ulozit_pracovni_dobu(
                date, start_time, end_time, lunch_duration, self.employee_manager.get_vybrani_zamestnanci()
            )

            if success:
                logger.info(
                    f"Úspěšně uložen záznam: Datum {date}, Začátek {start_time}, "
                    + f"Konec {end_time}, Oběd {lunch_duration}"
                )
                return True, "Záznam byl úspěšně uložen"
            else:
                raise Exception("Nepodařilo se uložit záznam")

        except ValueError as e:
            logger.error(f"Chyba validace: {e}")
            return False, str(e)
        except Exception as e:
            logger.error(f"Nepodařilo se uložit záznam: {e}")
            return False, f"Nepodařilo se uložit záznam: {e}"


"Nepodařilo se uložit záznam: {e}"
