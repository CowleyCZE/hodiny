"""VoiceProcessor: extrakce strukturovaných příkazů (čas, datum, akce) z řeči / textu.

Funkce:
 - volání (volitelného) Gemini API s retry + rate limiting
 - regulární extrakce entit (start/end, oběd, datum, typ akce, zaměstnanec)
 - validace dat a jednotný výstup vhodný pro další zpracování
"""

import os
import re
import time
from collections import deque
from datetime import datetime, timedelta
from functools import lru_cache

import requests
from requests_cache import CachedSession
from tenacity import retry, retry_if_exception_type, stop_after_attempt, wait_exponential

from config import Config
from employee_management import EmployeeManager
from utils.logger import setup_logger

logger = setup_logger("voice_processor")


class RateLimiter:
    """Simple sliding window rate limiter (in-memory)."""

    def __init__(self, max_requests, time_window):
        self.max_requests = max_requests
        self.time_window = time_window
        self.requests = deque()

    def can_make_request(self):
        """Vrátí True pokud lze provést další požadavek."""
        now = time.time()
        while self.requests and self.requests[0] < now - self.time_window:
            self.requests.popleft()
        return len(self.requests) < self.max_requests

    def add_request(self):
        """Zaloguje timestamp provedeného požadavku."""
        self.requests.append(time.time())


class VoiceProcessor:
    """Zpracovává hlasové nebo textové příkazy a extrahuje z nich entity."""

    def __init__(self):
        self.gemini_api_key = getattr(Config, "GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY"))
        self.gemini_api_url = getattr(Config, "GEMINI_API_URL", os.environ.get("GEMINI_API_URL"))
        self.employee_manager = EmployeeManager(data_path="data")
        self.default_lunch_duration = 1.0
        self.rate_limiter = RateLimiter(
            getattr(Config, "RATE_LIMIT_REQUESTS", 5),
            getattr(Config, "RATE_LIMIT_WINDOW", 60),
        )
        self.session = requests.Session()

    def init_cache_session(self):
        """Inicializace cache session (idempotentní)."""
        if not isinstance(self.session, CachedSession):
            self.session = CachedSession(
                "gemini_cache",
                expire_after=getattr(Config, "GEMINI_CACHE_TTL", 300),
                allowable_methods=["GET", "POST"],
                stale_if_error=True,
            )

    @lru_cache(maxsize=100)
    def _load_employees(self):
        """Načte seznam zaměstnanců s cachováním."""
        try:
            return self.employee_manager.get_all_employees()
        except Exception as e:
            logger.error("Chyba při načítání zaměstnanců: %s", e)
            return []

    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=4, max=10),
        retry=retry_if_exception_type(requests.exceptions.RequestException),
    )
    def _call_gemini_api(self, audio_file_path):
        """Volání Gemini API (transkripce) s rate limiting a retry."""
        if not os.path.exists(audio_file_path):
            raise FileNotFoundError(f"Audio soubor nebyl nalezen: {audio_file_path}")
        if not self.rate_limiter.can_make_request():
            return {"error": "Překročen rate limit pro API požadavky"}
        self.rate_limiter.add_request()

        with open(audio_file_path, "rb") as audio_file:
            files = {"audio": audio_file}
            headers = {"Authorization": f"Bearer {self.gemini_api_key}"}
            if not self.gemini_api_url:
                return {"error": "Gemini API URL není nakonfigurována"}

            try:
                response = self.session.post(
                    self.gemini_api_url,
                    headers=headers,
                    files=files,
                    timeout=getattr(Config, "GEMINI_REQUEST_TIMEOUT", 30),
                )
                response.raise_for_status()
                return response.json()
            except requests.exceptions.RequestException as e:
                logger.error("Síťová chyba při volání Gemini API: %s", e, exc_info=True)
                raise
            except Exception as e:
                logger.error("Neočekávaná chyba při volání Gemini API: %s", e, exc_info=True)
                return {"error": f"API volání selhalo: {e}"}

    def _extract_entities(self, text):
        """Regex extrakce entit z textu."""
        text = text.lower()
        action = self._extract_action(text)

        entities = {
            "date": self._extract_date(text),
            "start_time": None,
            "end_time": None,
            "lunch_duration": self.default_lunch_duration,
            "action": action,
            "is_free_day": False,
            "employee": None,
            "time_period": None,
        }

        if action == "record_free_day":
            entities.update({"is_free_day": True, "start_time": "00:00", "end_time": "00:00", "lunch_duration": 0.0})
        elif action == "record_time":
            entities.update(self._extract_time(text))
            entities["lunch_duration"] = self._extract_lunch(text) or self.default_lunch_duration
        elif action == "get_stats":
            entities["employee"] = self._extract_employee(text)
            entities["time_period"] = self._extract_time_period(text)

        if not entities["date"] and action in ["record_time", "record_free_day"]:
            entities["date"] = datetime.now().strftime("%Y-%m-%d")

        return entities

    def _extract_action(self, text):
        action_patterns = {
            "get_stats": [r"statistik[ay]", r"přehled"],
            "record_free_day": [r"voln[oý]", r"dovolen[áa]", r"sick\s*day", r"nepřítomnost"],
            "record_time": [r"práce", r"pracovní\s*dob[au]", r"zapiš", r"zaznamenej"],
        }
        for action, patterns in action_patterns.items():
            if any(re.search(p, text, re.IGNORECASE) for p in patterns):
                return action
        return None

    def _extract_time(self, text):
        time_patterns = [
            r"od\s+(\d{1,2}):\d{2}\s+do\s+(\d{1,2}):\d{2}",
            r"(\d{1,2}):\d{2}\s*-\s*(\d{1,2}):\d{2}",
        ]
        for pattern in time_patterns:
            match = re.search(pattern, text)
            if match:
                start, end = int(match.group(1)), int(match.group(2))
                if 0 <= start <= 23 and 0 <= end <= 23:
                    return {"start_time": f"{start:02d}:00", "end_time": f"{end:02d}:00"}
        return {}

    def _extract_lunch(self, text):
        match = re.search(r"ob[ěe]d\s+(\d+(?:[.,]\d+)?)\s*h", text)
        if match:
            try:
                duration = float(match.group(1).replace(",", "."))
                return duration if 0 <= duration <= 4 else None
            except ValueError:
                return None
        return None

    def _extract_date(self, text):
        date_patterns = {
            r"\bdnes\b": lambda: datetime.now(),
            r"\bvčera\b": lambda: datetime.now() - timedelta(days=1),
            r"\bzítra\b": lambda: datetime.now() + timedelta(days=1),
        }
        for pattern, func in date_patterns.items():
            if re.search(pattern, text):
                return func().strftime("%Y-%m-%d")

        date_str_match = re.search(r"\b(\d{1,2}[./]\d{1,2}[./]\d{4})\b", text)
        if date_str_match:
            return self._normalize_date(date_str_match.group(1))
        return None

    def _extract_employee(self, text):
        employee_names = [emp["name"] for emp in self._load_employees()]
        for name in employee_names:
            if re.search(re.escape(name), text, re.IGNORECASE):
                logger.info("Nalezen zaměstnanec '%s' pro statistiky.", name)
                return name
        return None

    def _extract_time_period(self, text):
        periods = {"week": r"týden", "month": r"měsíc", "year": r"rok"}
        for period, pattern in periods.items():
            if re.search(pattern, text, re.IGNORECASE):
                return period
        return None

    def _normalize_date(self, date_str):
        for fmt in ("%d.%m.%Y", "%d/%m/%Y"):
            try:
                return datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        return None

    def _validate_data(self, data):
        if not data.get("action"):
            return False, ["Neznámá akce"]
        if data["action"] == "record_time" and not (data.get("start_time") and data.get("end_time")):
            return False, ["Chybí čas začátku nebo konce"]
        return True, []

    def process_command(self, text=None, audio_file_path=None):
        """Hlavní metoda pro zpracování příkazu (text nebo audio)."""
        try:
            if audio_file_path:
                api_response = self._call_gemini_api(audio_file_path)
                if "error" in api_response:
                    return {"success": False, "error": api_response["error"]}
                text = api_response.get("text", "")

            if not text:
                return {"success": False, "error": "Prázdný vstupní text"}

            entities = self._extract_entities(text)
            is_valid, errors = self._validate_data(entities)
            if not is_valid:
                return {"success": False, "errors": errors, "original_text": text}

            entities.update(
                {
                    "success": True,
                    "processed_at": datetime.now().isoformat(),
                    "original_text": text,
                }
            )
            return entities
        except Exception as e:
            logger.error("Kritická chyba při zpracování příkazu: %s", e, exc_info=True)
            return {"success": False, "error": "Interní chyba při zpracování", "details": str(e)}
