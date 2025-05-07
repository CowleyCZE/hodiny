# utils/voice_processor.py
import os
import re
import json
import requests
import logging
from datetime import datetime, timedelta
from pathlib import Path
from config import Config
from employee_management import EmployeeManager
from utils.logger import setup_logger
from functools import lru_cache
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
from requests_cache import CachedSession
from collections import deque
import time

logger = setup_logger("voice_processor")

class RateLimiter:
    def __init__(self, max_requests, time_window):
        self.max_requests = max_requests
        self.time_window = time_window
        self.requests = deque()

    def can_make_request(self):
        now = time.time()
        while self.requests and self.requests[0] < now - self.time_window:
            self.requests.popleft()
        return len(self.requests) < self.max_requests

    def add_request(self):
        self.requests.append(time.time())

class VoiceProcessor:
    def __init__(self):
        """Inicializace procesoru hlasu"""
        self.gemini_api_key = Config.GEMINI_API_KEY
        self.gemini_api_url = Config.GEMINI_API_URL
        self.employee_manager = EmployeeManager(data_path="data")
        self.employee_list = self._load_employees()
        self.currency_options = ["CZK", "EUR"]
        
        # Inicializace rate limiteru
        self.rate_limiter = RateLimiter(
            Config.RATE_LIMIT_REQUESTS,
            Config.RATE_LIMIT_WINDOW
        )
        
        # Inicializace session pro HTTP požadavky
        self.session = requests.Session()

    def init_cache_session(self):
        """Inicializace cache session pro produkční prostředí"""
        if not isinstance(self.session, CachedSession):
            self.session = CachedSession(
                'gemini_cache',
                expire_after=Config.GEMINI_CACHE_TTL,
                allowable_methods=['GET', 'POST'],
                stale_if_error=True
            )

    @lru_cache(maxsize=100)
    def _load_employees(self):
        """Načte seznam zaměstnanců z konfiguračního souboru s cachováním"""
        try:
            return self.employee_manager.get_all_employees()
        except Exception as e:
            logger.error(f"Chyba při načítání zaměstnanců: {e}")
            return []

    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=4, max=10),
        retry=retry_if_exception_type((requests.exceptions.RequestException, Exception))
    )
    def _call_gemini_api(self, audio_file_path):
        """
        Komunikace s Gemini API pro transkripci a analýzu hlasu s retry mechanismem
        a rate limitingem
        """
        try:
            if not os.path.exists(audio_file_path):
                raise FileNotFoundError(f"Audio soubor nebyl nalezen: {audio_file_path}")

            # Kontrola rate limitu
            if not self.rate_limiter.can_make_request():
                return {"error": "Překročen rate limit pro API požadavky"}

            self.rate_limiter.add_request()

            with open(audio_file_path, "rb") as audio_file:
                files = {"audio": audio_file}
                headers = {
                    "Authorization": f"Bearer {self.gemini_api_key}",
                    "X-Request-ID": str(datetime.now().timestamp())
                }
                
                # Použití základní session pro testy nebo cache session pro produkci
                response = self.session.post(
                    self.gemini_api_url,
                    headers=headers,
                    files=files,
                    timeout=30
                )
                
                response.raise_for_status()
                return response.json()

        except requests.exceptions.RequestException as e:
            logger.error(f"Síťová chyba při volání Gemini API: {e}", exc_info=True)
            raise
        except Exception as e:
            logger.error(f"Neočekávaná chyba při volání Gemini API: {e}", exc_info=True)
            return {"error": f"API volání selhalo: {str(e)}"}

    def _extract_entities(self, text):
        """Extrahuje entity z textu pomocí regulárních výrazů"""
        text = text.lower()
        entities = {
            "employee": None,
            "date": None,
            "amount": None,
            "currency": None,
            "action": None,
            "time_period": None,
            "start_time": None,
            "end_time": None
        }

        # Detekce akce - musí být první
        action_patterns = {
            "add_advance": [
                r"záloh[au]",
                r"přid[aáe][tj]\s*záloh[au]",
                r"vypl[aá][tc]?\s*záloh[au]",
                r"nov[áé]\s*záloh[au]"
            ],
            "record_time": [
                r"práce",
                r"pracovní\s*dob[au]",
                r"odpracovan[éý]",
                r"zapiš\s*(?:čas|hodiny)",
                r"zaznamenej.*dob[au]",
                r"přidej.*(?:čas|hodiny|dobu)"
            ],
            "get_stats": [
                r"statistik[ay]",
                r"přehled",
                r"součet",
                r"celkem",
                r"ukaž",
                r"zobraz"
            ]
        }

        for action, patterns in action_patterns.items():
            if any(re.search(pattern, text, re.IGNORECASE) for pattern in patterns):
                entities["action"] = action
                break

        # Extrakce časů pro záznam pracovní doby
        if entities["action"] == "record_time":
            # Rozpoznání časů ve formátu "od X do Y" nebo "X-Y"
            time_patterns = [
                # Od sedmi do osmi
                (r"od\s+(\d{1,2}|sedmi|osmi|devíti|desíti|jedenácti|dvanácti|třinácti|čtrnácti|patnácti|šestnácti|sedmnácti|osmnácti|devatenácti|dvaceti|dvaceti ?jedné|dvaceti ?dvou|dvaceti ?tří)\s+(?:hodin(?:y)?)?(?:\s+)?do\s+(\d{1,2}|sedmi|osmi|devíti|desíti|jedenácti|dvanácti|třinácti|čtrnácti|patnácti|šestnácti|sedmnácti|osmnácti|devatenácti|dvaceti|dvaceti ?jedné|dvaceti ?dvou|dvaceti ?tří)(?:\s+hodin)?", lambda x, y: (self._word_to_hour(x), self._word_to_hour(y))),
                # 7-17 nebo 7:00-17:00
                (r"(\d{1,2})(?::\d{2})?\s*[-–]\s*(\d{1,2})(?::\d{2})?", lambda x, y: (int(x), int(y))),
                # Od 7 do 17
                (r"od\s+(\d{1,2})(?::\d{2})?\s+do\s+(\d{1,2})(?::\d{2})?", lambda x, y: (int(x), int(y)))
            ]

            for pattern, converter in time_patterns:
                match = re.search(pattern, text)
                if match:
                    start, end = converter(match.group(1), match.group(2))
                    if 0 <= start <= 23 and 0 <= end <= 23:
                        entities["start_time"] = f"{start:02d}:00"
                        entities["end_time"] = f"{end:02d}:00"
                        break

        # Extrakce data
        date_patterns = [
            (r"\bdnes\b", datetime.now()),
            (r"\bvčera\b", datetime.now() - timedelta(days=1)),
            (r"\bzítra\b", datetime.now() + timedelta(days=1)),
            (r"\b(\d{4}-\d{2}-\d{2})\b", None),
            (r"\b(\d{1,2}\.\s*\d{1,2}\.\s*\d{4})\b", None),
            (r"\b(\d{1,2}/\d{1,2}/\d{4})\b", None),
        ]

        for pattern, date_value in date_patterns:
            match = re.search(pattern, text)
            if match:
                if date_value:
                    entities["date"] = date_value.strftime("%Y-%m-%d")
                else:
                    date_str = match.group(1).replace(" ", "")
                    entities["date"] = self._normalize_date(date_str)
                break

        # Pokud není zadáno datum, použijeme dnešek
        if not entities["date"] and entities["action"] == "record_time":
            entities["date"] = datetime.now().strftime("%Y-%m-%d")

        # Extrakce zaměstnance
        employees = self.employee_manager.get_selected_employees()
        if not employees:  # Pokud nejsou vybraní, použijeme všechny
            employees = self._load_employees()
        
        for employee in employees:
            emp_pattern = r"\b" + re.escape(employee.lower()) + r"\b"
            if re.search(emp_pattern, text.lower()):
                entities["employee"] = employee
                break

        # Extrakce částky a měny
        amount_match = re.search(r"(\d+(?:[.,]\d+)?)\s*(czk|kč|eur|€)", text, re.IGNORECASE)
        if amount_match:
            amount = float(amount_match.group(1).replace(",", "."))
            currency = amount_match.group(2).lower()
            currency_map = {
                "kč": "CZK", 
                "czk": "CZK",
                "€": "EUR",
                "eur": "EUR"
            }
            entities["amount"] = amount
            entities["currency"] = currency_map.get(currency, currency.upper())

        # Extrakce časového období
        period_patterns = {
            "week": [r"týden", r"týdně", r"týdenní"],
            "month": [r"měsíc", r"měsíční", r"měsíčně"],
            "year": [r"rok", r"roční", r"ročně"]
        }

        for period, patterns in period_patterns.items():
            if any(re.search(pattern, text, re.IGNORECASE) for pattern in patterns):
                entities["time_period"] = period
                break

        return entities

    def _word_to_hour(self, word):
        """Převede slovní vyjádření hodiny na číslo"""
        word = word.lower().strip()
        if word.isdigit():
            return int(word)
            
        hour_map = {
            "sedmi": 7, "osmi": 8, "devíti": 9, "desíti": 10,
            "jedenácti": 11, "dvanácti": 12, "třinácti": 13,
            "čtrnácti": 14, "patnácti": 15, "šestnácti": 16,
            "sedmnácti": 17, "osmnácti": 18, "devatenácti": 19,
            "dvaceti": 20, "dvaceti jedné": 21, "dvacetijedné": 21,
            "dvaceti dvou": 22, "dvacetidvou": 22,
            "dvaceti tří": 23, "dvacetitří": 23
        }
        
        return hour_map.get(word, 0)

    def _normalize_date(self, date_str):
        """Převede různé formáty data na standardní formát YYYY-MM-DD"""
        for fmt in ["%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"]:
            try:
                return datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        return None

    def _validate_data(self, data):
        """Validace extrahovaných dat"""
        errors = []
        
        # Validace zaměstnance pro relevantní akce
        if data.get("action") in ["record_time", "add_advance"]:
            if not data.get("employee"):
                errors.append("Neznámý zaměstnanec")
        
        # Validace data
        if data.get("date") and not re.match(r"\d{4}-\d{2}-\d{2}", data["date"]):
            errors.append("Neplatný formát data")
        
        # Validace částky a měny pro zálohy
        if data.get("action") == "add_advance":
            if not data.get("amount") or data["amount"] <= 0:
                errors.append("Částka musí být kladná")
            if not data.get("currency") or data["currency"] not in self.currency_options:
                errors.append(f"Neplatná měna. Použijte {', '.join(self.currency_options)}")
        
        # Validace akce
        if not data.get("action"):
            errors.append("Neznámá akce")
        
        return len(errors) == 0, errors

    def process_voice_command(self, audio_file_path):
        """
        Zpracuje hlasový příkaz s vylepšeným error handlingem a retry logikou
        """
        try:
            # 1. Převod hlasu na text přes Gemini API
            gemini_response = self._call_gemini_api(audio_file_path)
            
            if "error" in gemini_response:
                logger.error(f"Chyba v Gemini API odpovědi: {gemini_response['error']}")
                return {"success": False, "error": gemini_response["error"]}
            
            # 2. Extrahujeme entity z textu
            text = gemini_response.get("text", "")
            if not text:
                return {"success": False, "error": "Prázdná odpověď od Gemini API"}
                
            entities = self._extract_entities(text)
            
            # 3. Validace dat
            is_valid, validation_errors = self._validate_data(entities)
            
            if not is_valid:
                return {
                    "success": False, 
                    "errors": validation_errors,
                    "original_text": text
                }
            
            # 4. Přidání dodatečných informací
            entities.update({
                "success": True,  # Explicitně nastavíme na True když validace prošla
                "confidence": gemini_response.get("confidence", 0.8),
                "processed_at": datetime.now().isoformat(),
                "original_text": text
            })
            
            return entities
            
        except Exception as e:
            logger.error(f"Kritická chyba při zpracování hlasového příkazu: {e}", exc_info=True)
            return {
                "success": False, 
                "error": "Interní chyba při zpracování",
                "details": str(e)
            }

    def process_voice_text(self, text):
        """
        Zpracuje textový příkaz podobně jako hlasový příkaz, ale přeskočí krok s Gemini API
        """
        try:
            if not text:
                return {"success": False, "error": "Prázdný textový vstup"}
                
            # Extrahujeme entity z textu
            entities = self._extract_entities(text)
            
            # Validace dat
            is_valid, validation_errors = self._validate_data(entities)
            
            if not is_valid:
                return {
                    "success": False, 
                    "errors": validation_errors,
                    "original_text": text
                }
            
            # Přidání dodatečných informací
            entities.update({
                "success": True,
                "confidence": 1.0,  # Pro textový vstup máme 100% jistotu textu
                "processed_at": datetime.now().isoformat(),
                "original_text": text
            })
            
            return entities
            
        except Exception as e:
            logger.error(f"Kritická chyba při zpracování textového příkazu: {e}", exc_info=True)
            return {
                "success": False, 
                "error": "Interní chyba při zpracování",
                "details": str(e)
            }
