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
        self.default_lunch_duration = 1.0  # Přednastavená délka oběda na 1 hodinu
        
        # Inicializace rate limiteru
        self.rate_limiter = RateLimiter(
            Config.RATE_LIMIT_REQUESTS,
            Config.RATE_LIMIT_WINDOW
        )
        
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
                    timeout=Config.GEMINI_REQUEST_TIMEOUT
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
            "date": None,
            "start_time": None,
            "end_time": None,
            "lunch_duration": self.default_lunch_duration,
            "action": None,
            "is_free_day": False
        }

        # Detekce akce
        action_patterns = {
            "record_time": [
                r"práce",
                r"pracovní\s*dob[au]",
                r"odpracovan[éý]",
                r"zapiš\s*(?:čas|hodiny)",
                r"zaznamenej.*dob[au]",
                r"přidej.*(?:čas|hodiny|dobu)"
            ],
            "record_free_day": [
                r"voln[oý]",
                r"dovolen[áa]",
                r"sick\s*day",
                r"náhradní\s*volno",
                r"nepřítomnost"
            ],
            "get_stats": [
                r"statistik[ay]",
                r"ukaž\s*(mi)?\s*statistik[uy]",
                r"jak[ée]\s*jsou\s*statistik[ay]",
                r"přehled"
            ]
        }

        # Priorita akcí - get_stats má vyšší prioritu
        # Pokud je nalezena "get_stats", nastaví se a ostatní se přeskočí.
        # Jinak se pokračuje s ostatními akcemi.
        # Toto je zjednodušené řešení konfliktů.
        if "get_stats" in action_patterns:
            if any(re.search(pattern, text, re.IGNORECASE) for pattern in action_patterns["get_stats"]):
                entities["action"] = "get_stats"

        if not entities["action"]: # Pokud nebyla detekována akce get_stats, zkusíme ostatní
            for action_key, patterns in action_patterns.items():
                if action_key == "get_stats": # Tuto akci jsme již zkontrolovali
                    continue
                if any(re.search(pattern, text, re.IGNORECASE) for pattern in patterns):
                    # Pro record_time a record_free_day je základní akce "record_time"
                    entities["action"] = "record_time"
                    if action_key == "record_free_day":
                        entities["is_free_day"] = True
                        entities["start_time"] = "00:00"
                        entities["end_time"] = "00:00"
                        entities["lunch_duration"] = 0.0
                    break # Našli jsme akci, můžeme přerušit

        # Pokud je akce 'get_stats', extrahujeme specifické entity
        if entities["action"] == "get_stats":
            # Extrakce časového období
            time_period_patterns = {
                "week": [r"týden", r"týdenní"],
                "month": [r"měsíc", r"měsíční"],
                "year": [r"rok", r"roční"]
            }
            for period_key, patterns in time_period_patterns.items():
                if any(re.search(pattern, text, re.IGNORECASE) for pattern in patterns):
                    entities["time_period"] = period_key
                    break
            
            # Extrakce jména zaměstnance (zjednodušený přístup)
            # Načteme seznam zaměstnanců (jména jsou ve formátu "Příjmení Jméno")
            employee_list_dicts = self._load_employees() # Vrací seznam slovníků [{'name': 'Jméno Příjmení', 'selected': True/False}]
            employee_names = [emp['name'] for emp in employee_list_dicts]

            for name in employee_names:
                # Vytvoříme regex pro jméno, case-insensitive
                # Jméno může být víceslovné, např. "Jan Novák"
                # Musíme ošetřit speciální znaky v regexu, pokud by jména obsahovala např. tečky
                safe_name_pattern = re.escape(name)
                # Hledáme jméno v kontextu statistik, např. "statistiky pro Jan Novák", "Jan Novák přehled"
                # nebo jen samotné jméno, pokud je kontext jasný z akce "get_stats"
                patterns_with_name = [
                    rf"pro\s+{safe_name_pattern}",
                    rf"{safe_name_pattern}\s*statistik[ay]",
                    rf"{safe_name_pattern}\s*přehled",
                    rf"přehled\s*(?:pro\s*)?{safe_name_pattern}",
                    rf"statistik[ay]\s*(?:pro\s*)?{safe_name_pattern}"
                ]
                # Přidáme i samotné jméno jako vzor, ale s nižší prioritou (pokud ostatní selžou)
                # To by mohlo být problematické, pokud jméno je běžné slovo.
                # Prozatím se držíme kontextových vzorů.
                # Pokud by se mělo hledat jen samotné jméno:
                # if re.search(rf"\b{safe_name_pattern}\b", text, re.IGNORECASE):
                # entities["employee"] = name
                # break

                if any(re.search(pattern, text, re.IGNORECASE) for pattern in patterns_with_name):
                    entities["employee"] = name
                    logger.info(f"Nalezen zaměstnanec '{name}' pro statistiky.")
                    break 
            if not entities.get("employee"):
                logger.info("Pro akci 'get_stats' nebyl specifikován/nalezen žádný konkrétní zaměstnanec.")


        # Pokud je to volný den (a akce není get_stats), přeskočíme extrakci časů
        if not entities["is_free_day"] and entities["action"] != "get_stats":
            # Extrakce časů pro pracovní dobu
            time_patterns = [
                # Od sedmi do osmi
                (r"od\s+(\d{1,2}|sedmi|osmi|devíti|desíti|jedenácti|dvanácti|třinácti|čtrnácti|patnácti|šestnácti|sedmnácti|osmnácti|devatenácti|dvaceti|dvaceti ?jedné|dvaceti ?dvou|dvaceti ?tří)\s+(?:hodin(?:y)?)?(?:\s+)?do\s+(\d{1,2}|sedmi|osmi|devíti|desíti|jedenácti|dvanácti|třinácti|čtrnácti|patnácti|šestnácti|sedmnácti|osmnácti|devatenácti|dvaceti|dvaceti ?jedné|dvaceti ?dvou|dvaceti ?tří)", lambda x, y: (self._word_to_hour(x), self._word_to_hour(y))),
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

        # Pro volný den není délka oběda potřeba
        if not entities["is_free_day"]:
            # Extrakce délky oběda (pokud je explicitně uvedena)
            lunch_match = re.search(r"ob[ěe]d\s+(\d+(?:[.,]\d+)?)\s*(?:hodiny?|hodin|h)?", text)
            if lunch_match:
                lunch_duration = float(lunch_match.group(1).replace(",", "."))
                if 0 <= lunch_duration <= 4:  # Kontrola rozumného rozsahu
                    entities["lunch_duration"] = lunch_duration

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
        
        # Validace data
        if data.get("date") and not re.match(r"\d{4}-\d{2}-\d{2}", data["date"]):
            errors.append("Neplatný formát data")
        
        # Validace časů pro záznam pracovní doby
        if data.get("action") == "record_time":
            if not data.get("start_time"):
                errors.append("Chybí čas začátku")
            if not data.get("end_time"):
                errors.append("Chybí čas konce")
            if data.get("lunch_duration") is not None and (data["lunch_duration"] < 0 or data["lunch_duration"] > 4):
                errors.append("Délka oběda musí být mezi 0 a 4 hodinami")
        
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
        Zpracuje textový příkaz podobně jako hlasový příkaz, ale přeskočí krok s Gemini API.
        Používá stejné metody pro extrakci entit a validaci jako hlasový vstup.
        
        Args:
            text (str): Textový příkaz k zpracování
            
        Returns:
            dict: Výsledek zpracování s extrahovanými entitami nebo chybou
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
            result = {
                "success": True,
                "entities": entities,
                "confidence": 1.0,  # Pro textový vstup máme 100% jistotu textu
                "processed_at": datetime.now().isoformat(),
                "original_text": text
            }
            
            return result
            
        except Exception as e:
            logger.error(f"Kritická chyba při zpracování textového příkazu: {e}", exc_info=True)
            return {
                "success": False, 
                "error": "Interní chyba při zpracování",
                "details": str(e)
            }
