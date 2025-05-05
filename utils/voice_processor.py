# voice_processor.py
import os
import re
import json
import logging
import google.generativeai as genai
from datetime import datetime, timedelta
from pathlib import Path
from config import Config
from employee_management import EmployeeManager
from utils.logger import setup_logger

logger = setup_logger("voice_processor")

class VoiceProcessor:
    def __init__(self):
        """Inicializace procesoru hlasu"""
        self.gemini_api_key = Config.GEMINI_API_KEY
        self.employee_manager = EmployeeManager(data_path="data/employee_config.json")
        self.employee_list = self._load_employees()
        self.currency_options = ["CZK", "EUR"]
        
        # Konfigurace Gemini API
        try:
            genai.configure(api_key=self.gemini_api_key)
            self.model = genai.GenerativeModel("gemini-pro")  # Použití Gemini Pro
            logger.info("Gemini API úspěšně nakonfigurováno")
        except Exception as e:
            logger.error(f"Chyba při konfiguraci Gemini API: {e}")
            self.model = None

    def _load_robots(self):
        """Načte seznam zaměstnanců z konfiguračního souboru"""
        try:
            return self.employee_manager.get_all_employees()
        except Exception as e:
            logger.error(f"Chyba při načítání zaměstnanců: {e}")
            return []

    def _call_gemini_stt(self, audio_file_path):
        """
        Komunikace s Gemini API pro transkripci hlasu
        Vrátí strukturovaná data s intencí a entitami
        """
        try:
            # Nahrání zvukového souboru do Gemini
            uploaded_file = genai.upload_file(path=audio_file_path, display_name="voice_command")
            logger.info(f"Soubor nahrán do Gemini: {uploaded_file.name}")
            
            # Prompt pro extrakci textu a entit
            prompt = """Převeď následující hlasový vstup na text a extrahuj entity:
                - Zaměstnanec (např. "Jan Novák")
                - Datum (např. "dnes", "včera", "2025-03-15")
                - Čas (např. "08:00", "osm hodin")
                - Částka (např. "500 Kč", "200 €")
                - Akce (např. "zaznamenat pracovní dobu", "přidat zálohu")
                - Časové období (např. "týden", "měsíc")

                Odpověď vrať ve formátu JSON s klíči:
                {
                    "intent": "record_time | add_advance | get_stats",
                    "entities": {
                        "employee": "Jméno zaměstnance",
                        "date": "Datum v formátu YYYY-MM-DD",
                        "start_time": "Čas začátku práce",
                        "end_time": "Čas konce práce",
                        "amount": "Částka zálohy",
                        "currency": "Měna (CZK/EUR)",
                        "time_period": "Časové období"
                    },
                    "confidence": "Spolehlivost rozpoznání (0.0-1.0)"
                }
            """
            
            # Analyza pomocí Gemini API
            if not self.model:
                raise Exception("Gemini model není k dispozici")
                
            response = self.model.generate_content([prompt, uploaded_file])
            
            # Získání JSON odpovědi
            response_text = response.text.strip()
            if response_text.startswith("```json"):
                response_text = response_text[7:-3].strip()  # Odstranění markdown zápisu
            
            result = json.loads(response_text)
            return result
            
        except json.JSONDecodeError as e:
            logger.error(f"Chyba při parsování JSON odpovědi: {e}")
            return {"success": False, "error": "Neplatný formát odpovědi Gemini API"}
        except Exception as e:
            logger.error(f"Chyba při volání Gemini API: {e}", exc_info=True)
            return {"success": False, "error": "API volání selhalo"}

    def _extract_entities(self, text):
        """
        Extrahuje entity z textu pomocí regulárních výrazů (záložní metoda)
        """
        entities = {
            "employee": None,
            "date": None,
            "amount": None,
            "currency": None,
            "action": None,
            "time_period": None
        }

        # Extrakce zaměstnance
        for employee in self.employee_list:
            if employee.lower() in text.lower():
                entities["employee"] = employee
                break

        # Extrakce data
        date_match = re.search(r"(\d{4}-\d{2}-\d{2}|\d{2}\.\d{2}\.\d{4}|\d{2}\/\d{2}\/\d{4})|dnes|včera", text)
        if date_match:
            if date_match.group(0) == "dnes":
                entities["date"] = datetime.now().strftime("%Y-%m-%d")
            elif date_match.group(0) == "včera":
                entities["date"] = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
            else:
                entities["date"] = self._normalize_date(date_match.group(0))

        # Extrakce částky
        amount_match = re.search(r"(\d+[\.,]?\d*)\s?(CZK|EUR|Kč|€)", text)
        if amount_match:
            entities["amount"] = float(amount_match.group(1).replace(",", "."))
            currency_map = {"Kč": "CZK", "€": "EUR"}
            entities["currency"] = currency_map.get(amount_match.group(2), amount_match.group(2))

        # Extrakce akce
        if any(word in text.lower() for word in ["záloha", "přidat", "přidej"]):
            entities["action"] = "add_advance"
        elif any(word in text.lower() for word in ["vykaz", "práce", "odpracoval"]):
            entities["action"] = "record_time"
        elif any(word in text.lower() for word in ["statistika", "součet", "celkem"]):
            entities["action"] = "get_stats"

        # Extrakce časového období
        period_match = re.search(r"(týden|měsíc|rok)", text)
        if period_match:
            entities["time_period"] = period_match.group(0)

        return entities

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
        
        if not data.get("employee"):
            errors.append("Neznámý zaměstnanec")
            
        if data.get("date") and not re.match(r"\d{4}-\d{2}-\d{2}", data["date"]):
            errors.append("Neplatný formát data")
            
        if data.get("amount") and data["amount"] <= 0:
            errors.append("Částka musí být kladná")
            
        if data.get("currency") and data["currency"] not in self.currency_options:
            errors.append(f"Neplatná měna. Použijte {', '.join(self.currency_options)}")
            
        if not data.get("action"):
            errors.append("Neznámá akce")
            
        return len(errors) == 0, errors

    def process_voice_audio(self, audio_file_path):
        """
        Zpracuje hlasový příkaz přes Gemini API
        Vrátí strukturovaná data pro aplikaci
        """
        try:
            # 1. Převod hlasu na text přes Gemini
            gemini_response = self._call_gemini_stt(audio_file_path)
            
            if "error" in gemini_response:
                return gemini_response
                
            # 2. Extrahujeme entity z textu
            text = gemini_response.get("text", "")
            entities = self._extract_entities(text)
            
            # 3. Validace dat
            is_valid, validation_errors = self._validate_data(entities)
            
            if not is_valid:
                return {"success": False, "errors": validation_errors}
                
            # 4. Přidání informace o úspěšnosti
            entities["success"] = True
            entities["confidence"] = gemini_response.get("confidence", 0.8)
            
            return entities
            
        except Exception as e:
            logger.error(f"Chyba při zpracování hlasového příkazu: {e}", exc_info=True)
            return {"success": False, "error": "Interní chyba při zpracování"}

    def process_voice_text(self, text):
        """
        Zpracuje textový příkaz (záložní metoda)
        """
        try:
            # 1. Extrahujeme entity z textu
            entities = self._extract_entities(text)
            
            # 2. Validace dat
            is_valid, validation_errors = self._validate_data(entities)
            
            if not is_valid:
                return {"success": False, "errors": validation_errors}
                
            # 3. Přidání informace o úspěšnosti
            entities["success"] = True
            entities["confidence"] = 0.9  # Výchozí spolehlivost pro text
            
            return entities
            
        except Exception as e:
            logger.error(f"Chyba při zpracování textového příkazu: {e}", exc_info=True)
            return {"success": False, "error": "Interní chyba při zpracování"}
