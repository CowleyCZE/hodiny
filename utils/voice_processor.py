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

logger = setup_logger("voice_processor")

class VoiceProcessor:
    def __init__(self):
        """Inicializace procesoru hlasu"""
        self.gemini_api_key = Config.GEMINI_API_KEY
        self.employee_manager = EmployeeManager(data_path="data/employee_config.json")
        self.employee_list = self._load_employees()
        self.currency_options = ["CZK", "EUR"]
        
    def _load_employees(self):
        """Načte seznam zaměstnanců z konfiguračního souboru"""
        try:
            return self.employee_manager.get_all_employees()
        except Exception as e:
            logger.error(f"Chyba při načítání zaměstnanců: {e}")
            return []

    def _call_gemini_api(self, audio_file_path):
        """
        Komunikace s Gemini API pro transkripci a analýzu hlasu
        Vrátí strukturovaná data s intencí a entitami
        """
        try:
            # Simulace Gemini API volání
            # Ve skutečnosti by se použil skutečný API endpoint
            with open(audio_file_path, "rb") as audio_file:
                files = {"audio": audio_file}
                data = {"api_key": self.gemini_api_key}
                
                # Toto je simulace - ve skutečnosti by se použil:
                # response = requests.post(
                #     "https://gemini-api-endpoint.com/speech-to-text",
                #     headers={"Authorization": f"Bearer {self.gemini_api_key}"},
                #     files=files,
                #     data=data
                # )
                
                # Ukázková odpověď pro testování
                return {
                    "intent": "record_time",
                    "entities": {
                        "employee": "Jan Novák",
                        "date": "2025-03-15",
                        "start_time": "08:00",
                        "end_time": "16:00"
                    },
                    "confidence": 0.92
                }
                
        except Exception as e:
            logger.error(f"Chyba při volání Gemini API: {e}")
            return {"error": "API volání selhalo"}

    def _extract_entities(self, text):
        """
        Extrahuje entity z textu pomocí regulárních výrazů
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

    def process_voice_command(self, audio_file_path):
        """
        Zpracuje hlasový příkaz
        Vrátí strukturovaná data pro aplikaci
        """
        try:
            # 1. Převod hlasu na text přes Gemini API
            gemini_response = self._call_gemini_api(audio_file_path)
            
            if "error" in gemini_response:
                return {"success": False, "error": gemini_response["error"]}
                
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
