# employee_management.py
import json
import os
from pathlib import Path
from typing import List, Dict, Union
import logging
from utils.logger import setup_logger

logger = setup_logger("employee_management")

class EmployeeManager:
    """Třída pro správu zaměstnanců"""
    
    def __init__(self, data_path: Union[str, Path]):
        """Inicializace správce zaměstnanců"""
        if not isinstance(data_path, (str, Path)):
            raise TypeError("data_path musí být řetězec nebo Path objekt")
        
        if not data_path:
            raise ValueError("data_path nesmí být prázdný")
        
        self.data_path = Path(data_path)
        self.config_file = self.data_path / "employee_config.json"
        self.zamestnanci = []
        self.vybrani_zamestnanci = []
        self.load_config()
        logger.info("Inicializován správce zaměstnanců")

    def load_config(self) -> None:
        """Načte konfiguraci ze souboru"""
        try:
            if not self.config_file.exists():
                logger.warning(f"Konfigurační soubor {self.config_file} nenalezen")
                self.zamestnanci = []
                self.vybrani_zamestnanci = []
                return

            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            self.zamestnanci = config.get('zamestnanci', [])
            self.vybrani_zamestnanci = config.get('vybrani_zamestnanci', [])
            logger.info(f"Načteno {len(self.zamestnanci)} zaměstnanců a {len(self.vybrani_zamestnanci)} vybraných zaměstnanců")
            
        except Exception as e:
            logger.error(f"Chyba při načítání konfigurace: {str(e)}")
            self.zamestnanci = []
            self.vybrani_zamestnanci = []

    def save_config(self) -> bool:
        """Uloží aktuální konfiguraci do souboru"""
        try:
            config_dir = self.data_path
            config_dir.mkdir(parents=True, exist_ok=True)
            
            config = {
                'zamestnanci': self.zamestnanci,
                'vybrani_zamestnanci': self.vybrani_zamestnanci
            }
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
                
            logger.info("Konfigurace zaměstnanců úspěšně uložena")
            return True
            
        except Exception as e:
            logger.error(f"Chyba při ukládání konfigurace: {str(e)}")
            return False

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

    def pridat_zamestnance(self, jmeno: str) -> bool:
        """Přidá nového zaměstnance"""
        try:
            # Validace jména
            self._validate_employee_name(jmeno)
            
            if jmeno in self.zamestnanci:
                logger.warning(f"Zaměstnanec {jmeno} již existuje")
                return False
                
            self.zamestnanci.append(jmeno)
            self.zamestnanci.sort()
            
            if self.save_config():
                logger.info(f"Přidán zaměstnanec: {jmeno}")
                return True
            return False
            
        except ValueError as e:
            logger.error(str(e))
            raise
        except Exception as e:
            logger.error(f"Neočekávaná chyba při přidávání zaměstnance: {str(e)}")
            return False

    def prepinat_zamestnance(self, jmeno: str) -> bool:
        """Přepíná výběr zaměstnance"""
        if jmeno not in self.zamestnanci:
            logger.warning(f"Zaměstnanec {jmeno} neexistuje")
            return False
            
        if jmeno in self.vybrani_zamestnanci:
            self.vybrani_zamestnanci.remove(jmeno)
            logger.info(f"Odebrán vybraný zaměstnanec: {jmeno}")
        else:
            self.vybrani_zamestnanci.append(jmeno)
            self.vybrani_zamestnanci.sort()
            logger.info(f"Přidán vybraný zaměstnanec: {jmeno}")
        
        return self.save_config()

    def upravit_zamestnance(self, stare_jmeno: str, nove_jmeno: str) -> bool:
        """Upraví jméno zaměstnance"""
        try:
            # Validace nového jména
            self._validate_employee_name(nove_jmeno)
            
            if stare_jmeno not in self.zamestnanci:
                raise ValueError(f"Zaměstnanec {stare_jmeno} neexistuje")
                
            if stare_jmeno == nove_jmeno:
                logger.info(f"Zaměstnanec {stare_jmeno} - žádná změna")
                return True
                
            if nove_jmeno in self.zamestnanci:
                raise ValueError(f"Zaměstnanec {nove_jmeno} již existuje")
                
            # Aktualizace jména v seznamu zaměstnanců
            index = self.zamestnanci.index(stare_jmeno)
            self.zamestnanci[index] = nove_jmeno
            
            # Aktualizace jména ve vybraných zaměstnancích
            if stare_jmeno in self.vybrani_zamestnanci:
                self.vybrani_zamestnanci.remove(stare_jmeno)
                self.vybrani_zamestnanci.append(nove_jmeno)
                self.vybrani_zamestnanci.sort()
            
            self.save_config()
            logger.info(f"Zaměstnanec {stare_jmeno} přejmenován na {nove_jmeno}")
            return True
            
        except ValueError as e:
            logger.error(str(e))
            raise
        except Exception as e:
            logger.error(f"Neočekávaná chyba při úpravě zaměstnance: {str(e)}")
            return False

    def smazat_zamestnance(self, index: int) -> bool:
        """Smaže zaměstnance podle indexu"""
        try:
            if not isinstance(index, int):
                raise ValueError("Index musí být celé číslo")
                
            if 1 <= index <= len(self.zamestnanci):
                jmeno = self.zamestnanci.pop(index - 1)
                
                # Pokud byl zaměstnanec vybrán, odstraň jej i odsud
                if jmeno in self.vybrani_zamestnanci:
                    self.vybrani_zamestnanci.remove(jmeno)
                    
                self.save_config()
                logger.info(f"Smazán zaměstnanec: {jmeno}")
                return True
                
            logger.warning(f"Nepodařilo se smazat zaměstnance s indexem: {index}")
            return False
            
        except Exception as e:
            logger.error(f"Neočekávaná chyba při mazání zaměstnance: {str(e)}")
            return False

    def get_all_employees(self) -> List[Dict[str, Union[str, bool]]]:
        """Vrátí seznam všech zaměstnanců s informací o výběru"""
        try:
            if not isinstance(self.zamestnanci, list) or not isinstance(self.vybrani_zamestnanci, list):
                logger.error("Seznamy zaměstnanců nejsou platné")
                return []
                
            return [
                {"name": jmeno, "selected": jmeno in self.vybrani_zamestnanci}
                for jmeno in sorted(self.zamestnanci)
            ]
            
        except Exception as e:
            logger.error(f"Chyba při získávání seznamu zaměstnanců: {str(e)}")
            return []

    def get_vybrani_zamestnanci(self) -> List[str]:
        """Vrátí seznam vybraných zaměstnanců"""
        return sorted(self.vybrani_zamestnanci)

    def pridat_vybraneho_zamestnance(self, jmeno: str) -> bool:
        """Přidá zaměstnance do vybraných"""
        if jmeno not in self.zamestnanci:
            logger.warning(f"Zaměstnanec {jmeno} neexistuje")
            return False
            
        if jmeno in self.vybrani_zamestnanci:
            logger.warning(f"Zaměstnanec {jmeno} je již vybrán")
            return False
            
        self.vybrani_zamestnanci.append(jmeno)
        self.vybrani_zamestnanci.sort()
        self.save_config()
        logger.info(f"Přidán vybraný zaměstnanec: {jmeno}")
        return True

    def odebrat_vybraneho_zamestnance(self, jmeno: str) -> bool:
        """Odebere zaměstnance z vybraných"""
        if jmeno not in self.vybrani_zamestnanci:
            logger.warning(f"Zaměstnanec {jmeno} není ve vybraných")
            return False
            
        self.vybrani_zamestnanci.remove(jmeno)
        self.save_config()
        logger.info(f"Odebrán vybraný zaměstnanec: {jmeno}")
        return True

    def upravit_zamestnance_podle_jmena(self, stare_jmeno: str, nove_jmeno: str) -> bool:
        """Upraví jméno zaměstnance podle jména"""
        return self.upravit_zamestnance(stare_jmeno, nove_jmeno)

    def smazat_zamestnance_podle_jmena(self, jmeno: str) -> bool:
        """Smaže zaměstnance podle jména"""
        if jmeno in self.zamestnanci:
            index = self.zamestnanci.index(jmeno)
            return self.smazat_zamestnance(index + 1)
        return False
