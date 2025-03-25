import os
import json
from pathlib import Path
from utils.logger import setup_logger

logger = setup_logger('employee_management')

class EmployeeManager:
    def __init__(self, data_path):
        if not isinstance(data_path, (str, Path)):
            raise TypeError("data_path musí být řetězec nebo Path objekt")
        if isinstance(data_path, Path):
            data_path = str(data_path)
        if not data_path:
            raise ValueError("data_path nesmí být prázdný")
            
        self.data_path = data_path
        self.zamestnanci = []
        self.vybrani_zamestnanci = []
        self.config_file = os.path.join(self.data_path, 'employee_config.json')
        self.load_config()
        logger.info("Inicializována třída EmployeeManagement")

    def load_config(self):
        """Načte konfiguraci ze souboru"""
        if not os.path.exists(self.config_file):
            logger.warning(f"Konfigurační soubor {self.config_file} nenalezen")
            self.zamestnanci = []
            self.vybrani_zamestnanci = []
            return

        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                
                # Validace struktury načtených dat
                if not isinstance(config, dict):
                    raise ValueError("Neplatný formát konfiguračního souboru")
                    
                zamestnanci = config.get('zamestnanci', [])
                vybrani = config.get('vybrani_zamestnanci', [])
                
                # Validace typů a hodnot
                if not isinstance(zamestnanci, list) or not isinstance(vybrani, list):
                    raise ValueError("Neplatný formát seznamů")
                    
                if not all(isinstance(z, str) for z in zamestnanci):
                    raise ValueError("Seznam zaměstnanců obsahuje neplatné hodnoty")
                    
                if not all(isinstance(z, str) for z in vybrani):
                    raise ValueError("Seznam vybraných zaměstnanců obsahuje neplatné hodnoty")
                    
                if not all(z in zamestnanci for z in vybrani):
                    raise ValueError("Vybraní zaměstnanci nejsou podmnožinou všech zaměstnanců")
                
                self.zamestnanci = sorted(zamestnanci)
                self.vybrani_zamestnanci = vybrani
                
        except json.JSONDecodeError:
            logger.error(f"Chyba při čtení JSON z {self.config_file}")
            self.zamestnanci = []
            self.vybrani_zamestnanci = []
        except ValueError as e:
            logger.error(f"Chyba validace dat: {str(e)}")
            self.zamestnanci = []
            self.vybrani_zamestnanci = []
        except Exception as e:
            logger.error(f"Neočekávaná chyba při načítání konfigurace: {str(e)}")
            self.zamestnanci = []
            self.vybrani_zamestnanci = []

    def _validate_employee_name(self, name):
        """Validuje jméno zaměstnance"""
        if not name or not isinstance(name, str):
            raise ValueError("Neplatné jméno zaměstnance")
            
        name = name.strip()
        if len(name) < 2:
            raise ValueError("Jméno zaměstnance musí mít alespoň 2 znaky")
            
        if any(char.isdigit() for char in name):
            raise ValueError("Jméno zaměstnance nesmí obsahovat čísla")
            
        if not name.replace(' ', '').isalpha():
            raise ValueError("Jméno může obsahovat pouze písmena a mezery")
            
        return name

    def _validate_lists(self):
        """Validuje interní seznamy zaměstnanců"""
        if not isinstance(self.zamestnanci, list) or not isinstance(self.vybrani_zamestnanci, list):
            raise ValueError("Neplatné typy seznamů")
            
        if not all(isinstance(z, str) for z in self.zamestnanci):
            raise ValueError("Seznam zaměstnanců obsahuje neplatné hodnoty")
            
        if not all(isinstance(z, str) for z in self.vybrani_zamestnanci):
            raise ValueError("Seznam vybraných zaměstnanců obsahuje neplatné hodnoty")
            
        if not all(z in self.zamestnanci for z in self.vybrani_zamestnanci):
            raise ValueError("Vybraní zaměstnanci nejsou podmnožinou všech zaměstnanců")

    def save_config(self):
        """Uloží konfiguraci do souboru"""
        try:
            self._validate_lists()
            # Ensure data directory exists
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump({
                    'zamestnanci': sorted(self.zamestnanci),
                    'vybrani_zamestnanci': sorted(self.vybrani_zamestnanci)
                }, f, ensure_ascii=False, indent=4)
            logger.info("Konfigurace úspěšně uložena")
            return True
        except Exception as e:
            logger.error(f"Chyba při ukládání konfigurace: {str(e)}")
            return False

    def pridat_zamestnance(self, zamestnanec):
        """Přidá nového zaměstnance do seznamu"""
        try:
            zamestnanec = self._validate_employee_name(zamestnanec)
            if zamestnanec not in self.zamestnanci:
                self.zamestnanci.append(zamestnanec)
                self.zamestnanci.sort()
                self.save_config()
                logger.info(f"Přidán zaměstnanec: {zamestnanec}")
                return True
            logger.warning(f"Zaměstnanec již existuje: {zamestnanec}")
            return False
        except ValueError as e:
            logger.error(str(e))
            return False

    def pridat_vybraneho_zamestnance(self, zamestnanec):
        """Přidá zaměstnance do seznamu vybraných"""
        try:
            zamestnanec = self._validate_employee_name(zamestnanec)
            if zamestnanec in self.zamestnanci and zamestnanec not in self.vybrani_zamestnanci:
                self.vybrani_zamestnanci.append(zamestnanec)
                self.vybrani_zamestnanci.sort()
                self.save_config()
                logger.info(f"Přidán vybraný zaměstnanec: {zamestnanec}")
                return True
            logger.warning(f"Nepodařilo se přidat vybraného zaměstnance: {zamestnanec}")
            return False
        except ValueError as e:
            logger.error(str(e))
            return False

    def odebrat_vybraneho_zamestnance(self, zamestnanec):
        """Odebere zaměstnance ze seznamu vybraných"""
        if not zamestnanec or not isinstance(zamestnanec, str):
            logger.error("Neplatné jméno zaměstnance")
            return False
            
        zamestnanec = zamestnanec.strip()
        if len(zamestnanec) < 2:
            logger.error("Jméno zaměstnance musí mít alespoň 2 znaky")
            return False
            
        if zamestnanec in self.vybrani_zamestnanci:
            self.vybrani_zamestnanci.remove(zamestnanec)
            self.save_config()
            logger.info(f"Odebrán vybraný zaměstnanec: {zamestnanec}")
            return True
        logger.warning(f"Nepodařilo se odebrat vybraného zaměstnance: {zamestnanec}")
        return False

    def get_vybrani_zamestnanci(self):
        """Vrátí seznam vybraných zaměstnanců"""
        if not isinstance(self.vybrani_zamestnanci, list):
            logger.error("Seznam vybraných zaměstnanců není platný")
            return []
        return sorted(self.vybrani_zamestnanci)

    def get_nevybrani_zamestnanci(self):
        """Vrátí seznam nevybraných zaměstnanců"""
        if not isinstance(self.zamestnanci, list) or not isinstance(self.vybrani_zamestnanci, list):
            logger.error("Seznamy zaměstnanců nejsou platné")
            return []
        return sorted([z for z in self.zamestnanci if z not in self.vybrani_zamestnanci])

    def upravit_zamestnance(self, index, novy_nazev):
        """Upraví jméno zaměstnance"""
        try:
            if not isinstance(index, int):
                raise ValueError("Index musí být celé číslo")
                
            novy_nazev = self._validate_employee_name(novy_nazev)

            if not (1 <= index <= len(self.zamestnanci)):
                raise ValueError(f"Index {index} je mimo rozsah")

            stary_nazev = self.zamestnanci[index - 1]
            if novy_nazev != stary_nazev:
                if novy_nazev in self.zamestnanci:
                    raise ValueError(f"Zaměstnanec s jménem {novy_nazev} již existuje")
                    
                self.zamestnanci[index - 1] = novy_nazev
                if stary_nazev in self.vybrani_zamestnanci:
                    self.vybrani_zamestnanci.remove(stary_nazev)
                    self.vybrani_zamestnanci.append(novy_nazev)
                    self.vybrani_zamestnanci.sort()
                
                self.zamestnanci.sort()
                self.save_config()
                
            logger.info(f"Upraven zaměstnanec: {stary_nazev} -> {novy_nazev}")
            return True
            
        except ValueError as e:
            logger.error(str(e))
            return False
        except Exception as e:
            logger.error(f"Neočekávaná chyba při úpravě zaměstnance: {str(e)}")
            return False

    def smazat_zamestnance(self, index):
        """Smaže zaměstnance ze seznamu"""
        if not isinstance(index, int):
            logger.error("Index musí být celé číslo")
            return False
            
        if 1 <= index <= len(self.zamestnanci):
            zamestnanec = self.zamestnanci.pop(index - 1)
            if zamestnanec in self.vybrani_zamestnanci:
                self.vybrani_zamestnanci.remove(zamestnanec)
            self.save_config()
            logger.info(f"Smazán zaměstnanec: {zamestnanec}")
            return True
        logger.warning(f"Nepodařilo se smazat zaměstnance s indexem: {index}")
        return False

    def get_all_employees(self):
        """Vrátí seznam všech zaměstnanců s informací o jejich označení"""
        # Validace existence seznamů
        if self.zamestnanci is None or self.vybrani_zamestnanci is None:
            logger.error("Seznamy zaměstnanců nejsou inicializovány")
            return []

        # Validace typů seznamů
        if not isinstance(self.zamestnanci, list) or not isinstance(self.vybrani_zamestnanci, list):
            logger.error("Seznamy zaměstnanců nejsou platné")
            return []

        # Validace obsahu seznamů
        if not all(isinstance(z, str) for z in self.zamestnanci):
            logger.error("Seznam zaměstnanců obsahuje neplatné hodnoty")
            return []
            
        if not all(isinstance(z, str) for z in self.vybrani_zamestnanci):
            logger.error("Seznam vybraných zaměstnanců obsahuje neplatné hodnoty")
            return []

        # Validace konzistence dat
        if not all(z in self.zamestnanci for z in self.vybrani_zamestnanci):
            logger.error("Vybraní zaměstnanci nejsou podmnožinou všech zaměstnanců")
            return []
            
        try:
            return [
                {
                    'name': name,
                    'selected': name in self.vybrani_zamestnanci
                } 
                for name in sorted(self.zamestnanci)
            ]
        except Exception as e:
            logger.error(f"Chyba při vytváření seznamu zaměstnanců: {str(e)}")
            return []

    def get_employee_row(self, employee_name):
        workbook = None
        try:
            workbook = self.nacti_nebo_vytvor_excel()
            sheet = workbook[self.ZALOHY_SHEET_NAME]
            for row in range(self.EMPLOYEE_START_ROW, sheet.max_row + 1):
                if sheet.cell(row=row, column=1).value == employee_name:
                    return row
            return None
        finally:
            if workbook:
                workbook.close()

