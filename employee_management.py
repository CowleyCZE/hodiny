import os
import json
import logging

class EmployeeManager:
    def __init__(self, data_path):
        self.data_path = data_path
        self.zamestnanci = []
        self.vybrani_zamestnanci = []
        self.config_file = os.path.join(self.data_path, 'employee_config.json')
        self.load_config()
        logging.info("Inicializována třída EmployeeManagement")

    def load_config(self):
        """Načte konfiguraci ze souboru"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.zamestnanci = sorted(config.get('zamestnanci', []))  # Seřazení podle abecedy
                    self.vybrani_zamestnanci = config.get('vybrani_zamestnanci', [])
            except json.JSONDecodeError:
                logging.error(f"Chyba při čtení JSON z {self.config_file}")
                self.zamestnanci = []
                self.vybrani_zamestnanci = []
            except Exception as e:
                logging.error(f"Neočekávaná chyba při načítání konfigurace: {str(e)}")
                self.zamestnanci = []
                self.vybrani_zamestnanci = []
        else:
            logging.warning(f"Konfigurační soubor {self.config_file} nenalezen")
            self.zamestnanci = []
            self.vybrani_zamestnanci = []

    def save_config(self):
        """Uloží konfiguraci do souboru"""
        try:
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump({
                    'zamestnanci': sorted(self.zamestnanci),  # Seřazení podle abecedy
                    'vybrani_zamestnanci': self.vybrani_zamestnanci
                }, f, ensure_ascii=False, indent=4)
            logging.info("Konfigurace úspěšně uložena")
        except Exception as e:
            logging.error(f"Chyba při ukládání konfigurace: {str(e)}")

    def pridat_zamestnance(self, zamestnanec):
        """Přidá nového zaměstnance do seznamu"""
        if not zamestnanec or not isinstance(zamestnanec, str):
            logging.error("Neplatné jméno zaměstnance")
            return False
        
        if zamestnanec not in self.zamestnanci:
            self.zamestnanci.append(zamestnanec)
            self.zamestnanci.sort()  # Seřazení podle abecedy
            self.save_config()
            logging.info(f"Přidán zaměstnanec: {zamestnanec}")
            return True
        logging.warning(f"Zaměstnanec již existuje: {zamestnanec}")
        return False

    def pridat_vybraneho_zamestnance(self, zamestnanec):
        """Přidá zaměstnance do seznamu vybraných"""
        if zamestnanec in self.zamestnanci and zamestnanec not in self.vybrani_zamestnanci:
            self.vybrani_zamestnanci.append(zamestnanec)
            self.vybrani_zamestnanci.sort()  # Seřazení podle abecedy
            self.save_config()
            logging.info(f"Přidán vybraný zaměstnanec: {zamestnanec}")
            return True
        logging.warning(f"Nepodařilo se přidat vybraného zaměstnance: {zamestnanec}")
        return False

    def odebrat_vybraneho_zamestnance(self, zamestnanec):
        """Odebere zaměstnance ze seznamu vybraných"""
        if zamestnanec in self.vybrani_zamestnanci:
            self.vybrani_zamestnanci.remove(zamestnanec)
            self.save_config()
            logging.info(f"Odebrán vybraný zaměstnanec: {zamestnanec}")
            return True
        logging.warning(f"Nepodařilo se odebrat vybraného zaměstnance: {zamestnanec}")
        return False

    def get_vybrani_zamestnanci(self):
        """Vrátí seznam vybraných zaměstnanců"""
        return sorted(self.vybrani_zamestnanci)  # Seřazení podle abecedy

    def get_nevybrani_zamestnanci(self):
        """Vrátí seznam nevybraných zaměstnanců"""
        return sorted([z for z in self.zamestnanci if z not in self.vybrani_zamestnanci])

    def upravit_zamestnance(self, index, novy_nazev):
        """Upraví jméno zaměstnance"""
        if not novy_nazev or not isinstance(novy_nazev, str):
            logging.error("Neplatné nové jméno zaměstnance")
            return False

        if 1 <= index <= len(self.zamestnanci):
            stary_nazev = self.zamestnanci[index - 1]
            if novy_nazev != stary_nazev and novy_nazev in self.zamestnanci:
                logging.error(f"Zaměstnanec s jménem {novy_nazev} již existuje")
                return False
            
            self.zamestnanci[index - 1] = novy_nazev
            # Aktualizace ve vybraných zaměstnancích
            if stary_nazev in self.vybrani_zamestnanci:
                self.vybrani_zamestnanci.remove(stary_nazev)
                self.vybrani_zamestnanci.append(novy_nazev)
                self.vybrani_zamestnanci.sort()  # Seřazení podle abecedy
            
            self.zamestnanci.sort()  # Seřazení podle abecedy
            self.save_config()
            logging.info(f"Upraven zaměstnanec: {stary_nazev} -> {novy_nazev}")
            return True
        logging.warning(f"Nepodařilo se upravit zaměstnance s indexem: {index}")
        return False

    def smazat_zamestnance(self, index):
        """Smaže zaměstnance ze seznamu"""
        if 1 <= index <= len(self.zamestnanci):
            zamestnanec = self.zamestnanci.pop(index - 1)
            if zamestnanec in self.vybrani_zamestnanci:
                self.vybrani_zamestnanci.remove(zamestnanec)
            self.save_config()
            logging.info(f"Smazán zaměstnanec: {zamestnanec}")
            return True
        logging.warning(f"Nepodařilo se smazat zaměstnance s indexem: {index}")
        return False

    def get_all_employees(self):
        """Vrátí seznam všech zaměstnanců s informací o jejich označení"""
        return [{'name': name, 'selected': name in self.vybrani_zamestnanci} 
                for name in sorted(self.zamestnanci)]  # Seřazení podle abecedy

