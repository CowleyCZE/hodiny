import os
import json
import logging

class EmployeeManager:
    def __init__(self, data_path):
        self.data_path = data_path
        self.zamestnanci = []
        self.vybrani_zamestnanci = []
        self.data_path = data_path
        self.config_file = os.path.join(self.data_path, 'employee_config.json')
        self.load_config()
        logging.info("Inicializována třída EmployeeManagement")

    def pridat_zamestnance(self, name):
        if name not in self.zamestnanci:
            self.zamestnanci.append(name)
            self.save_config()
            return True
        return False

    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    self.zamestnanci = config.get('zamestnanci', [])
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
        try:
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            with open(self.config_file, 'w') as f:
                json.dump({
                    'zamestnanci': self.zamestnanci,
                    'vybrani_zamestnanci': self.vybrani_zamestnanci
                }, f)
            logging.info("Konfigurace úspěšně uložena")
        except Exception as e:
            logging.error(f"Chyba při ukládání konfigurace: {str(e)}")

    def pridat_vybraneho_zamestnance(self, zamestnanec):
        if zamestnanec in self.zamestnanci and zamestnanec not in self.vybrani_zamestnanci:
            self.vybrani_zamestnanci.append(zamestnanec)
            self.save_config()
            logging.info(f"Přidán vybraný zaměstnanec: {zamestnanec}")
            return True
        logging.warning(f"Nepodařilo se přidat vybraného zaměstnance: {zamestnanec}")
        return False

    def odebrat_vybraneho_zamestnance(self, zamestnanec):
        if zamestnanec in self.vybrani_zamestnanci:
            self.vybrani_zamestnanci.remove(zamestnanec)
            self.save_config()
            logging.info(f"Odebrán vybraný zaměstnanec: {zamestnanec}")
            return True
        logging.warning(f"Nepodařilo se odebrat vybraného zaměstnance: {zamestnanec}")
        return False

    def get_vybrani_zamestnanci(self):
        return self.vybrani_zamestnanci

    def upravit_zamestnance(self, index, novy_nazev):
        if 1 <= index <= len(self.zamestnanci):
            self.zamestnanci[index - 1] = novy_nazev
            self.save_config()
            logging.info(f"Upraven zaměstnanec: {novy_nazev}")
            return True
        logging.warning(f"Nepodařilo se upravit zaměstnance s indexem: {index}")
        return False

    def smazat_zamestnance(self, index):
        if 1 <= index <= len(self.zamestnanci):
            zamestnanec = self.zamestnanci.pop(index - 1)
            if zamestnanec in self.vybrani_zamestnanci:
                self.vybrani_zamestnanci.remove(zamestnanec)
            self.save_config()
            logging.info(f"Smazán zaměstnanec: {zamestnanec}")
            return True
        logging.warning(f"Nepodařilo se smazat zaměstnance s indexem: {index}")
        return False
