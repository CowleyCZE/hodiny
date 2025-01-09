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
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.zamestnanci = config.get('zamestnanci', [])
                    self.vybrani_zamestnanci = config.get('vybrani_zamestnanci', [])
            else:
                self.zamestnanci = []
                self.vybrani_zamestnanci = []
        except Exception as e:
            logging.error(f"Chyba při načítání konfigurace: {str(e)}")
            self.zamestnanci = []
            self.vybrani_zamestnanci = []

    def save_config(self):
        try:
            config = {
                'zamestnanci': self.zamestnanci,
                'vybrani_zamestnanci': self.vybrani_zamestnanci
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.error(f"Chyba při ukládání konfigurace: {str(e)}")

    def pridat_zamestnance(self, name):
        if name not in self.zamestnanci:
            self.zamestnanci.append(name)
            self.save_config()
            logging.info(f"Přidán nový zaměstnanec: {name}")
            return True
        logging.warning(f"Zaměstnanec {name} již existuje")
        return False

    def smazat_zamestnance(self, name):
        if name in self.zamestnanci:
            self.zamestnanci.remove(name)
            if name in self.vybrani_zamestnanci:
                self.vybrani_zamestnanci.remove(name)
            self.save_config()
            logging.info(f"Smazán zaměstnanec: {name}")
            return True
        logging.warning(f"Zaměstnanec {name} neexistuje")
        return False

    def upravit_zamestnance(self, old_name, new_name):
        if old_name in self.zamestnanci:
            index = self.zamestnanci.index(old_name)
            self.zamestnanci[index] = new_name
            if old_name in self.vybrani_zamestnanci:
                index = self.vybrani_zamestnanci.index(old_name)
                self.vybrani_zamestnanci[index] = new_name
            self.save_config()
            logging.info(f"Upraven zaměstnanec: {old_name} na {new_name}")
            return True
        logging.warning(f"Zaměstnanec {old_name} neexistuje")
        return False

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

    def get_employees(self):
        return self.get_zamestnanci()

    def get_zamestnanci(self):
        return sorted(self.zamestnanci)

    def get_vybrani_zamestnanci(self):
        return sorted(self.vybrani_zamestnanci)

    def get_all_zamestnanci(self):
        vybrani = self.get_vybrani_zamestnanci()
        ostatni = [z for z in self.get_zamestnanci() if z not in vybrani]
        return vybrani + ostatni