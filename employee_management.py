import json
import os
import logging

class EmployeeManagement:
    def __init__(self):
        self.zamestnanci = []
        self.vybrani_zamestnanci = []
        self.config_file = 'employee_config.json'
        self.load_config()
        logging.info("Inicializována třída EmployeeManagement")

    def load_config(self):
        if os.path.exists(self.config_file):
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                self.zamestnanci = config.get('zamestnanci', [])
                self.vybrani_zamestnanci = config.get('vybrani_zamestnanci', [])
            logging.info(f"Načtena konfigurace: {len(self.zamestnanci)} zaměstnanců, {len(self.vybrani_zamestnanci)} vybraných")
        else:
            logging.warning(f"Konfigurační soubor {self.config_file} nenalezen")

    def pridat_zamestnance(self, jmeno):
        logging.info(f"Pokus o přidání zaměstnance: {jmeno}")
        if jmeno and jmeno not in self.zamestnanci:
            self.zamestnanci.append(jmeno)
            self.save_config()
            logging.info(f"Přidán nový zaměstnanec: {jmeno}")
            return True
        logging.warning(f"Nepodařilo se přidat zaměstnance: {jmeno}")
        return False

    def save_config(self):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump({
                    'zamestnanci': self.zamestnanci,
                    'vybrani_zamestnanci': self.vybrani_zamestnanci
                }, f, ensure_ascii=False, indent=2)
            logging.info(f"Konfigurace uložena do souboru: {self.config_file}")
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

    def oznacit_zamestnance(self, cislo):
        if 1 <= cislo <= len(self.zamestnanci):
            zamestnanec = self.zamestnanci[cislo - 1]
            if zamestnanec in self.vybrani_zamestnanci:
                return self.odebrat_vybraneho_zamestnance(zamestnanec)
            else:
                return self.pridat_vybraneho_zamestnance(zamestnanec)
        logging.error(f"Pokus o označení/odznačení zaměstnance s neplatným číslem: {cislo}")
        return False

    def get_vybrani_zamestnanci(self):
        logging.info(f"Vrácen seznam vybraných zaměstnanců: {len(self.vybrani_zamestnanci)} položek")
        return self.vybrani_zamestnanci