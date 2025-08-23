"""Správa seznamu zaměstnanců a výběru pro zápis docházky.

Odděleno od Excel logiky – perzistence do JSON souboru `employee_config.json`.
"""

import json
from pathlib import Path

from utils.logger import setup_logger

logger = setup_logger("employee_management")


class EmployeeManager:
    def __init__(self, data_path):
        self.data_path = Path(data_path)
        self.config_file = self.data_path / "employee_config.json"
        self.zamestnanci = []
        self.vybrani_zamestnanci = []
        self.load_config()
        logger.info("EmployeeManager inicializován.")

    def _sort_selected_employees(self):
        """Preferuje pevně 'Čáp Jakub' na začátku (firemní priorita)."""
        self.vybrani_zamestnanci.sort(key=lambda x: (x != "Čáp Jakub", x))

    def load_config(self):
        """Načte konfiguraci (tichý fallback na prázdné seznamy)."""
        if not self.config_file.exists():
            logger.warning(f"Konfigurační soubor {self.config_file} nenalezen.")
            return

        try:
            with open(self.config_file, "r", encoding="utf-8") as f:
                config = json.load(f)

            self.zamestnanci = sorted(config.get("zamestnanci", []))
            self.vybrani_zamestnanci = config.get("vybrani_zamestnanci", [])

            if "Čáp Jakub" in self.zamestnanci and "Čáp Jakub" not in self.vybrani_zamestnanci:
                self.vybrani_zamestnanci.append("Čáp Jakub")

            self._sort_selected_employees()
        except (json.JSONDecodeError, Exception) as e:
            logger.error(f"Chyba při načítání konfigurace: {e}", exc_info=True)
            self.zamestnanci, self.vybrani_zamestnanci = [], []

    def _validate_employee_name(self, name):
        """Trim + základní validace délky a absence číslic."""
        name = name.strip()
        if not (2 <= len(name) <= 100) or any(char.isdigit() for char in name):
            raise ValueError("Neplatné jméno zaměstnance.")
        return name

    def save_config(self):
        """Uloží konfiguraci; řadí zaměstnance i vybraný seznam deterministicky."""
        try:
            self._sort_selected_employees()
            self.data_path.mkdir(parents=True, exist_ok=True)
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(
                    {"zamestnanci": sorted(self.zamestnanci), "vybrani_zamestnanci": self.vybrani_zamestnanci},
                    f,
                    ensure_ascii=False,
                    indent=4,
                )
            return True
        except Exception as e:
            logger.error(f"Chyba při ukládání konfigurace: {e}", exc_info=True)
            return False

    def pridat_zamestnance(self, zamestnanec):
        try:
            zamestnanec = self._validate_employee_name(zamestnanec)
            if zamestnanec not in self.zamestnanci:
                self.zamestnanci.append(zamestnanec)
                if zamestnanec == "Čáp Jakub":
                    self.vybrani_zamestnanci.append(zamestnanec)
                self.save_config()
                return True
            return False
        except ValueError as e:
            logger.error(e)
            return False

    def pridat_vybraneho_zamestnance(self, zamestnanec):
        if zamestnanec in self.zamestnanci and zamestnanec not in self.vybrani_zamestnanci:
            self.vybrani_zamestnanci.append(zamestnanec)
            return self.save_config()
        return False

    def odebrat_vybraneho_zamestnance(self, zamestnanec):
        if zamestnanec == "Čáp Jakub":
            logger.warning("Nelze odebrat 'Čáp Jakub' z výběru.")
            return False
        if zamestnanec in self.vybrani_zamestnanci:
            self.vybrani_zamestnanci.remove(zamestnanec)
            return self.save_config()
        return False

    def upravit_zamestnance_podle_jmena(self, old_name, new_name):
        try:
            validated_new_name = self._validate_employee_name(new_name)
            if validated_new_name in self.zamestnanci:
                raise ValueError(f"Zaměstnanec '{validated_new_name}' již existuje.")

            if old_name in self.zamestnanci:
                self.zamestnanci[self.zamestnanci.index(old_name)] = validated_new_name
                if old_name in self.vybrani_zamestnanci:
                    self.vybrani_zamestnanci[self.vybrani_zamestnanci.index(old_name)] = validated_new_name
                return self.save_config()
            return False
        except ValueError as e:
            logger.error(e)
            return False

    def smazat_zamestnance_podle_jmena(self, zamestnanec):
        if zamestnanec in self.zamestnanci:
            self.zamestnanci.remove(zamestnanec)
            if zamestnanec in self.vybrani_zamestnanci:
                self.vybrani_zamestnanci.remove(zamestnanec)
            return self.save_config()
        return False

    def get_all_employees(self):
        """Vrací seznam slovníků se stavem výběru (pro UI)."""
        return [{"name": name, "selected": name in self.vybrani_zamestnanci} for name in sorted(self.zamestnanci)]

    def get_vybrani_zamestnanci(self):
        """Seznam aktuálně vybraných zaměstnanců (preferenční řazení)."""
        return sorted(self.vybrani_zamestnanci, key=lambda x: (x != "Čáp Jakub", x))

    def set_vybrani_zamestnanci(self, employees_list):
        """Nastaví seznam vybraných zaměstnanců."""
        if not isinstance(employees_list, list):
            raise ValueError("Seznam zaměstnanců musí být typu list")

        # Validace - všichni zaměstnanci musí být v seznamu dostupných zaměstnanců
        for employee in employees_list:
            if employee not in self.zamestnanci:
                logger.warning(f"Zaměstnanec '{employee}' není v seznamu dostupných zaměstnanců")

        # Filtruj pouze platné zaměstnance
        valid_employees = [emp for emp in employees_list if emp in self.zamestnanci]

        # Zajisti, že "Čáp Jakub" je vždy zahrnut, pokud existuje
        if "Čáp Jakub" in self.zamestnanci and "Čáp Jakub" not in valid_employees:
            valid_employees.append("Čáp Jakub")

        self.vybrani_zamestnanci = valid_employees
        self._sort_selected_employees()
        return self.save_config()
