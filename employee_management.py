import json
import os
from pathlib import Path

from utils.logger import setup_logger

logger = setup_logger("employee_management")


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
        self.config_file = os.path.join(self.data_path, "employee_config.json")
        self.load_config()
        logger.info("Inicializována třída EmployeeManagement")

    def _sort_selected_employees(self):
        """Seřadí vybrané zaměstnance tak, že 'Čáp Jakub' je vždy první a ostatní jsou seřazeni abecedně"""
        cap_jakub = "Čáp Jakub"
        others = sorted([x for x in self.vybrani_zamestnanci if x != cap_jakub])
        if cap_jakub in self.vybrani_zamestnanci:
            self.vybrani_zamestnanci = [cap_jakub] + others
        else:
            self.vybrani_zamestnanci = others

    def load_config(self):
        """Načte konfiguraci ze souboru"""
        if not os.path.exists(self.config_file):
            logger.warning(f"Konfigurační soubor {self.config_file} nenalezen")
            self.zamestnanci = []
            self.vybrani_zamestnanci = []
            return

        try:
            with open(self.config_file, "r", encoding="utf-8") as f:
                config = json.load(f)

                if not isinstance(config, dict):
                    raise ValueError("Neplatný formát konfiguračního souboru")

                zamestnanci = config.get("zamestnanci", [])
                vybrani = config.get("vybrani_zamestnanci", [])

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

                # Automaticky přidat Čáp Jakub do vybraných pokud existuje
                if "Čáp Jakub" in self.zamestnanci and "Čáp Jakub" not in self.vybrani_zamestnanci:
                    self.vybrani_zamestnanci.append("Čáp Jakub")
                
                self._sort_selected_employees()

        except json.JSONDecodeError:
            logger.error("Chyba při načítání konfiguračního souboru")
            self.zamestnanci = []
            self.vybrani_zamestnanci = []
        except Exception as e:
            logger.error(f"Neočekávaná chyba při načítání konfigurace: {str(e)}", exc_info=True)
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

        if not name.replace(" ", "").isalpha():
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
            self._sort_selected_employees()
            
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)

            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(
                    {"zamestnanci": sorted(self.zamestnanci), "vybrani_zamestnanci": self.vybrani_zamestnanci},
                    f,
                    ensure_ascii=False,
                    indent=4,
                )
            logger.info("Konfigurace úspěšně uložena")
            return True
        except Exception as e:
            logger.error(f"Chyba při ukládání konfigurace: {str(e)}", exc_info=True)
            return False

    def pridat_zamestnance(self, zamestnanec):
        """Přidá nového zaměstnance do seznamu"""
        try:
            zamestnanec = self._validate_employee_name(zamestnanec)
            if zamestnanec not in self.zamestnanci:
                self.zamestnanci.append(zamestnanec)
                self.zamestnanci.sort()
                
                # Automaticky přidat Čáp Jakub do vybraných
                if zamestnanec == "Čáp Jakub":
                    self.vybrani_zamestnanci.append(zamestnanec)
                    self._sort_selected_employees()
                
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
                self._sort_selected_employees()
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

        # Zabránit odebrání Čáp Jakub ze seznamu vybraných
        if zamestnanec == "Čáp Jakub":
            logger.warning("Nelze odebrat zaměstnance 'Čáp Jakub' ze seznamu vybraných")
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

    def _upravit_zamestnance_podle_indexu(self, index, novy_nazev):
        """Interní metoda pro úpravu jména zaměstnance podle 1-based indexu."""
        try:
            if not isinstance(index, int):
                raise ValueError("Index musí být celé číslo")

            novy_nazev = self._validate_employee_name(novy_nazev)

            if not (1 <= index <= len(self.zamestnanci)):
                raise ValueError(f"Index {index} je mimo rozsah (1 až {len(self.zamestnanci)})")

            stary_nazev = self.zamestnanci[index - 1] # Převedení 1-based na 0-based
            
            # Pokud se název nemění, není co dělat
            if novy_nazev == stary_nazev:
                logger.info(f"Jméno zaměstnance '{stary_nazev}' se nemění.")
                return True

            # Kontrola, zda nový název již neexistuje (pokud se liší od starého)
            if novy_nazev in self.zamestnanci:
                raise ValueError(f"Zaměstnanec s jménem '{novy_nazev}' již existuje")

            self.zamestnanci[index - 1] = novy_nazev
            
            # Úprava v seznamu vybraných zaměstnanců
            if stary_nazev in self.vybrani_zamestnanci:
                self.vybrani_zamestnanci.remove(stary_nazev)
                self.vybrani_zamestnanci.append(novy_nazev)
            
            # Pokud je nové jméno "Čáp Jakub" a není ve vybraných, přidat ho
            if novy_nazev == "Čáp Jakub" and novy_nazev not in self.vybrani_zamestnanci:
                self.vybrani_zamestnanci.append(novy_nazev)
            
            # Seřadit seznamy
            self.zamestnanci.sort() # Seřadí hlavní seznam zaměstnanců
            self._sort_selected_employees() # Seřadí seznam vybraných zaměstnanců
            
            if self.save_config():
                logger.info(f"Zaměstnanec úspěšně upraven: '{stary_nazev}' -> '{novy_nazev}'")
                return True
            else:
                # Pokud save_config selže, vrátíme změny zpět, aby byla zachována konzistence
                logger.error(f"Nepodařilo se uložit konfiguraci po úpravě '{stary_nazev}' na '{novy_nazev}'. Změny se vrací.")
                self.zamestnanci[index - 1] = stary_nazev # Vrácení původního jména
                if novy_nazev in self.vybrani_zamestnanci: # Pokud byl nový název přidán do vybraných
                    self.vybrani_zamestnanci.remove(novy_nazev)
                    if stary_nazev not in self.vybrani_zamestnanci: # Pokud starý název nebyl ve vybraných (což by nemělo nastat, pokud byl nahrazen)
                         self.vybrani_zamestnanci.append(stary_nazev)
                self.zamestnanci.sort()
                self._sort_selected_employees()
                return False

        except ValueError as e: # Chyby z _validate_employee_name nebo kontroly indexu/existence
            logger.error(f"Chyba při úpravě zaměstnance: {str(e)}")
            return False
        except Exception as e: # Neočekávané chyby
            logger.error(f"Neočekávaná chyba při úpravě zaměstnance ('{stary_nazev}' na '{novy_nazev}'): {str(e)}", exc_info=True)
            return False

    def upravit_zamestnance_podle_jmena(self, old_name, new_name):
        """
        Veřejná metoda pro úpravu jména zaměstnance.
        Najde zaměstnance podle `old_name` a změní jeho jméno na `new_name`.
        """
        try:
            # Validace starého a nového jména (nové jméno se validuje i v _upravit_zamestnance_podle_indexu)
            validated_old_name = self._validate_employee_name(old_name)
            validated_new_name = self._validate_employee_name(new_name)

            if validated_old_name not in self.zamestnanci:
                logger.warning(f"Zaměstnanec '{validated_old_name}' nebyl nalezen pro úpravu.")
                return False
            
            # Najdeme 0-based index starého jména
            try:
                index_0_based = self.zamestnanci.index(validated_old_name)
            except ValueError: # Mělo by být již pokryto kontrolou `if validated_old_name not in self.zamestnanci`
                logger.warning(f"Zaměstnanec '{validated_old_name}' nebyl nalezen v seznamu (vnitřní chyba).")
                return False

            # _upravit_zamestnance_podle_indexu očekává 1-based index
            return self._upravit_zamestnance_podle_indexu(index_0_based + 1, validated_new_name)

        except ValueError as e: # Chyba z _validate_employee_name pro old_name nebo new_name
            logger.error(f"Chyba při validaci jména pro úpravu ('{old_name}' na '{new_name}'): {str(e)}")
            return False
        except Exception as e: # Neočekávané chyby
            logger.error(f"Neočekávaná chyba při úpravě zaměstnance podle jména ('{old_name}' na '{new_name}'): {str(e)}", exc_info=True)
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
        if self.zamestnanci is None or self.vybrani_zamestnanci is None:
            logger.error("Seznamy zaměstnanců nejsou inicializovány")
            return []

        if not isinstance(self.zamestnanci, list) or not isinstance(self.vybrani_zamestnanci, list):
            logger.error("Seznamy zaměstnanců nejsou platné")
            return []

        if not all(isinstance(z, str) for z in self.zamestnanci):
            logger.error("Seznam zaměstnanců obsahuje neplatné hodnoty")
            return []

        if not all(isinstance(z, str) for z in self.vybrani_zamestnanci):
            logger.error("Seznam vybraných zaměstnanců obsahuje neplatné hodnoty")
            return []

        if not all(z in self.zamestnanci for z in self.vybrani_zamestnanci):
            logger.error("Vybraní zaměstnanci nejsou podmnožinou všech zaměstnanců")
            return []

        try:
            return [{"name": name, "selected": name in self.vybrani_zamestnanci} for name in sorted(self.zamestnanci)]
        except Exception as e:
            logger.error(f"Chyba při vytváření seznamu zaměstnanců: {str(e)}", exc_info=True)
            return []

    def get_selected_employees(self):
        """Vrátí seznam vybraných zaměstnanců"""
        return self.vybrani_zamestnanci
