# zalohy_manager.py
from config import Config
import logging
import os
from datetime import datetime
from pathlib import Path # Import Path

from openpyxl import Workbook, load_workbook

# Předpokládá existenci utils.logger
try:
    from utils.logger import setup_logger
    logger = setup_logger("zalohy_manager")
except ImportError:
    # Fallback na základní logger, pokud utils.logger není dostupný
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger("zalohy_manager")


class ZalohyManager:
    """
    Správce záloh pro zaměstnance v aktivním Excel souboru.
    """
    def __init__(self, base_path, active_filename):
        """
        Inicializuje ZalohyManager.

        Args:
            base_path (Path): Cesta k adresáři s Excel soubory.
            active_filename (str): Název aktuálně používaného souboru.
        """
        if not active_filename:
            logger.error("Nelze inicializovat ZalohyManager bez názvu aktivního souboru.")
            # Můžeme vyvolat chybu nebo nastavit cestu na None a kontrolovat později
            raise ValueError("Chybí název aktivního souboru pro ZalohyManager.")

        self.base_path = Path(base_path)
        self.active_filename = active_filename
        # Sestavení plné cesty k aktivnímu souboru
        self.active_file_path = self.base_path / self.active_filename

        self.ZALOHY_SHEET_NAME = Config.EXCEL_ADVANCES_SHEET_NAME
        self.EMPLOYEE_START_ROW = Config.EXCEL_EMPLOYEE_START_ROW
        self.VALID_CURRENCIES = ["EUR", "CZK"]
        # VALID_OPTIONS se nyní načítají dynamicky z Excelu, ale můžeme je nechat pro případnou validaci
        # self.VALID_OPTIONS = ["Option 1", "Option 2"]

        # Zajistíme existenci adresáře (i když by měl z init_app)
        self.base_path.mkdir(parents=True, exist_ok=True)
        logger.info(f"ZalohyManager inicializován pro soubor: {self.active_file_path}")


    def _get_active_workbook(self, read_only=False):
        """
        Načte aktivní workbook (Excel soubor) ze souborového systému.

        Tato metoda se používá interně v rámci `ZalohyManager` k získání instance
        sešitu pro čtení nebo zápis. Nepoužívá žádný mechanismus cachování na úrovni
        instance `ZalohyManager`; každý požadavek na sešit vede k jeho novému načtení
        ze souboru.

        Životní cyklus sešitu spravovaného touto metodou:
        1. Otevření/Načtení: Sešit je načten ze souboru `self.active_file_path`
           pomocí `load_workbook()`. Parametr `read_only` určuje, zda bude
           otevřen pouze pro čtení nebo pro čtení/zápis. `data_only=True` zajišťuje,
           že se načtou hodnoty buněk místo vzorců.
        2. Použití: Vrácený objekt `Workbook` je pak použit volající metodou.
        3. Zavření: **Je plnou zodpovědností volající metody zajistit, aby byl
           načtený sešit správně uzavřen pomocí `workbook.close()` po dokončení
           všech operací.** Tato metoda sama o sobě sešit nezavírá.

        Args:
            read_only (bool): Pokud True, sešit se otevře pouze pro čtení.
                              Výchozí je False (pro čtení i zápis).

        Returns:
            openpyxl.Workbook: Načtený objekt sešitu.

        Raises:
            FileNotFoundError: Pokud aktivní soubor (`self.active_filename`) neexistuje.
            IOError: Pokud dojde k chybě při načítání souboru.
        """
        if not self.active_file_path.exists():
             logger.error(f"Aktivní soubor '{self.active_filename}' nebyl nalezen na cestě '{self.active_file_path}'.")
             raise FileNotFoundError(f"Aktivní soubor '{self.active_filename}' neexistuje na cestě '{self.active_file_path}'.")
        try:
            # data_only=True načte hodnoty místo vzorců, což je obvykle žádoucí.
            logger.debug(f"Načítání workbooku: {self.active_file_path} (read_only={read_only})")
            wb = load_workbook(filename=self.active_file_path, read_only=read_only, data_only=True)
            logger.debug(f"Workbook '{self.active_filename}' úspěšně načten.")
            return wb
        except Exception as e:
            logger.error(f"Chyba při načítání workbooku '{self.active_filename}' z '{self.active_file_path}': {e}", exc_info=True)
            raise IOError(f"Nepodařilo se otevřít soubor '{self.active_filename}'. Zkontrolujte, zda soubor existuje a není poškozen.")

    def _save_workbook(self, workbook):
         """
         Uloží daný workbook (sešit) do aktivního souboru na disku.

         Tato metoda je volána poté, co byly v sešitu provedeny změny,
         které je třeba persistentně uložit.

         Args:
             workbook (openpyxl.Workbook): Objekt sešitu, který má být uložen.

         Raises:
             IOError: Pokud dojde k chybě při ukládání souboru (např. kvůli oprávněním,
                      nedostatku místa na disku nebo pokud je soubor používán jiným procesem).
         """
         if workbook is None:
             logger.error("Pokus o uložení None workbooku. Operace byla přeskočena.")
             # Můžeme zde vyvolat ValueError, pokud je to považováno za kritickou chybu.
             # raise ValueError("Nelze uložit None workbook.")
             return

         try:
              logger.debug(f"Pokus o uložení změn do workbooku: {self.active_file_path}")
              workbook.save(self.active_file_path)
              logger.info(f"Workbook '{self.active_filename}' úspěšně uložen do '{self.active_file_path}'.")
         except Exception as e:
              logger.error(f"Chyba při ukládání workbooku '{self.active_filename}' do '{self.active_file_path}': {e}", exc_info=True)
              raise IOError(f"Nepodařilo se uložit změny do souboru '{self.active_filename}'. Zkontrolujte oprávnění a zda soubor není používán jiným procesem.")


    # --- Validační metody (zůstávají stejné, komentáře nejsou požadovány pro změnu) ---
    def validate_amount(self, amount):
        """Validuje částku zálohy"""
        if not isinstance(amount, (int, float)):
            raise ValueError("Částka musí být číslo")
        if amount <= 0:
            raise ValueError("Částka musí být větší než 0")
        if amount > 1000000:  # Limit
            raise ValueError("Částka je příliš vysoká")
        return True

    def validate_currency(self, currency):
        """Validuje měnu"""
        if not isinstance(currency, str):
            raise ValueError("Měna musí být textový řetězec")
        if currency not in self.VALID_CURRENCIES:
            raise ValueError(f"Neplatná měna. Povolené měny jsou: {', '.join(self.VALID_CURRENCIES)}")
        return True

    def validate_employee_name(self, employee_name):
        """Validuje jméno zaměstnance"""
        if not isinstance(employee_name, str) or not employee_name.strip():
            raise ValueError("Jméno zaměstnance nemůže být prázdné")
        if len(employee_name) > 100:
            raise ValueError("Jméno zaměstnance je příliš dlouhé")
        # Můžeme přidat kontrolu na znaky, pokud je potřeba
        return True

    # Validace option se nyní dělá v app.py proti dynamicky načteným možnostem
    # def validate_option(self, option): ...

    def validate_date(self, date_str):
        """Validuje formát data"""
        try:
            datetime.strptime(date_str, "%Y-%m-%d")
            return True
        except ValueError:
            raise ValueError("Neplatný formát data. Použijte formát YYYY-MM-DD")

    # --- Metody pro práci se zálohami ---

    def _ensure_zalohy_sheet(self, workbook):
         """Zajistí existenci listu 'Zálohy' a vrátí ho."""
         if self.ZALOHY_SHEET_NAME not in workbook.sheetnames:
             sheet = workbook.create_sheet(self.ZALOHY_SHEET_NAME)
             # Nastavíme výchozí názvy možností
             sheet["B80"] = Config.DEFAULT_ADVANCE_OPTION_1
             sheet["D80"] = Config.DEFAULT_ADVANCE_OPTION_2
             sheet["F80"] = Config.DEFAULT_ADVANCE_OPTION_3
             sheet["H80"] = Config.DEFAULT_ADVANCE_OPTION_4
             logger.info(f"Vytvořen list '{self.ZALOHY_SHEET_NAME}' v souboru {self.active_filename}.")
             return sheet
         else:
             return workbook[self.ZALOHY_SHEET_NAME]

    def _get_employee_row(self, sheet, employee_name):
        """Najde řádek pro daného zaměstnance v listu."""
        # Projdeme řádky od EMPLOYEE_START_ROW
        for row in range(self.EMPLOYEE_START_ROW, sheet.max_row + 1):
            if sheet.cell(row=row, column=1).value == employee_name:
                return row
        return None # Zaměstnanec nenalezen

    def _get_next_empty_row(self, sheet):
         """Najde první prázdný řádek pro nového zaměstnance."""
         row = self.EMPLOYEE_START_ROW
         while sheet.cell(row=row, column=1).value is not None:
              row += 1
         return row


    def add_or_update_employee_advance(self, employee_name, amount, currency, option, date):
        """
        Přidá nebo aktualizuje zálohu pro daného zaměstnance v aktivním Excel souboru.

        Životní cyklus operace se sešitem:
        1. Validace vstupních dat (jméno, částka, měna, datum).
        2. Načtení sešitu: Volá `_get_active_workbook(read_only=False)` pro získání
           instance sešitu otevřené pro zápis. Tento sešit není cachován
           na úrovni `ZalohyManager`.
        3. Zajištění listu: Metoda `_ensure_zalohy_sheet()` zkontroluje existenci
           listu "Zálohy"; pokud neexistuje, vytvoří ho.
        4. Zpracování dat: Najde nebo vytvoří řádek pro zaměstnance, určí správný
           sloupec na základě `option` a `currency`, a aktualizuje buňku
           s novou hodnotou zálohy. Také zaznamená datum zálohy.
        5. Uložení sešitu: Po úspěšné aktualizaci dat v paměti je volána metoda
           `_save_workbook()`, která zapíše všechny změny z instance sešitu
           zpět do souboru na disku.
        6. Zavření sešitu: Ve `finally` bloku je zajištěno, že sešit je vždy
           uzavřen pomocí `workbook.close()`, aby se uvolnily systémové prostředky
           a předešlo se problémům s přístupem k souboru.

        Args:
            employee_name (str): Jméno zaměstnance.
            amount (float): Částka zálohy.
            currency (str): Měna ("EUR" nebo "CZK").
            option (str): Typ zálohy (název odpovídá hodnotě v B80 nebo D80 v listu Zálohy).
            date (str): Datum zálohy ve formátu "YYYY-MM-DD".

        Returns:
            bool: True, pokud byla operace úspěšná.

        Raises:
            ValueError: Pokud jsou vstupní data nevalidní.
            FileNotFoundError: Pokud aktivní Excel soubor neexistuje.
            IOError: Pokud dojde k chybě při čtení nebo zápisu souboru.
            RuntimeError: Pro jiné neočekávané chyby.
        """
        workbook = None # Inicializace pro finally blok
        try:
            # Fáze 1: Validace vstupů
            self.validate_employee_name(employee_name)
            self.validate_amount(amount)
            self.validate_currency(currency)
            # self.validate_option(option) se typicky děje v app.py proti dynamicky načteným možnostem
            self.validate_date(date)

            # Fáze 2: Načteme workbook pro zápis
            # read_only=False znamená, že můžeme provádět změny.
            # Volající (tato metoda) je zodpovědný za uzavření workbooku.
            workbook = self._get_active_workbook(read_only=False)
            
            # Fáze 3: Zajistíme existenci listu 'Zálohy'
            sheet = self._ensure_zalohy_sheet(workbook)

            # Získáme aktuální názvy možností záloh přímo z listu
            option1_value = sheet["B80"].value or Config.DEFAULT_ADVANCE_OPTION_1
            option2_value = sheet["D80"].value or Config.DEFAULT_ADVANCE_OPTION_2
            option3_value = sheet["F80"].value or Config.DEFAULT_ADVANCE_OPTION_3
            option4_value = sheet["H80"].value or Config.DEFAULT_ADVANCE_OPTION_4

            # Najdeme řádek zaměstnance nebo vytvoříme nový
            row = self._get_employee_row(sheet, employee_name)
            if row is None:
                row = self._get_next_empty_row(sheet)
                sheet.cell(row=row, column=1, value=employee_name)
                logger.info(f"Přidán nový zaměstnanec '{employee_name}' do listu '{self.ZALOHY_SHEET_NAME}' na řádek {row}.")

            # Určíme sloupec podle možnosti a měny (indexy od 1)
            if option == option1_value:
                column_index = 2 if currency == "EUR" else 3 # B nebo C
            elif option == option2_value:
                column_index = 4 if currency == "EUR" else 5 # D nebo E
            elif option == option3_value:
                column_index = 6 if currency == "EUR" else 7 # F nebo G
            elif option == option4_value:
                column_index = 8 if currency == "EUR" else 9 # H nebo I
            else:
                logger.error(f"Neznámá možnost zálohy '{option}' při ukládání. Používají se sloupce pro '{option1_value}'.")
                # Můžeme vyvolat chybu nebo použít fallback
                # raise ValueError(f"Neplatná volba zálohy: {option}")
                column_index = 2 if currency == "EUR" else 3 # Fallback

            # Aktualizujeme hodnotu zálohy
            target_cell = sheet.cell(row=row, column=column_index)
            current_value = target_cell.value or 0
            try:
                current_value_float = float(current_value)
                amount_float = float(amount)
                new_value = current_value_float + amount_float
                target_cell.value = new_value
                target_cell.number_format = '#,##0.00' # Nastavení formátu
            except (ValueError, TypeError):
                logger.error(f"Neplatná číselná hodnota v buňce {target_cell.coordinate} ('{current_value}') nebo částka '{amount}'. Záloha nebude přičtena.")
                raise ValueError("Neplatná číselná hodnota pro zálohu.")

            # Přidání data zálohy do sloupce Z (index 25 -> sloupec 26)
            date_column_index = 26
            date_cell = sheet.cell(row=row, column=date_column_index)
            try:
                date_obj = datetime.strptime(date, "%Y-%m-%d").date()
                date_cell.value = date_obj
                date_cell.number_format = 'DD.MM.YYYY'
            except ValueError:
                logger.error(f"Neplatný formát data '{date}' pro zálohu. Datum nebude uloženo.")
                date_cell.value = None

            # Uložíme změny
            self._save_workbook(workbook)
            logger.info(f"Záloha pro {employee_name} ({amount} {currency}, {option}, {date}) úspěšně uložena do {self.active_filename}")
            return True

        except (FileNotFoundError, ValueError, IOError) as e:
            logger.error(f"Chyba při ukládání zálohy do {self.active_filename}: {e}", exc_info=False) # False pro stručnější log u validací
            # Propagujeme chybu výše, aby ji mohlo zachytit app.py
            raise e
        except Exception as e:
            logger.error(f"Neočekávaná chyba při ukládání zálohy do {self.active_filename}: {e}", exc_info=True)
            # Propagujeme jako obecnou chybu
            raise RuntimeError(f"Neočekávaná chyba při ukládání zálohy: {e}")

        finally:
            # Fáze 6: Zajistíme uzavření workbooku, pokud byl otevřen
            if workbook: # Ověření, zda workbook byl úspěšně načten
                try:
                    workbook.close()
                    logger.debug(f"Workbook '{self.active_filename}' uzavřen v add_or_update_employee_advance.")
                except Exception as close_err:
                    logger.warning(f"Chyba při zavírání workbooku '{self.active_filename}' v add_or_update_employee_advance: {close_err}")

    # Metody get_employee_advances a get_option_names mohou být upraveny, pokud je potřeba číst data
    # Zde je příklad pro get_option_names (get_employee_advances by bylo podobné)
    def get_option_names(self):
         """
         Získá názvy dvou hlavních možností záloh z listu 'Zálohy' v aktivním Excel souboru.

         Tyto názvy jsou typicky uloženy v buňkách B80 a D80.

         Životní cyklus operace se sešitem:
         1. Načtení sešitu: Volá `_get_active_workbook(read_only=True)` pro získání
            instance sešitu otevřené pouze pro čtení. Tím se minimalizuje riziko
            nechtěných změn a může to být efektivnější.
         2. Čtení dat: Pokud list "Zálohy" existuje, přečte hodnoty z buněk B80 a D80.
            Pokud hodnoty neexistují, použijí se výchozí názvy ("Option 1", "Option 2").
         3. Zavření sešitu: Ve `finally` bloku je zajištěno, že sešit je vždy
            uzavřen pomocí `workbook.close()`, aby se uvolnily systémové prostředky.

         Returns:
             tuple: Dvojice stringů (option1_name, option2_name).
                    V případě chyby nebo pokud list/buňky neexistují, vrací výchozí názvy.
         """
         workbook = None # Inicializace pro finally blok
         try:
              # Fáze 1: Načtení workbooku v režimu read-only
              # Volající (tato metoda) je zodpovědný za uzavření.
              workbook = self._get_active_workbook(read_only=True)
              
              # Fáze 2: Čtení dat
              if self.ZALOHY_SHEET_NAME in workbook.sheetnames:
                   sheet = workbook[self.ZALOHY_SHEET_NAME]
                   option1 = sheet["B80"].value or Config.DEFAULT_ADVANCE_OPTION_1
                   option2 = sheet["D80"].value or Config.DEFAULT_ADVANCE_OPTION_2
                   option3 = sheet["F80"].value or Config.DEFAULT_ADVANCE_OPTION_3
                   option4 = sheet["H80"].value or Config.DEFAULT_ADVANCE_OPTION_4
                   return str(option1).strip(), str(option2).strip(), str(option3).strip(), str(option4).strip()
              else:
                   logger.warning(f"List '{self.ZALOHY_SHEET_NAME}' nenalezen v souboru '{self.active_filename}' při čtení názvů možností. Používají se výchozí názvy.")
                   return Config.DEFAULT_ADVANCE_OPTION_1, Config.DEFAULT_ADVANCE_OPTION_2 # Výchozí hodnoty
         except (FileNotFoundError, IOError) as e:
              # Tyto chyby jsou již logovány v _get_active_workbook, zde jen specifičtější kontext
              logger.error(f"Chyba souboru při čtení názvů možností z '{self.active_filename}': {e}")
              return Config.DEFAULT_ADVANCE_OPTION_1, Config.DEFAULT_ADVANCE_OPTION_2 # Výchozí v případě chyby souboru
         except Exception as e:
              logger.error(f"Neočekávaná chyba při čtení názvů možností z '{self.active_filename}': {e}", exc_info=True)
              return Config.DEFAULT_ADVANCE_OPTION_1, Config.DEFAULT_ADVANCE_OPTION_2 # Výchozí v případě jiné chyby
         finally:
              # Fáze 3: Zajistíme uzavření workbooku, pokud byl otevřen
              if workbook: # Ověření, zda workbook byl úspěšně načten
                   try:
                        workbook.close()
                        logger.debug(f"Read-only workbook '{self.active_filename}' uzavřen v get_option_names.")
                   except Exception as close_err:
                        logger.warning(f"Chyba při zavírání read-only workbooku '{self.active_filename}' v get_option_names: {close_err}")


# Blok if __name__ == "__main__": není relevantní pro tuto třídu v kontextu aplikace
