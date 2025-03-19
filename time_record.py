# Standard library imports
import tkinter as tk
from tkinter import ttk, simpledialog, messagebox
from datetime import datetime
import logging
import os
import re
import sys
from typing import Optional, Union
from utils.logger import setup_logger

logger = setup_logger('time_record')

# Third party imports
try:
    from tkcalendar import Calendar
except ImportError as e:
    logger.error(f"Chyba při importu tkcalendar: {e}")
    try:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "tkcalendar"])
        from tkcalendar import Calendar
    except Exception as e:
        messagebox.showerror("Error", f"Nepodařilo se nainstalovat tkcalendar: {e}")
        raise

# Local imports
from excel_manager import ExcelManager
from employee_management import EmployeeManager  # Import správné třídy
from settings import load_settings  # Import funkce pro načítání nastavení
from config import Config
from utils.logger import setup_logger

logger = setup_logger('time_record')

def check_dependencies():
    try:
        import tkcalendar
    except ImportError:
        print("Instaluji chybějící závislosti...")
        import subprocess
        subprocess.check_call(["pip", "install", "tkcalendar"])

class TimeRecord:
    def __init__(self, employee_manager):
        self.employee_manager = employee_manager
        self.excel_manager = ExcelManager(Config.EXCEL_BASE_PATH, Config.EXCEL_FILE_NAME)
        self.settings = None
        self.pracovni_doba = 0
        self.vybrane_datum = None
        self.zacatek = None
        self.konec = None
        self.obed = None
        self.time_pattern = re.compile(r'^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$')

    def get_current_settings(self):
        """Načte aktuální nastavení ze souboru"""
        try:
            self.settings = load_settings()
            logging.info("Nastavení úspěšně načteno")
        except Exception as e:
            logging.error(f"Chyba při načítání nastavení: {e}")
            self.settings = Config.get_default_settings()

    def validate_time(self, time_str: str) -> bool:
        """Validuje formát času"""
        if not time_str:
            raise ValueError("Čas nesmí být prázdný")
        
        if not self.time_pattern.match(time_str):
            raise ValueError("Neplatný formát času. Použijte formát HH:MM (00:00-23:59)")
        
        return True

    def validate_lunch_duration(self, duration: Union[int, float]) -> bool:
        """Validuje délku obědové pauzy"""
        if duration is None:
            raise ValueError("Délka oběda nesmí být prázdná")
        
        if not isinstance(duration, (int, float)):
            raise ValueError("Délka oběda musí být číslo")
            
        if duration < 0 or duration > 4:
            raise ValueError("Délka oběda musí být mezi 0 a 4 hodinami")
            
        return True

    def validate_date(self, date: datetime) -> bool:
        """Validuje vybrané datum"""
        if not date:
            raise ValueError("Datum musí být vybráno")
            
        if not isinstance(date, datetime):
            raise ValueError("Neplatný formát data")
            
        today = datetime.now().date()
        selected_date = date.date()
            
        # Kontrola víkendu
        if selected_date.weekday() >= 5:
            raise ValueError("Nelze vybrat víkend")
            
        if selected_date > today:
            raise ValueError("Nelze vybrat datum v budoucnosti")
            
        if (today - selected_date).days > 365:
            raise ValueError("Nelze vybrat datum starší než jeden rok")
            
        return True

    def validate_time_range(self, start_time: str, end_time: str) -> bool:
        """Validuje rozsah časů"""
        if not all([start_time, end_time]):
            raise ValueError("Časy nesmí být prázdné")

        start = datetime.strptime(start_time, "%H:%M")
        end = datetime.strptime(end_time, "%H:%M")
        
        if end <= start:
            raise ValueError("Konec práce musí být později než začátek")
            
        duration = (end - start).total_seconds() / 3600
        if duration > 24:
            raise ValueError("Pracovní doba nesmí přesáhnout 24 hodin")
            
        if duration < 0.5:
            raise ValueError("Pracovní doba musí být alespoň 30 minut")
            
        return True

    def zaznam_pracovni_doby(self, master):
        self.zaznam_okno = tk.Toplevel(master)
        self.zaznam_okno.title("Záznam pracovní doby")

        ttk.Button(self.zaznam_okno, text="Vybrat datum", command=self.vybrat_datum).pack(pady=10)
        ttk.Button(self.zaznam_okno, text="Uložit", command=self.ulozit_zaznam).pack(pady=10)

    def vybrat_datum(self):
        self.kalendar_okno = tk.Toplevel(self.zaznam_okno)
        self.kalendar = Calendar(self.kalendar_okno, selectmode="day")
        self.kalendar.pack(pady=10)
        ttk.Button(self.kalendar_okno, text="Vybrat", command=self.nastavit_datum).pack(pady=10)

    def nastavit_datum(self):
        self.vybrane_datum = self.kalendar.selection_get()
        self.kalendar_okno.destroy()
        logging.info(f"Datum vybráno: {self.vybrane_datum}")
        messagebox.showinfo("Datum vybráno", f"Vybráno datum: {self.vybrane_datum}")

        # Načte přednastavené časy ze settings.json
        self.zadat_casy()

    def zadat_casy(self):
        # Načteme aktuální nastavení před zobrazením dialogů
        self.get_current_settings()
        
        default_start_time = self.settings.get('start_time', '07:00')
        default_end_time = self.settings.get('end_time', '18:00')
        default_lunch_duration = self.settings.get('lunch_duration', 1)

        logging.info(f"Použití přednastavených hodnot - začátek: {default_start_time}, "
                    f"konec: {default_end_time}, oběd: {default_lunch_duration}")

        while True:
            try:
                self.zacatek = simpledialog.askstring(
                    "Začátek práce", 
                    "Zadejte čas začátku (HH:MM):", 
                    initialvalue=default_start_time
                )
                
                if self.zacatek is None:  # Uživatel stiskl Cancel
                    return
                    
                self.validate_time(self.zacatek)
                break
            except ValueError as e:
                messagebox.showerror("Chyba", str(e))
                logging.warning(f"Neplatný čas začátku: {e}")

        while True:
            try:
                self.konec = simpledialog.askstring(
                    "Konec práce", 
                    "Zadejte čas konce (HH:MM):", 
                    initialvalue=default_end_time
                )
                
                if self.konec is None:  # Uživatel stiskl Cancel
                    return
                    
                self.validate_time(self.konec)
                self.validate_time_range(self.zacatek, self.konec)
                break
            except ValueError as e:
                messagebox.showerror("Chyba", str(e))
                logging.warning(f"Neplatný čas konce: {e}")

        while True:
            try:
                self.obed = simpledialog.askfloat(
                    "Délka oběda", 
                    "Zadejte délku oběda v hodinách:",
                    initialvalue=default_lunch_duration
                )
                
                if self.obed is None:  # Uživatel stiskl Cancel
                    return
                    
                self.validate_lunch_duration(self.obed)
                break
            except ValueError as e:
                messagebox.showerror("Chyba", str(e))
                logging.warning(f"Neplatná délka oběda: {e}")

        try:
            zacatek_cas = datetime.strptime(self.zacatek, "%H:%M")
            konec_cas = datetime.strptime(self.konec, "%H:%M")
            self.pracovni_doba = (konec_cas - zacatek_cas).total_seconds() / 3600 - self.obed
            
            if self.pracovni_doba < 0:
                raise ValueError("Celková pracovní doba nemůže být záporná")
                
            logging.info(f"Čisté odpracované hodiny: {self.pracovni_doba:.2f}")
            messagebox.showinfo("Odpracované hodiny", f"Čisté odpracované hodiny: {self.pracovni_doba:.2f}")
        except ValueError as e:
            logging.error(f"Chyba při výpočtu pracovní doby: {e}")
            messagebox.showerror("Chyba", str(e))
            self.pracovni_doba = 0

    def ulozit_zaznam(self):
        try:
            if not hasattr(self, 'vybrane_datum') or not self.vybrane_datum:
                raise ValueError("Není vybráno datum")

            self.validate_date(self.vybrane_datum)
            datum = self.vybrane_datum.strftime('%Y-%m-%d')

            if not all([self.zacatek, self.konec, self.obed is not None]):
                raise ValueError("Nejsou vyplněny všechny časové údaje")
                
            if self.pracovni_doba <= 0:
                raise ValueError("Celková pracovní doba musí být kladná")

            logging.info("Ukládání pracovní doby...")
            self.excel_manager.ulozit_pracovni_dobu(
                datum, self.zacatek, self.konec, self.obed, 
                self.employee_manager.vybrani_zamestnanci
            )
            messagebox.showinfo("Uloženo", "Záznam byl úspěšně uložen do Excel souboru.")
            logging.info(f"Úspěšně uloženo: Datum {datum}, Začátek {self.zacatek}, " +
                        f"Konec {self.konec}, Oběd {self.obed}")
        except ValueError as e:
            messagebox.showerror("Chyba validace", str(e))
            logging.error(f"Chyba validace: {e}")
        except Exception as e:
            messagebox.showerror("Chyba", f"Nepodařilo se uložit záznam: {e}")
            logging.error(f"Nepodařilo se uložit záznam: {e}")

# Příklad použití
if __name__ == "__main__":
    employee_manager = EmployeeManager(Config.DATA_PATH)
    time_record = TimeRecord(employee_manager)

    root = tk.Tk()
    root.title("TimeRecord")
    ttk.Button(root, text="Záznam pracovní doby", command=lambda: time_record.zaznam_pracovni_doby(root)).pack(pady=20)
    root.mainloop()
