import tkinter as tk
from tkinter import ttk, simpledialog, messagebox
from tkcalendar import Calendar
from datetime import datetime
import logging
import json
import os
from excel_manager import ExcelManager
from zalohy_manager import ZalohyManager

# Aktualizace cesty k nastavení
SETTINGS_FILE_PATH = '/home/Cowley/hodiny/data/settings.json'

# Načtení nastavení z JSON souboru
def load_settings():
    if os.path.exists(SETTINGS_FILE_PATH):
        with open(SETTINGS_FILE_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    else:
        return {
            'start_time': '07:00',
            'end_time': '18:00',
            'lunch_duration': 1
        }

class TimeRecord:
    def __init__(self, employee_manager):
        self.employee_manager = employee_manager
        self.excel_manager = ExcelManager()
        self.settings = load_settings()  # Načte nastavení při inicializaci

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
        # Načtení přednastavených hodnot ze settings
        default_start_time = self.settings.get('start_time', '07:00')
        default_end_time = self.settings.get('end_time', '18:00')
        default_lunch_duration = self.settings.get('lunch_duration', 1)

        # Dialogová okna s předvyplněnými hodnotami
        self.zacatek = simpledialog.askstring(
        "Začátek práce", "Zadejte čas začátku (HH:MM):", initialvalue=default_start_time)
        self.konec = simpledialog.askstring(
        "Konec práce", "Zadejte čas konce (HH:MM):", initialvalue=default_end_time)
        self.obed = simpledialog.askfloat(
        "Délka oběda", "Zadejte délku oběda v hodinách:", initialvalue=default_lunch_duration)

        if self.zacatek and self.konec:
            try:
                zacatek_cas = datetime.strptime(self.zacatek, "%H:%M")
                konec_cas = datetime.strptime(self.konec, "%H:%M")
                pracovni_doba = (konec_cas - zacatek_cas).total_seconds() / 3600 - self.obed
                logging.info(f"Čisté odpracované hodiny: {pracovni_doba:.2f}")
                messagebox.showinfo("Odpracované hodiny", f"Čisté odpracované hodiny: {pracovni_doba:.2f}")
            except ValueError as e:
                logging.error(f"Chyba při zpracování časů: {e}")
                messagebox.showerror("Chyba", f"Chyba při zpracování časů: {e}")
        else:
            pracovni_doba = 0

        self.pracovni_doba = pracovni_doba

    def ulozit_zaznam(self):
        if not hasattr(self, 'vybrane_datum') or not hasattr(self, 'zacatek') or not hasattr(self, 'konec') or not hasattr(self, 'obed'):
            messagebox.showerror("Chyba", "Nejsou zadány všechny potřebné údaje.")
            logging.error("Nejsou zadány všechny potřebné údaje.")
            return

        try:
            datum = self.vybrane_datum.strftime('%Y-%m-%d')

            if not (datum and self.zacatek and self.konec and self.obed):
                logging.error("Nejsou předány platné parametry pro uložení.")
                messagebox.showerror("Chyba", "Chybí některý z údajů.")
                return

            logging.info("Pokus o uložení pracovní doby.")
            self.excel_manager.ulozit_pracovni_dobu(
                datum, self.zacatek, self.konec, self.obed, self.employee_manager.vybrani_zamestnanci
            )
            messagebox.showinfo("Uloženo", "Záznam byl úspěšně uložen do Excel souboru.")
            logging.info(f"Úspěšně uloženo: Datum {datum}, Začátek {self.zacatek}, Konec {self.konec}, Oběd {self.obed}")
        except Exception as e:
            messagebox.showerror("Chyba", f"Nepodařilo se uložit záznam: {e}")
            logging.error(f"Nepodařilo se uložit záznam: {e}")
        return

class EmployeeManager:
    def __init__(self, employees_file):
        self.employees_file = employees_file
        self.load_employees()

    def load_employees(self):
        try:
            if os.path.exists(self.employees_file):
                with open(self.employees_file, 'r', encoding='utf-8') as f:
                    self.employees = json.load(f)
            else:
                self.employees = []
        except Exception as e:
            logging.error(f"Chyba při načítání zaměstnanců: {e}")
            self.employees = []

    def add_employee(self, name, role):
        self.employees.append({'name': name, 'role': role})
        self.save_employees()

    def save_employees(self):
        try:
            with open(self.employees_file, 'w', encoding='utf-8') as f:
                json.dump(self.employees, f, ensure_ascii=False, indent=4)
        except Exception as e:
            logging.error(f"Chyba při ukládání zaměstnanců: {e}")

# Příklad použití
if __name__ == "__main__":
    excel_path = '/home/Cowley/hodiny/excel_data'
    employees_file = '/home/Cowley/hodiny/employees.json'
    employee_manager = EmployeeManager(employees_file)
    time_record = TimeRecord(employee_manager)

    root = tk.Tk()
    root.title("TimeRecord")
    ttk.Button(root, text="Záznam pracovní doby", command=lambda: time_record.zaznam_pracovni_doby(root)).pack(pady=20)
    root.mainloop()
