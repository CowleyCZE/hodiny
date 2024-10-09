import tkinter as tk
from tkinter import ttk, simpledialog, messagebox
from tkcalendar import Calendar
from datetime import datetime
import logging
from excel_manager import ExcelManager
from zalohy_manager import ZalohyManager

class TimeRecord:
    def __init__(self, employee_manager):
        self.employee_manager = employee_manager
        self.excel_manager = ExcelManager()

    def zaznam_pracovni_doby(self, master):
        self.zaznam_okno = tk.Toplevel(master)
        self.zaznam_okno.title("Záznam pracovní doby")

        ttk.Button(self.zaznam_okno, text="Vybrát datum", command=self.vybrat_datum).pack(pady=10)
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
        
        self.zadat_casy()

    def zadat_casy(self):
        self.zacatek = simpledialog.askstring("Začátek práce", "Zadejte čas začátku (HH:MM) nebo X:")
        self.konec = simpledialog.askstring("Konec práce", "Zadejte čas konce (HH:MM) nebo X:")
        self.obed = simpledialog.askfloat("Délka oběda", "Zadejte délku oběda v hodinách:")

        if self.zacatek != 'X' and self.konec != 'X':
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
            self.excel_manager.ulozit_pracovni_dobu(self.vybrane_datum, self.zacatek, self.konec, self.obed, self.employee_manager.vybrani_zamestnanci)
            messagebox.showinfo("Uloženo", "Záznam byl úspěšně uložen do Excel souboru.")
            logging.info(f"Úspěšně uloženo: Datum {self.vybrane_datum}, Začátek {self.zacatek}, Konec {self.konec}, Oběd {self.obed}")
        except Exception as e:
            messagebox.showerror("Chyba", f"Nepodařilo se uložit záznam: {e}")
            logging.error(f"Nepodařilo se uložit záznam: {e}")