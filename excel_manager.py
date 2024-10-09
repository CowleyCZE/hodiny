import os
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import logging

logging.basicConfig(filename='evidence_pracovni_doby.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

class ExcelManager:
    def __init__(self):
        self.excel_cesta = "Hodiny_Cap.xlsx"
        self.TEMPLATE_SHEET_NAME = 'Týden'

    def nacti_nebo_vytvor_excel(self):
        try:
            if os.path.exists(self.excel_cesta):
                try:
                    workbook = load_workbook(self.excel_cesta)
                    logging.info(f"Načten existující Excel soubor: {self.excel_cesta}")
                except Exception as e:
                    logging.warning(f"Nelze načíst existující soubor, vytvářím nový: {e}")
                    workbook = Workbook()
                    workbook.save(self.excel_cesta)
                    logging.info(f"Vytvořen nový Excel soubor se stejným názvem: {self.excel_cesta}")
            else:
                workbook = Workbook()
                workbook.save(self.excel_cesta)
                logging.info(f"Vytvořen nový Excel soubor: {self.excel_cesta}")
            return workbook
        except Exception as e:
            logging.error(f"Chyba při načítání nebo vytváření Excel souboru: {e}")
            raise

    def ziskej_nebo_vytvor_list(self, workbook, datum):
        try:
            cislo_tydne = datum.isocalendar()[1]
            nazev_listu = f"Týden {cislo_tydne}"

            if nazev_listu not in workbook.sheetnames:
                if self.TEMPLATE_SHEET_NAME in workbook.sheetnames:
                    sablona = workbook[self.TEMPLATE_SHEET_NAME]
                    novy_list = workbook.copy_worksheet(sablona)
                    novy_list.title = nazev_listu
                    novy_list['A80'] = nazev_listu
                else:
                    novy_list = workbook.create_sheet(title=nazev_listu)
                    self.inicializuj_list(novy_list, datum)
                logging.info(f"Vytvořen nový list '{nazev_listu}'.")
            else:
                novy_list = workbook[nazev_listu]
                logging.info(f"List '{nazev_listu}' již existuje.")

            return novy_list
        except Exception as e:
            logging.error(f"Chyba při získávání nebo vytváření listu: {e}")
            raise

    def inicializuj_list(self, sheet, datum):
        # Nastavení hlavičky a data pro každý den v týdnu
        dny = ["Pondělí", "Úterý", "Středa", "Čtvrtek", "Pátek", "Sobota", "Neděle"]
        prvni_den_tydne = datum - timedelta(days=datum.weekday())
        for i, den in enumerate(dny):
            sheet.cell(row=6, column=2 + i * 2, value=den)
            datum_bunky = prvni_den_tydne + timedelta(days=i)
            sheet.cell(row=80, column=2 + i * 2, value=datum_bunky.strftime("%d.%m.%Y"))

    def ulozit_pracovni_dobu(self, datum, zacatek, konec, obed, vybrani_zamestnanci):
        try:
            workbook = self.nacti_nebo_vytvor_excel()
            sheet = self.ziskej_nebo_vytvor_list(workbook, datum)

            den_v_tydnu = datum.weekday()
            sheet.cell(row=7, column=2 + den_v_tydnu * 2, value=zacatek)
            sheet.cell(row=7, column=3 + den_v_tydnu * 2, value=konec)
            sheet.cell(row=80, column=2 + datum.weekday() * 2, value=datum.strftime("%d.%m.%Y"))

            if zacatek != 'X' and konec != 'X':
                zacatek_cas = datetime.strptime(zacatek, "%H:%M")
                konec_cas = datetime.strptime(konec, "%H:%M")
                pracovni_doba = max((konec_cas - zacatek_cas).total_seconds() / 3600 - obed, 0)
                sheet.cell(row=8, column=2 + den_v_tydnu * 2, value=pracovni_doba)
                
                # Zápis pracovní doby pro vybrané zaměstnance
                for i, zamestnanec in enumerate(vybrani_zamestnanci):
                    row = 9 + i  # Začínáme od řádku 9 pro zaměstnance
                    sheet.cell(row=row, column=1, value=zamestnanec)
                    sheet.cell(row=row, column=2 + den_v_tydnu * 2, value=pracovni_doba)
            else:
                sheet.cell(row=8, column=2 + den_v_tydnu * 2, value='X')
                sheet.cell(row=9, column=2 + den_v_tydnu * 2, value='X')
                
                # Zápis 'X' pro vybrané zaměstnance v případě nepracovního dne
                for i, zamestnanec in enumerate(vybrani_zamestnanci):
                    row = 10 + i
                    sheet.cell(row=row, column=1, value=zamestnanec)
                    sheet.cell(row=row, column=2 + den_v_tydnu * 2, value='X')

            workbook.save(self.excel_cesta)
            logging.info(f"Data úspěšně uložena do souboru: {self.excel_cesta}")
        except Exception as e:
            logging.error(f"Nepodařilo se uložit pracovní dobu: {e}")
            raise

    def nacti_data_pro_tyden(self, datum):
        try:
            workbook = self.nacti_nebo_vytvor_excel()
            sheet = self.ziskej_nebo_vytvor_list(workbook, datum)

            data = []
            for i in range(7):  # Pro každý den v týdnu
                den_data = {
                    "datum": sheet.cell(row=80, column=2 + i * 2).value,
                    "zacatek": sheet.cell(row=7, column=2 + i * 2).value,
                    "konec": sheet.cell(row=7, column=3 + i * 2).value,
                    "pracovni_doba": sheet.cell(row=8, column=2 + i * 2).value
                }
                data.append(den_data)

            return data
        except Exception as e:
            logging.error(f"Chyba při načítání dat pro týden: {e}")
            raise

if __name__ == "__main__":
    # Zde můžete přidat testovací kód pro ověření funkčnosti ExcelManageru
    pass
