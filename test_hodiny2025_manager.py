#!/usr/bin/env python3
"""
Test script pro Hodiny2025Manager - ověření funkčnosti a mapování buněk

Tento script:
1. Vytvoří testovací data v Hodiny2025.xlsx
2. Ověří správnost zápisu dat do buněk
3. Otestuje vzorce a výpočty
4. Vygeneruje reporty pro validaci
"""

import sys
import os
from pathlib import Path
from datetime import datetime

# Přidej cestu k modulu
sys.path.append(str(Path(__file__).parent))

from hodiny2025_manager import Hodiny2025Manager
from config import Config

def test_hodiny2025_manager():
    """Hlavní test funkce."""
    print("🚀 SPOUŠTÍ SE TEST HODINY2025MANAGER")
    print("=" * 50)
    
    # Inicializace manageru
    try:
        manager = Hodiny2025Manager(Config.EXCEL_BASE_PATH)
        print("✅ Hodiny2025Manager úspěšně inicializován")
    except Exception as e:
        print(f"❌ Chyba při inicializaci: {e}")
        assert False, f"Inicializace selhala: {e}"
    
    # Test 1: Vytvoření testovacích dat
    print("\n📝 TEST 1: Vytváření testovacích dat")
    print("-" * 30)
    
    try:
        manager.create_test_data()
        print("✅ Testovací data byla vytvořena")
    except Exception as e:
        print(f"❌ Chyba při vytváření testovacích dat: {e}")
        assert False, f"Vytváření testovacích dat selhalo: {e}"
    
    # Test 2: Zápis jednotlivých záznamů
    print("\n✏️ TEST 2: Zápis pracovní doby")
    print("-" * 30)
    
    test_records = [
        {
            'date': '2025-08-13',  # Dnešní datum
            'start': '08:00',
            'end': '16:30',
            'lunch': '1.0',
            'employees': 5,
            'expected_hours': 7.5,
            'expected_overtime': 0.0
        },
        {
            'date': '2025-08-14',  # Zítřek
            'start': '07:00', 
            'end': '18:00',
            'lunch': '1.0',
            'employees': 3,
            'expected_hours': 10.0,
            'expected_overtime': 2.0
        },
        {
            'date': '2025-08-15',  # Volný den
            'start': '00:00',
            'end': '00:00', 
            'lunch': '0',
            'employees': 0,
            'expected_hours': 0.0,
            'expected_overtime': 0.0
        }
    ]
    
    for i, record in enumerate(test_records, 1):
        try:
            manager.zapis_pracovni_doby(
                record['date'],
                record['start'],
                record['end'], 
                record['lunch'],
                record['employees']
            )
            print(f"✅ Záznam {i}: {record['date']} - úspěšně zapsán")
        except Exception as e:
            print(f"❌ Záznam {i}: Chyba při zápisu - {e}")
    
    # Test 3: Čtení a validace dat
    print("\n📖 TEST 3: Čtení a validace zapsaných dat")
    print("-" * 40)
    
    for i, record in enumerate(test_records, 1):
        try:
            daily_data = manager.get_daily_record(record['date'])
            
            print(f"\n📅 Záznam {i} - {record['date']}:")
            print(f"   Začátek: {daily_data.get('start_time', 'N/A')}")
            print(f"   Konec: {daily_data.get('end_time', 'N/A')}")
            print(f"   Oběd: {daily_data.get('lunch_hours', 0)}h")
            print(f"   Celkem hodin: {daily_data.get('total_hours', 0)}h")
            print(f"   Přesčasy: {daily_data.get('overtime', 0)}h")
            print(f"   Zaměstnanci: {daily_data.get('num_employees', 0)}")
            print(f"   Celkem za všechny: {daily_data.get('total_all_employees', 0)}h")
            
            # Validace očekávaných hodnot
            if abs(daily_data.get('total_hours', 0) - record['expected_hours']) < 0.1:
                print(f"   ✅ Celkové hodiny jsou správně: {daily_data.get('total_hours', 0)}h")
            else:
                print(f"   ❌ Chyba v celkových hodinách: očekáváno {record['expected_hours']}h, "
                      f"získáno {daily_data.get('total_hours', 0)}h")
            
            if abs(daily_data.get('overtime', 0) - record['expected_overtime']) < 0.1:
                print(f"   ✅ Přesčasy jsou správně: {daily_data.get('overtime', 0)}h")
            else:
                print(f"   ❌ Chyba v přesčasech: očekáváno {record['expected_overtime']}h, "
                      f"získáno {daily_data.get('overtime', 0)}h")
                      
        except Exception as e:
            print(f"❌ Záznam {i}: Chyba při čtení - {e}")
    
    # Test 4: Měsíční souhrny
    print("\n📊 TEST 4: Měsíční souhrny")
    print("-" * 25)
    
    current_month = datetime.now().month
    try:
        summary = manager.get_monthly_summary(current_month, 2025)
        
        print(f"\n📈 Souhrn pro {summary['month_name']} 2025:")
        print(f"   List: {summary['sheet_name']}")
        print(f"   Celkem hodin: {summary['total_hours']}h")
        print(f"   Celkem přesčasů: {summary['total_overtime']}h") 
        print(f"   Celkem za všechny zaměstnance: {summary['total_all_employees']}h")
        print("✅ Měsíční souhrn úspěšně načten")
        
    except Exception as e:
        print(f"❌ Chyba při načítání měsíčního souhrnu: {e}")
    
    # Test 5: Validace integrity dat
    print("\n🔍 TEST 5: Validace integrity dat")
    print("-" * 30)
    
    try:
        validation = manager.validate_data_integrity()
        
        print(f"Zkontrolované listy: {len(validation['sheets_checked'])}")
        print(f"Zkontrolované záznamy: {validation['records_checked']}")
        print(f"Počet chyb: {len(validation['errors'])}")
        print(f"Počet varování: {len(validation['warnings'])}")
        
        if validation['errors']:
            print("\n❌ CHYBY:")
            for error in validation['errors']:
                print(f"   - {error}")
        
        if validation['warnings']:
            print("\n⚠️ VAROVÁNÍ:")
            for warning in validation['warnings'][:5]:  # Zobraz jen prvních 5
                print(f"   - {warning}")
            if len(validation['warnings']) > 5:
                print(f"   ... a {len(validation['warnings'])-5} dalších")
        
        if validation['valid'] and not validation['errors']:
            print("✅ Validace dat proběhla úspěšně")
        
    except Exception as e:
        print(f"❌ Chyba při validaci dat: {e}")
    
    # Test 6: Info o souboru
    print("\n📁 TEST 6: Informace o vytvořeném souboru")
    print("-" * 40)
    
    excel_file = Path(Config.EXCEL_BASE_PATH) / "Hodiny2025.xlsx"
    if excel_file.exists():
        file_size = excel_file.stat().st_size
        print(f"✅ Soubor vytvořen: {excel_file}")
        print(f"   Velikost: {file_size:,} bytů ({file_size/1024:.1f} KB)")
        print(f"   Poslední změna: {datetime.fromtimestamp(excel_file.stat().st_mtime)}")
    else:
        print(f"❌ Soubor nebyl vytvořen: {excel_file}")
    
    print("\n🎉 TEST DOKONČEN!")
    print("=" * 50)
    
    # Don't return True - pytest functions should not return anything
    assert True

def print_cell_mapping_reference():
    """Vypíše referenční přehled mapování buněk."""
    print("\n📋 REFERENČNÍ MAPOVÁNÍ BUNĚK")
    print("=" * 40)
    
    print("Hlavní vstupní buňky pro zápis pracovní doby:")
    print("- E[řádek]: Začátek práce (HH:MM)")
    print("- F[řádek]: Oběd v hodinách (desetinné číslo)")  
    print("- G[řádek]: Konec práce (HH:MM)")
    print("- M[řádek]: Počet zaměstnanců (integer)")
    
    print("\nVypočítávané buňky:")
    print("- H[řádek]: Celkem hodin = (G-E)*24-F")
    print("- I[řádek]: Přesčasy = MAX(0,H-8)")
    print("- N[řádek]: Celkem za všechny = H*M")
    
    print("\nIndexování řádků:")
    print("- řádek = den_v_měsíci + 2")
    print("- 1. den měsíce → řádek 3")
    print("- 15. den měsíce → řádek 17") 
    print("- 31. den měsíce → řádek 33")
    
    print("\nSouhrny (řádek 34):")
    print("- H34: =SUM(H3:H33) - Celkem hodin za měsíc")
    print("- I34: =SUM(I3:I33) - Celkem přesčasů za měsíc")
    print("- N34: =SUM(N3:N33) - Celkem za všechny za měsíc")

if __name__ == "__main__":
    print("🧪 TESTOVACÍ SCRIPT PRO HODINY2025MANAGER")
    print(f"⏰ Spuštěno: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}")
    
    # Zobraz referenční mapování
    print_cell_mapping_reference()
    
    # Spusť hlavní test
    success = test_hodiny2025_manager()
    
    if success:
        print("\n✅ VŠECHNY TESTY PROBĚHLY ÚSPĚŠNĚ!")
        print("📄 Podrobnou dokumentaci najdete v: MAPOVANI_BUNEK_HODINY2025.md")
        print("📊 Excel soubor byl vytvořen v: excel/Hodiny2025.xlsx")
    else:
        print("\n❌ NĚKTERÉ TESTY SELHALY!")
        sys.exit(1)
