#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script di test per verifica del programma di gestione turni
"""

from gestione_turni import Addetto, Turno, TurnoManager
from datetime import datetime

def test_addetto():
    """Test della classe Addetto"""
    print("\n=== TEST CLASSE ADDETTO ===")

    addetto = Addetto("Mario Rossi", 40, 160, True)
    print(f"✓ Addetto creato: {addetto}")

    # Test giorni riposo
    addetto.aggiungi_giorno_riposo(6)  # domenica
    print(f"✓ Domenica aggiunta come giorno di riposo")

    # Test ferie
    data_feria = datetime(2025, 1, 15)
    addetto.aggiungi_ferie(data_feria)
    print(f"✓ Feria aggiunta: {data_feria.strftime('%d/%m/%Y')}")

    # Test disponibilità
    data_test = datetime(2025, 1, 13)  # lunedì
    print(f"✓ Può lavorare il 13/01/2025? {addetto.puo_lavorare(data_test)}")

    return True

def test_turno():
    """Test della classe Turno"""
    print("\n=== TEST CLASSE TURNO ===")

    turno1 = Turno("Mattina", "08:00", "14:00")
    turno2 = Turno("Pomeriggio", "14:00", "20:00")
    turno3 = Turno("Sera", "20:00", "21:00")

    print(f"✓ {turno1}")
    print(f"✓ {turno2}")
    print(f"✓ {turno3}")

    return True

def test_manager():
    """Test della classe TurnoManager"""
    print("\n=== TEST CLASSE TURNOMANAGER ===")

    manager = TurnoManager()
    print(f"✓ Manager creato per {manager._nome_mese()} {manager.anno}")

    # Aggiungi addetti
    addetto1 = Addetto("Mario Rossi", 40, 160, True)
    addetto2 = Addetto("Luigi Bianchi", 36, 144, False)

    manager.aggiungi_addetto(addetto1)
    manager.aggiungi_addetto(addetto2)
    print(f"✓ Aggiunti {len(manager.addetti)} addetti")

    # Aggiungi turni
    turno1 = Turno("Mattina", "08:00", "14:00")
    turno2 = Turno("Pomeriggio", "14:00", "20:00")

    manager.aggiungi_turno(turno1)
    manager.aggiungi_turno(turno2)
    print(f"✓ Aggiunti {len(manager.turni)} turni")

    # Test giorni festivi
    data_natale = datetime(2025, 12, 25)
    print(f"✓ 25/12 è festivo? {manager.is_festivo(data_natale)}")

    # Test giorni domeniche
    data_domenica = datetime(2025, 1, 12)
    print(f"✓ 12/01/2025 è domenica? {manager.is_domenica(data_domenica)}")

    # Test get giorni mese
    giorni = manager.get_giorni_mese()
    print(f"✓ Giorni lavorativi nel mese: {len(giorni)}")

    # Test pianificazione
    print("\nAvvio pianificazione...")
    if manager.pianifica_turni():
        print("✓ Pianificazione completata")

        # Verifica statistiche
        stats = manager.genera_statistiche()
        print(f"\nStatistiche generate:")
        print(f"  - Ore totali per addetto: {stats['ore_totali_per_addetto']}")
        print(f"  - Giorni lavorati: {stats['giorni_lavorati_per_addetto']}")

        # Test export Excel
        try:
            percorso = manager.esporta_excel("test_turni.xlsx")
            print(f"\n✓ Excel esportato: {percorso}")
            return True
        except Exception as e:
            print(f"✗ Errore esportazione: {e}")
            return False
    else:
        print("✗ Errore nella pianificazione")
        return False

def main():
    """Funzione principale di test"""
    print("="*60)
    print("   TEST GESTIONE TURNI".center(60))
    print("="*60)

    try:
        if test_addetto() and test_turno() and test_manager():
            print("\n" + "="*60)
            print("   TUTTI I TEST COMPLETATI CON SUCCESSO ✓".center(60))
            print("="*60)
            return 0
        else:
            print("\n✗ Alcuni test sono falliti")
            return 1

    except Exception as e:
        print(f"\n✗ Errore durante i test: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    exit(main())
