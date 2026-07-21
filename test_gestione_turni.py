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

    addetto = Addetto("Mario Rossi", 40, 45, True)
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
    addetto1 = Addetto("Mario Rossi", 40, 45, True)
    addetto2 = Addetto("Luigi Bianchi", 36, 40, False)

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
        print(f"  - Ore per settimana: {stats['ore_per_settimana']}")

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

def test_ferie_rispettate():
    """La pianificazione non deve assegnare turni nei giorni di ferie"""
    print("\n=== TEST RISPETTO FERIE ===")

    manager = TurnoManager()
    manager.mese = 1
    manager.anno = 2025

    addetto = Addetto("Mario Rossi", 20, 40, False)
    giorno_ferie = datetime(2025, 1, 15)
    addetto.aggiungi_ferie(giorno_ferie)

    manager.aggiungi_addetto(addetto)
    manager.aggiungi_turno(Turno("Mattina", "08:00", "14:00"))

    assert manager.pianifica_turni(), "La pianificazione deve riuscire"

    assert giorno_ferie not in addetto.turni_assegnati, \
        f"Turno assegnato il {giorno_ferie.strftime('%d/%m/%Y')} nonostante le ferie"
    assert addetto.nome not in manager.pianificazione.get(giorno_ferie, {}), \
        f"{addetto.nome} presente in pianificazione il giorno di ferie"

    # puo_lavorare deve funzionare sia con datetime che con date
    assert not addetto.puo_lavorare(giorno_ferie), \
        "puo_lavorare deve restituire False per un datetime in ferie"
    assert not addetto.puo_lavorare(giorno_ferie.date()), \
        "puo_lavorare deve restituire False per una date in ferie"

    print("✓ Le ferie vengono rispettate dalla pianificazione")
    return True


def test_ore_minime_settimanali():
    """Il minimo contrattuale va garantito settimana per settimana"""
    print("\n=== TEST ORE MINIME SETTIMANALI ===")

    manager = TurnoManager()
    manager.mese = 1
    manager.anno = 2025

    # Due addetti e un solo turno da 6h: senza la Fase 2 per settimana,
    # ognuno riceverebbe circa metà delle ore
    addetto1 = Addetto("Mario Rossi", 24, 40, False)
    addetto2 = Addetto("Luigi Bianchi", 24, 40, False)
    manager.aggiungi_addetto(addetto1)
    manager.aggiungi_addetto(addetto2)
    manager.aggiungi_turno(Turno("Mattina", "08:00", "14:00"))

    assert manager.pianifica_turni(), "La pianificazione deve riuscire"

    for addetto in (addetto1, addetto2):
        for num_settimana, giorni_settimana in manager.get_settimane_mese().items():
            if not manager._settimana_completa_nel_mese(giorni_settimana[0]):
                continue  # settimane a cavallo di due mesi: minimo non esigibile
            ore = addetto.get_ore_settimana(num_settimana)
            assert ore >= addetto.ore_contratto, \
                f"{addetto.nome}: settimana {num_settimana} ha {ore}h, minimo {addetto.ore_contratto}h"
            assert ore <= addetto.ore_max_settimanale, \
                f"{addetto.nome}: settimana {num_settimana} ha {ore}h, massimo {addetto.ore_max_settimanale}h"

    print("✓ Minimo e massimo settimanale rispettati per ogni settimana completa")
    return True


def test_festivita():
    """Le festività nazionali italiane, incluse Pasqua e Pasquetta mobili, sono riconosciute"""
    print("\n=== TEST FESTIVITÀ ===")

    manager = TurnoManager()

    # Festività a data fissa
    for mese, giorno in [(1, 1), (1, 6), (4, 25), (5, 1), (6, 2), (8, 15), (11, 1), (12, 8), (12, 25), (12, 26)]:
        assert manager.is_festivo(datetime(2025, mese, giorno)), f"{giorno}/{mese} deve essere festivo"

    # Pasqua e Pasquetta: mobili, cambiano ogni anno
    assert TurnoManager.calcola_pasqua(2025) == datetime(2025, 4, 20).date(), "Pasqua 2025 è il 20 aprile"
    assert TurnoManager.calcola_pasqua(2026) == datetime(2026, 4, 5).date(), "Pasqua 2026 è il 5 aprile"
    assert manager.is_festivo(datetime(2025, 4, 21)), "Pasquetta 2025 (21/04) deve essere festiva"
    assert manager.is_festivo(datetime(2026, 4, 6)), "Pasquetta 2026 (06/04) deve essere festiva"
    assert not manager.is_festivo(datetime(2026, 4, 20)), "Il 20/04/2026 non è festivo (Pasqua non è fissa)"

    # Un giorno feriale qualunque non è festivo
    assert not manager.is_festivo(datetime(2025, 3, 12)), "Il 12/03 non deve essere festivo"

    print("✓ Festività fisse e mobili riconosciute correttamente")
    return True


def main():
    """Funzione principale di test"""
    print("="*60)
    print("   TEST GESTIONE TURNI".center(60))
    print("="*60)

    try:
        if (test_addetto() and test_turno() and test_manager()
                and test_ferie_rispettate() and test_ore_minime_settimanali()
                and test_festivita()):
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
