#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GESTIONE TURNI NEGOZIO - PROGRAMMA PRINCIPALE
Permette di scegliere tra interfaccia a riga di comando (CLI) e interfaccia grafica (GUI)
"""

import sys


def main():
    """Menu di scelta tra CLI e GUI"""
    print("\n" + "="*60)
    print("   GESTIONE TURNI NEGOZIO".center(60))
    print("="*60)
    print("\nScegli modalità di utilizzo:\n")
    print("1. Interfaccia Grafica (GUI) - PyQt6")
    print("2. Interfaccia Testuale (CLI) - Menu Interattivo")
    print("0. Esci")
    print("\n" + "-"*60)

    scelta = input("Seleziona un'opzione: ").strip()

    if scelta == '1':
        try:
            from gui_turni import main as main_gui
            main_gui()
        except ImportError:
            print("\nErrore: PyQt6 non è installato.")
            print("Installa con: pip install PyQt6")
            sys.exit(1)

    elif scelta == '2':
        from gestione_turni import MenuInterattivo
        menu = MenuInterattivo()
        menu.menu_principale()

    elif scelta == '0':
        print("\nArrivederci!")
        sys.exit(0)

    else:
        print("Opzione non valida")
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nProgramma interrotto dall'utente.")
        sys.exit(0)
    except Exception as e:
        print(f"\nErrore non gestito: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
