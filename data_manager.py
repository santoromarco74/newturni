#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DATA MANAGER - GESTIONE PERSISTENZA DATI
Salva e carica addetti e turni da file JSON
"""

import json
import os
from datetime import datetime
from typing import List, Dict, Any
from gestione_turni import Addetto, Turno


class DataManager:
    """Gestisce il salvataggio e caricamento dei dati"""

    def __init__(self, nome_file: str = "dati_turni.json"):
        """
        Inizializza il DataManager

        Args:
            nome_file: Nome del file JSON per salvare i dati
        """
        self.nome_file = nome_file
        self.dati = {
            'addetti': [],
            'turni': [],
            'ultimo_aggiornamento': None
        }

    def salva_dati(self, addetti: List[Addetto], turni: List[Turno]) -> bool:
        """
        Salva addetti e turni nel file JSON

        Args:
            addetti: Lista di addetti
            turni: Lista di turni

        Returns:
            True se il salvataggio è riuscito, False altrimenti
        """
        try:
            dati = {
                'addetti': self._serializza_addetti(addetti),
                'turni': self._serializza_turni(turni),
                'ultimo_aggiornamento': datetime.now().isoformat()
            }

            with open(self.nome_file, 'w', encoding='utf-8') as f:
                json.dump(dati, f, indent=2, ensure_ascii=False)

            return True
        except Exception as e:
            print(f"Errore durante il salvataggio: {e}")
            return False

    def carica_dati(self) -> tuple:
        """
        Carica addetti e turni dal file JSON

        Returns:
            Tupla (addetti, turni) se il caricamento è riuscito, ([], []) altrimenti
        """
        if not os.path.exists(self.nome_file):
            return [], []

        try:
            with open(self.nome_file, 'r', encoding='utf-8') as f:
                dati = json.load(f)

            addetti = self._deserializza_addetti(dati.get('addetti', []))
            turni = self._deserializza_turni(dati.get('turni', []))

            return addetti, turni
        except Exception as e:
            print(f"Errore durante il caricamento: {e}")
            return [], []

    def _serializza_addetti(self, addetti: List[Addetto]) -> List[Dict[str, Any]]:
        """Serializza gli addetti in formato JSON"""
        risultato = []

        for addetto in addetti:
            dati_addetto = {
                'nome': addetto.nome,
                'ore_contratto': addetto.ore_contratto,
                'ore_max_settimanale': addetto.ore_max_settimanale,
                'straordinario': addetto.straordinario,
                'giorni_riposo': sorted(list(addetto.giorni_riposo)),
                'ferie_permessi': [d.isoformat() for d in addetto.ferie_permessi]
            }
            risultato.append(dati_addetto)

        return risultato

    def _deserializza_addetti(self, dati: List[Dict[str, Any]]) -> List[Addetto]:
        """Deserializza gli addetti dal formato JSON"""
        risultato = []

        for dati_addetto in dati:
            try:
                addetto = Addetto(
                    nome=dati_addetto['nome'],
                    ore_contratto=dati_addetto['ore_contratto'],
                    ore_max_settimanale=dati_addetto['ore_max_settimanale'],
                    straordinario=dati_addetto['straordinario']
                )

                # Aggiungi giorni di riposo
                for giorno in dati_addetto.get('giorni_riposo', []):
                    addetto.aggiungi_giorno_riposo(giorno)

                # Aggiungi ferie
                for feria_str in dati_addetto.get('ferie_permessi', []):
                    try:
                        feria = datetime.fromisoformat(feria_str).date()
                        addetto.aggiungi_ferie(feria)
                    except ValueError:
                        # Ignora date non valide
                        pass

                risultato.append(addetto)
            except Exception as e:
                print(f"Errore nel caricamento addetto: {e}")
                continue

        return risultato

    def _serializza_turni(self, turni: List[Turno]) -> List[Dict[str, Any]]:
        """Serializza i turni in formato JSON"""
        risultato = []

        for turno in turni:
            dati_turno = {
                'nome': turno.nome,
                'ora_inizio': turno.ora_inizio,
                'ora_fine': turno.ora_fine
            }
            risultato.append(dati_turno)

        return risultato

    def _deserializza_turni(self, dati: List[Dict[str, Any]]) -> List[Turno]:
        """Deserializza i turni dal formato JSON"""
        risultato = []

        for dati_turno in dati:
            try:
                turno = Turno(
                    nome=dati_turno['nome'],
                    ora_inizio=dati_turno['ora_inizio'],
                    ora_fine=dati_turno['ora_fine']
                )
                risultato.append(turno)
            except Exception as e:
                print(f"Errore nel caricamento turno: {e}")
                continue

        return risultato

    def elimina_dati(self) -> bool:
        """Elimina il file dei dati salvati"""
        try:
            if os.path.exists(self.nome_file):
                os.remove(self.nome_file)
            return True
        except Exception as e:
            print(f"Errore durante l'eliminazione: {e}")
            return False

    def esiste_file_dati(self) -> bool:
        """Verifica se esiste il file dei dati"""
        return os.path.exists(self.nome_file)
