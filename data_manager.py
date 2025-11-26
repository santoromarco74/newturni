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
            'pianificazione': {},
            'ultimo_aggiornamento': None
        }

    def salva_dati(self, addetti: List[Addetto], turni: List[Turno], pianificazione: Dict = None) -> bool:
        """
        Salva addetti, turni e pianificazione nel file JSON

        Args:
            addetti: Lista di addetti
            turni: Lista di turni
            pianificazione: Dizionario della pianificazione dei turni (opzionale)

        Returns:
            True se il salvataggio è riuscito, False altrimenti
        """
        try:
            dati = {
                'addetti': self._serializza_addetti(addetti),
                'turni': self._serializza_turni(turni),
                'pianificazione': self._serializza_pianificazione(pianificazione) if pianificazione else {},
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
        Carica addetti, turni e pianificazione dal file JSON

        Returns:
            Tupla (addetti, turni, pianificazione) se il caricamento è riuscito, ([], [], {}) altrimenti
        """
        if not os.path.exists(self.nome_file):
            return [], [], {}

        try:
            with open(self.nome_file, 'r', encoding='utf-8') as f:
                dati = json.load(f)

            addetti = self._deserializza_addetti(dati.get('addetti', []))
            turni = self._deserializza_turni(dati.get('turni', []))
            pianificazione = self._deserializza_pianificazione(dati.get('pianificazione', {}), turni)

            return addetti, turni, pianificazione
        except Exception as e:
            print(f"Errore durante il caricamento: {e}")
            return [], [], {}

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

    def _serializza_pianificazione(self, pianificazione: Dict) -> Dict:
        """
        Serializza la pianificazione in formato JSON

        La pianificazione è un dict: {data_datetime: {nome_addetto: Turno}}
        Converte le date datetime in stringhe ISO e i turni nel loro nome
        """
        if not pianificazione:
            return {}

        risultato = {}
        for data, assegnazioni in pianificazione.items():
            # Converte la data datetime in stringa ISO
            data_str = data.isoformat()
            assegnazioni_serializzate = {}

            for nome_addetto, turno in assegnazioni.items():
                # Salva solo il nome del turno (il turno completo può essere recuperato dalla lista turni)
                assegnazioni_serializzate[nome_addetto] = turno.nome if hasattr(turno, 'nome') else str(turno)

            risultato[data_str] = assegnazioni_serializzate

        return risultato

    def _deserializza_pianificazione(self, dati: Dict, turni: List[Turno]) -> Dict:
        """
        Deserializza la pianificazione dal formato JSON

        Converte le stringhe ISO in date datetime e i nomi dei turni negli oggetti Turno
        """
        if not dati:
            return {}

        # Crea un dizionario per cercare rapidamente turni per nome
        turni_per_nome = {turno.nome: turno for turno in turni}

        risultato = {}
        for data_str, assegnazioni in dati.items():
            try:
                # Converte la stringa ISO in datetime
                data = datetime.fromisoformat(data_str)
                assegnazioni_deserializzate = {}

                for nome_addetto, nome_turno in assegnazioni.items():
                    # Recupera l'oggetto Turno dal dizionario
                    if nome_turno in turni_per_nome:
                        assegnazioni_deserializzate[nome_addetto] = turni_per_nome[nome_turno]

                if assegnazioni_deserializzate:
                    risultato[data] = assegnazioni_deserializzate
            except (ValueError, KeyError) as e:
                print(f"Errore nel caricamento pianificazione per {data_str}: {e}")
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
