#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GESTIONE TURNI NEGOZIO - INTERFACCIA GRAFICA PyQt6
Interfaccia grafica per il programma di pianificazione turni
"""

import sys
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QTableWidget, QTableWidgetItem, QDialog,
    QLineEdit, QSpinBox, QCheckBox, QComboBox, QDateEdit,
    QMessageBox, QTabWidget, QDateTimeEdit, QListWidget, QListWidgetItem,
    QHeaderView, QCalendarWidget
)
from PyQt6.QtCore import Qt, QDate, QDateTime
from PyQt6.QtGui import QColor, QFont

from gestione_turni import Addetto, Turno, TurnoManager


def center_dialog_on_parent(dialog, parent):
    """Centra un dialog sulla finestra padre"""
    if parent:
        parent_geometry = parent.frameGeometry()
        center_point = parent_geometry.center()
        dialog.move(center_point.x() - dialog.width() // 2,
                   center_point.y() - dialog.height() // 2)


class DialogAggiungiAddetto(QDialog):
    """Dialog per aggiungere un nuovo addetto"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Aggiungi Addetto")
        self.setGeometry(0, 0, 400, 450)  # Geometria per il calcolo della grandezza
        self.addetto = None

        layout = QVBoxLayout()

        # Nome
        layout.addWidget(QLabel("Nome Addetto:"))
        self.nome_input = QLineEdit()
        layout.addWidget(self.nome_input)

        # Ore contratto
        layout.addWidget(QLabel("Ore Settimanali Contratto (MIN):"))
        self.ore_contratto_spin = QSpinBox()
        self.ore_contratto_spin.setMinimum(0)
        self.ore_contratto_spin.setMaximum(60)
        self.ore_contratto_spin.setValue(40)
        layout.addWidget(self.ore_contratto_spin)

        # Ore max settimanale
        layout.addWidget(QLabel("Ore Massime Settimanali (MAX):"))
        self.ore_max_spin = QSpinBox()
        self.ore_max_spin.setMinimum(0)
        self.ore_max_spin.setMaximum(60)
        self.ore_max_spin.setValue(45)
        layout.addWidget(self.ore_max_spin)

        # Straordinario
        self.straordinario_check = QCheckBox("Può fare straordinario?")
        layout.addWidget(self.straordinario_check)

        # Giorni riposo
        layout.addWidget(QLabel("Giorni di Riposo Settimanale:"))
        self.giorni_riposo = []
        giorni_nomi = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì", "Sabato", "Domenica"]
        for i, giorno in enumerate(giorni_nomi):
            check = QCheckBox(giorno)
            check.setProperty("giorno_idx", i)
            self.giorni_riposo.append(check)
            layout.addWidget(check)

        # Bottoni
        btn_layout = QHBoxLayout()
        btn_ok = QPushButton("OK")
        btn_cancel = QPushButton("Annulla")
        btn_ok.clicked.connect(self.accetta)
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

        self.setLayout(layout)

        # Centra il dialog sulla finestra padre
        center_dialog_on_parent(self, self.parent())

    def accetta(self):
        """Convalida e accetta il dialog"""
        nome = self.nome_input.text().strip()

        if not nome:
            QMessageBox.warning(self, "Errore", "Inserisci il nome dell'addetto")
            return

        try:
            ore_contratto = self.ore_contratto_spin.value()
            ore_max = self.ore_max_spin.value()

            if ore_contratto <= 0:
                QMessageBox.warning(self, "Errore", "Ore contratto deve essere maggiore di 0")
                return

            if ore_max <= 0:
                QMessageBox.warning(self, "Errore", "Ore massime deve essere maggiore di 0")
                return

            if ore_contratto > ore_max:
                QMessageBox.warning(self, "Errore", "Ore contratto non possono superare ore massime")
                return

            self.addetto = Addetto(nome, ore_contratto, ore_max, self.straordinario_check.isChecked())

            # Aggiungi giorni riposo
            for check in self.giorni_riposo:
                if check.isChecked():
                    self.addetto.aggiungi_giorno_riposo(check.property("giorno_idx"))

            self.accept()
        except ValueError as e:
            QMessageBox.warning(self, "Errore di Validazione", str(e))


class DialogAggiungiTurno(QDialog):
    """Dialog per aggiungere un nuovo turno"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Aggiungi Turno")
        self.setGeometry(0, 0, 300, 250)
        self.turno = None

        layout = QVBoxLayout()

        # Nome turno
        layout.addWidget(QLabel("Nome Turno:"))
        self.nome_input = QLineEdit()
        layout.addWidget(self.nome_input)

        # Ora inizio
        layout.addWidget(QLabel("Ora Inizio (HH:MM):"))
        self.ora_inizio_input = QLineEdit("08:00")
        layout.addWidget(self.ora_inizio_input)

        # Ora fine
        layout.addWidget(QLabel("Ora Fine (HH:MM):"))
        self.ora_fine_input = QLineEdit("14:00")
        layout.addWidget(self.ora_fine_input)

        # Bottoni
        btn_layout = QHBoxLayout()
        btn_ok = QPushButton("OK")
        btn_cancel = QPushButton("Annulla")
        btn_ok.clicked.connect(self.accetta)
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

        self.setLayout(layout)

        # Centra il dialog sulla finestra padre
        center_dialog_on_parent(self, self.parent())

    def accetta(self):
        """Convalida e accetta il dialog"""
        nome = self.nome_input.text().strip()
        ora_inizio = self.ora_inizio_input.text().strip()
        ora_fine = self.ora_fine_input.text().strip()

        if not nome or not ora_inizio or not ora_fine:
            QMessageBox.warning(self, "Errore", "Completa tutti i campi")
            return

        try:
            self.turno = Turno(nome, ora_inizio, ora_fine)
            self.accept()
        except ValueError as e:
            QMessageBox.warning(self, "Errore di Validazione", str(e))


class FinestraPrincipale(QMainWindow):
    """Finestra principale dell'applicazione"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gestione Turni Negozio")

        # Calcola dimensioni responsive in base allo schermo
        try:
            screen = QApplication.primaryScreen()
            screen_geometry = screen.availableGeometry()
            width = int(screen_geometry.width() * 0.85)  # 85% della larghezza schermo
            height = int(screen_geometry.height() * 0.90)  # 90% dell'altezza schermo
            x = (screen_geometry.width() - width) // 2   # Centra orizzontalmente
            y = (screen_geometry.height() - height) // 2 # Centra verticalmente
            self.setGeometry(x, y, width, height)
        except:
            # Fallback a dimensioni fisse se c'è errore
            self.setGeometry(50, 50, 1200, 800)

        self.manager = TurnoManager()

        # Carica i dati salvati
        if self.manager.carica_dati():
            print("✓ Dati caricati dal salvataggio precedente")
        else:
            print("⚠ Nessun salvataggio precedente trovato")

        # Crea il widget centrale con tab
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # Crea i tab
        self.tab_addetti = self.crea_tab_addetti()
        self.tab_turni = self.crea_tab_turni()
        self.tab_pianificazione = self.crea_tab_pianificazione()
        self.tab_statistiche = self.crea_tab_statistiche()

        self.tabs.addTab(self.tab_addetti, "Addetti")
        self.tabs.addTab(self.tab_turni, "Turni")
        self.tabs.addTab(self.tab_pianificazione, "Pianificazione")
        self.tabs.addTab(self.tab_statistiche, "Statistiche")

        # Popola le tabelle con i dati caricati
        self.aggiorna_tabella_addetti()
        self.aggiorna_tabella_turni()

    def crea_tab_addetti(self):
        """Crea il tab per la gestione addetti"""
        widget = QWidget()
        layout = QVBoxLayout()

        # Barra di ricerca
        ricerca_layout = QHBoxLayout()
        ricerca_layout.addWidget(QLabel("Cerca:"))
        self.ricerca_addetti_input = QLineEdit()
        self.ricerca_addetti_input.setPlaceholderText("Digita nome addetto...")
        self.ricerca_addetti_input.textChanged.connect(self.filtra_addetti)
        ricerca_layout.addWidget(self.ricerca_addetti_input)
        layout.addLayout(ricerca_layout)

        # Tabella addetti
        layout.addWidget(QLabel("Addetti Registrati:"))
        self.tabella_addetti = QTableWidget()
        self.tabella_addetti.setColumnCount(5)
        self.tabella_addetti.setHorizontalHeaderLabels(
            ["Nome", "Ore Contratto (min)", "Ore Max (sett)", "Straordinario", "Giorni Riposo"]
        )
        self.tabella_addetti.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.tabella_addetti)

        # Bottoni
        btn_layout = QHBoxLayout()
        btn_aggiungi = QPushButton("Aggiungi Addetto")
        btn_rimuovi = QPushButton("Rimuovi Addetto")
        btn_aggiungi.clicked.connect(self.aggiungi_addetto)
        btn_rimuovi.clicked.connect(self.rimuovi_addetto)
        btn_layout.addWidget(btn_aggiungi)
        btn_layout.addWidget(btn_rimuovi)
        layout.addLayout(btn_layout)

        widget.setLayout(layout)
        return widget

    def crea_tab_turni(self):
        """Crea il tab per la gestione turni"""
        widget = QWidget()
        layout = QVBoxLayout()

        # Tabella turni
        layout.addWidget(QLabel("Turni Disponibili:"))
        self.tabella_turni = QTableWidget()
        self.tabella_turni.setColumnCount(4)
        self.tabella_turni.setHorizontalHeaderLabels(
            ["Nome", "Ora Inizio", "Ora Fine", "Ore"]
        )
        self.tabella_turni.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.tabella_turni)

        # Bottoni
        btn_layout = QHBoxLayout()
        btn_aggiungi = QPushButton("Aggiungi Turno")
        btn_rimuovi = QPushButton("Rimuovi Turno")
        btn_aggiungi.clicked.connect(self.aggiungi_turno)
        btn_rimuovi.clicked.connect(self.rimuovi_turno)
        btn_layout.addWidget(btn_aggiungi)
        btn_layout.addWidget(btn_rimuovi)
        layout.addLayout(btn_layout)

        # Impostazioni mese/anno
        impostazioni_layout = QHBoxLayout()
        impostazioni_layout.addWidget(QLabel("Mese:"))
        self.mese_combo = QComboBox()
        mesi = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
                "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]
        self.mese_combo.addItems(mesi)
        self.mese_combo.setCurrentIndex(datetime.now().month - 1)
        self.mese_combo.currentIndexChanged.connect(self.aggiorna_mese)
        impostazioni_layout.addWidget(self.mese_combo)

        impostazioni_layout.addWidget(QLabel("Anno:"))
        self.anno_spin = QSpinBox()
        self.anno_spin.setMinimum(2020)
        self.anno_spin.setMaximum(2050)
        self.anno_spin.setValue(datetime.now().year)
        self.anno_spin.valueChanged.connect(self.aggiorna_anno)
        impostazioni_layout.addWidget(self.anno_spin)

        impostazioni_layout.addStretch()

        # Bottone pianifica
        btn_pianifica = QPushButton("Pianifica Turni")
        btn_pianifica.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        btn_pianifica.clicked.connect(self.pianifica_turni)
        impostazioni_layout.addWidget(btn_pianifica)

        layout.addLayout(impostazioni_layout)

        widget.setLayout(layout)
        return widget

    def crea_tab_pianificazione(self):
        """Crea il tab per visualizzare la pianificazione"""
        widget = QWidget()
        layout = QVBoxLayout()

        layout.addWidget(QLabel("Pianificazione Turni:"))
        self.tabella_pianificazione = QTableWidget()
        self.tabella_pianificazione.setColumnCount(4)
        self.tabella_pianificazione.setHorizontalHeaderLabels(
            ["Data", "Giorno", "Addetto", "Turno (Orario)"]
        )
        self.tabella_pianificazione.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.tabella_pianificazione)

        # Bottoni
        btn_layout = QHBoxLayout()
        btn_aggiorna = QPushButton("Aggiorna Visualizzazione")
        btn_export = QPushButton("Esporta Excel")
        btn_aggiorna.clicked.connect(self.aggiorna_pianificazione)
        btn_export.clicked.connect(self.esporta_excel)
        btn_layout.addWidget(btn_aggiorna)
        btn_layout.addWidget(btn_export)
        layout.addLayout(btn_layout)

        widget.setLayout(layout)
        return widget

    def crea_tab_statistiche(self):
        """Crea il tab per visualizzare le statistiche"""
        widget = QWidget()
        layout = QVBoxLayout()

        # Area statistiche
        self.statistiche_text = QLabel()
        self.statistiche_text.setFont(QFont("Courier", 10))
        self.statistiche_text.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        scroll_area = QWidget()
        scroll_layout = QVBoxLayout()
        scroll_layout.addWidget(self.statistiche_text)
        scroll_area.setLayout(scroll_layout)

        layout.addWidget(scroll_area)

        # Bottone aggiorna
        btn_aggiorna = QPushButton("Aggiorna Statistiche")
        btn_aggiorna.clicked.connect(self.aggiorna_statistiche)
        layout.addWidget(btn_aggiorna)

        widget.setLayout(layout)
        return widget

    def aggiungi_addetto(self):
        """Apre il dialog per aggiungere un addetto"""
        dialog = DialogAggiungiAddetto(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            if dialog.addetto:
                # Controlla duplicati
                if any(a.nome.lower() == dialog.addetto.nome.lower() for a in self.manager.addetti):
                    QMessageBox.warning(self, "Errore", "Addetto già esistente!")
                    return

                self.manager.aggiungi_addetto(dialog.addetto)
                # Salva automaticamente
                if self.manager.salva_dati():
                    QMessageBox.information(self, "Successo", f"Addetto '{dialog.addetto.nome}' aggiunto e salvato!")
                else:
                    QMessageBox.warning(self, "Avviso", "Addetto aggiunto ma errore nel salvataggio")
                self.aggiorna_tabella_addetti()

    def rimuovi_addetto(self):
        """Rimuove l'addetto selezionato"""
        riga_selezionata = self.tabella_addetti.currentRow()
        if riga_selezionata == -1:
            QMessageBox.warning(self, "Errore", "Seleziona un addetto")
            return

        nome = self.tabella_addetti.item(riga_selezionata, 0).text()
        risposta = QMessageBox.question(self, "Conferma", f"Rimuovere '{nome}'?")
        if risposta == QMessageBox.StandardButton.Yes:
            self.manager.rimuovi_addetto(nome)
            # Salva automaticamente
            if self.manager.salva_dati():
                QMessageBox.information(self, "Successo", f"Addetto '{nome}' rimosso e salvato!")
            else:
                QMessageBox.warning(self, "Avviso", "Addetto rimosso ma errore nel salvataggio")
            self.aggiorna_tabella_addetti()

    def filtra_addetti(self):
        """Filtra la tabella degli addetti in base al testo cercato"""
        testo_ricerca = self.ricerca_addetti_input.text().lower()

        for riga in range(self.tabella_addetti.rowCount()):
            nome_addetto = self.tabella_addetti.item(riga, 0).text()
            # Mostra la riga se il nome contiene il testo cercato, altrimenti la nasconde
            self.tabella_addetti.setRowHidden(riga, testo_ricerca not in nome_addetto.lower())

    def aggiorna_tabella_addetti(self):
        """Aggiorna la tabella degli addetti"""
        self.tabella_addetti.setRowCount(len(self.manager.addetti))

        for i, addetto in enumerate(self.manager.addetti):
            # Riapplica il filtro se è già stato impostato
            if hasattr(self, 'ricerca_addetti_input'):
                self.tabella_addetti.setRowHidden(i, False)  # Mostra tutte le righe inizialmente
            # Nome
            self.tabella_addetti.setItem(i, 0, QTableWidgetItem(addetto.nome))

            # Ore contratto
            self.tabella_addetti.setItem(i, 1, QTableWidgetItem(str(addetto.ore_contratto)))

            # Ore max
            self.tabella_addetti.setItem(i, 2, QTableWidgetItem(str(addetto.ore_max_settimanale)))

            # Straordinario
            self.tabella_addetti.setItem(i, 3, QTableWidgetItem("Sì" if addetto.straordinario else "No"))

            # Giorni riposo
            giorni_nomi = ["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"]
            giorni_riposo = [giorni_nomi[g] for g in sorted(addetto.giorni_riposo)]
            self.tabella_addetti.setItem(i, 4, QTableWidgetItem(", ".join(giorni_riposo) if giorni_riposo else "-"))

        # Riapplica il filtro di ricerca se presente
        if hasattr(self, 'ricerca_addetti_input'):
            self.filtra_addetti()

    def aggiungi_turno(self):
        """Apre il dialog per aggiungere un turno"""
        dialog = DialogAggiungiTurno(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            if dialog.turno:
                # Controlla duplicati
                if any(t.nome.lower() == dialog.turno.nome.lower() for t in self.manager.turni):
                    QMessageBox.warning(self, "Errore", "Turno già esistente!")
                    return

                self.manager.aggiungi_turno(dialog.turno)
                # Salva automaticamente
                if self.manager.salva_dati():
                    QMessageBox.information(self, "Successo", f"Turno '{dialog.turno.nome}' aggiunto e salvato!")
                else:
                    QMessageBox.warning(self, "Avviso", "Turno aggiunto ma errore nel salvataggio")
                self.aggiorna_tabella_turni()

    def rimuovi_turno(self):
        """Rimuove il turno selezionato"""
        riga_selezionata = self.tabella_turni.currentRow()
        if riga_selezionata == -1:
            QMessageBox.warning(self, "Errore", "Seleziona un turno")
            return

        nome = self.tabella_turni.item(riga_selezionata, 0).text()
        risposta = QMessageBox.question(self, "Conferma", f"Rimuovere '{nome}'?")
        if risposta == QMessageBox.StandardButton.Yes:
            self.manager.rimuovi_turno(nome)
            # Salva automaticamente
            if self.manager.salva_dati():
                QMessageBox.information(self, "Successo", f"Turno '{nome}' rimosso e salvato!")
            else:
                QMessageBox.warning(self, "Avviso", "Turno rimosso ma errore nel salvataggio")
            self.aggiorna_tabella_turni()

    def aggiorna_tabella_turni(self):
        """Aggiorna la tabella dei turni"""
        self.tabella_turni.setRowCount(len(self.manager.turni))

        for i, turno in enumerate(self.manager.turni):
            self.tabella_turni.setItem(i, 0, QTableWidgetItem(turno.nome))
            self.tabella_turni.setItem(i, 1, QTableWidgetItem(turno.ora_inizio))
            self.tabella_turni.setItem(i, 2, QTableWidgetItem(turno.ora_fine))
            self.tabella_turni.setItem(i, 3, QTableWidgetItem(f"{turno.ore}h"))

    def aggiorna_mese(self):
        """Aggiorna il mese nel manager"""
        self.manager.mese = self.mese_combo.currentIndex() + 1
        # Auto-aggiorna statistiche quando cambia il mese
        self.aggiorna_statistiche()

    def aggiorna_anno(self):
        """Aggiorna l'anno nel manager"""
        self.manager.anno = self.anno_spin.value()
        # Auto-aggiorna statistiche quando cambia l'anno
        self.aggiorna_statistiche()

    def pianifica_turni(self):
        """Avvia la pianificazione dei turni"""
        if not self.manager.addetti:
            QMessageBox.warning(self, "Errore", "Aggiungere almeno un addetto")
            return

        if not self.manager.turni:
            QMessageBox.warning(self, "Errore", "Aggiungere almeno un turno")
            return

        if self.manager.pianifica_turni():
            # Calcola statistiche per il messaggio di successo
            giorni_pianificati = len([d for d in self.manager.pianificazione if self.manager.pianificazione[d]])
            giorni_totali = len(self.manager.pianificazione)
            addetti_con_turni = len(set([nome for d in self.manager.pianificazione for nome in self.manager.pianificazione[d]]))

            messaggio = f"""Pianificazione completata con successo!

✓ Giorni con assegnazioni: {giorni_pianificati}/{giorni_totali}
✓ Addetti con turni assegnati: {addetti_con_turni}/{len(self.manager.addetti)}
✓ Mese: {self.manager._nome_mese(self.manager.mese)} {self.manager.anno}"""

            # Salva automaticamente
            if self.manager.salva_dati():
                messaggio += "\n✓ Dati salvati"

            QMessageBox.information(self, "Successo", messaggio)
            self.aggiorna_pianificazione()
            self.aggiorna_statistiche()
        else:
            QMessageBox.critical(self, "Errore", "Errore durante la pianificazione")

    def aggiorna_pianificazione(self):
        """Aggiorna la tabella della pianificazione"""
        if not self.manager.pianificazione:
            self.tabella_pianificazione.setRowCount(0)
            return

        giorni = self.manager.get_giorni_mese()
        righe = 0

        # Conta le righe necessarie
        for data in giorni:
            assegnazioni = self.manager.pianificazione.get(data, {})
            if assegnazioni:
                righe += len(assegnazioni)
            else:
                righe += 1

        self.tabella_pianificazione.setRowCount(righe)

        riga = 0
        for data in giorni:
            assegnazioni = self.manager.pianificazione.get(data, {})

            # Colore
            if self.manager.is_festivo(data):
                colore = QColor(255, 200, 200)
            elif self.manager.is_domenica(data):
                colore = QColor(255, 255, 200)
            else:
                colore = QColor(255, 255, 255)

            if assegnazioni:
                for nome_addetto, turno in assegnazioni.items():
                    data_str = data.strftime("%d/%m/%Y")
                    giorno_str = self.manager._nome_giorno_italiano(data.weekday())
                    turno_str = f"{turno.nome} ({turno.ora_inizio}-{turno.ora_fine})"

                    item_data = QTableWidgetItem(data_str)
                    item_data.setBackground(colore)
                    item_giorno = QTableWidgetItem(giorno_str)
                    item_giorno.setBackground(colore)
                    item_addetto = QTableWidgetItem(nome_addetto)
                    item_addetto.setBackground(colore)
                    item_turno = QTableWidgetItem(turno_str)
                    item_turno.setBackground(colore)

                    self.tabella_pianificazione.setItem(riga, 0, item_data)
                    self.tabella_pianificazione.setItem(riga, 1, item_giorno)
                    self.tabella_pianificazione.setItem(riga, 2, item_addetto)
                    self.tabella_pianificazione.setItem(riga, 3, item_turno)

                    riga += 1
            else:
                data_str = data.strftime("%d/%m/%Y")
                giorno_str = self.manager._nome_giorno_italiano(data.weekday())

                item_data = QTableWidgetItem(data_str)
                item_data.setBackground(colore)
                item_giorno = QTableWidgetItem(giorno_str)
                item_giorno.setBackground(colore)
                item_none = QTableWidgetItem("-")
                item_none.setBackground(colore)
                item_turno = QTableWidgetItem("Nessun turno")
                item_turno.setBackground(colore)

                self.tabella_pianificazione.setItem(riga, 0, item_data)
                self.tabella_pianificazione.setItem(riga, 1, item_giorno)
                self.tabella_pianificazione.setItem(riga, 2, item_none)
                self.tabella_pianificazione.setItem(riga, 3, item_turno)

                riga += 1

    def aggiorna_statistiche(self):
        """Aggiorna le statistiche"""
        if not self.manager.pianificazione:
            self.statistiche_text.setText("Nessuna pianificazione disponibile")
            return

        stats = self.manager.genera_statistiche()

        testo = "=== STATISTICHE PIANIFICAZIONE ===\n\n"

        testo += "--- ORE TOTALI PER ADDETTO (mese) ---\n"
        for nome, ore in sorted(stats['ore_totali_per_addetto'].items()):
            testo += f"{nome:20} {ore:5.0f}h totali nel mese\n"

        testo += "\n--- ORE PER SETTIMANA ---\n"
        for nome, ore_settimane in sorted(stats['ore_per_settimana'].items()):
            addetto = next(a for a in self.manager.addetti if a.nome == nome)
            if ore_settimane:
                dettagli = ", ".join([f"Sett {s}: {o:.0f}h" for s, o in sorted(ore_settimane.items())])
                testo += f"{nome:20} {dettagli}\n"
                media = sum(ore_settimane.values()) / len(ore_settimane)
                testo += f"{'':20} Media: {media:.1f}h (contratto: {addetto.ore_contratto}h min, max {addetto.ore_max_settimanale}h)\n"

        testo += "\n--- GIORNI LAVORATI ---\n"
        for nome, giorni in sorted(stats['giorni_lavorati_per_addetto'].items()):
            testo += f"{nome:20} {giorni:3} giorni\n"

        if stats['dettaglio_domeniche']:
            testo += "\n--- DOMENICHE LAVORATE ---\n"
            for nome, giorni in sorted(stats['dettaglio_domeniche'].items()):
                testo += f"{nome:20} {giorni:3} domeniche\n"

        self.statistiche_text.setText(testo)

    def esporta_excel(self):
        """Esporta la pianificazione in Excel"""
        if not self.manager.pianificazione:
            QMessageBox.warning(self, "Errore", "Nessuna pianificazione disponibile")
            return

        try:
            percorso = self.manager.esporta_excel()
            QMessageBox.information(self, "Successo", f"File esportato:\n{percorso}")
        except Exception as e:
            QMessageBox.critical(self, "Errore", f"Errore durante l'esportazione:\n{str(e)}")


def main():
    """Punto di entrata del programma GUI"""
    app = QApplication(sys.argv)
    finestra = FinestraPrincipale()
    finestra.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
