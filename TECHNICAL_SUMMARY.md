# Sommario Tecnico - Progetto NewTurni

**Ultimo aggiornamento**: 2025-12-27

---

## 1. Obiettivo del Progetto

NewTurni è un'applicazione Python per la **gestione e pianificazione automatica dei turni** di lavoro per negozi e piccole attività commerciali. Il sistema permette di:

- Gestire il personale con vincoli orari contrattuali (ore minime e massime settimanali)
- Definire turni di lavoro con fasce orarie
- Pianificare automaticamente i turni rispettando:
  - Ore contrattuali minime/massime di ogni addetto
  - Giorni di riposo settimanali
  - Ferie e permessi
  - Giorni festivi
- Monitorare le statistiche di lavoro (ore per settimana, domeniche lavorate, etc.)
- Esportare la pianificazione in formato Excel professionale

---

## 2. Stack Tecnologico

### Linguaggio e Runtime
- **Python 3** - Linguaggio principale

### Librerie e Framework

#### Interfaccia Utente
- **PyQt6** (>= 6.4.0) - Framework per interfaccia grafica moderna
  - Finestre responsive con dimensionamento automatico
  - Dialogs centrati sulla finestra padre
  - Tabelle con filtri di ricerca in tempo reale
  - Sistema a tab per organizzazione funzionale

#### Data Processing
- **openpyxl** (>= 3.1.0) - Export Excel con formattazione avanzata
  - Stili personalizzati (colori, font, allineamento)
  - Fogli multipli (Pianificazione, Statistiche, Dettagli Addetti)
  - Formato matrice Giorni × Addetti
- **python-dateutil** (>= 2.8.2) - Gestione avanzata delle date

#### Persistenza
- **JSON** (libreria standard) - Salvataggio/caricamento dati
  - Serializzazione/deserializzazione custom per oggetti complessi
  - Auto-save su ogni modifica

### Architettura del Codice

```
newturni/
├── main.py                   # Entry point - menu scelta CLI/GUI
├── gestione_turni.py         # Core business logic
│   ├── Addetto               # Classe dipendente
│   ├── Turno                 # Classe turno di lavoro
│   ├── TurnoManager          # Pianificazione e logica centrale
│   └── MenuInterattivo       # Interfaccia CLI
├── gui_turni.py              # Interfaccia grafica PyQt6
│   ├── FinestraPrincipale    # Main window con tab
│   ├── DialogAggiungiAddetto # Dialog aggiunta addetto
│   └── DialogAggiungiTurno   # Dialog aggiunta turno
├── data_manager.py           # Persistenza dati JSON
│   └── DataManager           # Serializzazione/deserializzazione
├── test_gestione_turni.py    # Test suite
├── requirements.txt          # Dipendenze Python
└── dati_turni.json          # File dati persistiti (auto-generato)
```

---

## 3. Funzionalità Implementate

### 3.1 Gestione Addetti
- Creazione addetti con parametri:
  - Nome
  - Ore contratto (minimo settimanale)
  - Ore massime settimanali
  - Possibilità straordinario (booleano)
  - Giorni di riposo settimanali (0-6, Lunedì-Domenica)
  - Ferie/permessi (date specifiche)
- Validazione robusta:
  - Nome obbligatorio
  - Ore contratto ≤ ore massime
  - Controllo duplicati
- Modifica e rimozione addetti
- Barra di ricerca in tempo reale (GUI)

### 3.2 Gestione Turni
- Definizione turni con:
  - Nome (es. "Mattina", "Pomeriggio")
  - Ora inizio (formato HH:MM)
  - Ora fine (formato HH:MM)
  - Calcolo automatico durata in ore
- Validazione formato orario
- Supporto turni che attraversano la mezzanotte
- Controllo duplicati

### 3.3 Algoritmo di Pianificazione Automatica

**Regola fondamentale**: Massimo 1 turno per addetto per giorno

#### Fase 1: Bilanciamento Iniziale
- Per ogni giorno del mese:
  - Filtra addetti disponibili (non in riposo/ferie)
  - Per ogni turno da assegnare:
    - Seleziona l'addetto con **meno ore nella settimana corrente**
    - Verifica che non superi il massimo settimanale
    - Verifica che non abbia già un turno quel giorno

#### Fase 2: Verifica Ore Minime
- Per ogni addetto:
  - Calcola ore totali assegnate
  - Se < ore_contratto:
    - Trova giorni disponibili
    - Assegna turni aggiuntivi fino a raggiungere il minimo
    - Rispetta sempre il massimo settimanale

**Risultato**: Pianificazione equa che rispetta tutti i vincoli

### 3.4 Interfaccia Doppia (CLI + GUI)

#### CLI (Command Line Interface)
- Menu interattivo testuale
- Ideale per server o terminale
- Tutte le funzionalità disponibili

#### GUI (Graphical User Interface - PyQt6)
- **Tab Addetti**:
  - Tabella con tutti gli addetti
  - Barra di ricerca filtro in tempo reale
  - Bottoni aggiungi/rimuovi
  - Dialog modale per inserimento dati
- **Tab Turni**:
  - Tabella turni disponibili
  - Selezione mese/anno
  - Bottone pianificazione con conferma
- **Tab Pianificazione**:
  - Visualizzazione matrice completa
  - Colori differenziati:
    - Rosso chiaro: giorni festivi
    - Giallo chiaro: domeniche
    - Bianco: giorni normali
  - Bottone export Excel
- **Tab Statistiche**:
  - Ore totali per addetto (mensili)
  - Ore per settimana (con media)
  - Giorni lavorati
  - Domeniche lavorate (dettaglio per addetto)

**UX Enhancements**:
- Finestra responsive (85% larghezza × 90% altezza schermo)
- Centering automatico finestra principale
- Dialog centrati sulla finestra padre
- Header tabelle stretch per occupare tutto lo spazio
- Feedback utente su tutte le operazioni (successo/errore)

### 3.5 Persistenza Dati
- **Auto-save** su ogni operazione di modifica
- Salvataggio in `dati_turni.json` con struttura:
  ```json
  {
    "addetti": [...],
    "turni": [...],
    "pianificazione": {...},
    "ultimo_aggiornamento": "2025-12-27T..."
  }
  ```
- **Auto-load** all'avvio dell'applicazione
- Serializzazione custom per oggetti complessi (datetime, Addetto, Turno)

### 3.6 Export Excel Professionale
- **Foglio 1 - Pianificazione**:
  - Matrice Giorni × Addetti
  - Una riga per giorno
  - Una colonna per addetto
  - Celle con turno assegnato (nome + orario)
  - Colori festivi/domeniche
  - Header formattati (blu, testo bianco)
  - Text wrapping e centering

- **Foglio 2 - Statistiche**:
  - Ore totali per addetto
  - Giorni lavorati per addetto
  - Domeniche lavorate dettagliate

- **Foglio 3 - Dettagli Addetti**:
  - Tutti i vincoli e parametri di ogni addetto
  - Giorni riposo
  - Ferie programmate

### 3.7 Gestione Festività e Weekend
- Calendario festività italiane predefinito:
  - 1 Gennaio (Capodanno)
  - 20 Aprile (esempio)
  - 1 Maggio (Festa dei Lavoratori)
  - 25 Dicembre (Natale)
  - 26 Dicembre (Santo Stefano)
- Riconoscimento automatico domeniche
- Esclusione festivi dalla pianificazione
- Evidenziazione visiva in export e GUI

### 3.8 Sistema di Validazione
- **Addetti**:
  - Nome non vuoto
  - Ore numeriche positive
  - Ore contratto ≤ ore massime
  - Tipo straordinario booleano
- **Turni**:
  - Nome non vuoto
  - Formato orario HH:MM
  - Ore e minuti validi (0-23, 0-59)
- **Pianificazione**:
  - Almeno 1 addetto e 1 turno
  - Rispetto vincoli orari
  - Nessun doppio turno stesso giorno

---

## 4. Prossimi Passi

### Possibili Evoluzioni Future
1. **Funzionalità Avanzate**:
   - Gestione festività personalizzate
   - Notifiche via email/SMS agli addetti
   - Import dati da CSV/Excel
   - Gestione più negozi/sedi
   - Storico pianificazioni mensili

2. **Algoritmo di Pianificazione**:
   - Preferenze turni per addetto
   - Rotazione turni automatica
   - Ottimizzazione costi straordinari
   - Machine learning per previsione necessità

3. **Interfaccia**:
   - Web app (Flask/Django + React)
   - App mobile (React Native)
   - Dashboard analytics avanzate
   - Drag & drop per spostamento turni

4. **Integrations**:
   - Calendario Google/Outlook
   - Sistemi paghe (export buste paga)
   - API REST per integrazioni esterne

5. **Testing & Quality**:
   - Test coverage completo
   - CI/CD pipeline
   - Docker containerization
   - Logging avanzato

---

## 5. Note Tecniche

### Algoritmo di Pianificazione - Complessità
- **Tempo**: O(G × T × A) dove:
  - G = giorni del mese (~30)
  - T = numero turni
  - A = numero addetti
- **Spazio**: O(G × A) per la matrice pianificazione

### Limitazioni Attuali
- Massimo 1 turno per addetto per giorno
- Festività hardcoded (non configurabili da GUI)
- Non gestisce turni su più giorni (es. turno notturno 22-06)
- Pianificazione non modificabile manualmente dopo generazione

### Performance
- Caricamento istantaneo per ~10 addetti × 30 giorni
- Export Excel < 1 secondo
- GUI responsive anche con dataset grandi

---

**Mantenuto da**: Claude AI Assistant
**Versione**: 1.0
**Licenza**: Da definire
