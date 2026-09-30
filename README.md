# 🏃 Garmin Training Planner

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)]()

**Sistema completo per pianificare, creare, caricare e scaricare allenamenti strutturati da Garmin Connect**

Un'applicazione desktop Python che consente di creare workout strutturati per running, cycling e swimming utilizzando un DSL (Domain Specific Language) personalizzato in Excel, caricarli automaticamente su Garmin Connect, e scaricare dati storici per analisi e training AI.

---

## ✨ Caratteristiche Principali

### 📤 Upload & Scheduling
- ✅ **Creazione workout strutturati** tramite DSL personalizzato in Excel
- ✅ **Upload automatico** su Garmin Connect
- ✅ **Pianificazione calendario** con date specifiche
- ✅ **Gestione multi-sport** (running, cycling, swimming)
- ✅ **Zone personalizzate** (HR, Pace, Power, Swimming Pace)
- ✅ **Ripetizioni annidate** e strutture complesse

### 📥 Download & Analisi (NUOVO v2.0!)
- ✅ **Download workout pianificati** dalla Workout Library Garmin
- ✅ **Download attività completate** con dati reali
- ✅ **Filtro sport intelligente** (running, cycling, swimming)
- ✅ **Selezione periodo visuale** con calendario DateEntry
- ✅ **Export Excel multi-sheet** con formato AI-ready
- ✅ **Generazione automatica DSL** dalle attività completate
- ✅ **Nome file intelligente** auto-generato basato su filtri
- ✅ **Statistiche dettagliate** per sport e periodo

### 🎨 Un'unica interfaccia (v3)
- Tema **chiaro/scuro** (segue il Mac, pulsante ☾/☀ per cambiarlo)
- Lista degli allenamenti **divisa per settimana**, con durata e km stimati e stato (da caricare, su Garmin, pianificato, errore)
- **Editor dell'allenamento** con profilo grafico dell'intensità, lista degli step, modelli rapidi e testo DSL con colori e controllo mentre scrivi
- Scheda **Parametri** (zone e ritmi) e **Panoramica** con i km per settimana
- Supporto running, ciclismo e nuoto

### 🔐 Sicurezza
- ✅ Autenticazione con supporto MFA
- ✅ Nessuna password memorizzata
- ✅ Sessione salvata **fuori dal progetto** in `~/.garminplanner/tokens` (permessi 600), quindi non finisce mai su git
- ⚠️ I file token danno accesso al tuo account Garmin: non condividerli e non caricarli online

---

## 📦 Requisiti

### Python
- Python 3.9 o superiore
- tkinter (incluso nella distribuzione standard)

### Dipendenze
```bash
pip install -r requirements.txt
```
(pandas, openpyxl, tkcalendar, garth). Nota: `garth`, la libreria non ufficiale per Garmin Connect, è stata dichiarata *deprecated* dal suo autore; la versione è fissata in `requirements.txt`.

---

## 🚀 Installazione

### 1. Clona il repository
```bash
git clone https://github.com/yesnoj/garminPlannerCGPT.git
cd garminPlannerCGPT
```

### 2. Installa le dipendenze
```bash
python -m pip install -r requirements.txt
```
Usa lo **stesso `python`** con cui avvii l'app (su Windows spesso `py -m pip install -r requirements.txt`).
Se il tema `sv-ttk` manca l'app funziona lo stesso, con un tema chiaro/scuro di riserva, e lo segnala nella barra di stato.

### 3. Avvia l'applicazione
```bash
python garmin_planner.py      # oppure: python launcher.py
```
Le vecchie interfacce (Classic, Advanced, Multisport) restano per ora in `legacy/`.

---

## 📖 Guida Completa

### 🎯 PARTE 1: Upload Workout su Garmin

#### 1.1 Primo Avvio

```bash
python garmin_planner.py
```

- **Apri…** un piano Excel esistente, oppure **Nuovo piano** per crearne uno da un modello.
- A sinistra trovi gli allenamenti divisi per settimana; cerca per testo o filtra (da caricare, pianificati, con errori, da oggi in poi).
- Selezionando un allenamento si apre l'**editor** a destra:
  - titolo, data, sport, settimana e sessione;
  - **profilo** dell'allenamento: la larghezza è la durata, l'altezza l'intensità, il colore il tipo di step (passa col mouse per i dettagli);
  - scheda **Step**: aggiungi step o ripetizioni, **Modelli rapidi** (ripetute, soglia, allunghi…), modifica con doppio clic, **trascina col mouse** per riordinare (rilascia al centro di una ripetizione per metterci dentro lo step), oppure usa ↑ ↓ e ⇥ ⇤;
  - scheda **Testo DSL**: il testo con i colori, gli errori evidenziati in rosso mentre scrivi;
  - **Salva allenamento** aggiorna anche il file Excel (gli altri fogli del file restano intatti).
- Nella lista: **trascina** un allenamento per riordinarlo, anche in un'altra settimana (la data si sposta dello stesso numero di settimane); **tasto destro** per Duplica, Duplica nella settimana successiva, Sposta su/giù, Ordina per data e le operazioni Garmin. Selezionando una settimana intera, Duplica copia tutta la settimana.
- Scorciatoie: ⌘O apri, ⌘S salva, ⌘N nuovo allenamento, ⌘D duplica, ⌘↑/⌘↓ sposta (Ctrl su Windows/Linux).

#### 1.2 Login a Garmin Connect

1. Clicca su **Accedi…** in alto a destra
2. Inserisci email e password Garmin Connect (o **Usa sessione salvata**)
3. Se richiesto, inserisci il codice MFA
4. La sessione viene salvata in `~/.garminplanner/tokens` per i prossimi utilizzi

#### 1.3 Creare Workout con Excel

##### Genera il Template
1. Clicca su **"📄 Genera Excel"** (oppure parti da `examples/esempio_running.xlsx`)
2. Salva il file template nella directory desiderata

##### Struttura Excel

**Sheet "Workouts"** (una riga per allenamento):

| Week | Date | Session | Sport | Description | Steps |
|------|------|---------|-------|-------------|-------|
| 1 | 2026-10-05 | 1 | Running | Facile 6 km | `interval: 6km @ easy_range` |
| 1 | 2026-10-07 | 2 | Running | 5×1000 m @4:50-5:00 rec 2' | *(vedi esempio multiriga sotto)* |
| 1 | 2026-10-10 | 3 | Running | Lungo 12 km | `interval: 12km @ long_range` |

Le colonne `WorkoutId`, `WorkoutScheduleId` e `ScheduledDate` vengono compilate dal programma dopo il caricamento su Garmin.

**Sheet "Parameters"** (zone e ritmi riutilizzabili):

| Key | Metric | Expression | Notes |
|-----|--------|------------|-------|
| Z1 | pace | 7:00-7:30 | Recupero |
| Z2 | pace | 6:35-7:00 | Facile |
| easy_range | pace | 6:35-7:00 | Chiave personalizzata |
| HR_Z2 | hr | 70-80% | % di HR_max |
| HR_max | hr | 179 | Frequenza massima |
| pace_tolerance | pace | 5 | ± secondi per i ritmi singoli |

Le chiavi si cercano **senza distinzione tra maiuscole e minuscole** (`Easy_Range` = `easy_range`).

#### 1.4 DSL - Linguaggio dei Workout

Una riga per step, nel formato:

```
tipo: durata @ target
```

**Tipi di step:** `warmup`, `interval`, `recovery` (recupero attivo), `rest` (riposo da fermo), `cooldown`.

**Durata o distanza:**

| Scrivi | Significa |
|---|---|
| `10min`, `10'`, `1.5min`, `10 minuti` | minuti |
| `30sec`, `30s`, `90s` | secondi |
| `1h`, `1.5h` | ore |
| `1:30`, `1:05:00` | mm:ss / h:mm:ss |
| `400m`, `1.5km`, `21.1km` | distanza (**`10m` = 10 metri**, non minuti!) |
| `lap-button` | finché non premi Lap |

**Target (dopo `@`, facoltativo):**

| Scrivi | Significa |
|---|---|
| `5:00` | ritmo ± `pace_tolerance` |
| `4:50-5:00` | intervallo di ritmo |
| `Z1` … `Z5` | zona ritmo dal foglio Parameters |
| `HR_Z1` … `HR_Z5`, `140-160` | frequenza cardiaca |
| `easy_range`, `hmp`, … | qualunque chiave del foglio Parameters |
| `250W`, `200-250W`, `Power_Z2` | potenza (bici) |
| `90rpm`, `Cadence_Easy` | cadenza (bici) |
| `1:45/100m`, `Swim_Z2` | ritmo nuoto |
| `open` o niente | nessun target |

**Ripetizioni:** `repeat N:` e sotto, **rientrati**, gli step da ripetere (2 o 4 spazi o tab, basta che siano coerenti). Si possono annidare.

```
warmup: 15min @ Z2
repeat 5:
  interval: 1000m @ 4:50-5:00
  recovery: 2min @ Z1
cooldown: 10min @ Z1
```

```
warmup: lap-button @ Z2
repeat 2:
  repeat 6:
    interval: 1min @ 4:30
    recovery: 1min @ Z1
  rest: 3min
```

Righe vuote e righe che iniziano con `#` vengono ignorate; tutto ciò che segue `--` su una riga è un commento.

##### Controllo errori

Prima di ogni caricamento il programma **controlla tutti i workout selezionati**. Se qualcosa non è valido (durata non riconosciuta, zona inesistente, chiave mancante in Parameters, `repeat` senza step, data mancante…) **non viene caricato niente** e compare l'elenco dei problemi con riga Excel e riga del DSL, ad esempio:

```
Riga Excel 5 W2S1 «Facile 6 km»: riga 1: zona 'Z7' non valida (usa Z1-Z5)  →  «interval: 6km @ Z7»
```

Il foglio **"Esempi DSL"** del template contiene altri esempi pronti.

#### 1.5 Upload su Garmin

1. Seleziona uno o più allenamenti (o un'intera settimana) nella lista
2. Scegli l'operazione nella barra "Garmin":
   - **Carica e pianifica**: crea il workout (se serve) e lo mette nel calendario alla sua data
   - **Solo carica**: solo nella libreria Garmin
   - **Togli dal calendario**: rimuove la pianificazione, il workout resta nella libreria
   - **Elimina da Garmin**: cancella definitivamente il workout (resta nel tuo Excel)
3. Prima di inviare, tutti gli allenamenti selezionati vengono controllati: se c'è un errore non parte nulla
4. Gli ID Garmin vengono salvati nell'Excel dopo ogni operazione, anche se si interrompe a metà: un nuovo tentativo non crea doppioni

---

### 📥 PARTE 2: Download Dati da Garmin (NUOVO!)

#### 2.1 Aprire il Download Dialog

Clicca sul pulsante **"📥 Download"** nella GUI principale.

Si aprirà la finestra "Download Workout e Attività".

#### 2.2 Configurazione Download

##### 📅 Seleziona Periodo

**Opzione A - Quick Select:**
- **Ultimo mese**: Scarica dati dell'ultimo mese
- **Ultimi 3 mesi**: Scarica dati degli ultimi 3 mesi
- **Ultimo anno**: Scarica dati dell'ultimo anno

**Opzione B - Range Personalizzato:**
- Usa i calendari **Dal:** e **Al:** per selezionare date specifiche
- Esempio: Dal 2025-01-01 Al 2025-12-31

##### 📊 Seleziona Cosa Scaricare

- **🗓️ Solo workout pianificati**: Scarica solo i workout dalla tua libreria Garmin
- **✅ Solo attività completate**: Scarica solo le corse/uscite effettivamente fatte
- **📥 Entrambi (consigliato)**: Scarica sia workout pianificati che attività completate

##### 🏃 Filtro Sport

- **Tutti gli sport**: Nessun filtro, scarica tutto
- **Corsa**: Solo attività/workout di running
- **Ciclismo**: Solo attività/workout di cycling
- **Nuoto**: Solo attività/workout di swimming

**Nota:** Il filtro viene applicato DOPO il download, quindi la velocità è la stessa indipendentemente dal filtro.

##### 📄 Formato Export

Scegli il formato di output più adatto alle tue esigenze:

**📊 Excel (.xlsx) - Consigliato per AI:**
- File Excel multi-sheet (Workouts, Activities, Summary, Parameters)
- Colonna "Steps" con **DSL generato automaticamente** dalle attività
- Pronto per essere caricato a Claude/ChatGPT per analisi
- Statistiche aggregate e zone pre-configurate
- Ideale per training AI e generazione piani personalizzati

**📋 JSON (.json) - Dati grezzi completi:**
- Formato JSON con tutti i dati originali dall'API Garmin
- Nessuna elaborazione o trasformazione
- Include tutti i campi e metadati
- Ideale per sviluppatori e analisi programmatiche
- Mantiene la struttura originale dell'API

**Quale scegliere?**
- 🤖 **Vuoi usare l'AI per generare piani?** → Scegli Excel
- 📊 **Vuoi analizzare i dati con Claude?** → Scegli Excel
- 💻 **Sei uno sviluppatore?** → Scegli JSON
- 🔧 **Hai bisogno dei dati grezzi?** → Scegli JSON

##### 💾 Nome File

Il sistema genera automaticamente un nome file intelligente basato sui tuoi filtri:

**Esempi Excel:**
```
garmin_workouts_2025_running.xlsx        # Solo workout, anno 2025, running
garmin_activities_2026-01_swimming.xlsx  # Solo attività, gennaio 2026, nuoto
garmin_data_2024-2025_cycling.xlsx       # Entrambi, 2 anni, ciclismo
garmin_workouts_2025-2026_running.xlsx   # Solo workout, ultimo anno, running
```

**Esempi JSON:**
```
garmin_data_2025_running.json            # Tutti i dati, anno 2025, running
garmin_activities_2026-01.json           # Solo attività, gennaio 2026, tutti sport
```

**Puoi modificare il nome manualmente** prima di fare il download!

#### 2.3 Avvia Download

1. Clicca sul pulsante **"📥 Download"**
2. Attendi il completamento (può richiedere 10-60 secondi)
3. Vedrai un messaggio di conferma con le statistiche

**Esempio messaggio:**
```
✅ Download completato!

Workout pianificati: 170
Attività completate: 138
Totale: 308

Salvato in: garmin_data_2025_running.xlsx
```

---

### 📊 PARTE 3: Struttura File Scaricato

A seconda del formato scelto, il file avrà strutture diverse:

---

#### 📊 Formato EXCEL (.xlsx)

Il file Excel generato contiene **4 sheet** ottimizzati per l'analisi AI:

#### Sheet 1: "Workouts" (Workout Pianificati)

| Week | Date | Session | Sport | Description | Distance (km) | Duration (min) | Notes | WorkoutId |
|------|------|---------|-------|-------------|---------------|----------------|-------|-----------|
| 5 | 2026-01-26 | 1 | running | W9S27 - Activation | 2.5 | 30 | Workout vigilia gara | 1454947666 |
| 48 | 2025-12-15 | 2 | running | Interval Training | 8.0 | 45 | 8x400m Z5 | 1454947601 |

**Colonne:**
- **Week**: Settimana dell'anno (calcolata automaticamente)
- **Date**: Data di creazione/pianificazione del workout
- **Session**: Numero progressivo sessione
- **Sport**: running, cycling, swimming
- **Description**: Nome descrittivo del workout
- **Distance (km)**: Distanza stimata (se disponibile)
- **Duration (min)**: Durata stimata (se disponibile)
- **Notes**: Note o descrizione estesa
- **WorkoutId**: ID univoco Garmin

**Nota:** Gli step DSL non sono disponibili per i workout pianificati perché l'API Garmin non fornisce la struttura dettagliata nella Workout Library.

#### Sheet 2: "Activities" (Attività Completate)

| Week | Date | Time | Session | Sport | Description | Distance (km) | Duration | Avg Pace | Avg HR | Calories | Steps | ActivityId |
|------|------|------|---------|-------|-------------|---------------|----------|----------|--------|----------|-------|------------|
| 3 | 2026-01-21 | 19:26 | 1 | running | Modena Corsa | 7.84 | 47:15 | 6:01 | 145 | 456 | `warmup: 5min @ Z2` / `interval: 37min @ 6:01` / `cooldown: 5min @ Z1` | 21622466117 |
| 3 | 2026-01-18 | 18:45 | 2 | running | Long Run | 10.01 | 63:15 | 6:19 | 138 | 672 | `warmup: 6min @ Z2` / `interval: 51min @ 6:19` / `cooldown: 6min @ Z1` | 21609834521 |

**Colonne:**
- **Week**: Settimana dell'anno della corsa
- **Date**: Data dell'attività
- **Time**: Ora di inizio
- **Session**: Numero progressivo
- **Sport**: running, cycling, swimming, lap_swimming
- **Description**: Nome dell'attività
- **Distance (km)**: Distanza effettiva percorsa
- **Duration**: Durata effettiva (mm:ss)
- **Avg Pace (min/km)**: Pace medio (calcolato automaticamente)
- **Avg HR (bpm)**: Frequenza cardiaca media
- **Calories**: Calorie bruciate
- **Steps**: **DSL GENERATO AUTOMATICAMENTE** ⭐ (questa è la colonna più importante!)
- **ActivityId**: ID univoco Garmin

**🎯 COLONNA "STEPS" - VALORE PER L'AI:**

Questa colonna contiene la **sintassi DSL ricostruita automaticamente** dalle metriche reali dell'attività:

```
Activity reale: 7.84 km in 47:15 (pace 6:01/km)
↓ Algoritmo di conversione ↓
DSL generato (una riga per step, ricaricabile così com'è):
warmup: 5min @ Z2
interval: 37min @ 6:01
cooldown: 5min @ Z1
```

**Perché è importante:**
- ✅ L'AI vede **esempi reali** di sintassi DSL
- ✅ Ogni attività diventa un **training example**
- ✅ L'AI impara il formato **senza bisogno del parser code**
- ✅ Puoi chiedere all'AI di generare nuovi workout usando questi esempi

**Esempio prompt per AI:**
```
"Ho fatto 138 corse nell'ultimo anno. Guarda la colonna Steps 
nel foglio Activities per vedere come sono strutturati gli allenamenti.
Creami un piano di 8 settimane per una mezza maratona usando 
la stessa sintassi."
```

L'AI analizzerà i tuoi allenamenti reali e genererà un piano personalizzato!

#### Sheet 3: "Summary" (Statistiche)

```
Metric                    Value
Download Date             2026-01-26 23:01:07
Period Start              2025-01-26
Period End                2026-01-26
Total Workouts            170
Total Activities          138

Workouts by Sport
  - running               170

Activities by Sport
  - running               138

Total Distance (km)       1042.07
Total Duration (hours)    109.67
Average Distance (km)     7.55
```

Contiene statistiche aggregate su tutto il periodo scaricato.

#### Sheet 4: "Parameters" (Zone Definitions)

```
Key      Metric   Expression   Notes
Z1       pace     6:30-7:00    Recovery
Z2       pace     5:50-6:20    Easy
Z3       pace     5:20-5:45    Tempo
Z4       pace     4:50-5:15    Threshold
Z5       pace     4:20-4:45    VO2max

FTP_Z1   power    0.55         Active Recovery
FTP_Z2   power    0.56-0.75    Endurance
FTP_Z3   power    0.76-0.90    Tempo
FTP_Z4   power    0.91-1.05    Lactate Threshold
FTP_Z5   power    1.06-1.20    VO2max
```

Fornisce esempi di zone per l'AI, così può capire come usare Z1, Z2, etc. nei workout generati.

---

#### 📋 Formato JSON (.json)

Il file JSON mantiene la struttura originale dei dati Garmin:

```json
{
  "download_info": {
    "date": "2026-01-26T23:01:07",
    "period": {
      "start": "2025-01-26",
      "end": "2026-01-26"
    }
  },
  "scheduled_workouts": {
    "total": 170,
    "summary": {
      "by_sport": {"running": 170},
      "by_month": {"2025-01": 10, "2025-02": 15, ...}
    },
    "data": [
      {
        "workoutId": 1454947666,
        "workoutName": "W9S27 - Activation",
        "sportType": {"sportTypeKey": "running"},
        "createdDate": "2026-01-26T17:10:38.0",
        "estimatedDurationInSecs": 1800,
        "estimatedDistanceInMeters": 5000,
        ...
      }
    ]
  },
  "completed_activities": {
    "total": 138,
    "summary": {
      "by_sport": {"running": 138},
      "by_month": {"2025-01": 4, "2025-02": 12, ...}
    },
    "data": [
      {
        "activityId": 21622466117,
        "activityName": "Modena Corsa",
        "activityType": {"typeKey": "running"},
        "startTimeLocal": "2026-01-21 19:26:39",
        "distance": 7835.93,
        "duration": 2834.78,
        "averageHR": 145,
        "averageSpeed": 2.764,
        "calories": 456,
        ...
      }
    ]
  }
}
```

**Struttura JSON:**
- **download_info**: Metadati del download (data, periodo)
- **scheduled_workouts**: Array completo dei workout con tutti i campi API
- **completed_activities**: Array completo delle attività con tutti i campi API
- **summary**: Statistiche aggregate per sport e mese

**Quando usare JSON:**
- Analisi programmatica con Python/JavaScript
- Import in database
- Integrazione con altri sistemi
- Backup completo dei dati
- Sviluppo di tool personalizzati

---

### 🤖 PARTE 4: Uso del File Excel con AI

#### 4.1 Workflow Base

1. **Scarica i tuoi dati** dall'ultimo anno (o più)
2. **Carica l'Excel** a Claude/ChatGPT
3. **Chiedi all'AI di analizzare** la colonna "Steps" nel foglio Activities
4. **Genera nuovi workout** basati sui tuoi esempi reali

#### 4.2 Esempi di Prompt

**Generazione Piano Allenamento:**
```
"Analizza i miei allenamenti nel foglio Activities (colonna Steps).
Crea un piano di 12 settimane per preparare una mezza maratona,
usando la stessa sintassi DSL che vedi negli esempi.
Voglio 4 allenamenti a settimana."
```

**Analisi Progressione:**
```
"Guarda le mie corse nel foglio Activities.
Analizza la mia progressione negli ultimi 6 mesi.
Quali sono i miei pace medi per distanza?"
```

**Generazione Workout Specifico:**
```
"Guardando i miei allenamenti passati, crea un workout interval
per migliorare la mia soglia. Usa la sintassi della colonna Steps."
```

**Confronto Periodi:**
```
"Confronta i miei allenamenti del Q1 2025 vs Q4 2024.
Distanza totale, pace medio, frequenza settimanale."
```

#### 4.3 Perché Funziona

L'AI può:
- ✅ **Vedere la sintassi reale** senza bisogno di documentazione
- ✅ **Imparare dai tuoi dati** effettivi (pace, distanze, strutture)
- ✅ **Generare workout compatibili** pronti per l'upload
- ✅ **Personalizzare i piani** sul tuo livello attuale

**Il file Excel è auto-esplicativo** - contiene sia i dati che gli esempi di formato!

---

## 🔧 Risoluzione Problemi

### Errore Login Garmin
- Verifica email e password
- Controlla il codice MFA se abilitato
- Elimina la cartella `~/.garminplanner/tokens` e rifai il login con email e password
- Per vedere i JSON scambiati con Garmin avvia con `GARMINPLANNER_DEBUG=1 python launcher.py`

### Workout Non Visibile su Garmin
- Verifica di aver fatto il login
- Controlla che il WorkoutId sia stato salvato nell'Excel
- Ricarica/sincronizza l'app Garmin sul dispositivo

### Errore nella Sintassi DSL
- Leggi il messaggio: indica riga Excel, riga del DSL e il motivo
- Formato: `tipo: durata @ target`, una riga per step (niente `;` né `5x(...)`)
- Sotto `repeat N:` gli step vanno rientrati
- `10m` sono 10 metri: per i minuti scrivi `10min`

### DateEntry Non Funziona
```bash
pip install tkcalendar
```

### Download Non Trova Dati
- Verifica di essere connesso a Garmin Connect (fai login prima)
- Controlla di avere effettivamente dati nel periodo selezionato
- Prova con "Entrambi" e "Tutti gli sport" per un test completo
- Estendi il periodo (es. "Ultimo anno")

### Excel Corrotto o Errore "At least one sheet must be visible"
- Aggiorna `download_dialog.py` all'ultima versione
- Verifica di avere pandas e openpyxl aggiornati: `pip install --upgrade pandas openpyxl`

### Download Lento
- È normale! L'API Garmin può richiedere 30-60 secondi per scaricare dati di un anno
- Per periodi molto lunghi (2+ anni) può richiedere 1-2 minuti

---

## 📁 Struttura Progetto

```
garminPlannerCGPT/
├── garmin_planner.py                   # App (finestra principale)
├── workout_editor.py                   # Editor allenamento: profilo, step, testo DSL
├── workout_model.py                    # Albero step <-> DSL, stime durata/km/intensità
├── app_theme.py                        # Tema chiaro/scuro e colori
├── launcher.py                         # Avvio (compatibilità)
├── gui_helpers.py                      # Finestra errori di validazione
├── legacy/                             # Vecchie interfacce v2 (riserva)
├── download_dialog.py                  # Dialog download
├── garmin_service.py                   # Login/sessione Garmin Connect
├── garmin_client.py                    # Chiamate API workout-service
├── dsl_parser.py                       # Parser DSL + validazione
├── excel_utils.py                      # Template e formattazione Excel
├── requirements.txt / requirements-dev.txt
├── tests/                              # Test pytest (parser, token)
└── examples/
    ├── esempio_running.xlsx
    └── esempio_multisport.xlsx
```

---

## 🛠️ Sviluppo

### Setup Ambiente di Sviluppo

```bash
# Crea virtual environment
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows

# Installa dipendenze
pip install -r requirements.txt

# Installa dipendenze di sviluppo
pip install -r requirements-dev.txt
```

### Eseguire i Test

```bash
python -m pytest -q
```
I test non richiedono né la GUI né la connessione a Garmin.

### Build Eseguibile

```bash
python -m PyInstaller training_planner_gui.spec

# Risultato in dist/ (GarminTrainingPlanner.app su macOS, .exe su Windows)
```

---

## 🤝 Contribuire

I contributi sono benvenuti! Per favore:

1. Fork il progetto
2. Crea un branch per la tua feature (`git checkout -b feature/AmazingFeature`)
3. Commit le modifiche (`git commit -m 'Add some AmazingFeature'`)
4. Push al branch (`git push origin feature/AmazingFeature`)
5. Apri una Pull Request

---

## 📝 Changelog

### v3.2.0 (2026-09-30)
- 🌗 Tema scuro leggibile anche senza sv-ttk (tema di riserva completo) e barra del titolo scura su Windows
- 🧰 Barra degli strumenti che va a capo se la finestra è stretta


### v3.1.0 (2026-09-30)
- 🖱️ Trascinamento col mouse per riordinare gli step (anche dentro/fuori dalle ripetizioni) e gli allenamenti nella lista (anche tra settimane)
- 📋 Menu col tasto destro nella lista; "Duplica nella settimana successiva"; Sposta su/giù; Ordina per data; ⌘D
- 🌗 Tema scuro esteso a calendario delle date, menu, tendine, finestra errori e download


### v3.0.0 (2026-09-30)
- 🎨 Nuova interfaccia unica al posto di Classic/Advanced/Multisport (spostate in `legacy/`)
- 🌗 Tema chiaro/scuro (sv-ttk), segue le impostazioni del sistema
- 📅 Lista per settimane con durata, km stimati e stato di ogni allenamento; ricerca e filtri
- ✏️ Editor integrato: profilo grafico dell'intensità, step con dialog guidato, modelli rapidi, testo DSL evidenziato con controllo live
- 📊 Panoramica con i km per settimana; gestione parametri con dialog
- 💾 Il salvataggio conserva gli altri fogli del file Excel (es. Impostazioni, Esempi DSL)
- 🔁 "Carica e pianifica" salta gli allenamenti già in calendario alla stessa data


### v2.1.0 (2026-09-30)
**Sicurezza**
- 🔐 Token di sessione spostati in `~/.garminplanner/tokens` (permessi 600); migrazione automatica dalla vecchia cartella `./garminconnect`
- 🔐 Nessun token, dato personale o file di sistema nel repository (`.gitignore`)
- 🔐 Password cancellata dal campo dopo il login; log dei JSON solo con `GARMINPLANNER_DEBUG=1`

**Affidabilità**
- ✅ Il parser segnala gli errori (riga Excel + riga DSL) invece di creare step da 60 secondi o senza target
- ✅ Validazione di tutti i workout selezionati prima di contattare Garmin
- ✅ Gli ID Garmin vengono salvati nell'Excel anche se il caricamento si interrompe a metà (niente doppioni al nuovo tentativo)
- ✅ Chiavi di Parameters senza distinzione maiuscole/minuscole (Swim_Z*, Cadence_* ora funzionano)
- ✅ Durate decimali e in ore, indentazione libera sotto `repeat`, step `rest` inviati come Riposo
- ✅ Il download genera il DSL nel formato reale (ricaricabile)
- 🧹 Rimossi `dsl_parser_multisport.py` ed `excel_utils_multisport.py` (duplicati mai usati); caratteri accentati corretti
- 🧪 Test automatici (`tests/`)


### v2.0.0 (2026-01-26) ⭐ MAJOR RELEASE
**Nuove Funzionalità:**
- ✨ **Download workout e attività** da Garmin Connect con interfaccia dedicata
- ✨ **Filtro sport intelligente** (running, cycling, swimming)
- ✨ **Selezione periodo visuale** con calendario DateEntry
- ✨ **Export multi-formato**: Excel (.xlsx) o JSON (.json) ⭐ NEW
- ✨ **Export Excel multi-sheet** (Workouts, Activities, Summary, Parameters)
- ✨ **Export JSON completo** con dati grezzi originali API
- ✨ **Generazione automatica DSL** dalle attività completate per training AI
- ✨ **Nome file intelligente** auto-generato ed editabile
- ✨ **Statistiche dettagliate** per sport e mese
- ✨ **Colonna "Steps"** con sintassi DSL per ogni attività (AI-ready)

**Miglioramenti:**
- 🔧 Gestione valori `None` dall'API Garmin
- 🔧 Calcolo automatico settimana dell'anno
- 🔧 Gestione errori migliorata con logging dettagliato
- 🔧 Limite di 100 righe per sheet Excel per stabilità
- 🔧 Layout GUI ottimizzato per accogliere nuovo pulsante Download
- 🔧 Scelta formato export (Excel/JSON) con selezione radio button

**Bug Fix:**
- 🐛 Fix: Gestione encoding emoji su Windows
- 🐛 Fix: Layout GUI con DateEntry
- 🐛 Fix: Confronto None con valori numerici
- 🐛 Fix: Sheet Excel vuoti causavano corruzione file

### v1.5.0 (2025-11-27)
- ✨ Aggiunti: Preset workout per multi-sport
- ✨ Aggiunto: Visual builder con drag-and-drop
- 🔧 Migliorato: Parser DSL per swimming
- 🐛 Fix: Conversione pace nuoto per API Garmin

### v1.0.0 (2025-10-01)
- 🎉 Release iniziale
- ✅ Upload workout su Garmin Connect
- ✅ Pianificazione calendario
- ✅ DSL per running e cycling

---

## 📜 Licenza

Questo progetto è rilasciato sotto licenza MIT. Vedi il file `LICENSE` per i dettagli.

---

## 🙏 Ringraziamenti

- **Garth**: Libreria Python per Garmin Connect API
- **Pandas**: Gestione dati Excel
- **tkcalendar**: Widget calendario per Tkinter
- **OpenPyXL**: Engine per creazione file Excel
- Community Garmin Connect per la documentazione delle API

---

## 📧 Supporto

Per bug, richieste di funzionalità o domande:
- 🐛 [Issues](https://github.com/yesnoj/garminPlannerCGPT/issues)
- 💬 [Discussions](https://github.com/yesnoj/garminPlannerCGPT/discussions)
- 📧 Email: tuoemail@example.com

---

## ⚠️ Disclaimer

Questo progetto non è affiliato, associato, autorizzato, approvato da, o in alcun modo ufficialmente connesso con Garmin Ltd. o le sue sussidiarie o affiliate. Il nome "Garmin" e i relativi nomi, marchi, emblemi e immagini sono marchi registrati dei rispettivi proprietari.

Usa questo software a tuo rischio. Gli autori non sono responsabili per eventuali danni o perdite derivanti dall'uso di questo software.

---

## 🌟 Se Ti Piace il Progetto

Se trovi utile questo progetto:
- ⭐ Metti una stella su GitHub
- 🐛 Segnala bug o richiedi funzionalità
- 🤝 Contribuisci con codice o documentazione
- 📢 Condividi con altri atleti!

---

## 💡 Tips & Tricks

### Per Massimizzare l'Uso con AI

1. **Scarica dati dell'ultimo anno** per avere dataset completo
2. **Usa "Entrambi"** per confrontare workout pianificati vs eseguiti
3. **La colonna Steps nelle Activities** è oro per l'AI - contiene la sintassi reale
4. **Carica l'Excel insieme al prompt** specifico per risultati migliori
5. **Salva i piani generati** e caricali con questo software!

### Per Migliori Performance

1. **Filtra per sport** se hai molti dati misti
2. **Limita il periodo** a 6-12 mesi per download più veloci
3. **Fai backup regolari** dei tuoi Excel scaricati
4. **Usa nomi file descrittivi** per organizzare i download

### Workflow Consigliato

```
1. Scarica dati ultimo anno → garmin_activities_2025_running.xlsx
2. Analizza con AI → Genera piano 12 settimane
3. AI crea Excel con workout → piano_mezza_maratona.xlsx
4. Carica workout su Garmin → Upload + Pianifica
5. Segui il piano e completa allenamenti
6. Ri-scarica dopo 3 mesi → Analizza progressione
7. Ripeti ciclo! 🔄
```

---

**Buon allenamento! 🏃‍♂️🚴‍♂️🏊‍♂️**

*Versione 2.0.0 - Gennaio 2026*