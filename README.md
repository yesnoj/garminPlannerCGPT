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

### 🎨 Interfacce Multiple
- **Classic Edition**: Interfaccia essenziale e veloce
- **Advanced Edition**: Editor visuale con drag-and-drop
- **Multi-Sport Edition**: Supporto completo per triathlon

### 🔐 Sicurezza
- ✅ Autenticazione con supporto MFA
- ✅ Sessioni salvate localmente
- ✅ Nessuna password memorizzata

---

## 📦 Requisiti

### Python
- Python 3.8 o superiore
- tkinter (incluso nella distribuzione standard)

### Dipendenze
```bash
pip install pandas openpyxl garth tkcalendar
```

---

## 🚀 Installazione

### 1. Clona il repository
```bash
git clone https://github.com/tuousername/garmin-training-planner.git
cd garmin-training-planner
```

### 2. Installa le dipendenze
```bash
pip install -r requirements.txt
```

### 3. Avvia l'applicazione
```bash
# GUI Classic
python training_planner_gui.py

# GUI Advanced (con editor visuale)
python training_planner_gui_advanced.py

# GUI Multisport (per triathlon)
python training_planner_multisport_gui.py

# Launcher unificato (scegli la GUI all'avvio)
python launcher.py
```

---

## 📖 Guida Completa

### 🎯 PARTE 1: Upload Workout su Garmin

#### 1.1 Primo Avvio

```bash
python launcher.py
```

Scegli la versione GUI:
- 📊 **Standard**: Funzionalità essenziali
- ⚡ **Advanced**: Editor visuale avanzato
- 🏊 **Multisport**: Supporto triathlon completo

#### 1.2 Login a Garmin Connect

1. Clicca su **"🔐 Login"**
2. Inserisci email e password Garmin Connect
3. Se richiesto, inserisci il codice MFA
4. La sessione viene salvata automaticamente per i prossimi utilizzi

#### 1.3 Creare Workout con Excel

##### Genera il Template
1. Clicca su **"📄 Genera Excel"**
2. Salva il file template nella directory desiderata

##### Struttura Excel

**Sheet "Workouts":**
| Week | Date | Session | Sport | Description | Steps |
|------|------|---------|-------|-------------|-------|
| 1 | 2025-01-27 | 1 | running | Easy Run | warmup 10min Z2; run 30min Z3; cooldown 5min Z1 |
| 1 | 2025-01-29 | 2 | running | Intervals | warmup 15min Z2; 5x(run 5min Z4; recover 2min Z1); cooldown 10min Z1 |
| 2 | 2025-02-01 | 3 | cycling | Threshold | warmup 10min FTP_Z2; bike 20min FTP_Z4; cooldown 10min FTP_Z1 |

**Sheet "Parameters":**
| Key | Metric | Expression | Notes |
|-----|--------|------------|-------|
| Z1 | pace | 6:30-7:00 | Recovery |
| Z2 | pace | 5:50-6:20 | Easy |
| Z3 | pace | 5:20-5:45 | Tempo |
| Z4 | pace | 4:50-5:15 | Threshold |
| Z5 | pace | 4:20-4:45 | VO2max |
| FTP_Z1 | power | 0.55 | Active Recovery |
| FTP_Z2 | power | 0.56-0.75 | Endurance |

#### 1.4 DSL Syntax - Linguaggio Workout

##### Step Base
```
<tipo> <durata> <target>
```

**Tipi di step:**
- `warmup` - Riscaldamento
- `run` / `bike` / `swim` - Allenamento principale
- `recover` - Recupero
- `cooldown` - Defaticamento

**Durata:**
- Tempo: `10min`, `30sec`, `2h`
- Distanza: `5km`, `400m`, `100m`
- Manuale: `lap-button`

**Target:**

**Running:**
- Zone pace: `Z1`, `Z2`, `Z3`, `Z4`, `Z5`
- Pace specifico: `5:00`, `4:30-4:45`
- Zone HR: `HR_Z2`, `HR_Z3`
- Aperto: `open`

**Cycling:**
- Potenza: `200W`, `180-220W`
- Zone potenza: `FTP_Z2`, `FTP_Z3`
- Cadenza: `90rpm`, `100rpm`
- Aperto: `open`

**Swimming:**
- Pace: `1:30/100m`, `1:20-1:35/100m`
- Aperto: `open`

##### Ripetizioni

**Sintassi:**
```
<numero>x(<step1>; <step2>; ...)
```

**Esempi:**
```
5x(run 1km Z4; recover 400m Z1)
8x(run 400m Z5; recover 200m Z1)
3x(bike 5min 250W; recover 3min 150W)
```

##### Ripetizioni Annidate

```
3x(
    run 1km Z4; 
    4x(run 200m Z5; recover 100m Z1); 
    recover 400m Z2
)
```

##### Esempi Pratici

**Easy Run:**
```
warmup 10min Z2; run 30min Z3; cooldown 5min Z1
```

**Interval Training:**
```
warmup 15min Z2; 8x(run 400m Z5; recover 200m Z1); cooldown 10min Z1
```

**Tempo Run:**
```
warmup 15min Z2; run 20min Z4; cooldown 10min Z1
```

**Long Run Progressivo:**
```
warmup 10min Z2; run 20min Z3; run 20min Z4; run 10min Z5; cooldown 10min Z1
```

**Cycling Intervals:**
```
warmup 15min FTP_Z2; 5x(bike 5min 250W; recover 3min 150W); cooldown 10min FTP_Z1
```

**Swimming Workout:**
```
warmup 200m open; 10x(swim 100m 1:30/100m; rest 20sec); cooldown 200m open
```

**Fartlek:**
```
warmup 10min Z2; 6x(run 3min Z4; run 2min Z2); cooldown 10min Z1
```

#### 1.5 Upload su Garmin

1. **Carica Excel** con i workout creati
2. Seleziona i workout dalla lista
3. Scegli l'operazione:
   - **📤 Carica**: Solo upload nella libreria Garmin
   - **📤📅 Carica+Pianifica**: Upload e programmazione nel calendario
   - **📅✖ Rimuovi Piano**: Rimuove programmazione (ma non elimina il workout)
   - **🗑 Cancella**: Elimina definitivamente da Garmin

4. Attendi la conferma
5. I workout sono ora disponibili sul tuo dispositivo Garmin!

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
| 3 | 2026-01-21 | 19:26 | 1 | running | Modena Corsa | 7.84 | 47:15 | 6:01 | 145 | 456 | **warmup 5min Z2; run 37min 6:01; cooldown 5min Z1** | 21622466117 |
| 3 | 2026-01-18 | 18:45 | 2 | running | Long Run | 10.01 | 63:15 | 6:19 | 138 | 672 | **warmup 6min Z2; run 51min 6:19; cooldown 6min Z1** | 21609834521 |

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
DSL generato: warmup 5min Z2; run 37min 6:01; cooldown 5min Z1
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
- Elimina la cartella `./garminconnect` e riprova

### Workout Non Visibile su Garmin
- Verifica di aver fatto il login
- Controlla che il WorkoutId sia stato salvato nell'Excel
- Ricarica/sincronizza l'app Garmin sul dispositivo

### Errore nella Sintassi DSL
- Verifica gli spazi: `run 5min Z2` (non `run5minZ2`)
- Usa punto e virgola per separare step: `warmup 10min Z2; run 20min Z3`
- Le parentesi devono essere bilanciate: `5x(...)`

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
garmin-training-planner/
├── launcher.py                          # Launcher unificato
├── training_planner_gui.py             # GUI Classic
├── training_planner_gui_advanced.py    # GUI Advanced
├── training_planner_multisport_gui.py  # GUI Multi-Sport
├── download_dialog.py                  # Dialog download (v2.0) ⭐ NEW
├── garmin_service.py                   # API Garmin Connect
├── dsl_parser.py                       # Parser DSL
├── excel_utils.py                      # Utilità Excel
├── requirements.txt                    # Dipendenze Python
├── README.md                           # Questa guida
└── examples/
    ├── example_plan_running.xlsx
    ├── example_plan_cycling.xlsx
    └── example_plan_multisport.xlsx
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
pip install pytest black flake8
```

### Eseguire i Test

```bash
pytest tests/
```

### Build Eseguibile

```bash
# Windows
python -m PyInstaller training_planner_gui.spec

# Il file .exe sarà in dist/
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
- 🐛 [Issues](https://github.com/tuousername/garmin-training-planner/issues)
- 💬 [Discussions](https://github.com/tuousername/garmin-training-planner/discussions)
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