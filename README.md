# 🏃 Training Planner - Garmin Workout Manager

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Un'applicazione desktop Python per creare, gestire e sincronizzare piani di allenamento strutturati con **Garmin Connect**. Supporta **Running**, **Cycling** e **Swimming** con un linguaggio DSL intuitivo per definire workout complessi.

![Training Planner Screenshot](docs/screenshot.png)

---

## 📋 Indice

- [Caratteristiche](#-caratteristiche)
- [Installazione](#-installazione)
- [Quick Start](#-quick-start)
- [Struttura del File Excel](#-struttura-del-file-excel)
- [Guida al DSL (Domain Specific Language)](#-guida-al-dsl-domain-specific-language)
- [Versioni della GUI](#-versioni-della-gui)
- [Esempi Completi di Workout](#-esempi-completi-di-workout)
- [Parametri e Zone](#-parametri-e-zone)
- [Troubleshooting](#-troubleshooting)
- [Struttura del Progetto](#-struttura-del-progetto)
- [Contribuire](#-contribuire)
- [Licenza](#-licenza)

---

## ✨ Caratteristiche

### Funzionalità Principali
- 📊 **Gestione workout via Excel** - Pianifica i tuoi allenamenti in un foglio Excel strutturato
- 🔄 **Sincronizzazione Garmin Connect** - Upload automatico dei workout sul tuo account Garmin
- 📅 **Pianificazione calendario** - Programma automaticamente i workout nelle date desiderate
- 🎨 **Editor Visuale** - Costruisci workout con drag-and-drop (versione Advanced)
- 🏊 **Multi-Sport** - Supporto completo per Running, Cycling e Swimming

### Sport Supportati

| Sport | Target Supportati |
|-------|-------------------|
| 🏃 **Running** | Pace (min/km), Zone ritmo (Z1-Z5), HR zones |
| 🚴 **Cycling** | Potenza (Watt), Zone FTP, Cadenza (RPM), HR zones |
| 🏊 **Swimming** | Pace nuoto (min/100m), Zone swim, HR zones |

### DSL Potente e Flessibile
- Sintassi semplice e leggibile
- Supporto per ripetizioni (anche annidate)
- Zone personalizzabili
- Tolleranze configurabili

---

## 🚀 Installazione

### Prerequisiti
- Python 3.8 o superiore
- pip (package manager Python)

### Installazione Dipendenze

```bash
# Clona il repository
git clone https://github.com/tuousername/training-planner.git
cd training-planner

# Crea un virtual environment (opzionale ma consigliato)
python -m venv venv
source venv/bin/activate  # Linux/Mac
# oppure
venv\Scripts\activate  # Windows

# Installa le dipendenze
pip install pandas openpyxl garth tkcalendar
```

### Dipendenze
| Pacchetto | Descrizione |
|-----------|-------------|
| `pandas` | Gestione dati Excel |
| `openpyxl` | Lettura/scrittura file .xlsx |
| `garth` | Autenticazione e API Garmin Connect |
| `tkcalendar` | Widget calendario per la GUI (opzionale) |

---

## 🎯 Quick Start

### 1. Avvia il Launcher

```bash
python launcher.py
```

Scegli la versione della GUI:
- **Standard** - Funzionalità essenziali
- **Advanced** - Editor visuale e funzioni avanzate  
- **Multisport** - Supporto completo per tutti gli sport

### 2. Genera un File Excel di Esempio

Clicca su **"Genera Excel di esempio"** per creare un file template con workout pre-configurati.

### 3. Modifica i Workout

Apri il file Excel e modifica:
- **Foglio "Workouts"** - I tuoi allenamenti
- **Foglio "Parameters"** - Le tue zone personalizzate

### 4. Connettiti a Garmin

1. Inserisci le credenziali Garmin Connect
2. Clicca **"Login"**
3. Se richiesto, inserisci il codice MFA ricevuto via email

### 5. Carica i Workout

1. Seleziona i workout nella lista
2. Clicca **"📤 Carica su Garmin"** per caricarli
3. Oppure **"📤📅 Carica + Pianifica"** per caricarli e pianificarli automaticamente

---

## 📁 Struttura del File Excel

Il file Excel contiene tre fogli principali:

### Foglio "Workouts"

| Colonna | Descrizione | Esempio |
|---------|-------------|---------|
| `Week` | Numero settimana | 1 |
| `Date` | Data dell'allenamento | 2025-05-15 |
| `Session` | Numero sessione | 1 |
| `Sport` | Tipo di sport | Running, Cycling, Swimming |
| `Description` | Nome del workout | "Ripetute 4×5' @ Z4" |
| `Steps` | Definizione DSL del workout | (vedi sotto) |
| `WorkoutId` | ID Garmin (auto-generato) | - |
| `WorkoutScheduleId` | ID pianificazione (auto) | - |
| `ScheduledDate` | Data pianificata (auto) | - |

### Foglio "Parameters"

| Colonna | Descrizione | Esempio |
|---------|-------------|---------|
| `Key` | Nome del parametro | Z1, HR_max, Power_Z4 |
| `Metric` | Tipo di metrica | pace, hr, power |
| `Expression` | Valore o range | 6:00, 185, 200-250W |
| `Notes` | Descrizione | Zona recupero |

### Foglio "Esempi DSL"

Contiene esempi di sintassi DSL per riferimento rapido.

---

## 📝 Guida al DSL (Domain Specific Language)

Il DSL (Domain Specific Language) è il linguaggio utilizzato per definire gli step dei workout nella colonna "Steps".

### Struttura Base

```
tipo_step: durata @ target
```

Dove:
- **tipo_step**: warmup, interval, recovery, rest, cooldown
- **durata**: tempo o distanza
- **target**: zona o valore specifico (opzionale)

### Tipi di Step

| Tipo | Descrizione | Uso |
|------|-------------|-----|
| `warmup` | Riscaldamento | Inizio workout |
| `interval` | Intervallo intenso | Fase principale |
| `recovery` | Recupero attivo | Tra ripetute (corsa lenta) |
| `rest` | Riposo completo | Fermo, senza movimento |
| `cooldown` | Defaticamento | Fine workout |

### Formati Durata

#### Tempo
```dsl
warmup: 10min @ Z2      # 10 minuti
interval: 30sec @ Z5    # 30 secondi
warmup: 10' @ Z2        # 10 minuti (notazione alternativa)
interval: 30s @ Z5      # 30 secondi (notazione alternativa)
```

#### Distanza
```dsl
interval: 1000m @ Z4    # 1000 metri
interval: 5km @ Z3      # 5 chilometri
interval: 400m @ Z5     # 400 metri
```

#### Lap Button
```dsl
warmup: lap-button @ Z2  # Premi lap per continuare
```

### Target (Zone e Valori)

#### Running - Pace
```dsl
# Zone predefinite (da Parameters)
interval: 5min @ Z1     # Zona 1 - Recupero
interval: 5min @ Z2     # Zona 2 - Aerobica
interval: 5min @ Z3     # Zona 3 - Tempo
interval: 5min @ Z4     # Zona 4 - Soglia
interval: 5min @ Z5     # Zona 5 - VO2max

# Pace specifico (min:sec per km)
interval: 3km @ 5:00           # 5:00/km fisso
interval: 3km @ 5:00-5:30      # Range 5:00-5:30/km

# Parametri custom (definiti in Parameters)
warmup: 10min @ easy_range     # Range ritmo facile
interval: 5km @ marathon       # Ritmo maratona
interval: 1km @ threshold      # Ritmo soglia
```

#### Running/Cycling/Swimming - Heart Rate
```dsl
# Zone HR (percentuali di HR_max)
interval: 20min @ HR_Z1    # 60-70% HR_max
interval: 20min @ HR_Z2    # 70-80% HR_max
interval: 20min @ HR_Z3    # 80-85% HR_max
interval: 20min @ HR_Z4    # 85-90% HR_max
interval: 20min @ HR_Z5    # 90-100% HR_max

# HR assoluto (bpm)
interval: 15min @ 150-165  # Range 150-165 bpm
```

#### Cycling - Potenza
```dsl
# Zone potenza (da Parameters, basate su FTP)
interval: 8min @ Power_Z1    # Recovery (0-55% FTP)
interval: 8min @ Power_Z2    # Endurance (56-75% FTP)
interval: 8min @ Power_Z3    # Tempo/Sweet Spot (76-90% FTP)
interval: 8min @ Power_Z4    # Threshold (91-105% FTP)
interval: 8min @ Power_Z5    # VO2max (106-120% FTP)

# Potenza diretta (Watt)
interval: 10min @ 200W        # 200 Watt fisso
interval: 10min @ 200-250W    # Range 200-250 Watt
```

#### Cycling - Cadenza
```dsl
# Cadenza da Parameters
interval: 5min @ Cadence_Easy      # 70-80 rpm
interval: 5min @ Cadence_Tempo     # 85-95 rpm
interval: 5min @ Cadence_Sprint    # 100-110 rpm

# Cadenza diretta (rpm)
interval: 3min @ 90rpm        # 90 rpm fisso
interval: 3min @ 85-95rpm     # Range 85-95 rpm
```

#### Swimming - Pace
```dsl
# Zone pace nuoto (min:sec per 100m)
interval: 400m @ Swim_Z1    # Recovery/Warm-up
interval: 400m @ Swim_Z2    # Endurance
interval: 200m @ Swim_Z3    # Tempo
interval: 100m @ Swim_Z4    # Threshold
interval: 50m @ Swim_Z5     # Sprint

# Pace custom
warmup: 300m @ swim_easy         # Pace facile
interval: 200m @ swim_threshold  # Pace soglia

# Pace diretto (min:sec per 100m)
interval: 100m @ 1:45            # 1:45/100m
interval: 100m @ 1:40-1:50       # Range
```

### Ripetizioni (Repeat)

Le ripetizioni permettono di definire blocchi che si ripetono più volte.

#### Sintassi Base
```dsl
repeat N:
  step1
  step2
```

**IMPORTANTE**: Gli step dentro il repeat devono essere indentati con **2 spazi**.

#### Esempio Semplice
```dsl
warmup: 10min @ Z2
repeat 4:
  interval: 1km @ Z4
  recovery: 2min @ Z1
cooldown: 5min @ Z1
```

Questo crea:
1. Warmup 10 minuti
2. 4 ripetizioni di: 1km veloce + 2min recupero
3. Cooldown 5 minuti

#### Ripetizioni Annidate
```dsl
warmup: 15min @ Z2
repeat 2:
  repeat 6:
    interval: 1min @ Z5
    recovery: 1min @ Z1
  rest: 3min
cooldown: 10min @ Z1
```

Questo crea:
1. Warmup 15 minuti
2. 2 serie da:
   - 6 ripetizioni di: 1min veloce + 1min recupero
   - 3 minuti di riposo completo tra le serie
3. Cooldown 10 minuti

### Commenti

Puoi aggiungere commenti usando `--`:

```dsl
warmup: 10min @ Z2 -- riscaldamento graduale
interval: 5km @ marathon -- ritmo gara
```

I commenti vengono ignorati dal parser ma rimangono visibili nel campo Steps.

---

## 🖥️ Versioni della GUI

### Standard (`training_planner_gui.py`)

La versione base con tutte le funzionalità essenziali:
- ✅ Caricamento/salvataggio Excel
- ✅ Upload workout su Garmin
- ✅ Pianificazione calendario
- ✅ Modifica parametri
- ✅ Editor DSL testuale

**Consigliata per**: Utenti che preferiscono lavorare direttamente con il DSL.

### Advanced (`training_planner_gui_advanced.py`)

Versione con funzionalità avanzate:
- ✅ Tutte le funzionalità Standard
- ✅ **Editor Visuale** con drag-and-drop
- ✅ Libreria preset workout
- ✅ Anteprima workout in tempo reale
- ✅ Creazione nuovo workout guidata

**Consigliata per**: Utenti che preferiscono un'interfaccia visuale.

### Multisport (`training_planner_multisport_gui.py`)

Versione completa per triatleti e atleti multisport:
- ✅ Tutte le funzionalità Advanced
- ✅ **Preset organizzati per sport** (Running, Cycling, Swimming)
- ✅ Zone specifiche per ogni sport
- ✅ Parametri multisport preconfigurati

**Consigliata per**: Triatleti e chi pratica più sport.

---

## 📚 Esempi Completi di Workout

### 🏃 Running

#### Ripetute Classiche 10×400m
```dsl
warmup: 15min @ Z2
repeat 10:
  interval: 400m @ Z5
  rest: 90sec
cooldown: 10min @ Z1
```

#### Lungo con Progressione
```dsl
warmup: 10min @ easy_range
interval: 30min @ Z2
interval: 20min @ Z3
interval: 10min @ Z4
cooldown: 10min @ Z1
```

#### Fartlek 8×(1' veloce / 2' recupero)
```dsl
warmup: 10min @ Z2
repeat 8:
  interval: 1min @ Z5
  recovery: 2min @ Z1
cooldown: 10min @ Z1
```

#### Piramidale
```dsl
warmup: 10min @ Z2
interval: 400m @ Z5
rest: 90sec
interval: 800m @ Z4
rest: 2min
interval: 1200m @ Z4
rest: 3min
interval: 800m @ Z4
rest: 2min
interval: 400m @ Z5
cooldown: 10min @ Z1
```

#### Progressivo con HR
```dsl
interval: 10min @ HR_Z1
interval: 15min @ HR_Z2
interval: 10min @ HR_Z3
interval: 5min @ HR_Z4
cooldown: 5min @ HR_Z1
```

### 🚴 Cycling

#### FTP Intervals
```dsl
warmup: 15min @ Power_Z2
repeat 4:
  interval: 8min @ Power_Z4
  recovery: 4min @ Power_Z1
cooldown: 10min @ Power_Z1
```

#### Sweet Spot Training
```dsl
warmup: 15min @ Power_Z2
repeat 3:
  interval: 15min @ Power_Z3
  recovery: 5min @ Power_Z1
cooldown: 10min @ Power_Z1
```

#### VO2max Intervals
```dsl
warmup: 20min @ Power_Z2
repeat 5:
  interval: 3min @ Power_Z5
  recovery: 3min @ Power_Z1
cooldown: 15min @ Power_Z1
```

#### Endurance + Cadence Drills
```dsl
warmup: 10min @ Power_Z2
interval: 45min @ Power_Z2
repeat 6:
  interval: 2min @ Cadence_Sprint
  recovery: 2min @ Cadence_Easy
cooldown: 10min @ Power_Z1
```

### 🏊 Swimming

#### Threshold Set
```dsl
warmup: 400m @ Swim_Z1
repeat 6:
  interval: 100m @ Swim_Z4
  rest: 20sec
cooldown: 200m @ Swim_Z1
```

#### Pyramid Swim
```dsl
warmup: 400m @ Swim_Z1
interval: 100m @ Swim_Z4
rest: 30sec
interval: 200m @ Swim_Z4
rest: 45sec
interval: 300m @ Swim_Z3
rest: 60sec
interval: 200m @ Swim_Z4
rest: 45sec
interval: 100m @ Swim_Z5
cooldown: 200m @ Swim_Z1
```

#### Sprint Set
```dsl
warmup: 400m @ Swim_Z1
repeat 10:
  interval: 50m @ Swim_Z5
  rest: 30sec
interval: 200m @ swim_easy
repeat 5:
  interval: 25m @ Swim_Z5
  rest: 20sec
cooldown: 200m @ Swim_Z1
```

---

## ⚙️ Parametri e Zone

### Configurazione Running

```
| Key            | Expression  | Notes                                    |
|----------------|-------------|------------------------------------------|
| Z1             | 6:30        | Zona 1: Recupero attivo (min/km)         |
| Z2             | 6:00        | Zona 2: Endurance/Aerobica               |
| Z3             | 5:30        | Zona 3: Tempo/Ritmo gara                 |
| Z4             | 5:00        | Zona 4: Soglia anaerobica                |
| Z5             | 4:30        | Zona 5: VO2max/Ripetute                  |
| easy_range     | 6:00-6:30   | Range ritmo facile                       |
| marathon       | 5:20        | Ritmo maratona                           |
| threshold      | 5:10        | Ritmo soglia                             |
| pace_tolerance | 5           | Tolleranza ±sec per pace singoli         |
```

### Configurazione Cycling

```
| Key            | Expression  | Notes                                    |
|----------------|-------------|------------------------------------------|
| FTP            | 250         | Functional Threshold Power (Watt)        |
| Power_Z1       | 0-140W      | Recovery (0-55% FTP)                     |
| Power_Z2       | 141-195W    | Endurance (56-75% FTP)                   |
| Power_Z3       | 196-225W    | Tempo/Sweet Spot (76-90% FTP)            |
| Power_Z4       | 226-250W    | Threshold (91-105% FTP)                  |
| Power_Z5       | 251-300W    | VO2max (106-120% FTP)                    |
| Cadence_Easy   | 70-80rpm    | Cadenza recupero                         |
| Cadence_Tempo  | 85-95rpm    | Cadenza gara                             |
| Cadence_Sprint | 100-110rpm  | Cadenza sprint                           |
```

### Configurazione Swimming

```
| Key            | Expression  | Notes                                    |
|----------------|-------------|------------------------------------------|
| Swim_Z1        | 2:30        | Recovery/Warm-up (min:sec per 100m)      |
| Swim_Z2        | 2:10        | Endurance                                |
| Swim_Z3        | 1:55        | Tempo                                    |
| Swim_Z4        | 1:45        | Threshold                                |
| Swim_Z5        | 1:30        | Sprint                                   |
| swim_easy      | 2:20        | Pace facile riscaldamento                |
| swim_threshold | 1:50        | Pace soglia                              |
| swim_tolerance | 5           | Tolleranza ±sec per pace singoli         |
```

### Configurazione Heart Rate (Universale)

```
| Key            | Expression  | Notes                                    |
|----------------|-------------|------------------------------------------|
| HR_max         | 185         | Frequenza cardiaca massima (PERSONALIZZA)|
| HR_Z1          | 60-70%      | Zona FC 1: Recupero                      |
| HR_Z2          | 70-80%      | Zona FC 2: Aerobica                      |
| HR_Z3          | 80-85%      | Zona FC 3: Tempo                         |
| HR_Z4          | 85-90%      | Zona FC 4: Soglia                        |
| HR_Z5          | 90-100%     | Zona FC 5: VO2max                        |
| hr_tolerance   | 5           | Tolleranza ±bpm per HR singoli           |
```

### Come Funzionano le Tolleranze

Quando specifichi un valore singolo (es. `5:00` per il pace), il sistema applica automaticamente la tolleranza configurata:

```
# Con pace_tolerance = 5 secondi:
interval: 3km @ 5:00
# Diventa range: 4:55 - 5:05

# Con hr_tolerance = 5 bpm:
interval: 20min @ 150
# Diventa range: 145 - 155 bpm
```

Se specifichi già un range, la tolleranza viene ignorata:
```
interval: 3km @ 5:00-5:30  # Usa esattamente 5:00-5:30
```

---

## 🔧 Troubleshooting

### Errore "File non trovato"

**Problema**: Il launcher non trova i file GUI.

**Soluzione**: Assicurati che tutti i file `.py` siano nella stessa cartella:
- `launcher.py`
- `training_planner_gui.py`
- `training_planner_gui_advanced.py`
- `training_planner_multisport_gui.py`
- `dsl_parser.py`
- `excel_utils.py`
- `garmin_client.py`
- `garmin_service.py`

### Errore Login Garmin

**Problema**: "401 Unauthorized" o "403 Forbidden"

**Soluzioni**:
1. Verifica che email e password siano corretti
2. Se hai MFA attivo, inserisci il codice quando richiesto
3. Prova a cancellare la cartella `./garminconnect` e rifare il login

### Errore "Sessione scaduta"

**Problema**: La sessione salvata non funziona più.

**Soluzione**: Clicca su "Login" invece di "Usa sessione salvata" per creare una nuova sessione.

### Workout non appare su Garmin

**Problema**: Il workout viene caricato ma non compare nel calendario.

**Soluzioni**:
1. Verifica che la data nel campo `Date` sia nel formato corretto (YYYY-MM-DD)
2. Usa "Carica + Pianifica" invece di solo "Carica"
3. Controlla su Garmin Connect web se il workout è nella libreria

### Errore parsing DSL

**Problema**: Il workout non viene parsato correttamente.

**Soluzioni**:
1. Verifica l'indentazione: usa **esattamente 2 spazi** per i blocchi repeat
2. Controlla che non ci siano caratteri speciali non supportati
3. Verifica la sintassi dei target (es. `Z1` non `z1`)

### Excel non si apre / Errore formattazione

**Problema**: Errore durante il salvataggio Excel.

**Soluzioni**:
1. Chiudi il file Excel se è aperto in un altro programma
2. Verifica di avere i permessi di scrittura nella cartella
3. Prova a salvare con un nome diverso

---

## 📂 Struttura del Progetto

```
training-planner/
├── launcher.py                         # Launcher unificato
├── training_planner_gui.py             # GUI Standard
├── training_planner_gui_advanced.py    # GUI Advanced con editor visuale
├── training_planner_multisport_gui.py  # GUI Multisport completa
├── dsl_parser.py                       # Parser del linguaggio DSL
├── excel_utils.py                      # Utility per gestione Excel
├── garmin_client.py                    # Client API Garmin Connect
├── garmin_service.py                   # Servizio autenticazione Garmin
├── garminconnect/                      # Cartella sessione (auto-generata)
├── README.md                           # Questa documentazione
└── requirements.txt                    # Dipendenze Python
```

### Descrizione Moduli

| File | Descrizione |
|------|-------------|
| `launcher.py` | Entry point principale. Permette di scegliere quale GUI avviare. |
| `dsl_parser.py` | Converte il DSL testuale in struttura JSON compatibile con Garmin API. |
| `excel_utils.py` | Genera file Excel template e gestisce la formattazione. |
| `garmin_client.py` | Wrapper per le chiamate API a Garmin Connect. |
| `garmin_service.py` | Gestisce autenticazione e sessione Garmin. |

---

## 🤝 Contribuire

Contribuzioni, issue e feature request sono benvenute!

### Come Contribuire

1. Forka il repository
2. Crea un branch per la tua feature (`git checkout -b feature/AmazingFeature`)
3. Committa le modifiche (`git commit -m 'Add some AmazingFeature'`)
4. Pusha sul branch (`git push origin feature/AmazingFeature`)
5. Apri una Pull Request

### Idee per Contribuzioni

- [ ] Supporto per altri formati export (CSV, JSON)
- [ ] Integrazione con altri servizi (Strava, TrainingPeaks)
- [ ] App mobile companion
- [ ] Grafici e statistiche workout
- [ ] Import workout da Garmin Connect

---

## 📄 Licenza

Questo progetto è distribuito sotto licenza MIT. Vedi il file `LICENSE` per maggiori dettagli.

---

## 🙏 Ringraziamenti

- [garth](https://github.com/matin/garth) - Libreria per autenticazione Garmin
- [pandas](https://pandas.pydata.org/) - Gestione dati
- [openpyxl](https://openpyxl.readthedocs.io/) - Gestione file Excel
- [tkcalendar](https://github.com/j4321/tkcalendar) - Widget calendario

---

## 📞 Supporto

Per problemi o domande:
- Apri una [Issue](https://github.com/tuousername/training-planner/issues) su GitHub
- Controlla la sezione [Troubleshooting](#-troubleshooting)

---

**Made with ❤️ for runners, cyclists and swimmers**
