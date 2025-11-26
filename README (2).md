# Training Planner - Garmin Workout Manager

Applicazione desktop Python per creare, gestire e sincronizzare allenamenti con Garmin Connect usando un linguaggio DSL (Domain Specific Language) semplificato.

## 🎯 Caratteristiche Principali

- ✅ **Interfaccia grafica intuitiva** (Tkinter)
- ✅ **DSL semplificato** per definire workout con sintassi leggibile
- ✅ **Gestione parametri personalizzati** (zone di ritmo, FC, toleranze)
- ✅ **Sincronizzazione con Garmin Connect** (upload, pianificazione, cancellazione)
- ✅ **Supporto ripetizioni annidate** (serie complesse)
- ✅ **Import/Export Excel** con formattazione automatica
- ✅ **Autenticazione MFA** per Garmin Connect

## 📋 Requisiti

```bash
pip install pandas openpyxl garth tkinter
```

- **Python 3.8+**
- **Garmin Connect account** (per sincronizzazione workout)

## 🚀 Avvio Rapido

```bash
python training_planner_gui.py
```

### Workflow Base

1. **Genera Excel di esempio** - Crea un template con workout pre-compilati
2. **Modifica workout** - Personalizza allenamenti usando la sintassi DSL
3. **Login Garmin** - Autentica con email/password (supporta MFA)
4. **Upload workout** - Carica gli allenamenti su Garmin Connect
5. **Pianifica** - Assegna workout a date specifiche nel calendario Garmin

## 📝 Sintassi DSL

### Struttura Base

```
tipo_step: durata @ target
```

### Tipi di Step

| Tipo       | Descrizione            | Esempio                    |
|------------|------------------------|----------------------------|
| `warmup`   | Riscaldamento          | `warmup: 10min @ Z2`       |
| `interval` | Intervallo intenso     | `interval: 1km @ Z4`       |
| `recovery` | Recupero attivo (corsa)| `recovery: 2min @ Z1`      |
| `rest`     | Riposo completo (fermo)| `rest: 90sec`              |
| `cooldown` | Defaticamento          | `cooldown: 5min @ Z1`      |

### Durate

| Formato      | Tipo     | Esempio  | Valore           |
|--------------|----------|----------|------------------|
| `XXmin`      | Minuti   | `10min`  | 10 minuti        |
| `XX'`        | Minuti   | `15'`    | 15 minuti        |
| `XXsec`      | Secondi  | `30sec`  | 30 secondi       |
| `XXs`        | Secondi  | `45s`    | 45 secondi       |
| `lap-button` | Manuale  | `lap-button` | Premi lap per continuare |

⚠️ **IMPORTANTE**: Per i minuti usa `min` o `'`, NON solo `m` (che è riservato ai metri)

### Distanze

| Formato | Tipo        | Esempio | Valore      |
|---------|-------------|---------|-------------|
| `XXm`   | Metri       | `400m`  | 400 metri   |
| `XXkm`  | Chilometri  | `5km`   | 5 chilometri|

### Target (Zone/Ritmi)

#### Zone Standard (da Parameters)
```
@ Z1    # Zona 1 (recupero)
@ Z2    # Zona 2 (aerobica)
@ Z3    # Zona 3 (tempo)
@ Z4    # Zona 4 (soglia)
@ Z5    # Zona 5 (VO2max)
```

#### Ritmi Personalizzati
```
@ marathon      # Ritmo custom definito in Parameters
@ threshold     # Soglia anaerobica
@ recovery      # Ritmo recupero
@ easy_range    # Range di ritmo facile
```

#### Ritmi Assoluti
```
@ 5:00          # Ritmo fisso 5:00 min/km (usa pace_tolerance)
@ 5:00-5:30     # Range di ritmo esplicito 5:00-5:30 min/km
```

#### Frequenza Cardiaca
```
@ HR_Z1         # Zona FC 1 (% di HR_max)
@ HR_Z2         # Zona FC 2 (% di HR_max)
@ 150-165       # Range FC assoluto in bpm
@ 80%           # 80% di HR_max (usa hr_tolerance)
@ 70-85%        # Range percentuale di HR_max
```

### Ripetizioni

#### Semplici
```
repeat 4:
  interval: 1km @ Z4
  recovery: 2min @ Z1
```

#### Annidate
```
repeat 2:
  repeat 6:
    interval: 1min @ threshold
    recovery: 1min @ recovery
  rest: 3min
```

⚠️ **Indentazione**: Usa **2 spazi** per ogni livello di ripetizione

## 📊 Foglio Parameters

Il foglio "Parameters" nell'Excel definisce zone personalizzate e parametri globali:

### Zone di Ritmo

```
Key         | Metric | Expression | Notes
------------|--------|------------|---------------------------
Z1          | pace   | 6:30       | Zona 1: Recupero attivo
Z2          | pace   | 6:00       | Zona 2: Aerobica
Z3          | pace   | 5:30       | Zona 3: Tempo
Z4          | pace   | 5:00       | Zona 4: Soglia
Z5          | pace   | 4:30       | Zona 5: VO2max
```

### Ritmi Personalizzati

#### Valori Singoli (usano tolerance)
```
Key         | Metric | Expression | Notes
------------|--------|------------|---------------------------
recovery    | pace   | 7:00       | Ritmo recupero (usa pace_tolerance)
marathon    | pace   | 5:20       | Ritmo maratona (usa pace_tolerance)
threshold   | pace   | 5:10       | Soglia anaerobica (usa pace_tolerance)
```

#### Range Fissi (ignorano tolerance)
```
Key           | Metric | Expression | Notes
--------------|--------|------------|---------------------------
easy_range    | pace   | 6:00-6:30  | Range fisso ritmo facile
tempo_range   | pace   | 5:20-5:40  | Range fisso ritmo tempo
fartlek       | pace   | 4:30-6:00  | Range ampio per fartlek
```

### Zone Frequenza Cardiaca

```
Key         | Metric | Expression | Notes
------------|--------|------------|---------------------------
HR_Z1       | hr     | 60-70%     | Zona FC 1 (60-70% di HR_max)
HR_Z2       | hr     | 70-80%     | Zona FC 2 (70-80% di HR_max)
HR_Z3       | hr     | 80-85%     | Zona FC 3 (80-85% di HR_max)
HR_Z4       | hr     | 85-90%     | Zona FC 4 (85-90% di HR_max)
HR_Z5       | hr     | 90-100%    | Zona FC 5 (90-100% di HR_max)
HR_max      | hr     | 185        | FC massima (PERSONALIZZA!)
```

### Toleranze

```
Key            | Metric | Expression | Notes
---------------|--------|------------|---------------------------
pace_tolerance | pace   | 5          | ±5 sec per ritmi singoli
hr_tolerance   | hr     | 5          | ±5 bpm per FC singoli
```

**Come funzionano le toleranze:**
- `marathon: 5:20` con `pace_tolerance: 5` → diventa range `5:15-5:25`
- `150bpm` con `hr_tolerance: 5` → diventa range `145-155`
- I range espliciti (`5:00-5:30`) ignorano le toleranze

## 📋 Esempi di Workout Completi

### Ripetute 4×5' @ Soglia
```
warmup: 10min @ Z2
repeat 4:
  interval: 5min @ Z4
  recovery: 2min @ Z1
cooldown: 5min @ Z1
```

### Lungo con Accelerazioni
```
warmup: 10min @ easy_range
interval: 40min @ Z2
repeat 4:
  interval: 100m @ Z5
  rest: 60sec
cooldown: 5min @ Z1
```

### Piramidale
```
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

### Serie Annidate
```
warmup: lap-button @ Z2
repeat 2:
  repeat 6:
    interval: 1min @ threshold
    recovery: 1min @ recovery
  rest: 3min
cooldown: 5min @ Z1
```

## 🔧 Funzionalità GUI

### Gestione Workout
- ✅ Visualizzazione tabellare di tutti i workout
- ✅ Editor inline per modifiche rapide
- ✅ Anteprima JSON del workout espanso
- ✅ Colorazione righe (bianco=non caricato, giallo=caricato, verde=pianificato)

### Garmin Connect
- ✅ **Login con credenziali** - Crea/aggiorna sessione (supporta MFA)
- ✅ **Login con sessione salvata** - Riutilizza sessione esistente
- ✅ **Upload workout** - Carica sulla libreria Garmin
- ✅ **Pianifica workout** - Assegna a data specifica
- ✅ **Rimuovi pianificazione** - Toglie dal calendario
- ✅ **Cancella workout** - Rimuove definitivamente da Garmin

### Opzioni
- ☑️ **Prefisso WnSn** - Aggiunge "W1S2 - " al nome workout (Week/Session)

## 🏗️ Architettura

```
training_planner_gui.py     # GUI principale (Tkinter)
├── dsl_parser.py            # Parser DSL → JSON Garmin
├── excel_utils.py           # Import/Export Excel + formattazione
├── garmin_service.py        # Wrapper autenticazione Garmin
└── garmin_client.py         # API client Garmin Connect
```

### Moduli Chiave

**dsl_parser.py**
- `expand_repeat_lines()` - Espande ripetizioni mantenendo struttura
- `parse_duration_part()` - Parsea durate/distanze (⚠️ fix disambiguazione m/min)
- `parse_target()` - Parsea target (ritmo, HR, potenza)
- `build_garmin_workout_from_excel_row()` - Genera JSON workout completo

**garmin_service.py**
- `login_with_credentials()` - Login con email/password + MFA
- `login_with_saved_session()` - Riprendi sessione salvata
- `save_workout()` - Crea workout su Garmin
- `schedule_workout()` - Pianifica workout su data
- `unschedule_workout()` - Rimuove pianificazione
- `delete_workout()` - Cancella workout definitivamente

## 🐛 Fix Recenti

### v1.1 - Disambiguazione Metri/Minuti (2025-11-26)

**Problema**: Il regex per i metri `r"^(\d+(?:\.\d+)?)\s*m(?:\s|$)"` poteva confondere "10m" con 10 metri invece di riconoscere "10min" come minuti.

**Soluzione**: Aggiunto negative lookahead `(?!i)` per impedire match su "min":

```python
# PRIMA (bug potenziale)
m = re.match(r"^(\d+(?:\.\d+)?)\s*m(?:\s|$)", s)

# DOPO (corretto)
m = re.match(r"^(\d+(?:\.\d+)?)\s*m(?!i)(?:\s|$)", s)
```

**Comportamento garantito:**
- ✅ `400m` → 400 metri (distanza)
- ✅ `10min` → 10 minuti (tempo = 600 sec)
- ✅ `15 min` → 15 minuti (tempo = 900 sec)

Dettagli: [FIX_PARSING_METRI_MINUTI.md](FIX_PARSING_METRI_MINUTI.md)

## 🔒 Sicurezza

- Le credenziali Garmin sono gestite tramite **garth** con OAuth2
- La sessione è salvata localmente in `./garminconnect/`
- Supporto nativo per **autenticazione a 2 fattori (MFA)**
- Nessuna password salvata in chiaro

## 📚 Documentazione Aggiuntiva

- `FIX_PARSING_METRI_MINUTI.md` - Dettagli fix disambiguazione m/min
- Foglio "Esempi DSL" nell'Excel - Guida completa alla sintassi

## 🤝 Contributi

Progetto sviluppato da Francesco con assistenza di Claude (Anthropic).

## 📄 Licenza

Uso personale - © 2025 Francesco

---

## ⚠️ Note Importanti

1. **Indentazione ripetizioni**: Usa sempre **2 spazi** per livello
2. **Minuti**: Specifica con `min` o `'`, NON con solo `m`
3. **Metri**: Specifica con `m` o `km`
4. **HR_max**: Personalizza in Parameters per zone FC corrette
5. **Toleranze**: Si applicano solo a valori singoli, non ai range
6. **Excel aperto**: Chiudi il file Excel prima di salvare dalla GUI

## 🆘 Troubleshooting

**Login Garmin fallisce**
- Verifica email/password corrette
- Se hai MFA attivo, inserisci il codice quando richiesto
- Prova "Login (crea/aggiorna sessione)" invece di "Usa sessione salvata"

**Workout non viene caricato**
- Controlla la sintassi DSL nel campo Steps
- Verifica che i parametri usati esistano in "Parameters"
- Guarda la console per messaggi di errore dettagliati

**Errore 403 nella dis-pianificazione**
- Garmin limita alcune API: rimuovi manualmente da Garmin Connect web
- Alternativa: cancella l'intero workout (rimuove automaticamente la pianificazione)

**Excel non si salva**
- Chiudi il file Excel se è aperto in un altro programma
- Verifica permessi scrittura sulla cartella

---

**Ultima modifica**: 2025-11-26  
**Versione**: 1.1
