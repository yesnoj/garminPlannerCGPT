# 🏃‍♂️ Garmin Training Planner

[![Python](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)]()

**Un'applicazione desktop completa per pianificare, creare e sincronizzare automaticamente i tuoi allenamenti di corsa con Garmin Connect.**

Gestisci i tuoi workout in Excel con un DSL (Domain Specific Language) intuitivo, e caricali automaticamente sul tuo account Garmin Connect con un click. Perfetto per runner che vogliono programmare settimane o mesi di allenamenti in modo organizzato e professionale.

![Training Planner Screenshot](https://via.placeholder.com/800x500?text=Training+Planner+GUI)

---

## 📋 Indice

- [Caratteristiche](#-caratteristiche)
- [Installazione](#-installazione)
- [Guida Rapida](#-guida-rapida)
- [DSL: Linguaggio di Definizione Workout](#-dsl-linguaggio-di-definizione-workout)
- [Parametri e Zone](#-parametri-e-zone)
- [Workflow Completo](#-workflow-completo)
- [Esempi Pratici](#-esempi-pratici)
- [FAQ](#-faq)
- [Troubleshooting](#-troubleshooting)
- [Contribuire](#-contribuire)
- [Licenza](#-licenza)

---

## ✨ Caratteristiche

### 🎯 Core Features

- **📊 Gestione Excel Intuitiva**: Crea e modifica workout in Excel con sintassi semplice
- **🔄 Sincronizzazione Garmin**: Carica automaticamente workout su Garmin Connect
- **📅 Pianificazione Automatica**: Schedula workout su date specifiche
- **🏷️ Zone Personalizzabili**: Definisci le tue zone di ritmo e frequenza cardiaca
- **🔁 Ripetizioni Annidate**: Supporto per allenamenti complessi (serie di ripetute)
- **💾 Autosalvataggio**: Salva automaticamente le modifiche su Excel
- **🔐 Sessioni Persistenti**: Login a Garmin con sessione salvata (no credenziali ripetute)
- **⚡ Loading Dinamico**: Indicatori di progresso per operazioni lunghe

### 🎨 UI Features

- **Codifica Colori**: 
  - 🟢 Verde = Workout caricato E pianificato
  - 🟡 Giallo = Workout caricato ma NON pianificato
  - ⚪ Bianco = Workout NON caricato
- **Editor Integrato**: Modifica workout direttamente nell'app
- **Visualizzazione JSON**: Vedi la struttura JSON prima di caricare
- **Multi-selezione**: Gestisci più workout contemporaneamente

---

## 🚀 Installazione

### Prerequisiti

- Python 3.7 o superiore
- pip (gestore pacchetti Python)
- Account Garmin Connect

### Step 1: Clona il Repository

```bash
git clone https://github.com/tuousername/garmin-training-planner.git
cd garmin-training-planner
```

### Step 2: Installa le Dipendenze

```bash
pip install -r requirements.txt
```

**requirements.txt**:
```
pandas>=1.3.0
openpyxl>=3.0.0
garth>=0.4.0
tkinter  # Solitamente incluso in Python
```

### Step 3: Avvia l'Applicazione

```bash
python training_planner_gui.py
```

---

## 🎯 Guida Rapida

### 1️⃣ Genera un Excel di Esempio

1. Avvia l'applicazione
2. Clicca su **"Genera Excel di esempio"**
3. Salva il file (es. `mio_piano.xlsx`)
4. Apri il file per vedere la struttura e gli esempi

### 2️⃣ Configura le Tue Zone

Apri il foglio **"Parameters"** e personalizza:

```
Key          | Metric | Expression | Notes
-------------|--------|------------|---------------------------
Z1           | pace   | 6:30       | Zona 1: Recupero
Z2           | pace   | 6:00       | Zona 2: Endurance
Z3           | pace   | 5:30       | Zona 3: Tempo
Z4           | pace   | 5:00       | Zona 4: Soglia
Z5           | pace   | 4:30       | Zona 5: VO2max
HR_max       | hr     | 185        | Frequenza cardiaca massima
```

### 3️⃣ Crea i Tuoi Workout

Nel foglio **"Workouts"**, aggiungi le tue sessioni:

| Week | Date       | Session | Sport   | Description           | Steps                    |
|------|------------|---------|---------|----------------------|--------------------------|
| 1    | 2025-12-01 | 1       | Running | Ripetute 5×1km @ Z4  | warmup: 15min @ Z2<br>repeat 5:<br>  interval: 1km @ Z4<br>  recovery: 2min @ Z1<br>cooldown: 10min @ Z1 |

### 4️⃣ Connettiti a Garmin

1. Inserisci email e password Garmin
2. Clicca **"Login (crea/aggiorna sessione)"**
3. Se richiesto, inserisci il codice MFA via email
4. ✅ Status diventa verde "Connesso"

### 5️⃣ Carica e Pianifica

1. Seleziona uno o più workout (Ctrl+Click per multi-selezione)
2. Clicca **"Carica + pianifica selezionati"**
3. Attendi il completamento (vedrai un loading dinamico)
4. ✅ I workout sono ora su Garmin Connect!

---

## 📝 DSL: Linguaggio di Definizione Workout

Il DSL (Domain Specific Language) è una sintassi semplice per definire allenamenti complessi.

### Struttura Base

```
tipo: durata @ target
```

- **tipo**: warmup, interval, recovery, rest, cooldown
- **durata**: tempo (10min, 30sec) o distanza (1km, 400m) o lap-button
- **target**: zona (Z1-Z5), ritmo (5:00), HR (150bpm, HR_Z3), parametro custom

### Tipi di Step

| Tipo      | Descrizione                  | Esempio                   |
|-----------|------------------------------|---------------------------|
| `warmup`  | Riscaldamento                | `warmup: 15min @ Z2`      |
| `interval`| Intervallo intenso           | `interval: 5min @ Z4`     |
| `recovery`| Recupero attivo (corri lento)| `recovery: 2min @ Z1`     |
| `rest`    | Riposo completo (FERMO)      | `rest: 90sec`             |
| `cooldown`| Defaticamento                | `cooldown: 10min @ Z1`    |

### Durate

| Formato       | Descrizione           | Esempio              |
|---------------|-----------------------|----------------------|
| `Xmin`        | Minuti                | `10min`, `45min`     |
| `Xm`, `X'`    | Minuti (alternativo)  | `10m`, `45'`         |
| `Xsec`, `Xs`  | Secondi               | `30sec`, `90s`       |
| `Xkm`         | Chilometri            | `5km`, `10km`        |
| `Xm`          | Metri                 | `400m`, `1000m`      |
| `lap-button`  | Premi LAP per avanzare| `lap-button @ Z2`    |

### Target di Ritmo

| Formato         | Descrizione                    | Esempio                |
|-----------------|--------------------------------|------------------------|
| `Z1` - `Z5`     | Zone predefinite (Parameters)  | `@ Z4`                 |
| `min:sec`       | Ritmo fisso (usa tolerance)    | `@ 5:00`               |
| `min:sec-min:sec`| Range ritmo esplicito         | `@ 5:00-5:30`          |
| `parametro`     | Ritmo custom (Parameters)      | `@ marathon`, `@ easy_range` |

### Target di Frequenza Cardiaca

| Formato         | Descrizione                    | Esempio                |
|-----------------|--------------------------------|------------------------|
| `HR_Z1` - `HR_Z5`| Zone FC (% di HR_max)        | `@ HR_Z3`              |
| `Xbpm`, `X`     | FC assoluta (usa tolerance)    | `@ 150`, `@ 160bpm`    |
| `X-Ybpm`, `X-Y` | Range FC esplicito             | `@ 140-160`, `@ 150-165bpm` |

### Target Altri

| Formato       | Descrizione          | Esempio              |
|---------------|----------------------|----------------------|
| `Xrpm`        | Cadenza              | `@ 85rpm`            |
| `X-Yrpm`      | Range cadenza        | `@ 80-90rpm`         |
| `XW`          | Potenza (Watt)       | `@ 200W`             |
| `X-YW`        | Range potenza        | `@ 180-220W`         |
| `open`        | Nessun target        | `@ open`             |

### Ripetizioni

#### Sintassi Base

```
repeat N:
  step1
  step2
  ...
```

⚠️ **IMPORTANTE**: Usa **2 spazi** per l'indentazione, non tab!

#### Esempio Semplice

```
warmup: 10min @ Z2
repeat 4:
  interval: 1km @ Z4
  recovery: 2min @ Z1
cooldown: 5min @ Z1
```

Questo crea: riscaldamento → (1km veloce + 2min recupero) × 4 → defaticamento

#### Ripetizioni Annidate

```
warmup: 15min @ Z2
repeat 2:
  repeat 6:
    interval: 1min @ threshold
    recovery: 1min @ Z1
  rest: 3min
cooldown: 10min @ Z1
```

Questo crea: 2 serie da 6 ripetute di 1min, con 3min di riposo tra le serie.

---

## ⚙️ Parametri e Zone

### Foglio "Parameters"

Il foglio Parameters contiene tutte le tue zone personalizzate e configurazioni.

#### Zone di Ritmo (Pace)

```
Key  | Metric | Expression | Notes
-----|--------|------------|---------------------------
Z1   | pace   | 6:30       | Zona 1: Recupero attivo
Z2   | pace   | 6:00       | Zona 2: Endurance/Aerobica
Z3   | pace   | 5:30       | Zona 3: Tempo/Ritmo gara
Z4   | pace   | 5:00       | Zona 4: Soglia anaerobica
Z5   | pace   | 4:30       | Zona 5: VO2max/Ripetute
```

#### Zone di Frequenza Cardiaca

```
Key     | Metric | Expression | Notes
--------|--------|------------|---------------------------
HR_max  | hr     | 185        | FC massima (PERSONALIZZA!)
HR_Z1   | hr     | 60-70%     | Zona FC 1: 60-70% di HR_max
HR_Z2   | hr     | 70-80%     | Zona FC 2: 70-80% di HR_max
HR_Z3   | hr     | 80-85%     | Zona FC 3: 80-85% di HR_max
HR_Z4   | hr     | 85-90%     | Zona FC 4: 85-90% di HR_max
HR_Z5   | hr     | 90-100%    | Zona FC 5: 90-100% di HR_max
```

💡 **Le zone HR sono calcolate automaticamente in percentuale di HR_max!**

#### Ritmi Personalizzati

##### Singoli (usano tolerance)

```
Key       | Metric | Expression | Notes
----------|--------|------------|---------------------------
recovery  | pace   | 7:00       | Ritmo recupero lento
marathon  | pace   | 5:20       | Ritmo maratona
threshold | pace   | 5:10       | Ritmo soglia anaerobica
interval  | pace   | 4:20       | Ritmo intervalli veloci
```

Quando usi `@ marathon`, diventa automaticamente `5:15-5:25` (se pace_tolerance = 5 sec).

##### Range (ignorano tolerance)

```
Key            | Metric | Expression | Notes
---------------|--------|------------|---------------------------
easy_range     | pace   | 6:00-6:30  | Range ritmo facile (fisso)
tempo_range    | pace   | 5:20-5:40  | Range ritmo tempo (fisso)
threshold_range| pace   | 4:50-5:10  | Range ritmo soglia (fisso)
fartlek        | pace   | 4:30-6:00  | Range ampio per fartlek
```

Quando usi `@ easy_range`, usa esattamente `6:00-6:30`, senza tolleranza aggiuntiva.

#### Tolleranze

```
Key            | Metric | Expression | Notes
---------------|--------|------------|---------------------------
pace_tolerance | pace   | 5          | Tolleranza ±sec per ritmi singoli
hr_tolerance   | hr     | 5          | Tolleranza ±bpm per FC singole
```

**Come funzionano le tolleranze:**

- **Ritmo singolo**: `@ 5:00` con `pace_tolerance = 5` → diventa `4:55-5:05`
- **Ritmo range**: `@ 5:00-5:30` → resta `5:00-5:30` (tolleranza ignorata)
- **FC singola**: `@ 150` con `hr_tolerance = 5` → diventa `145-155`
- **FC range**: `@ 140-160` → resta `140-160` (tolleranza ignorata)

---

## 🔄 Workflow Completo

### Scenario: Piano Allenamento 8 Settimane per Maratona

#### Step 1: Crea il Piano Base

```excel
Week | Date       | Session | Sport   | Description
-----|------------|---------|---------|------------------
1    | 2025-01-06 | 1       | Running | Lungo base
1    | 2025-01-08 | 2       | Running | Ripetute brevi
1    | 2025-01-10 | 3       | Running | Tempo run
2    | 2025-01-13 | 1       | Running | Lungo progressivo
2    | 2025-01-15 | 2       | Running | Intervalli 1km
...  | ...        | ...     | ...     | ...
```

#### Step 2: Definisci gli Allenamenti

**Workout Settimana 1, Sessione 1** - Lungo base:
```
warmup: 10min @ easy_range
interval: 90min @ Z2
cooldown: 5min @ Z1
```

**Workout Settimana 1, Sessione 2** - Ripetute brevi:
```
warmup: 15min @ Z2
repeat 8:
  interval: 400m @ Z5
  rest: 90sec
cooldown: 10min @ Z1
```

**Workout Settimana 1, Sessione 3** - Tempo run:
```
warmup: 10min @ Z2
interval: 20min @ threshold
interval: 5min @ Z2
interval: 20min @ threshold
cooldown: 10min @ Z1
```

#### Step 3: Carica su Garmin

1. Seleziona tutti i workout della settimana 1 (Shift+Click)
2. Clicca **"Carica + pianifica selezionati"**
3. Attendi il completamento
4. Verifica su Garmin Connect app o web

#### Step 4: Gestione Modifiche

**Se vuoi modificare un workout già caricato:**

1. Modifica gli Steps in Excel
2. Seleziona il workout
3. Clicca **"Cancella da Garmin selezionati"**
4. Clicca **"Carica + pianifica selezionati"**

**Se vuoi solo spostare la data:**

1. Cambia la colonna Date in Excel
2. Seleziona il workout
3. Clicca **"Rimuovi pianificazione selezionati"**
4. Clicca **"Carica + pianifica selezionati"**

---

## 📚 Esempi Pratici

### Esempio 1: Progressivo Semplice

**Obiettivo**: Inizia piano, aumenta gradualmente il ritmo.

```
interval: 10min @ Z1
interval: 15min @ Z2
interval: 10min @ Z3
interval: 5min @ Z4
```

**Garmin riceverà**: 4 step consecutivi con ritmi crescenti.

---

### Esempio 2: Fartlek Strutturato

**Obiettivo**: Alternare ritmi veloci e lenti in modo ripetitivo.

```
warmup: 10min @ Z2
repeat 6:
  interval: 3min @ Z4
  recovery: 2min @ Z1
cooldown: 10min @ Z1
```

**Garmin riceverà**: Riscaldamento → (3min veloce + 2min lento) × 6 → defaticamento.

---

### Esempio 3: Piramidale Ascendente

**Obiettivo**: Aumentare progressivamente la distanza.

```
warmup: 15min @ Z2
interval: 400m @ Z5
rest: 90sec
interval: 800m @ Z4
rest: 2min
interval: 1200m @ Z4
rest: 3min
interval: 1600m @ Z3
rest: 3min
cooldown: 10min @ Z1
```

**Garmin riceverà**: Ogni step è separato con riposo completo (rest = fermo).

---

### Esempio 4: Serie di Ripetute con Riposo Lungo

**Obiettivo**: 2 serie da 5 ripetute veloci, con lungo riposo tra le serie.

```
warmup: 15min @ Z2
repeat 2:
  repeat 5:
    interval: 800m @ Z5
    recovery: 90sec @ Z1
  rest: 5min
cooldown: 10min @ Z1
```

**Garmin riceverà**: Struttura annidata con RepeatGroupDTO.

---

### Esempio 5: Allenamento con Frequenza Cardiaca

**Obiettivo**: Controllare l'intensità tramite FC invece che ritmo.

```
interval: 10min @ HR_Z1
interval: 20min @ HR_Z2
interval: 15min @ HR_Z3
interval: 10min @ HR_Z4
interval: 5min @ HR_Z2
```

**Garmin riceverà**: Target in bpm calcolati automaticamente da HR_max.

---

### Esempio 6: Lungo con Inserti Veloci

**Obiettivo**: Corsa lunga con accelerazioni finali.

```
warmup: 10min @ easy_range
interval: 70min @ Z2
repeat 4:
  interval: 200m @ Z5
  rest: 60sec
cooldown: 10min @ Z1
```

**Garmin riceverà**: Lungo aerobico + serie di sprint brevi.

---

### Esempio 7: Threshold Run con Gestione LAP

**Obiettivo**: Riscaldamento variabile premendo LAP, poi soglia e defaticamento.

```
warmup: lap-button @ Z2
interval: 30min @ threshold
cooldown: lap-button @ Z1
```

**Garmin riceverà**: I primi e l'ultimo step richiedono pressione del tasto LAP.

---

### Esempio 8: Allenamento Multi-Zona Complesso

**Obiettivo**: Combinare zone, parametri custom e ripetizioni.

```
warmup: 10min @ easy_range
interval: 20min @ marathon
repeat 3:
  interval: 5min @ threshold
  recovery: 3min @ recovery
interval: 10min @ tempo_range
cooldown: 5min @ Z1
```

**Garmin riceverà**: Mix di zone standard, ritmi custom e range personalizzati.

---

## ❓ FAQ

### Q: Posso usare l'app per ciclismo o nuoto?

**A**: Sì! Cambia la colonna `Sport` in:
- `Running` per corsa
- `Cycling` per bici
- `Swimming` per nuoto

Il DSL funziona allo stesso modo, ma cambia i target (es. potenza per bici, ritmo per nuoto).

---

### Q: Cosa succede se modifico un workout già caricato?

**A**: Garmin non aggiorna automaticamente. Devi:
1. Cancellare il workout da Garmin
2. Ricaricarlo con le modifiche

Oppure crea un nuovo workout con nome diverso.

---

### Q: Posso caricare workout senza pianificarli?

**A**: Sì! Usa **"Carica workout selezionati su Garmin"** invece di "Carica + pianifica". I workout andranno nella libreria ma senza data.

---

### Q: Come funziona il prefisso "W1S2 - "?

**A**: Se abiliti la checkbox "Prefisso nome WnSn", il nome del workout diventa:
```
W1S2 - Ripetute 4×5' @ Z4
```
Dove `W1` = Week 1, `S2` = Session 2.

Utile per riconoscere i workout in Garmin Connect!

---

### Q: Cosa significa "rest" vs "recovery"?

**A**:
- **rest**: Riposo completo (ti fermi, non corri) → utile per recuperi lunghi
- **recovery**: Recupero attivo (corri lento) → utile tra ripetute

In Garmin, `rest` è uno step di tipo "rest", `recovery` è uno step normale con ritmo lento.

---

### Q: Posso usare "m" per minuti invece di metri?

**A**: ⚠️ **ATTENZIONE**: "m" può essere ambiguo!
- `10m` → metri (per evitare confusione con minuti)
- `10min` → minuti (raccomandato)
- `10'` → minuti (alternativo)

**Best practice**: Usa sempre `min` per minuti e `m` per metri.

---

### Q: Le zone HR devono essere percentuali?

**A**: No! Puoi anche usare valori assoluti:

**Percentuali** (raccomandato):
```
HR_Z3 | hr | 80-85%
```

**Assoluti**:
```
HR_Z3 | hr | 148-157
```

Ma le percentuali si ricalcolano automaticamente se cambi HR_max!

---

### Q: Posso caricare workout per altre persone?

**A**: Sì, basta fare login con le loro credenziali Garmin. Ma ricorda di salvare sessioni separate per ogni account.

---

### Q: Il file Excel deve avere colonne specifiche?

**A**: Sì, colonne obbligatorie nel foglio "Workouts":
- `Week`, `Date`, `Session`, `Sport`, `Description`, `Steps`

Colonne opzionali (gestite automaticamente):
- `WorkoutId`, `WorkoutScheduleId`, `ScheduledDate`

---

### Q: Posso usare l'app offline?

**A**: Sì per:
- Creare/modificare workout in Excel
- Generare file di esempio
- Modificare parametri

No per:
- Caricare workout su Garmin
- Pianificare workout
- Cancellare workout

---

## 🛠️ Troubleshooting

### Problema: "Errore nel login a Garmin"

**Causa**: Credenziali errate, MFA non completato, o problema di rete.

**Soluzione**:
1. Verifica email e password
2. Se hai MFA attivo, inserisci il codice ricevuto via email
3. Controlla la connessione internet
4. Prova a fare logout e re-login

---

### Problema: "403 Forbidden" quando rimuovo pianificazione

**Causa**: L'API Garmin non permette la rimozione via API in alcuni casi.

**Soluzione**:
Rimuovi manualmente la pianificazione da Garmin Connect (app o web), poi cancella il workout dall'app.

---

### Problema: Workout caricato ma non compare su Garmin Connect

**Causa**: Cache o sincronizzazione ritardata.

**Soluzione**:
1. Aspetta 1-2 minuti
2. Ricarica la pagina Garmin Connect
3. Forza la sincronizzazione sul dispositivo Garmin

---

### Problema: "Data non valida per workout"

**Causa**: Formato data errato in Excel.

**Soluzione**:
Usa formato `YYYY-MM-DD` o `DD/MM/YYYY` (Excel lo converte automaticamente). Esempio: `2025-12-25` o `25/12/2025`.

---

### Problema: DSL non riconosciuto

**Causa**: Sintassi errata o indentazione sbagliata.

**Soluzione**:
1. Verifica la sintassi base: `tipo: durata @ target`
2. Controlla l'indentazione: usa 2 spazi, non tab
3. Clicca "Mostra JSON steps espanso" per vedere gli errori

---

### Problema: Zone personalizzate non funzionano

**Causa**: Chiave non definita in Parameters o Expression errata.

**Soluzione**:
1. Verifica che la Key esista nel foglio Parameters
2. Controlla che Expression sia nel formato corretto (es. `5:00` per ritmo, `150` per HR)
3. Ricarica il file Excel nell'app

---

### Problema: "WorkoutId non valido" quando cancello

**Causa**: Workout non è stato caricato su Garmin (colonna WorkoutId vuota).

**Soluzione**:
Seleziona solo workout con WorkoutId compilato (colore giallo o verde nella GUI).

---

### Problema: Loading si blocca

**Causa**: Operazione molto lunga o problema di rete.

**Soluzione**:
1. Aspetta almeno 5-10 minuti per operazioni con molti workout
2. Se ancora bloccato, chiudi e riapri l'app
3. Riprova con meno workout selezionati

---

### Problema: "Sessione scaduta"

**Causa**: Token Garmin scaduto (scade dopo ~24 ore).

**Soluzione**:
Clicca di nuovo "Login (crea/aggiorna sessione)" per rinnovare il token.

---

## 💡 Tips & Tricks

### Tip 1: Usa Prefissi per Organizzare

Abilita "Prefisso nome WnSn" per avere workout denominati:
```
W1S1 - Lungo base
W1S2 - Ripetute brevi
W2S1 - Tempo run
```

Più facili da trovare in Garmin Connect!

---

### Tip 2: Crea Template Workout

Salva workout frequenti come template:

```excel
Week | Date | Session | Sport   | Description        | Steps
-----|------|---------|---------|-------------------|-------
TMPL | -    | -       | Running | Template Ripetute | warmup: 15min @ Z2
                                                       repeat 6:
                                                         interval: 1km @ Z4
                                                         recovery: 2min @ Z1
                                                       cooldown: 10min @ Z1
```

Poi copia e modifica per ogni settimana.

---

### Tip 3: Usa Parametri per Aggiornamenti Rapidi

Invece di cambiare ogni workout, aggiorna solo i parametri:

```
Z4 | pace | 5:00  →  Z4 | pace | 4:55
```

Poi ricarca i workout. Le zone si aggiornano automaticamente!

---

### Tip 4: Backup Regolari

L'app salva su Excel, ma fai backup regolari del file:
```
mio_piano_backup_2025-01-15.xlsx
mio_piano_backup_2025-02-01.xlsx
```

---

### Tip 5: Multi-selezione con Pattern

Per selezionare tutti i workout di una settimana:
1. Clicca il primo workout della settimana
2. Shift+Click l'ultimo workout della settimana
3. Tutti i workout intermedi sono selezionati!

---

### Tip 6: Verifica JSON Prima di Caricare

Clicca "Mostra JSON steps espanso" per vedere la struttura esatta che andrà su Garmin. Utile per debug!

---

### Tip 7: Usa ScheduledDate per Tracking

La colonna `ScheduledDate` si compila automaticamente. Usa Excel per filtrare:
```
=FILTER(Workouts, Workouts[ScheduledDate]<>"")
```

Vedi tutti i workout già pianificati!

---

## 🤝 Contribuire

Contributi, issue e feature request sono benvenuti!

### Come Contribuire

1. Fai fork del progetto
2. Crea un branch per la tua feature (`git checkout -b feature/AmazingFeature`)
3. Committa le modifiche (`git commit -m 'Add some AmazingFeature'`)
4. Pusha il branch (`git push origin feature/AmazingFeature`)
5. Apri una Pull Request

### Aree di Miglioramento

- [ ] Supporto per altri sport (triathlon, nuoto in piscina)
- [ ] Import da altri formati (Strava, TrainingPeaks)
- [ ] Grafici di visualizzazione piano allenamento
- [ ] Export in PDF del piano
- [ ] Notifiche desktop per workout pianificati
- [ ] Sincronizzazione bidirezionale (import da Garmin)

---

## 📄 Licenza

Questo progetto è distribuito sotto licenza MIT. Vedi il file `LICENSE` per maggiori dettagli.

---

## 🙏 Ringraziamenti

- **garth**: Libreria Python per l'integrazione con Garmin Connect
- **pandas & openpyxl**: Per la gestione Excel
- **Comunità Garmin**: Per documentazione e supporto API

---

## 📧 Contatti

**Autore**: [Il Tuo Nome]  
**Email**: tua.email@example.com  
**GitHub**: [@tuousername](https://github.com/tuousername)

**Link Progetto**: [https://github.com/tuousername/garmin-training-planner](https://github.com/tuousername/garmin-training-planner)

---

## 🚀 Versioni Future

### v2.0 (Pianificate)

- 🌐 **Web App**: Versione browser-based
- 📱 **Mobile App**: iOS e Android
- 📊 **Analytics**: Statistiche allenamenti completati
- 🔔 **Notifiche**: Promemoria workout giornalieri
- 🌍 **Multi-lingua**: Inglese, Italiano, Spagnolo, Francese

### v1.1 (In Sviluppo)

- ✅ Loading dinamico (COMPLETATO)
- 🔄 Auto-sync periodica
- 📝 Log dettagliati operazioni
- 🎨 Temi UI (light/dark)

---

<div align="center">

**⭐ Se ti piace questo progetto, lascia una stella su GitHub! ⭐**

Made with ❤️ by runners, for runners

</div>
