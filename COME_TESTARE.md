# v3.2 – Cosa è cambiato e come provarlo

## 1. Prima di tutto: sicurezza (da fare tu)

I file `garminconnect/oauth1_token.json` e `oauth2_token.json` sono pubblici su GitHub
e permettono di accedere al tuo account Garmin.

1. **Cambia la password Garmin** (così i token pubblicati smettono di funzionare).
2. **Togli i file dalla cronologia git** (non basta cancellarli con un commit):
   ```bash
   pip install git-filter-repo
   cd garminPlannerCGPT        # il tuo clone
   git filter-repo --invert-paths \
     --path garminconnect \
     --path garmin_workouts_2025-2026_running.json \
     --path garmin_workouts_2025-2026_running.xlsx \
     --path AllenamentoCorridaFrank.xlsx --path test.xlsx \
     --path '~$garmin_workouts_2025-2026_running.xlsx' --path .DS_Store \
     --path __pycache__ --path download_dialog_backup.py
   git remote add origin https://github.com/yesnoj/garminPlannerCGPT.git   # filter-repo rimuove il remote
   git push --force origin main
   ```
   Tieni una copia dei tuoi Excel personali fuori dal repo prima di farlo.

## 2. Installare questa versione

Copia il contenuto dello zip nella cartella del progetto (sovrascrivendo), poi:

```bash
python -m pip install -r requirements.txt   # su Windows: py -m pip install -r requirements.txt
python garmin_planner.py
```

## 3. Checklist di prova (10 minuti)

| # | Cosa fare | Risultato atteso |
|---|---|---|
| 1 | Apri il piano della Mezza | lista divisa in 23 settimane, durata e km per settimana, nessuna riga rossa |
| 2 | Clicca un allenamento di ripetute | a destra il profilo (blocchi rossi alti = ripetute) e gli step colorati |
| 3 | Doppio clic su uno step, cambia il target in `Z9` | il dialog segnala l'errore e non fa confermare |
| 4 | "Modelli rapidi" → "Allunghi 6×20\"", poi **Salva allenamento** | il file Excel si aggiorna; i fogli Impostazioni/Esempi DSL restano |
| 5 | Scheda "Testo DSL": scrivi una riga sbagliata | riga evidenziata in rosso e messaggio sotto al titolo |
| 6 | Pulsante ☾/☀ in alto a destra | passa da tema chiaro a scuro |
| 7 | Accedi… → Usa sessione salvata | "● Connesso a Garmin" in verde |
| 8 | Seleziona **un** allenamento futuro → Carica e pianifica | stato "✓ Pianificato", visibile nel calendario Garmin |
| 9 | Stesso allenamento → Elimina da Garmin | stato torna "○ Da caricare" |
| 10 | `python -m pytest -q` | tutti i test passano |
| 11 | Tema scuro: apri il calendario della data, "Modelli rapidi", la tendina Sport | tutto scuro |
| 12 | Nell'editor trascina il Defaticamento al centro della riga "Ripeti" | lo step entra nella ripetizione (bordo evidenziato mentre trascini) |
| 13 | Nella lista trascina un allenamento su un'altra riga della stessa settimana | cambia solo l'ordine |
| 14 | Trascinalo sull'intestazione di un'altra settimana | passa a quella settimana, data spostata di conseguenza |
| 15 | Tasto destro su "Settimana 3" → Duplica nella settimana successiva | le 3 sedute copiate nella settimana 4, 7 giorni dopo |

Le vecchie interfacce restano disponibili in `legacy/` se ti servisse tornare indietro.

## 4. Cosa è cambiato nella v2.1 (sicurezza e parser)

- **Parser**: ogni errore è segnalato con riga Excel e riga DSL; prima diventava in silenzio
  uno step di 60 secondi o senza target. Per input validi il JSON inviato a Garmin è identico
  a prima (verificato su 287 workout), tranne gli step `rest` e i target Swim_/Cadence_ che
  prima venivano persi.
- **Nuove durate**: `1.5min`, `90s`, `1h`, `1:30`, `40′`. Attenzione: `10m` = 10 metri.
- **Indentazione** sotto `repeat` libera (2/4 spazi o tab); prima con 4 spazi gli step sparivano.
- **Upload**: validazione completa prima di contattare Garmin; ID salvati nell'Excel anche se
  il caricamento si interrompe a metà (niente doppioni).
- **Sessione Garmin** in `~/.garminplanner/tokens` con permessi 600, salvata anche dopo il
  rinnovo automatico; password cancellata dal campo dopo il login.
- **Log** dei JSON Garmin solo con `GARMINPLANNER_DEBUG=1`.
- **Repo**: rimossi token, dati personali, file di sistema e i due moduli duplicati mai usati
  (`dsl_parser_multisport.py`, `excel_utils_multisport.py`); aggiunti `.gitignore`,
  `requirements.txt`, `LICENSE` (MIT), `examples/` anonimi e `tests/`.
- Caratteri accentati "rotti" (Ã¨ → è) corretti nei messaggi.

## 5. Novità della v3.0 (interfaccia)

- Un'unica finestra al posto di Classic/Advanced/Multisport, tema chiaro/scuro.
- Lista per settimane con stato colorato, ricerca e filtri; selezionando una settimana
  intera le operazioni Garmin valgono per tutti i suoi allenamenti.
- Editor integrato con profilo grafico, dialog guidato per gli step (anteprima e controllo
  in tempo reale), modelli rapidi, spostamenti dentro/fuori dalle ripetizioni, testo DSL
  con colori.
- Panoramica con i km per settimana.
- Il salvataggio non cancella piu' gli altri fogli dell'Excel. Nota: se il file aveva le date
  calcolate da formule (foglio Impostazioni), dopo un salvataggio dall'app le date diventano
  valori fissi.
- Se modifichi un allenamento gia' su Garmin, l'app ti ricorda che va eliminato e ricaricato
  (Garmin non aggiorna i workout da solo).
