import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment
from copy import copy
from datetime import date


def format_workbook_dates_and_steps(path: str):
    """
    Applica formattazione al file Excel:
    - colonna Date e ScheduledDate come data DD/MM/YYYY
    - wrap text per Steps e Description
    - altezza riga in base al numero di righe di Steps
    - larghezze colonne per Workouts, Parameters e Esempi DSL
    - colonna Expression del foglio Parameters forzata a TESTO
    """
    wb = load_workbook(path)

    # ----- Sheet Workouts -----
    if "Workouts" in wb.sheetnames:
        ws = wb["Workouts"]

        # Trova colonne
        date_col_idx = sched_col_idx = desc_col_idx = steps_col_idx = None
        for cell in ws[1]:
            header = str(cell.value).strip().lower() if cell.value is not None else ""
            if header == "date":
                date_col_idx = cell.column
            elif header == "scheduleddate":
                sched_col_idx = cell.column
            elif header == "description":
                desc_col_idx = cell.column
            elif header == "steps":
                steps_col_idx = cell.column

        # Formato data DD/MM/YYYY
        if date_col_idx is not None:
            date_letter = get_column_letter(date_col_idx)
            for row_idx in range(2, ws.max_row + 1):
                cell = ws[f"{date_letter}{row_idx}"]
                if cell.value is not None:
                    cell.number_format = "DD/MM/YYYY"

        if sched_col_idx is not None:
            sched_letter = get_column_letter(sched_col_idx)
            for row_idx in range(2, ws.max_row + 1):
                cell = ws[f"{sched_letter}{row_idx}"]
                if cell.value is not None:
                    cell.number_format = "DD/MM/YYYY"

        # Larghezze colonne
        if desc_col_idx is not None:
            ws.column_dimensions[get_column_letter(desc_col_idx)].width = 23.22
        if steps_col_idx is not None:
            ws.column_dimensions[get_column_letter(steps_col_idx)].width = 82.89

        # Wrap text + altezza riga per Steps
        if steps_col_idx is not None:
            steps_letter = get_column_letter(steps_col_idx)
            for row_idx in range(2, ws.max_row + 1):
                cell = ws[f"{steps_letter}{row_idx}"]
                val = cell.value
                if val is None:
                    continue
                text = str(val)
                alignment = copy(cell.alignment)
                alignment.wrap_text = True
                cell.alignment = alignment
                num_lines = text.count("\n") + 1
                ws.row_dimensions[row_idx].height = 14.4 * num_lines

        # Wrap text per Description
        if desc_col_idx is not None:
            desc_letter = get_column_letter(desc_col_idx)
            for row_idx in range(2, ws.max_row + 1):
                cell = ws[f"{desc_letter}{row_idx}"]
                if cell.value is not None:
                    alignment = copy(cell.alignment)
                    alignment.wrap_text = True
                    cell.alignment = alignment

    # ----- Sheet Parameters -----
    if "Parameters" in wb.sheetnames:
        ws_params = wb["Parameters"]
        ws_params.column_dimensions["A"].width = 18   # Key
        ws_params.column_dimensions["B"].width = 10   # Metric
        ws_params.column_dimensions["C"].width = 18   # Expression
        ws_params.column_dimensions["D"].width = 40   # Notes

        # 🔴 FORZA Expression (colonna C) come TESTO ('@')
        expr_col_letter = "C"
        for row_idx in range(2, ws_params.max_row + 1):
            cell = ws_params[f"{expr_col_letter}{row_idx}"]
            cell.number_format = "@"

    # ----- Sheet Esempi DSL -----
    if "Esempi DSL" in wb.sheetnames:
        ws_esempi = wb["Esempi DSL"]
        ws_esempi.column_dimensions["A"].width = 25   # Categoria
        ws_esempi.column_dimensions["B"].width = 80   # Esempio
        ws_esempi.column_dimensions["C"].width = 50   # Descrizione
        
        # Wrap text per tutte le colonne
        for row in ws_esempi.iter_rows(min_row=2):
            for cell in row:
                if cell.value:
                    alignment = copy(cell.alignment)
                    alignment.wrap_text = True
                    alignment.vertical = "top"
                    cell.alignment = alignment

    wb.save(path)


def generate_training_excel(output_path: str = "training_plan_formattato.xlsx"):
    """
    Crea un file Excel con:
      - sheet 'Workouts' con allenamenti di esempio
      - sheet 'Parameters' con zone aggiornate (HR basate su %, tolerance)
      - sheet 'Esempi DSL' con guida alla sintassi
    """
    workouts_data = [
        {
            "Week": 1,
            "Date": date(2025, 5, 5),
            "Session": 1,
            "Sport": "Running",
            "Description": "Ripetute 4×5' @ Z4",
            "Steps": (
                "warmup: 10min @ Z2\n"
                "repeat 4:\n"
                "  interval: 5min @ Z4\n"
                "  recovery: 2min @ Z1\n"
                "cooldown: 5min @ Z1"
            ),
            "WorkoutId": "",
            "WorkoutScheduleId": "",
            "ScheduledDate": "",
        },
        {
            "Week": 1,
            "Date": date(2025, 5, 7),
            "Session": 2,
            "Sport": "Running",
            "Description": "Lungo + accelerazioni",
            "Steps": (
                "warmup: 10min @ easy_range\n"
                "interval: 40min @ Z2\n"
                "repeat 4:\n"
                "  interval: 100m @ Z5\n"
                "  rest: 60sec\n"
                "cooldown: 5min @ Z1"
            ),
            "WorkoutId": "",
            "ScheduledDate": "",
        },
        {
            "Week": 1,
            "Date": date(2025, 5, 9),
            "Session": 3,
            "Sport": "Running",
            "Description": "Progressivo con HR",
            "Steps": (
                "interval: 10min @ HR_Z1\n"
                "interval: 15min @ HR_Z2\n"
                "interval: 10min @ HR_Z3"
            ),
            "WorkoutId": "",
            "ScheduledDate": "",
        },
        {
            "Week": 2,
            "Date": date(2025, 5, 12),
            "Session": 1,
            "Sport": "Running",
            "Description": "Ripetute annidate 2×(6×1' veloci)",
            "Steps": (
                "warmup: lap-button @ Z2\n"
                "repeat 2:\n"
                "  repeat 6:\n"
                "    interval: 1min @ threshold\n"
                "    recovery: 1min @ recovery\n"
                "  rest: 3min\n"
                "cooldown: 5min @ Z1"
            ),
            "WorkoutId": "",
            "ScheduledDate": "",
        },
        {
            "Week": 2,
            "Date": date(2025, 5, 14),
            "Session": 2,
            "Sport": "Running",
            "Description": "10×400m con riposo completo",
            "Steps": (
                "warmup: 15min @ Z2\n"
                "repeat 10:\n"
                "  interval: 400m @ Z5\n"
                "  rest: 2min\n"
                "cooldown: 10min @ Z1"
            ),
            "WorkoutId": "",
            "ScheduledDate": "",
        },
        {
            "Week": 2,
            "Date": date(2025, 5, 16),
            "Session": 3,
            "Sport": "Running",
            "Description": "Serie di sprint brevi",
            "Steps": (
                "warmup: 10min @ easy_range\n"
                "repeat 8:\n"
                "  interval: 100m @ Z5\n"
                "  rest: 60sec\n"
                "cooldown: 5min @ Z1"
            ),
            "WorkoutId": "",
            "ScheduledDate": "",
        },
        {
            "Week": 3,
            "Date": date(2025, 5, 19),
            "Session": 1,
            "Sport": "Running",
            "Description": "Piramidale con riposo",
            "Steps": (
                "warmup: 10min @ Z2\n"
                "interval: 400m @ Z5\n"
                "rest: 90sec\n"
                "interval: 800m @ Z4\n"
                "rest: 2min\n"
                "interval: 1200m @ Z4\n"
                "rest: 3min\n"
                "interval: 800m @ Z4\n"
                "rest: 2min\n"
                "interval: 400m @ Z5\n"
                "rest: 90sec\n"
                "cooldown: 10min @ Z1"
            ),
            "WorkoutId": "",
            "ScheduledDate": "",
        },
    ]

    parameters_data = [
        # ZONE RITMO
        {"Key": "Z1", "Metric": "pace", "Expression": "6:30", 
         "Notes": "Zona 1: Recupero attivo"},
        {"Key": "Z2", "Metric": "pace", "Expression": "6:00", 
         "Notes": "Zona 2: Endurance/Aerobica"},
        {"Key": "Z3", "Metric": "pace", "Expression": "5:30", 
         "Notes": "Zona 3: Tempo/Ritmo gara"},
        {"Key": "Z4", "Metric": "pace", "Expression": "5:00", 
         "Notes": "Zona 4: Soglia anaerobica"},
        {"Key": "Z5", "Metric": "pace", "Expression": "4:30", 
         "Notes": "Zona 5: VO2max/Ripetute"},
        
        # RITMI PERSONALIZZATI (singoli - usano pace_tolerance)
        {"Key": "recovery", "Metric": "pace", "Expression": "7:00", 
         "Notes": "Ritmo di recupero lento"},
        {"Key": "marathon", "Metric": "pace", "Expression": "5:20", 
         "Notes": "Ritmo maratona"},
        {"Key": "threshold", "Metric": "pace", "Expression": "5:10", 
         "Notes": "Ritmo soglia anaerobica"},
        {"Key": "interval", "Metric": "pace", "Expression": "4:20", 
         "Notes": "Ritmo intervalli veloci"},
        
        # RANGE PERSONALIZZATI (range fissi - ignorano tolerance)
        {"Key": "easy_range", "Metric": "pace", "Expression": "6:00-6:30", 
         "Notes": "Range ritmo facile (fisso)"},
        {"Key": "tempo_range", "Metric": "pace", "Expression": "5:20-5:40", 
         "Notes": "Range ritmo tempo (fisso)"},
        {"Key": "threshold_range", "Metric": "pace", "Expression": "4:50-5:10", 
         "Notes": "Range ritmo soglia (fisso)"},
        {"Key": "fartlek", "Metric": "pace", "Expression": "4:30-6:00", 
         "Notes": "Range ampio per fartlek (fisso)"},
        
        # TOLLERANZA RITMO
        {"Key": "pace_tolerance", "Metric": "pace", "Expression": "5", 
         "Notes": "Tolleranza ±sec per ritmi singoli (es. 5:00 diventa 4:55-5:05)"},
        
        # ZONE FC (percentuali di HR_max)
        {"Key": "HR_Z1", "Metric": "hr", "Expression": "60-70%", 
         "Notes": "Zona FC 1: Recupero (60-70% di HR_max)"},
        {"Key": "HR_Z2", "Metric": "hr", "Expression": "70-80%", 
         "Notes": "Zona FC 2: Aerobica (70-80% di HR_max)"},
        {"Key": "HR_Z3", "Metric": "hr", "Expression": "80-85%", 
         "Notes": "Zona FC 3: Tempo (80-85% di HR_max)"},
        {"Key": "HR_Z4", "Metric": "hr", "Expression": "85-90%", 
         "Notes": "Zona FC 4: Soglia (85-90% di HR_max)"},
        {"Key": "HR_Z5", "Metric": "hr", "Expression": "90-100%", 
         "Notes": "Zona FC 5: VO2max (90-100% di HR_max)"},
        
        # FC MASSIMA
        {"Key": "HR_max", "Metric": "hr", "Expression": "185", 
         "Notes": "Frequenza cardiaca massima - PERSONALIZZA!"},
        
        # TOLLERANZA HR
        {"Key": "hr_tolerance", "Metric": "hr", "Expression": "5", 
         "Notes": "Tolleranza ±bpm per HR singoli (es. 150 diventa 145-155)"},
    ]

    esempi_dsl_data = [
        # DURATE
        {"Categoria": "DURATE", "Esempio": "warmup: 10min @ Z2", "Descrizione": "Durata in minuti (anche: 10m, 10')"},
        {"Categoria": "DURATE", "Esempio": "interval: 30sec @ Z5", "Descrizione": "Durata in secondi (anche: 30s)"},
        {"Categoria": "DURATE", "Esempio": "warmup: lap-button @ Z2", "Descrizione": "Premi lap per continuare"},
        
        # DISTANZE
        {"Categoria": "DISTANZE", "Esempio": "interval: 1000m @ 5:00", "Descrizione": "Distanza in metri"},
        {"Categoria": "DISTANZE", "Esempio": "interval: 5km @ Z3", "Descrizione": "Distanza in chilometri"},
        {"Categoria": "DISTANZE", "Esempio": "interval: 400m @ Z5", "Descrizione": "Distanze brevi per ripetute"},
        
        # TARGET RITMO
        {"Categoria": "TARGET RITMO", "Esempio": "interval: 3km @ 5:00", "Descrizione": "Ritmo fisso min:sec/km (usa tolerance)"},
        {"Categoria": "TARGET RITMO", "Esempio": "interval: 3km @ 5:00-5:30", "Descrizione": "Range di ritmo esplicito"},
        {"Categoria": "TARGET RITMO", "Esempio": "interval: 3km @ Z3", "Descrizione": "Zona ritmo da Parameters"},
        {"Categoria": "TARGET RITMO", "Esempio": "interval: 5km @ marathon", "Descrizione": "Ritmo custom da Parameters"},
        {"Categoria": "TARGET RITMO", "Esempio": "warmup: 10min @ easy_range", "Descrizione": "Range custom da Parameters"},
        
        # TARGET HR
        {"Categoria": "TARGET HR", "Esempio": "interval: 10min @ HR_Z3", "Descrizione": "Zona FC da Parameters (% di HR_max)"},
        {"Categoria": "TARGET HR", "Esempio": "recovery: 5min @ HR_Z1", "Descrizione": "Recupero con controllo FC"},
        {"Categoria": "TARGET HR", "Esempio": "interval: 20min @ 150-165", "Descrizione": "Range FC assoluto in bpm"},
        
        # TIPI DI STEP
        {"Categoria": "TIPI STEP", "Esempio": "warmup: 10min @ Z2", "Descrizione": "Riscaldamento"},
        {"Categoria": "TIPI STEP", "Esempio": "interval: 5min @ Z4", "Descrizione": "Intervallo intenso"},
        {"Categoria": "TIPI STEP", "Esempio": "recovery: 2min @ Z1", "Descrizione": "Recupero attivo (corsa lenta)"},
        {"Categoria": "TIPI STEP", "Esempio": "rest: 90sec", "Descrizione": "Riposo completo (FERMO)"},
        {"Categoria": "TIPI STEP", "Esempio": "cooldown: 5min @ Z1", "Descrizione": "Defaticamento"},
        
        # RIPETIZIONI SEMPLICI
        {"Categoria": "RIPETIZIONI", "Esempio": "repeat 4:\n  interval: 1km @ Z4\n  recovery: 2min @ Z1", 
         "Descrizione": "Ripeti 4 volte: 1km veloce + 2min recupero attivo"},
        {"Categoria": "RIPETIZIONI", "Esempio": "repeat 6:\n  interval: 400m @ Z5\n  rest: 90sec", 
         "Descrizione": "Ripeti 6 volte: 400m + riposo FERMO"},
        {"Categoria": "RIPETIZIONI", "Esempio": "repeat 10:\n  interval: 400m @ Z5\n  rest: 2min", 
         "Descrizione": "10×400m con riposo completo tra le ripetute"},
        {"Categoria": "RIPETIZIONI", "Esempio": "repeat 8:\n  interval: 100m @ Z5\n  rest: 60sec", 
         "Descrizione": "Serie di sprint brevi con recupero completo"},
        
        # RIPETIZIONI ANNIDATE
        {"Categoria": "RIPETIZIONI ANNIDATE", 
         "Esempio": "repeat 2:\n  repeat 6:\n    interval: 1min @ 4:30\n    recovery: 1min @ Z1\n  rest: 3min", 
         "Descrizione": "2 serie da 6 ripetute di 1min + 3min riposo FERMO tra serie"},
        
        # PIRAMIDALE
        {"Categoria": "PIRAMIDALE", 
         "Esempio": "interval: 400m @ Z5\nrest: 90sec\ninterval: 800m @ Z4\nrest: 2min\ninterval: 1200m @ Z4\nrest: 3min", 
         "Descrizione": "Piramidale crescente con riposo completo tra step"},
        {"Categoria": "PIRAMIDALE", 
         "Esempio": "interval: 1km @ Z4\nrest: 2min\ninterval: 2km @ Z3\nrest: 3min\ninterval: 1km @ Z4", 
         "Descrizione": "Piramidale su-giù con riposi fissi"},
        
        # WORKOUT COMPLETI
        {"Categoria": "WORKOUT COMPLETO", 
         "Esempio": "warmup: 15min @ Z2\nrepeat 10:\n  interval: 400m @ Z5\n  rest: 2min\ncooldown: 10min @ Z1", 
         "Descrizione": "Esempio: 10×400m con riposo completo"},
        {"Categoria": "WORKOUT COMPLETO", 
         "Esempio": "warmup: 10min @ easy_range\nrepeat 8:\n  interval: 100m @ Z5\n  rest: 60sec\ncooldown: 5min @ Z1", 
         "Descrizione": "Esempio: sprint brevi con riposo"},
        {"Categoria": "WORKOUT COMPLETO", 
         "Esempio": "warmup: lap-button @ HR_Z2\nrepeat 4:\n  interval: 5min @ HR_Z4\n  recovery: 2min @ HR_Z1\ncooldown: lap-button @ Z1", 
         "Descrizione": "Esempio: ripetute con controllo FC"},
        {"Categoria": "WORKOUT COMPLETO", 
         "Esempio": "interval: 10min @ Z1\ninterval: 20min @ Z2\ninterval: 10min @ Z3\ninterval: 5min @ Z4", 
         "Descrizione": "Esempio: progressivo senza ripetizioni"},
        
        # NOTE IMPORTANTI
        {"Categoria": "NOTE", "Esempio": "INDENTAZIONE: usare 2 spazi per ogni livello di repeat", 
         "Descrizione": "Le ripetizioni annidate richiedono corretta indentazione"},
        {"Categoria": "NOTE", "Esempio": "TOLERANCE: pace_tolerance e hr_tolerance si applicano ai valori singoli", 
         "Descrizione": "I range espliciti ignorano le tolerance"},
        {"Categoria": "NOTE", "Esempio": "HR_max: personalizza in Parameters per calcolare le zone FC automaticamente", 
         "Descrizione": "Le zone HR_Z1-HR_Z5 sono percentuali di HR_max"},
    ]

    df_workouts = pd.DataFrame(workouts_data)
    df_parameters = pd.DataFrame(parameters_data)
    df_esempi = pd.DataFrame(esempi_dsl_data)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_workouts.to_excel(writer, sheet_name="Workouts", index=False)
        df_parameters.to_excel(writer, sheet_name="Parameters", index=False)
        df_esempi.to_excel(writer, sheet_name="Esempi DSL", index=False)

    # Applica formattazione
    format_workbook_dates_and_steps(output_path)