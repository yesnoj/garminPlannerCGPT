"""Funzioni di supporto condivise dalle interfacce (Classic, Advanced, Multisport)."""
import tkinter as tk
from tkinter import ttk
from typing import List, Optional

import pandas as pd

from dsl_parser import validate_workout_rows


def show_validation_errors(parent, errors: List[str], title: str = "Workout non validi") -> None:
    """Mostra gli errori di validazione in una finestra scorrevole (copiabile)."""
    import app_theme
    c = app_theme.CURRENT.c if app_theme.CURRENT else None
    win = tk.Toplevel(parent)
    if c:
        win.configure(bg=c["bg"])
    win.title(title)
    win.geometry("820x420")
    win.transient(parent)

    frame = ttk.Frame(win, padding=12)
    frame.pack(fill="both", expand=True)

    ttk.Label(
        frame,
        text=(f"{len(errors)} problema/i trovato/i. Nessun workout e' stato caricato su Garmin.\n"
              "Correggi la colonna Steps (o il foglio Parameters) e riprova."),
        foreground=c["err"] if c else "#b00020",
        font=("", 11, "bold"),
        justify="left",
    ).pack(anchor="w", pady=(0, 8))

    text_frame = ttk.Frame(frame)
    text_frame.pack(fill="both", expand=True)
    txt = tk.Text(text_frame, wrap="word", font=("Courier", 11), relief="flat", padx=8, pady=6)
    if c:
        txt.configure(bg=c["card"], fg=c["fg"], insertbackground=c["fg"])
    sb = ttk.Scrollbar(text_frame, orient="vertical", command=txt.yview)
    txt.configure(yscrollcommand=sb.set)
    txt.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    txt.insert("1.0", "\n\n".join(f"• {e}" for e in errors))
    txt.configure(state="disabled")

    ttk.Button(frame, text="Chiudi", command=win.destroy).pack(anchor="e", pady=(8, 0))
    win.grab_set()
    parent.wait_window(win)


def confirm_workouts_valid(
    parent,
    df_workouts: pd.DataFrame,
    indices: List[int],
    df_parameters: Optional[pd.DataFrame],
    require_date: bool = False,
) -> bool:
    """Valida i workout selezionati prima dell'upload. Ritorna False (e mostra gli errori) se ci sono problemi."""
    errors = validate_workout_rows(df_workouts, indices, df_parameters, require_date=require_date)
    if errors:
        show_validation_errors(parent, errors)
        return False
    return True
