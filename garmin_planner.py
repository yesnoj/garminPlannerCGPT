"""
Garmin Training Planner – interfaccia unica.

Avvio:  python garmin_planner.py   (oppure python launcher.py)
"""
from __future__ import annotations

import datetime as dt
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, List, Optional

import pandas as pd
from openpyxl import Workbook, load_workbook

import workout_model as wm
from app_theme import Theme
from download_dialog import show_download_dialog
from dsl_parser import build_garmin_workout_from_excel_row, validate_workout_rows
from excel_utils import format_workbook_dates_and_steps, generate_training_excel
from garmin_service import GarminService, TOKEN_STORE_DIR
from gui_helpers import confirm_workouts_valid
from tree_dnd import TreeDragDrop
from workout_editor import WorkoutEditor, _is_blank, _to_date, _to_int, center_on_parent

APP_NAME = "Garmin Training Planner"
APP_VERSION = "3.0"
WORKOUT_COLUMNS = ["Week", "Date", "Session", "Sport", "Description", "Steps",
                   "WorkoutId", "WorkoutScheduleId", "ScheduledDate"]
PARAM_COLUMNS = ["Key", "Metric", "Expression", "Notes"]
GIORNI = ["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"]
DEFAULT_STEPS = "warmup: 10min @ Z2\ninterval: 20min @ Z2\ncooldown: 5min @ Z1"


def _clean_id(v) -> str:
    if _is_blank(v):
        return ""
    s = str(v).strip()
    return s[:-2] if s.endswith(".0") else s


# ===========================================================================
# Dialog di supporto
# ===========================================================================

class LoadingDialog(tk.Toplevel):
    def __init__(self, parent, title):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", lambda: None)
        f = ttk.Frame(self, padding=24)
        f.pack()
        self.lbl = ttk.Label(f, text="Attendere…", width=40)
        self.lbl.pack(anchor="w")
        pb = ttk.Progressbar(f, mode="indeterminate", length=320)
        pb.pack(pady=(12, 0))
        pb.start(12)
        self.update_idletasks()
        x = parent.winfo_rootx() + parent.winfo_width() // 2 - self.winfo_width() // 2
        y = parent.winfo_rooty() + parent.winfo_height() // 3
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        self.grab_set()

    def update_message(self, msg: str):
        self.after(0, lambda: self.lbl.configure(text=msg))

    def close(self):
        self.grab_release()
        self.destroy()


class LoginDialog(tk.Toplevel):
    def __init__(self, app: "PlannerApp"):
        super().__init__(app.root)
        self.app = app
        self.title("Accedi a Garmin Connect")
        self.transient(app.root)
        self.resizable(False, False)
        f = ttk.Frame(self, padding=22)
        f.pack(fill="both")
        ttk.Label(f, text="Garmin Connect", style="H2.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(f, text="La password non viene salvata: resta solo la sessione, in\n"
                          f"{TOKEN_STORE_DIR}", style="Small.TLabel").grid(row=1, column=0, columnspan=2,
                                                                           sticky="w", pady=(2, 14))
        ttk.Label(f, text="Email").grid(row=2, column=0, sticky="w")
        self.e_mail = ttk.Entry(f, width=34)
        self.e_mail.grid(row=2, column=1, pady=4)
        ttk.Label(f, text="Password").grid(row=3, column=0, sticky="w")
        self.e_pw = ttk.Entry(f, width=34, show="•")
        self.e_pw.grid(row=3, column=1, pady=4)
        b = ttk.Frame(f)
        b.grid(row=4, column=0, columnspan=2, sticky="we", pady=(16, 0))
        self.btn_session = ttk.Button(b, text="Usa sessione salvata", command=self._session)
        self.btn_session.pack(side="left")
        if not GarminService.has_saved_session():
            self.btn_session.state(["disabled"])
        ttk.Button(b, text="Accedi", style="Accent.TButton", command=self._login).pack(side="right")
        ttk.Button(b, text="Annulla", command=self.destroy).pack(side="right", padx=8)
        self.bind("<Return>", lambda e: self._login())
        self.bind("<Escape>", lambda e: self.destroy())
        self.e_mail.focus_set()
        center_on_parent(self, app.root)
        self.grab_set()

    def _login(self):
        mail, pw = self.e_mail.get().strip(), self.e_pw.get()
        self.e_pw.delete(0, "end")
        if not mail or not pw:
            messagebox.showerror("Accesso", "Inserisci email e password.", parent=self)
            return
        try:
            self.config(cursor="watch")
            self.update()
            self.app.garmin.login_with_credentials(mail, pw, parent=self)
        except Exception as e:
            self.config(cursor="")
            messagebox.showerror("Accesso non riuscito", str(e), parent=self)
            return
        self.app._set_connected(True)
        self.destroy()

    def _session(self):
        try:
            self.app.garmin.login_with_saved_session()
        except Exception as e:
            messagebox.showerror("Sessione", str(e), parent=self)
            return
        self.app._set_connected(True)
        legacy = getattr(self.app.garmin, "migrated_from", None)
        self.destroy()
        if legacy:
            messagebox.showwarning(
                "Sessione spostata",
                "La sessione e' stata copiata nella nuova cartella sicura:\n"
                f"{TOKEN_STORE_DIR}\n\nCancella la vecchia cartella, che contiene i tuoi token:\n{legacy}")


class ParamDialog(tk.Toplevel):
    def __init__(self, parent, values: Optional[Dict] = None):
        super().__init__(parent)
        self.result = None
        self.title("Parametro")
        self.transient(parent)
        self.resizable(False, False)
        v = values or {}
        f = ttk.Frame(self, padding=18)
        f.pack()
        self.vars = {}
        hints = {"Key": "es. easy_range, Z2, HR_Z3", "Metric": "pace · hr · power · cadence",
                 "Expression": "es. 6:30-7:00 · 70-80% · 250", "Notes": ""}
        for i, col in enumerate(PARAM_COLUMNS):
            ttk.Label(f, text={"Key": "Chiave", "Metric": "Tipo", "Expression": "Valore", "Notes": "Note"}[col]
                      ).grid(row=i * 2, column=0, sticky="w")
            var = tk.StringVar(value="" if _is_blank(v.get(col)) else str(v.get(col)))
            if col == "Metric":
                w = ttk.Combobox(f, textvariable=var, values=["pace", "hr", "power", "cadence"], width=30)
            else:
                w = ttk.Entry(f, textvariable=var, width=34)
            w.grid(row=i * 2, column=1, pady=(6, 0))
            ttk.Label(f, text=hints[col], style="Small.TLabel").grid(row=i * 2 + 1, column=1, sticky="w")
            self.vars[col] = var
        b = ttk.Frame(f)
        b.grid(row=10, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(b, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(b, text="OK", style="Accent.TButton", command=self._ok).pack(side="right", padx=8)
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        center_on_parent(self, parent)
        self.grab_set()
        self.wait_window(self)

    def _ok(self):
        vals = {k: v.get().strip() for k, v in self.vars.items()}
        if not vals["Key"]:
            messagebox.showerror("Parametro", "La chiave e' obbligatoria.", parent=self)
            return
        self.result = vals
        self.destroy()


# ===========================================================================
# Panoramica (volume settimanale)
# ===========================================================================

class WeeklyChart(tk.Canvas):
    def __init__(self, parent, theme: Theme):
        super().__init__(parent, highlightthickness=0, bd=0)
        self.theme = theme
        self.data = []  # (label, km, secs, n, first_date)
        self.bind("<Configure>", lambda e: self.redraw())
        theme.on_change(self.redraw)

    def set_data(self, data):
        self.data = data
        self.redraw()

    def redraw(self):
        c = self.theme.c
        self.delete("all")
        self.configure(bg=c["card"])
        w, h = self.winfo_width(), self.winfo_height()
        if w < 80 or not self.data:
            self.create_text(w / 2, h / 2, text="Nessun dato", fill=c["muted"], font=self.theme.font)
            return
        L, R, T, B = 44, 14, 30, 42
        vmax = max(d[1] for d in self.data) or 1
        vmax = (int(vmax / 10000) + 1) * 10  # km, arrotondato a 10
        for k in range(0, vmax + 1, 10 if vmax <= 60 else 20):
            y = h - B - (h - T - B) * k / vmax
            self.create_line(L, y, w - R, y, fill=c["grid"])
            self.create_text(L - 8, y, text=f"{k}", anchor="e", fill=c["muted"], font=self.theme.font_small)
        self.create_text(L, T - 12, text="km", anchor="w", fill=c["muted"], font=self.theme.font_small)
        n = len(self.data)
        slot = (w - L - R) / n
        bw = max(4, min(38, slot * 0.66))
        today = dt.date.today()
        for i, (lab, meters, secs, cnt, d0) in enumerate(self.data):
            km = meters / 1000
            x = L + slot * i + slot / 2
            top = h - B - (h - T - B) * km / vmax
            current = d0 is not None and d0 <= today < d0 + dt.timedelta(days=7)
            col = c["ok"] if current else c["accent"]
            self.create_rectangle(x - bw / 2, top, x + bw / 2, h - B, fill=col, outline="")
            if km > 0 and bw >= 16:
                self.create_text(x, top - 8, text=f"{km:.0f}", fill=c["fg"], font=self.theme.font_small)
            if n <= 14 or i % 2 == 0:
                self.create_text(x, h - B + 12, text=lab, fill=c["muted"], font=self.theme.font_small)
                if d0 is not None and (n <= 16 or i % 4 == 0):
                    self.create_text(x, h - B + 26, text=d0.strftime("%d/%m"), fill=c["muted"],
                                     font=self.theme.font_small)


# ===========================================================================
# App
# ===========================================================================

class PlannerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.theme = Theme(root)
        self.garmin = GarminService()
        self.df_workouts: Optional[pd.DataFrame] = None
        self.df_parameters: Optional[pd.DataFrame] = None
        self.excel_path: Optional[str] = None
        self._errors: Dict[int, str] = {}
        self._est: Dict[int, tuple] = {}
        self._current_idx: Optional[int] = None
        self._ignore_select = False

        root.title(APP_NAME)
        root.geometry("1440x900")
        root.minsize(1180, 720)
        self._build()
        self._bind_keys()
        root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.theme.on_change(self._retheme)
        self._retheme()
        self._refresh_everything()
        root.after(80, self._init_sash)
        if self.theme.using_fallback:
            self._set_status("Suggerimento: per la grafica migliore installa il tema sv-ttk  →  "
                             "python -m pip install sv-ttk darkdetect")

    def _init_sash(self):
        self.root.update_idletasks()
        w = self.paned.winfo_width()
        if w > 200:
            self.paned.sashpos(0, int(w * 0.47))

    # ================================================================ UI
    def _build(self):
        t = self.theme
        # --- header
        head = ttk.Frame(self.root, padding=(18, 14, 18, 6))
        head.pack(fill="x")
        left = ttk.Frame(head)
        left.pack(side="left")
        ttk.Label(left, text=APP_NAME, style="Title.TLabel").pack(anchor="w")
        self.lbl_file = ttk.Label(left, text="Nessun piano aperto", style="Muted.TLabel")
        self.lbl_file.pack(anchor="w")
        right = ttk.Frame(head)
        right.pack(side="right")
        self.btn_theme = ttk.Button(right, text="", width=3, command=self.theme.toggle)
        self.btn_theme.pack(side="right")
        self.btn_login = ttk.Button(right, text="Accedi…", command=self.login)
        self.btn_login.pack(side="right", padx=8)
        self.lbl_conn = ttk.Label(right, text="● Non connesso")
        self.lbl_conn.pack(side="right", padx=8)
        self.var_prefix = tk.BooleanVar(value=False)
        ttk.Checkbutton(right, text="Nome con prefisso W/S", variable=self.var_prefix,
                        style="Switch.TCheckbutton").pack(side="right", padx=(0, 18))

        # --- toolbar (due gruppi: se la finestra e' stretta il gruppo Garmin va a capo)
        bar = ttk.Frame(self.root, padding=(18, 4, 18, 10))
        bar.pack(fill="x")
        left = ttk.Frame(bar)
        self._btn(left, "Nuovo piano", self.new_plan)
        self._btn(left, "Apri…", self.open_plan)
        self._btn(left, "Salva", self.save_plan)
        self._sep(left)
        ttk.Label(left, text="Allenamento:", style="Muted.TLabel").pack(side="left", padx=(0, 6))
        self._btn(left, "+ Nuovo", self.new_workout)
        self._btn(left, "Duplica", self.duplicate_workout)
        self._btn(left, "Elimina", self.delete_workout)
        garmin = ttk.Frame(bar)
        self._sep(garmin)
        ttk.Label(garmin, text="Garmin:", style="Muted.TLabel").pack(side="left", padx=(0, 6))
        self._btn(garmin, "Carica e pianifica", self.upload_and_schedule, accent=True)
        self._btn(garmin, "Solo carica", self.upload)
        self._btn(garmin, "Togli dal calendario", self.unschedule)
        self._btn(garmin, "Elimina da Garmin", self.delete_from_garmin)
        self._btn(garmin, "Scarica…", self.download)
        left.grid(row=0, column=0, sticky="w")
        garmin.grid(row=0, column=1, sticky="w")
        self._toolbar_wrapped = False

        def reflow(_e=None):
            need = left.winfo_reqwidth() + garmin.winfo_reqwidth() + 8
            wrap = need > bar.winfo_width() - 36
            if wrap != self._toolbar_wrapped:
                self._toolbar_wrapped = wrap
                garmin.grid_configure(row=1 if wrap else 0, column=0 if wrap else 1, pady=(6, 0) if wrap else 0)
        bar.bind("<Configure>", reflow)

        # --- corpo
        self.paned = ttk.PanedWindow(self.root, orient="horizontal")
        self.paned.pack(fill="both", expand=True, padx=18)
        self._build_list(self.paned)
        self._build_right(self.paned)

        # --- barra di stato
        self.lbl_status = ttk.Label(self.root, text="", style="Muted.TLabel", padding=(18, 6))
        self.lbl_status.pack(fill="x", side="bottom")

    def _btn(self, parent, text, cmd, accent=False):
        b = ttk.Button(parent, text=text, command=cmd, style="Accent.TButton" if accent else "TButton")
        b.pack(side="left", padx=2)
        return b

    def _sep(self, parent):
        ttk.Separator(parent, orient="vertical").pack(side="left", fill="y", padx=10)

    def _build_list(self, paned):
        frame = ttk.Frame(paned, padding=(0, 0, 10, 0))
        paned.add(frame, weight=5)
        top = ttk.Frame(frame)
        top.pack(fill="x", pady=(0, 8))
        self.var_search = tk.StringVar()
        ent = ttk.Entry(top, textvariable=self.var_search)
        ent.pack(side="left", fill="x", expand=True)
        self._placeholder(ent, self.var_search, "Cerca allenamento…")
        self.var_filter = tk.StringVar(value="Tutti")
        cb = ttk.Combobox(top, textvariable=self.var_filter, state="readonly", width=16,
                          values=["Tutti", "Da caricare", "Su Garmin", "Pianificati", "Con errori", "Da oggi in poi"])
        cb.pack(side="left", padx=(8, 0))
        cb.bind("<<ComboboxSelected>>", lambda e: self._populate_tree())
        self.var_search.trace_add("write", lambda *a: self._populate_tree())

        wrap = ttk.Frame(frame)
        wrap.pack(fill="both", expand=True)
        cols = ("desc", "dur", "km", "stato")
        self.tree = ttk.Treeview(wrap, columns=cols, selectmode="extended")
        self.tree.heading("#0", text="Giorno")
        self.tree.heading("desc", text="Allenamento")
        self.tree.heading("dur", text="Durata")
        self.tree.heading("km", text="Km")
        self.tree.heading("stato", text="Stato")
        self.tree.column("#0", width=150, stretch=False)
        self.tree.column("desc", width=280, stretch=True)
        self.tree.column("dur", width=70, anchor="e", stretch=False)
        self.tree.column("km", width=75, anchor="e", stretch=False)
        self.tree.column("stato", width=120, stretch=False)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Delete>", lambda e: self.delete_workout())
        self.dnd = TreeDragDrop(self.tree, self.theme, on_drop=self._on_workout_drop,
                                can_drag=lambda iid: iid.startswith("w:"),
                                allow_inside=lambda iid: iid.startswith("week:"))
        self.menu_list = tk.Menu(self.tree, tearoff=False)
        for label, cmd in (("Duplica", self.duplicate_workout),
                           ("Duplica nella settimana successiva", lambda: self.duplicate_workout(shift_weeks=1)),
                           (None, None),
                           ("Sposta su", lambda: self.move_selected(-1)),
                           ("Sposta giu'", lambda: self.move_selected(+1)),
                           ("Ordina tutto per data", self.sort_by_date),
                           (None, None),
                           ("Carica e pianifica su Garmin", self.upload_and_schedule),
                           ("Togli dal calendario Garmin", self.unschedule),
                           ("Elimina da Garmin", self.delete_from_garmin),
                           (None, None),
                           ("Elimina dal piano", self.delete_workout)):
            if label is None:
                self.menu_list.add_separator()
            else:
                self.menu_list.add_command(label=label, command=cmd)
        for seq in ("<Button-3>", "<Button-2>", "<Control-Button-1>"):
            self.tree.bind(seq, self._list_context_menu)

        ttk.Label(frame, text="Trascina gli allenamenti per riordinarli (anche in un'altra settimana) · "
                              "tasto destro per duplicare, spostare, ordinare",
                  style="Small.TLabel").pack(anchor="w", pady=(6, 0))

        # schermata iniziale
        self.welcome = ttk.Frame(wrap, padding=30)
        ttk.Label(self.welcome, text="Nessun piano aperto", style="H2.TLabel").pack(pady=(40, 6))
        ttk.Label(self.welcome, text="Apri un file Excel con i tuoi allenamenti\n"
                                     "oppure creane uno nuovo da un modello.",
                  style="Muted.TLabel", justify="center").pack()
        b = ttk.Frame(self.welcome)
        b.pack(pady=18)
        ttk.Button(b, text="Apri piano…", style="Accent.TButton", command=self.open_plan).pack(side="left", padx=4)
        ttk.Button(b, text="Nuovo piano", command=self.new_plan).pack(side="left", padx=4)

    def _build_right(self, paned):
        self.nb = ttk.Notebook(paned)
        paned.add(self.nb, weight=6)
        self.editor = WorkoutEditor(self.nb, self.theme, lambda: self.df_parameters,
                                    self._on_editor_save, self._on_editor_dirty)
        self.nb.add(self.editor, text="  Allenamento  ")

        # parametri
        pf = ttk.Frame(self.nb, padding=16)
        self.nb.add(pf, text="  Parametri  ")
        ttk.Label(pf, text="Zone e ritmi riutilizzabili negli step (es. @ Z2, @ easy_range)",
                  style="Muted.TLabel").pack(anchor="w", pady=(0, 8))
        pb = ttk.Frame(pf)
        pb.pack(fill="x", pady=(0, 8))
        ttk.Button(pb, text="+ Parametro", style="Accent.TButton", command=self.add_param).pack(side="left")
        ttk.Button(pb, text="Modifica", command=self.edit_param).pack(side="left", padx=4)
        ttk.Button(pb, text="Elimina", command=self.delete_param).pack(side="left")
        self.ptree = ttk.Treeview(pf, columns=("metric", "expr", "notes"), selectmode="browse")
        for col, txt, w in (("#0", "Chiave", 170), ("metric", "Tipo", 80), ("expr", "Valore", 140),
                            ("notes", "Note", 320)):
            self.ptree.heading(col, text=txt)
            self.ptree.column(col, width=w, stretch=(col == "notes"))
        self.ptree.pack(fill="both", expand=True)
        self.ptree.bind("<Double-1>", lambda e: self.edit_param())

        # panoramica
        of = ttk.Frame(self.nb, padding=16)
        self.nb.add(of, text="  Panoramica  ")
        self.lbl_overview = ttk.Label(of, text="", style="H2.TLabel", wraplength=600, justify="left")
        self.lbl_overview.pack(anchor="w")
        ttk.Label(of, text="Chilometri stimati per settimana (in verde la settimana corrente)",
                  style="Muted.TLabel").pack(anchor="w", pady=(2, 10))
        self.chart = WeeklyChart(of, self.theme)
        self.chart.pack(fill="both", expand=True)

    def _placeholder(self, entry, var, text):
        def on_focus_in(_):
            if var.get() == text:
                var.set("")
                entry.configure(foreground=self.theme.c["fg"])

        def on_focus_out(_):
            if not var.get():
                var.set(text)
                entry.configure(foreground=self.theme.c["muted"])
        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)
        self._search_placeholder = text
        var.set(text)
        entry.configure(foreground=self.theme.c["muted"])

    def _list_context_menu(self, e):
        iid = self.tree.identify_row(e.y)
        if not iid:
            return "break"
        if iid not in self.tree.selection():
            self.tree.selection_set(iid)
        try:
            self.menu_list.tk_popup(e.x_root, e.y_root)
        finally:
            self.menu_list.grab_release()
        return "break"

    def _bind_keys(self):
        mod = "Command" if self.root.tk.call("tk", "windowingsystem") == "aqua" else "Control"
        self.root.bind_all(f"<{mod}-d>", lambda e: self.duplicate_workout())
        self.root.bind_all(f"<{mod}-Up>", lambda e: self.move_selected(-1))
        self.root.bind_all(f"<{mod}-Down>", lambda e: self.move_selected(+1))
        self.root.bind_all(f"<{mod}-o>", lambda e: self.open_plan())
        self.root.bind_all(f"<{mod}-s>", lambda e: self._save_shortcut())
        self.root.bind_all(f"<{mod}-n>", lambda e: self.new_workout())

    def _retheme(self):
        c = self.theme.c
        self.root.configure(bg=c["bg"])
        self.btn_theme.configure(text="☀" if self.theme.mode == "dark" else "☾")
        self.tree.tag_configure("week", background=c["week_bg"], font=self.theme.font_bold)
        self.tree.tag_configure("error", foreground=c["err"])
        self.tree.tag_configure("scheduled", foreground=c["fg"])
        self.tree.tag_configure("past", foreground=c["muted"])
        self._set_connected(self.garmin.authenticated)

    # ============================================================ stato
    def _set_connected(self, ok: bool):
        c = self.theme.c
        if ok:
            self.lbl_conn.configure(text="● Connesso a Garmin", foreground=c["ok"])
            self.btn_login.configure(text="Riconnetti…")
        else:
            self.lbl_conn.configure(text="● Non connesso", foreground=c["muted"])
            self.btn_login.configure(text="Accedi…")

    def _set_status(self, text=None):
        if text:
            self.lbl_status.configure(text=text)
            return
        if self.df_workouts is None:
            self.lbl_status.configure(text=f"{APP_NAME} {APP_VERSION}")
            return
        n = len(self.df_workouts)
        weeks = self.df_workouts["Week"].dropna().nunique() if "Week" in self.df_workouts else 0
        km = sum(v[1] for v in self._est.values()) / 1000
        secs = sum(v[0] for v in self._est.values())
        sched = sum(1 for i in range(n) if _clean_id(self.df_workouts.iloc[i].get("WorkoutScheduleId")))
        errs = len(self._errors)
        self.lbl_status.configure(
            text=f"{n} allenamenti · {weeks} settimane · ~{km:.0f} km · ~{secs / 3600:.0f} h"
                 f" · {sched} pianificati su Garmin" + (f" · {errs} con errori" if errs else " · nessun errore"))

    # ============================================================ file
    def new_plan(self):
        if not self._confirm_discard():
            return
        path = filedialog.asksaveasfilename(title="Nuovo piano", defaultextension=".xlsx",
                                            initialfile="piano_allenamenti.xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            generate_training_excel(path, multisport=True)
        except Exception as e:
            messagebox.showerror("Nuovo piano", f"Impossibile creare il file:\n{e}")
            return
        self._load(path)

    def open_plan(self):
        if not self._confirm_discard():
            return
        path = filedialog.askopenfilename(title="Apri piano", filetypes=[("Excel", "*.xlsx *.xlsm")])
        if path:
            self._load(path)

    def _load(self, path: str):
        try:
            df = pd.read_excel(path, sheet_name="Workouts")
            try:
                params = pd.read_excel(path, sheet_name="Parameters",
                                       dtype={"Key": str, "Metric": str, "Expression": str, "Notes": str})
            except ValueError:
                params = pd.DataFrame(columns=PARAM_COLUMNS)
        except Exception as e:
            messagebox.showerror("Apri piano", f"Impossibile leggere il file:\n{e}")
            return
        for col in WORKOUT_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        for col in ("WorkoutId", "WorkoutScheduleId"):
            df[col] = df[col].map(_clean_id).astype(object)
        df["ScheduledDate"] = df["ScheduledDate"].map(lambda v: "" if _is_blank(v) else str(_to_date(v) or v)).astype(object)
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce", dayfirst=False)
        for col in ("Description", "Steps", "Sport"):
            df[col] = df[col].astype(object)
        self.df_workouts = df.reset_index(drop=True)
        self.df_parameters = params
        self.excel_path = path
        self.lbl_file.configure(text=path)
        self.editor.load(None)
        self._current_idx = None
        self._refresh_everything()
        if self._errors:
            self._set_status(f"Piano caricato: {len(self._errors)} allenamenti hanno errori (in rosso). "
                             "Filtra 'Con errori' per vederli.")

    def save_plan(self):
        if self.df_workouts is None:
            return
        if self.editor.dirty:
            self.editor.save()
        path = self.excel_path
        if not path:
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
            if not path:
                return
        if self._write_excel(path, show_errors=True):
            self.excel_path = path
            self.lbl_file.configure(text=path)
            self._set_status(f"Salvato: {os.path.basename(path)}")

    def _save_shortcut(self):
        if self.editor.dirty:
            self.editor.save()
        else:
            self.save_plan()

    def _write_excel(self, path: str, show_errors: bool = False) -> bool:
        """Riscrive Workouts e Parameters lasciando intatti gli altri fogli del file."""
        if self.df_workouts is None:
            return False
        try:
            df = self.df_workouts.copy()
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df["ScheduledDate"] = df["ScheduledDate"].map(lambda v: _to_date(v))
            for col in ("WorkoutId", "WorkoutScheduleId"):
                df[col] = df[col].map(_clean_id)
            params = self.df_parameters if self.df_parameters is not None else pd.DataFrame(columns=PARAM_COLUMNS)

            wb = load_workbook(path) if os.path.exists(path) else Workbook()
            if not os.path.exists(path):
                wb.remove(wb.active)
            for name, data in (("Workouts", df), ("Parameters", params)):
                pos = wb.sheetnames.index(name) if name in wb.sheetnames else len(wb.sheetnames)
                if name in wb.sheetnames:
                    wb.remove(wb[name])
                ws = wb.create_sheet(name, pos)
                ws.append(list(data.columns))
                for row in data.itertuples(index=False):
                    out = []
                    for v in row:
                        if _is_blank(v):
                            out.append(None)
                        elif isinstance(v, pd.Timestamp):
                            out.append(v.to_pydatetime())
                        else:
                            out.append(v)
                    ws.append(out)
            if "Workouts" in wb.sheetnames:
                wb.active = wb.sheetnames.index("Workouts")
            wb.save(path)
            format_workbook_dates_and_steps(path)
            return True
        except PermissionError:
            messagebox.showwarning("Salvataggio", f"Il file e' aperto in un altro programma (Excel?).\n"
                                                  f"Chiudilo e salva di nuovo:\n{path}")
        except Exception as e:
            if show_errors:
                messagebox.showerror("Salvataggio", f"Errore nel salvataggio:\n{e}")
            else:
                messagebox.showwarning("Salvataggio automatico", f"Non riesco a salvare:\n{e}")
        return False

    def _autosave(self):
        if self.excel_path and self.df_workouts is not None:
            self._write_excel(self.excel_path)

    def _confirm_discard(self) -> bool:
        if self.editor.dirty:
            ans = messagebox.askyesnocancel("Modifiche non salvate", "Salvare le modifiche all'allenamento?")
            if ans is None:
                return False
            if ans:
                self.editor.save()
        return True

    def _on_close(self):
        if self._confirm_discard():
            self.root.destroy()

    # ============================================================ lista
    def _refresh_everything(self):
        self._recompute_all()
        self._populate_tree()
        self._populate_params()
        self._update_overview()
        self._set_status()

    def _recompute_row(self, i: int):
        row = self.df_workouts.iloc[i]
        errs = validate_workout_rows(self.df_workouts, [i], self.df_parameters)
        if errs:
            self._errors[i] = errs[0].split(": ", 1)[-1]
        else:
            self._errors.pop(i, None)
        steps = "" if _is_blank(row.get("Steps")) else str(row.get("Steps"))
        self._est[i] = wm.estimate_text(steps, str(row.get("Sport") or ""), self.df_parameters)

    def _recompute_all(self):
        self._errors, self._est = {}, {}
        if self.df_workouts is None:
            return
        for i in range(len(self.df_workouts)):
            self._recompute_row(i)

    def _status_of(self, i: int):
        row = self.df_workouts.iloc[i]
        if i in self._errors:
            return "⚠ Errore", "error"
        if _clean_id(row.get("WorkoutScheduleId")) or not _is_blank(row.get("ScheduledDate")):
            return "✓ Pianificato", "scheduled"
        if _clean_id(row.get("WorkoutId")):
            return "● Su Garmin", "uploaded"
        return "○ Da caricare", "draft"

    def _row_matches(self, i: int) -> bool:
        row = self.df_workouts.iloc[i]
        q = self.var_search.get().strip().lower()
        if q and q != self._search_placeholder.lower():
            hay = f"{row.get('Description', '')} {row.get('Steps', '')}".lower()
            if q not in hay:
                return False
        f = self.var_filter.get()
        _, st = self._status_of(i)
        if f == "Da caricare":
            return st == "draft"
        if f == "Su Garmin":
            return st in ("uploaded", "scheduled")
        if f == "Pianificati":
            return st == "scheduled"
        if f == "Con errori":
            return st == "error"
        if f == "Da oggi in poi":
            d = _to_date(row.get("Date"))
            return d is None or d >= dt.date.today()
        return True

    def _week_key(self, i: int):
        w = self.df_workouts.iloc[i].get("Week")
        return None if _is_blank(w) else _to_int(w)

    def _populate_tree(self):
        self._ignore_select = True
        sel = set(self.tree.selection())
        self.tree.delete(*self.tree.get_children())
        has = self.df_workouts is not None
        if has:
            self.welcome.place_forget()
        else:
            self.welcome.place(relx=0, rely=0, relwidth=1, relheight=1)
            self._ignore_select = False
            return
        weeks: Dict = {}
        order = []
        for i in range(len(self.df_workouts)):
            if not self._row_matches(i):
                continue
            k = self._week_key(i)
            if k not in weeks:
                weeks[k] = []
                order.append(k)
            weeks[k].append(i)
        today = dt.date.today()
        for k in order:
            idxs = weeks[k]
            secs = sum(self._est.get(i, (0, 0))[0] for i in idxs)
            meters = sum(self._est.get(i, (0, 0))[1] for i in idxs)
            dates = [d for d in (_to_date(self.df_workouts.iloc[i].get("Date")) for i in idxs) if d]
            span = f"{min(dates):%d/%m}–{max(dates):%d/%m}" if dates else ""
            n_sched = sum(1 for i in idxs if self._status_of(i)[1] == "scheduled")
            label = f"Settimana {k}" if k is not None else "Senza settimana"
            wid = f"week:{k}"
            self.tree.insert("", "end", iid=wid, text=label, open=True, tags=("week",),
                             values=(f"dal {span.replace('–', ' al ')}" if span else "", wm.fmt_duration(secs),
                                     wm.fmt_km(meters).replace(" km", ""), f"{n_sched}/{len(idxs)} pianif."))
            for i in idxs:
                self._insert_row(wid, i, today)
        for iid in sel:
            if self.tree.exists(iid):
                self.tree.selection_add(iid)
        self._ignore_select = False

    def _insert_row(self, parent, i, today=None):
        today = today or dt.date.today()
        row = self.df_workouts.iloc[i]
        d = _to_date(row.get("Date"))
        day = f"{GIORNI[d.weekday()]} {d:%d/%m}" if d else "—"
        secs, meters = self._est.get(i, (0, 0))
        status, tag = self._status_of(i)
        tags = [tag]
        if d and d < today and tag != "error":
            tags.append("past")
        desc = "" if _is_blank(row.get("Description")) else str(row.get("Description"))
        sport = wm.sport_key(str(row.get("Sport") or ""))
        icon = {"cycling": "🚴 ", "swimming": "🏊 "}.get(sport, "")
        self.tree.insert(parent, "end", iid=f"w:{i}", text=f"  {day}",
                         values=(icon + desc, wm.fmt_duration(secs), wm.fmt_km(meters).replace(" km", ""), status),
                         tags=tuple(tags))

    def _refresh_row(self, i: int):
        self._recompute_row(i)
        iid = f"w:{i}"
        if self.tree.exists(iid):
            row = self.df_workouts.iloc[i]
            d = _to_date(row.get("Date"))
            secs, meters = self._est.get(i, (0, 0))
            status, tag = self._status_of(i)
            desc = "" if _is_blank(row.get("Description")) else str(row.get("Description"))
            self.tree.item(iid, text=f"  {GIORNI[d.weekday()]} {d:%d/%m}" if d else "  —",
                           values=(desc, wm.fmt_duration(secs), wm.fmt_km(meters).replace(" km", ""), status),
                           tags=(tag,))
        self._set_status()

    def _selected_indices(self, require=True) -> List[int]:
        out = []
        for iid in self.tree.selection():
            if iid.startswith("w:"):
                out.append(int(iid[2:]))
            elif iid.startswith("week:"):
                out.extend(int(c[2:]) for c in self.tree.get_children(iid))
        out = sorted(set(out))
        if require and not out:
            messagebox.showinfo(APP_NAME, "Seleziona uno o piu' allenamenti (o una settimana intera) nella lista.")
        return out

    def _on_select(self, _e=None):
        if self._ignore_select:
            return
        sel = [s for s in self.tree.selection() if s.startswith("w:")]
        if len(sel) != 1 or len(self.tree.selection()) != 1:
            return
        i = int(sel[0][2:])
        if i == self._current_idx:
            return
        if not self._confirm_discard():
            return
        self._current_idx = i
        self.editor.load(self.df_workouts.iloc[i].to_dict())
        self.nb.select(0)

    def _on_editor_dirty(self, dirty: bool):
        title = APP_NAME + (" – modifiche non salvate" if dirty else "")
        self.root.title(title)

    def _on_editor_save(self, values: Dict):
        i = self._current_idx
        if i is None or self.df_workouts is None:
            return
        old_steps = str(self.df_workouts.at[i, "Steps"])
        for k, v in values.items():
            if k == "Date":
                v = pd.Timestamp(v) if v else pd.NaT
            self.df_workouts.at[i, k] = v
        self._recompute_row(i)
        self._populate_tree()
        self.tree.selection_set(f"w:{i}")
        self._update_overview()
        self._autosave()
        msg = "Allenamento salvato."
        if _clean_id(self.df_workouts.at[i, "WorkoutId"]) and old_steps != values["Steps"]:
            msg += (" Attenzione: la versione su Garmin non cambia da sola: usa 'Elimina da Garmin' "
                    "e poi 'Carica e pianifica' per aggiornarla.")
        self._set_status(msg)

    # ============================================================ allenamenti
    def _require_plan(self) -> bool:
        if self.df_workouts is None:
            messagebox.showinfo(APP_NAME, "Apri prima un piano o creane uno nuovo.")
            return False
        return True

    def _insert_df_row(self, pos: int, row: Dict) -> int:
        top = self.df_workouts.iloc[:pos]
        bottom = self.df_workouts.iloc[pos:]
        new = pd.DataFrame([row], columns=self.df_workouts.columns)
        self.df_workouts = pd.concat([top, new, bottom], ignore_index=True)
        return pos

    def new_workout(self):
        if not self._require_plan() or not self._confirm_discard():
            return
        sel = self._selected_indices(require=False)
        base = self.df_workouts.iloc[sel[-1]].to_dict() if sel else (
            self.df_workouts.iloc[-1].to_dict() if len(self.df_workouts) else {})
        d = _to_date(base.get("Date")) or dt.date.today()
        row = {c: "" for c in self.df_workouts.columns}
        row.update({"Week": base.get("Week", 1), "Date": pd.Timestamp(d + dt.timedelta(days=1)),
                    "Session": _to_int(base.get("Session"), 0) + 1, "Sport": base.get("Sport") or "Running",
                    "Description": "Nuovo allenamento", "Steps": DEFAULT_STEPS})
        pos = (sel[-1] + 1) if sel else len(self.df_workouts)
        i = self._insert_df_row(pos, row)
        self._after_structure_change(i, focus_title=True)

    def duplicate_workout(self, shift_weeks: int = 0):
        """Duplica gli allenamenti selezionati (anche un'intera settimana).
        Con shift_weeks=1 la copia va nella settimana successiva, stesso giorno."""
        if not self._require_plan() or not self._confirm_discard():
            return
        sel = self._selected_indices()
        if not sel:
            return
        copies = []
        for i in sel:
            row = self.df_workouts.iloc[i].to_dict()
            row.update({"WorkoutId": "", "WorkoutScheduleId": "", "ScheduledDate": ""})
            if shift_weeks:
                if not _is_blank(row.get("Week")):
                    row["Week"] = _to_int(row["Week"]) + shift_weeks
                d = _to_date(row.get("Date"))
                if d:
                    row["Date"] = pd.Timestamp(d + dt.timedelta(days=7 * shift_weeks))
            else:
                row["Description"] = f"{row.get('Description', '')} (copia)"
            copies.append(row)
        # posizione: dopo l'ultima riga della settimana di destinazione (se esiste), altrimenti dopo la selezione
        pos = sel[-1] + 1
        if shift_weeks:
            wk = copies[-1].get("Week")
            same = [j for j in range(len(self.df_workouts)) if not _is_blank(wk) and self._week_key(j) == _to_int(wk)]
            if same:
                pos = same[-1] + 1
        for k, row in enumerate(copies):
            self._insert_df_row(pos + k, row)
        where = "nella settimana successiva" if shift_weeks else ""
        self._after_structure_change(pos)
        self._set_status(f"Duplicati {len(copies)} allenamenti {where}".strip() + ".")

    def _move_row(self, i: int, dest: int, new_week=None) -> int:
        """Sposta la riga i in posizione dest (indice prima della rimozione). Ritorna il nuovo indice."""
        row = self.df_workouts.iloc[i].to_dict()
        msg = ""
        if new_week is not None and self._week_key(i) != new_week and not _is_blank(row.get("Week")):
            delta = new_week - _to_int(row["Week"])
            row["Week"] = new_week
            d = _to_date(row.get("Date"))
            if d:
                row["Date"] = pd.Timestamp(d + dt.timedelta(days=7 * delta))
                msg = f" Data spostata al {row['Date']:%d/%m/%Y}."
            if _clean_id(row.get("WorkoutScheduleId")):
                msg += " Era gia' nel calendario Garmin: usa 'Togli dal calendario' e ripianificalo."
        self.df_workouts = self.df_workouts.drop(index=i).reset_index(drop=True)
        if dest > i:
            dest -= 1
        self._insert_df_row(dest, row)
        self._move_msg = msg
        return dest

    def _on_workout_drop(self, src: str, target: str, where: str):
        if self.df_workouts is None or not src.startswith("w:"):
            return
        if not self._confirm_discard():
            return
        i = int(src[2:])
        if target.startswith("week:"):
            weeks = list(self.tree.get_children(""))
            wpos = weeks.index(target)
            if where == "before" and wpos > 0:          # linea sopra l'intestazione = fine settimana precedente
                target, where = weeks[wpos - 1], "inside"
            k = target.split(":", 1)[1]
            week = None if k == "None" else int(k)
            rows = [int(c[2:]) for c in self.tree.get_children(target)]
            if not rows:
                return
            dest = rows[-1] + 1 if where == "inside" else rows[0]
        else:
            j = int(target[2:])
            week = self._week_key(j)
            dest = j if where == "before" else j + 1
        new = self._move_row(i, dest, week)
        self._after_structure_change(new)
        self._set_status("Allenamento spostato." + getattr(self, "_move_msg", ""))

    def move_selected(self, delta: int):
        """Sposta l'allenamento selezionato su/giu' di una posizione nella lista (anche tra settimane)."""
        if self.df_workouts is None:
            return
        sel = [s for s in self.tree.selection() if s.startswith("w:")]
        if len(sel) != 1 or not self._confirm_discard():
            return
        order = [iid for wk in self.tree.get_children("") for iid in self.tree.get_children(wk)]
        pos = order.index(sel[0])
        nb = pos + delta
        if not 0 <= nb < len(order):
            return
        self._on_workout_drop(sel[0], order[nb], "before" if delta < 0 else "after")

    def sort_by_date(self):
        if not self._require_plan() or not self._confirm_discard():
            return
        tmp = self.df_workouts.copy()
        tmp["_d"] = pd.to_datetime(tmp["Date"], errors="coerce")
        tmp["_s"] = pd.to_numeric(tmp["Session"], errors="coerce")
        self.df_workouts = tmp.sort_values(["_d", "_s"], kind="stable", na_position="last").drop(
            columns=["_d", "_s"]).reset_index(drop=True)
        self._after_structure_change(None)
        self._set_status("Allenamenti ordinati per data.")

    def delete_workout(self):
        if not self._require_plan():
            return
        sel = self._selected_indices()
        if not sel:
            return
        on_garmin = [i for i in sel if _clean_id(self.df_workouts.iloc[i].get("WorkoutId"))]
        msg = f"Eliminare {len(sel)} allenamento/i dal piano?"
        if on_garmin:
            msg += (f"\n\n{len(on_garmin)} sono gia' su Garmin e resteranno li'. "
                    "Per toglierli usa prima 'Elimina da Garmin'.")
        if not messagebox.askyesno("Elimina", msg):
            return
        self.df_workouts = self.df_workouts.drop(index=sel).reset_index(drop=True)
        self._current_idx = None
        self.editor.load(None)
        self._after_structure_change(None)

    def _after_structure_change(self, select: Optional[int], focus_title=False):
        self._current_idx = None
        self._recompute_all()
        self._populate_tree()
        self._update_overview()
        self._set_status()
        self._autosave()
        if select is not None:
            iid = f"w:{select}"
            if self.tree.exists(iid):
                self.tree.selection_set(iid)
                self.tree.see(iid)
                self._on_select()
                if focus_title:
                    self.editor.ent_title.focus_set()
                    self.editor.ent_title.select_range(0, "end")

    # ============================================================ parametri
    def _populate_params(self):
        self.ptree.delete(*self.ptree.get_children())
        if self.df_parameters is None:
            return
        for i, r in self.df_parameters.iterrows():
            vals = ["" if _is_blank(r.get(c)) else str(r.get(c)) for c in PARAM_COLUMNS]
            self.ptree.insert("", "end", iid=str(i), text=vals[0], values=vals[1:])

    def _params_changed(self):
        self.df_parameters = self.df_parameters.reset_index(drop=True)
        self._populate_params()
        self._refresh_everything()
        self.editor._refresh_all(source="text")
        self._autosave()

    def add_param(self):
        if not self._require_plan():
            return
        dlg = ParamDialog(self.root)
        if dlg.result:
            self.df_parameters = pd.concat([self.df_parameters, pd.DataFrame([dlg.result])], ignore_index=True)
            self._params_changed()

    def edit_param(self):
        sel = self.ptree.selection()
        if not sel:
            return
        i = int(sel[0])
        dlg = ParamDialog(self.root, self.df_parameters.iloc[i].to_dict())
        if dlg.result:
            for k, v in dlg.result.items():
                self.df_parameters.at[i, k] = v
            self._params_changed()

    def delete_param(self):
        sel = self.ptree.selection()
        if not sel:
            return
        i = int(sel[0])
        key = self.df_parameters.iloc[i]["Key"]
        if messagebox.askyesno("Elimina parametro", f"Eliminare '{key}'?\nGli allenamenti che lo usano "
                                                    "diventeranno non validi."):
            self.df_parameters = self.df_parameters.drop(index=i)
            self._params_changed()

    # ============================================================ panoramica
    def _update_overview(self):
        if self.df_workouts is None:
            self.lbl_overview.configure(text="")
            self.chart.set_data([])
            return
        weeks: Dict = {}
        order = []
        for i in range(len(self.df_workouts)):
            k = self._week_key(i)
            if k not in weeks:
                weeks[k] = [0.0, 0.0, 0, None]
                order.append(k)
            s, m = self._est.get(i, (0, 0))
            weeks[k][0] += m
            weeks[k][1] += s
            weeks[k][2] += 1
            d = _to_date(self.df_workouts.iloc[i].get("Date"))
            if d:
                monday = d - dt.timedelta(days=d.weekday())
                weeks[k][3] = min(weeks[k][3], monday) if weeks[k][3] else monday
        data = [(f"S{k}" if k is not None else "–", v[0], v[1], v[2], v[3]) for k, v in
                ((k, weeks[k]) for k in order)]
        tot_km = sum(d[1] for d in data) / 1000
        tot_h = sum(d[2] for d in data) / 3600
        peak = max(data, key=lambda d: d[1]) if data else None
        self.lbl_overview.configure(
            text=f"{len(data)} settimane · {len(self.df_workouts)} allenamenti · ~{tot_km:.0f} km · ~{tot_h:.0f} h"
                 + (f"\nSettimana piu' carica: {peak[0]} con {peak[1] / 1000:.0f} km" if peak else ""))
        self.chart.set_data(data)

    # ============================================================ Garmin
    def login(self):
        LoginDialog(self)

    def _require_garmin(self) -> bool:
        if not self._require_plan():
            return False
        if self.garmin.client is None:
            if messagebox.askyesno("Garmin", "Non sei connesso a Garmin Connect. Accedere adesso?"):
                dlg = LoginDialog(self)
                self.root.wait_window(dlg)
            if self.garmin.client is None:
                return False
        return self._confirm_discard()

    def _run(self, operation, title):
        loading = LoadingDialog(self.root, title)
        result = {"ok": False, "msg": ""}

        def worker():
            try:
                result["ok"], result["msg"] = operation(loading)
            except Exception as e:  # noqa: BLE001 - mostrato all'utente
                result["ok"], result["msg"] = False, str(e)

        th = threading.Thread(target=worker, daemon=True)
        th.start()
        while th.is_alive():
            self.root.update()
            th.join(timeout=0.05)
        loading.close()
        # salva SEMPRE gli ID ottenuti, anche se l'operazione si e' interrotta a meta'
        self._autosave()
        self._refresh_everything()
        (messagebox.showinfo if result["ok"] else messagebox.showerror)(
            "Garmin" if result["ok"] else "Errore Garmin", result["msg"])

    def _date_str(self, row) -> str:
        d = _to_date(row.get("Date"))
        if not d:
            raise ValueError(f"Data mancante per '{row.get('Description')}'")
        return d.isoformat()

    def upload(self):
        if not self._require_garmin():
            return
        idxs = self._selected_indices()
        if not idxs or not confirm_workouts_valid(self.root, self.df_workouts, idxs, self.df_parameters):
            return
        prefix = self.var_prefix.get()

        def op(loading):
            created = 0
            for n, i in enumerate(idxs, 1):
                loading.update_message(f"Caricamento {n}/{len(idxs)}…")
                row = self.df_workouts.iloc[i]
                if _clean_id(row.get("WorkoutId")):
                    continue
                data = build_garmin_workout_from_excel_row(row, row.get("Steps", ""), self.df_parameters,
                                                           use_prefix=prefix)
                resp = self.garmin.save_workout(data)
                wid = _clean_id((resp or {}).get("workoutId") or (resp or {}).get("workout_id"))
                if wid:
                    self.df_workouts.at[i, "WorkoutId"] = wid
                    created += 1
            skipped = len(idxs) - created
            return True, f"Caricati {created} allenamenti nella libreria Garmin." + (
                f"\n{skipped} erano gia' su Garmin." if skipped else "")

        self._run(op, "Caricamento su Garmin")

    def upload_and_schedule(self):
        if not self._require_garmin():
            return
        idxs = self._selected_indices()
        if not idxs or not confirm_workouts_valid(self.root, self.df_workouts, idxs, self.df_parameters,
                                                  require_date=True):
            return
        prefix = self.var_prefix.get()

        def op(loading):
            done = 0
            for n, i in enumerate(idxs, 1):
                row = self.df_workouts.iloc[i]
                date_str = self._date_str(row)
                wid = _clean_id(row.get("WorkoutId"))
                if not wid:
                    loading.update_message(f"Creazione {n}/{len(idxs)}…")
                    data = build_garmin_workout_from_excel_row(row, row.get("Steps", ""), self.df_parameters,
                                                               use_prefix=prefix)
                    resp = self.garmin.save_workout(data)
                    wid = _clean_id((resp or {}).get("workoutId") or (resp or {}).get("workout_id"))
                    if not wid:
                        continue
                    self.df_workouts.at[i, "WorkoutId"] = wid
                if _clean_id(row.get("WorkoutScheduleId")) and str(row.get("ScheduledDate")) == date_str:
                    continue  # gia' pianificato in quella data
                loading.update_message(f"Pianificazione {n}/{len(idxs)}…")
                resp = self.garmin.schedule_workout(wid, date_str)
                sid = ""
                if isinstance(resp, dict):
                    sid = _clean_id(resp.get("workoutScheduleId") or resp.get("workout_schedule_id")
                                    or resp.get("eventId") or resp.get("id"))
                if sid:
                    self.df_workouts.at[i, "WorkoutScheduleId"] = sid
                self.df_workouts.at[i, "ScheduledDate"] = date_str
                done += 1
            return True, f"Pianificati {done} allenamenti nel calendario Garmin."

        self._run(op, "Caricamento e pianificazione")

    def unschedule(self):
        if not self._require_garmin():
            return
        idxs = self._selected_indices()
        if not idxs:
            return

        def op(loading):
            removed = 0
            for n, i in enumerate(idxs, 1):
                row = self.df_workouts.iloc[i]
                sid = _clean_id(row.get("WorkoutScheduleId"))
                sdate = "" if _is_blank(row.get("ScheduledDate")) else str(row.get("ScheduledDate"))
                if not sid:
                    continue
                loading.update_message(f"Rimozione dal calendario {n}/{len(idxs)}…")
                self.garmin.unschedule_workout(sid, sdate)
                self.df_workouts.at[i, "WorkoutScheduleId"] = ""
                self.df_workouts.at[i, "ScheduledDate"] = ""
                removed += 1
            return True, (f"Tolti {removed} allenamenti dal calendario Garmin "
                          "(restano nella libreria)." if removed else "Nessun allenamento era in calendario.")

        self._run(op, "Rimozione dal calendario")

    def delete_from_garmin(self):
        if not self._require_garmin():
            return
        idxs = self._selected_indices()
        if not idxs:
            return
        if not messagebox.askyesno("Elimina da Garmin",
                                   f"Eliminare DEFINITIVAMENTE {len(idxs)} allenamento/i da Garmin Connect?\n"
                                   "Restano nel tuo piano Excel e potrai ricaricarli."):
            return

        def op(loading):
            deleted = 0
            for n, i in enumerate(idxs, 1):
                row = self.df_workouts.iloc[i]
                wid = _clean_id(row.get("WorkoutId"))
                if not wid:
                    continue
                sid = _clean_id(row.get("WorkoutScheduleId"))
                if sid:
                    loading.update_message(f"Rimozione dal calendario {n}/{len(idxs)}…")
                    try:
                        self.garmin.unschedule_workout(sid, str(row.get("ScheduledDate") or ""))
                    except Exception:  # noqa: BLE001 - la cancellazione del workout rimuove comunque tutto
                        pass
                loading.update_message(f"Eliminazione {n}/{len(idxs)}…")
                self.garmin.delete_workout(wid)
                for col in ("WorkoutId", "WorkoutScheduleId", "ScheduledDate"):
                    self.df_workouts.at[i, col] = ""
                deleted += 1
            return True, f"Eliminati {deleted} allenamenti da Garmin Connect."

        self._run(op, "Eliminazione da Garmin")

    def download(self):
        if self.garmin.client is None:
            if not messagebox.askyesno("Garmin", "Non sei connesso a Garmin Connect. Accedere adesso?"):
                return
            self.root.wait_window(LoginDialog(self))
            if self.garmin.client is None:
                return
        result = show_download_dialog(self.root, self.garmin)
        if result:
            messagebox.showinfo("Download completato", f"Dati salvati in:\n{result}")


def main():
    root = tk.Tk()
    PlannerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
