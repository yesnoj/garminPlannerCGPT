"""
Editor di un allenamento: dati generali, grafico dell'intensita', lista degli step
(con dialog di modifica e modelli rapidi) e testo DSL con evidenziazione e controllo live.
"""
from __future__ import annotations

import datetime as dt
import re
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Callable, Dict, List, Optional

import pandas as pd

import workout_model as wm
from app_theme import STEP_COLORS, Theme, sv_ttk
from tree_dnd import TreeDragDrop


def sv_loaded() -> bool:
    return sv_ttk is not None
from dsl_parser import DSLError, build_garmin_workout_from_excel_row, get_parameter_value

try:
    from tkcalendar import DateEntry
except ImportError:  # pragma: no cover
    DateEntry = None

SPORTS = ["Running", "Cycling", "Swimming"]
DURATION_UNITS = ["min", "sec", "h", "km", "m", "mm:ss", "lap-button"]
_UNIT_ALIASES = {
    "mins": "min", "minuti": "min", "minutes": "min", "minute": "min", "'": "min", "′": "min", "’": "min",
    "s": "sec", "secs": "sec", "secondi": "sec", "seconds": "sec", '"': "sec", "″": "sec",
    "hr": "h", "ore": "h", "ora": "h", "hours": "h", "mt": "m", "metri": "m",
}


def split_duration(text: str):
    """'10min' -> ('10', 'min'); '1:30' -> ('1:30', 'mm:ss'); 'lap-button' -> ('', 'lap-button')."""
    t = (text or "").strip().lower()
    if t.startswith("lap"):
        return "", "lap-button"
    if re.match(r"^\d+:\d{2}(:\d{2})?$", t):
        return t, "mm:ss"
    m = re.match(r"^([\d.,]+)\s*(.*)$", t)
    if not m:
        return t, "min"
    unit = _UNIT_ALIASES.get(m.group(2), m.group(2) or "min")
    return m.group(1), unit if unit in DURATION_UNITS else "min"


def join_duration(value: str, unit: str) -> str:
    if unit == "lap-button":
        return "lap-button"
    if unit == "mm:ss":
        return value.strip()
    return f"{value.strip()}{unit}"


def center_on_parent(win: tk.Toplevel, parent: tk.Misc):
    """Centra un dialog sulla finestra principale."""
    win.update_idletasks()
    top = parent.winfo_toplevel()
    x = top.winfo_rootx() + (top.winfo_width() - win.winfo_reqwidth()) // 2
    y = top.winfo_rooty() + (top.winfo_height() - win.winfo_reqheight()) // 3
    win.geometry(f"+{max(x, 0)}+{max(y, 0)}")
    import app_theme
    if app_theme.CURRENT is not None:
        app_theme.set_windows_titlebar(win, app_theme.CURRENT.mode == "dark")


def _is_blank(v) -> bool:
    try:
        if v is None or pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    return str(v).strip().lower() in ("", "nan", "none", "<na>", "nat")


def _to_date(v) -> Optional[dt.date]:
    if _is_blank(v):
        return None
    try:
        return pd.to_datetime(v).date()
    except (ValueError, TypeError):
        return None


def _to_int(v, default=1) -> int:
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return default


class _Tooltip:
    """Suggerimento al passaggio del mouse."""

    def __init__(self, widget, text, theme):
        self.w, self.text, self.theme, self.tip = widget, text, theme, None
        widget.bind("<Enter>", self._show, add="+")
        widget.bind("<Leave>", self._hide, add="+")

    def _show(self, _e):
        c = self.theme.c
        self.tip = tk.Toplevel(self.w)
        self.tip.wm_overrideredirect(True)
        tk.Label(self.tip, text=self.text, bg=c["card"], fg=c["fg"], padx=6, pady=3,
                 highlightthickness=1, highlightbackground=c["border"], font=self.theme.font_small).pack()
        self.tip.wm_geometry(f"+{self.w.winfo_rootx()}+{self.w.winfo_rooty() + self.w.winfo_height() + 4}")

    def _hide(self, _e):
        if self.tip:
            self.tip.destroy()
            self.tip = None


# ===========================================================================
# Grafico
# ===========================================================================

class WorkoutGraph(tk.Canvas):
    """Profilo dell'allenamento: larghezza = durata, altezza = intensita', colore = tipo di step."""

    PAD_L, PAD_R, PAD_T, PAD_B = 8, 8, 10, 22

    def __init__(self, parent, theme: Theme, height=150, **kw):
        super().__init__(parent, height=height, highlightthickness=0, bd=0, **kw)
        self.theme = theme
        self.segments: List[wm.Segment] = []
        self.highlight: Optional[str] = None
        self._tip = None
        self.bind("<Configure>", lambda e: self.redraw())
        self.bind("<Motion>", self._on_motion)
        self.bind("<Leave>", lambda e: self._hide_tip())
        theme.on_change(self.redraw)

    def set_segments(self, segs: List[wm.Segment], highlight: Optional[str] = None):
        self.segments = segs
        self.highlight = highlight
        self.redraw()

    def redraw(self):
        c = self.theme.c
        self.delete("all")
        self.configure(bg=c["card"])
        w, h = self.winfo_width(), self.winfo_height()
        if w < 50:
            return
        x0, x1 = self.PAD_L, w - self.PAD_R
        y0, y1 = self.PAD_T, h - self.PAD_B
        for frac in (0.25, 0.5, 0.75, 1.0):
            y = y1 - (y1 - y0) * frac
            self.create_line(x0, y, x1, y, fill=c["grid"])
        self.create_line(x0, y1, x1, y1, fill=c["border"])
        total = sum(s.seconds for s in self.segments)
        if total <= 0:
            self.create_text((x0 + x1) / 2, (y0 + y1) / 2, text="Aggiungi degli step per vedere il profilo",
                             fill=c["muted"], font=self.theme.font)
            return
        # tacche temporali
        step_min = 5 if total <= 3600 else 10 if total <= 7200 else 20
        t = step_min * 60
        while t < total:
            x = x0 + (x1 - x0) * t / total
            self.create_line(x, y1, x, y1 + 4, fill=c["muted"])
            self.create_text(x, y1 + 12, text=f"{t // 60}'", fill=c["muted"], font=self.theme.font_small)
            t += step_min * 60
        self._boxes = []
        x = x0
        for s in self.segments:
            wpx = (x1 - x0) * s.seconds / total
            top = y1 - max(6, (y1 - y0) * s.intensity)
            color = STEP_COLORS.get(s.type, "#888")
            outline = c["fg"] if self.highlight and s.label == self.highlight else ""
            rid = self.create_rectangle(x + 0.5, top, x + max(wpx - 0.5, 1.5), y1, fill=color,
                                        outline=outline, width=2 if outline else 0)
            self._boxes.append((x, x + wpx, s))
            x += wpx

    # --- tooltip ---
    def _on_motion(self, e):
        for a, b, s in getattr(self, "_boxes", []):
            if a <= e.x <= b:
                txt = (f"{wm.STEP_LABELS.get(s.type, s.type)} · {wm.fmt_duration(s.seconds)}"
                       f"{' · ' + wm.fmt_km(s.meters) if s.meters else ''}"
                       f"{' (stima)' if s.estimated else ''}\n{s.label}")
                self._show_tip(e.x_root + 12, e.y_root + 12, txt)
                return
        self._hide_tip()

    def _show_tip(self, x, y, text):
        c = self.theme.c
        if self._tip is None:
            self._tip = tk.Toplevel(self)
            self._tip.wm_overrideredirect(True)
            self._tip_lbl = tk.Label(self._tip, justify="left", padx=8, pady=5, font=self.theme.font_small)
            self._tip_lbl.pack()
        self._tip_lbl.configure(text=text, bg=c["card"], fg=c["fg"], highlightthickness=1,
                                highlightbackground=c["border"])
        self._tip.wm_geometry(f"+{x}+{y}")

    def _hide_tip(self):
        if self._tip is not None:
            self._tip.destroy()
            self._tip = None


# ===========================================================================
# Dialog step / ripetizione
# ===========================================================================

class StepDialog(tk.Toplevel):
    """Crea o modifica uno step. Ritorna uno wm.Step in self.result (None se annullato)."""

    def __init__(self, parent, theme: Theme, step: Optional[wm.Step], params, sport="running"):
        super().__init__(parent)
        self.theme = theme
        self.params = params
        self.result: Optional[wm.Step] = None
        self.title("Modifica step" if step else "Nuovo step")
        self.transient(parent)
        self.resizable(False, False)
        step = step or wm.Step("interval", "5min", "")
        self._comment = step.comment

        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)

        ttk.Label(body, text="Tipo", style="H2.TLabel").grid(row=0, column=0, sticky="w")
        self.var_type = tk.StringVar(value=step.type if step.type in wm.STEP_TYPES else "interval")
        types = ttk.Frame(body)
        types.grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 14))
        for t in wm.STEP_TYPES:
            ttk.Radiobutton(types, text=wm.STEP_LABELS[t], value=t, variable=self.var_type,
                            style="Toggle.TButton",
                            command=self._update).pack(side="left", padx=(0, 4))

        ttk.Label(body, text="Durata o distanza", style="H2.TLabel").grid(row=2, column=0, sticky="w")
        val, unit = split_duration(step.duration)
        self.var_val = tk.StringVar(value=val)
        self.var_unit = tk.StringVar(value=unit)
        dur = ttk.Frame(body)
        dur.grid(row=3, column=0, columnspan=3, sticky="w", pady=(4, 14))
        self.ent_val = ttk.Entry(dur, textvariable=self.var_val, width=10)
        self.ent_val.pack(side="left")
        cb_unit = ttk.Combobox(dur, textvariable=self.var_unit, values=DURATION_UNITS, width=11, state="readonly")
        cb_unit.pack(side="left", padx=6)
        cb_unit.bind("<<ComboboxSelected>>", lambda e: self._update())
        ttk.Label(dur, text="es. 10 min · 400 m · 1,5 km · 1:30", style="Small.TLabel").pack(side="left", padx=6)

        ttk.Label(body, text="Target", style="H2.TLabel").grid(row=4, column=0, sticky="w")
        self.var_target = tk.StringVar(value=step.target)
        tgt = ttk.Frame(body)
        tgt.grid(row=5, column=0, columnspan=3, sticky="we", pady=(4, 2))
        self.cb_target = ttk.Combobox(tgt, textvariable=self.var_target, values=self._target_choices(), width=24)
        self.cb_target.pack(side="left")
        self.cb_target.bind("<<ComboboxSelected>>", lambda e: self._update())
        self.lbl_target_info = ttk.Label(tgt, text="", style="Small.TLabel")
        self.lbl_target_info.pack(side="left", padx=8)
        ttk.Label(body, text="Vuoto = nessun target · 5:00 · 4:50-5:00 · Z1-Z5 · HR_Z2 · 140-155 · chiavi del foglio Parametri",
                  style="Small.TLabel").grid(row=6, column=0, columnspan=3, sticky="w", pady=(0, 14))

        prev = ttk.Frame(body)
        prev.grid(row=7, column=0, columnspan=3, sticky="we")
        ttk.Label(prev, text="Riga DSL:", style="Muted.TLabel").pack(side="left")
        self.lbl_preview = ttk.Label(prev, text="", font=theme.font_mono)
        self.lbl_preview.pack(side="left", padx=6)
        self.lbl_status = ttk.Label(body, text="")
        self.lbl_status.grid(row=8, column=0, columnspan=3, sticky="w", pady=(4, 0))

        btns = ttk.Frame(body)
        btns.grid(row=9, column=0, columnspan=3, sticky="e", pady=(18, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right")
        self.btn_ok = ttk.Button(btns, text="OK", style="Accent.TButton", command=self._ok)
        self.btn_ok.pack(side="right", padx=8)

        for v in (self.var_val, self.var_target):
            v.trace_add("write", lambda *a: self._update())
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        self._update()
        self.ent_val.focus_set()
        self.ent_val.select_range(0, "end")
        center_on_parent(self, parent)
        self.grab_set()
        self.wait_window(self)

    def _target_choices(self) -> List[str]:
        items = ["", "Z1", "Z2", "Z3", "Z4", "Z5", "HR_Z1", "HR_Z2", "HR_Z3", "HR_Z4", "HR_Z5"]
        if self.params is not None and "Key" in self.params.columns:
            skip = {"pace_tolerance", "hr_tolerance", "hr_max", "swim_tolerance", "ftp"}
            for k in self.params["Key"].dropna().astype(str):
                k = k.strip()
                if k and k.lower() not in skip and k not in items:
                    items.append(k)
        return items

    def _step(self) -> wm.Step:
        unit = self.var_unit.get()
        self.ent_val.configure(state="disabled" if unit == "lap-button" else "normal")
        return wm.Step(self.var_type.get(), join_duration(self.var_val.get().replace(",", "."), unit),
                       self.var_target.get().strip(), self._comment)

    def _update(self):
        st = self._step()
        self.lbl_preview.configure(text=st.to_line())
        err = wm.validate_line(st.to_line(), self.params)
        c = self.theme.c
        if err:
            self.lbl_status.configure(text=f"✗ {err}", foreground=c["err"])
            self.btn_ok.state(["disabled"])
        else:
            self.lbl_status.configure(text="✓ Step valido", foreground=c["ok"])
            self.btn_ok.state(["!disabled"])
        expr = get_parameter_value(st.target, self.params) if st.target else None
        self.lbl_target_info.configure(text=f"= {expr}" if expr else "")

    def _ok(self):
        st = self._step()
        if wm.validate_line(st.to_line(), self.params):
            return
        self.result = st
        self.destroy()


class RepeatDialog(tk.Toplevel):
    def __init__(self, parent, count: int = 4, title="Ripetizione"):
        super().__init__(parent)
        self.result: Optional[int] = None
        self.title(title)
        self.transient(parent)
        self.resizable(False, False)
        body = ttk.Frame(self, padding=18)
        body.pack()
        ttk.Label(body, text="Numero di ripetizioni").grid(row=0, column=0, sticky="w")
        self.var = tk.IntVar(value=count)
        sp = ttk.Spinbox(body, from_=1, to=99, textvariable=self.var, width=6)
        sp.grid(row=0, column=1, padx=10)
        btns = ttk.Frame(body)
        btns.grid(row=1, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="OK", style="Accent.TButton", command=self._ok).pack(side="right", padx=8)
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        sp.focus_set()
        center_on_parent(self, parent)
        self.grab_set()
        self.wait_window(self)

    def _ok(self):
        try:
            n = int(self.var.get())
        except (tk.TclError, ValueError):
            return
        if n >= 1:
            self.result = n
            self.destroy()


# ===========================================================================
# Editor
# ===========================================================================

class WorkoutEditor(ttk.Frame):
    """
    Pannello di modifica di un allenamento.
    on_save(values) riceve un dict con Description, Date, Week, Session, Sport, Steps.
    """

    def __init__(self, parent, theme: Theme, get_params: Callable[[], Optional[pd.DataFrame]],
                 on_save: Callable[[Dict], None], on_dirty: Callable[[bool], None] = lambda d: None):
        super().__init__(parent, padding=(16, 12))
        self.theme = theme
        self.get_params = get_params
        self.on_save = on_save
        self.on_dirty = on_dirty
        self.nodes: List[wm.Node] = []
        self._orig: Optional[Dict] = None
        self._dirty = False
        self._loading = False
        self._text_job = None
        self._iid_to_id: Dict[str, int] = {}
        self._build()
        theme.on_change(self._retheme)
        # messaggio quando nessun allenamento e' selezionato
        self.empty = ttk.Frame(self, padding=40)
        ttk.Label(self.empty, text="Nessun allenamento selezionato", style="H2.TLabel").pack(pady=(80, 6))
        ttk.Label(self.empty, text="Scegli un allenamento dalla lista per vederlo e modificarlo,\n"
                                   "oppure creane uno nuovo con '+ Allenamento'.",
                  style="Muted.TLabel", justify="center").pack()
        self.load(None)

    # ------------------------------------------------------------------ UI
    def _build(self):
        t = self.theme
        # --- intestazione
        self.var_desc = tk.StringVar()
        self.ent_title = ttk.Entry(self, textvariable=self.var_desc, font=t.font_h2)
        self.ent_title.pack(fill="x")
        meta = ttk.Frame(self)
        meta.pack(fill="x", pady=(10, 0))
        ttk.Label(meta, text="Data").grid(row=0, column=0, sticky="w")
        if DateEntry is not None:
            self.ent_date = DateEntry(meta, width=11, date_pattern="dd/mm/yyyy", locale="it_IT",
                                      firstweekday="monday", showweeknumbers=False)
            t.register_date_entry(self.ent_date)
        else:
            self.ent_date = ttk.Entry(meta, width=12)
        self.ent_date.grid(row=0, column=1, sticky="w", padx=(6, 16))
        ttk.Label(meta, text="Sport").grid(row=0, column=2, sticky="w")
        self.var_sport = tk.StringVar(value="Running")
        cb = ttk.Combobox(meta, textvariable=self.var_sport, values=SPORTS, width=10, state="readonly")
        cb.grid(row=0, column=3, sticky="w", padx=(6, 16))
        cb.bind("<<ComboboxSelected>>", lambda e: (self._build_presets_menu(), self._refresh_all(), self._mark_dirty()))
        ttk.Label(meta, text="Sett.").grid(row=0, column=4, sticky="w")
        self.var_week = tk.StringVar()
        ttk.Spinbox(meta, from_=0, to=99, width=3, textvariable=self.var_week).grid(row=0, column=5, sticky="w",
                                                                                   padx=(6, 16))
        ttk.Label(meta, text="Sess.").grid(row=0, column=6, sticky="w")
        self.var_session = tk.StringVar()
        ttk.Spinbox(meta, from_=0, to=20, width=3, textvariable=self.var_session).grid(row=0, column=7, sticky="w",
                                                                                      padx=(6, 0))

        # --- statistiche
        stats = ttk.Frame(self)
        stats.pack(fill="x", pady=(14, 6))
        self.lbl_dur = self._stat(stats, "Durata stimata")
        self.lbl_dist = self._stat(stats, "Distanza stimata")
        self.lbl_nsteps = self._stat(stats, "Step")
        self.lbl_valid = ttk.Label(self, text="", style="Ok.TLabel", justify="left")
        self.lbl_valid.pack(fill="x", pady=(0, 8))
        self.lbl_valid.bind("<Configure>", lambda e: self.lbl_valid.configure(wraplength=max(200, e.width - 4)))

        # --- grafico
        frame_graph = ttk.Frame(self)
        frame_graph.pack(fill="x")
        self.graph = WorkoutGraph(frame_graph, t, height=150)
        self.graph.pack(fill="x")
        legend = ttk.Frame(self)
        legend.pack(fill="x", pady=(4, 8))
        for k in wm.STEP_TYPES:
            sw = tk.Canvas(legend, width=10, height=10, highlightthickness=0, bd=0)
            sw.create_rectangle(0, 0, 10, 10, fill=STEP_COLORS[k], outline="")
            sw.pack(side="left", padx=(0, 4))
            ttk.Label(legend, text=wm.STEP_LABELS[k], style="Small.TLabel").pack(side="left", padx=(0, 12))
            sw.configure(bg=t.c["bg"])
            setattr(self, f"_sw_{k}", sw)

        # --- pulsanti
        bottom = ttk.Frame(self)
        bottom.pack(side="bottom", fill="x", pady=(10, 0))
        self.lbl_dirty = ttk.Label(bottom, text="", style="Muted.TLabel")
        self.lbl_dirty.pack(side="left")
        self.btn_save = ttk.Button(bottom, text="Salva allenamento", style="Accent.TButton", command=self.save)
        self.btn_save.pack(side="right")
        self.btn_revert = ttk.Button(bottom, text="Annulla modifiche", command=self.revert)
        self.btn_revert.pack(side="right", padx=8)

        # --- schede step / testo
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)
        self._build_steps_tab()
        self._build_text_tab()
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self._on_tab())

        for v in (self.var_desc, self.var_week, self.var_session):
            v.trace_add("write", lambda *a: self._mark_dirty())
        self.ent_date.bind("<<DateEntrySelected>>", lambda e: self._mark_dirty())
        self.ent_date.bind("<KeyRelease>", lambda e: self._mark_dirty())

    def _stat(self, parent, label):
        box = ttk.Frame(parent)
        box.pack(side="left", padx=(0, 22))
        ttk.Label(box, text=label, style="Small.TLabel").pack(anchor="w")
        val = ttk.Label(box, text="–", style="Stat.TLabel")
        val.pack(anchor="w")
        return val

    def _build_steps_tab(self):
        tab = ttk.Frame(self.nb, padding=(0, 8, 0, 0))
        self.nb.add(tab, text="  Step  ")
        bar = ttk.Frame(tab)
        bar.pack(fill="x", pady=(0, 6))
        ttk.Button(bar, text="+ Step", style="Accent.TButton", command=self.add_step).pack(side="left")
        ttk.Button(bar, text="+ Ripetizione", command=self.add_repeat).pack(side="left", padx=4)
        self.mb_presets = ttk.Menubutton(bar, text="Modelli rapidi")
        self.mb_presets.pack(side="left", padx=4)
        self.menu_presets = tk.Menu(self.mb_presets, tearoff=False)
        self.mb_presets["menu"] = self.menu_presets
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=8)
        self._tool(bar, "Modifica", self.edit_selected)
        self._tool(bar, "Duplica", self.duplicate_selected)
        self._tool(bar, "Elimina", self.delete_selected)
        # spostamenti: a destra, compatti, con suggerimento
        for txt, cmd, tip in (("⇤", self._outdent, "Porta fuori dalla ripetizione"),
                              ("⇥", self._indent, "Metti dentro la ripetizione precedente"),
                              ("↓", lambda: self._move(+1), "Sposta giu'"),
                              ("↑", lambda: self._move(-1), "Sposta su")):
            b = ttk.Button(bar, text=txt, width=3, command=cmd)
            b.pack(side="right", padx=1)
            _Tooltip(b, tip, self.theme)

        cols = ("dur", "target", "stima")
        wrap = ttk.Frame(tab)
        wrap.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(wrap, columns=cols, selectmode="browse")
        self.tree.heading("#0", text="Step")
        self.tree.heading("dur", text="Durata")
        self.tree.heading("target", text="Target")
        self.tree.heading("stima", text="Stima")
        self.tree.column("#0", width=220, stretch=True)
        self.tree.column("dur", width=110, anchor="w")
        self.tree.column("target", width=150, anchor="w")
        self.tree.column("stima", width=120, anchor="w")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", lambda e: None if self.dnd.recently_dropped() else self.edit_selected())
        self.tree.bind("<Return>", lambda e: self.edit_selected())
        self.tree.bind("<Delete>", lambda e: self.delete_selected())
        self.tree.bind("<BackSpace>", lambda e: self.delete_selected())
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._refresh_graph())
        # trascinamento con il mouse: riordina gli step, anche dentro/fuori dalle ripetizioni
        self.dnd = TreeDragDrop(
            self.tree, self.theme, on_drop=self._on_step_drop,
            allow_inside=lambda iid: isinstance(wm.get(self.nodes, self._iid_to_id.get(iid, -1)), wm.Repeat),
            can_drop=self._can_step_drop)
        ttk.Label(tab, text="Trascina gli step col mouse per riordinarli · rilascia al centro di una "
                            "ripetizione per metterli dentro · doppio clic per modificare",
                  style="Small.TLabel").pack(anchor="w", pady=(6, 0))
        self._tag_colors()

    def _tool(self, parent, text, cmd, width=None):
        b = ttk.Button(parent, text=text, command=cmd, width=width)
        b.pack(side="left", padx=2)
        return b

    def _build_text_tab(self):
        tab = ttk.Frame(self.nb, padding=(0, 8, 0, 0))
        self.nb.add(tab, text="  Testo DSL  ")
        wrap = ttk.Frame(tab)
        wrap.pack(fill="both", expand=True)
        self.text = tk.Text(wrap, wrap="none", undo=True, font=self.theme.font_mono, height=10,
                            padx=10, pady=8, relief="flat", borderwidth=0)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=sb.set)
        self.text.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.text.bind("<<Modified>>", self._on_text_modified)
        ttk.Label(tab, text="Una riga per step: tipo: durata @ target  ·  'repeat N:' con gli step rientrati sotto",
                  style="Small.TLabel").pack(anchor="w", pady=(6, 0))
        self._retheme_text()

    def _tag_colors(self):
        c = self.theme.c
        for k, col in c["tint"].items():
            self.tree.tag_configure(k, background=col)
        self.tree.tag_configure("repeat", font=self.theme.font_bold)
        self.tree.tag_configure("invalid", foreground=c["err"])

    def _retheme_text(self):
        c = self.theme.c
        self.text.configure(bg=c["card"], fg=c["fg"], insertbackground=c["fg"],
                            selectbackground=c["sel"], selectforeground=c["fg"])
        for k, col in STEP_COLORS.items():
            self.text.tag_configure(f"kw_{k}", foreground=col, font=self.theme.font_mono + ("bold",))
        self.text.tag_configure("kw_repeat", foreground=c["accent"], font=self.theme.font_mono + ("bold",))
        self.text.tag_configure("target", foreground=c["info"])
        self.text.tag_configure("comment", foreground=c["muted"])
        self.text.tag_configure("errline", background="#5a1d1d" if self.theme.mode == "dark" else "#fde2e2")

    def _retheme(self):
        self._tag_colors()
        self._retheme_text()
        for k in wm.STEP_TYPES:
            getattr(self, f"_sw_{k}").configure(bg=self.theme.c["bg"])
        self._validate()

    # ------------------------------------------------------------- dati
    def load(self, values: Optional[Dict]):
        """Carica un allenamento (dict con le colonne della riga) o svuota l'editor (None)."""
        self._loading = True
        self._orig = dict(values) if values else None
        if values:
            self.empty.place_forget()
        else:
            self.empty.place(relx=0, rely=0, relwidth=1, relheight=1)
            self.empty.lift()
        state = "!disabled" if values else "disabled"
        for w in (self.btn_save, self.btn_revert, self.ent_title):
            w.state([state])
        v = values or {}
        self.var_desc.set("" if _is_blank(v.get("Description")) else str(v.get("Description")))
        d = _to_date(v.get("Date"))
        if DateEntry is not None:
            if d:
                self.ent_date.set_date(d)
            else:
                self.ent_date.delete(0, "end")
        else:
            self.ent_date.delete(0, "end")
            if d:
                self.ent_date.insert(0, d.strftime("%d/%m/%Y"))
        self.var_week.set("" if _is_blank(v.get("Week")) else str(_to_int(v.get("Week"))))
        self.var_session.set("" if _is_blank(v.get("Session")) else str(_to_int(v.get("Session"))))
        sport = str(v.get("Sport") or "Running")
        self.var_sport.set({"cycling": "Cycling", "swimming": "Swimming"}.get(wm.sport_key(sport), "Running"))
        steps = "" if _is_blank(v.get("Steps")) else str(v.get("Steps"))
        try:
            self.nodes = wm.from_dsl(steps) if steps.strip() else []
            self._set_text(wm.to_dsl(self.nodes))
        except DSLError:
            self.nodes = []
            self._set_text(steps)  # l'utente sistema nel testo
            if values:
                self.nb.select(1)
        self._build_presets_menu()
        self._refresh_tree()
        self._refresh_graph()
        self._validate()
        self._loading = False
        self._set_dirty(False)

    def get_values(self) -> Dict:
        if DateEntry is not None:
            try:
                d = self.ent_date.get_date() if self.ent_date.get().strip() else None
            except ValueError:
                d = None
        else:
            try:
                d = dt.datetime.strptime(self.ent_date.get().strip(), "%d/%m/%Y").date()
            except ValueError:
                d = None
        return {
            "Description": self.var_desc.get().strip(),
            "Date": d,
            "Week": _to_int(self.var_week.get(), None) if self.var_week.get().strip() else None,
            "Session": _to_int(self.var_session.get(), None) if self.var_session.get().strip() else None,
            "Sport": self.var_sport.get(),
            "Steps": self._current_steps_text(),
        }

    def _current_steps_text(self) -> str:
        return self.text.get("1.0", "end").rstrip("\n")

    @property
    def dirty(self) -> bool:
        return self._dirty

    def _mark_dirty(self):
        if not self._loading and self._orig is not None:
            self._set_dirty(True)

    def _set_dirty(self, value: bool):
        self._dirty = value
        self.lbl_dirty.configure(text="● Modifiche non salvate" if value else "")
        self.on_dirty(value)

    def save(self):
        if self._orig is None:
            return
        err = self._validate()
        if err and not messagebox.askyesno(
                "Allenamento non valido",
                f"L'allenamento contiene un errore:\n\n{err}\n\nSalvarlo comunque? "
                "(non potra' essere caricato su Garmin finche' non lo correggi)", parent=self):
            return
        values = self.get_values()
        self.on_save(values)
        self._orig = dict(values)
        self._set_dirty(False)

    def revert(self):
        if self._orig is not None:
            self.load(self._orig)

    # ------------------------------------------------------ refresh
    def _refresh_all(self, source: str = "model"):
        if source == "model":
            self._set_text(wm.to_dsl(self.nodes))
        self._refresh_tree()
        self._refresh_graph()
        self._validate()

    def _refresh_tree(self, select_id: Optional[int] = None):
        if select_id is None:
            sel = self._selected()
            select_id = sel.id if sel else None
        self.tree.delete(*self.tree.get_children())
        self._iid_to_id.clear()
        params = self.get_params()
        sport = self.var_sport.get()

        def add(parent, nodes):
            for n in nodes:
                if isinstance(n, wm.Repeat):
                    segs = wm.segments([n], sport, params)
                    secs, meters, _ = wm.totals(segs)
                    iid = self.tree.insert(parent, "end", text=f"  Ripeti {n.count}×",
                                           values=("", "", self._est(secs, meters)),
                                           open=True, tags=("repeat",))
                    self._iid_to_id[iid] = n.id
                    add(iid, n.children)
                else:
                    segs = wm.segments([n], sport, params)
                    secs, meters, _ = wm.totals(segs)
                    bad = wm.validate_line(n.to_line(), params)
                    tags = (n.type,) + (("invalid",) if bad else ())
                    iid = self.tree.insert(parent, "end",
                                           text=f"  {wm.STEP_LABELS.get(n.type, n.type)}",
                                           values=(n.duration, n.target or "—", self._est(secs, meters) if not bad else "errore"),
                                           tags=tags)
                    self._iid_to_id[iid] = n.id
        add("", self.nodes)
        if select_id is not None:
            for iid, nid in self._iid_to_id.items():
                if nid == select_id:
                    self.tree.selection_set(iid)
                    self.tree.see(iid)
                    break

    @staticmethod
    def _est(secs, meters):
        parts = [wm.fmt_duration(secs)] if secs else []
        if meters:
            parts.append(wm.fmt_km(meters))
        return " · ".join(parts)

    def _refresh_graph(self):
        params = self.get_params()
        sport = self.var_sport.get()
        segs = wm.segments(self.nodes, sport, params)
        sel = self._selected()
        self.graph.set_segments(segs, sel.to_line() if isinstance(sel, wm.Step) else None)
        secs, meters, approx = wm.totals(segs)
        tilde = "~" if approx else ""
        self.lbl_dur.configure(text=f"{tilde}{wm.fmt_duration(secs)}" if secs else "–")
        self.lbl_dist.configure(text=f"{tilde}{wm.fmt_km(meters)}" if meters else "–")
        self.lbl_nsteps.configure(text=str(len(list(wm.iter_steps(self.nodes)))))

    def _validate(self) -> Optional[str]:
        """Valida tutto l'allenamento come fara' l'upload. Ritorna il messaggio d'errore o None."""
        self.text.tag_remove("errline", "1.0", "end")
        if self._orig is None:
            self.lbl_valid.configure(text="Seleziona un allenamento dalla lista", style="Muted.TLabel")
            return None
        v = self.get_values()
        steps = v["Steps"]
        err = None
        try:
            build_garmin_workout_from_excel_row(pd.Series(v), steps, self.get_params())
        except DSLError as e:
            err = str(e)
            if e.line_no:
                self.text.tag_add("errline", f"{e.line_no}.0", f"{e.line_no}.end+1c")
        if err:
            self.lbl_valid.configure(text=f"✗ {err}", style="Err.TLabel")
        else:
            self.lbl_valid.configure(text="✓ Pronto per Garmin", style="Ok.TLabel")
        return err

    # ------------------------------------------------------ testo DSL
    def _set_text(self, s: str):
        self._suppress_text = True
        self.text.delete("1.0", "end")
        self.text.insert("1.0", s)
        self.text.edit_modified(False)
        self._highlight()
        self._suppress_text = False

    def _on_text_modified(self, _e=None):
        if not self.text.edit_modified():
            return
        self.text.edit_modified(False)
        if getattr(self, "_suppress_text", False):
            return
        self._highlight()
        self._mark_dirty()
        if self._text_job:
            self.after_cancel(self._text_job)
        self._text_job = self.after(350, self._text_changed)

    def _text_changed(self):
        self._text_job = None
        try:
            self.nodes = wm.from_dsl(self._current_steps_text())
            self._refresh_tree()
            self._refresh_graph()
        except DSLError:
            pass
        self._validate()

    def _highlight(self):
        t = self.text
        for tag in [f"kw_{k}" for k in STEP_COLORS] + ["kw_repeat", "target", "comment"]:
            t.tag_remove(tag, "1.0", "end")
        lines = t.get("1.0", "end").split("\n")
        for i, line in enumerate(lines, start=1):
            m = re.match(r"^(\s*)(repeat\s+\d+\s*:?)", line, re.I)
            if m:
                t.tag_add("kw_repeat", f"{i}.{len(m.group(1))}", f"{i}.{m.end(2)}")
                continue
            m = re.match(r"^(\s*)([A-Za-z]+)\s*:", line)
            if m:
                k = wm._ALIASES.get(m.group(2).lower(), m.group(2).lower())
                if k in STEP_COLORS:
                    t.tag_add(f"kw_{k}", f"{i}.{len(m.group(1))}", f"{i}.{m.end(2)}")
            at = line.find("@")
            cm = line.find("--")
            if at >= 0:
                t.tag_add("target", f"{i}.{at}", f"{i}.{cm if cm > at else len(line)}")
            if cm >= 0:
                t.tag_add("comment", f"{i}.{cm}", f"{i}.{len(line)}")

    def _on_tab(self):
        # passando alla lista, riallinea il modello al testo (se valido)
        if self.nb.index("current") == 0:
            try:
                self.nodes = wm.from_dsl(self._current_steps_text())
            except DSLError:
                pass
            self._refresh_tree()
            self._refresh_graph()

    # ------------------------------------------------------ azioni step
    def _selected(self) -> Optional[wm.Node]:
        sel = self.tree.selection()
        if not sel:
            return None
        return wm.get(self.nodes, self._iid_to_id.get(sel[0], -1))

    def _require_loaded(self) -> bool:
        if self._orig is None:
            return False
        if self.nb.index("current") == 1:
            try:
                self.nodes = wm.from_dsl(self._current_steps_text())
            except DSLError as e:
                messagebox.showerror("Testo non valido", f"Correggi prima il testo DSL:\n{e}", parent=self)
                return False
        return True

    def _changed(self, select: Optional[wm.Node] = None):
        self._set_text(wm.to_dsl(self.nodes))
        self._refresh_tree(select.id if select else None)
        self._refresh_graph()
        self._validate()
        self._mark_dirty()

    def add_step(self):
        if not self._require_loaded():
            return
        dlg = StepDialog(self.winfo_toplevel(), self.theme, None, self.get_params(), self.var_sport.get())
        if dlg.result:
            sel = self._selected()
            wm.insert_after(self.nodes, sel.id if sel else None, dlg.result, inside=isinstance(sel, wm.Repeat))
            self._changed(dlg.result)

    def add_repeat(self):
        if not self._require_loaded():
            return
        dlg = RepeatDialog(self.winfo_toplevel(), 4, "Nuova ripetizione")
        if dlg.result:
            rep = wm.Repeat(dlg.result, [wm.Step("interval", "400m", "Z5"), wm.Step("recovery", "90sec", "Z1")])
            sel = self._selected()
            wm.insert_after(self.nodes, sel.id if sel else None, rep)
            self._changed(rep)

    def insert_preset(self, factory):
        if not self._require_loaded():
            return
        sel = self._selected()
        ref = sel.id if sel else None
        last = None
        for n in factory():
            wm.insert_after(self.nodes, ref, n, inside=isinstance(sel, wm.Repeat) and last is None)
            ref = n.id
            last = n
            sel = None
        self._changed(last)

    def edit_selected(self):
        if not self._require_loaded():
            return
        n = self._selected()
        if n is None:
            return
        if isinstance(n, wm.Repeat):
            dlg = RepeatDialog(self.winfo_toplevel(), n.count)
            if dlg.result:
                n.count = dlg.result
                self._changed(n)
        else:
            dlg = StepDialog(self.winfo_toplevel(), self.theme, n, self.get_params(), self.var_sport.get())
            if dlg.result:
                n.type, n.duration, n.target, n.comment = (dlg.result.type, dlg.result.duration,
                                                          dlg.result.target, dlg.result.comment)
                self._changed(n)

    def duplicate_selected(self):
        if not self._require_loaded():
            return
        n = self._selected()
        if n is None:
            return
        new = n.copy()
        wm.insert_after(self.nodes, n.id, new)
        self._changed(new)

    def delete_selected(self):
        if not self._require_loaded():
            return
        n = self._selected()
        if n is None:
            return
        if isinstance(n, wm.Repeat) and n.children and not messagebox.askyesno(
                "Elimina ripetizione", f"Eliminare la ripetizione e i suoi {len(n.children)} step?", parent=self):
            return
        lst, i = wm.find(self.nodes, n.id)
        wm.remove(self.nodes, n.id)
        nxt = lst[min(i, len(lst) - 1)] if lst else None
        self._changed(nxt)

    def _can_step_drop(self, src, target, where) -> bool:
        node = wm.get(self.nodes, self._iid_to_id.get(src, -1))
        return node is not None and not wm.contains(node, self._iid_to_id.get(target, -1))

    def _on_step_drop(self, src, target, where):
        if not self._require_loaded():
            return
        sid, tid = self._iid_to_id.get(src), self._iid_to_id.get(target)
        node = wm.get(self.nodes, sid)
        if node is not None and wm.move_to(self.nodes, sid, tid, where):
            self._changed(node)

    def _move(self, delta):
        if not self._require_loaded():
            return
        n = self._selected()
        if n and wm.move(self.nodes, n.id, delta):
            self._changed(n)

    def _indent(self):
        if not self._require_loaded():
            return
        n = self._selected()
        if n and wm.indent(self.nodes, n.id):
            self._changed(n)

    def _outdent(self):
        if not self._require_loaded():
            return
        n = self._selected()
        if n and wm.outdent(self.nodes, n.id):
            self._changed(n)

    def _build_presets_menu(self):
        self.menu_presets.delete(0, "end")
        sport = wm.sport_key(self.var_sport.get())
        for name, factory in wm.PRESETS.get(sport, []):
            self.menu_presets.add_command(label=name, command=lambda f=factory: self.insert_preset(f))
