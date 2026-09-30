"""Tema grafico dell'app: Sun Valley (sv-ttk) chiaro/scuro + palette colori condivisa."""
import tkinter as tk
from tkinter import ttk, font as tkfont

try:
    import sv_ttk
except ImportError:  # l'app funziona anche senza, con il tema ttk di sistema
    sv_ttk = None

try:
    import darkdetect
except ImportError:
    darkdetect = None


PALETTES = {
    "light": {
        "bg": "#fafafa", "card": "#ffffff", "fg": "#1c1c1c", "muted": "#6b7280",
        "border": "#e3e5e8", "grid": "#eef0f3", "accent": "#005fb8",
        "ok": "#15803d", "err": "#b91c1c", "warn": "#b45309", "info": "#1d4ed8",
        "sel": "#dbeafe", "week_bg": "#f1f3f6",
        "tint": {"warmup": "#fff7e6", "interval": "#fdecec", "recovery": "#eaf2fe",
                 "rest": "#f3f4f6", "cooldown": "#e7f7f5", "repeat": "#f4f4f6"},
        "status": {"error": "#b91c1c", "scheduled": "#15803d", "uploaded": "#1d4ed8", "draft": "#6b7280"},
    },
    "dark": {
        "bg": "#1c1c1c", "card": "#262626", "fg": "#f5f5f5", "muted": "#9ca3af",
        "border": "#3a3a3a", "grid": "#303030", "accent": "#57c8ff",
        "ok": "#4ade80", "err": "#f87171", "warn": "#fbbf24", "info": "#60a5fa",
        "sel": "#1e3a5f", "week_bg": "#2a2a2a",
        "tint": {"warmup": "#3a2e17", "interval": "#3b1f1f", "recovery": "#1c2a40",
                 "rest": "#2b2b2b", "cooldown": "#163330", "repeat": "#2d2d30"},
        "status": {"error": "#f87171", "scheduled": "#4ade80", "uploaded": "#60a5fa", "draft": "#9ca3af"},
    },
}

STEP_COLORS = {
    "warmup": "#f59e0b",
    "interval": "#ef4444",
    "recovery": "#3b82f6",
    "rest": "#9ca3af",
    "cooldown": "#14b8a6",
}


CURRENT = None  # tema attivo (usato dai dialog che non ricevono il tema come parametro)


def calendar_options(c: dict) -> dict:
    """Colori del calendario a comparsa di tkcalendar (che non segue il tema ttk)."""
    return dict(
        background=c["card"], foreground=c["fg"], bordercolor=c["border"],
        headersbackground=c["week_bg"], headersforeground=c["fg"],
        normalbackground=c["card"], normalforeground=c["fg"],
        weekendbackground=c["card"], weekendforeground=c["fg"],
        othermonthbackground=c["bg"], othermonthforeground=c["muted"],
        othermonthwebackground=c["bg"], othermonthweforeground=c["muted"],
        selectbackground=c["accent"], selectforeground="#ffffff" if c is PALETTES["light"] else "#000000",
        disabledbackground=c["bg"], disabledforeground=c["muted"],
        tooltipbackground=c["card"], tooltipforeground=c["fg"],
    )


def style_date_entry(widget) -> None:
    """Applica i colori del tema a un DateEntry (anche ai cambi di tema successivi)."""
    if CURRENT is None:
        return
    CURRENT.register_date_entry(widget)


def _fallback_style(style: ttk.Style, c: dict, font, font_bold):
    """
    Tema chiaro/scuro completo costruito sul tema 'clam' di Tk.
    Usato quando sv-ttk non e' installato: tutti i widget ttk prendono i colori della palette.
    """
    if style.theme_use() != "clam":
        style.theme_use("clam")
    bg, card, fg, muted, border = c["bg"], c["card"], c["fg"], c["muted"], c["border"]
    accent, sel = c["accent"], c["sel"]
    hover = c["week_bg"]
    on_accent = "#ffffff" if c is PALETTES["light"] else "#0b1a24"
    style.configure(".", background=bg, foreground=fg, fieldbackground=card, bordercolor=border,
                    darkcolor=bg, lightcolor=bg, troughcolor=bg, focuscolor=accent,
                    selectbackground=sel, selectforeground=fg, insertcolor=fg, font=font,
                    arrowcolor=fg)
    style.map(".", foreground=[("disabled", muted)])
    style.configure("TFrame", background=bg)
    style.configure("TLabel", background=bg, foreground=fg)
    style.configure("TLabelframe", background=bg, bordercolor=border)
    style.configure("TLabelframe.Label", background=bg, foreground=fg)
    style.configure("TSeparator", background=border)
    style.configure("TPanedwindow", background=bg)
    style.configure("Sash", sashthickness=6, gripcount=0, background=bg)
    # pulsanti
    style.configure("TButton", background=card, foreground=fg, bordercolor=border, relief="solid",
                    borderwidth=1, padding=(8, 3), width=0, lightcolor=card, darkcolor=card)
    style.map("TButton", background=[("disabled", bg), ("pressed", sel), ("active", hover)],
              foreground=[("disabled", muted)], bordercolor=[("focus", accent)])
    style.configure("Accent.TButton", background=accent, foreground=on_accent, bordercolor=accent,
                    lightcolor=accent, darkcolor=accent)
    style.map("Accent.TButton", background=[("disabled", border), ("pressed", accent), ("active", accent)],
              foreground=[("disabled", muted)])
    style.configure("Toggle.TButton", background=card, foreground=fg, bordercolor=border, relief="solid",
                    borderwidth=1, padding=(8, 3))
    style.map("Toggle.TButton", background=[("selected", accent), ("active", hover)],
              foreground=[("selected", on_accent)])
    style.configure("TMenubutton", background=card, foreground=fg, bordercolor=border, arrowcolor=fg,
                    relief="solid", borderwidth=1, padding=(8, 3))
    style.map("TMenubutton", background=[("active", hover)])
    for name in ("TCheckbutton", "Switch.TCheckbutton", "TRadiobutton"):
        style.configure(name, background=bg, foreground=fg, indicatorbackground=card,
                        indicatorforeground=fg, upperbordercolor=border, lowerbordercolor=border)
        style.map(name, background=[("active", bg)], indicatorbackground=[("selected", accent)],
                  indicatorforeground=[("selected", on_accent)])
    # campi
    for name in ("TEntry", "TCombobox", "TSpinbox"):
        style.configure(name, fieldbackground=card, foreground=fg, background=card, bordercolor=border,
                        lightcolor=card, darkcolor=card, insertcolor=fg, arrowcolor=fg, padding=4)
        style.map(name, fieldbackground=[("readonly", card), ("disabled", bg)],
                  foreground=[("disabled", muted)], bordercolor=[("focus", accent)],
                  lightcolor=[("focus", accent)], background=[("active", hover)],
                  selectbackground=[("focus", sel)], selectforeground=[("focus", fg)])
    # schede
    style.configure("TNotebook", background=bg, bordercolor=border, tabmargins=(0, 2, 0, 0))
    style.configure("TNotebook.Tab", background=bg, foreground=muted, bordercolor=border,
                    lightcolor=bg, padding=(14, 6))
    style.map("TNotebook.Tab", background=[("selected", card), ("active", hover)],
              foreground=[("selected", fg)], lightcolor=[("selected", card)])
    # tabelle
    style.configure("Treeview", background=card, fieldbackground=card, foreground=fg, bordercolor=border,
                    lightcolor=card, darkcolor=card)
    style.map("Treeview", background=[("selected", accent)], foreground=[("selected", on_accent)])
    style.configure("Treeview.Heading", background=bg, foreground=muted, bordercolor=border,
                    lightcolor=bg, darkcolor=bg, relief="flat", font=font)
    style.map("Treeview.Heading", background=[("active", hover)])
    # scrollbar / progress
    for name in ("Vertical.TScrollbar", "Horizontal.TScrollbar"):
        style.configure(name, background=hover, troughcolor=bg, bordercolor=bg, arrowcolor=muted,
                        lightcolor=hover, darkcolor=hover, gripcount=0)
        style.map(name, background=[("active", border)])
    style.configure("TProgressbar", background=accent, troughcolor=card, bordercolor=border,
                    lightcolor=accent, darkcolor=accent)


def set_windows_titlebar(window, dark: bool) -> None:
    """Barra del titolo scura/chiara su Windows 10/11 (senza effetto su macOS e Linux)."""
    import sys
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        window.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
        value = ctypes.c_int(1 if dark else 0)
        for attr in (20, 19):  # DWMWA_USE_IMMERSIVE_DARK_MODE (nuovo e vecchio id)
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(value),
                                                          ctypes.sizeof(value)) == 0:
                break
        # forza il ridisegno della cornice
    except Exception:
        pass


class Theme:
    """Gestisce tema chiaro/scuro e notifica i widget che disegnano a mano (canvas, tag)."""

    def __init__(self, root: tk.Tk, mode: str = "auto"):
        global CURRENT
        self.root = root
        self._listeners = []
        self._date_entries = []
        CURRENT = self
        if mode == "auto":
            mode = "dark" if (darkdetect and darkdetect.isDark()) else "light"
        self.mode = mode
        self._setup_fonts()
        self.apply(mode)

    # --- font ---
    def _setup_fonts(self):
        base = tkfont.nametofont("TkDefaultFont")
        fam = base.actual("family")
        size = max(base.actual("size"), 11) if base.actual("size") > 0 else 11
        self.font = (fam, size)
        self.font_small = (fam, size - 1)
        self.font_bold = (fam, size, "bold")
        self.font_title = (fam, size + 7, "bold")
        self.font_h2 = (fam, size + 2, "bold")
        mono = "Menlo" if "Menlo" in tkfont.families() else "Courier"
        self.font_mono = (mono, size)

    # --- colori ---
    @property
    def c(self):
        return PALETTES[self.mode]

    def apply(self, mode: str):
        self.mode = mode
        style = ttk.Style()
        c = self.c
        if sv_ttk is not None:
            sv_ttk.set_theme(mode)
        else:
            _fallback_style(style, c, self.font, self.font_bold)
        set_windows_titlebar(self.root, mode == "dark")
        style.configure("Title.TLabel", font=self.font_title)
        style.configure("H2.TLabel", font=self.font_h2)
        style.configure("Muted.TLabel", foreground=c["muted"])
        style.configure("Small.TLabel", foreground=c["muted"], font=self.font_small)
        style.configure("Ok.TLabel", foreground=c["ok"], font=self.font_bold)
        style.configure("Err.TLabel", foreground=c["err"], font=self.font_bold)
        style.configure("Stat.TLabel", font=self.font_h2)
        style.configure("Treeview", rowheight=28)
        style.configure("Card.TFrame", background=c["card"])
        # widget "classici" (tk) che non seguono il tema ttk
        root = self.root
        root.option_add("*Text.background", c["card"])
        root.option_add("*Text.foreground", c["fg"])
        root.option_add("*Listbox.background", c["card"])
        root.option_add("*Listbox.foreground", c["fg"])
        root.option_add("*Menu.background", c["card"])
        root.option_add("*Menu.foreground", c["fg"])
        root.option_add("*Menu.activeBackground", c["sel"])
        root.option_add("*Menu.activeForeground", c["fg"])
        root.option_add("*Toplevel.background", c["bg"])
        root.option_add("*TCombobox*Listbox.background", c["card"])
        root.option_add("*TCombobox*Listbox.foreground", c["fg"])
        root.option_add("*TCombobox*Listbox.selectBackground", c["accent"])
        for w in list(self._date_entries):
            try:
                w.configure(**calendar_options(c))
            except tk.TclError:
                self._date_entries.remove(w)
        self._recolor_tk_widgets(root)
        for listener in list(self._listeners):
            try:
                listener()
            except tk.TclError:
                self._listeners.remove(listener)

    @property
    def using_fallback(self) -> bool:
        return sv_ttk is None

    def register_date_entry(self, widget):
        self._date_entries.append(widget)
        try:
            widget.configure(**calendar_options(self.c))
        except tk.TclError:
            pass

    def _recolor_tk_widgets(self, widget):
        """Ricolora ricorsivamente i widget tk gia' creati (Toplevel, Menu, Text, Listbox)."""
        c = self.c
        for child in widget.winfo_children():
            cls = child.winfo_class()
            try:
                if cls in ("Toplevel", "Tk"):
                    child.configure(bg=c["bg"])
                    set_windows_titlebar(child, self.mode == "dark")
                elif cls == "Menu":
                    child.configure(bg=c["card"], fg=c["fg"], activebackground=c["sel"], activeforeground=c["fg"])
                elif cls == "Listbox":
                    child.configure(bg=c["card"], fg=c["fg"])
            except tk.TclError:
                pass
            self._recolor_tk_widgets(child)

    def toggle(self):
        self.apply("light" if self.mode == "dark" else "dark")

    def on_change(self, callback):
        self._listeners.append(callback)
