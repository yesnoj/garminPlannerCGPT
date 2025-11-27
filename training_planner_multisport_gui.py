"""
Training Planner - MULTI-SPORT GUI con Editor Visuale
Versione completa con supporto per Running, Cycling e Swimming.
Include preset specifici per ogni sport e gestione zone multi-sport.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import json
from datetime import datetime
import threading
from typing import Optional, List, Dict, Any
try:
    from tkcalendar import DateEntry
    HAS_CALENDAR = True
except ImportError:
    HAS_CALENDAR = False


from excel_utils_multisport import generate_training_excel_multisport, format_workbook_dates_and_steps
from dsl_parser_multisport import expand_repeat_lines, build_garmin_workout_from_excel_row
from garmin_service import GarminService


# ============================================================================
# PRESET LIBRARY - Blocchi predefiniti comuni
# ============================================================================

# ============================================================================
# PRESET LIBRARY - Blocchi predefiniti MULTI-SPORT
# ============================================================================

# Preset organizzati per sport
PRESET_BLOCKS_BY_SPORT = {
    "RUNNING": {
        "🔥 Warmup 10min Z2": {
            "type": "warmup",
            "duration": "10min",
            "target": "Z2",
            "icon": "🔥"
        },
        "🔥 Warmup 15min Z2": {
            "type": "warmup",
            "duration": "15min",
            "target": "Z2",
            "icon": "🔥"
        },
        "⚡ 5min @ Z4": {
            "type": "interval",
            "duration": "5min",
            "target": "Z4",
            "icon": "⚡"
        },
        "🚀 1min @ Z5": {
            "type": "interval",
            "duration": "1min",
            "target": "Z5",
            "icon": "🚀"
        },
        "🏃 400m @ Z5": {
            "type": "interval",
            "duration": "400m",
            "target": "Z5",
            "icon": "🏃"
        },
        "🏃 1km @ Z4": {
            "type": "interval",
            "duration": "1km",
            "target": "Z4",
            "icon": "🏃"
        },
        "🧘 Recovery 2min Z1": {
            "type": "recovery",
            "duration": "2min",
            "target": "Z1",
            "icon": "🧘"
        },
        "❄️ Cooldown 10min Z1": {
            "type": "cooldown",
            "duration": "10min",
            "target": "Z1",
            "icon": "❄️"
        },
    },
    
    "CYCLING": {
        "🔥 Warmup 15min Power_Z2": {
            "type": "warmup",
            "duration": "15min",
            "target": "Power_Z2",
            "icon": "🔥"
        },
        "⚡ 8min @ Power_Z4 (FTP)": {
            "type": "interval",
            "duration": "8min",
            "target": "Power_Z4",
            "icon": "⚡"
        },
        "💪 10min @ Power_Z3 (Sweet Spot)": {
            "type": "interval",
            "duration": "10min",
            "target": "Power_Z3",
            "icon": "💪"
        },
        "🚀 5min @ Power_Z5 (VO2max)": {
            "type": "interval",
            "duration": "5min",
            "target": "Power_Z5",
            "icon": "🚀"
        },
        "🔄 3min @ Cadence_Sprint": {
            "type": "interval",
            "duration": "3min",
            "target": "Cadence_Sprint",
            "icon": "🔄"
        },
        "🧘 Recovery 4min Power_Z1": {
            "type": "recovery",
            "duration": "4min",
            "target": "Power_Z1",
            "icon": "🧘"
        },
        "❄️ Cooldown 10min Power_Z1": {
            "type": "cooldown",
            "duration": "10min",
            "target": "Power_Z1",
            "icon": "❄️"
        },
    },
    
    "SWIMMING": {
        "🔥 Warmup 400m Swim_Z1": {
            "type": "warmup",
            "duration": "400m",
            "target": "Swim_Z1",
            "icon": "🔥"
        },
        "⚡ 200m @ Swim_Z3 (Tempo)": {
            "type": "interval",
            "duration": "200m",
            "target": "Swim_Z3",
            "icon": "⚡"
        },
        "💪 100m @ Swim_Z4 (Threshold)": {
            "type": "interval",
            "duration": "100m",
            "target": "Swim_Z4",
            "icon": "💪"
        },
        "🚀 50m @ Swim_Z5 (Sprint)": {
            "type": "interval",
            "duration": "50m",
            "target": "Swim_Z5",
            "icon": "🚀"
        },
        "🧘 100m @ swim_easy": {
            "type": "recovery",
            "duration": "100m",
            "target": "swim_easy",
            "icon": "🧘"
        },
        "❄️ Cooldown 200m Swim_Z1": {
            "type": "cooldown",
            "duration": "200m",
            "target": "Swim_Z1",
            "icon": "❄️"
        },
    },
    
    "COMMON": {
        "⏸️ Rest 2min": {
            "type": "rest",
            "duration": "2min",
            "target": "",
            "icon": "⏸️"
        },
        "⏸️ Rest 60sec": {
            "type": "rest",
            "duration": "60sec",
            "target": "",
            "icon": "⏸️"
        },
        "⏸️ Rest 30sec": {
            "type": "rest",
            "duration": "30sec",
            "target": "",
            "icon": "⏸️"
        },
        "👆 Lap Button": {
            "type": "interval",
            "duration": "lap-button",
            "target": "open",
            "icon": "👆"
        },
    }
}

# Compatibilità: mantieni PRESET_BLOCKS per codice esistente
PRESET_BLOCKS = {}
for sport_blocks in PRESET_BLOCKS_BY_SPORT.values():
    PRESET_BLOCKS.update(sport_blocks)

# Repeat blocks predefiniti MULTI-SPORT
PRESET_REPEAT_BLOCKS_BY_SPORT = {
    "RUNNING": {
        "🔁 4x (1km Z4 + 2min rec)": {
            "repetitions": 4,
            "steps": [
                {"type": "interval", "duration": "1km", "target": "Z4"},
                {"type": "recovery", "duration": "2min", "target": "Z1"}
            ],
            "icon": "🔁",
            "description": "Running: Ripetute classiche 4x1km"
        },
        "🔁 6x (400m Z5 + 90sec rest)": {
            "repetitions": 6,
            "steps": [
                {"type": "interval", "duration": "400m", "target": "Z5"},
                {"type": "rest", "duration": "90sec", "target": ""}
            ],
            "icon": "🔁",
            "description": "Running: Sprint brevi 6x400m"
        },
        "🔁 10x (200m Z5 + 60sec rest)": {
            "repetitions": 10,
            "steps": [
                {"type": "interval", "duration": "200m", "target": "Z5"},
                {"type": "rest", "duration": "60sec", "target": ""}
            ],
            "icon": "🔁",
            "description": "Running: Sprint molto brevi"
        },
        "🔁🔁 2x [6x (1min Z5 + 1min rec) + 3min rest]": {
            "repetitions": 2,
            "steps": [
                {
                    "type": "nested_repeat",
                    "repetitions": 6,
                    "steps": [
                        {"type": "interval", "duration": "1min", "target": "Z5"},
                        {"type": "recovery", "duration": "1min", "target": "Z1"}
                    ]
                },
                {"type": "rest", "duration": "3min", "target": ""}
            ],
            "icon": "🔁🔁",
            "description": "Running: 2 serie da 6 ripetute"
        },
    },
    
    "CYCLING": {
        "🔁 4x (8min Power_Z4 + 4min rec)": {
            "repetitions": 4,
            "steps": [
                {"type": "interval", "duration": "8min", "target": "Power_Z4"},
                {"type": "recovery", "duration": "4min", "target": "Power_Z1"}
            ],
            "icon": "🔁",
            "description": "Cycling: FTP intervals classici"
        },
        "🔁 3x (10min Power_Z3 + 5min rec)": {
            "repetitions": 3,
            "steps": [
                {"type": "interval", "duration": "10min", "target": "Power_Z3"},
                {"type": "recovery", "duration": "5min", "target": "Power_Z1"}
            ],
            "icon": "🔁",
            "description": "Cycling: Sweet spot intervals"
        },
        "🔁 6x (3min Cadence_Sprint + 2min easy)": {
            "repetitions": 6,
            "steps": [
                {"type": "interval", "duration": "3min", "target": "Cadence_Sprint"},
                {"type": "recovery", "duration": "2min", "target": "Cadence_Easy"}
            ],
            "icon": "🔁",
            "description": "Cycling: Cadence drills"
        },
        "🔁🔁 3x [4x (2min Power_Z5 + 2min rec) + 5min rest]": {
            "repetitions": 3,
            "steps": [
                {
                    "type": "nested_repeat",
                    "repetitions": 4,
                    "steps": [
                        {"type": "interval", "duration": "2min", "target": "Power_Z5"},
                        {"type": "recovery", "duration": "2min", "target": "Power_Z2"}
                    ]
                },
                {"type": "rest", "duration": "5min", "target": ""}
            ],
            "icon": "🔁🔁",
            "description": "Cycling: VO2max intervals"
        },
    },
    
    "SWIMMING": {
        "🔁 6x (100m Swim_Z4 + 20sec rest)": {
            "repetitions": 6,
            "steps": [
                {"type": "interval", "duration": "100m", "target": "Swim_Z4"},
                {"type": "rest", "duration": "20sec", "target": ""}
            ],
            "icon": "🔁",
            "description": "Swimming: Threshold set"
        },
        "🔁 10x (50m Swim_Z5 + 30sec rest)": {
            "repetitions": 10,
            "steps": [
                {"type": "interval", "duration": "50m", "target": "Swim_Z5"},
                {"type": "rest", "duration": "30sec", "target": ""}
            ],
            "icon": "🔁",
            "description": "Swimming: Sprint set"
        },
        "🔁 4x (200m Swim_Z3 + 45sec rest)": {
            "repetitions": 4,
            "steps": [
                {"type": "interval", "duration": "200m", "target": "Swim_Z3"},
                {"type": "rest", "duration": "45sec", "target": ""}
            ],
            "icon": "🔁",
            "description": "Swimming: Tempo intervals"
        },
        "🔁🔁 3x [4x (25m Swim_Z5 + 15sec) + 60sec rest]": {
            "repetitions": 3,
            "steps": [
                {
                    "type": "nested_repeat",
                    "repetitions": 4,
                    "steps": [
                        {"type": "interval", "duration": "25m", "target": "Swim_Z5"},
                        {"type": "rest", "duration": "15sec", "target": ""}
                    ]
                },
                {"type": "rest", "duration": "60sec", "target": ""}
            ],
            "icon": "🔁🔁",
            "description": "Swimming: Sprint short bursts"
        },
    }
}

# Compatibilità: mantieni PRESET_REPEAT_BLOCKS per codice esistente
PRESET_REPEAT_BLOCKS = {}
for sport_blocks in PRESET_REPEAT_BLOCKS_BY_SPORT.values():
    PRESET_REPEAT_BLOCKS.update(sport_blocks)


# ============================================================================
# WORKOUT STEP - Rappresentazione di uno step
# ============================================================================

class WorkoutStep:
    """Rappresenta un singolo step dell'allenamento."""
    
    def __init__(self, step_type: str = "interval", duration: str = "5min", 
                 target: str = "Z2", indent: int = 0):
        self.step_type = step_type
        self.duration = duration
        self.target = target
        self.indent = indent  # Per gestire repeat annidati
        
    def to_dsl(self) -> str:
        """Converte lo step in formato DSL."""
        indent_str = "  " * self.indent
        
        if self.target:
            return f"{indent_str}{self.step_type}: {self.duration} @ {self.target}"
        else:
            return f"{indent_str}{self.step_type}: {self.duration}"
    
    def get_icon(self) -> str:
        """Ritorna l'icona rappresentativa dello step."""
        icons = {
            "warmup": "🔥",
            "cooldown": "❄️",
            "interval": "⚡",
            "recovery": "🧘",
            "rest": "⏸️"
        }
        return icons.get(self.step_type, "▪️")
    
    def get_display_text(self) -> str:
        """Testo da mostrare nella lista visuale."""
        icon = self.get_icon()
        indent_str = "  " * self.indent
        
        if self.target:
            return f"{indent_str}{icon} {self.step_type.upper()}: {self.duration} @ {self.target}"
        else:
            return f"{indent_str}{icon} {self.step_type.upper()}: {self.duration}"
    
    @staticmethod
    def from_dsl(dsl_line: str) -> 'WorkoutStep':
        """Crea uno step da una riga DSL."""
        # Conta indentazione
        indent = 0
        while dsl_line.startswith("  "):
            indent += 1
            dsl_line = dsl_line[2:]
        
        # Parse tipo step
        if ":" not in dsl_line:
            return WorkoutStep(indent=indent)
        
        type_part, rest = dsl_line.split(":", 1)
        step_type = type_part.strip().lower()
        rest = rest.strip()
        
        # Parse durata e target
        if "@" in rest:
            duration, target = rest.split("@", 1)
            duration = duration.strip()
            target = target.strip()
        else:
            duration = rest
            target = ""
        
        return WorkoutStep(step_type, duration, target, indent)


class RepeatBlock:
    """Rappresenta un blocco repeat."""
    
    def __init__(self, repetitions: int = 4, indent: int = 0):
        self.repetitions = repetitions
        self.indent = indent
        self.steps: List[WorkoutStep] = []
    
    def to_dsl(self) -> str:
        """Converte il repeat block in formato DSL."""
        indent_str = "  " * self.indent
        lines = [f"{indent_str}repeat {self.repetitions}:"]
        
        for step in self.steps:
            if isinstance(step, RepeatBlock):
                # Repeat innestato
                lines.append(step.to_dsl())
            else:
                lines.append(step.to_dsl())
        
        return "\n".join(lines)
    
    def get_display_text(self) -> str:
        """Testo da mostrare nella lista visuale."""
        indent_str = "  " * self.indent
        return f"{indent_str}🔁 REPEAT {self.repetitions}x"


# ============================================================================
# VISUAL WORKOUT BUILDER
# ============================================================================

class VisualWorkoutBuilder(tk.Toplevel):
    """Editor visuale per costruire workout con drag-and-drop."""
    
    def __init__(self, parent, initial_dsl: str = "", callback=None):
        super().__init__(parent)
        self.title("Visual Workout Builder")
        self.geometry("1300x800")
        
        self.callback = callback  # Funzione da chiamare al salvataggio
        self.steps: List[Any] = []  # Lista di WorkoutStep e RepeatBlock
        
        self._build_ui()
        
        # Carica DSL iniziale se fornito
        if initial_dsl:
            self._load_from_dsl(initial_dsl)
        
        # Centra la finestra
        self.transient(parent)
        self.grab_set()
    
    def _build_ui(self):
        """Costruisce l'interfaccia."""
        # Top toolbar
        toolbar = ttk.Frame(self, padding=5)
        toolbar.pack(side=tk.TOP, fill=tk.X)
        
        ttk.Button(toolbar, text="💾 Salva", command=self._save_workout).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="❌ Annulla", command=self.destroy).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        ttk.Button(toolbar, text="🔁 Aggiungi Repeat Block", command=self._add_repeat_block).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="📄 Mostra DSL", command=self._show_dsl).pack(side=tk.LEFT, padx=2)
        
        # Main container
        main_container = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left: Preset library
        left_frame = ttk.LabelFrame(main_container, text="📚 Libreria Preset", padding=5)
        main_container.add(left_frame, weight=1)
        
        # Preset list
        preset_canvas = tk.Canvas(left_frame, bg="white")
        preset_scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=preset_canvas.yview)
        self.preset_frame = ttk.Frame(preset_canvas)
        
        preset_canvas.create_window((0, 0), window=self.preset_frame, anchor="nw")
        preset_canvas.configure(yscrollcommand=preset_scrollbar.set)
        
        preset_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        preset_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self._populate_preset_library()
        
        # Update scroll region
        self.preset_frame.update_idletasks()
        preset_canvas.configure(scrollregion=preset_canvas.bbox("all"))
        
        # Right: Workout builder
        right_frame = ttk.LabelFrame(main_container, text="🏗️ Costruttore Workout", padding=5)
        main_container.add(right_frame, weight=2)
        
        # Workout steps list
        list_container = ttk.Frame(right_frame)
        list_container.pack(fill=tk.BOTH, expand=True)
        
        self.steps_listbox = tk.Listbox(list_container, height=25, font=("Courier", 10))
        list_scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.steps_listbox.yview)
        self.steps_listbox.configure(yscrollcommand=list_scrollbar.set)
        
        self.steps_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        list_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Bind double-click per editing
        self.steps_listbox.bind("<Double-Button-1>", self._edit_selected_step)
        
        # Bottom buttons
        btn_frame = ttk.Frame(right_frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(btn_frame, text="✏️ Modifica", command=self._edit_selected_step).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="🗑️ Elimina", command=self._delete_selected_step).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="⬆️ Sposta Su", command=self._move_step_up).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="⬇️ Sposta Giù", command=self._move_step_down).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="➡️ Aumenta Indent", command=self._increase_indent).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="⬅️ Riduci Indent", command=self._decrease_indent).pack(side=tk.LEFT, padx=2)
    
    def _populate_preset_library(self):
        """Popola la libreria di preset MULTI-SPORT organizzata per sport."""
        
        # ========== RUNNING ==========
        ttk.Label(
            self.preset_frame,
            text="🏃 RUNNING",
            font=("", 10, "bold"),
            foreground="blue"
        ).pack(fill=tk.X, pady=(5, 2))
        
        for name, config in PRESET_BLOCKS_BY_SPORT["RUNNING"].items():
            btn = ttk.Button(
                self.preset_frame,
                text=name,
                command=lambda c=config: self._add_preset_step(c),
                width=35
            )
            btn.pack(fill=tk.X, pady=1, padx=5)
        
        # Running repeat blocks
        for name, config in PRESET_REPEAT_BLOCKS_BY_SPORT["RUNNING"].items():
            btn = ttk.Button(
                self.preset_frame,
                text=name,
                command=lambda c=config: self._add_preset_repeat(c),
                width=35
            )
            btn.pack(fill=tk.X, pady=1, padx=5)
            ttk.Label(
                self.preset_frame,
                text=config['description'],
                font=("", 7),
                foreground="gray"
            ).pack(fill=tk.X, padx=10)
        
        ttk.Separator(self.preset_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=8)
        
        # ========== CYCLING ==========
        ttk.Label(
            self.preset_frame,
            text="🚴 CYCLING",
            font=("", 10, "bold"),
            foreground="green"
        ).pack(fill=tk.X, pady=(5, 2))
        
        for name, config in PRESET_BLOCKS_BY_SPORT["CYCLING"].items():
            btn = ttk.Button(
                self.preset_frame,
                text=name,
                command=lambda c=config: self._add_preset_step(c),
                width=35
            )
            btn.pack(fill=tk.X, pady=1, padx=5)
        
        # Cycling repeat blocks
        for name, config in PRESET_REPEAT_BLOCKS_BY_SPORT["CYCLING"].items():
            btn = ttk.Button(
                self.preset_frame,
                text=name,
                command=lambda c=config: self._add_preset_repeat(c),
                width=35
            )
            btn.pack(fill=tk.X, pady=1, padx=5)
            ttk.Label(
                self.preset_frame,
                text=config['description'],
                font=("", 7),
                foreground="gray"
            ).pack(fill=tk.X, padx=10)
        
        ttk.Separator(self.preset_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=8)
        
        # ========== SWIMMING ==========
        ttk.Label(
            self.preset_frame,
            text="🏊 SWIMMING",
            font=("", 10, "bold"),
            foreground="teal"
        ).pack(fill=tk.X, pady=(5, 2))
        
        for name, config in PRESET_BLOCKS_BY_SPORT["SWIMMING"].items():
            btn = ttk.Button(
                self.preset_frame,
                text=name,
                command=lambda c=config: self._add_preset_step(c),
                width=35
            )
            btn.pack(fill=tk.X, pady=1, padx=5)
        
        # Swimming repeat blocks
        for name, config in PRESET_REPEAT_BLOCKS_BY_SPORT["SWIMMING"].items():
            btn = ttk.Button(
                self.preset_frame,
                text=name,
                command=lambda c=config: self._add_preset_repeat(c),
                width=35
            )
            btn.pack(fill=tk.X, pady=1, padx=5)
            ttk.Label(
                self.preset_frame,
                text=config['description'],
                font=("", 7),
                foreground="gray"
            ).pack(fill=tk.X, padx=10)
        
        ttk.Separator(self.preset_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=8)
        
        # ========== COMMON (Rest, Lap Button) ==========
        ttk.Label(
            self.preset_frame,
            text="⚙️ COMMON",
            font=("", 10, "bold"),
            foreground="gray"
        ).pack(fill=tk.X, pady=(5, 2))
        
        for name, config in PRESET_BLOCKS_BY_SPORT["COMMON"].items():
            btn = ttk.Button(
                self.preset_frame,
                text=name,
                command=lambda c=config: self._add_preset_step(c),
                width=35
            )
            btn.pack(fill=tk.X, pady=1, padx=5)
    
    def _add_preset_step(self, config: Dict[str, str]):
        """Aggiunge uno step preset al workout."""
        step = WorkoutStep(
            step_type=config["type"],
            duration=config["duration"],
            target=config.get("target", "")
        )
        self.steps.append(step)
        self._refresh_steps_list()
    
    def _add_preset_repeat(self, config: Dict[str, Any]):
        """Aggiunge un blocco repeat preset al workout."""
        block = RepeatBlock(repetitions=config["repetitions"])
        
        # Costruisci gli step del repeat
        for step_config in config["steps"]:
            if step_config.get("type") == "nested_repeat":
                # Repeat innestato
                nested_block = RepeatBlock(
                    repetitions=step_config["repetitions"],
                    indent=1  # Indentato di un livello
                )
                
                for nested_step_config in step_config["steps"]:
                    nested_step = WorkoutStep(
                        step_type=nested_step_config["type"],
                        duration=nested_step_config["duration"],
                        target=nested_step_config.get("target", ""),
                        indent=2  # Indentato di due livelli
                    )
                    nested_block.steps.append(nested_step)
                
                block.steps.append(nested_block)
            else:
                # Step normale
                step = WorkoutStep(
                    step_type=step_config["type"],
                    duration=step_config["duration"],
                    target=step_config.get("target", ""),
                    indent=1  # Indentato di un livello
                )
                block.steps.append(step)
        
        self.steps.append(block)
        self._refresh_steps_list()
    
    def _add_repeat_block(self):
        """Aggiunge un blocco repeat."""
        reps = simpledialog.askinteger(
            "Ripetizioni",
            "Numero di ripetizioni:",
            parent=self,
            minvalue=1,
            maxvalue=50,
            initialvalue=4
        )
        
        if reps:
            block = RepeatBlock(repetitions=reps)
            self.steps.append(block)
            self._refresh_steps_list()
    
    def _refresh_steps_list(self):
        """Aggiorna la visualizzazione della lista steps."""
        self.steps_listbox.delete(0, tk.END)
        
        for item in self.steps:
            if isinstance(item, RepeatBlock):
                self.steps_listbox.insert(tk.END, item.get_display_text())
                # Aggiungi gli step del repeat
                for step in item.steps:
                    if isinstance(step, RepeatBlock):
                        # Repeat innestato
                        self.steps_listbox.insert(tk.END, step.get_display_text())
                        for nested_step in step.steps:
                            self.steps_listbox.insert(tk.END, nested_step.get_display_text())
                    else:
                        self.steps_listbox.insert(tk.END, step.get_display_text())
            else:
                self.steps_listbox.insert(tk.END, item.get_display_text())
    
    def _count_listbox_rows(self, item) -> int:
        """Conta quante righe occupa un item nella listbox."""
        if isinstance(item, RepeatBlock):
            # 1 riga per l'header del repeat + tutte le righe dei suoi step
            count = 1
            for step in item.steps:
                if isinstance(step, RepeatBlock):
                    # Repeat innestato: ricorsione
                    count += self._count_listbox_rows(step)
                else:
                    count += 1
            return count
        else:
            # Step normale: 1 riga
            return 1
    
    def _edit_selected_step(self, event=None):
        """Modifica lo step selezionato."""
        selection = self.steps_listbox.curselection()
        if not selection:
            return
        
        idx = selection[0]
        
        # Usa approccio ricorsivo per trovare l'item
        self._find_and_edit_item(self.steps, idx, 0)
    
    def _find_and_edit_item(self, items_list, target_idx, current_flat_idx):
        """Trova ricorsivamente un item e lo modifica."""
        for item in items_list:
            if isinstance(item, RepeatBlock):
                if current_flat_idx == target_idx:
                    self._edit_repeat_block(item)
                    return True
                
                current_flat_idx += 1
                
                # Ricorsione negli step del repeat
                result = self._find_and_edit_item_helper(item.steps, target_idx, current_flat_idx)
                if result['found']:
                    return True
                current_flat_idx = result['flat_idx']
            else:
                if current_flat_idx == target_idx:
                    self._edit_step(item)
                    return True
                current_flat_idx += 1
        
        return False
    
    def _find_and_edit_item_helper(self, items_list, target_idx, current_flat_idx):
        """Helper per ricorsione."""
        for item in items_list:
            if isinstance(item, RepeatBlock):
                if current_flat_idx == target_idx:
                    self._edit_repeat_block(item)
                    return {'found': True, 'flat_idx': current_flat_idx}
                
                current_flat_idx += 1
                
                result = self._find_and_edit_item_helper(item.steps, target_idx, current_flat_idx)
                if result['found']:
                    return result
                current_flat_idx = result['flat_idx']
            else:
                if current_flat_idx == target_idx:
                    self._edit_step(item)
                    return {'found': True, 'flat_idx': current_flat_idx}
                current_flat_idx += 1
        
        return {'found': False, 'flat_idx': current_flat_idx}
    
    def _edit_step(self, step: WorkoutStep):
        """Apre dialog per modificare uno step."""
        dialog = StepEditorDialog(self, step)
        self.wait_window(dialog)
        
        if dialog.result:
            step.step_type = dialog.result["type"]
            step.duration = dialog.result["duration"]
            step.target = dialog.result["target"]
            self._refresh_steps_list()
    
    def _edit_repeat_block(self, block: RepeatBlock):
        """Modifica un repeat block."""
        reps = simpledialog.askinteger(
            "Ripetizioni",
            "Numero di ripetizioni:",
            parent=self,
            minvalue=1,
            maxvalue=50,
            initialvalue=block.repetitions
        )
        
        if reps:
            block.repetitions = reps
            self._refresh_steps_list()
    
    def _delete_selected_step(self):
        """Elimina lo step selezionato."""
        selection = self.steps_listbox.curselection()
        if not selection:
            return
        
        idx = selection[0]
        
        # Usa approccio ricorsivo per trovare ed eliminare
        result = self._delete_item_recursive(self.steps, idx, 0)
        
        if result['deleted']:
            self._refresh_steps_list()
    
    def _delete_item_recursive(self, items_list, target_idx, current_flat_idx):
        """
        Trova e elimina ricorsivamente un item.
        Returns: {'deleted': bool, 'flat_idx': int}
        """
        for i, item in enumerate(items_list):
            if isinstance(item, RepeatBlock):
                # Check header del repeat
                if current_flat_idx == target_idx:
                    # Rimuovi il RepeatBlock
                    items_list.pop(i)
                    return {'deleted': True, 'flat_idx': current_flat_idx}
                
                current_flat_idx += 1
                
                # Ricorsione negli step del repeat
                result = self._delete_item_recursive(item.steps, target_idx, current_flat_idx)
                if result['deleted']:
                    return result
                current_flat_idx = result['flat_idx']
                
            else:
                # Step normale
                if current_flat_idx == target_idx:
                    # Rimuovi lo step
                    items_list.pop(i)
                    return {'deleted': True, 'flat_idx': current_flat_idx}
                
                current_flat_idx += 1
        
        return {'deleted': False, 'flat_idx': current_flat_idx}
    
    def _move_step_up(self):
        """Sposta lo step selezionato in su."""
        selection = self.steps_listbox.curselection()
        if not selection or selection[0] == 0:
            return
        
        idx = selection[0]
        
        # Usa approccio ricorsivo per trovare e spostare
        result = self._move_item_up_recursive(self.steps, idx, 0)
        
        if result['moved']:
            self._refresh_steps_list()
            self.steps_listbox.selection_set(result['new_idx'])
            self.steps_listbox.see(result['new_idx'])
    
    def _move_item_up_recursive(self, items_list, target_idx, current_flat_idx):
        """
        Trova e sposta ricorsivamente un item verso l'alto.
        Returns: {'moved': bool, 'new_idx': int, 'flat_idx': int}
        """
        for i, item in enumerate(items_list):
            if isinstance(item, RepeatBlock):
                # Check header del repeat
                if current_flat_idx == target_idx:
                    if i > 0:
                        # Swappa con precedente
                        prev_item = items_list[i-1]
                        items_list[i], items_list[i-1] = items_list[i-1], items_list[i]
                        new_idx = target_idx - self._count_listbox_rows(prev_item)
                        return {'moved': True, 'new_idx': new_idx, 'flat_idx': current_flat_idx}
                    else:
                        return {'moved': False, 'new_idx': target_idx, 'flat_idx': current_flat_idx}
                
                current_flat_idx += 1
                
                # Ricorsione negli step del repeat
                result = self._move_item_up_recursive(item.steps, target_idx, current_flat_idx)
                if result['moved']:
                    return result
                current_flat_idx = result['flat_idx']
                
            else:
                # Step normale
                if current_flat_idx == target_idx:
                    if i > 0:
                        # Swappa con precedente
                        prev_item = items_list[i-1]
                        items_list[i], items_list[i-1] = items_list[i-1], items_list[i]
                        new_idx = target_idx - self._count_listbox_rows(prev_item)
                        return {'moved': True, 'new_idx': new_idx, 'flat_idx': current_flat_idx}
                    else:
                        return {'moved': False, 'new_idx': target_idx, 'flat_idx': current_flat_idx}
                
                current_flat_idx += 1
        
        return {'moved': False, 'new_idx': target_idx, 'flat_idx': current_flat_idx}
    
    def _move_step_down(self):
        """Sposta lo step selezionato in giù."""
        selection = self.steps_listbox.curselection()
        if not selection:
            return
        
        idx = selection[0]
        
        # Usa approccio ricorsivo per trovare e spostare
        result = self._move_item_down_recursive(self.steps, idx, 0)
        
        if result['moved']:
            self._refresh_steps_list()
            self.steps_listbox.selection_set(result['new_idx'])
            self.steps_listbox.see(result['new_idx'])
    
    def _move_item_down_recursive(self, items_list, target_idx, current_flat_idx):
        """
        Trova e sposta ricorsivamente un item verso il basso.
        Returns: {'moved': bool, 'new_idx': int, 'flat_idx': int}
        """
        for i, item in enumerate(items_list):
            if isinstance(item, RepeatBlock):
                # Check header del repeat
                if current_flat_idx == target_idx:
                    if i < len(items_list) - 1:
                        # Swappa con successivo
                        next_item = items_list[i+1]
                        items_list[i], items_list[i+1] = items_list[i+1], items_list[i]
                        new_idx = target_idx + self._count_listbox_rows(next_item)
                        return {'moved': True, 'new_idx': new_idx, 'flat_idx': current_flat_idx}
                    else:
                        return {'moved': False, 'new_idx': target_idx, 'flat_idx': current_flat_idx}
                
                current_flat_idx += 1
                
                # Ricorsione negli step del repeat
                result = self._move_item_down_recursive(item.steps, target_idx, current_flat_idx)
                if result['moved']:
                    return result
                current_flat_idx = result['flat_idx']
                
            else:
                # Step normale
                if current_flat_idx == target_idx:
                    if i < len(items_list) - 1:
                        # Swappa con successivo
                        next_item = items_list[i+1]
                        items_list[i], items_list[i+1] = items_list[i+1], items_list[i]
                        new_idx = target_idx + self._count_listbox_rows(next_item)
                        return {'moved': True, 'new_idx': new_idx, 'flat_idx': current_flat_idx}
                    else:
                        return {'moved': False, 'new_idx': target_idx, 'flat_idx': current_flat_idx}
                
                current_flat_idx += 1
        
        return {'moved': False, 'new_idx': target_idx, 'flat_idx': current_flat_idx}
    
    def _increase_indent(self):
        """Aumenta l'indentazione dello step selezionato."""
        selection = self.steps_listbox.curselection()
        if not selection:
            return
        
        idx = selection[0]
        
        # Usa approccio ricorsivo
        result = self._change_indent_recursive(self.steps, idx, 0, +1)
        
        if result['changed']:
            self._refresh_steps_list()
            self.steps_listbox.selection_set(idx)
    
    def _decrease_indent(self):
        """Riduce l'indentazione dello step selezionato."""
        selection = self.steps_listbox.curselection()
        if not selection:
            return
        
        idx = selection[0]
        
        # Usa approccio ricorsivo
        result = self._change_indent_recursive(self.steps, idx, 0, -1)
        
        if result['changed']:
            self._refresh_steps_list()
            self.steps_listbox.selection_set(idx)
    
    def _change_indent_recursive(self, items_list, target_idx, current_flat_idx, delta):
        """
        Trova ricorsivamente un item e cambia il suo indent.
        delta: +1 per aumentare, -1 per diminuire
        Returns: {'changed': bool, 'flat_idx': int}
        """
        for item in items_list:
            if isinstance(item, RepeatBlock):
                # Check header del repeat
                if current_flat_idx == target_idx:
                    if delta > 0 or item.indent > 0:
                        item.indent += delta
                        return {'changed': True, 'flat_idx': current_flat_idx}
                    return {'changed': False, 'flat_idx': current_flat_idx}
                
                current_flat_idx += 1
                
                # Ricorsione negli step del repeat
                result = self._change_indent_recursive(item.steps, target_idx, current_flat_idx, delta)
                if result['changed']:
                    return result
                current_flat_idx = result['flat_idx']
                
            else:
                # Step normale
                if current_flat_idx == target_idx:
                    if delta > 0 or item.indent > 0:
                        item.indent += delta
                        return {'changed': True, 'flat_idx': current_flat_idx}
                    return {'changed': False, 'flat_idx': current_flat_idx}
                
                current_flat_idx += 1
        
        return {'changed': False, 'flat_idx': current_flat_idx}
    
    def _load_from_dsl(self, dsl_text: str):
        """Carica un workout da DSL esistente."""
        self.steps.clear()
        lines = dsl_text.strip().split("\n")
        
        current_repeat_stack = []  # Stack per gestire repeat annidati
        
        for line in lines:
            if not line.strip():
                continue
            
            # Conta indent
            indent = 0
            original_line = line
            while line.startswith("  "):
                indent += 1
                line = line[2:]
            
            # Check se è repeat
            if line.strip().lower().startswith("repeat"):
                import re
                m = re.match(r"repeat\s+(\d+)\s*:", line.strip(), re.IGNORECASE)
                if m:
                    reps = int(m.group(1))
                    new_repeat = RepeatBlock(repetitions=reps, indent=indent)
                    
                    # Determina dove aggiungere il repeat
                    if indent == 0:
                        # Top-level repeat
                        self.steps.append(new_repeat)
                        current_repeat_stack = [new_repeat]
                    else:
                        # Nested repeat
                        if current_repeat_stack:
                            parent = current_repeat_stack[-1]
                            parent.steps.append(new_repeat)
                            current_repeat_stack.append(new_repeat)
                    continue
            
            # Altrimenti è uno step normale
            step = WorkoutStep.from_dsl(original_line)
            
            # Determina dove aggiungere lo step
            if current_repeat_stack and indent > current_repeat_stack[0].indent:
                # Aggiungi all'ultimo repeat nello stack
                target_repeat = current_repeat_stack[-1]
                target_repeat.steps.append(step)
                
                # Pop dallo stack se siamo tornati a un livello precedente
                while len(current_repeat_stack) > 1 and indent <= current_repeat_stack[-1].indent:
                    current_repeat_stack.pop()
            else:
                # Step top-level
                current_repeat_stack = []
                self.steps.append(step)
        
        self._refresh_steps_list()
    
    def _to_dsl(self) -> str:
        """Converte il workout corrente in DSL."""
        lines = []
        
        for item in self.steps:
            if isinstance(item, RepeatBlock):
                lines.append(item.to_dsl())
            else:
                lines.append(item.to_dsl())
        
        return "\n".join(lines)
    
    def _show_dsl(self):
        """Mostra il DSL generato."""
        dsl = self._to_dsl()
        
        dialog = tk.Toplevel(self)
        dialog.title("DSL Generato")
        dialog.geometry("600x400")
        
        text = tk.Text(dialog, wrap=tk.NONE)
        text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        text.insert("1.0", dsl)
        text.config(state=tk.DISABLED)
        
        ttk.Button(dialog, text="Chiudi", command=dialog.destroy).pack(pady=5)
    
    def _save_workout(self):
        """Salva il workout e chiude."""
        if self.callback:
            dsl = self._to_dsl()
            self.callback(dsl)
        self.destroy()


# ============================================================================
# STEP EDITOR DIALOG
# ============================================================================

class StepEditorDialog(tk.Toplevel):
    """Dialog per modificare un singolo step."""
    
    def __init__(self, parent, step: WorkoutStep):
        super().__init__(parent)
        self.title("Modifica Step")
        self.geometry("400x250")
        
        self.step = step
        self.result = None
        
        self._build_ui()
        
        # Centra
        self.transient(parent)
        self.grab_set()
    
    def _build_ui(self):
        """Costruisce l'interfaccia."""
        frame = ttk.Frame(self, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Step type
        ttk.Label(frame, text="Tipo Step:").grid(row=0, column=0, sticky="w", pady=5)
        self.type_var = tk.StringVar(value=self.step.step_type)
        type_combo = ttk.Combobox(
            frame,
            textvariable=self.type_var,
            values=["warmup", "cooldown", "interval", "recovery", "rest"],
            state="readonly"
        )
        type_combo.grid(row=0, column=1, sticky="we", pady=5, padx=(5, 0))
        
        # Duration
        ttk.Label(frame, text="Durata:").grid(row=1, column=0, sticky="w", pady=5)
        self.duration_var = tk.StringVar(value=self.step.duration)
        duration_entry = ttk.Entry(frame, textvariable=self.duration_var)
        duration_entry.grid(row=1, column=1, sticky="we", pady=5, padx=(5, 0))
        
        ttk.Label(frame, text="(es: 10min, 5km, 400m, lap-button)", font=("", 8)).grid(
            row=2, column=1, sticky="w", padx=(5, 0)
        )
        
        # Target
        ttk.Label(frame, text="Target:").grid(row=3, column=0, sticky="w", pady=5)
        self.target_var = tk.StringVar(value=self.step.target)
        target_entry = ttk.Entry(frame, textvariable=self.target_var)
        target_entry.grid(row=3, column=1, sticky="we", pady=5, padx=(5, 0))
        
        ttk.Label(frame, text="(es: Z2, 5:00, HR_Z3, open)", font=("", 8)).grid(
            row=4, column=1, sticky="w", padx=(5, 0)
        )
        
        frame.columnconfigure(1, weight=1)
        
        # Buttons
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="💾 Salva", command=self._save).pack(side=tk.RIGHT, padx=2)
        ttk.Button(btn_frame, text="❌ Annulla", command=self.destroy).pack(side=tk.RIGHT, padx=2)
    
    def _save(self):
        """Salva le modifiche."""
        self.result = {
            "type": self.type_var.get(),
            "duration": self.duration_var.get().strip(),
            "target": self.target_var.get().strip()
        }
        self.destroy()


# ============================================================================
# LOADING DIALOG (riutilizzato dalla GUI principale)
# ============================================================================

class LoadingDialog:
    """Finestra di loading modale con animazione."""
    
    def __init__(self, parent, title="Operazione in corso..."):
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.transient(parent)
        self.top.grab_set()
        
        self.top.geometry("300x100")
        self.top.resizable(False, False)
        
        frame = ttk.Frame(self.top, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        self.label = ttk.Label(frame, text="Attendere...", font=("", 10))
        self.label.pack(pady=(0, 10))
        
        self.progress = ttk.Progressbar(frame, mode='indeterminate', length=250)
        self.progress.pack()
        self.progress.start(10)
        
        self.top.protocol("WM_DELETE_WINDOW", lambda: None)
        
        self.top.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (self.top.winfo_width() // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (self.top.winfo_height() // 2)
        self.top.geometry(f"+{x}+{y}")
    
    def update_message(self, message: str):
        """Aggiorna il messaggio mostrato."""
        self.label.config(text=message)
        self.top.update_idletasks()
    
    def close(self):
        """Chiude la finestra di loading."""
        self.progress.stop()
        self.top.grab_release()
        self.top.destroy()


# ============================================================================
# NEW WORKOUT DIALOG
# ============================================================================

class NewWorkoutDialog(tk.Toplevel):
    """Dialog per creare un nuovo workout."""
    
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Nuovo Workout")
        self.geometry("450x300")
        
        self.result = None
        
        self._build_ui()
        
        # Centra
        self.transient(parent)
        self.grab_set()
        
        # Centra rispetto al parent
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (self.winfo_width() // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")
    
    def _build_ui(self):
        """Costruisce l'interfaccia."""
        frame = ttk.Frame(self, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Week
        ttk.Label(frame, text="Week:").grid(row=0, column=0, sticky="w", pady=8)
        self.entry_week = ttk.Entry(frame, width=10)
        self.entry_week.grid(row=0, column=1, sticky="w", pady=8)
        self.entry_week.insert(0, "1")
        
        # Date
        ttk.Label(frame, text="Date (YYYY-MM-DD):").grid(row=1, column=0, sticky="w", pady=8)
        self.entry_date = ttk.Entry(frame, width=25)
        self.entry_date.grid(row=1, column=1, sticky="we", pady=8)
        
        # Data di default: oggi
        from datetime import date
        today = date.today().isoformat()
        self.entry_date.insert(0, today)
        
        ttk.Label(frame, text="(es: 2025-05-15)", font=("", 8), foreground="gray").grid(
            row=1, column=2, sticky="w", padx=(5, 0)
        )
        
        # Session
        ttk.Label(frame, text="Session:").grid(row=2, column=0, sticky="w", pady=8)
        self.entry_session = ttk.Entry(frame, width=10)
        self.entry_session.grid(row=2, column=1, sticky="w", pady=8)
        self.entry_session.insert(0, "1")
        
        # Sport
        ttk.Label(frame, text="Sport:").grid(row=3, column=0, sticky="w", pady=8)
        self.sport_var = tk.StringVar(value="Running")
        sport_combo = ttk.Combobox(
            frame,
            textvariable=self.sport_var,
            values=["Running", "Cycling", "Swimming"],
            state="readonly",
            width=22
        )
        sport_combo.grid(row=3, column=1, sticky="w", pady=8)
        
        # Description
        ttk.Label(frame, text="Description:").grid(row=4, column=0, sticky="nw", pady=8)
        self.entry_description = tk.Text(frame, width=30, height=3)
        self.entry_description.grid(row=4, column=1, columnspan=2, sticky="we", pady=8)
        self.entry_description.insert("1.0", "Nuovo allenamento")
        
        frame.columnconfigure(1, weight=1)
        
        # Buttons
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        ttk.Button(btn_frame, text="✅ Crea", command=self._create).pack(side=tk.RIGHT, padx=2)
        ttk.Button(btn_frame, text="❌ Annulla", command=self.destroy).pack(side=tk.RIGHT, padx=2)
    
    def _create(self):
        """Crea il workout."""
        # Validazione
        week_str = self.entry_week.get().strip()
        date_str = self.entry_date.get().strip()
        session_str = self.entry_session.get().strip()
        sport = self.sport_var.get()
        description = self.entry_description.get("1.0", tk.END).strip()
        
        if not date_str:
            messagebox.showerror("Errore", "Inserisci una data.", parent=self)
            return
        
        # Valida formato data
        try:
            datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            messagebox.showerror(
                "Errore",
                "Formato data non valido.\nUsa: YYYY-MM-DD (es: 2025-05-15)",
                parent=self
            )
            return
        
        if not description:
            messagebox.showerror("Errore", "Inserisci una descrizione.", parent=self)
            return
        
        # Salva risultato
        self.result = {
            "week": int(week_str) if week_str else 1,
            "date": date_str,
            "session": int(session_str) if session_str else 1,
            "sport": sport,
            "description": description
        }
        
        self.destroy()


# ============================================================================
# MAIN GUI (versione avanzata)
# ============================================================================

class TrainingPlannerAdvancedGUI:
    """GUI avanzata con editor visuale per workout."""
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Training Planner - Multi-Sport Edition 🏃🚴🏊")
        
        self.excel_path: Optional[str] = None
        self.df_workouts: Optional[pd.DataFrame] = None
        self.df_parameters: Optional[pd.DataFrame] = None
        
        self.garmin = GarminService()
        self.use_prefix_var = tk.BooleanVar(value=True)
        
        self._build_ui()
    
    def _build_ui(self):
        """Costruisce l'interfaccia principale."""
        # Top toolbar
        toolbar = ttk.Frame(self.root, padding=5)
        toolbar.pack(side=tk.TOP, fill=tk.X)
        
        ttk.Button(toolbar, text="📄 Genera Excel", command=self._generate_excel).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="📁 Carica Excel", command=self._load_excel).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="💾 Salva Excel", command=self._save_excel).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        ttk.Button(toolbar, text="➕ Nuovo Workout", command=self._create_new_workout).pack(side=tk.LEFT, padx=2)
        
        self.lbl_file = ttk.Label(toolbar, text="Nessun file caricato")
        self.lbl_file.pack(side=tk.LEFT, padx=10)
        
        # Garmin frame
        garmin_frame = ttk.LabelFrame(self.root, text="⌚ Garmin Connect", padding=10)
        garmin_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        
        garmin_frame.columnconfigure(1, weight=1)
        
        # Authentication
        ttk.Label(garmin_frame, text="Email:").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=(0, 5))
        self.entry_email = ttk.Entry(garmin_frame, width=35)
        self.entry_email.grid(row=0, column=1, columnspan=2, sticky="we", padx=(0, 10), pady=(0, 5))
        
        ttk.Label(garmin_frame, text="Password:").grid(row=1, column=0, sticky="e", padx=(0, 8), pady=(0, 8))
        self.entry_password = ttk.Entry(garmin_frame, width=35, show="*")
        self.entry_password.grid(row=1, column=1, columnspan=2, sticky="we", padx=(0, 10), pady=(0, 8))
        
        # Login buttons
        btn_frame_login = ttk.Frame(garmin_frame)
        btn_frame_login.grid(row=0, column=3, rowspan=2, sticky="nsew", padx=(10, 0))
        
        ttk.Button(btn_frame_login, text="🔐 Login", command=self._login_credentials).pack(side=tk.TOP, fill=tk.X, pady=(0, 5))
        ttk.Button(btn_frame_login, text="♻️ Usa Sessione", command=self._login_session).pack(side=tk.TOP, fill=tk.X)
        
        self.lbl_garmin_status = ttk.Label(garmin_frame, text="● Non connesso", foreground="red", font=("", 9, "bold"))
        self.lbl_garmin_status.grid(row=0, column=4, rowspan=2, sticky="w", padx=(15, 0))
        
        # Separator
        ttk.Separator(garmin_frame, orient="horizontal").grid(row=2, column=0, columnspan=5, sticky="ew", pady=10)
        
        # Options
        chk_prefix = ttk.Checkbutton(garmin_frame, text="☑ Prefisso WnSn", variable=self.use_prefix_var)
        chk_prefix.grid(row=3, column=0, columnspan=5, sticky="w", pady=(0, 10))
        
        ttk.Separator(garmin_frame, orient="horizontal").grid(row=4, column=0, columnspan=5, sticky="ew", pady=(0, 10))
        
        # Operations
        btn_ops = ttk.Frame(garmin_frame)
        btn_ops.grid(row=5, column=0, columnspan=5, sticky="ew")
        
        for i in range(4):
            btn_ops.columnconfigure(i, weight=1)
        
        ttk.Button(btn_ops, text="📤 Carica", command=self._upload_workouts).grid(row=0, column=0, sticky="ew", padx=(0, 5))
        ttk.Button(btn_ops, text="📤📅 Carica+Pianifica", command=self._upload_and_schedule).grid(row=0, column=1, sticky="ew", padx=(0, 5))
        ttk.Button(btn_ops, text="📅✖ Rimuovi Piano", command=self._unschedule).grid(row=0, column=2, sticky="ew", padx=(0, 5))
        ttk.Button(btn_ops, text="🗑 Cancella", command=self._delete_workouts).grid(row=0, column=3, sticky="ew")
        
        # Main container
        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left: Workouts list
        left_frame = ttk.Frame(main_pane, padding=5)
        main_pane.add(left_frame, weight=1)
        
        ttk.Label(left_frame, text="📋 Lista Workouts").pack(anchor="w")
        
        cols = ("Week", "Date", "Session", "Sport", "Description", "ScheduledDate")
        self.tree_workouts = ttk.Treeview(left_frame, columns=cols, show="headings", height=15, selectmode="extended")
        
        headers = {"Week": "Week", "Date": "Date", "Session": "Sess", "Sport": "Sport", 
                   "Description": "Description", "ScheduledDate": "Sched."}
        widths = {"Week": 50, "Date": 90, "Session": 50, "Sport": 80, "Description": 220, "ScheduledDate": 90}
        
        for c in cols:
            self.tree_workouts.heading(c, text=headers[c])
            self.tree_workouts.column(c, width=widths[c], anchor="w")
        
        self.tree_workouts.tag_configure("scheduled", background="#d9fdd3")
        self.tree_workouts.tag_configure("uploaded", background="#fff4ce")
        self.tree_workouts.tag_configure("not_uploaded", background="#ffffff")
        
        self.tree_workouts.pack(fill=tk.BOTH, expand=True)
        self.tree_workouts.bind("<<TreeviewSelect>>", self._on_select_workout)
        
        # Right: Workout editor
        right_frame = ttk.Frame(main_pane, padding=5)
        main_pane.add(right_frame, weight=2)
        
        notebook = ttk.Notebook(right_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Tab: Visual Editor
        visual_frame = ttk.Frame(notebook, padding=5)
        notebook.add(visual_frame, text="🎨 Editor Visuale")
        
        # Metadata
        meta_frame = ttk.Frame(visual_frame)
        meta_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(meta_frame, text="Week:").grid(row=0, column=0, sticky="w")
        self.entry_week = ttk.Entry(meta_frame, width=5)
        self.entry_week.grid(row=0, column=1, sticky="w", padx=(0, 10))
        
        ttk.Label(meta_frame, text="Session:").grid(row=0, column=2, sticky="w")
        self.entry_session = ttk.Entry(meta_frame, width=5)
        self.entry_session.grid(row=0, column=3, sticky="w", padx=(0, 10))
        
        ttk.Label(meta_frame, text="Date:").grid(row=0, column=4, sticky="w")
        
        # Frame per Date + bottone calendario
        date_frame = ttk.Frame(meta_frame)
        date_frame.grid(row=0, column=5, sticky="w")
        
        # Date picker
        if HAS_CALENDAR:
            self.entry_date = DateEntry(
                date_frame,
                width=12,
                background='darkblue',
                foreground='white',
                borderwidth=2,
                date_pattern='yyyy-mm-dd',
                showweeknumbers=False
            )
            self.entry_date.pack(side=tk.LEFT)
            
            # Bottone per aprire calendario
            btn_cal = ttk.Button(date_frame, text="📅", width=3,
                                command=lambda: self.entry_date.drop_down())
            btn_cal.pack(side=tk.LEFT, padx=(2, 0))
        else:
            self.entry_date = ttk.Entry(date_frame, width=12)
            self.entry_date.pack(side=tk.LEFT)
        
        ttk.Label(visual_frame, text="Description:").pack(anchor="w")
        self.entry_description = ttk.Entry(visual_frame)
        self.entry_description.pack(fill=tk.X, pady=(0, 5))
        
        # Bottoni Salva/Elimina (visibili anche in Editor Visuale)
        btn_frame_visual = ttk.Frame(visual_frame)
        btn_frame_visual.pack(fill=tk.X, pady=5)
        
        btn_save_visual = ttk.Button(
            btn_frame_visual,
            text="💾 Salva Modifiche",
            command=self._save_workout_changes
        )
        btn_save_visual.pack(side=tk.LEFT, padx=2)
        
        btn_delete_visual = ttk.Button(
            btn_frame_visual,
            text="🗑️ Elimina Workout",
            command=self._delete_workout_from_list
        )
        btn_delete_visual.pack(side=tk.LEFT, padx=2)
        
        # Visual builder button
        btn_visual = ttk.Button(
            visual_frame,
            text="🎨 Apri Visual Builder",
            command=self._open_visual_builder
        )
        btn_visual.pack(pady=5)
        
        # Steps preview (read-only)
        ttk.Label(visual_frame, text="Steps Preview:").pack(anchor="w")
        self.text_steps_preview = tk.Text(visual_frame, wrap="none", height=15, state=tk.DISABLED)
        self.text_steps_preview.pack(fill=tk.BOTH, expand=True)
        
        # Tab: DSL Editor (testo)
        dsl_frame = ttk.Frame(notebook, padding=5)
        notebook.add(dsl_frame, text="📝 Editor DSL")
        
        ttk.Label(dsl_frame, text="Steps DSL (editing avanzato):").pack(anchor="w")
        self.text_steps_dsl = tk.Text(dsl_frame, wrap="none", height=20)
        self.text_steps_dsl.pack(fill=tk.BOTH, expand=True)
        
        btn_save = ttk.Button(dsl_frame, text="💾 Salva Modifiche", command=self._save_workout_changes)
        btn_save.pack(side=tk.LEFT, pady=5, padx=2)
        
        btn_delete = ttk.Button(dsl_frame, text="🗑️ Elimina Workout", command=self._delete_workout_from_list)
        btn_delete.pack(side=tk.LEFT, pady=5, padx=2)
        
        # Tab: Parameters
        params_frame = ttk.Frame(notebook, padding=5)
        notebook.add(params_frame, text="⚙️ Parameters")
        
        self.tree_params = ttk.Treeview(
            params_frame,
            columns=("Key", "Metric", "Expression", "Notes"),
            show="headings",
            height=15
        )
        
        for col in ["Key", "Metric", "Expression", "Notes"]:
            self.tree_params.heading(col, text=col)
        
        self.tree_params.column("Key", width=120)
        self.tree_params.column("Metric", width=70)
        self.tree_params.column("Expression", width=120)
        self.tree_params.column("Notes", width=220)
        
        self.tree_params.pack(fill=tk.BOTH, expand=True)
        
        ttk.Button(params_frame, text="✏️ Modifica", command=self._edit_parameter).pack(pady=5)
    
    # ========== Excel Operations ==========
    
    def _create_new_workout(self):
        """Crea un nuovo workout e lo aggiunge all'Excel."""
        if self.df_workouts is None:
            messagebox.showerror("Errore", "Devi prima caricare o generare un file Excel.")
            return
        
        # Dialog per i dettagli del nuovo workout
        dialog = NewWorkoutDialog(self.root)
        self.root.wait_window(dialog)
        
        if not dialog.result:
            return  # Utente ha annullato
        
        # Crea nuova riga
        new_row = {
            "Week": dialog.result["week"],
            "Date": dialog.result["date"],
            "Session": dialog.result["session"],
            "Sport": dialog.result["sport"],
            "Description": dialog.result["description"],
            "Steps": "",  # Vuoto, verrà compilato con il visual builder
            "WorkoutId": "",
            "ScheduledDate": "",
        }
        
        # Aggiungi WorkoutScheduleId se presente nel DataFrame
        if "WorkoutScheduleId" in self.df_workouts.columns:
            new_row["WorkoutScheduleId"] = ""
        
        # Aggiungi al DataFrame usando pd.concat invece di append
        new_row_df = pd.DataFrame([new_row])
        self.df_workouts = pd.concat([self.df_workouts, new_row_df], ignore_index=True)
        
        # Aggiorna la tree view
        self._populate_workouts_tree()
        
        # Seleziona il nuovo workout
        new_idx = len(self.df_workouts) - 1
        self.tree_workouts.selection_set(str(new_idx))
        self.tree_workouts.see(str(new_idx))
        
        # Trigger l'evento di selezione per popolare i campi
        self._on_select_workout()
        
        # Autosave
        self._autosave()
        
        # Messaggio con opzione di aprire subito il builder
        if messagebox.askyesno(
            "Workout Creato",
            f"Nuovo workout creato!\n\n"
            f"Vuoi aprire subito il Visual Builder\n"
            f"per definire gli step del workout?",
            icon='question'
        ):
            self._open_visual_builder()
    
    def _generate_excel(self):
        """Genera file Excel di esempio."""
        path = filedialog.asksaveasfilename(
            title="Salva Excel di esempio",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not path:
            return
        
        try:
            generate_training_excel_multisport(path)
            messagebox.showinfo("OK", f"File di esempio creato:\n{path}")
        except Exception as e:
            messagebox.showerror("Errore", f"Errore nella generazione:\n{e}")
    
    def _load_excel(self):
        """Carica file Excel."""
        path = filedialog.askopenfilename(
            title="Seleziona Excel",
            filetypes=[("Excel", "*.xlsx *.xls")]
        )
        if not path:
            return
        
        try:
            self.df_workouts = pd.read_excel(path, sheet_name="Workouts")
            
            try:
                self.df_parameters = pd.read_excel(
                    path,
                    sheet_name="Parameters",
                    dtype={"Key": str, "Metric": str, "Expression": str, "Notes": str}
                )
            except:
                self.df_parameters = pd.DataFrame(columns=["Key", "Metric", "Expression", "Notes"])
            
            for col in ["WorkoutId", "ScheduledDate", "WorkoutScheduleId"]:
                if col not in self.df_workouts.columns:
                    self.df_workouts[col] = ""
            
            self.df_workouts["WorkoutId"] = self.df_workouts["WorkoutId"].astype("string")
            self.df_workouts["ScheduledDate"] = self.df_workouts["ScheduledDate"].astype("string")
            
            self.excel_path = path
            self.lbl_file.config(text=path)
            
            self._populate_workouts_tree()
            self._populate_params_tree()
            
            # Clear fields
            self.entry_week.delete(0, tk.END)
            self.entry_date.delete(0, tk.END)
            self.entry_session.delete(0, tk.END)
            self.entry_description.delete(0, tk.END)
            self.text_steps_dsl.delete("1.0", tk.END)
            
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile caricare Excel:\n{e}")
    
    def _save_excel(self):
        """Salva file Excel."""
        if self.df_workouts is None:
            messagebox.showerror("Errore", "Nessun allenamento caricato.")
            return
        
        path = filedialog.asksaveasfilename(
            title="Salva Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not path:
            return
        
        try:
            self._write_excel(path)
            messagebox.showinfo("OK", f"File salvato:\n{path}")
        except Exception as e:
            messagebox.showerror("Errore", f"Errore salvataggio:\n{e}")
    
    def _write_excel(self, path: str):
        """Scrive il DataFrame su Excel."""
        if self.df_workouts is None:
            return
        
        df_temp = self.df_workouts.copy()
        
        if "Date" in df_temp.columns:
            df_temp["Date"] = pd.to_datetime(df_temp["Date"], errors="coerce", dayfirst=True)
        
        if "ScheduledDate" in df_temp.columns:
            def to_date_only(val):
                if pd.isna(val) or val == "" or str(val).strip().lower() in ("nan", "<na>"):
                    return None
                try:
                    dt = pd.to_datetime(val, errors="coerce")
                    return dt.date() if pd.notna(dt) else None
                except:
                    return None
            df_temp["ScheduledDate"] = df_temp["ScheduledDate"].apply(to_date_only)
        
        # Preserva fogli esistenti
        existing_sheets = {}
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path)
            if "Esempi DSL" in wb.sheetnames:
                df_esempi = pd.read_excel(path, sheet_name="Esempi DSL")
                existing_sheets["Esempi DSL"] = df_esempi
            wb.close()
        except:
            pass
        
        with pd.ExcelWriter(path, engine="openpyxl", mode='w') as writer:
            df_temp.to_excel(writer, sheet_name="Workouts", index=False)
            
            if self.df_parameters is not None:
                self.df_parameters.to_excel(writer, sheet_name="Parameters", index=False)
            
            for sheet_name, df in existing_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        format_workbook_dates_and_steps(path)
    
    def _autosave(self):
        """Salva automaticamente su file caricato."""
        if not self.excel_path or self.df_workouts is None:
            return
        
        try:
            self._write_excel(self.excel_path)
        except Exception as e:
            messagebox.showwarning("Autosave fallito", f"Non riesco a salvare:\n{e}")
    
    # ========== Workouts Management ==========
    
    def _populate_workouts_tree(self):
        """Popola la tree view dei workout."""
        for item in self.tree_workouts.get_children():
            self.tree_workouts.delete(item)
        
        if self.df_workouts is None:
            return
        
        for idx, row in self.df_workouts.iterrows():
            week = row.get("Week", "")
            date = self._format_date(row.get("Date", ""))
            session = row.get("Session", "")
            sport = row.get("Sport", "")
            desc = row.get("Description", "")
            sched = row.get("ScheduledDate", "")
            
            if pd.isna(sched) or str(sched).strip().lower() in ("nan", "<na>"):
                sched = ""
            
            vals = (week, date, session, sport, desc, sched)
            iid = str(idx)
            self.tree_workouts.insert("", tk.END, iid=iid, values=vals)
            self._apply_row_style(idx)
    
    def _format_date(self, value) -> str:
        """Formatta data per display."""
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return ""
        
        if isinstance(value, (pd.Timestamp, datetime)):
            return value.date().isoformat()
        
        from datetime import date as date_cls
        if isinstance(value, date_cls):
            return value.isoformat()
        
        s = str(value).strip()
        if " " in s:
            s = s.split(" ")[0]
        return s
    
    def _apply_row_style(self, idx: int):
        """Applica colore riga in base allo stato."""
        if self.df_workouts is None:
            return
        
        row = self.df_workouts.iloc[idx]
        iid = str(idx)
        
        workout_id = str(row.get("WorkoutId", "")).strip()
        sched_date = str(row.get("ScheduledDate", "")).strip()
        
        if workout_id.lower() in ("nan", "<na>"):
            workout_id = ""
        if sched_date.lower() in ("nan", "<na>"):
            sched_date = ""
        
        if sched_date:
            tag = "scheduled"
        elif workout_id:
            tag = "uploaded"
        else:
            tag = "not_uploaded"
        
        self.tree_workouts.item(iid, tags=(tag,))
    
    def _refresh_tree_row(self, idx: int):
        """Aggiorna riga nella tree."""
        if self.df_workouts is None:
            return
        
        row = self.df_workouts.iloc[idx]
        iid = str(idx)
        
        week = row.get("Week", "")
        date = self._format_date(row.get("Date", ""))
        session = row.get("Session", "")
        sport = row.get("Sport", "")
        desc = row.get("Description", "")
        sched = row.get("ScheduledDate", "")
        
        if pd.isna(sched) or str(sched).strip().lower() in ("nan", "<na>"):
            sched = ""
        
        vals = (week, date, session, sport, desc, sched)
        
        if iid in self.tree_workouts.get_children():
            self.tree_workouts.item(iid, values=vals)
            self._apply_row_style(idx)
    
    def _on_select_workout(self, event=None):
        """Gestisce selezione workout."""
        if self.df_workouts is None:
            return
        
        sel = self.tree_workouts.selection()
        if not sel:
            return
        
        idx = int(sel[-1])
        row = self.df_workouts.iloc[idx]
        
        # Popola campi
        self.entry_week.delete(0, tk.END)
        self.entry_week.insert(0, str(row.get("Week", "")))
        
        self.entry_date.delete(0, tk.END)
        self.entry_date.insert(0, self._format_date(row.get("Date", "")))
        
        self.entry_session.delete(0, tk.END)
        self.entry_session.insert(0, str(row.get("Session", "")))
        
        self.entry_description.delete(0, tk.END)
        self.entry_description.insert(0, str(row.get("Description", "")))
        
        steps_text = str(row.get("Steps", ""))
        
        # DSL editor
        self.text_steps_dsl.delete("1.0", tk.END)
        self.text_steps_dsl.insert("1.0", steps_text)
        
        # Preview (read-only)
        self.text_steps_preview.config(state=tk.NORMAL)
        self.text_steps_preview.delete("1.0", tk.END)
        self.text_steps_preview.insert("1.0", steps_text)
        self.text_steps_preview.config(state=tk.DISABLED)
    
    def _open_visual_builder(self):
        """Apre il visual builder."""
        sel = self.tree_workouts.selection()
        if not sel:
            messagebox.showwarning("Attenzione", "Seleziona un workout.")
            return
        
        # Get current DSL
        current_dsl = self.text_steps_dsl.get("1.0", tk.END).strip()
        
        # Callback per salvare
        def save_callback(new_dsl: str):
            # Aggiorna i campi di testo
            self.text_steps_dsl.delete("1.0", tk.END)
            self.text_steps_dsl.insert("1.0", new_dsl)
            
            self.text_steps_preview.config(state=tk.NORMAL)
            self.text_steps_preview.delete("1.0", tk.END)
            self.text_steps_preview.insert("1.0", new_dsl)
            self.text_steps_preview.config(state=tk.DISABLED)
            
            # 🔥 SALVA AUTOMATICAMENTE NEL DATAFRAME
            sel = self.tree_workouts.selection()
            if sel and self.df_workouts is not None:
                for iid in sel:
                    idx = int(iid)
                    self.df_workouts.at[idx, "Steps"] = new_dsl
                
                # Autosave su file
                self._autosave()
                
                # Notifica visiva
                messagebox.showinfo(
                    "✅ Salvato",
                    "Workout salvato automaticamente!\n\n"
                    "Gli step sono stati aggiornati nell'Excel.",
                    parent=self.root
                )
        
        # Apri builder
        builder = VisualWorkoutBuilder(self.root, initial_dsl=current_dsl, callback=save_callback)
    
    def _save_workout_changes(self):
        """Salva modifiche al workout corrente."""
        if self.df_workouts is None:
            return
        
        sel = self.tree_workouts.selection()
        if not sel:
            messagebox.showwarning("Attenzione", "Seleziona un workout.")
            return
        
        week_str = self.entry_week.get().strip()
        date_val = self.entry_date.get().strip()
        session_str = self.entry_session.get().strip()
        desc = self.entry_description.get().strip()
        steps_text = self.text_steps_dsl.get("1.0", tk.END).rstrip("\n")
        
        for iid in sel:
            idx = int(iid)
            
            if week_str:
                try:
                    self.df_workouts.at[idx, "Week"] = int(week_str)
                except:
                    self.df_workouts.at[idx, "Week"] = week_str
            
            if date_val:
                self.df_workouts.at[idx, "Date"] = date_val
            
            if session_str:
                try:
                    self.df_workouts.at[idx, "Session"] = int(session_str)
                except:
                    self.df_workouts.at[idx, "Session"] = session_str
            
            self.df_workouts.at[idx, "Description"] = desc
            self.df_workouts.at[idx, "Steps"] = steps_text
            
            self._refresh_tree_row(idx)
        
        self._autosave()
        messagebox.showinfo("OK", "Workout aggiornato!")
    
    def _delete_workout_from_list(self):
        """Elimina il workout selezionato dalla lista Excel (non da Garmin)."""
        if self.df_workouts is None:
            return
        
        sel = self.tree_workouts.selection()
        if not sel:
            messagebox.showwarning("Attenzione", "Seleziona un workout da eliminare.")
            return
        
        # Conferma
        if not messagebox.askyesno(
            "Conferma Eliminazione",
            "Vuoi eliminare questo workout dalla lista Excel?\n\n"
            "Nota: se il workout è già caricato su Garmin,\n"
            "dovrai cancellarlo separatamente con il bottone '🗑 Cancella'."
        ):
            return
        
        # Elimina tutte le righe selezionate
        indices_to_delete = sorted([int(iid) for iid in sel], reverse=True)
        
        for idx in indices_to_delete:
            self.df_workouts = self.df_workouts.drop(idx).reset_index(drop=True)
        
        # Aggiorna tree e autosave
        self._populate_workouts_tree()
        self._autosave()
        
        # Clear campi editor
        self.entry_week.delete(0, tk.END)
        self.entry_date.delete(0, tk.END)
        self.entry_session.delete(0, tk.END)
        self.entry_description.delete(0, tk.END)
        self.text_steps_dsl.delete("1.0", tk.END)
        self.text_steps_preview.config(state=tk.NORMAL)
        self.text_steps_preview.delete("1.0", tk.END)
        self.text_steps_preview.config(state=tk.DISABLED)
        
        messagebox.showinfo("OK", f"Eliminati {len(indices_to_delete)} workout dalla lista!")
    
    # ========== Parameters ==========
    
    def _populate_params_tree(self):
        """Popola tree dei parametri."""
        for item in self.tree_params.get_children():
            self.tree_params.delete(item)
        
        if self.df_parameters is None or self.df_parameters.empty:
            return
        
        for idx, row in self.df_parameters.iterrows():
            vals = (
                row.get("Key", ""),
                row.get("Metric", ""),
                row.get("Expression", ""),
                row.get("Notes", "")
            )
            self.tree_params.insert("", tk.END, iid=str(idx), values=vals)
    
    def _edit_parameter(self):
        """Modifica parametro selezionato."""
        if self.df_parameters is None or self.df_parameters.empty:
            messagebox.showwarning("Attenzione", "Nessun parametro.")
            return
        
        sel = self.tree_params.selection()
        if not sel:
            messagebox.showwarning("Attenzione", "Seleziona un parametro.")
            return
        
        idx = int(sel[0])
        row = self.df_parameters.iloc[idx]
        
        # Dialog semplice
        win = tk.Toplevel(self.root)
        win.title("Modifica Parametro")
        win.geometry("400x200")
        
        ttk.Label(win, text="Key:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        e_key = ttk.Entry(win)
        e_key.grid(row=0, column=1, sticky="we", padx=5, pady=5)
        e_key.insert(0, str(row.get("Key", "")))
        
        ttk.Label(win, text="Metric:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        e_metric = ttk.Entry(win)
        e_metric.grid(row=1, column=1, sticky="we", padx=5, pady=5)
        e_metric.insert(0, str(row.get("Metric", "")))
        
        ttk.Label(win, text="Expression:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        e_expr = ttk.Entry(win)
        e_expr.grid(row=2, column=1, sticky="we", padx=5, pady=5)
        e_expr.insert(0, str(row.get("Expression", "")))
        
        ttk.Label(win, text="Notes:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        e_notes = ttk.Entry(win)
        e_notes.grid(row=3, column=1, sticky="we", padx=5, pady=5)
        e_notes.insert(0, str(row.get("Notes", "")))
        
        win.columnconfigure(1, weight=1)
        
        def save():
            self.df_parameters.at[idx, "Key"] = e_key.get().strip()
            self.df_parameters.at[idx, "Metric"] = e_metric.get().strip()
            self.df_parameters.at[idx, "Expression"] = e_expr.get().strip()
            self.df_parameters.at[idx, "Notes"] = e_notes.get().strip()
            self._populate_params_tree()
            self._autosave()
            win.destroy()
        
        ttk.Button(win, text="💾 Salva", command=save).grid(row=4, column=0, columnspan=2, pady=10)
    
    # ========== Garmin Operations ==========
    
    def _login_credentials(self):
        """Login con credenziali."""
        email = self.entry_email.get().strip()
        password = self.entry_password.get().strip()
        
        if not email or not password:
            messagebox.showerror("Errore", "Inserisci email e password.")
            return
        
        try:
            self.garmin.login_with_credentials(email, password)
            self.lbl_garmin_status.config(text="● Connesso", foreground="green")
            messagebox.showinfo("OK", "Login effettuato!")
        except Exception as e:
            self.lbl_garmin_status.config(text="● Non connesso", foreground="red")
            messagebox.showerror("Errore", f"Errore login:\n{e}")
    
    def _login_session(self):
        """Login con sessione salvata."""
        try:
            self.garmin.login_with_saved_session()
            self.lbl_garmin_status.config(text="● Connesso", foreground="green")
            messagebox.showinfo("OK", "Login con sessione effettuato!")
        except Exception as e:
            self.lbl_garmin_status.config(text="● Non connesso", foreground="red")
            messagebox.showerror("Errore", f"Errore sessione:\n{e}")
    
    def _get_selected_indices(self) -> List[int]:
        """Ottiene indici workout selezionati."""
        if self.df_workouts is None:
            messagebox.showerror("Errore", "Nessun Excel caricato.")
            return []
        
        sel = self.tree_workouts.selection()
        if not sel:
            messagebox.showerror("Errore", "Seleziona almeno un workout.")
            return []
        
        return [int(iid) for iid in sel]
    
    def _run_with_loading(self, operation_func, title="Operazione in corso..."):
        """Esegue operazione con loading dialog."""
        loading = LoadingDialog(self.root, title)
        result = {"success": False, "message": ""}
        
        def worker():
            try:
                success, message = operation_func(loading)
                result["success"] = success
                result["message"] = message
            except Exception as e:
                result["success"] = False
                result["message"] = str(e)
            finally:
                self.root.after(0, loading.close)
        
        thread = threading.Thread(target=worker, daemon=True)
        thread.start()
        
        while thread.is_alive():
            self.root.update()
            thread.join(timeout=0.1)
        
        if result["success"]:
            messagebox.showinfo("OK", result["message"])
        else:
            messagebox.showerror("Errore", result["message"])
    
    def _upload_workouts(self):
        """Carica workout su Garmin."""
        if self.garmin.client is None:
            messagebox.showerror("Errore", "Non connesso a Garmin.")
            return
        
        idxs = self._get_selected_indices()
        if not idxs:
            return
        
        def operation(loading):
            created = 0
            total = len(idxs)
            
            for i, idx in enumerate(idxs, 1):
                loading.update_message(f"Caricamento {i}/{total}...")
                
                row = self.df_workouts.iloc[idx]
                steps_text = str(row.get("Steps", ""))
                
                workout_data = build_garmin_workout_from_excel_row(
                    row, steps_text, self.df_parameters,
                    use_prefix=self.use_prefix_var.get()
                )
                
                resp = self.garmin.save_workout(workout_data)
                workout_id = resp.get("workoutId") or resp.get("workout_id")
                
                if workout_id:
                    self.df_workouts.at[idx, "WorkoutId"] = str(workout_id)
                    self.root.after(0, lambda idx=idx: self._refresh_tree_row(idx))
                    created += 1
            
            self.root.after(0, self._autosave)
            return True, f"Caricati {created} workout!"
        
        self._run_with_loading(operation, "Caricamento workout")
    
    def _upload_and_schedule(self):
        """Carica e pianifica workout."""
        if self.garmin.client is None:
            messagebox.showerror("Errore", "Non connesso a Garmin.")
            return
        
        idxs = self._get_selected_indices()
        if not idxs:
            return
        
        def operation(loading):
            total = 0
            num = len(idxs)
            
            for i, idx in enumerate(idxs, 1):
                loading.update_message(f"Elaborazione {i}/{num}...")
                
                row = self.df_workouts.iloc[idx]
                steps_text = str(row.get("Steps", ""))
                
                date_val = row.get("Date", "")
                if isinstance(date_val, str):
                    date_str = date_val.strip()
                else:
                    try:
                        date_str = pd.to_datetime(date_val).date().isoformat()
                    except:
                        raise ValueError(f"Data non valida per workout {idx}")
                
                # Valida formato
                try:
                    datetime.strptime(date_str, "%Y-%m-%d")
                except:
                    raise ValueError(f"Formato data non valido: {date_str}")
                
                # Crea se serve
                workout_id = str(row.get("WorkoutId", "")).strip()
                if workout_id.lower() in ("nan", "<na>", ""):
                    workout_id = ""
                
                if not workout_id:
                    loading.update_message(f"Creazione workout {i}/{num}...")
                    
                    workout_data = build_garmin_workout_from_excel_row(
                        row, steps_text, self.df_parameters,
                        use_prefix=self.use_prefix_var.get()
                    )
                    
                    resp = self.garmin.save_workout(workout_data)
                    workout_id = str(resp.get("workoutId") or resp.get("workout_id") or "").strip()
                    
                    if not workout_id:
                        continue
                    
                    self.df_workouts.at[idx, "WorkoutId"] = workout_id
                    self.root.after(0, lambda idx=idx: self._refresh_tree_row(idx))
                
                # Pianifica
                loading.update_message(f"Pianificazione {i}/{num}...")
                
                resp_sched = self.garmin.schedule_workout(workout_id, date_str)
                
                schedule_id = ""
                if isinstance(resp_sched, dict):
                    schedule_id = str(
                        resp_sched.get("workoutScheduleId") or
                        resp_sched.get("workout_schedule_id") or
                        resp_sched.get("eventId") or
                        resp_sched.get("id") or ""
                    ).strip()
                
                if schedule_id:
                    self.df_workouts.at[idx, "WorkoutScheduleId"] = schedule_id
                
                self.df_workouts.at[idx, "ScheduledDate"] = date_str
                self.root.after(0, lambda idx=idx: self._refresh_tree_row(idx))
                total += 1
            
            self.root.after(0, self._autosave)
            return True, f"Pianificati {total} workout!"
        
        self._run_with_loading(operation, "Caricamento e pianificazione")
    
    def _unschedule(self):
        """Rimuove pianificazione."""
        if self.garmin.client is None:
            messagebox.showerror("Errore", "Non connesso a Garmin.")
            return
        
        idxs = self._get_selected_indices()
        if not idxs:
            return
        
        def operation(loading):
            removed = 0
            num = len(idxs)
            
            for i, idx in enumerate(idxs, 1):
                loading.update_message(f"Rimozione {i}/{num}...")
                
                row = self.df_workouts.iloc[idx]
                
                schedule_id = str(row.get("WorkoutScheduleId", "")).strip()
                sched_date = str(row.get("ScheduledDate", "")).strip()
                
                if schedule_id.lower() in ("nan", "<na>", ""):
                    schedule_id = ""
                if sched_date.lower() in ("nan", "<na>", ""):
                    sched_date = ""
                
                if not schedule_id or not sched_date:
                    continue
                
                try:
                    self.garmin.unschedule_workout(schedule_id, sched_date)
                    self.df_workouts.at[idx, "ScheduledDate"] = ""
                    self.df_workouts.at[idx, "WorkoutScheduleId"] = ""
                    self.root.after(0, lambda idx=idx: self._refresh_tree_row(idx))
                    removed += 1
                except Exception as e:
                    if "403" in str(e):
                        raise RuntimeError(f"403 Forbidden - API non autorizzata per scheduleId {schedule_id}")
                    raise
            
            if removed:
                self.root.after(0, self._autosave)
                return True, f"Rimossi {removed} workout!"
            else:
                return True, "Nessun workout rimosso."
        
        self._run_with_loading(operation, "Rimozione pianificazione")
    
    def _delete_workouts(self):
        """Cancella workout da Garmin."""
        if self.garmin.client is None:
            messagebox.showerror("Errore", "Non connesso a Garmin.")
            return
        
        idxs = self._get_selected_indices()
        if not idxs:
            return
        
        if not messagebox.askyesno("Conferma", "Cancellare DEFINITIVAMENTE i workout da Garmin?"):
            return
        
        def operation(loading):
            deleted = 0
            num = len(idxs)
            
            for i, idx in enumerate(idxs, 1):
                loading.update_message(f"Cancellazione {i}/{num}...")
                
                row = self.df_workouts.iloc[idx]
                
                workout_id = str(row.get("WorkoutId", "")).strip()
                if workout_id.lower() in ("nan", "<na>", ""):
                    continue
                
                # Rimuovi pianificazione se esiste
                schedule_id = str(row.get("WorkoutScheduleId", "")).strip()
                sched_date = str(row.get("ScheduledDate", "")).strip()
                
                if schedule_id and schedule_id.lower() not in ("nan", "<na>", "") and \
                   sched_date and sched_date.lower() not in ("nan", "<na>", ""):
                    try:
                        loading.update_message(f"Rimozione pianificazione {i}/{num}...")
                        self.garmin.unschedule_workout(schedule_id, sched_date)
                    except:
                        pass
                
                # Cancella workout
                loading.update_message(f"Cancellazione workout {i}/{num}...")
                self.garmin.delete_workout(workout_id)
                
                self.df_workouts.at[idx, "WorkoutId"] = ""
                if "WorkoutScheduleId" in self.df_workouts.columns:
                    self.df_workouts.at[idx, "WorkoutScheduleId"] = ""
                if "ScheduledDate" in self.df_workouts.columns:
                    self.df_workouts.at[idx, "ScheduledDate"] = ""
                
                self.root.after(0, lambda idx=idx: self._refresh_tree_row(idx))
                deleted += 1
            
            if deleted:
                self.root.after(0, self._autosave)
                return True, f"Cancellati {deleted} workout!"
            else:
                return True, "Nessun workout cancellato."
        
        self._run_with_loading(operation, "Cancellazione workout")


# ============================================================================
# MAIN
# ============================================================================

def main():
    root = tk.Tk()
    
    # Dimensioni finestra
    root.geometry("1400x800")
    root.minsize(1200, 700)
    app = TrainingPlannerAdvancedGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()