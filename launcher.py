#!/usr/bin/env python3
"""
Training Planner - Launcher Unificato
Mostra un popup all'avvio per scegliere quale versione usare.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sys
import subprocess
from pathlib import Path


class LauncherDialog:
    """Dialog di selezione versione GUI."""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Training Planner - Launcher")
        self.root.resizable(False, False)
        
        # Variabile per la scelta
        self.selected = None
        
        # Configurazione finestra
        self._setup_ui()
        self._center_window()
        
    def _setup_ui(self):
        """Crea l'interfaccia del launcher."""
        
        # Frame principale con padding
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = ttk.Label(
            header_frame,
            text="🏃 Training Planner",
            font=("", 16, "bold")
        )
        title_label.pack()
        
        subtitle_label = ttk.Label(
            header_frame,
            text="Seleziona la versione da avviare",
            font=("", 10)
        )
        subtitle_label.pack(pady=(5, 0))
        
        # Separatore
        ttk.Separator(main_frame, orient="horizontal").pack(fill=tk.X, pady=10)
        
        # Opzioni
        options_frame = ttk.Frame(main_frame)
        options_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Versione Standard
        self._create_option_button(
            options_frame,
            "📊 Standard",
            "Versione base con funzionalità essenziali\n"
            "✅ Upload workout su Garmin\n"
            "✅ Pianificazione calendario\n"
            "✅ Gestione Excel semplificata",
            "standard",
            0
        )
        
        # Versione Advanced
        self._create_option_button(
            options_frame,
            "⚡ Advanced",
            "Versione con funzionalità avanzate\n"
            "✅ Tutte le funzionalità Standard\n"
            "✅ Editor DSL integrato\n"
            "✅ Anteprima workout dettagliata\n"
            "✅ Export/Import configurazioni",
            "advanced",
            1
        )
        
        # Versione Multisport
        self._create_option_button(
            options_frame,
            "🏊 Multisport",
            "Supporto completo multisport\n"
            "✅ Running (ritmo, zone)\n"
            "✅ Cycling (potenza, cadenza)\n"
            "✅ Swimming (pace nuoto)\n"
            "✅ Zone HR universali",
            "multisport",
            2
        )
        
        # Footer
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Checkbox "Ricorda scelta"
        self.remember_var = tk.BooleanVar(value=False)
        remember_check = ttk.Checkbutton(
            footer_frame,
            text="Ricorda la mia scelta (avvia automaticamente)",
            variable=self.remember_var
        )
        remember_check.pack(pady=(0, 10))
        
        # Bottone Annulla
        cancel_btn = ttk.Button(
            footer_frame,
            text="Annulla",
            command=self._cancel
        )
        cancel_btn.pack(side=tk.RIGHT)
        
        # Info versione
        version_label = ttk.Label(
            footer_frame,
            text="v2.0 - Unified Edition",
            font=("", 8),
            foreground="gray"
        )
        version_label.pack(side=tk.LEFT)
        
    def _create_option_button(self, parent, title, description, value, row):
        """Crea un bottone opzione con descrizione."""
        
        # Frame per l'opzione
        option_frame = ttk.LabelFrame(parent, text=title, padding=10)
        option_frame.grid(row=row, column=0, sticky="ew", pady=5)
        
        # Descrizione
        desc_label = ttk.Label(
            option_frame,
            text=description,
            justify=tk.LEFT,
            wraplength=350
        )
        desc_label.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Bottone Avvia
        launch_btn = ttk.Button(
            option_frame,
            text="Avvia →",
            command=lambda: self._launch(value),
            width=12
        )
        launch_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
    def _center_window(self):
        """Centra la finestra sullo schermo."""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
    def _launch(self, version):
        """Avvia la versione selezionata."""
        self.selected = version
        
        # Salva preferenza se richiesto
        if self.remember_var.get():
            self._save_preference(version)
        
        self.root.destroy()
        
    def _cancel(self):
        """Annulla e chiudi."""
        self.selected = None
        self.root.destroy()
        
    def _save_preference(self, version):
        """Salva la preferenza in un file."""
        try:
            pref_file = Path.home() / ".training_planner_preference"
            with open(pref_file, 'w') as f:
                f.write(version)
            print(f"✅ Preferenza salvata: {version}")
        except Exception as e:
            print(f"⚠️  Impossibile salvare preferenza: {e}")
    
    def show(self):
        """Mostra il dialog e ritorna la scelta."""
        self.root.mainloop()
        return self.selected


def load_preference():
    """Carica la preferenza salvata se esiste."""
    try:
        pref_file = Path.home() / ".training_planner_preference"
        if pref_file.exists():
            with open(pref_file, 'r') as f:
                return f.read().strip()
    except Exception:
        pass
    return None


def clear_preference():
    """Cancella la preferenza salvata."""
    try:
        pref_file = Path.home() / ".training_planner_preference"
        if pref_file.exists():
            pref_file.unlink()
            print("✅ Preferenza cancellata")
    except Exception as e:
        print(f"⚠️  Errore cancellazione preferenza: {e}")


def launch_gui(version):
    """Avvia la GUI selezionata."""
    
    # Mappa versione -> file
    gui_files = {
        'standard': 'training_planner_gui.py',
        'advanced': 'training_planner_gui_advanced.py',
        'multisport': 'training_planner_multisport_gui.py'
    }
    
    gui_file = gui_files.get(version)
    if not gui_file:
        print(f"❌ Versione non valida: {version}")
        return False
    
    # Verifica che il file esista
    if not Path(gui_file).exists():
        messagebox.showerror(
            "Errore",
            f"File non trovato: {gui_file}\n\n"
            "Assicurati di eseguire questo script dalla directory del progetto."
        )
        return False
    
    # Avvia la GUI
    print(f"🚀 Avvio {version.title()} GUI...")
    print(f"   File: {gui_file}")
    
    try:
        # Avvia come subprocess per permettere al launcher di chiudersi
        subprocess.Popen([sys.executable, gui_file])
        return True
    except Exception as e:
        messagebox.showerror(
            "Errore Avvio",
            f"Impossibile avviare la GUI:\n\n{e}"
        )
        return False


def main():
    """Funzione principale del launcher."""
    
    print("╔══════════════════════════════════════════════════════════╗")
    print("║        Training Planner - Launcher Unificato            ║")
    print("╚══════════════════════════════════════════════════════════╝")
    print()
    
    # Check argomenti da riga di comando
    if len(sys.argv) > 1:
        arg = sys.argv[1].lower()
        
        # Reset preferenza
        if arg in ('--reset', '-r'):
            clear_preference()
            print("Riavvia il launcher per scegliere una nuova versione.")
            return
        
        # Avvio diretto
        if arg in ('standard', 'advanced', 'multisport', 's', 'a', 'm'):
            version_map = {'s': 'standard', 'a': 'advanced', 'm': 'multisport'}
            version = version_map.get(arg, arg)
            
            if launch_gui(version):
                print(f"✅ {version.title()} GUI avviata!")
            return
        
        # Help
        if arg in ('--help', '-h'):
            print("Uso:")
            print("  python launcher.py              # Mostra dialog di selezione")
            print("  python launcher.py standard     # Avvia versione Standard")
            print("  python launcher.py advanced     # Avvia versione Advanced")
            print("  python launcher.py multisport   # Avvia versione Multisport")
            print("  python launcher.py --reset      # Cancella preferenza salvata")
            print()
            print("Alias brevi: s=standard, a=advanced, m=multisport")
            return
    
    # Controlla se c'è una preferenza salvata
    preferred = load_preference()
    if preferred:
        print(f"ℹ️  Preferenza salvata trovata: {preferred.title()}")
        print(f"🚀 Avvio automatico...")
        print(f"   (Usa 'python launcher.py --reset' per cambiare)\n")
        
        if launch_gui(preferred):
            print(f"✅ {preferred.title()} GUI avviata!")
        return
    
    # Mostra dialog di selezione
    launcher = LauncherDialog()
    selected = launcher.show()
    
    if selected:
        if launch_gui(selected):
            print(f"✅ {selected.title()} GUI avviata!")
    else:
        print("ℹ️  Avvio annullato")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n⚠️  Avvio interrotto dall'utente")
    except Exception as e:
        print(f"\n❌ Errore: {e}")
        import traceback
        traceback.print_exc()
