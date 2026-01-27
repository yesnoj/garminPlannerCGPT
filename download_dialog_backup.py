#!/usr/bin/env python3
"""
Dialog per download workout/attività da Garmin Connect
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import json
from pathlib import Path
import garth  # Importo garth per chiamate dirette API


class DownloadDialog:
    """Finestra dialogo per scaricare workout da Garmin."""
    
    SPORT_TYPES = {
        "Tutti gli sport": None,
        "Corsa": "running",
        "Ciclismo": "cycling",
        "Nuoto": "swimming",
        "Triathlon": "multi_sport",
        "Altro": "other"
    }
    
    def __init__(self, parent, garmin_client):
        """
        Args:
            parent: Finestra genitore
            garmin_client: Istanza di GarminClient già autenticato
        """
        self.parent = parent
        self.garmin = garmin_client
        self.result = None
        
        # Crea finestra modale
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Download Workout da Garmin")
        self.dialog.geometry("550x650")  # Aumentato a 650 per campo nome file editabile
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self._build_ui()
        
        # Centra la finestra
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (self.dialog.winfo_width() // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f"+{x}+{y}")
    
    def _build_ui(self):
        """Costruisce l'interfaccia."""
        main_frame = ttk.Frame(self.dialog, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Titolo
        title = ttk.Label(
            main_frame,
            text="📥 Download Workout e Attività",
            font=("", 12, "bold")
        )
        title.pack(pady=(0, 15))
        
        # ========== SELEZIONE DATE ==========
        date_frame = ttk.LabelFrame(main_frame, text="📅 Periodo", padding=10)
        date_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Data inizio
        ttk.Label(date_frame, text="Dal:").grid(row=0, column=0, sticky="w", pady=5)
        
        today = datetime.now()
        start_default = today - timedelta(days=365)  # Ultimo anno
        
        self.date_start = DateEntry(
            date_frame,
            width=20,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            year=start_default.year,
            month=start_default.month,
            day=start_default.day,
            date_pattern='yyyy-mm-dd'
        )
        self.date_start.grid(row=0, column=1, sticky="w", padx=(10, 0), pady=5)
        self.date_start.bind('<<DateEntrySelected>>', lambda e: self._update_filename())
        
        # Data fine
        ttk.Label(date_frame, text="Al:").grid(row=1, column=0, sticky="w", pady=5)
        
        self.date_end = DateEntry(
            date_frame,
            width=20,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            year=today.year,
            month=today.month,
            day=today.day,
            date_pattern='yyyy-mm-dd'
        )
        self.date_end.grid(row=1, column=1, sticky="w", padx=(10, 0), pady=5)
        self.date_end.bind('<<DateEntrySelected>>', lambda e: self._update_filename())
        
        # Quick select
        quick_frame = ttk.Frame(date_frame)
        quick_frame.grid(row=2, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(
            quick_frame,
            text="Ultimo mese",
            command=lambda: self._set_quick_range(30)
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            quick_frame,
            text="Ultimi 3 mesi",
            command=lambda: self._set_quick_range(90)
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            quick_frame,
            text="Ultimo anno",
            command=lambda: self._set_quick_range(365)
        ).pack(side=tk.LEFT, padx=2)
        
        # ========== TIPO DI DATI ==========
        type_frame = ttk.LabelFrame(main_frame, text="📊 Cosa scaricare", padding=10)
        type_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.download_type = tk.StringVar(value="both")
        
        ttk.Radiobutton(
            type_frame,
            text="🗓️  Solo workout pianificati",
            variable=self.download_type,
            value="workouts",
            command=self._update_filename
        ).pack(anchor="w", pady=2)
        
        ttk.Radiobutton(
            type_frame,
            text="✅ Solo attività completate",
            variable=self.download_type,
            value="activities",
            command=self._update_filename
        ).pack(anchor="w", pady=2)
        
        ttk.Radiobutton(
            type_frame,
            text="📥 Entrambi (consigliato)",
            variable=self.download_type,
            value="both",
            command=self._update_filename
        ).pack(anchor="w", pady=2)
        
        # ========== FILTRO SPORT ==========
        sport_frame = ttk.LabelFrame(main_frame, text="🏃 Filtro Sport", padding=10)
        sport_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(sport_frame, text="Sport:").grid(row=0, column=0, sticky="w", pady=5)
        
        self.sport_combo = ttk.Combobox(
            sport_frame,
            values=list(self.SPORT_TYPES.keys()),
            state="readonly",
            width=30
        )
        self.sport_combo.set("Tutti gli sport")
        self.sport_combo.grid(row=0, column=1, sticky="w", padx=(10, 0), pady=5)
        
        # Bind per aggiornare il nome file quando cambia lo sport
        self.sport_combo.bind('<<ComboboxSelected>>', lambda e: self._update_filename())
        
        ttk.Label(
            sport_frame,
            text="ℹ️  Il filtro sport viene applicato dopo il download",
            foreground="gray",
            font=("", 8)
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(5, 0))
        
        # ========== DESTINAZIONE ==========
        dest_frame = ttk.LabelFrame(main_frame, text="💾 Salvataggio", padding=10)
        dest_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Nome file editabile
        ttk.Label(dest_frame, text="Nome file:").grid(row=0, column=0, sticky="w", pady=5)
        
        self.filename_entry = ttk.Entry(dest_frame, width=45)
        self.filename_entry.grid(row=0, column=1, sticky="we", padx=(10, 0), pady=5)
        
        dest_frame.columnconfigure(1, weight=1)
        
        # Genera nome iniziale
        self._update_filename()
        
        ttk.Label(
            dest_frame,
            text="💡 Il nome viene generato automaticamente in base ai filtri selezionati",
            foreground="gray",
            font=("", 8)
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(5, 0))
        
        # ========== PULSANTI AZIONE ==========
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X)
        
        ttk.Button(
            btn_frame,
            text="📥 Download",
            command=self._start_download
        ).pack(side=tk.LEFT, padx=(0, 5), ipadx=20)
        
        ttk.Button(
            btn_frame,
            text="❌ Annulla",
            command=self.dialog.destroy
        ).pack(side=tk.LEFT, ipadx=20)
        
        # Status label
        self.status_label = ttk.Label(
            main_frame,
            text="",
            foreground="blue",
            font=("", 9)
        )
        self.status_label.pack(pady=(10, 0))
    
    def _set_quick_range(self, days):
        """Imposta range veloce."""
        end = datetime.now()
        start = end - timedelta(days=days)
        
        self.date_start.set_date(start)
        self.date_end.set_date(end)
        self._update_filename()  # Aggiorna nome file dopo cambio date
    
    def _update_filename(self):
        """Genera nome file intelligente basato sui filtri selezionati."""
        try:
            # Ottieni date
            start_date = self.date_start.get_date()
            end_date = self.date_end.get_date()
            
            # Formato date per nome file
            if start_date.year == end_date.year:
                if start_date.month == end_date.month:
                    # Stesso mese: garmin_2025-01
                    date_part = f"{start_date.year}-{start_date.month:02d}"
                else:
                    # Stesso anno: garmin_2025
                    date_part = f"{start_date.year}"
            else:
                # Anni diversi: garmin_2024-2025
                date_part = f"{start_date.year}-{end_date.year}"
            
            # Tipo di dati
            download_type = self.download_type.get()
            if download_type == "workouts":
                type_part = "workouts"
            elif download_type == "activities":
                type_part = "activities"
            else:
                type_part = "data"  # "both" -> generico
            
            # Sport
            sport_selected = self.sport_combo.get()
            sport_filter = self.SPORT_TYPES.get(sport_selected)
            
            if sport_filter:  # Se c'è un filtro sport specifico
                # Usa il valore dell'API (running, cycling, swimming)
                sport_part = f"_{sport_filter}"
            else:
                sport_part = ""  # Tutti gli sport -> nessun suffisso
            
            # Componi nome file
            filename = f"garmin_{type_part}_{date_part}{sport_part}.json"
            
            # Aggiorna entry
            self.filename_entry.delete(0, tk.END)
            self.filename_entry.insert(0, filename)
            
        except Exception as e:
            # Fallback su nome generico
            self.filename_entry.delete(0, tk.END)
            self.filename_entry.insert(0, "garmin_data.json")
    
    def _start_download(self):
        """Avvia il download."""
        if self.garmin.client is None:
            messagebox.showerror("Errore", "Non sei connesso a Garmin Connect.")
            return
        
        # Valida date
        start_date = self.date_start.get_date()
        end_date = self.date_end.get_date()
        
        if start_date > end_date:
            messagebox.showerror("Errore", "La data di inizio deve essere precedente alla data di fine.")
            return
        
        # Formato ISO
        start_str = start_date.strftime("%Y-%m-%d")
        end_str = end_date.strftime("%Y-%m-%d")
        
        # Tipo di download
        download_type = self.download_type.get()
        
        # Sport filter
        sport_selected = self.sport_combo.get()
        sport_filter = self.SPORT_TYPES.get(sport_selected)
        
        # Nome file dall'entry (editabile dall'utente)
        filename = self.filename_entry.get().strip()
        
        # Valida nome file
        if not filename:
            messagebox.showerror("Errore", "Inserisci un nome file valido.")
            return
        
        # Aggiungi .json se manca
        if not filename.endswith('.json'):
            filename = filename + '.json'
        
        # Disabilita UI durante download
        self.status_label.config(text="⏳ Download in corso...", foreground="blue")
        self.dialog.update()
        
        try:
            # Download
            workouts = []
            activities = []
            
            if download_type in ("workouts", "both"):
                self.status_label.config(text="📥 Download workout pianificati...")
                self.dialog.update()
                workouts = self._download_workouts(start_str, end_str)
            
            if download_type in ("activities", "both"):
                self.status_label.config(text="📥 Download attività completate...")
                self.dialog.update()
                activities = self._download_activities(start_str, end_str)
            
            # Filtra per sport se necessario
            if sport_filter:
                workouts = self._filter_by_sport(workouts, sport_filter, "workout")
                activities = self._filter_by_sport(activities, sport_filter, "activity")
            
            # Salva risultati
            self.status_label.config(text="💾 Salvataggio dati...")
            self.dialog.update()
            
            self._save_data(workouts, activities, filename, start_str, end_str)
            
            # Successo
            total = len(workouts) + len(activities)
            message = f"✅ Download completato!\n\n"
            message += f"Workout pianificati: {len(workouts)}\n"
            message += f"Attività completate: {len(activities)}\n"
            message += f"Totale: {total}\n\n"
            message += f"Salvato in: {filename}"
            
            messagebox.showinfo("Successo", message)
            self.result = filename
            self.dialog.destroy()
            
        except Exception as e:
            self.status_label.config(text="", foreground="red")
            messagebox.showerror("Errore", f"Errore durante il download:\n{str(e)}")
    
    def _download_workouts(self, start_date, end_date):
        """Scarica workout pianificati."""
        all_workouts = []
        
        # Prova tutti e 3 gli endpoint come nello script originale
        endpoints = [
            f"/workout-service/schedule/{start_date}/{end_date}",
            "/workout-service/workouts",  # Workout Library - AGGIUNTO!
            f"/workout-service/workouts/calendar/{start_date}/{end_date}",
        ]
        
        for url in endpoints:
            try:
                # Usa garth.connectapi direttamente (come nello script funzionante)
                response = garth.connectapi(url, method="GET")
                
                if hasattr(response, 'json'):
                    data = response.json()
                else:
                    data = response
                
                if data:
                    if isinstance(data, list):
                        all_workouts.extend(data)
                    elif isinstance(data, dict):
                        if 'workouts' in data:
                            all_workouts.extend(data['workouts'])
                        else:
                            all_workouts.append(data)
            except:
                pass  # Ignora errori endpoint non disponibili
        
        # Rimuovi duplicati per workoutId
        seen = set()
        unique = []
        for w in all_workouts:
            wid = w.get('workoutId') or w.get('workout_id')
            if wid and wid not in seen:
                seen.add(wid)
                unique.append(w)
        
        return unique
    
    def _download_activities(self, start_date, end_date):
        """Scarica attività completate."""
        try:
            url = f"/activitylist-service/activities/search/activities?startDate={start_date}&endDate={end_date}&limit=999"
            # Usa garth.connectapi direttamente (come nello script funzionante)
            response = garth.connectapi(url, method="GET")
            
            if hasattr(response, 'json'):
                data = response.json()
            else:
                data = response
            
            if data:
                return data if isinstance(data, list) else data.get('activityList', [])
            return []
        except:
            return []
    
    def _filter_by_sport(self, items, sport_filter, item_type):
        """Filtra per sport type."""
        filtered = []
        
        for item in items:
            sport_key = None
            
            if item_type == "workout":
                sport_info = item.get('sportType', {})
                if isinstance(sport_info, dict):
                    sport_key = sport_info.get('sportTypeKey', '').lower()
            else:  # activity
                sport_info = item.get('activityType', {})
                if isinstance(sport_info, dict):
                    sport_key = sport_info.get('typeKey', '').lower()
            
            # Match
            if sport_key and sport_filter.lower() in sport_key:
                filtered.append(item)
        
        return filtered
    
    def _save_data(self, workouts, activities, filename, start_date, end_date):
        """Salva dati in JSON."""
        def create_summary(items, date_field='scheduledDate'):
            """Crea statistiche."""
            summary = {'by_sport': {}, 'by_month': {}, 'total': len(items)}
            
            for item in items:
                # Sport
                sport_key = 'unknown'
                if 'sportType' in item:
                    sport_info = item['sportType']
                    if isinstance(sport_info, dict):
                        sport_key = sport_info.get('sportTypeKey', 'unknown')
                elif 'activityType' in item:
                    sport_info = item['activityType']
                    if isinstance(sport_info, dict):
                        sport_key = sport_info.get('typeKey', 'unknown')
                
                summary['by_sport'][sport_key] = summary['by_sport'].get(sport_key, 0) + 1
                
                # Date
                date_str = item.get(date_field) or item.get('startTimeLocal') or item.get('createdDate')
                if date_str:
                    try:
                        date = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
                        month = date.strftime('%Y-%m')
                        summary['by_month'][month] = summary['by_month'].get(month, 0) + 1
                    except:
                        pass
            
            return summary
        
        data = {
            'download_info': {
                'date': datetime.now().isoformat(),
                'period': {
                    'start': start_date,
                    'end': end_date
                }
            },
            'scheduled_workouts': {
                'total': len(workouts),
                'data': workouts,
                'summary': create_summary(workouts, 'scheduledDate')
            },
            'completed_activities': {
                'total': len(activities),
                'data': activities,
                'summary': create_summary(activities, 'startTimeLocal')
            }
        }
        
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    
    def show(self):
        """Mostra il dialog e attende la chiusura."""
        self.dialog.wait_window()
        return self.result


def show_download_dialog(parent, garmin_client):
    """
    Funzione helper per mostrare il dialog di download.
    
    Args:
        parent: Finestra genitore
        garmin_client: GarminClient autenticato
    
    Returns:
        str: Path del file salvato o None se annullato
    """
    dialog = DownloadDialog(parent, garmin_client)
    return dialog.show()