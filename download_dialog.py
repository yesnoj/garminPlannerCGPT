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
import pandas as pd  # Per export Excel


class DownloadDialog:
    """Finestra dialogo per scaricare workout da Garmin."""
    
    SPORT_TYPES = {
        "Tutti gli sport": None,
        "Corsa": "running",
        "Ciclismo": "cycling",
        "Nuoto": "swimming"
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
        self.dialog.geometry("550x700")# Aumentato a 700 per campo nome file editabile
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
        
        # ========== FORMATO EXPORT ==========
        format_frame = ttk.LabelFrame(main_frame, text="📄 Formato Export", padding=10)
        format_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.export_format = tk.StringVar(value="excel")
        
        ttk.Radiobutton(
            format_frame,
            text="📊 Excel (.xlsx) - Consigliato per AI",
            variable=self.export_format,
            value="excel",
            command=self._update_filename
        ).pack(anchor="w", pady=2)
        
        ttk.Radiobutton(
            format_frame,
            text="📋 JSON (.json) - Dati grezzi completi",
            variable=self.export_format,
            value="json",
            command=self._update_filename
        ).pack(anchor="w", pady=2)
        
        ttk.Label(
            format_frame,
            text="ℹ️  Excel genera DSL dalle attività, JSON mantiene i dati originali",
            foreground="gray",
            font=("", 8)
        ).pack(anchor="w", pady=(5, 0))
        
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
            
            # Formato export
            export_format = self.export_format.get()
            extension = ".xlsx" if export_format == "excel" else ".json"
            
            # Componi nome file con estensione appropriata
            filename = f"garmin_{type_part}_{date_part}{sport_part}{extension}"
            
            # Aggiorna entry
            self.filename_entry.delete(0, tk.END)
            self.filename_entry.insert(0, filename)
            
        except Exception as e:
            # Fallback su nome generico con formato appropriato
            export_format = self.export_format.get() if hasattr(self, 'export_format') else "excel"
            extension = ".xlsx" if export_format == "excel" else ".json"
            self.filename_entry.delete(0, tk.END)
            self.filename_entry.insert(0, f"garmin_data{extension}")
    
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
        
        # Determina formato export e aggiungi estensione se manca
        export_format = self.export_format.get()
        if export_format == "excel":
            if not filename.endswith('.xlsx'):
                filename = filename + '.xlsx'
        else:  # json
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
            
            # Determina formato
            export_format = self.export_format.get()
            
            try:
                self._save_data(workouts, activities, filename, start_str, end_str, export_format)
            except Exception as save_error:
                # Log errore dettagliato
                import traceback
                error_details = traceback.format_exc()
                print(f"❌ ERRORE DURANTE SALVATAGGIO:")
                print(error_details)
                raise Exception(f"Errore nel salvare il file:\n{str(save_error)}\n\nDettagli tecnici:\n{error_details[:500]}")
            
            
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
    
    def _save_data(self, workouts, activities, filename, start_date, end_date, export_format="excel"):
        """Salva dati in Excel o JSON basato sul formato scelto."""
        
        if export_format == "json":
            # Export JSON (formato originale)
            self._save_data_json(workouts, activities, filename, start_date, end_date)
        else:
            # Export Excel (con generazione DSL)
            self._save_data_excel(workouts, activities, filename, start_date, end_date)
    
    def _save_data_json(self, workouts, activities, filename, start_date, end_date):
        """Salva dati in formato JSON (formato originale completo)."""
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
    
    def _save_data_excel(self, workouts, activities, filename, start_date, end_date):
        """Salva dati in Excel con formato compatibile per AI e re-upload."""
        
        # Limita numero righe per evitare problemi memoria
        MAX_ROWS = 100
        workouts = workouts[:MAX_ROWS]
        activities = activities[:MAX_ROWS]
        
        # Crea ExcelWriter
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            
            # ========== SHEET 1: WORKOUTS ==========
            if workouts:
                workout_rows = []
                for idx, w in enumerate(workouts, 1):
                    # Estrai informazioni base
                    name = w.get('workoutName', 'Unnamed Workout')
                    workout_id = w.get('workoutId', '')
                    
                    # Sport type
                    sport_type = w.get('sportType', {})
                    sport = sport_type.get('sportTypeKey', 'unknown') if isinstance(sport_type, dict) else 'unknown'
                    
                    # Data - usa createdDate come fallback
                    date_str = ''
                    week_num = ''
                    
                    created_date = w.get('createdDate', '')
                    if created_date:
                        try:
                            date_obj = datetime.fromisoformat(created_date.replace('Z', '+00:00'))
                            date_str = date_obj.strftime('%Y-%m-%d')
                            week_num = date_obj.isocalendar()[1]  # Week dell'anno
                        except:
                            pass
                    
                    # Metriche stimate
                    est_duration = w.get('estimatedDurationInSecs') or 0
                    est_distance = w.get('estimatedDistanceInMeters') or 0
                    
                    duration_min = f"{int(est_duration/60)}" if est_duration and est_duration > 0 else ""
                    distance_km = f"{est_distance/1000:.2f}" if est_distance and est_distance > 0 else ""
                    
                    # Description
                    description = w.get('description', '')
                    
                    workout_rows.append({
                        'Week': week_num,
                        'Date': date_str,
                        'Session': idx,
                        'Sport': sport,
                        'Description': name,
                        'Distance (km)': distance_km,
                        'Duration (min)': duration_min,
                        'Notes': description if description else 'Dalla libreria Garmin',
                        'WorkoutId': workout_id
                    })
                
                df_workouts = pd.DataFrame(workout_rows)
                df_workouts.to_excel(writer, sheet_name='Workouts', index=False)
            else:
                # Sheet vuoto con una riga di esempio per evitare errori Excel
                df_empty = pd.DataFrame([{
                    'Week': '',
                    'Date': '',
                    'Session': '',
                    'Sport': '',
                    'Description': '(Nessun workout nel periodo selezionato)',
                    'Distance (km)': '',
                    'Duration (min)': '',
                    'Notes': '',
                    'WorkoutId': ''
                }])
                df_empty.to_excel(writer, sheet_name='Workouts', index=False)
            
            # ========== SHEET 2: ACTIVITIES ==========
            if activities:
                activity_rows = []
                for idx, a in enumerate(activities, 1):
                    try:
                        # Estrai informazioni base
                        name = a.get('activityName', 'Activity')
                        activity_id = a.get('activityId', '')
                        
                        # Sport type
                        activity_type = a.get('activityType', {})
                        sport = activity_type.get('typeKey', 'unknown') if isinstance(activity_type, dict) else 'unknown'
                        
                        # Data
                        date_str = ''
                        time_str = ''
                        week_num = ''
                        
                        start_time = a.get('startTimeLocal', '')
                        if start_time:
                            try:
                                date_obj = datetime.fromisoformat(start_time.replace('Z', '+00:00'))
                                date_str = date_obj.strftime('%Y-%m-%d')
                                time_str = date_obj.strftime('%H:%M')
                                week_num = date_obj.isocalendar()[1]
                            except:
                                date_str = str(start_time)[:10] if start_time else ''
                                time_str = ''
                        
                        # Metriche (gestisci None)
                        distance = a.get('distance') or 0  # metri
                        duration = a.get('duration') or 0  # secondi
                        avg_hr = a.get('averageHR', '')
                        calories = a.get('calories', '')
                        avg_speed = a.get('averageSpeed') or 0  # m/s
                        
                        # Converti in formato leggibile
                        distance_km = f"{distance/1000:.2f}" if distance and distance > 0 else ""
                        duration_min = f"{int(duration//60)}:{int(duration%60):02d}" if duration and duration > 0 else ""
                        
                        # Calcola pace (min/km)
                        pace_str = ""
                        if distance and duration and distance > 0 and duration > 0:
                            pace_sec_per_km = (duration / (distance / 1000))
                            pace_min = int(pace_sec_per_km / 60)
                            pace_sec = int(pace_sec_per_km % 60)
                            pace_str = f"{pace_min}:{pace_sec:02d}"
                        
                        # Genera Steps DSL equivalente
                        steps_dsl = self._activity_to_dsl(a, sport)
                        
                        activity_rows.append({
                            'Week': week_num,
                            'Date': date_str,
                            'Time': time_str,
                            'Session': idx,
                            'Sport': sport,
                            'Description': name,
                            'Distance (km)': distance_km,
                            'Duration': duration_min,
                            'Avg Pace (min/km)': pace_str,
                            'Avg HR (bpm)': avg_hr if avg_hr else '',
                            'Calories': calories if calories else '',
                            'Steps': steps_dsl,
                            'ActivityId': activity_id
                        })
                    except Exception as e:
                        # Skip questa attività ma continua
                        print(f"⚠️ Skip attività {idx}: {str(e)}")
                        continue
                
                if activity_rows:
                    df_activities = pd.DataFrame(activity_rows)
                    df_activities.to_excel(writer, sheet_name='Activities', index=False)
                else:
                    # Nessuna attività valida
                    df_empty = pd.DataFrame([{
                        'Week': '',
                        'Date': '',
                        'Time': '',
                        'Session': '',
                        'Sport': '',
                        'Description': '(Errore processando attività)',
                        'Distance (km)': '',
                        'Duration': '',
                        'Avg Pace (min/km)': '',
                        'Avg HR (bpm)': '',
                        'Calories': '',
                        'Steps': '',
                        'ActivityId': ''
                    }])
                    df_empty.to_excel(writer, sheet_name='Activities', index=False)
            else:
                # Sheet vuoto con una riga di esempio
                df_empty = pd.DataFrame([{
                    'Week': '',
                    'Date': '',
                    'Time': '',
                    'Session': '',
                    'Sport': '',
                    'Description': '(Nessuna attività nel periodo selezionato)',
                    'Distance (km)': '',
                    'Duration': '',
                    'Avg Pace (min/km)': '',
                    'Avg HR (bpm)': '',
                    'Calories': '',
                    'Steps': '',
                    'ActivityId': ''
                }])
                df_empty.to_excel(writer, sheet_name='Activities', index=False)
            
            # ========== SHEET 3: SUMMARY ==========
            summary_data = {
                'Metric': [
                    'Download Date',
                    'Period Start',
                    'Period End',
                    'Total Workouts',
                    'Total Activities',
                    '',
                    'Workouts by Sport',
                ],
                'Value': [
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    start_date,
                    end_date,
                    len(workouts),
                    len(activities),
                    '',
                    '',
                ]
            }
            
            # Aggiungi breakdown per sport (workouts)
            sport_counts_workouts = {}
            for w in workouts:
                sport_type = w.get('sportType', {})
                sport = sport_type.get('sportTypeKey', 'unknown') if isinstance(sport_type, dict) else 'unknown'
                sport_counts_workouts[sport] = sport_counts_workouts.get(sport, 0) + 1
            
            for sport, count in sorted(sport_counts_workouts.items()):
                summary_data['Metric'].append(f"  - {sport}")
                summary_data['Value'].append(count)
            
            summary_data['Metric'].append('')
            summary_data['Value'].append('')
            summary_data['Metric'].append('Activities by Sport')
            summary_data['Value'].append('')
            
            # Aggiungi breakdown per sport (activities)
            sport_counts_activities = {}
            total_distance = 0
            total_duration = 0
            
            for a in activities:
                activity_type = a.get('activityType', {})
                sport = activity_type.get('typeKey', 'unknown') if isinstance(activity_type, dict) else 'unknown'
                sport_counts_activities[sport] = sport_counts_activities.get(sport, 0) + 1
                
                total_distance += a.get('distance', 0)
                total_duration += a.get('duration', 0)
            
            for sport, count in sorted(sport_counts_activities.items()):
                summary_data['Metric'].append(f"  - {sport}")
                summary_data['Value'].append(count)
            
            # Aggiungi statistiche totali activities
            if activities:
                summary_data['Metric'].extend(['', 'Total Distance (km)', 'Total Duration (hours)', 'Average Distance (km)'])
                summary_data['Value'].extend([
                    '',
                    f"{total_distance/1000:.2f}",
                    f"{total_duration/3600:.2f}",
                    f"{(total_distance/1000)/len(activities):.2f}"
                ])
            
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            
            # ========== SHEET 4: PARAMETERS ==========
            # Zone definitions di esempio per l'AI
            parameters_data = {
                'Key': ['Z1', 'Z2', 'Z3', 'Z4', 'Z5', '', 'FTP_Z1', 'FTP_Z2', 'FTP_Z3', 'FTP_Z4', 'FTP_Z5'],
                'Metric': ['pace', 'pace', 'pace', 'pace', 'pace', '', 'power', 'power', 'power', 'power', 'power'],
                'Expression': ['6:30-7:00', '5:50-6:20', '5:20-5:45', '4:50-5:15', '4:20-4:45', '', '0.55', '0.56-0.75', '0.76-0.90', '0.91-1.05', '1.06-1.20'],
                'Notes': ['Recovery', 'Easy', 'Tempo', 'Threshold', 'VO2max', '', 'Active Recovery', 'Endurance', 'Tempo', 'Lactate Threshold', 'VO2max']
            }
            
            df_parameters = pd.DataFrame(parameters_data)
            df_parameters.to_excel(writer, sheet_name='Parameters', index=False)
    
    def _activity_to_dsl(self, activity, sport):
        """Converte un'attività completata in DSL equivalente approssimativo."""
        try:
            duration = activity.get('duration') or 0  # secondi
            distance = activity.get('distance') or 0  # metri
            
            if not duration or duration == 0:
                return ""
            
            # Calcola pace/velocità
            if distance and distance > 0:
                pace_sec_per_km = (duration / (distance / 1000))
                pace_min = int(pace_sec_per_km / 60)
                pace_sec = int(pace_sec_per_km % 60)
                pace_str = f"{pace_min}:{pace_sec:02d}"
            else:
                pace_str = "open"
            
            # Durata totale in minuti
            duration_min = int(duration / 60)
            
            if duration_min < 15:
                # Workout corto - un solo blocco
                step_type = "run" if sport == "running" else "bike" if sport == "cycling" else "swim"
                return f"{step_type} {duration_min}min {pace_str}"
            
            # Struttura con warmup/main/cooldown
            warmup_min = max(5, int(duration_min * 0.10))
            cooldown_min = max(5, int(duration_min * 0.10))
            main_min = duration_min - warmup_min - cooldown_min
            
            if main_min <= 0:
                step_type = "run" if sport == "running" else "bike" if sport == "cycling" else "swim"
                return f"{step_type} {duration_min}min {pace_str}"
            
            step_type = "run" if sport == "running" else "bike" if sport == "cycling" else "swim"
            
            # Genera DSL strutturato
            dsl = f"warmup {warmup_min}min Z2; {step_type} {main_min}min {pace_str}; cooldown {cooldown_min}min Z1"
            return dsl
            
        except Exception as e:
            return ""
    
    
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