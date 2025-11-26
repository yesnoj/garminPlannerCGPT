"""
Servizio per gestire l'interazione con Garmin Connect.
Usa garth direttamente senza wrapper custom.
"""
import garth
from pathlib import Path
from typing import Dict, Any
import garmin_client


TOKEN_STORE_DIR = "./garminconnect"


class GarminService:
    """Wrapper per gestire autenticazione e operazioni Garmin."""
    
    def __init__(self):
        self.authenticated = False
    
    @staticmethod
    def has_saved_session() -> bool:
        """Controlla se esiste una sessione salvata."""
        token_dir = Path(TOKEN_STORE_DIR)
        return token_dir.exists() and any(token_dir.iterdir())
    
    def login_with_credentials(self, email: str, password: str) -> None:
        """
        Effettua login con email e password.
        Salva la sessione per riutilizzo futuro.
        Supporta MFA se richiesto.
        """
        # Crea directory se non esiste
        Path(TOKEN_STORE_DIR).mkdir(exist_ok=True)
        
        # Funzione per richiedere MFA code
        def mfa_prompt():
            """Chiede il codice MFA all'utente via GUI."""
            from tkinter import simpledialog, Tk
            root = Tk()
            root.withdraw()  # Nascondi finestra principale
            code = simpledialog.askstring(
                "Codice MFA richiesto",
                "Garmin ha inviato un codice di verifica via email.\nInserisci il codice:",
                parent=root
            )
            root.destroy()
            return code
        
        # Login con supporto MFA
        try:
            garth.login(email, password, prompt_mfa=mfa_prompt)
            garth.save(TOKEN_STORE_DIR)
            self.authenticated = True
        except Exception as e:
            raise RuntimeError(f"Errore durante il login: {e}")
    
    def login_with_saved_session(self) -> None:
        """
        Riprende la sessione salvata.
        Lancia eccezione se la sessione non esiste o è invalida.
        """
        if not self.has_saved_session():
            raise RuntimeError("Nessuna sessione salvata trovata.")
        
        garth.resume(TOKEN_STORE_DIR)
        self.authenticated = True
    
    def ensure_authenticated(self) -> None:
        """Verifica che sia autenticato."""
        if not self.authenticated:
            raise RuntimeError("Non sei connesso a Garmin.")
    
    # ========== WORKOUT OPERATIONS ==========
    
    def save_workout(self, workout_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Crea un nuovo workout su Garmin Connect.
        
        Args:
            workout_data: Dizionario JSON del workout
            
        Returns:
            Response JSON con workoutId
        """
        self.ensure_authenticated()
        
        # DEBUG: stampa il JSON inviato
        import json
        print("\n=== WORKOUT JSON INVIATO ===")
        print(json.dumps(workout_data, indent=2, ensure_ascii=False))
        print("=== END ===\n")
        
        resp = garmin_client.save_workout(workout_data)
        
        print("\n=== GARMIN RESPONSE ===")
        print(resp)
        print("=== END ===\n")
        
        return resp
    
    def schedule_workout(self, workout_id: str, date_str: str) -> Dict[str, Any]:
        """
        Pianifica un workout su una data specifica.
        
        Args:
            workout_id: ID del workout
            date_str: Data in formato YYYY-MM-DD
            
        Returns:
            Response JSON
        """
        self.ensure_authenticated()
        return garmin_client.schedule_workout(workout_id, date_str)
    
    def unschedule_workout(self, schedule_id: str, date_str: str) -> None:
        """
        Rimuove la pianificazione di un workout.

        Args:
            schedule_id: ID della pianificazione (workoutScheduleId)
            date_str: Data in formato YYYY-MM-DD (solo per messaggi)
        """
        self.ensure_authenticated()
        garmin_client.remove_workout_schedule(schedule_id, date_str)

    
    # Proprietà per compatibilità con il codice esistente
    @property
    def client(self):
        """Proprietà per compatibilità - ritorna self se autenticato."""
        return self if self.authenticated else None