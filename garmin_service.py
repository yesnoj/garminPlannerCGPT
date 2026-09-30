"""
Servizio per gestire l'interazione con Garmin Connect.
Usa garth direttamente senza wrapper custom.

I token di sessione sono salvati FUORI dal progetto, in ~/.garminplanner/tokens
(permessi 700/600), cosi' non finiscono per sbaglio su git.
Si puo' cambiare cartella con la variabile d'ambiente GARMINPLANNER_TOKEN_DIR.
"""
import json
import os
import shutil
import stat
from pathlib import Path
from typing import Dict, Any, Optional

import garth

import garmin_client


TOKEN_FILES = ("oauth1_token.json", "oauth2_token.json")

TOKEN_STORE_DIR = Path(
    os.environ.get("GARMINPLANNER_TOKEN_DIR", Path.home() / ".garminplanner" / "tokens")
).expanduser()

# Vecchia posizione (dentro la cartella del progetto): usata solo per la migrazione.
LEGACY_TOKEN_DIRS = (
    Path(__file__).resolve().parent / "garminconnect",
    Path.cwd() / "garminconnect",
)

_DEBUG = os.environ.get("GARMINPLANNER_DEBUG", "") not in ("", "0")


def _debug(*args):
    if _DEBUG:
        print(*args)


def _secure_token_dir() -> Path:
    """Crea la cartella token con permessi 700."""
    TOKEN_STORE_DIR.mkdir(parents=True, exist_ok=True)
    try:
        os.chmod(TOKEN_STORE_DIR, stat.S_IRWXU)
        os.chmod(TOKEN_STORE_DIR.parent, stat.S_IRWXU)
    except OSError:
        pass  # es. Windows: i permessi POSIX non sono supportati
    return TOKEN_STORE_DIR


def _save_tokens() -> None:
    """Salva i token garth e li rende leggibili solo dall'utente (600)."""
    token_dir = _secure_token_dir()
    garth.save(str(token_dir))
    for name in TOKEN_FILES:
        f = token_dir / name
        if f.exists():
            try:
                os.chmod(f, stat.S_IRUSR | stat.S_IWUSR)
            except OSError:
                pass


def _migrate_legacy_tokens() -> Optional[Path]:
    """
    Se esiste una sessione nella vecchia cartella ./garminconnect e non in quella nuova,
    la copia nella nuova posizione. Ritorna la cartella legacy trovata (o None).
    """
    if all((TOKEN_STORE_DIR / n).exists() for n in TOKEN_FILES):
        return None
    for legacy in LEGACY_TOKEN_DIRS:
        if all((legacy / n).exists() for n in TOKEN_FILES):
            token_dir = _secure_token_dir()
            for n in TOKEN_FILES:
                shutil.copy2(legacy / n, token_dir / n)
                try:
                    os.chmod(token_dir / n, stat.S_IRUSR | stat.S_IWUSR)
                except OSError:
                    pass
            return legacy
    return None


class GarminService:
    """Wrapper per gestire autenticazione e operazioni Garmin."""

    def __init__(self):
        self.authenticated = False
        self.migrated_from: Optional[Path] = None

    @staticmethod
    def has_saved_session() -> bool:
        """Controlla se esiste una sessione salvata (entrambi i file token)."""
        if all((TOKEN_STORE_DIR / n).exists() for n in TOKEN_FILES):
            return True
        return any(all((d / n).exists() for n in TOKEN_FILES) for d in LEGACY_TOKEN_DIRS)

    def login_with_credentials(self, email: str, password: str, parent=None) -> None:
        """
        Effettua login con email e password.
        Salva la sessione per riutilizzo futuro.
        Supporta MFA se richiesto (dialog agganciato alla finestra `parent`, se fornita).
        """
        def mfa_prompt():
            """Chiede il codice MFA all'utente via GUI."""
            from tkinter import simpledialog, Tk
            owner = parent
            temp_root = None
            if owner is None:
                temp_root = Tk()
                temp_root.withdraw()
                owner = temp_root
            try:
                code = simpledialog.askstring(
                    "Codice MFA richiesto",
                    "Garmin ha inviato un codice di verifica via email.\nInserisci il codice:",
                    parent=owner,
                )
            finally:
                if temp_root is not None:
                    temp_root.destroy()
            if not code:
                raise RuntimeError("Codice MFA non inserito.")
            return code.strip()

        try:
            garth.login(email, password, prompt_mfa=mfa_prompt)
        except Exception as e:
            self.authenticated = False
            raise RuntimeError(
                "Login a Garmin non riuscito. Controlla email e password "
                f"(dettaglio: {type(e).__name__})."
            ) from e
        _save_tokens()
        self.authenticated = True

    def login_with_saved_session(self) -> None:
        """
        Riprende la sessione salvata.
        Lancia eccezione se la sessione non esiste o e' invalida.
        """
        self.migrated_from = _migrate_legacy_tokens()
        if not all((TOKEN_STORE_DIR / n).exists() for n in TOKEN_FILES):
            raise RuntimeError("Nessuna sessione salvata trovata: effettua il login con email e password.")

        try:
            garth.resume(str(TOKEN_STORE_DIR))
            # Verifica che la sessione sia ancora valida (e fa scattare il refresh del token)
            garth.client.username
        except Exception as e:
            self.authenticated = False
            raise RuntimeError(
                "Sessione salvata scaduta o non valida: effettua di nuovo il login con email e password."
            ) from e
        _save_tokens()  # salva l'eventuale token aggiornato
        self.authenticated = True

    def ensure_authenticated(self) -> None:
        """Verifica che sia autenticato."""
        if not self.authenticated:
            raise RuntimeError("Non sei connesso a Garmin.")

    # ========== WORKOUT OPERATIONS ==========

    def save_workout(self, workout_data: Dict[str, Any]) -> Dict[str, Any]:
        """Crea un nuovo workout su Garmin Connect. Ritorna la risposta JSON con workoutId."""
        self.ensure_authenticated()
        _debug("\n=== WORKOUT JSON INVIATO ===\n" + json.dumps(workout_data, indent=2, ensure_ascii=False))
        resp = garmin_client.save_workout(workout_data)
        _debug("\n=== GARMIN RESPONSE ===\n", resp)
        return resp

    def schedule_workout(self, workout_id: str, date_str: str) -> Dict[str, Any]:
        """Pianifica un workout su una data specifica (YYYY-MM-DD)."""
        self.ensure_authenticated()
        return garmin_client.schedule_workout(workout_id, date_str)

    def unschedule_workout(self, schedule_id: str, date_str: str) -> None:
        """Rimuove la pianificazione di un workout (workoutScheduleId)."""
        self.ensure_authenticated()
        garmin_client.remove_workout_schedule(schedule_id, date_str)

    def delete_workout(self, workout_id: str) -> None:
        """Cancella definitivamente un workout dalla libreria Garmin."""
        self.ensure_authenticated()
        garmin_client.delete_workout_definition(workout_id)

    # Proprieta' per compatibilita' con il codice esistente
    @property
    def client(self):
        """Proprieta' per compatibilita' - ritorna self se autenticato."""
        return self if self.authenticated else None
