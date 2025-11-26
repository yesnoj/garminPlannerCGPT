"""
Garmin Connect API Client - VERSIONE DEFINITIVA
Usa garth.client.request() per avere automaticamente CSRF token e cookie di sessione.
"""
import garth
from typing import Dict, Any
import json

# Endpoint costanti
_WORKOUT_SERVICE_ENDPOINT = "/workout-service"


def save_workout(workout: Dict[str, Any]) -> Dict[str, Any]:
    """
    Crea un nuovo workout su Garmin Connect.
    
    Args:
        workout: Dizionario JSON del workout
        
    Returns:
        Response JSON con workoutId
    """
    url = f"{_WORKOUT_SERVICE_ENDPOINT}/workout"
    response = garth.connectapi(url, method="POST", json=workout)
    
    if hasattr(response, 'json'):
        return response.json()
    return response


def schedule_workout(workout_id: str, date: str) -> Dict[str, Any]:
    """
    Pianifica un workout su una data specifica.
    Stampa a video la risposta completa di Garmin così possiamo capire
    quali campi usare poi per cancellare la pianificazione.
    """
    url = f"{_WORKOUT_SERVICE_ENDPOINT}/schedule/{workout_id}"
    json_data = {"date": date}
    response = garth.connectapi(url, method="POST", json=json_data)

    print("\n=== GARMIN RESPONSE (schedule_workout) ===")
    try:
        if hasattr(response, "json"):
            data = response.json()
            try:
                print(json.dumps(data, indent=2, ensure_ascii=False))
            except Exception:
                # se ci sono problemi con caratteri strani, stampiamo grezzo
                print(data)
            print("=== END SCHEDULE RESPONSE ===\n")
            return data
        else:
            # in alcune versioni garth può restituire già un dizionario
            print(response)
            print("=== END SCHEDULE RESPONSE ===\n")
            return response
    except Exception as e:
        print(f"Errore nel leggere la risposta JSON: {e}")
        # se è un oggetto tipo requests.Response, proviamo a stampare il testo
        if hasattr(response, "text"):
            print("Response.text:", response.text)
        else:
            print("Response object:", response)
        print("=== END SCHEDULE RESPONSE (ERROR) ===\n")
        raise



def remove_workout_schedule(workout_id: str, date: str) -> Any:
    """
    Rimuove la pianificazione di un workout da Garmin Connect.

    Usa l'endpoint:
    DELETE https://connectapi.garmin.com/workout-service/schedule/{workoutScheduleId}

    Args:
        workout_id: ID della pianificazione (workoutScheduleId)
        date: Data in formato YYYY-MM-DD (solo informativa)

    Returns:
        Response object oppure dict, a seconda di garth.connectapi

    Raises:
        RuntimeError: Se la richiesta fallisce
    """
    # IMPORTANTE: endpoint relativo, come nelle altre funzioni
    url = f"{_WORKOUT_SERVICE_ENDPOINT}/schedule/{workout_id}"

    print(f"🔄 Rimozione pianificazione workout {workout_id} del {date}...")
    print(f"   URL (relative): {url}")
    print("   Host: connectapi.garmin.com (via garth.connectapi)")
    print("   Method: DELETE")

    try:
        # Usiamo garth.connectapi come per la POST di schedule_workout
        response = garth.connectapi(url, method="DELETE")

        # Alcune versioni di garth restituiscono un oggetto Response, altre un dict
        status_code = getattr(response, "status_code", None)

        if status_code is None:
            # Nessun status_code: se non è esploso, consideriamo la cosa un successo
            print("✅ Workout dis-pianificato (nessun status_code, nessuna eccezione)")
            return response

        if status_code in (200, 204):
            print(f"✅ SUCCESS! Workout dis-pianificato (Status {status_code})")
            return response

        # Altri status code -> fallo esplodere per finire nel blocco except
        print(f"⚠️  Status {status_code}")
        if hasattr(response, "raise_for_status"):
            response.raise_for_status()
        else:
            raise RuntimeError(f"Status code inatteso: {status_code}")

    except Exception as e:
        error_msg = str(e)

        if "403" in error_msg:
            raise RuntimeError(
                "Errore 403 Forbidden - Possibili cause:\n"
                "1. CSRF token non valido o scaduto\n"
                "2. Permessi insufficienti\n"
                "3. Sessione scaduta\n\n"
                "Soluzione: rifare login con garth.login()\n\n"
                f"Workout ID: {workout_id}, Date: {date}\n"
                f"Dettagli: {e}"
            )
        elif "401" in error_msg:
            raise RuntimeError(
                "Errore 401 Unauthorized - Token OAuth2 non accettato per questa "
                "chiamata.\n\n"
                "Se l'errore persiste dopo aver rifatto il login, è probabile che "
                "Garmin abbia cambiato le regole di accesso per questo endpoint.\n\n"
                f"Workout ID: {workout_id}, Date: {date}\n"
                f"Dettagli: {e}"
            )
        elif "404" in error_msg:
            raise RuntimeError(
                "Errore 404 Not Found:\n"
                f"1. WorkoutScheduleId non esiste: {workout_id}\n"
                f"2. Workout non è (più) pianificato per {date}\n"
                "3. Endpoint non corretto\n\n"
                f"Dettagli: {e}"
            )
        else:
            raise RuntimeError(
                "Errore nella rimozione pianificazione:\n"
                f"Workout ID: {workout_id}, Date: {date}\n\n"
                f"Dettagli: {e}"
            )



# Mantieni il nome originale per compatibilità
def remove_workout(workout_id: str, date: str) -> Any:
    """
    Alias per remove_workout_schedule per compatibilità con il codice esistente.
    
    Args:
        workout_id: ID del workout
        date: Data in formato YYYY-MM-DD
        
    Returns:
        Response object
    """
    return remove_workout_schedule(workout_id, date)