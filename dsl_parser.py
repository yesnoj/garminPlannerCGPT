import re
from typing import List, Tuple, Optional, Dict, Any, Union

import pandas as pd


# ------------------------------------------------------------
# 1) Espansione dei repeat mantenendo la struttura
# ------------------------------------------------------------

def _count_leading_spaces(s: str) -> int:
    """Conta gli spazi iniziali in una stringa."""
    return len(s) - len(s.lstrip(" "))


def expand_repeat_lines(steps_text: str) -> List[Union[str, Dict[str, Any]]]:
    """
    Prende il testo degli steps e ritorna una lista mista:
    - Stringhe per step normali
    - Dict {"repeat": N, "steps": [...]} per blocchi repeat
    """
    raw_lines = steps_text.splitlines()
    while raw_lines and not raw_lines[-1].strip():
        raw_lines.pop()
    
    expanded, _ = _expand_block_structured(raw_lines, 0, 0)
    return expanded


def _expand_block_structured(lines: List[str], start: int, indent: int) -> Tuple[List[Any], int]:
    """Espande ricorsivamente mantenendo la struttura repeat."""
    result: List[Any] = []
    i = start

    while i < len(lines):
        raw = lines[i]
        if not raw.strip():
            i += 1
            continue

        current_indent = _count_leading_spaces(raw)
        if current_indent < indent:
            break
        if current_indent > indent:
            i += 1
            continue

        content = raw[indent:].rstrip()

        # repeat N:
        m = re.match(r"^repeat\s+(\d+)\s*:\s*$", content, flags=re.IGNORECASE)
        if m:
            reps = int(m.group(1))
            child_steps, new_i = _expand_block_structured(lines, i + 1, indent + 2)
            result.append({
                "repeat": reps,
                "steps": child_steps
            })
            i = new_i
            continue

        # Riga normale
        result.append(content)
        i += 1

    return result, i


# ------------------------------------------------------------
# 2) Parse durata
# ------------------------------------------------------------

def parse_duration_part(base_str: str) -> Tuple[str, Optional[float], Optional[str]]:
    """
    Parsea la parte prima di '@' per ricavare:
      - tipo condizione ('time' o 'distance' o 'lap.button')
      - valore numerico (secondi o metri)
      - unitÃ  preferita
    """
    s = base_str.strip().lower()

    # lap-button
    if "lap-button" in s or "lap button" in s:
        return "lap.button", None, None

    # ORDINE IMPORTANTE: distanze PRIMA di tempi per evitare confusione tra "m" (metri) e "m" (minuti abbreviato)
    
    # Metri: "1000m", "200m", "1500 m" - DEVE essere prima di "min"
    # IMPORTANTE: negative lookahead (?!i) per NON catturare "min" come metri
    m = re.match(r"^(\d+(?:\.\d+)?)\s*m(?!i)(?:\s|$)", s)
    if m:
        meters = float(m.group(1))
        return "distance", meters, "meter"
    
    # Chilometri: "5km", "5 km", "1.5km"
    m = re.match(r"^(\d+(?:\.\d+)?)\s*km(?:\s|$)", s)
    if m:
        km = float(m.group(1))
        meters = km * 1000.0
        return "distance", meters, "kilometer"

    # ORA i tempi
    # Minuti: "10min", "10 min", "40'", "35 '"
    m = re.match(r"^(\d+)\s*(?:min|'|â€²)(?:\s|$)", s)
    if m:
        minutes = int(m.group(1))
        return "time", float(minutes * 60), None

    # Secondi: "30sec", "30s", "30 s"
    m = re.match(r"^(\d+)\s*(?:sec|s)(?:\s|$)", s)
    if m:
        seconds = int(m.group(1))
        return "time", float(seconds), None

    # fallback
    return "time", 60.0, None


# ------------------------------------------------------------
# 3) Mappa tipo di step
# ------------------------------------------------------------

STEP_TYPE_IDS = {
    'warmup': 1,
    'cooldown': 2,
    'interval': 3,
    'recovery': 4,
    'rest': 5,
    'repeat': 6,
}

CONDITION_TYPE_IDS = {
    'lap.button': 1,
    'time': 2,
    'distance': 3,
}

SPORT_TYPE_IDS = {
    'running': 1,
    'cycling': 2,
    'swimming': 4,
}

TARGET_TYPE_IDS = {
    'no.target': 1,
    'power.zone': 2,
    'cadence': 3,
    'heart.rate.zone': 4,
    'speed.zone': 5,
    'pace.zone': 6,
}


def map_step_type(step_key: str) -> Dict[str, Any]:
    """Mappa la parola chiave DSL al stepType Garmin."""
    k = step_key.strip().lower()
    
    if k == "warmup":
        type_key = "warmup"
    elif k == "cooldown":
        type_key = "cooldown"
    elif k in ("recovery", "rest"):
        type_key = "recovery"
    else:
        type_key = "interval"
    
    return {
        "stepTypeId": STEP_TYPE_IDS[type_key],
        "stepTypeKey": type_key
    }


# ------------------------------------------------------------
# 4) Funzioni helper per Parameters
# ------------------------------------------------------------

def get_pace_tolerance(df_parameters: Optional[pd.DataFrame] = None) -> float:
    """Legge pace_tolerance da Parameters (default: 10 sec)."""
    if df_parameters is not None:
        tolerance_row = df_parameters[df_parameters['Key'] == 'pace_tolerance']
        if not tolerance_row.empty:
            expr = str(tolerance_row.iloc[0]['Expression']).strip()
            try:
                return float(expr)
            except (ValueError, TypeError):
                pass
    return 10.0


def get_hr_tolerance(df_parameters: Optional[pd.DataFrame] = None) -> float:
    """Legge hr_tolerance da Parameters (default: 5 bpm)."""
    if df_parameters is not None:
        tolerance_row = df_parameters[df_parameters['Key'] == 'hr_tolerance']
        if not tolerance_row.empty:
            expr = str(tolerance_row.iloc[0]['Expression']).strip()
            try:
                return float(expr)
            except (ValueError, TypeError):
                pass
    return 5.0


def get_hr_max(df_parameters: Optional[pd.DataFrame] = None) -> Optional[float]:
    """Legge HR_max da Parameters."""
    if df_parameters is not None:
        hr_max_row = df_parameters[df_parameters['Key'] == 'HR_max']
        if not hr_max_row.empty:
            expr = str(hr_max_row.iloc[0]['Expression']).strip()
            try:
                return float(expr)
            except (ValueError, TypeError):
                pass
    return None


def get_parameter_value(key: str, df_parameters: Optional[pd.DataFrame] = None) -> Optional[str]:
    """Cerca un parametro nel DataFrame Parameters."""
    if df_parameters is None:
        return None
    param_row = df_parameters[df_parameters['Key'] == key]
    if not param_row.empty:
        return str(param_row.iloc[0]['Expression']).strip()
    return None


def parse_pace_expression(expr: str) -> Tuple[float, Optional[float]]:
    """
    Parsea espressione di ritmo da Parameters.
    - "5:00" -> (300, None)
    - "5:00-5:30" -> (300, 330)
    - "300" -> (300, None)
    """
    expr = expr.strip()
    if '-' in expr:
        parts = expr.split('-')
        if len(parts) == 2:
            min_val = parse_single_pace(parts[0].strip())
            max_val = parse_single_pace(parts[1].strip())
            return (min_val, max_val)
    return (parse_single_pace(expr), None)


def parse_single_pace(s: str) -> float:
    """Parsea singolo valore di ritmo: '5:00' -> 300 sec."""
    if ':' in s:
        parts = s.split(':')
        if len(parts) == 2:
            return float(int(parts[0]) * 60 + int(parts[1]))
    return float(s)


def pace_seconds_to_mps(pace_seconds: float) -> float:
    """Converte sec/km a m/s: 300 sec/km -> 3.33 m/s."""
    return 1000.0 / pace_seconds if pace_seconds > 0 else 1.0


def parse_hr_expression(expr: str, df_parameters: Optional[pd.DataFrame] = None) -> Tuple[float, Optional[float]]:
    """
    Parsea espressione HR: '150', '150-165', '80%', '70-85%'
    Se sono percentuali, usa HR_max da Parameters.
    """
    expr = expr.strip()
    
    # Percentuale singola: "80%"
    m = re.match(r"^(\d+(?:\.\d+)?)%$", expr)
    if m:
        pct = float(m.group(1))
        hr_max = get_hr_max(df_parameters)
        if hr_max:
            hr_value = hr_max * (pct / 100.0)
            # Singolo valore: applica hr_tolerance
            tolerance = get_hr_tolerance(df_parameters)
            return (hr_value - tolerance, hr_value + tolerance)
        else:
            # Fallback se HR_max non definito
            return (150.0, 160.0)
    
    # Range percentuale: "70-85%"
    m = re.match(r"^(\d+(?:\.\d+)?)-(\d+(?:\.\d+)?)%$", expr)
    if m:
        pct_min = float(m.group(1))
        pct_max = float(m.group(2))
        hr_max = get_hr_max(df_parameters)
        if hr_max:
            hr_min = hr_max * (pct_min / 100.0)
            hr_max_val = hr_max * (pct_max / 100.0)
            return (hr_min, hr_max_val)
        else:
            # Fallback
            return (140.0, 170.0)
    
    # Range assoluto: "150-165"
    if '-' in expr:
        parts = expr.split('-')
        if len(parts) == 2:
            return (float(parts[0]), float(parts[1]))
    
    # Singolo valore assoluto: "150"
    hr_value = float(expr)
    tolerance = get_hr_tolerance(df_parameters)
    return (hr_value - tolerance, hr_value + tolerance)


# ------------------------------------------------------------
# 5) Parse target
# ------------------------------------------------------------

def parse_target(target_str: str, df_parameters: Optional[pd.DataFrame] = None) -> Dict[str, Any]:
    """
    Parsea il target dopo '@'.
    Supporta: ritmo (5:00, Z1-Z5, custom), HR (HR1-HR5, 140-160bpm), potenza, cadenza.
    """
    s = target_str.strip().lower()
    
    if not s or s == "open":
        return {
            "targetType": {
                "workoutTargetTypeId": TARGET_TYPE_IDS['no.target'],
                "workoutTargetTypeKey": "no.target"
            },
            "targetValueOne": None,
            "targetValueTwo": None,
            "zoneNumber": None
        }
    
    # Zone di ritmo standard: Z1, Z2, Z3, Z4, Z5
    m = re.match(r"^z(\d)$", s)
    if m:
        zone_num = m.group(1)
        zone_key = f"Z{zone_num}"
        zone_expr = get_parameter_value(zone_key, df_parameters)
        
        if zone_expr:
            min_sec, max_sec = parse_pace_expression(zone_expr)
            
            if max_sec is None:
                # Singolo valore: applica pace_tolerance
                tolerance_sec = get_pace_tolerance(df_parameters)
                slower_sec = min_sec + tolerance_sec
                faster_sec = min_sec - tolerance_sec
                slower_mps = pace_seconds_to_mps(slower_sec)
                faster_mps = pace_seconds_to_mps(faster_sec)
            else:
                # Range esplicito
                slower_mps = pace_seconds_to_mps(max(min_sec, max_sec))
                faster_mps = pace_seconds_to_mps(min(min_sec, max_sec))
            
            return {
                "targetType": {
                    "workoutTargetTypeId": TARGET_TYPE_IDS['pace.zone'],
                    "workoutTargetTypeKey": "pace.zone"
                },
                "targetValueOne": slower_mps,
                "targetValueTwo": faster_mps,
                "zoneNumber": None
            }
        else:
            # Usa zone Garmin predefinite
            zone = int(zone_num)
            return {
                "targetType": {
                    "workoutTargetTypeId": TARGET_TYPE_IDS['pace.zone'],
                    "workoutTargetTypeKey": "pace.zone"
                },
                "targetValueOne": None,
                "targetValueTwo": None,
                "zoneNumber": zone
            }
    
    # Zone HR: HR1, HR2, HR_Z1, etc.
    m = re.match(r"^hr[_]?z?(\d)$", s)
    if m:
        zone_num = m.group(1)
        for zone_key in [f"HR{zone_num}", f"HR_Z{zone_num}", f"HRZ{zone_num}"]:
            zone_expr = get_parameter_value(zone_key, df_parameters)
            if zone_expr:
                min_hr, max_hr = parse_hr_expression(zone_expr, df_parameters)
                return {
                    "targetType": {
                        "workoutTargetTypeId": TARGET_TYPE_IDS['heart.rate.zone'],
                        "workoutTargetTypeKey": "heart.rate.zone"
                    },
                    "targetValueOne": min_hr,
                    "targetValueTwo": max_hr,
                    "zoneNumber": None
                }
        
        zone = int(zone_num)
        return {
            "targetType": {
                "workoutTargetTypeId": TARGET_TYPE_IDS['heart.rate.zone'],
                "workoutTargetTypeKey": "heart.rate.zone"
            },
            "targetValueOne": None,
            "targetValueTwo": None,
            "zoneNumber": zone
        }
    
    # Ritmo assoluto: "5:00" o "5:00-5:30"
    m = re.match(r"^(\d+):(\d+)(?:-(\d+):(\d+))?$", s)
    if m:
        min1 = int(m.group(1))
        sec1 = int(m.group(2))
        total_sec = min1 * 60 + sec1
        
        if m.group(3) and m.group(4):
            # Range
            min2 = int(m.group(3))
            sec2 = int(m.group(4))
            total_sec2 = min2 * 60 + sec2
            pace1_mps = pace_seconds_to_mps(total_sec)
            pace2_mps = pace_seconds_to_mps(total_sec2)
            
            return {
                "targetType": {
                    "workoutTargetTypeId": TARGET_TYPE_IDS['pace.zone'],
                    "workoutTargetTypeKey": "pace.zone"
                },
                "targetValueOne": min(pace1_mps, pace2_mps),
                "targetValueTwo": max(pace1_mps, pace2_mps),
                "zoneNumber": None
            }
        else:
            # Singolo valore: applica pace_tolerance
            tolerance_sec = get_pace_tolerance(df_parameters)
            slower_sec = total_sec + tolerance_sec
            faster_sec = total_sec - tolerance_sec
            slower_mps = pace_seconds_to_mps(slower_sec)
            faster_mps = pace_seconds_to_mps(faster_sec)
            
            return {
                "targetType": {
                    "workoutTargetTypeId": TARGET_TYPE_IDS['pace.zone'],
                    "workoutTargetTypeKey": "pace.zone"
                },
                "targetValueOne": slower_mps,
                "targetValueTwo": faster_mps,
                "zoneNumber": None
            }
    
    # NUOVO: Parametri custom di ritmo (recovery, marathon, threshold, easy_range, ecc.)
    # Cerca qualsiasi parametro in Parameters che non sia giÃ  stato gestito
    custom_pace = get_parameter_value(s, df_parameters)
    if custom_pace:
        # Trovato un parametro custom, prova a parsarlo come ritmo
        try:
            min_sec, max_sec = parse_pace_expression(custom_pace)
            
            if max_sec is None:
                # Singolo valore: applica pace_tolerance
                tolerance_sec = get_pace_tolerance(df_parameters)
                slower_sec = min_sec + tolerance_sec
                faster_sec = min_sec - tolerance_sec
                slower_mps = pace_seconds_to_mps(slower_sec)
                faster_mps = pace_seconds_to_mps(faster_sec)
            else:
                # Range esplicito
                slower_mps = pace_seconds_to_mps(max(min_sec, max_sec))
                faster_mps = pace_seconds_to_mps(min(min_sec, max_sec))
            
            return {
                "targetType": {
                    "workoutTargetTypeId": TARGET_TYPE_IDS['pace.zone'],
                    "workoutTargetTypeKey": "pace.zone"
                },
                "targetValueOne": slower_mps,
                "targetValueTwo": faster_mps,
                "zoneNumber": None
            }
        except:
            pass  # Se non riesce a parsare, continua con gli altri check
    
    # Potenza: "200W" o "200-250W"
    m = re.match(r"^(\d+)(?:-(\d+))?w$", s)
    if m:
        power1 = float(m.group(1))
        power2 = float(m.group(2)) if m.group(2) else power1 * 1.05
        return {
            "targetType": {
                "workoutTargetTypeId": TARGET_TYPE_IDS['power.zone'],
                "workoutTargetTypeKey": "power.zone"
            },
            "targetValueOne": power1,
            "targetValueTwo": power2,
            "zoneNumber": None
        }
    
    # Cadenza: "85rpm" o "85-90rpm"
    m = re.match(r"^(\d+)(?:-(\d+))?rpm$", s)
    if m:
        cadence1 = float(m.group(1))
        cadence2 = float(m.group(2)) if m.group(2) else cadence1 + 5
        return {
            "targetType": {
                "workoutTargetTypeId": TARGET_TYPE_IDS['cadence'],
                "workoutTargetTypeKey": "cadence"
            },
            "targetValueOne": cadence1,
            "targetValueTwo": cadence2,
            "zoneNumber": None
        }
    
    # HR range: "140-160bpm" o "140-160"
    m = re.match(r"^(\d+)(?:-(\d+))?(?:bpm)?$", s)
    if m and int(m.group(1)) > 20:
        hr1 = float(m.group(1))
        hr2 = float(m.group(2)) if m.group(2) else hr1 + 10
        return {
            "targetType": {
                "workoutTargetTypeId": TARGET_TYPE_IDS['heart.rate.zone'],
                "workoutTargetTypeKey": "heart.rate.zone"
            },
            "targetValueOne": hr1,
            "targetValueTwo": hr2,
            "zoneNumber": None
        }
    
    # Fallback
    return {
        "targetType": {
            "workoutTargetTypeId": TARGET_TYPE_IDS['no.target'],
            "workoutTargetTypeKey": "no.target"
        },
        "targetValueOne": None,
        "targetValueTwo": None,
        "zoneNumber": None
    }


# ------------------------------------------------------------
# 6) Parse singola riga di step
# ------------------------------------------------------------

def parse_step_line(
    line: str,
    order: int,
    df_parameters: Optional[pd.DataFrame] = None,
) -> Dict[str, Any]:
    """Parsea una riga di step e ritorna un ExecutableStepDTO."""
    main_part = line.split("--", 1)[0].rstrip()

    m = re.match(r"^(warmup|cooldown|interval|recovery|rest|step)\s*:\s*(.+)$",
                 main_part, flags=re.IGNORECASE)
    if not m:
        step_type_key = "interval"
        base_and_target = main_part
    else:
        step_type_key = m.group(1).strip().lower()
        base_and_target = m.group(2).strip()

    # Separa durata da target
    target_str = ""
    if "@" in base_and_target:
        base_str, target_str = base_and_target.split("@", 1)
        base_str = base_str.strip()
        target_str = target_str.strip()
    else:
        base_str = base_and_target.strip()

    # Parse durata/distanza
    condition_type_key, end_condition_value, preferred_unit_key = parse_duration_part(base_str)

    # Parse target
    target_info = parse_target(target_str, df_parameters)

    # Step base
    step: Dict[str, Any] = {
        "type": "ExecutableStepDTO",
        "stepId": None,
        "stepOrder": order,
        "childStepId": None,
        "description": line.strip(),
        "stepType": map_step_type(step_type_key),
        "endCondition": {
            "conditionTypeId": CONDITION_TYPE_IDS[condition_type_key],
            "conditionTypeKey": condition_type_key,
        },
        "targetType": target_info["targetType"],
        "targetValueOne": target_info["targetValueOne"],
        "targetValueTwo": target_info["targetValueTwo"],
        "zoneNumber": target_info["zoneNumber"],
    }

    # Aggiungi endConditionValue solo se non Ã¨ lap.button
    if end_condition_value is not None:
        step["endConditionValue"] = end_condition_value
    
    # Se distanza, aggiungi preferredEndConditionUnit
    if preferred_unit_key:
        unit_ids = {'meter': 1, 'kilometer': 2}
        step["preferredEndConditionUnit"] = {
            "unitId": unit_ids.get(preferred_unit_key, 1),
            "unitKey": preferred_unit_key,
            "factor": 100.0 if preferred_unit_key == "meter" else 100000.0
        }

    return step


# ------------------------------------------------------------
# 7) Sport type
# ------------------------------------------------------------

def map_sport_type(row: pd.Series) -> Dict[str, Any]:
    """Mappa Sport in sportType Garmin."""
    raw = str(row.get("Sport", "") or "").strip().lower()
    key = "running"
    if "bike" in raw or "bici" in raw or "cycling" in raw:
        key = "cycling"
    elif "swim" in raw or "nuoto" in raw:
        key = "swimming"
    return {
        "sportTypeId": SPORT_TYPE_IDS[key],
        "sportTypeKey": key
    }


# ------------------------------------------------------------
# 8) Costruzione workout steps con RepeatGroupDTO
# ------------------------------------------------------------

def build_workout_steps(
    structured_steps: List[Any],
    df_parameters: Optional[pd.DataFrame] = None,
    start_order: int = 1,
    start_child_id: int = 1
) -> List[Dict[str, Any]]:
    """
    Converte la lista strutturata in workout steps Garmin.
    Gestisce sia ExecutableStepDTO che RepeatGroupDTO.
    """
    steps = []
    order = start_order
    child_id = start_child_id
    
    for item in structured_steps:
        if isinstance(item, dict) and "repeat" in item:
            # RepeatGroupDTO
            repeat_count = item["repeat"]
            nested_steps = build_workout_steps(
                item["steps"],
                df_parameters,
                start_order=1,  # Reset ordine dentro il repeat
                start_child_id=child_id + 1
            )
            
            repeat_step = {
                "type": "RepeatGroupDTO",
                "stepId": None,
                "stepOrder": order,
                "numberOfIterations": repeat_count,
                "smartRepeat": False,
                "childStepId": child_id,
                "workoutSteps": nested_steps
            }
            
            steps.append(repeat_step)
            order += 1
            child_id += 1
            
        elif isinstance(item, str):
            # ExecutableStepDTO
            step = parse_step_line(item, order, df_parameters)
            steps.append(step)
            order += 1
    
    return steps


# ------------------------------------------------------------
# 9) Costruzione workout completo
# ------------------------------------------------------------

def build_garmin_workout_from_excel_row(
    row: pd.Series,
    steps_text: str,
    df_parameters: Optional[pd.DataFrame] = None,
    use_prefix: bool = False,
) -> Dict[str, Any]:
    """Costruisce il workout JSON Garmin da Excel."""

    # Nome base dal campo Description
    workout_name = str(row.get("Description", "") or "").strip()
    if not workout_name:
        workout_name = "Workout"

    # Prefisso opzionale W{Week}S{Session}
    if use_prefix:
        week = str(row.get("Week", "") or "").strip()
        session = str(row.get("Session", "") or "").strip()

        parts = []
        if week:
            parts.append(f"W{week}")
        if session:
            parts.append(f"S{session}")

        if parts:
            prefix = "".join(parts) + " - "
            workout_name = prefix + workout_name

    sport_type = map_sport_type(row)

    # Espande mantenendo struttura repeat
    structured_steps = expand_repeat_lines(steps_text)

    # Converti in workout steps Garmin
    steps = build_workout_steps(structured_steps, df_parameters)

    segment = {
        "segmentOrder": 1,
        "sportType": sport_type,
        "workoutSteps": steps,
    }

    workout: Dict[str, Any] = {
        "workoutId": None,
        "ownerId": None,
        "workoutName": workout_name,
        "description": workout_name,
        "sportType": sport_type,
        "workoutSegments": [segment],
    }

    return workout