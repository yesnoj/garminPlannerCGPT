"""
Parser del DSL dei workout (colonna "Steps" dell'Excel) -> JSON Garmin Connect.

Formato supportato (una riga per step, i blocchi repeat si indentano):

    warmup: 15min @ Z2
    repeat 5:
      interval: 1000m @ 4:50-5:00
      recovery: 2min @ Z1
    cooldown: 10min @ Z1

Novita' rispetto alla versione precedente:
- ogni errore viene segnalato (DSLError) con il numero di riga, invece di
  trasformarsi in silenzio in uno step di 60 secondi o senza target;
- durate decimali e in ore (1.5min, 90s, 1h, 1:30), apice tipografico (40′);
- indentazione libera sotto "repeat" (2 o 4 spazi, tab), purche' coerente;
- le chiavi del foglio Parameters sono cercate senza distinzione maiuscole/minuscole
  (prima Swim_Z*, Cadence_* o Easy_Range venivano ignorate);
- "rest:" produce uno step di tipo Riposo (prima veniva inviato come Recupero);
- validate_workout_rows() controlla un intero piano prima del caricamento.

Per input validi il JSON prodotto e' identico a quello della versione precedente
(eccetto il tipo degli step "rest").
"""
import math
import re
from typing import List, Tuple, Optional, Dict, Any, Union

import pandas as pd


# ------------------------------------------------------------
# 0) Errori
# ------------------------------------------------------------

class DSLError(ValueError):
    """Errore di sintassi nel DSL. `line_no` e' 1-based (None se non applicabile)."""

    def __init__(self, message: str, line_no: Optional[int] = None, line: str = ""):
        self.message = message
        self.line_no = line_no
        self.line = line
        where = f"riga {line_no}: " if line_no else ""
        snippet = f"  →  «{line.strip()}»" if line.strip() else ""
        super().__init__(f"{where}{message}{snippet}")


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip().lower() in ("", "nan", "none", "<na>")


# ------------------------------------------------------------
# 1) Espansione dei repeat mantenendo la struttura
# ------------------------------------------------------------

_TAB_WIDTH = 2  # un tab vale un livello di indentazione (2 spazi)


def _count_leading_spaces(s: str) -> int:
    """Conta gli spazi iniziali in una stringa (i tab valgono 2 spazi)."""
    s = s.replace("\t", " " * _TAB_WIDTH)
    return len(s) - len(s.lstrip(" "))


_REPEAT_RE = re.compile(r"^repeat\s+(-?\d+)\s*:?\s*$", re.IGNORECASE)


def _parse_structure(steps_text: str) -> List[Any]:
    """
    Converte il testo in una struttura annidata.
    Gli step normali sono tuple (line_no, testo); i blocchi repeat sono
    dict {"repeat": N, "steps": [...], "line_no": n}.
    Solleva DSLError su indentazione incoerente o repeat vuoti.
    """
    lines = []
    for i, raw in enumerate(str(steps_text).splitlines(), start=1):
        if raw.strip() == "" or raw.strip().startswith("#"):
            continue
        lines.append((i, _count_leading_spaces(raw), raw.strip(), raw))

    if not lines:
        raise DSLError("la colonna Steps e' vuota")

    pos = 0

    def parse_block(indent: int) -> List[Any]:
        nonlocal pos
        items: List[Any] = []
        while pos < len(lines):
            line_no, ind, content, raw = lines[pos]
            if ind < indent:
                break
            if ind > indent:
                raise DSLError("indentazione inattesa (la riga e' rientrata ma non e' dentro un 'repeat')",
                               line_no, raw)
            m = _REPEAT_RE.match(content)
            if m:
                reps = int(m.group(1))
                if reps < 1:
                    raise DSLError(f"numero di ripetizioni non valido ({reps})", line_no, raw)
                pos += 1
                if pos >= len(lines) or lines[pos][1] <= indent:
                    raise DSLError("'repeat' senza step: rientra gli step sotto il repeat", line_no, raw)
                child_indent = lines[pos][1]
                children = parse_block(child_indent)
                items.append({"repeat": reps, "steps": children, "line_no": line_no})
                continue
            if content.lower().startswith("repeat"):
                raise DSLError("sintassi repeat non valida: usa 'repeat N:'", line_no, raw)
            items.append((line_no, content))
            pos += 1
        return items

    first_indent = lines[0][1]
    result = parse_block(first_indent)
    if pos < len(lines):  # righe meno rientrate della prima
        line_no, _, _, raw = lines[pos]
        raise DSLError("indentazione incoerente rispetto alla prima riga", line_no, raw)
    return result


def _strip_line_numbers(items: List[Any]) -> List[Union[str, Dict[str, Any]]]:
    out: List[Any] = []
    for it in items:
        if isinstance(it, dict):
            out.append({"repeat": it["repeat"], "steps": _strip_line_numbers(it["steps"])})
        else:
            out.append(it[1])
    return out


def expand_repeat_lines(steps_text: str) -> List[Union[str, Dict[str, Any]]]:
    """
    Prende il testo degli steps e ritorna una lista mista:
    - Stringhe per step normali
    - Dict {"repeat": N, "steps": [...]} per blocchi repeat
    Solleva DSLError se la struttura non e' valida.
    """
    return _strip_line_numbers(_parse_structure(steps_text))


# ------------------------------------------------------------
# 2) Parse durata
# ------------------------------------------------------------

_NUM = r"(\d+(?:[.,]\d+)?)"


def _num(s: str) -> float:
    return float(s.replace(",", "."))


def parse_duration_strict(base_str: str) -> Tuple[str, Optional[float], Optional[str]]:
    """
    Parsea la durata/distanza di uno step.
    Ritorna (tipo condizione, valore in secondi o metri, unita' preferita).
    Solleva ValueError se il formato non e' riconosciuto.
    """
    s = base_str.strip().lower()

    if not s:
        raise ValueError("durata mancante (es. 10min, 1km, 400m, lap-button)")

    if s in ("lap-button", "lap button", "lap", "lapbutton"):
        return "lap.button", None, None

    # Metri: "1000m", "200 m", "400mt"  (NB: "m" = METRI, i minuti si scrivono "min")
    m = re.match(rf"^{_NUM}\s*(?:m|mt|metri|meters?)$", s)
    if m:
        return "distance", _num(m.group(1)), "meter"

    # Chilometri: "5km", "1.5 km"
    m = re.match(rf"^{_NUM}\s*km$", s)
    if m:
        return "distance", _num(m.group(1)) * 1000.0, "kilometer"

    # Ore: "1h", "1.5 h", "2 ore"
    m = re.match(rf"^{_NUM}\s*(?:h|hr|ora|ore|hours?)$", s)
    if m:
        return "time", _num(m.group(1)) * 3600.0, None

    # Minuti: "10min", "10 min", "1.5min", "10'", "10′", "10 minuti"
    m = re.match(rf"^{_NUM}\s*(?:min|mins|minuti|minutes?|'|′|’)$", s)
    if m:
        return "time", _num(m.group(1)) * 60.0, None

    # Secondi: "30sec", "30s", "30\"", "30″", "30 secondi"
    m = re.match(rf"^{_NUM}\s*(?:s|sec|secs|secondi|seconds?|\"|″|”|'')$", s)
    if m:
        return "time", _num(m.group(1)), None

    # mm:ss oppure h:mm:ss
    m = re.match(r"^(?:(\d+):)?(\d{1,2}):(\d{2})$", s)
    if m:
        h = int(m.group(1) or 0)
        mi = int(m.group(2))
        sec = int(m.group(3))
        if sec >= 60 or (m.group(1) and mi >= 60):
            raise ValueError(f"durata non valida '{base_str.strip()}'")
        return "time", float(h * 3600 + mi * 60 + sec), None

    raise ValueError(
        f"durata non riconosciuta '{base_str.strip()}' "
        "(usa es. 10min, 30sec, 1h, 1:30, 400m, 5km, lap-button)"
    )


def parse_duration_part(base_str: str) -> Tuple[str, Optional[float], Optional[str]]:
    """Versione tollerante (compatibilita'): in caso di errore ritorna 60 secondi."""
    try:
        return parse_duration_strict(base_str)
    except ValueError:
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

# parola chiave DSL -> stepTypeKey Garmin
_STEP_KEYWORDS = {
    "warmup": "warmup", "riscaldamento": "warmup",
    "cooldown": "cooldown", "defaticamento": "cooldown",
    "interval": "interval", "step": "interval", "run": "interval",
    "bike": "interval", "swim": "interval", "active": "interval",
    "recovery": "recovery", "recover": "recovery", "recupero": "recovery",
    "rest": "rest", "riposo": "rest",
}


# Se Garmin rifiutasse gli step di tipo "rest", metti False per tornare al
# comportamento precedente (rest inviato come recovery).
REST_AS_REST = True


def map_step_type(step_key: str) -> Dict[str, Any]:
    """Mappa la parola chiave DSL al stepType Garmin."""
    type_key = _STEP_KEYWORDS.get(step_key.strip().lower(), "interval")
    if type_key == "rest" and not REST_AS_REST:
        type_key = "recovery"
    return {
        "stepTypeId": STEP_TYPE_IDS[type_key],
        "stepTypeKey": type_key
    }


# ------------------------------------------------------------
# 4) Funzioni helper per Parameters
# ------------------------------------------------------------

def get_parameter_value(key: str, df_parameters: Optional[pd.DataFrame] = None) -> Optional[str]:
    """Cerca un parametro nel foglio Parameters (senza distinzione maiuscole/minuscole)."""
    if df_parameters is None or "Key" not in df_parameters.columns:
        return None
    keys = df_parameters["Key"].astype(str).str.strip().str.lower()
    param_row = df_parameters[keys == str(key).strip().lower()]
    if not param_row.empty:
        val = param_row.iloc[0]["Expression"]
        if _is_blank(val):
            return None
        return str(val).strip()
    return None


def _get_float_param(key: str, df_parameters: Optional[pd.DataFrame], default):
    expr = get_parameter_value(key, df_parameters)
    if expr is not None:
        try:
            return float(expr)
        except (ValueError, TypeError):
            pass
    return default


def get_pace_tolerance(df_parameters: Optional[pd.DataFrame] = None) -> float:
    """Legge pace_tolerance da Parameters (default: 10 sec)."""
    return _get_float_param('pace_tolerance', df_parameters, 10.0)


def get_hr_tolerance(df_parameters: Optional[pd.DataFrame] = None) -> float:
    """Legge hr_tolerance da Parameters (default: 5 bpm)."""
    return _get_float_param('hr_tolerance', df_parameters, 5.0)


def get_hr_max(df_parameters: Optional[pd.DataFrame] = None) -> Optional[float]:
    """Legge HR_max da Parameters."""
    return _get_float_param('HR_max', df_parameters, None)


def parse_single_pace(s: str) -> float:
    """Parsea singolo valore di ritmo: '5:00' -> 300 sec."""
    s = s.strip()
    if ':' in s:
        parts = s.split(':')
        if len(parts) == 2:
            mi, sec = int(parts[0]), int(parts[1])
            if sec >= 60:
                raise ValueError(f"ritmo non valido '{s}' (i secondi devono essere < 60)")
            return float(mi * 60 + sec)
    return float(s)


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


def pace_seconds_to_mps(pace_seconds: float) -> float:
    """Converte sec/km a m/s: 300 sec/km -> 3.33 m/s."""
    return 1000.0 / pace_seconds if pace_seconds > 0 else 1.0


def parse_hr_expression(expr: str, df_parameters: Optional[pd.DataFrame] = None) -> Tuple[float, Optional[float]]:
    """
    Parsea espressione HR: '150', '150-165', '80%', '70-85%'
    Se sono percentuali, usa HR_max da Parameters. Il risultato e' ordinato (min, max).
    """
    expr = expr.strip()

    m = re.match(r"^(\d+(?:\.\d+)?)%$", expr)
    if m:
        pct = float(m.group(1))
        hr_max = get_hr_max(df_parameters)
        if not hr_max:
            raise ValueError("zona FC in percentuale ma HR_max non e' definito in Parameters")
        hr_value = hr_max * (pct / 100.0)
        tolerance = get_hr_tolerance(df_parameters)
        return (hr_value - tolerance, hr_value + tolerance)

    m = re.match(r"^(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)%$", expr)
    if m:
        pct_min, pct_max = sorted((float(m.group(1)), float(m.group(2))))
        hr_max = get_hr_max(df_parameters)
        if not hr_max:
            raise ValueError("zona FC in percentuale ma HR_max non e' definito in Parameters")
        return (hr_max * (pct_min / 100.0), hr_max * (pct_max / 100.0))

    if '-' in expr:
        parts = expr.split('-')
        if len(parts) == 2:
            a, b = sorted((float(parts[0]), float(parts[1])))
            return (a, b)

    hr_value = float(expr)
    tolerance = get_hr_tolerance(df_parameters)
    return (hr_value - tolerance, hr_value + tolerance)


def parse_power_expression(expr: str) -> Tuple[float, Optional[float]]:
    """Parsea espressione potenza: '250W', '200-250W', '250'."""
    expr = expr.strip().upper().replace('W', '')
    if '-' in expr:
        parts = expr.split('-')
        if len(parts) == 2:
            return (float(parts[0]), float(parts[1]))
    return (float(expr), None)


def parse_cadence_expression(expr: str) -> Tuple[float, Optional[float]]:
    """Parsea espressione cadenza: '90rpm', '85-95rpm', '90'."""
    expr = expr.strip().lower().replace('rpm', '')
    if '-' in expr:
        parts = expr.split('-')
        if len(parts) == 2:
            return (float(parts[0]), float(parts[1]))
    return (float(expr), None)


def parse_swim_pace_expression(expr: str, df_parameters: Optional[pd.DataFrame] = None) -> Tuple[float, Optional[float]]:
    """
    Parsea espressione pace nuoto (min:sec per 100m): '1:45', '1:40-1:50'
    Converte in m/s per API Garmin.
    """
    expr = expr.strip().lower().replace("/100m", "").strip()

    if '-' in expr and ':' in expr:
        parts = expr.split('-')
        if len(parts) == 2:
            pace1_sec = parse_single_pace(parts[0].strip())
            pace2_sec = parse_single_pace(parts[1].strip())
            mps1 = 100.0 / pace1_sec if pace1_sec > 0 else 1.0
            mps2 = 100.0 / pace2_sec if pace2_sec > 0 else 1.0
            return (min(mps1, mps2), max(mps1, mps2))

    if ':' in expr:
        pace_sec = parse_single_pace(expr)
        tolerance_param = get_parameter_value('swim_tolerance', df_parameters)
        tolerance = float(tolerance_param) if tolerance_param else 5.0
        slower_sec = pace_sec + tolerance
        faster_sec = pace_sec - tolerance
        slower_mps = 100.0 / slower_sec if slower_sec > 0 else 1.0
        faster_mps = 100.0 / faster_sec if faster_sec > 0 else 1.0
        return (slower_mps, faster_mps)

    raise ValueError(f"ritmo nuoto non valido '{expr}' (usa es. 1:45 o 1:40-1:50)")


# ------------------------------------------------------------
# 5) Parse target
# ------------------------------------------------------------

def _target(key: str, one=None, two=None, zone=None) -> Dict[str, Any]:
    return {
        "targetType": {
            "workoutTargetTypeId": TARGET_TYPE_IDS[key],
            "workoutTargetTypeKey": key,
        },
        "targetValueOne": one,
        "targetValueTwo": two,
        "zoneNumber": zone,
    }


def _no_target() -> Dict[str, Any]:
    return _target("no.target")


def _pace_target_from_expr(expr: str, df_parameters) -> Dict[str, Any]:
    min_sec, max_sec = parse_pace_expression(expr)
    if max_sec is None:
        tolerance_sec = get_pace_tolerance(df_parameters)
        slower_mps = pace_seconds_to_mps(min_sec + tolerance_sec)
        faster_mps = pace_seconds_to_mps(min_sec - tolerance_sec)
    else:
        slower_mps = pace_seconds_to_mps(max(min_sec, max_sec))
        faster_mps = pace_seconds_to_mps(min(min_sec, max_sec))
    return _target("pace.zone", slower_mps, faster_mps)


def parse_target_strict(target_str: str, df_parameters: Optional[pd.DataFrame] = None) -> Dict[str, Any]:
    """
    Parsea il target dopo '@'. Solleva ValueError se il target non e' riconosciuto.
    Supporta: ritmo (5:00, 5:00-5:30, Z1-Z5, chiavi custom), FC (HR_Z1-5, 140-160),
    potenza (200W, Power_Z*), cadenza (90rpm, Cadence_*), nuoto (1:45/100m, Swim_*).
    """
    raw = target_str.strip()
    s = raw.lower()

    if not s or s in ("open", "libero", "none"):
        return _no_target()

    # Zone di ritmo: Z1..Z5
    m = re.match(r"^z(\d+)$", s)
    if m:
        zone = int(m.group(1))
        if not 1 <= zone <= 5:
            raise ValueError(f"zona '{raw}' non valida (usa Z1-Z5)")
        zone_expr = get_parameter_value(f"Z{zone}", df_parameters)
        if zone_expr:
            return _pace_target_from_expr(zone_expr, df_parameters)
        return _target("pace.zone", zone=zone)  # zone Garmin predefinite

    # Zone FC: HR1, HR_Z1, HRZ1
    m = re.match(r"^hr[_]?z?(\d+)$", s)
    if m:
        zone = int(m.group(1))
        if not 1 <= zone <= 5:
            raise ValueError(f"zona FC '{raw}' non valida (usa HR_Z1-HR_Z5)")
        for zone_key in (f"HR{zone}", f"HR_Z{zone}", f"HRZ{zone}"):
            zone_expr = get_parameter_value(zone_key, df_parameters)
            if zone_expr:
                min_hr, max_hr = parse_hr_expression(zone_expr, df_parameters)
                return _target("heart.rate.zone", min_hr, max_hr)
        return _target("heart.rate.zone", zone=zone)

    # Ritmo assoluto: "5:00" o "5:00-5:30"
    m = re.match(r"^(\d+):(\d+)(?:\s*-\s*(\d+):(\d+))?$", s)
    if m:
        if int(m.group(2)) >= 60 or (m.group(4) and int(m.group(4)) >= 60):
            raise ValueError(f"ritmo '{raw}' non valido (i secondi devono essere < 60)")
        total_sec = int(m.group(1)) * 60 + int(m.group(2))
        if m.group(3):
            total_sec2 = int(m.group(3)) * 60 + int(m.group(4))
            p1, p2 = pace_seconds_to_mps(total_sec), pace_seconds_to_mps(total_sec2)
            return _target("pace.zone", min(p1, p2), max(p1, p2))
        tolerance_sec = get_pace_tolerance(df_parameters)
        return _target("pace.zone",
                       pace_seconds_to_mps(total_sec + tolerance_sec),
                       pace_seconds_to_mps(total_sec - tolerance_sec))

    # Ritmo nuoto diretto: "1:45/100m" o "1:40-1:50/100m"
    if s.endswith("/100m"):
        mps1, mps2 = parse_swim_pace_expression(s, df_parameters)
        return _target("pace.zone", mps1, mps2)

    # Zone POTENZA: Power_Z1 ...
    if re.match(r"^power[_]?z(\d)$", s):
        zone_expr = get_parameter_value(s.replace("powerz", "power_z"), df_parameters) \
            or get_parameter_value(s, df_parameters)
        if not zone_expr:
            raise ValueError(f"'{raw}' non e' definito nel foglio Parameters")
        power1, power2 = parse_power_expression(zone_expr)
        if power2 is None:
            power2 = power1 * 1.05
        return _target("power.zone", power1, power2)

    # Cadenza custom: Cadence_Easy ...
    if s.startswith('cadence'):
        custom_cadence = get_parameter_value(s, df_parameters)
        if not custom_cadence:
            raise ValueError(f"'{raw}' non e' definito nel foglio Parameters")
        cad1, cad2 = parse_cadence_expression(custom_cadence)
        if cad2 is None:
            cad2 = cad1 + 5
        return _target("cadence", cad1, cad2)

    # Nuoto custom: Swim_Z2, swim_easy ...
    if s.startswith('swim'):
        custom_swim = get_parameter_value(s, df_parameters)
        if not custom_swim:
            raise ValueError(f"'{raw}' non e' definito nel foglio Parameters")
        mps1, mps2 = parse_swim_pace_expression(custom_swim, df_parameters)
        return _target("pace.zone", mps1, mps2)

    # Parametri custom di ritmo (easy_range, hmp, threshold, ...)
    custom_pace = get_parameter_value(s, df_parameters)
    if custom_pace:
        try:
            return _pace_target_from_expr(custom_pace, df_parameters)
        except ValueError:
            raise ValueError(f"il parametro '{raw}' = '{custom_pace}' non e' un ritmo valido")

    # Potenza: "200W" o "200-250W"
    m = re.match(r"^(\d+)(?:\s*-\s*(\d+))?\s*w$", s)
    if m:
        power1 = float(m.group(1))
        power2 = float(m.group(2)) if m.group(2) else power1 * 1.05
        return _target("power.zone", power1, power2)

    # Cadenza: "85rpm" o "85-90rpm"
    m = re.match(r"^(\d+)(?:\s*-\s*(\d+))?\s*rpm$", s)
    if m:
        cadence1 = float(m.group(1))
        cadence2 = float(m.group(2)) if m.group(2) else cadence1 + 5
        return _target("cadence", cadence1, cadence2)

    # FC assoluta: "140-160bpm", "140-160", "150"
    m = re.match(r"^(\d+)(?:\s*-\s*(\d+))?\s*(?:bpm)?$", s)
    if m:
        hr1 = float(m.group(1))
        hr2 = float(m.group(2)) if m.group(2) else hr1 + 10
        if min(hr1, hr2) < 30 or max(hr1, hr2) > 250:
            raise ValueError(f"frequenza cardiaca '{raw}' fuori scala (30-250 bpm)")
        return _target("heart.rate.zone", min(hr1, hr2), max(hr1, hr2))

    raise ValueError(
        f"target '{raw}' non riconosciuto: non e' un ritmo (5:00), una zona (Z1-Z5, HR_Z1-5) "
        "ne' una chiave del foglio Parameters"
    )


def parse_target(target_str: str, df_parameters: Optional[pd.DataFrame] = None) -> Dict[str, Any]:
    """Versione tollerante (compatibilita'): target non riconosciuto -> nessun target."""
    try:
        return parse_target_strict(target_str, df_parameters)
    except (ValueError, TypeError):
        return _no_target()


# ------------------------------------------------------------
# 6) Parse singola riga di step
# ------------------------------------------------------------

_STEP_LINE_RE = re.compile(r"^([a-zA-Zàèéìòù]+)\s*:\s*(.*)$")


def parse_step_line(
    line: str,
    order: int,
    df_parameters: Optional[pd.DataFrame] = None,
    line_no: Optional[int] = None,
) -> Dict[str, Any]:
    """Parsea una riga di step e ritorna un ExecutableStepDTO. Solleva DSLError."""
    main_part = line.split("--", 1)[0].rstrip()

    m = _STEP_LINE_RE.match(main_part.strip())
    if not m:
        raise DSLError("formato step non valido: usa 'tipo: durata @ target' "
                       "(es. interval: 1km @ 5:00)", line_no, line)
    step_type_key = m.group(1).strip().lower()
    if step_type_key not in _STEP_KEYWORDS:
        raise DSLError(f"tipo di step '{m.group(1)}' sconosciuto "
                       "(usa warmup, interval, recovery, rest, cooldown)", line_no, line)
    base_and_target = m.group(2).strip()

    target_str = ""
    if "@" in base_and_target:
        base_str, target_str = base_and_target.split("@", 1)
        base_str = base_str.strip()
        target_str = target_str.strip()
        if not target_str:
            raise DSLError("'@' senza target", line_no, line)
    else:
        base_str = base_and_target.strip()

    try:
        condition_type_key, end_condition_value, preferred_unit_key = parse_duration_strict(base_str)
        if end_condition_value is not None and end_condition_value <= 0:
            raise ValueError("la durata deve essere maggiore di zero")
    except ValueError as e:
        raise DSLError(str(e), line_no, line) from None

    try:
        target_info = parse_target_strict(target_str, df_parameters)
    except (ValueError, TypeError) as e:
        raise DSLError(str(e), line_no, line) from None

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

    if end_condition_value is not None:
        step["endConditionValue"] = end_condition_value

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
    raw_val = row.get("Sport", "")
    raw = "" if _is_blank(raw_val) else str(raw_val).strip().lower()
    key = "running"
    if "bike" in raw or "bici" in raw or "cycling" in raw or "cicl" in raw:
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
    Converte la struttura (da _parse_structure o expand_repeat_lines) in workout steps Garmin.
    Gestisce sia ExecutableStepDTO che RepeatGroupDTO.
    """
    steps = []
    order = start_order
    child_id = start_child_id

    for item in structured_steps:
        if isinstance(item, dict) and "repeat" in item:
            nested_steps = build_workout_steps(
                item["steps"],
                df_parameters,
                start_order=1,
                start_child_id=child_id + 1
            )
            steps.append({
                "type": "RepeatGroupDTO",
                "stepId": None,
                "stepOrder": order,
                "numberOfIterations": item["repeat"],
                "smartRepeat": False,
                "childStepId": child_id,
                "workoutSteps": nested_steps
            })
            order += 1
            child_id += 1
        else:
            if isinstance(item, tuple):
                line_no, text = item
            else:
                line_no, text = None, item
            steps.append(parse_step_line(text, order, df_parameters, line_no=line_no))
            order += 1

    return steps


# ------------------------------------------------------------
# 9) Costruzione workout completo
# ------------------------------------------------------------

def _clean_int_str(value: Any) -> str:
    if _is_blank(value):
        return ""
    if isinstance(value, float) and not math.isnan(value) and value.is_integer():
        return str(int(value))
    s = str(value).strip()
    return s[:-2] if s.endswith(".0") else s


def build_garmin_workout_from_excel_row(
    row: pd.Series,
    steps_text: str,
    df_parameters: Optional[pd.DataFrame] = None,
    use_prefix: bool = False,
) -> Dict[str, Any]:
    """Costruisce il workout JSON Garmin da una riga Excel. Solleva DSLError se gli step non sono validi."""

    desc = row.get("Description", "")
    workout_name = "" if _is_blank(desc) else str(desc).strip()
    if not workout_name:
        workout_name = "Workout"

    if use_prefix:
        parts = []
        week = _clean_int_str(row.get("Week", ""))
        session = _clean_int_str(row.get("Session", ""))
        if week:
            parts.append(f"W{week}")
        if session:
            parts.append(f"S{session}")
        if parts:
            workout_name = "".join(parts) + " - " + workout_name

    sport_type = map_sport_type(row)

    if _is_blank(steps_text):
        raise DSLError("la colonna Steps e' vuota")

    structured_steps = _parse_structure(steps_text)
    steps = build_workout_steps(structured_steps, df_parameters)

    segment = {
        "segmentOrder": 1,
        "sportType": sport_type,
        "workoutSteps": steps,
    }

    return {
        "workoutId": None,
        "ownerId": None,
        "workoutName": workout_name,
        "description": workout_name,
        "sportType": sport_type,
        "workoutSegments": [segment],
    }


# ------------------------------------------------------------
# 10) Validazione di un intero piano (prima dell'upload)
# ------------------------------------------------------------

def _row_label(df_workouts: pd.DataFrame, idx: int) -> str:
    row = df_workouts.iloc[idx]
    parts = [f"Riga Excel {idx + 2}"]
    week, session = _clean_int_str(row.get("Week", "")), _clean_int_str(row.get("Session", ""))
    if week or session:
        parts.append(f"W{week}S{session}")
    desc = row.get("Description", "")
    if not _is_blank(desc):
        parts.append(f"«{str(desc).strip()[:40]}»")
    return " ".join(parts)


def _date_problem(value: Any) -> Optional[str]:
    if _is_blank(value):
        return "data mancante nella colonna Date"
    try:
        pd.to_datetime(value)
        return None
    except (ValueError, TypeError):
        return f"data non valida nella colonna Date ({value})"


def validate_workout_rows(
    df_workouts: pd.DataFrame,
    indices: List[int],
    df_parameters: Optional[pd.DataFrame] = None,
    require_date: bool = False,
) -> List[str]:
    """
    Controlla i workout indicati (indici posizionali) senza contattare Garmin.
    Ritorna una lista di messaggi d'errore leggibili (vuota se tutto e' valido).
    """
    errors: List[str] = []
    for idx in indices:
        row = df_workouts.iloc[idx]
        label = _row_label(df_workouts, idx)
        try:
            build_garmin_workout_from_excel_row(row, row.get("Steps", ""), df_parameters)
        except DSLError as e:
            errors.append(f"{label}: {e}")
        except Exception as e:  # errore inatteso: meglio mostrarlo che caricare dati sbagliati
            errors.append(f"{label}: errore inatteso ({e})")
        if require_date:
            problem = _date_problem(row.get("Date", ""))
            if problem:
                errors.append(f"{label}: {problem}")
    return errors
