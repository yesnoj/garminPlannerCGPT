"""
Modello dati dell'editor: albero di step/ripetizioni <-> testo DSL,
stima di durata, distanza e intensita' (per il grafico).

Non dipende da tkinter: e' testabile da solo.
"""
from __future__ import annotations

import itertools
import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple, Union

import pandas as pd

from dsl_parser import (
    DSLError, _parse_structure, get_parameter_value, parse_duration_strict,
    parse_step_line, parse_target_strict, get_hr_max,
)

# ---------------------------------------------------------------------------
# Etichette
# ---------------------------------------------------------------------------

STEP_TYPES = ["warmup", "interval", "recovery", "rest", "cooldown"]
STEP_LABELS = {
    "warmup": "Riscaldamento",
    "interval": "Lavoro",
    "recovery": "Recupero",
    "rest": "Riposo",
    "cooldown": "Defaticamento",
}
_ALIASES = {
    "riscaldamento": "warmup", "defaticamento": "cooldown", "step": "interval",
    "run": "interval", "bike": "interval", "swim": "interval", "active": "interval",
    "recover": "recovery", "recupero": "recovery", "riposo": "rest",
}

_ids = itertools.count(1)


def _new_id() -> int:
    return next(_ids)


# ---------------------------------------------------------------------------
# Nodi
# ---------------------------------------------------------------------------

@dataclass
class Step:
    type: str = "interval"
    duration: str = "10min"
    target: str = ""
    comment: str = ""
    id: int = field(default_factory=_new_id)

    kind = "step"

    def to_line(self) -> str:
        line = f"{self.type}: {self.duration}"
        if self.target:
            line += f" @ {self.target}"
        if self.comment:
            line += f"  -- {self.comment}"
        return line

    def copy(self) -> "Step":
        return Step(self.type, self.duration, self.target, self.comment)


@dataclass
class Repeat:
    count: int = 4
    children: List[Union[Step, "Repeat"]] = field(default_factory=list)
    id: int = field(default_factory=_new_id)

    kind = "repeat"

    def copy(self) -> "Repeat":
        return Repeat(self.count, [c.copy() for c in self.children])


Node = Union[Step, Repeat]

_LINE_RE = re.compile(r"^\s*([A-Za-zàèéìòù]+)\s*:\s*(.*?)\s*$")


def parse_line(text: str) -> Step:
    """'interval: 1km @ 5:00 -- nota' -> Step. Non valida il contenuto (lo fa il parser)."""
    comment = ""
    if "--" in text:
        text, comment = text.split("--", 1)
        comment = comment.strip()
    m = _LINE_RE.match(text)
    if not m:
        return Step("interval", text.strip(), "", comment)
    kind = m.group(1).lower()
    kind = _ALIASES.get(kind, kind)
    rest = m.group(2)
    target = ""
    if "@" in rest:
        rest, target = rest.split("@", 1)
    return Step(kind, rest.strip(), target.strip(), comment)


def from_dsl(text: str) -> List[Node]:
    """Testo DSL -> lista di nodi. Solleva DSLError se la struttura non e' valida."""
    def convert(items) -> List[Node]:
        out: List[Node] = []
        for it in items:
            if isinstance(it, dict):
                out.append(Repeat(it["repeat"], convert(it["steps"])))
            else:
                out.append(parse_line(it[1]))
        return out

    return convert(_parse_structure(text))


def to_dsl(nodes: List[Node], indent: int = 0) -> str:
    """Lista di nodi -> testo DSL (2 spazi per livello)."""
    lines: List[str] = []
    pad = "  " * indent
    for n in nodes:
        if isinstance(n, Repeat):
            lines.append(f"{pad}repeat {n.count}:")
            sub = to_dsl(n.children, indent + 1)
            if sub:
                lines.append(sub)
        else:
            lines.append(pad + n.to_line())
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Operazioni sull'albero
# ---------------------------------------------------------------------------

def find(nodes: List[Node], node_id: int) -> Optional[Tuple[List[Node], int]]:
    """Ritorna (lista che contiene il nodo, indice) oppure None."""
    for i, n in enumerate(nodes):
        if n.id == node_id:
            return nodes, i
        if isinstance(n, Repeat):
            r = find(n.children, node_id)
            if r:
                return r
    return None


def parent_of(nodes: List[Node], node_id: int) -> Optional[Repeat]:
    """Ripetizione che contiene direttamente il nodo (None se e' al livello principale)."""
    for n in nodes:
        if isinstance(n, Repeat):
            if any(c.id == node_id for c in n.children):
                return n
            p = parent_of(n.children, node_id)
            if p is not None:
                return p
    return None


def get(nodes: List[Node], node_id: int) -> Optional[Node]:
    r = find(nodes, node_id)
    return r[0][r[1]] if r else None


def insert_after(nodes: List[Node], ref_id: Optional[int], new: Node, inside: bool = False) -> None:
    """Inserisce `new` dopo il nodo ref (o dentro, in fondo, se ref e' una ripetizione e inside=True)."""
    if ref_id is None:
        nodes.append(new)
        return
    ref = get(nodes, ref_id)
    if inside and isinstance(ref, Repeat):
        ref.children.append(new)
        return
    lst, i = find(nodes, ref_id)
    lst.insert(i + 1, new)


def remove(nodes: List[Node], node_id: int) -> None:
    r = find(nodes, node_id)
    if r:
        del r[0][r[1]]


def move(nodes: List[Node], node_id: int, delta: int) -> bool:
    r = find(nodes, node_id)
    if not r:
        return False
    lst, i = r
    j = i + delta
    if 0 <= j < len(lst):
        lst[i], lst[j] = lst[j], lst[i]
        return True
    return False


def indent(nodes: List[Node], node_id: int) -> bool:
    """Sposta il nodo dentro la ripetizione che lo precede."""
    r = find(nodes, node_id)
    if not r:
        return False
    lst, i = r
    if i > 0 and isinstance(lst[i - 1], Repeat):
        node = lst.pop(i)
        lst[i - 1].children.append(node)
        return True
    return False


def outdent(nodes: List[Node], node_id: int) -> bool:
    """Porta il nodo fuori dalla ripetizione che lo contiene (subito dopo di essa)."""
    parent = parent_of(nodes, node_id)
    if parent is None:
        return False
    node = get(nodes, node_id)
    remove(nodes, node_id)
    glst, gi = find(nodes, parent.id)
    glst.insert(gi + 1, node)
    return True


def contains(node: Node, other_id: int) -> bool:
    """True se other_id e' il nodo stesso o un suo discendente."""
    if node.id == other_id:
        return True
    if isinstance(node, Repeat):
        return any(contains(c, other_id) for c in node.children)
    return False


def move_to(nodes: List[Node], node_id: int, target_id: int, where: str) -> bool:
    """
    Sposta un nodo prima/dopo un altro nodo, oppure dentro una ripetizione (in fondo).
    where: "before" | "after" | "inside". Ritorna False se lo spostamento non e' valido.
    """
    node = get(nodes, node_id)
    target = get(nodes, target_id)
    if node is None or target is None or node_id == target_id:
        return False
    if contains(node, target_id):          # una ripetizione non puo' finire dentro se stessa
        return False
    if where == "inside" and not isinstance(target, Repeat):
        return False
    remove(nodes, node_id)
    if where == "inside":
        target.children.append(node)
        return True
    lst, i = find(nodes, target_id)
    lst.insert(i if where == "before" else i + 1, node)
    return True


def iter_steps(nodes: List[Node]):
    for n in nodes:
        if isinstance(n, Repeat):
            yield from iter_steps(n.children)
        else:
            yield n


# ---------------------------------------------------------------------------
# Stima durata / distanza / intensita'
# ---------------------------------------------------------------------------

DEFAULT_SPEED = {"running": 1000 / 390, "cycling": 25 / 3.6, "swimming": 100 / 120}  # 6:30/km, 25 km/h, 2:00/100m
DEFAULT_INTENSITY = {"warmup": 0.45, "interval": 0.8, "recovery": 0.3, "rest": 0.12, "cooldown": 0.38}
LAP_BUTTON_SECONDS = 300


@dataclass
class Segment:
    type: str
    seconds: float
    meters: float
    intensity: float          # 0..1 per l'altezza nel grafico
    label: str                # testo dello step
    estimated: bool = False   # durata/distanza stimate (lap-button o senza ritmo)


def sport_key(sport: str) -> str:
    s = (sport or "").lower()
    if any(k in s for k in ("bike", "bici", "cycl", "cicl")):
        return "cycling"
    if any(k in s for k in ("swim", "nuoto")):
        return "swimming"
    return "running"


def _easy_speed(sport: str, params: Optional[pd.DataFrame]) -> float:
    if sport == "running":
        for key in ("easy_range", "Z2"):
            try:
                t = parse_target_strict(key, params) if get_parameter_value(key, params) else None
            except ValueError:
                t = None
            if t and t["targetValueOne"]:
                return (t["targetValueOne"] + t["targetValueTwo"]) / 2
    return DEFAULT_SPEED[sport]


def _intensity_from_target(t: dict, sport: str, easy_speed: float, params) -> Optional[float]:
    key = t["targetType"]["workoutTargetTypeKey"]
    v1, v2 = t["targetValueOne"], t["targetValueTwo"]
    if key == "pace.zone":
        if v1 and v2:
            speed = (v1 + v2) / 2
            # easy -> ~0.45, easy*1.35 (ripetute veloci) -> ~1.0
            return max(0.1, min(1.0, 0.45 + (speed / easy_speed - 1) * 1.6))
        if t["zoneNumber"]:
            return 0.2 + 0.16 * t["zoneNumber"]
    if key == "heart.rate.zone":
        if v1 and v2:
            hr_max = get_hr_max(params) or 185
            pct = ((v1 + v2) / 2) / hr_max
            return max(0.1, min(1.0, (pct - 0.55) / 0.4))
        if t["zoneNumber"]:
            return 0.2 + 0.16 * t["zoneNumber"]
    if key == "power.zone" and v1:
        ftp = None
        try:
            ftp = float(get_parameter_value("FTP", params) or 0) or None
        except ValueError:
            pass
        watts = ((v1 or 0) + (v2 or v1)) / 2
        return max(0.1, min(1.0, watts / (ftp or 250) * 0.8))
    return None


def _speed_from_target(t: dict) -> Optional[float]:
    if t["targetType"]["workoutTargetTypeKey"] == "pace.zone" and t["targetValueOne"] and t["targetValueTwo"]:
        return (t["targetValueOne"] + t["targetValueTwo"]) / 2
    return None


def segments(nodes: List[Node], sport: str = "running", params: Optional[pd.DataFrame] = None) -> List[Segment]:
    """Espande le ripetizioni e stima ogni step. Gli step non validi vengono saltati."""
    sp = sport_key(sport)
    easy = _easy_speed(sp, params)
    out: List[Segment] = []

    def walk(items: List[Node]):
        for n in items:
            if isinstance(n, Repeat):
                for _ in range(max(0, n.count)):
                    walk(n.children)
                continue
            try:
                cond, value, _unit = parse_duration_strict(n.duration)
                t = parse_target_strict(n.target, params)
            except ValueError:
                continue
            speed = _speed_from_target(t)
            estimated = False
            if n.type == "rest":
                speed_eff = 0.0
            else:
                speed_eff = speed or easy * (1.0 if n.type != "interval" else 1.15)
                if speed is None:
                    estimated = True
            if cond == "time":
                secs, meters = value, value * speed_eff
            elif cond == "distance":
                meters = value
                secs = value / speed_eff if speed_eff else 0.0
            else:  # lap-button
                secs, meters, estimated = LAP_BUTTON_SECONDS, LAP_BUTTON_SECONDS * speed_eff, True
            inten = _intensity_from_target(t, sp, easy, params)
            if inten is None:
                inten = DEFAULT_INTENSITY.get(n.type, 0.5)
            if n.type == "rest":
                inten = DEFAULT_INTENSITY["rest"]
            out.append(Segment(n.type, secs, meters, inten, n.to_line(), estimated))

    walk(nodes)
    return out


def totals(segs: List[Segment]) -> Tuple[float, float, bool]:
    """(secondi, metri, stima approssimativa?)"""
    return (sum(s.seconds for s in segs), sum(s.meters for s in segs), any(s.estimated for s in segs))


def estimate_text(steps_text: str, sport: str, params) -> Tuple[float, float]:
    """Durata (s) e distanza (m) stimate di un testo DSL; (0, 0) se non valido."""
    try:
        nodes = from_dsl(steps_text)
    except DSLError:
        return 0.0, 0.0
    secs, meters, _ = totals(segments(nodes, sport, params))
    return secs, meters


# ---------------------------------------------------------------------------
# Formattazione
# ---------------------------------------------------------------------------

def fmt_duration(seconds: float) -> str:
    seconds = int(round(seconds))
    if seconds <= 0:
        return "–"
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    if h:
        return f"{h}h{m:02d}'"
    if m < 10 and s:
        return f"{m}'{s:02d}\""
    return f"{m + (1 if s >= 30 else 0)}'"


def fmt_km(meters: float) -> str:
    if meters <= 0:
        return "–"
    if meters < 1000:
        return f"{int(round(meters))} m"
    return f"{meters / 1000:.1f} km".replace(".", ",")


def validate_line(line: str, params) -> Optional[str]:
    """Messaggio d'errore per una singola riga di step (None se valida)."""
    try:
        parse_step_line(line, 1, params)
        return None
    except DSLError as e:
        return e.message


# ---------------------------------------------------------------------------
# Modelli rapidi
# ---------------------------------------------------------------------------

def _s(t, d, tg=""):
    return Step(t, d, tg)


PRESETS = {
    "running": [
        ("Riscaldamento 15' facile", lambda: [_s("warmup", "15min", "Z2")]),
        ("Defaticamento 10'", lambda: [_s("cooldown", "10min", "Z1")]),
        ("Corsa facile 40'", lambda: [_s("interval", "40min", "Z2")]),
        ("Allunghi 6×20\"", lambda: [Repeat(6, [_s("interval", "20sec"), _s("recovery", "1min", "Z1")])]),
        ("Ripetute 5×1000 m", lambda: [Repeat(5, [_s("interval", "1000m", "Z5"), _s("recovery", "2min", "Z1")])]),
        ("Ripetute 8×400 m", lambda: [Repeat(8, [_s("interval", "400m", "Z5"), _s("recovery", "90sec", "Z1")])]),
        ("Soglia 3×10'", lambda: [Repeat(3, [_s("interval", "10min", "Z4"), _s("recovery", "3min", "Z1")])]),
        ("Fartlek 6×(2' + 2')", lambda: [Repeat(6, [_s("interval", "2min", "Z4"), _s("recovery", "2min", "Z2")])]),
        ("Progressivo 3 km + 3 km + 2 km", lambda: [_s("interval", "3km", "Z2"), _s("interval", "3km", "Z3"),
                                                   _s("interval", "2km", "Z4")]),
    ],
    "cycling": [
        ("Riscaldamento 15'", lambda: [_s("warmup", "15min", "")]),
        ("Defaticamento 10'", lambda: [_s("cooldown", "10min", "")]),
        ("Sweet spot 3×12'", lambda: [Repeat(3, [_s("interval", "12min", ""), _s("recovery", "4min", "")])]),
        ("VO2max 5×4'", lambda: [Repeat(5, [_s("interval", "4min", ""), _s("recovery", "4min", "")])]),
    ],
    "swimming": [
        ("Riscaldamento 300 m", lambda: [_s("warmup", "300m", "")]),
        ("Defaticamento 200 m", lambda: [_s("cooldown", "200m", "")]),
        ("10×100 m rec 20\"", lambda: [Repeat(10, [_s("interval", "100m", ""), _s("rest", "20sec", "")])]),
        ("4×200 m rec 30\"", lambda: [Repeat(4, [_s("interval", "200m", ""), _s("rest", "30sec", "")])]),
    ],
}
