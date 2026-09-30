import pandas as pd
import pytest

import workout_model as wm
from dsl_parser import DSLError

PARAMS = pd.DataFrame([
    {"Key": "Z1", "Metric": "pace", "Expression": "7:00-7:30"},
    {"Key": "Z2", "Metric": "pace", "Expression": "6:30-7:00"},
    {"Key": "Z5", "Metric": "pace", "Expression": "4:50-5:00"},
])

TEXT = ("warmup: 15min @ Z2\n"
        "repeat 5:\n"
        "  interval: 1000m @ Z5\n"
        "  recovery: 2min @ Z1  -- corsetta\n"
        "cooldown: 10min @ Z1")


def test_roundtrip_dsl():
    nodes = wm.from_dsl(TEXT)
    assert wm.to_dsl(nodes) == TEXT
    assert isinstance(nodes[1], wm.Repeat) and nodes[1].count == 5
    assert nodes[1].children[1].comment == "corsetta"


def test_roundtrip_normalizes_indentation():
    nodes = wm.from_dsl("repeat 2:\n    interval: 1km @ Z2\n    rest: 1min")
    assert wm.to_dsl(nodes) == "repeat 2:\n  interval: 1km @ Z2\n  rest: 1min"


def test_invalid_structure_raises():
    with pytest.raises(DSLError):
        wm.from_dsl("repeat 3:")


def test_tree_operations():
    nodes = wm.from_dsl(TEXT)
    rep, cool = nodes[1], nodes[2]
    assert wm.move(nodes, cool.id, -1)
    assert nodes[1] is cool
    wm.move(nodes, cool.id, +1)
    assert wm.indent(nodes, cool.id)          # dentro la ripetizione
    assert rep.children[-1] is cool
    assert wm.parent_of(nodes, cool.id) is rep
    assert wm.outdent(nodes, cool.id)         # e di nuovo fuori, subito dopo
    assert nodes[2] is cool and wm.parent_of(nodes, cool.id) is None
    new = wm.Step("interval", "5min", "Z2")
    wm.insert_after(nodes, rep.id, new, inside=True)
    assert rep.children[-1] is new
    wm.remove(nodes, new.id)
    assert wm.get(nodes, new.id) is None


def test_estimates():
    segs = wm.segments(wm.from_dsl(TEXT), "running", PARAMS)
    assert len(segs) == 3 + 5 * 2 - 1           # warmup + 5×(2) + cooldown
    secs, meters, _ = wm.totals(segs)
    assert 55 * 60 < secs < 65 * 60          # 15' + 5×(~4'55" + 2') + 10'
    assert 9000 < meters < 11000
    inten = {s.type: s.intensity for s in segs}
    assert inten["interval"] > inten["warmup"] > inten["recovery"]


def test_rest_has_no_distance():
    segs = wm.segments(wm.from_dsl("rest: 2min"), "running", PARAMS)
    assert segs[0].meters == 0 and segs[0].seconds == 120


def test_format():
    assert wm.fmt_duration(3900) == "1h05'"
    assert wm.fmt_duration(270) == "4'30\""
    assert wm.fmt_duration(2400) == "40'"
    assert wm.fmt_km(8450) == "8,4 km" or wm.fmt_km(8450) == "8,5 km"
    assert wm.fmt_km(400) == "400 m"


def test_presets_are_valid_dsl():
    from dsl_parser import build_garmin_workout_from_excel_row
    for sport, items in wm.PRESETS.items():
        for name, factory in items:
            text = wm.to_dsl(factory())
            build_garmin_workout_from_excel_row(pd.Series({"Description": name, "Sport": sport}), text, PARAMS)


def test_move_to_drag_and_drop():
    nodes = wm.from_dsl(TEXT)
    warm, rep, cool = nodes
    work, rec = rep.children
    assert wm.move_to(nodes, cool.id, warm.id, "before")            # defaticamento in cima
    assert nodes[0] is cool
    assert wm.move_to(nodes, cool.id, rep.id, "inside")             # dentro la ripetizione
    assert rep.children[-1] is cool
    assert wm.move_to(nodes, work.id, rec.id, "after")              # riordino dentro la ripetizione
    assert rep.children[:2] == [rec, work]
    assert wm.move_to(nodes, rec.id, rep.id, "after")               # fuori dalla ripetizione
    assert nodes[-1] is rec and wm.parent_of(nodes, rec.id) is None
    assert not wm.move_to(nodes, rep.id, work.id, "before")         # ripetizione dentro se stessa: no
    assert not wm.move_to(nodes, warm.id, rec.id, "inside")         # dentro uno step: no
    assert wm.to_dsl(nodes).count("\n") == 4                        # nessuno step perso
