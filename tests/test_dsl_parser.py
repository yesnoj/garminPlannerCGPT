"""Test del parser DSL: python -m pytest -q"""
from pathlib import Path

import pandas as pd
import pytest

from dsl_parser import (
    DSLError, build_garmin_workout_from_excel_row, expand_repeat_lines,
    parse_duration_strict, parse_target_strict, validate_workout_rows,
)

ROOT = Path(__file__).resolve().parents[1]
ROW = pd.Series({"Description": "Test", "Sport": "Running", "Week": 1.0, "Session": 2.0})
PARAMS = pd.DataFrame([
    {"Key": "Z1", "Metric": "pace", "Expression": "7:00-7:30"},
    {"Key": "Z2", "Metric": "pace", "Expression": "6:00"},
    {"Key": "Easy_Range", "Metric": "pace", "Expression": "6:30-7:00"},
    {"Key": "Swim_Z2", "Metric": "pace", "Expression": "1:50"},
    {"Key": "Cadence_Easy", "Metric": "cadence", "Expression": "85-90"},
    {"Key": "HR_Z2", "Metric": "hr", "Expression": "70-80%"},
    {"Key": "HR_max", "Metric": "hr", "Expression": "180"},
    {"Key": "pace_tolerance", "Metric": "pace", "Expression": "5"},
])


def steps(text, params=PARAMS):
    return build_garmin_workout_from_excel_row(ROW, text, params)["workoutSegments"][0]["workoutSteps"]


# ---------- durate ----------

@pytest.mark.parametrize("text, expected", [
    ("10min", ("time", 600.0, None)),
    ("1.5min", ("time", 90.0, None)),
    ("10'", ("time", 600.0, None)),
    ("40′", ("time", 2400.0, None)),
    ("30sec", ("time", 30.0, None)),
    ("90s", ("time", 90.0, None)),
    ("1h", ("time", 3600.0, None)),
    ("1:30", ("time", 90.0, None)),
    ("1:05:00", ("time", 3900.0, None)),
    ("400m", ("distance", 400.0, "meter")),
    ("10m", ("distance", 10.0, "meter")),        # m = METRI
    ("1.5km", ("distance", 1500.0, "kilometer")),
    ("21.1km", ("distance", 21100.0, "kilometer")),
    ("lap-button", ("lap.button", None, None)),
])
def test_durations(text, expected):
    assert parse_duration_strict(text) == expected


@pytest.mark.parametrize("text", ["", "abc", "10 miles", "1:75", "5x"])
def test_bad_durations(text):
    with pytest.raises(ValueError):
        parse_duration_strict(text)


# ---------- target ----------

def test_custom_keys_are_case_insensitive():
    for key in ("Easy_Range", "easy_range", "EASY_RANGE"):
        assert parse_target_strict(key, PARAMS)["targetType"]["workoutTargetTypeKey"] == "pace.zone"
    assert parse_target_strict("Swim_Z2", PARAMS)["targetType"]["workoutTargetTypeKey"] == "pace.zone"
    assert parse_target_strict("Cadence_Easy", PARAMS)["targetType"]["workoutTargetTypeKey"] == "cadence"


def test_pace_single_uses_tolerance():
    t = parse_target_strict("5:00", PARAMS)
    assert t["targetValueOne"] == pytest.approx(1000 / 305)
    assert t["targetValueTwo"] == pytest.approx(1000 / 295)


def test_hr_zone_percent_and_reversed_range():
    t = parse_target_strict("HR_Z2", PARAMS)
    assert (t["targetValueOne"], t["targetValueTwo"]) == pytest.approx((126, 144))
    t = parse_target_strict("160-140", PARAMS)
    assert (t["targetValueOne"], t["targetValueTwo"]) == (140, 160)


def test_open_means_no_target():
    assert parse_target_strict("open", PARAMS)["targetType"]["workoutTargetTypeKey"] == "no.target"


@pytest.mark.parametrize("target", ["Z7", "Z0", "4:60", "xyz", "threshold", "10"])
def test_bad_targets(target):
    with pytest.raises(ValueError):
        parse_target_strict(target, PARAMS)


# ---------- struttura / repeat ----------

@pytest.mark.parametrize("indent", ["  ", "    ", "\t"])
def test_repeat_any_consistent_indentation(indent):
    s = steps(f"warmup: 10min @ Z2\nrepeat 5:\n{indent}interval: 1km @ 5:00\n{indent}recovery: 2min @ Z1\ncooldown: 5min @ Z1")
    assert [x["type"] for x in s] == ["ExecutableStepDTO", "RepeatGroupDTO", "ExecutableStepDTO"]
    assert s[1]["numberOfIterations"] == 5
    assert len(s[1]["workoutSteps"]) == 2


def test_nested_repeat():
    s = steps("repeat 2:\n  repeat 6:\n    interval: 1min @ 4:30\n    recovery: 1min @ Z1\n  rest: 3min")
    outer = s[0]
    assert outer["numberOfIterations"] == 2
    assert outer["workoutSteps"][0]["numberOfIterations"] == 6
    assert outer["workoutSteps"][1]["stepType"]["stepTypeKey"] == "rest"


def test_expand_repeat_lines_public_api():
    assert expand_repeat_lines("repeat 2:\n  interval: 1km @ Z2") == [{"repeat": 2, "steps": ["interval: 1km @ Z2"]}]


@pytest.mark.parametrize("text, fragment", [
    ("warmup 15min Z2; 5x(run 5min Z4)", "formato step non valido"),
    ("repeat 4:", "senza step"),
    ("repeat 0:\n  interval: 1km", "ripetizioni non valido"),
    ("interval: 1km @ Z2\n  interval: 1km @ Z2", "indentazione inattesa"),
    ("sprint: 1km @ Z2", "sconosciuto"),
    ("interval: 1km @", "senza target"),
    ("nan", "vuota"),
])
def test_errors_are_reported_with_reason(text, fragment):
    with pytest.raises(DSLError) as exc:
        steps(text)
    assert fragment in str(exc.value)


def test_error_reports_line_number():
    with pytest.raises(DSLError) as exc:
        steps("warmup: 10min @ Z2\nrepeat 3:\n  interval: 1km @ Z9\n  recovery: 2min @ Z1")
    assert exc.value.line_no == 3


# ---------- workout completo ----------

def test_workout_name_prefix_uses_integers():
    w = build_garmin_workout_from_excel_row(ROW, "interval: 5km @ Z2", PARAMS, use_prefix=True)
    assert w["workoutName"] == "W1S2 - Test"


def test_empty_description_does_not_become_nan():
    w = build_garmin_workout_from_excel_row(pd.Series({"Description": float("nan")}), "interval: 5km", PARAMS)
    assert w["workoutName"] == "Workout"


@pytest.mark.parametrize("name", ["esempio_running.xlsx", "esempio_multisport.xlsx"])
def test_example_files_are_fully_valid(name):
    x = pd.read_excel(ROOT / "examples" / name, sheet_name=None)
    w, p = x["Workouts"], x["Parameters"]
    assert validate_workout_rows(w, list(range(len(w))), p, require_date=True) == []
    for _, r in x["Esempi DSL"].iterrows():
        if r["Categoria"] == "NOTE":
            continue
        build_garmin_workout_from_excel_row(pd.Series({"Description": "x"}), r["Esempio"], p)


def test_every_step_with_target_has_a_target():
    """Una riga con '@ qualcosa' non deve mai finire su Garmin senza target."""
    x = pd.read_excel(ROOT / "examples" / "esempio_multisport.xlsx", sheet_name=None)
    w, p = x["Workouts"], x["Parameters"]

    def walk(items):
        for it in items:
            if it["type"] == "RepeatGroupDTO":
                yield from walk(it["workoutSteps"])
            else:
                yield it

    for _, r in w.iterrows():
        wo = build_garmin_workout_from_excel_row(r, r["Steps"], p)
        for st in walk(wo["workoutSegments"][0]["workoutSteps"]):
            if "@" in st["description"] and "open" not in st["description"].lower():
                assert st["targetType"]["workoutTargetTypeKey"] != "no.target", st["description"]


def test_validate_reports_row_and_date():
    df = pd.DataFrame([
        {"Week": 1, "Session": 1, "Description": "ok", "Date": "2026-10-05", "Steps": "interval: 5km @ Z2"},
        {"Week": 1, "Session": 2, "Description": "ko", "Date": None, "Steps": "interval: 5 km @ Z9"},
    ])
    errs = validate_workout_rows(df, [0, 1], PARAMS, require_date=True)
    assert len(errs) == 2
    assert all("Riga Excel 3" in e for e in errs)
