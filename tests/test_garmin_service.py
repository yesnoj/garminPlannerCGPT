import os
import stat

import pytest


@pytest.fixture
def service(tmp_path, monkeypatch):
    monkeypatch.setenv("GARMINPLANNER_TOKEN_DIR", str(tmp_path / "tokens"))
    import importlib
    import garmin_service
    importlib.reload(garmin_service)
    return garmin_service


def test_token_dir_is_outside_project(service):
    project = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    assert not str(service.TOKEN_STORE_DIR).startswith(project)


@pytest.mark.skipif(os.name == "nt", reason="permessi POSIX")
def test_tokens_saved_private(service, monkeypatch):
    def fake_save(path):
        for n in service.TOKEN_FILES:
            with open(os.path.join(path, n), "w") as f:
                f.write("{}")
    monkeypatch.setattr(service.garth, "save", fake_save)
    service._save_tokens()
    for n in service.TOKEN_FILES:
        mode = stat.S_IMODE(os.stat(service.TOKEN_STORE_DIR / n).st_mode)
        assert mode == 0o600
    assert stat.S_IMODE(os.stat(service.TOKEN_STORE_DIR).st_mode) == 0o700


def test_legacy_tokens_are_migrated(service, tmp_path, monkeypatch):
    legacy = tmp_path / "old" / "garminconnect"
    legacy.mkdir(parents=True)
    for n in service.TOKEN_FILES:
        (legacy / n).write_text("{}")
    monkeypatch.setattr(service, "LEGACY_TOKEN_DIRS", (legacy,))
    assert service.GarminService.has_saved_session()
    assert service._migrate_legacy_tokens() == legacy
    assert all((service.TOKEN_STORE_DIR / n).exists() for n in service.TOKEN_FILES)


def test_has_saved_session_ignores_stray_files(service, monkeypatch):
    monkeypatch.setattr(service, "LEGACY_TOKEN_DIRS", ())
    service.TOKEN_STORE_DIR.mkdir(parents=True)
    (service.TOKEN_STORE_DIR / ".DS_Store").write_text("x")
    assert not service.GarminService.has_saved_session()
