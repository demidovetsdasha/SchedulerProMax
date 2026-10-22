"""Конфиг: где приложение ищет файлы (VS Code / exe) и разбор faculties.json."""
from __future__ import annotations

import json
import sys

import pytest

from conftest import ROOT
from schedule_app import config
from schedule_app.config import app_dir, load_config, normalize


# ------------------------------------------------------------------ app_dir

def test_app_dir_from_sources_is_project_root():
    assert app_dir() == ROOT
    assert config.CONFIG_PATH == ROOT / "config" / "faculties.json"


def test_app_dir_frozen_exe_is_exe_folder(tmp_path, monkeypatch):
    exe = tmp_path / "ScheduleParser" / "ScheduleParser.exe"
    exe.parent.mkdir()
    exe.touch()
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(exe))
    assert app_dir() == exe.parent.resolve()


def test_app_dir_frozen_macos_app_uses_folder_next_to_app(tmp_path, monkeypatch):
    exe = tmp_path / "ScheduleParser.app" / "Contents" / "MacOS" / "ScheduleParser"
    exe.parent.mkdir(parents=True)
    exe.touch()
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(exe))
    assert app_dir() == tmp_path.resolve()


def test_app_dir_frozen_macos_prefers_config_inside_bundle(tmp_path, monkeypatch):
    exe = tmp_path / "ScheduleParser.app" / "Contents" / "MacOS" / "ScheduleParser"
    (exe.parent / "config").mkdir(parents=True)
    exe.touch()
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(exe))
    assert app_dir() == exe.parent.resolve()


# -------------------------------------------------------------- load_config

def test_project_config_loads(cfg):
    assert cfg.template.exists(), "шаблон из конфига не найден"
    assert cfg.max_columns_per_page > 0
    assert cfg.table_font_min <= cfg.table_font_max
    codes = [f.code for f in cfg.faculties]
    assert len(codes) == len(set(codes)), "коды факультетов повторяются"
    for fac in cfg.faculties:
        assert fac.name and fac.header_name and fac.deputy_dean, fac.code
        assert fac.aliases, f"у {fac.code} нет aliases — факультет не определится автоматически"


def test_load_config_defaults(tmp_path):
    path = tmp_path / "faculties.json"
    path.write_text(json.dumps({"faculties": []}), encoding="utf-8")
    cfg = load_config(path)
    assert cfg.max_columns_per_page == 8
    assert cfg.week1_monday is None
    assert cfg.faculties == []
    assert cfg.path == path


def test_faculty_by_code(cfg):
    assert cfg.faculty_by_code("ФКЯиК").name == "Факультет китайского языка и культуры"
    assert cfg.faculty_by_code("нет такого") is None


# ----------------------------------------------------------- detect_faculty

@pytest.mark.parametrize("text, code", [
    ("Расписание_занятий_ФКЯиК,_специалист,_очная_дневная,_2_й_курс,_9.xlsx", "ФКЯиК"),
    ("Расписание_занятий_ФРЯ,_специалист,_очная_дневная,_5_й_курс,_8_я.xlsx", "ФРЯ"),
    ("Расписание_занятий_перев_ф_т,_специалист,_очная_дневная,_4_й_курс.xlsx", "ПФ"),
    ("Расписание_занятий_ф_т_ан_яз,_специалист,_очная_дневная,_2_й_курс.xlsx", "ФАЯ"),
    ("Расписание_занятий_ф_т_нем_яз,_специалист,_очная_дневная,_3_й_курс.xlsx", "ФНЯ"),
    ("ф_т_ан_яз_спе_дн_2к_зимний_10_11_2025_16_11_2025.xlsx", "ФАЯ"),
    ("Расписание занятий: ф-т межкультурных коммуникаций, 3-й курс", "ФМК"),
])
def test_detect_faculty(cfg, text, code):
    fac = cfg.detect_faculty(text)
    assert fac is not None and fac.code == code


def test_detect_faculty_unknown(cfg):
    assert cfg.detect_faculty("Расписание чего-то непонятного") is None


def test_normalize():
    assert normalize("  Ф_Т.Ан-Яз  ") == "ф т ан яз"
    assert normalize("Ёлка") == "елка"
