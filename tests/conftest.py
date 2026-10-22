"""Общие фикстуры: пути к примерам расписаний из testing/ и загруженный конфиг."""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

SAMPLES_DIR = ROOT / "testing" / "тестирование"
SAMPLE_FILES = sorted(SAMPLES_DIR.glob("*.xls*"))


def sample(part: str) -> Path:
    """Пример расписания, в имени которого есть `part`."""
    matches = [p for p in SAMPLE_FILES if part in p.name]
    assert len(matches) == 1, f"ожидался один файл с «{part}», найдено: {matches}"
    return matches[0]


@pytest.fixture(scope="session")
def cfg():
    from schedule_app.config import load_config
    return load_config()


@pytest.fixture(scope="session")
def load_sheet():
    """Кэширует разбор файлов: openpyxl читает их заметно долго."""
    from schedule_app.excel_parser import load_excel
    cache = {}

    def _load(path: Path):
        if path not in cache:
            cache[path] = load_excel(path).sheets[0]
        return cache[path]
    return _load
