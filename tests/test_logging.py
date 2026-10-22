"""Файловый лог: config/app.log рядом с faculties.json, запасной путь при отсутствии прав."""
from __future__ import annotations

import logging

import pytest

from conftest import ROOT
from schedule_app import logging_setup
from schedule_app.logging_setup import LOGGER_NAME, setup_file_logging


@pytest.fixture(autouse=True)
def clean_logger():
    """Каждый тест начинает с логгера без хэндлеров."""
    logger = logging.getLogger(LOGGER_NAME)

    def reset():
        for h in list(logger.handlers):
            h.close()
            logger.removeHandler(h)
    reset()
    yield
    reset()


def test_default_log_path_is_next_to_config():
    assert logging_setup.CONFIG_PATH.parent / "app.log" == ROOT / "config" / "app.log"


def test_log_written_to_file(tmp_path):
    log_file = tmp_path / "config" / "app.log"
    logger, path = setup_file_logging(log_file)
    logger.error("проверка")
    for h in logger.handlers:
        h.flush()
    assert path == log_file
    text = log_file.read_text(encoding="utf-8")
    assert "Запуск приложения" in text
    assert "[ERROR] проверка" in text


def test_log_falls_back_when_folder_not_writable(tmp_path):
    blocker = tmp_path / "config"
    blocker.write_text("это файл, а не папка")  # mkdir внутри него упадёт с OSError
    fallback = tmp_path / "home" / "app.log"
    _, path = setup_file_logging(blocker / "app.log", fallback=fallback)
    assert path == fallback
    assert fallback.exists()


def test_no_duplicate_handlers(tmp_path):
    setup_file_logging(tmp_path / "a.log")
    setup_file_logging(tmp_path / "a.log")
    assert len(logging.getLogger(LOGGER_NAME).handlers) == 1
