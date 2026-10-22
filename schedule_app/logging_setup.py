"""Логирование в файл config/app.log рядом с faculties.json."""
from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

from .config import CONFIG_PATH

LOGGER_NAME = "schedule_app"
FALLBACK_LOG_FILE = Path.home() / ".schedule_parser" / "app.log"


def _handler(log_file: Path) -> RotatingFileHandler:
    log_file.parent.mkdir(parents=True, exist_ok=True)
    return RotatingFileHandler(log_file, maxBytes=2_000_000, backupCount=3, encoding="utf-8")


def setup_file_logging(log_file: Path | None = None,
                       fallback: Path = FALLBACK_LOG_FILE) -> tuple[logging.Logger, Path]:
    """Подключает файловый лог (с ротацией) и возвращает логгер и путь к файлу."""
    log_file = log_file or CONFIG_PATH.parent / "app.log"
    try:
        handler = _handler(log_file)
    except OSError:
        # Папка приложения недоступна на запись (например, Program Files)
        log_file = fallback
        handler = _handler(log_file)
    handler.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    ))

    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(logging.DEBUG)
    # Не дублируем хэндлеры при повторном создании окна
    if not logger.handlers:
        logger.addHandler(handler)
    else:
        handler.close()

    logger.info(f"=== Запуск приложения (файл: {log_file}) ===")
    return logger, log_file
