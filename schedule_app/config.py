"""Загрузка конфигурации факультетов (config/faculties.json)."""
from __future__ import annotations

import datetime as dt
import json
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path


def app_dir() -> Path:
    """Папка приложения: рядом с exe / .app (PyInstaller) или корень проекта."""
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        # macOS: …/ScheduleParser.app/Contents/MacOS — ищем конфиг и рядом с .app
        if exe_dir.parent.parent.suffix == ".app" and not (exe_dir / "config").exists():
            return exe_dir.parent.parent.parent
        return exe_dir
    return Path(__file__).resolve().parent.parent


def resource_dir() -> Path:
    """Папка со встроенными ресурсами (для PyInstaller --onefile это _MEIPASS)."""
    return Path(getattr(sys, "_MEIPASS", app_dir()))


CONFIG_PATH = app_dir() / "config" / "faculties.json"


@dataclass
class Faculty:
    code: str
    name: str
    header_name: str
    deputy_dean: str
    aliases: list[str] = field(default_factory=list)


@dataclass
class AppConfig:
    template: Path
    max_columns_per_page: int
    week1_monday: dt.date | None
    table_font_max: float
    table_font_min: float
    self_study_text: str
    head_ckop_title: str
    head_ckop_name: str
    faculties: list[Faculty]
    path: Path

    def faculty_by_code(self, code: str) -> Faculty | None:
        return next((f for f in self.faculties if f.code == code), None)

    def detect_faculty(self, *texts: str) -> Faculty | None:
        """Определяет факультет по заголовку листа / имени файла."""
        haystack = " ".join(normalize(t) for t in texts if t)
        for fac in self.faculties:
            for alias in fac.aliases:
                if re.search(r"(?<!\w)" + re.escape(normalize(alias)), haystack):
                    return fac
        return None


def normalize(text: str) -> str:
    text = unicodedata.normalize("NFC", text).lower().replace("ё", "е")
    text = re.sub(r"[._\-,:;()]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _resolve(path_str: str) -> Path:
    path = Path(path_str)
    if path.is_absolute():
        return path
    for base in (app_dir(), resource_dir()):
        if (base / path).exists():
            return base / path
    return app_dir() / path


def load_config(path: Path = CONFIG_PATH) -> AppConfig:
    if not path.exists():
        bundled = resource_dir() / "config" / "faculties.json"
        if bundled.exists():
            path = bundled
    with open(path, encoding="utf-8") as f:
        raw = json.load(f)
    head = raw.get("head_ckop", {})
    return AppConfig(
        template=_resolve(raw.get("template", "templates/template.docx")),
        max_columns_per_page=int(raw.get("max_columns_per_page", 8)),
        week1_monday=dt.date.fromisoformat(raw["week1_monday"]) if raw.get("week1_monday") else None,
        table_font_max=float(raw.get("table_font_max_pt", 8.5)),
        table_font_min=float(raw.get("table_font_min_pt", 4)),
        self_study_text=raw.get("self_study_text", "День самостоятельной работы (дср), ауд"),
        head_ckop_title=head.get("title", "НАЧАЛЬНИК ЦКОП"),
        head_ckop_name=head.get("name", "В.Д.СИНЯК"),
        faculties=[
            Faculty(
                code=f["code"],
                name=f["name"],
                header_name=f.get("header_name", f["name"]),
                deputy_dean=f.get("deputy_dean", ""),
                aliases=f.get("aliases", []),
            )
            for f in raw.get("faculties", [])
        ],
        path=path,
    )
