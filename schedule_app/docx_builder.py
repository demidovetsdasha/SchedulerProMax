"""Формирование .docx по шаблону.

Из шаблона берётся первый «блок страницы»: таблица-шапка (РАСПИСАНИЕ ЗАНЯТИЙ /
УТВЕРЖДАЮ), разделительные абзацы, таблица расписания и подписи. Для каждой
страницы (порции групп) блок копируется, шапка и подписи заполняются, а таблица
расписания строится заново в стиле шаблона.
"""
from __future__ import annotations

import copy
import datetime as dt
import re
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

from . import tahoma_metrics
from .excel_parser import DAY_NAMES, Block, Column, Day, MergeRange, SheetSchedule

FONT = "Tahoma"
FONT_SIZE = 17            # половины пункта (8,5 pt), как в шаблоне
MIN_ROW_H = 129           # минимальная высота строки таблицы (twips), как в шаблоне
# Высота, которую на странице занимают шапка (РАСПИСАНИЕ / УТВЕРЖДАЮ), разделители
# и подписи; остальное — под таблицу расписания (twips, по замеру страницы шаблона).
PAGE_OVERHEAD = 2850
HEADER_FILL = "D3D3D3"
THIN, THICK = 4, 12       # толщина границ (восьмые пункта)
TABLE_WIDTH = 21724       # ширина таблицы расписания в шаблоне (twips)
DAY_COL_W, TIME_COL_W = 580, 758
STUDY_DAYS = DAY_NAMES[:6]          # понедельник – суббота
DSR_TEXT = "День самостоятельной работы (дср), ауд"


@dataclass
class DocumentParams:
    course: int
    faculty_header: str           # "факультета межкультурных коммуникаций"
    date_start: dt.date
    date_end: dt.date
    head_ckop_title: str
    head_ckop_name: str
    deputy_dean: str
    max_columns_per_page: int = 8
    academic_year: str | None = None      # "2026-2027"; по умолчанию — по дате начала недели
    approval_year: int | None = None      # год в «УТВЕРЖДАЮ»; по умолчанию — текущий
    table_font_max: float = 8.5           # кегль таблицы, pt (подбирается ≤ max, ≥ min)
    table_font_min: float = 4.0
    self_study_text: str = DSR_TEXT       # что ставить в учебный день без занятий


def academic_year(d: dt.date) -> str:
    start = d.year if d.month >= 8 else d.year - 1
    return f"{start}-{start + 1}"


# ------------------------------------------------------------------ pagination

def paginate(columns: list[Column], max_cols: int) -> list[list[Column]]:
    """Делит столбцы на страницы, не разрывая подгруппы одной группы."""
    by_group: list[list[Column]] = []
    for col in columns:
        if by_group and by_group[-1][0].group == col.group:
            by_group[-1].append(col)
        else:
            by_group.append([col])
    pages: list[list[Column]] = []
    for grp in by_group:
        if pages and len(pages[-1]) + len(grp) <= max_cols:
            pages[-1].extend(grp)
        else:
            pages.append(list(grp))
    return pages


# --------------------------------------------------------------- XML helpers

def _set_paragraph_text(p, text: str) -> None:
    """Заменяет текст абзаца, сохраняя формат первого текстового прогона."""
    runs = p.findall(qn("w:r"))
    proto = next((r for r in runs if r.find(qn("w:t")) is not None), runs[0] if runs else None)
    for r in runs:
        p.remove(r)
    run = copy.deepcopy(proto) if proto is not None else parse_xml(f"<w:r {nsdecls('w')}/>")
    for child in list(run):
        if child.tag != qn("w:rPr"):
            run.remove(child)
    t = parse_xml(f'<w:t {nsdecls("w")} xml:space="preserve"></w:t>')
    t.text = text
    run.append(t)
    p.append(run)


def _para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def _borders(top: int, left: int, bottom: int, right: int) -> str:
    def b(side, size):
        return f'<w:{side} w:val="single" w:sz="{size}" w:space="0" w:color="auto"/>'
    return f"<w:tcBorders>{b('top', top)}{b('left', left)}{b('bottom', bottom)}{b('right', right)}</w:tcBorders>"


def _cell(width: int, text: str = "", *, bold=False, fill="FFFFFF", borders=(THIN,) * 4,
          span: int = 1, vmerge: str | None = None, vertical=False, size: int = FONT_SIZE):
    span_xml = f'<w:gridSpan w:val="{span}"/>' if span > 1 else ""
    vm_xml = ""
    if vmerge == "restart":
        vm_xml = '<w:vMerge w:val="restart"/>'
    elif vmerge == "continue":
        vm_xml = "<w:vMerge/>"
    dir_xml = '<w:textDirection w:val="btLr"/>' if vertical else ""
    rpr = (f'<w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/>'
           + ("<w:b/><w:bCs/>" if bold else "")
           + f'<w:color w:val="000000"/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>')
    runs = ""
    for i, line in enumerate(text.split("\n") if text else []):
        br = "<w:br/>" if i else ""
        runs += f'<w:r>{rpr}{br}<w:t xml:space="preserve">{_esc(line)}</w:t></w:r>'
    xml = (
        f"<w:tc {nsdecls('w')}><w:tcPr>"
        f'<w:tcW w:w="{width}" w:type="dxa"/>{span_xml}{vm_xml}{_borders(*borders)}'
        f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>{dir_xml}<w:vAlign w:val="center"/>'
        f'</w:tcPr><w:p><w:pPr><w:jc w:val="center"/>{rpr}</w:pPr>{runs}</w:p></w:tc>'
    )
    return parse_xml(xml)


def _esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _row(height: int, header=False):
    extra = "<w:tblHeader/>" if header else "<w:cantSplit/>"
    return parse_xml(f'<w:tr {nsdecls("w")}><w:trPr>{extra}<w:trHeight w:val="{height}"/></w:trPr></w:tr>')


# ------------------------------------------------------------ table building

@dataclass
class _OutRow:
    day: Day
    block: Block
    row: int                 # номер строки Excel
    first_in_day: bool = False
    first_in_block: bool = False
    block_size: int = 1      # сколько строк этой пары попало в Word


@dataclass
class _CellSpec:
    grid_col: int            # первый столбец сетки
    width: int               # twips
    text: str
    borders: tuple
    span: int = 1
    vmerge: str | None = None
    bold: bool = False
    vertical: bool = False
    fill: str = "FFFFFF"
    fixed_size: float | None = None   # кегль, который не подбирается (строка заголовка)


def _key(sheet: SheetSchedule, row: int, col: Column) -> MergeRange | None:
    """Диапазон объединения Excel для ячейки (None — обычная ячейка)."""
    lesson = sheet.lesson(row, col.index)
    return lesson.merge if lesson else None


def _is_blank(sheet: SheetSchedule, row: int, col: Column) -> bool:
    lesson = sheet.lesson(row, col.index)
    return lesson is None or (not lesson.text and lesson.merge is None)


def _study_days(sheet: SheetSchedule) -> list[Day]:
    """Пн–Сб в порядке недели; дни, которых нет в Excel, добавляются пустыми."""
    by_name = {d.name: d for d in sheet.days}
    first_time = next((b.time for d in sheet.days for b in d.blocks if b.time), "")
    days = [by_name.get(name) or Day(name, [Block(pair="", time=first_time, rows=[-1])])
            for name in STUDY_DAYS]
    return days + [d for d in sheet.days if d.name not in STUDY_DAYS]


def _is_self_study(text: str) -> bool:
    return text.lower().replace("ё", "е").startswith("день самостоятельной работы")


def _self_study_columns(sheet: SheetSchedule, day: Day, cols: list[Column]) -> set[int]:
    """Подгруппы (Column.index), у которых в этот учебный день нет занятий
    (пусто или в Excel стоит только «День самостоятельной работы»)."""
    if day.name not in STUDY_DAYS:
        return set()
    rows = [r for b in day.blocks for r in b.rows]
    result = set()
    for c in cols:
        texts = [les.text for r in rows if (les := sheet.lesson(r, c.index)) and les.text]
        if all(_is_self_study(t) for t in texts):
            result.add(c.index)
    return result


def _layout(sheet: SheetSchedule, cols: list[Column], days: list[Day]) -> list[_OutRow]:
    """Строки Word = строки Excel (геометрия сохраняется).

    Пара пропускается, если у групп страницы в ней нет занятий. Внутри пары
    строка опускается, только если ничего не добавляет для групп страницы —
    все её ячейки пусты или продолжают вертикальное объединение строки выше
    (например, дисциплины по выбору есть лишь у групп с другой страницы).
    """
    out: list[_OutRow] = []
    for day in days:
        day_rows: list[_OutRow] = []
        for block in day.blocks:
            if all(_is_blank(sheet, r, c) for r in block.rows for c in cols):
                continue
            kept: list[int] = []
            for r in block.rows:
                prev = kept[-1] if kept else None
                redundant = all(
                    _is_blank(sheet, r, c)
                    or (prev is not None and _key(sheet, r, c) is not None
                        and _key(sheet, r, c) == _key(sheet, prev, c))
                    for c in cols
                )
                if not redundant:
                    kept.append(r)
            kept = kept or block.rows[:1]
            for i, r in enumerate(kept):
                day_rows.append(_OutRow(day, block, r, first_in_block=i == 0, block_size=len(kept)))
        if not day_rows and day.blocks and day.name != "Воскресенье":
            # день без занятий — одна пустая строка, чтобы день был виден
            block = day.blocks[0]
            day_rows.append(_OutRow(day, block, block.rows[0], first_in_block=True))
        if day_rows:
            day_rows[0].first_in_day = True
            out.extend(day_rows)
    return out


def _table_specs(sheet: SheetSchedule, cols: list[Column],
                 self_study_text: str = DSR_TEXT) -> tuple[list[int], list[list[_CellSpec]]]:
    """Логическая раскладка таблицы: ширины столбцов и ячейки по строкам (0-я — заголовок)."""
    n = len(cols)
    data_w = (TABLE_WIDTH - DAY_COL_W - TIME_COL_W) // n
    widths = [DAY_COL_W, TIME_COL_W] + [data_w] * n

    header = []
    heads = ["День недели", "Время пары"] + [c.name for c in cols]
    for i, (w, text) in enumerate(zip(widths, heads)):
        lft = THICK if i == 0 else THIN
        rgt = THICK if i == n + 1 else THIN
        header.append(_CellSpec(i, w, text, (THICK, lft, THIN, rgt), fill=HEADER_FILL,
                                fixed_size=FONT_SIZE / 2))
    table = [header]

    days = _study_days(sheet)
    rows = _layout(sheet, cols, days)
    # В учебный день без занятий у подгруппы — «День самостоятельной работы».
    # Общий ключ дня объединяет такие ячейки соседних подгрупп и все строки дня в одну.
    self_study = {d.name: _self_study_columns(sheet, d, cols) for d in days}

    def cell(orow: _OutRow, col: Column) -> tuple[object, str]:
        if col.index in self_study[orow.day.name]:
            return ("self-study", orow.day.name), self_study_text
        lesson = sheet.lesson(orow.row, col.index)
        return (lesson.merge, lesson.text) if lesson else (None, "")

    for i, orow in enumerate(rows):
        prev = rows[i - 1] if i > 0 and not orow.first_in_day else None
        nxt = rows[i + 1] if i + 1 < len(rows) and not rows[i + 1].first_in_day else None
        top = THICK if orow.first_in_day else THIN
        bottom = THICK if i == len(rows) - 1 else THIN
        cells = []

        # День недели — объединён на все строки дня
        day_vm = "continue" if prev is not None else ("restart" if nxt is not None else None)
        cells.append(_CellSpec(0, DAY_COL_W, orow.day.name if orow.first_in_day else "",
                               (top, THICK, bottom, THIN), vmerge=day_vm, bold=True, vertical=True))

        # Время пары — объединено на все строки пары, как в Excel
        time_vm = None
        if orow.block_size > 1:
            time_vm = "restart" if orow.first_in_block else "continue"
        cells.append(_CellSpec(1, TIME_COL_W, orow.block.time if orow.first_in_block else "",
                               (top, THIN, bottom, THIN), vmerge=time_vm, bold=True))

        # Ячейки групп — те же объединения, что и в Excel
        pc = 0
        while pc < n:
            key, text = cell(orow, cols[pc])
            span = 1
            if key is not None:
                while pc + span < n and cell(orow, cols[pc + span])[0] == key:
                    span += 1
            cont = prev is not None and key is not None and cell(prev, cols[pc])[0] == key
            down = nxt is not None and key is not None and cell(nxt, cols[pc])[0] == key
            if cont:
                text = ""
            rgt = THICK if pc + span == n else THIN
            cells.append(_CellSpec(pc + 2, data_w * span, text, (top, THIN, bottom, rgt), span=span,
                                   vmerge="continue" if cont else ("restart" if down else None)))
            pc += span
        table.append(cells)
    return widths, table


# ------------------------------------------------------- подбор размера шрифта

# Коэффициенты модели высоты подобраны по замерам отрендеренных страниц.
LINE_FACTOR = 1.30          # высота строки Tahoma в долях кегля
CHAR_SCALE = 1.0            # поправка к ширине символов
ROW_PAD = 30.0              # постоянная добавка к высоте каждой строки таблицы (twips)
WRAP_FACTOR = 1.0           # доля ширины ячейки, доступная тексту
FIT_SAFETY = 1.03           # запас на погрешность модели (до ~3% по замерам)
_WIDE = set("шщжмюыфШЩЖМЮЫФmwMW")


def _char_em(ch: str) -> float:
    """Приблизительная ширина символа Tahoma в долях кегля."""
    if ch == " ":
        return 0.313
    if ch in ".,:;'!|":
        return 0.33
    if ch in "()[]-/":
        return 0.38
    if ch.isdigit():
        return 0.546
    if ch in _WIDE:
        return 0.76
    if ch.isupper():
        return 0.66
    return 0.55


def _text_width(text: str, size_pt: float, bold: bool) -> float:
    """Ширина текста в twips по метрикам Tahoma (неизвестные символы — приблизительно)."""
    table = tahoma_metrics.BOLD if bold else tahoma_metrics.REGULAR
    k = 1.17 if bold else 1.0
    em = sum(table.get(c) or _char_em(c) * k for c in text)
    return em * size_pt * CHAR_SCALE * 20


def _line_count(text: str, width_tw: float, size_pt: float, bold: bool) -> int:
    """Число строк после переноса по словам (длинные слова рвутся)."""
    if not text:
        return 1
    avail = width_tw * WRAP_FACTOR
    space = _text_width(" ", size_pt, bold)
    total = 0
    for para in text.split("\n"):
        lines, cur = 1, 0.0
        for word in para.split(" "):
            w = _text_width(word, size_pt, bold)
            if cur and cur + space + w <= avail:
                cur += space + w
                continue
            if cur:
                lines += 1
            while w > avail:            # слово длиннее ячейки
                lines += 1
                w -= avail
            cur = w
        total += lines
    return total


def _cell_need(cell: _CellSpec, size_pt: float) -> float:
    size_pt = cell.fixed_size or size_pt
    line_h = size_pt * LINE_FACTOR * 20
    if cell.vertical:
        # текст повёрнут: строки укладываются поперёк ширины столбца
        lines_fit = max(1, int(cell.width // line_h))
        return _text_width(cell.text, size_pt, cell.bold) / lines_fit + line_h * 0.5
    return _line_count(cell.text, cell.width, size_pt, cell.bold) * line_h


def _table_height(table: list[list[_CellSpec]], size_pt: float) -> float:
    heights = []
    for cells in table:
        need = [_cell_need(c, size_pt) for c in cells if c.vmerge is None]
        heights.append(max([MIN_ROW_H, *need]) + ROW_PAD)
    # вертикально объединённые ячейки: текст распределяется на все их строки
    for r, cells in enumerate(table):
        for c in cells:
            if c.vmerge != "restart":
                continue
            rows = [r]
            for r2 in range(r + 1, len(table)):
                if any(x.grid_col == c.grid_col and x.vmerge == "continue" for x in table[r2]):
                    rows.append(r2)
                else:
                    break
            deficit = _cell_need(c, size_pt) - sum(heights[i] for i in rows)
            if deficit > 0:
                for i in rows:
                    heights[i] += deficit / len(rows)
    return sum(heights) + len(table) * 10   # + толщина границ


def _fit_font_size(table: list[list[_CellSpec]], available_tw: float,
                   max_pt: float, min_pt: float) -> float:
    """Наибольший кегль (шаг 0,5 pt), при котором таблица помещается на страницу."""
    size = max_pt
    while size > min_pt:
        if _table_height(table, size) * FIT_SAFETY <= available_tw:
            return size
        size -= 0.5
    return min_pt


def _render_table(proto_tbl, widths: list[int], table: list[list[_CellSpec]], size_pt: float):
    tbl = copy.deepcopy(proto_tbl)
    for child in list(tbl):
        if child.tag in (qn("w:tblGrid"), qn("w:tr")):
            tbl.remove(child)
    tbl.append(parse_xml(f"<w:tblGrid {nsdecls('w')}>" + "".join(
        f'<w:gridCol w:w="{w}"/>' for w in widths) + "</w:tblGrid>"))
    half_pt = int(round(size_pt * 2))
    for r, cells in enumerate(table):
        tr = _row(68 if r == 0 else MIN_ROW_H, header=r == 0)
        for c in cells:
            size = int(round(c.fixed_size * 2)) if c.fixed_size else half_pt
            tr.append(_cell(c.width, c.text, bold=c.bold, fill=c.fill, borders=c.borders,
                            span=c.span, vmerge=c.vmerge, vertical=c.vertical, size=size))
        tbl.append(tr)
    return tbl


# ------------------------------------------------------------------ template

class _TemplateBlock:
    def __init__(self, body):
        children = [c for c in body if c.tag != qn("w:sectPr")]
        tables = [i for i, c in enumerate(children) if c.tag == qn("w:tbl")]
        if len(tables) < 2:
            raise ValueError("В шаблоне не найдены таблица-шапка и таблица расписания.")
        h_idx, s_idx = tables[0], tables[1]
        sig_idx = next((i for i in range(s_idx + 1, len(children))
                        if children[i].tag == qn("w:p") and "НАЧАЛЬНИК" in _para_text(children[i]).upper()),
                       None)
        if sig_idx is None:
            raise ValueError("В шаблоне не найден абзац подписи «НАЧАЛЬНИК ЦКОП».")
        end = sig_idx + 1
        if end < len(children) and children[end].tag == qn("w:p"):
            end += 1   # строка с фамилиями
        self.elements = children[h_idx:end]
        self.header_offset = 0
        self.schedule_offset = s_idx - h_idx
        self.sig_title_offset = sig_idx - h_idx
        self.sig_names_offset = end - 1 - h_idx


def _fill_header(header_tbl, p: DocumentParams) -> None:
    cells = header_tbl.findall(f".//{qn('w:tc')}")
    left = cells[0].findall(qn("w:p"))
    texts = [_para_text(x) for x in left]
    course_i = next((i for i, t in enumerate(texts) if "КУРС" in t.upper()), 1)
    dates_i = next((i for i, t in enumerate(texts) if "учебный год" in t), 2)
    _set_paragraph_text(left[course_i], f"{p.course} КУРСА {p.faculty_header}")
    _set_paragraph_text(
        left[dates_i],
        f"на {p.academic_year or academic_year(p.date_start)} учебный год "
        f"с {p.date_start:%d.%m.%Y} по {p.date_end:%d.%m.%Y}",
    )
    if len(cells) > 1:
        for para in cells[1].findall(qn("w:p")):
            text = _para_text(para)
            if "«" in text and re.search(r"\d{4}\s*$", text):
                # меняем только год, не трогая формат прогонов
                for t in reversed(list(para.iter(qn("w:t")))):
                    if t.text and re.search(r"\d", t.text):
                        t.text = re.sub(r"\d+\s*$", "", t.text)
                for t in reversed(list(para.iter(qn("w:t")))):
                    if t.text is not None:
                        t.text = t.text + str(p.approval_year or dt.date.today().year)
                        break


def _fill_signatures(title_p, names_p, p: DocumentParams) -> None:
    # Заголовки: «НАЧАЛЬНИК ЦКОП <tabs> ЗАМЕСТИТЕЛЬ ДЕКАНА ФАКУЛЬТЕТА» — меняем только первый текст
    first = next((t for t in title_p.iter(qn("w:t")) if t.text and t.text.strip()), None)
    if first is not None and "ЦКОП" in first.text.upper():
        first.text = p.head_ckop_title

    # Фамилии: пересобираем абзац «ИМЯ ЦКОП + пробелы + 5 табуляций + ЗАМ. ДЕКАНА»
    runs = names_p.findall(qn("w:r"))
    rpr = next((r.find(qn("w:rPr")) for r in runs if r.find(qn("w:rPr")) is not None), None)
    for r in runs:
        names_p.remove(r)

    def run(inner: str):
        r = parse_xml(f"<w:r {nsdecls('w')}>{inner}</w:r>")
        if rpr is not None:
            r.insert(0, copy.deepcopy(rpr))
        names_p.append(r)

    run(f'<w:t xml:space="preserve">{_esc(p.head_ckop_name)}                                </w:t>')
    for _ in range(5):
        run("<w:tab/>")
    run(f"<w:t>{_esc(p.deputy_dean)}</w:t>")


def _page_break():
    # минимальная высота абзаца, чтобы он не отнимал место у таблицы
    rpr = '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>'
    return parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
                     f'{rpr}</w:pPr><w:r>{rpr}<w:br w:type="page"/></w:r></w:p>')


def _available_height(sect) -> float:
    """Высота страницы под таблицу расписания (twips)."""
    page_h, top, bottom = 16840, 426, 142           # значения шаблона (A3, альбом)
    if sect is not None:
        pg, mar = sect.find(qn("w:pgSz")), sect.find(qn("w:pgMar"))
        if pg is not None and pg.get(qn("w:h")):
            page_h = int(pg.get(qn("w:h")))
        if mar is not None:
            top = abs(int(mar.get(qn("w:top"), top)))
            bottom = abs(int(mar.get(qn("w:bottom"), bottom)))
    return page_h - top - bottom - PAGE_OVERHEAD


# ------------------------------------------------------------------- public

def build_document(template: Path, sheet: SheetSchedule, pages: list[list[str]],
                   params: DocumentParams, output: Path) -> list[float]:
    """pages — наборы подгрупп по страницам, имена как в Excel ("204/1", "204/2", ...).

    Каждая страница — отдельная полная таблица со своей шапкой и подписями.
    Возвращает подобранный кегль таблицы для каждой страницы (pt).
    Шрифт шапки и подписей не меняется.
    """
    doc = Document(str(template))
    body = doc.element.body
    block = _TemplateBlock(body)
    proto = [copy.deepcopy(e) for e in block.elements]

    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)
    sect = body.find(qn("w:sectPr"))
    available = _available_height(sect)

    # столбцы каждой страницы — в порядке листа Excel; пустые страницы пропускаются
    page_columns = [[c for c in sheet.columns if c.name in set(names)] for names in pages]
    page_columns = [cols for cols in page_columns if cols]
    if not page_columns:
        raise ValueError("Ни на одной странице не выбраны подгруппы.")

    sizes: list[float] = []
    for p_i, page_cols in enumerate(page_columns):
        els = [copy.deepcopy(e) for e in proto]
        _fill_header(els[block.header_offset], params)
        widths, table = _table_specs(sheet, page_cols, params.self_study_text)
        size = _fit_font_size(table, available, params.table_font_max, params.table_font_min)
        sizes.append(size)
        els[block.schedule_offset] = _render_table(els[block.schedule_offset], widths, table, size)
        _fill_signatures(els[block.sig_title_offset], els[block.sig_names_offset], params)
        if p_i < len(page_columns) - 1:
            els.append(_page_break())
        for e in els:
            sect.addprevious(e) if sect is not None else body.append(e)

    output.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output))
    return sizes
