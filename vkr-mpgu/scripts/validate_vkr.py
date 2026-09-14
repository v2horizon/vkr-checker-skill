#!/usr/bin/env python3
"""
Валидатор ВКР для МПГУ (версия 6.33).

Проверяет готовый .docx на:
- соответствие требованиям методички МПГУ (шрифт, размер, цвет, поля, интервал,
  абзацный отступ, выравнивание) по ЭФФЕКТИВНОМУ форматированию: run → стиль
  символа → стиль абзаца → basedOn → docDefaults;
- структурные элементы (титул, аннотация, содержание, введение, главы, заключение,
  список литературы) и формальные признаки (номер страницы, нумерация глав, точки
  в заголовках, ссылки до точки, пустая последняя страница);
- соответствие ссылок [N] списку литературы, его сквозную нумерацию и алфавитный порядок;
- маркеры незавершённого черновика (общий модуль vkr_markers.py), в том числе
  в колонтитулах и сносках; код листингов не проверяется;
- личные местоимения и признаки ИИ-текста.

Код и листинги (стиль VKR Listing или моноширинный шрифт), подписи, список
литературы, титульный лист, оглавление и пункты списков не участвуют в метриках
основного текста.

Использование:
    python validate_vkr.py path/to/vkr.docx
    python validate_vkr.py path/to/vkr.docx --profile mpgu-09-project
    python validate_vkr.py path/to/vkr.docx --profile mpgu-09-regular --json
    python validate_vkr.py path/to/vkr.docx --profile mpgu-09-project -o audit/automated-validation.json

Коды возврата: 0 — ошибок нет; 1 — найдены ошибки; 2 — неверный профиль, нет
файла, повреждённый DOCX или не установлен python-docx.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import re
import statistics
import sys
from pathlib import Path


VALIDATOR_VERSION = "6.33"

# python-docx импортируется лениво (_import_docx): --help и сообщение об
# отсутствии зависимости работают без него.
_DOCX_DOCUMENT = None


def _import_docx():
    global _DOCX_DOCUMENT
    if _DOCX_DOCUMENT is None:
        from docx import Document as _Document  # noqa: WPS433 — ленивый импорт
        _DOCX_DOCUMENT = _Document
    return _DOCX_DOCUMENT


def _load_markers_module():
    """Общий модуль маркеров (SPEC 7), загружается по пути рядом со скриптом."""
    path = Path(__file__).resolve().with_name("vkr_markers.py")
    existing = sys.modules.get("vkr_markers")
    if existing is not None and Path(getattr(existing, "__file__", "") or "").resolve() == path:
        return existing
    spec = importlib.util.spec_from_file_location("vkr_markers", path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


vkr_markers = _load_markers_module()


# ========== КОНСТАНТЫ: ТРЕБОВАНИЯ МЕТОДИЧКИ МПГУ ==========

REQUIREMENTS = {
    "font_name": "Times New Roman",
    "font_size_pt": 14,
    "line_spacing": 1.5,
    "page_width_mm": 210,
    "page_height_mm": 297,
    "top_margin_mm": 20,
    "bottom_margin_mm": 20,
    "left_margin_mm": 35,
    "right_margin_mm": 10,
    "first_line_indent_cm": 1.25,
    "min_pages": 40,
    "max_pages_regular": 60,
    "max_pages_project": 50,
}

VALIDATION_PROFILES = (
    "generic",
    "mpgu-09-regular",
    "mpgu-09-project",
)

# Рекомендуемый объём всей работы по профилям (методичка 09.03.02: обычная ВКР
# 50–60 стр., проектная 40–50 стр.). Для generic диапазон задаёт программа.
PROFILE_PAGE_RANGES = {
    "mpgu-09-regular": (50, REQUIREMENTS["max_pages_regular"]),
    "mpgu-09-project": (REQUIREMENTS["min_pages"], REQUIREMENTS["max_pages_project"]),
}


def _is_mpgu(profile: str) -> bool:
    return str(profile or "").startswith("mpgu-")


# ========== СПИСКИ ДЛЯ ПРОВЕРКИ ТЕКСТА ==========

PERSONAL_PRONOUNS = [
    # 1-е лицо единственное
    r"\bя\b", r"\bмне\b", r"\bменя\b", r"\bмной\b", r"\bмною\b",
    # 1-е лицо множественное
    r"\bмы\b", r"\bнас\b", r"\bнам\b", r"\bнами\b",
    # Притяжательные 1-го лица ед.ч. — все падежи и роды
    r"\bмой\b", r"\bмоя\b", r"\bмоё\b", r"\bмои\b",
    r"\bмоего\b", r"\bмоему\b", r"\bмоим\b", r"\bмоём\b",
    r"\bмоей\b", r"\bмоею\b", r"\bмоих\b", r"\bмоими\b",
    # Притяжательные 1-го лица мн.ч. — все падежи и роды
    r"\bнаш\b", r"\bнаша\b", r"\bнаше\b", r"\bнаши\b",
    r"\bнашего\b", r"\bнашему\b", r"\bнашим\b", r"\bнашем\b",
    r"\bнашей\b", r"\bнашею\b", r"\bнаших\b", r"\bнашими\b",
]

# Важно: даже «наш» в «наш взгляд» — запрещён в МПГУ. Список soft оставлен
# пустым для совместимости.
PRONOUNS_SOFT = []

AI_CLICHES_PATH = Path(__file__).resolve().parent.parent / "data" / "ai_cliches.json"


def _load_ai_cliches():
    """Загрузка централизованного словаря клише из data/ai_cliches.json.
    Единый источник истины с ai_detection_heuristic.py.
    Валидатор использует объединённый список strong + medium.
    При ошибке загрузки — минимальный встроенный fallback.
    """
    try:
        with open(AI_CLICHES_PATH, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
        return data.get("strong", []) + data.get("medium", [])
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(
            f"ВНИМАНИЕ: не удалось загрузить {AI_CLICHES_PATH}: {e}",
            file=sys.stderr,
        )
        print("Используется минимальный fallback-список клише.", file=sys.stderr)
        return [
            "стоит отметить",
            "важно отметить",
            "важно подчеркнуть",
            "следует отметить",
            "необходимо отметить",
            "в свою очередь",
            "таким образом",
            "на сегодняшний день",
            "в современном мире",
            "давайте рассмотрим",
        ]


AI_CLICHES = _load_ai_cliches()

STATE_FILE_NAMES = ("vkr-state.md", ".vkr-state.md", "VKR-STATE.md")
# Сколько уровней вверх от каталога DOCX искать state-файл (final/vkr.docx →
# корень проекта на уровень выше). Поиск останавливается на корне проекта —
# каталоге с vkr-project.json, чтобы не взять allowlist чужого проекта.
STATE_SEARCH_LEVELS = 3


def _load_cliche_allowlist(document_path=None):
    """v6.26-friends: загрузка персонального cliche_allowlist из state-файла.

    Ищет `vkr-state.md` в каталоге DOCX и выше, не более чем на
    STATE_SEARCH_LEVELS уровней и не выше корня проекта (каталог с
    `vkr-project.json`). Без пути к документу поиск начинается с текущего
    каталога. Понимает YAML-список (с отступом и без) и inline-формат
    `cliche_allowlist: ["a", "b"]`. Возвращает фразы в нижнем регистре,
    которые НЕ считаются ИИ-клише у конкретного пользователя.

    Пустой список → проверка работает как раньше (без allowlist).
    """
    state_path = None
    try:
        start = Path(document_path).resolve().parent if document_path else Path.cwd()
        for parent in [start] + list(start.parents)[:STATE_SEARCH_LEVELS]:
            found = next((parent / name for name in STATE_FILE_NAMES if (parent / name).is_file()), None)
            if found is not None:
                state_path = found
                break
            if (parent / "vkr-project.json").is_file():
                break
    except OSError:
        return []

    if state_path is None:
        return []
    try:
        text = state_path.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError):
        return []

    block = re.search(
        r"^[ \t]*cliche_allowlist[ \t]*:[ \t]*\r?\n((?:[ \t]*-[ \t]+.+(?:\r?\n|$))+)",
        text,
        re.MULTILINE,
    )
    if not block:
        inline = re.search(
            r"^[ \t]*cliche_allowlist[ \t]*:[ \t]*\[(.*?)\][ \t]*$",
            text,
            re.MULTILINE,
        )
        if inline:
            return [
                item.strip().strip('"').strip("'").lower()
                for item in inline.group(1).split(",")
                if item.strip().strip('"').strip("'")
            ]
        return []

    items = []
    for line in block.group(1).splitlines():
        line = line.strip()
        if line.startswith("-"):
            phrase = line[1:].strip().strip('"').strip("'")
            if phrase:
                items.append(phrase.lower())
    return items


# Заполняется в validate_vkr() по state-файлу рядом с проверяемым DOCX.
CLICHE_ALLOWLIST = []

STRUCTURAL_HEADERS = [
    "аннотация",
    "содержание",
    "оглавление",
    "введение",
    "заключение",
    "список использованной литературы",
    "список использованных источников",
    "список литературы",
    "библиографический список",
    "литература",
    "библиография",
    "приложение",
]

BIBLIOGRAPHY_HEADINGS = {
    "список использованной литературы",
    "список использованных источников",
    "список литературы",
    "библиографический список",
    "литература",
    "библиография",
    "список источников",
    "список использованной литературы и источников",
    "список использованных источников и литературы",
    "список источников и литературы",
    "список литературы и источников",
    "список использованной литературы и электронных ресурсов",
}

STRUCTURE_KEYS = {
    "annotation": {"аннотация", "реферат", "abstract"},
    "toc": {"содержание", "оглавление"},
    "intro": {"введение"},
    "conclusion": {"заключение"},
    "bibliography": BIBLIOGRAPHY_HEADINGS,
}

MONOSPACE_FONTS = {
    "courier new", "courier", "consolas", "lucida console", "lucida sans typewriter",
    "menlo", "monaco", "dejavu sans mono", "liberation mono", "source code pro",
    "jetbrains mono", "fira code", "fira mono", "cascadia code", "cascadia mono",
    "pt mono", "roboto mono", "ubuntu mono", "inconsolata", "sf mono", "noto mono",
    "noto sans mono", "andale mono", "hack", "ibm plex mono", "droid sans mono",
    "anonymous pro", "space mono", "freemono", "nimbus mono l", "ocr a extended",
    "courier 10 pitch",
}
SYMBOL_FONTS = {"symbol", "wingdings", "wingdings 2", "wingdings 3", "webdings", "cambria math", "segoe ui symbol", "marlett"}

SERVICE_MARKERS = (
    "_______",
    "[Для обновления",
    "⚠️",
    "Ф.И.О.",
    "выполнена мной совершенно самостоятельно",
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{" + W_NS + "}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
M_MATH = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"

TWIPS_PER_MM = 1440 / 25.4
TWIPS_PER_CM = 1440 / 2.54


def check(condition: bool, message: str, severity: str = "error"):
    """Создать запись результата проверки."""
    return {
        "ok": bool(condition),
        "message": message,
        "severity": severity,  # "error" | "warning" | "info"
    }


def mm_from_emu(emu: int) -> float:
    """EMU → мм. 914400 EMU = 1 дюйм = 25.4 мм."""
    return round(emu / 914400 * 25.4, 1)


def pt_from_emu(emu: int) -> float:
    """EMU → pt."""
    return round(emu / 914400 * 72, 1)


# ========== ТЕКСТОВЫЕ УТИЛИТЫ ==========

def _normalize(text: str) -> str:
    text = (text or "").replace("\u00a0", " ").replace("ё", "е").replace("Ё", "Е")
    return re.sub(r"\s+", " ", text).strip().casefold()


def _heading_key(text: str) -> str:
    return _normalize(text).rstrip(" .:;")


def _is_on(value) -> bool:
    return value is None or str(value).strip().casefold() not in {"0", "false", "off", "none"}


def _toggle(element):
    if element is None:
        return None
    return _is_on(element.get(W + "val"))


def _int_attr(element, name, default=None):
    if element is None:
        return default
    value = element.get(W + name)
    if value is None:
        return default
    try:
        return int(float(value))
    except ValueError:
        return default


def _is_monospace(font_name) -> bool:
    name = str(font_name or "").casefold().strip()
    return bool(name) and (
        name in MONOSPACE_FONTS or name.endswith(" mono") or " mono " in name
        or "courier" in name or name.startswith("consol")
    )


TOC_LINE_RE = re.compile(r"(?:\t|\.{3,}|…{2,}|\s{3,})\s*\d{1,3}\s*$")

ROMAN_RE = re.compile(r"^M{0,3}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$")

CHAPTER_START_RE = re.compile(
    r"^\s*глава\s+([IVXLCDM]+|\d+|перв\w*|втор\w*|трет\w*|четв[её]рт\w*|пят\w*|шест\w*)(?=$|[\s.:,)—–-])",
    re.IGNORECASE,
)

CAPTION_RE = re.compile(
    r"^\s*(таблица|рисунок|рис\.|листинг|схема|диаграмма)\s+(?:[А-ЯA-Z]\.)?\d+(?:[.\-]\d+)*"
    r"\s*(?:$|[—–\-.:]\s*\S?.*$)",
    re.IGNORECASE | re.DOTALL,
)

APPENDIX_LABEL_RE = re.compile(
    r"^\s*приложени[ея](?:\s+(?:[А-ЯA-Z]|\d{1,2}))?\s*(?:$|[.—–\-:]\s*(?:\S.*)?$)",
    re.IGNORECASE | re.DOTALL,
)

# Сокращения, после точки которых ссылка [N] и конец заголовка допустимы.
ABBREVIATION_END_RE = re.compile(
    r"(?:^|[\s(«\"'„“\[/—–-])(?:"
    r"и\s*др|и\s*пр|и\s*т\.\s*д|и\s*т\.\s*п|т\.\s*е|т\.\s*д|т\.\s*п|т\.\s*к|т\.\s*н|"
    r"см|ср|напр|г|гг|в|вв|с|стр|р|рис|табл|им|ст|ч|п|пп|гл|разд|т|тт|вып|изд|ред|сост|пер|"
    r"англ|нем|фр|лат|др|пр|акад|проф|доц|канд|млн|млрд|тыс|руб|коп|долл|мин|сек|обл|ул|пос|"
    r"кв|экз|л|№|"
    r"et\s+al|etc|vol|vols|p|pp|no|nos|ed|eds|cf|ibid|op\.\s*cit|e\.\s*g|i\.\s*e|fig|ch|sec|"
    r"univ|dept|jr|sr|dr|prof|inc|ltd|co"
    r")\.\s*$",
    re.IGNORECASE,
)
INITIALS_END_RE = re.compile(r"(?:^|[\s(«\"„“])(?:[А-ЯЁA-Z]\.\s*){1,3}$")
SECTION_NUMBER_END_RE = re.compile(r"(?:^|\s)(?:п\.\s*|пп\.\s*|§\s*)?\d+(?:\.\d+)+\.\s*$")


def _ends_with_abbreviation(text: str) -> bool:
    stripped = (text or "").rstrip()
    if not stripped.endswith("."):
        return False
    return bool(
        ABBREVIATION_END_RE.search(stripped) or INITIALS_END_RE.search(stripped)
        or SECTION_NUMBER_END_RE.search(stripped)
    )


def _ends_with_ellipsis(text: str) -> bool:
    stripped = (text or "").rstrip()
    return stripped.endswith("...") or stripped.endswith("…")


def _is_service_text(text: str) -> bool:
    lowered = (text or "").casefold()
    return any(marker.casefold() in lowered for marker in SERVICE_MARKERS)

# ========== МОДЕЛЬ ДОКУМЕНТА: ЭФФЕКТИВНОЕ ФОРМАТИРОВАНИЕ ==========

_EXCLUDED_RUN_ANCESTORS = {W + "del", W + "moveFrom", W + "txbxContent", MC + "Fallback"}

THEME_COLOR_KEYS = {
    "dark1": "dk1", "text1": "dk1", "light1": "lt1", "background1": "lt1",
    "dark2": "dk2", "text2": "dk2", "light2": "lt2", "background2": "lt2",
    "accent1": "accent1", "accent2": "accent2", "accent3": "accent3",
    "accent4": "accent4", "accent5": "accent5", "accent6": "accent6",
    "hyperlink": "hlink", "followedHyperlink": "folHlink",
}


def _parse_xml(blob):
    from lxml import etree  # python-docx зависит от lxml

    parser = etree.XMLParser(resolve_entities=False, huge_tree=True, recover=True)
    return etree.fromstring(blob, parser)


def _part_root(part):
    element = getattr(part, "element", None)
    if element is not None:
        return element
    blob = getattr(part, "blob", None)
    if not blob:
        return None
    try:
        return _parse_xml(blob)
    except Exception:  # noqa: BLE001 — повреждённая часть не должна ронять проверку
        return None


def _related_parts(doc, suffix):
    parts = []
    try:
        rels = doc.part.rels
    except AttributeError:
        return parts
    for rel in rels.values():
        try:
            if rel.is_external or not str(rel.reltype).endswith(suffix):
                continue
            parts.append((rel.rId, rel.target_part))
        except (AttributeError, KeyError, ValueError):
            continue
    return parts


class _StyleSheet:
    def __init__(self, doc):
        root = doc.styles.element
        self.by_id = {}
        self.names = {}
        self.default_para = None
        self.default_char = None
        for style in root.iterchildren(W + "style"):
            sid = style.get(W + "styleId")
            if sid is None or sid in self.by_id:
                continue
            self.by_id[sid] = style
            name = style.find(W + "name")
            self.names[sid] = ((name.get(W + "val") if name is not None else None) or sid).strip()
            if style.get(W + "default") is not None and _is_on(style.get(W + "default")):
                kind = style.get(W + "type")
                if kind == "paragraph" and self.default_para is None:
                    self.default_para = sid
                elif kind == "character" and self.default_char is None:
                    self.default_char = sid
        defaults = root.find(W + "docDefaults")
        self.rpr_default = defaults.find(W + "rPrDefault/" + W + "rPr") if defaults is not None else None
        self.ppr_default = defaults.find(W + "pPrDefault/" + W + "pPr") if defaults is not None else None
        self._chains = {}
        self._rpr_chains = {}
        self._ppr_chains = {}

    def chain(self, sid):
        if sid in self._chains:
            return self._chains[sid]
        result = []
        seen = set()
        current = sid
        while current and current not in seen and current in self.by_id and len(result) < 30:
            seen.add(current)
            style = self.by_id[current]
            result.append(style)
            based = style.find(W + "basedOn")
            current = based.get(W + "val") if based is not None else None
        self._chains[sid] = result
        return result

    def rpr_chain(self, sid):
        if sid not in self._rpr_chains:
            self._rpr_chains[sid] = [s.find(W + "rPr") for s in self.chain(sid) if s.find(W + "rPr") is not None]
        return self._rpr_chains[sid]

    def ppr_chain(self, sid):
        if sid not in self._ppr_chains:
            self._ppr_chains[sid] = [s.find(W + "pPr") for s in self.chain(sid) if s.find(W + "pPr") is not None]
        return self._ppr_chains[sid]

    def name(self, sid):
        return self.names.get(sid, "") if sid else ""

    def chain_names(self, sid):
        return [self.names.get(s.get(W + "styleId"), "").casefold() for s in self.chain(sid)]


class _Theme:
    def __init__(self, doc):
        self.fonts = {}
        self.colors = {}
        for _rid, part in _related_parts(doc, "/theme"):
            root = _part_root(part)
            if root is None:
                continue
            for family in ("majorFont", "minorFont"):
                node = root.find(".//" + A + "fontScheme/" + A + family + "/" + A + "latin")
                if node is not None and node.get("typeface"):
                    self.fonts[family[:5]] = node.get("typeface")
            scheme = root.find(".//" + A + "clrScheme")
            if scheme is not None:
                for child in scheme:
                    key = child.tag.split("}")[-1]
                    srgb = child.find(A + "srgbClr")
                    sysclr = child.find(A + "sysClr")
                    if srgb is not None and srgb.get("val"):
                        self.colors[key] = srgb.get("val").upper()
                    elif sysclr is not None and (sysclr.get("lastClr") or sysclr.get("val") == "windowText"):
                        self.colors[key] = (sysclr.get("lastClr") or "000000").upper()
            break

    def font(self, theme_value):
        value = str(theme_value or "").casefold()
        return self.fonts.get("major" if value.startswith("major") else "minor")

    def color(self, theme_color, tint=None, shade=None):
        base = self.colors.get(THEME_COLOR_KEYS.get(str(theme_color or ""), ""))
        if not base or not re.fullmatch(r"[0-9A-F]{6}", base):
            return None
        channels = [int(base[i:i + 2], 16) for i in (0, 2, 4)]
        try:
            if shade:
                factor = int(shade, 16) / 255
                channels = [round(c * factor) for c in channels]
            if tint:
                factor = int(tint, 16) / 255
                channels = [round(c * factor + 255 * (1 - factor)) for c in channels]
        except ValueError:
            pass
        return "".join(f"{max(0, min(255, c)):02X}" for c in channels)


class _Numbering:
    def __init__(self, doc):
        self.num_to_abstract = {}
        self.formats = {}
        for _rid, part in _related_parts(doc, "/numbering"):
            root = _part_root(part)
            if root is None:
                continue
            for abstract in root.iterchildren(W + "abstractNum"):
                aid = abstract.get(W + "abstractNumId")
                for level in abstract.iterchildren(W + "lvl"):
                    fmt = level.find(W + "numFmt")
                    self.formats[(aid, level.get(W + "ilvl") or "0")] = fmt.get(W + "val") if fmt is not None else None
            for num in root.iterchildren(W + "num"):
                ref = num.find(W + "abstractNumId")
                if ref is not None:
                    self.num_to_abstract[num.get(W + "numId")] = ref.get(W + "val")
            break

    def fmt(self, num_id, ilvl="0"):
        abstract = self.num_to_abstract.get(str(num_id))
        if abstract is None:
            return None
        return self.formats.get((abstract, str(ilvl or "0")))


class _Run:
    __slots__ = ("el", "text", "hyperlink")

    def __init__(self, el, text, hyperlink):
        self.el = el
        self.text = text
        self.hyperlink = hyperlink


def _run_text(run_el):
    parts = []
    for child in run_el:
        tag = child.tag
        if tag == W + "t":
            parts.append(child.text or "")
        elif tag in (W + "tab", W + "ptab"):
            parts.append("\t")
        elif tag == W + "br":
            if child.get(W + "type") in (None, "textWrapping"):
                parts.append("\n")
        elif tag == W + "cr":
            parts.append("\n")
        elif tag == W + "noBreakHyphen":
            parts.append("-")
    return "".join(parts)


def _collect_runs(p_el):
    runs = []
    for run_el in p_el.iter(W + "r"):
        hyperlink = False
        skip = False
        ancestor = run_el.getparent()
        while ancestor is not None and ancestor is not p_el:
            if ancestor.tag in _EXCLUDED_RUN_ANCESTORS:
                skip = True
                break
            if ancestor.tag == W + "hyperlink":
                hyperlink = True
            ancestor = ancestor.getparent()
        if skip:
            continue
        runs.append(_Run(run_el, _run_text(run_el), hyperlink))
    return runs


class _Para:
    __slots__ = (
        "model", "el", "idx", "block", "container", "location", "style_id", "runs", "text",
        "section", "in_toc_field", "in_toc_sdt", "cache",
    )

    def __init__(self, model, el, container, location=""):
        self.model = model
        self.el = el
        self.idx = -1
        self.block = -1
        self.container = container
        self.location = location
        ppr = el.find(W + "pPr")
        style = ppr.find(W + "pStyle") if ppr is not None else None
        sid = style.get(W + "val") if style is not None else None
        if sid not in model.styles.by_id:
            sid = model.styles.default_para
        self.style_id = sid
        self.runs = _collect_runs(el)
        self.text = "".join(run.text for run in self.runs)
        self.section = 0
        self.in_toc_field = False
        self.in_toc_sdt = False
        self.cache = {}


class _Section:
    def __init__(self, index, sect_pr):
        self.index = index
        self.el = sect_pr
        pg_sz = sect_pr.find(W + "pgSz") if sect_pr is not None else None
        pg_mar = sect_pr.find(W + "pgMar") if sect_pr is not None else None
        self.width = _int_attr(pg_sz, "w")
        self.height = _int_attr(pg_sz, "h")
        self.orient = pg_sz.get(W + "orient") if pg_sz is not None else None
        self.top = _int_attr(pg_mar, "top")
        self.bottom = _int_attr(pg_mar, "bottom")
        self.left = _int_attr(pg_mar, "left")
        self.right = _int_attr(pg_mar, "right")
        self.gutter = _int_attr(pg_mar, "gutter", 0) or 0
        title_pg = sect_pr.find(W + "titlePg") if sect_pr is not None else None
        self.title_pg = bool(_toggle(title_pg))
        pg_num = sect_pr.find(W + "pgNumType") if sect_pr is not None else None
        self.pg_start = _int_attr(pg_num, "start")
        self.pg_fmt = pg_num.get(W + "fmt") if pg_num is not None else None
        kind = sect_pr.find(W + "type") if sect_pr is not None else None
        self.break_type = (kind.get(W + "val") if kind is not None else None) or "nextPage"
        self.footer_refs = {}
        self.header_refs = {}
        if sect_pr is not None:
            for ref in sect_pr.iterchildren(W + "footerReference", W + "headerReference"):
                target = self.footer_refs if ref.tag == W + "footerReference" else self.header_refs
                rid = ref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                target[ref.get(W + "type") or "default"] = rid
        self.first_idx = None
        self.last_idx = None


def _mm(twips):
    return None if twips is None else round(twips / TWIPS_PER_MM, 1)


def _is_toc_sdt(sdt):
    pr = sdt.find(W + "sdtPr")
    if pr is None:
        return False
    gallery = pr.find(W + "docPartObj/" + W + "docPartGallery")
    if gallery is None:
        gallery = pr.find(W + "docPartList/" + W + "docPartGallery")
    return gallery is not None and "table of contents" in str(gallery.get(W + "val") or "").casefold()


def _field_types(root):
    """Типы полей в порядке документа: PAGE, TOC, NUMPAGES…"""
    types = []
    stack = []
    for node in root.iter(W + "fldChar", W + "instrText", W + "fldSimple"):
        if node.tag == W + "fldSimple":
            code = (node.get(W + "instr") or "").strip()
            if code:
                types.append(code.split()[0].upper())
            continue
        if node.tag == W + "instrText":
            if stack and stack[-1] is not None:
                stack[-1].append(node.text or "")
            continue
        kind = node.get(W + "fldCharType")
        if kind == "begin":
            stack.append([])
        elif kind == "separate":
            if stack and stack[-1] is not None:
                code = "".join(stack[-1]).strip()
                if code:
                    types.append(code.split()[0].upper())
                stack[-1] = None
        elif kind == "end" and stack:
            top = stack.pop()
            if top is not None:
                code = "".join(top).strip()
                if code:
                    types.append(code.split()[0].upper())
    return types


class _DocModel:
    """Однократный разбор DOCX: абзацы тела (включая w:sdt), таблицы, секции,
    колонтитулы, сноски и области работы (титул, оглавление, список литературы,
    приложения, последний лист)."""

    def __init__(self, doc):
        self.doc = doc
        self.styles = _StyleSheet(doc)
        self.theme = _Theme(doc)
        self.numbering = _Numbering(doc)
        self.body = []
        self.blocks = []
        self.table_paras = []
        self.floating_paras = []
        self.sections = []
        self.part_paras = []
        self._part_cache = {}
        self._by_element = {}
        self._walk_body()
        self._mark_toc_fields()
        self._collect_parts()
        self._compute_regions()

    # ---------- разбор ----------

    def _new_para(self, el, container, location=""):
        para = _Para(self, el, container, location)
        self._by_element[el] = para
        for box in el.iter(W + "txbxContent"):
            if box.getparent() is not None:
                for inner in box.iter(W + "p"):
                    if inner not in self._by_element:
                        floating = _Para(self, inner, "textbox", "надпись")
                        self._by_element[inner] = floating
                        self.floating_paras.append(floating)
        return para

    def _walk_body(self):
        body = self.doc.element.body
        state = {"section": 0}

        def visit(container, in_toc_sdt):
            for child in container:
                tag = child.tag
                if tag == W + "p":
                    para = self._new_para(child, "body")
                    para.idx = len(self.body)
                    para.block = len(self.blocks)
                    para.section = state["section"]
                    para.in_toc_sdt = in_toc_sdt
                    self.body.append(para)
                    self.blocks.append(("p", para))
                    ppr = child.find(W + "pPr")
                    sect = ppr.find(W + "sectPr") if ppr is not None else None
                    if sect is not None:
                        self._add_section(sect)
                        state["section"] += 1
                elif tag == W + "tbl":
                    cell_paras = []
                    for p_el in child.iter(W + "p"):
                        if p_el in self._by_element:
                            continue
                        para = self._new_para(p_el, "table", "таблица")
                        para.section = state["section"]
                        cell_paras.append(para)
                    self.table_paras.extend(cell_paras)
                    self.blocks.append(("tbl", child, cell_paras))
                elif tag == W + "sdt":
                    content = child.find(W + "sdtContent")
                    if content is not None:
                        visit(content, in_toc_sdt or _is_toc_sdt(child))
                elif tag in (W + "customXml", W + "smartTag"):
                    visit(child, in_toc_sdt)
                elif tag == MC + "AlternateContent":
                    choice = child.find(MC + "Choice")
                    if choice is not None:
                        visit(choice, in_toc_sdt)

        visit(body, False)
        final = body.find(W + "sectPr")
        if final is not None or not self.sections or (self.body and self.body[-1].section >= len(self.sections)):
            self._add_section(final)
        for para in self.body:
            section = self.sections[min(para.section, len(self.sections) - 1)]
            if section.first_idx is None:
                section.first_idx = para.idx
            section.last_idx = para.idx

    def _add_section(self, sect_pr):
        self.sections.append(_Section(len(self.sections), sect_pr))

    def _mark_toc_fields(self):
        """Абзац относится к полю TOC, если в нём есть видимый текст внутри поля.

        Word иногда ставит конец поля в начало следующего абзаца (например,
        «Введение»): такой абзац оглавлением не считается.
        """
        stack = []

        def toc_open():
            return any(entry.get("type") == "TOC" for entry in stack)

        for para in self.body:
            in_toc = False
            for node in para.el.iter(W + "fldChar", W + "instrText", W + "fldSimple", W + "t"):
                tag = node.tag
                if tag == W + "t":
                    if node.text and node.text.strip() and toc_open():
                        in_toc = True
                    continue
                if tag == W + "fldSimple":
                    if (node.get(W + "instr") or "").strip().upper().startswith("TOC"):
                        in_toc = True
                    continue
                if tag == W + "instrText":
                    if stack and stack[-1].get("state") == "code":
                        stack[-1]["code"] += node.text or ""
                        if stack[-1]["code"].strip().upper().startswith("TOC"):
                            stack[-1]["type"] = "TOC"
                    continue
                kind = node.get(W + "fldCharType")
                if kind == "begin":
                    stack.append({"code": "", "state": "code", "type": None})
                elif kind == "separate" and stack:
                    stack[-1]["state"] = "result"
                elif kind == "end" and stack:
                    stack.pop()
            para.in_toc_field = in_toc

    def _collect_parts(self):
        seen = set()
        for section in self.sections:
            for kind, refs in (("header", section.header_refs), ("footer", section.footer_refs)):
                for ref_type, rid in refs.items():
                    root = self._part_by_rid(rid)
                    if root is None or id(root) in seen:
                        continue
                    seen.add(id(root))
                    label = "верхний колонтитул" if kind == "header" else "нижний колонтитул"
                    for p_el in root.iter(W + "p"):
                        if p_el not in self._by_element:
                            self.part_paras.append(self._new_para(p_el, kind, label))
        for suffix, item_tag, label in (("/footnotes", "footnote", "сноска"), ("/endnotes", "endnote", "концевая сноска")):
            for _rid, part in _related_parts(self.doc, suffix):
                root = _part_root(part)
                if root is None:
                    continue
                for note in root.iterchildren(W + item_tag):
                    if note.get(W + "type") in ("separator", "continuationSeparator", "continuationNotice"):
                        continue
                    for p_el in note.iter(W + "p"):
                        if p_el not in self._by_element:
                            self.part_paras.append(self._new_para(p_el, "note", label))

    def _part_by_rid(self, rid):
        if not rid:
            return None
        if rid in self._part_cache:
            return self._part_cache[rid]
        root = None
        try:
            root = _part_root(self.doc.part.related_parts[rid])
        except (KeyError, AttributeError):
            root = None
        self._part_cache[rid] = root
        return root

    def para_for(self, element):
        return self._by_element.get(element)

    # ---------- стили и эффективные свойства абзаца ----------

    def style_name(self, para):
        return self.styles.name(para.style_id).casefold()

    def style_names(self, para):
        return self.styles.chain_names(para.style_id)

    def _ppr_chain(self, para):
        chain = para.cache.get("ppr")
        if chain is None:
            chain = []
            direct = para.el.find(W + "pPr")
            if direct is not None:
                chain.append(direct)
            chain.extend(self.styles.ppr_chain(para.style_id))
            if self.styles.ppr_default is not None:
                chain.append(self.styles.ppr_default)
            para.cache["ppr"] = chain
        return chain

    def _first_ppr(self, para, tag, attrs=None):
        for ppr in self._ppr_chain(para):
            el = ppr.find(W + tag)
            if el is None:
                continue
            if attrs is None or any(el.get(W + attr) is not None for attr in attrs):
                return el
        return None

    def alignment(self, para):
        if "jc" not in para.cache:
            el = self._first_ppr(para, "jc")
            value = (el.get(W + "val") if el is not None else None) or "left"
            value = {"start": "left", "end": "right", "distribute": "both", "lowKashida": "both",
                     "mediumKashida": "both", "highKashida": "both", "thaiDistribute": "both"}.get(value, value)
            para.cache["jc"] = value
        return para.cache["jc"]

    def line_spacing(self, para):
        """(line, rule): line в 240-х долях строки для auto или в twips для exact/atLeast."""
        el = self._first_ppr(para, "spacing", ("line",))
        if el is None:
            return 240, "auto"
        return _int_attr(el, "line", 240), (el.get(W + "lineRule") or "auto")

    def first_line_cm(self, para):
        el = self._first_ppr(para, "ind", ("firstLine", "hanging", "firstLineChars", "hangingChars"))
        if el is None:
            return 0.0
        for attr, sign in (("hangingChars", -1), ("firstLineChars", 1)):
            chars = _int_attr(el, attr)
            if chars:
                return sign * chars / 100 * 14 * 2.54 / 72
        hanging = _int_attr(el, "hanging")
        if hanging is not None:
            return -hanging / TWIPS_PER_CM
        return (_int_attr(el, "firstLine", 0) or 0) / TWIPS_PER_CM

    def spacing_pt(self, para, side):
        el = self._first_ppr(para, "spacing", (side, side + "Lines", side + "Autospacing"))
        if el is None:
            return 0.0
        if el.get(W + side + "Autospacing") is not None and _is_on(el.get(W + side + "Autospacing")):
            return 14.0
        lines = _int_attr(el, side + "Lines")
        if lines:
            return lines / 100 * 14 * 1.15
        return (_int_attr(el, side, 0) or 0) / 20

    def page_break_before(self, para):
        el = self._first_ppr(para, "pageBreakBefore")
        return bool(_toggle(el))

    def outline_level(self, para):
        el = self._first_ppr(para, "outlineLvl")
        level = _int_attr(el, "val")
        return level if level is not None and level < 9 else None

    def heading_level(self, para):
        """Уровень заголовка 1..9 по outlineLvl (с учётом стиля) или имени стиля."""
        if "heading" not in para.cache:
            level = self.outline_level(para)
            result = level + 1 if level is not None else None
            if result is None:
                for name in self.style_names(para)[:1]:
                    match = re.match(r"^(?:heading|заголовок)\s*([1-9])$", name)
                    if match:
                        result = int(match.group(1))
            para.cache["heading"] = result
        return para.cache["heading"]

    def num_pr(self, para):
        """(numId, ilvl) эффективной нумерации абзаца или None."""
        if "numpr" not in para.cache:
            result = None
            for ppr in self._ppr_chain(para):
                numpr = ppr.find(W + "numPr")
                if numpr is None:
                    continue
                num_id = numpr.find(W + "numId")
                if num_id is None:
                    continue
                value = num_id.get(W + "val")
                if value not in (None, "0"):
                    ilvl = numpr.find(W + "ilvl")
                    result = (value, ilvl.get(W + "val") if ilvl is not None else "0")
                break
            para.cache["numpr"] = result
        return para.cache["numpr"]

    def is_list_item(self, para):
        if self.num_pr(para) is not None:
            return True
        names = self.style_names(para)[:1]
        return any(
            name.startswith("list") or name in {"абзац списка", "маркированный список", "нумерованный список"}
            for name in names
        )

    # ---------- эффективные свойства run ----------

    def _rpr_chain(self, para, run):
        rpr = run.el.find(W + "rPr")
        rstyle_el = rpr.find(W + "rStyle") if rpr is not None else None
        rstyle = rstyle_el.get(W + "val") if rstyle_el is not None else self.styles.default_char
        key = (rstyle, para.style_id)
        style_chain = self.styles._rpr_chains.get(("run",) + key)
        if style_chain is None:
            style_chain = []
            if rstyle and rstyle in self.styles.by_id:
                style_chain.extend(self.styles.rpr_chain(rstyle))
            style_chain.extend(self.styles.rpr_chain(para.style_id))
            if self.styles.rpr_default is not None:
                style_chain.append(self.styles.rpr_default)
            self.styles._rpr_chains[("run",) + key] = style_chain
        return ([rpr] if rpr is not None else []) + style_chain

    def run_font(self, para, run):
        slot = "hAnsi" if re.search(r"[^\x00-\x7f]", run.text or "") else "ascii"
        for rpr in self._rpr_chain(para, run):
            fonts = rpr.find(W + "rFonts")
            if fonts is None:
                continue
            theme = fonts.get(W + slot + "Theme")
            if theme:
                name = self.theme.font(theme)
                if name:
                    return name
            value = fonts.get(W + slot)
            if value:
                return value
        return "Times New Roman"

    def run_size(self, para, run):
        for rpr in self._rpr_chain(para, run):
            size = rpr.find(W + "sz")
            value = _int_attr(size, "val")
            if value:
                return value / 2
        return 10.0

    def run_color(self, para, run):
        """Цвет в HEX или None для «Авто» (чёрный)."""
        for rpr in self._rpr_chain(para, run):
            color = rpr.find(W + "color")
            if color is None:
                continue
            value = (color.get(W + "val") or "").strip()
            if re.fullmatch(r"[0-9A-Fa-f]{6}", value):
                return value.upper()
            theme = color.get(W + "themeColor")
            if theme:
                return self.theme.color(theme, color.get(W + "themeTint"), color.get(W + "themeShade"))
            return None
        return None

    def run_toggle(self, para, run, tag):
        for rpr in self._rpr_chain(para, run):
            value = _toggle(rpr.find(W + tag))
            if value is not None:
                return value
        return False

    def is_hyperlink_run(self, para, run):
        if run.hyperlink:
            return True
        rpr = run.el.find(W + "rPr")
        rstyle = rpr.find(W + "rStyle") if rpr is not None else None
        name = self.styles.name(rstyle.get(W + "val")).casefold() if rstyle is not None else ""
        return "hyperlink" in name

    def visible_runs(self, para):
        runs = para.cache.get("visible")
        if runs is None:
            runs = [run for run in para.runs if run.text.strip() and not self.run_toggle(para, run, "vanish")]
            para.cache["visible"] = runs
        return runs

    def all_bold(self, para):
        runs = self.visible_runs(para)
        return bool(runs) and all(self.run_toggle(para, run, "b") for run in runs)

    # ---------- классификация абзацев ----------

    def is_toc(self, para):
        if para.in_toc_field or para.in_toc_sdt:
            return True
        name = self.style_name(para)
        if re.match(r"^(?:toc|оглавление)\s*\d$", name) or name in {"toc heading", "заголовок оглавления"}:
            return True
        return bool(para.container == "body" and TOC_LINE_RE.search(para.text.strip()) and len(para.text.strip()) <= 250)

    def is_code(self, para):
        if "code" not in para.cache:
            names = self.style_names(para)
            result = any(
                "listing" in name or "листинг" in name or name.startswith("code") or "source code" in name
                or name in {"html preformatted", "plain text", "программный код", "код"}
                for name in names
            )
            if not result:
                total = mono = 0
                for run in self.visible_runs(para):
                    size = len(run.text.strip())
                    total += size
                    if _is_monospace(self.run_font(para, run)):
                        mono += size
                result = total > 0 and mono / total >= 0.6
            para.cache["code"] = result
        return para.cache["code"]

    def text_without_code(self, para):
        """Текст абзаца без кода: абзац-листинг пуст, моноширинные фрагменты заменены пробелом."""
        if self.is_code(para):
            return ""
        parts = []
        for run in para.runs:
            if run.text.strip() and _is_monospace(self.run_font(para, run)):
                parts.append(" ")
            else:
                parts.append(run.text)
        return "".join(parts)

    def is_caption(self, para):
        name = self.style_name(para)
        if name in {"vkr caption", "caption"}:
            return True
        text = para.text.strip()
        return bool(text) and len(text) <= 300 and bool(CAPTION_RE.match(text))

    def structural_key(self, para):
        if para.container != "body":
            return None
        text = para.text.strip()
        if not text or len(text) > 120 or "\t" in text:
            return None
        key = _heading_key(text)
        for name, variants in STRUCTURE_KEYS.items():
            if key in variants:
                return name
        return None

    def is_appendix_label(self, para):
        if para.container != "body" or self.is_toc(para):
            return False
        name = self.style_name(para)
        if name == "vkr appendix label":
            return True
        text = para.text.strip()
        if not text or len(text) > 150 or not APPENDIX_LABEL_RE.match(text):
            return False
        key = _heading_key(text)
        if re.fullmatch(r"приложени[ея](?: (?:[а-яa-z]|\d{1,2}))?", key):
            return True
        # «Приложение 1. Анкета» — только если абзац оформлен как заголовок
        return (
            self.heading_level(para) is not None or self.all_bold(para)
            or self.alignment(para) in {"center", "right"}
        )

    def is_chapter_heading(self, para):
        if "chapter" in para.cache:
            return para.cache["chapter"]
        result = False
        text = para.text.strip()
        if (
            para.container == "body" and not self.is_toc(para) and text and len(text) <= 250
            and CHAPTER_START_RE.match(text) and not TOC_LINE_RE.search(text)
        ):
            level = self.heading_level(para)
            if level is not None and level <= 2:
                result = True
            elif any("глава" in name or "chapter" in name for name in self.style_names(para)):
                result = True
            else:
                first_line = text.split("\n")[0].strip()
                upper = first_line.upper() == first_line and re.search(r"[А-ЯЁ]", first_line)
                result = bool(self.all_bold(para) or self.alignment(para) == "center" or upper)
        para.cache["chapter"] = result
        return result

    def is_heading_like(self, para):
        if self.heading_level(para) is not None or self.structural_key(para) or self.is_chapter_heading(para):
            return True
        if self.is_appendix_label(para) or self.style_name(para) in {"title", "subtitle", "vkr appendix title", "vkr title"}:
            return True
        text = para.text.strip()
        if not text or len(text) > 150:
            return False
        ends_like_sentence = text.endswith(".") and not _ends_with_abbreviation(text)
        return not ends_like_sentence and (self.all_bold(para) or (self.alignment(para) == "center" and len(text) <= 120))

    # ---------- области работы ----------

    def _compute_regions(self):
        body = self.body
        count = len(body)
        self.title_end = self._find_title_end()
        self.conclusion_idx = next(
            (p.idx for p in body[self.title_end:] if not self.is_toc(p) and self.structural_key(p) == "conclusion"),
            None,
        )
        bib_candidates = [p.idx for p in body if not self.is_toc(p) and self.structural_key(p) == "bibliography"]
        after_conclusion = [idx for idx in bib_candidates if self.conclusion_idx is not None and idx > self.conclusion_idx]
        if after_conclusion:
            self.bib_start = after_conclusion[0]
        elif bib_candidates:
            # без «Заключения»: последний заголовок, оформленный как заголовок
            # (подзаголовок «Литература» обычным абзацем внутри списка — не начало раздела)
            styled = [
                idx for idx in bib_candidates
                if self.heading_level(body[idx]) is not None or self.all_bold(body[idx])
                or self.alignment(body[idx]) == "center"
            ]
            self.bib_start = styled[-1] if styled else bib_candidates[-1]
        else:
            self.bib_start = None
        self.last_page_start = self._find_last_page_start()
        appendix_from = self.bib_start if self.bib_start is not None else (self.conclusion_idx or self.title_end)
        self.appendix_start = next(
            (p.idx for p in body[appendix_from:] if self.is_appendix_label(p)),
            None,
        )
        if self.appendix_start is not None and self.last_page_start is not None and self.appendix_start > self.last_page_start:
            self.appendix_start = None
        self.bib_end = None
        if self.bib_start is not None:
            self.bib_end = self._find_bibliography_end()
        ends = [value for value in (self.bib_start, self.appendix_start, self.last_page_start) if value is not None]
        self.main_end = min(ends) if ends else count

    def _find_title_end(self):
        first_structural = next(
            (p.idx for p in self.body if self.is_toc(p) or self.structural_key(p) or self.is_chapter_heading(p)),
            None,
        )
        # VKR Title генератор ставит и на титул, и на последний лист: титул — только
        # абзацы этого стиля до первого структурного заголовка.
        titled = [
            p.idx for p in self.body
            if self.style_name(p) == "vkr title" and (first_structural is None or p.idx < first_structural)
        ]
        if titled:
            return max(titled) + 1
        if first_structural is not None:
            return first_structural
        if len(self.sections) > 1:
            first = self.sections[0]
            if first.last_idx is not None and first.last_idx < len(self.body) - 1:
                text_size = sum(len(p.text) for p in self.body[: first.last_idx + 1])
                if text_size < 2000:
                    return first.last_idx + 1
        size = 0
        for para in self.body[:40]:
            size += len(para.text)
            if size > 1500:
                break
            if any(br.get(W + "type") == "page" for br in para.el.iter(W + "br")):
                return para.idx + 1
        return 0

    def _find_last_page_start(self):
        oath = "выполнена мной совершенно самостоятельно"
        start_from = max(self.title_end, self.conclusion_idx or 0)
        for index, block in enumerate(self.blocks):
            if block[0] != "tbl":
                continue
            if not any(oath in p.text.casefold() for p in block[2]):
                continue
            previous = [b[1] for b in self.blocks[max(0, index - 4):index] if b[0] == "p"]
            for para in reversed(previous):
                if _heading_key(para.text) == "выпускная квалификационная работа" and para.idx >= start_from:
                    return para.idx
            following = [b[1] for b in self.blocks[index:] if b[0] == "p"]
            if following and following[0].idx >= start_from:
                return following[0].idx
        # Последний лист в стиле VKR Title (генератор 6.33) после основного текста
        for para in self.body[start_from:]:
            if self.style_name(para) == "vkr title" and para.idx >= self.title_end and (
                self.conclusion_idx is None or para.idx > self.conclusion_idx
            ):
                return para.idx
        return None

    def _find_bibliography_end(self):
        start = self.bib_start
        first_block = self.body[start].block
        for block in self.blocks[first_block + 1:]:
            if block[0] == "tbl":
                following = [b[1] for b in self.blocks[self.blocks.index(block):] if b[0] == "p"]
                return following[0].idx if following else len(self.body)
            para = block[1]
            if self.last_page_start is not None and para.idx >= self.last_page_start:
                return para.idx
            if self.is_appendix_label(para) or self.structural_key(para) in {"annotation", "toc", "intro", "conclusion"}:
                return para.idx
            level = self.heading_level(para)
            if level == 1 and not re.match(r"^\s*\d", para.text) and self.num_pr(para) is None:
                return para.idx
        return len(self.body)

    def region(self, para):
        """title | toc | main | bibliography | appendix | last_page | other."""
        if para.container != "body":
            return para.container
        if self.is_toc(para):
            return "toc"
        idx = para.idx
        if idx < self.title_end:
            return "title"
        if self.last_page_start is not None and idx >= self.last_page_start:
            return "last_page"
        if self.bib_start is not None and self.bib_start <= idx < (self.bib_end if self.bib_end is not None else len(self.body)):
            return "bibliography"
        if self.appendix_start is not None and idx >= self.appendix_start:
            return "appendix"
        if idx >= self.main_end:
            return "other"
        return "main"

    def main_paras(self):
        return [p for p in self.body[self.title_end:self.main_end] if not self.is_toc(p)]

    def chapter_headings(self):
        return [p for p in self.body[self.title_end:self.main_end] if self.is_chapter_heading(p)]

    def body_text_paras(self):
        """Абзацы для метрик основного текста (интервал): без заголовков, кода, подписей, служебных строк."""
        result = []
        service_styles = {"vkr title", "vkr bibliography", "vkr appendix label", "vkr appendix title", "vkr caption", "vkr listing"}
        for para in self.main_paras():
            text = para.text.strip()
            if not text or self.is_code(para) or self.is_caption(para) or self.is_heading_like(para):
                continue
            if _is_service_text(text) or self.style_name(para) in service_styles:
                continue
            result.append(para)
        return result

    # ---------- колонтитулы ----------

    def footer_root(self, section_index, kind="default"):
        for index in range(section_index, -1, -1):
            rid = self.sections[index].footer_refs.get(kind)
            if rid:
                return self._part_by_rid(rid)
        return None

    def header_root(self, section_index, kind="default"):
        for index in range(section_index, -1, -1):
            rid = self.sections[index].header_refs.get(kind)
            if rid:
                return self._part_by_rid(rid)
        return None


_ACTIVE_MODEL = None


def _model(doc):
    if _ACTIVE_MODEL is not None and _ACTIVE_MODEL[0] is doc:
        return _ACTIVE_MODEL[1]
    return _DocModel(doc)

# ========== ПРОВЕРКИ: СТРАНИЦА И ФОРМАТИРОВАНИЕ ==========

def _real_sections(model):
    return [section for section in model.sections if section.el is not None]


def check_page_size(doc: Document) -> list:
    """Проверка формата A4 во всех секциях; альбомная A4 допустима."""
    model = _model(doc)
    sections = _real_sections(model)
    if not sections:
        return [check(False, "В документе нет секций — невозможно проверить формат страницы")]

    results = []
    for index, section in enumerate(sections, 1):
        width, height = _mm(section.width), _mm(section.height)
        if width is None or height is None:
            results.append(check(
                False,
                f"Секция {index}: размер страницы не задан — Word подставит формат по умолчанию; "
                f"задай A4 ({REQUIREMENTS['page_width_mm']} × {REQUIREMENTS['page_height_mm']} мм)",
            ))
            continue
        short_side, long_side = sorted((width, height))
        is_a4 = (
            abs(short_side - REQUIREMENTS["page_width_mm"]) < 1
            and abs(long_side - REQUIREMENTS["page_height_mm"]) < 1
        )
        landscape = width > height
        suffix = " (альбомная ориентация)" if landscape else ""
        results.append(check(
            is_a4,
            f"Секция {index}: {width} × {height} мм{suffix} "
            f"(требуется A4: {REQUIREMENTS['page_width_mm']} × "
            f"{REQUIREMENTS['page_height_mm']} мм)",
            severity="error",
        ))
    return results


def check_margins(doc: Document) -> list:
    """Проверка полей всех секций. Для альбомной секции Word поворачивает поля:
    принимаются повёрнутые варианты (35/10/20/20 и 10/35/20/20) и исходный набор."""
    model = _model(doc)
    sections = _real_sections(model)
    if not sections:
        return [check(False, "В документе нет секций — невозможно проверить поля")]

    expected = {
        "Верхнее": REQUIREMENTS["top_margin_mm"],
        "Нижнее": REQUIREMENTS["bottom_margin_mm"],
        "Левое": REQUIREMENTS["left_margin_mm"],
        "Правое": REQUIREMENTS["right_margin_mm"],
    }
    results = []
    for index, section in enumerate(sections, start=1):
        left = section.left + section.gutter if section.left is not None else None
        actual = {
            "Верхнее": _mm(section.top),
            "Нижнее": _mm(section.bottom),
            "Левое": _mm(left),
            "Правое": _mm(section.right),
        }
        landscape = (section.width or 0) > (section.height or 0)
        if landscape and None not in actual.values():
            t, b, l, r = (expected[key] for key in ("Верхнее", "Нижнее", "Левое", "Правое"))
            variants = [
                {"Верхнее": l, "Нижнее": r, "Левое": b, "Правое": t},
                {"Верхнее": r, "Нижнее": l, "Левое": t, "Правое": b},
                dict(expected),
            ]
            match = next(
                (variant for variant in variants if all(abs(actual[key] - variant[key]) < 1 for key in variant)),
                None,
            )
            if match is not None:
                results.append(check(
                    True,
                    f"Секция {index} (альбомная): поля {actual['Верхнее']}/{actual['Нижнее']}/"
                    f"{actual['Левое']}/{actual['Правое']} мм соответствуют повёрнутым полям методички",
                ))
                continue
            expected_for_section = variants[0]
        else:
            expected_for_section = expected
        for label, value in actual.items():
            target = expected_for_section[label]
            if value is None:
                results.append(check(False, f"Секция {index}, {label.lower()} поле: не задано (требуется {target} мм)"))
                continue
            results.append(check(
                abs(value - target) < 1,
                f"Секция {index}, {label.lower()} поле: {value} мм (требуется {target} мм)",
            ))
    return results


def _body_runs_for_font_stats(model):
    for para in model.body:
        if model.is_code(para) or model.is_toc(para):
            continue
        for run in model.visible_runs(para):
            yield para, run


def check_fonts(doc: Document) -> list:
    """Шрифт и размер по эффективному форматированию (run → стили → docDefaults).

    Код (моноширинные шрифты, стиль VKR Listing) и таблицы не учитываются.
    """
    model = _model(doc)
    results = []
    font_counts = {}
    size_counts = {}
    heading_fonts = {}
    heading_total = 0

    for para, run in _body_runs_for_font_stats(model):
        name = model.run_font(para, run)
        if _is_monospace(name):
            continue
        weight = len(run.text.strip())
        font_counts[name] = font_counts.get(name, 0) + weight
        size = model.run_size(para, run)
        size_counts[size] = size_counts.get(size, 0) + weight
        if model.heading_level(para) is not None or model.is_chapter_heading(para) or model.structural_key(para):
            heading_total += weight
            heading_fonts[name] = heading_fonts.get(name, 0) + weight

    required = REQUIREMENTS["font_name"]
    if font_counts:
        main_font = max(font_counts, key=font_counts.get)
        results.append(check(
            main_font.casefold() == required.casefold(),
            f"Основной шрифт: {main_font}"
            + ("" if main_font.casefold() == required.casefold() else f" (требуется {required})"),
        ))
        other_fonts = [
            font for font in sorted(font_counts, key=font_counts.get, reverse=True)
            if font.casefold() != required.casefold() and font != main_font
            and font.casefold() not in SYMBOL_FONTS
        ]
        if other_fonts:
            results.append(check(
                False,
                f"В документе найдены другие шрифты: {', '.join(other_fonts[:5])}",
                severity="warning",
            ))
        wrong_heading = {
            font: weight for font, weight in heading_fonts.items()
            if font.casefold() != required.casefold() and font.casefold() not in SYMBOL_FONTS
        }
        if heading_total and sum(wrong_heading.values()) / heading_total >= 0.2:
            results.append(check(
                False,
                f"Заголовки набраны шрифтом {', '.join(sorted(wrong_heading, key=wrong_heading.get, reverse=True))} "
                f"(требуется {required}; проверь стили «Заголовок 1/2»)",
            ))

    if size_counts:
        main_size = max(size_counts, key=size_counts.get)
        results.append(check(
            abs(main_size - REQUIREMENTS["font_size_pt"]) < 0.5,
            f"Основной размер шрифта: {main_size:g} pt (требуется {REQUIREMENTS['font_size_pt']} pt)",
        ))
    return results


def _is_code_paragraph(doc: Document, paragraph) -> bool:
    """Совместимость: абзац python-docx — код (стиль листинга или моноширинный шрифт)."""
    model = _model(doc)
    para = model.para_for(paragraph._p)
    return bool(para is not None and model.is_code(para))


def _spacing_key(line, rule):
    if rule == "auto":
        ratio = (line or 240) / 240
        if abs(ratio - 1.5) < 0.05:
            return "1.5"
        if abs(ratio - 1.0) < 0.05:
            return "1.0 (single)"
        if abs(ratio - 2.0) < 0.05:
            return "2.0 (double)"
        return f"{ratio:.2f}"
    label = "точно" if rule == "exact" else "минимум"
    return f"{label} {line / 20:g} pt"


def check_line_spacing(doc: Document) -> list:
    """Междустрочный интервал основного текста (методичка МПГУ: 1,5).

    Интервал берётся из эффективного форматирования абзаца: прямое → стиль
    абзаца → basedOn → docDefaults. Титул, оглавление, код, подписи, заголовки,
    таблицы, список литературы и приложения не учитываются.
    """
    model = _model(doc)
    spacing_counts = {}
    paragraphs = model.body_text_paras()
    for para in paragraphs:
        key = _spacing_key(*model.line_spacing(para))
        spacing_counts[key] = spacing_counts.get(key, 0) + 1

    non_empty = len(paragraphs)
    if non_empty == 0:
        return [check(True, "Пустой документ — интервал не проверен", severity="info")]

    main_spacing = max(spacing_counts, key=spacing_counts.get)
    main_count = spacing_counts[main_spacing]
    main_ratio = main_count / non_empty
    if main_spacing == "1.5" and main_ratio >= 0.8:
        results = [check(
            True,
            f"Междустрочный интервал 1,5 применён в {main_count}/{non_empty} абзацах "
            f"({main_ratio*100:.0f}%) — соответствует методичке",
        )]
    elif main_spacing == "1.5":
        others = {key: value for key, value in spacing_counts.items() if key != "1.5"}
        results = [check(
            False,
            f"Междустрочный интервал 1,5 применён только в {main_ratio*100:.0f}% абзацев. "
            f"Остальные: {others}. Проверь в Word: выделить всё (Ctrl+A) → межстрочный интервал 1,5.",
            severity="warning",
        )]
    else:
        results = [check(
            False,
            f"Основной междустрочный интервал — «{main_spacing}» ({main_count} абзацев). "
            f"Методичка МПГУ требует 1,5. Выдели всё и поменяй в Word (или исправь стиль «Обычный»).",
            severity="error",
        )]
    results.extend(_check_paragraph_gaps(model, paragraphs))
    return results


def _check_paragraph_gaps(model, paragraphs) -> list:
    """Методичка P0184: «Лишние пробелы между абзацами отсутствуют» (эвристика, предупреждение)."""
    candidates = [p for p in paragraphs if not model.is_list_item(p)]
    if not candidates:
        return []
    spaced = []
    for para in candidates:
        if bool(_toggle(model._first_ppr(para, "contextualSpacing"))):
            continue
        gap = model.spacing_pt(para, "before") + model.spacing_pt(para, "after")
        if gap > 3:
            spaced.append(gap)
    if len(spaced) / len(candidates) >= 0.5:
        typical = statistics.median(spaced)
        return [check(
            False,
            f"Между абзацами основного текста есть интервал (~{typical:g} pt в {len(spaced)} из {len(candidates)} абзацев). "
            "Методичка: «Лишние пробелы между абзацами отсутствуют» — поставь интервал до и после 0 pt в стиле «Обычный».",
            severity="warning",
        )]
    return []


def _indent_candidates(model):
    result = []
    for para in model.body_text_paras():
        text = para.text.strip()
        if len(text) < 20 or model.is_list_item(para):
            continue
        if model.alignment(para) in {"center", "right"}:
            continue
        if text.startswith("(") and text.endswith(")") and len(text) < 80:
            continue
        result.append(para)
    return result


def check_first_line_indent(doc: Document) -> list:
    """Абзацный отступ 1,25 см по эффективному форматированию.

    Не учитываются титул, оглавление, заголовки, код, подписи, пункты списков
    (numPr, стили List …), центрированные и выровненные вправо абзацы, список
    литературы и приложения. Допуск 0,1 см.
    """
    model = _model(doc)
    expected = REQUIREMENTS["first_line_indent_cm"]
    correct = 0
    wrong = 0
    wrong_examples = []
    for para in _indent_candidates(model):
        indent = model.first_line_cm(para)
        if abs(indent - expected) <= 0.1:
            correct += 1
        else:
            wrong += 1
            if len(wrong_examples) < 3:
                wrong_examples.append(f"{para.text.strip()[:40]}… ({indent:.2f} см)")

    total = correct + wrong
    if total == 0:
        return [check(True, "Нет абзацев основного текста для проверки отступа", severity="info")]
    ratio = correct / total
    if ratio >= 0.85:
        return [check(True, f"Отступ первой строки 1,25 см — в {correct}/{total} абзацах ({ratio*100:.0f}%)")]
    if ratio >= 0.6:
        return [check(
            False,
            f"Отступ 1,25 см применён только в {ratio*100:.0f}% абзацев. "
            f"Примеры с неверным отступом: {', '.join(wrong_examples)}. "
            f"В Word: Ctrl+A → абзац → первая строка → 1,25 см.",
            severity="warning",
        )]
    return [check(
        False,
        f"Отступ 1,25 см НЕ соблюдён ({correct}/{total} абзацев корректны). "
        f"Примеры: {', '.join(wrong_examples)}. Методичка требует 1,25 см для красной строки.",
        severity="error",
    )]


def _check_alignment(model) -> list:
    left = []
    justified = 0
    for para in model.body_text_paras():
        text = para.text.strip()
        if len(text) < 100 or model.is_list_item(para):
            continue
        alignment = model.alignment(para)
        if alignment == "both":
            justified += 1
        elif alignment == "left":
            left.append(text[:50])
    considered = justified + len(left)
    if not left:
        return [check(True, "Основной текст выровнен по ширине", severity="info")]
    ratio = len(left) / considered if considered else 0
    message = (
        f"Выравнивание основного текста влево в {len(left)} из {considered} абзацев "
        f"(методичка: «Выравнивание основного текста работы — по ширине»). "
        f"Примеры: {'; '.join(left[:3])}"
    )
    if len(left) >= 3 and ratio >= 0.1:
        return [check(False, message, severity="error")]
    return [check(False, message, severity="warning")]


def check_font_color(doc: Document) -> list:
    """Цвет шрифта (методичка: «Цвет — чёрный») по эффективному форматированию.

    «Авто» и тематический «Текст 1» считаются чёрными; «Авто» у run перекрывает
    синий цвет стиля заголовка. Гиперссылки выносятся в предупреждение.
    """
    model = _model(doc)
    nonblack = {}
    hyperlinks = {}
    for para in model.body:
        if model.is_toc(para):
            # Word не показывает стиль «Гиперссылка» у строк поля TOC (проверено по PDF Word)
            continue
        style_name = model.styles.name(para.style_id) or "Normal"
        for run in model.visible_runs(para):
            rgb = model.run_color(para, run)
            if rgb is None:
                continue
            r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
            if r <= 30 and g <= 30 and b <= 30:
                continue
            target = hyperlinks if model.is_hyperlink_run(para, run) else nonblack
            key = (f"#{rgb}", style_name)
            target[key] = target.get(key, 0) + 1

    results = []
    if not nonblack and not hyperlinks:
        return [check(True, "Цвет шрифта везде чёрный (или близко к нему)")]
    by_color = {}
    for (hex_repr, style), count in nonblack.items():
        entry = by_color.setdefault(hex_repr, {"count": 0, "styles": []})
        entry["count"] += count
        entry["styles"].append(style)
    for hex_repr, entry in sorted(by_color.items(), key=lambda item: -item[1]["count"])[:5]:
        styles = ", ".join(f"«{style}»" for style in sorted(set(entry["styles"])))
        results.append(check(
            False,
            f"Найден не-чёрный цвет {hex_repr}: {entry['count']} run'ов "
            f"(стиль {styles}). Методичка: «Цвет — чёрный».",
            severity="error",
        ))
    if hyperlinks:
        total = sum(hyperlinks.values())
        colors = ", ".join(sorted({key[0] for key in hyperlinks}))
        results.append(check(
            False,
            f"Гиперссылки окрашены ({colors}): {total} фрагментов. Методичка требует чёрный цвет текста — "
            f"сними цвет и подчёркивание у ссылок.",
            severity="warning",
        ))
    if not nonblack:
        results.insert(0, check(True, "Цвет основного текста чёрный"))
    return results


# ========== ПРОВЕРКИ: СОДЕРЖАНИЕ И СТРУКТУРА ==========

def _section_text(model, start_idx, stop_keys=("toc", "intro", "conclusion", "bibliography")):
    parts = []
    for para in model.body[start_idx:]:
        if model.is_toc(para) or model.structural_key(para) in stop_keys or model.is_chapter_heading(para):
            break
        parts.append(model.text_without_code(para))
    return "\n".join(parts).strip()


def check_annotation_content(doc: Document) -> list:
    """Проверка содержимого аннотации (методичка МПГУ, P0034–P0040).

    В аннотации должно быть указано: название работы; сведения об объёме
    (количество страниц); количество использованных источников; перечень
    ключевых слов; сведения об объекте исследования, цели работы и её
    актуальности. Объём — не более 0,5 страницы (эвристика ~1000 знаков).
    """
    model = _model(doc)
    start = next(
        (p.idx for p in model.body if not model.is_toc(p) and model.structural_key(p) == "annotation"),
        None,
    )
    if start is None:
        return [check(True, "Аннотация не найдена (проверить через check_structure)", severity="info")]

    annotation_text = _section_text(model, start + 1)
    results = []
    if len(annotation_text) < 100:
        return [check(
            False,
            f"Аннотация слишком короткая ({len(annotation_text)} симв.) или пуста. "
            "По методичке — до 0,5 страницы с названием работы, объёмом, числом источников, "
            "ключевыми словами, объектом, целью и актуальностью.",
            severity="warning",
        )]

    if not re.search(r"\b\d{1,3}\s*(?:стр\w*|с\.)", annotation_text, re.IGNORECASE):
        results.append(check(
            False,
            "В аннотации не указан объём работы (например, «Работа изложена на 55 страницах»). "
            "Методичка требует указания количества страниц.",
            severity="warning",
        ))
    if not re.search(r"\b\d{1,3}\s*(?:источник\w*|наименован\w*)|источник\w*[\s:—–-]*\d{1,3}\b", annotation_text, re.IGNORECASE):
        results.append(check(
            False,
            "В аннотации не указано число использованных источников "
            "(например, «Список литературы включает 45 источников»).",
            severity="warning",
        ))
    if not re.search(r"ключевые\s+слова|key\s*words", annotation_text, re.IGNORECASE):
        results.append(check(
            False,
            "В аннотации нет перечня ключевых слов (методичка: «перечень ключевых слов»).",
            severity="warning",
        ))
    for label, pattern in (
        ("цель", r"\bцел(?:ь|и|ью|ей|ям|ями|ях)\b"),
        ("объект", r"\bобъект\w*"),
        ("актуальность", r"\bактуальн\w*"),
        ("название работы", r"\b(?:вкр|работ[аыеу]|исследовани[еяю])\b"),
    ):
        if not re.search(pattern, annotation_text, re.IGNORECASE):
            results.append(check(False, f"В аннотации явно не обозначено поле «{label}».", severity="warning"))

    if len(annotation_text) > 1000:
        results.append(check(
            False,
            f"Аннотация слишком длинная ({len(annotation_text)} симв. при ориентире не более 1000 симв. ≈ 0,5 страницы). "
            "Сократи до ориентира, сохранив обязательные элементы (тема, объём, число источников, ключевые слова, "
            "объект/цель/актуальность), и проверь фактическую вёрстку.",
            severity="warning",
        ))
    if not any(not item["ok"] for item in results):
        results.append(check(
            True,
            f"Аннотация содержит объём работы, число источников, ключевые слова, объект, цель и актуальность "
            f"({len(annotation_text)} симв.) — формально корректна",
        ))
    return results


def check_structure(doc: Document) -> list:
    """Обязательные структурные части: самостоятельный заголовок, а не упоминание в тексте.

    Оглавление засчитывается по заголовку «Содержание/Оглавление», полю TOC или
    автособираемому оглавлению Word (w:sdt). Строки оглавления заголовками не считаются.
    """
    model = _model(doc)
    found_keys = set()
    for para in model.body:
        key = model.structural_key(para)
        if key == "toc" or (key and not model.is_toc(para)):
            found_keys.add(key)
    if any(model.is_toc(p) and (p.in_toc_field or p.in_toc_sdt) for p in model.body):
        found_keys.add("toc")
    has_chapter = bool(model.chapter_headings()) or any(model.is_chapter_heading(p) for p in model.body)

    required = (
        ("Аннотация", "annotation" in found_keys),
        ("Содержание / Оглавление", "toc" in found_keys),
        ("Введение", "intro" in found_keys),
        ("Хотя бы одна Глава", has_chapter),
        ("Заключение", "conclusion" in found_keys),
        ("Список литературы", "bibliography" in found_keys),
    )
    return [
        check(found, f"Раздел '{title}' — {'найден' if found else 'НЕ НАЙДЕН'}", severity="info" if found else "error")
        for title, found in required
    ]


def _main_text_chunks(model, include_appendix=False):
    chunks = []
    for para in model.body:
        region = model.region(para)
        if region == "main" or (include_appendix and region == "appendix"):
            text = model.text_without_code(para)
            if text.strip():
                chunks.append(text)
    return chunks


def check_personal_pronouns(doc: Document) -> list:
    """Личные местоимения в тексте работы.

    Перед проверкой удаляются инициалы («И.Я.», «Я.Л.»): иначе \\bя\\b ловит
    «Я» в инициалах педагогов. Титул, оглавление, код и список литературы не
    проверяются; в приложениях (анкеты, стенограммы) находка — предупреждение.
    """
    model = _model(doc)
    initial_pattern = re.compile(r"\b[А-ЯЁA-Z]\.\s?[А-ЯЁA-Z]\.(?:\s?[А-ЯЁA-Z]\.)?")
    single_initial_pattern = re.compile(r"\b[А-ЯЁA-Z]\.")

    def scan(chunks):
        text = single_initial_pattern.sub(" ", initial_pattern.sub(" ", "\n".join(chunks)))
        found = []
        for pattern in PERSONAL_PRONOUNS:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                label = pattern.replace(r"\b", "").strip("^$\\")
                found.append(f"{label}: {len(matches)} раз")
        return found

    main_found = scan(_main_text_chunks(model))
    appendix_chunks = [model.text_without_code(p) for p in model.body if model.region(p) == "appendix"]
    appendix_found = scan(appendix_chunks)
    results = []
    if main_found:
        results.append(check(
            False,
            f"Найдены ЗАПРЕЩЁННЫЕ местоимения: {', '.join(main_found)}. Замени на безличные конструкции!",
        ))
    else:
        results.append(check(True, "Запрещённых личных местоимений не найдено"))
    if appendix_found:
        results.append(check(
            False,
            f"В приложениях есть личные местоимения: {', '.join(appendix_found)}. Допустимо в анкетах и "
            "цитируемых материалах; авторский текст перепиши безлично.",
            severity="warning",
        ))
    return results


def check_ai_cliches(doc: Document) -> list:
    """Проверка признаков ИИ-текста: характерные клише.

    Фразы из персонального CLICHE_ALLOWLIST (state-файл, поле cliche_allowlist)
    не учитываются.
    """
    model = _model(doc)
    full_text = "\n".join(_main_text_chunks(model, include_appendix=True)).lower()
    effective_cliches = [c for c in AI_CLICHES if c not in CLICHE_ALLOWLIST]
    found = []
    total_count = 0
    for cliche in effective_cliches:
        count = full_text.count(cliche)
        if count > 0:
            found.append((cliche, count))
            total_count += count

    results = []
    if CLICHE_ALLOWLIST:
        results.append(check(
            True,
            f"Активен персональный cliche_allowlist ({len(CLICHE_ALLOWLIST)} фраз "
            f"из state-файла) — эти клише не учитываются.",
            severity="info",
        ))
    top = ", ".join(f"«{c}» ({n})" for c, n in sorted(found, key=lambda x: -x[1])[:5])
    if total_count == 0:
        results.append(check(True, "ИИ-клише не найдены"))
    elif total_count <= 3:
        results.append(check(True, f"ИИ-клише встречаются редко ({total_count} шт. всего). Топ: {top}. Это в пределах нормы.", severity="info"))
    elif total_count <= 8:
        results.append(check(False, f"Повышенное содержание ИИ-клише: {total_count} шт. Топ: {top}. Рекомендуется сократить.", severity="warning"))
    else:
        results.append(check(
            False,
            f"КРИТИЧНО: много ИИ-клише ({total_count} шт.). Топ: {top}. Работа воспринимается как ИИ-текст. "
            f"Переработай начала абзацев (см. humanizer-techniques.md).",
        ))
    return results


def check_paragraph_rhythm(doc: Document) -> list:
    """Проверка ритма абзацев — не слишком ли одинаковые по длине."""
    model = _model(doc)
    paragraphs = [text for text in _main_text_chunks(model) if len(text) > 50]
    if len(paragraphs) < 10:
        return [check(True, "Слишком мало абзацев для анализа ритма", severity="info")]
    lengths = [len(p) for p in paragraphs]
    mean = statistics.mean(lengths)
    stdev = statistics.stdev(lengths) if len(lengths) > 1 else 0
    cv = stdev / mean if mean > 0 else 0
    if cv < 0.25:
        return [check(
            False,
            f"Слишком ровный ритм абзацев (CV={cv:.2f}, требуется >0.3). "
            f"Средняя длина {mean:.0f} символов, разброс маленький. "
            f"Признак ИИ-текста! Разбей или объедини некоторые абзацы, чтобы было разнообразие.",
            severity="warning",
        )]
    if cv < 0.35:
        return [check(
            True,
            f"Ритм абзацев приемлемый, но монотонный (CV={cv:.2f}). "
            f"Можно улучшить, разбавив короткими (2 предложения) и длинными (7-8 предложений) абзацами.",
            severity="info",
        )]
    return [check(True, f"Хороший ритм абзацев (CV={cv:.2f})")]


def check_consecutive_same_starts(doc: Document) -> list:
    """Проверка: нет ли подряд абзацев с одинаковыми первыми словами."""
    model = _model(doc)
    paragraphs = [text.strip() for text in _main_text_chunks(model) if len(text) > 50]
    consecutive_count = 0
    max_consecutive = 0
    prev_start = ""
    for para in paragraphs:
        start = " ".join(para.split()[:2]).lower()
        if start == prev_start and start:
            consecutive_count += 1
            max_consecutive = max(max_consecutive, consecutive_count)
        else:
            consecutive_count = 1
        prev_start = start
    if max_consecutive >= 3:
        return [check(
            False,
            f"Найдено {max_consecutive} абзацев подряд, начинающихся одинаково. "
            f"Это признак шаблонного текста. Варьируй начала абзацев.",
            severity="warning",
        )]
    return [check(True, "Начала абзацев разнообразны")]


# ========== ССЫЛКИ НА ИСТОЧНИКИ ==========

# Квадратные скобки, перед которыми нет буквы, цифры, «_», «]» или «)»:
# arr[25], matrix[n][30] и f(x)[1] — индексы, а не ссылки.
CITATION_BRACKET_RE = re.compile(r"(?<![\w\])])\[([^\[\]\n]{1,200})\]")
CITATION_PREFIX_RE = re.compile(
    r"^\s*(?:цит\.?\s*по(?:\s*(?:кн|ст)\.?)?|см\.?(?:\s*также)?|ср\.?|подробнее\s+см\.?|"
    r"cit\.?\s*(?:by|from|in)?|see(?:\s+also)?|cf\.?)\s*:?\s*",
    re.IGNORECASE,
)


def _parse_citation(content: str) -> list:
    """Номера источников из содержимого скобок; номера страниц и тома не считаются."""
    numbers = []
    for part in content.split(";"):
        part = CITATION_PREFIX_RE.sub("", part, count=1).strip()
        match = re.match(r"^(\d{1,3})(?:\s*[-–—]\s*(\d{1,3}))?(?=$|[\s,])", part)
        if not match:
            continue
        first = int(match.group(1))
        if match.group(2):
            last = int(match.group(2))
            if first < last <= first + 20:
                numbers.extend(range(first, last + 1))
            else:
                numbers.append(first)
        else:
            numbers.append(first)
        rest = part[match.end():]
        tokens = rest.split(",")
        for index, token in enumerate(tokens):
            token = token.strip()
            if index == 0 and not token:
                continue
            if re.fullmatch(r"\d{1,3}", token):
                numbers.append(int(token))
                continue
            break
    return [n for n in numbers if 1 <= n <= 500]


def _iter_citations(text: str):
    """(match, [номера]) для каждой ссылки в тексте."""
    for match in CITATION_BRACKET_RE.finditer(text or ""):
        numbers = _parse_citation(match.group(1))
        if numbers:
            yield match, numbers


def _extract_citation_numbers(text: str) -> list:
    """Извлечь номера источников из всех типов ссылок.

    Покрывает: [15], [15, с. 23], [15, с. 23–25], [15, p. 23], [15, т. 2, с. 45],
    [15; 17], [15, 17, 19], [Цит. по: 27, с. 235], [См.: 12; 13].
    Номера страниц, томов и частей номерами источников не считаются;
    индексы кода (arr[25], matrix[n][30]) не считаются ссылками.
    Возвращает список всех номеров (с повторениями) в порядке появления.
    """
    numbers = []
    for _match, found in _iter_citations(text):
        numbers.extend(found)
    return numbers


def _citation_texts(model):
    """Тексты, где ссылки [N] считаются упоминанием источника: основной текст,
    приложения, таблицы, сноски (без кода, оглавления и самого списка литературы)."""
    chunks = []
    for para in model.body:
        if model.region(para) in {"main", "appendix", "other"}:
            chunks.append(model.text_without_code(para))
    for para in model.table_paras + model.floating_paras:
        chunks.append(model.text_without_code(para))
    for para in model.part_paras:
        if para.container == "note":
            chunks.append(model.text_without_code(para))
    return chunks


def check_citations(doc: Document) -> list:
    """Наличие ссылок в тексте и отсутствие ссылок в заключении (методичка P0016)."""
    model = _model(doc)
    citations = _extract_citation_numbers("\n".join(_citation_texts(model)))
    results = []
    if len(citations) < 20:
        results.append(check(
            True,
            f"В тексте найдено {len(citations)} ссылок. "
            f"Это ниже рабочего ориентира скилла (20+), но методички ИМО не "
            f"задают формального минимума. Сверь число с требованиями кафедры и "
            f"проверь, что каждый источник из списка упомянут в тексте.",
            severity="info",
        ))
    else:
        results.append(check(
            True,
            f"Ссылок в тексте: {len(citations)} (уникальных источников: {len(set(citations))})",
            severity="info",
        ))

    conclusion_citations = 0
    if model.conclusion_idx is not None:
        for para in model.body[model.conclusion_idx + 1:]:
            if model.is_toc(para):
                continue
            if model.structural_key(para) or model.is_appendix_label(para) or model.region(para) != "main":
                break
            conclusion_citations += len(_extract_citation_numbers(model.text_without_code(para)))
    if conclusion_citations > 0:
        results.append(check(
            False,
            f"В заключении найдено {conclusion_citations} ссылок. "
            "По методичке МПГУ в заключении ссылок и цитат БЫТЬ НЕ ДОЛЖНО!",
        ))
    else:
        results.append(check(True, "В заключении нет ссылок — хорошо"))
    return results


SURNAME_NOT_SURNAMES = {
    "теоретическая", "теоретический", "теоретическое", "практическая", "практический", "практическое",
    "аналитическая", "аналитический", "аналитическое", "эмпирическая", "эмпирический", "эмпирическое",
    "общественная", "общественный", "общественное", "педагогическая", "педагогический", "педагогическое",
    "психологическая", "психологический", "психологическое", "методическая", "методический", "методическое",
    "техническая", "технический", "техническое", "электронная", "электронный", "электронное",
    "российская", "российский", "российское", "статистическая", "статистический", "статистическое",
    "педагогических", "практических", "теоретических", "исследовательская", "исследовательский",
    "информационная", "информационный", "информационное", "московский", "московская", "московское",
    "петербургский", "петербургская", "санкт-петербургский", "государственный", "государственная",
    "государственное", "федеральный", "федеральная", "федеральное", "университетский", "университетская",
    "городской", "городская", "городское", "данный", "данная", "данное", "данные", "данных",
    "самый", "самая", "самое", "описательный", "описательная", "сравнительный", "сравнительная",
    "деятельностный", "деятельностная", "системный", "системная", "системное", "системные", "системных",
    "когнитивный", "когнитивная", "комплексный", "комплексная", "комплексное", "комплексные",
    "индустриальный", "индустриальная", "индустриальное", "культурный", "культурная", "культурное",
    "традиционный", "традиционная", "традиционное", "функциональный", "функциональная", "функциональное",
    "оперативный", "оперативная", "оперативное", "учебный", "учебная", "учебное", "служебный", "служебная",
    "служебное", "организационный", "организационная", "организационное", "финансовый", "финансовая",
    "финансовое", "социальный", "социальная", "социальное", "экологический", "экологическая", "экологическое",
    "цифровой", "цифровая", "цифровое", "цифровые", "цифровых", "структурный", "структурная", "структурное",
    "европейский", "европейская", "европейское", "православный", "православная", "качественный",
    "качественная", "качественное", "количественный", "количественная", "количественное", "современный",
    "современная", "современное", "современных", "научный", "научная", "научное", "научных",
    "образовательный", "образовательная", "образовательное", "нормативный", "нормативная", "нормативное",
    "типовой", "типовая", "типовое", "основной", "основная", "основное", "основные", "основных",
    "дополнительный", "дополнительная", "дополнительное", "внутренний", "внутренняя", "внутреннее",
    "внешний", "внешняя", "внешнее", "программный", "программная", "программное", "сетевой", "сетевая",
    "сетевое", "мобильный", "мобильная", "мобильное",
}


def check_surnames_format(doc: Document) -> list:
    """Фамилии в тексте работы должны быть с инициалами перед фамилией через пробел.

    Титульный лист (полное ФИО автора и руководителя), оглавление, список
    литературы, последний лист, код и сноски не проверяются.
    """
    model = _model(doc)
    full_text = "\n".join(_main_text_chunks(model, include_appendix=True))
    surname_pattern = re.compile(
        r"\b([А-ЯЁ][а-яё]{2,})(ов|ова|ёв|ёва|ев|ева|ин|ина|ский|ская|цкий|цкая)\b"
    )
    surnames_without_initials = []
    malformed_initials = []
    # Слово с заглавной буквы в начале абзаца или предложения («Клиентская часть — …») фамилией не
    # считается, если основа того же слова нигде больше не встречается как фамилия: с инициалами или
    # с заглавной буквы в середине предложения.
    surname_stems = set()
    sentence_start_candidates = []
    for match in surname_pattern.finditer(full_text):
        if match.group(0).lower() in SURNAME_NOT_SURNAMES:
            continue
        stem = match.group(1)
        start = max(0, match.start() - 15)
        context_before = full_text[start:match.start()]
        if re.search(r"[А-ЯЁ]\.\s*[А-ЯЁ]\.\s+$", context_before):
            surname_stems.add(stem)
            continue
        after = full_text[match.end():match.end() + 10]
        if re.search(r"^\s*[А-ЯЁ]\.\s*[А-ЯЁ]\.", after):
            malformed_initials.append(match.group(0) + " + инициалы после фамилии")
            surname_stems.add(stem)
            continue
        if re.search(r"[А-ЯЁ]\.\s*[А-ЯЁ]\.$", context_before):
            malformed_initials.append("инициалы без пробела + " + match.group(0))
            surname_stems.add(stem)
            continue
        if re.search(r"(?<![А-ЯЁа-яёA-Za-z])[А-ЯЁ]\.\s*$", context_before):
            surname_stems.add(stem)  # один инициал перед фамилией
            continue
        line_start = full_text.rfind("\n", 0, match.start()) + 1
        raw_prefix = full_text[line_start:match.start()].rstrip()
        prefix = raw_prefix.rstrip("«\"„“'([—–-•·*").rstrip()
        # Начало абзаца, пункта списка, предложения или цитаты в кавычках.
        if not prefix or prefix[-1] in ".!?…" or raw_prefix[-1:] in "«„“\"":
            sentence_start_candidates.append((match.group(0), stem))
            continue
        surname_stems.add(stem)
        surnames_without_initials.append(match.group(0))
    surnames_without_initials.extend(word for word, stem in sentence_start_candidates if stem in surname_stems)

    unique = sorted(set(surnames_without_initials))[:10]
    malformed = sorted(set(malformed_initials))[:10]
    if unique or malformed:
        details = []
        if unique:
            details.append("возможно без инициалов: " + ", ".join(unique))
        if malformed:
            details.append("неверный порядок/пробел: " + ", ".join(malformed))
        return [check(
            False,
            "Оформление фамилий требует проверки: " + "; ".join(details) + ". "
            "По методичке инициалы должны стоять ПЕРЕД фамилией через пробел. Проверь вручную!",
            severity="warning",
        )]
    return [check(True, "Фамилии в целом оформлены с инициалами")]


def _looks_like_chapter_heading(para, text: str) -> bool:
    """Совместимость: абзац python-docx — заголовок главы (стиль Heading, отдельный
    полужирный/центрированный абзац), а не строка оглавления или фраза «Глава II посвящена…»."""
    document = getattr(getattr(para, "part", None), "document", None)
    if document is None:
        return False
    model = _model(document)
    item = model.para_for(para._p)
    return bool(item is not None and model.is_chapter_heading(item))


def _chapters(model):
    """[(заголовок, [абзацы главы])] до следующей главы или структурного раздела."""
    headings = model.chapter_headings()
    chapters = []
    for number, heading in enumerate(headings):
        stop = headings[number + 1].idx if number + 1 < len(headings) else model.main_end
        if model.conclusion_idx is not None and heading.idx < model.conclusion_idx < stop:
            stop = model.conclusion_idx
        paras = [
            p for p in model.body[heading.idx + 1:stop]
            if not model.is_toc(p) and not model.structural_key(p)
        ]
        chapters.append((heading, paras))
    return chapters


def check_chapter_lengths(doc: Document, profile: str = "generic") -> list:
    """Длина глав и общий объём (эвристика ~1800 знаков на страницу)."""
    model = _model(doc)
    chapters = _chapters(model)
    if not chapters:
        return [check(True, "Главы не найдены (возможно, нестандартные заголовки)", severity="info")]

    chars_per_page = 1800
    results = []
    total_chars = 0
    for number, (heading, paras) in enumerate(chapters, 1):
        chars = sum(len(model.text_without_code(p).strip()) for p in paras)
        total_chars += chars
        pages = chars / chars_per_page
        title = heading.text.strip().replace("\n", " ")[:40]
        if pages < 5:
            results.append(check(False, f"Глава {number} «{title}...»: ~{pages:.1f} стр. — очень мало по эвристике баланса. "
                                        f"Проверь, раскрыта ли функция главы и согласована ли такая структура с научным руководителем.", severity="warning"))
        elif pages < 8:
            results.append(check(False, f"Глава {number} «{title}...»: ~{pages:.1f} стр. — ниже рабочего ориентира 8 стр. "
                                        f"Это не формальный порог допуска: оцени полноту главы по её задачам и профилю программы.", severity="warning"))
        elif pages > 25:
            results.append(check(False, f"Глава {number} «{title}...»: ~{pages:.1f} стр. — слишком много для рабочего ориентира "
                                        f"баланса (12-20 стр.); проверь структуру по плану", severity="warning"))
        else:
            results.append(check(True, f"Глава {number}: ~{pages:.1f} стр. (в рабочем диапазоне эвристики)", severity="info"))

    main_chars = sum(len(model.text_without_code(p).strip()) for p in model.main_paras())
    estimated_pages = main_chars / chars_per_page + 6  # титул, аннотация, содержание, список литературы
    low, high = PROFILE_PAGE_RANGES.get(profile, (35, 65))
    in_range = low * 0.8 <= estimated_pages <= high * 1.2
    results.append(check(
        in_range,
        f"Оценка объёма работы: ~{estimated_pages:.0f} стр. (главы ~{total_chars / chars_per_page:.1f} стр.); "
        f"ориентир профиля {profile}: {low}–{high} стр. Точный объём смотри в Word — оценка по знакам приблизительная.",
        severity="info" if in_range else "warning",
    ))
    return results


def _caption_numbers(model, kind):
    pattern = {
        "table": re.compile(r"^\s*таблица\s+((?:[А-ЯA-Z]\.)?\d+(?:\.\d+)*)", re.IGNORECASE),
        "figure": re.compile(r"^\s*(?:рисунок|рис\.)\s+((?:[А-ЯA-Z]\.)?\d+(?:\.\d+)*)", re.IGNORECASE),
    }[kind]
    numbers = []
    for para in model.body + model.table_paras:
        if para.container == "body" and model.is_toc(para):
            continue
        if not model.is_caption(para):
            continue
        match = pattern.match(para.text.strip())
        if match:
            numbers.append(match.group(1))
    return numbers


def _reference_text(model):
    chunks = []
    for para in model.body:
        if model.is_toc(para) or model.is_caption(para):
            continue
        chunks.append(model.text_without_code(para))
    for para in model.table_paras + model.part_paras:
        if para.container in {"table", "note"} and not model.is_caption(para):
            chunks.append(model.text_without_code(para))
    return "\n".join(chunks)


def check_figure_table_references(doc: Document) -> list:
    """Все ли таблицы и рисунки упомянуты в тексте.

    Подписью считается абзац «Таблица N» / «Таблица N — Название» / «Рисунок N. …»
    или стиль подписи; предложение «Таблица 1 показывает, что…» — упоминание.
    """
    model = _model(doc)
    results = []
    unique_tables = set(_caption_numbers(model, "table"))
    unique_figures = set(_caption_numbers(model, "figure"))
    reference_text = _reference_text(model)
    number = r"((?:[А-ЯA-Z]\.)?\d+(?:\.\d+)*)"
    table_refs = set(re.findall(r"\bтабл(?:ица|ицы|ице|ицу|ицей|ицах|ицам|ицами|\.)?\s*" + number, reference_text, re.IGNORECASE))
    figure_refs = set(re.findall(r"\bрис(?:унок|унка|унке|унку|унком|унки|унков|унках|унками|\.)?\s*" + number, reference_text, re.IGNORECASE))

    if unique_tables:
        missing = unique_tables - table_refs
        if missing:
            results.append(check(
                False,
                f"Таблицы {sorted(missing)} не упомянуты в тексте. "
                f"Каждая таблица должна иметь ссылку в тексте («в таблице 1», «см. табл. 1»).",
                severity="warning",
            ))
        else:
            results.append(check(True, f"Все {len(unique_tables)} таблиц(ы) упомянуты в тексте"))
    if unique_figures:
        missing = unique_figures - figure_refs
        if missing:
            results.append(check(
                False,
                f"Рисунки {sorted(missing)} не упомянуты в тексте. "
                f"Каждый рисунок должен иметь ссылку в тексте («на рисунке 1», «см. рис. 1»).",
                severity="warning",
            ))
        else:
            results.append(check(True, f"Все {len(unique_figures)} рисунок(ов) упомянуты в тексте"))
    if not unique_tables and not unique_figures:
        results.append(check(True, "Таблиц и рисунков не найдено (для IT-ВКР странно — обычно есть)", severity="info"))
    return results

TOC_PLACEHOLDER_RE = re.compile(
    r"для обновления содержания|щ[её]лкните правой|нет элементов оглавления|no table of contents entries",
    re.IGNORECASE,
)


def check_toc_filled(doc: Document, profile: str = "generic") -> list:
    """Оглавление реально заполнено (не остался плейсхолдер генератора).

    Оглавление — поле TOC, автособираемое оглавление Word (w:sdt) или строки со
    стилями «Оглавление N»; заглушка ищется только внутри этой области, а не по
    всему тексту. По методичке (P0013) «Аннотацию» и само «Содержание» в
    оглавление не вносят: список начинается с «Введения».
    """
    model = _model(doc)
    toc_paras = [p for p in model.body if model.is_toc(p)]
    heading = next((p for p in model.body if model.structural_key(p) == "toc"), None)
    if not toc_paras and heading is None:
        return [check(True, "Содержание не найдено — проверка пропущена", severity="info")]

    if not toc_paras and heading is not None:
        # Ручное оглавление после заголовка: строки до следующего раздела.
        for para in model.body[heading.idx + 1:]:
            if model.structural_key(para) or model.is_chapter_heading(para):
                break
            if para.text.strip():
                toc_paras.append(para)

    toc_lines = [p.text.strip() for p in toc_paras if p.text.strip() and model.style_name(p) not in {"toc heading", "заголовок оглавления"}]
    toc_lines = [line for line in toc_lines if _heading_key(line) not in STRUCTURE_KEYS["toc"]]
    toc_text = "\n".join(toc_lines)
    results = []
    if TOC_PLACEHOLDER_RE.search(toc_text):
        results.append(check(
            False,
            "Оглавление НЕ обновлено — остался плейсхолдер «[Для обновления содержания…]». "
            "Открой файл в Word, щёлкни правой кнопкой по оглавлению → «Обновить поле» → "
            "«Обновить целиком» (или F9). Без этого работа уйдёт на нормоконтроль с пустым "
            "оглавлением.",
            severity="error",
        ))
        return results
    entries = [line for line in toc_lines if TOC_LINE_RE.search(line)]
    if len(toc_text) < 50:
        results.append(check(
            False,
            f"Оглавление подозрительно пустое (всего {len(toc_text)} символов). "
            "Проверь в Word: должны быть строки с заголовками и номерами страниц.",
            severity="warning",
        ))
    elif len(entries) >= 3:
        results.append(check(True, f"Оглавление заполнено (найдено {len(entries)} строк с номерами страниц)"))
    else:
        results.append(check(
            False,
            "Оглавление заполнено текстом, но без типичных номеров страниц в конце строк. "
            "Проверь в Word — возможно, нужен F9 для автогенерации.",
            severity="warning",
        ))

    if entries:
        first_key = _heading_key(re.split(r"\t|\.{3,}|…{2,}", entries[0])[0])
        extra = [
            label for label, keys in (("Аннотация", STRUCTURE_KEYS["annotation"]), ("Содержание/Оглавление", STRUCTURE_KEYS["toc"]))
            if any(_heading_key(re.split(r"\t|\.{3,}|…{2,}", line)[0]) in keys for line in entries)
        ]
        if extra:
            results.append(check(
                False,
                f"В оглавление внесены: {', '.join(extra)}. Методичка: «не нужно в список оглавления вносить "
                f"аннотацию и оглавление! Список должен начинаться с Введения». Убери у этих заголовков стиль "
                f"«Заголовок 1» или исключи их из поля TOC.",
                severity="error" if _is_mpgu(profile) else "warning",
            ))
        elif first_key not in STRUCTURE_KEYS["intro"] and _is_mpgu(profile):
            results.append(check(
                False,
                f"Оглавление начинается не с «Введения», а с «{entries[0][:40]}».",
                severity="warning",
            ))
    return results


def check_chapter_conclusions(doc: Document) -> list:
    """В каждой главе есть «Выводы по Главе X» (конвенция кафедр; предупреждение)."""
    model = _model(doc)
    chapter_count = len(model.chapter_headings())
    if chapter_count == 0:
        return [check(True, "Главы не найдены — проверка выводов пропущена", severity="info")]
    full_text = "\n".join(_main_text_chunks(model)).lower()
    conclusions_re = re.compile(
        r"\bвыводы\s+по\s+(?:главе\s+[ivxlcm\d]+|первой\s+главе|второй\s+главе|третьей\s+главе)\b",
        re.IGNORECASE,
    )
    matches = conclusions_re.findall(full_text)
    if len(matches) >= chapter_count:
        return [check(True, f"Найдено {len(matches)} блоков «Выводы по Главе» на {chapter_count} глав(ы) — достаточно")]
    if matches:
        return [check(
            False,
            f"Найдено {len(matches)} блоков «Выводы по Главе», а глав — {chapter_count}. "
            f"Если утверждённый план или научрук требуют выводы по каждой главе, добавь недостающие блоки.",
            severity="warning",
        )]
    return [check(
        False,
        f"Ни одного блока «Выводы по Главе X» не найдено (глав в работе: {chapter_count}). "
        f"Это общеакадемическая конвенция (методичка МПГУ напрямую не требует, "
        f"но практически все кафедры её ожидают). Добавь короткий блок «Выводы по Главе N» "
        f"в конце каждой главы — 3-5 тезисов на 0.5 страницы.",
        severity="warning",
    )]


def check_pilot_sample_size(doc: Document, profile: str = "generic") -> list:
    """Эвристически проверяет размер выборки пилота.

    В базовой методичке порог 5 человек относится к частной образовательной
    практике. Для других пилотов минимум задаёт программа или метод проверки.
    """
    model = _model(doc)
    full_text = " ".join(_main_text_chunks(model, include_appendix=True))
    results = []
    private_practice = bool(re.search(r"частн\w*\s+практик|репетитор", full_text, re.IGNORECASE))
    people = r"(?:человек|участник|респондент|школьник|студент|обучающ|учащ|пользовател)"
    patterns = [
        r"(?:частн\w*\s+практик|репетитор)[^\d\n]{0,200}?(\d{1,3})\s*" + people,
        r"(?:пилот\w*|апробаци\w+|эксперимент\w*)[^\d\n]{0,120}?(\d{1,3})\s*" + people,
        r"(\d{1,3})\s*" + people + r"[^\d\n]{0,120}?(?:пилот\w*|апробаци\w+|эксперимент\w*)",
        r"выборк\w+[^\d\n]{0,30}?(\d{1,3})\s*(?:человек|участник|респондент|школьник|студент|обучающ|учащ)",
    ]
    found_numbers = []
    for pattern in patterns:
        for match in re.finditer(pattern, full_text, re.IGNORECASE):
            number = int(match.group(1))
            if 1 <= number <= 500:
                found_numbers.append(number)

    if not found_numbers:
        if profile == "mpgu-09-project":
            return [check(
                False,
                "Для проектной ВКР 09.03.02 не найдено конкретное число участников "
                "пилота. Укажи, где, когда и с кем проведён тестовый запуск. Для "
                "частной образовательной практики требуется не менее 5 участников; "
                "для других баз методичка числовой минимум не задаёт.",
                severity="warning",
            )]
        return [check(
            True,
            "Конкретное число участников пилота не найдено — проверка пропущена. "
            "Для частной образовательной практики базовая методичка требует не "
            "менее 5 человек; в остальных случаях проверь профиль программы.",
            severity="info",
        )]
    min_n = min(found_numbers)
    if private_practice and min_n < 5:
        results.append(check(
            False,
            f"Для пилота в частной практике обнаружена выборка {min_n} человек. "
            f"Базовая методичка требует не менее 5 участников для этого случая.",
            severity="error",
        ))
    elif min_n < 5:
        results.append(check(
            True,
            f"Выборка пилота: {min_n} человек. Универсальный минимум методичкой "
            f"не задан; обоснуй размер методом проверки и сверь с кафедрой.",
            severity="info",
        ))
    else:
        results.append(check(True, f"Выборка пилота: {min_n}+ человек; числовой порог сверь с профилем программы", severity="info"))
    return results


def check_pedagogical_frame(doc: Document, profile: str = "generic") -> list:
    """Проверяет образовательную рамку с учётом выбранного профиля."""
    model = _model(doc)
    full_text = "\n".join(_main_text_chunks(model, include_appendix=True)).lower()
    pedagogical_markers = [
        "целевая аудитория", "учащи", "обучающи", "школьник", "студент", "преподавател",
        "дидактическ", "образовательн", "учебн", "апробаци", "методическ", "педагогическ",
        "дистанционн", "пре-пост", "пост-тест", "пре-тест", "усвоени",
    ]
    found = [marker for marker in pedagogical_markers if marker in full_text]
    if len(found) >= 5:
        return [check(True, f"Педагогический контекст есть ({len(found)} маркеров: {', '.join(found[:5])}…)")]
    if len(found) >= 2:
        return [check(
            profile != "mpgu-09-project",
            f"Есть отдельные образовательные маркеры ({', '.join(found)}). "
            f"Для проектной ВКР 09.03.02 должны быть явно описаны образовательный "
            f"продукт, целевая аудитория, методические основания и результаты пилота.",
            severity="warning" if profile == "mpgu-09-project" else "info",
        )]
    return [check(
        profile != "mpgu-09-project",
        "Образовательный контекст не выражен. В профиле mpgu-09-project "
        "методичка 09.03.02 требует образовательный продукт, целевую аудиторию "
        "и методическое обоснование. Для generic и mpgu-09-regular эта "
        "эвристика не создаёт ошибку.",
        severity="error" if profile == "mpgu-09-project" else "info",
    )]


def check_mpgu_09_profile(doc: Document, profile: str = "generic") -> list:
    """Проверяет специальные требования методички 09.03.02 к проектной ВКР."""
    if profile == "generic":
        return [check(True, "Универсальный профиль: специальные требования 09.03.02 не применены", severity="info")]
    if profile == "mpgu-09-regular":
        return [check(
            True,
            "Обычная ВКР 09.03.02: фиксированное число глав и обязательный пилот "
            "методичкой не установлены",
            severity="info",
        )]
    model = _model(doc)
    chapter_headings = model.chapter_headings()
    results = [check(
        len(chapter_headings) == 3,
        f"Обнаружено глав: {len(chapter_headings)}; для проектной ВКР 09.03.02 "
        "методичка задаёт три главы: теоретическую, аналитическую и проектную. "
        "Иная структура допустима только по утверждённому плану научного руководителя.",
        severity="error",
    )]
    full_text = "\n".join(_main_text_chunks(model, include_appendix=True)).lower()
    pilot_found = bool(re.search(r"\b(?:пилот\w*|апробаци\w*|тестов\w*\s+внедрен\w*)\b", full_text, re.IGNORECASE))
    results.append(check(
        pilot_found,
        "Описание пилотирования или апробации найдено"
        if pilot_found
        else "Для проектной ВКР 09.03.02 не найдено описание пилотирования, "
             "апробации или тестового внедрения. Технические тесты могут дополнять "
             "пилот, но не заменяют его без согласованного исключения.",
        severity="error",
    ))
    return results


def check_tables(doc: Document) -> list:
    """Оформление таблиц по методичке МПГУ (эвристики, предупреждения).

    «Таблица N» — отдельной подписью над таблицей, без точки после номера;
    каждая таблица упомянута в тексте. Служебные таблицы клятвы и титула не считаются.
    """
    model = _model(doc)
    oath_marker = "выполнена мной совершенно самостоятельно"
    dash_marker = "___________________"
    tables = []
    for block in model.blocks:
        if block[0] != "tbl":
            continue
        cells = [p.text.strip().casefold() for p in block[2]]
        total_cells = len(cells) or 1
        service = sum(1 for text in cells if oath_marker in text or dash_marker in text)
        if cells and service / total_cells >= 0.6:
            continue
        following = next((b[1] for b in model.blocks[model.blocks.index(block):] if b[0] == "p"), None)
        if following is not None and model.region(following) in {"title", "last_page"}:
            continue
        tables.append(block)
    total = len(tables)
    if total == 0:
        return [check(True, "Содержательных таблиц в работе не найдено (в большинстве ВКР-проектов ожидается 1-3 таблицы)", severity="info")]

    captions = [p for p in model.body if not model.is_toc(p) and model.is_caption(p) and re.match(r"^\s*таблица\b", p.text.strip(), re.IGNORECASE)]
    results = []
    if len(captions) < total:
        results.append(check(
            False,
            f"Найдено {total} таблиц в документе, но меток «Таблица N» только {len(captions)}. "
            f"Каждая таблица должна быть подписана меткой «Таблица N» в правом верхнем углу и упомянута в тексте.",
            severity="warning",
        ))
    else:
        results.append(check(True, f"Найдено {total} таблиц и {len(captions)} меток «Таблица N» — согласовано"))
    dotted = [p.text.strip() for p in captions if re.match(r"^\s*таблица\s+(?:[А-ЯA-Z]\.)?\d+(?:\.\d+)*\.(?:\s|$)", p.text.strip(), re.IGNORECASE)]
    if dotted:
        results.append(check(
            False,
            f"После номера таблицы не должна стоять точка (найдено {len(dotted)} случаев). "
            f"Правильно: «Таблица 1» (без точки), заголовок на следующей строке.",
            severity="warning",
        ))
    reference_text = _reference_text(model)
    mentions = re.findall(
        r"\b(?:в\s+таблице|из\s+таблицы|согласно\s+таблице|данные\s+таблицы|как\s+видно\s+из\s+таблицы|"
        r"см\.\s*табл\.?|в\s+табл\.?|таблица\s+\d+\s+(?:показывает|содержит|демонстрирует|иллюстрирует|отражает|представляет))\s*\d*",
        reference_text,
        re.IGNORECASE,
    )
    if not mentions:
        results.append(check(
            False,
            "В тексте работы нет явных упоминаний таблиц («см. табл. 1», «в таблице 2 представлено...»). "
            "Каждая таблица должна быть упомянута в основном тексте.",
            severity="warning",
        ))
    return results


# ========== СПИСОК ЛИТЕРАТУРЫ ==========

def _bibliography_entries(model):
    """Записи списка литературы: [(номер или None, текст, группа, признак автонумерации)]."""
    if model.bib_start is None:
        return None
    end = model.bib_end if model.bib_end is not None else len(model.body)
    entries = []
    group = 0
    auto_counter = 0
    for para in model.body[model.bib_start + 1:end]:
        text = para.text.strip()
        if not text:
            continue
        match = re.match(r"^\s*(\d{1,4})\s*[.)]\s*(\S.*)$", text, re.DOTALL)
        numpr = model.num_pr(para)
        if match:
            entries.append((int(match.group(1)), match.group(2), group, False))
            continue
        if numpr is not None and model.numbering.fmt(*numpr) not in ("bullet", "none"):
            auto_counter += 1
            entries.append((auto_counter, text, group, True))
            continue
        looks_like_subheading = len(text) <= 80 and not re.search(r"\d{4}|[–—-]\s*\d|\bс\.\s*\d|URL", text)
        if looks_like_subheading:
            group += 1
            continue
        entries.append((None, text, group, False))
    return entries


def _alphabet_key(text):
    """Ключ первого слова записи, согласованный с format_bibliography.sort_text_key (SPEC 3):
    регистр не важен, ё = е, латиница с диакритикой приводится к базовой букве,
    кириллица раньше латиницы. Записи, начинающиеся с цифр, не сравниваются."""
    import unicodedata

    cleaned = (text or "").casefold().replace("ё", "е")
    match = re.search(r"[^\W\d_]+", cleaned)
    if not match or re.search(r"\d", cleaned[:match.start()]):
        return None
    letters = []
    for char in match.group(0):
        if "а" <= char <= "я" or "a" <= char <= "z":
            letters.append(char)
            continue
        base = unicodedata.normalize("NFKD", char)[0]
        letters.append(base if "a" <= base <= "z" else char)
    word = "".join(letters)
    script = 0 if "а" <= word[0] <= "я" else (1 if "a" <= word[0] <= "z" else 2)
    return script, word


def check_cross_references(doc: Document, profile: str = "generic") -> list:
    """Соответствие ссылок [N, с. XX] в тексте номерам в списке литературы.

    Раздел ищется по заголовку (не по строке оглавления): «Список использованной
    литературы», «Список использованных источников», «Список литературы»,
    «Библиографический список», «Литература» и др. Номера записей берутся из
    текста «N.» или из автонумерации Word (numPr). Ссылки собираются из текста,
    таблиц и сносок, без кода.
    """
    model = _model(doc)
    cited_numbers = set(_extract_citation_numbers("\n".join(_citation_texts(model))))
    entries = _bibliography_entries(model)
    if entries is None:
        return [check(False, "Раздел «Список литературы» не найден — проверить оформление невозможно.", severity="warning")]
    listed = [entry[0] for entry in entries if entry[0] is not None]
    listed_numbers = set(listed)
    if not listed_numbers:
        return [check(
            False,
            "В списке литературы не найдено пронумерованных записей. "
            "Проверить формат вручную.",
            severity="warning",
        )]

    results = [check(
        True,
        f"Источников в списке: {len(listed_numbers)}. Номеров в ссылках по тексту: {len(cited_numbers)}",
        severity="info",
    )]
    missing_in_list = cited_numbers - listed_numbers
    unused_in_text = listed_numbers - cited_numbers
    if missing_in_list:
        sample = sorted(missing_in_list)[:10]
        more = f" (ещё {len(missing_in_list) - 10})" if len(missing_in_list) > 10 else ""
        results.append(check(
            False,
            f"Ссылки в тексте указывают на номера {sample}{more}, которых НЕТ в списке литературы. "
            f"Вероятно, список был пересортирован после расстановки ссылок. Проверь соответствие номеров!",
            severity="error",
        ))
    if unused_in_text:
        sample = sorted(unused_in_text)[:10]
        more = f" (ещё {len(unused_in_text) - 10})" if len(unused_in_text) > 10 else ""
        results.append(check(
            False,
            f"В списке литературы есть источники {sample}{more}, которые НЕ упоминаются в тексте. "
            f"Методичка требует, чтобы все источники из списка были отражены в работе.",
            severity="warning",
        ))
    if not missing_in_list and not unused_in_text:
        results.append(check(
            True,
            "Все ссылки в тексте соответствуют номерам в списке литературы, "
            "и все источники из списка используются в тексте.",
        ))
    return results


def check_bibliography_order(doc: Document, profile: str = "generic") -> list:
    """Сквозная нумерация без дублей и алфавитный порядок списка литературы.

    Методичка (P0044/P0192): «Список составляется в алфавитном порядке».
    Кириллица раньше латиницы, «ё» = «е», сравнивается первое слово записи
    (фамилия автора или первое слово заглавия); внутри подзаголовков-групп
    порядок проверяется отдельно. Для профилей mpgu-* нарушение — ошибка,
    для generic — предупреждение.
    """
    model = _model(doc)
    entries = _bibliography_entries(model)
    if not entries:
        return [check(True, "Список литературы не найден или пуст — порядок не проверен", severity="info")]
    severity = "error" if _is_mpgu(profile) else "warning"
    results = []

    # Номера в порядке документа: набранные «N.» и автонумерация Word (счётчик
    # продолжается через набранные вручную номера, как в Word).
    numbered = [entry for entry in entries if entry[0] is not None]
    if numbered and not all(entry[3] for entry in numbered):
        numbers = [entry[0] for entry in numbered]
        duplicates = sorted({n for n in numbers if numbers.count(n) > 1})
        if duplicates:
            results.append(check(
                False,
                f"В списке литературы повторяются номера {duplicates[:10]}: ссылки [N] становятся неоднозначными.",
                severity=severity,
            ))
        expected = list(range(numbers[0], numbers[0] + len(numbers)))
        if not duplicates and (numbers != expected or numbers[0] != 1):
            results.append(check(
                False,
                f"Нумерация списка литературы не сквозная: {numbers[:12]}{'…' if len(numbers) > 12 else ''} "
                f"(ожидается 1, 2, 3 …).",
                severity=severity,
            ))

    inversions = []
    groups = {}
    for entry in entries:
        groups.setdefault(entry[2], []).append(entry)
    for group_entries in groups.values():
        keyed = [(entry, _alphabet_key(entry[1])) for entry in group_entries]
        keyed = [(entry, key) for entry, key in keyed if key is not None]
        for (left, left_key), (right, right_key) in zip(keyed, keyed[1:]):
            if left_key > right_key:
                inversions.append((left[1], right[1]))
    if inversions:
        examples = "; ".join(f"«{a[:30]}…» стоит перед «{b[:30]}…»" for a, b in inversions[:3])
        results.append(check(
            False,
            f"Список литературы не в алфавитном порядке ({len(inversions)} нарушений): {examples}. "
            f"Методичка: «Список составляется в алфавитном порядке» (кириллица, затем латиница). "
            f"После сортировки перенумеруй ссылки [N] в тексте.",
            severity=severity,
        ))
    else:
        results.append(check(True, "Список литературы упорядочен по алфавиту", severity="info"))
    if numbered and not any(not item["ok"] and "омер" in item["message"] for item in results):
        results.insert(0, check(True, "Нумерация списка литературы сквозная", severity="info"))
    return results


# ========== МАРКЕРЫ ==========

def check_unresolved_placeholders(doc: Document) -> list:
    """Маркеры незаполненного черновика (общий модуль vkr_markers, SPEC 7).

    Проверяются текст, таблицы, надписи, колонтитулы и сноски. Абзацы кода
    (VKR Listing, моноширинный шрифт) и моноширинные фрагменты не проверяются:
    history[x], {{ user.name }} и # TODO в листинге — не маркеры.
    """
    model = _model(doc)
    found = []
    for para in model.body + model.table_paras + model.floating_paras + model.part_paras:
        text = model.text_without_code(para)
        if not text.strip():
            continue
        for marker in vkr_markers.find_markers(text, docx_toc_placeholder=True):
            label = marker if para.container in {"body", "table"} else f"{marker} ({para.location})"
            found.append(label)
    unique = list(dict.fromkeys(found))
    if unique:
        sample = ", ".join(unique[:10])
        suffix = f"; ещё {len(unique) - 10}" if len(unique) > 10 else ""
        return [check(
            False,
            f"Остались маркеры незаполненного черновика: {sample}{suffix}. "
            "Замени их подтверждёнными данными или удали зависимые утверждения.",
            severity="error",
        )]
    return [check(True, "Незаполненные маркеры черновика не найдены", severity="info")]


# ========== НОМЕРА СТРАНИЦ И ФОРМАЛЬНЫЕ ПРИЗНАКИ ==========

def _has_page_field(root) -> bool:
    return root is not None and "PAGE" in _field_types(root)


def _section_is_title_only(model, index):
    section = model.sections[index]
    if section.first_idx is None or section.last_idx is None:
        return True
    return section.last_idx < model.title_end


def _check_page_numbers(model, profile="generic") -> list:
    results = []
    sections = model.sections
    if not sections:
        return [check(False, "Нет секций — номера страниц не проверены")]
    title_section = len(sections) > 1 and _section_is_title_only(model, 0)
    missing = []
    first_missing = []
    in_header = []
    title_number = False
    for index, section in enumerate(sections):
        if index == 0 and title_section:
            kind = "first" if section.title_pg else "default"
            title_number = _has_page_field(model.footer_root(0, kind)) or _has_page_field(model.header_root(0, kind))
            continue
        if section.first_idx is None and index == len(sections) - 1 and index > 0:
            continue
        if not _has_page_field(model.footer_root(index, "default")):
            missing.append(index + 1)
            if _has_page_field(model.header_root(index, "default")):
                in_header.append(index + 1)
        if index > 0 and section.title_pg and not _has_page_field(model.footer_root(index, "first")):
            first_missing.append(index + 1)
    if not title_section:
        first = sections[0]
        if first.title_pg:
            title_number = _has_page_field(model.footer_root(0, "first")) or _has_page_field(model.header_root(0, "first"))
        else:
            title_number = _has_page_field(model.footer_root(0, "default"))

    if missing:
        hint = ""
        if in_header:
            hint = f" (в секциях {', '.join(map(str, in_header))} номер в верхнем колонтитуле; методичка: внизу страницы)"
        results.append(check(
            False,
            "Нет поля PAGE в секциях: " + ", ".join(map(str, missing)) + hint
            + ". Вставь номер: «Вставка → Номер страницы → Внизу страницы».",
        ))
    else:
        results.append(check(True, "Во всех нижних колонтитулах найдено поле PAGE"))
    if first_missing:
        results.append(check(
            False,
            f"В секциях {', '.join(map(str, first_missing))} включён «Особый колонтитул для первой страницы» без номера: "
            f"первая страница секции останется без номера. Сними флажок «Особый колонтитул для первой страницы» "
            f"в этом разделе или вставь номер и в первый колонтитул.",
        ))
    if title_number:
        results.append(check(
            False,
            "На титульном листе выводится номер страницы. Методичка: «на титульном листе цифра 1 не ставится» — "
            "включи «Особый колонтитул для первой страницы» или вынеси титул в отдельный раздел без номера.",
        ))
    elif not missing or title_section:
        results.append(check(True, "Номер страницы на титульном листе не выводится"))

    restart_severity = "error" if _is_mpgu(profile) else "warning"
    if title_section and len(sections) > 1 and sections[1].pg_start == 1:
        results.append(check(
            False,
            "После титульного листа нумерация начинается заново с 1: аннотация получит номер 1 вместо 2. "
            "Методичка: титул — первая страница. Убери «Начать с 1» во втором разделе.",
            severity=restart_severity,
        ))
    restarts = [
        f"{index + 1} (с {section.pg_start})" for index, section in enumerate(sections)
        if section.pg_start is not None and index > 0 and not (index == 1 and title_section and section.pg_start == 1)
    ]
    if restarts:
        results.append(check(
            False,
            f"Нумерация страниц перезапускается в секциях: {', '.join(restarts)}. Методичка: все страницы "
            f"получают сквозной порядковый номер — в «Формат номеров страниц» выбери «продолжить».",
            severity=restart_severity,
        ))
    formats = sorted({section.pg_fmt for section in sections if section.pg_fmt not in (None, "decimal")})
    if formats:
        results.append(check(False, f"Нестандартный формат номеров страниц: {', '.join(formats)} (ожидаются арабские цифры)", severity="warning"))
    return results


def _break_after_content(para):
    """В абзаце после последнего видимого текста стоит разрыв страницы."""
    after_break = False
    for node in para.el.iter(W + "t", W + "br", W + "drawing", W + "pict", W + "object"):
        if node.tag == W + "br":
            if node.get(W + "type") == "page":
                after_break = True
        elif node.tag == W + "t":
            if (node.text or "").strip():
                after_break = False
        else:
            after_break = False
    return after_break


def _section_break_in(para):
    ppr = para.el.find(W + "pPr")
    sect = ppr.find(W + "sectPr") if ppr is not None else None
    if sect is None:
        return False
    kind = sect.find(W + "type")
    return (kind.get(W + "val") if kind is not None else "nextPage") != "continuous"


def _starts_new_page(model, para):
    if model.page_break_before(para):
        return True
    for node in para.el.iter(W + "br", W + "t"):
        if node.tag == W + "t" and (node.text or "").strip():
            break
        if node.tag == W + "br" and node.get(W + "type") == "page":
            return True
    block_index = para.block
    while block_index > 0:
        block_index -= 1
        block = model.blocks[block_index]
        if block[0] == "tbl":
            return False
        previous = block[1]
        if _break_after_content(previous) or _section_break_in(previous):
            return True
        if _has_visual_content(previous):
            return False
    return True


def _has_visual_content(para):
    return bool(
        para.text.strip()
        or para.el.find(".//" + W + "drawing") is not None
        or para.el.find(".//" + W + "pict") is not None
        or para.el.find(".//" + W + "object") is not None
        or para.el.find(".//" + M_MATH + "oMath") is not None
    )


def _check_trailing_empty_page(model) -> list:
    blocks = model.blocks
    last_content = None
    for index in range(len(blocks) - 1, -1, -1):
        block = blocks[index]
        if block[0] == "tbl" or _has_visual_content(block[1]):
            last_content = index
            break
    if last_content is None:
        return []
    breaks = False
    if blocks[last_content][0] == "p":
        para = blocks[last_content][1]
        breaks = _break_after_content(para) or (_section_break_in(para) and last_content < len(blocks) - 1)
    for position in range(last_content + 1, len(blocks)):
        para = blocks[position][1]
        if any(br.get(W + "type") == "page" for br in para.el.iter(W + "br")) or model.page_break_before(para):
            breaks = True
        if _section_break_in(para) and position < len(blocks) - 1:
            breaks = True
    if not breaks:
        return [check(True, "В конце документа нет пустой страницы", severity="info")]
    numbered = _has_page_field(model.footer_root(len(model.sections) - 1, "default"))
    return [check(
        False,
        "В конце документа пустая страница" + (" с номером" if numbered else "")
        + ": после последнего содержимого стоит разрыв страницы или раздела. Удали лишний разрыв.",
        severity="error" if numbered else "warning",
    )]


def _chapter_numbering_problems(text):
    first_line = text.split("\n")[0].strip()
    match = CHAPTER_START_RE.match(first_line)
    if not match:
        return None, False
    token = match.group(1)
    normalized = _normalize(text)
    duplicate = bool(re.match(r"^глава\s+\S+\.?\s+глава\s+\S+", normalized))
    if re.fullmatch(r"[IVXLCDM]+", token) and ROMAN_RE.match(token):
        return None, duplicate
    if token.isdigit():
        return "арабская", duplicate
    return "не римская", duplicate


def check_document_mechanics(doc: Document, profile: str = "generic") -> list:
    """Формальные проверки: номера страниц, титул без номера, нумерация глав,
    точки в заголовках, ссылка до точки, разрыв перед главой, выравнивание,
    пустая последняя страница."""
    model = _model(doc)
    results = _check_page_numbers(model, profile)

    bad_chapters = []
    duplicate_prefix = []
    dotted_headings = []
    chapters_without_break = []
    for para in model.body:
        if model.is_toc(para) or para.idx < model.title_end:
            continue
        text = para.text.strip()
        if not text:
            continue
        region = model.region(para)
        is_chapter = model.is_chapter_heading(para) and para.idx < model.main_end
        if is_chapter:
            problem, duplicate = _chapter_numbering_problems(text)
            if problem:
                bad_chapters.append(text.replace("\n", " ")[:60])
            if duplicate:
                duplicate_prefix.append(text.replace("\n", " ")[:60])
            if not _starts_new_page(model, para):
                chapters_without_break.append(text.replace("\n", " ")[:60])
        is_heading = is_chapter or model.heading_level(para) is not None or bool(
            # «Заключение.» обычным абзацем по ширине — текст, а не заголовок
            model.structural_key(para) and (model.all_bold(para) or model.alignment(para) == "center")
        )
        if is_heading and region in {"main", "appendix", "bibliography", "other"}:
            if region == "bibliography" and para.idx != model.bib_start:
                continue
            last_line = text.split("\n")[-1].strip()
            if (
                last_line.endswith(".") and not _ends_with_ellipsis(last_line)
                and not _ends_with_abbreviation(last_line)
                and not re.fullmatch(r"(?:глава\s+\S+|\d+(?:\.\d+)*)\.", last_line, re.IGNORECASE)
            ):
                dotted_headings.append(text.replace("\n", " ")[:60])

    numbering_severity = "error" if _is_mpgu(profile) else "warning"
    results.append(check(
        not bad_chapters,
        "Главы имеют римскую нумерацию" if not bad_chapters
        else "Главы нумеруются не римскими цифрами (методичка: «Нумерация глав – римскими цифрами», "
             "например «Глава I.»): " + "; ".join(bad_chapters[:3]),
        severity="error" if not bad_chapters else numbering_severity,
    ))
    if duplicate_prefix:
        results.append(check(False, "Префикс главы повторяется: " + "; ".join(duplicate_prefix[:3])))
    results.append(check(
        not dotted_headings,
        "В конце заголовков нет точки" if not dotted_headings else "Точка в конце заголовка: " + "; ".join(dotted_headings[:3]),
    ))

    citation_after_period = []
    for para in model.body:
        if model.region(para) not in {"main", "appendix"}:
            continue
        text = model.text_without_code(para)
        for match, _numbers in _iter_citations(text):
            before = text[:match.start()]
            stripped = before.rstrip()
            if not stripped.endswith("."):
                continue
            if _ends_with_ellipsis(stripped) or _ends_with_abbreviation(stripped):
                continue
            citation_after_period.append(text.strip()[:80])
            break
    results.append(check(
        not citation_after_period,
        "Ссылки [N] стоят до точки" if not citation_after_period
        else "Ссылка стоит после точки (методичка: «Ссылка на источник указывается до точки»): " + "; ".join(citation_after_period[:3]),
    ))
    results.append(check(
        not chapters_without_break,
        "Каждая глава начинается с новой страницы" if not chapters_without_break
        else "Нет разрыва страницы перед главой: " + "; ".join(chapters_without_break[:3]),
        severity="warning",
    ))
    results.extend(_check_alignment(model))
    results.extend(_check_trailing_empty_page(model))
    return results

# ========== ГЛАВНАЯ ФУНКЦИЯ ==========

def _error_report(code: str, message: str, path=None, profile=None) -> dict:
    report = {
        "status": "error",
        "error_code": code,
        "error": message,
        "validator_version": VALIDATOR_VERSION,
    }
    if path is not None:
        report["path"] = str(path)
    if profile is not None:
        report["profile"] = profile
    return report


def _run_check(function, *args):
    try:
        return function(*args)
    except Exception as error:  # noqa: BLE001 — сбой одной проверки не должен скрывать остальные
        return [check(
            False,
            f"Проверка не выполнена из-за внутренней ошибки валидатора: {type(error).__name__}: {error}",
            severity="error",
        )]


def validate_vkr(
    docx_path: str,
    verbose: bool = False,
    profile: str = "generic",
) -> dict:
    """Валидация ВКР. Возвращает dict с результатами.

    При ошибке ввода — {"status": "error", "error_code": …, "error": …}:
    UNKNOWN_PROFILE, FILE_NOT_FOUND, PYTHON_DOCX_MISSING, INVALID_DOCX.
    """
    global CLICHE_ALLOWLIST, _ACTIVE_MODEL
    if profile not in VALIDATION_PROFILES:
        return _error_report(
            "UNKNOWN_PROFILE",
            f"Неизвестный профиль: {profile}. Допустимо: {', '.join(VALIDATION_PROFILES)}",
            docx_path,
            profile,
        )
    path = Path(docx_path)
    if not path.is_file():
        return _error_report("FILE_NOT_FOUND", f"Файл не найден: {docx_path}", path, profile)
    try:
        document_class = _import_docx()
    except ImportError:
        return _error_report(
            "PYTHON_DOCX_MISSING",
            "python-docx не установлен. Установи: python -m pip install python-docx",
            path,
            profile,
        )
    try:
        doc = document_class(str(path))
        model = _DocModel(doc)
    except Exception as error:  # noqa: BLE001 — любой сбой разбора = повреждённый DOCX
        return _error_report("INVALID_DOCX", f"Не удалось открыть файл: {error}", path, profile)

    CLICHE_ALLOWLIST = _load_cliche_allowlist(path)
    _ACTIVE_MODEL = (doc, model)
    try:
        all_checks = {
            "Формат страницы": _run_check(check_page_size, doc),
            "Поля страницы": _run_check(check_margins, doc),
            "Шрифт и размер": _run_check(check_fonts, doc),
            "Междустрочный интервал": _run_check(check_line_spacing, doc),
            "Отступ первой строки": _run_check(check_first_line_indent, doc),
            "Цвет шрифта": _run_check(check_font_color, doc),
            "Содержимое аннотации": _run_check(check_annotation_content, doc),
            "Структура работы": _run_check(check_structure, doc),
            "Выводы по главам": _run_check(check_chapter_conclusions, doc),
            "Оглавление заполнено": _run_check(check_toc_filled, doc, profile),
            "Незаполненные маркеры": _run_check(check_unresolved_placeholders, doc),
            "Нумерация и формальные признаки": _run_check(check_document_mechanics, doc, profile),
            "Профиль методички": _run_check(check_mpgu_09_profile, doc, profile),
            "Образовательный контекст (по профилю)": _run_check(check_pedagogical_frame, doc, profile),
            "Выборка пилота (для проекта)": _run_check(check_pilot_sample_size, doc, profile),
            "Длина глав и объём": _run_check(check_chapter_lengths, doc, profile),
            "Ссылки на таблицы и рисунки": _run_check(check_figure_table_references, doc),
            "Оформление таблиц": _run_check(check_tables, doc),
            "Личные местоимения": _run_check(check_personal_pronouns, doc),
            "ИИ-клише": _run_check(check_ai_cliches, doc),
            "Ритм абзацев": _run_check(check_paragraph_rhythm, doc),
            "Начала абзацев": _run_check(check_consecutive_same_starts, doc),
            "Ссылки на источники": _run_check(check_citations, doc),
            "Соответствие ссылок и списка литературы": _run_check(check_cross_references, doc, profile),
            "Порядок и нумерация списка литературы": _run_check(check_bibliography_order, doc, profile),
            "Оформление фамилий": _run_check(check_surnames_format, doc),
        }
    finally:
        _ACTIVE_MODEL = None

    total_errors = 0
    total_warnings = 0
    total_passed = 0
    for group in all_checks.values():
        for result in group:
            if result["ok"]:
                total_passed += 1
            elif result["severity"] == "error":
                total_errors += 1
            elif result["severity"] == "warning":
                total_warnings += 1

    return {
        "status": "ok",
        "validator_version": VALIDATOR_VERSION,
        "path": str(path),
        "profile": profile,
        "document_sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
        "summary": {
            "passed": total_passed,
            "warnings": total_warnings,
            "errors": total_errors,
        },
        "checks": all_checks,
    }


def print_report(report: dict, verbose: bool = False):
    """Вывод отчёта в консоль."""
    if "error" in report:
        print(f"ОШИБКА: {report['error']}", file=sys.stderr)
        return

    print(f"\n{'='*70}")
    print(f"  ВАЛИДАЦИЯ ВКР: {report['path']}")
    print(f"  ПРОФИЛЬ: {report.get('profile', 'generic')}   ВАЛИДАТОР: {report.get('validator_version', VALIDATOR_VERSION)}")
    print(f"{'='*70}\n")

    for group_name, results in report["checks"].items():
        visible = results if verbose else [result for result in results if not result["ok"]]
        print(f"\n{group_name}")
        print("-" * 70)
        if not visible:
            print("  OK")
        for result in visible:
            if result["ok"]:
                icon = "OK"
            elif result["severity"] == "warning":
                icon = "ПРЕДУПРЕЖДЕНИЕ"
            elif result["severity"] == "info":
                icon = "ИНФО"
            else:
                icon = "ОШИБКА"
            print(f"  {icon} {result['message']}")

    print(f"\n{'='*70}")
    s = report["summary"]
    print(f"  ИТОГО: OK {s['passed']}  ПРЕДУПРЕЖДЕНИЯ {s['warnings']}  ОШИБКИ {s['errors']}")
    print(f"{'='*70}\n")

    if s["errors"] > 0:
        print("ОШИБКА: есть критические ошибки; работу нужно доработать перед сдачей.")
    elif s["warnings"] > 0:
        print("ПРЕДУПРЕЖДЕНИЕ: рекомендуется проверить и улучшить отмеченные места.")
    else:
        print("OK: работа прошла базовую валидацию.")
    print()


def _exit_code(report: dict) -> int:
    if report.get("status") == "error" or "error" in report:
        return 2
    return 1 if report.get("summary", {}).get("errors", 0) else 0


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except (ValueError, OSError):
                pass
    parser = argparse.ArgumentParser(
        description="Валидатор ВКР для МПГУ. Коды возврата: 0 — ошибок нет, 1 — найдены ошибки, "
                    "2 — неверный профиль, нет файла, повреждённый DOCX или нет python-docx."
    )
    parser.add_argument("docx_path", help="Путь к .docx файлу ВКР")
    parser.add_argument("--verbose", "-v", action="store_true", help="Подробный вывод")
    parser.add_argument("--json", action="store_true", help="Только JSON-отчёт в stdout")
    parser.add_argument("-o", "--output", type=Path, help="Записать JSON-отчёт в файл (UTF-8, каталоги создаются)")
    parser.add_argument(
        "--profile",
        default="generic",
        help="Профиль требований: " + ", ".join(VALIDATION_PROFILES) + " (по умолчанию generic)",
    )
    args = parser.parse_args(argv)

    report = validate_vkr(args.docx_path, args.verbose, args.profile)
    code = _exit_code(report)
    payload = json.dumps(report, ensure_ascii=False, indent=2) + "\n"

    if args.output:
        try:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(payload, encoding="utf-8")
        except OSError as error:
            message = f"не удалось записать отчёт {args.output}: {error}"
            if args.json:
                print(json.dumps(_error_report("OUTPUT_WRITE_FAILED", message), ensure_ascii=False, indent=2))
            else:
                print(f"ОШИБКА: {message}", file=sys.stderr)
            return 2
    if args.json:
        sys.stdout.write(payload)
    elif code == 2:
        print_report(report, args.verbose)
    else:
        print_report(report, args.verbose)
    return code


if __name__ == "__main__":
    raise SystemExit(main())
