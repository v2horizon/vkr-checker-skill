#!/usr/bin/env python3
"""
ai_detection_heuristic.py — стайл-чекер на известные паттерны ИИ-генерации (vkr-mpgu 6.33).

ЧТО ЭТОТ СКРИПТ ДЕЛАЕТ (честная декларация):
  Проверяет текст на наличие известных паттернов, характерных для LLM-генерации:
  низкая burstiness, типичные ИИ-клише и N-граммы, бедное лексическое разнообразие,
  монотонный ритм абзацев, неудачное распределение авторских якорьков.

ЧТО ЭТОТ СКРИПТ НЕ ДЕЛАЕТ:
  - НЕ симулирует реальный Антиплагиат.Вуз (его архитектура не раскрыта публично)
  - НЕ гарантирует, что LOW-результат == работа пройдёт детектор
  - НЕ заменяет ручную проверку текста

ЦЕЛЬ: Быстрый самопроверочный чек «не похоже ли очевидно на LLM». Если скрипт
показывает HIGH/CRITICAL — работа точно не пройдёт. Если LOW — риск снижен,
но не нулевой. Это необходимое, но недостаточное условие.

ЧТО АНАЛИЗИРУЕТСЯ:
  Разделы определяются по заголовкам: стили Heading 1/2 («Заголовок 1/2») или
  уровень структуры; без стилей — отдельные абзацы, целиком совпадающие с
  названием раздела («Введение», «Глава I…», «Заключение», варианты заголовка
  списка литературы, «Приложение N», «Приложения»). В анализ идут введение,
  главы и заключение. Титул, аннотация, содержание (стили TOC, поле TOC,
  sdt-оглавление, строки «Название<TAB>34»), список литературы, приложения,
  последний лист, код (стиль VKR Listing, моноширинный шрифт) и подписи
  (VKR Caption, «Таблица N», «Рисунок N — …», «Источник: …») исключаются.

МЕТРИКИ (все — эвристические, пороги откалиброваны на примерах человек vs LLM):
    1. Burstiness:           CV длин абзацев, CV длин предложений
    2. Плотность N-грамм:    сильные и средние ИИ-клише из data/ai_cliches.json
    3. Лексическое:          TTR, MTLD (упрощённый), overused-слова
    4. Ритм:                 длинные серии однородных абзацев
    5. Якорьки:              плотность (не слишком много, но и есть)
    6. Starter variety:      разнообразие начал предложений
    7. Passive voice ratio:  LLM любит пассивные конструкции
    8. Readability proxy:    длина предложений vs средняя длина слова

Использование:
    python ai_detection_heuristic.py vkr.docx
    python ai_detection_heuristic.py vkr.docx --section chapter1
    python ai_detection_heuristic.py vkr.docx --per-section --json -o report.json

Коды выхода: 0 — анализ выполнен (уровень риска смотри в отчёте);
1 — недостаточно данных (INSUFFICIENT_DATA), раздел --section не найден, в
--per-section есть ненайденные или недостаточные разделы; 2 — нет файла,
повреждённый DOCX, нет python-docx, ошибка записи -o, неверные аргументы.

Персональный cliche_allowlist берётся из vkr-state.md в каталоге DOCX или выше
(не более трёх уровней и не выше каталога с vkr-project.json). Из того же
state-файла читаются личные якорьки раздела personal_style_profile («Якорьки
для ВКР», «Оборот для критики/сомнения», «Оборот для технического вывода» —
фразы в кавычках): метрика «Якорьки» считает их вместе со встроенным списком
общих оборотов (в отчёте — anchors.personal_anchors и anchor_source).

Модуль также содержит общий разбор структуры DOCX (DocumentStructure), который
используют content_ownership_check.py и verify_sources.py.

НЕ доверяй одному результату. Прогоняй до и после правок, сравнивай тренд.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import statistics
import sys
from collections import OrderedDict
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

sys.dont_write_bytecode = True
SCRIPTS_DIR = Path(__file__).resolve().parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import vkr_common as common  # noqa: E402

TOOL_VERSION = "6.33"
MIN_WORDS_FULL = 300
MIN_WORDS_SECTION = 120


# ========== СЛОВАРИ ==========

AI_CLICHES_PATH = Path(__file__).resolve().parent.parent / "data" / "ai_cliches.json"


def _load_ai_ngrams():
    """Загрузка N-грамм из централизованного data/ai_cliches.json.
    Единый источник истины с validate_vkr.py. При сбое загрузки —
    минимальный встроенный fallback, чтобы детектор не падал.

    Возвращает кортеж (strong_list, medium_list).
    """
    try:
        with open(AI_CLICHES_PATH, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
        return data.get("strong", []), data.get("medium", [])
    except (OSError, ValueError) as e:
        print(
            f"ВНИМАНИЕ: не удалось загрузить {AI_CLICHES_PATH}: {e}",
            file=sys.stderr,
        )
        print("Используется минимальный fallback-список N-грамм.", file=sys.stderr)
        fallback_strong = [
            "стоит отметить",
            "важно отметить",
            "важно подчеркнуть",
            "следует отметить",
            "необходимо отметить",
            "в современном мире",
            "давайте рассмотрим",
            "как мы можем видеть",
            "в эпоху стремительного",
            "благодаря развитию информационных технологий",
        ]
        fallback_medium = [
            "в свою очередь",
            "вместе с тем",
            "таким образом",
            "на сегодняшний день",
            "существует множество",
            "играет важную роль",
            "обладает рядом преимуществ",
            "современные технологии позволяют",
            "рассматриваемый подход",
            "позволяет значительно оптимизировать",
        ]
        return fallback_strong, fallback_medium


# N-граммы, типичные для ИИ-текста на русском академическом
AI_NGRAMS_STRONG, AI_NGRAMS_MEDIUM = _load_ai_ngrams()
CLICHE_ALLOWLIST: List[str] = []
# Личные якорьки из personal_style_profile state-файла анализируемого DOCX (заполняет main).
PERSONAL_ANCHORS: List[str] = []


# ========== ПЕРСОНАЛЬНЫЙ CLICHE_ALLOWLIST ==========

STATE_FILE_NAMES = ("vkr-state.md", ".vkr-state.md", "VKR-STATE.md")
# Уровни вверх от каталога DOCX (final/vkr.docx → корень проекта); поиск
# останавливается на каталоге с vkr-project.json — чужой проект не читается.
STATE_SEARCH_LEVELS = 3
PROJECT_CONFIG_NAME = "vkr-project.json"

_ALLOWLIST_KEY_RE = re.compile(r"^[ \t]*cliche_allowlist[ \t]*:(?P<rest>.*)$")
_YAML_ITEM_RE = re.compile(r"^[ \t]*-(?:[ \t]+(?P<value>.*?))?[ \t]*$")


def normalize_phrase(value: str) -> str:
    """Фраза для сравнения: без кавычек, пробелы схлопнуты, регистр и ё/е не важны."""
    value = str(value or "").strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
        value = value[1:-1]
    return re.sub(r"\s+", " ", value).strip().casefold().replace("ё", "е")


def _strip_yaml_comment(value: str) -> str:
    result = []
    quote = None
    for index, char in enumerate(value):
        if quote:
            if char == quote:
                quote = None
        elif char in "\"'":
            quote = char
        elif char == "#" and (index == 0 or value[index - 1] in " \t"):
            break
        result.append(char)
    return "".join(result).strip()


def _split_inline_list(body: str) -> List[str]:
    items: List[str] = []
    current: List[str] = []
    quote = None
    for char in body:
        if quote:
            current.append(char)
            if char == quote:
                quote = None
        elif char in "\"'":
            quote = char
            current.append(char)
        elif char == ",":
            items.append("".join(current))
            current = []
        else:
            current.append(char)
    items.append("".join(current))
    return items


def parse_cliche_allowlist(text: str) -> List[str]:
    """Разбирает поле ``cliche_allowlist`` state-файла.

    Понимает YAML-список с отступом и без (``cliche_allowlist:`` + строки ``- …``),
    inline-список ``cliche_allowlist: ["a", 'b']`` (в том числе на нескольких
    строках) и одиночное значение. Закомментированные строки ``# …`` не читаются.
    """
    lines = str(text or "").splitlines()
    for index, line in enumerate(lines):
        match = _ALLOWLIST_KEY_RE.match(line)
        if not match:
            continue
        rest = _strip_yaml_comment(match.group("rest"))
        phrases: List[str] = []
        if rest.startswith("["):
            body = rest[1:]
            cursor = index
            while "]" not in body and cursor + 1 < len(lines):
                cursor += 1
                body += " " + _strip_yaml_comment(lines[cursor])
            body = body.split("]", 1)[0]
            phrases = [normalize_phrase(item) for item in _split_inline_list(body)]
        elif rest and rest not in ("~", "null", "None"):
            phrases = [normalize_phrase(rest)]
        elif not rest:
            for following in lines[index + 1:]:
                stripped = following.strip()
                if not stripped or stripped.startswith("#"):
                    continue
                item = _YAML_ITEM_RE.match(following)
                if not item:
                    break
                phrases.append(normalize_phrase(_strip_yaml_comment(item.group("value") or "")))
        return [phrase for phrase in phrases if phrase]
    return []


def find_state_file(document_path) -> Optional[Path]:
    """vkr-state.md в каталоге DOCX или выше, не выше корня проекта (vkr-project.json)."""
    try:
        start = Path(document_path).resolve().parent
    except (OSError, RuntimeError, ValueError):
        return None
    for parent in [start] + list(start.parents)[:STATE_SEARCH_LEVELS]:
        for name in STATE_FILE_NAMES:
            candidate = parent / name
            try:
                if candidate.is_file():
                    return candidate
            except OSError:
                continue
        try:
            if (parent / PROJECT_CONFIG_NAME).is_file():
                return None
        except OSError:
            return None
    return None


def load_cliche_allowlist(document_path) -> Tuple[List[str], Optional[Path]]:
    """Возвращает (фразы allowlist, путь state-файла или None)."""
    state = find_state_file(document_path)
    if state is None:
        return [], None
    try:
        text = state.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError):
        return [], state
    return parse_cliche_allowlist(text), state


def _load_cliche_allowlist(document_path) -> list:
    """Совместимость с 6.32: только список фраз."""
    return load_cliche_allowlist(document_path)[0]


def allowlist_info(phrases: Sequence[str], state: Optional[Path]) -> Dict[str, Any]:
    return {"state_file": str(state) if state else None, "phrases": list(phrases)}


# ========== ЛИЧНЫЕ ЯКОРЬКИ ИЗ personal_style_profile ==========

_MD_HEADING_RE = re.compile(r"^[ \t]{0,3}(#{1,6})[ \t]+(.*)$")
_PROFILE_TITLE_RE = re.compile(r"personal_style_profile|личный\s+стилистический\s+профиль", re.IGNORECASE)
_ANCHORS_LABEL_RE = re.compile(r"^\*{0,2}\s*якор(?:ьки|ёк|ек)\b", re.IGNORECASE)
_TURN_LABEL_RE = re.compile(r"^\*{0,2}\s*оборот\s+для\s+[^:*]+:\*{0,2}\s*(?P<value>.*)$", re.IGNORECASE)
_QUOTED_RE = re.compile(r"[«\"“„]([^«»\"“”„\n]{2,120})[»\"”“]")


def _quoted_phrases(text: str) -> List[str]:
    """Фразы в кавычках; шаблонные заглушки вида [«якорёк 1»] пропускаются."""
    phrases = []
    for match in _QUOTED_RE.finditer(text):
        before = text[:match.start()].rstrip()
        if before.endswith("["):
            continue
        phrases.append(normalize_phrase(match.group(1)))
    return [phrase for phrase in phrases if phrase]


def parse_personal_anchors(text: str) -> List[str]:
    """Личные якорьки из раздела personal_style_profile state-файла.

    Берутся фразы в кавычках из списка «Якорьки для ВКР» и строк «Оборот для критики/сомнения»,
    «Оборот для технического вывода» (шаблон — references/state-file-pattern.md). Заглушки шаблона
    в квадратных скобках и закомментированные строки не читаются.
    """
    anchors: List[str] = []
    profile_level = 0
    in_anchor_list = False
    in_fence = False
    for line in str(text or "").splitlines():
        stripped = line.strip()
        if stripped.startswith(("```", "~~~")):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        heading = _MD_HEADING_RE.match(line)
        if heading:
            level = len(heading.group(1))
            if _PROFILE_TITLE_RE.search(heading.group(2)):
                profile_level = level
            elif profile_level and level <= profile_level:
                profile_level = 0
            in_anchor_list = False
            continue
        if not profile_level or not stripped or stripped.startswith("<!--"):
            continue
        if _ANCHORS_LABEL_RE.match(stripped):
            in_anchor_list = True
            continue
        turn = _TURN_LABEL_RE.match(stripped)
        if turn:
            in_anchor_list = False
            anchors.extend(_quoted_phrases(turn.group("value"))[:1])
            continue
        if stripped.startswith("**"):
            in_anchor_list = False
            continue
        if in_anchor_list and stripped[:1] in "-*+":
            anchors.extend(_quoted_phrases(stripped)[:1])
    unique: List[str] = []
    for phrase in anchors:
        if phrase not in unique:
            unique.append(phrase)
    return unique


def load_personal_anchors(state: Optional[Path]) -> List[str]:
    if state is None:
        return []
    try:
        return parse_personal_anchors(state.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeDecodeError):
        return []


AI_WORDS_OVERUSED = {
    "ключевой", "важный", "значительный", "существенный", "комплексный",
    "эффективный", "инновационный", "актуальный", "современный", "системный",
    "оптимальный", "уникальный", "универсальный",
}

# Слова-индикаторы авторской позиции (якорьки)
# ВНИМАНИЕ (v6.16): это **fallback**-список общих оборотов для случаев,
# когда в state-файле НЕТ personal_style_profile (первая сессия, минимальный intake).
# Фразы здесь **общие**, использовать их буквально — риск кросс-матчинга
# (разные пользователи скилла получат одни и те же обороты в похожих позициях).
#
# Правильное поведение ИИ-агента: сначала пытаться читать personal_style_profile из
# state-файла пользователя и использовать ЕГО индивидуальные якорьки.
# Этот список — только для грубой оценки наличия/отсутствия авторской позиции
# в тексте, не для рекомендаций.
#
# Удалены в v6.11: «на наш взгляд», «по нашему мнению» (запрещено методичкой).
# Удалены в v6.16: «представляется обоснованным», «представляется целесообразным»,
# «в контексте настоящей работы» — они же фигурируют в humanizer-techniques.md
# (раздел «Примеры якорьков») и intake-interview.md как примеры с явным
# предупреждением «не копировать буквально».
# Оставлять их в положительном списке — значит награждать буквальное копирование.
AUTHOR_ANCHORS = [
    "в рассматриваемом контексте",
    "данное соображение учитывалось",
    "необходимо оговориться",
    "при всех оговорках",
    "если следовать данной логике",
    "с точки зрения практической",
    "с учётом сказанного",
    "с позиции научной",
    "представляется небеспочвенным",
]

# Личные местоимения (не должны встречаться): (метка, регулярное выражение)
PERSONAL_PRONOUNS = [
    ("я", r"\bя\b"), ("мне", r"\bмне\b"), ("меня", r"\bменя\b"),
    ("мной", r"\bмно(?:й|ю)\b"),
    ("мы", r"\bмы\b"), ("нас", r"\bнас\b"), ("нам", r"\bнам\b"), ("нами", r"\bнами\b"),
    ("наш", r"\bнаш(?:а|е|и|его|ему|ей|ем|им|ими|их|у)?\b"),
]


# ========== СТРУКТУРА DOCX (общая для анализаторов и verify_sources.py) ==========

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _w(tag: str) -> str:
    return "{%s}%s" % (W_NS, tag)


_P, _R, _T, _TAB, _BR, _CR = _w("p"), _w("r"), _w("t"), _w("tab"), _w("br"), _w("cr")
_NO_BREAK_HYPHEN = _w("noBreakHyphen")
_TBL, _TR, _TC = _w("tbl"), _w("tr"), _w("tc")
_SDT, _SDT_PR, _SDT_CONTENT = _w("sdt"), _w("sdtPr"), _w("sdtContent")
_CUSTOM_XML = _w("customXml")
_PPR, _RPR, _PSTYLE, _RSTYLE = _w("pPr"), _w("rPr"), _w("pStyle"), _w("rStyle")
_NUMPR, _NUMID, _OUTLINE = _w("numPr"), _w("numId"), _w("outlineLvl")
_RFONTS, _BOLD, _VAL = _w("rFonts"), _w("b"), _w("val")
_FLDCHAR, _INSTR, _FLDSIMPLE = _w("fldChar"), _w("instrText"), _w("fldSimple")
_MC_ALTERNATE, _MC_CHOICE, _MC_FALLBACK = (
    "{%s}AlternateContent" % MC_NS, "{%s}Choice" % MC_NS, "{%s}Fallback" % MC_NS,
)
_SKIP_IN_PARAGRAPH = {_PPR, _w("del"), _w("moveFrom"), _MC_FALLBACK}

MONO_FONTS = {
    "consolas", "courier new", "courier", "lucida console", "lucida sans typewriter", "menlo", "monaco",
    "dejavu sans mono", "liberation mono", "source code pro", "jetbrains mono", "fira code", "fira mono",
    "roboto mono", "ubuntu mono", "pt mono", "cascadia code", "cascadia mono", "sf mono", "inconsolata",
    "noto sans mono", "andale mono", "droid sans mono", "hack", "anonymous pro", "ibm plex mono",
}
CODE_STYLE_RE = re.compile(
    r"(?:vkr listing|listing|листинг|source code|html preformatted|плоский текст|plain text|\bcode\b)",
    re.IGNORECASE,
)
CAPTION_STYLE_RE = re.compile(r"^(?:vkr caption|caption|название объекта)$", re.IGNORECASE)
TOC_STYLE_RE = re.compile(r"^(?:toc|оглавление)(?:\s|\d|$)", re.IGNORECASE)
# Строка оглавления: «Название<TAB>34», «Название ........ 34», «Название … 34».
TOC_LINE_RE = re.compile(r"(?:\t|\.{3,}|…+|_{3,})\s*\d{1,4}\s*$")
TOC_LINE_MAX_CHARS = 250
OATH_RE = re.compile(r"выполнена\s+мной\s+(?:совершенно\s+)?самостоятельно", re.IGNORECASE)
LAST_SHEET_KEY = "выпускная квалификационная работа"

ANNOTATION_RE = re.compile(r"(?:аннотация|реферат|abstract)")
CONTENTS_RE = re.compile(r"(?:содержание|оглавление)")
BIBLIOGRAPHY_RE = re.compile(
    r"(?:список\s+(?:использованн\w+\s+|цитируем\w+\s+)?(?:литературы|источников)"
    r"(?:\s+и\s+(?:литературы|источников))?"
    r"|библиографическ\w+\s+список|библиография"
    r"|(?:использованная\s+)?литература(?:\s+и\s+источники)?|источники\s+и\s+литература)"
)
_CHAPTER_RE = re.compile(
    r"глава\s+(?P<num>\d{1,2}|[ivxlc]{1,7}|первая|вторая|третья|четвертая|пятая|шестая)(?P<rest>(?!\w).*)"
)
_APPENDIX_RE = re.compile(r"приложение(?:\s+(?:№\s*)?(?P<num>\d{1,2}|[а-яa-z])(?!\w))?(?P<rest>.*)")
_SEPARATOR_TITLE_RE = re.compile(r"^\s*[.:—–-]\s*(?P<title>.*)$")
CHAPTER_CONCLUSIONS_RE = re.compile(
    r"^выводы\s+(?:по\s+)?(?:главе|первой|второй|третьей|четвертой|пятой|[ivx\d]+\s+главе)"
)
SUBSECTION_NUMBER_RE = re.compile(r"^\d+\.\d+")
LEADING_NUMBER_RE = re.compile(r"^(?P<num>\d{1,2})(?:\.|\)|\s)\s*\S")
CAPTION_TEXT_RE = re.compile(
    r"^(?:(?:таблица|рисунок|рис\.|листинг|схема|диаграмма)\s*(?:№\s*)?[\dA-Za-zА-Яа-яЁё]{1,4}(?:\.\d+)*"
    r"(?:\s*[.:—–-]\s*.{0,250})?|источник\s*:.{0,300})$",
    re.IGNORECASE,
)
FIGURE_PLACEHOLDER_RE = re.compile(r"^\[\s*Здесь\s+вставить[^\]]*\]$", re.IGNORECASE)
ORDINAL_CHAPTERS = {"первая": 1, "вторая": 2, "третья": 3, "четвертая": 4, "пятая": 5, "шестая": 6}
ROMAN_VALUES = {"i": 1, "v": 5, "x": 10, "l": 50, "c": 100}

SECTION_NAME_RE = re.compile(r"^(?:introduction|conclusion|chapter[1-9]\d?)$")
BODY_REGION_KINDS = ("introduction", "chapter", "conclusion")
AFTER_BODY_KINDS = ("conclusion", "bibliography", "appendix", "appendix_divider", "last_sheet", "other")
NON_TEXT_REGION_KINDS = ("annotation", "contents", "bibliography", "appendix", "appendix_divider", "last_sheet", "other")


class DependencyError(RuntimeError):
    """Нет python-docx (код выхода 2)."""


def require_docx():
    try:
        from docx import Document
    except ImportError as error:
        raise DependencyError("python-docx не установлен: python -m pip install python-docx") from error
    return Document


def open_document(source):
    """Путь → python-docx Document; готовый Document возвращается как есть."""
    if hasattr(source, "element") and hasattr(source, "part"):
        return source
    Document = require_docx()
    return Document(str(source))


def looks_like_toc_line(text: str) -> bool:
    """Короткая строка с табуляцией или заполнителем и номером страницы в конце."""
    raw = str(text or "")
    return len(raw.strip()) <= TOC_LINE_MAX_CHARS and bool(TOC_LINE_RE.search(raw))


def is_inside_skill_dir(path) -> bool:
    """SPEC 10: скрипты не пишут внутрь каталога скилла."""
    try:
        target = os.path.normcase(str(Path(path).resolve()))
        skill = os.path.normcase(str(common.SKILL_DIR.resolve()))
    except (OSError, RuntimeError, ValueError):
        return False
    return target == skill or target.startswith(skill.rstrip("\\/") + os.sep)


def normalize_space(text: str) -> str:
    # \s в str-шаблонах покрывает и неразрывные пробелы (U+00A0, U+202F, U+2009).
    return re.sub(r"\s+", " ", str(text or "")).strip()


def heading_key(text: str) -> str:
    """Текст для сравнения с названием раздела: регистр, ё/е, пробелы и точка в конце не важны."""
    key = normalize_space(text).casefold().replace("ё", "е")
    return re.sub(r"[\s.:;]+$", "", key)


def _to_int(value: Any) -> Optional[int]:
    try:
        return int(str(value))
    except (TypeError, ValueError):
        return None


def roman_to_int(value: str) -> Optional[int]:
    value = value.casefold()
    if not value or any(ch not in ROMAN_VALUES for ch in value):
        return None
    total = 0
    for index, char in enumerate(value):
        current = ROMAN_VALUES[char]
        following = ROMAN_VALUES[value[index + 1]] if index + 1 < len(value) else 0
        total += -current if current < following else current
    canonical = ""
    rest = total
    for number, letters in ((100, "c"), (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")):
        while rest >= number:
            canonical += letters
            rest -= number
    return total if canonical == value and 0 < total <= 60 else None


def _chapter_number(token: str) -> Optional[int]:
    if token.isdigit():
        number = int(token)
        return number if number > 0 else None
    if token in ORDINAL_CHAPTERS:
        return ORDINAL_CHAPTERS[token]
    return roman_to_int(token)


def _match_chapter(key: str, raw: str, strict: bool) -> Optional[Dict[str, Any]]:
    match = _CHAPTER_RE.fullmatch(key)
    if not match:
        return None
    number = _chapter_number(match.group("num"))
    if number is None:
        return None
    rest = match.group("rest")
    if re.match(r"^\.\d", rest):  # «Глава 1.1» — не заголовок главы
        return None
    if rest.strip():
        separated = _SEPARATOR_TITLE_RE.match(rest)
        if separated:
            title = separated.group("title")
        elif not strict and rest[:1].isspace():
            title = rest.strip()
        else:
            return None
        # Обычный абзац из нескольких предложений — не заголовок.
        if strict and re.search(r"[.!?]\s+\S", title):
            return None
    display = re.sub(r"^\s*глава\s+\S+?(?:\s*[.:—–-]\s*|\s+|$)", "", normalize_space(raw), flags=re.IGNORECASE)
    return {"number": number, "title": display}


def _match_appendix(key: str, strict: bool) -> Optional[Dict[str, Any]]:
    match = _APPENDIX_RE.fullmatch(key)
    if not match:
        return None
    rest = match.group("rest")
    number = match.group("num")
    if not rest.strip() or _SEPARATOR_TITLE_RE.match(rest):
        return {"number": number}
    if not strict and number and rest[:1].isspace():
        return {"number": number}
    return None


def classify_block(block: Dict[str, Any]) -> Optional[Tuple[str, Dict[str, Any]]]:
    """Вид границы раздела для блока или None.

    Виды: annotation, contents, introduction, chapter, conclusion, bibliography,
    appendix, appendix_divider, last_sheet, heading (заголовок без известного
    названия). Строки оглавления, пункты списков, код и подписи границами не бывают.
    Абзац без стиля заголовка считается границей, только если целиком совпадает с
    названием раздела («Список требований…» или «Приложение разработано…» — нет).
    """
    if block.get("type") == "tbl":
        if any(OATH_RE.search(cell) for row in block.get("rows", []) for cell in row):
            return "last_sheet", {"oath": True}
        return None
    raw = block.get("text") or ""
    if block.get("toc") or not raw.strip() or looks_like_toc_line(raw):
        return None
    key = heading_key(raw)
    if not key:
        return None
    level = block.get("heading_level")
    styled = level is not None
    if block.get("appendix_label"):
        return "appendix", {"number": None}
    if block.get("appendix_title") or block.get("bibliography_style") or block.get("code") or block.get("caption"):
        return None
    if block.get("in_list") and not styled:
        return None
    if block.get("title_style"):
        return ("last_sheet", {}) if key == LAST_SHEET_KEY else None
    if len(key) > 300 and not styled:
        return None
    loose = styled or bool(block.get("bold")) or bool(block.get("upper"))
    if ANNOTATION_RE.fullmatch(key):
        return "annotation", {}
    if CONTENTS_RE.fullmatch(key):
        return "contents", {}
    if key == "введение":
        return "introduction", {}
    if key == "заключение":
        return "conclusion", {}
    if BIBLIOGRAPHY_RE.fullmatch(key):
        return "bibliography", {}
    if key == "приложения":
        return "appendix_divider", {}
    chapter = _match_chapter(key, raw, strict=not loose)
    if chapter is not None:
        return "chapter", chapter
    appendix = _match_appendix(key, strict=not loose)
    if appendix is not None:
        return "appendix", appendix
    if key == LAST_SHEET_KEY:
        return "last_sheet", {}
    if styled and not CHAPTER_CONCLUSIONS_RE.match(key):
        return "heading", {"level": level}
    return None


def _new_region(kind: str, number: Optional[int] = None, heading: Optional[Dict[str, Any]] = None,
                info: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    return {
        "kind": kind,
        "number": number,
        "heading": normalize_space(heading["text"]) if heading and heading.get("type") == "p" else "",
        "heading_index": heading.get("index") if heading else None,
        "implicit": bool(info and info.get("implicit")),
        "blocks": [],
    }


def build_regions(blocks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Делит блоки на области: front (титул), annotation, contents, introduction,
    chapter, conclusion, bibliography, appendix, appendix_divider, last_sheet, other."""
    classified = [classify_block(block) for block in blocks]
    explicit_chapters = any(found is not None and found[0] == "chapter" for found in classified)
    regions = [_new_region("front")]
    current = regions[0]
    seen_body = False
    max_chapter = 0
    for block, found in zip(blocks, classified):
        kind = found[0] if found else None
        info = dict(found[1]) if found else {}
        if kind == "last_sheet" and not seen_body:
            kind = None  # «Выпускная квалификационная работа» на титуле
        if kind == "heading":
            key = heading_key(block.get("text") or "")
            if (info.get("level") == 1 and current["kind"] in ("introduction", "chapter")
                    and not explicit_chapters and not SUBSECTION_NUMBER_RE.match(key)):
                leading = LEADING_NUMBER_RE.match(key)
                number = int(leading.group("num")) if leading else None
                if number is None or number <= max_chapter:
                    number = max_chapter + 1
                kind, info = "chapter", {"number": number, "implicit": True}
            elif info.get("level") == 1 and current["kind"] in AFTER_BODY_KINDS:
                kind, info = "other", {}
            else:
                kind = None
        if kind is None:
            current["blocks"].append(block)
            continue
        number = None
        if kind == "chapter":
            number = info.get("number") or (max_chapter + 1)
            max_chapter = max(max_chapter, number)
        current = _new_region(kind, number, block, info)
        regions.append(current)
        if kind in BODY_REGION_KINDS or kind == "bibliography":
            seen_body = True
    return regions


def region_name(region: Dict[str, Any]) -> Optional[str]:
    if region["kind"] == "chapter":
        return f"chapter{region['number']}"
    if region["kind"] in ("introduction", "conclusion"):
        return region["kind"]
    return None


def analysis_text(block: Dict[str, Any]) -> Optional[str]:
    """Текст абзаца для анализа или None для заголовков, кода, подписей и служебных строк."""
    if block.get("type") != "p":
        return None
    if (block.get("toc") or block.get("code") or block.get("caption") or block.get("title_style")
            or block.get("appendix_label") or block.get("appendix_title")
            or block.get("heading_level") is not None):
        return None
    raw = block.get("text") or ""
    if looks_like_toc_line(raw):
        return None
    text = normalize_space(raw)
    if not text:
        return None
    if CAPTION_TEXT_RE.match(text) or FIGURE_PLACEHOLDER_RE.match(text):
        return None
    if len(text) < 80 and CHAPTER_CONCLUSIONS_RE.match(heading_key(text)):
        return None
    return text


class DocumentStructure:
    """Блоки тела DOCX по порядку (абзацы и таблицы) и разделы ВКР.

    Учитывает стили с наследованием (basedOn), уровень структуры, автонумерацию
    (numPr), поле TOC на нескольких абзацах, sdt-оглавление, моноширинный шрифт.
    """

    def __init__(self, doc) -> None:
        self.doc = doc
        self._styles, self._default_style = self._load_styles(doc)
        self._props_cache: Dict[Optional[str], Dict[str, Any]] = {}
        self._fields: List[Dict[str, Any]] = []
        self.blocks: List[Dict[str, Any]] = []
        self.skipped_toc_sdt = 0
        self._collect(doc.element.body)
        for index, block in enumerate(self.blocks):
            block["index"] = index
        self.regions = build_regions(self.blocks)

    # --- стили -------------------------------------------------------------
    @staticmethod
    def _load_styles(doc) -> Tuple[Dict[str, Dict[str, Any]], Optional[str]]:
        styles: Dict[str, Dict[str, Any]] = {}
        default = None
        try:
            root = doc.styles.element
        except Exception:  # pragma: no cover - пакет без стилей
            return styles, None
        for element in root.findall(_w("style")):
            style_id = element.get(_w("styleId"))
            if not style_id:
                continue
            name = element.find(_w("name"))
            based = element.find(_w("basedOn"))
            info: Dict[str, Any] = {
                "name": (name.get(_VAL) if name is not None else None) or style_id,
                "based": based.get(_VAL) if based is not None else None,
                "outline": None, "num": None, "font": None, "bold": None,
            }
            ppr = element.find(_PPR)
            if ppr is not None:
                outline = ppr.find(_OUTLINE)
                if outline is not None:
                    info["outline"] = _to_int(outline.get(_VAL))
                num_pr = ppr.find(_NUMPR)
                if num_pr is not None and num_pr.find(_NUMID) is not None:
                    info["num"] = num_pr.find(_NUMID).get(_VAL)
            rpr = element.find(_RPR)
            if rpr is not None:
                fonts = rpr.find(_RFONTS)
                if fonts is not None:
                    info["font"] = fonts.get(_w("ascii")) or fonts.get(_w("hAnsi"))
                bold = rpr.find(_BOLD)
                if bold is not None:
                    info["bold"] = (bold.get(_VAL) or "true").casefold() not in ("0", "false", "off")
            if element.get(_w("type")) == "paragraph" and (element.get(_w("default")) or "").casefold() in ("1", "true", "on"):
                default = style_id
            styles[style_id] = info
        return styles, default

    def _props(self, style_id: Optional[str]) -> Dict[str, Any]:
        key = style_id if style_id in self._styles else self._default_style
        if key in self._props_cache:
            return self._props_cache[key]
        chain: List[Dict[str, Any]] = []
        seen = set()
        current = key
        while current and current in self._styles and current not in seen:
            seen.add(current)
            chain.append(self._styles[current])
            current = self._styles[current]["based"]
        props = {
            "names": [str(item["name"]).casefold() for item in chain],
            "outline": next((item["outline"] for item in chain if item["outline"] is not None), None),
            "num": next((item["num"] for item in chain if item["num"] is not None), None),
            "font": next((item["font"] for item in chain if item["font"]), None),
            "bold": next((item["bold"] for item in chain if item["bold"] is not None), None),
        }
        self._props_cache[key] = props
        return props

    # --- обход тела ----------------------------------------------------------
    def _collect(self, parent) -> None:
        for child in parent:
            tag = child.tag
            if tag == _P:
                self.blocks.append(self._read_paragraph(child))
            elif tag == _TBL:
                self.blocks.append(self._read_table(child))
            elif tag == _SDT:
                gallery = child.find(f"{_SDT_PR}/{_w('docPartObj')}/{_w('docPartGallery')}")
                if gallery is not None and "table of contents" in (gallery.get(_VAL) or "").casefold():
                    self.skipped_toc_sdt += 1
                    continue
                content = child.find(_SDT_CONTENT)
                if content is not None:
                    self._collect(content)
            elif tag == _CUSTOM_XML:
                self._collect(child)

    def _new_state(self) -> Dict[str, Any]:
        return {"parts": [], "toc": any(field["toc"] for field in self._fields), "mono": [], "bold": []}

    def _walk(self, element, state: Dict[str, Any], props: Dict[str, Any]) -> None:
        for child in element:
            tag = child.tag
            if not isinstance(tag, str) or tag in _SKIP_IN_PARAGRAPH:
                continue
            if tag == _R:
                self._run(child, state, props)
            elif tag == _FLDSIMPLE:
                if (child.get(_w("instr")) or "").strip().upper().startswith("TOC"):
                    state["toc"] = True
                self._walk(child, state, props)
            elif tag == _MC_ALTERNATE:
                choice = child.find(_MC_CHOICE)
                if choice is not None:
                    self._walk(choice, state, props)
            elif tag in (_TBL,):
                continue
            else:
                self._walk(child, state, props)

    def _run(self, run, state: Dict[str, Any], props: Dict[str, Any]) -> None:
        font = None
        bold = None
        rpr = run.find(_RPR)
        if rpr is not None:
            fonts = rpr.find(_RFONTS)
            if fonts is not None:
                font = fonts.get(_w("ascii")) or fonts.get(_w("hAnsi"))
            bold_el = rpr.find(_BOLD)
            if bold_el is not None:
                bold = (bold_el.get(_VAL) or "true").casefold() not in ("0", "false", "off")
            run_style = rpr.find(_RSTYLE)
            if run_style is not None and run_style.get(_VAL) in self._styles:
                char_props = self._props(run_style.get(_VAL))
                font = font or char_props["font"]
                if bold is None:
                    bold = char_props["bold"]
        font = font or props["font"]
        if bold is None:
            bold = props["bold"]
        mono = bool(font) and str(font).strip().casefold() in MONO_FONTS
        for child in run:
            tag = child.tag
            if tag == _FLDCHAR:
                kind = child.get(_w("fldCharType"))
                if kind == "begin":
                    self._fields.append({"instr": "", "toc": False, "separated": False})
                elif kind == "separate" and self._fields:
                    self._fields[-1]["separated"] = True
                elif kind == "end" and self._fields:
                    self._fields.pop()
                continue
            if tag == _INSTR:
                if self._fields:
                    self._fields[-1]["instr"] += child.text or ""
                    if self._fields[-1]["instr"].lstrip().upper().startswith("TOC"):
                        self._fields[-1]["toc"] = True
                        state["toc"] = True
                continue
            if any(field["toc"] for field in self._fields):
                state["toc"] = True
            if self._fields and not all(field["separated"] for field in self._fields):
                continue  # текст кода поля, а не результата
            if tag == _T:
                text = child.text or ""
            elif tag == _TAB:
                text = "\t"
            elif tag in (_BR, _CR):
                text = "" if child.get(_w("type")) in ("page", "column") else "\n"
            elif tag == _NO_BREAK_HYPHEN:
                text = "-"
            else:
                continue
            if text:
                state["parts"].append(text)
                if text.strip():
                    state["mono"].append(mono)
                    state["bold"].append(bool(bold))

    def _read_paragraph(self, element) -> Dict[str, Any]:
        style_id = None
        direct_outline = None
        direct_num = None
        ppr = element.find(_PPR)
        if ppr is not None:
            pstyle = ppr.find(_PSTYLE)
            if pstyle is not None:
                style_id = pstyle.get(_VAL)
            outline = ppr.find(_OUTLINE)
            if outline is not None:
                direct_outline = _to_int(outline.get(_VAL))
            num_pr = ppr.find(_NUMPR)
            if num_pr is not None and num_pr.find(_NUMID) is not None:
                direct_num = num_pr.find(_NUMID).get(_VAL)
        props = self._props(style_id)
        state = self._new_state()
        self._walk(element, state, props)
        text = "".join(state["parts"])
        names = props["names"]
        name = names[0] if names else ""
        outline = direct_outline if direct_outline is not None else props["outline"]
        num = direct_num if direct_num is not None else props["num"]
        in_list = num is not None and str(num).strip() not in ("", "0")
        if not in_list and name.startswith(("list number", "list bullet")):
            in_list = True
        level = None
        match = re.match(r"^(?:heading|заголовок)\s*([1-9])$", name)
        if match:
            level = int(match.group(1))
        elif outline is not None and 0 <= outline < 9:
            level = outline + 1
        letters = [ch for ch in text if ch.isalpha()]
        return {
            "type": "p",
            "text": text,
            "style": name,
            "heading_level": level,
            "in_list": in_list,
            "toc": bool(state["toc"]) or bool(TOC_STYLE_RE.match(name)),
            "code": any(CODE_STYLE_RE.search(item) for item in names) or (bool(state["mono"]) and all(state["mono"])),
            "caption": any(CAPTION_STYLE_RE.match(item) for item in names),
            "title_style": name == "vkr title",
            "bibliography_style": name == "vkr bibliography",
            "appendix_label": name == "vkr appendix label",
            "appendix_title": name == "vkr appendix title",
            "bold": bool(state["bold"]) and all(state["bold"]),
            "upper": len(letters) >= 4 and all(ch.isupper() for ch in letters),
        }

    def _read_table(self, element) -> Dict[str, Any]:
        rows: List[List[str]] = []
        default_props = self._props(None)
        for row in element.findall(_TR):
            cells: List[str] = []
            for cell in row.findall(_TC):
                texts = []
                for paragraph in cell.iter(_P):
                    state = self._new_state()
                    self._walk(paragraph, state, default_props)
                    text = normalize_space("".join(state["parts"]))
                    if text:
                        texts.append(text)
                cells.append(" ".join(texts))
            rows.append(cells)
        return {"type": "tbl", "rows": rows, "text": " ".join(cell for row in rows for cell in row if cell)}

    # --- разделы -------------------------------------------------------------
    def section_regions(self) -> "OrderedDict[str, List[Dict[str, Any]]]":
        result: "OrderedDict[str, List[Dict[str, Any]]]" = OrderedDict()
        for region in self.regions:
            name = region_name(region)
            if name:
                result.setdefault(name, []).append(region)
        return result

    def section_names(self) -> List[str]:
        return list(self.section_regions().keys())

    def has_body(self) -> bool:
        return bool(self.section_names())

    def section_texts(self, name: str) -> List[str]:
        texts: List[str] = []
        for region in self.section_regions().get(str(name).strip().casefold(), []):
            for block in region["blocks"]:
                text = analysis_text(block)
                if text:
                    texts.append(text)
        return texts

    def body_items(self) -> List[Tuple[str, str]]:
        """(раздел, текст абзаца) введения, глав и заключения в порядке документа."""
        items: List[Tuple[str, str]] = []
        for region in self.regions:
            name = region_name(region)
            if not name:
                continue
            for block in region["blocks"]:
                text = analysis_text(block)
                if text:
                    items.append((name, text))
        return items

    def fallback_texts(self, include_tables: bool = True) -> List[str]:
        """Текст документа с нераспознанной структурой без служебных областей."""
        texts: List[str] = []
        for region in self.regions:
            if region["kind"] in NON_TEXT_REGION_KINDS:
                continue
            for block in region["blocks"]:
                if block.get("type") == "tbl":
                    if include_tables:
                        texts.extend(cell for row in block["rows"] for cell in row if cell)
                    continue
                text = analysis_text(block)
                if text:
                    texts.append(text)
        return texts

    def expected_sections(self) -> List[str]:
        chapters = [int(name[len("chapter"):]) for name in self.section_names() if name.startswith("chapter")]
        top = max(chapters) if chapters else 1
        return ["introduction"] + [f"chapter{number}" for number in range(1, top + 1)] + ["conclusion"]

    def bibliography_regions(self) -> List[Dict[str, Any]]:
        return [region for region in self.regions if region["kind"] == "bibliography"]

    def summary(self) -> Dict[str, Any]:
        return {
            "structure": "recognized" if self.has_body() else "unrecognized",
            "sections_found": self.section_names(),
            "implicit_chapters": [region_name(r) for r in self.regions if r["kind"] == "chapter" and r["implicit"]],
            "bibliography_found": bool(self.bibliography_regions()),
            "appendices": sum(1 for region in self.regions if region["kind"] == "appendix"),
        }


def read_structure(source) -> DocumentStructure:
    if isinstance(source, DocumentStructure):
        return source
    return DocumentStructure(open_document(source))


def count_words(text: str) -> int:
    return len(re.findall(r"\b\w+\b", text, re.UNICODE))


# ========== АНАЛИЗ ТЕКСТА ==========

def split_sentences(text: str) -> list:
    """Разбить текст на предложения."""
    sentences = re.split(r'(?<=[.!?])\s+', text)
    return [s.strip() for s in sentences if s.strip() and len(s) > 5]


def split_paragraphs_from_doc(doc) -> list:
    """Содержательное тело ВКР (введение, главы, заключение) без титула, TOC,
    библиографии, приложений, кода и подписей. Для нераспознанной структуры —
    весь текст без служебных областей, включая таблицы: вызывающая сторона
    проверит достаточность объёма и не выдаст ложный LOW."""
    structure = read_structure(doc)
    if structure.has_body():
        return [text for _, text in structure.body_items() if len(text) > 30]
    return [text for text in structure.fallback_texts() if len(text) >= 30]


def extract_section(doc, section_name: str) -> tuple:
    """Раздел по заголовкам: (полный текст, абзацы длиннее 30 символов)."""
    structure = read_structure(doc)
    texts = structure.section_texts(section_name)
    full_text = "\n\n".join(texts)
    paragraphs = [text for text in texts if len(text.strip()) > 30]
    return full_text, paragraphs


def compute_burstiness(paragraphs: list) -> dict:
    """
    Burstiness = CV (coefficient of variation) длин.
    Человеческий текст: CV > 0.35
    ИИ-текст: CV < 0.25 обычно.
    """
    if not paragraphs:
        return {"error": "Нет абзацев для анализа"}

    # Длина абзацев в символах
    para_lens = [len(p) for p in paragraphs]
    para_mean = statistics.mean(para_lens) if para_lens else 0
    para_stdev = statistics.stdev(para_lens) if len(para_lens) > 1 else 0
    para_cv = para_stdev / para_mean if para_mean > 0 else 0

    # Длина предложений в словах
    all_sentences = []
    for p in paragraphs:
        all_sentences.extend(split_sentences(p))

    sent_lens = [len(s.split()) for s in all_sentences]
    sent_mean = statistics.mean(sent_lens) if sent_lens else 0
    sent_stdev = statistics.stdev(sent_lens) if len(sent_lens) > 1 else 0
    sent_cv = sent_stdev / sent_mean if sent_mean > 0 else 0

    # Оценка
    verdict = "human-like"
    if para_cv < 0.25 or sent_cv < 0.30:
        verdict = "AI-like (too uniform)"
    elif para_cv < 0.35 or sent_cv < 0.40:
        verdict = "slightly suspicious"

    return {
        "paragraph_count": len(paragraphs),
        "paragraph_mean_chars": round(para_mean, 1),
        "paragraph_stdev_chars": round(para_stdev, 1),
        "paragraph_cv": round(para_cv, 3),
        "sentence_count": len(all_sentences),
        "sentence_mean_words": round(sent_mean, 1),
        "sentence_stdev_words": round(sent_stdev, 1),
        "sentence_cv": round(sent_cv, 3),
        "verdict": verdict,
    }


def _normalize_for_match(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "").casefold().replace("ё", "е"))


def compute_ngram_density(text: str, allowlist: Optional[Sequence[str]] = None) -> dict:
    """Плотность характерных ИИ-N-грамм (фразы из allowlist не считаются)."""
    allowed = {normalize_phrase(item) for item in (CLICHE_ALLOWLIST if allowlist is None else allowlist)}
    text_lower = _normalize_for_match(text)
    total_words = len(text.split())

    strong_hits = []
    for ng in AI_NGRAMS_STRONG:
        normalized = normalize_phrase(ng)
        if normalized in allowed:
            continue
        count = len(re.findall(rf"(?<!\w){re.escape(normalized)}(?!\w)", text_lower))
        if count > 0:
            strong_hits.append((ng, count))

    medium_hits = []
    for ng in AI_NGRAMS_MEDIUM:
        normalized = normalize_phrase(ng)
        if normalized in allowed:
            continue
        count = len(re.findall(rf"(?<!\w){re.escape(normalized)}(?!\w)", text_lower))
        if count > 0:
            medium_hits.append((ng, count))

    total_strong = sum(c for _, c in strong_hits)
    total_medium = sum(c for _, c in medium_hits)

    # Плотность: клише на 1000 слов
    strong_per_1k = (total_strong / total_words * 1000) if total_words > 0 else 0
    medium_per_1k = (total_medium / total_words * 1000) if total_words > 0 else 0

    # Оценка
    verdict = "acceptable"
    if strong_per_1k > 0.5:
        verdict = "AI-like (too many strong cliches)"
    elif strong_per_1k > 0.2 or medium_per_1k > 3.0:
        verdict = "slightly suspicious"

    return {
        "total_words": total_words,
        "strong_cliches_count": total_strong,
        "medium_cliches_count": total_medium,
        "strong_per_1k_words": round(strong_per_1k, 2),
        "medium_per_1k_words": round(medium_per_1k, 2),
        "top_strong": sorted(strong_hits, key=lambda x: -x[1])[:5],
        "top_medium": sorted(medium_hits, key=lambda x: -x[1])[:5],
        "verdict": verdict,
    }


def compute_lexical_diversity(text: str) -> dict:
    """
    Лексическое разнообразие: TTR (type-token ratio) и MTLD-like метрика.
    Человеческий текст на 10000 слов: TTR ~0.35-0.55
    ИИ-текст: TTR ниже, словарь беднее.
    """
    # Нормализация: только слова, в нижнем регистре
    words = re.findall(r'\b[а-яёa-z]+\b', text.lower())
    total_words = len(words)

    if total_words < 100:
        return {"error": "Слишком мало слов для оценки"}

    unique_words = set(words)
    ttr = len(unique_words) / total_words

    # MTLD-подобная метрика: средняя длина «отрезка» с TTR > 0.72
    # (упрощённый вариант)
    def mtld_segment(tokens, threshold=0.72):
        types = set()
        token_count = 0
        factors = 0
        for w in tokens:
            types.add(w)
            token_count += 1
            if token_count > 0 and len(types) / token_count < threshold:
                factors += 1
                types = set()
                token_count = 0
        if token_count > 0:
            factors += (1 - len(types) / token_count) / (1 - threshold)
        return len(tokens) / factors if factors > 0 else len(tokens)

    mtld = mtld_segment(words)

    # Проверка на пересыщение "слов-паразитов" ИИ
    overused_count = sum(1 for w in words if w in AI_WORDS_OVERUSED)
    overused_per_1k = (overused_count / total_words * 1000) if total_words > 0 else 0

    verdict = "diverse"
    if ttr < 0.18 or mtld < 45:
        verdict = "AI-like (poor vocabulary)"
    elif ttr < 0.25 or mtld < 60:
        verdict = "slightly poor"

    if overused_per_1k > 12:
        verdict += "; overused buzzwords"

    return {
        "total_words": total_words,
        "unique_words": len(unique_words),
        "type_token_ratio": round(ttr, 3),
        "mtld_approx": round(mtld, 1),
        "overused_words_per_1k": round(overused_per_1k, 2),
        "verdict": verdict,
    }


def compute_author_anchors(text: str, personal: Optional[Sequence[str]] = None) -> dict:
    """Плотность авторских якорьков — не перебор ли?

    Считаются личные якорьки из personal_style_profile state-файла (``personal``; по умолчанию —
    найденные для анализируемого DOCX) и встроенный список общих оборотов.
    """
    text_lower = _normalize_for_match(text)
    total_words = len(text.split())
    personal_list = [normalize_phrase(item) for item in (PERSONAL_ANCHORS if personal is None else personal) if normalize_phrase(item)]

    anchor_hits = []
    seen = set()
    for anchor in list(personal_list) + list(AUTHOR_ANCHORS):
        normalized = normalize_phrase(anchor)
        if normalized in seen:
            continue
        seen.add(normalized)
        count = len(re.findall(rf"(?<!\w){re.escape(normalized)}(?!\w)", text_lower))
        if count > 0:
            anchor_hits.append((anchor, count))

    total = sum(c for _, c in anchor_hits)
    per_1k = (total / total_words * 1000) if total_words > 0 else 0

    verdict = "balanced"
    if per_1k > 1.5:
        verdict = "too many (sounds like PhD, not bachelor)"
    elif per_1k < 0.2 and total_words > 3000:
        verdict = "no author voice (sounds AI-generic)"

    return {
        "total_anchors": total,
        "unique_anchors_used": len(anchor_hits),
        "anchors_per_1k_words": round(per_1k, 2),
        "hits": sorted(anchor_hits, key=lambda x: -x[1]),
        "personal_anchors": personal_list,
        "anchor_source": "personal_style_profile+builtin" if personal_list else "builtin",
        "verdict": verdict,
    }


def check_pronouns(text: str) -> dict:
    """Проверка личных местоимений (должно быть 0), включая «наш/наша/наши…».

    Важно: инициалы типа «И.Я.», «Я.Л.» удаляются из текста перед
    проверкой, иначе regex \\bя\\b с re.IGNORECASE триггерит на «Я»
    в инициалах педагогов-классиков (Лернер И.Я., Коломинский Я.Л.).
    """
    # Фильтрация инициалов: «И.Я.», «И. Я.», «А.Б.В.»
    initial_pattern = re.compile(
        r"\b[А-ЯЁA-Z]\.\s?[А-ЯЁA-Z]\.(?:\s?[А-ЯЁA-Z]\.)?"
    )
    clean_text = initial_pattern.sub(" ", text)
    # Одиночные инициалы перед фамилией
    clean_text = re.sub(r"\b[А-ЯЁA-Z]\.", " ", clean_text)

    found = []
    for label, pat in PERSONAL_PRONOUNS:
        matches = re.findall(pat, clean_text, re.IGNORECASE)
        if matches:
            found.append((label, len(matches)))

    return {
        "found_pronouns": found,
        "total": sum(c for _, c in found),
        "verdict": "clean" if not found else "BAD: personal pronouns present",
    }


def check_paragraph_rhythm(paragraphs: list) -> dict:
    """Проверка на монотонность ритма — не идут ли подряд одинаковые абзацы."""
    if len(paragraphs) < 5:
        return {"skipped": "Too few paragraphs"}

    lens = [len(p) for p in paragraphs]
    # Ищем группы подряд идущих абзацев с похожей длиной (±20%)
    max_run = 1
    current_run = 1
    for i in range(1, len(lens)):
        prev = lens[i-1]
        curr = lens[i]
        if prev > 0 and abs(curr - prev) / prev < 0.20:
            current_run += 1
            max_run = max(max_run, current_run)
        else:
            current_run = 1

    verdict = "good"
    if max_run >= 5:
        verdict = "AI-like (long run of similar paragraphs)"
    elif max_run >= 4:
        verdict = "slightly monotonous"

    return {
        "max_consecutive_similar_paragraphs": max_run,
        "verdict": verdict,
    }


def check_sentence_starter_variety(paragraphs: list) -> dict:
    """
    LLM часто начинает предложения одинаковыми словами.
    Проверка: разнообразие первых 1-2 слов предложений.
    """
    if not paragraphs:
        return {"error": "No paragraphs"}

    all_sentences = []
    for p in paragraphs:
        all_sentences.extend(split_sentences(p))

    if len(all_sentences) < 10:
        return {"skipped": "Too few sentences"}

    starters = []
    for s in all_sentences:
        words = s.split()[:2]
        starter = " ".join(words).lower()
        # Убираем пунктуацию
        starter = re.sub(r'[^\w\s]', '', starter)
        if starter:
            starters.append(starter)

    unique_starters = set(starters)
    variety_ratio = len(unique_starters) / len(starters) if starters else 0

    # Top-3 повторяющихся
    starter_counts = {}
    for s in starters:
        starter_counts[s] = starter_counts.get(s, 0) + 1
    top_repeated = sorted(
        [(s, c) for s, c in starter_counts.items() if c >= 3],
        key=lambda x: -x[1]
    )[:5]

    verdict = "diverse"
    if variety_ratio < 0.50:
        verdict = "AI-like (monotonous starters)"
    elif variety_ratio < 0.65:
        verdict = "slightly monotonous"

    return {
        "total_sentences": len(starters),
        "unique_starters": len(unique_starters),
        "variety_ratio": round(variety_ratio, 3),
        "top_repeated": top_repeated,
        "verdict": verdict,
    }


# Краткие страдательные причастия, типичные для научного стиля (ё приводится к е).
PASSIVE_PARTICIPLE_STEMS = (
    "проведен", "получен", "выявлен", "установлен", "реализован", "рассмотрен", "выполнен",
    "построен", "определен", "изучен", "разработан", "предложен", "применен", "внедрен",
    "описан", "представлен", "приведен", "рассчитан", "сформулирован", "обоснован",
    "апробирован", "достигнут", "отражен", "использован", "создан", "сделан", "проанализирован",
)
# «был/будет + краткое причастие» (-ан/-ян/-ен/-ыт/-ит/-ят/-ут); существительное
# или прилагательное рядом с «был» («Работа была интересной», «Он был в школе»)
# пассивом не считается.
PASSIVE_RE = re.compile(
    r"\b(?:был|была|было|были|будет|будут|бывает|бывают)\s+\w+?(?:ан|ян|ен|ыт|ит|ят|ут)[аоы]?\b"
    r"|\b(?:" + "|".join(PASSIVE_PARTICIPLE_STEMS) + r")[аоы]?\b"
)


def check_passive_voice_ratio(text: str) -> dict:
    """
    Пассивный залог: LLM его переиспользует. Человек чередует.
    Эвристика: «был/была/было/были/будет + краткое причастие» или отдельные
    краткие причастия («проведён/проведено/получены»); ё и е не различаются.
    """
    passive_markers = PASSIVE_RE.findall(text.lower().replace("ё", "е"))

    total_sentences = len(split_sentences(text))
    if total_sentences == 0:
        return {"skipped": "No sentences"}

    ratio = len(passive_markers) / total_sentences

    verdict = "balanced"
    if ratio > 0.70:
        verdict = "AI-like (excessive passive voice)"
    elif ratio > 0.50:
        verdict = "slightly heavy passive"
    elif ratio < 0.15:
        # В научном тексте так низко не бывает
        verdict = "unusually low passive (unusual for academic text)"

    return {
        "passive_markers_count": len(passive_markers),
        "sentences_count": total_sentences,
        "ratio": round(ratio, 3),
        "verdict": verdict,
    }


def check_readability_proxy(paragraphs: list) -> dict:
    """
    Простая proxy-метрика читаемости.
    LLM часто пишет либо слишком ровно (все предложения 15-20 слов),
    либо неестественно сложно (все 25+).

    Считаем: средняя длина слова и распределение длин предложений.
    """
    all_sentences = []
    for p in paragraphs:
        all_sentences.extend(split_sentences(p))
    if len(all_sentences) < 10:
        return {"skipped": "Too few sentences"}

    # Распределение длин предложений
    sent_lens_words = [len(s.split()) for s in all_sentences]

    # Доля «коротких» (<8 слов) и «длинных» (>25 слов)
    short_ratio = sum(1 for l in sent_lens_words if l < 8) / len(sent_lens_words)
    long_ratio = sum(1 for l in sent_lens_words if l > 25) / len(sent_lens_words)
    medium_ratio = 1 - short_ratio - long_ratio

    # Средняя длина слова
    all_words = []
    for s in all_sentences:
        all_words.extend([w for w in re.findall(r'\b[а-яёa-z]+\b', s.lower()) if len(w) > 1])
    avg_word_len = sum(len(w) for w in all_words) / len(all_words) if all_words else 0

    verdict = "natural"
    if short_ratio < 0.05 and long_ratio < 0.10:
        # Всё в диапазоне 8-25 слов = слишком ровно
        verdict = "AI-like (all sentences same length range)"
    elif medium_ratio > 0.90:
        verdict = "slightly uniform"

    return {
        "short_sentences_pct": round(short_ratio * 100, 1),
        "medium_sentences_pct": round(medium_ratio * 100, 1),
        "long_sentences_pct": round(long_ratio * 100, 1),
        "avg_word_length_chars": round(avg_word_len, 2),
        "verdict": verdict,
    }


# ========== AGGREGATE RISK SCORE ==========

# Посекционные множители весов: теоретическая глава звучит академичнее по природе,
# проектная должна быть живее и с конкретикой.
# Ключ — тип секции. Значения — множители для категорий риска.
# Если множитель < 1.0 — секция более терпима к этому виду риска.
# Если множитель > 1.0 — секция жёстче штрафуется за этот вид риска.
SECTION_WEIGHT_MULTIPLIERS = {
    # Полный документ — стандартные веса
    "full": {
        "burstiness": 1.0, "ngrams": 1.0, "lexical": 1.0, "rhythm": 1.0,
        "starters": 1.0, "passive": 1.0, "readability": 1.0, "anchors": 1.0,
    },
    # Введение — строгие пороги, это визитка работы. Клише и пассив особенно штрафуются.
    "introduction": {
        "burstiness": 1.0, "ngrams": 1.3, "lexical": 1.0, "rhythm": 1.0,
        "starters": 1.0, "passive": 1.3, "readability": 1.0, "anchors": 0.7,
    },
    # Заключение — без ссылок и цитат, по методичке формальное и сухое.
    # Якорьки тут наоборот штрафуются (если «представляется обоснованным» в заключении — звучит фальшиво).
    "conclusion": {
        "burstiness": 0.8, "ngrams": 1.2, "lexical": 0.8, "rhythm": 0.8,
        "starters": 0.8, "passive": 1.2, "readability": 0.8, "anchors": 1.5,
    },
    # Теоретическая глава — по природе академичная, шаблонность более допустима.
    # Burstiness, ритм и разнообразие стартов тут смягчены — в теории это норма.
    "chapter1": {
        "burstiness": 0.7, "ngrams": 1.0, "lexical": 0.8, "rhythm": 0.7,
        "starters": 0.7, "passive": 1.0, "readability": 0.8, "anchors": 0.8,
    },
    # Аналитическая глава — с собственными данными, должна быть живее теоретической.
    "chapter2": {
        "burstiness": 1.0, "ngrams": 1.0, "lexical": 1.0, "rhythm": 1.0,
        "starters": 1.0, "passive": 1.0, "readability": 1.0, "anchors": 1.0,
    },
    # Проектная глава — главный полигон конкретики (версии, цифры, решения).
    # Здесь шаблонность бьёт сильнее: если описание продукта звучит как ИИ — это «красный флаг».
    "chapter3": {
        "burstiness": 1.2, "ngrams": 1.3, "lexical": 1.2, "rhythm": 1.2,
        "starters": 1.2, "passive": 1.0, "readability": 1.0, "anchors": 1.0,
    },
}

# Посекционные пороги уровней риска. Теоретическая глава может иметь больший score
# при том же уровне риска — потому что теория по природе более шаблонна.
SECTION_LEVEL_THRESHOLDS = {
    # "full" / "introduction" / "chapter2" / "chapter3" — стандартные
    "full":         {"LOW": 20, "MODERATE": 40, "HIGH": 60},
    "introduction": {"LOW": 18, "MODERATE": 36, "HIGH": 55},  # строже: визитка работы
    "conclusion":   {"LOW": 18, "MODERATE": 36, "HIGH": 55},  # строже
    "chapter1":     {"LOW": 25, "MODERATE": 48, "HIGH": 68},  # мягче: теория академична
    "chapter2":     {"LOW": 20, "MODERATE": 40, "HIGH": 60},
    "chapter3":     {"LOW": 22, "MODERATE": 42, "HIGH": 62},
}


def aggregate_risk(results: dict, section_type: str = "full") -> dict:
    """Сводная оценка риска в баллах от 0 до 100.

    `section_type`: 'full' | 'introduction' | 'conclusion' | 'chapterN'.
    Разные секции имеют разные веса штрафов и разные пороги уровней риска,
    потому что теоретическая глава по природе шаблоннее проектной, и одни
    и те же пороги на весь документ дают искажённую картину. Для глав после
    третьей применяются стандартные веса и пороги.
    """
    weights = SECTION_WEIGHT_MULTIPLIERS.get(section_type, SECTION_WEIGHT_MULTIPLIERS["full"])
    thresholds = SECTION_LEVEL_THRESHOLDS.get(section_type, SECTION_LEVEL_THRESHOLDS["full"])

    score = 0.0
    flags = []

    # Burstiness
    b = results.get("burstiness", {})
    if "too uniform" in b.get("verdict", ""):
        score += 30 * weights["burstiness"]
        flags.append("LOW burstiness (too uniform)")
    elif "suspicious" in b.get("verdict", ""):
        score += 15 * weights["burstiness"]
        flags.append("Moderate burstiness risk")

    # N-grams
    n = results.get("ngrams", {})
    if "too many strong" in n.get("verdict", ""):
        score += 30 * weights["ngrams"]
        flags.append("Too many AI strong cliches")
    elif "suspicious" in n.get("verdict", ""):
        score += 15 * weights["ngrams"]
        flags.append("Elevated AI cliches")

    # Lexical diversity
    l = results.get("lexical", {})
    if "AI-like" in l.get("verdict", ""):
        score += 20 * weights["lexical"]
        flags.append("Poor lexical diversity")
    elif "poor" in l.get("verdict", "") and "AI" not in l.get("verdict", ""):
        score += 10 * weights["lexical"]
        flags.append("Somewhat limited vocabulary")
    if "overused" in l.get("verdict", ""):
        score += 5 * weights["lexical"]
        flags.append("Overused AI buzzwords")

    # Paragraph rhythm
    r = results.get("rhythm", {})
    if "AI-like" in r.get("verdict", ""):
        score += 15 * weights["rhythm"]
        flags.append("Monotonous paragraph rhythm")

    # Sentence starter variety
    starters = results.get("starters", {})
    if "AI-like" in starters.get("verdict", ""):
        score += 15 * weights["starters"]
        flags.append("Monotonous sentence starters")
    elif "slightly monotonous" in starters.get("verdict", ""):
        score += 7 * weights["starters"]
        flags.append("Somewhat repetitive sentence starts")

    # Passive voice
    passive = results.get("passive", {})
    if "excessive passive" in passive.get("verdict", ""):
        score += 10 * weights["passive"]
        flags.append("Excessive passive voice (AI tendency)")

    # Readability
    readability = results.get("readability", {})
    if "same length range" in readability.get("verdict", ""):
        score += 10 * weights["readability"]
        flags.append("All sentences same length range (AI tendency)")

    # Author anchors (in BOTH directions)
    a = results.get("anchors", {})
    if "too many" in a.get("verdict", ""):
        score += 10 * weights["anchors"]
        flags.append("Too many author anchors (sounds like PhD)")
    elif "no author voice" in a.get("verdict", ""):
        score += 10 * weights["anchors"]
        flags.append("No author voice (sounds AI-generic)")

    # Pronouns — одинаково по всем секциям (хард-правило методички)
    p = results.get("pronouns", {})
    if p.get("total", 0) > 0:
        score += 5
        flags.append("Personal pronouns found")

    # Score capping и округление
    score = min(round(score, 1), 100)

    # Посекционные пороги уровня риска
    if score < thresholds["LOW"]:
        risk_level = "LOW"
    elif score < thresholds["MODERATE"]:
        risk_level = "MODERATE"
    elif score < thresholds["HIGH"]:
        risk_level = "HIGH"
    else:
        risk_level = "CRITICAL"

    return {
        "risk_score": score,
        "risk_level": risk_level,
        "section_type": section_type,
        "flags": flags,
    }


def analyze_text(text: str, paragraphs: list, section_type: str = "full") -> dict:
    results = {
        "burstiness": compute_burstiness(paragraphs),
        "ngrams": compute_ngram_density(text),
        "lexical": compute_lexical_diversity(text),
        "anchors": compute_author_anchors(text),
        "pronouns": check_pronouns(text),
        "rhythm": check_paragraph_rhythm(paragraphs),
        "starters": check_sentence_starter_variety(paragraphs),
        "passive": check_passive_voice_ratio(text),
        "readability": check_readability_proxy(paragraphs),
    }
    results["risk"] = aggregate_risk(results, section_type=section_type)
    return results


def insufficient_result(section_type: str, message: str) -> dict:
    return {
        "risk": {
            "risk_score": None,
            "risk_level": "INSUFFICIENT_DATA",
            "section_type": section_type,
            "flags": [message],
        }
    }


# ========== ОТЧЁТ ==========

def print_report(results: dict):
    """Текстовый отчёт."""
    print("\n" + "=" * 70)
    print("  ЭВРИСТИЧЕСКИЙ ДЕТЕКТОР ИИ-ТЕКСТА")
    print("=" * 70)

    risk = results["risk"]
    level = risk["risk_level"]

    print(f"\nРиск: {level} ({risk['risk_score']}/100)")
    if risk["flags"]:
        print("\nТревожные сигналы:")
        for f in risk["flags"]:
            print(f"  • {f}")

    print("\nBurstiness (рваность текста):")
    b = results["burstiness"]
    print(f"   Абзацев: {b.get('paragraph_count', '?')}")
    print(f"   Средняя длина абзаца: {b.get('paragraph_mean_chars', '?')} символов")
    print(f"   CV абзацев: {b.get('paragraph_cv', '?')} (цель: >0.35)")
    print(f"   CV предложений: {b.get('sentence_cv', '?')} (цель: >0.40)")
    print(f"   Вердикт: {b.get('verdict', '?')}")

    print("\nПлотность ИИ-клише:")
    n = results["ngrams"]
    print(f"   Сильных клише: {n.get('strong_cliches_count', '?')} ({n.get('strong_per_1k_words', '?')}/1000 слов)")
    print(f"   Средних клише: {n.get('medium_cliches_count', '?')} ({n.get('medium_per_1k_words', '?')}/1000 слов)")
    if n.get("top_strong"):
        print("   Топ сильных:")
        for ng, c in n["top_strong"]:
            print(f"     • «{ng}»: {c} раз")
    print(f"   Вердикт: {n.get('verdict', '?')}")

    print("\nЛексическое разнообразие:")
    l = results["lexical"]
    print(f"   Type-Token Ratio: {l.get('type_token_ratio', '?')} (цель: >0.25)")
    print(f"   MTLD-approx: {l.get('mtld_approx', '?')} (цель: >60)")
    print(f"   «Буззворды» на 1k слов: {l.get('overused_words_per_1k', '?')} (норма ≤12)")
    print(f"   Вердикт: {l.get('verdict', '?')}")

    print("\nАвторские якорьки:")
    a = results["anchors"]
    if a.get("personal_anchors"):
        print(f"   Личные якорьки из personal_style_profile: {len(a['personal_anchors'])}")
    print(f"   Всего использований: {a.get('total_anchors', '?')}")
    print(f"   Разных якорьков: {a.get('unique_anchors_used', '?')} (оптимум 2-3)")
    print(f"   Плотность на 1k слов: {a.get('anchors_per_1k_words', '?')} (норма 0.3-1.0)")
    print(f"   Вердикт: {a.get('verdict', '?')}")

    print("\nРитм абзацев:")
    r = results["rhythm"]
    print(f"   Макс. серия похожих абзацев: {r.get('max_consecutive_similar_paragraphs', '?')}")
    print(f"   Вердикт: {r.get('verdict', '?')}")

    print("\nРазнообразие начал предложений:")
    s = results.get("starters", {})
    if "error" in s or "skipped" in s:
        print(f"   {s.get('error') or s.get('skipped')}")
    else:
        print(f"   Уникальных стартов: {s.get('unique_starters', '?')} из {s.get('total_sentences', '?')}")
        print(f"   Коэффициент разнообразия: {s.get('variety_ratio', '?')} (цель: >0.65)")
        if s.get("top_repeated"):
            print("   Топ повторяющихся:")
            for starter, cnt in s["top_repeated"][:3]:
                print(f"     • «{starter}»: {cnt} раз")
        print(f"   Вердикт: {s.get('verdict', '?')}")

    print("\nПассивный залог:")
    passive = results.get("passive", {})
    if "skipped" in passive:
        print(f"   {passive.get('skipped')}")
    else:
        print(f"   Маркеры пассива: {passive.get('passive_markers_count', '?')} "
              f"({passive.get('ratio', '?')} на предложение)")
        print(f"   Вердикт: {passive.get('verdict', '?')}")

    print("\nРаспределение длин предложений:")
    rd = results.get("readability", {})
    if "skipped" in rd:
        print(f"   {rd.get('skipped')}")
    else:
        print(f"   Коротких (<8 слов): {rd.get('short_sentences_pct', '?')}%")
        print(f"   Средних (8-25 слов): {rd.get('medium_sentences_pct', '?')}%")
        print(f"   Длинных (>25 слов): {rd.get('long_sentences_pct', '?')}%")
        print(f"   Вердикт: {rd.get('verdict', '?')}")

    print("\nЛичные местоимения:")
    p = results["pronouns"]
    print(f"   {p.get('verdict', '?')}")
    if p.get("found_pronouns"):
        print("   " + ", ".join(f"«{word}»: {count}" for word, count in p["found_pronouns"]))

    print("\n" + "=" * 70)

    # Рекомендации
    print("\nРЕКОМЕНДАЦИИ:")
    if level == "LOW":
        print("   Поверхностные ИИ-маркеры не обнаружены. Это НЕОБХОДИМОЕ,")
        print("   но НЕ ДОСТАТОЧНОЕ условие — реальный детектор Антиплагиат.Вуз")
        print("   использует нейросетевой классификатор, который этот скрипт")
        print("   не симулирует. Низкий результат здесь ≠ гарантия прохождения.")
    elif level == "MODERATE":
        print("   Текст близок к норме, но есть что улучшить. Пройдись по флагам выше.")
    elif level == "HIGH":
        print("   ВНИМАНИЕ: текст может быть помечен как ИИ. Переработай:")
        print("   - Разбей монотонные блоки")
        print("   - Убери лишние клише")
        print("   - Добавь конкретные имена/цифры/технические детали")
    else:  # CRITICAL
        print("   ОПАСНО: текст сильно похож на ИИ-генерацию.")
        print("   Требуется серьёзная переработка перед сдачей.")
    print()


def _scope_line(meta: dict) -> str:
    if meta.get("structure") == "unrecognized":
        return ("[ВНИМАНИЕ: разделы ВКР не распознаны — анализируется весь текст документа "
                "без содержания, списка литературы, приложений, кода и подписей]")
    return ("[Анализируются: " + ", ".join(meta.get("sections_analyzed") or meta.get("sections_found") or [])
            + "; титул, аннотация, содержание, список литературы, приложения, код и подписи исключены]")


# ========== MAIN ==========

def _write_output(path: Path, payload: dict) -> Optional[str]:
    if is_inside_skill_dir(path):
        return f"{path} внутри каталога скилла: укажи файл в каталоге проекта, например <PROJECT_DIR>/exports/"
    try:
        common.atomic_write_text(Path(path), common.dump_json(payload))
    except OSError as error:
        return f"не удалось записать {path}: {error}"
    return None


def main(argv: Optional[Sequence[str]] = None) -> int:
    common.reconfigure_stdio()
    parser = argparse.ArgumentParser(
        description="Эвристический детектор ИИ-текста для ВКР",
        epilog="Коды выхода: 0 — анализ выполнен; 1 — INSUFFICIENT_DATA, раздел не найден или "
               "в --per-section есть ненайденные/недостаточные разделы; 2 — нет файла, повреждённый "
               "DOCX, нет python-docx, ошибка записи -o.",
    )
    parser.add_argument("docx_path", help="Путь к .docx файлу")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--section",
                      help="Анализировать только раздел: introduction | conclusion | chapterN (chapter1, chapter2, …)")
    mode.add_argument("--per-section", action="store_true",
                      help="Проанализировать введение, каждую главу и заключение отдельно с их порогами")
    parser.add_argument("--json", action="store_true", help="Только JSON в stdout")
    parser.add_argument("-o", "--output", type=Path, help="Записать JSON-отчёт в файл (UTF-8, каталоги создаются)")
    args = parser.parse_args(argv)

    path = Path(args.docx_path)
    meta: Dict[str, Any] = {"tool": "ai_detection_heuristic", "version": TOOL_VERSION, "docx": str(path)}

    def finish(payload: dict, code: int, text_printer=None) -> int:
        payload.setdefault("meta", meta)
        payload["exit_code"] = code
        if args.output:
            problem = _write_output(args.output, payload)
            if problem:
                print(f"ОШИБКА: {problem}", file=sys.stderr)
                return 2
        if args.json:
            sys.stdout.write(common.dump_json(payload))
        elif text_printer is not None:
            text_printer()
        return code

    def fail(code: int, error_code: str, message: str) -> int:
        payload = {"status": "error", "error": {"code": error_code, "message": message}}
        return finish(payload, code, lambda: print(f"ОШИБКА: {message}", file=sys.stderr))

    if args.section and not SECTION_NAME_RE.match(args.section.strip().casefold()):
        return fail(2, "SECTION_INVALID", f"неизвестный раздел «{args.section}»: introduction, conclusion, chapter1, chapter2, …")
    if not path.is_file():
        return fail(2, "FILE_NOT_FOUND", f"файл не найден: {path}")
    try:
        structure = read_structure(path)
    except DependencyError as error:
        return fail(2, "DEPENDENCY_MISSING", str(error))
    except Exception as error:  # повреждённый пакет, не DOCX
        return fail(2, "DOCX_INVALID", f"не удалось открыть DOCX {path}: {type(error).__name__}: {error}")

    global CLICHE_ALLOWLIST, PERSONAL_ANCHORS
    phrases, state = load_cliche_allowlist(path)
    CLICHE_ALLOWLIST = phrases
    PERSONAL_ANCHORS = load_personal_anchors(state)
    meta.update(structure.summary())
    meta["cliche_allowlist"] = allowlist_info(phrases, state)
    meta["personal_anchors"] = {"state_file": str(state) if state else None, "phrases": list(PERSONAL_ANCHORS)}

    if args.per_section:
        found = structure.section_names()
        missing = [name for name in structure.expected_sections() if name not in found]
        sections: "OrderedDict[str, dict]" = OrderedDict()
        insufficient: List[str] = []
        for name in found:
            text, paragraphs = extract_section(structure, name)
            words = count_words(text)
            if words < MIN_WORDS_SECTION or not paragraphs:
                sections[name] = insufficient_result(
                    name, f"Для анализа раздела нужно минимум {MIN_WORDS_SECTION} слов; найдено {words}")
                insufficient.append(name)
            else:
                sections[name] = analyze_text(text, paragraphs, section_type=name)
            sections[name]["words"] = words
            sections[name]["paragraphs"] = len(paragraphs)
        if not found:
            status = "insufficient_data"
        elif missing or insufficient:
            status = "incomplete"
        else:
            status = "ok"
        payload = {
            "status": status,
            "mode": "per_section",
            "sections": sections,
            "missing_sections": missing,
            "insufficient_sections": insufficient,
        }
        code = 0 if status == "ok" else 1

        def show() -> None:
            print(f"\n{'=' * 70}")
            print("  ПОСЕКЦИОННЫЙ АНАЛИЗ (разные пороги по типу главы)")
            print(f"{'=' * 70}\n")
            if not found:
                print("Разделы ВКР (введение, главы, заключение) не найдены: проверь заголовки.")
            for sname, res in sections.items():
                risk = res["risk"]
                score = "—" if risk["risk_score"] is None else f"{risk['risk_score']}/100"
                print(f"{sname.upper()}: {res['words']} слов")
                print(f"   Score: {score} — Уровень: {risk['risk_level']}")
                if risk["flags"]:
                    print(f"   Флаги: {', '.join(risk['flags'][:5])}")
                print()
            if missing:
                print("НЕ НАЙДЕНЫ РАЗДЕЛЫ: " + ", ".join(missing))
            if insufficient:
                print(f"НЕДОСТАТОЧНО ДАННЫХ (< {MIN_WORDS_SECTION} слов): " + ", ".join(insufficient))
            if missing or insufficient or not found:
                print("Код выхода 1: посекционный анализ неполный.\n")
            print("СПРАВКА: пороги различаются: теоретическая глава (chapter1) толерантнее к шаблонности,")
            print("    проектная (chapter3) и введение/заключение — строже. Это нормально:")
            print("    теория по природе академичная, а проектное описание должно быть живым.")

        return finish(payload, code, show)

    if args.section:
        section = args.section.strip().casefold()
        meta["mode"] = "section"
        meta["section"] = section
        if section not in structure.section_names():
            found = ", ".join(structure.section_names()) or "нет"
            message = f"раздел {section} не найден; найдены: {found}"
            payload = insufficient_result(section, message)
            payload["status"] = "section_not_found"
            return finish(payload, 1, lambda: print(f"ОШИБКА: {message}", file=sys.stderr))
        text, paragraphs = extract_section(structure, section)
        section_type = section
        minimum = MIN_WORDS_SECTION
        meta["sections_analyzed"] = [section]
    else:
        meta["mode"] = "full"
        paragraphs = split_paragraphs_from_doc(structure)
        text = "\n\n".join(paragraphs)
        section_type = "full"
        minimum = MIN_WORDS_FULL
        meta["sections_analyzed"] = structure.section_names()

    word_count = count_words(text)
    meta["words"] = word_count
    meta["paragraphs"] = len(paragraphs)
    if word_count < minimum or not paragraphs:
        message = f"Для анализа нужно минимум {minimum} слов; найдено {word_count}"
        payload = insufficient_result(section_type, message)
        payload["status"] = "insufficient_data"
        return finish(payload, 1, lambda: print(f"INSUFFICIENT_DATA: {message}. Код выхода 1.", file=sys.stderr))

    results = analyze_text(text, paragraphs, section_type=section_type)
    results["status"] = "ok"

    def show_full() -> None:
        if args.section:
            print(f"\n[Анализ раздела: {args.section}, применены пороги для этого типа]")
        print(_scope_line(meta))
        if phrases:
            print(f"[Активен персональный cliche_allowlist: {len(phrases)} фраз из {state}]")
        print_report(results)

    return finish(results, 0, show_full)


if __name__ == "__main__":
    raise SystemExit(main())
