#!/usr/bin/env python3
"""
verify_sources.py — реестр источников ВКР: извлечение из DOCX, сводка статусов,
поисковые запросы для проверки (vkr-mpgu 6.33).

Скрипт НЕ ищет в интернете: запросы выполняет ИИ-агент через доступный
web_search/web_fetch, а результат записывает в <PROJECT_DIR>/sources.json
(status confirmed и объект verification, см. references/source-verification.md).

    python verify_sources.py --extract <DOCX> [-o <PROJECT_DIR>/audit/sources-extracted.json] [--force] [--json]
    python verify_sources.py --report <PROJECT_DIR>/sources.json [--json] [-o REPORT.json]
    python verify_sources.py --queries <PROJECT_DIR>/sources.json [--json] [-o QUERIES.json]
    python verify_sources.py --mark <PROJECT_DIR>/sources.json --id ID --status confirmed --method catalog --url URL [--note TEXT]

--extract ищет раздел списка литературы по заголовку («Список использованной
литературы», «Список использованных источников», «Список литературы»,
«Библиографический список», «Литература», «… и источников»; регистр и двоеточие
не важны), не принимая за него строки оглавления. Раздел заканчивается на
следующем заголовке, «Приложение N», «Приложения», последнем листе или клятве.
Запись — абзац с номером («1.», «1)», «[1]»), с автонумерацией Word (numPr) или со
стилем VKR Bibliography. Ненумерованный абзац приклеивается к предыдущей записи,
если это очевидное продолжение («– URL: …», строчная буква, висящий знак в конце
записи); короткий полужирный или прописной абзац без года — подзаголовок группы
(SUBHEADING_SKIPPED); иначе он попадает в warnings. Результат — записи
{id, raw_text, status: "pending", hints}; существующий файл перезаписывается
только с --force, а реестр проекта sources.json со структурированными записями —
никогда.

Тот же разбор записей (parse_bibliography_blocks) использует import_docx.py.

--mark записывает результат проверки в реестр: status (verified пишется как
confirmed) и verification {checked_at — текущее время ISO 8601, method, url,
note, fields_sha256 — хеш библиографических полей записи на момент проверки}.
Для confirmed обязательны --method и --url или --note; для suspicious и
rejected — --note с причиной; --status pending снимает verification. Если
реквизиты confirmed-записи потом правили (хеш не совпадает), --report относит её
к confirmed_stale, --queries снова строит для неё запросы, build_vkr.py
предупреждает SOURCE_VERIFICATION_STALE.

--queries: точная фраза заглавия включает подзаголовок (subtitle), отдельно —
«фамилия + заглавие» без кавычек.

Коды выхода: 0 — успех (--report: все источники confirmed с полным verification);
1 — --extract не нашёл ни одной записи; --report: есть непроверенные, отклонённые,
подозрительные записи, confirmed без verification или с устаревшей проверкой,
ошибки реестра или реестр пуст;
--queries: реестр пуст; 2 — нет файла, некорректный JSON, повреждённый DOCX, нет
python-docx, выходной файл существует или защищён, ошибка записи; --mark: неверные
аргументы, запись не найдена или id неоднозначен.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

sys.dont_write_bytecode = True
SCRIPTS_DIR = Path(__file__).resolve().parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import vkr_common as common  # noqa: E402
import ai_detection_heuristic as docx_structure  # noqa: E402  (общий разбор структуры DOCX)

TOOL_VERSION = "6.33"
DEFAULT_EXTRACT_NAME = "sources-extracted.json"
PROJECT_SEARCH_LEVELS = 3
SOURCE_ID_RE = re.compile(r"^[A-Za-z0-9_.:-]+$")
REGISTRY_SERVICE_KEYS = {"id", "raw_text", "status", "hints", "notes", "found_url"}

normalize_space = docx_structure.normalize_space
DependencyError = docx_structure.DependencyError


# ========== ИЗВЛЕЧЕНИЕ СПИСКА ЛИТЕРАТУРЫ ИЗ DOCX ==========

ENTRY_NUMBER_RE = re.compile(r"^\s*(?:\[(?P<bracket>\d{1,3})\]|(?P<num>\d{1,3})\s*[.)])\s*(?=\S)")
CONTINUATION_START_RE = re.compile(
    r"^(?:[–—\-/:;,.)\]]|//|\(\s*дата|url\b|doi\b|https?://|www\.|режим\s+доступа|isbn\b|"
    r"с\.\s*\d|p\.\s*\d|pp\.\s*\d|vol\.|no\.|№|т\.\s*\d|вып\.)",
    re.IGNORECASE,
)
TERMINAL_END_RE = re.compile(r"[.!?…]\s*$")
HANGING_END_RE = re.compile(r"(?:[–—\-/:,;(«\"]|//)\s*$")


def looks_like_continuation(previous_raw: str, text: str) -> bool:
    """Очевидное продолжение разорванной записи списка литературы."""
    stripped = text.lstrip()
    if not stripped or not previous_raw:
        return False
    if CONTINUATION_START_RE.match(stripped):
        return True
    first = stripped[0]
    if first.isalpha() and first.islower():
        return True
    if first.isdigit() and not TERMINAL_END_RE.search(previous_raw):
        return True
    return bool(HANGING_END_RE.search(previous_raw))


GROUP_HEADING_MAX_CHARS = 100
ENTRY_MARKS_RE = re.compile(r"\s[/–—]\s|//|https?://|\[\s*(?:текст|электронный\s+ресурс)", re.IGNORECASE)


def looks_like_group_heading(block: Dict[str, Any], text: str) -> bool:
    """Подзаголовок группы записей без стиля заголовка («НОРМАТИВНЫЕ АКТЫ», полужирное «Электронные ресурсы»)."""
    if not (block.get("bold") or block.get("upper")) or len(text) > GROUP_HEADING_MAX_CHARS:
        return False
    return not (re.search(r"(?<!\d)(?:1[89]|20)\d{2}(?!\d)", text) or ENTRY_MARKS_RE.search(text))


def _short(text: str, limit: int = 120) -> str:
    text = normalize_space(text)
    return text if len(text) <= limit else text[: limit - 1] + "…"


def _entries_from_blocks(blocks: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
    entries: List[Dict[str, Any]] = []
    warnings: List[Dict[str, Any]] = []
    previous: Optional[Dict[str, Any]] = None
    for block in blocks:
        if block.get("type") == "tbl":
            if block.get("text"):
                warnings.append({"code": "TABLE_SKIPPED", "text": _short(block["text"]),
                                 "message": "таблица внутри списка литературы пропущена"})
            previous = None
            continue
        raw = block.get("text") or ""
        text = normalize_space(raw)
        if not text or block.get("toc"):
            continue
        if block.get("heading_level") is not None:
            warnings.append({"code": "SUBHEADING_SKIPPED", "text": _short(text),
                             "message": "подзаголовок внутри списка литературы пропущен"})
            previous = None
            continue
        numbered = ENTRY_NUMBER_RE.match(text)
        auto = bool(block.get("in_list") or block.get("bibliography_style"))
        if numbered:
            number = int(numbered.group("bracket") or numbered.group("num"))
            body = text[numbered.end():].strip()
            previous = {"number": number, "raw_text": body, "source": "text_number"}
            entries.append(previous)
            continue
        if auto and not (previous is not None and (CONTINUATION_START_RE.match(text) or text[0].islower())):
            previous = {"number": None, "raw_text": text, "source": "numbering"}
            entries.append(previous)
            continue
        if previous is not None and looks_like_continuation(previous["raw_text"], text):
            previous["raw_text"] = normalize_space(previous["raw_text"] + " " + text)
            warnings.append({"code": "CONTINUATION_JOINED", "text": _short(text),
                             "message": "абзац без номера приклеен к предыдущей записи как продолжение"})
            previous.setdefault("joined", 0)
            previous["joined"] += 1
            continue
        if looks_like_group_heading(block, text):
            warnings.append({"code": "SUBHEADING_SKIPPED", "text": _short(text),
                             "message": "подзаголовок внутри списка литературы пропущен"})
            previous = None
            continue
        warnings.append({"code": "UNNUMBERED_PARAGRAPH", "text": _short(text),
                         "message": "абзац без номера и автонумерации не считается записью"})
    return {"entries": entries, "warnings": warnings}


def _assign_ids(entries: List[Dict[str, Any]], warnings: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    records: List[Dict[str, Any]] = []
    used: set = set()
    for entry in entries:
        number = entry["number"]
        if number is not None and number > 0 and number not in used:
            source_id = number
        else:
            source_id = (max(used) if used else 0) + 1
            if number is not None:
                warnings.append({"code": "NUMBER_DUPLICATE", "text": _short(entry["raw_text"]),
                                 "message": f"номер {number} уже встречался: запись получила id {source_id}"})
        used.add(source_id)
        hints = parse_source_hints(entry["raw_text"])
        hints.pop("raw", None)
        records.append({"id": source_id, "raw_text": entry["raw_text"], "status": "pending", "hints": hints})
    return records


def parse_bibliography_blocks(blocks: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
    """Общий разбор записей списка литературы: ``--extract`` и ``import_docx.py``.

    ``blocks`` — абзацы и таблицы раздела по порядку: ``{"type": "p"|"tbl", "text",
    "toc", "heading_level", "in_list", "bibliography_style"}``. Запись — абзац с
    номером («1.», «1)», «[1]»), с автонумерацией Word или стилем VKR Bibliography;
    очевидное продолжение разорванной записи приклеивается (CONTINUATION_JOINED).
    Возвращает ``{"entries": [{number, raw_text, source, joined?}], "records":
    [{id, raw_text, status, hints}], "warnings": [{code, text, message}]}``;
    повторный номер получает следующий свободный id с NUMBER_DUPLICATE, абзац без
    номера, не похожий на продолжение, записью не становится (UNNUMBERED_PARAGRAPH).
    """
    parsed = _entries_from_blocks(list(blocks))
    warnings = list(parsed["warnings"])
    records = _assign_ids(parsed["entries"], warnings)
    return {"entries": parsed["entries"], "records": records, "warnings": warnings}


def extract_bibliography(source) -> Dict[str, Any]:
    """Записи списка литературы DOCX: {found, heading, records, warnings}."""
    structure = docx_structure.read_structure(source)
    regions = structure.bibliography_regions()
    if not regions:
        return {"found": False, "heading": "", "records": [], "warnings": []}
    conclusion_positions = [r["heading_index"] for r in structure.regions if r["kind"] == "conclusion"]
    after = conclusion_positions[-1] if conclusion_positions else None
    preferred = [r for r in regions if after is None or (r["heading_index"] or 0) > after]
    ordered = preferred + [r for r in regions if r not in preferred]
    chosen = None
    parsed = None
    for region in ordered:
        candidate = parse_bibliography_blocks(region["blocks"])
        if candidate["entries"]:
            chosen, parsed = region, candidate
            break
    if chosen is None:
        chosen = ordered[0]
        parsed = parse_bibliography_blocks(chosen["blocks"])
    warnings = list(parsed["warnings"])
    records = parsed["records"]
    if len(regions) > 1:
        warnings.append({"code": "BIBLIOGRAPHY_HEADINGS_MULTIPLE", "text": chosen["heading"],
                         "message": f"найдено заголовков списка литературы: {len(regions)}; использован «{chosen['heading']}»"})
    return {"found": True, "heading": chosen["heading"], "records": records, "warnings": warnings}


def extract_sources_from_docx(docx_path) -> list:
    """Совместимость с 6.32: только записи."""
    return extract_bibliography(docx_path)["records"]


# ========== ПОДСКАЗКИ ИЗ СЫРОГО ТЕКСТА ==========

AUTHOR_RE = re.compile(
    r"^(?P<last>[А-ЯЁA-Z][А-ЯЁа-яёA-Za-z'’\-]*[а-яёa-z'’])(?:\s*,\s*|\s+)(?P<initials>[А-ЯЁA-Z]\.\s*(?:[А-ЯЁA-Z]\.)?)"
)
TITLE_STOP_RE = re.compile(
    r"\s*\[[^\]]{0,60}\]"
    r"|\s+//\s*|\s+/\s+|\s*\.\s*[–—-]\s+"
    r"|\s*:\s*(?=(?:электрон|дис\.|дисс|автореф|учеб|монограф|сб\.|сборник|материал|федер|приказ|закон|"
    r"постановлен|указ\b|распоряжен|гост\b|стандарт|справочн|словар|пособ|практикум|хрестомат|"
    r"курс\s+лекций|методическ|научн\w*\s+(?:журнал|электрон)))",
    re.IGNORECASE,
)
ACCESS_DATE_RE = re.compile(r"\(\s*дата\s+обращения[^)]*\)", re.IGNORECASE)
URL_RE = re.compile(r"https?://[^\s<>«»\"]+", re.IGNORECASE)
DOI_RE = re.compile(r"(?:\bdoi\s*:?\s*|https?://(?:dx\.)?doi\.org/)(10\.\d{4,9}/[^\s,;]+)", re.IGNORECASE)
ISBN_RE = re.compile(r"\bisbn[\s:]*((?:97[89][\s-]?)?(?:\d[\s-]?){9}[\dXx])", re.IGNORECASE)
YEAR_RE = re.compile(r"(?<!\d)(1[89]\d{2}|20\d{2})(?!\d)")
DESIGNATION_RE = re.compile(r"\b(ГОСТ(?:\s+Р)?(?:\s+ИСО(?:/МЭК)?)?\s+[\d.]+(?:[-–]\d{2,4})?)", re.IGNORECASE)
DOC_NUMBER_RE = re.compile(r"№\s*([\w\-–/]+)")
NORMATIVE_RE = re.compile(
    r"(?:\bфедеральн\w*\s+закон\b|\bфедер\.\s*закон\b|\bприказ\b|\bпостановлени\w*\b|\bкодекс\b|"
    r"\bгост\b|\bфгос\b|\bуказ\b|\bраспоряжени\w*\b|\bконституци\w*\b)",
    re.IGNORECASE,
)


def parse_source_hints(raw_text: str) -> dict:
    """Эвристика по строке списка литературы: автор, заглавие, год, DOI, ISBN, URL, тип."""
    raw_text = normalize_space(raw_text)
    hints: Dict[str, Any] = {"raw": raw_text}
    text = re.sub(r"^\s*(?:\[\d{1,3}\]|\d{1,3}\s*[.)])\s*", "", raw_text)

    author = AUTHOR_RE.match(text)
    body = text
    if author:
        hints["author_last"] = author.group("last")
        hints["author_initials"] = re.sub(r"\s+", "", author.group("initials"))
        body = text[author.end():].strip()

    parts = TITLE_STOP_RE.split(body, maxsplit=1)
    title = parts[0].strip(" .,:;«»\"")
    remainder = body[len(parts[0]):] if parts else ""
    if title:
        hints["title"] = title

    cleaned = URL_RE.sub(" ", ACCESS_DATE_RE.sub(" ", remainder))
    year = YEAR_RE.search(cleaned) or YEAR_RE.search(URL_RE.sub(" ", ACCESS_DATE_RE.sub(" ", text)))
    if year:
        hints["year"] = int(year.group(1))

    doi = DOI_RE.search(raw_text)
    if doi:
        hints["doi"] = doi.group(1).rstrip(".")
    url = next((m.group(0).rstrip(".,;)") for m in URL_RE.finditer(raw_text) if "doi.org/" not in m.group(0).casefold()), None)
    if url:
        hints["url"] = url
    isbn = ISBN_RE.search(raw_text)
    if isbn:
        hints["isbn"] = re.sub(r"\s+", "", isbn.group(1))
    designation = DESIGNATION_RE.search(raw_text)
    if designation:
        hints["designation"] = normalize_space(designation.group(1))
        if title and title.casefold().startswith(hints["designation"].casefold()):
            rest = title[len(hints["designation"]):].lstrip(" .:–—-")
            if rest:
                hints["title"] = title = rest
    number = DOC_NUMBER_RE.search(raw_text)
    if number:
        hints["doc_number"] = number.group(1).rstrip(".,")

    lowered = raw_text.casefold()
    if NORMATIVE_RE.search(raw_text):
        hints["type"] = "normative"
    elif "автореф" in lowered:
        hints["type"] = "autoreferat"
    elif re.search(r"\bдисс?\.|\bдиссертаци", lowered):
        hints["type"] = "dissertation"
    elif " // " in raw_text:
        hints["type"] = "article"
    elif "[электронный ресурс]" in lowered or "[electronic resource]" in lowered or "url:" in lowered:
        hints["type"] = "electronic"
    else:
        hints["type"] = "book"
    probe = title or text
    if re.search(r"[A-Za-z]", probe) and not re.search(r"[А-Яа-яЁё]", probe):
        hints["language"] = "en"
    return hints


def _clean_title(value: Any) -> str:
    title = normalize_space(value)
    title = re.sub(r"\s*\[[^\]]{0,60}\]", "", title)
    return title.strip(" .,:;«»\"")


def source_hints(source: dict) -> dict:
    """Подсказки: эвристика raw_text, поверх — структурированные поля реестра."""
    raw = source.get("raw_text")
    if raw:
        hints = parse_source_hints(str(raw))
    else:
        hints = dict(source.get("hints") or {})
    if source.get("subtitle"):
        hints["subtitle"] = _clean_title(source["subtitle"])
    if source.get("title"):
        hints["title"] = _clean_title(source["title"])
        probe = hints["title"]
        if re.search(r"[A-Za-z]", probe) and not re.search(r"[А-Яа-яЁё]", probe):
            hints["language"] = "en"
        else:
            hints.pop("language", None)
    for key in ("year", "doi", "url", "isbn", "type"):
        if source.get(key):
            hints[key] = source[key]
    if source.get("designation"):
        hints["designation"] = source["designation"]
    if source.get("number") and str(source.get("type") or "") == "normative":
        hints["doc_number"] = source["number"]
    authors = source.get("authors") or []
    if isinstance(authors, (str, dict)):
        authors = [authors]
    if authors:
        first = authors[0]
        if isinstance(first, dict):
            hints["author_last"] = str(first.get("last") or first.get("name") or "").strip()
            hints["author_initials"] = str(first.get("initials") or "").strip()
        else:
            value = normalize_space(first)
            match = re.match(r"^([А-ЯЁA-Z][А-ЯЁа-яёA-Za-z'’\-]+)\s*,?\s*((?:[А-ЯЁA-Z]\.\s*){1,2})$", value)
            reverse = re.match(r"^((?:[А-ЯЁA-Z]\.\s*){1,2})\s*([А-ЯЁA-Z][А-ЯЁа-яёA-Za-z'’\-]+)$", value)
            if match:
                hints["author_last"], hints["author_initials"] = match.group(1), match.group(2).replace(" ", "")
            elif reverse:
                hints["author_last"], hints["author_initials"] = reverse.group(2), reverse.group(1).replace(" ", "")
            else:
                hints["author_last"], hints["author_initials"] = value, ""
    return hints


# ========== ПОИСКОВЫЕ ЗАПРОСЫ ==========

def build_search_queries(hints: dict) -> list:
    """Запросы для проверки источника: точное заглавие без служебных пометок (вместе с
    подзаголовком, если он есть: общее заглавие «Информатика» само по себе бесполезно),
    автор и заглавие, отдельно «фамилия + заглавие» без кавычек, год, DOI, ISBN и URL."""
    queries: List[Dict[str, str]] = []

    def add(engine: str, query: str, purpose: str) -> None:
        query = normalize_space(query)
        if query and all(item["query"] != query for item in queries):
            queries.append({"engine": engine, "query": query, "purpose": purpose})

    title = _clean_title(hints.get("title") or "")
    subtitle = _clean_title(hints.get("subtitle") or "")
    if subtitle and subtitle.casefold() in title.casefold():
        subtitle = ""
    phrase = f"{title}: {subtitle}" if title and subtitle else title
    author_last = normalize_space(hints.get("author_last") or "")
    author = normalize_space(f"{author_last} {hints.get('author_initials', '')}")
    year = hints.get("year")
    kind = hints.get("type") or "book"
    latin = hints.get("language") == "en"

    if title:
        add("web_search", f'"{phrase}"',
            "точное заглавие" + (" с подзаголовком" if subtitle else "") + ": существует ли работа")
        if author:
            add("web_search", f'{author} "{phrase}"',
                "автор и заглавие" + (" (Google Scholar, сайт издателя)" if latin else ""))
            add("web_search", normalize_space(f"{author_last} {title} {subtitle}"),
                "фамилия автора и заглавие без кавычек: найдёт издание при другой пунктуации и порядке слов")
        if year:
            add("web_search", f'"{phrase}" {year}', "год и издание")
        if not latin and kind in ("article", "collection_article", "conference"):
            add("web_search", f'"{phrase}" site:cyberleninka.ru', "КиберЛенинка")
        if not latin and kind in ("article", "collection_article", "conference", "book", "dissertation", "autoreferat"):
            add("web_search", f'"{phrase}" site:elibrary.ru', "eLibrary (РИНЦ)")
        if kind in ("dissertation", "autoreferat"):
            add("web_search", f'"{phrase}" site:dissercat.com', "каталог диссертаций")
    else:
        raw = _short(hints.get("raw") or "", 100)
        if raw:
            add("web_search", raw, "поиск по описанию: заглавие не выделено")
    if author and not title:
        add("web_search", author + (f" {year}" if year else ""), "автор")
    if kind in ("normative", "standard"):
        if hints.get("designation"):
            add("web_search", f'"{hints["designation"]}"', "обозначение стандарта: действующая редакция")
        if hints.get("doc_number"):
            add("web_search", f'"№ {hints["doc_number"]}"' + (f' "{title}"' if title else ""),
                "официальный текст: pravo.gov.ru, КонсультантПлюс, Гарант")
    if hints.get("doi"):
        add("web_fetch", f"https://doi.org/{hints['doi']}", "DOI разрешается и ведёт на эту работу")
        add("web_search", f'"{hints["doi"]}"', "DOI в каталогах")
    if hints.get("isbn"):
        add("web_search", f"ISBN {hints['isbn']}", "ISBN: издание, год, издательство")
    if hints.get("url"):
        add("web_fetch", str(hints["url"]), "URL открывается и содержит работу")
    return queries


# ========== РЕЕСТР: ОТЧЁТ ==========

class RegistryError(Exception):
    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code
        self.message = message


def load_registry(path: Path) -> List[Any]:
    try:
        data = common.read_json(Path(path))
    except common.ProjectDataError as error:
        raise RegistryError(error.code, str(error)) from error
    if not isinstance(data, list):
        raise RegistryError("REGISTRY_SHAPE", f"{path}: ожидается JSON-массив источников (SPEC 3)")
    return data


# Не входят в хеш реквизитов: служебные поля и сам результат проверки.
FIELDS_HASH_EXCLUDED_KEYS = frozenset({"id", "status", "verification", "hints", "notes", "found_url", "recommendation"})


def bibliographic_fields_sha256(source: dict) -> str:
    """SHA-256 канонического JSON библиографических полей записи (без id, status, verification и служебных полей).

    ``--mark`` сохраняет его в ``verification.fields_sha256``; другой хеш у confirmed-записи
    значит, что реквизиты правили после проверки.
    """
    fields = {key: value for key, value in source.items() if key not in FIELDS_HASH_EXCLUDED_KEYS}
    return common.sha256_bytes(common.canonical_json_bytes(fields))


def verification_stale(source: Any) -> bool:
    """Реквизиты записи изменены после ``--mark`` (verification.fields_sha256 не совпадает)."""
    if not isinstance(source, dict):
        return False
    verification = source.get("verification")
    if not isinstance(verification, dict):
        return False
    recorded = str(verification.get("fields_sha256") or "").strip()
    return bool(recorded) and recorded != bibliographic_fields_sha256(source)


def verification_missing(source: dict) -> List[str]:
    """Чего не хватает в verification у confirmed-источника (правило doctor/build)."""
    verification = source.get("verification")
    if not isinstance(verification, dict):
        return ["verification"]
    missing = []
    if verification.get("method") not in common.VERIFICATION_METHODS:
        missing.append("method")
    if common.parse_iso8601(verification.get("checked_at")) is None:
        missing.append("checked_at")
    if not (str(verification.get("url") or "").strip() or str(verification.get("note") or "").strip()):
        missing.append("url|note")
    return missing


def describe_source(source: dict) -> str:
    raw = normalize_space(source.get("raw_text") or "")
    if raw:
        return _short(raw)
    parts = []
    authors = source.get("authors") or []
    if isinstance(authors, (str, dict)):
        authors = [authors]
    if authors:
        first = authors[0]
        if isinstance(first, dict):
            parts.append(normalize_space(f"{first.get('last') or first.get('name') or ''} {first.get('initials') or ''}"))
        else:
            parts.append(normalize_space(first))
    if source.get("title"):
        parts.append(normalize_space(source["title"]))
    if source.get("year"):
        parts.append(f"({source['year']})")
    return _short(" ".join(part for part in parts if part)) or "(нет title и raw_text)"


def source_label(source: Any, index: int) -> str:
    if isinstance(source, dict) and source.get("id") is not None and str(source.get("id")).strip():
        return f"[{source['id']}]"
    return f"[#{index}]"


def _note_and_url(source: dict) -> Dict[str, str]:
    verification = source.get("verification") if isinstance(source.get("verification"), dict) else {}
    note = str(verification.get("note") or source.get("notes") or "").strip()
    url = str(verification.get("url") or source.get("found_url") or "").strip()
    result = {"note": note, "url": url}
    if source.get("recommendation"):
        result["recommendation"] = str(source["recommendation"]).strip()
    return result


def build_report(sources: List[Any], registry: str) -> Dict[str, Any]:
    groups: Dict[str, List[Dict[str, Any]]] = {"confirmed": [], "suspicious": [], "pending": [], "rejected": [], "unknown": []}
    without_verification: List[Dict[str, Any]] = []
    stale: List[Dict[str, Any]] = []
    problems: List[Dict[str, Any]] = []
    seen_ids: Dict[str, int] = {}
    for index, source in enumerate(sources, start=1):
        if not isinstance(source, dict):
            problems.append({"code": "SOURCE_SHAPE", "index": index, "message": f"#{index}: запись не JSON-объект"})
            continue
        source_id = source.get("id")
        label = source_label(source, index)
        if source_id is None:
            problems.append({"code": "SOURCE_ID_MISSING", "index": index, "message": f"#{index}: нет id"})
        elif isinstance(source_id, bool) or not (isinstance(source_id, int) or (isinstance(source_id, str) and SOURCE_ID_RE.match(source_id))):
            problems.append({"code": "SOURCE_ID_INVALID", "index": index,
                             "message": f"#{index}: id {source_id!r} — нужна строка [A-Za-z0-9_.:-]+ или целое"})
        else:
            key = str(source_id).casefold()
            if key in seen_ids:
                problems.append({"code": "SOURCE_ID_DUPLICATE", "index": index,
                                 "message": f"#{index}: id {source_id} уже есть в позиции #{seen_ids[key]}"})
            else:
                seen_ids[key] = index
        status = common.normalize_source_status(source.get("status"))
        entry = {"id": source_id, "index": index, "label": label, "status": source.get("status"),
                 "description": describe_source(source)}
        entry.update(_note_and_url(source))
        groups[status if status in common.SOURCE_STATUSES else "unknown"].append(entry)
        if status == "confirmed":
            missing = verification_missing(source)
            if missing:
                without_verification.append(dict(entry, missing=missing))
            elif verification_stale(source):
                stale.append(entry)
    failing = bool(groups["suspicious"] or groups["pending"] or groups["rejected"] or groups["unknown"]
                   or without_verification or stale or problems or not sources)
    return {
        "tool": "verify_sources",
        "version": TOOL_VERSION,
        "mode": "report",
        "registry": registry,
        "status": "fail" if failing else "ok",
        "exit_code": 1 if failing else 0,
        "total": len(sources),
        "counts": {
            "confirmed": len(groups["confirmed"]),
            "confirmed_without_verification": len(without_verification),
            "confirmed_stale": len(stale),
            "suspicious": len(groups["suspicious"]),
            "pending": len(groups["pending"]),
            "rejected": len(groups["rejected"]),
            "unknown_status": len(groups["unknown"]),
        },
        "suspicious": groups["suspicious"],
        "pending": groups["pending"],
        "rejected": groups["rejected"],
        "unknown_status": groups["unknown"],
        "confirmed_without_verification": without_verification,
        "confirmed_stale": stale,
        "problems": problems,
    }


def _print_entries(title: str, entries: List[Dict[str, Any]], show_missing: bool = False) -> None:
    if not entries:
        return
    print(f"\n{title}:")
    for entry in entries:
        status = f" (status: {entry['status']})" if entry.get("status") not in (None, "") else " (status не указан)"
        print(f"  {entry['label']} {entry['description']}{status}")
        if show_missing:
            if entry["missing"] == ["verification"]:
                print("       Нет объекта verification {method, checked_at, url|note}")
            else:
                print(f"       Не хватает в verification: {', '.join(entry['missing'])}")
        if entry.get("note"):
            print(f"       Заметка: {entry['note']}")
        if entry.get("url"):
            print(f"       URL: {entry['url']}")
        if entry.get("recommendation"):
            print(f"       Рекомендация: {entry['recommendation']}")


def print_report(report: Dict[str, Any]) -> None:
    counts = report["counts"]
    print("\n" + "=" * 70)
    print("  ВЕРИФИКАЦИЯ ИСТОЧНИКОВ")
    print("=" * 70)
    print(f"Реестр: {report['registry']}")
    print(f"\nВсего записей: {report['total']}")
    print(f"  confirmed (подтверждено): {counts['confirmed']}"
          + (f", из них без полного verification: {counts['confirmed_without_verification']}"
             if counts["confirmed_without_verification"] else "")
          + (f", с устаревшей проверкой (реквизиты изменены после --mark): {counts['confirmed_stale']}"
             if counts["confirmed_stale"] else ""))
    print(f"  suspicious (сомнительные): {counts['suspicious']}")
    print(f"  pending (ожидают проверки): {counts['pending']}")
    print(f"  rejected (отклонены): {counts['rejected']}")
    print(f"  неизвестный статус: {counts['unknown_status']}")
    if report["total"] == 0:
        print("\nРеестр пуст: добавь источники в sources.json (схема — references/gost-citations.md).")
    _print_entries("ПОДОЗРИТЕЛЬНЫЕ (suspicious) — перепроверить или заменить", report["suspicious"])
    _print_entries("ОЖИДАЮТ ПРОВЕРКИ (pending)", report["pending"])
    _print_entries("ОТКЛОНЕНЫ (rejected) — убрать из текста и списка", report["rejected"])
    _print_entries("НЕИЗВЕСТНЫЙ СТАТУС — исправить на pending | confirmed | suspicious | rejected",
                   report["unknown_status"])
    _print_entries("CONFIRMED БЕЗ ПОЛНОГО verification {method, checked_at, url|note}",
                   report["confirmed_without_verification"], show_missing=True)
    _print_entries("CONFIRMED С УСТАРЕВШЕЙ ПРОВЕРКОЙ — реквизиты изменены после --mark: перепроверь и отметь заново",
                   report["confirmed_stale"])
    if report["problems"]:
        print("\nОШИБКИ РЕЕСТРА:")
        for problem in report["problems"]:
            print(f"  {problem['message']}")
    print()
    if report["status"] == "ok":
        print("Итог: все источники confirmed с полным verification (код 0).")
    else:
        print("Итог: проверка не пройдена (код 1).")
    print()


# ========== CLI ==========

def default_extract_path(docx_path: Path) -> Path:
    resolved = Path(docx_path).resolve()
    for parent in [resolved.parent] + list(resolved.parents)[:PROJECT_SEARCH_LEVELS]:
        if (parent / "vkr-project.json").is_file():
            return parent / "audit" / DEFAULT_EXTRACT_NAME
    return resolved.parent / DEFAULT_EXTRACT_NAME


def protected_registry_reason(output: Path) -> Optional[str]:
    """Причина отказа писать в существующий реестр проекта sources.json."""
    output = Path(output)
    if output.name.casefold() != "sources.json" or not output.exists():
        return None
    suggestion = "запиши извлечение в другой файл, например -o <PROJECT_DIR>/audit/sources-extracted.json"
    try:
        data = common.read_json(output)
    except common.ProjectDataError as error:
        return f"{output} уже существует и не читается ({error.detail}); реестр проекта не перезаписывается — {suggestion}"
    if not isinstance(data, list):
        return f"{output} уже существует и не является извлечённым списком; реестр проекта не перезаписывается — {suggestion}"
    structured = [
        entry for entry in data
        if isinstance(entry, dict) and (
            set(entry) - REGISTRY_SERVICE_KEYS
            or common.normalize_source_status(entry.get("status")) not in ("", "pending")
        )
    ]
    if structured:
        return (f"{output} — реестр проекта со структурированными или проверенными записями ({len(structured)}); "
                f"--extract не перезаписывает его даже с --force — {suggestion}")
    return None


def _emit(args, payload: Dict[str, Any], text_printer=None) -> int:
    code = int(payload.get("exit_code", 0))
    if args.output and payload.get("mode") != "extract":
        if docx_structure.is_inside_skill_dir(args.output):
            print(f"ОШИБКА: {args.output} внутри каталога скилла: укажи файл в каталоге проекта", file=sys.stderr)
            return 2
        try:
            common.atomic_write_text(Path(args.output), common.dump_json(payload))
        except OSError as error:
            print(f"ОШИБКА: не удалось записать {args.output}: {error}", file=sys.stderr)
            return 2
    if args.json:
        sys.stdout.write(common.dump_json(payload))
    elif text_printer is not None:
        text_printer()
    return code


def _error(args, mode: str, code: int, error_code: str, message: str, **extra: Any) -> int:
    payload: Dict[str, Any] = {"tool": "verify_sources", "version": TOOL_VERSION, "mode": mode,
                               "status": "error", "exit_code": code,
                               "error": {"code": error_code, "message": message}}
    payload.update(extra)
    if args.json:
        sys.stdout.write(common.dump_json(payload))
    else:
        print(f"ОШИБКА: {message}", file=sys.stderr)
    return code


def run_extract(args) -> int:
    docx_path = Path(args.extract)
    output = Path(args.output) if args.output else default_extract_path(docx_path)
    extra = {"docx": str(docx_path), "output": str(output)}
    if not docx_path.is_file():
        return _error(args, "extract", 2, "FILE_NOT_FOUND", f"файл не найден: {docx_path}", **extra)
    if docx_structure.is_inside_skill_dir(output):
        return _error(args, "extract", 2, "OUTPUT_INSIDE_SKILL",
                      f"{output} внутри каталога скилла: укажи -o в каталоге проекта, "
                      "например <PROJECT_DIR>/audit/sources-extracted.json", **extra)
    refusal = protected_registry_reason(output)
    if refusal:
        return _error(args, "extract", 2, "REGISTRY_PROTECTED", refusal, **extra)
    if output.exists() and not args.force:
        return _error(args, "extract", 2, "OUTPUT_EXISTS",
                      f"выходной файл уже существует: {output}. Укажи другой путь -o или добавь --force.", **extra)
    try:
        result = extract_bibliography(docx_path)
    except DependencyError as error:
        return _error(args, "extract", 2, "DEPENDENCY_MISSING", str(error), **extra)
    except Exception as error:  # повреждённый пакет, не DOCX
        return _error(args, "extract", 2, "DOCX_INVALID",
                      f"не удалось открыть DOCX {docx_path}: {type(error).__name__}: {error}", **extra)

    records = result["records"]
    payload: Dict[str, Any] = {
        "tool": "verify_sources", "version": TOOL_VERSION, "mode": "extract",
        "docx": str(docx_path), "output": str(output), "heading": result["heading"],
        "records": len(records), "warnings": result["warnings"],
    }
    if not records:
        if not result["found"]:
            message = ("не найден заголовок списка литературы («Список использованной литературы», «Список "
                       "использованных источников», «Список литературы», «Библиографический список», «Литература»); "
                       "строки оглавления не учитываются")
        else:
            unnumbered = sum(1 for item in result["warnings"] if item["code"] == "UNNUMBERED_PARAGRAPH")
            message = (f"в разделе «{result['heading']}» нет записей: запись — абзац с номером («1.», «1)»), "
                       f"с автонумерацией Word или стилем VKR Bibliography; абзацев без номера: {unnumbered}")
        payload.update(status="empty", exit_code=1, message=message)

        def show_empty() -> None:
            print(f"ОШИБКА: извлечено 0 записей — {message}. Файл {output} не создан.", file=sys.stderr)
            for item in result["warnings"][:10]:
                print(f"  {item['code']}: {item['text']}", file=sys.stderr)

        return _emit(args, payload, show_empty)

    try:
        common.atomic_write_text(output, common.dump_json(records))
    except OSError as error:
        return _error(args, "extract", 2, "OUTPUT_WRITE_FAILED", f"не удалось записать {output}: {error}", **extra)
    payload.update(status="ok", exit_code=0)
    script = Path(__file__).resolve()

    def show_ok() -> None:
        print(f"OK: извлечено записей: {len(records)} из раздела «{result['heading']}» -> {output}")
        if result["warnings"]:
            print(f"\nПредупреждения ({len(result['warnings'])}):")
            for item in result["warnings"]:
                print(f"  {item['code']}: {item['text']} — {item['message']}")
        print("\nДальше:")
        print(f'  python "{script}" --queries "{output.resolve()}"')
        print("  Перенеси записи в <PROJECT_DIR>/sources.json: разложи raw_text на поля по gost-citations.md;")
        print("  результат проверки каждого источника запиши командой")
        print(f'  python "{script}" --mark <PROJECT_DIR>/sources.json --id ID --status confirmed --method catalog --url URL')

    return _emit(args, payload, show_ok)


def run_report(args) -> int:
    path = Path(args.report)
    try:
        sources = load_registry(path)
    except RegistryError as error:
        code = "FILE_NOT_FOUND" if error.code == "FILE_MISSING" else error.code
        return _error(args, "report", 2, code, error.message, registry=str(path))
    report = build_report(sources, str(path))
    return _emit(args, report, lambda: print_report(report))


def run_queries(args) -> int:
    path = Path(args.queries)
    try:
        sources = load_registry(path)
    except RegistryError as error:
        code = "FILE_NOT_FOUND" if error.code == "FILE_MISSING" else error.code
        return _error(args, "queries", 2, code, error.message, registry=str(path))
    items: List[Dict[str, Any]] = []
    skipped = {"confirmed": 0, "rejected": 0, "not_object": 0}
    for index, source in enumerate(sources, start=1):
        if not isinstance(source, dict):
            skipped["not_object"] += 1
            continue
        status = common.normalize_source_status(source.get("status"))
        if status == "rejected":
            skipped["rejected"] += 1
            continue
        missing = verification_missing(source) if status == "confirmed" else []
        stale = status == "confirmed" and not missing and verification_stale(source)
        if status == "confirmed" and not missing and not stale:
            skipped["confirmed"] += 1
            continue
        hints = source_hints(source)
        if missing:
            reason = "confirmed без полного verification: " + ", ".join(missing)
        elif stale:
            reason = "confirmed, но реквизиты изменены после проверки (verification.fields_sha256)"
        else:
            reason = f"status: {source.get('status') or 'не указан'}"
        items.append({
            "id": source.get("id"),
            "index": index,
            "label": source_label(source, index),
            "description": describe_source(source),
            "reason": reason,
            "hints": {key: hints[key] for key in ("title", "subtitle", "author_last", "author_initials", "year", "doi", "isbn", "url", "type", "language") if hints.get(key)},
            "queries": build_search_queries(hints),
        })
    payload = {
        "tool": "verify_sources", "version": TOOL_VERSION, "mode": "queries", "registry": str(path),
        "status": "ok" if sources else "empty", "exit_code": 0 if sources else 1,
        "total": len(sources), "items": items, "skipped": skipped,
    }

    def show() -> None:
        if not sources:
            print("Реестр пуст: запросы строить не для чего (код 1).")
            return
        print("Поисковые запросы для проверки источников. Скрипт сам не ищет: выполни их через web_search/web_fetch.")
        for item in items:
            print(f"\n{item['label']} {item['description']}  ({item['reason']})")
            for query in item["queries"]:
                print(f"  {query['engine']}: {query['query']}  — {query['purpose']}")
        print(f"\nПропущено: confirmed с полным verification — {skipped['confirmed']}, rejected — {skipped['rejected']}"
              + (f", не объекты — {skipped['not_object']}" if skipped["not_object"] else "") + ".")

    return _emit(args, payload, show)


def _find_record(sources: List[Any], source_id: Optional[str], index: Optional[int]) -> Dict[str, Any]:
    if index is not None:
        if index < 1 or index > len(sources) or not isinstance(sources[index - 1], dict):
            raise RegistryError("SOURCE_NOT_FOUND", f"нет записи-объекта в позиции #{index}")
        return sources[index - 1]
    key = str(source_id).strip().casefold()
    matches = [item for item in sources if isinstance(item, dict) and item.get("id") is not None
               and str(item.get("id")).strip().casefold() == key]
    if not matches:
        raise RegistryError("SOURCE_NOT_FOUND", f"в реестре нет записи с id {source_id}")
    if len(matches) > 1:
        raise RegistryError("SOURCE_ID_DUPLICATE", f"id {source_id} встречается {len(matches)} раза: исправь реестр")
    return matches[0]


def run_mark(args) -> int:
    path = Path(args.mark)
    extra = {"registry": str(path)}
    if (args.id is None) == (args.index is None):
        return _error(args, "mark", 2, "USAGE", "укажи ровно одно: --id ID или --index N", **extra)
    status = common.normalize_source_status(args.status)
    if status not in common.SOURCE_STATUSES:
        return _error(args, "mark", 2, "USAGE", "--status: " + ", ".join(common.SOURCE_STATUSES), **extra)
    method = (args.method or "").strip()
    url = (args.url or "").strip()
    note = (args.note or "").strip()
    if method and method not in common.VERIFICATION_METHODS:
        return _error(args, "mark", 2, "USAGE", "--method: " + ", ".join(common.VERIFICATION_METHODS), **extra)
    if status == "confirmed" and not (method and (url or note)):
        return _error(args, "mark", 2, "USAGE",
                      "для confirmed нужны --method и --url или --note (чем и где подтверждено)", **extra)
    if status in ("suspicious", "rejected") and not note:
        return _error(args, "mark", 2, "USAGE", f"для {status} нужна --note с причиной", **extra)
    if docx_structure.is_inside_skill_dir(path):
        return _error(args, "mark", 2, "OUTPUT_INSIDE_SKILL", f"{path} внутри каталога скилла: реестр должен быть в проекте", **extra)
    try:
        sources = load_registry(path)
        record = _find_record(sources, args.id, args.index)
    except RegistryError as error:
        code = "FILE_NOT_FOUND" if error.code == "FILE_MISSING" else error.code
        return _error(args, "mark", 2, code, error.message, **extra)
    record["status"] = status
    if status == "pending":
        record.pop("verification", None)
    else:
        verification: Dict[str, Any] = {"checked_at": common.utc_now_iso()}
        if method:
            verification["method"] = method
        verification["url"] = url
        verification["note"] = note
        # Реквизиты на момент проверки: их правка потом видна в --report и build_vkr.py (SOURCE_VERIFICATION_STALE).
        verification["fields_sha256"] = bibliographic_fields_sha256(record)
        record["verification"] = verification
    try:
        common.atomic_write_json(path, sources)
    except OSError as error:
        return _error(args, "mark", 2, "OUTPUT_WRITE_FAILED", f"не удалось записать {path}: {error}", **extra)
    warnings = []
    if status == "confirmed" and not record.get("title"):
        warnings.append("запись без title: build_vkr.py не оформит её, пока raw_text не разложен на поля по gost-citations.md")
    payload = {"tool": "verify_sources", "version": TOOL_VERSION, "mode": "mark", "registry": str(path),
               "status": "ok", "exit_code": 0, "record": record, "warnings": warnings}

    def show() -> None:
        label = source_label(record, args.index or 0)
        print(f"OK: {label} {describe_source(record)} -> status {status}")
        if record.get("verification"):
            print("    verification: " + ", ".join(f"{k}={v}" for k, v in record["verification"].items() if v))
        for warning in warnings:
            print(f"ВНИМАНИЕ: {warning}")

    return _emit(args, payload, show)


def main(argv: Optional[Sequence[str]] = None) -> int:
    common.reconfigure_stdio()
    parser = argparse.ArgumentParser(
        description="Реестр источников ВКР: извлечение из DOCX, отчёт по статусам, поисковые запросы "
                    "(скрипт сам в интернет не обращается)",
        epilog="Коды выхода: 0 — успех; 1 — --extract: 0 записей, --report: есть непроверенные/ошибки или "
               "реестр пуст, --queries: реестр пуст; 2 — нет файла, некорректный JSON или DOCX, нет python-docx, "
               "выходной файл существует или защищён, неверные аргументы --mark.",
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--extract", metavar="DOCX",
                       help="Извлечь список литературы из DOCX в JSON (записи {id, raw_text, status, hints})")
    group.add_argument("--report", metavar="REGISTRY", help="Сводка по статусам реестра sources.json")
    group.add_argument("--queries", metavar="REGISTRY", help="Поисковые запросы для непроверенных источников")
    group.add_argument("--mark", metavar="REGISTRY",
                       help="Записать результат проверки: status и verification (checked_at ставит скрипт)")
    parser.add_argument("--id", help="--mark: id записи")
    parser.add_argument("--index", type=int, help="--mark: номер позиции записи (1…), если id нет")
    parser.add_argument("--status", default="confirmed", help="--mark: pending | confirmed | suspicious | rejected")
    parser.add_argument("--method", help="--mark: catalog | doi | publisher | original | web_search")
    parser.add_argument("--url", help="--mark: где подтверждено")
    parser.add_argument("--note", help="--mark: что сверено или причина suspicious/rejected")
    parser.add_argument("-o", "--output",
                        help="--extract: файл записей (по умолчанию <PROJECT_DIR>/audit/sources-extracted.json, "
                             "иначе рядом с DOCX); --report/--queries: JSON-отчёт в файл")
    parser.add_argument("--force", action="store_true",
                        help="--extract: перезаписать существующий файл (реестр проекта sources.json со "
                             "структурированными записями не перезаписывается никогда)")
    parser.add_argument("--json", action="store_true", help="Только JSON в stdout")
    args = parser.parse_args(argv)
    if args.extract:
        return run_extract(args)
    if args.report:
        return run_report(args)
    if args.mark:
        return run_mark(args)
    return run_queries(args)


if __name__ == "__main__":
    raise SystemExit(main())
