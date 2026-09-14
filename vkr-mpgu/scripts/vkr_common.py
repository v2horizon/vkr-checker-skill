#!/usr/bin/env python3
"""Общие функции скилла vkr-mpgu: JSON, время, снимок аудита, словари статусов,
реестр замечаний и запуск валидатора.

Модуль без внешних зависимостей. Его используют ``vkr_audit.py``,
``vkr_project_doctor.py``, ``vkr_memory.py`` и ``init_vkr_project.py``.
Служебные записи аудита (хеши, попытки, execution_id, метки времени) создают
только эти инструменты; агент их руками не пишет.
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import os
import posixpath
import re
import sys
import tempfile
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from collections import OrderedDict
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

SKILL_VERSION = "6.33-production"
PROTOCOL_VERSION = "6.33"
VALIDATOR_VERSION = "6.33"

SCRIPTS_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPTS_DIR.parent

# ---------------------------------------------------------------------------
# Словари
# ---------------------------------------------------------------------------

VKR_TYPES = ("project", "regular")
PROFILES = ("generic", "mpgu-09-project", "mpgu-09-regular")
PROFILE_TYPES = {"mpgu-09-project": "project", "mpgu-09-regular": "regular"}
MODES = ("standard", "express-14d", "express-7d", "express-4d")
AUDIT_INTENSITIES = ("balanced", "strict", "maximum")
INTENSITY_RANK = {name: rank for rank, name in enumerate(AUDIT_INTENSITIES)}
TITLE_PAGE_KEYS = (
    "university", "institute", "department", "program_code", "program_name", "program_profile",
    "work_type", "author", "group", "supervisor", "supervisor_title", "head_of_department", "head_title",
    "city", "year",
)
# Поля титула, которые сборка (create_vkr_docx.py) при пустом значении заменяет маркером [ЗАПОЛНИТЬ: …].
TITLE_PAGE_BUILD_REQUIRED = (
    "institute", "program_code", "program_name", "program_profile", "supervisor_title", "supervisor",
    "department", "head_title", "head_of_department",
)
# Обязательные поля title_page для doctor (TITLE_PAGE_INCOMPLETE): поля сборки и автор.
TITLE_PAGE_REQUIRED = ("author",) + TITLE_PAGE_BUILD_REQUIRED
# Плоские поля титула в vkr-project.json и intake 6.32 → ключи title_page.
LEGACY_FLAT_TITLE_KEYS = OrderedDict((
    ("student_name", "author"), ("institution", "university"), ("institute", "institute"),
    ("program_code", "program_code"), ("program_name", "program_name"), ("supervisor", "supervisor"),
    ("department", "department"), ("group", "group"),
))

CLAIM_STATUSES = ("pending", "confirmed", "partial", "unsupported", "invalidated")
CLAIM_STATUS_SYNONYMS = {"supported": "confirmed", "verified": "confirmed"}
SOURCE_STATUSES = ("pending", "confirmed", "suspicious", "rejected")
SOURCE_STATUS_SYNONYMS = {"verified": "confirmed"}
VERIFICATION_METHODS = ("catalog", "doi", "publisher", "original", "web_search")

RUN_STATUSES = ("planned", "running", "reported", "complete", "invalidated", "abandoned")
RUN_OPEN_STATUSES = ("planned", "running", "reported")
TASK_STATUSES = (
    "planned", "running", "pass", "fail", "partial", "not_applicable",
    "not_recheckable", "interrupted", "error", "cancelled",
)
REPORT_STATUSES = ("pass", "fail", "partial", "not_applicable", "not_recheckable")
UNCLOSED_ATTEMPT_STATUSES = ("fail", "partial", "not_recheckable")
FINDING_STATUSES = ("open", "resolved", "rejected", "accepted_risk", "waived")
SEVERITIES = ("BLOCKER", "MAJOR", "MINOR", "INFO")
HIGH_SEVERITIES = ("BLOCKER", "MAJOR")
EVIDENCE_LEVELS = ("E0", "E1", "E2", "E3", "E4")
FIX_CLASSES = ("AUTO_SAFE", "AUTO_CONTENT", "USER_DATA_REQUIRED", "SUPERVISOR_APPROVAL", "MANUAL_WORD")
RECHECK_VERDICTS = ("pass", "fail", "partial", "not_recheckable", "invalid_finding")
AUDIT_KINDS = ("primary", "targeted_recheck", "blind_regression")
INDEPENDENCE_LEVELS = ("isolated_contexts", "diverse_models", "diverse_providers", "degraded_independence")
DEGRADED = "degraded_independence"
REVIEW_ORDERS = ("forward", "backward", "evidence_first", "adversarial")
READINESS_LEVELS = ("READY_TO_SUBMIT", "READY_FOR_SUPERVISOR_REVIEW", "NOT_READY")

BASE_GATES = ("MET", "SRC", "LOG", "LNG", "STY", "EVD", "TEC", "DOC", "OWN")
SUBJECTIVE_GATES = ("STY", "LNG", "LOG")
FULL_GATE_LENS = "full_gate"
SPECIALIZED_ROLES = OrderedDict(
    (
        ("SRC-BIB", ("SRC", "bibliography")),
        ("SRC-CIT", ("SRC", "claim_support")),
        ("SRC-CROSSREF", ("SRC", "cross_references")),
        ("DOC-STRUCTURE", ("DOC", "structure")),
        ("DOC-VISUAL", ("DOC", "visual_render")),
        ("MET-FORMAT", ("MET", "format_requirements")),
        ("OWN-QUESTIONS", ("OWN", "question_generation")),
        ("OWN-DEFENSE", ("OWN", "interactive_defense")),
    )
)
ROLE_SCOPES = {
    "MET": "Методист: приоритет документов, выбранный профиль, структура, обязательные части, "
           "три главы проектной ВКР 09.03.02, образовательный продукт, пилот, аннотация, заключение "
           "и допустимость отклонений.",
    "SRC": "Библиограф и фактчекер: существование каждого источника, авторы, название, год, "
           "DOI/ISBN/URL, страницы, точность цитат, соответствие источника поддерживаемому тезису "
           "и связь ссылок со списком.",
    "LOG": "Аудитор научной логики: связь темы, проблемы, цели, задач, методов, глав, результатов и "
           "выводов; противоречия, необоснованная причинность и скачки аргументации.",
    "LNG": "Корректор: орфография, пунктуация, грамматика, синтаксис, управление, согласование, "
           "терминология, повторы и изменение смысла из-за языковой ошибки.",
    "STY": "Аудитор академического стиля: ИИ-клише, монотонность, повторяющиеся начала, чрезмерная "
           "шаблонность, несоответствие уровню бакалавра; эвристики не выдаются за доказательство авторства.",
    "EVD": "Аудитор продукта и данных: соответствие текста коду, функциям, датам, исходным данным, "
           "расчётам, участникам, процедуре пилота и реальным результатам.",
    "TEC": "Технический аудитор: корректность архитектуры, кода, алгоритмов, API, версий, ограничений "
           "и описания программного продукта; соответствие текста предоставленным исходникам.",
    "DOC": "Нормоконтролёр: DOCX, A4, поля, стили, нумерация, оглавление, таблицы, рисунки, приложения, "
           "титул, метаданные, плейсхолдеры и визуальный рендер.",
    "OWN": "Экзаменатор: способность пользователя объяснить понятия, решения, источники, цифры, методы "
           "и результаты; вопросы комиссии и Test-30 по записанным ответам пользователя.",
}
LENS_DESCRIPTIONS = {
    FULL_GATE_LENS: "полная проверка всей области базовой роли",
    "bibliography": "реквизиты и оформление записей списка литературы, дубли, существование источников",
    "claim_support": "поддерживает ли источник конкретный тезис; точность цитат и страниц",
    "cross_references": "соответствие ссылок [N] в тексте записям списка литературы",
    "structure": "структура DOCX: разделы, заголовки, нумерация, оглавление, приложения",
    "visual_render": "визуальный рендер: поля, шрифты, таблицы, рисунки, колонтитулы",
    "format_requirements": "формальные требования методички к оформлению",
    "question_generation": "вопросы комиссии и критерии оценки ответов",
    "interactive_defense": "оценка записанных ответов пользователя (Test-30) из evidence/defense/",
}
REVIEW_ORDER_DESCRIPTIONS = {
    "forward": "прямая трассировка: от темы, цели и задач к результатам и выводам",
    "backward": "обратная трассировка: от заключения и выводов назад к доказательствам",
    "evidence_first": "сначала доказательства (источники, данные, код), затем опирающийся на них текст",
    "adversarial": "состязательно: ищи самое слабое место и пытайся опровергнуть ключевые утверждения",
}
AUDITOR_ID_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$")
FORBIDDEN_AUDITOR_TOKENS = {"main", "author", "student", "self", "user", "coordinator", "claude", "assistant"}
DOCUMENT_EXTENSIONS = {".docx", ".doc", ".docm", ".dot", ".dotx", ".dotm", ".pdf", ".odt", ".ott", ".fodt", ".rtf", ".xps"}
LOCK_FILE_MAX_BYTES = 1024
BUILD_MANIFEST_PATH = "exports/build-manifest.json"
DOCX_FIELDS_PATH = "exports/docx-fields.json"
FINAL_DOCX_PATH = "final/vkr.docx"
RUN_CLOSED_STATUSES = ("complete", "reported", "invalidated", "abandoned")
NEGATIVE_VERDICTS = ("fail", "partial")
EVIDENCE_GROUPS = ("methodology", "sources", "product", "pilot", "figures", "defense", "approvals")
EVIDENCE_ENTRY_KEYS = ("id", "path", "url", "kind", "title", "note", "checked_at")
EVIDENCE_LOCAL_PREFIXES = ("evidence/", "sources/materials/")
METHODOLOGY_ID = "methodology-09-03-02-2024"
METHODOLOGY_ASSET = "methodology-09-03-02-2024.docx"
METHODOLOGY_PROJECT_PATH = "sources/materials/methodology-09-03-02-2024.docx"
STATE_LOGICAL_ID = "state/project-state"
# Служебные разделы state, которые не входят в снимок: точные названия шаблона
# (без учёта регистра и пробелов). Любой другой раздел — содержательный.
STATE_EXCLUDED_TITLES = (
    "Независимый аудит",
    "История сессий",
    "Журнал текущей работы",
    "Handoff и продолжение",
    "Память и продолжение",
)
IGNORED_FILE_NAMES = {"thumbs.db", "desktop.ini", ".ds_store"}


def is_ignored_file_name(name: str) -> bool:
    """Lock-файлы Word, временные файлы и системный мусор не являются входами."""
    lowered = name.casefold()
    return (
        name.startswith("~$")
        or lowered in IGNORED_FILE_NAMES
        or lowered.endswith((".tmp", ".temp", ".lock"))
        or name.startswith(".~lock")
    )


# ---------------------------------------------------------------------------
# Консоль, время, хеши
# ---------------------------------------------------------------------------


def reconfigure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except (OSError, ValueError):  # pragma: no cover - нестандартный поток
                pass


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds")


_ISO_RE = re.compile(
    r"^(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})"
    r"(?:[Tt ](?P<hour>\d{2}):(?P<minute>\d{2})"
    r"(?::(?P<second>\d{2})(?:[.,](?P<fraction>\d+))?)?"
    r"(?P<tz>[Zz]|[+-]\d{2}(?::?\d{2})?)?)?$"
)


def parse_iso8601(value: Any, *, require_timezone: bool = False) -> Optional[datetime]:
    """Разбирает ISO 8601 одинаково на Python 3.9–3.13.

    Принимает ``Z``, смещения ``+03:00``/``+0300``/``+03``, дробную часть любой
    длины (лишние знаки отбрасываются) и дату без времени. Возвращает ``None``
    для некорректного значения.
    """
    if isinstance(value, datetime):
        parsed = value
    else:
        raw = str(value or "").strip()
        match = _ISO_RE.match(raw)
        if not match:
            return None
        parts = match.groupdict()
        fraction = (parts.get("fraction") or "")[:6].ljust(6, "0")
        tzinfo = None
        tz = parts.get("tz")
        if tz:
            if tz in ("Z", "z"):
                tzinfo = timezone.utc
            else:
                sign = 1 if tz[0] == "+" else -1
                digits = tz[1:].replace(":", "")
                hours = int(digits[:2])
                minutes = int(digits[2:4]) if len(digits) > 2 else 0
                if hours > 23 or minutes > 59:
                    return None
                tzinfo = timezone(sign * timedelta(hours=hours, minutes=minutes))
        try:
            parsed = datetime(
                int(parts["year"]), int(parts["month"]), int(parts["day"]),
                int(parts.get("hour") or 0), int(parts.get("minute") or 0),
                int(parts.get("second") or 0), int(fraction), tzinfo=tzinfo,
            )
        except ValueError:
            return None
    if require_timezone and parsed.tzinfo is None:
        return None
    return parsed


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def canonical_json_bytes(value: Any) -> bytes:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")


def content_sha256(path: Path) -> str:
    """Хеш входа снимка и сборки.

    Для ``.json`` — SHA-256 канонической сериализации: отступы, BOM и порядок
    ключей объекта не важны, порядок элементов массивов важен. Нечитаемый JSON
    хешируется побайтно (ошибку разбора сообщают обычные проверки JSON).
    """
    path = Path(path)
    if path.suffix.casefold() != ".json":
        return sha256_file(path)
    raw = path.read_bytes()
    try:
        parsed = json.loads(raw.decode("utf-8-sig"))
    except (UnicodeDecodeError, ValueError, RecursionError):
        return sha256_bytes(raw)
    return sha256_bytes(canonical_json_bytes(parsed))


# ---------------------------------------------------------------------------
# Чтение и атомарная запись
# ---------------------------------------------------------------------------


class ProjectDataError(ValueError):
    """Ошибка входных данных с указанием файла и машинного кода."""

    def __init__(self, path: str, message: str, code: str = "JSON_INVALID") -> None:
        self.path = path
        self.code = code
        self.detail = message
        super().__init__(f"{path}: {message}")


_MISSING = object()


def read_text_utf8(path: Path, label: Optional[str] = None) -> str:
    label = label or str(path)
    try:
        data = Path(path).read_bytes()
    except OSError as error:
        raise ProjectDataError(label, f"не удалось прочитать файл: {error}", "IO_ERROR") from error
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError as error:
        raise ProjectDataError(
            label,
            f"файл не в кодировке UTF-8 (байт {error.start}); пересохрани его в UTF-8",
            "ENCODING_INVALID",
        ) from error
    return text


def read_json(path: Path, default: Any = _MISSING, label: Optional[str] = None) -> Any:
    label = label or str(path)
    if not Path(path).is_file():
        if default is _MISSING:
            raise ProjectDataError(label, "файл отсутствует", "FILE_MISSING")
        return default
    text = read_text_utf8(path, label)
    try:
        return json.loads(text)
    except json.JSONDecodeError as error:
        raise ProjectDataError(
            label, f"некорректный JSON: {error.msg} (строка {error.lineno}, столбец {error.colno})", "JSON_INVALID"
        ) from error


def read_jsonl(path: Path, label: Optional[str] = None) -> List[Tuple[int, Any, Optional[str]]]:
    """Возвращает ``(номер строки, объект или None, ошибка или None)``."""
    label = label or str(path)
    if not Path(path).is_file():
        return []
    text = read_text_utf8(path, label)
    rows: List[Tuple[int, Any, Optional[str]]] = []
    for number, line in enumerate(text.splitlines(), start=1):
        if not line.strip():
            continue
        try:
            rows.append((number, json.loads(line), None))
        except json.JSONDecodeError as error:
            rows.append((number, None, f"некорректный JSON: {error.msg}"))
    return rows


def dump_json(data: Any) -> str:
    return json.dumps(data, ensure_ascii=False, indent=2) + "\n"


def atomic_write_text(path: Path, text: str) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=str(path.parent))
    try:
        with os.fdopen(handle, "w", encoding="utf-8", newline="\n") as stream:
            stream.write(text)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, str(path))
    except BaseException:
        try:
            os.unlink(temporary)
        except OSError:
            pass
        raise


def atomic_write_json(path: Path, data: Any) -> None:
    atomic_write_text(path, dump_json(data))


def emit_json(data: Any, output: Optional[Path] = None) -> None:
    """Печатает JSON в stdout или пишет UTF-8 без BOM в файл ``-o``."""
    payload = dump_json(data)
    if output is not None:
        try:
            atomic_write_text(Path(output), payload)
        except OSError as error:
            # Например, `-o nul` в Windows: устройство нельзя подменить атомарно. Код 2, а не трассировка.
            message = f"не удалось записать отчёт в {output}: {error}"
            sys.stderr.write(message + "\n")
            sys.stdout.write(dump_json({"status": "error", "error_code": "OUTPUT_UNWRITABLE", "error": message}))
            raise SystemExit(2)
    else:
        sys.stdout.write(payload)


# ---------------------------------------------------------------------------
# Пути и логические идентификаторы
# ---------------------------------------------------------------------------


def normalize_logical_id(raw: Any) -> str:
    value = unicodedata.normalize("NFC", str(raw or "")).replace("\\", "/").strip()
    while value.startswith("./"):
        value = value[2:]
    return value


def is_url(value: Any) -> bool:
    return bool(re.match(r"^[A-Za-z][A-Za-z0-9+.-]*://", str(value or "").strip()))


def resolve_inside(root: Path, raw: Any) -> Optional[Path]:
    """Разрешает относительный путь внутри корня проекта или возвращает None."""
    value = normalize_logical_id(raw)
    if not value or value.startswith("/") or re.match(r"^[A-Za-z]:", value) or is_url(value):
        return None
    candidate = Path(root) / value
    try:
        resolved = candidate.resolve()
    except (OSError, RuntimeError):
        return None
    try:
        resolved.relative_to(root)
    except ValueError:
        return None
    return resolved


def logical_id_for(root: Path, path: Path) -> str:
    return unicodedata.normalize("NFC", Path(path).relative_to(root).as_posix())


def dedupe_key(logical_id: str) -> str:
    return logical_id.casefold() if os.name == "nt" else logical_id


def _is_word_lock(path: Path) -> bool:
    try:
        return path.name.startswith("~$") and path.stat().st_size < LOCK_FILE_MAX_BYTES
    except OSError:
        return False


def final_listing(root: Path) -> Tuple[List[Path], List[str]]:
    """DOCX в ``final/`` и посторонние документы или подкаталоги.

    Lock-файл Word ``~$…`` меньше 1 КБ игнорируется; настоящий документ с таким
    именем считается документом.
    """
    directory = Path(root) / "final"
    docs: List[Path] = []
    extras: List[str] = []
    if not directory.is_dir():
        return docs, extras
    for path in sorted(directory.iterdir()):
        if path.is_dir():
            extras.append(f"final/{path.name}/")
            continue
        if not path.is_file() or _is_word_lock(path):
            continue
        suffix = path.suffix.casefold()
        if suffix == ".docx":
            docs.append(path)
        elif suffix in DOCUMENT_EXTENSIONS:
            extras.append(f"final/{path.name}")
    return docs, extras


def final_docx_files(root: Path) -> List[Path]:
    return final_listing(root)[0]


# ---------------------------------------------------------------------------
# Проекция state и входы снимка
# ---------------------------------------------------------------------------

_FENCE_RE = re.compile(r"^[ \t]{0,3}(`{3,}|~{3,})")
_HEADING_RE = re.compile(r"^[ \t]{0,3}(#{1,6})[ \t]+(.*?)[ \t#]*$")
_LAST_UPDATE_PREFIX_RE = re.compile(r"^[\s>*_\-+#`]*")


def title_key(title: str) -> str:
    return re.sub(r"\s+", "", str(title or "").strip("*_` ")).casefold()


_EXCLUDED_TITLE_KEYS = {title_key(item) for item in STATE_EXCLUDED_TITLES}


def _is_last_update_line(line: str) -> bool:
    cleaned = _LAST_UPDATE_PREFIX_RE.sub("", line).casefold()
    return cleaned.startswith("последнее обновление")


def state_projection(text: str) -> str:
    """Проекция ``vkr-state.md`` для снимка аудита (SPEC 9.1).

    Удаляются служебные разделы ``##`` с точными названиями из
    ``STATE_EXCLUDED_TITLES`` (без учёта регистра и пробелов) и строки
    «Последнее обновление…». Переводы строк — LF, хвостовые пробелы и пустые
    строки в конце не учитываются.
    """
    text = str(text or "").lstrip("\ufeff").replace("\r\n", "\n").replace("\r", "\n")
    result: List[str] = []
    skipping = False
    fence: Optional[str] = None
    for line in text.split("\n"):
        fence_match = _FENCE_RE.match(line)
        if fence is not None:
            if fence_match and fence_match.group(1)[0] == fence[0] and len(fence_match.group(1)) >= len(fence):
                fence = None
            if not skipping:
                result.append(line.rstrip())
            continue
        if fence_match:
            fence = fence_match.group(1)
            if not skipping:
                result.append(line.rstrip())
            continue
        heading = _HEADING_RE.match(line)
        if heading and len(heading.group(1)) <= 2:
            skipping = len(heading.group(1)) == 2 and title_key(heading.group(2)) in _EXCLUDED_TITLE_KEYS
        if skipping or _is_last_update_line(line):
            continue
        result.append(line.rstrip())
    return "\n".join(result).strip("\n") + "\n"


def _problem(code: str, severity: str, path: str, message: str) -> Dict[str, str]:
    return {"code": code, "severity": severity, "path": path, "message": message}


def evidence_entries(root: Path, index: Any) -> Tuple[List[Dict[str, Any]], List[Dict[str, str]]]:
    """Проверяет ``evidence/index.json`` и возвращает нормализованные записи.

    Запись: ``{"id", "path"}`` или ``{"id", "url"}``; дополнительно ``kind``,
    ``title``, ``note``, ``checked_at``. Локальный путь — только внутри
    ``evidence/`` или ``sources/materials/``.
    """
    problems: List[Dict[str, str]] = []
    entries: List[Dict[str, Any]] = []
    if index is None:
        return entries, problems
    if not isinstance(index, dict):
        problems.append(_problem("EVIDENCE_SHAPE", "ERROR", "evidence/index.json", "Ожидается JSON-объект с группами"))
        return entries, problems
    seen_ids: Dict[str, str] = {}
    for group, items in index.items():
        label = f"evidence/index.json:{group}"
        if group not in EVIDENCE_GROUPS:
            problems.append(_problem(
                "EVIDENCE_GROUP_UNKNOWN", "WARNING", label,
                "Неизвестная группа; допустимы: " + ", ".join(EVIDENCE_GROUPS),
            ))
        if not isinstance(items, list):
            problems.append(_problem("EVIDENCE_LIST_SHAPE", "ERROR", label, f"{group} должен быть массивом"))
            continue
        for position, item in enumerate(items, start=1):
            item_label = f"{label}[{position}]"
            if not isinstance(item, dict):
                problems.append(_problem("EVIDENCE_ENTRY_INVALID", "ERROR", item_label, "Запись должна быть объектом {id, path|url}"))
                continue
            unknown = sorted(set(item) - set(EVIDENCE_ENTRY_KEYS))
            if unknown:
                problems.append(_problem(
                    "EVIDENCE_ENTRY_UNKNOWN_KEYS", "ERROR", item_label,
                    "Неизвестные ключи: " + ", ".join(unknown) + "; допустимы: " + ", ".join(EVIDENCE_ENTRY_KEYS),
                ))
            entry_id = str(item.get("id") or "").strip()
            if not entry_id:
                problems.append(_problem("EVIDENCE_ENTRY_INVALID", "ERROR", item_label, "Нет id"))
                continue
            key = entry_id.casefold()
            if key in seen_ids:
                problems.append(_problem("EVIDENCE_ID_DUPLICATE", "ERROR", item_label, f"id {entry_id} уже есть в {seen_ids[key]}"))
                continue
            seen_ids[key] = item_label
            path_value = str(item.get("path") or "").strip()
            url_value = str(item.get("url") or "").strip()
            if bool(path_value) == bool(url_value):
                problems.append(_problem("EVIDENCE_ENTRY_INVALID", "ERROR", item_label, "Нужен ровно один из ключей path или url"))
                continue
            entry: Dict[str, Any] = {"id": entry_id, "group": group, "label": item_label, "raw": item}
            if url_value:
                if not is_url(url_value):
                    problems.append(_problem("EVIDENCE_URL_INVALID", "ERROR", item_label, "url должен начинаться со схемы, например https://"))
                    continue
                entry["url"] = url_value
                entries.append(entry)
                continue
            logical = normalize_logical_id(path_value)
            resolved = resolve_inside(root, logical)
            if resolved is None:
                problems.append(_problem("EVIDENCE_PATH_OUTSIDE", "ERROR", item_label, f"Путь {path_value} выходит за корень проекта или не является относительным"))
                continue
            logical = logical_id_for(root, resolved)
            if not logical.casefold().startswith(EVIDENCE_LOCAL_PREFIXES):
                problems.append(_problem(
                    "EVIDENCE_PATH_INVALID", "ERROR", item_label,
                    f"Доказательство {logical} должно лежать в evidence/ или sources/materials/, а не в служебных файлах проекта",
                ))
                continue
            if not resolved.is_file():
                problems.append(_problem("EVIDENCE_FILE_MISSING", "ERROR", item_label, f"Файл {logical} отсутствует"))
                continue
            entry["path"] = logical
            entry["file"] = resolved
            entries.append(entry)
    return entries, problems


def default_title_page() -> Dict[str, str]:
    page = {key: "" for key in TITLE_PAGE_KEYS}
    page["work_type"] = "Выпускная квалификационная работа"
    page["city"] = "Москва"
    page["year"] = str(datetime.now().year)
    return page


def title_page_missing(title_page: Any) -> List[str]:
    """Обязательные поля титула (``TITLE_PAGE_REQUIRED``), которые не заполнены."""
    if not isinstance(title_page, dict):
        return ["title_page"]
    missing: List[str] = []
    for key in TITLE_PAGE_REQUIRED:
        if key == "author":
            authors = title_page.get("authors")
            if isinstance(authors, list) and any(str(item or "").strip() for item in authors):
                continue
        if not str(title_page.get(key) or "").strip():
            missing.append(key)
    return missing


def migrate_flat_title_page(config: Dict[str, Any]) -> Optional[Dict[str, str]]:
    """Переносит плоские поля титула 6.32 (student_name, institution, …) в title_page.

    Срабатывает, только если title_page нет; плоские ключи удаляются. Возвращает
    созданный title_page или None, если переносить нечего.
    """
    if isinstance(config.get("title_page"), dict) or not any(key in config for key in LEGACY_FLAT_TITLE_KEYS):
        return None
    page = default_title_page()
    for legacy, target in LEGACY_FLAT_TITLE_KEYS.items():
        if legacy in config:
            value = " ".join(str(config.pop(legacy) or "").split())
            if value:
                page[target] = value
    config["title_page"] = page
    return page


def approval_problem(root: Path, approval: Any, expected_sha: Optional[str] = None) -> Optional[str]:
    """Почему файл не годится как одобрение руководителя для waive (None — годится).

    Одобрение — непустой файл внутри evidence/, зарегистрированный в
    evidence/index.json → approvals: так он входит в снимок и его видят аудиторы
    финальных волн. ``expected_sha`` сверяется при воспроизведении решения.
    """
    root = Path(root).resolve()
    logical = normalize_logical_id(approval)
    path = resolve_inside(root, logical)
    if path is None or not logical_id_for(root, path).startswith("evidence/"):
        return f"файл одобрения {logical} должен лежать внутри evidence/"
    if not path.is_file() or path.stat().st_size == 0:
        return f"файл одобрения {logical} отсутствует или пуст"
    try:
        index = read_json(root / "evidence" / "index.json", default=None, label="evidence/index.json")
    except ProjectDataError:
        index = None
    entries, _problems = evidence_entries(root, index)
    key = dedupe_key(logical_id_for(root, path))
    if not any(entry.get("group") == "approvals" and dedupe_key(str(entry.get("path") or "")) == key for entry in entries):
        return (
            f"файл одобрения {logical} не зарегистрирован в evidence/index.json → approvals "
            "(запись {id, path}), поэтому не входит в снимок и не виден аудиторам"
        )
    if expected_sha is not None and sha256_file(path) != expected_sha:
        return f"файл одобрения {logical} изменён после решения"
    return None


def project_profile(project: Any) -> str:
    if isinstance(project, dict) and project.get("profile") in PROFILES:
        return str(project["profile"])
    return "generic"


def collect_snapshot(root: Path, profile: str) -> Tuple[List[Dict[str, str]], List[Dict[str, str]]]:
    """Вычисляет входы снимка и список проблем, мешающих честному снимку."""
    root = Path(root).resolve()
    problems: List[Dict[str, str]] = []
    collected: "OrderedDict[str, Dict[str, str]]" = OrderedDict()

    def add(logical_id: str, kind: str, digest: str, override: bool = False) -> None:
        key = dedupe_key(logical_id)
        if key in collected and not override:
            return
        collected[key] = {"logical_id": logical_id, "kind": kind, "sha256": digest}

    def add_file(relative: str, kind: str) -> None:
        path = root / relative
        if path.is_file():
            add(relative, kind, content_sha256(path))

    add_file("vkr-project.json", "config")
    state_path = root / "vkr-state.md"
    if state_path.is_file():
        try:
            projection = state_projection(read_text_utf8(state_path, "vkr-state.md"))
        except ProjectDataError as error:
            problems.append(_problem(error.code, "ERROR", "vkr-state.md", error.detail))
        else:
            add(STATE_LOGICAL_ID, "state_projection", sha256_bytes(projection.encode("utf-8")))
    add_file("plan.md", "plan")
    add_file("sources.json", "sources")
    add_file("memory/claims-register.json", "claims")
    add_file("evidence/index.json", "evidence_index")

    drafts = root / "drafts"
    if drafts.is_dir():
        for path in sorted(drafts.glob("*.md")):
            if path.is_file() and not is_ignored_file_name(path.name):
                add(logical_id_for(root, path), "draft", sha256_file(path))

    index_path = root / "evidence" / "index.json"
    index: Any = None
    if index_path.is_file():
        try:
            index = read_json(index_path, label="evidence/index.json")
        except ProjectDataError as error:
            problems.append(_problem(error.code, "ERROR", "evidence/index.json", error.detail))
    entries, entry_problems = evidence_entries(root, index)
    problems.extend(entry_problems)
    evidence_kinds: Dict[str, str] = {}
    for entry in entries:
        if "file" in entry:
            kind = "methodology" if entry["group"] == "methodology" else f"evidence_{entry['group']}"
            evidence_kinds[dedupe_key(entry["path"])] = kind
            add(entry["path"], kind, content_sha256(entry["file"]), override=True)

    materials = root / "sources" / "materials"
    if materials.is_dir():
        for path in sorted(item for item in materials.rglob("*") if item.is_file()):
            if is_ignored_file_name(path.name) or "__pycache__" in path.parts:
                continue
            logical = logical_id_for(root, path)
            kind = evidence_kinds.get(dedupe_key(logical), "source_material")
            if profile.startswith("mpgu-") and path.name.casefold() == METHODOLOGY_ASSET:
                kind = "methodology"
            add(logical, kind, content_sha256(path))

    for path in final_docx_files(root):
        add(logical_id_for(root, path), "docx", sha256_file(path))
    add_file("audit/automated-validation.json", "validation_report")
    add_file(BUILD_MANIFEST_PATH, "build_manifest")
    # Число страниц по Word: файл входит в снимок, иначе его правку после снимка никто не заметит.
    add_file(DOCX_FIELDS_PATH, "docx_fields")

    inputs = sorted(collected.values(), key=lambda item: (item["logical_id"], item["kind"], item["sha256"]))
    return inputs, problems


def collect_inputs(root: Path, profile: str) -> List[Dict[str, str]]:
    """Входы снимка ``[{logical_id, kind, sha256}]`` (SPEC 9.1)."""
    inputs, _problems = collect_snapshot(root, profile)
    return inputs


def _draft_text_key(path: Path) -> str:
    text = path.read_bytes().decode("utf-8-sig", errors="replace")
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in text.splitlines()]
    normalized = re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()
    return sha256_bytes(normalized.encode("utf-8"))


def snapshot_content_key(root: Path) -> str:
    """Ключ текста работы для снимка: черновики без пробельных различий и отпечаток final/vkr.docx.

    accept-risk субъективных gate засчитывает волны только на снимках с разными
    ключами: правка plan.md или лишний пробел в черновике новой волны не дают.
    """
    root = Path(root).resolve()
    drafts: List[List[str]] = []
    folder = root / "drafts"
    if folder.is_dir():
        for path in sorted(folder.glob("*.md")):
            if path.is_file() and not is_ignored_file_name(path.name):
                drafts.append([logical_id_for(root, path), _draft_text_key(path)])
    final = root / FINAL_DOCX_PATH
    fingerprint = docx_text_fingerprint(final)[0] if final.is_file() else None
    return sha256_bytes(canonical_json_bytes({"drafts": drafts, "docx": fingerprint}))


def tool_versions() -> Dict[str, str]:
    return {"protocol": PROTOCOL_VERSION, "validator": VALIDATOR_VERSION}


def build_snapshot(root: Path, profile: str) -> Tuple[Dict[str, Any], List[Dict[str, str]]]:
    inputs, problems = collect_snapshot(root, profile)
    snapshot = {
        "protocol_version": PROTOCOL_VERSION,
        "validation_profile": profile,
        "inputs": inputs,
        "tool_versions": tool_versions(),
    }
    return snapshot, problems


def canonical_snapshot_sha(snapshot: Any) -> Optional[str]:
    """SHA-256 канонической проекции снимка (алгоритм не менялся с 6.30)."""
    if not isinstance(snapshot, dict):
        return None
    inputs = snapshot.get("inputs", [])
    if not isinstance(inputs, list):
        return None
    normalized_inputs = []
    for item in inputs:
        if not isinstance(item, dict):
            return None
        normalized_inputs.append(
            {
                "logical_id": unicodedata.normalize("NFC", str(item.get("logical_id") or "")),
                "kind": unicodedata.normalize("NFC", str(item.get("kind") or "")),
                "sha256": str(item.get("sha256") or "").casefold(),
            }
        )
    projection = {
        "protocol_version": snapshot.get("protocol_version"),
        "validation_profile": snapshot.get("validation_profile"),
        "inputs": sorted(normalized_inputs, key=lambda item: (item["logical_id"], item["kind"], item["sha256"])),
        "tool_versions": snapshot.get("tool_versions", {}),
    }
    encoded = json.dumps(projection, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def diff_inputs(old: Any, new: Any) -> Dict[str, List[str]]:
    def as_map(items: Any) -> Dict[str, Tuple[str, str, str]]:
        result: Dict[str, Tuple[str, str, str]] = {}
        if isinstance(items, list):
            for item in items:
                if isinstance(item, dict):
                    logical = normalize_logical_id(item.get("logical_id"))
                    result[dedupe_key(logical)] = (logical, str(item.get("kind") or ""), str(item.get("sha256") or "").casefold())
        return result

    before, after = as_map(old), as_map(new)
    return {
        "added": sorted(after[key][0] for key in after if key not in before),
        "removed": sorted(before[key][0] for key in before if key not in after),
        "changed": sorted(after[key][0] for key in after if key in before and before[key][1:] != after[key][1:]),
    }


# ---------------------------------------------------------------------------
# Нормализация значений
# ---------------------------------------------------------------------------


def normalize_status(value: Any) -> str:
    return str(value if value is not None else "").strip().casefold()


def normalize_source_status(value: Any) -> str:
    status = normalize_status(value)
    return SOURCE_STATUS_SYNONYMS.get(status, status)


def normalize_claim_status(value: Any) -> str:
    status = normalize_status(value)
    return CLAIM_STATUS_SYNONYMS.get(status, status)


def auditor_key(value: Any) -> str:
    """auditor_id для сравнений: без учёта регистра, пробелов и ширины символов."""
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(value or ""))).casefold()


def auditor_tokens(value: Any) -> List[str]:
    """Лексемы auditor_id: разбиение по ``. _ -``, границам букв и цифр и смене регистра.

    ``ClaudeMain`` → claude, main; ``authorReview`` → author, review; ``GPTMain`` → gpt, main.
    """
    raw = str(value or "")
    tokens = re.findall(r"[A-Z]+(?![a-z])|[A-Z]?[a-z]+|[0-9]+", raw)
    return [token.casefold() for token in tokens]


def auditor_id_problem(value: Any) -> Optional[str]:
    """Причина, по которой строка не годится как auditor_id, или None."""
    raw = str(value if value is not None else "")
    if not AUDITOR_ID_RE.match(raw):
        return (
            f"auditor_id {raw!r} недопустим: только латинские буквы, цифры и . _ - (до 64 символов, "
            "первый — буква или цифра), без пробелов и невидимых символов"
        )
    return None


def auditor_collapsed(value: Any) -> str:
    """auditor_id без разделителей: ``M.A.I.N`` → ``main``, ``ma-in`` → ``main``."""
    return re.sub(r"[._\-\s]+", "", unicodedata.normalize("NFKC", str(value or ""))).casefold()


def is_forbidden_auditor(value: Any) -> bool:
    """auditor_id обозначает автора, координатора или ИИ-ассистента основного контекста.

    Проверяются и лексемы (``ClaudeMain`` → claude, main), и склейка без
    разделителей (``M.A.I.N``, ``ma-in`` → main): точками и дефисами запрет не обходится.
    """
    if any(token in FORBIDDEN_AUDITOR_TOKENS for token in auditor_tokens(value)):
        return True
    return auditor_collapsed(value) in FORBIDDEN_AUDITOR_TOKENS


def role_gate_lens(role: str) -> Optional[Tuple[str, str]]:
    role = str(role or "").strip().upper()
    if role in BASE_GATES:
        return role, FULL_GATE_LENS
    return SPECIALIZED_ROLES.get(role)


def has_product_evidence(root: Path) -> bool:
    try:
        index = read_json(Path(root) / "evidence" / "index.json", default=None)
    except ProjectDataError:
        return False
    return isinstance(index, dict) and isinstance(index.get("product"), list) and bool(index.get("product"))


def not_applicable_allowed(gate: str, preset: str, vkr_type: str, has_docx: bool, has_product: bool) -> Tuple[bool, str]:
    """Когда аудитор вправе вернуть not_applicable (SPEC 9.4)."""
    gate = str(gate or "").upper()
    if gate in ("EVD", "TEC"):
        if vkr_type == "regular" and not has_product:
            return True, ""
        return False, f"{gate}: not_applicable допустим только для vkr_type=regular без материалов продукта"
    if gate == "DOC":
        if preset == "final":
            return False, "DOC: на финальной волне not_applicable недопустим"
        if has_docx:
            return False, "DOC: DOCX уже собран, not_applicable недопустим"
        return True, ""
    if gate == "OWN":
        return False, "OWN: not_applicable недопустим (на финале — никогда)"
    return False, f"{gate}: not_applicable для этой роли не предусмотрен"


# ---------------------------------------------------------------------------
# Manifest и реестр замечаний
# ---------------------------------------------------------------------------

MANIFEST_SCHEMA = "vkr-audit-manifest"
REPORT_SCHEMA_PREFIX = "vkr-audit-report/"
RUN_NUMBER_RE = re.compile(r"^R(\d+)-")
TOLERATED_REPLAY_CODES = ("AUDIT_WAIVER_INVALID", "AUDIT_DECISION_INVALID")


def manifest_skeleton(profile: str) -> Dict[str, Any]:
    return {
        "schema": MANIFEST_SCHEMA,
        "protocol_version": PROTOCOL_VERSION,
        "skill_version": SKILL_VERSION,
        "validation_profile": profile,
        "snapshot_manifest_sha256": None,
        "artifact_sha256": None,
        "seq": 0,
        "snapshots": [],
        "runs": [],
        "decisions": [],
    }


def findings_skeleton() -> Dict[str, Any]:
    return {"schema": "vkr-audit-findings", "protocol_version": PROTOCOL_VERSION, "generated_at": None, "findings": []}


def as_list(value: Any) -> List[Any]:
    return value if isinstance(value, list) else []


def safe_int(value: Any) -> Optional[int]:
    if isinstance(value, bool) or not isinstance(value, int):
        return None
    return value


def manifest_is_native(manifest: Any) -> bool:
    """Manifest создан инструментами 6.33 (есть схема)."""
    return isinstance(manifest, dict) and manifest.get("schema") == MANIFEST_SCHEMA


def _run_native_markers(run: Any) -> List[str]:
    if not isinstance(run, dict):
        return []
    markers = [key for key in ("planned_seq", "snapshot_seq", "closed_seq") if key in run]
    for task in as_list(run.get("planned_tasks")):
        if not isinstance(task, dict):
            continue
        for attempt in as_list(task.get("attempts")):
            if isinstance(attempt, dict) and ("seq" in attempt or "report_sha256" in attempt):
                markers.append(f"{task.get('task_id')}.attempts")
                break
    return markers


def native_markers(manifest: Any) -> List[str]:
    """Следы записей vkr_audit.py в manifest (для manifest без схемы — признак подделки)."""
    markers: List[str] = []
    if not isinstance(manifest, dict):
        return markers
    markers.extend(key for key in ("seq", "snapshots", "decisions", "legacy_runs") if key in manifest)
    for run in as_list(manifest.get("runs")):
        markers.extend(f"{run.get('run_id') if isinstance(run, dict) else '?'}:{item}" for item in _run_native_markers(run))
    return markers


def legacy_runs_sha(legacy_runs: Any) -> str:
    encoded = json.dumps(legacy_runs, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return sha256_bytes(encoded)


def migrate_legacy_manifest(manifest: Dict[str, Any], profile: str) -> Dict[str, Any]:
    """Первичная миграция manifest 6.32: все прежние run переносятся в legacy_runs."""
    migrated = manifest_skeleton(profile)
    legacy = list(as_list(manifest.get("runs")))
    migrated["legacy_runs"] = legacy
    migrated["legacy_runs_sha256"] = legacy_runs_sha(legacy)
    migrated["migrated_from"] = {
        "protocol_version": manifest.get("protocol_version"),
        "skill_version": manifest.get("skill_version"),
        "legacy_split_at_seq": 0,
        "at": utc_now_iso(),
    }
    return migrated


def _tampered(where: str, message: str) -> Dict[str, str]:
    return _problem("AUDIT_RUN_TAMPERED", "ERROR", where, message)


def manifest_integrity_problems(manifest: Any) -> List[Dict[str, str]]:
    """Структурная целостность manifest 6.33: run, задачи, попытки и непрерывность seq.

    Проверка ловит правки отдельных полей (удалённый planned_seq, null в runs,
    удалённый run, исправленный статус). Согласованную подделку всех файлов
    аудита без внешнего якоря она не обнаруживает.
    """
    label = "audit/manifest.json"
    problems: List[Dict[str, str]] = []
    if not isinstance(manifest, dict):
        return [_problem("MANIFEST_INVALID", "ERROR", label, "manifest должен быть JSON-объектом")]
    if not manifest_is_native(manifest):
        markers = native_markers(manifest)
        if markers:
            problems.append(_tampered(label, "у manifest нет схемы 6.33, но есть записи vkr_audit.py (" + ", ".join(markers[:5]) + ")"))
        return problems
    used: List[int] = []
    seq_total = safe_int(manifest.get("seq"))
    if seq_total is None or seq_total < 0:
        problems.append(_tampered(label, "нет целого счётчика seq"))
    snapshot_by_seq: Dict[int, str] = {}
    snapshots = manifest.get("snapshots")
    if not isinstance(snapshots, list):
        problems.append(_tampered(label, "snapshots должен быть массивом"))
        snapshots = []
    for index, entry in enumerate(snapshots):
        seq = safe_int(entry.get("seq")) if isinstance(entry, dict) else None
        sha = str(entry.get("snapshot_manifest_sha256") or "") if isinstance(entry, dict) else ""
        if seq is None or not re.fullmatch(r"[0-9a-f]{64}", sha):
            problems.append(_tampered(f"{label}:snapshots[{index}]", "запись снимка повреждена"))
            continue
        used.append(seq)
        snapshot_by_seq[seq] = sha
    if snapshots and isinstance(snapshots[-1], dict) and manifest.get("snapshot_manifest_sha256") != snapshots[-1].get("snapshot_manifest_sha256"):
        problems.append(_problem(
            "AUDIT_SNAPSHOT_HASH_MISMATCH", "ERROR", label,
            "snapshot_manifest_sha256 в manifest не совпадает с последней записью журнала snapshots",
        ))
    runs = manifest.get("runs")
    if not isinstance(runs, list):
        problems.append(_tampered(label, "runs должен быть массивом"))
        runs = []
    run_ids: set = set()
    task_ids: set = set()
    for index, run in enumerate(runs, start=1):
        where = f"{label}:runs[{index}]"
        if not isinstance(run, dict):
            problems.append(_tampered(where, "запись run не является объектом"))
            continue
        reasons: List[str] = []
        run_id = run.get("run_id")
        if not isinstance(run_id, str) or not run_id:
            reasons.append("нет run_id")
        elif run_id in run_ids:
            reasons.append(f"повтор run_id {run_id}")
        else:
            run_ids.add(run_id)
            where = f"{label}:{run_id}"
        if run.get("kind") not in AUDIT_KINDS:
            reasons.append(f"неизвестный kind {run.get('kind')!r}")
        planned = safe_int(run.get("planned_seq"))
        snap_seq = safe_int(run.get("snapshot_seq"))
        closed = safe_int(run.get("closed_seq"))
        status = normalize_status(run.get("status"))
        sha = str(run.get("snapshot_manifest_sha256") or "")
        if planned is None:
            reasons.append("нет целого planned_seq")
        else:
            used.append(planned)
        if snap_seq is None:
            reasons.append("нет целого snapshot_seq")
        if not re.fullmatch(r"[0-9a-f]{64}", sha):
            reasons.append("нет snapshot_manifest_sha256")
        elif snap_seq is not None:
            if snapshot_by_seq.get(snap_seq) != sha:
                reasons.append("snapshot_seq не указывает на этот снимок в журнале snapshots")
            elif planned is not None and snap_seq >= planned:
                reasons.append("снимок записан позже планирования run")
        if status not in RUN_STATUSES:
            reasons.append(f"неизвестный статус {run.get('status')!r}")
        if status in RUN_CLOSED_STATUSES:
            if closed is None:
                reasons.append("закрытый run без целого closed_seq")
        elif run.get("closed_seq") is not None:
            reasons.append("открытый run с closed_seq")
        if closed is not None:
            used.append(closed)
            if planned is not None and closed <= planned:
                reasons.append("closed_seq не позже planned_seq")
        last_close = closed
        for reclose in as_list(run.get("reclosed")):
            reclose_seq = safe_int(reclose.get("seq")) if isinstance(reclose, dict) else None
            if reclose_seq is None:
                reasons.append("повреждена запись reclosed")
            else:
                used.append(reclose_seq)
                if last_close is not None and reclose_seq <= last_close:
                    reasons.append("reclosed раньше закрытия")
        tasks = run.get("planned_tasks")
        if not isinstance(tasks, list) or not tasks:
            reasons.append("нет planned_tasks")
            tasks = []
        for task in tasks:
            if not isinstance(task, dict) or not isinstance(task.get("task_id"), str) or not task.get("task_id"):
                reasons.append("задача без task_id")
                continue
            task_id = task["task_id"]
            if task_id in task_ids:
                reasons.append(f"повтор task_id {task_id}")
            task_ids.add(task_id)
            if task.get("kind") != run.get("kind"):
                reasons.append(f"{task_id}: kind задачи не совпадает с run")
            if task.get("snapshot_manifest_sha256") != run.get("snapshot_manifest_sha256"):
                reasons.append(f"{task_id}: снимок задачи не совпадает с run")
            attempts = task.get("attempts")
            if not isinstance(attempts, list):
                reasons.append(f"{task_id}: attempts должен быть массивом")
                continue
            for number, attempt in enumerate(attempts, start=1):
                attempt_seq = safe_int(attempt.get("seq")) if isinstance(attempt, dict) else None
                if attempt_seq is None:
                    reasons.append(f"{task_id}: попытка {number} без seq")
                    continue
                used.append(attempt_seq)
                if safe_int(attempt.get("attempt")) != number:
                    reasons.append(f"{task_id}: нарушена нумерация попыток")
                if planned is not None and attempt_seq <= planned:
                    reasons.append(f"{task_id}: попытка раньше планирования run")
                if closed is not None and attempt_seq >= closed:
                    reasons.append(f"{task_id}: попытка после закрытия run")
            task_status = normalize_status(task.get("status"))
            if attempts and isinstance(attempts[-1], dict):
                last = attempts[-1]
                if (task_status != normalize_status(last.get("status")) or task.get("report_path") != last.get("report_path")
                        or task.get("report_sha256") != last.get("report_sha256")):
                    reasons.append(f"{task_id}: статус или отчёт задачи не совпадает с последней попыткой")
            elif not attempts and task_status not in ("planned", "cancelled"):
                reasons.append(f"{task_id}: статус {task.get('status')!r} без записанной попытки")
        problems.extend(_tampered(where, reason) for reason in reasons)
    for index, decision in enumerate(as_list(manifest.get("decisions"))):
        decision_seq = safe_int(decision.get("seq")) if isinstance(decision, dict) else None
        if decision_seq is None:
            problems.append(_tampered(f"{label}:decisions[{index}]", "решение без seq"))
        else:
            used.append(decision_seq)
    if seq_total is not None:
        counts: Dict[int, int] = {}
        for value in used:
            counts[value] = counts.get(value, 0) + 1
        duplicates = sorted(value for value, count in counts.items() if count > 1)
        missing = sorted(set(range(1, seq_total + 1)) - set(counts))
        outside = sorted(value for value in counts if value < 1 or value > seq_total)
        if duplicates or missing or outside:
            parts = []
            if missing:
                parts.append("пропущены " + ", ".join(map(str, missing[:10])))
            if duplicates:
                parts.append("повторены " + ", ".join(map(str, duplicates[:10])))
            if outside:
                parts.append("вне счётчика " + ", ".join(map(str, outside[:10])))
            problems.append(_problem(
                "AUDIT_SEQ_GAP", "ERROR", label,
                "Журнал seq manifest непрерывен только при штатной записи: " + "; ".join(parts)
                + " (удалённые или вставленные вручную записи аудита)",
            ))
    legacy = manifest.get("legacy_runs")
    if legacy is not None:
        if not isinstance(legacy, list) or manifest.get("legacy_runs_sha256") != legacy_runs_sha(legacy):
            problems.append(_tampered(label, "legacy_runs изменены после миграции 6.32"))
        elif any(_run_native_markers(item) for item in legacy):
            problems.append(_tampered(label, "в legacy_runs перенесены записи vkr_audit.py 6.33"))
    return problems


def next_run_number(root: Path, manifest: Dict[str, Any]) -> int:
    """Номер нового run: больше любого прежнего, в том числе удалённого из manifest."""
    names: List[Any] = [
        run.get("run_id") for run in as_list(manifest.get("runs")) + as_list(manifest.get("legacy_runs")) if isinstance(run, dict)
    ]
    for folder in ("reports", "briefs"):
        directory = Path(root) / "audit" / folder
        if directory.is_dir():
            names.extend(child.name for child in directory.iterdir())
    highest = 0
    for name in names:
        match = RUN_NUMBER_RE.match(str(name or ""))
        if match:
            highest = max(highest, int(match.group(1)))
    return highest + 1


def report_files_status(root: Path, manifest: Any) -> Tuple[List[str], List[str]]:
    """Отчёты в audit/reports без записи в manifest: (сироты, отчёты 6.32)."""
    root = Path(root)
    reports_dir = root / "audit" / "reports"
    if not reports_dir.is_dir():
        return [], []
    referenced: set = set()
    legacy_refs: set = set()
    native_runs = as_list(manifest.get("runs")) if manifest_is_native(manifest) else []
    legacy_runs = as_list(manifest.get("legacy_runs")) if manifest_is_native(manifest) else as_list(manifest.get("runs") if isinstance(manifest, dict) else [])
    for run in native_runs:
        for task in as_list(run.get("planned_tasks") if isinstance(run, dict) else []):
            for attempt in as_list(task.get("attempts") if isinstance(task, dict) else []):
                if isinstance(attempt, dict):
                    referenced.add(dedupe_key(normalize_logical_id(attempt.get("report_path"))))
    for run in legacy_runs:
        for task in as_list(run.get("planned_tasks") if isinstance(run, dict) else []):
            if not isinstance(task, dict):
                continue
            legacy_refs.add(dedupe_key(normalize_logical_id(task.get("report_path"))))
            for attempt in as_list(task.get("attempts")):
                if isinstance(attempt, dict):
                    legacy_refs.add(dedupe_key(normalize_logical_id(attempt.get("report_path"))))
    orphans: List[str] = []
    legacy: List[str] = []
    for path in sorted(reports_dir.rglob("*.json")):
        if not path.is_file():
            continue
        logical = logical_id_for(root, path)
        key = dedupe_key(logical)
        if key in referenced:
            continue
        try:
            data = read_json(path, label=logical)
        except ProjectDataError:
            data = None
        schema = data.get("report_schema") if isinstance(data, dict) else None
        native_report = isinstance(schema, str) and schema.startswith(REPORT_SCHEMA_PREFIX)
        if not native_report and (key in legacy_refs or path.parent == reports_dir):
            legacy.append(logical)
        else:
            orphans.append(logical)
    return orphans, legacy


def iter_attempts(manifest: Any) -> Iterable[Tuple[Dict[str, Any], Dict[str, Any], Dict[str, Any]]]:
    if not isinstance(manifest, dict):
        return
    for run in as_list(manifest.get("runs")):
        if not isinstance(run, dict):
            continue
        for task in as_list(run.get("planned_tasks")):
            if not isinstance(task, dict):
                continue
            for attempt in as_list(task.get("attempts")):
                if isinstance(attempt, dict):
                    yield run, task, attempt


def load_attempt_report(
    root: Path, run: Dict[str, Any], task: Dict[str, Any], attempt: Dict[str, Any]
) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """Читает канонический отчёт попытки и сверяет его с manifest по всем полям."""
    relative = normalize_logical_id(attempt.get("report_path"))
    if not relative.startswith("audit/reports/") or not relative.endswith(".json"):
        return None, "report_path должен указывать на audit/reports/<run>/<task>.json"
    path = resolve_inside(root, relative)
    if path is None or not path.is_file():
        return None, f"отчёт {relative} отсутствует"
    expected = str(attempt.get("report_sha256") or "").casefold()
    if not re.fullmatch(r"[0-9a-f]{64}", expected) or sha256_file(path) != expected:
        return None, f"report sha mismatch: {relative}"
    try:
        report = read_json(path, label=relative)
    except ProjectDataError as error:
        return None, str(error)
    if not isinstance(report, dict):
        return None, f"{relative}: отчёт должен быть объектом"
    expectations = (
        ("run_id", run.get("run_id")),
        ("task_id", task.get("task_id")),
        ("execution_id", attempt.get("execution_id")),
        ("snapshot_manifest_sha256", run.get("snapshot_manifest_sha256")),
        ("artifact_sha256", run.get("artifact_sha256")),
        ("kind", run.get("kind")),
        ("auditor_role", task.get("auditor_role")),
        ("base_gate", task.get("base_gate")),
        ("audit_lens", task.get("audit_lens")),
        ("replica_id", task.get("replica_id")),
        ("independence", run.get("independence")),
        ("attempt", attempt.get("attempt")),
        ("task_status", attempt.get("status")),
        ("status", attempt.get("report_status")),
    )
    mismatches = [key for key, expected_value in expectations if report.get(key) != expected_value]
    if auditor_key(report.get("auditor_id")) != auditor_key(attempt.get("auditor_id")):
        mismatches.append("auditor_id")
    report_ids = [item.get("finding_id") for item in as_list(report.get("findings")) if isinstance(item, dict)]
    if report_ids != as_list(attempt.get("finding_ids")):
        mismatches.append("finding_ids")
    if mismatches:
        return None, f"{relative}: поля отчёта не совпадают с manifest ({', '.join(mismatches)})"
    return report, None


def snapshot_is_newer(run: Dict[str, Any], source: Dict[str, Any]) -> bool:
    run_seq = safe_int(run.get("snapshot_seq"))
    source_seq = safe_int(source.get("snapshot_seq"))
    if run_seq is None or source_seq is None:
        return False
    return run_seq > source_seq and run.get("snapshot_manifest_sha256") != source.get("snapshot_manifest_sha256")


def preset_matrix(preset: str, intensity: str, vkr_type: str, subsection_index: int = 0) -> "OrderedDict[str, int]":
    """Роли и число реплик пресета по матрице continuous-audit.md."""
    strict = intensity in ("strict", "maximum")
    maximum = intensity == "maximum"
    project_type = vkr_type == "project"
    matrix: "OrderedDict[str, int]" = OrderedDict()

    def add(role: str, count: int = 1) -> None:
        matrix[role] = max(matrix.get(role, 0), count)

    double = 2 if maximum else 1
    if preset == "intake":
        for role in ("MET", "EVD", "TEC"):
            add(role)
    elif preset == "plan":
        add("MET"); add("LOG", 2); add("EVD"); add("TEC")
    elif preset == "subsection":
        add("LOG"); add("SRC"); add("LNG")
        if maximum or (strict and subsection_index % 2 == 1):
            add("STY")
    elif preset == "sources":
        add("SRC-BIB", 2); add("SRC-CIT", double)
    elif preset == "product":
        add("TEC", 2); add("EVD"); add("LOG"); add("OWN-QUESTIONS")
    elif preset == "chapter":
        add("MET", double); add("LOG", 2); add("SRC", double); add("LNG"); add("STY"); add("OWN-QUESTIONS")
        if project_type:
            add("EVD", double); add("TEC", double)
    elif preset == "pilot":
        add("EVD", 2); add("LOG"); add("SRC")
    elif preset == "docx":
        for role in ("DOC-STRUCTURE", "DOC-VISUAL", "MET-FORMAT", "SRC-CROSSREF"):
            add(role)
    elif preset == "draft":
        for role in ("MET", "SRC", "LOG", "LNG", "STY", "EVD", "TEC", "DOC", "OWN-QUESTIONS"):
            add(role)
    elif preset == "prefinal":
        for role in BASE_GATES:
            add(role)
    elif preset == "final":
        for role in BASE_GATES:
            add(role, double)
        if strict:
            for role in ("MET", "LOG", "SRC", "EVD", "TEC", "DOC"):
                add(role, 2)
    return matrix


def blind_replica_count(preset: str, intensity: str) -> int:
    return 2 if (preset == "final" and intensity == "maximum") else 1


def project_intensity(project: Any) -> str:
    if isinstance(project, dict) and project.get("audit_intensity") in AUDIT_INTENSITIES:
        return str(project["audit_intensity"])
    mode = str(project.get("mode") or "") if isinstance(project, dict) else ""
    return "balanced" if mode.startswith("express") else "strict"


def final_wave_requirements(kind: str, intensity: str, vkr_type: str) -> Dict[str, int]:
    """Минимальное число разных пройденных реплик full_gate на каждый базовый gate финальной волны."""
    if kind == "blind_regression":
        count = blind_replica_count("final", intensity)
        return {gate: count for gate in BASE_GATES}
    matrix = preset_matrix("final", intensity, vkr_type)
    return {gate: matrix.get(gate, 1) for gate in BASE_GATES}


def gate_waves_since(manifest: Dict[str, Any], gate: str, from_planned_seq: int, until_seq: int) -> List[Dict[str, Any]]:
    """Закрытые волны, проверявшие gate, запланированные не раньше from_planned_seq."""
    waves = []
    for run in as_list(manifest.get("runs")):
        if not isinstance(run, dict):
            continue
        planned = safe_int(run.get("planned_seq"))
        closed = safe_int(run.get("closed_seq"))
        if planned is None or closed is None or planned < from_planned_seq or closed >= until_seq:
            continue
        if normalize_status(run.get("status")) not in ("complete", "reported"):
            continue
        if any(
            isinstance(task, dict) and str(task.get("base_gate") or "").upper() == gate and as_list(task.get("attempts"))
            for task in as_list(run.get("planned_tasks"))
        ):
            waves.append(run)
    return waves


def accept_risk_problem(
    finding: Dict[str, Any], manifest: Dict[str, Any], registry: Dict[str, Dict[str, Any]], until_seq: int
) -> Optional[str]:
    """Причина, по которой accept-risk недопустим, или None."""
    if finding.get("status") != "open":
        return f"{finding.get('finding_id')}: статус {finding.get('status')}, принять риск можно только для open"
    severity = finding.get("severity")
    if severity in HIGH_SEVERITIES:
        return f"{finding.get('finding_id')}: {severity} нельзя закрыть accept-risk — исправь, получи invalid_finding с E3/E4 или waive по SUPERVISOR_APPROVAL"
    gate = finding.get("base_gate")
    if severity == "MINOR" and gate in SUBJECTIVE_GATES:
        source = finding.get("source") or {}
        from_seq = (safe_int(source.get("run_planned_seq")) or 0) + 1
        key_by_sha: Dict[Any, str] = {}
        for entry in as_list(manifest.get("snapshots")):
            if isinstance(entry, dict) and isinstance(entry.get("content_key"), str) and entry.get("content_key"):
                key_by_sha.setdefault(entry.get("snapshot_manifest_sha256"), entry["content_key"])
        source_key = key_by_sha.get(source.get("snapshot_manifest_sha256"))
        waves = [
            run for run in gate_waves_since(manifest, gate, from_seq, until_seq)
            if key_by_sha.get(run.get("snapshot_manifest_sha256")) not in (None, source_key)
        ]
        snapshots = {key_by_sha[run.get("snapshot_manifest_sha256")] for run in waves}
        if len(snapshots) < 3:
            return (
                f"{finding.get('finding_id')}: для {gate} MINOR нужны три закрытые волны {gate} на трёх снимках с изменённым "
                f"текстом работы (черновики или final/vkr.docx) после исходной без новых замечаний уровня E2+ "
                f"(сейчас снимков: {len(snapshots)})"
            )
        later_runs = {run.get("run_id") for run in waves}
        fresh = [
            other.get("finding_id") for other in registry.values()
            if other.get("base_gate") == gate
            and (other.get("source") or {}).get("run_id") in later_runs
            and other.get("evidence_level") in ("E2", "E3", "E4")
        ]
        if fresh:
            return f"{finding.get('finding_id')}: в последующих волнах {gate} есть новые замечания E2+: {', '.join(sorted(fresh))}"
    return None


def replay_findings(root: Path, manifest: Any) -> Tuple["OrderedDict[str, Dict[str, Any]]", List[Dict[str, str]]]:
    """Восстанавливает реестр замечаний из manifest и канонических отчётов.

    ``audit/findings.json`` — только представление этого результата; doctor и
    инструменты пересчитывают реестр сами и не доверяют файлу.
    """
    root = Path(root)
    registry: "OrderedDict[str, Dict[str, Any]]" = OrderedDict()
    problems: List[Dict[str, str]] = []
    if not manifest_is_native(manifest):
        return registry, problems
    events: List[Tuple[int, int, str, Tuple[Any, ...]]] = []
    for run, task, attempt in iter_attempts(manifest):
        seq = safe_int(attempt.get("seq"))
        if seq is None:
            problems.append(_problem("AUDIT_ATTEMPT_INVALID", "ERROR", "audit/manifest.json", f"{task.get('task_id')}: у попытки нет seq"))
            continue
        events.append((seq, 0, "attempt", (run, task, attempt)))
    for decision in as_list(manifest.get("decisions")):
        seq = safe_int(decision.get("seq")) if isinstance(decision, dict) else None
        if seq is None:
            problems.append(_problem("AUDIT_DECISION_INVALID", "ERROR", "audit/manifest.json", "Решение без seq"))
            continue
        events.append((seq, 1, "decision", (decision,)))
    events.sort(key=lambda item: (item[0], item[1]))

    for seq, _order, event_type, payload in events:
        if event_type == "attempt":
            run, task, attempt = payload
            report, error = load_attempt_report(root, run, task, attempt)
            if error:
                problems.append(_problem("AUDIT_REPORT_MISMATCH", "ERROR", str(attempt.get("report_path") or "audit/reports"), error))
                continue
            assert report is not None
            for item in as_list(report.get("findings")):
                if not isinstance(item, dict):
                    continue
                finding_id = str(item.get("finding_id") or "")
                if not finding_id:
                    problems.append(_problem("AUDIT_REPORT_MISMATCH", "ERROR", str(attempt.get("report_path")), "замечание без finding_id"))
                    continue
                if finding_id in registry:
                    problems.append(_problem("AUDIT_FINDING_DUPLICATE", "ERROR", str(attempt.get("report_path")), f"{finding_id} уже зарегистрирован"))
                    continue
                record = {key: item.get(key) for key in (
                    "finding_id", "severity", "location", "criterion", "rule_source", "observed", "expected",
                    "evidence_level", "evidence", "proposed_fix", "fix_class", "needs_user_fact",
                )}
                record.update(
                    {
                        "base_gate": task.get("base_gate"),
                        "auditor_role": task.get("auditor_role"),
                        "audit_lens": task.get("audit_lens"),
                        "status": "open",
                        "source": {
                            "run_id": run.get("run_id"),
                            "run_planned_seq": run.get("planned_seq"),
                            "task_id": task.get("task_id"),
                            "kind": task.get("kind"),
                            "auditor_id": attempt.get("auditor_id"),
                            "execution_id": attempt.get("execution_id"),
                            "snapshot_manifest_sha256": run.get("snapshot_manifest_sha256"),
                            "snapshot_seq": run.get("snapshot_seq"),
                            "artifact_sha256": run.get("artifact_sha256"),
                            "report_path": attempt.get("report_path"),
                            "recorded_at": attempt.get("updated_at"),
                            "seq": seq,
                        },
                        "history": [],
                        "resolution": None,
                    }
                )
                registry[finding_id] = record
            for recheck in as_list(report.get("rechecks")):
                if not isinstance(recheck, dict):
                    continue
                finding = registry.get(str(recheck.get("finding_id") or ""))
                if finding is None:
                    problems.append(_problem("AUDIT_RECHECK_UNKNOWN", "ERROR", str(attempt.get("report_path")), f"перепроверка неизвестного {recheck.get('finding_id')}"))
                    continue
                _apply_recheck(finding, recheck, run, task, attempt, seq)
        else:
            (decision,) = payload
            _apply_decision(root, decision, registry, manifest, seq, problems)
    return registry, problems


def fatal_replay_problems(problems: List[Dict[str, str]]) -> List[Dict[str, str]]:
    """Проблемы реестра, при которых инструмент не может продолжать (подделка отчётов)."""
    return [item for item in problems if item.get("code") not in TOLERATED_REPLAY_CODES]


def _apply_recheck(finding: Dict[str, Any], recheck: Dict[str, Any], run: Dict[str, Any], task: Dict[str, Any], attempt: Dict[str, Any], seq: int) -> None:
    verdict = normalize_status(recheck.get("verdict"))
    snapshot_sha = run.get("snapshot_manifest_sha256")
    entry = {
        "seq": seq,
        "at": attempt.get("updated_at"),
        "action": "recheck",
        "verdict": verdict,
        "run_id": run.get("run_id"),
        "task_id": task.get("task_id"),
        "auditor_id": attempt.get("auditor_id"),
        "snapshot_manifest_sha256": snapshot_sha,
        "evidence_level": recheck.get("evidence_level"),
        "evidence": recheck.get("evidence"),
    }
    source = finding.get("source") or {}
    negative_here = any(
        item.get("action") == "recheck" and item.get("verdict") in NEGATIVE_VERDICTS and item.get("snapshot_manifest_sha256") == snapshot_sha
        for item in finding["history"]
    )
    resolution = {
        "run_id": run.get("run_id"),
        "task_id": task.get("task_id"),
        "auditor_id": attempt.get("auditor_id"),
        "snapshot_manifest_sha256": snapshot_sha,
        "at": attempt.get("updated_at"),
        "seq": seq,
    }
    if verdict == "pass":
        if finding["status"] == "open" and negative_here:
            entry["note"] = "на этом снимке уже есть отрицательный вердикт: pass не закрывает замечание (конфликт вердиктов)"
        elif finding["status"] == "open" and snapshot_is_newer(run, source):
            finding["status"] = "resolved"
            finding["resolution"] = dict(resolution, verdict="pass")
        elif finding["status"] == "open":
            entry["note"] = "снимок не новее исходного: pass не закрывает замечание"
    elif verdict == "invalid_finding":
        independent = auditor_key(attempt.get("auditor_id")) != auditor_key(source.get("auditor_id"))
        strong = str(recheck.get("evidence_level") or "").upper() in ("E3", "E4")
        if finding["status"] == "open" and negative_here:
            entry["note"] = "на этом снимке уже есть отрицательный вердикт: invalid_finding не закрывает замечание (конфликт вердиктов)"
        elif finding["status"] == "open" and independent and strong:
            finding["status"] = "rejected"
            finding["resolution"] = dict(resolution, verdict="invalid_finding", evidence_level=recheck.get("evidence_level"))
        elif finding["status"] == "open":
            entry["note"] = "invalid_finding требует E3/E4 от другого аудитора"
    elif verdict in NEGATIVE_VERDICTS and finding["status"] in ("resolved", "rejected"):
        finding["status"] = "open"
        finding["resolution"] = None
        entry["note"] = "замечание открыто повторно"
    finding["history"].append(entry)


def _apply_decision(root: Path, decision: Dict[str, Any], registry: Dict[str, Dict[str, Any]], manifest: Dict[str, Any], seq: int, problems: List[Dict[str, str]]) -> None:
    finding = registry.get(str(decision.get("finding_id") or ""))
    action = decision.get("action")
    if finding is None:
        problems.append(_problem("AUDIT_DECISION_INVALID", "ERROR", "audit/manifest.json", f"решение по неизвестному {decision.get('finding_id')}"))
        return
    entry = {"seq": seq, "at": decision.get("at"), "action": action}
    if action == "accept_risk":
        reason = str(decision.get("reason") or "").strip()
        problem = accept_risk_problem(finding, manifest, registry, seq) if reason else "нет обоснования"
        if problem:
            problems.append(_problem("AUDIT_DECISION_INVALID", "ERROR", "audit/manifest.json", problem))
            return
        finding["status"] = "accepted_risk"
        finding["resolution"] = {"action": "accept_risk", "reason": reason, "at": decision.get("at"), "seq": seq}
        entry["reason"] = reason
    elif action == "waive":
        approval = normalize_logical_id(decision.get("approval"))
        problem = None
        if finding.get("fix_class") != "SUPERVISOR_APPROVAL" or finding.get("status") != "open":
            problem = f"{finding['finding_id']}: waive только для открытого замечания с fix_class SUPERVISOR_APPROVAL"
        else:
            reason = approval_problem(root, approval, str(decision.get("approval_sha256") or ""))
            if reason:
                problem = f"{finding['finding_id']}: {reason}"
        if problem:
            problems.append(_problem("AUDIT_WAIVER_INVALID", "ERROR", approval or "evidence/", problem))
            return
        finding["status"] = "waived"
        finding["resolution"] = {
            "action": "waive", "approval": approval, "approval_sha256": decision.get("approval_sha256"),
            "at": decision.get("at"), "seq": seq,
        }
        entry["approval"] = approval
    else:
        problems.append(_problem("AUDIT_DECISION_INVALID", "ERROR", "audit/manifest.json", f"неизвестное решение {action}"))
        return
    finding["history"].append(entry)


def run_blocking_reasons(run: Dict[str, Any], registry: Dict[str, Dict[str, Any]]) -> Dict[str, List[str]]:
    """Текущие причины, по которым run мешает готовности (по актуальному реестру)."""
    result: Dict[str, List[str]] = {"pending": [], "unclosed": [], "open_high": []}
    status = normalize_status(run.get("status"))
    if status == "invalidated":
        return result
    for task in as_list(run.get("planned_tasks")):
        if not isinstance(task, dict):
            continue
        attempts = [attempt for attempt in as_list(task.get("attempts")) if isinstance(attempt, dict)]
        if not attempts:
            if status != "abandoned" and normalize_status(task.get("status")) != "cancelled":
                result["pending"].append(str(task.get("task_id")))
            continue
        for attempt in attempts:
            attempt_status = normalize_status(attempt.get("status"))
            if attempt_status not in UNCLOSED_ATTEMPT_STATUSES:
                continue
            ids = [str(item) for item in as_list(attempt.get("finding_ids")) + as_list(task.get("recheck_finding_ids"))]
            still_open = [finding_id for finding_id in ids if (registry.get(finding_id) or {}).get("status") == "open"]
            if still_open or not ids:
                result["unclosed"].append(
                    f"{task.get('task_id')}#{attempt.get('attempt')}={attempt_status}"
                    + (f" (открыты {', '.join(still_open)})" if still_open else "")
                )
    result["open_high"] = [
        f"{item['finding_id']} ({item.get('severity')})" for item in registry.values()
        if (item.get("source") or {}).get("run_id") == run.get("run_id")
        and item.get("status") == "open" and item.get("severity") in HIGH_SEVERITIES
    ]
    return result


def findings_view(registry: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
    view = findings_skeleton()
    view["findings"] = list(registry.values())
    view["counts"] = {
        status: sum(1 for item in registry.values() if item.get("status") == status) for status in FINDING_STATUSES
    }
    view["open_high"] = sorted(
        item["finding_id"] for item in registry.values()
        if item.get("status") == "open" and item.get("severity") in HIGH_SEVERITIES
    )
    return view


def views_equal(left: Any, right: Any) -> bool:
    def strip(value: Any) -> Any:
        if isinstance(value, dict):
            return {key: strip(item) for key, item in value.items() if key != "generated_at"}
        if isinstance(value, list):
            return [strip(item) for item in value]
        return value

    return strip(left) == strip(right)


# ---------------------------------------------------------------------------
# Текст DOCX и манифест сборки
# ---------------------------------------------------------------------------

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
_M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
FINGERPRINT_SCHEMA = "vkr-docx-text/2"
INPUT_HASH_SCHEMA = "json-canonical/1"
_TOC_FIELD_CODES = {"TOC"}
_PAGE_FIELD_CODES = {"PAGE", "NUMPAGES", "SECTIONPAGES", "PAGEREF"}
# Результат поля номера страницы: цифры, римские цифры, пробелы и знаки вокруг номера.
_PAGE_RESULT_RE = re.compile(r"^[\s0-9IVXLCDMivxlcdm.,:;()\[\]\-‐‑‒–—―]*$")
_NOTE_SERVICE_TYPES = ("separator", "continuationSeparator", "continuationNotice")
_MEDIA_PREFIXES = ("word/media/", "word/embeddings/")
_HIDDEN_OPEN, _HIDDEN_CLOSE, _HIDDEN_MARK = "⟦", "⟧", "⟦¶⟧"


def _toggle_on(element: Any) -> bool:
    if element is None:
        return False
    return (element.get(_W + "val") or "true").strip().casefold() not in ("0", "false", "off", "none")


def _field_code(instruction: str) -> str:
    return (instruction.strip().split() or [""])[0].upper()


def _alternate_choice(node: Any) -> Any:
    choice = node.find(_MC + "Choice")
    return choice if choice is not None else node.find(_MC + "Fallback")


def _join_parts(parts: List[Tuple[str, bool]]) -> str:
    chunks: List[str] = []
    hidden = False
    for text, is_hidden in parts:
        if is_hidden != hidden:
            chunks.append(_HIDDEN_OPEN if is_hidden else _HIDDEN_CLOSE)
            hidden = is_hidden
        chunks.append(text)
    if hidden:
        chunks.append(_HIDDEN_CLOSE)
    return re.sub(r"\s+", " ", "".join(chunks)).strip()


class _TextCollector:
    """Текст одной XML-части DOCX (тело, колонтитул, сноска, комментарий) для отпечатка.

    Результаты полей TOC пропускаются целиком; результат полей номеров страниц
    (PAGE, NUMPAGES, SECTIONPAGES, PAGEREF) пропускается, только если похож на
    номер. Скрытый текст (w:vanish) отмечается границами, надписи — отдельными
    строками после абзаца-якоря, удалённые правки (w:del) не входят.
    """

    def __init__(self) -> None:
        self.stack: List[Dict[str, Any]] = []

    def _mode(self) -> Tuple[str, Optional[Dict[str, Any]]]:
        page_frame = None
        for frame in self.stack:
            if not frame["separated"]:
                continue
            code = _field_code(frame["instruction"])
            if code in _TOC_FIELD_CODES:
                return "skip", None
            page_frame = frame if code in _PAGE_FIELD_CODES else None
        return ("page", page_frame) if page_frame is not None else ("text", None)

    @staticmethod
    def _flush(frame: Dict[str, Any], parts: List[Tuple[str, bool]]) -> None:
        text = "".join(frame["buffer"])
        if text and not _PAGE_RESULT_RE.match(text):
            parts.append((text, False))
        frame["buffer"] = []

    def block(self, node: Any, lines: List[str]) -> None:
        for child in node:
            tag = child.tag
            if tag == _W + "p":
                self.paragraph(child, lines)
            elif tag == _W + "sdt":
                galleries = [item.get(_W + "val") or "" for item in child.iter(_W + "docPartGallery")]
                if any("tableofcontents" in re.sub(r"\s+", "", value).casefold() for value in galleries):
                    continue
                content = child.find(_W + "sdtContent")
                if content is not None:
                    self.block(content, lines)
            elif tag == _MC + "AlternateContent":
                chosen = _alternate_choice(child)
                if chosen is not None:
                    self.block(chosen, lines)
            elif tag in (_W + "tbl", _W + "tr", _W + "tc", _W + "customXml", _W + "ins", _W + "moveTo", _W + "txbxContent"):
                self.block(child, lines)

    def paragraph(self, node: Any, lines: List[str]) -> None:
        parts: List[Tuple[str, bool]] = []
        extra: List[str] = []
        self.inline(node, parts, extra)
        for frame in self.stack:
            if frame["buffer"]:
                self._flush(frame, parts)
        text = _join_parts(parts)
        properties = node.find(_W + "pPr")
        mark = properties.find(_W + "rPr") if properties is not None else None
        if text and mark is not None and _toggle_on(mark.find(_W + "vanish")):
            text += " " + _HIDDEN_MARK
        if text:
            lines.append(text)
        lines.extend(extra)

    def inline(self, node: Any, parts: List[Tuple[str, bool]], extra: List[str]) -> None:
        for child in node:
            tag = child.tag
            if tag == _W + "r":
                self.run(child, parts, extra)
            elif tag == _W + "fldSimple":
                code = _field_code(child.get(_W + "instr") or "")
                if code in _TOC_FIELD_CODES:
                    continue
                if code in _PAGE_FIELD_CODES:
                    buffer: List[Tuple[str, bool]] = []
                    self.inline(child, buffer, extra)
                    self._flush({"buffer": [text for text, _hidden in buffer]}, parts)
                    continue
                self.inline(child, parts, extra)
            elif tag == _W + "sdt":
                content = child.find(_W + "sdtContent")
                if content is not None:
                    self.inline(content, parts, extra)
            elif tag == _MC + "AlternateContent":
                chosen = _alternate_choice(child)
                if chosen is not None:
                    self.inline(chosen, parts, extra)
            elif tag in (_M + "oMath", _M + "oMathPara"):
                parts.append(("".join(item.text or "" for item in child.iter(_M + "t")), False))
            elif tag in (_W + "hyperlink", _W + "smartTag", _W + "ins", _W + "moveTo", _W + "customXml", _W + "dir", _W + "bdo"):
                self.inline(child, parts, extra)

    def run(self, run: Any, parts: List[Tuple[str, bool]], extra: List[str]) -> None:
        properties = run.find(_W + "rPr")
        hidden = properties is not None and _toggle_on(properties.find(_W + "vanish"))
        for item in run:
            kind = item.tag
            if kind == _W + "fldChar":
                char_type = item.get(_W + "fldCharType")
                if char_type == "begin":
                    self.stack.append({"instruction": "", "separated": False, "buffer": []})
                elif char_type == "separate" and self.stack:
                    self.stack[-1]["separated"] = True
                elif char_type == "end" and self.stack:
                    self._flush(self.stack.pop(), parts)
                continue
            if kind == _W + "instrText":
                if self.stack and not self.stack[-1]["separated"]:
                    self.stack[-1]["instruction"] += item.text or ""
                continue
            if kind in (_W + "drawing", _W + "pict", _W + "object", _MC + "AlternateContent"):
                self.embedded(item, extra)
                continue
            text: Optional[str] = None
            if kind == _W + "t":
                text = item.text or ""
            elif kind in (_W + "tab", _W + "br", _W + "cr", _W + "ptab"):
                text = " "
            elif kind == _W + "noBreakHyphen":
                text = "-"
            elif kind == _W + "sym":
                try:
                    text = chr(int(item.get(_W + "char") or "", 16))
                except ValueError:
                    text = "?"
            elif kind in (_W + "footnoteReference", _W + "endnoteReference"):
                text = "[^]"
            if text is None:
                continue
            mode, frame = self._mode()
            if mode == "skip":
                continue
            if mode == "page" and frame is not None:
                frame["buffer"].append(text)
                continue
            parts.append((text, hidden))

    def embedded(self, node: Any, extra: List[str]) -> None:
        for child in node:
            if child.tag == _MC + "Fallback":
                continue
            if child.tag == _W + "txbxContent":
                lines: List[str] = []
                _TextCollector().block(child, lines)
                extra.extend("textbox: " + line for line in lines)
                continue
            self.embedded(child, extra)


def _part_relationships(archive: zipfile.ZipFile, part_name: str) -> Dict[str, Tuple[str, str]]:
    """rId → (тип связи, имя части в пакете); у внешних связей имя пустое."""
    directory, base = posixpath.split(part_name)
    try:
        root = ET.fromstring(archive.read(posixpath.join(directory, "_rels", base + ".rels")))
    except KeyError:
        return {}
    result: Dict[str, Tuple[str, str]] = {}
    for relation in root.iter(_PKG_REL + "Relationship"):
        target = relation.get("Target") or ""
        if (relation.get("TargetMode") or "").casefold() == "external":
            resolved = ""
        elif target.startswith("/"):
            resolved = target.lstrip("/")
        else:
            resolved = posixpath.normpath(posixpath.join(directory, target))
        result[relation.get("Id") or ""] = (relation.get("Type") or "", resolved)
    return result


def docx_text_lines(path: Path) -> List[str]:
    """Нормализованный сдаваемый текст DOCX для отпечатка (схема ``FINGERPRINT_SCHEMA``).

    Входят: тело (абзацы, таблицы, надписи, скрытый текст с отметкой, формулы),
    колонтитулы, сноски, концевые сноски, комментарии и SHA-256 медиа-частей, на
    которые ссылаются эти части. Не входят: оглавление (результат поля TOC,
    sdt-оглавление), номера страниц в результатах PAGE/NUMPAGES/SECTIONPAGES/
    PAGEREF и удалённые правки. Отпечаток не меняется после обновления полей в
    Word и очистки метаданных.
    """
    lines: List[str] = []
    media: set = set()
    with zipfile.ZipFile(str(path)) as archive:
        names = set(archive.namelist())

        def parse(part_name: str) -> Any:
            return ET.fromstring(archive.read(part_name))

        def note_media(part_name: str, root: Any) -> None:
            part_relations = _part_relationships(archive, part_name)
            for element in root.iter():
                for key, value in element.attrib.items():
                    if key.startswith(_R):
                        target = part_relations.get(value, ("", ""))[1]
                        if target.startswith(_MEDIA_PREFIXES) and target in names:
                            media.add(target)

        document_name = "word/document.xml"
        document = parse(document_name)
        relations = _part_relationships(archive, document_name)
        body = document.find(_W + "body")
        if body is not None:
            _TextCollector().block(body, lines)
        note_media(document_name, document)

        seen: List[str] = []
        for section in document.iter(_W + "sectPr"):
            references = sorted(
                ("header" if item.tag == _W + "headerReference" else "footer", item.get(_W + "type") or "default", item.get(_R + "id") or "")
                for item in section
                if item.tag in (_W + "headerReference", _W + "footerReference")
            )
            for kind, ref_type, rid in references:
                target = relations.get(rid, ("", ""))[1]
                if not target or target in seen or target not in names:
                    continue
                seen.append(target)
                part = parse(target)
                part_lines: List[str] = []
                _TextCollector().block(part, part_lines)
                lines.extend(f"{kind}/{ref_type}: {line}" for line in part_lines)
                note_media(target, part)

        for suffix, element_name in (("/footnotes", "footnote"), ("/endnotes", "endnote"), ("/comments", "comment")):
            for target in sorted({item[1] for item in relations.values() if item[0].endswith(suffix) and item[1] in names}):
                part = parse(target)
                for note in part.findall(_W + element_name):
                    if (note.get(_W + "type") or "normal") in _NOTE_SERVICE_TYPES:
                        continue
                    note_lines: List[str] = []
                    _TextCollector().block(note, note_lines)
                    lines.extend(f"{element_name}: {line}" for line in note_lines)
                note_media(target, part)

        lines.extend("media: " + digest for digest in sorted(sha256_bytes(archive.read(name)) for name in media))
    return lines


def docx_text_fingerprint(path: Path) -> Tuple[Optional[str], Optional[str]]:
    """SHA-256 сдаваемого текста DOCX (``docx_text_lines``); не меняется после обновления полей в Word и очистки метаданных."""
    try:
        lines = docx_text_lines(Path(path))
    except (OSError, KeyError, zipfile.BadZipFile, ET.ParseError, ValueError) as error:
        return None, f"{type(error).__name__}: {error}"
    return sha256_bytes("\n".join(lines).encode("utf-8")), None


# ---------------------------------------------------------------------------
# Число страниц из самого DOCX (docProps/app.xml)
# ---------------------------------------------------------------------------

_EP = "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}"
APP_PROPS_PART = "docProps/app.xml"
# Насколько статистика app.xml может расходиться с текстом документа и всё ещё считаться свежей:
# Word считает и оглавление с номерами страниц (в отпечаток текста они не входят), поэтому ~7 % — норма.
_APP_STATS_MIN_RATIO = 0.5
_APP_STATS_MAX_RATIO = 2.0


def _app_int(root: Any, local: str) -> Optional[int]:
    element = root.find(_EP + local)
    if element is None:
        return None
    try:
        return int((element.text or "").strip())
    except ValueError:
        return None


def docx_app_pages(path: Path) -> Tuple[Optional[int], Optional[str]]:
    """(число страниц, причина недоверия) из ``docProps/app.xml`` сдаваемого DOCX.

    ``<Pages>`` записывает Word при сохранении файла; очистка метаданных его не
    трогает. Значение принимается только тогда, когда статистику записал Word
    **для этого текста**: сборка python-docx переносит ``docProps/app.xml`` из
    шаблона (``<Pages>1</Pages>``, ``<Words>0</Words>``), поэтому свежесть
    статистики проверяется по объёму текста — расхождение больше чем в два раза
    означает, что числа остались от другого документа. Недоверие — не ошибка:
    вызывающий считает, что надёжного числа страниц нет.
    """
    path = Path(path)
    try:
        with zipfile.ZipFile(str(path)) as archive:
            try:
                data = archive.read(APP_PROPS_PART)
            except KeyError:
                return None, f"в DOCX нет {APP_PROPS_PART} (число страниц пишет туда Word при сохранении)"
        root = ET.fromstring(data)
    except (OSError, zipfile.BadZipFile, ET.ParseError, ValueError) as error:
        return None, f"не удалось прочитать {APP_PROPS_PART}: {type(error).__name__}: {error}"
    application = root.find(_EP + "Application")
    application_text = (application.text or "").strip() if application is not None else ""
    if "word" not in application_text.casefold():
        return None, (f"{APP_PROPS_PART}: статистику записал не Word (Application «{application_text}») — "
                      "число страниц в DOCX не проверить")
    pages = _app_int(root, "Pages")
    if pages is None or pages <= 0:
        return None, f"{APP_PROPS_PART}: нет числа страниц (<Pages>) — открой файл в Word и сохрани"
    characters = _app_int(root, "CharactersWithSpaces")
    if characters is None:
        characters = _app_int(root, "Characters")
    try:
        actual = len("\n".join(docx_text_lines(path)))
    except (OSError, KeyError, zipfile.BadZipFile, ET.ParseError, ValueError) as error:
        return None, f"не удалось прочитать текст DOCX: {type(error).__name__}: {error}"
    if characters is None:
        return None, f"{APP_PROPS_PART}: нет объёма текста (<Characters>), свежесть <Pages> не проверить"
    if actual and not (_APP_STATS_MIN_RATIO * actual <= characters <= _APP_STATS_MAX_RATIO * actual):
        return None, (f"{APP_PROPS_PART}: статистика ({characters} знаков) не соответствует тексту ({actual} знаков) — "
                      "файл собран заново после Word, число страниц в нём осталось от шаблона")
    return pages, None


def build_input_paths(root: Path, figure_paths: Iterable[str]) -> List[str]:
    """Входы сборки DOCX: черновики, реестр источников, конфигурация и рисунки."""
    root = Path(root).resolve()
    paths: List[str] = []
    drafts = root / "drafts"
    if drafts.is_dir():
        paths.extend(logical_id_for(root, path) for path in sorted(drafts.glob("*.md")) if path.is_file())
    for relative in ("sources.json", "vkr-project.json"):
        if (root / relative).is_file():
            paths.append(relative)
    for figure in sorted({normalize_logical_id(item) for item in figure_paths if item}):
        resolved = resolve_inside(root, figure)
        if resolved is not None and resolved.is_file():
            paths.append(logical_id_for(root, resolved))
    seen: set = set()
    unique: List[str] = []
    for item in paths:
        key = dedupe_key(item)
        if key not in seen:
            seen.add(key)
            unique.append(item)
    return unique


def build_manifest_payload(root: Path, docx_path: Path, figure_paths: Iterable[str], build_version: str) -> Dict[str, Any]:
    root = Path(root).resolve()
    fingerprint, error = docx_text_fingerprint(docx_path)
    if error:
        raise ProjectDataError(str(docx_path), f"не удалось прочитать собранный DOCX: {error}", "IO_ERROR")
    try:
        docx_logical = logical_id_for(root.resolve(), Path(docx_path).resolve())
    except ValueError:
        docx_logical = str(docx_path)
    return {
        "schema": "vkr-build-manifest",
        "fingerprint_schema": FINGERPRINT_SCHEMA,
        "input_hash": INPUT_HASH_SCHEMA,
        "build_version": build_version,
        "built_at": utc_now_iso(),
        "inputs": [
            {"path": relative, "sha256": content_sha256(root / relative)}
            for relative in build_input_paths(root, figure_paths)
        ],
        "docx": {
            "path": docx_logical,
            "sha256": sha256_file(docx_path),
            "text_fingerprint": fingerprint,
        },
    }


def build_manifest_problems(root: Path) -> List[str]:
    """Почему final/vkr.docx не соответствует черновикам (пусто — соответствует)."""
    root = Path(root).resolve()
    path = root / BUILD_MANIFEST_PATH
    if not path.is_file():
        return ["нет exports/build-manifest.json: final/vkr.docx собран не build_vkr.py 6.33"]
    try:
        manifest = read_json(path, label=BUILD_MANIFEST_PATH)
    except ProjectDataError as error:
        return [str(error)]
    if not isinstance(manifest, dict) or not isinstance(manifest.get("inputs"), list) or not isinstance(manifest.get("docx"), dict):
        return ["exports/build-manifest.json повреждён"]
    if manifest.get("fingerprint_schema") != FINGERPRINT_SCHEMA or manifest.get("input_hash") != INPUT_HASH_SCHEMA:
        return ["exports/build-manifest.json записан прежней версией build_vkr.py (другой отпечаток текста): пересобери DOCX"]
    problems: List[str] = []
    recorded: set = set()
    for item in manifest["inputs"]:
        if not isinstance(item, dict):
            problems.append("запись входа сборки повреждена")
            continue
        logical = normalize_logical_id(item.get("path"))
        recorded.add(dedupe_key(logical))
        resolved = resolve_inside(root, logical)
        if resolved is None or not resolved.is_file():
            problems.append(f"{logical} удалён после сборки")
        elif content_sha256(resolved) != item.get("sha256"):
            problems.append(f"{logical} изменён после сборки")
    for relative in build_input_paths(root, []):
        if dedupe_key(relative) not in recorded:
            problems.append(f"{relative} появился после сборки")
    docx = manifest["docx"]
    if normalize_logical_id(docx.get("path")) != FINAL_DOCX_PATH:
        problems.append("последняя сборка записана не в final/vkr.docx")
    final = root / FINAL_DOCX_PATH
    if not final.is_file():
        problems.append("нет final/vkr.docx")
    else:
        fingerprint, error = docx_text_fingerprint(final)
        if error:
            problems.append(f"не удалось прочитать текст final/vkr.docx: {error}")
        elif fingerprint != docx.get("text_fingerprint"):
            problems.append("текст final/vkr.docx не совпадает с последней сборкой из черновиков")
    return problems


# ---------------------------------------------------------------------------
# Валидатор
# ---------------------------------------------------------------------------

_VALIDATOR_MODULE: Any = None


def _load_validator_module() -> Any:
    global _VALIDATOR_MODULE
    if _VALIDATOR_MODULE is not None:
        return _VALIDATOR_MODULE
    path = SCRIPTS_DIR / "validate_vkr.py"
    if not path.is_file():
        raise ImportError(f"validate_vkr.py не найден: {path}")
    spec = importlib.util.spec_from_file_location("vkr_validate_vkr_runtime", str(path))
    if spec is None or spec.loader is None:
        raise ImportError("не удалось загрузить validate_vkr.py")
    module = importlib.util.module_from_spec(spec)
    previous = sys.dont_write_bytecode
    sys.dont_write_bytecode = True
    try:
        spec.loader.exec_module(module)
    except SystemExit as error:  # старый валидатор завершал процесс без python-docx
        raise ImportError(f"validate_vkr.py завершился при импорте (код {error.code})") from error
    finally:
        sys.dont_write_bytecode = previous
    _VALIDATOR_MODULE = module
    return module


_INTEGRITY_MODULE: Any = None


def docx_package_problems(path: Path) -> Optional[List[str]]:
    """Структурные проблемы пакета DOCX, из-за которых Word не откроет файл (``docx_integrity.py``).

    Пустой список — проблем нет; None — модуль проверки не установлен.
    """
    global _INTEGRITY_MODULE
    if _INTEGRITY_MODULE is None:
        module_path = SCRIPTS_DIR / "docx_integrity.py"
        if not module_path.is_file():
            return None
        spec = importlib.util.spec_from_file_location("vkr_docx_integrity_runtime", str(module_path))
        if spec is None or spec.loader is None:
            return None
        module = importlib.util.module_from_spec(spec)
        previous = sys.dont_write_bytecode
        sys.dont_write_bytecode = True
        try:
            spec.loader.exec_module(module)
        finally:
            sys.dont_write_bytecode = previous
        _INTEGRITY_MODULE = module
    return [str(item) for item in _INTEGRITY_MODULE.docx_package_problems(str(path))]


def run_validator(docx_path: Path, profile: str) -> Dict[str, Any]:
    """Запускает ``validate_vkr.validate_vkr`` в текущем процессе.

    Возвращает ``{"available": bool, "report": dict | None, "error": str | None}``.
    ``available=False`` означает, что проверку выполнить нельзя (нет python-docx
    или валидатора): это не PASS.
    """
    if importlib.util.find_spec("docx") is None:
        return {"available": False, "report": None, "error": "python-docx не установлен"}
    try:
        module = _load_validator_module()
    except ImportError as error:
        return {"available": False, "report": None, "error": str(error)}
    function = getattr(module, "validate_vkr", None)
    if not callable(function):
        return {"available": False, "report": None, "error": "в validate_vkr.py нет функции validate_vkr"}
    try:
        report = function(str(docx_path), False, profile)
    except Exception as error:  # отчёт об ошибке вместо трассировки
        return {"available": True, "report": {"status": "error", "error": f"{type(error).__name__}: {error}"}, "error": None}
    if not isinstance(report, dict):
        return {"available": True, "report": {"status": "error", "error": "валидатор вернул не объект"}, "error": None}
    return {"available": True, "report": report, "error": None}


def validator_error_messages(report: Any, limit: int = 5) -> List[str]:
    messages: List[str] = []
    checks = report.get("checks") if isinstance(report, dict) else None
    if isinstance(checks, dict):
        for group, results in checks.items():
            for result in as_list(results):
                if isinstance(result, dict) and not result.get("ok") and str(result.get("severity") or "").casefold() == "error":
                    messages.append(f"{group}: {result.get('message')}")
                    if len(messages) >= limit:
                        return messages
    return messages


def validator_error_count(report: Any) -> Optional[int]:
    summary = report.get("summary") if isinstance(report, dict) else None
    errors = summary.get("errors") if isinstance(summary, dict) else None
    return safe_int(errors)


def validator_warning_messages(report: Any, limit: int = 5) -> List[str]:
    """Первые предупреждения валидатора: они не блокируют gate, но агент обязан их прочитать."""
    messages: List[str] = []
    checks = report.get("checks") if isinstance(report, dict) else None
    if isinstance(checks, dict):
        for group, results in checks.items():
            for result in as_list(results):
                if isinstance(result, dict) and not result.get("ok") and str(result.get("severity") or "").casefold() == "warning":
                    messages.append(f"{group}: {result.get('message')}")
                    if len(messages) >= limit:
                        return messages
    return messages


def validator_warning_count(report: Any) -> Optional[int]:
    summary = report.get("summary") if isinstance(report, dict) else None
    warnings = summary.get("warnings") if isinstance(summary, dict) else None
    return safe_int(warnings)


# ---------------------------------------------------------------------------
# Блокировка каталога аудита
# ---------------------------------------------------------------------------


def pid_alive(pid: int) -> bool:
    """Жив ли процесс с таким PID на этом компьютере (без сигналов процессу)."""
    if pid <= 0:
        return False
    if os.name == "nt":
        import ctypes
        from ctypes import wintypes

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.OpenProcess.restype = wintypes.HANDLE
        kernel32.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
        kernel32.GetExitCodeProcess.argtypes = (wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD))
        kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
        handle = kernel32.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
        if not handle:
            return ctypes.get_last_error() == 5  # ERROR_ACCESS_DENIED: процесс есть, но чужой
        try:
            code = wintypes.DWORD()
            if kernel32.GetExitCodeProcess(handle, ctypes.byref(code)):
                return code.value == 259  # STILL_ACTIVE
            return True
        finally:
            kernel32.CloseHandle(handle)
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return False
    return True


class AuditLock:
    """Межпроцессная блокировка ``audit/.vkr_audit.lock``.

    Держится только на время чтения и записи служебных файлов; хеширование
    входов выполняется до захвата. Блокировка процесса, которого уже нет,
    снимается сразу; ожидание живого владельца — до ``timeout`` секунд.
    """

    def __init__(self, root: Path, timeout: float = 120.0, foreign_stale_after: float = 900.0) -> None:
        self.path = Path(root) / "audit" / ".vkr_audit.lock"
        self.timeout = timeout
        self.foreign_stale_after = foreign_stale_after
        self.handle: Optional[int] = None

    def _stale(self) -> bool:
        import socket
        import time

        try:
            raw = self.path.read_text(encoding="utf-8", errors="replace")
            age = time.time() - self.path.stat().st_mtime
        except OSError:
            return False
        try:
            parsed: Any = json.loads(raw)
        except ValueError:
            parsed = None
        if isinstance(parsed, dict):
            owner: Dict[str, Any] = parsed
        else:
            # Прежний формат — только PID (строка «12345» разбирается и как JSON-число).
            digits = re.findall(r"\d+", raw)
            owner = {"pid": int(digits[0])} if digits else {}
        pid = safe_int(owner.get("pid"))
        host = owner.get("host")
        if pid is None:
            return age > 5.0
        if host and host != socket.gethostname():
            return age > self.foreign_stale_after
        return not pid_alive(pid)

    def __enter__(self) -> "AuditLock":
        import socket
        import time

        self.path.parent.mkdir(parents=True, exist_ok=True)
        deadline = time.monotonic() + self.timeout
        payload = json.dumps({"pid": os.getpid(), "host": socket.gethostname(), "at": utc_now_iso()}).encode("utf-8")
        while True:
            try:
                self.handle = os.open(str(self.path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.write(self.handle, payload)
                return self
            except FileExistsError:
                if self._stale():
                    try:
                        self.path.unlink()
                    except OSError:
                        pass
                    continue
                if time.monotonic() > deadline:
                    raise ProjectDataError(
                        "audit/.vkr_audit.lock",
                        f"каталог аудита занят другим процессом дольше {int(self.timeout)} с; повтори команду позже",
                        "AUDIT_LOCKED",
                    )
                time.sleep(0.1)

    def __exit__(self, *exc: Any) -> None:
        if self.handle is not None:
            os.close(self.handle)
            self.handle = None
        try:
            self.path.unlink()
        except OSError:
            pass
