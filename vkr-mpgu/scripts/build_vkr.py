#!/usr/bin/env python3
"""Сборка ВКР из черновиков: drafts/*.md + sources.json → final/vkr.docx (vkr-mpgu 6.33).

    python build_vkr.py <PROJECT_DIR> [--json] [-o REPORT] [--output final/vkr.docx]
                        [--content-only] [--group-normative-first]

Читает ``vkr-project.json`` (topic, title_page, vkr_type, profile), черновики
``drafts/annotation.md``, ``introduction.md``, ``chapter-N.md``, ``conclusion.md``,
``appendix-N.md`` (формат — ``references/drafts-format.md``) и реестр
``sources.json``. Пишет ``final/vkr.docx``, а после его успешной записи —
``exports/build-manifest.json``, ``exports/citation-map.json``
({"build_id", "numbers": {"1": id}, "entries": {"1": текст записи}}) и
``exports/vkr-content.json`` (вход рендерера ``create_vkr_docx.py``) с общим
``build_id``; прежний DOCX сохраняется в ``backups/checkpoints/<UTC>-build/``.

Порядок записи: сначала проверка, что ``final/vkr.docx`` и служебные файлы можно
перезаписать (DOCX открыт в Word — код 2 OUTPUT_LOCKED, ничего не изменено);
DOCX пишется во временный файл и атомарно подменяется; служебные файлы
подменяются только после этого. ``--content-only`` — проверка содержания без
DOCX: пишет только ``exports/vkr-content-preview.json``, файлы последней сборки не
трогает. Сборка в другой путь (``--output exports/…``) пишет DOCX и тот же
preview-файл.

Список литературы — все записи реестра, кроме rejected и неструктурированных,
в алфавитном порядке (раздел 3 SPEC); ссылки ``[@id, с. 5]`` и legacy
``[3, с. 17]`` переписываются в ``[N, с. 5]``.

После сборки ``final/vkr.docx`` отчёт (``next_steps``) и итоговое сообщение
называют следующие шаги: ``update_docx_fields.py <PROJECT_DIR>/final/vkr.docx``
(без Word — ручной F9 и сохранение) → ``clean_docx_metadata.py
<PROJECT_DIR>/final/vkr.docx --in-place`` → ``vkr_audit.py <PROJECT_DIR> validate``.

Коды: 0 — собрано (предупреждения допустимы); 1 — ошибки содержания, ничего не
перезаписано; 2 — ошибка использования, ввода-вывода (в том числе DOCX открыт в
Word) или нет python-docx.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import shutil
import sys
import tempfile
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

sys.dont_write_bytecode = True
SCRIPTS_DIR = Path(__file__).resolve().parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import format_bibliography as fb  # noqa: E402
import update_docx_fields as fields_tool  # noqa: E402  (проверка, что DOCX не открыт в Word)
import verify_sources as sources_tool  # noqa: E402  (устаревшая проверка источника: fields_sha256)
import vkr_common as common  # noqa: E402
import vkr_markers  # noqa: E402

BUILD_VERSION = "6.33"
ROOT_BUILD_MANIFEST = common.BUILD_MANIFEST_PATH
SUPPORTED_IMAGE_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"}
CITATION_MAP_SCHEMA = "vkr-citation-map/2"
CONTENT_PATH = "exports/vkr-content.json"
CITATION_MAP_PATH = "exports/citation-map.json"
PREVIEW_CONTENT_PATH = "exports/vkr-content-preview.json"

# ---------------------------------------------------------------------------
# Разбор Markdown-черновика
# ---------------------------------------------------------------------------

FENCE_RE = re.compile(r"^[ \t]{0,3}(`{3,}|~{3,})\s*([^\s`]*)\s*$")
HEADING_RE = re.compile(r"^[ \t]{0,3}(#{1,6})[ \t]+(.*?)[ \t]*#*[ \t]*$")
LIST_RE = re.compile(r"^([ \t]*)([-*+•–—]|\d{1,3}[.)])[ \t]+(.*\S)[ \t]*$")
TABLE_LINE_RE = re.compile(r"^[ \t]*\|.*\|?[ \t]*$")
TABLE_SEP_RE = re.compile(r"^[ \t]*\|?[ \t]*:?-{3,}:?[ \t]*(\|[ \t]*:?-{3,}:?[ \t]*)*\|?[ \t]*$")
FIGURE_RE = re.compile(r"^[ \t]*!\[(?P<caption>[^\]]*)\]\((?P<path>[^)]*)\)[ \t]*$")
TABLE_CAPTION_RE = re.compile(r"^Таблица(?:\s+\d+)?\s*[:.—–-]\s*(?P<title>.+)$", re.IGNORECASE)
LISTING_CAPTION_RE = re.compile(r"^Листинг(?:\s+\d+)?\s*[:.—–-]\s*(?P<title>.+)$", re.IGNORECASE)
SOURCE_LINE_RE = re.compile(r"^Источник\s*:\s*(?P<source>.+)$", re.IGNORECASE)
CONCLUSION_RE = re.compile(r"^выводы\s+по\s+(?:главе|первой\s+главе|второй\s+главе|третьей\s+главе)(?:\s+[IVXLCDM\d]+)?\.?$", re.IGNORECASE)
CHAPTER_FILE_RE = re.compile(r"^chapter-(\d+)\.md$")
APPENDIX_FILE_RE = re.compile(r"^appendix-(\d+)\.md$")
COMMENT_RE = re.compile(r"<!--.*?-->", re.DOTALL)


def _strip_comments(lines: List[str], code_lines: set) -> List[str]:
    """Удаляет HTML-комментарии вне кода, сохраняя число строк."""
    out = list(lines)
    text_positions = [i for i in range(len(lines)) if i not in code_lines]
    joined = "\n".join(lines[i] if i not in code_lines else "\x00" * len(lines[i]) for i in range(len(lines)))

    def blank(match: "re.Match[str]") -> str:
        return "".join("\n" if ch == "\n" else " " for ch in match.group(0))

    cleaned = COMMENT_RE.sub(blank, joined)
    cleaned_lines = cleaned.split("\n")
    for i in text_positions:
        out[i] = cleaned_lines[i]
    return out


def _split_cells(line: str) -> List[str]:
    text = line.strip()
    if text.startswith("|"):
        text = text[1:]
    if text.endswith("|") and not text.endswith("\\|"):
        text = text[:-1]
    cells, current, escape = [], "", False
    for ch in text:
        if escape:
            current += ch if ch == "|" else "\\" + ch
            escape = False
        elif ch == "\\":
            escape = True
        elif ch == "|":
            cells.append(current.strip())
            current = ""
        else:
            current += ch
    if escape:
        current += "\\"
    cells.append(current.strip())
    return cells


def parse_markdown(text: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """Черновик → блоки с номерами строк (start, end — 0-based, end не включается).

    Возвращает (blocks, problems); problems — {"level": "error"|"warning", "line", "code", "message"}.
    Виды блоков: heading (level, text), paragraph (text), list (ordered, items),
    table (title, headers, rows, source), listing (caption, code, language), figure (caption, path).
    """
    problems: List[Dict[str, Any]] = []
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    code_lines = set()
    i = 0
    while i < len(lines):
        match = FENCE_RE.match(lines[i])
        if match:
            fence = match.group(1)
            j = i + 1
            while j < len(lines):
                close = FENCE_RE.match(lines[j])
                if close and close.group(1)[0] == fence[0] and len(close.group(1)) >= len(fence) and not close.group(2):
                    break
                j += 1
            code_lines.update(range(i, min(j + 1, len(lines))))
            i = j + 1
            continue
        i += 1
    clean = _strip_comments(lines, code_lines)

    blocks: List[Dict[str, Any]] = []
    i = 0
    n = len(clean)

    def is_special(idx: int) -> bool:
        line = clean[idx]
        return bool(FENCE_RE.match(line) or HEADING_RE.match(line) or TABLE_LINE_RE.match(line)
                    or FIGURE_RE.match(line) or LIST_RE.match(line))

    while i < n:
        line = clean[i]
        if not line.strip():
            i += 1
            continue
        fence = FENCE_RE.match(line) if i in code_lines else None
        if fence:
            j = i + 1
            code: List[str] = []
            closed = False
            while j < n:
                close = FENCE_RE.match(lines[j])
                if close and close.group(1)[0] == fence.group(1)[0] and len(close.group(1)) >= len(fence.group(1)) and not close.group(2):
                    closed = True
                    break
                code.append(lines[j])
                j += 1
            if not closed:
                problems.append({"level": "error", "line": i + 1, "code": "CODE_BLOCK_UNCLOSED",
                                 "message": "блок кода ``` не закрыт"})
            block = {"kind": "listing", "caption": "", "code": "\n".join(code), "language": fence.group(2),
                     "start": i, "end": min(j + 1, n)}
            prev = blocks[-1] if blocks else None
            if prev and prev["kind"] == "paragraph" and prev.get("single_line"):
                cap = LISTING_CAPTION_RE.match(prev["text"])
                if cap:
                    blocks.pop()
                    block["caption"] = cap.group("title").strip()
                    block["start"] = prev["start"]
            blocks.append(block)
            i = min(j + 1, n)
            continue
        heading = HEADING_RE.match(line)
        if heading:
            blocks.append({"kind": "heading", "level": len(heading.group(1)), "text": heading.group(2).strip(),
                           "start": i, "end": i + 1})
            i += 1
            continue
        figure = FIGURE_RE.match(line)
        if figure:
            path = figure.group("path").strip()
            path = re.sub(r"\s+\"[^\"]*\"$", "", path).strip()
            blocks.append({"kind": "figure", "caption": figure.group("caption").strip(), "path": path,
                           "start": i, "end": i + 1})
            i += 1
            continue
        if TABLE_LINE_RE.match(line) and line.strip().startswith("|"):
            j = i
            rows_raw = []
            while j < n and clean[j].strip().startswith("|"):
                rows_raw.append(clean[j])
                j += 1
            headers: List[str] = []
            body = rows_raw
            if len(rows_raw) >= 2 and TABLE_SEP_RE.match(rows_raw[1]):
                headers = _split_cells(rows_raw[0])
                body = rows_raw[2:]
            elif TABLE_SEP_RE.match(rows_raw[0]):
                body = rows_raw[1:]
            rows = [_split_cells(r) for r in body if not TABLE_SEP_RE.match(r)]
            block = {"kind": "table", "title": "", "headers": headers, "rows": rows, "source": "",
                     "start": i, "end": j}
            prev = blocks[-1] if blocks else None
            if prev and prev["kind"] == "paragraph" and prev.get("single_line"):
                cap = TABLE_CAPTION_RE.match(prev["text"])
                if cap:
                    blocks.pop()
                    block["title"] = cap.group("title").strip()
                    block["start"] = prev["start"]
            k = j
            while k < n and not clean[k].strip() and k - j < 2:
                k += 1
            if k < n:
                src = SOURCE_LINE_RE.match(clean[k].strip())
                if src and (k + 1 >= n or not clean[k + 1].strip() or is_special(k + 1)):
                    block["source"] = src.group("source").strip()
                    block["end"] = k + 1
                    j = k + 1
            if not block["title"]:
                problems.append({"level": "warning", "line": i + 1, "code": "TABLE_NO_CAPTION",
                                 "message": "у таблицы нет строки «Таблица: Название» перед ней"})
            blocks.append(block)
            i = j
            continue
        item = LIST_RE.match(line)
        if item:
            ordered = item.group(2)[0].isdigit()
            number = int(re.match(r"\d+", item.group(2)).group(0)) if ordered else 0
            prev = blocks[-1] if blocks else None
            continue_prev = False
            if prev is not None and prev["kind"] == "list" and prev["ordered"] == ordered:
                if prev["end"] == i:
                    # Пункты на соседних строках — один список; «1.» после «2.»/«3.» начинает новый.
                    continue_prev = not (ordered and number == 1 and prev["last_number"] >= 2)
                else:
                    # Через пустую строку список продолжается, только если номер следующий
                    # (маркированный — всегда); новый нумерованный список начинается с «1.».
                    continue_prev = (not ordered) or number == prev["last_number"] + 1
            if continue_prev:
                block = prev
            else:
                block = {"kind": "list", "ordered": ordered, "items": [], "last_number": 0, "start": i, "end": i + 1}
                blocks.append(block)
            if len(item.group(1).replace("\t", "    ")) >= 2:
                problems.append({"level": "warning", "line": i + 1, "code": "LIST_NESTED",
                                 "message": "вложенные списки не поддерживаются — пункт оформлен на первом уровне"})
            text_parts = [item.group(3).strip()]
            j = i + 1
            while j < n and clean[j].strip() and not is_special(j):
                text_parts.append(clean[j].strip())
                j += 1
            block["items"].append(" ".join(text_parts))
            block["last_number"] = number
            block["end"] = j
            i = j
            continue
        j = i
        para_lines = []
        while j < n and clean[j].strip() and not (j > i and is_special(j)):
            if j == i and is_special(j):
                break
            para_lines.append(clean[j].strip())
            j += 1
        if not para_lines:
            para_lines = [line.strip()]
            j = i + 1
        blocks.append({"kind": "paragraph", "text": " ".join(para_lines), "single_line": len(para_lines) == 1,
                       "start": i, "end": j})
        i = j
    for block in blocks:
        if block["kind"] == "paragraph" and block.get("single_line"):
            if TABLE_CAPTION_RE.match(block["text"]) and block["text"].lower().startswith("таблица") and ":" in block["text"][:12]:
                problems.append({"level": "warning", "line": block["start"] + 1, "code": "CAPTION_WITHOUT_TABLE",
                                 "message": "строка «Таблица: …» не стоит перед таблицей"})
            if LISTING_CAPTION_RE.match(block["text"]) and ":" in block["text"][:12]:
                problems.append({"level": "warning", "line": block["start"] + 1, "code": "CAPTION_WITHOUT_LISTING",
                                 "message": "строка «Листинг: …» не стоит перед блоком кода"})
    return blocks, problems


def block_to_content(block: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Блок черновика → блок входа рендерера."""
    kind = block["kind"]
    if kind == "paragraph":
        return {"type": "text", "content": block["text"]}
    if kind == "heading":
        return {"type": "text", "content": block["text"]}
    if kind == "list":
        return {"type": "list", "ordered": block["ordered"], "items": list(block["items"])}
    if kind == "table":
        out = {"type": "table", "title": block["title"], "headers": block["headers"], "rows": block["rows"]}
        if block.get("source"):
            out["source"] = block["source"]
        return out
    if kind == "listing":
        out = {"type": "listing", "caption": block["caption"], "code": block["code"]}
        if block.get("language"):
            out["language"] = block["language"]
        return out
    if kind == "figure":
        out = {"type": "figure", "caption": block["caption"]}
        if block["path"]:
            out["path"] = block["path"]
        else:
            out["placeholder"] = True
        return out
    return None


def block_texts(block: Dict[str, Any]) -> List[str]:
    """Тексты блока, в которых ищутся ссылки (код листинга не входит)."""
    kind = block.get("kind") or block.get("type")
    if kind in ("paragraph", "text", "heading"):
        return [block.get("text", block.get("content", ""))]
    if kind == "list":
        return list(block.get("items", []))
    if kind == "table":
        texts = [block.get("title", "")] + list(block.get("headers", []))
        for row in block.get("rows", []):
            texts.extend(row)
        texts.append(block.get("source", ""))
        return texts
    if kind in ("listing", "figure"):
        return [block.get("caption", "")]
    return []


# ---------------------------------------------------------------------------
# Проект
# ---------------------------------------------------------------------------

class Report:
    def __init__(self) -> None:
        self.errors: List[Dict[str, Any]] = []
        self.warnings: List[Dict[str, Any]] = []

    def add(self, level: str, code: str, message: str, file: str = "", line: Optional[int] = None) -> None:
        entry: Dict[str, Any] = {"code": code, "message": message}
        if file:
            entry["file"] = file
        if line:
            entry["line"] = line
        (self.errors if level == "error" else self.warnings).append(entry)


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def write_json_atomic(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=str(path.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as handle:
            json.dump(value, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        os.replace(tmp, str(path))
    except BaseException:
        if os.path.exists(tmp):
            os.unlink(tmp)
        raise


def draft_files(drafts: Path) -> Dict[str, Any]:
    """Файлы черновиков в порядке сборки."""
    chapters: Dict[int, Path] = {}
    appendices: Dict[int, Path] = {}
    other: List[Path] = []
    for path in sorted(drafts.glob("*.md")) if drafts.is_dir() else []:
        name = path.name
        if name in ("annotation.md", "introduction.md", "conclusion.md"):
            continue
        m = CHAPTER_FILE_RE.match(name)
        if m:
            chapters[int(m.group(1))] = path
            continue
        m = APPENDIX_FILE_RE.match(name)
        if m:
            appendices[int(m.group(1))] = path
            continue
        other.append(path)
    return {"chapters": dict(sorted(chapters.items())), "appendices": dict(sorted(appendices.items())), "other": other}


def significant(blocks: List[Dict[str, Any]]) -> bool:
    return any(b["kind"] != "heading" for b in blocks)


def split_chapter(blocks: List[Dict[str, Any]], rel: str, number: int, report: Report) -> Dict[str, Any]:
    chapter = {"title": "", "blocks": [], "paragraphs": [], "conclusion": []}
    target = chapter["blocks"]
    seen_h1 = False
    in_conclusion = False
    para_index = 0
    for block in blocks:
        if block["kind"] == "heading" and block["level"] == 1:
            if seen_h1:
                report.add("warning", "HEADING_DUPLICATE", "второй заголовок # в файле главы оформлен как текст", rel, block["start"] + 1)
                target.append(block_to_content(block))
                continue
            seen_h1 = True
            chapter["title"] = fb_strip_chapter(block["text"])
            prefix = re.match(r"^\s*глава\s+([IVXLCDM]+|\d+)", block["text"], re.IGNORECASE)
            if prefix:
                value = prefix.group(1)
                given = int(value) if value.isdigit() else roman_to_int(value.upper())
                if given != number:
                    report.add("warning", "CHAPTER_NUMBER_MISMATCH",
                               f"в заголовке «{block['text']}» номер {value}, а файл chapter-{number}.md — номер берётся из имени файла",
                               rel, block["start"] + 1)
            continue
        if block["kind"] == "heading" and block["level"] == 2:
            title = block["text"].strip()
            if CONCLUSION_RE.match(title):
                in_conclusion = True
                target = chapter["conclusion"]
                continue
            if in_conclusion:
                report.add("warning", "PARAGRAPH_AFTER_CONCLUSION",
                           "параграф после «Выводы по главе» — выводы будут напечатаны в конце главы", rel, block["start"] + 1)
                in_conclusion = False
            para_index += 1
            numbered = re.match(r"^\s*(\d+)\.(\d+)\.?\s+", title)
            if numbered and (int(numbered.group(1)), int(numbered.group(2))) != (number, para_index):
                report.add("warning", "PARAGRAPH_NUMBER_MISMATCH",
                           f"в заголовке «{title}» номер не совпадает с порядком — build поставит {number}.{para_index}.",
                           rel, block["start"] + 1)
            paragraph = {"title": strip_para_number(title), "blocks": []}
            chapter["paragraphs"].append(paragraph)
            target = paragraph["blocks"]
            continue
        if block["kind"] == "heading":
            report.add("warning", "HEADING_LEVEL_UNSUPPORTED",
                       f"заголовок уровня {block['level']} не поддерживается — оформлен как абзац", rel, block["start"] + 1)
        target.append(block_to_content(block))
    if not seen_h1:
        report.add("error", "CHAPTER_TITLE_MISSING", "нет заголовка «# Название главы»", rel)
    return chapter


ROMAN_VALUES = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}


def roman_to_int(value: str) -> int:
    total, prev = 0, 0
    for ch in reversed(value):
        cur = ROMAN_VALUES.get(ch, 0)
        total = total - cur if cur < prev else total + cur
        prev = max(prev, cur)
    return total


def fb_strip_chapter(text: str) -> str:
    return re.sub(r"^\s*глава\s+(?:[IVXLCDM]+|\d+)\s*(?:[.:)]|[-–—])?\s*", "", text, count=1, flags=re.IGNORECASE).strip().rstrip(".").strip()


def strip_para_number(text: str) -> str:
    return re.sub(r"^\s*\d+(?:\.\d+)+\.?\s+|^\s*\d+\.\s+", "", text, count=1).strip().rstrip(".").strip()


def section_blocks(blocks: List[Dict[str, Any]], rel: str, report: Report) -> Tuple[str, List[Dict[str, Any]]]:
    title = ""
    out = []
    for block in blocks:
        if block["kind"] == "heading" and block["level"] == 1 and not title:
            title = block["text"]
            continue
        if block["kind"] == "heading":
            report.add("warning", "HEADING_LEVEL_UNSUPPORTED",
                       f"подзаголовок «{block['text']}» в этом разделе не поддерживается — оформлен как абзац",
                       rel, block["start"] + 1)
        out.append(block_to_content(block))
    return title, out


def sort_citation_numbers(text: str) -> str:
    """Части числовой ссылки — по возрастанию номера: «[11; 9, с. 5]» → «[9, с. 5; 11]».

    Локатор остаётся при своём номере, общий префикс («см.», «Цит. по:») — в начале
    ссылки. Не переставляются ссылки с @id (не разрешены), с префиксами у нескольких
    частей и перечисления через запятую («[3, 5]»).
    """
    citations = fb.find_citations(text)
    if not citations:
        return text
    out: List[str] = []
    pos = 0
    for citation in citations:
        parts = citation["parts"]
        rendered = citation["text"]
        raw_parts = citation["text"][1:-1].split(";")
        simple = (len(parts) > 1 and len(raw_parts) == len(parts)
                  and all(part["kind"] in ("num", "range") and not part.get("comma") for part in parts)
                  and not any(str(part["prefix"]).strip() for part in parts[1:]))
        if simple:
            order = sorted(range(len(parts)), key=lambda index: (
                parts[index]["ref"][0] if parts[index]["kind"] == "range" else parts[index]["ref"], index))
            if order != list(range(len(parts))):
                prefix = str(parts[0]["prefix"])
                bodies = [raw.strip() for raw in raw_parts]
                if prefix:
                    bodies[0] = raw_parts[0].lstrip()[len(prefix):].strip()
                sorted_bodies = [bodies[index] for index in order]
                sorted_bodies[0] = prefix + sorted_bodies[0]
                rendered = "[" + "; ".join(sorted_bodies) + "]"
        out.append(text[pos:citation["start"]])
        out.append(rendered)
        pos = citation["end"]
    out.append(text[pos:])
    return "".join(out)


def verification_complete(source: Any) -> bool:
    """То же правило, что у doctor: method из списка, checked_at в ISO 8601, url или note."""
    verification = source.get("verification") if isinstance(source, dict) else None
    if not isinstance(verification, dict):
        return False
    return (verification.get("method") in common.VERIFICATION_METHODS
            and common.parse_iso8601(verification.get("checked_at")) is not None
            and bool(str(verification.get("url") or "").strip() or str(verification.get("note") or "").strip()))


def collect_project(root: Path, report: Report, group_normative_first: bool = False) -> Optional[Dict[str, Any]]:
    """Читает проект и возвращает {content, citation_map, stats}; ошибки — в report."""
    config_path = root / "vkr-project.json"
    if not config_path.is_file():
        raise UsageError(f"нет файла {config_path} — сначала init_vkr_project.py")
    try:
        config = read_json(config_path)
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise UsageError(f"не удалось прочитать vkr-project.json: {error}") from error
    if not isinstance(config, dict):
        raise UsageError("vkr-project.json должен быть объектом")

    vkr_type = str(config.get("vkr_type") or "").strip()
    profile = str(config.get("profile") or "").strip()
    if vkr_type not in ("project", "regular"):
        inferred = "project" if profile == "mpgu-09-project" else "regular"
        report.add("warning", "VKR_TYPE_MISSING",
                   f"vkr_type не задан или неизвестен («{vkr_type}») — проверки как для {inferred}", "vkr-project.json")
        vkr_type = inferred
    topic = str(config.get("topic") or "").strip()
    if not topic:
        report.add("error", "TOPIC_MISSING", "в vkr-project.json не заполнена тема (topic)", "vkr-project.json")
    title_page = config.get("title_page") if isinstance(config.get("title_page"), dict) else {}
    title_page = dict(title_page)
    members = config.get("members") if isinstance(config.get("members"), list) else []
    member_names = []
    for member in members:
        if isinstance(member, str) and member.strip():
            member_names.append(member.strip())
        elif isinstance(member, dict):
            name = member.get("full_name") or member.get("name") or member.get("fio")
            if isinstance(name, str) and name.strip():
                member_names.append(name.strip())
    if config.get("collective") is True and len(member_names) >= 2:
        title_page.pop("author", None)
        title_page["authors"] = member_names
    if not (str(title_page.get("author") or "").strip() or title_page.get("authors")):
        report.add("error", "AUTHOR_MISSING", "в vkr-project.json → title_page не заполнен author (ФИО полностью)", "vkr-project.json")

    drafts = root / "drafts"
    files = draft_files(drafts)
    for path in files["other"]:
        report.add("warning", "DRAFT_IGNORED", "файл не входит в сборку (ожидаются annotation, introduction, chapter-N, conclusion, appendix-N)",
                   f"drafts/{path.name}")

    parsed: Dict[str, Tuple[List[Dict[str, Any]], str]] = {}

    def load(rel: str, required: bool) -> Optional[List[Dict[str, Any]]]:
        path = root / rel
        if not path.is_file():
            if required:
                report.add("error", "DRAFT_REQUIRED_MISSING", "обязательный черновик отсутствует", rel)
            return None
        try:
            text = path.read_text(encoding="utf-8-sig")
        except (OSError, UnicodeDecodeError) as error:
            report.add("error", "DRAFT_UNREADABLE", f"не удалось прочитать как UTF-8: {error}", rel)
            return None
        blocks, problems = parse_markdown(text)
        for problem in problems:
            report.add(problem["level"], problem["code"], problem["message"], rel, problem["line"])
        stripped = vkr_markers.strip_markdown_code(text)
        for number, line in enumerate(stripped.split("\n"), 1):
            for marker in vkr_markers.find_markers(line):
                report.add("warning", "MARKER", f"маркер незавершённого текста: {marker}", rel, number)
        if not significant(blocks):
            report.add("error" if required else "warning", "DRAFT_EMPTY",
                       "черновик пуст: нет текста кроме заголовков и комментариев", rel)
        parsed[rel] = (blocks, text)
        return blocks

    content: Dict[str, Any] = {"title": topic, "title_page": title_page}
    for key in ("annotation", "introduction"):
        rel = f"drafts/{key}.md"
        blocks = load(rel, required=True)
        content[key] = section_blocks(blocks, rel, report)[1] if blocks is not None else []

    # Три главы — требование методички 09.03.02 для проектной ВКР; то же правило в doctor.
    required_chapters = 3 if profile == "mpgu-09-project" else 2
    numbers = list(files["chapters"].keys())
    for expected in range(1, max(numbers + [required_chapters]) + 1):
        if expected not in files["chapters"]:
            if expected <= required_chapters:
                report.add("error", "DRAFT_REQUIRED_MISSING",
                           f"нет главы {expected}: для профиля {profile or 'generic'} нужно не меньше {required_chapters} глав",
                           f"drafts/chapter-{expected}.md")
            else:
                report.add("error", "DRAFT_CHAPTER_GAP", f"пропущен номер главы {expected} — главы нумеруются подряд",
                           f"drafts/chapter-{expected}.md")
    content["chapters"] = []
    chapter_rels: List[str] = []
    for number, path in files["chapters"].items():
        rel = f"drafts/{path.name}"
        blocks = load(rel, required=True)
        if blocks is None:
            continue
        content["chapters"].append(split_chapter(blocks, rel, number, report))
        chapter_rels.append(rel)

    rel = "drafts/conclusion.md"
    blocks = load(rel, required=True)
    content["conclusion"] = section_blocks(blocks, rel, report)[1] if blocks is not None else []

    content["appendices"] = []
    appendix_rels: List[str] = []
    expected_app = 1
    for number, path in files["appendices"].items():
        rel = f"drafts/{path.name}"
        if number != expected_app:
            report.add("warning", "APPENDIX_NUMBER_GAP", f"приложения нумеруются подряд: ожидался appendix-{expected_app}.md", rel)
        expected_app = number + 1
        blocks = load(rel, required=False)
        if blocks is None:
            continue
        title, app_blocks = section_blocks(blocks, rel, report)
        title = re.sub(r"^\s*приложение\s+(?:[А-ЯЁA-Z]|\d+)\s*(?:[.:)]|[-–—])?\s*", "", title, count=1, flags=re.IGNORECASE).strip()
        if not title:
            report.add("error", "APPENDIX_TITLE_MISSING", "у приложения нет заголовка «# Название»", rel)
        content["appendices"].append({"title": title, "label": f"Приложение {number}", "blocks": app_blocks})
        appendix_rels.append(rel)

    # --- рисунки -------------------------------------------------------------
    def walk_blocks():
        for key in ("annotation", "introduction", "conclusion"):
            for block in content.get(key, []):
                yield f"drafts/{key}.md", block
        for rel_c, chapter in zip(chapter_rels, content["chapters"]):
            for block in chapter["blocks"]:
                yield rel_c, block
            for para in chapter["paragraphs"]:
                for block in para["blocks"]:
                    yield rel_c, block
            for block in chapter["conclusion"]:
                yield rel_c, block
        for rel_a, app in zip(appendix_rels, content["appendices"]):
            for block in app["blocks"]:
                yield rel_a, block

    root_resolved = root.resolve()
    for rel_file, block in walk_blocks():
        if block["type"] != "figure" or not block.get("path"):
            continue
        target = (root / block["path"]).resolve()
        if root_resolved != target and root_resolved not in target.parents:
            report.add("error", "FIGURE_OUTSIDE_PROJECT", f"путь рисунка вне проекта: {block['path']}", rel_file)
        elif not target.is_file():
            report.add("error", "FIGURE_NOT_FOUND", f"файл рисунка не найден: {block['path']}", rel_file)
        elif target.suffix.lower() not in SUPPORTED_IMAGE_EXT:
            report.add("error", "FIGURE_FORMAT", f"формат {target.suffix} не вставляется в DOCX — экспортируйте PNG или JPEG", rel_file)
        else:
            block["path"] = target.relative_to(root_resolved).as_posix()

    # --- источники ------------------------------------------------------------
    sources_path = root / "sources.json"
    raw_sources: List[Any] = []
    if sources_path.is_file():
        try:
            loaded = read_json(sources_path)
        except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
            raise UsageError(f"не удалось прочитать sources.json: {error}") from error
        if isinstance(loaded, dict) and isinstance(loaded.get("sources"), list):
            loaded = loaded["sources"]
        if not isinstance(loaded, list):
            raise UsageError("sources.json должен быть массивом записей")
        raw_sources = loaded
    else:
        report.add("warning", "SOURCES_MISSING", "нет sources.json — список литературы будет пуст", "sources.json")

    records: List[Dict[str, Any]] = []
    for index, raw in enumerate(raw_sources, 1):
        entry: Dict[str, Any] = {"index": index, "raw": raw, "normalized": None, "error": ""}
        if not isinstance(raw, dict):
            entry["error"] = f"запись #{index} не объект"
            entry["id"] = index
        else:
            entry["id"] = raw.get("id", index)
            try:
                entry["normalized"] = fb.normalize_source(raw, index)
            except fb.BibliographyError as error:
                entry["error"] = str(error)
        raw_status = raw.get("status") if isinstance(raw, dict) else None
        entry["status"] = common.normalize_source_status(raw_status) or "pending"
        if not str(raw_status or "").strip():
            report.add("warning", "SOURCE_STATUS_UNKNOWN",
                       f"источник {entry['id']}: статус не задан — считается pending (doctor требует status)", "sources.json")
        elif entry["status"] not in common.SOURCE_STATUSES:
            report.add("warning", "SOURCE_STATUS_UNKNOWN",
                       f"источник {entry['id']}: неизвестный статус «{entry['status']}» (допустимы {', '.join(fb.STATUS_VALUES)})",
                       "sources.json")
        records.append(entry)
    seen: Dict[str, int] = {}
    for entry in records:
        key = str(entry["id"])
        if key in seen:
            report.add("error", "SOURCE_ID_DUPLICATE", f"повторяющийся id «{entry['id']}» (записи #{seen[key]} и #{entry['index']})", "sources.json")
        seen[key] = entry["index"]
    by_id = {str(e["id"]): e for e in records}
    by_int = {int(str(e["id"])): e for e in records if str(e["id"]).isdigit()}

    cited: Dict[str, List[str]] = {}
    conclusion_warned = False
    for rel_file, block in walk_blocks():
        for text in block_texts(block):
            for kind, ref in fb.citation_refs(text):
                entry = by_id.get(str(ref)) if kind == "id" else by_int.get(int(ref))
                shown = f"@{ref}" if kind == "id" else str(ref)
                if entry is None:
                    report.add("error", "SOURCE_UNKNOWN", f"ссылка на источник {shown}, которого нет в sources.json", rel_file)
                    continue
                cited.setdefault(str(entry["id"]), []).append(rel_file)
                if rel_file == "drafts/conclusion.md" and not conclusion_warned:
                    conclusion_warned = True
                    report.add("warning", "CITATION_IN_CONCLUSION", "в заключении не должно быть ссылок и цитат (методичка)", rel_file)
    for key, files_list in cited.items():
        entry = by_id[key]
        where = files_list[0]
        if entry["status"] == "rejected":
            report.add("error", "SOURCE_REJECTED", f"процитирован отклонённый источник {entry['id']} (status=rejected)", where)
        elif entry["normalized"] is None:
            report.add("error", "SOURCE_INVALID", f"процитирован источник {entry['id']} с ошибкой в записи: {entry['error']}", where)
        elif not fb.is_structured(entry["normalized"]):
            report.add("error", "SOURCE_UNSTRUCTURED",
                       f"процитирован источник {entry['id']}, у которого есть только raw_text — разберите запись на поля", where)
    unconfirmed = sorted(str(k) for k in cited if by_id[k]["status"] in ("pending", "suspicious"))
    if unconfirmed:
        report.add("warning", "SOURCES_NOT_CONFIRMED", "процитированы неподтверждённые источники: " + ", ".join(unconfirmed), "sources.json")
    unverified = sorted(str(k) for k in cited if by_id[k]["status"] == "confirmed" and not verification_complete(by_id[k]["raw"]))
    if unverified:
        report.add("warning", "SOURCE_VERIFICATION_MISSING",
                   "у confirmed-источников нет verification {method, checked_at, url|note} (на стадии final doctor "
                   "считает это ошибкой): " + ", ".join(unverified), "sources.json")
    stale = sorted(str(k) for k in cited
                   if by_id[k]["status"] == "confirmed" and sources_tool.verification_stale(by_id[k]["raw"]))
    if stale:
        report.add("warning", "SOURCE_VERIFICATION_STALE",
                   "у confirmed-источников реквизиты изменены после проверки (verification.fields_sha256 не совпадает) — "
                   "сверь запись с источником и отметь заново verify_sources.py --mark: " + ", ".join(stale), "sources.json")

    included: List[Dict[str, Any]] = []
    for entry in records:
        key = str(entry["id"])
        if entry["status"] == "rejected":
            continue
        if entry["normalized"] is None or not fb.is_structured(entry["normalized"]):
            if key not in cited:
                report.add("warning", "SOURCE_SKIPPED",
                           f"источник {entry['id']} не процитирован и не оформляется ({entry['error'] or 'только raw_text'}) — в список не вошёл",
                           "sources.json")
            continue
        try:
            fb.format_source(entry["normalized"], [])
        except fb.BibliographyError as error:
            level = "error" if key in cited else "warning"
            report.add(level, "SOURCE_INVALID", f"источник {entry['id']}: {error}", "sources.json")
            continue
        normalized = dict(entry["normalized"])
        normalized["id"] = entry["id"]
        included.append(normalized)
        if key not in cited:
            report.add("warning", "SOURCE_NOT_CITED",
                       f"источник {entry['id']} есть в реестре, но не процитирован — методичка требует соответствия списка и текста",
                       "sources.json")
    if report.errors:
        return None

    try:
        numbering = fb.number_sources(included, group_normative_first=group_normative_first)
    except fb.BibliographyError as error:
        report.add("error", "BIBLIOGRAPHY_FAILED", str(error), "sources.json")
        return None
    for message in numbering["warnings"]:
        report.add("warning", "SOURCE_FORMAT", message, "sources.json")
    resolver = fb.make_resolver(numbering["mapping"], [e["id"] for e in numbering["entries"]])
    problems: List[str] = []

    def rewrite(text: str) -> str:
        return sort_citation_numbers(fb.rewrite_citations(text, resolver, problems))

    for _, block in walk_blocks():
        btype = block["type"]
        if btype == "text":
            block["content"] = rewrite(block["content"])
        elif btype == "list":
            block["items"] = [rewrite(x) for x in block["items"]]
        elif btype == "table":
            block["title"] = rewrite(block.get("title", ""))
            block["headers"] = [rewrite(x) for x in block.get("headers", [])]
            block["rows"] = [[rewrite(x) for x in row] for row in block.get("rows", [])]
            if block.get("source"):
                block["source"] = rewrite(block["source"])
        elif btype in ("listing", "figure"):
            block["caption"] = rewrite(block.get("caption", ""))
    if problems:
        for message in sorted(set(problems)):
            report.add("error", "SOURCE_UNKNOWN", message)
        return None
    content["bibliography"] = [f"{e['number']}. {e['text']}" for e in numbering["entries"]]
    content["group_normative_first"] = bool(group_normative_first)
    # entries — текст записи под каждым номером: import_docx.py --update сверяет по нему список литературы DOCX.
    citation_map = {"schema": CITATION_MAP_SCHEMA,
                    "numbers": {str(e["number"]): e["id"] for e in numbering["entries"]},
                    "entries": {str(e["number"]): e["text"] for e in numbering["entries"]}}
    stats = {
        "vkr_type": vkr_type,
        "chapters": len(content["chapters"]),
        "paragraphs": sum(len(c["paragraphs"]) for c in content["chapters"]),
        "appendices": len(content["appendices"]),
        "tables": sum(1 for _, b in walk_blocks() if b["type"] == "table"),
        "figures": sum(1 for _, b in walk_blocks() if b["type"] == "figure"),
        "listings": sum(1 for _, b in walk_blocks() if b["type"] == "listing"),
        "lists": sum(1 for _, b in walk_blocks() if b["type"] == "list"),
        "sources_in_list": len(numbering["entries"]),
        "sources_cited": len(cited),
    }
    return {"content": content, "citation_map": citation_map, "stats": stats}


class UsageError(Exception):
    """Ошибка использования или ввода-вывода (код 2)."""


def utc_stamp() -> str:
    return dt.datetime.now(dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def backup_existing(root: Path, target: Path) -> Optional[str]:
    if not target.is_file():
        return None
    base = root / "backups" / "checkpoints"
    stamp = utc_stamp()
    folder = base / f"{stamp}-build"
    counter = 2
    while folder.exists():
        folder = base / f"{stamp}-build-{counter}"
        counter += 1
    folder.mkdir(parents=True)
    destination = folder / target.name
    shutil.copy2(str(target), str(destination))
    return destination.relative_to(root).as_posix()


def figure_paths(content: Dict[str, Any]) -> List[str]:
    """Пути рисунков собранного содержания (входы манифеста сборки)."""
    paths: List[str] = []

    def visit(blocks: Any) -> None:
        for block in blocks or []:
            if isinstance(block, dict) and block.get("type") == "figure" and block.get("path"):
                paths.append(str(block["path"]))

    for key in ("annotation", "introduction", "conclusion"):
        visit(content.get(key))
    for chapter in content.get("chapters") or []:
        visit(chapter.get("blocks"))
        visit(chapter.get("conclusion"))
        for paragraph in chapter.get("paragraphs") or []:
            visit(paragraph.get("blocks"))
    for appendix in content.get("appendices") or []:
        visit(appendix.get("blocks"))
    return paths


def final_docx_next_steps(root: Path) -> List[Dict[str, str]]:
    """Канонические шаги 6.33 после сборки final/vkr.docx: поля Word → метаданные → валидация."""
    project = root.resolve()
    docx = project / "final" / "vkr.docx"

    def command(script: str, *args: str) -> str:
        return " ".join([f'python "{SCRIPTS_DIR / script}"'] + list(args))

    return [
        {"step": "update_fields",
         "command": command("update_docx_fields.py", f'"{docx}"'),
         "purpose": "обновить оглавление и поля через Microsoft Word",
         "manual": "код 1 (Word недоступен): открыть final/vkr.docx в Word, Ctrl+A, F9 "
                   "(оглавление — «Обновить целиком»), сохранить под тем же именем"},
        {"step": "clean_metadata",
         "command": command("clean_docx_metadata.py", f'"{docx}"', "--in-place"),
         "purpose": "очистить метаданные сдаваемого файла на месте"},
        {"step": "validate",
         "command": command("vkr_audit.py", f'"{project}"', "validate"),
         "purpose": "проверить final/vkr.docx валидатором → audit/automated-validation.json"},
    ]


def new_build_id() -> str:
    """Идентификатор сборки: общий для build-manifest.json и citation-map.json одной сборки."""
    return f"{utc_stamp()}-{uuid.uuid4().hex[:12]}"


def project_label(root: Path, path: Path) -> str:
    try:
        return Path(path).resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        return str(path)


def write_blockers(root: Path, out_path: Path, final_build: bool) -> List[Tuple[str, str]]:
    """Что помешает записать результат сборки; проверка до любых записей, файлы не меняются.

    Открытый в Word DOCX (Windows) не даёт атомарно заменить файл — сборка отказывает
    сразу, а не после перезаписи части файлов.
    """
    blockers: List[Tuple[str, str]] = []
    label = project_label(root, out_path)
    reason = fields_tool.write_blocker(out_path)
    if reason:
        blockers.append(("OUTPUT_LOCKED", f"{label}: {reason} — закрой файл в Word и повтори сборку; ничего не изменено"))
    targets = [ROOT_BUILD_MANIFEST, CITATION_MAP_PATH, CONTENT_PATH] if final_build else [PREVIEW_CONTENT_PATH]
    for relative in targets:
        reason = fields_tool.write_blocker(root / relative)
        if reason:
            blockers.append(("OUTPUT_LOCKED", f"{relative}: {reason} — закрой программу, в которой открыт файл, "
                                              "и повтори сборку; ничего не изменено"))
    for directory in {out_path.parent, root / "exports"}:
        reason = fields_tool.directory_blocker(directory)
        if reason:
            blockers.append(("IO_ERROR", f"{reason}; ничего не изменено"))
    return blockers


def stage_json(path: Path, value: Any) -> Path:
    """Пишет JSON во временный файл рядом с целевым; подмена — os.replace после записи DOCX."""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=str(path.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as handle:
            json.dump(value, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
    except BaseException:
        if os.path.exists(tmp):
            os.unlink(tmp)
        raise
    return Path(tmp)


def restore_previous_docx(out_path: Path, backup: Optional[Path]) -> Optional[str]:
    """Возвращает прежний DOCX после сбоя записи служебных файлов; текст ошибки или None."""
    try:
        if backup is None:
            if out_path.exists():
                out_path.unlink()
            return None
        fd, tmp = tempfile.mkstemp(prefix=".vkr-restore-", suffix=".tmp", dir=str(out_path.parent))
        os.close(fd)
        shutil.copy2(str(backup), tmp)
        os.replace(tmp, str(out_path))
        return None
    except OSError as error:
        return str(error)


def run_build(root: Path, *, output: Optional[str] = None, content_only: bool = False,
              group_normative_first: bool = False) -> Tuple[int, Dict[str, Any]]:
    """Сборка. Порядок записи: проверка возможности записи → DOCX во временный файл →
    атомарная подмена DOCX → build-manifest.json, citation-map.json, vkr-content.json
    (атомарно, с общим build_id). ``--content-only`` пишет только
    exports/vkr-content-preview.json и не трогает файлы последней сборки."""
    report = Report()
    result: Dict[str, Any] = {"tool": "build_vkr", "version": BUILD_VERSION, "project": str(root),
                              "status": "error", "exit_code": 2, "outputs": {}, "stats": {},
                              "errors": report.errors, "warnings": report.warnings, "next_steps": []}
    if not root.is_dir():
        report.add("error", "PROJECT_NOT_FOUND", f"каталог проекта не найден: {root}")
        return 2, result
    try:
        collected = collect_project(root, report, group_normative_first=group_normative_first)
    except UsageError as error:
        report.add("error", "INPUT_ERROR", str(error))
        return 2, result
    if collected is None or report.errors:
        result.update(status="fail", exit_code=1)
        return 1, result
    content = collected["content"]
    result["stats"] = collected["stats"]
    if content_only:
        # Проверка содержания без DOCX: файлы последней сборки (citation-map, build-manifest,
        # vkr-content) описывают final/vkr.docx, поэтому здесь не перезаписываются.
        blocker = fields_tool.write_blocker(root / PREVIEW_CONTENT_PATH)
        if blocker:
            report.add("error", "OUTPUT_LOCKED", f"{PREVIEW_CONTENT_PATH}: {blocker}")
            return 2, result
        try:
            write_json_atomic(root / PREVIEW_CONTENT_PATH, content)
        except OSError as error:
            report.add("error", "IO_ERROR", f"не удалось записать {PREVIEW_CONTENT_PATH}: {error}")
            return 2, result
        result["outputs"]["content"] = PREVIEW_CONTENT_PATH
        result.update(status="ok", exit_code=0)
        return 0, result

    out_path = Path(output) if output else Path("final") / "vkr.docx"
    if not out_path.is_absolute():
        out_path = root / out_path
    final_build = out_path.resolve() == (root / "final" / "vkr.docx").resolve()
    blockers = write_blockers(root, out_path, final_build)
    if blockers:
        for code, message in blockers:
            report.add("error", code, message)
        return 2, result
    try:
        import create_vkr_docx as generator  # noqa: WPS433
    except ImportError as error:
        report.add("error", "GENERATOR_MISSING", f"не найден create_vkr_docx.py: {error}")
        return 2, result

    renderer_warnings: List[str] = []
    staged: List[Tuple[Path, Path]] = []
    docx_tmp: Optional[Path] = None
    backup_rel: Optional[str] = None
    replaced = False
    try:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp_name = tempfile.mkstemp(prefix=".vkr-build-", suffix=".tmp", dir=str(out_path.parent))
        os.close(fd)
        docx_tmp = Path(tmp_name)
        try:
            generator.build_vkr_docx(content, docx_tmp, base_dir=root, warnings=renderer_warnings)
        except ImportError:
            report.add("error", "DEPENDENCY_MISSING", "не установлен python-docx: pip install python-docx")
            return 2, result
        except (generator.InputError, generator.ContentError) as error:
            for message in str(error).split("\n"):
                report.add("error", "RENDER_FAILED", message)
            result.update(status="fail", exit_code=1)
            return 1, result
        for message in renderer_warnings:
            report.add("warning", "RENDER", message)

        others = [p.name for p in out_path.parent.glob("*.docx")
                  if p.name != out_path.name and not p.name.startswith("~$")]
        if others:
            report.add("warning", "FINAL_OTHER_DOCX",
                       f"в {out_path.parent.name}/ есть другие DOCX: {', '.join(sorted(others))} — сдаётся один файл final/vkr.docx")

        # Служебные файлы готовятся до подмены DOCX и подменяются только после неё.
        exports = root / "exports"
        if final_build:
            build_id = new_build_id()
            # Манифест сборки: doctor сверяет по нему final/vkr.docx с черновиками (FINAL_DOCX_OUTDATED).
            manifest = common.build_manifest_payload(root, docx_tmp, figure_paths(content), BUILD_VERSION)
            manifest["docx"]["path"] = common.FINAL_DOCX_PATH
            manifest["build_id"] = build_id
            # Предупреждения сборки сохраняются в манифесте: в новом чате их иначе не увидеть.
            manifest["warnings"] = list(report.warnings)
            citation_map = dict(collected["citation_map"])
            citation_map["build_id"] = build_id
            plan = [(root / ROOT_BUILD_MANIFEST, manifest), (root / CITATION_MAP_PATH, citation_map), (root / CONTENT_PATH, content)]
        else:
            plan = [(root / PREVIEW_CONTENT_PATH, content)]
        for target, value in plan:
            staged.append((stage_json(target, value), target))

        backup_rel = backup_existing(root, out_path)
        try:
            os.replace(str(docx_tmp), str(out_path))
        except OSError as error:
            if backup_rel:
                shutil.rmtree(str((root / backup_rel).parent), ignore_errors=True)
                backup_rel = None
            report.add("error", "OUTPUT_LOCKED",
                       f"не удалось заменить {project_label(root, out_path)} ({error}) — закрой файл в Word и повтори "
                       "сборку; служебные файлы сборки не изменены")
            return 2, result
        replaced = True
        docx_tmp = None
        try:
            for tmp, target in staged:
                os.replace(str(tmp), str(target))
        except OSError as error:
            restore_error = restore_previous_docx(out_path, root / backup_rel if backup_rel else None)
            message = f"DOCX записан, но служебные файлы сборки не записаны ({error}); "
            message += ("прежний DOCX возвращён — повтори сборку" if restore_error is None
                        else f"вернуть прежний DOCX не удалось ({restore_error}) — повтори сборку")
            report.add("error", "IO_ERROR", message)
            return 2, result
        staged = []
    except common.ProjectDataError as error:
        report.add("error", "IO_ERROR", f"не удалось прочитать собранный DOCX для манифеста: {error}")
        return 2, result
    except OSError as error:
        report.add("error", "IO_ERROR",
                   f"не удалось записать результат ({error}); если {out_path.name} открыт в Word — закрой его"
                   + ("" if replaced else "; ничего не изменено"))
        return 2, result
    finally:
        if docx_tmp is not None and docx_tmp.exists():
            docx_tmp.unlink()
        for tmp, _target in staged:
            if tmp.exists():
                tmp.unlink()

    result["outputs"]["docx"] = project_label(root, out_path)
    result["outputs"]["backup"] = backup_rel
    if final_build:
        result["outputs"]["content"] = CONTENT_PATH
        result["outputs"]["citation_map"] = CITATION_MAP_PATH
        result["outputs"]["build_manifest"] = ROOT_BUILD_MANIFEST
        result["build_id"] = build_id
        result["next_steps"] = final_docx_next_steps(root)
    else:
        result["outputs"]["content"] = PREVIEW_CONTENT_PATH
    result.update(status="ok", exit_code=0)
    return 0, result


def _configure_streams() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except (ValueError, OSError):
                pass


def main(argv: Optional[List[str]] = None) -> int:
    _configure_streams()
    parser = argparse.ArgumentParser(description="Сборка ВКР: drafts/*.md + sources.json → final/vkr.docx")
    parser.add_argument("project_dir", help="Каталог проекта ВКР")
    parser.add_argument("--json", action="store_true", help="Отчёт в stdout как JSON")
    parser.add_argument("-o", dest="report", help="Записать JSON-отчёт в файл")
    parser.add_argument("--output", help="Путь DOCX (по умолчанию final/vkr.docx; относительный — от каталога проекта)")
    parser.add_argument("--content-only", action="store_true",
                        help="Проверка содержания без DOCX: только exports/vkr-content-preview.json; "
                             "citation-map.json, build-manifest.json и vkr-content.json не меняются")
    parser.add_argument("--group-normative-first", action="store_true",
                        help="Нормативные акты и стандарты — в начало списка литературы")
    args = parser.parse_args(argv)
    root = Path(args.project_dir)
    code, result = run_build(root, output=args.output, content_only=args.content_only,
                             group_normative_first=args.group_normative_first)
    if args.report:
        try:
            report_path = Path(args.report)
            report_path.parent.mkdir(parents=True, exist_ok=True)
            report_path.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        except OSError as error:
            print(f"ОШИБКА: не удалось записать отчёт {args.report}: {error}", file=sys.stderr)
            code = 2
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return code
    for entry in result["errors"]:
        where = entry.get("file", "")
        if entry.get("line"):
            where += f":{entry['line']}"
        print(f"ОШИБКА [{entry['code']}] {where + ': ' if where else ''}{entry['message']}", file=sys.stderr)
    for entry in result["warnings"]:
        where = entry.get("file", "")
        if entry.get("line"):
            where += f":{entry['line']}"
        print(f"ПРЕДУПРЕЖДЕНИЕ [{entry['code']}] {where + ': ' if where else ''}{entry['message']}", file=sys.stderr)
    if code == 0:
        outputs = result["outputs"]
        print(f"OK: собрано {outputs.get('docx') or outputs.get('content')}")
        if outputs.get("backup"):
            print(f"   прежний DOCX сохранён: {outputs['backup']}")
        if result["next_steps"]:
            print("   Дальше:")
            for number, step in enumerate(result["next_steps"], 1):
                print(f"   {number}. {step['command']}")
                if step.get("manual"):
                    print(f"      {step['manual']}")
        elif outputs.get("docx"):
            print("   Это промежуточный DOCX: сдаётся только final/vkr.docx (сборка без --output).")
    elif code == 1:
        print("СБОРКА ОСТАНОВЛЕНА: исправьте ошибки содержания; файлы не перезаписаны.", file=sys.stderr)
    return code


if __name__ == "__main__":
    raise SystemExit(main())
