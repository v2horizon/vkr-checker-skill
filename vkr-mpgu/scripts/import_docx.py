#!/usr/bin/env python3
"""Импорт DOCX в черновики проекта ВКР (vkr-mpgu 6.33).

    python import_docx.py <PROJECT_DIR> --docx <FILE> [--update [--overwrite-drafts]] [--force [--replace-sources]] [--json] [-o REPORT]

Без ``--update`` — новый проект после init (или ``--force``): DOCX делится на
``drafts/annotation.md``, ``introduction.md``, ``chapter-N.md``, ``conclusion.md``,
``appendix-N.md``; список литературы становится реестром ``sources.json``
(``{"id": N, "raw_text": …, "status": "pending"}``), ссылки ``[N]`` остаются
числовыми. Записи списка разбираются той же функцией, что у
``verify_sources.py --extract`` (``parse_bibliography_blocks``): разорванная
Enter запись склеивается, id — номер записи. Если номера неоднозначны (номер
повторяется, абзац без номера не похож на продолжение, смешаны номера в тексте и
автонумерация Word) — код 1, импорт остановлен и ничего не записано. Титульный
лист, содержание (стили TOC, поле TOC, sdt-оглавление) и последний лист с
клятвой пропускаются. Код (стиль VKR Listing или моноширинный шрифт) →
fenced-блоки (пустые абзацы между моноширинными строками — пустые строки того же
листинга), таблицы → pipe-таблицы, рисунки → ``evidence/figures/``.

``--force`` не заменяет реестр со структурированными записями или статусами
confirmed/suspicious/rejected (правило ``verify_sources.py --extract``): реестр
остаётся, список литературы DOCX пишется в ``audit/sources-extracted.json``
(``{id, raw_text, status: "pending", hints}``) с предупреждением. Заменить такой
реестр — только ``--force --replace-sources`` (прежний — в резервной копии).

``--update`` — после ручной правки собранного DOCX в Word: переносит изменения
текста в существующие черновики поблочно (неизменённые блоки, комментарии и
переносы строк остаются), числа ``[N]`` переводит обратно в ``[@id]`` по
``exports/citation-map.json``; ``sources.json`` не меняется. До записи:
(0) сама карта сверяется с реестром: источник каждого номера (``numbers``)
оформляется тем же кодом, что и сборка, и сравнивается с ``entries`` — правка
``numbers`` (номера переставлены на другие ``id``) или правка ``sources.json``
после сборки дают код 1 «карта ссылок не соответствует реестру: пересобери» без
изменений черновиков; (1) записи списка литературы DOCX сверяются с ``entries``
карты последней сборки (нормализованный текст по номерам) — расхождение (карта от
другой сборки, список правили в Word) даёт код 1 без изменений черновиков;
(2) черновики, которые
импорт изменил бы, сверяются с хешами ``exports/build-manifest.json`` (того же
``build_id``, что у карты) и с ``exports/import-sync.json`` прошлого
``--update`` этой сборки — изменённые после сборки черновики дают код 1 со
списком файлов; ``--overwrite-drafts`` записывает версию из DOCX (прежние
черновики — в резервной копии).

Всегда пишет ``audit/import-report.json``; перезаписываемые файлы копируются в
``backups/checkpoints/<UTC>-import/``.

Коды: 0 — импорт выполнен; 1 — импорт остановлен проверкой (``status: fail``:
неоднозначные номера списка литературы, расхождение с последней сборкой,
черновики изменены после сборки; ничего не записано) или выполнен, но есть
нераспознанные элементы с потерей содержимого (``status: incomplete``: формулы,
диаграммы, сноски…), см. отчёт; 2 — ошибка использования, ввода-вывода или нет
python-docx.
"""

from __future__ import annotations

import argparse
import datetime as dt
import difflib
import hashlib
import json
import re
import shutil
import sys
import unicodedata
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

sys.dont_write_bytecode = True
SCRIPTS_DIR = Path(__file__).resolve().parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import build_vkr  # noqa: E402
import format_bibliography as fb  # noqa: E402
import vkr_common as common  # noqa: E402

IMPORT_VERSION = "6.33"
EXTRACTED_SOURCES = "audit/sources-extracted.json"
CITATION_MAP_PATH = "exports/citation-map.json"
IMPORT_SYNC_PATH = "exports/import-sync.json"
BIBLIOGRAPHY_MISMATCH = ("список литературы в DOCX не совпадает с последней сборкой — добавь/исправь источник в "
                         "sources.json и пересобери; правки текста перенеси после этого")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
V = "urn:schemas-microsoft-com:vml"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
C_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


def w(tag: str) -> str:
    return f"{{{W}}}{tag}"


MONO_FONTS = {
    "consolas", "courier new", "courier", "lucida console", "lucida sans typewriter", "menlo", "monaco",
    "dejavu sans mono", "liberation mono", "source code pro", "jetbrains mono", "fira code", "fira mono",
    "roboto mono", "ubuntu mono", "pt mono", "cascadia code", "cascadia mono", "sf mono", "inconsolata",
    "noto sans mono", "andale mono", "droid sans mono", "hack", "anonymous pro", "ibm plex mono",
}
CODE_STYLE_RE = re.compile(r"(vkr listing|listing|листинг|source code|code|html preformatted|плоский текст|plain text)", re.IGNORECASE)
SECTION_PATTERNS = (
    ("annotation", re.compile(r"^аннотация\.?$", re.IGNORECASE)),
    ("contents", re.compile(r"^(?:содержание|оглавление)\.?$", re.IGNORECASE)),
    ("introduction", re.compile(r"^введение\.?$", re.IGNORECASE)),
    ("chapter", re.compile(r"^глава\s+(?P<num>[IVXLCDM]+|\d+)\b\.?\s*(?P<title>.*)$", re.IGNORECASE)),
    ("conclusion", re.compile(r"^заключение\.?$", re.IGNORECASE)),
    ("bibliography", re.compile(
        r"^(?:список\s+(?:использованн\w+\s+|цитируем\w+\s+)?(?:литературы|источников(?:\s+и\s+литературы)?)"
        r"|библиографическ\w+\s+список|библиография|литература)\.?$", re.IGNORECASE)),
    ("appendix", re.compile(r"^приложени[ея](?:\s+(?P<num>[А-ЯЁA-Z]|\d+))?\s*(?:[.:—–-]\s*(?P<title>.*))?$", re.IGNORECASE)),
)
CONCLUSION_HEADING_RE = re.compile(
    r"^выводы\s+по\s+(?:главе|первой\s+главе|второй\s+главе|третьей\s+главе)(?:\s+[IVXLCDM\d]+)?\.?$", re.IGNORECASE)
PARA_NUMBER_HEADING_RE = re.compile(r"^\d+\.\d+\.?\s+\S")
TABLE_LABEL_RE = re.compile(r"^таблица\s+(?P<num>[\w.]+)\s*(?:[.:—–-]\s*(?P<title>.*))?$", re.IGNORECASE)
FIGURE_CAPTION_RE = re.compile(r"^(?:рисунок|рис\.)\s*(?P<num>[\w.]+)\s*(?:[.:—–-]\s*(?P<title>.*))?$", re.IGNORECASE)
LISTING_CAPTION_RE = re.compile(r"^листинг\s+(?P<num>[\w.]+)\s*(?:[.:—–-]\s*(?P<title>.*))?$", re.IGNORECASE)
SOURCE_RE = re.compile(r"^источник\s*:\s*(?P<source>.+)$", re.IGNORECASE)
FIGURE_PLACEHOLDER_RE = re.compile(r"^\[\s*Здесь\s+вставить[^\]]*\]$", re.IGNORECASE)
TOC_LINE_RE = re.compile(r"(?:\.{3,}|…{2,}|\t)\s*\d+\s*$")
OATH_RE = re.compile(r"выполнена\s+мной\s+совершенно\s+самостоятельно", re.IGNORECASE)
BIB_NUMBER_RE = re.compile(r"^\s*(\d{1,3})[.)]\s+")
ROMAN = {1: "I", 2: "II", 3: "III", 4: "IV", 5: "V", 6: "VI", 7: "VII", 8: "VIII", 9: "IX", 10: "X"}
IMAGE_EXT = {"image/png": ".png", "image/jpeg": ".jpg", "image/gif": ".gif", "image/bmp": ".bmp",
             "image/tiff": ".tif", "image/x-emf": ".emf", "image/x-wmf": ".wmf", "image/svg+xml": ".svg"}


class UsageError(Exception):
    """Код 2."""


# ---------------------------------------------------------------------------
# Чтение DOCX
# ---------------------------------------------------------------------------

class DocxReader:
    """Проходит тело документа по порядку, отслеживая поля (TOC) между абзацами."""

    def __init__(self, doc) -> None:
        self.doc = doc
        self.part = doc.part
        self.field_stack: List[Dict[str, Any]] = []
        self.unrecognized: List[Dict[str, Any]] = []
        self.warnings: List[str] = []
        self._style_cache: Dict[str, Dict[str, Any]] = {}
        self._numbering = self._load_numbering()
        self._footnotes = self._load_notes("footnotes")

    # --- стили ----------------------------------------------------------------
    def _style_info(self, style) -> Dict[str, Any]:
        if style is None:
            return {"name": "", "font": None, "bold": None, "align": None, "outline": None, "num": None}
        key = style.style_id
        if key in self._style_cache:
            return self._style_cache[key]
        base = self._style_info(style.base_style) if style.base_style is not None else {
            "name": "", "font": None, "bold": None, "align": None, "outline": None, "num": None}
        element = style.element
        ppr = element.find(w("pPr"))
        rpr = element.find(w("rPr"))
        info = dict(base)
        info["name"] = style.name or ""
        if rpr is not None:
            fonts = rpr.find(w("rFonts"))
            if fonts is not None and (fonts.get(w("ascii")) or fonts.get(w("hAnsi"))):
                info["font"] = fonts.get(w("ascii")) or fonts.get(w("hAnsi"))
            bold = rpr.find(w("b"))
            if bold is not None:
                info["bold"] = bold.get(w("val")) not in ("0", "false", "off")
        if ppr is not None:
            jc = ppr.find(w("jc"))
            if jc is not None:
                info["align"] = jc.get(w("val"))
            outline = ppr.find(w("outlineLvl"))
            if outline is not None:
                info["outline"] = int(outline.get(w("val"), "9"))
            num_pr = ppr.find(w("numPr"))
            if num_pr is not None:
                num_id = num_pr.find(w("numId"))
                ilvl = num_pr.find(w("ilvl"))
                if num_id is not None:
                    info["num"] = (num_id.get(w("val")), int(ilvl.get(w("val"), "0")) if ilvl is not None else 0)
        self._style_cache[key] = info
        return info

    def _paragraph_style(self, p_el):
        from docx.enum.style import WD_STYLE_TYPE

        ppr = p_el.find(w("pPr"))
        style_id = None
        if ppr is not None and ppr.find(w("pStyle")) is not None:
            style_id = ppr.find(w("pStyle")).get(w("val"))
        try:
            return self.part.get_style(style_id, WD_STYLE_TYPE.PARAGRAPH)
        except Exception:
            return None

    def _run_style_font(self, r_el) -> Optional[str]:
        from docx.enum.style import WD_STYLE_TYPE

        rpr = r_el.find(w("rPr"))
        if rpr is None or rpr.find(w("rStyle")) is None:
            return None
        try:
            style = self.part.get_style(rpr.find(w("rStyle")).get(w("val")), WD_STYLE_TYPE.CHARACTER)
        except Exception:
            return None
        while style is not None:
            fonts = style.element.find(w("rPr"))
            fonts = fonts.find(w("rFonts")) if fonts is not None else None
            if fonts is not None and (fonts.get(w("ascii")) or fonts.get(w("hAnsi"))):
                return fonts.get(w("ascii")) or fonts.get(w("hAnsi"))
            style = style.base_style
        return None

    # --- нумерация ------------------------------------------------------------
    def _load_numbering(self) -> Dict[str, Any]:
        result: Dict[str, Any] = {"nums": {}, "abstract": {}}
        try:
            element = self.part.numbering_part.element
        except Exception:
            return result
        for abstract in element.findall(w("abstractNum")):
            levels = {}
            for lvl in abstract.findall(w("lvl")):
                fmt = lvl.find(w("numFmt"))
                levels[int(lvl.get(w("ilvl"), "0"))] = fmt.get(w("val")) if fmt is not None else "decimal"
            result["abstract"][abstract.get(w("abstractNumId"))] = levels
        for num in element.findall(w("num")):
            abstract_id = num.find(w("abstractNumId"))
            overrides = {}
            for override in num.findall(w("lvlOverride")):
                lvl = override.find(w("lvl"))
                if lvl is not None and lvl.find(w("numFmt")) is not None:
                    overrides[int(override.get(w("ilvl"), "0"))] = lvl.find(w("numFmt")).get(w("val"))
            result["nums"][num.get(w("numId"))] = {
                "abstract": abstract_id.get(w("val")) if abstract_id is not None else None, "overrides": overrides}
        return result

    def num_format(self, num_id: str, ilvl: int) -> Optional[str]:
        num = self._numbering["nums"].get(num_id)
        if not num:
            return None
        if ilvl in num["overrides"]:
            return num["overrides"][ilvl]
        return self._numbering["abstract"].get(num["abstract"], {}).get(ilvl, "decimal")

    def _load_notes(self, kind: str) -> Dict[str, str]:
        notes: Dict[str, str] = {}
        for rel in self.part.rels.values():
            if rel.reltype.endswith("/" + kind):
                try:
                    from lxml import etree

                    root = etree.fromstring(rel.target_part.blob)
                except Exception:
                    return notes
                for note in root.findall(w(kind[:-1])):
                    texts = ["".join(t.text or "" for t in note.iter(w("t")))]
                    notes[note.get(w("id"))] = " ".join(texts).strip()
        return notes

    # --- абзац ----------------------------------------------------------------
    def read_paragraph(self, p_el) -> Dict[str, Any]:
        style = self._paragraph_style(p_el)
        sinfo = self._style_info(style)
        ppr = p_el.find(w("pPr"))
        align = sinfo["align"]
        outline = sinfo["outline"]
        num = sinfo["num"]
        if ppr is not None:
            if ppr.find(w("jc")) is not None:
                align = ppr.find(w("jc")).get(w("val"))
            if ppr.find(w("outlineLvl")) is not None:
                outline = int(ppr.find(w("outlineLvl")).get(w("val"), "9"))
            num_pr = ppr.find(w("numPr"))
            if num_pr is not None and num_pr.find(w("numId")) is not None:
                ilvl = num_pr.find(w("ilvl"))
                num = (num_pr.find(w("numId")).get(w("val")), int(ilvl.get(w("val"), "0")) if ilvl is not None else 0)
        if num is not None and num[0] == "0":
            num = None
        info: Dict[str, Any] = {"type": "p", "style": sinfo["name"], "align": align, "outline": outline,
                                "runs": [], "images": [], "flags": set(), "toc": False, "bold_all": True,
                                "has_text_run": False}
        if self.field_stack and any(f["toc"] for f in self.field_stack):
            info["toc"] = True
        self._walk(p_el, info, sinfo)
        text = "".join(t for t, _ in info["runs"])
        info["text"] = text
        mono_chars = sum(len(t.strip()) for t, mono in info["runs"] if mono)
        all_chars = sum(len(t.strip()) for t, _ in info["runs"])
        style_code = bool(CODE_STYLE_RE.search(sinfo["name"] or ""))
        info["is_code"] = style_code or (all_chars > 0 and mono_chars == all_chars)
        if num is not None:
            fmt = self.num_format(num[0], num[1])
            info["num"] = {"id": num[0], "ilvl": num[1], "ordered": fmt not in (None, "bullet", "none")}
        else:
            info["num"] = None
        if not info["has_text_run"]:
            info["bold_all"] = False
        return info

    def _walk(self, element, info: Dict[str, Any], sinfo: Dict[str, Any]) -> None:
        for child in element:
            tag = child.tag
            if not isinstance(tag, str):
                continue
            if tag == f"{{{MC}}}Fallback" or tag in (w("pPr"), w("rPr"), w("del"), w("moveFrom")):
                continue
            if tag == f"{{{MC}}}AlternateContent":
                choice = child.find(f"{{{MC}}}Choice")
                if choice is not None:
                    self._walk(choice, info, sinfo)
                continue
            if tag == w("r"):
                self._run(child, info, sinfo)
                continue
            if tag == w("fldSimple"):
                instr = (child.get(w("instr")) or "").strip()
                if instr.upper().startswith("TOC"):
                    info["toc"] = True
                    continue
                self._walk(child, info, sinfo)
                continue
            if tag in (f"{{{M}}}oMath", f"{{{M}}}oMathPara"):
                text = "".join(t.text or "" for t in child.iter(f"{{{M}}}t"))
                info["flags"].add("equation")
                self.unrecognized.append({"kind": "equation", "text": text[:200], "lost_content": True})
                continue
            self._walk(child, info, sinfo)

    def _run(self, r_el, info: Dict[str, Any], sinfo: Dict[str, Any]) -> None:
        rpr = r_el.find(w("rPr"))
        font = None
        bold = None
        if rpr is not None:
            fonts = rpr.find(w("rFonts"))
            if fonts is not None:
                font = fonts.get(w("ascii")) or fonts.get(w("hAnsi"))
            if rpr.find(w("b")) is not None:
                bold = rpr.find(w("b")).get(w("val")) not in ("0", "false", "off")
        font = font or self._run_style_font(r_el) or sinfo["font"]
        if bold is None:
            bold = bool(sinfo["bold"])
        mono = bool(font) and font.strip().casefold() in MONO_FONTS
        in_result = True
        for child in r_el:
            tag = child.tag
            if tag == w("fldChar"):
                kind = child.get(w("fldCharType"))
                if kind == "begin":
                    self.field_stack.append({"instr": "", "separated": False, "toc": False})
                elif kind == "separate" and self.field_stack:
                    self.field_stack[-1]["separated"] = True
                elif kind == "end" and self.field_stack:
                    self.field_stack.pop()
                continue
            if tag == w("instrText"):
                if self.field_stack:
                    self.field_stack[-1]["instr"] += child.text or ""
                    if self.field_stack[-1]["instr"].strip().upper().startswith("TOC"):
                        self.field_stack[-1]["toc"] = True
                        info["toc"] = True
                continue
            in_result = not self.field_stack or all(f["separated"] for f in self.field_stack)
            if any(f["toc"] for f in self.field_stack):
                info["toc"] = True
            if not in_result:
                continue
            text = ""
            if tag == w("t"):
                text = child.text or ""
            elif tag == w("tab"):
                text = "\t"
            elif tag in (w("br"), w("cr")):
                if child.get(w("type")) not in ("page", "column"):
                    text = "\n"
            elif tag == w("noBreakHyphen"):
                text = "-"
            elif tag == w("sym"):
                text = "?"
            elif tag in (w("drawing"), w("pict"), w("object")):
                self._graphic(child, info)
                continue
            elif tag in (w("footnoteReference"), w("endnoteReference")):
                note_id = child.get(w("id"))
                note_text = self._footnotes.get(note_id, "")
                info["flags"].add("footnote")
                self.unrecognized.append({"kind": "footnote", "text": note_text[:300], "lost_content": True})
                continue
            if text:
                info["runs"].append((text, mono))
                if text.strip():
                    info["has_text_run"] = True
                    if not bold:
                        info["bold_all"] = False

    def _graphic(self, element, info: Dict[str, Any]) -> None:
        chart = list(element.iter(f"{{{C_CHART}}}chart"))
        if chart:
            self.unrecognized.append({"kind": "chart", "text": "диаграмма Word (chart)", "lost_content": True})
            return
        if any(el.tag.startswith(f"{{{DGM}}}") for el in element.iter() if isinstance(el.tag, str)):
            self.unrecognized.append({"kind": "smartart", "text": "SmartArt", "lost_content": True})
            return
        textboxes = list(element.iter(w("txbxContent")))
        if textboxes:
            text = " ".join("".join(t.text or "" for t in box.iter(w("t"))) for box in textboxes).strip()
            if text:
                info["runs"].append((" " + text + " ", False))
                info["has_text_run"] = True
                self.warnings.append(f"текст из надписи (text box) перенесён в абзац: «{text[:60]}»")
        ids = [blip.get(f"{{{R}}}embed") for blip in element.iter(f"{{{A}}}blip") if blip.get(f"{{{R}}}embed")]
        ids += [img.get(f"{{{R}}}id") for img in element.iter(f"{{{V}}}imagedata") if img.get(f"{{{R}}}id")]
        if element.tag == w("object") and not ids:
            self.unrecognized.append({"kind": "ole_object", "text": "встроенный объект OLE", "lost_content": True})
            return
        for rid in ids:
            if rid not in info["images"]:
                info["images"].append(rid)

    def image_blob(self, rid: str) -> Tuple[bytes, str]:
        part = self.part.related_parts[rid]
        ext = IMAGE_EXT.get(part.content_type) or Path(str(part.partname)).suffix or ".bin"
        return part.blob, ext

    # --- таблица ----------------------------------------------------------------
    def read_table(self, tbl_el) -> Dict[str, Any]:
        rows: List[List[str]] = []
        images = 0
        for tr in tbl_el.findall(w("tr")):
            cells: List[str] = []
            for tc in tr.findall(w("tc")):
                span = tc.find(w("tcPr"))
                repeat = 1
                if span is not None and span.find(w("gridSpan")) is not None:
                    repeat = int(span.find(w("gridSpan")).get(w("val"), "1"))
                texts = []
                for p_el in tc.iter(w("p")):
                    para = self.read_paragraph(p_el)
                    images += len(para["images"])
                    if para["text"].strip():
                        texts.append(render_inline(para))
                if tc.find(".//" + w("tbl")) is not None:
                    self.warnings.append("вложенная таблица развернута в текст ячейки")
                cell = " ".join(t.strip() for t in texts).replace("\n", " ")
                cells.extend([cell] * repeat)
            rows.append(cells)
        if images:
            self.unrecognized.append({"kind": "image_in_table", "text": f"{images} изображений в ячейках таблицы",
                                      "lost_content": True})
        return {"type": "tbl", "rows": rows}

    def items(self) -> List[Dict[str, Any]]:
        out: List[Dict[str, Any]] = []
        self._collect(self.doc.element.body, out)
        return out

    def _collect(self, parent, out: List[Dict[str, Any]]) -> None:
        for child in parent:
            tag = child.tag
            if tag == w("p"):
                out.append(self.read_paragraph(child))
            elif tag == w("tbl"):
                out.append(self.read_table(child))
            elif tag == w("sdt"):
                gallery = child.find(f"{w('sdtPr')}/{w('docPartObj')}/{w('docPartGallery')}")
                if gallery is not None and "table of contents" in (gallery.get(w("val")) or "").casefold():
                    continue
                content = child.find(w("sdtContent"))
                if content is not None:
                    self._collect(content, out)
            elif tag in (w("customXml"),):
                self._collect(child, out)


def render_inline(para: Dict[str, Any]) -> str:
    """Текст абзаца: моноширинные фрагменты внутри обычного текста → `код`."""
    segments: List[List[Any]] = []
    for piece, mono in para["runs"]:
        if segments and segments[-1][1] == mono:
            segments[-1][0] += piece
        else:
            segments.append([piece, mono])
    text = ""
    for piece, mono in segments:
        if mono and piece.strip() and not para.get("is_code"):
            lead = piece[: len(piece) - len(piece.lstrip())]
            trail = piece[len(piece.rstrip()):]
            text += f"{lead}`{piece.strip()}`{trail}"
        else:
            text += piece
    text = text.replace("\t", " ").replace("\n", " ")
    return re.sub(r"[  ]{2,}", " ", text).strip()


# ---------------------------------------------------------------------------
# Разделы
# ---------------------------------------------------------------------------

def _is_heading_like(item: Dict[str, Any]) -> bool:
    text = item["text"].strip()
    if not text or len(text) > 250 or item.get("num"):
        return False
    style = (item.get("style") or "").casefold()
    letters = [ch for ch in text if ch.isalpha()]
    upper = bool(letters) and all(ch.isupper() for ch in letters) and len(letters) > 3
    return (style.startswith("heading") or "заголов" in style or "title" in style or "vkr appendix" in style
            or item.get("outline") is not None and item["outline"] <= 2
            or item.get("bold_all") or item.get("align") in ("center", "right") or upper)


def heading_level(item: Dict[str, Any]) -> Optional[int]:
    style = (item.get("style") or "").casefold()
    match = re.match(r"^(?:heading|заголовок)\s*(\d)$", style)
    if match:
        return int(match.group(1))
    if item.get("outline") is not None and item["outline"] < 9:
        return item["outline"] + 1
    return None


def classify_section(item: Dict[str, Any]) -> Optional[Tuple[str, Dict[str, str]]]:
    if item["type"] != "p" or item.get("toc"):
        return None
    text = re.sub(r"\s+", " ", item["text"]).strip()
    if not text or TOC_LINE_RE.search(item["text"]):
        return None
    level = heading_level(item)
    if not (_is_heading_like(item) or level == 1):
        return None
    if text.casefold().rstrip(".") == "приложения":
        return "appendix_divider", {}
    for kind, pattern in SECTION_PATTERNS:
        match = pattern.match(text)
        if match:
            if kind == "appendix" and len(text) > 150:
                return None
            return kind, {k: (v or "") for k, v in match.groupdict().items()}
    return None


def roman_to_int(value: str) -> int:
    return build_vkr.roman_to_int(value.upper()) if not value.isdigit() else int(value)


def split_sections(items: List[Dict[str, Any]], report: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Делит элементы на разделы. Титул (до первого раздела) и последний лист отбрасываются."""
    # последний лист: «Выпускная квалификационная работа» + таблица с текстом клятвы
    cut = len(items)
    for index in range(len(items) - 1, -1, -1):
        item = items[index]
        if item["type"] == "tbl" and any(OATH_RE.search(" ".join(row)) for row in item["rows"]):
            cut = index
            for back in range(index - 1, max(-1, index - 4), -1):
                if items[back]["type"] == "p" and items[back]["text"].strip().casefold() == "выпускная квалификационная работа":
                    cut = back
                    break
            report["skipped"].append("последний лист (клятва)")
            break
    items = items[:cut]
    sections: List[Dict[str, Any]] = []
    current: Optional[Dict[str, Any]] = None
    title_page: List[str] = []
    seen_intro = seen_conclusion = False
    chapter_no = appendix_no = 0
    for item in items:
        if item["type"] == "p" and item.get("toc"):
            continue
        found = classify_section(item)
        level = heading_level(item) if item["type"] == "p" else None
        if found is None and item["type"] == "p" and level == 1 and seen_intro and not seen_conclusion \
                and item["text"].strip() and current is not None:
            found = ("chapter", {"num": "", "title": item["text"].strip()})
            report["warnings"].append(f"заголовок «{item['text'].strip()[:60]}» без «Глава N» принят за главу")
        if found and found[0] == "appendix_divider":
            report["skipped"].append("заголовок-разделитель «Приложения»")
            continue
        if found:
            kind, groups = found
            if kind == "introduction":
                seen_intro = True
            if kind == "conclusion":
                seen_conclusion = True
            current = {"kind": kind, "items": [], "title": "", "heading": item["text"].strip()}
            if kind == "chapter":
                chapter_no += 1
                current["number"] = chapter_no
                current["title"] = build_vkr.fb_strip_chapter(item["text"].strip())
                if groups.get("num") and roman_to_int(groups["num"]) != chapter_no:
                    report["warnings"].append(
                        f"глава «{item['text'].strip()[:40]}» получила номер {chapter_no} по порядку")
            elif kind == "appendix":
                appendix_no += 1
                current["number"] = appendix_no
                current["title"] = (groups.get("title") or "").strip()
            sections.append(current)
            continue
        if current is None:
            if item["type"] == "p" and item["text"].strip():
                title_page.append(item["text"].strip())
            continue
        current["items"].append(item)
    report["title_page_text"] = title_page[:40]
    return sections


# ---------------------------------------------------------------------------
# Блоки раздела
# ---------------------------------------------------------------------------

def _caption_title(match: "re.Match[str]") -> str:
    return (match.group("title") or "").strip().rstrip(".").strip()


def _is_code_item(item: Dict[str, Any]) -> bool:
    return item["type"] == "p" and bool(item["is_code"]) and not item["num"]


def _is_blank_item(item: Dict[str, Any]) -> bool:
    return item["type"] == "p" and not item["text"].strip() and not item["images"] and not item["num"]


def _code_lines(item: Dict[str, Any]) -> List[str]:
    # Разрыв строки (Shift+Enter) внутри абзаца кода — отдельная строка листинга.
    return [line.rstrip() for line in "".join(t for t, _ in item["runs"]).split("\n")]


def items_to_blocks(section: Dict[str, Any], reader: DocxReader, ctx: Dict[str, Any]) -> List[Dict[str, Any]]:
    items = section["items"]
    blocks: List[Dict[str, Any]] = []
    kind = section["kind"]
    i = 0
    if kind == "appendix" and not section["title"]:
        while i < len(items) and items[i]["type"] == "p" and not items[i]["text"].strip() and not items[i]["images"]:
            i += 1
        if i < len(items) and items[i]["type"] == "p" and _is_heading_like(items[i]):
            section["title"] = items[i]["text"].strip()
            i += 1
    pending_table_title: Optional[str] = None
    while i < len(items):
        item = items[i]
        if item["type"] == "tbl":
            rows = [[cell for cell in row] for row in item["rows"] if any(c.strip() for c in row)]
            headers = rows[0] if rows else []
            body = rows[1:] if rows else []
            block = {"kind": "table", "title": pending_table_title or "", "headers": headers, "rows": body, "source": ""}
            pending_table_title = None
            nxt = i + 1
            while nxt < len(items) and items[nxt]["type"] == "p" and not items[nxt]["text"].strip() and not items[nxt]["images"]:
                nxt += 1
            if nxt < len(items) and items[nxt]["type"] == "p":
                src = SOURCE_RE.match(items[nxt]["text"].strip())
                if src:
                    block["source"] = src.group("source").strip()
                    i = nxt
            blocks.append(block)
            i += 1
            continue
        text = item["text"].strip()
        if not text and not item["images"] and not item["is_code"]:
            i += 1
            continue
        if _is_code_item(item):
            code_lines: List[str] = []
            while i < len(items):
                if _is_code_item(items[i]):
                    code_lines.extend(_code_lines(items[i]))
                    i += 1
                    continue
                # Пустые абзацы между моноширинными строками — пустые строки того же листинга.
                gap = i
                while gap < len(items) and _is_blank_item(items[gap]):
                    gap += 1
                if gap > i and gap < len(items) and _is_code_item(items[gap]):
                    code_lines.extend([""] * (gap - i))
                    i = gap
                    continue
                break
            while code_lines and not code_lines[-1].strip():
                code_lines.pop()
            while code_lines and not code_lines[0].strip():
                code_lines.pop(0)
            caption = ""
            captioned = False  # есть абзац «Листинг N…», даже без названия
            if blocks and blocks[-1]["kind"] == "paragraph":
                match = LISTING_CAPTION_RE.match(blocks[-1]["text"])
                if match:
                    caption = _caption_title(match)
                    captioned = True
                    blocks.pop()
            if not captioned:
                nxt = i
                while nxt < len(items) and items[nxt]["type"] == "p" and not items[nxt]["text"].strip() and not items[nxt]["is_code"]:
                    nxt += 1
                if nxt < len(items) and items[nxt]["type"] == "p":
                    match = LISTING_CAPTION_RE.match(items[nxt]["text"].strip())
                    if match:
                        caption = _caption_title(match)
                        captioned = True
                        i = nxt + 1
            if code_lines:
                if not captioned:
                    previous = blocks[-1] if blocks else None
                    before = blocks[-2] if len(blocks) > 1 else None
                    if previous is not None and previous["kind"] == "listing" and not previous.get("captioned"):
                        ctx["warnings"].append(
                            f"два листинга без подписи подряд в разделе «{section.get('heading', '')[:40]}»: если это "
                            "один листинг, объедините fenced-блоки в черновике")
                    elif (previous is not None and previous["kind"] == "paragraph" and len(previous["text"]) <= 120
                          and before is not None and before["kind"] == "listing" and not before.get("captioned")):
                        ctx["warnings"].append(
                            f"листинг разделён абзацем без моноширинного шрифта «{previous['text'][:60]}» в разделе "
                            f"«{section.get('heading', '')[:40]}»: если это один листинг, объедините фрагменты в черновике")
                blocks.append({"kind": "listing", "caption": caption, "code": "\n".join(code_lines), "language": "",
                               "captioned": captioned})
            continue
        if item["num"]:
            num_id = item["num"]["id"]
            ordered = item["num"]["ordered"]
            list_items = []
            while i < len(items) and items[i]["type"] == "p" and items[i]["num"] and items[i]["num"]["id"] == num_id:
                if items[i]["text"].strip():
                    list_items.append(render_inline(items[i]))
                if items[i]["num"]["ilvl"] > 0:
                    ctx["warnings"].append("вложенный список импортирован на первом уровне")
                i += 1
            if list_items:
                blocks.append({"kind": "list", "ordered": ordered, "items": list_items})
            continue
        if item["images"] or FIGURE_PLACEHOLDER_RE.match(text):
            paths = [save_image(reader, rid, ctx) for rid in item["images"]] or [""]
            nxt = i + 1
            while nxt < len(items) and items[nxt]["type"] == "p" and not items[nxt]["text"].strip() and not items[nxt]["images"]:
                nxt += 1
            caption = ""
            if nxt < len(items) and items[nxt]["type"] == "p":
                match = FIGURE_CAPTION_RE.match(items[nxt]["text"].strip())
                if match:
                    caption = _caption_title(match)
                    i = nxt
            rest = FIGURE_PLACEHOLDER_RE.sub("", text).strip() if not item["images"] else text
            if rest:
                blocks.append({"kind": "paragraph", "text": render_inline(item)})
            for path in paths:
                blocks.append({"kind": "figure", "caption": caption, "path": path})
            if not caption:
                ctx["warnings"].append("рисунок без подписи «Рисунок N — Название»")
            i += 1
            continue
        table_label = TABLE_LABEL_RE.match(text)
        if table_label:
            title = _caption_title(table_label)
            nxt = i + 1
            if not title and nxt < len(items) and items[nxt]["type"] == "p" and items[nxt]["text"].strip() \
                    and nxt + 1 < len(items) and items[nxt + 1]["type"] == "tbl":
                title = items[nxt]["text"].strip().rstrip(".")
                nxt += 1
            if nxt < len(items) and items[nxt]["type"] == "tbl":
                pending_table_title = title
                i = nxt
                continue
        if kind == "chapter" and item["type"] == "p":
            level = heading_level(item)
            if CONCLUSION_HEADING_RE.match(text) and _is_heading_like(item):
                blocks.append({"kind": "heading", "level": 2, "text": "Выводы по главе"})
                i += 1
                continue
            if level == 2 or (PARA_NUMBER_HEADING_RE.match(text) and _is_heading_like(item) and len(text) < 200 and level is None):
                blocks.append({"kind": "heading", "level": 2, "text": build_vkr.strip_para_number(text)})
                i += 1
                continue
            if level is not None and level >= 3:
                ctx["warnings"].append(f"заголовок уровня {level} «{text[:40]}» импортирован как абзац")
        blocks.append({"kind": "paragraph", "text": render_inline(item)})
        i += 1
    return blocks


def save_image(reader: DocxReader, rid: str, ctx: Dict[str, Any]) -> str:
    try:
        blob, ext = reader.image_blob(rid)
    except KeyError:
        ctx["unrecognized"].append({"kind": "image", "text": f"не найдено изображение {rid}", "lost_content": True})
        return ""
    digest = hashlib.sha256(blob).hexdigest()
    figures: Path = ctx["root"] / "evidence" / "figures"
    known = ctx.setdefault("figure_hashes", None)
    if known is None:
        known = {}
        if figures.is_dir():
            for path in sorted(figures.iterdir()):
                if path.is_file():
                    known[hashlib.sha256(path.read_bytes()).hexdigest()] = path
        ctx["figure_hashes"] = known
    if digest in known:
        return known[digest].relative_to(ctx["root"]).as_posix()
    if ext in (".emf", ".wmf", ".svg"):
        ctx["warnings"].append(f"изображение {ext} сохранено, но build вставляет только PNG/JPEG — экспортируйте его в PNG")
    # Файл пишется только при фиксации импорта (write_pending_images): отказ импорта ничего не оставляет.
    pending: Dict[Path, bytes] = ctx.setdefault("pending_images", {})
    number = 1
    while (figures / f"import-{number}{ext}").exists() or figures / f"import-{number}{ext}" in pending:
        number += 1
    target = figures / f"import-{number}{ext}"
    pending[target] = blob
    known[digest] = target
    ctx["new_figures"].append(target.relative_to(ctx["root"]).as_posix())
    return target.relative_to(ctx["root"]).as_posix()


def write_pending_images(ctx: Dict[str, Any]) -> None:
    for target, blob in ctx.get("pending_images", {}).items():
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(blob)
    ctx["pending_images"] = {}


# ---------------------------------------------------------------------------
# Markdown и слияние
# ---------------------------------------------------------------------------

def render_block(block: Dict[str, Any]) -> List[str]:
    kind = block["kind"]
    if kind == "heading":
        return ["#" * block["level"] + " " + block["text"]]
    if kind == "paragraph":
        return [block["text"]]
    if kind == "list":
        if block["ordered"]:
            return [f"{n}. {item}" for n, item in enumerate(block["items"], 1)]
        return [f"- {item}" for item in block["items"]]
    if kind == "table":
        lines = []
        if block.get("title"):
            lines += [f"Таблица: {block['title']}", ""]
        width = max([len(block["headers"])] + [len(r) for r in block["rows"]] + [1])

        def row(cells: List[str]) -> str:
            cells = list(cells) + [""] * (width - len(cells))
            return "| " + " | ".join(c.replace("|", "\\|") for c in cells) + " |"

        lines.append(row(block["headers"]))
        lines.append("|" + "|".join(["---"] * width) + "|")
        lines += [row(r) for r in block["rows"]]
        if block.get("source"):
            lines += ["", f"Источник: {block['source']}"]
        return lines
    if kind == "listing":
        code = block["code"]
        longest = max([len(m.group(0)) for m in re.finditer(r"`+", code)] + [2])
        fence = "`" * max(3, longest + 1)
        lines = []
        if block.get("caption"):
            lines += [f"Листинг: {block['caption']}", ""]
        lines += [fence + (block.get("language") or "")] + code.split("\n") + [fence]
        return lines
    if kind == "figure":
        return [f"![{block['caption']}]({block['path']})"]
    return []


def render_markdown(blocks: List[Dict[str, Any]]) -> str:
    out: List[str] = []
    for block in blocks:
        if out:
            out.append("")
        out.extend(render_block(block))
    return "\n".join(out).rstrip() + "\n"


def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def canonical_citations(text: str, resolve, order=None) -> str:
    """Ссылки → [@id, лок.] для сравнения и записи; нерешённые остаются как есть.

    ``order(id)`` — ключ сортировки частей множественной ссылки (номер в последней
    сборке): build_vkr.py печатает номера по возрастанию, поэтому «[@b; @a]» в
    черновике и «[1; 2]» в DOCX — одна и та же ссылка. Общий префикс остаётся в начале;
    ссылки с префиксами у нескольких частей не переставляются.
    """
    citations = fb.find_citations(text)
    if not citations:
        return text
    out, pos = [], 0
    for citation in citations:
        rendered: List[Tuple[Any, str, str]] = []
        ok = True
        for part in citation["parts"]:
            refs = part["ref"] if part["kind"] == "range" else [part["ref"]]
            kind = "num" if part["kind"] == "range" else part["kind"]
            prefix = _norm(part["prefix"])
            locator = _norm(part["loc"]).lstrip(",").strip()
            for ref in refs:
                ident = resolve(kind, ref)
                if ident is None:
                    ok = False
                    break
                body = f"@{ident}" + (f", {locator}" if locator else "")
                rendered.append((ident, prefix, body))
                prefix = ""
            if not ok:
                break
        if ok and order is not None and len(rendered) > 1 and not any(prefix for _, prefix, _ in rendered[1:]):
            first_prefix = rendered[0][1]
            rendered = sorted(rendered, key=lambda item: order(item[0]))
            rendered = [(ident, first_prefix if index == 0 else "", body) for index, (ident, _p, body) in enumerate(rendered)]
        out.append(text[pos:citation["start"]])
        if ok:
            out.append("[" + "; ".join((prefix + " " if prefix else "") + body for _, prefix, body in rendered) + "]")
        else:
            out.append(citation["text"])
        pos = citation["end"]
    out.append(text[pos:])
    return "".join(out)


def map_block_texts(block: Dict[str, Any], fn) -> Dict[str, Any]:
    result = dict(block)
    kind = block["kind"]
    if kind == "paragraph":
        result["text"] = fn(block["text"])
    elif kind == "list":
        result["items"] = [fn(x) for x in block["items"]]
    elif kind == "table":
        result["title"] = fn(block.get("title", ""))
        result["headers"] = [fn(x) for x in block.get("headers", [])]
        result["rows"] = [[fn(x) for x in row] for row in block.get("rows", [])]
        result["source"] = fn(block.get("source", ""))
    elif kind in ("listing", "figure"):
        result["caption"] = fn(block.get("caption", ""))
    return result


def signature(block: Dict[str, Any], file_kind: str, first_h1: bool) -> str:
    kind = block["kind"]
    if kind == "heading":
        text = _norm(block["text"])
        if block["level"] == 1:
            if file_kind in ("annotation", "introduction", "conclusion"):
                return f"H1:{file_kind}"
            if file_kind == "chapter":
                return "H1:" + build_vkr.fb_strip_chapter(text).casefold()
            if file_kind == "appendix":
                return "H1:" + re.sub(r"^\s*приложение\s+(?:[А-ЯЁA-Z]|\d+)\s*(?:[.:)]|[-–—])?\s*", "", text, flags=re.IGNORECASE).casefold().rstrip(".")
        if block["level"] == 2 and file_kind == "chapter":
            if build_vkr.CONCLUSION_RE.match(text):
                return "H2:выводы по главе"
            return "H2:" + build_vkr.strip_para_number(text).casefold()
        return "P:" + text
    if kind == "paragraph":
        return "P:" + _norm(block["text"])
    if kind == "list":
        return f"L:{int(block['ordered'])}:" + "␟".join(_norm(x) for x in block["items"])
    if kind == "table":
        return "T:" + "␟".join([_norm(block.get("title", "")), _norm(block.get("source", ""))]
                                    + [_norm(x) for x in block.get("headers", [])]
                                    + ["␞".join(_norm(c) for c in row) for row in block.get("rows", [])])
    if kind == "listing":
        code = "\n".join(line.rstrip() for line in block["code"].split("\n")).strip("\n")
        return "C:" + _norm(block.get("caption", "")) + "␟" + code
    if kind == "figure":
        return "F:" + _norm(block.get("caption", "")) + "␟" + block.get("path", "")
    return kind


def merge_draft(old_text: str, new_blocks: List[Dict[str, Any]], file_kind: str, canon_old, canon_new) -> Tuple[str, int]:
    """Переносит изменённые блоки в черновик; возвращает (текст, число изменённых блоков)."""
    old_blocks, _ = build_vkr.parse_markdown(old_text)
    lines = old_text.replace("\r\n", "\n").split("\n")
    old_sig = [signature(map_block_texts(b, canon_old), file_kind, False) for b in old_blocks]
    new_canon = [map_block_texts(b, canon_new) for b in new_blocks]
    new_sig = [signature(b, file_kind, False) for b in new_canon]
    matcher = difflib.SequenceMatcher(a=old_sig, b=new_sig, autojunk=False)
    edits: List[Tuple[int, int, List[str], int]] = []
    changed = 0
    for order, (tag, i1, i2, j1, j2) in enumerate(matcher.get_opcodes()):
        if tag == "equal":
            continue
        changed += max(i2 - i1, j2 - j1)
        replacement_blocks = new_canon[j1:j2]
        if tag == "replace" and i2 - i1 == j2 - j1:
            for offset, block in enumerate(replacement_blocks):
                old = old_blocks[i1 + offset]
                if block["kind"] == "listing" and old["kind"] == "listing" and old.get("language"):
                    block["language"] = old["language"]
                if block["kind"] == "heading" and old["kind"] == "heading" and block["level"] == 1 and file_kind in ("annotation", "introduction", "conclusion"):
                    block["text"] = old["text"]
        rendered: List[str] = []
        for block in replacement_blocks:
            if rendered:
                rendered.append("")
            rendered.extend(render_block(block))
        if i1 < i2:
            start, end = old_blocks[i1]["start"], old_blocks[i2 - 1]["end"]
            covered = set()
            for block in old_blocks[i1:i2]:
                covered.update(range(block["start"], block["end"]))
            kept = [lines[k] for k in range(start, end) if k not in covered and lines[k].strip()]
            edits.append((start, end, rendered + ([""] + kept if kept else []), order))
        else:
            at = old_blocks[i1 - 1]["end"] if i1 > 0 else (old_blocks[0]["start"] if old_blocks else len(lines))
            edits.append((at, at, ([""] if at > 0 else []) + rendered + ([""] if i1 == 0 and old_blocks else []), order))
    # С конца файла; при одинаковом начале сначала более поздняя правка, чтобы вставка
    # перед ней не сдвинула её индексы.
    for start, end, replacement, _order in sorted(edits, key=lambda e: (e[0], e[3]), reverse=True):
        lines[start:end] = replacement
    if changed:
        lines = _collapse_blank_lines(lines)
    return "\n".join(lines).rstrip("\n") + "\n", changed


def _collapse_blank_lines(lines: List[str]) -> List[str]:
    """Не больше одной пустой строки подряд вне fenced-кода."""
    out: List[str] = []
    fence: Optional[str] = None
    for line in lines:
        match = build_vkr.FENCE_RE.match(line)
        if fence is None and match:
            fence = match.group(1)
        elif fence is not None and match and match.group(1)[0] == fence[0] and len(match.group(1)) >= len(fence) and not match.group(2):
            fence = None
            out.append(line)
            continue
        if fence is None and not line.strip() and out and not out[-1].strip():
            continue
        out.append(line)
    return out


# ---------------------------------------------------------------------------
# Проект
# ---------------------------------------------------------------------------

def utc_stamp() -> str:
    return dt.datetime.now(dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def write_lf(path: Path, text: str) -> None:
    """UTF-8 без BOM, переводы строк LF (Path.write_text(newline=) есть только с Python 3.10)."""
    with open(str(path), "w", encoding="utf-8", newline="\n") as handle:
        handle.write(text)


def register_figures(root: Path, report: Dict[str, Any]) -> None:
    """Добавляет извлечённые рисунки в evidence/index.json → figures ({id, path, kind})."""
    index_path = root / "evidence" / "index.json"
    if not index_path.is_file():
        report["warnings"].append("нет evidence/index.json — извлечённые рисунки не зарегистрированы в индексе доказательств")
        return
    try:
        index = json.loads(index_path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as error:
        report["warnings"].append(f"evidence/index.json не прочитан ({error}) — рисунки не зарегистрированы")
        return
    if not isinstance(index, dict) or not isinstance(index.get("figures", []), list):
        report["warnings"].append("evidence/index.json имеет неожиданную структуру — рисунки не зарегистрированы")
        return
    used = {str(item.get("id", "")).casefold() for group in index.values() if isinstance(group, list)
            for item in group if isinstance(item, dict)}
    known_paths = {str(item.get("path", "")) for item in index.get("figures", []) if isinstance(item, dict)}
    figures = index.setdefault("figures", [])
    for rel in report["figures"]:
        if rel in known_paths:
            continue
        ident = "figure-" + Path(rel).stem
        counter = 2
        while ident.casefold() in used:
            ident = f"figure-{Path(rel).stem}-{counter}"
            counter += 1
        used.add(ident.casefold())
        figures.append({"id": ident, "path": rel, "kind": "figure"})
    write_lf(index_path, json.dumps(index, ensure_ascii=False, indent=2) + "\n")


def registry_protection(sources_path: Path) -> Optional[str]:
    """Почему ``--force`` не заменяет реестр (None — заменить можно).

    Правило ``verify_sources.py --extract``: реестр защищён, если он не читается, не
    массив или в нём есть записи с полями сверх служебных (id, raw_text, status,
    hints, notes, found_url) либо со статусом, отличным от pending.
    """
    import verify_sources  # noqa: WPS433 — нужен только при повторном импорте

    try:
        data = json.loads(sources_path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as error:
        return f"sources.json не читается ({error})"
    if not isinstance(data, list):
        return "sources.json не является массивом записей"
    protected = [
        entry for entry in data
        if isinstance(entry, dict) and (
            set(entry) - verify_sources.REGISTRY_SERVICE_KEYS
            or common.normalize_source_status(entry.get("status")) not in ("", "pending")
        )
    ]
    if protected:
        return f"в sources.json {len(protected)} структурированных или проверенных записей (confirmed/suspicious/rejected)"
    return None


# ---------------------------------------------------------------------------
# Список литературы: общий разбор с verify_sources.py --extract
# ---------------------------------------------------------------------------

def _short(text: str, limit: int = 90) -> str:
    text = _norm(text)
    return text if len(text) <= limit else text[: limit - 1] + "…"


def bibliography_block(item: Dict[str, Any]) -> Dict[str, Any]:
    """Элемент DOCX → блок общего разбора ``verify_sources.parse_bibliography_blocks``."""
    if item["type"] == "tbl":
        return {"type": "tbl", "text": " ".join(cell.strip() for row in item["rows"] for cell in row if cell.strip())}
    style = (item.get("style") or "").casefold()
    letters = [ch for ch in item["text"] if ch.isalpha()]
    return {
        "type": "p",
        "text": item["text"],
        "toc": bool(item.get("toc")),
        "heading_level": heading_level(item),
        "in_list": bool(item.get("num")) or style.startswith(("list number", "list bullet")),
        "bibliography_style": style == "vkr bibliography",
        "bold": bool(item.get("bold_all")),
        "upper": len(letters) >= 4 and all(ch.isupper() for ch in letters),
    }


def parse_docx_bibliography(sections: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Записи всех разделов «Список литературы» DOCX тем же разбором, что у ``verify_sources.py --extract``."""
    import verify_sources  # noqa: WPS433

    items = [item for section in sections if section["kind"] == "bibliography" for item in section["items"]]
    parsed = verify_sources.parse_bibliography_blocks([bibliography_block(item) for item in items])
    parsed["found"] = any(section["kind"] == "bibliography" for section in sections)
    return parsed


def bibliography_problems(parsed: Dict[str, Any]) -> List[str]:
    """Почему номера записей нельзя однозначно сопоставить ссылкам [N] в тексте (пусто — можно)."""
    problems: List[str] = []
    groups: Dict[int, List[str]] = {}
    for entry in parsed["entries"]:
        if entry["number"] is not None:
            groups.setdefault(entry["number"], []).append(entry["raw_text"])
    for number, texts in sorted(groups.items()):
        if len(texts) > 1:
            shown = "; ".join(f"«{_short(text, 60)}»" for text in texts[:3])
            problems.append(f"номер {number} стоит у {len(texts)} записей ({shown}) — ссылку [{number}] нельзя связать "
                            "с одной записью")
    for entry, record in zip(parsed["entries"], parsed["records"]):
        number = entry["number"]
        if number is not None and len(groups.get(number, [])) == 1 and record["id"] != number:
            problems.append(f"номер {number} у записи «{_short(entry['raw_text'], 60)}» недопустим")
    for warning in parsed["warnings"]:
        if warning["code"] == "UNNUMBERED_PARAGRAPH":
            problems.append(f"абзац без номера «{warning['text']}» — не ясно, отдельная это запись, продолжение "
                            "предыдущей или подзаголовок")
    kinds = {entry["source"] for entry in parsed["entries"]}
    if "text_number" in kinds and "numbering" in kinds:
        problems.append("часть записей пронумерована в тексте, часть — автонумерацией Word: номера автонумерованных "
                        "записей в DOCX не определить")
    return problems


def bibliography_notes(parsed: Dict[str, Any]) -> List[str]:
    """Предупреждения разбора, не мешающие импорту (склейка, подзаголовки, номера не подряд)."""
    notes = [f"список литературы: {w['message']}: «{w['text']}»" for w in parsed["warnings"]
             if w["code"] not in ("UNNUMBERED_PARAGRAPH", "NUMBER_DUPLICATE")]
    ids = [record["id"] for record in parsed["records"]]
    if ids and ids != list(range(1, len(ids) + 1)):
        shown = ", ".join(str(value) for value in ids[:15]) + ("…" if len(ids) > 15 else "")
        notes.append(f"номера списка литературы в DOCX идут не подряд ({shown}) — id записей реестра равны этим номерам")
    return notes


def load_citation_map(root: Path) -> Tuple[Optional[Dict[str, Any]], str, int]:
    """(карта, ошибка, код) для ``--update``; карта сборки 6.33 содержит build_id и entries."""
    path = root / CITATION_MAP_PATH
    if not path.is_file():
        return None, "нет exports/citation-map.json — --update работает только с DOCX, собранным build_vkr.py", 2
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as error:
        return None, f"не удалось прочитать citation-map.json: {error}", 2
    if not isinstance(data, dict) or not isinstance(data.get("numbers"), dict):
        return None, "exports/citation-map.json повреждён (нет numbers) — пересобери DOCX: build_vkr.py", 2
    if not isinstance(data.get("entries"), dict) or not str(data.get("build_id") or "").strip():
        return None, ("exports/citation-map.json записан прежней версией build_vkr.py: в нём нет build_id и текста "
                      "записей списка литературы, поэтому нумерацию DOCX не сверить — пересобери DOCX (build_vkr.py) и "
                      "перенеси правки из Word после этого; если черновики потеряны — import_docx.py --force "
                      "(ссылки останутся номерами [N], failure-recovery.md)"), 1
    if {str(key) for key in data["numbers"]} != {str(key) for key in data["entries"]}:
        return None, "exports/citation-map.json повреждён (номера numbers и entries различаются) — пересобери DOCX: build_vkr.py", 2
    return data, "", 0


def citation_map_registry_problems(root: Path, citation_map: Dict[str, Any]) -> List[str]:
    """Расхождения карты ссылок с реестром источников.

    ``numbers`` (номер → id) и ``entries`` (номер → текст записи) пишет одна и та же
    сборка, но правкой ``numbers`` номера можно переставить на другие источники: тогда
    ``import --update`` заменил бы ``[N]`` в черновиках чужими ``[@id]``, а список
    литературы DOCX по-прежнему совпадал бы с ``entries``. Поэтому источник каждого
    номера оформляется тем же кодом, что и сборка (format_bibliography), и сверяется
    с ``entries[N]``.
    """
    path = root / "sources.json"
    try:
        loaded = json.loads(path.read_text(encoding="utf-8-sig")) if path.is_file() else []
    except (OSError, ValueError) as error:
        return [f"не удалось прочитать sources.json: {error}"]
    if isinstance(loaded, dict):
        loaded = loaded.get("sources", [])
    if not isinstance(loaded, list):
        return ["sources.json не является массивом записей"]
    by_id: Dict[str, Tuple[int, Dict[str, Any]]] = {}
    for index, raw in enumerate(loaded, 1):
        if isinstance(raw, dict):
            by_id.setdefault(str(raw.get("id", index)), (index, raw))

    def same(left: str, right: str) -> bool:
        return unicodedata.normalize("NFC", _norm(left)) == unicodedata.normalize("NFC", _norm(right))

    def order(key: str) -> Tuple[int, int, str]:
        return (0, int(key), "") if str(key).isdigit() else (1, 0, str(key))

    entries = {str(key): str(value) for key, value in (citation_map.get("entries") or {}).items()}
    numbers = {str(key): str(value) for key, value in citation_map["numbers"].items()}
    problems: List[str] = []
    for key in sorted(numbers, key=order):
        ident = numbers[key]
        found = by_id.get(ident)
        if found is None:
            problems.append(f"номер {key} указывает на источник {ident}, которого нет в sources.json")
            continue
        index, raw = found
        try:
            text = fb.format_source(fb.normalize_source(raw, index), [])
        except fb.BibliographyError as error:
            problems.append(f"номер {key}: источник {ident} из sources.json не оформляется по ГОСТ ({error})")
            continue
        if not same(text, entries.get(key, "")):
            problems.append(f"номер {key}: в сборке «{_short(entries.get(key, ''))}», а источник {ident} из "
                            f"sources.json оформляется как «{_short(text)}»")
    return problems


def bibliography_differences(parsed: Dict[str, Any], citation_map: Dict[str, Any]) -> List[str]:
    """Расхождения списка литературы DOCX с записями последней сборки (нормализованный текст по номерам)."""
    details = list(bibliography_problems(parsed))
    in_docx: Dict[str, str] = {}
    for entry, record in zip(parsed["entries"], parsed["records"]):
        in_docx.setdefault(str(entry["number"] if entry["number"] is not None else record["id"]), entry["raw_text"])
    built = {str(key): str(value) for key, value in citation_map["entries"].items()}

    def order(key: str) -> Tuple[int, int, str]:
        return (0, int(key), "") if key.isdigit() else (1, 0, key)

    def same(left: str, right: str) -> bool:
        return unicodedata.normalize("NFC", _norm(left)) == unicodedata.normalize("NFC", _norm(right))

    for key in sorted(set(in_docx) | set(built), key=order):
        if key not in in_docx:
            details.append(f"записи {key} из сборки нет в DOCX: «{_short(built[key])}»")
        elif key not in built:
            details.append(f"запись {key} есть только в DOCX: «{_short(in_docx[key])}»")
        elif not same(in_docx[key], built[key]):
            details.append(f"запись {key} отличается: в DOCX «{_short(in_docx[key])}», в сборке «{_short(built[key])}»")
    return details


# ---------------------------------------------------------------------------
# Черновики, изменённые после сборки
# ---------------------------------------------------------------------------

def load_build_manifest(root: Path) -> Tuple[Optional[Dict[str, Any]], str]:
    path = root / common.BUILD_MANIFEST_PATH
    if not path.is_file():
        return None, "нет exports/build-manifest.json"
    try:
        manifest = common.read_json(path, label=common.BUILD_MANIFEST_PATH)
    except common.ProjectDataError as error:
        return None, f"exports/build-manifest.json не читается ({error})"
    if not isinstance(manifest, dict) or not isinstance(manifest.get("inputs"), list):
        return None, "exports/build-manifest.json повреждён"
    if manifest.get("input_hash") != common.INPUT_HASH_SCHEMA:
        return None, "exports/build-manifest.json записан прежней версией build_vkr.py"
    return manifest, ""


def load_import_sync(root: Path, build_id: str) -> Dict[str, str]:
    """Хеши черновиков после прошлого ``--update`` той же сборки (эти правки пришли из DOCX)."""
    try:
        data = json.loads((root / IMPORT_SYNC_PATH).read_text(encoding="utf-8-sig"))
    except (OSError, ValueError):
        return {}
    if not isinstance(data, dict) or data.get("build_id") != build_id or not isinstance(data.get("drafts"), dict):
        return {}
    return {common.dedupe_key(common.normalize_logical_id(key)): str(value) for key, value in data["drafts"].items()}


def drafts_changed_after_build(root: Path, rels: List[str], manifest: Dict[str, Any], synced: Dict[str, str]) -> List[str]:
    """Черновики из ``rels``, чей текст не совпадает с последней сборкой (и с прошлым ``--update``)."""
    recorded: Dict[str, Any] = {}
    for item in manifest.get("inputs") or []:
        if isinstance(item, dict) and item.get("path"):
            recorded[common.dedupe_key(common.normalize_logical_id(item["path"]))] = item.get("sha256")
    changed: List[str] = []
    for rel in rels:
        key = common.dedupe_key(rel)
        path = root / rel
        if path.is_file():
            digest = common.content_sha256(path)
            if digest == recorded.get(key) or digest == synced.get(key):
                continue
            changed.append(f"{rel} ({'изменён' if key in recorded else 'создан'} после сборки)")
        elif key in recorded:
            changed.append(f"{rel} (удалён после сборки)")
    return changed


def write_import_sync(root: Path, build_id: str, rels: List[str], docx_path: Path) -> None:
    path = root / IMPORT_SYNC_PATH
    drafts: Dict[str, str] = {}
    try:
        previous = json.loads(path.read_text(encoding="utf-8-sig"))
        if isinstance(previous, dict) and previous.get("build_id") == build_id and isinstance(previous.get("drafts"), dict):
            drafts.update({str(key): str(value) for key, value in previous["drafts"].items()})
    except (OSError, ValueError):
        pass
    for rel in rels:
        if (root / rel).is_file():
            drafts[rel] = common.content_sha256(root / rel)
    payload = {"schema": "vkr-import-sync", "build_id": build_id, "updated_at": common.utc_now_iso(),
               "docx": str(docx_path), "drafts": dict(sorted(drafts.items()))}
    common.atomic_write_text(path, common.dump_json(payload))


def is_template_draft(text: str) -> bool:
    blocks, _ = build_vkr.parse_markdown(text)
    return not any(b["kind"] != "heading" for b in blocks)


def section_file(section: Dict[str, Any]) -> Optional[Tuple[str, str]]:
    kind = section["kind"]
    if kind in ("annotation", "introduction", "conclusion"):
        return f"drafts/{kind}.md", kind
    if kind == "chapter":
        return f"drafts/chapter-{section['number']}.md", "chapter"
    if kind == "appendix":
        return f"drafts/appendix-{section['number']}.md", "appendix"
    return None


def section_markdown(section: Dict[str, Any], blocks: List[Dict[str, Any]]) -> str:
    kind = section["kind"]
    names = {"annotation": "Аннотация", "introduction": "Введение", "conclusion": "Заключение"}
    if kind in names:
        head = f"# {names[kind]}"
    elif kind == "chapter":
        head = f"# Глава {ROMAN.get(section['number'], section['number'])}. {section['title'] or '[Название]'}"
    else:
        head = f"# {section['title'] or '[ЗАПОЛНИТЬ: название приложения]'}"
    return head + "\n\n" + render_markdown(blocks) if blocks else head + "\n"


def backup_writer(root: Path, report: Dict[str, Any]):
    """Копирование перезаписываемых файлов в backups/checkpoints/<UTC>-import[-N]/."""
    base = root / "backups" / "checkpoints"
    stamp = utc_stamp()
    folder = base / f"{stamp}-import"
    counter = 2
    while folder.exists():  # два импорта в одну секунду не должны смешивать копии
        folder = base / f"{stamp}-import-{counter}"
        counter += 1

    def backup(rel: str) -> None:
        source = root / rel
        if not source.is_file():
            return
        target = folder / rel
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(source), str(target))
        report["backup"] = folder.relative_to(root).as_posix()

    return backup


def _import_new(root: Path, report: Dict[str, Any], outputs: Dict[str, Tuple[str, List[Dict[str, Any]], Dict[str, Any]]],
                bibliography: Dict[str, Any], ctx: Dict[str, Any], force: bool, replace_sources: bool) -> Optional[int]:
    """Новый проект: черновики, рисунки и реестр пишутся только после всех проверок."""
    drafts_dir = root / "drafts"
    existing = sorted(drafts_dir.glob("*.md")) if drafts_dir.is_dir() else []
    occupied = [p for p in existing if not is_template_draft(p.read_text(encoding="utf-8-sig"))]
    sources_path = root / "sources.json"
    registry_filled = False
    if sources_path.is_file():
        try:
            loaded = json.loads(sources_path.read_text(encoding="utf-8-sig"))
            registry_filled = bool(loaded.get("sources") if isinstance(loaded, dict) else loaded)
        except ValueError:
            registry_filled = True
    if (occupied or registry_filled) and not force:
        what = [f"drafts/{p.name}" for p in occupied] + (["sources.json"] if registry_filled else [])
        report["errors"].append("проект уже содержит текст: " + ", ".join(what)
                                + " — используйте --update для правок из Word или --force (с резервной копией; "
                                "реестр со структурированными или проверенными записями --force не заменяет, "
                                "для замены добавьте --replace-sources)")
        return 2
    problems = bibliography_problems(bibliography)
    if problems:
        report["sources"] = {"file": "sources.json", "imported": 0, "docx_entries": len(bibliography["records"]),
                             "problems": problems}
        report["errors"].append(
            "номера списка литературы DOCX неоднозначны: " + "; ".join(problems)
            + ". Импорт остановлен, файлы проекта не изменены. Исправь список в копии DOCX (у каждой записи свой "
              "номер; разорванную Enter запись склей; подзаголовок оформи стилем «Заголовок 2») и повтори импорт; "
              "разбор без записи в проект — verify_sources.py --extract")
        report.update(status="fail", exit_code=1)
        return 1
    report["warnings"].extend(bibliography_notes(bibliography))
    protection = registry_protection(sources_path) if registry_filled and not replace_sources else None
    backup = backup_writer(root, report)
    try:
        write_pending_images(ctx)
        report["figures"] = list(ctx["new_figures"])
        for path in existing:
            rel = f"drafts/{path.name}"
            backup(rel)
            if rel not in outputs:
                path.unlink()
                report["files"]["removed"].append(rel)
        for rel, (file_kind, blocks, section) in outputs.items():
            target = root / rel
            target.parent.mkdir(parents=True, exist_ok=True)
            existed = target.is_file()
            write_lf(target, section_markdown(section, blocks))
            report["files"]["updated" if existed else "created"].append(rel)
        records = bibliography["records"]
        entries = [{"id": record["id"], "raw_text": record["raw_text"], "status": "pending"} for record in records]
        if protection is None:
            if sources_path.is_file():
                backup("sources.json")
            sources_path.write_text(json.dumps(entries, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            report["sources"] = {"file": "sources.json", "imported": len(entries)}
            if not bibliography["found"]:
                report["warnings"].append("список литературы в DOCX не найден — sources.json пуст")
        else:
            # Реестр со структурированными/проверенными записями --force не заменяет (как verify_sources --extract).
            report["sources"] = {"file": "sources.json", "imported": 0, "registry_preserved": True,
                                 "reason": protection, "docx_entries": len(entries)}
            if entries:
                backup(EXTRACTED_SOURCES)
                common.atomic_write_text(root / EXTRACTED_SOURCES, common.dump_json(records))
                report["sources"]["extracted"] = EXTRACTED_SOURCES
                report["warnings"].append(
                    f"sources.json не заменён: {protection}. Список литературы DOCX ({len(entries)} записей) записан в "
                    f"{EXTRACTED_SOURCES}; ссылки [N] в черновиках — номера этого списка, а не id реестра: сверь записи "
                    "с реестром (source-verification.md → «Перенос извлечённых записей в реестр») или замени реестр: "
                    "--force --replace-sources (прежний сохранится в резервной копии)")
            else:
                report["warnings"].append(f"sources.json не заменён: {protection}; список литературы в DOCX не найден")
    except OSError as error:
        report["errors"].append(f"ошибка записи: {error}")
        return 2
    return None


def _import_update(root: Path, docx_path: Path, report: Dict[str, Any],
                   outputs: Dict[str, Tuple[str, List[Dict[str, Any]], Dict[str, Any]]], bibliography: Dict[str, Any],
                   citation_map: Dict[str, Any], registry_ids: List[Any], ctx: Dict[str, Any],
                   overwrite_drafts: bool) -> Optional[int]:
    """Правки из Word → черновики. До записи: список литературы DOCX сверяется с картой последней
    сборки, черновики — с build-manifest (изменённые после сборки не затираются без --overwrite-drafts)."""
    report["sources"] = {"file": "sources.json", "imported": 0, "note": "реестр не изменяется в режиме --update",
                         "docx_entries": len(bibliography["records"])}
    registry_problems = citation_map_registry_problems(root, citation_map)
    if registry_problems:
        report["sources"]["citation_map_problems"] = registry_problems
        shown = "; ".join(registry_problems[:5]) + (f"; и ещё {len(registry_problems) - 5}"
                                                    if len(registry_problems) > 5 else "")
        report["errors"].append(
            "карта ссылок не соответствует реестру: пересобери DOCX (build_vkr.py). Расхождения: " + shown
            + ". Черновики не изменены: номера [N] из этого DOCX указали бы не на те источники"
        )
        report.update(status="fail", exit_code=1)
        return 1
    differences = bibliography_differences(bibliography, citation_map)
    if differences:
        report["sources"]["differences"] = differences
        shown = "; ".join(differences[:5]) + (f"; и ещё {len(differences) - 5}" if len(differences) > 5 else "")
        report["errors"].append(
            f"{BIBLIOGRAPHY_MISMATCH}. Расхождения: {shown}. Черновики не изменены: номера [N] из этого DOCX "
            "указали бы не на те источники. При пересборке DOCX с правками сохранится в "
            "backups/checkpoints/<UTC>-build/vkr.docx — оттуда правки и переносятся")
        report.update(status="fail", exit_code=1)
        return 1

    by_number = {str(k): v for k, v in citation_map["numbers"].items()}
    number_of = {str(v): int(k) for k, v in citation_map["numbers"].items() if str(k).isdigit()}
    ids_text = {str(v) for v in registry_ids}
    int_ids = {int(str(v)): v for v in registry_ids if str(v).isdigit()}

    def order(ident: Any) -> Tuple[int, str]:
        return number_of.get(str(ident), 10 ** 9), str(ident)

    def canon_new(text: str) -> str:
        return canonical_citations(text, lambda kind, ref: by_number.get(str(ref)) if kind == "num" else (ref if str(ref) in ids_text else None), order)

    def canon_old(text: str) -> str:
        return canonical_citations(text, lambda kind, ref: (ref if str(ref) in ids_text else None) if kind == "id" else int_ids.get(int(ref)), order)

    unresolved = set()
    for rel, (file_kind, blocks, _section) in outputs.items():
        for block in blocks:
            for text in build_vkr.block_texts(block):
                for kind, ref in fb.citation_refs(text):
                    if kind == "num" and str(ref) not in by_number:
                        unresolved.add(str(ref))
    if unresolved:
        report["warnings"].append("в DOCX есть номера ссылок без записи в citation-map.json: "
                                  + ", ".join(sorted(unresolved, key=int)) + " — они оставлены числами")

    planned: Dict[str, Tuple[str, Optional[str]]] = {}
    try:
        for rel, (file_kind, blocks, section) in outputs.items():
            target = root / rel
            if not target.is_file():
                converted = [map_block_texts(b, canon_new) for b in blocks]
                planned[rel] = ("created", section_markdown(section, converted))
                continue
            old_text = target.read_text(encoding="utf-8-sig")
            heading_block = {"kind": "heading", "level": 1, "text": section_markdown(section, []).strip()[2:]}
            new_text, changed = merge_draft(old_text, [heading_block] + blocks, file_kind, canon_old, canon_new)
            planned[rel] = ("updated", new_text) if changed and new_text != old_text else ("unchanged", None)
    except (OSError, UnicodeDecodeError) as error:
        report["errors"].append(f"не удалось прочитать черновик: {error}")
        return 2

    # Перезапись существующего черновика — только если он не менялся после сборки. Отсутствующий
    # черновик создаётся из DOCX без проверки: терять нечего (восстановление, failure-recovery.md).
    updates = [rel for rel, (action, _text) in planned.items() if action == "updated"]
    creates = [rel for rel, (action, _text) in planned.items() if action == "created"]
    build_id = str(citation_map.get("build_id"))
    blockers: List[str] = []
    manifest, manifest_problem = load_build_manifest(root) if updates or creates else (None, "")
    if updates:
        if manifest is None:
            blockers.append(f"{manifest_problem} — нельзя проверить, менялись ли после сборки черновики: {', '.join(updates)}")
        elif str(manifest.get("build_id") or "") != build_id:
            blockers.append("exports/citation-map.json и exports/build-manifest.json записаны разными сборками — нельзя "
                            f"проверить, менялись ли после сборки черновики: {', '.join(updates)}")
        else:
            changed_drafts = drafts_changed_after_build(root, updates, manifest, load_import_sync(root, build_id))
            if changed_drafts:
                report["drafts_changed_after_build"] = changed_drafts
                blockers.append("черновики изменены после сборки: " + ", ".join(changed_drafts))
    if creates and manifest is not None:
        deleted = [rel for rel in drafts_changed_after_build(root, creates, manifest, {}) if rel.endswith("(удалён после сборки)")]
        if deleted:
            report["warnings"].append("черновики удалены после сборки и созданы заново из DOCX: " + ", ".join(deleted))
    if blockers and not overwrite_drafts:
        report["errors"].append(
            "; ".join(blockers) + " — импорт затёр бы эти правки, черновики не изменены. Если главная правка в "
            "черновике — пересобери DOCX (build_vkr.py) и повтори правки в Word по новой сборке; если главная правка "
            "в Word — повтори импорт с --overwrite-drafts (прежние черновики уйдут в backups/checkpoints/<UTC>-import/)")
        report.update(status="fail", exit_code=1)
        return 1
    if blockers:
        report["warnings"].append("--overwrite-drafts: " + "; ".join(blockers)
                                  + " — записана версия из DOCX, прежние черновики в резервной копии")

    backup = backup_writer(root, report)
    drafts_dir = root / "drafts"
    try:
        write_pending_images(ctx)
        report["figures"] = list(ctx["new_figures"])
        for rel, (action, text) in planned.items():
            if action == "unchanged":
                report["files"]["unchanged"].append(rel)
                continue
            target = root / rel
            if action == "updated":
                backup(rel)
            target.parent.mkdir(parents=True, exist_ok=True)
            write_lf(target, str(text))
            report["files"][action].append(rel)
        for path in sorted(drafts_dir.glob("*.md")) if drafts_dir.is_dir() else []:
            rel = f"drafts/{path.name}"
            if rel not in outputs:
                report["warnings"].append(f"{rel}: раздела нет в DOCX — черновик не изменён")
    except OSError as error:
        report["errors"].append(f"ошибка записи: {error}")
        return 2
    try:
        write_import_sync(root, build_id, list(planned), docx_path)
    except OSError as error:
        report["warnings"].append(f"не удалось записать {IMPORT_SYNC_PATH} ({error}): повторный --update до пересборки "
                                  "может потребовать --overwrite-drafts")
    return None


def run_import(root: Path, docx_path: Path, *, update: bool, force: bool,
               replace_sources: bool = False, overwrite_drafts: bool = False) -> Tuple[int, Dict[str, Any]]:
    report: Dict[str, Any] = {
        "tool": "import_docx", "version": IMPORT_VERSION, "mode": "update" if update else "new",
        "project": str(root), "docx": str(docx_path), "status": "error", "exit_code": 2,
        "files": {"created": [], "updated": [], "unchanged": [], "removed": []},
        "sections": [], "sources": {}, "bibliography": {}, "figures": [], "backup": None, "skipped": [],
        "unrecognized": [], "warnings": [], "errors": [],
    }
    if update and replace_sources:
        report["errors"].append("--replace-sources не сочетается с --update: в режиме --update реестр источников не меняется")
        return 2, report
    if overwrite_drafts and not update:
        report["errors"].append("--overwrite-drafts работает только с --update: новый импорт заменяет черновики "
                                "только с --force")
        return 2, report
    if not root.is_dir():
        report["errors"].append(f"каталог проекта не найден: {root}")
        return 2, report
    if not (root / "vkr-project.json").is_file():
        report["errors"].append("нет vkr-project.json — сначала init_vkr_project.py")
        return 2, report
    if not docx_path.is_file():
        report["errors"].append(f"DOCX не найден: {docx_path}")
        return 2, report
    try:
        from docx import Document
    except ImportError:
        report["errors"].append("не установлен python-docx: pip install python-docx")
        return 2, report
    try:
        doc = Document(str(docx_path))
    except Exception as error:  # битый DOCX
        report["errors"].append(f"не удалось открыть DOCX: {type(error).__name__}: {error}")
        return 2, report

    citation_map: Dict[str, Any] = {}
    registry_ids: List[Any] = []
    if update:
        loaded_map, message, map_code = load_citation_map(root)
        if loaded_map is None:
            report["errors"].append(message)
            if map_code == 1:
                report.update(status="fail", exit_code=1)
            return map_code, report
        citation_map = loaded_map
        try:
            loaded = json.loads((root / "sources.json").read_text(encoding="utf-8-sig")) if (root / "sources.json").is_file() else []
            if isinstance(loaded, dict):
                loaded = loaded.get("sources", [])
            registry_ids = [src.get("id", n) for n, src in enumerate(loaded, 1) if isinstance(src, dict)]
        except (OSError, ValueError) as error:
            report["errors"].append(f"не удалось прочитать sources.json: {error}")
            return 2, report

    reader = DocxReader(doc)
    items = reader.items()
    sections = split_sections(items, report)
    ctx: Dict[str, Any] = {"root": root, "warnings": report["warnings"], "unrecognized": reader.unrecognized,
                           "new_figures": [], "pending_images": {}}
    if not sections:
        report["errors"].append("в DOCX не найдено ни одного раздела (Введение, Глава, Заключение…)")
        return 2, report
    bibliography = parse_docx_bibliography(sections)
    report["bibliography"] = {"found": bibliography["found"], "entries": len(bibliography["records"]),
                              "warnings": bibliography["warnings"]}

    outputs: Dict[str, Tuple[str, List[Dict[str, Any]], Dict[str, Any]]] = {}
    for section in sections:
        if section["kind"] == "contents":
            leftovers = [it for it in section["items"] if it["type"] == "p" and it["text"].strip()]
            if leftovers:
                report["warnings"].append(f"в разделе «Содержание» пропущено {len(leftovers)} абзацев без стиля оглавления")
            continue
        if section["kind"] == "bibliography":
            continue
        blocks = items_to_blocks(section, reader, ctx)
        located = section_file(section)
        if located is None:
            continue
        rel, file_kind = located
        if rel in outputs:
            report["warnings"].append(f"раздел «{section['heading'][:40]}» встречается повторно — объединён с {rel}")
            outputs[rel][1].extend(blocks)
            continue
        outputs[rel] = (file_kind, blocks, section)
    for entry in reader.unrecognized:
        report["unrecognized"].append(entry)
    for message in reader.warnings:
        report["warnings"].append(message)
    if "drafts/annotation.md" not in outputs and not update:
        outputs["drafts/annotation.md"] = ("annotation", [{"kind": "paragraph", "text": "[ЗАПОЛНИТЬ: аннотация не найдена в DOCX]"}],
                                           {"kind": "annotation", "title": "", "heading": ""})
        report["warnings"].append("аннотация не найдена — создан drafts/annotation.md с маркером")
    for rel, (file_kind, blocks, section) in outputs.items():
        report["sections"].append({
            "file": rel, "heading": section.get("heading", ""),
            "paragraphs": sum(1 for b in blocks if b["kind"] == "heading" and b["level"] == 2
                              and not CONCLUSION_HEADING_RE.match(b["text"])),
            "blocks": len(blocks),
            "citations": sum(len(fb.citation_refs(t)) for b in blocks for t in build_vkr.block_texts(b)),
        })

    if update:
        stopped = _import_update(root, docx_path, report, outputs, bibliography, citation_map, registry_ids, ctx,
                                 overwrite_drafts)
    else:
        stopped = _import_new(root, report, outputs, bibliography, ctx, force, replace_sources)
    if stopped is not None:
        return stopped, report

    if report["figures"]:
        register_figures(root, report)
    lost = [u for u in report["unrecognized"] if u.get("lost_content")]
    code = 1 if lost else 0
    report["status"] = "incomplete" if lost else "ok"
    report["exit_code"] = code
    return code, report


def _configure_streams() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except (ValueError, OSError):
                pass


def main(argv: Optional[List[str]] = None) -> int:
    _configure_streams()
    parser = argparse.ArgumentParser(description="Импорт DOCX в черновики проекта ВКР")
    parser.add_argument("project_dir", help="Каталог проекта ВКР (после init)")
    parser.add_argument("--docx", required=True, help="Импортируемый DOCX")
    parser.add_argument("--update", action="store_true", help="Перенести правки из собранного DOCX в существующие черновики")
    parser.add_argument("--force", action="store_true",
                        help="Перезаписать черновики (с резервной копией); sources.json заменяется, только если в нём "
                             "нет структурированных или проверенных записей, иначе список литературы DOCX пишется "
                             "в audit/sources-extracted.json")
    parser.add_argument("--replace-sources", action="store_true",
                        help="Вместе с --force: заменить sources.json записями raw_text из DOCX, даже если в нём есть "
                             "структурированные или проверенные записи (прежний реестр — в резервной копии)")
    parser.add_argument("--overwrite-drafts", action="store_true",
                        help="Вместе с --update: записать правки из Word и в черновики, изменённые после сборки "
                             "(прежние черновики — в backups/checkpoints/<UTC>-import/)")
    parser.add_argument("--json", action="store_true", help="Отчёт в stdout как JSON")
    parser.add_argument("-o", dest="report", help="Дополнительно записать отчёт в файл")
    args = parser.parse_args(argv)
    root = Path(args.project_dir)
    code, report = run_import(root, Path(args.docx), update=args.update, force=args.force,
                              replace_sources=args.replace_sources, overwrite_drafts=args.overwrite_drafts)
    payload = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    targets = []
    if root.is_dir():
        targets.append(root / "audit" / "import-report.json")
    if args.report:
        targets.append(Path(args.report))
    for target in targets:
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(payload, encoding="utf-8")
        except OSError as error:
            print(f"ОШИБКА: не удалось записать отчёт {target}: {error}", file=sys.stderr)
            code = 2
    if args.json:
        print(payload, end="")
        return code
    for message in report["errors"]:
        print(f"ОШИБКА: {message}", file=sys.stderr)
    for message in report["warnings"]:
        print(f"ПРЕДУПРЕЖДЕНИЕ: {message}", file=sys.stderr)
    for entry in report["unrecognized"]:
        mark = "ПОТЕРЯ СОДЕРЖИМОГО" if entry.get("lost_content") else "НЕ РАСПОЗНАНО"
        print(f"{mark}: {entry['kind']}: {entry.get('text', '')}", file=sys.stderr)
    if report.get("status") == "fail":
        print("ИМПОРТ ОСТАНОВЛЕН: черновики и реестр не изменены; причины — выше и в audit/import-report.json",
              file=sys.stderr)
    elif code != 2:
        files = report["files"]
        print(f"OK: создано {len(files['created'])}, обновлено {len(files['updated'])}, "
              f"без изменений {len(files['unchanged'])}; отчёт audit/import-report.json")
    return code


if __name__ == "__main__":
    raise SystemExit(main())
