#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx_integrity.py -- структурная проверка пакета .docx без Microsoft Word.

Ловит то, из-за чего Word отказывается открыть файл («Файл поврежден»), хотя
python-docx и validate_vkr.py его спокойно читают:

1. каждая XML-часть пакета (*.xml, *.rels) -- корректный XML без DTD;
2. в каждой XML-части объявлены все префиксы имён элементов и атрибутов, а
   каждый префикс, перечисленный в mc:Ignorable, mc:ProcessContent,
   mc:MustUnderstand и mc:Choice/@Requires, объявлен в области видимости
   элемента (Word проверяет это строго: префикс из mc:Ignorable без
   объявления xmlns -- повреждённый файл);
3. у каждой части есть тип содержимого в [Content_Types].xml -- Override по
   имени части или Default по расширению;
4. каждая внутренняя связь в *.rels указывает на существующую часть
   (связи TargetMode="External" не проверяются);
5. в пакете есть главный документ (связь officeDocument в _rels/.rels) и нет
   частей, имена которых совпадают без учёта регистра;
6. в частях с текстом (document.xml, колонтитулы, сноски, комментарии) частые
   элементы WordprocessingML стоят на своём месте: w:t, w:br, w:sym, w:drawing --
   внутри w:r; w:tab -- внутри w:r или w:tabs; w:r -- внутри w:p, w:hyperlink,
   w:sdtContent и других контейнеров прогонов; w:p -- не внутри w:r или w:p;
   w:tc -- внутри w:tr, w:tr -- внутри w:tbl, w:tbl -- не внутри w:r. Такой файл
   python-docx читает, а Word отвечает «Ошибка при попытке открытия файла».

Проверка структурная: она не заменяет открытие файла в Word, но находит
известные причины отказа Word без запуска Word.

ИСПОЛЬЗОВАНИЕ
=============
    python docx_integrity.py final/vkr.docx
    python docx_integrity.py final/vkr.docx --json

КОДЫ ВЫХОДА
===========
0 -- проблем не найдено;
1 -- найдены структурные проблемы пакета;
2 -- файл не найден, не читается или не является ZIP-архивом.

API
===
    docx_package_problems(path) -> list[str]        # пустой список -- проблем нет
    package_problems_from_parts(parts) -> list[str] # parts: {имя записи ZIP: bytes}

Только стандартная библиотека Python (python-docx не нужен).
"""
import argparse
import json
import posixpath
import re
import sys
import zipfile
from collections import OrderedDict
from pathlib import Path
from urllib.parse import unquote
from xml.parsers import expat

MCE_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"
XML_NS = "http://www.w3.org/XML/1998/namespace"
XMLNS_NS = "http://www.w3.org/2000/xmlns/"

CONTENT_TYPES = "[Content_Types].xml"
MAX_MESSAGES_PER_PART = 20


class PackageReadError(Exception):
    """Файл не найден, не читается или не является ZIP-архивом."""


class _DoctypeFound(Exception):
    pass


# --------------------------------------------------------------------------
# Чтение пакета
# --------------------------------------------------------------------------

def read_package(path):
    """Читает ZIP в OrderedDict {имя записи: bytes}. Повторяющиеся имена
    записей сохраняются как отдельная проблема (см. package_problems_from_parts):
    для этого возвращается ещё и список всех имён в порядке архива."""
    path = Path(path)
    if not path.is_file():
        raise PackageReadError("файл не найден: %s" % path)
    try:
        with zipfile.ZipFile(str(path), "r") as archive:
            names = [info.filename for info in archive.infolist()]
            parts = OrderedDict()
            for name in names:
                if name not in parts:
                    parts[name] = archive.read(name)
    except zipfile.BadZipFile as exc:
        raise PackageReadError("не ZIP-архив (повреждённый .docx): %s" % exc)
    except (OSError, RuntimeError, ValueError) as exc:
        raise PackageReadError("не удалось прочитать файл: %s" % exc)
    return parts, names


# --------------------------------------------------------------------------
# XML-части: корректность и объявления префиксов
# --------------------------------------------------------------------------

def _is_xml_part(name):
    lower = name.lower()
    return lower.endswith(".xml") or lower.endswith(".rels")


def xml_part_problems(name, data):
    """Проблемы одной XML-части: разбор, DTD, необъявленные префиксы."""
    problems = []
    seen = set()

    def add(message):
        if message in seen:
            return
        seen.add(message)
        problems.append(message)

    # Области видимости: список словарей префикс -> URI. Разбор без обработки
    # пространств имён expat'ом, чтобы видеть исходные префиксы и xmlns-атрибуты.
    scopes = [{"xml": XML_NS, "xmlns": XMLNS_NS}]

    def start(tag, attrs):
        parent = scopes[-1]
        current = parent
        for key, value in attrs.items():
            if key == "xmlns" or key.startswith("xmlns:"):
                if current is parent:
                    current = dict(parent)
                current[key[6:] if key.startswith("xmlns:") else ""] = value
        scopes.append(current)

        element_prefix = tag.split(":", 1)[0] if ":" in tag else None
        if element_prefix is not None and element_prefix not in current:
            add("%s: префикс «%s» элемента <%s> не объявлен (xmlns:%s)"
                % (name, element_prefix, tag, element_prefix))

        for key, value in attrs.items():
            if key == "xmlns" or key.startswith("xmlns:"):
                continue
            if ":" not in key:
                continue
            prefix, local = key.split(":", 1)
            if prefix not in current:
                add("%s: префикс «%s» атрибута %s не объявлен (xmlns:%s)" % (name, prefix, key, prefix))
                continue
            uri = current[prefix]
            if uri == MCE_NS and local in ("Ignorable", "MustUnderstand", "ProcessContent"):
                for token in value.split():
                    ref = token.split(":", 1)[0] if local == "ProcessContent" else token
                    if ref and ref not in current:
                        add("%s: префикс «%s» из mc:%s не объявлен в области видимости <%s> (xmlns:%s)"
                            % (name, ref, local, tag, ref))
            elif uri == XSI_NS and local == "type" and ":" in value:
                ref = value.split(":", 1)[0]
                if ref not in current:
                    add("%s: префикс «%s» в значении xsi:type=\"%s\" не объявлен" % (name, ref, value))

        if element_prefix is not None and tag.split(":", 1)[1] == "Choice" \
                and current.get(element_prefix) == MCE_NS and "Requires" in attrs:
            for token in attrs["Requires"].split():
                if token not in current:
                    add("%s: префикс «%s» из mc:Choice/@Requires не объявлен (xmlns:%s)"
                        % (name, token, token))

    def end(_tag):
        scopes.pop()

    def doctype(*_args):
        raise _DoctypeFound()

    parser = expat.ParserCreate()
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.StartDoctypeDeclHandler = doctype
    try:
        parser.Parse(data, True)
    except _DoctypeFound:
        return ["%s: объявление DTD (<!DOCTYPE>) в части OOXML недопустимо" % name]
    except expat.ExpatError as exc:
        return ["%s: некорректный XML (%s)" % (name, exc)]
    if len(problems) > MAX_MESSAGES_PER_PART:
        extra = len(problems) - MAX_MESSAGES_PER_PART
        problems = problems[:MAX_MESSAGES_PER_PART] + ["%s: и ещё %d сообщений того же рода" % (name, extra)]
    return problems


# --------------------------------------------------------------------------
# Место элемента в дереве WordprocessingML
# --------------------------------------------------------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Части, в которых лежит размеченный текст: их схему Word проверяет при открытии.
BODY_PART_RE = re.compile(
    r"^word/(document\d*\.xml|header\d*\.xml|footer\d*\.xml|footnotes\.xml|endnotes\.xml|comments\.xml)$",
    re.IGNORECASE,
)

# Допустимые родители элемента (локальные имена в пространстве w:). Списки шире
# минимально необходимого: сюда входят все контейнеры схемы WordprocessingML, в
# которых элемент встречается законно, — проверка ищет грубые нарушения
# («висячий» w:t прямо в w:p), а не неполноту разметки.
_RUN_PARENTS = (
    "p", "hyperlink", "sdtContent", "smartTag", "ins", "del", "fldSimple", "customXml",
    "moveTo", "moveFrom", "dir", "bdo", "rt", "rubyBase",
)
ALLOWED_PARENTS = {
    # содержимое текстового прогона — только внутри w:r
    "t": ("r",),
    "br": ("r",),
    "sym": ("r",),
    "drawing": ("r",),
    # w:tab — и содержимое прогона, и позиция табуляции в w:pPr/w:tabs (оглавление)
    "tab": ("r", "tabs"),
    "r": _RUN_PARENTS,
    "tc": ("tr", "sdtContent", "customXml"),
    "tr": ("tbl", "sdtContent", "customXml"),
}
# Родители, которых у элемента быть не может (остальные допустимы).
FORBIDDEN_PARENTS = {
    "p": ("r", "p"),
    "tbl": ("r",),
}


def element_placement_problems(name, data):
    """Элементы WordprocessingML не на своём месте (Word: «Ошибка при попытке открытия файла»).

    Проверяется только положение частых элементов относительно родителя: python-docx
    и валидатор такой файл читают, а Word отказывается его открыть.
    """
    problems = []
    seen = set()
    stack = []

    prefix = W_NS + "}"

    def start(tag, _attrs):
        parent = stack[-1] if stack else None
        # Родитель из чужого пространства имён (mc:Choice, mc:Fallback, w14:…, корень части):
        # его правила вложенности этой проверке неизвестны — не трогаем.
        parent_local = parent[len(prefix):] if parent and parent.startswith(prefix) else None
        if tag.startswith(prefix) and parent_local is not None:
            local = tag[len(prefix):]
            parent_name = "w:" + parent_local
            allowed = ALLOWED_PARENTS.get(local)
            forbidden = FORBIDDEN_PARENTS.get(local)
            bad = False
            if allowed is not None and parent_local not in allowed:
                bad = True
                expected = "внутри " + ", ".join("<w:%s>" % item for item in allowed)
            elif forbidden is not None and parent_local in forbidden:
                bad = True
                expected = "не внутри " + ", ".join("<w:%s>" % item for item in forbidden)
            if bad:
                message = ("%s: <w:%s> стоит внутри <%s> — по схеме WordprocessingML он должен быть %s; "
                           "Word откажется открыть такой файл (пересобери DOCX)" % (name, local, parent_name, expected))
                if message not in seen:
                    seen.add(message)
                    problems.append(message)
        stack.append(tag)

    def end(_tag):
        stack.pop()

    def doctype(*_args):
        raise _DoctypeFound()

    parser = expat.ParserCreate(namespace_separator="}")
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.StartDoctypeDeclHandler = doctype
    try:
        parser.Parse(data, True)
    except (expat.ExpatError, _DoctypeFound):
        return []  # некорректный XML уже отмечен xml_part_problems
    if len(problems) > MAX_MESSAGES_PER_PART:
        extra = len(problems) - MAX_MESSAGES_PER_PART
        problems = problems[:MAX_MESSAGES_PER_PART] + ["%s: и ещё %d сообщений того же рода" % (name, extra)]
    return problems


def _parse_simple(data):
    """Разбор [Content_Types].xml или *.rels: возвращает список (локальное имя
    элемента, атрибуты) для всех элементов. None -- XML некорректен."""
    items = []

    def start(tag, attrs):
        local = tag.rsplit("}", 1)[-1]
        items.append((local, attrs))

    def doctype(*_args):
        raise _DoctypeFound()

    parser = expat.ParserCreate(namespace_separator="}")
    parser.StartElementHandler = start
    parser.StartDoctypeDeclHandler = doctype
    try:
        parser.Parse(data, True)
    except (expat.ExpatError, _DoctypeFound):
        return None
    return items


# --------------------------------------------------------------------------
# [Content_Types].xml и связи
# --------------------------------------------------------------------------

def _norm_part(name):
    return unquote(name.lstrip("/")).lower()


def _relationship_problems(rels_name, data, existing):
    problems = []
    items = _parse_simple(data)
    if items is None:
        return problems  # сама некорректность XML уже отмечена xml_part_problems
    folder, _file = posixpath.split(rels_name)
    if posixpath.basename(folder).lower() != "_rels":
        return problems
    source_dir = posixpath.dirname(folder)  # "" для _rels/.rels
    for local, attrs in items:
        if local != "Relationship":
            continue
        if (attrs.get("TargetMode") or "").lower() == "external":
            continue
        rel_id = attrs.get("Id", "?")
        rel_type = (attrs.get("Type") or "").rsplit("/", 1)[-1] or "?"
        target = attrs.get("Target")
        if not target:
            problems.append("%s: связь %s (%s) без Target" % (rels_name, rel_id, rel_type))
            continue
        target_path = target.split("#", 1)[0]
        if not target_path:
            continue
        if target_path.startswith("/"):
            resolved = posixpath.normpath(target_path.lstrip("/"))
        else:
            resolved = posixpath.normpath(posixpath.join(source_dir, target_path)) if source_dir \
                else posixpath.normpath(target_path)
        if resolved == ".." or resolved.startswith("../"):
            problems.append("%s: связь %s (%s) указывает за пределы пакета: %s"
                            % (rels_name, rel_id, rel_type, target))
            continue
        if _norm_part(resolved) not in existing:
            problems.append("%s: связь %s (%s) указывает на отсутствующую часть %s"
                            % (rels_name, rel_id, rel_type, resolved))
    return problems


def package_problems_from_parts(parts, all_names=None):
    """parts: {имя записи ZIP: bytes}; all_names -- все имена в порядке архива
    (с повторами, если они были). Возвращает список проблем."""
    problems = []
    names = list(all_names) if all_names is not None else list(parts.keys())

    seen_exact = set()
    seen_folded = {}
    for name in names:
        if name in seen_exact:
            problems.append("в ZIP повторяется запись %s" % name)
            continue
        seen_exact.add(name)
        folded = _norm_part(name)
        if folded in seen_folded and not name.endswith("/"):
            problems.append("части %s и %s совпадают без учёта регистра" % (seen_folded[folded], name))
        seen_folded.setdefault(folded, name)

    file_names = [n for n in parts if not n.endswith("/")]
    existing = set(_norm_part(n) for n in file_names)

    for name in file_names:
        if _is_xml_part(name):
            problems.extend(xml_part_problems(name, parts[name]))
        if BODY_PART_RE.match(name.replace("\\", "/")):
            problems.extend(element_placement_problems(name, parts[name]))

    # [Content_Types].xml
    if CONTENT_TYPES not in parts:
        problems.append("нет [Content_Types].xml")
    else:
        items = _parse_simple(parts[CONTENT_TYPES])
        if items is not None:
            defaults = set()
            overrides = set()
            for local, attrs in items:
                if local == "Default" and attrs.get("Extension"):
                    defaults.add(attrs["Extension"].lower())
                elif local == "Override" and attrs.get("PartName"):
                    overrides.add(_norm_part(attrs["PartName"]))
            for name in file_names:
                if name == CONTENT_TYPES:
                    continue
                if _norm_part(name) in overrides:
                    continue
                base = posixpath.basename(name)
                ext = base.rsplit(".", 1)[-1].lower() if "." in base else ""
                if ext and ext in defaults:
                    continue
                problems.append("часть %s: нет типа содержимого в [Content_Types].xml "
                                "(ни Override, ни Default для расширения «%s»)" % (name, ext))

    # связи
    for name in file_names:
        if name.lower().endswith(".rels"):
            problems.extend(_relationship_problems(name, parts[name], existing))

    # главный документ
    root_rels = None
    for name in file_names:
        if name.lower() == "_rels/.rels":
            root_rels = name
            break
    if root_rels is None:
        problems.append("нет _rels/.rels (связей пакета)")
    else:
        items = _parse_simple(parts[root_rels]) or []
        if not any(local == "Relationship"
                   and (attrs.get("Type") or "").rsplit("/", 1)[-1] == "officeDocument"
                   for local, attrs in items):
            problems.append("_rels/.rels: нет связи officeDocument (главного документа)")
    return problems


def docx_package_problems(path):
    """Список структурных проблем пакета .docx (пустой -- проблем нет).
    Нечитаемый файл или не ZIP -- тоже проблема (одна строка)."""
    try:
        parts, names = read_package(path)
    except PackageReadError as exc:
        return [str(exc)]
    return package_problems_from_parts(parts, names)


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def _configure_output():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except (ValueError, OSError):
                pass


def main(argv=None):
    _configure_output()
    parser = argparse.ArgumentParser(
        prog="docx_integrity.py",
        description="Структурная проверка пакета .docx: XML, объявления префиксов "
                    "(mc:Ignorable и др.), [Content_Types].xml, связи .rels.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("docx", type=Path, help="Проверяемый .docx")
    parser.add_argument("--json", action="store_true", help="Печатать только JSON в stdout")
    args = parser.parse_args(argv)

    try:
        parts, names = read_package(args.docx)
    except PackageReadError as exc:
        if args.json:
            print(json.dumps({"status": "error", "file": str(args.docx), "error": str(exc)},
                             ensure_ascii=False, indent=2))
        else:
            print("ОШИБКА: %s" % exc, file=sys.stderr)
        return 2

    problems = package_problems_from_parts(parts, names)
    if args.json:
        print(json.dumps({"status": "problems" if problems else "ok", "file": str(args.docx),
                          "problems": problems}, ensure_ascii=False, indent=2))
    elif problems:
        print("НАЙДЕНЫ ПРОБЛЕМЫ ПАКЕТА: %s" % args.docx)
        for item in problems:
            print("  - %s" % item)
    else:
        print("OK: структура пакета в порядке: %s" % args.docx)
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
