#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
clean_docx_metadata.py -- очищает внутренние метаданные .docx перед сдачей.

Реализация -- только zipfile + xml.etree.ElementTree из стандартной библиотеки
Python; python-docx не требуется и нигде не импортируется. Результат -- обычный
.docx: его открывают и Word, и python-docx (см. tests/test_templates_metadata_633.py
и tests/test_docx_integrity_633.py).

Переписываются только части, в которых что-то изменилось; остальные копируются
побайтно. В переписанных частях сохраняются все объявления xmlns исходника на
тех же элементах и с теми же префиксами -- в том числе префиксы, упомянутые
только в mc:Ignorable/mc:ProcessContent/mc:MustUnderstand: без их объявлений
Microsoft Word считает файл повреждённым. Перед записью результат проходит
структурную самопроверку scripts/docx_integrity.py (XML, объявления префиксов,
[Content_Types].xml, связи .rels); не прошёл -- файл не записывается.

ЧТО ОЧИЩАЕТСЯ
=============
- docProps/core.xml: creator, lastModifiedBy, title, subject, description,
  keywords, category, contentStatus -- обнуляются; revision=1; created и
  modified -- одно и то же время очистки (текущее время UTC с суффиксом Z;
  с --seed -- дата последнего сохранения из самого документа, если она там
  есть, иначе даты не меняются). Дата создания не сдвигается в прошлое и не
  выдумывается: скрипт не подделывает историю файла.
- docProps/app.xml: Company, Manager -- обнуляются; HyperlinkBase --
  удаляется; Template -- принудительно "Normal.dotm"; TotalTime -- "0".
- docProps/custom.xml удаляется целиком вместе со связью в _rels/.rels и
  Override в [Content_Types].xml (нестандартные custom properties).
- Комментарии: word/comments.xml, commentsExtended.xml, commentsIds.xml,
  commentsExtensible.xml, word/people.xml -- удаляются вместе со связями в
  word/_rels/document.xml.rels и Override в [Content_Types].xml; якоря
  commentRangeStart/commentRangeEnd/commentReference вычищаются из
  document.xml, колонтитулов и сносок.
- Исправления рецензирования (принимаются все, по всем частям word/*.xml):
  w:ins и w:moveTo -- обёртка снимается, содержимое остаётся; w:del и
  w:moveFrom -- удаляются вместе с содержимым; w:moveFromRangeStart/End и
  w:moveToRangeStart/End -- удаляются; pPrChange/rPrChange/sectPrChange/
  tblPrChange/tcPrChange/trPrChange/tblGridChange/numberingChange (история
  прежнего форматирования) -- удаляются. <w:trackChanges/> в settings.xml
  снимается, чтобы дальнейшее редактирование не писало новые правки поверх.
- --strip-rsid: атрибуты w:rsid* удаляются из всех элементов всех частей
  word/*.xml, список <w:rsids> в settings.xml удаляется целиком.

РЕЖИМЫ
======
    python clean_docx_metadata.py final/vkr.docx --in-place
    python clean_docx_metadata.py final/vkr.docx --check          # ничего не пишет
    python clean_docx_metadata.py final/vkr.docx -o exports/vkr-clean.docx
    python clean_docx_metadata.py final/vkr.docx -o exports/vkr-clean.docx --json
    python clean_docx_metadata.py final/vkr.docx --in-place --strip-rsid --seed s1

Пути — от корня проекта ВКР. Сдаётся единственный final/vkr.docx: его чистят
на месте (--in-place), копии пишут только в exports/. Без -o и --in-place
результат ложится рядом с входом (input-clean.docx); для файла в каталоге
final/ это запрещено (код 2): второй DOCX в final/ не пропустит финальная
проверка проекта.

-o/--in-place пишут через временный файл в целевом каталоге и атомарно
подменяют результат (os.replace) -- это одинаково безопасно и когда -o
совпадает со входным файлом, и когда совпадает с ним только по регистру
(на файловой системе Windows это один и тот же файл).

--seed делает результат детерминированным: created = modified = реальная
дата последнего сохранения из самого входного документа (если её нет --
даты не меняются), служебные даты записей ZIP -- та же дата (или 1980-01-01,
минимальная дата формата ZIP), порядок записей ZIP -- порядок входного
архива. Значение seed на содержимое не влияет и никаких дат не порождает. Одинаковые вход и seed дают побайтно одинаковый выход --
НО только в пределах одного окружения Python/zlib (см. «Ограничения»). Без
--seed created = modified = реальное текущее время UTC, и повторный запуск
в другую секунду даёт новый файл.

Если чистить нечего (поля core.xml уже пусты, нет comments/revisions/custom.xml,
created == modified), запись на месте (-o на тот же файл или --in-place) НЕ
выполняется: код 0, в отчёте "status": "already_clean", байты файла не меняются.
Иначе повторный запуск после финального снимка переписал бы даты, изменил sha
сдаваемого DOCX и обнулил волны аудита (AUDIT_SNAPSHOT_STALE) без единой правки
в работе. Запись в другой файл (-o exports/...) выполняется всегда.

С --seed дата последнего сохранения берётся из самого документа: если она
больше текущего времени (сбитые часы на компьютере, где сохраняли файл), в
отчёт и в вывод добавляется предупреждение. Даты при этом не меняются.

КОДЫ ВЫХОДА
===========
0 -- успех (очищено; чистить нечего -- "already_clean"; либо --check не нашёл
     метаданных для очистки);
1 -- --check нашёл метаданные для очистки, либо (в режиме очистки)
     самопроверка результата нашла остаток метаданных или структурную
     проблему пакета (docx_integrity.py) -- в этом случае выходной файл
     НЕ записывается/не подменяется, входной остаётся как был;
2 -- использование, ввод-вывод или повреждённый .docx (не ZIP, нет
     word/document.xml, конфликт -o/--in-place/--check, нет рядом
     docx_integrity.py и т. п.).

ОГРАНИЧЕНИЯ (честно)
=====================
- Отметка удаления абзацного знака (w:pPr/w:rPr/w:del, "удалить разрыв
  абзаца") снимается, но два абзаца НЕ объединяются -- Word при «Принять
  все исправления» их сливает, этот скрипт только убирает пометку. Для
  обычных вставленных/удалённых предложений и слов внутри абзаца это не
  встречается и роли не играет.
- Метаданные вложенных объектов (встроенные книги/таблицы OLE) не
  затрагиваются -- скрипт работает только с частями word/*.xml и docProps/*.
- История версий облачного хранилища (OneDrive/SharePoint) вне файла .docx
  и этим скриптом не очищается.
- Байт-в-байт воспроизводимость с --seed гарантирована для повторных
  запусков в одном и том же окружении Python/zlib; между разными сборками
  Python (например PY39 и PYCODEX) сжатие DEFLATE может отличаться побайтно
  даже при одинаковом содержимом.
"""
import argparse
import datetime
import hashlib
import importlib.util
import io
import json
import os
import re
import sys
import tempfile
import time
import zipfile
from collections import OrderedDict, defaultdict
from pathlib import Path

import xml.etree.ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"
EP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"


def w(local):
    return "{%s}%s" % (W_NS, local)


UNWRAP_TAGS = {w("ins"), w("moveTo")}
DROP_CONTENT_TAGS = {w("del"), w("moveFrom")}
DROP_ANCHOR_TAGS = {
    w("commentRangeStart"), w("commentRangeEnd"), w("commentReference"),
    w("moveFromRangeStart"), w("moveFromRangeEnd"),
    w("moveToRangeStart"), w("moveToRangeEnd"),
}
DROP_CHANGE_TAGS = {
    w("pPrChange"), w("rPrChange"), w("sectPrChange"),
    w("tblPrChange"), w("tcPrChange"), w("trPrChange"),
    w("tblGridChange"), w("numberingChange"),
}
ALL_TARGET_TAGS = UNWRAP_TAGS | DROP_CONTENT_TAGS | DROP_ANCHOR_TAGS | DROP_CHANGE_TAGS

COMMENT_PART_NAMES = [
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsIds.xml",
    "word/commentsExtensible.xml",
    "word/people.xml",
]

CORE_TEXT_FIELDS = [
    (DC_NS, "title"), (DC_NS, "subject"), (DC_NS, "creator"), (DC_NS, "description"),
    (CP_NS, "keywords"), (CP_NS, "lastModifiedBy"), (CP_NS, "category"),
    (CP_NS, "contentStatus"),
]


class DocxStructureError(Exception):
    """Входной файл -- не ZIP или не похож на .docx (нет word/document.xml)."""


# --------------------------------------------------------------------------
# XML: разбор с сохранением объявлений пространств имён + сериализация
# --------------------------------------------------------------------------
#
# ElementTree.tostring объявляет только те пространства имён, которые реально
# встречаются в именах элементов и атрибутов. Префиксы, которые упомянуты лишь
# в значениях атрибутов -- mc:Ignorable="w14 w15 w16se ...", mc:ProcessContent,
# mc:MustUnderstand, mc:Choice/@Requires, xsi:type="dcterms:W3CDTF", -- при
# этом теряют объявления, и Microsoft Word считает такой файл повреждённым
# (python-docx его при этом читает). Поэтому сериализация своя: каждое
# объявление xmlns выводится на том же элементе, где оно было в исходной
# части, с тем же префиксом; недостающие объявления (например, после снятия
# обёртки w:ins, на которой было локальное xmlns) дописываются на месте.

XML_NS = "http://www.w3.org/XML/1998/namespace"
MCE_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_MCE_PREFIX_LIST_ATTRS = {
    "{%s}Ignorable" % MCE_NS, "{%s}MustUnderstand" % MCE_NS, "{%s}ProcessContent" % MCE_NS,
}
_MCE_CHOICE = "{%s}Choice" % MCE_NS
_QNAME_VALUE_RE = re.compile(r"^([A-Za-z_][\w.\-]*):[A-Za-z_*][\w.\-*]*$")
_WELL_KNOWN_PREFIXES = {
    W_NS: "w",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships": "r",
    MCE_NS: "mc",
}


class _NsInfo(object):
    """Где и какие объявления xmlns были в исходной части."""

    __slots__ = ("by_elem", "prefix_hint", "uri_hint")

    def __init__(self):
        self.by_elem = {}                 # элемент -> [(префикс, URI), ...] в исходном порядке
        self.prefix_hint = OrderedDict()  # префикс -> URI (первое объявление в части)
        self.uri_hint = OrderedDict()     # URI -> непустой префикс (первое объявление)


def _parse_xml(data):
    """Разбирает часть и запоминает объявления пространств имён по элементам.
    Некорректный или пустой XML -- ET.ParseError, как и раньше."""
    info = _NsInfo()
    pending = []
    root = None
    for event, payload in ET.iterparse(io.BytesIO(data), events=("start-ns", "start")):
        if event == "start-ns":
            prefix, uri = payload
            prefix = prefix or ""
            pending.append((prefix, uri))
            info.prefix_hint.setdefault(prefix, uri)
            if prefix and uri not in info.uri_hint:
                info.uri_hint[uri] = prefix
        else:
            if pending:
                info.by_elem[payload] = pending
                pending = []
            if root is None:
                root = payload
    return root, info


class _Scope(object):
    """Область видимости префиксов при сериализации."""

    __slots__ = ("p2u", "_elem", "_attr")

    def __init__(self, p2u):
        self.p2u = p2u
        self._elem = None
        self._attr = None

    def extend(self, decls):
        p2u = dict(self.p2u)
        for prefix, uri in decls:
            p2u[prefix] = uri
        return _Scope(p2u)

    def _build(self):
        elem, attr = {}, {}
        for prefix, uri in self.p2u.items():
            if not uri:
                continue  # xmlns="" -- отмена пространства имён по умолчанию
            if prefix == "":
                elem[uri] = ""
            else:
                elem.setdefault(uri, prefix)
                attr.setdefault(uri, prefix)
        self._elem, self._attr = elem, attr

    def elem_prefix(self, uri):
        if self._elem is None:
            self._build()
        return self._elem.get(uri)

    def attr_prefix(self, uri):
        if self._attr is None:
            self._build()
        return self._attr.get(uri)


def _new_prefix(uri, scope, info):
    candidate = info.uri_hint.get(uri)
    if candidate and candidate not in scope.p2u:
        return candidate
    candidate = _WELL_KNOWN_PREFIXES.get(uri)
    if candidate and candidate not in scope.p2u:
        return candidate
    index = 0
    while "ns%d" % index in scope.p2u or "ns%d" % index in info.prefix_hint:
        index += 1
    return "ns%d" % index


def _referenced_prefixes(tag, key, value):
    if key in _MCE_PREFIX_LIST_ATTRS or (key == "Requires" and tag == _MCE_CHOICE):
        return [token.split(":", 1)[0] for token in value.split() if token]
    match = _QNAME_VALUE_RE.match(value)
    if match:
        return [match.group(1)]
    return ()


def _escape_text(text):
    return (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace("\r", "&#13;"))


def _escape_attr(value):
    return (value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;").replace("\n", "&#10;").replace("\r", "&#13;")
            .replace("\t", "&#9;"))


def _write_element(elem, parent_scope, info, out):
    tag = elem.tag
    if not isinstance(tag, str):  # комментарии/PI парсер ElementTree не сохраняет
        return
    decls = info.by_elem.get(elem, ())
    scope = parent_scope.extend(decls) if decls else parent_scope
    extra = []

    if tag[:1] == "{":
        uri, local = tag[1:].split("}", 1)
        prefix = scope.elem_prefix(uri)
        if prefix is None:
            prefix = _new_prefix(uri, scope, info)
            extra.append((prefix, uri))
            scope = scope.extend([(prefix, uri)])
        qname = "%s:%s" % (prefix, local) if prefix else local
    else:
        qname = tag
        if scope.p2u.get(""):
            extra.append(("", ""))
            scope = scope.extend([("", "")])

    attrs = []
    for key, value in elem.attrib.items():
        if key[:1] == "{":
            uri, local = key[1:].split("}", 1)
            prefix = scope.attr_prefix(uri)
            if prefix is None:
                prefix = _new_prefix(uri, scope, info)
                extra.append((prefix, uri))
                scope = scope.extend([(prefix, uri)])
            attrs.append(("%s:%s" % (prefix, local), value))
        else:
            attrs.append((key, value))

    # Префиксы, на которые ссылаются значения атрибутов (mc:Ignorable и т. п.),
    # должны быть объявлены в области видимости этого элемента.
    for key, value in elem.attrib.items():
        for ref in _referenced_prefixes(tag, key, value):
            if ref in ("xml", "xmlns") or ref in scope.p2u:
                continue
            uri = info.prefix_hint.get(ref)
            if uri:
                extra.append((ref, uri))
                scope = scope.extend([(ref, uri)])

    out.append("<" + qname)
    for prefix, uri in list(decls) + extra:
        if prefix:
            out.append(' xmlns:%s="%s"' % (prefix, _escape_attr(uri)))
        else:
            out.append(' xmlns="%s"' % _escape_attr(uri))
    for name, value in attrs:
        out.append(' %s="%s"' % (name, _escape_attr(value)))
    text = elem.text
    if len(elem) == 0 and not text:
        out.append("/>")
        return
    out.append(">")
    if text:
        out.append(_escape_text(text))
    for child in elem:
        _write_element(child, scope, info, out)
        if child.tail:
            out.append(_escape_text(child.tail))
    out.append("</%s>" % qname)


def _serialize_xml(root, info):
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n']
    _write_element(root, _Scope({"xml": XML_NS}), info, out)
    return "".join(out).encode("utf-8")


# --------------------------------------------------------------------------
# Дерево-уровневые операции над word/*.xml (принять правки, снять якоря)
# --------------------------------------------------------------------------

def _parent_map(root):
    return {c: p for p in root.iter() for c in p}


def _unwrap(parent, elem):
    children = list(elem)
    siblings = list(parent)
    idx = siblings.index(elem)
    if children:
        children[-1].tail = (children[-1].tail or "") + (elem.tail or "")
    elif idx > 0:
        siblings[idx - 1].tail = (siblings[idx - 1].tail or "") + (elem.tail or "")
    else:
        parent.text = (parent.text or "") + (elem.tail or "")
    parent.remove(elem)
    for i, child in enumerate(children):
        parent.insert(idx + i, child)


def _drop(parent, elem):
    siblings = list(parent)
    idx = siblings.index(elem)
    if idx > 0:
        siblings[idx - 1].tail = (siblings[idx - 1].tail or "") + (elem.tail or "")
    else:
        parent.text = (parent.text or "") + (elem.tail or "")
    parent.remove(elem)


def _accept_revisions(root, counters):
    """Принимает все правки рецензирования в дереве. Заново вычисляет
    родителя перед каждой операцией -- дёшево для документов такого размера
    и застраховано от редкого случая вложенных w:ins/w:del."""
    while True:
        target = None
        for elem in root.iter():
            if elem.tag in ALL_TARGET_TAGS:
                target = elem
                break
        if target is None:
            return
        parent_map = _parent_map(root)
        parent = parent_map.get(target)
        if parent is None:
            return
        tag = target.tag
        if tag in UNWRAP_TAGS:
            _unwrap(parent, target)
            counters["insertions_accepted" if tag == w("ins") else "moves_accepted"] += 1
        elif tag in DROP_CONTENT_TAGS:
            _drop(parent, target)
            counters["deletions_removed" if tag == w("del") else "moves_removed"] += 1
        elif tag in DROP_ANCHOR_TAGS:
            _drop(parent, target)
            if tag in (w("commentRangeStart"), w("commentRangeEnd"), w("commentReference")):
                counters["comment_anchors_removed"] += 1
            else:
                counters["move_range_markers_removed"] += 1
        elif tag in DROP_CHANGE_TAGS:
            _drop(parent, target)
            counters["format_history_removed"] += 1


def _count_revisions(root):
    counts = defaultdict(int)
    for elem in root.iter():
        if elem.tag in ALL_TARGET_TAGS:
            counts[elem.tag] += 1
    return counts


def _strip_rsid_attrs(root):
    prefix = w("rsid")
    removed = 0
    for elem in root.iter():
        for key in list(elem.attrib):
            if key.startswith(prefix):
                del elem.attrib[key]
                removed += 1
    return removed


def _count_rsid_attrs(root):
    prefix = w("rsid")
    return sum(1 for elem in root.iter() for key in elem.attrib if key.startswith(prefix))


# --------------------------------------------------------------------------
# docProps/core.xml, docProps/app.xml
# --------------------------------------------------------------------------

def _fmt_z(dt):
    return dt.strftime("%Y-%m-%dT%H:%M:%SZ")


def _clean_core_tree(root, modified_dt, created_dt):
    cleared = []
    for ns, local in CORE_TEXT_FIELDS:
        el = root.find("{%s}%s" % (ns, local))
        if el is not None and el.text:
            cleared.append(local)
        if el is not None:
            el.text = ""
    rev_el = root.find("{%s}revision" % CP_NS)
    if rev_el is not None:
        if (rev_el.text or "") != "1":
            cleared.append("revision")
        rev_el.text = "1"
    # None -- в документе не было даты (режим --seed): даты не выдумываются и не меняются.
    created_el = root.find("{%s}created" % DCTERMS_NS)
    if created_el is not None and created_dt is not None:
        created_el.text = _fmt_z(created_dt)
    modified_el = root.find("{%s}modified" % DCTERMS_NS)
    if modified_el is not None and modified_dt is not None:
        modified_el.text = _fmt_z(modified_dt)
    return cleared


def _inspect_core_tree(root):
    nonempty = [local for ns, local in CORE_TEXT_FIELDS
                if (root.find("{%s}%s" % (ns, local)) is not None
                    and (root.find("{%s}%s" % (ns, local)).text or ""))]
    rev_el = root.find("{%s}revision" % CP_NS)
    revision_not_one = rev_el is not None and (rev_el.text or "") != "1"
    return nonempty, revision_not_one


def _clean_app_tree(root):
    changed = []

    def _clear(local):
        el = root.find("{%s}%s" % (EP_NS, local))
        if el is not None:
            if el.text:
                changed.append(local)
            el.text = ""

    def _remove(local):
        el = root.find("{%s}%s" % (EP_NS, local))
        if el is not None:
            root.remove(el)
            changed.append(local)

    def _force(local, value):
        el = root.find("{%s}%s" % (EP_NS, local))
        if el is not None:
            if (el.text or "") != value:
                changed.append(local)
            el.text = value

    _clear("Company")
    _clear("Manager")
    _remove("HyperlinkBase")
    _force("Template", "Normal.dotm")
    _force("TotalTime", "0")
    return changed


def _inspect_app_tree(root):
    found = []
    for local in ("Company", "Manager"):
        el = root.find("{%s}%s" % (EP_NS, local))
        if el is not None and (el.text or ""):
            found.append(local)
    if root.find("{%s}HyperlinkBase" % EP_NS) is not None:
        found.append("HyperlinkBase")
    template_el = root.find("{%s}Template" % EP_NS)
    if template_el is not None and (template_el.text or "") not in ("Normal.dotm", ""):
        found.append("Template")
    total_el = root.find("{%s}TotalTime" % EP_NS)
    if total_el is not None and (total_el.text or "0") not in ("0", ""):
        found.append("TotalTime")
    return found


# --------------------------------------------------------------------------
# settings.xml: снять w:trackChanges, по флагу -- список w:rsids
# --------------------------------------------------------------------------

def _clean_settings_tree(root, strip_rsid):
    track_changes_removed = False
    rsids_removed = False
    parent_map = _parent_map(root)
    tc = root.find(w("trackChanges"))
    if tc is not None:
        parent = parent_map.get(tc, root)
        parent.remove(tc)
        track_changes_removed = True
    if strip_rsid:
        rsids_el = root.find(w("rsids"))
        if rsids_el is not None:
            parent = parent_map.get(rsids_el, root)
            parent.remove(rsids_el)
            rsids_removed = True
    return track_changes_removed, rsids_removed


# --------------------------------------------------------------------------
# [Content_Types].xml, связи (_rels/.rels, word/_rels/document.xml.rels)
# --------------------------------------------------------------------------

def _remove_content_type_overrides(root, partnames):
    removed = []
    for child in list(root):
        if child.tag == "{%s}Override" % CT_NS and child.get("PartName") in partnames:
            root.remove(child)
            removed.append(child.get("PartName"))
    return removed


def _remove_relationships_by_target(root, targets):
    removed = []
    for rel in list(root):
        if rel.get("Target") in targets:
            root.remove(rel)
            removed.append(rel.get("Id"))
    return removed


# --------------------------------------------------------------------------
# Обёртка "проанализировать часть word/*.xml" -- общая для очистки и --check
# --------------------------------------------------------------------------

def _iter_word_xml_parts(parts, exclude):
    for name in parts:
        if name in exclude:
            continue
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue
        yield name


# --------------------------------------------------------------------------
# Публичный API: анализ (только чтение) и очистка
# --------------------------------------------------------------------------

class Issues:
    def __init__(self):
        self.core_nonempty_fields = []
        self.revision_not_one = False
        self.app_fields_present = []
        self.custom_xml_present = False
        self.comment_parts_present = []
        self.comment_anchors = 0
        self.tracked_changes = 0
        self.format_history = 0
        self.track_changes_setting_on = False
        self.rsid_attributes = 0
        self.rsid_counted_as_dirty = False

    def is_dirty(self):
        return bool(
            self.core_nonempty_fields or self.revision_not_one or self.app_fields_present
            or self.custom_xml_present or self.comment_parts_present or self.comment_anchors
            or self.tracked_changes or self.format_history or self.track_changes_setting_on
            or (self.rsid_counted_as_dirty and self.rsid_attributes)
        )

    def as_dict(self):
        return {
            "dirty": self.is_dirty(),
            "core_nonempty_fields": self.core_nonempty_fields,
            "revision_not_one": self.revision_not_one,
            "app_fields_present": self.app_fields_present,
            "custom_xml_present": self.custom_xml_present,
            "comment_parts_present": self.comment_parts_present,
            "comment_anchors": self.comment_anchors,
            "tracked_changes": self.tracked_changes,
            "format_history": self.format_history,
            "track_changes_setting_on": self.track_changes_setting_on,
            "rsid_attributes": self.rsid_attributes,
        }


def analyze_issues(parts, check_rsid=False):
    """Только читает `parts` (dict имя_части -> bytes), ничего не меняет."""
    issues = Issues()
    issues.rsid_counted_as_dirty = check_rsid

    if "docProps/core.xml" in parts:
        root, _ = _parse_xml(parts["docProps/core.xml"])
        issues.core_nonempty_fields, issues.revision_not_one = _inspect_core_tree(root)

    if "docProps/app.xml" in parts:
        root, _ = _parse_xml(parts["docProps/app.xml"])
        issues.app_fields_present = _inspect_app_tree(root)

    issues.custom_xml_present = "docProps/custom.xml" in parts
    issues.comment_parts_present = [n for n in COMMENT_PART_NAMES if n in parts]

    if "word/settings.xml" in parts:
        root, _ = _parse_xml(parts["word/settings.xml"])
        issues.track_changes_setting_on = root.find(w("trackChanges")) is not None

    for name in _iter_word_xml_parts(parts, exclude=set(issues.comment_parts_present)):
        try:
            root, _ = _parse_xml(parts[name])
        except ET.ParseError:
            continue
        counts = _count_revisions(root)
        for tag, n in counts.items():
            if tag in DROP_ANCHOR_TAGS and tag in (
                w("commentRangeStart"), w("commentRangeEnd"), w("commentReference")
            ):
                issues.comment_anchors += n
            elif tag in DROP_CHANGE_TAGS:
                issues.format_history += n
            else:
                issues.tracked_changes += n
        issues.rsid_attributes += _count_rsid_attrs(root)

    return issues


def clean_docx_parts(parts, seed=None, strip_rsid=False):
    """Возвращает (new_parts, report). `parts` -- OrderedDict имя -> bytes в
    исходном порядке архива; не изменяется."""
    if "word/document.xml" not in parts:
        raise DocxStructureError("в архиве нет word/document.xml -- это не .docx")

    # created = modified: дата создания не сдвигается в прошлое (не подделывается).
    future_warning = None
    if seed:
        modified_dt, created_dt = _deterministic_times(parts)
        deterministic = True
        if modified_dt is not None:
            now = datetime.datetime.now(datetime.timezone.utc).replace(tzinfo=None, microsecond=0)
            if modified_dt > now:
                future_warning = (
                    "дата последнего сохранения в документе (%s) больше текущего времени (%s): "
                    "проверь часы на компьютере, где документ сохраняли; дата не изменена"
                    % (_fmt_z(modified_dt), _fmt_z(now))
                )
    else:
        modified_dt = datetime.datetime.now(datetime.timezone.utc).replace(
            tzinfo=None, microsecond=0)
        created_dt = modified_dt
        deterministic = False

    report = {
        "warnings": [],
        "core_fields_cleared": [],
        "app_fields_cleared": [],
        "custom_xml_removed": False,
        "comment_parts_removed": [],
        "comment_anchors_removed": 0,
        "insertions_accepted": 0,
        "moves_accepted": 0,
        "deletions_removed": 0,
        "moves_removed": 0,
        "move_range_markers_removed": 0,
        "format_history_removed": 0,
        "track_changes_disabled": False,
        "rsids_list_removed": False,
        "rsid_attributes_removed": 0,
        "seed": seed,
        "deterministic": deterministic,
    }
    if future_warning:
        report["warnings"].append(future_warning)

    to_delete = set()
    if "docProps/custom.xml" in parts:
        to_delete.add("docProps/custom.xml")
        report["custom_xml_removed"] = True
    for name in COMMENT_PART_NAMES:
        if name in parts:
            to_delete.add(name)
            report["comment_parts_removed"].append(name)

    # Переписываются только те части, в которых что-то действительно изменилось;
    # остальные байты (styles.xml, theme1.xml, fontTable.xml ...) копируются как есть.
    new_parts = OrderedDict()
    for name, data in parts.items():
        if name in to_delete:
            continue

        if name == "docProps/core.xml":
            root, ns = _parse_xml(data)
            report["core_fields_cleared"] = _clean_core_tree(root, modified_dt, created_dt)
            data = _serialize_xml(root, ns)

        elif name == "docProps/app.xml":
            root, ns = _parse_xml(data)
            report["app_fields_cleared"] = _clean_app_tree(root)
            if report["app_fields_cleared"]:
                data = _serialize_xml(root, ns)

        elif name == "_rels/.rels" and report["custom_xml_removed"]:
            root, ns = _parse_xml(data)
            if _remove_relationships_by_target(root, {"docProps/custom.xml"}):
                data = _serialize_xml(root, ns)

        elif name == "[Content_Types].xml" and (
            report["custom_xml_removed"] or report["comment_parts_removed"]
        ):
            root, ns = _parse_xml(data)
            partnames = set()
            if report["custom_xml_removed"]:
                partnames.add("/docProps/custom.xml")
            for cp in report["comment_parts_removed"]:
                partnames.add("/" + cp)
            if _remove_content_type_overrides(root, partnames):
                data = _serialize_xml(root, ns)

        elif name == "word/_rels/document.xml.rels" and report["comment_parts_removed"]:
            root, ns = _parse_xml(data)
            targets = {cp.split("/", 1)[1] for cp in report["comment_parts_removed"]}
            if _remove_relationships_by_target(root, targets):
                data = _serialize_xml(root, ns)

        elif name == "word/settings.xml":
            root, ns = _parse_xml(data)
            tc_removed, rsids_removed = _clean_settings_tree(root, strip_rsid)
            report["track_changes_disabled"] = tc_removed
            report["rsids_list_removed"] = rsids_removed
            counters = defaultdict(int)
            _accept_revisions(root, counters)
            _merge_counters(report, counters)
            rsid_removed = _strip_rsid_attrs(root) if strip_rsid else 0
            report["rsid_attributes_removed"] += rsid_removed
            if tc_removed or rsids_removed or rsid_removed or any(counters.values()):
                data = _serialize_xml(root, ns)

        elif name.startswith("word/") and name.endswith(".xml"):
            try:
                root, ns = _parse_xml(data)
            except ET.ParseError:
                new_parts[name] = data
                continue
            counters = defaultdict(int)
            _accept_revisions(root, counters)
            _merge_counters(report, counters)
            rsid_removed = _strip_rsid_attrs(root) if strip_rsid else 0
            report["rsid_attributes_removed"] += rsid_removed
            if rsid_removed or any(counters.values()):
                data = _serialize_xml(root, ns)

        new_parts[name] = data

    return new_parts, report


def _merge_counters(report, counters):
    for key in ("insertions_accepted", "moves_accepted", "deletions_removed",
                "moves_removed", "move_range_markers_removed",
                "comment_anchors_removed", "format_history_removed"):
        report[key] += counters.get(key, 0)


ZIP_EPOCH = datetime.datetime(1980, 1, 1)  # минимальная дата ZIP; только для служебных дат записей архива

# Поля отчёта, любое непустое значение которых означает: очистка что-то нашла и изменила.
_REPORT_CHANGE_KEYS = (
    "core_fields_cleared", "app_fields_cleared", "custom_xml_removed", "comment_parts_removed",
    "comment_anchors_removed", "insertions_accepted", "moves_accepted", "deletions_removed",
    "moves_removed", "move_range_markers_removed", "format_history_removed",
    "track_changes_disabled", "rsids_list_removed", "rsid_attributes_removed",
)


def core_date_texts(parts):
    """(created, modified) так, как они записаны во входном docProps/core.xml (None -- поля нет)."""
    raw = parts.get("docProps/core.xml") if isinstance(parts, dict) else None
    if not raw:
        return None, None
    text = raw.decode("utf-8", errors="replace")

    def value(local):
        match = re.search(r"<[^>]*:%s\b[^>]*>\s*([^<]*?)\s*<" % local, text)
        return match.group(1).strip() if match else None

    return value("created"), value("modified")


def nothing_to_clean(parts, report):
    """Очистке нечего делать: следов нет и даты трогать не нужно (created == modified).

    В этом случае повторный запуск не должен переписывать файл: новые даты изменили бы
    байты сдаваемого DOCX, а значит и снимок аудита (AUDIT_SNAPSHOT_STALE) -- волны
    финальных проверок пришлось бы проводить заново без единой правки в работе.
    """
    if any(report.get(key) for key in _REPORT_CHANGE_KEYS):
        return False
    created, modified = core_date_texts(parts)
    return created == modified


def _deterministic_times(parts):
    """(modified, created) для --seed: реальная дата последнего сохранения из самого документа.

    Дата не выдумывается: берётся dcterms:modified (или dcterms:created) входного
    docProps/core.xml. Если даты в документе нет, возвращается (None, None) и даты
    в core.xml не меняются.
    """
    raw = parts.get("docProps/core.xml") if isinstance(parts, dict) else None
    if not raw:
        return None, None
    text = raw.decode("utf-8", errors="replace")
    for local in ("modified", "created"):
        match = re.search(r"<[^>]*:%s\b[^>]*>\s*([^<]+?)\s*<" % local, text)
        if not match:
            continue
        value = match.group(1).strip().replace("Z", "+00:00")
        try:
            parsed = datetime.datetime.fromisoformat(value)
        except ValueError:
            try:
                parsed = datetime.datetime.strptime(value[:19], "%Y-%m-%dT%H:%M:%S")
            except ValueError:
                continue
        if parsed.tzinfo is not None:
            parsed = parsed.astimezone(datetime.timezone.utc).replace(tzinfo=None)
        parsed = parsed.replace(microsecond=0)
        return parsed, parsed
    return None, None


# --------------------------------------------------------------------------
# Чтение/запись ZIP
# --------------------------------------------------------------------------

def read_docx(path):
    with zipfile.ZipFile(str(path), "r") as zin:
        infolist = zin.infolist()
        parts = OrderedDict((zi.filename, zin.read(zi.filename)) for zi in infolist)
    return parts, infolist


def write_docx(path, parts, source_infolist, deterministic, fixed_dt):
    info_by_name = {zi.filename: zi for zi in source_infolist}
    if deterministic:
        date_time = (fixed_dt.year, fixed_dt.month, fixed_dt.day,
                     fixed_dt.hour, fixed_dt.minute, fixed_dt.second)
    else:
        date_time = time.localtime()[:6]
    with zipfile.ZipFile(str(path), "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in parts.items():
            zi = zipfile.ZipInfo(name, date_time=date_time)
            src = info_by_name.get(name)
            if src is not None:
                zi.external_attr = src.external_attr
                zi.create_system = src.create_system
            else:
                zi.external_attr = 0o600 << 16
            zi.compress_type = zipfile.ZIP_DEFLATED
            zout.writestr(zi, data, compress_type=zipfile.ZIP_DEFLATED, compresslevel=6)


def _same_path(a, b):
    return os.path.normcase(os.path.abspath(str(a))) == os.path.normcase(os.path.abspath(str(b)))


class DocxIntegrityError(Exception):
    """Результат очистки не прошёл структурную проверку docx_integrity.py."""

    def __init__(self, message, problems):
        Exception.__init__(self, message)
        self.problems = problems


def _integrity_module():
    """scripts/docx_integrity.py, загружается по пути рядом со скриптом."""
    path = Path(__file__).resolve().with_name("docx_integrity.py")
    existing = sys.modules.get("docx_integrity")
    if existing is not None and Path(getattr(existing, "__file__", "") or "").resolve() == path:
        return existing
    spec = importlib.util.spec_from_file_location("docx_integrity", str(path))
    if spec is None or spec.loader is None or not path.is_file():
        raise ImportError("не найден %s" % path)
    module = importlib.util.module_from_spec(spec)
    saved = sys.dont_write_bytecode
    sys.dont_write_bytecode = True  # не писать __pycache__ внутрь каталога скилла
    try:
        spec.loader.exec_module(module)
    finally:
        sys.dont_write_bytecode = saved
    sys.modules["docx_integrity"] = module
    return module


def verify_written_docx(written_path, source_path):
    """Самопроверка записанного результата. Возвращает список проблем с
    пометкой, какие из них были и во входном файле; пустой список -- всё в порядке."""
    integrity = _integrity_module()
    problems = integrity.docx_package_problems(written_path)
    if not problems:
        return []
    before = set(integrity.docx_package_problems(source_path))
    return [p + (" (было и во входном файле)" if p in before else "") for p in problems]


def clean_metadata(input_path, output_path, seed=None):
    """Совместимость со старым API (6.31/6.32, всё ещё используется
    tests/test_audit_fixes_632.py): читает input_path, очищает метаданные и
    атомарно записывает результат в output_path -- в том числе когда
    output_path совпадает с input_path (полный эквивалент --in-place).

    Новый код пусть использует read_docx()/clean_docx_parts()/write_docx()
    или сам CLI -- здесь только удобная обёртка с прежней сигнатурой.
    В отличие от CLI эта обёртка пишет результат всегда, даже когда чистить
    нечего (проверка "already_clean" есть только в CLI): прежние вызывающие
    рассчитывают на существующий output_path.
    Возвращает отчёт (dict), которого у старой версии не было; старые
    вызывающие его игнорируют, поэтому это не ломает совместимость.
    Если результат не проходит структурную проверку (docx_integrity.py),
    файл не записывается и поднимается DocxIntegrityError.
    """
    input_path = Path(input_path)
    output_path = Path(output_path)
    parts, infolist = read_docx(input_path)
    new_parts, report = clean_docx_parts(parts, seed=seed)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(
        prefix=".%s-" % output_path.stem, suffix=".docx.tmp", dir=str(output_path.parent))
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        fixed_dt = (_deterministic_times(parts)[0] or ZIP_EPOCH) if report["deterministic"] else None
        write_docx(tmp_path, new_parts, infolist, report["deterministic"], fixed_dt)
        problems = verify_written_docx(tmp_path, input_path)
        if problems:
            raise DocxIntegrityError(
                "после очистки пакет DOCX не прошёл структурную проверку, файл не записан: "
                + "; ".join(problems[:10]), problems)
        os.replace(str(tmp_path), str(output_path))
    finally:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass
    report["input"] = str(input_path)
    report["output"] = str(output_path)
    return report


# --------------------------------------------------------------------------
# Человекочитаемый вывод
# --------------------------------------------------------------------------

def _print_issues_human(issues, input_path):
    if not issues.is_dirty():
        print("OK: метаданные для очистки не найдены: %s" % input_path)
        return
    print("НАЙДЕНО: файл %s содержит метаданные, которые очистит эта утилита:" % input_path)
    if issues.core_nonempty_fields:
        print("   core.xml: заполнены поля %s" % ", ".join(issues.core_nonempty_fields))
    if issues.revision_not_one:
        print("   core.xml: revision != 1")
    if issues.app_fields_present:
        print("   app.xml: %s" % ", ".join(issues.app_fields_present))
    if issues.custom_xml_present:
        print("   docProps/custom.xml присутствует (нестандартные свойства)")
    if issues.comment_parts_present:
        print("   комментарии: %s" % ", ".join(issues.comment_parts_present))
    if issues.comment_anchors:
        print("   якорей комментариев в тексте: %d" % issues.comment_anchors)
    if issues.tracked_changes:
        print("   непринятых исправлений рецензирования: %d" % issues.tracked_changes)
    if issues.format_history:
        print("   записей истории форматирования (*Change): %d" % issues.format_history)
    if issues.track_changes_setting_on:
        print("   включена запись исправлений (w:trackChanges)")
    if issues.rsid_counted_as_dirty and issues.rsid_attributes:
        print("   rsid-атрибутов: %d (--strip-rsid уберёт)" % issues.rsid_attributes)
    elif issues.rsid_attributes:
        print("   (справочно) rsid-атрибутов: %d -- уберёт --strip-rsid" % issues.rsid_attributes)


def _print_report_human(report):
    print("OK: метаданные очищены: %s" % report["output"])
    for message in report.get("warnings") or ():
        print("ПРЕДУПРЕЖДЕНИЕ: %s" % message)
    print("   created:  %s" % report["created"])
    print("   modified: %s" % report["modified"])
    print("   детерминированно (--seed): %s" % ("да" if report["deterministic"] else "нет"))
    if report["core_fields_cleared"]:
        print("   core.xml очищены: %s" % ", ".join(report["core_fields_cleared"]))
    if report["app_fields_cleared"]:
        print("   app.xml изменены: %s" % ", ".join(report["app_fields_cleared"]))
    if report["custom_xml_removed"]:
        print("   docProps/custom.xml удалён")
    if report["comment_parts_removed"]:
        print("   удалены части комментариев: %s" % ", ".join(report["comment_parts_removed"]))
    if report["comment_anchors_removed"]:
        print("   удалено якорей комментариев: %d" % report["comment_anchors_removed"])
    accepted = report["insertions_accepted"] + report["moves_accepted"]
    removed_rev = report["deletions_removed"] + report["moves_removed"]
    if accepted or removed_rev:
        print("   правки рецензирования: принято %d, удалено %d" % (accepted, removed_rev))
    if report["format_history_removed"]:
        print("   удалено записей истории форматирования: %d" % report["format_history_removed"])
    if report["track_changes_disabled"]:
        print("   запись исправлений (w:trackChanges) отключена")
    if report["strip_rsid"]:
        print("   rsid: удалено атрибутов %d%s" % (
            report["rsid_attributes_removed"],
            ", список w:rsids удалён" if report["rsids_list_removed"] else ""))
    print()
    print("ВНИМАНИЕ: после очистки повторно запусти финальную проверку и snapshot аудита.")
    print("Ограничения см. в --help (объединение абзацев при удалении абзацного знака,")
    print("облачная история версий, встроенные OLE-объекты -- не затрагиваются).")


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def _build_parser():
    parser = argparse.ArgumentParser(
        prog="clean_docx_metadata.py",
        description="Чистит метаданные .docx перед финальной сдачей "
                     "(core.xml, app.xml, custom.xml, комментарии, правки рецензирования).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("input", type=Path, help="Входной .docx")
    parser.add_argument("-o", "--output", type=Path,
                         help="Выходной .docx (по умолчанию: input-clean.docx)")
    parser.add_argument("--in-place", action="store_true",
                         help="Перезаписать входной файл (нельзя вместе с -o/--check)")
    parser.add_argument("--check", action="store_true",
                         help="Ничего не менять: только сообщить, что было бы очищено "
                              "(код выхода 1, если найдено; нельзя вместе с -o/--in-place)")
    parser.add_argument("--seed", type=str, default=None,
                         help="Детерминированный результат: одинаковый вход даёт одинаковый выход; "
                              "created = modified = дата последнего сохранения из самого документа "
                              "(дата не выдумывается). Без --seed created = modified = текущее время UTC")
    parser.add_argument("--strip-rsid", action="store_true",
                         help="Также удалить атрибуты w:rsid* и список w:rsids в settings.xml")
    parser.add_argument("--json", action="store_true",
                         help="Печатать только JSON в stdout (отчёт или {\"status\":\"error\"})")
    return parser


def _error(args, message, code, problems=None):
    if getattr(args, "json", False):
        payload = {"status": "error", "error": message}
        if problems is not None:
            payload["problems"] = problems
        print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    else:
        print("ОШИБКА: %s" % message, file=sys.stderr)
        for item in problems or ():
            print("  - %s" % item, file=sys.stderr)
    return code


def main(argv=None):
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    parser = _build_parser()
    args = parser.parse_args(argv)

    if args.check and (args.output or args.in_place):
        parser.error("--check нельзя сочетать с -o/--output или --in-place")
    if args.in_place and args.output:
        parser.error("--in-place и -o одновременно не задавать")

    if not args.input.exists():
        return _error(args, "файл не найден: %s" % args.input, 2)

    try:
        parts, infolist = read_docx(args.input)
    except zipfile.BadZipFile as exc:
        return _error(args, "повреждённый .docx (не ZIP-архив): %s" % exc, 2)
    except OSError as exc:
        return _error(args, "не удалось прочитать файл: %s" % exc, 2)

    if "word/document.xml" not in parts:
        return _error(args, "в архиве нет word/document.xml -- это не документ Word", 2)

    if args.check:
        issues = analyze_issues(parts, check_rsid=args.strip_rsid)
        if args.json:
            print(json.dumps(issues.as_dict(), ensure_ascii=False, indent=2))
        else:
            _print_issues_human(issues, args.input)
        return 1 if issues.is_dirty() else 0

    try:
        new_parts, report = clean_docx_parts(parts, seed=args.seed, strip_rsid=args.strip_rsid)
    except DocxStructureError as exc:
        return _error(args, str(exc), 2)

    residue = analyze_issues(new_parts, check_rsid=args.strip_rsid)
    if residue.is_dirty():
        return _error(
            args,
            "после очистки в результате остались метаданные (внутренняя проверка "
            "не пройдена), файл не записан: %s" % json.dumps(residue.as_dict(), ensure_ascii=False),
            1,
        )

    if args.in_place:
        output_path = args.input
    elif args.output:
        output_path = args.output
    else:
        if args.input.resolve().parent.name.casefold() == "final":
            # Копия рядом со сдаваемым файлом дала бы второй DOCX в final/ (doctor: FINAL_EXTRA_FILES).
            parser.error("для файла в final/ укажи --in-place или -o exports/<имя>.docx")
        output_path = args.input.with_name(args.input.stem + "-clean" + args.input.suffix)

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        return _error(args, "не удалось создать каталог для вывода: %s" % exc, 2)

    if _same_path(args.input, output_path) and nothing_to_clean(parts, report):
        # Чистить нечего: файл не переписывается, байты и снимок аудита остаются прежними.
        created, modified = core_date_texts(parts)
        report["status"] = "already_clean"
        report["created"] = created or ""
        report["modified"] = modified or ""
        report["input"] = str(args.input)
        report["output"] = str(output_path)
        report["in_place"] = True
        report["strip_rsid"] = args.strip_rsid
        if args.json:
            print(json.dumps(report, ensure_ascii=False, indent=2))
        else:
            print("OK: чистить нечего, файл не изменён: %s" % output_path)
            for message in report["warnings"]:
                print("ПРЕДУПРЕЖДЕНИЕ: %s" % message)
            print("   (повторная очистка переписала бы байты и потребовала нового snapshot аудита)")
        return 0

    fd, tmp_name = tempfile.mkstemp(
        prefix=".%s-" % output_path.stem, suffix=".docx.tmp", dir=str(output_path.parent))
    os.close(fd)
    tmp_path = Path(tmp_name)

    try:
        # modified/created для отчёта берём из уже очищенного core.xml, чтобы
        # не тащить datetime через report дважды.
        core_root, _ = _parse_xml(new_parts["docProps/core.xml"]) \
            if "docProps/core.xml" in new_parts else (None, None)
        if core_root is not None:
            m_el = core_root.find("{%s}modified" % DCTERMS_NS)
            c_el = core_root.find("{%s}created" % DCTERMS_NS)
            report["modified"] = m_el.text if m_el is not None else ""
            report["created"] = c_el.text if c_el is not None else ""
        else:
            report["modified"] = report["created"] = ""

        fixed_dt = None
        if report["deterministic"]:
            fixed_dt = _deterministic_times(parts)[0] or ZIP_EPOCH
        write_docx(tmp_path, new_parts, infolist, report["deterministic"], fixed_dt)
        try:
            problems = verify_written_docx(tmp_path, args.input)
        except ImportError as exc:
            return _error(args, "нет модуля самопроверки docx_integrity.py: %s" % exc, 2)
        if problems:
            # Word такой файл может не открыть -- исходник не трогаем, результат не пишем.
            return _error(
                args,
                "после очистки пакет DOCX не прошёл структурную проверку "
                "(docx_integrity.py), файл не записан",
                1,
                problems=problems,
            )
        os.replace(str(tmp_path), str(output_path))
    except OSError as exc:
        return _error(args, "не удалось записать результат: %s" % exc, 2)
    finally:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass

    report["status"] = "cleaned"
    report["input"] = str(args.input)
    report["output"] = str(output_path)
    report["in_place"] = _same_path(args.input, output_path)
    report["strip_rsid"] = args.strip_rsid

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        _print_report_human(report)

    return 0


if __name__ == "__main__":
    sys.exit(main())
