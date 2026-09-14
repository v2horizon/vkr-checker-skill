#!/usr/bin/env python3
"""Варианты с форматированием в docDefaults: Word COM не даёт писать docDefaults напрямую,
поэтому styles.xml правится здесь, а затем word_fixtures.ps1 пересохраняет файлы в Word.

python xml_variants.py <gen_project_word.docx> <каталог in>
Создаёт in/xml_<имя>.docx.
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def rewrite(src: Path, dst: Path, mutate):
    with zipfile.ZipFile(src) as archive:
        entries = [(info, archive.read(info.filename)) for info in archive.infolist()]
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as archive:
        for info, data in entries:
            if info.filename == "word/styles.xml":
                root = etree.fromstring(data)
                mutate(root)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            archive.writestr(info.filename, data)


def normal_style(root):
    for style in root.findall(W + "style"):
        name = style.find(W + "name")
        if name is not None and name.get(W + "val") == "Normal":
            return style
    raise RuntimeError("Normal not found")


def child(parent, tag):
    el = parent.find(W + tag)
    if el is None:
        el = etree.SubElement(parent, W + tag)
    return el


def defaults(root):
    dd = root.find(W + "docDefaults")
    return dd


def spacing_variant(line):
    def mutate(root):
        spacing = normal_style(root).find(W + "pPr/" + W + "spacing")
        for attr in ("line", "lineRule"):
            spacing.attrib.pop(W + attr, None)
        ppr = defaults(root).find(W + "pPrDefault/" + W + "pPr")
        sp = child(ppr, "spacing")
        sp.set(W + "line", str(line))
        sp.set(W + "lineRule", "auto")
    return mutate


def font_variant(font, half_points):
    def mutate(root):
        rpr = normal_style(root).find(W + "rPr")
        for tag in ("rFonts", "sz", "szCs"):
            el = rpr.find(W + tag)
            if el is not None:
                rpr.remove(el)
        drpr = defaults(root).find(W + "rPrDefault/" + W + "rPr")
        rfonts = drpr.find(W + "rFonts")
        for key in list(rfonts.attrib):
            del rfonts.attrib[key]
        for key in ("ascii", "hAnsi", "eastAsia", "cs"):
            rfonts.set(W + key, font)
        child(drpr, "sz").set(W + "val", str(half_points))
        child(drpr, "szCs").set(W + "val", str(half_points))
    return mutate


def indent_variant(first_line):
    def mutate(root):
        ppr = normal_style(root).find(W + "pPr")
        ind = ppr.find(W + "ind")
        if ind is not None:
            ppr.remove(ind)
        dppr = defaults(root).find(W + "pPrDefault/" + W + "pPr")
        child(dppr, "ind").set(W + "firstLine", str(first_line))
    return mutate


def main():
    src = Path(sys.argv[1])
    out = Path(sys.argv[2])
    rewrite(src, out / "xml_ok_docdefaults_spacing15.docx", spacing_variant(360))
    rewrite(src, out / "xml_bad_docdefaults_spacing_single.docx", spacing_variant(240))
    rewrite(src, out / "xml_ok_docdefaults_font_tnr14.docx", font_variant("Times New Roman", 28))
    rewrite(src, out / "xml_bad_docdefaults_font_arial12.docx", font_variant("Arial", 24))
    rewrite(src, out / "xml_ok_docdefaults_indent125.docx", indent_variant(709))
    rewrite(src, out / "xml_bad_docdefaults_indent0.docx", indent_variant(0))
    print("ok")


if __name__ == "__main__":
    main()
