"""W0 (6.33): очистка метаданных не должна делать DOCX «повреждённым» для Word.

Дефект: clean_docx_metadata.py сериализовал части через ElementTree.tostring,
который объявляет только реально использованные пространства имён. Префиксы,
перечисленные лишь в mc:Ignorable (w15, w16se, ... wp14), теряли объявления
xmlns, и Microsoft Word отказывался открыть файл («Файл поврежден»), хотя
python-docx и validate_vkr.py его читали.

Что проверяется здесь (unittest, без сети; PY39 и PYCODEX):
- scripts/docx_integrity.py находит синтетические поломки и не шумит на
  реальных документах (фикстура, сохранённая Word, шаблоны assets, фикстуры
  валидатора ok_*);
- clean_docx_metadata --in-place на каждом таком документе: проблем пакета нет,
  корневые объявления xmlns каждой части сохранены, непереписанные части
  совпадают побайтно;
- снятая обёртка w:ins со своими xmlns не оставляет необъявленных префиксов;
- самопроверка: сломанный результат не записывается (код 1), исходник цел;
- только при VKR_TEST_WORD=1 на Windows: очищенные документы открывает
  настоящий Word без восстановления (Word в отдельном процессе с тайм-аутом,
  останавливается только свой экземпляр).

Запуск:
    python -m unittest discover -s tests -t .
    python tests/test_docx_integrity_633.py
"""
from __future__ import annotations

import contextlib
import importlib.util
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import textwrap
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from unittest import mock
from xml.parsers import expat

sys.dont_write_bytecode = True

ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "vkr-mpgu"
SCRIPTS = SKILL / "scripts"
WORD_SAVED = ROOT / "tests" / "fixtures" / "build" / "vkr-word-saved.docx"
VALIDATOR_FIXTURES = ROOT / "tests" / "fixtures" / "validator"
ENV = dict(os.environ, PYTHONDONTWRITEBYTECODE="1", PYTHONIOENCODING="utf-8")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def load_module(name):
    spec = importlib.util.spec_from_file_location(name + "_under_test_w0", SCRIPTS / (name + ".py"))
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


cleaner = load_module("clean_docx_metadata")
integrity = load_module("docx_integrity")


def fixture_documents():
    docs = [WORD_SAVED]
    docs.extend(sorted((SKILL / "assets").glob("*.docx")))
    docs.extend(sorted(VALIDATOR_FIXTURES.glob("ok_*.docx")))
    return docs


def root_declarations(data):
    """{префикс: URI} объявлений xmlns на корневом элементе XML-части."""
    found = {}

    class _Stop(Exception):
        pass

    def start(_tag, attrs):
        for key, value in attrs.items():
            if key == "xmlns":
                found[""] = value
            elif key.startswith("xmlns:"):
                found[key[6:]] = value
        raise _Stop()

    parser = expat.ParserCreate()
    parser.StartElementHandler = start
    try:
        parser.Parse(data, True)
    except _Stop:
        pass
    return found


def run_cleaner(*args):
    buffer = io.StringIO()
    with contextlib.redirect_stdout(buffer), contextlib.redirect_stderr(io.StringIO()):
        code = cleaner.main([str(a) for a in args])
    return code, buffer.getvalue()


CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '</Types>'
)
ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="word/document.xml"/></Relationships>'
)


def document_xml(root_attrs, body):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<w:document xmlns:w="%s" xmlns:mc="%s" %s><w:body>%s<w:sectPr/></w:body></w:document>'
        % (W_NS, MC_NS, root_attrs, body)
    )


def write_package(path, parts):
    with zipfile.ZipFile(str(path), "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data.encode("utf-8") if isinstance(data, str) else data)


def minimal_parts(document):
    return {"[Content_Types].xml": CONTENT_TYPES, "_rels/.rels": ROOT_RELS, "word/document.xml": document}


def inject_revision_into_word_fixture(path):
    """Добавляет в начало document.xml абзац с непринятой вставкой w:ins."""
    with zipfile.ZipFile(str(path)) as archive:
        parts = [(info.filename, archive.read(info.filename)) for info in archive.infolist()]
    with zipfile.ZipFile(str(path), "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts:
            if name == "word/document.xml":
                insert = ('<w:p><w:ins w:id="900" w:author="W0" w:date="2024-01-01T00:00:00Z">'
                          '<w:r><w:t>Вставка W0</w:t></w:r></w:ins></w:p>')
                data = data.decode("utf-8").replace("<w:body>", "<w:body>" + insert, 1).encode("utf-8")
            archive.writestr(name, data)


def old_elementtree_serializer(root, info):
    """Сериализация 6.33 до исправления W0: ET.tostring с регистрацией
    префиксов -- объявляет только используемые пространства имён."""
    seen = set()
    for prefix, uri in info.prefix_hint.items():
        if uri not in seen:
            try:
                ET.register_namespace(prefix, uri)
            except ValueError:
                pass
            seen.add(uri)
    body = ET.tostring(root, encoding="unicode")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + body).encode("utf-8")


# ---------------------------------------------------------------------------
# docx_integrity.py
# ---------------------------------------------------------------------------

class IntegrityCheckerTests(unittest.TestCase):

    def setUp(self):
        self._td = tempfile.TemporaryDirectory(prefix="vkr633-integrity-")
        self.addCleanup(self._td.cleanup)
        self.tmp = Path(self._td.name)

    def test_real_documents_have_no_problems(self):
        for path in fixture_documents():
            with self.subTest(document=path.name):
                self.assertEqual([], integrity.docx_package_problems(path))

    def test_undeclared_ignorable_prefix_detected(self):
        path = self.tmp / "ignorable.docx"
        write_package(path, minimal_parts(document_xml(
            'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14 w15"',
            "<w:p/>")))
        problems = integrity.docx_package_problems(path)
        self.assertEqual(1, len(problems), problems)
        self.assertIn("«w15»", problems[0])
        self.assertIn("mc:Ignorable", problems[0])

    def test_declared_ignorable_prefixes_pass(self):
        path = self.tmp / "ignorable-ok.docx"
        write_package(path, minimal_parts(document_xml(
            'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
            'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 w15"',
            "<w:p/>")))
        self.assertEqual([], integrity.docx_package_problems(path))

    def test_must_understand_process_content_and_requires_detected(self):
        body = ('<w:p mc:MustUnderstand="w16" mc:ProcessContent="w17:thing">'
                '<mc:AlternateContent><mc:Choice Requires="wps"><w:r/></mc:Choice>'
                '<mc:Fallback><w:r/></mc:Fallback></mc:AlternateContent></w:p>')
        path = self.tmp / "mce.docx"
        write_package(path, minimal_parts(document_xml("", body)))
        joined = "\n".join(integrity.docx_package_problems(path))
        self.assertIn("«w16» из mc:MustUnderstand", joined)
        self.assertIn("«w17» из mc:ProcessContent", joined)
        self.assertIn("«wps» из mc:Choice/@Requires", joined)

    def test_prefix_declared_on_inner_element_is_in_scope(self):
        body = '<w:p><w:r xmlns:v="urn:v" mc:Ignorable="v"/></w:p><w:p><w:r mc:Ignorable="v"/></w:p>'
        path = self.tmp / "scope.docx"
        write_package(path, minimal_parts(document_xml("", body)))
        problems = integrity.docx_package_problems(path)
        # первое объявление видно только внутри своего w:r
        self.assertEqual(1, len(problems), problems)
        self.assertIn("«v»", problems[0])

    def test_text_outside_run_detected(self):
        # T2: <w:p><w:t>…</w:t></w:p> — python-docx читает, Word отвечает «Ошибка при попытке
        # открытия файла». Прежде docx_integrity такой файл пропускал.
        path = self.tmp / "text-outside-run.docx"
        write_package(path, minimal_parts(document_xml("", "<w:p><w:t>висячий</w:t></w:p>")))
        problems = integrity.docx_package_problems(path)
        self.assertEqual(1, len(problems), problems)
        self.assertIn("<w:t>", problems[0])
        self.assertIn("<w:p>", problems[0])
        self.assertIn("word/document.xml", problems[0])

    def test_misplaced_elements_detected(self):
        cases = {
            "прогон вне абзаца": ("<w:r><w:t>текст</w:t></w:r>", "<w:r>"),
            "абзац внутри прогона": ("<w:p><w:r><w:p/></w:r></w:p>", "<w:p>"),
            "ячейка вне строки": ("<w:tbl><w:tc><w:p/></w:tc></w:tbl>", "<w:tc>"),
            "строка вне таблицы": ("<w:tr><w:tc><w:p/></w:tc></w:tr>", "<w:tr>"),
            "таблица внутри прогона": ("<w:p><w:r><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:r></w:p>", "<w:tbl>"),
            "разрыв строки вне прогона": ("<w:p><w:br/></w:p>", "<w:br>"),
            "рисунок вне прогона": ("<w:p><w:drawing/></w:p>", "<w:drawing>"),
        }
        for name, (body, needle) in cases.items():
            with self.subTest(case=name):
                path = self.tmp / ("misplaced-%d.docx" % abs(hash(name)))
                write_package(path, minimal_parts(document_xml("", body)))
                problems = [p for p in integrity.docx_package_problems(path) if needle in p]
                self.assertTrue(problems, integrity.docx_package_problems(path))

    def test_legitimate_placements_pass(self):
        body = (
            '<w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9345"/></w:tabs></w:pPr>'
            '<w:hyperlink r:id="rId1"><w:r><w:t>ссылка</w:t></w:r></w:hyperlink>'
            '<w:r><w:tab/><w:t>текст</w:t><w:br/><w:sym w:font="Symbol" w:char="F0B7"/></w:r>'
            '<w:ins w:id="1"><w:r><w:t>вставка</w:t></w:r></w:ins>'
            '<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>'
            '<w:sdt><w:sdtContent><w:r><w:t>поле</w:t></w:r></w:sdtContent></w:sdt></w:p>'
            '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>ячейка</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
            '<w:sdt><w:sdtContent><w:tr><w:tc><w:p/></w:tc></w:tr></w:sdtContent></w:sdt>'
        )
        path = self.tmp / "legit.docx"
        write_package(path, minimal_parts(document_xml(
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"', body)))
        self.assertEqual([], integrity.docx_package_problems(path))

    def test_placement_checked_in_headers_footnotes_and_comments(self):
        body = "<w:p><w:t>висячий</w:t></w:p>"
        for part in ("word/header1.xml", "word/footer2.xml", "word/footnotes.xml", "word/endnotes.xml",
                     "word/comments.xml", "word/document.xml"):
            with self.subTest(part=part):
                self.assertTrue(integrity.element_placement_problems(part, document_xml("", body).encode("utf-8")), part)
        # части без размеченного текста в проверку не попадают (в styles.xml и numbering.xml своя схема)
        for part in ("word/styles.xml", "word/numbering.xml", "word/settings.xml", "docProps/app.xml"):
            self.assertIsNone(integrity.BODY_PART_RE.match(part), part)

    def test_no_placement_problems_in_any_repository_document(self):
        # Ложные срабатывания дороже пропуска: проверка идёт по всем DOCX репозитория.
        documents = sorted(ROOT.rglob("*.docx"))
        self.assertGreaterEqual(len(documents), 41, documents)
        for path in documents:
            with self.subTest(document=path.name):
                placement = [p for p in integrity.docx_package_problems(path) if "WordprocessingML" in p]
                self.assertEqual([], placement)

    def test_missing_content_type_detected(self):
        parts = minimal_parts(document_xml("", "<w:p/>"))
        parts["word/media/extra.bin"] = b"\x00\x01"
        path = self.tmp / "content-type.docx"
        write_package(path, parts)
        problems = integrity.docx_package_problems(path)
        self.assertEqual(1, len(problems), problems)
        self.assertIn("word/media/extra.bin", problems[0])

    def test_dangling_internal_relationship_detected_external_ignored(self):
        parts = minimal_parts(document_xml("", "<w:p/>"))
        parts["word/_rels/document.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
            'Target="styles.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
            'Target="https://example.invalid/x" TargetMode="External"/>'
            '</Relationships>')
        path = self.tmp / "rels.docx"
        write_package(path, parts)
        problems = integrity.docx_package_problems(path)
        self.assertEqual(1, len(problems), problems)
        self.assertIn("word/styles.xml", problems[0])
        self.assertIn("rId1", problems[0])

    def test_malformed_xml_and_missing_main_document_detected(self):
        parts = minimal_parts('<w:document xmlns:w="%s"><w:body></w:document>' % W_NS)
        path = self.tmp / "malformed.docx"
        write_package(path, parts)
        self.assertTrue(any("некорректный XML" in p for p in integrity.docx_package_problems(path)))

        parts = minimal_parts(document_xml("", "<w:p/>"))
        parts["_rels/.rels"] = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                                'relationships"/>')
        path = self.tmp / "no-main.docx"
        write_package(path, parts)
        self.assertTrue(any("officeDocument" in p for p in integrity.docx_package_problems(path)))

    def test_elementtree_roundtrip_of_word_part_is_detected(self):
        # Ровно тот класс поломки, что дал W0: часть Word, пересобранная ElementTree.
        with zipfile.ZipFile(str(WORD_SAVED)) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        root, info = cleaner._parse_xml(parts["word/document.xml"])
        parts["word/document.xml"] = old_elementtree_serializer(root, info)
        path = self.tmp / "et-roundtrip.docx"
        write_package(path, parts)
        problems = integrity.docx_package_problems(path)
        self.assertTrue(problems)
        self.assertTrue(all(p.startswith("word/document.xml:") and "mc:Ignorable" in p for p in problems), problems)

    def test_cli_exit_codes_and_json(self):
        script = str(SCRIPTS / "docx_integrity.py")
        ok = self.tmp / "ok.docx"
        shutil.copyfile(str(WORD_SAVED), str(ok))
        broken = self.tmp / "broken.docx"
        write_package(broken, minimal_parts(document_xml('mc:Ignorable="w15"', "<w:p/>")))
        not_zip = self.tmp / "not-zip.docx"
        not_zip.write_bytes(b"not a zip")

        def run(*args):
            result = subprocess.run([sys.executable, script] + [str(a) for a in args],
                                    capture_output=True, env=ENV, check=False)
            return result.returncode, result.stdout.decode("utf-8")

        code, out = run(ok, "--json")
        self.assertEqual(0, code, out)
        self.assertEqual({"status": "ok", "problems": []},
                         {k: v for k, v in json.loads(out).items() if k != "file"})
        code, out = run(broken, "--json")
        self.assertEqual(1, code, out)
        payload = json.loads(out)
        self.assertEqual("problems", payload["status"])
        self.assertTrue(any("«w15»" in p for p in payload["problems"]))
        code, out = run(broken)
        self.assertEqual(1, code)
        self.assertIn("w15", out)
        code, out = run(self.tmp / "absent.docx", "--json")
        self.assertEqual(2, code)
        self.assertEqual("error", json.loads(out)["status"])
        code, _out = run(not_zip)
        self.assertEqual(2, code)

    def test_module_uses_only_standard_library(self):
        source = (SCRIPTS / "docx_integrity.py").read_text(encoding="utf-8")
        for forbidden in ("import docx", "from docx", "lxml"):
            self.assertNotIn(forbidden, source)


# ---------------------------------------------------------------------------
# clean_docx_metadata.py: сериализация и самопроверка
# ---------------------------------------------------------------------------

class CleanerKeepsPackageIntactTests(unittest.TestCase):

    def setUp(self):
        self._td = tempfile.TemporaryDirectory(prefix="vkr633-clean-w0-")
        self.addCleanup(self._td.cleanup)
        self.tmp = Path(self._td.name)

    def _assert_clean_keeps_package(self, source, extra_args):
        target = self.tmp / ("clean-" + source.name)
        shutil.copyfile(str(source), str(target))
        code, out = run_cleaner(target, "--in-place", "--json", *extra_args)
        self.assertEqual(0, code, out)
        self.assertEqual([], integrity.docx_package_problems(target))
        with zipfile.ZipFile(str(source)) as before, zipfile.ZipFile(str(target)) as after:
            after_names = set(after.namelist())
            for name in before.namelist():
                if name not in after_names or not name.lower().endswith((".xml", ".rels")):
                    continue
                expected = root_declarations(before.read(name))
                actual = root_declarations(after.read(name))
                lost = {p: u for p, u in expected.items() if actual.get(p) != u}
                self.assertEqual({}, lost, "%s: утеряны объявления xmlns корня" % name)

    def test_clean_in_place_every_fixture_keeps_namespace_declarations(self):
        for source in fixture_documents():
            for extra in ([], ["--strip-rsid"]):
                with self.subTest(document=source.name, args=extra):
                    self._assert_clean_keeps_package(source, extra)

    def test_unchanged_parts_are_copied_byte_for_byte(self):
        target = self.tmp / "word-saved.docx"
        shutil.copyfile(str(WORD_SAVED), str(target))
        code, out = run_cleaner(target, "--in-place", "--json")
        self.assertEqual(0, code, out)
        with zipfile.ZipFile(str(WORD_SAVED)) as before, zipfile.ZipFile(str(target)) as after:
            for name in ("word/document.xml", "word/styles.xml", "word/settings.xml", "word/numbering.xml",
                         "word/fontTable.xml", "word/webSettings.xml", "word/theme/theme1.xml",
                         "word/footer1.xml", "[Content_Types].xml", "_rels/.rels"):
                with self.subTest(part=name):
                    self.assertEqual(before.read(name), after.read(name))

    def test_strip_rsid_rewrite_preserves_ignorable_declarations(self):
        target = self.tmp / "word-saved-rsid.docx"
        shutil.copyfile(str(WORD_SAVED), str(target))
        code, out = run_cleaner(target, "--in-place", "--strip-rsid", "--json")
        self.assertEqual(0, code, out)
        self.assertGreater(json.loads(out)["rsid_attributes_removed"], 0)
        with zipfile.ZipFile(str(target)) as archive:
            document = archive.read("word/document.xml")
        declared = root_declarations(document)
        ignorable = ET.fromstring(document).get("{%s}Ignorable" % MC_NS)
        self.assertTrue(ignorable)
        for prefix in ignorable.split():
            self.assertIn(prefix, declared)
        self.assertNotIn(b"w:rsidR=", document)

    def test_unwrapped_revision_redeclares_its_local_namespaces(self):
        body = (
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z" xmlns:x="urn:x" xmlns:v="urn:v">'
            '<w:r x:foo="1" mc:Ignorable="v"><w:t xml:space="preserve">Принятая вставка &amp; текст</w:t></w:r>'
            '</w:ins><w:del w:id="2" w:author="A" xmlns:z="urn:z"><w:r><w:delText>удалено</w:delText></w:r></w:del>'
            '<mc:AlternateContent><mc:Choice Requires="w14"><w:r><w:t>выбор</w:t></w:r></mc:Choice>'
            '<mc:Fallback><w:r><w:t>выбор</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p>'
        )
        source = self.tmp / "revisions.docx"
        write_package(source, minimal_parts(document_xml(
            'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"', body)))
        self.assertEqual([], integrity.docx_package_problems(source))
        code, out = run_cleaner(source, "--in-place", "--json")
        self.assertEqual(0, code, out)
        self.assertEqual(1, json.loads(out)["insertions_accepted"])
        self.assertEqual([], integrity.docx_package_problems(source))
        with zipfile.ZipFile(str(source)) as archive:
            document = archive.read("word/document.xml")
        root = ET.fromstring(document)
        run = root.find(".//{%s}r" % W_NS)
        self.assertEqual("1", run.get("{urn:x}foo"))
        self.assertEqual("v", run.get("{%s}Ignorable" % MC_NS))
        text = "".join(t.text or "" for t in root.iter("{%s}t" % W_NS))
        self.assertIn("Принятая вставка & текст", text)
        self.assertNotIn("удалено", document.decode("utf-8"))
        self.assertIsNone(root.find(".//{%s}ins" % W_NS))

    def test_self_check_refuses_to_write_broken_result(self):
        target = self.tmp / "guarded.docx"
        shutil.copyfile(str(WORD_SAVED), str(target))
        before = target.read_bytes()
        with mock.patch.object(cleaner, "_serialize_xml", old_elementtree_serializer):
            code, out = run_cleaner(target, "--in-place", "--strip-rsid", "--json")
        self.assertEqual(1, code, out)
        payload = json.loads(out)
        self.assertEqual("error", payload["status"])
        self.assertTrue(any("mc:Ignorable" in p for p in payload["problems"]), payload)
        self.assertEqual(before, target.read_bytes(), "исходный файл не должен меняться")
        self.assertEqual(["guarded.docx"], sorted(p.name for p in self.tmp.iterdir()))

    def test_compat_api_raises_instead_of_writing_broken_result(self):
        # У Word-фикстуры без правок переписываются только core.xml и app.xml
        # (без mc:Ignorable), поэтому подмешиваем вставку w:ins: тогда
        # переписывается и document.xml со своим mc:Ignorable.
        target = self.tmp / "compat.docx"
        shutil.copyfile(str(WORD_SAVED), str(target))
        inject_revision_into_word_fixture(target)
        self.assertEqual([], integrity.docx_package_problems(target))
        before = target.read_bytes()
        with mock.patch.object(cleaner, "_serialize_xml", old_elementtree_serializer):
            with self.assertRaises(cleaner.DocxIntegrityError) as ctx:
                cleaner.clean_metadata(target, target, seed="w0")
        self.assertTrue(ctx.exception.problems)
        self.assertEqual(before, target.read_bytes())
        # Без подмены тот же вызов проходит и даёт целый пакет.
        cleaner.clean_metadata(target, target, seed="w0")
        self.assertEqual([], integrity.docx_package_problems(target))


# ---------------------------------------------------------------------------
# Настоящий Word (только VKR_TEST_WORD=1 на Windows)
# ---------------------------------------------------------------------------

WORD_OPEN_SCRIPT = textwrap.dedent(r"""
    param([string]$ListFile, [string]$OutFile, [string]$PidFile)
    $ErrorActionPreference = 'Stop'
    $files = [System.IO.File]::ReadAllLines($ListFile, [System.Text.Encoding]::UTF8) | Where-Object { $_ }
    $before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $word = New-Object -ComObject Word.Application
    $mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { $_.Id })
    Set-Content -Path $PidFile -Value ($mine -join ",") -Encoding ASCII
    $results = @()
    $m = [Type]::Missing
    try {
        $word.Visible = $false
        $word.DisplayAlerts = 0
        foreach ($path in $files) {
            $item = [ordered]@{ path = $path }
            $doc = $null
            try {
                # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate,
                # Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair
                $doc = $word.Documents.Open($path, $false, $true, $false, $m, $m, $m, $m, $m, $m, $m, $false, $false)
                $item.status = "OPEN_OK"
                $item.pages = $doc.ComputeStatistics(2)
                $item.paragraphs = $doc.Paragraphs.Count
            } catch {
                $item.status = "OPEN_FAIL"
                $item.error = $_.Exception.Message
            } finally {
                if ($doc -ne $null) { $doc.Close($false) | Out-Null }
            }
            $results += [pscustomobject]$item
        }
    } finally {
        # Если новый процесс не появился, мы подключились к чужому Word: его не закрываем.
        if ($mine.Count -gt 0) { $word.Quit() | Out-Null }
        $json = ConvertTo-Json -InputObject $results -Depth 4
        [System.IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
    }
""")


def _kill_recorded_word(pid_file):
    try:
        raw = pid_file.read_text(encoding="ascii", errors="ignore")
    except OSError:
        return
    for token in raw.replace("\n", ",").split(","):
        token = token.strip()
        if token.isdigit():
            subprocess.run(["taskkill", "/PID", token, "/F"], stdout=subprocess.DEVNULL,
                           stderr=subprocess.DEVNULL, check=False)


@unittest.skipUnless(os.name == "nt" and os.environ.get("VKR_TEST_WORD") == "1",
                     "реальный Word проверяется только при VKR_TEST_WORD=1 на Windows")
class RealWordOpensCleanedDocumentsTests(unittest.TestCase):

    def test_cleaned_documents_open_in_word_without_repair(self):
        powershell = shutil.which("powershell.exe") or shutil.which("powershell")
        self.assertIsNotNone(powershell, "нужен PowerShell")
        with tempfile.TemporaryDirectory(prefix="vkr633-word-open-") as td:
            work = Path(td)
            original = work / "original.docx"
            shutil.copyfile(str(WORD_SAVED), str(original))
            cleaned = work / "cleaned.docx"
            shutil.copyfile(str(WORD_SAVED), str(cleaned))
            cleaned_rsid = work / "cleaned-strip-rsid.docx"
            shutil.copyfile(str(WORD_SAVED), str(cleaned_rsid))
            revision = work / "cleaned-revision.docx"
            shutil.copyfile(str(WORD_SAVED), str(revision))
            inject_revision_into_word_fixture(revision)
            template = work / "cleaned-template-mpgu.docx"
            shutil.copyfile(str(SKILL / "assets" / "template-mpgu.docx"), str(template))

            for path, extra in ((cleaned, []), (cleaned_rsid, ["--strip-rsid"]), (revision, ["--strip-rsid"]),
                                (template, [])):
                code, out = run_cleaner(path, "--in-place", "--json", *extra)
                self.assertEqual(0, code, out)

            documents = [original, cleaned, cleaned_rsid, revision, template]
            list_file = work / "files.txt"
            list_file.write_text("\n".join(str(p) for p in documents), encoding="utf-8")
            script = work / "open.ps1"
            script.write_text(WORD_OPEN_SCRIPT, encoding="utf-8-sig")
            out_file = work / "result.json"
            pid_file = work / "word-pid.txt"
            command = [powershell, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File",
                       str(script), "-ListFile", str(list_file), "-OutFile", str(out_file), "-PidFile", str(pid_file)]
            try:
                completed = subprocess.run(command, capture_output=True, timeout=300, check=False)
            except subprocess.TimeoutExpired:
                _kill_recorded_word(pid_file)
                self.fail("Word не завершил работу за 300 с; остановлен только свой экземпляр")
            self.assertTrue(out_file.is_file(), completed.stdout.decode("utf-8", "replace")
                            + completed.stderr.decode("utf-8", "replace"))
            results = json.loads(out_file.read_text(encoding="utf-8-sig"))
            by_name = {Path(item["path"]).name: item for item in results}
            report = json.dumps(results, ensure_ascii=False, indent=1)
            sys.stderr.write("\nWord open results:\n%s\n" % report)
            for path in documents:
                with self.subTest(document=path.name):
                    self.assertEqual("OPEN_OK", by_name[path.name]["status"], report)
            for path in (cleaned, cleaned_rsid):
                self.assertEqual(by_name["original.docx"]["pages"], by_name[path.name]["pages"], report)
                self.assertEqual(by_name["original.docx"]["paragraphs"], by_name[path.name]["paragraphs"], report)
            self.assertEqual(by_name["original.docx"]["paragraphs"] + 1,
                             by_name["cleaned-revision.docx"]["paragraphs"], report)


if __name__ == "__main__":
    unittest.main()
