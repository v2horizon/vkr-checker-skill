"""Список литературы 6.33: сортировка по методичке, форматы записей, ссылки, CLI.

Запуск: python -m unittest discover -s tests -t .   или   python tests/test_bibliography_633.py
"""

from __future__ import annotations

import importlib.util
import json
import os
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

os.environ.setdefault("PYTHONDONTWRITEBYTECODE", "1")
sys.dont_write_bytecode = True

ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "vkr-mpgu"
SCRIPTS = SKILL / "scripts"


def load_module(name: str):
    spec = importlib.util.spec_from_file_location(f"{name}_under_test_bib", SCRIPTS / f"{name}.py")
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


fb = load_module("format_bibliography")


def fmt(record: dict) -> str:
    return fb.format_source(fb.normalize_source(record, 1), [])


def order(records: list, **kwargs) -> list:
    result = fb.number_sources(fb.normalize_input_data(records), **kwargs)
    return [entry["id"] for entry in result["entries"]]


def run_cli(*args, input_text=None):
    env = dict(os.environ)
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    return subprocess.run(
        [sys.executable, str(SCRIPTS / "format_bibliography.py"), *map(str, args)],
        capture_output=True, text=True, encoding="utf-8", env=env, input=input_text,
    )


class DocumentationExamplesTest(unittest.TestCase):
    """Каждая пара «запись → строка» в gost-citations.md — реальный вывод скрипта."""

    PAIR = re.compile(r"```json\n(?P<json>(?:(?!```).)*?)\n```\n```text\n(?P<text>(?:(?!```).)*?)\n```", re.S)

    def test_every_documented_example_matches_script_output(self):
        doc = (SKILL / "references" / "gost-citations.md").read_text(encoding="utf-8")
        pairs = list(self.PAIR.finditer(doc))
        self.assertGreaterEqual(len(pairs), 40)
        for match in pairs:
            value = json.loads(match.group("json"))
            if isinstance(value, list):
                result = fb.number_sources(fb.normalize_input_data(value))
                actual = "\n".join(f"{e['number']}. {e['text']}" for e in result["entries"])
            else:
                actual = fmt(value)
            with self.subTest(example=match.group("json")[:80]):
                self.assertEqual(match.group("text"), actual)

    def test_key_samples_from_methodology_are_reproduced(self):
        doc = (SKILL / "references" / "gost-citations.md").read_text(encoding="utf-8")
        for sample in (
            "Гарифуллина, А.М. Аллюзии как трансляторы культурных кодов [Текст] / А.М. Гарифуллина // Вестник ВЭГУ. – 2001. – № 4 (54). – С. 74–79.",
            "Большой толковый словарь русского языка [Текст] / Сост. и гл. ред. С.А. Кузнецов. – СПб.: Норинт, 2000. – 1536 с.",
            "Иванов, А.Б. Название статьи [Текст] / А.Б. Иванов // Название сборника: сб. статей / под ред. В.Г. Петрова. – М.: Наука, 2020. – С. 45–58.",
            "Иванов, А.Б. Название статьи [Текст] / А.Б. Иванов // Аргументы и факты. – 2023. – 15 марта. – С. 3.",
        ):
            self.assertIn(sample, doc)

    def test_status_dictionary_matches_vkr_common(self):
        common = load_module("vkr_common")
        self.assertEqual(tuple(common.SOURCE_STATUSES), tuple(fb.STATUS_VALUES))
        for value in ("Verified", "CONFIRMED", "pending", "rejected", "suspicious"):
            self.assertEqual(common.normalize_source_status(value), fb.normalize_status(value))

    def test_documentation_has_single_ordering_rule(self):
        doc = (SKILL / "references" / "gost-citations.md").read_text(encoding="utf-8")
        self.assertNotIn("по убыванию юридической силы", doc.replace("«по юридической силе» скрипт не определяет", ""))
        self.assertNotIn("В конце — электронные ресурсы", doc)
        self.assertNotIn("--preserve-ids -o", doc)
        for status in fb.STATUS_VALUES:
            self.assertIn(f"`{status}`", doc)


class SortingTest(unittest.TestCase):
    RECORDS = [
        {"id": 3, "type": "book", "authors": ["Яковлев А.В."], "title": "Методика", "place": "М.", "publisher": "Наука", "year": 2019},
        {"id": "smith", "type": "article", "authors": ["Smith J."], "title": "Immersive learning", "journal": "Computers & Education", "year": 2019},
        {"id": "yolkin", "type": "article", "authors": ["Ёлкин Д.А."], "title": "Виртуальная реальность", "journal": "Информатика и образование", "year": 2021},
        {"id": "law", "type": "normative", "title": "Об образовании в Российской Федерации", "doc_type": "федер. закон"},
        {"id": "egorov", "type": "book", "authors": ["Егоров И.П."], "title": "Мышление", "place": "СПб.", "publisher": "Питер", "year": 2020},
        {"id": "zhukov-e", "type": "electronic", "authors": ["Жуков С.С."], "title": "Обзор гарнитур", "url": "https://example.org", "access_date": "01.02.2026"},
        {"id": "unity", "type": "software_doc", "authors": ["Unity Technologies"], "title": "Unity Manual", "url": "https://docs.unity3d.com", "access_date": "01.02.2026"},
        {"id": "quoted", "type": "book", "title": "«Цифровая школа»: опыт внедрения", "place": "М.", "publisher": "Наука", "year": 2020},
        {"id": "four", "type": "book", "authors": ["Абрамов А.А.", "Борисов Б.Б.", "Васильев В.В.", "Григорьев Г.Г."], "title": "Коллективный труд", "place": "М.", "publisher": "Наука", "year": 2018},
        {"id": "digit", "type": "article", "title": "1С:Предприятие в вузе", "journal": "Информатика и образование", "year": 2021},
    ]

    def test_order_follows_methodology_rule(self):
        # цифра первой; ё = е (Егоров < Ёлкин < Жуков); 4+ авторов — по заглавию (К);
        # закон — по заглавию (О); кавычки не учитываются (Ц); кириллица раньше латиницы;
        # русский электронный ресурс стоит по алфавиту, а не в конце (R8, R-F-05)
        self.assertEqual(
            ["digit", "egorov", "yolkin", "zhukov-e", "four", "law", "quoted", 3, "smith", "unity"],
            order(self.RECORDS),
        )

    def test_group_normative_first_is_an_option(self):
        self.assertEqual("law", order(self.RECORDS, group_normative_first=True)[0])
        self.assertEqual("digit", order(self.RECORDS)[0])

    def test_ids_opposite_to_alphabet_are_renumbered(self):
        # R3-04: id противоположны алфавиту, иначе сортировка и сохранение id неразличимы
        sources = fb.normalize_input_data([
            {"id": 1, "type": "book", "authors": ["Яковлев А.А."], "title": "Яблоки", "place": "М.", "publisher": "Наука", "year": 2020},
            {"id": 2, "type": "book", "authors": ["Абрамов Б.Б."], "title": "Арбузы", "place": "М.", "publisher": "Наука", "year": 2021},
        ])
        text, mapping = fb.format_bibliography(sources)
        lines = text.split("\n\n")
        self.assertTrue(lines[0].startswith("1. Абрамов, Б.Б. Арбузы"), lines)
        self.assertTrue(lines[1].startswith("2. Яковлев, А.А. Яблоки"), lines)
        self.assertEqual({2: 1, 1: 2}, mapping)
        result = fb.number_sources(sources)
        self.assertEqual({"2": 1, "1": 2}, result["mapping"])

    def test_same_author_sorted_by_title_then_year(self):
        records = [
            {"id": "b", "type": "book", "authors": ["Иванов А.Б."], "title": "Яблоко", "place": "М.", "publisher": "Н", "year": 2020},
            {"id": "a2", "type": "book", "authors": ["Иванов А.Б."], "title": "Арбуз", "place": "М.", "publisher": "Н", "year": 2021},
            {"id": "a1", "type": "book", "authors": ["Иванов А.Б."], "title": "Арбуз", "place": "М.", "publisher": "Н", "year": 2019},
            {"id": "c", "type": "book", "authors": ["Иванова С.В."], "title": "Арбуз", "place": "М.", "publisher": "Н", "year": 2019},
        ]
        self.assertEqual(["a1", "a2", "b", "c"], order(records))

    def test_duplicate_ids_are_an_error(self):
        with self.assertRaises(fb.BibliographyError):
            fb.number_sources(fb.normalize_input_data([
                {"id": 1, "type": "book", "title": "А"}, {"id": 1, "type": "book", "title": "Б"}]))
        with self.assertRaises(fb.BibliographyError):  # запись без id получает id = номер в реестре
            fb.number_sources(fb.normalize_input_data([{"type": "book", "title": "А"}, {"id": 1, "type": "book", "title": "Б"}]))

    def test_preserve_ids_requires_integer_ids(self):
        with self.assertRaises(fb.BibliographyError):
            fb.format_bibliography(fb.normalize_input_data([{"id": "src-2", "type": "book", "title": "А"}]), preserve_ids=True)
        text, mapping = fb.format_bibliography(
            fb.normalize_input_data([{"id": 5, "type": "book", "title": "Б"}, {"id": 1, "type": "book", "title": "А"}]),
            preserve_ids=True)
        self.assertTrue(text.startswith("5. Б"))
        self.assertEqual({}, mapping)


class FormattingTest(unittest.TestCase):
    def test_editors_as_strings_list_and_single_string(self):
        # R5: строки в editors не роняют скрипт
        for editors in (["С.А. Кузнецов"], "Кузнецов С.А.", [{"last": "Кузнецов", "initials": "С.А."}]):
            text = fmt({"type": "book", "title": "Сборник", "editors": editors, "place": "М.", "publisher": "Наука", "year": 2020})
            self.assertIn("под ред. С.А. Кузнецова", text)
        text = fmt({"type": "collection_article", "authors": ["Иванов А.Б."], "title": "Статья", "container_title": "Сборник",
                    "editor": "Петров В.Г.", "place": "М.", "publisher": "Наука", "year": 2020, "pages": "1-5"})
        self.assertIn("/ под ред. В.Г. Петрова", text)
        text = fmt({"type": "book", "title": "Сборник", "editors": ["Дюма А."], "editors_text": "под ред. А. Дюма", "place": "М.", "publisher": "Н", "year": 2020})
        self.assertIn("под ред. А. Дюма", text)

    def test_collection_conference_and_newspaper_samples(self):
        # R6: [Текст], родительный падеж редактора, дата газеты перед страницами, Vol./P. у иностранных
        text = fmt({"type": "conference", "authors": ["Smith J."], "title": "VR in classrooms", "container_title": "Proceedings",
                    "place": "New York", "publisher": "IEEE", "year": 2023, "volume": 2, "pages": "10-20"})
        self.assertIn("– Vol. 2. – P. 10–20.", text)
        self.assertNotIn("Т. 2", text)
        news = fmt({"type": "newspaper", "authors": ["Иванов А.Б."], "title": "Статья", "newspaper": "Газета", "year": 2023,
                    "date": "15 марта", "number": 11, "pages": 3})
        self.assertIn("[Текст]", news)
        self.assertTrue(news.endswith("– 2023. – 15 марта (№ 11). – С. 3."), news)

    def test_foreign_electronic_resource_uses_russian_marker(self):
        # R7, E-new2
        text = fmt({"type": "electronic", "authors": ["Fielding R."], "title": "HTTP Semantics", "url": "https://example.org", "access_date": "01.02.2026"})
        self.assertIn("[Электронный ресурс]", text)
        self.assertNotIn("Electronic resource", text)

    def test_language_by_share_of_cyrillic_not_first_letter(self):
        # E9
        book = fmt({"type": "book", "title": "ASP.NET Core MVC: руководство", "place": "М.", "publisher": "ДМК Пресс", "year": 2022, "pages": 300})
        self.assertIn("[Текст]", book)
        self.assertTrue(book.endswith("300 с."), book)
        article = fmt({"type": "article", "title": "Docker и Kubernetes: CI/CD для микросервисов", "journal": "Системный администратор",
                       "year": 2023, "number": 5, "pages": "10-15"})
        self.assertIn("№ 5. – С. 10–15.", article)
        quoted = fmt({"type": "book", "title": "«Цифровая школа»: опыт внедрения", "place": "М.", "publisher": "Наука", "year": 2020, "pages": 100})
        self.assertIn("[Текст]", quoted)
        english = fmt({"type": "book", "title": "Clean Code", "language": "en", "place": "Boston", "publisher": "Pearson", "year": 2008, "pages": 431})
        self.assertTrue(english.endswith("431 p."), english)

    def test_electronic_article_keeps_volume_number_pages_doi_and_url(self):
        # E7: DOI не теряется при наличии URL, разделитель «. – С.»
        text = fmt({"type": "electronic", "authors": ["Иванов А.Б."], "title": "Статья", "journal": "Журнал", "year": 2023, "volume": 15,
                    "number": 2, "pages": "45-58", "doi": "10.1/x", "url": "https://example.org", "access_date": "01.02.2024"})
        self.assertIn("– Т. 15, № 2. – С. 45–58. – DOI: 10.1/x. – URL: https://example.org (дата обращения: 01.02.2024).", text)

    def test_article_with_place_and_publisher_does_not_lose_them(self):
        text = fmt({"type": "article", "authors": ["Иванов А.Б."], "title": "Статья", "journal": "Сборник", "place": "М.", "publisher": "Наука", "year": 2020, "pages": "5-9"})
        self.assertIn("– М.: Наука, 2020. – С. 5–9.", text)

    def test_all_spec_types_are_supported(self):
        minimal = {
            "book": {"title": "Т"}, "article": {"title": "Т", "journal": "Ж"}, "collection_article": {"title": "Т", "container_title": "С"},
            "conference": {"title": "Т", "container_title": "С"}, "newspaper": {"title": "Т", "newspaper": "Г"},
            "electronic": {"title": "Т", "url": "https://e.org"}, "dissertation": {"title": "Т", "authors": ["Иванов А.Б."], "place": "М."},
            "autoreferat": {"title": "Т", "authors": ["Иванов А.Б."], "place": "М."}, "normative": {"title": "Т"},
            "standard": {"title": "Т", "designation": "ГОСТ 1"}, "software_doc": {"title": "Т", "url": "https://e.org"},
            "repository": {"title": "Т", "url": "https://e.org"}, "preprint": {"title": "Т", "url": "https://e.org"},
            "dataset": {"title": "Т", "url": "https://e.org"}, "patent": {"title": "Т", "number": "1"},
            "video": {"title": "Т", "url": "https://e.org"},
        }
        self.assertEqual(set(fb.CANONICAL_TYPES), set(minimal))
        for kind, fields in minimal.items():
            with self.subTest(kind=kind):
                text = fmt(dict(fields, type=kind))
                self.assertTrue(text.endswith("."), text)
        self.assertIn("Пат. 1", fmt({"type": "patent", "title": "Т", "number": "1"}))
        self.assertIn("репозиторий", fmt({"type": "repository", "title": "Т", "url": "https://e.org"}))
        with self.assertRaises(fb.BibliographyError):
            fmt({"type": "report", "title": "Т"})

    def test_incomplete_inputs_give_clear_output_or_error(self):
        # E6: книга без выходных данных, диссертация без автора, автор без инициалов
        warnings = []
        text = fb.format_source(fb.normalize_source({"type": "book", "authors": ["Иванов А.Б."], "title": "Книга"}, 1), warnings)
        self.assertEqual("Иванов, А.Б. Книга [Текст] / А.Б. Иванов. – [Б. м.]: [б. и.], [б. г.].", text)
        self.assertTrue(any("нет места" in w for w in warnings))
        with self.assertRaises(fb.BibliographyError):
            fmt({"type": "dissertation", "title": "Тема"})
        warnings = []
        text = fb.format_source(fb.normalize_source({"type": "book", "authors": ["Иванов"], "title": "Книга", "place": "М.", "publisher": "Н", "year": 2020}, 1), warnings)
        self.assertTrue(text.startswith("Иванов. Книга"), text)
        self.assertTrue(any("нет инициалов" in w for w in warnings))
        warnings = []
        fb.format_source(fb.normalize_source({"type": "electronic", "title": "Сайт", "url": "https://e.org"}, 1), warnings)
        self.assertTrue(any("даты обращения" in w for w in warnings))
        with self.assertRaises(fb.BibliographyError):
            fmt({"type": "electronic", "title": "Сайт"})
        with self.assertRaises(fb.BibliographyError):
            fb.format_source({"id": 5, "raw_text": "Иванов И.И. Книга. М., 2020."})

    def test_organization_authors(self):
        self.assertTrue(fmt({"type": "electronic", "authors": ["W3C"], "title": "WCAG", "url": "https://w3.org"}).startswith("WCAG [Электронный ресурс] / W3C."))
        self.assertTrue(fmt({"type": "book", "authors": ["Министерство просвещения РФ"], "title": "Рекомендации", "place": "М.", "publisher": "Н", "year": 2021}).startswith("Рекомендации [Текст] / Министерство просвещения РФ."))
        self.assertTrue(fmt({"type": "book", "authors": ["Римский-Корсаков Н.А."], "title": "Летопись", "place": "М.", "publisher": "Н", "year": 1909}).startswith("Римский-Корсаков, Н.А. Летопись"))

    def test_latin_surnames_and_initials_with_diacritics(self):  # B15 (F05)
        cases = {
            "Rößling G.": ("Rößling", "G."), "Müller J.": ("Müller", "J."), "G. Rößling": ("Rößling", "G."),
            "Müller, J.-P.": ("Müller", "J.-P."), "Łukasiewicz J.": ("Łukasiewicz", "J."), "Öztürk Ö.": ("Öztürk", "Ö."),
            "Иванов А.Б.": ("Иванов", "А.Б."), "O’Brien P.": ("O’Brien", "P."),
        }
        for text, (last, initials) in cases.items():
            with self.subTest(author=text):
                parsed = fb._parse_person_string(text)
                self.assertEqual((last, initials), (parsed["last"], parsed["initials"]))
                self.assertFalse(parsed.get("is_organization"))
        self.assertTrue(fb._parse_person_string("иванов а.б.").get("is_organization"))  # строчные «инициалы» — не инициалы
        entry = fmt({"type": "article", "authors": ["Rößling G.", "Müller J."], "title": "Visualizing sorting algorithms",
                     "journal": "Computers & Education", "year": 2020, "volume": 140, "pages": "1-12"})
        self.assertTrue(entry.startswith("Rößling, G. Visualizing sorting algorithms / G. Rößling, J. Müller // Computers & Education."), entry)
        result = fb.number_sources(fb.normalize_input_data([
            {"id": "mueller", "type": "book", "authors": ["Müller J."], "title": "Didaktik", "place": "Berlin", "publisher": "Springer", "year": 2019},
            {"id": "roessling", "type": "book", "authors": ["Rößling G."], "title": "Animation", "place": "Berlin", "publisher": "Springer", "year": 2018},
        ]))
        self.assertEqual(["mueller", "roessling"], [item["id"] for item in result["entries"]])


class CitationTest(unittest.TestCase):
    def test_citation_forms(self):
        text = ("См. [@ivanov2020], [@ivanov2020, с. 23–25], [@a; @b], [Цит. по: @c, с. 5], [3, с. 17], [1–3], "
                "`arr[2]`, [2023], [N], [ПРОВЕРИТЬ: @x], [x].")
        refs = fb.citation_refs(text)
        self.assertEqual([("id", "ivanov2020"), ("id", "ivanov2020"), ("id", "a"), ("id", "b"), ("id", "c"),
                          ("num", 3), ("num", 1), ("num", 2), ("num", 3)], refs)

    def test_rewrite_keeps_locators_and_prefix(self):
        numbers = {"ivanov2020": 4, "a": 1, "b": 2, "c": 9}
        resolve = lambda kind, ref: numbers.get(ref) if kind == "id" else {3: 7}.get(ref)
        text = "[@ivanov2020, с. 23–25] и [@a; @b], [Цит. по: @c, с. 5], [3, с. 17], `x[3]`."
        self.assertEqual("[4, с. 23–25] и [1; 2], [Цит. по: 9, с. 5], [7, с. 17], `x[3]`.", fb.rewrite_citations(text, resolve))
        problems = []
        self.assertEqual("[@missing]", fb.rewrite_citations("[@missing]", resolve, problems))
        self.assertTrue(problems)


class CliTest(unittest.TestCase):
    def write(self, folder: Path, name: str, value, bom: bool = False) -> Path:
        path = folder / name
        text = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False)
        path.write_bytes((b"\xef\xbb\xbf" if bom else b"") + text.encode("utf-8"))
        return path

    def test_exit_codes_and_json_output(self):
        with tempfile.TemporaryDirectory() as td:
            folder = Path(td)
            good = self.write(folder, "sources.json", SortingTest.RECORDS, bom=True)
            proc = run_cli(good, "--json", "--mapping", folder / "map.json")
            self.assertEqual(0, proc.returncode, proc.stderr)
            payload = json.loads(proc.stdout)
            self.assertEqual("ok", payload["status"])
            self.assertEqual("digit", payload["entries"][0]["id"])
            self.assertEqual(1, json.loads((folder / "map.json").read_text(encoding="utf-8"))["digit"])

            proc = run_cli(folder / "missing.json")
            self.assertEqual(2, proc.returncode)
            broken = self.write(folder, "broken.json", "[{")
            self.assertEqual(2, run_cli(broken).returncode)

            bad = self.write(folder, "bad.json", [{"id": 1, "type": "book", "editors": 5, "title": "Т"}])
            proc = run_cli(bad)
            self.assertEqual(1, proc.returncode)
            self.assertNotIn("Traceback", proc.stderr)
            self.assertIn("editors", proc.stderr)

            dup = self.write(folder, "dup.json", [{"id": 1, "type": "book", "title": "А"}, {"id": 1, "type": "book", "title": "Б"}])
            self.assertEqual(1, run_cli(dup, "--preserve-ids").returncode)

            unsorted = self.write(folder, "unsorted.json", [
                {"id": 1, "type": "book", "authors": ["Яковлев А.А."], "title": "Я", "place": "М.", "publisher": "Н", "year": 2020},
                {"id": 2, "type": "book", "authors": ["Абрамов Б.Б."], "title": "А", "place": "М.", "publisher": "Н", "year": 2020}])
            proc = run_cli(unsorted, "--preserve-ids")
            self.assertEqual(0, proc.returncode)
            self.assertIn("не алфавитный", proc.stderr)
            self.assertTrue(proc.stdout.startswith("1. Яковлев"))
            proc = run_cli(unsorted, "-o", folder / "out" / "list.txt")
            self.assertEqual(0, proc.returncode)
            self.assertTrue((folder / "out" / "list.txt").read_text(encoding="utf-8").startswith("1. Абрамов"))

    def test_help_works(self):
        proc = run_cli("--help")
        self.assertEqual(0, proc.returncode)
        self.assertIn("--group-normative-first", proc.stdout)


if __name__ == "__main__":
    unittest.main()
