"""Контракт документации 6.33 (WP6): документы совпадают с кодом и друг с другом.

Проверяется:
1. каждая команда ``python <SKILL_DIR>/scripts/X.py …`` (и сокращённые формы
   ``X.py …`` в строчном коде, и команды через переменные ``$A …``) ссылается на
   существующий скрипт, а все её флаги есть в ``--help`` скрипта или подкоманды;
2. ссылки на ``references/``, ``scripts/``, ``assets/``, ``data/``, ``agents/`` и
   относительные Markdown-ссылки существуют; дерево «Структура скилла» в SKILL.md
   совпадает с файлами;
3. версия ``6.33-production`` одинакова в SKILL.md, PORTABLE-PROMPT.md,
   INSTALL-RU.md, README.md и в коде;
4. в документах нет запрещённых наследий прежних версий;
5. шкала «срок → режим» одна и та же во всех файлах;
6. YAML frontmatter SKILL.md валиден, description не длиннее 1024 символов;
7. тексты универсальны: «Claude» встречается только как название платформы
   (Claude Code, Claude Desktop, Claude.ai, ``anthropic-skills:vkr-mpgu``,
   Claude Project в скобочном перечне контекстов), а не как обозначение
   исполнителя;
8. файлы точек входа для других ИИ-систем существуют, ссылаются на
   ``vkr-mpgu/SKILL.md``, называют версию и не хранят своей копии
   ``AI_AGENT_BOOTSTRAP``.

Запуск: ``python -m unittest discover -s tests -t .`` или
``python tests/test_docs_contract_633.py``. Тест не ходит в сеть и ничего не пишет.
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import unittest
from pathlib import Path
from typing import Dict, Iterator, List, Optional, Set, Tuple

os.environ["PYTHONDONTWRITEBYTECODE"] = "1"
sys.dont_write_bytecode = True

ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "vkr-mpgu"
SCRIPTS = SKILL / "scripts"
VERSION = "6.33-production"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

WP6_DOCS = (
    "vkr-mpgu/SKILL.md",
    "vkr-mpgu/INSTALL-RU.md",
    "vkr-mpgu/PORTABLE-PROMPT.md",
    "README.md",
    "vkr-mpgu/references/quickstart.md",
    "vkr-mpgu/references/writing-workflow.md",
    "vkr-mpgu/references/writing-workflow-express.md",
    "vkr-mpgu/references/draft-polish-workflow.md",
    "vkr-mpgu/references/intake-interview.md",
    "vkr-mpgu/references/submission-checklist.md",
    "vkr-mpgu/references/failure-recovery.md",
    "vkr-mpgu/references/friend-sharing.md",
    "vkr-mpgu/references/originality-techniques.md",
    "vkr-mpgu/references/requirements-precedence.md",
    "vkr-mpgu/references/structure-project.md",
    "vkr-mpgu/references/structure-regular.md",
    "vkr-mpgu/references/methodology-mpgu.md",
    "vkr-mpgu/references/defense.md",
    "vkr-mpgu/references/defense-mastery.md",
    "vkr-mpgu/references/platform-integration.md",
    "vkr-mpgu/references/project-templates.md",
    "vkr-mpgu/references/humanizer-techniques.md",
)

SUBCOMMAND_SCRIPTS = {
    "vkr_audit.py": ("validate", "snapshot", "plan", "brief", "record", "close", "findings", "accept-risk", "waive", "migrate-legacy", "next"),
    "vkr_memory.py": ("status", "record"),
}

FORBIDDEN = (
    ("final/vkr-final.docx", "сдаваемый файл один: final/vkr.docx"),
    ("final/vkr-submit.docx", "сдаваемый файл один: final/vkr.docx"),
    ("final/vkr-draft.docx", "промежуточные DOCX — только в exports/"),
    ("vkr-input.json", "JSON для генератора собирает build_vkr.py"),
    ("Всегда загружается автоматически", "ложное платформенное утверждение"),
    ("200K", "ложное платформенное утверждение о лимите контекста"),
    ("conversation_search", "инструмент конкретной платформы вместо файлов проекта"),
    ("Claude Projects", "инструмент конкретной платформы вместо файлов проекта"),
    ("WAITING_SUPERVISOR_APPROVAL", "статус удалён: SUPERVISOR_APPROVAL — fix_class, этап roadmap WAITING_SUPERVISOR"),
    ("подозрительно для нормоконтроля", "придуманный потолок оригинальности"),
    ("сокращённый стандарт", "название режима расходится со шкалой сроков"),
    ("audit/fix-plan.md", "файлы в audit/ вручную не создаются"),
)

DAY_RANGE = re.compile(r"(\d+)\s*[–-]\s*(\d+)\s*(?:дн|дня|дней)")
DAY_PLUS = re.compile(r"(\d+)\+\s*(?:дн|дня|дней)")
DAY_OR_MORE = re.compile(r"(\d+)\s*(?:дн|дня|дней)\w*\s+и\s+больше")
MODE_TOKEN = re.compile(r"(?<![\w-])(standard|express-14d|express-7d|express-4d)(?![\w-])")
ALLOWED_DAYS = {
    "standard": {(29, None), (15, 28), (15, None)},
    "express-14d": {(8, 14)},
    "express-7d": {(5, 7)},
    "express-4d": {(3, 4)},
}
CANONICAL_ROWS = (
    ("29+ дней", "standard"),
    ("15–28 дней", "standard"),
    ("8–14 дней", "express-14d"),
    ("5–7 дней", "express-7d"),
    ("3–4 дня", "express-4d"),
)
BASE_ROLES = ("MET", "SRC", "LOG", "LNG", "STY", "EVD", "TEC", "DOC", "OWN")
PROJECT_MD_NAMES = {
    "vkr-state.md", ".vkr-state.md", "plan.md", "handoff.md", "roadmap.md", "annotation.md",
    "introduction.md", "conclusion.md", "mastery-log.md",
}
PROJECT_MD_PATTERNS = (re.compile(r"chapter-(?:\d+|N)\.md"), re.compile(r"appendix-(?:\d+|N)\.md"), re.compile(r"test30-[\w-]*\.md"))

# Универсальность: имя платформы допустимо, обозначение исполнителя — нет.
# Пробел заменён на \s+, чтобы название пережило перенос строки в тексте.
CLAUDE_PLATFORM = re.compile(
    r"Claude\s+Code"               # агентный CLI
    r"|Claude\s+Desktop"           # десктоп-клиент
    r"|[Cc]laude\.ai"              # веб-интерфейс
    r"|anthropic-skills:vkr-mpgu"  # имя навыка, подключённого в аккаунте
)
CLAUDE_PROJECT = re.compile(r"Claude Projects?")
PAREN_SPAN = re.compile(r"\([^()]*\)")
CLAUDE_ANY = re.compile(r"Claude")
ENTRY_POINTS = (
    "AGENTS.md",
    "CLAUDE.md",
    "GEMINI.md",
    ".cursor/rules/vkr-mpgu.mdc",
    ".github/copilot-instructions.md",
)
LINK = re.compile(r"\]\(([^)\s]+)\)")


def rel(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def doc_files() -> List[Path]:
    files = [path for path in SKILL.rglob("*.md") if "__pycache__" not in path.parts]
    return sorted(files) + [ROOT / "README.md"]


def read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def skill_files() -> Set[str]:
    return {
        path.relative_to(SKILL).as_posix()
        for path in SKILL.rglob("*")
        if path.is_file() and "__pycache__" not in path.parts and path.suffix != ".pyc"
    }


_HELP: Dict[Tuple[str, Optional[str]], Tuple[int, str]] = {}


def script_help(script: str, subcommand: Optional[str] = None) -> Tuple[int, str]:
    key = (script, subcommand)
    if key not in _HELP:
        args = [sys.executable, str(SCRIPTS / script)]
        if subcommand:
            args += ["PROJECT_DIR", subcommand]
        args.append("--help")
        env = os.environ.copy()
        env["PYTHONDONTWRITEBYTECODE"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        result = subprocess.run(
            args, cwd=str(ROOT), env=env, capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=180,
        )
        _HELP[key] = (result.returncode, result.stdout + result.stderr)
    return _HELP[key]


def help_flags(script: str, subcommand: Optional[str] = None) -> Set[str]:
    code, text = script_help(script, subcommand)
    if code != 0:
        raise AssertionError(f"{script} {subcommand or ''} --help вернул {code}: {text[-400:]}")
    return set(re.findall(r"(?<![\w-])(--?[A-Za-z][\w-]*)", text))


class Command:
    def __init__(self, source: str, line: int, script: str, args: str) -> None:
        self.source = source
        self.line = line
        self.script = script
        self.args = args

    def label(self) -> str:
        return f"{self.source}:{self.line}: {self.script} {self.args}".strip()

    def tokens(self) -> List[str]:
        return self.args.split()

    def flags(self) -> List[str]:
        found = []
        for token in self.tokens():
            token = token.strip("[](),;")
            match = re.match(r"^(--?[A-Za-z][\w-]*)(?:=.*)?$", token)
            if match:
                found.append(match.group(1))
        return found

    def subcommand(self) -> Optional[str]:
        choices = SUBCOMMAND_SCRIPTS.get(self.script)
        if not choices:
            return None
        for token in self.tokens():
            if token in choices:
                return token
        return None


def _cut_comment(text: str) -> str:
    return re.split(r"\s#\s", text, maxsplit=1)[0]


def iter_commands(path: Path) -> Iterator[Command]:
    source = rel(path)
    text = read(path)
    variables: Dict[str, str] = {}
    for match in re.finditer(r'^\s*([A-Z][A-Z_]*)="python3? <SKILL_DIR>/scripts/([\w-]+\.py)(?: [^"]*)?"', text, re.MULTILINE):
        variables[match.group(1)] = match.group(2)
    in_fence = False
    for number, line in enumerate(text.splitlines(), 1):
        if re.match(r"^\s*(```|~~~)", line):
            in_fence = not in_fence
            continue
        for match in re.finditer(r"python3? <SKILL_DIR>/scripts/([\w-]+\.py)", line):
            before = line[: match.start()]
            rest = line[match.end():]
            if before.count("`") % 2 == 1:
                rest = rest.split("`", 1)[0]
            yield Command(source, number, match.group(1), _cut_comment(rest).rstrip('"'))
        if in_fence:
            var = re.match(r"^\s*\$([A-Z][A-Z_]*)\s+(.*)$", line)
            if var and var.group(1) in variables:
                yield Command(source, number, variables[var.group(1)], _cut_comment(var.group(2)))
        else:
            for span in re.findall(r"`([^`\n]+)`", line):
                short = re.match(r"^([\w-]+\.py)(\s.*)?$", span.strip())
                if short and not short.group(1).startswith("test_"):
                    yield Command(source, number, short.group(1), short.group(2) or "")


def parse_skill_tree(text: str) -> Set[str]:
    section = text.split("## Структура скилла", 1)[1]
    block = re.search(r"```[a-z]*\n(.*?)```", section, re.S)
    if not block:
        raise AssertionError("в разделе «Структура скилла» нет блока с деревом")
    lines = block.group(1).splitlines()
    if not lines or lines[0].strip() != "vkr-mpgu/":
        raise AssertionError("дерево должно начинаться с vkr-mpgu/")
    stack: List[str] = []
    files: Set[str] = set()
    for line in lines[1:]:
        match = re.match(r"^((?:│   |    )*)(?:├── |└── )(\S+)", line)
        if not match:
            continue
        depth = len(match.group(1)) // 4
        name = match.group(2)
        stack = stack[:depth]
        if name.endswith("/"):
            stack.append(name.rstrip("/"))
        else:
            files.add("/".join(stack + [name]))
    return files


def frontmatter(text: str) -> str:
    if not text.startswith("---\n"):
        raise AssertionError("SKILL.md должен начинаться с YAML frontmatter")
    end = text.find("\n---\n", 4)
    if end < 0:
        raise AssertionError("frontmatter SKILL.md не закрыт строкой ---")
    return text[4:end]


def parse_frontmatter_strict(block: str) -> Dict[str, object]:
    """Разбор без PyYAML для ограниченной формы: скаляры в кавычках и один вложенный объект."""
    data: Dict[str, object] = {}
    current: Optional[Dict[str, object]] = None
    for line in block.splitlines():
        if not line.strip():
            continue
        if "\t" in line:
            raise ValueError("табуляция во frontmatter")
        top = re.match(r'^([a-z_]+):\s*(.*)$', line)
        nested = re.match(r'^  ([a-z_]+):\s*(.+)$', line)
        if top:
            key, value = top.group(1), top.group(2)
            if key in data:
                raise ValueError(f"повторяющийся ключ {key}")
            if value == "":
                current = {}
                data[key] = current
            else:
                current = None
                data[key] = _scalar(value)
        elif nested and current is not None:
            current[nested.group(1)] = _scalar(nested.group(2))
        else:
            raise ValueError(f"строка frontmatter не разобрана: {line!r}")
    return data


def _scalar(value: str) -> object:
    value = value.strip()
    if value.startswith('"'):
        return json.loads(value)
    if value.startswith("'") or value.startswith(("[", "{", "|", ">", "&", "*", "!")):
        raise ValueError(f"неподдерживаемая форма значения: {value[:20]}")
    if ": " in value or " #" in value:
        raise ValueError(f"значение без кавычек содержит спецсимволы YAML: {value[:40]}")
    return value


def mode_day_mismatches(path: Path) -> List[str]:
    problems = []
    for number, line in enumerate(read(path).splitlines(), 1):
        for segment in re.split(r"[|;,]", line):
            modes = set(MODE_TOKEN.findall(segment))
            ranges: List[Tuple[int, Optional[int]]] = []
            ranges += [(int(a), int(b)) for a, b in DAY_RANGE.findall(segment)]
            ranges += [(int(a), None) for a in DAY_PLUS.findall(segment)]
            ranges += [(int(a), None) for a in DAY_OR_MORE.findall(segment)]
            if len(modes) == 1 and ranges:
                mode = next(iter(modes))
                for item in ranges:
                    if item not in ALLOWED_DAYS[mode]:
                        problems.append(f"{rel(path)}:{number}: {mode} ↔ {item}: {line.strip()[:120]}")
            if re.search(r"плотн\w*\s+(?:темп\w*\s+)?(?:стандарт|`?standard`?)", segment, re.IGNORECASE) or re.search(r"`?standard`?\s+в\s+плотном", segment):
                for item in ranges:
                    if item != (15, 28):
                        problems.append(f"{rel(path)}:{number}: «плотный standard» ↔ {item}: {line.strip()[:120]}")
    return problems


def _blank(match: "re.Match[str]") -> str:
    """Замена совпадения точками той же длины: номера строк не смещаются."""
    return "".join("\n" if char == "\n" else "·" for char in match.group(0))


def mask_allowed_claude(text: str) -> str:
    masked = CLAUDE_PLATFORM.sub(_blank, text)
    parts: List[str] = []
    last = 0
    for paren in PAREN_SPAN.finditer(masked):
        parts.append(masked[last:paren.start()])
        parts.append(CLAUDE_PROJECT.sub(_blank, paren.group(0)))
        last = paren.end()
    parts.append(masked[last:])
    return "".join(parts)


class CommandContractTest(unittest.TestCase):
    def test_every_documented_command_uses_existing_script_and_flags(self) -> None:
        problems: List[str] = []
        checked = 0
        for path in doc_files():
            for command in iter_commands(path):
                checked += 1
                if not (SCRIPTS / command.script).is_file():
                    problems.append(f"{command.label()}: скрипта нет в scripts/")
                    continue
                subcommand = command.subcommand()
                flags = help_flags(command.script, subcommand)
                for flag in command.flags():
                    if flag not in flags:
                        where = f"{command.script} {subcommand}" if subcommand else command.script
                        problems.append(f"{command.label()}: флага {flag} нет в --help {where}")
        self.assertGreater(checked, 80, "команды в документах не найдены: парсер сломан")
        self.assertEqual([], problems)

    def test_python_file_names_in_docs_exist(self) -> None:
        existing = {path.name for path in SCRIPTS.glob("*.py")}
        problems = []
        for path in doc_files():
            for number, line in enumerate(read(path).splitlines(), 1):
                for name in re.findall(r"(?<![\w/.-])([A-Za-z_][\w-]*\.py)\b", line):
                    if name.startswith("test_"):
                        if not (ROOT / "tests" / name).is_file():
                            problems.append(f"{rel(path)}:{number}: нет tests/{name}")
                    elif name not in existing:
                        problems.append(f"{rel(path)}:{number}: нет scripts/{name}")
        self.assertEqual([], problems)

    def test_parser_sees_variables_subcommands_and_inline_forms(self) -> None:
        commands = list(iter_commands(SKILL / "references" / "continuous-audit.md"))
        self.assertTrue(any(c.script == "vkr_audit.py" and c.subcommand() == "record" and "--auditor-id" in c.flags() for c in commands))
        broken = Command("x.md", 1, "vkr_audit.py", "<PROJECT_DIR> plan --kind primary --preserve-ids")
        self.assertNotIn("--preserve-ids", help_flags("vkr_audit.py", broken.subcommand()))
        inline = list(iter_commands(SKILL / "SKILL.md"))
        self.assertTrue(any(c.script == "clean_docx_metadata.py" and "--in-place" in c.flags() for c in inline))


class ReferenceContractTest(unittest.TestCase):
    def test_skill_paths_in_docs_exist(self) -> None:
        pattern = re.compile(r"(?<![\w./<-])((?:vkr-mpgu/)?(?:references|scripts|assets|data|agents)/[A-Za-z0-9_.-]+(?:/[A-Za-z0-9_.-]+)*)")
        problems = []
        for path in doc_files():
            for number, line in enumerate(read(path).splitlines(), 1):
                for match in pattern.finditer(line):
                    target = match.group(1).rstrip(".,;:")
                    if target.startswith("vkr-mpgu/"):
                        target = target[len("vkr-mpgu/"):]
                    if target.endswith("/") or "<" in target:
                        continue
                    if not (SKILL / target).exists():
                        problems.append(f"{rel(path)}:{number}: нет {target}")
        self.assertEqual([], problems)

    def test_bare_reference_names_and_markdown_links_exist(self) -> None:
        references = {path.name for path in (SKILL / "references").glob("*.md")}
        top = {path.name for path in SKILL.glob("*.md")} | {"README.md"}
        problems = []
        for path in doc_files():
            text = read(path)
            for number, line in enumerate(text.splitlines(), 1):
                for name in re.findall(r"`([a-z][a-z0-9-]*\.md)`", line):
                    if name in references or name in top or name in PROJECT_MD_NAMES:
                        continue
                    if any(p.fullmatch(name) for p in PROJECT_MD_PATTERNS):
                        continue
                    problems.append(f"{rel(path)}:{number}: справочника {name} нет")
                for target in re.findall(r"\]\(([^)\s]+)\)", line):
                    if target.startswith(("http://", "https://", "#", "mailto:")) or "<" in target:
                        continue
                    if not target.endswith(".md") and not target.startswith("./"):
                        continue
                    resolved = (path.parent / target.split("#", 1)[0]).resolve()
                    if not resolved.exists():
                        problems.append(f"{rel(path)}:{number}: ссылка {target} ведёт в пустоту")
        self.assertEqual([], problems)

    def test_skill_tree_matches_files(self) -> None:
        tree = parse_skill_tree(read(SKILL / "SKILL.md"))
        actual = skill_files()
        self.assertEqual(sorted(actual - tree), [], "файлы есть на диске, но не в дереве SKILL.md")
        self.assertEqual(sorted(tree - actual), [], "файлы есть в дереве SKILL.md, но не на диске")

    def test_references_table_lists_new_references(self) -> None:
        text = read(SKILL / "SKILL.md")
        table = text.split("## Что читать для какой задачи", 1)[1].split("\n## ", 1)[0]
        for name in ("drafts-format.md", "analyzers-cli.md", "source-verification.md", "docx-input-schema.md",
                     "quickstart.md", "platform-integration.md", "friend-sharing.md"):
            self.assertIn(f"references/{name}", table, name)


class VersionContractTest(unittest.TestCase):
    def test_version_is_the_same_everywhere(self) -> None:
        import vkr_common
        import init_vkr_project

        self.assertEqual(VERSION, vkr_common.SKILL_VERSION)
        self.assertEqual(VERSION, init_vkr_project.SKILL_VERSION)
        skill = read(SKILL / "SKILL.md")
        data = parse_frontmatter_strict(frontmatter(skill))
        self.assertEqual(VERSION, data["metadata"]["version"])  # type: ignore[index]
        body = skill.split("\n---\n", 1)[1]
        self.assertIn(f"`{VERSION}`", body)
        for relative in ("vkr-mpgu/PORTABLE-PROMPT.md", "vkr-mpgu/INSTALL-RU.md", "README.md"):
            text = read(ROOT / relative)
            self.assertIn(VERSION, text, relative)
            self.assertIn(VERSION, text.splitlines()[0] + text[:600], f"{relative}: версия должна быть в начале файла")
        for relative in ("vkr-mpgu/SKILL.md", "vkr-mpgu/PORTABLE-PROMPT.md", "vkr-mpgu/INSTALL-RU.md", "README.md"):
            stale = re.findall(r"(?<![\w.])6\.(?:2\d|3[0-2])-[a-z][a-z-]*", read(ROOT / relative))
            self.assertEqual([], stale, f"{relative}: упоминание прежней версии как текущей")


class LegacyContractTest(unittest.TestCase):
    def test_forbidden_legacy_strings_are_absent(self) -> None:
        problems = []
        for path in doc_files():
            for number, line in enumerate(read(path).splitlines(), 1):
                for needle, reason in FORBIDDEN:
                    if needle in line:
                        problems.append(f"{rel(path)}:{number}: «{needle}» — {reason}")
                if re.search(r"[\w-]+\.md:\d+", line):
                    problems.append(f"{rel(path)}:{number}: ссылка на номер строки .md — ссылайся на раздел")
                if re.search(r"final/(?!vkr\.docx)[\w.-]+\.docx", line):
                    problems.append(f"{rel(path)}:{number}: второй DOCX в final/ — промежуточные файлы в exports/")
        self.assertEqual([], problems)

    def test_scripts_reference_sections_not_line_numbers(self) -> None:
        problems = []
        for path in sorted(SCRIPTS.glob("*.py")):
            for number, line in enumerate(read(path).splitlines(), 1):
                if re.search(r"[\w-]+\.md:\d+", line):
                    problems.append(f"{rel(path)}:{number}: ссылка на номер строки .md — ссылайся на раздел")
        self.assertEqual([], problems)

    def test_preserve_ids_is_not_used_for_finalization(self) -> None:
        problems = []
        for path in doc_files():
            for command in iter_commands(path):
                if "--preserve-ids" in command.flags():
                    problems.append(command.label())
        for relative in WP6_DOCS:
            if "--preserve-ids" in read(ROOT / relative):
                problems.append(f"{relative}: упоминание --preserve-ids")
        self.assertEqual([], problems)

    def test_chapter_samples_use_roman_numerals(self) -> None:
        problems = []
        for relative in WP6_DOCS:
            for number, line in enumerate(read(ROOT / relative).splitlines(), 1):
                if re.search(r"Глава\s+\d+\.", line) or re.search(r"Выводы по [Гг]лаве\s+\d", line):
                    problems.append(f"{relative}:{number}: {line.strip()[:100]}")
        self.assertEqual([], problems)

    def test_audit_and_readiness_rules_are_documented(self) -> None:
        checklist = read(SKILL / "references" / "submission-checklist.md")
        for role in BASE_ROLES:
            self.assertIn(role, checklist, role)
        for needle in ("final/vkr.docx", "READY_TO_SUBMIT", "blind_regression", "evidence/defense/", "--in-place",
                       "каждый участник", "2 страницы"):
            self.assertIn(needle, checklist, needle)
        portable = read(SKILL / "PORTABLE-PROMPT.md")
        for needle in ("degraded_independence", "continuous-audit.md", "READY_FOR_SUPERVISOR_REVIEW", "vkr_audit.py"):
            self.assertIn(needle, portable, needle)
        express = read(SKILL / "references" / "writing-workflow-express.md")
        self.assertIn("balanced", express)
        self.assertIn("READY_FOR_SUPERVISOR_REVIEW", express)
        self.assertIn("[ПРОВЕРИТЬ", express)

    def test_web_file_list_contains_first_read_and_new_references(self) -> None:
        skill = read(SKILL / "SKILL.md")
        first = skill.split("**Прочитай в начале работы с проектом:**", 1)[1].split("**По задаче:**", 1)[0]
        required = set(re.findall(r"references/[\w-]+\.md", first))
        self.assertTrue(required, "в SKILL.md нет списка «Прочитай в начале работы»")
        required |= {f"references/{name}" for name in (
            "intake-interview.md", "project-initialization.md", "project-memory.md", "drafts-format.md",
            "docx-input-schema.md", "gost-citations.md", "source-verification.md", "analyzers-cli.md",
            "continuous-audit.md", "multi-agent-audit.md", "quality-control-loop.md", "state-file-pattern.md",
        )}
        required |= {"data/ai_cliches.json", "SKILL.md"}
        portable = read(SKILL / "PORTABLE-PROMPT.md")
        web = portable.split("## Файлы для веб-систем", 1)[1]
        self.assertEqual(sorted(name for name in required if name not in web), [])
        for relative in ("vkr-mpgu/INSTALL-RU.md", "vkr-mpgu/references/platform-integration.md", "README.md"):
            self.assertIn("Файлы для веб-систем", read(ROOT / relative), relative)


class FixBDocsTest(unittest.TestCase):
    """V2 Fix-B: установка переносит старую папку одной операцией, проверки версии, docx_integrity, профиль."""

    def test_install_block_moves_old_folder_atomically(self) -> None:  # B3
        blocks = []
        for relative in ("README.md", "vkr-mpgu/INSTALL-RU.md"):
            found = re.findall(r"```powershell\n(.*?)```", read(ROOT / relative), flags=re.S)
            install = [block for block in found if "Expand-Archive" in block]
            self.assertEqual(1, len(install), relative)
            blocks.append(install[0])
            self.assertIn("[IO.Directory]::Move($target, $backup)", install[0], relative)
            self.assertNotIn("Move-Item", install[0], relative)
            move = install[0].index("[IO.Directory]::Move")
            self.assertLess(install[0].index("throw", move), install[0].index("Expand-Archive"), relative)
        self.assertEqual(blocks[0], blocks[1])

    def test_bash_install_blocks_refuse_to_unpack_over_old_folder(self) -> None:  # финальная приёмка, находка 4
        # `unzip -n` молча пропускает существующие файлы, поэтому перед распаковкой блок обязан
        # убедиться, что старой папки нет, и остановиться с тем же сообщением, что и PowerShell.
        first_blocks = []
        for relative in ("README.md", "vkr-mpgu/INSTALL-RU.md"):
            blocks = [block for block in re.findall(r"```bash\n(.*?)```", read(ROOT / relative), flags=re.S) if "unzip -q -n" in block]
            self.assertTrue(blocks, relative)
            for block in blocks:
                guard = block.find("[ ! -e ")
                self.assertGreaterEqual(guard, 0, (relative, block))
                self.assertLess(guard, block.index("unzip -q -n"), relative)
                self.assertIn("Старая папка не перенесена, ничего не изменено", block, relative)
            first_blocks.append(blocks[0])
        self.assertEqual(first_blocks[0], first_blocks[1])

    def test_bash_version_check_for_codex(self) -> None:  # B6
        for relative in ("README.md", "vkr-mpgu/INSTALL-RU.md"):
            self.assertIn("grep -q '6.33-production' ~/.codex/skills/vkr-mpgu/SKILL.md && echo ok", read(ROOT / relative), relative)

    def test_quickstart_step4_runs_docx_integrity_with_exit_codes(self) -> None:  # B4
        step = read(SKILL / "references" / "quickstart.md").split("## Шаг 4.", 1)[1].split("\n## ", 1)[0]
        self.assertIn("python <SKILL_DIR>/scripts/docx_integrity.py <PROJECT_DIR>/final/vkr.docx", step)
        explained = step.split("`docx_integrity.py`", 1)[1].split("\n5.", 1)[0]
        for code in ("код 0", "код 1", "код 2"):
            self.assertIn(code, explained)

    def test_acceptance_v1_documentation_gaps_are_closed(self) -> None:  # B13, B9, B10, B11, B12
        cadence = read(SKILL / "references" / "continuous-audit.md")
        self.assertIn("`final` → `prefinal` → `draft` → `docx` → `chapter`", cadence)  # F13
        auto_safe = [line for line in cadence.splitlines() if line.startswith("|") and "`AUTO_SAFE`" in line]
        self.assertTrue(auto_safe and "без слепой регрессии" in auto_safe[0], auto_safe)  # F21
        self.assertIn("**`STY` в каждом втором подразделе (`strict`)** считает сам `plan`", cadence)  # F10
        quickstart = read(SKILL / "references" / "quickstart.md")
        step1 = quickstart.split("## Шаг 1.", 1)[1].split("\n## ", 1)[0]
        for needle in ("--preset intake", '"status": "external_changes_detected"', "record --event <EVENT_JSON> --all-changed"):
            self.assertIn(needle, step1)  # F20, F03
        self.assertIn("rechecks[].finding_status", quickstart)  # F09
        self.assertIn("handoff_stale", quickstart.split("## Шаг 9.", 1)[1])  # NC-1
        prompt = read(SKILL / "references" / "intake-interview.md").split("## Промпт-шаблон запуска интервью", 1)[1]
        self.assertNotIn("минут на 15-20", prompt)  # F01
        self.assertIn("45–60 минут", prompt)
        self.assertIn('"questions"', read(SKILL / "references" / "multi-agent-audit.md"))  # F08, F17
        for relative in ("references/writing-workflow.md", "references/methodology-mpgu.md"):
            self.assertIn("update_docx_fields.py", read(SKILL / relative), relative)  # F11
            self.assertIn("ANNOTATION_PAGES_MISMATCH", read(SKILL / relative), relative)

    def test_polish_stage_one_uses_profile_from_project(self) -> None:  # B5
        stage = read(SKILL / "references" / "draft-polish-workflow.md").split("### Этап 1", 1)[1].split("\n### ", 1)[0]
        self.assertIn("--profile <профиль из vkr-project.json>", stage)
        self.assertNotIn("--profile mpgu-09-project", stage)


class DeadlineScaleTest(unittest.TestCase):
    def test_canonical_table_in_skill(self) -> None:
        skill = read(SKILL / "SKILL.md")
        rows = [line for line in skill.splitlines() if line.startswith("|") and (DAY_RANGE.search(line) or DAY_PLUS.search(line))]
        pairs = []
        for row in rows:
            cells = [cell.strip() for cell in row.strip("|").split("|")]
            modes = MODE_TOKEN.findall(cells[1]) if len(cells) > 1 else []
            if modes:
                pairs.append((cells[0], modes[0]))
        self.assertEqual(list(CANONICAL_ROWS), pairs)

    def test_day_ranges_match_modes_in_all_docs(self) -> None:
        problems: List[str] = []
        for path in doc_files():
            problems.extend(mode_day_mismatches(path))
        self.assertEqual([], problems)

    def test_other_workflow_files_point_to_the_canonical_scale(self) -> None:
        for relative in ("references/writing-workflow-express.md", "references/intake-interview.md"):
            self.assertIn("SKILL.md", read(SKILL / relative), relative)


class FrontmatterTest(unittest.TestCase):
    def test_frontmatter_is_valid_yaml_with_short_description(self) -> None:
        block = frontmatter(read(SKILL / "SKILL.md"))
        data = parse_frontmatter_strict(block)
        try:
            import yaml  # type: ignore
        except ImportError:  # pragma: no cover - PyYAML необязателен для теста
            yaml = None
        if yaml is not None:
            self.assertEqual(data, yaml.safe_load(block))
        self.assertEqual("vkr-mpgu", data["name"])
        description = data["description"]
        self.assertIsInstance(description, str)
        self.assertLessEqual(len(description), 1024)
        self.assertGreater(len(description), 200)
        for trigger in ("ВКР", "диплом", "Антиплагиат", "список литературы", "титульный лист"):
            self.assertIn(trigger, description, trigger)


class UniversalWordingTest(unittest.TestCase):
    """Раунд «универсализация»: скилл читается как универсальный, а не как навык одной платформы."""

    def test_claude_is_used_only_as_a_platform_name(self) -> None:
        problems = []
        for path in doc_files():
            text = read(path)
            lines = text.splitlines()
            for match in CLAUDE_ANY.finditer(mask_allowed_claude(text)):
                number = text.count("\n", 0, match.start()) + 1
                context = lines[number - 1].strip()[:120] if number <= len(lines) else ""
                problems.append(
                    f"{rel(path)}:{number}: «Claude» не как название платформы — "
                    f"замени на «ИИ-агент»/«основной агент»/«ассистент»: {context}"
                )
        self.assertEqual([], problems)

    def test_masking_does_not_hide_a_real_mention(self) -> None:
        sample = "Claude Code и Claude.ai\nClaude помогает писать\n(Claude Project, Custom GPT)\n"
        masked = mask_allowed_claude(sample)
        found = [masked.count("\n", 0, m.start()) + 1 for m in CLAUDE_ANY.finditer(masked)]
        self.assertEqual([2], found)

    def test_entry_points_exist_and_point_to_the_skill(self) -> None:
        for relative in ENTRY_POINTS:
            path = ROOT / relative
            self.assertTrue(path.is_file(), f"нет файла точки входа {relative}")
            text = read(path)
            self.assertIn("vkr-mpgu/SKILL.md", text, relative)
            self.assertIn(VERSION, text, relative)
            self.assertIn("AI_AGENT_BOOTSTRAP", text, relative)
            self.assertNotIn("AI_AGENT_BOOTSTRAP_END", text,
                             f"{relative}: копия блока вместо ссылки на README")
        readme = read(ROOT / "README.md")
        self.assertIn("AI_AGENT_BOOTSTRAP_END", readme, "README — единственный источник блока")
        self.assertIn("## Использование в других ИИ-системах", readme)
        for platform in ("Codex", "Cursor", "Windsurf", "GitHub Copilot", "Gemini CLI", "PORTABLE-PROMPT.md"):
            self.assertIn(platform, readme, platform)
        self.assertIn("degraded_independence", readme)

    def test_entry_point_links_are_relative_and_exist(self) -> None:
        problems = []
        for relative in ENTRY_POINTS:
            path = ROOT / relative
            for number, line in enumerate(read(path).splitlines(), 1):
                for target in LINK.findall(line):
                    if target.startswith(("http://", "https://", "#", "mailto:")):
                        continue
                    if target.startswith("/") or ":" in target.split("#", 1)[0]:
                        problems.append(f"{relative}:{number}: ссылка не относительная: {target}")
                        continue
                    if not (path.parent / target.split("#", 1)[0]).resolve().exists():
                        problems.append(f"{relative}:{number}: ссылка {target} ведёт в пустоту")
        self.assertEqual([], problems)

    def test_cursor_rule_has_frontmatter_with_description(self) -> None:
        text = read(ROOT / ".cursor" / "rules" / "vkr-mpgu.mdc")
        self.assertTrue(text.startswith("---\n"), ".mdc должен начинаться с frontmatter")
        block = text[4:text.index("\n---\n", 4)]
        keys = dict(re.findall(r"^([A-Za-z]+):\s*(.*)$", block, re.M))
        self.assertTrue(keys.get("description", "").strip(), "во frontmatter нет description")
        self.assertEqual("false", keys.get("alwaysApply", "").strip())
        self.assertTrue(keys.get("globs", "").strip(), "во frontmatter нет globs")

    def test_entry_points_are_outside_the_skill_folder(self) -> None:
        for relative in ENTRY_POINTS:
            self.assertFalse(relative.startswith("vkr-mpgu/"), relative)
            self.assertNotIn(relative, skill_files(), relative)


if __name__ == "__main__":
    unittest.main(verbosity=2)
