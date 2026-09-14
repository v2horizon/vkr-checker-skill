#!/usr/bin/env python3
"""
content_ownership_check.py — Генератор вопросов комиссии по тексту ВКР (vkr-mpgu 6.33).

Главная проблема LLM-ассистированных работ: студент не владеет собственным
текстом. На защите комиссия задаёт 3-5 вопросов по конкретным абзацам —
и если студент не может обосновать написанное, оценка снижена или защита
отложена.

Этот скрипт решает именно эту проблему:
1. Извлекает абзацы из ВКР
2. Для каждого абзаца формулирует 2-3 типичных вопроса комиссии
3. Студент проходит по ним до защиты и тренирует ответы

Что берётся в вопросы: абзацы не короче 80 символов из введения, глав и
заключения. Разделы определяются так же, как в ai_detection_heuristic.py (стили
Heading 1/2, «Заголовок 1/2», отдельные абзацы «Введение», «Глава I…»,
«Заключение», варианты заголовка списка литературы, «Приложение N»). Титул,
аннотация, содержание, список литературы, приложения, последний лист, код и
подписи в вопросы не попадают.

Использование:
    python content_ownership_check.py vkr.docx                  # выборка до 25 абзацев (seed 42)
    python content_ownership_check.py vkr.docx --all            # полный прогон
    python content_ownership_check.py vkr.docx --section chapter3
    python content_ownership_check.py vkr.docx --sample 20 --seed 7
    python content_ownership_check.py vkr.docx --all --json -o report.json

Выборка и вопросы воспроизводимы: одинаковые файл, --sample и --seed дают
одинаковый результат. При выборке отчёт явно помечен «ВЫБОРКА» (JSON: sampling).

Коды выхода: 0 — вопросы сформированы; 1 — раздел --section не найден или нет ни
одного абзаца не короче 80 символов (INSUFFICIENT_DATA); 2 — нет файла,
повреждённый DOCX, нет python-docx, неверные аргументы, ошибка записи -o.

Цель — не формальная оценка, а подготовка к защите через самопроверку.
"""

from __future__ import annotations

import argparse
import random
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

sys.dont_write_bytecode = True
SCRIPTS_DIR = Path(__file__).resolve().parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import vkr_common as common  # noqa: E402
import ai_detection_heuristic as docx_structure  # noqa: E402  (общий разбор структуры DOCX и allowlist)

TOOL_VERSION = "6.33"
MIN_PARAGRAPH_CHARS = 80
PLACEHOLDER_RE = re.compile(r"\{[^{}]*\}")


# ========== ШАБЛОНЫ ВОПРОСОВ ==========

# Паттерны, которые срабатывают на определённые конструкции в тексте.
# Каждый паттерн → набор вопросов. Буквы ё/е в маркерах и тексте не различаются.
QUESTION_PATTERNS = [
    {
        # Конкретное решение / выбор технологии — только в главах 2-3,
        # где реально обсуждаются выборы. В аннотации/введении/заключении
        # слова «использован», «реализован» — обзорные, не про выбор.
        "markers": [
            r"\bвыбран(а|о|ы)?\b",
            r"\bреализован(а|о|ы)?\b",
            r"\bиспользован(а|о|ы)?\b",
            r"\bприменён(а|о|ы)?\b",
            r"\bисходя из\b",
        ],
        "questions": [
            "Почему вы выбрали именно это решение, а не альтернативу?",
            "Какие альтернативы рассматривались? Какие оставили и почему отвергли?",
            "Что было бы, если бы использовали другой подход?",
        ],
        "category": "technology-choice",
        "applicable_sections": {"chapter2", "chapter3"},
    },
    {
        # Цифровые утверждения — в аннотации/содержании числа это объём/номера,
        # в заключении — итоговые метрики. Вопросы про выборку/погрешность
        # уместны только в главах 2-3, где приводятся результаты.
        "markers": [
            r"\b\d{1,3}\s?%",
            r"\bоколо \d+\b",
            r"\bболее \d+\b",
            r"\bменее \d+\b",
            r"\bсоставляет \d+\b",
            r"\bсоставил[аи]? \d+\b",
        ],
        "questions": [
            "Откуда взята эта цифра? Какой источник?",
            "Как методологически измерялось? Какая выборка?",
            "Какая погрешность в этих данных?",
        ],
        "category": "quantitative-claim",
        "applicable_sections": {"chapter1", "chapter2", "chapter3"},
    },
    {
        # Метод / методология
        "markers": [
            r"\b(выбран|применён|использ\w+|реализован|предложен)\w*\s+\w*\s*(метод|подход|алгоритм|методолог)\w*\b",
            r"\b(метод|подход|алгоритм|методолог)\w*\s+\w*\s*(был|являлся|лёг в основу|применялся|использовался)\b",
            r"\bв качестве\s+\w*\s*(метод|подход|алгоритм|методолог)\w*\b",
        ],
        "questions": [
            "Расшифруйте, как этот метод работает пошагово",
            "Какие ограничения у этого метода?",
            "В каких случаях этот метод неприменим?",
        ],
        "category": "methodology",
        "applicable_sections": {"chapter1", "chapter2", "chapter3"},
    },
    {
        # Ссылки на авторов / источники — всегда есть в теории и анализе,
        # иногда во введении (актуальность). В заключении ссылок БЫТЬ НЕ ДОЛЖНО
        # по методичке, но если попали — всё равно проверить понимание.
        "markers": [
            r"\[[\d]+(?:,\s*с\.\s*\d+)?\]",
            r"\bпо мнению\b",
            r"\bсогласно\b",
            r"\bуказывает\b",
            r"\bотмечает\b",
            r"\bсчитает\b",
        ],
        "questions": [
            "Перескажите своими словами, что именно утверждает этот автор",
            "Согласны ли вы с этой позицией? Если нет — почему?",
            "Какой контраргумент существует в научной литературе?",
        ],
        "category": "source-reference",
        "applicable_sections": {"introduction", "chapter1", "chapter2", "chapter3"},
    },
    {
        # Результаты / выводы
        "markers": [
            r"\bрезультат\w*\b",
            r"\bполучен\w*\b",
            r"\bвыявлен\w*\b",
            r"\bустановлен\w*\b",
            r"\bпоказан[ао]?\b",
        ],
        "questions": [
            "Как проверялась достоверность этого результата?",
            "Что могло исказить результат? Какие факторы учитывались?",
            "Воспроизводим ли этот результат? Что нужно, чтобы повторить?",
        ],
        "category": "results",
        "applicable_sections": {"chapter2", "chapter3", "conclusion"},
    },
    {
        # Определения / ключевые понятия
        "markers": [
            r"\bпредставля\w+ собой\b",
            r"\bпод\s+\w+\s+понимается\b",
            r"\bпод\s+\w+\s+подразумевается\b",
            r"\b(под|в\s+\w+\s+смысле)\s+\w+\s+мы\s+понимаем\b",
            r"\bопределяется как\b",
            r"\bпо определению\b",
            r"\bданное понятие\b",
            r"\bданный термин\b",
        ],
        "questions": [
            "Дайте альтернативное определение — как ещё можно сформулировать?",
            "Чем это понятие отличается от смежных?",
            "Есть ли в литературе другое определение? Какое?",
        ],
        "category": "definition",
        "applicable_sections": {"chapter1", "chapter2"},
    },
    {
        # Критические замечания / позиция автора
        "markers": [
            r"\bвызывает вопросы\b",
            r"\bпредставляется\b",
            r"\bна наш взгляд\b",
            r"\bпо нашему мнению\b",
            r"\bследует отметить\b",
            r"\bтребует уточнения\b",
        ],
        "questions": [
            "Обоснуйте вашу позицию конкретнее. На какие данные она опирается?",
            "Какие контраргументы к вашей точке зрения существуют?",
            "Что бы убедило вас изменить эту позицию?",
        ],
        "category": "author-position",
        "applicable_sections": {"introduction", "chapter1", "chapter2", "chapter3", "conclusion"},
    },
    {
        # Пилотирование / эксперимент — ТОЛЬКО Глава 3. В аннотации/заключении
        # слово «выборка» упоминается обзорно, там вопрос про обоснование размера
        # бессмысленен.
        "markers": [
            r"\bпилот\w*\b",
            r"\bэксперимент\w*\b",
            r"\bапробац\w*\b",
            r"\bтестиро\w*\b",
            r"\bвыборк\w*\b",
        ],
        "questions": [
            "Почему именно такая выборка? Чем обоснован её размер?",
            "Что пошло не так во время пилотирования? Что вы изменили?",
            "Репрезентативны ли участники для генеральной совокупности?",
        ],
        "category": "pilot",
        "applicable_sections": {"chapter3"},
    },
    {
        # Архитектурные / системные решения
        "markers": [
            r"\bархитектур\w*\b",
            r"\bкомпонент\w*\b",
            r"\bмодул\w*\b",
            r"\bсервис\w*\b",
            r"\bбаза данных\b",
            r"\bAPI\b",
            r"\bсхем\w*\b",
        ],
        "questions": [
            "Нарисуйте схему этого на доске",
            "Как данные перемещаются между компонентами?",
            "Какое самое уязвимое место этой архитектуры?",
        ],
        "category": "architecture",
        # v6.25: архитектурные вопросы не имеют смысла во введении/аннотации/
        # заключении — там архитектура не описывается подробно.
        "applicable_sections": {"chapter2", "chapter3"},
    },
]

# Общие вопросы на весь текст — универсальные
GENERIC_QUESTIONS = [
    "Перескажите суть этого абзаца одним предложением",
    "Какая главная мысль? Зачем этот абзац в работе?",
    "Что конкретно нового вы добавили к существующей литературе в этом абзаце?",
    "Какой вопрос этот абзац закрывает и какой ставит?",
]

# Служебные строки, которые не бывают содержательными абзацами
SERVICE_MARKERS = [
    "⚠️",
    "выполнена мной совершенно самостоятельно",
    "живой ручкой",
    "Отсканированная подпись",
    "[Для обновления содержания",
    "___________________",
    "Ф.И.О.",
    "Проверка на объем заимствований",
    "Проверка на объём заимствований",
]
# Запись списка литературы, если раздел не распознан: «NN. Фамилия, И.О. …»,
# «NN. Заглавие [Текст] …», «NN. … // Журнал».
BIBLIOGRAPHY_ENTRY_RE = re.compile(
    r"^\d{1,3}[.)]\s+(?:[А-ЯЁA-Z][а-яёa-z'’\-]+,\s|.{0,250}?(?:\[(?:Текст|Электронный ресурс|Electronic resource)\]|\s//\s))"
)


def _norm(text: str) -> str:
    return str(text or "").replace("ё", "е").replace("Ё", "Е")


_COMPILED: Dict[str, "re.Pattern[str]"] = {}


def _marker_re(marker: str) -> "re.Pattern[str]":
    compiled = _COMPILED.get(marker)
    if compiled is None:
        compiled = re.compile(_norm(marker), re.IGNORECASE)
        _COMPILED[marker] = compiled
    return compiled


def is_service_paragraph(text: str) -> bool:
    return any(marker in text for marker in SERVICE_MARKERS)


def split_paragraphs(doc) -> list:
    """Значимые абзацы работы (≥ 80 символов) без служебных частей: титула,
    аннотации, содержания, списка литературы, приложений, клятвы, кода и подписей."""
    return [text for _, text in collect_paragraphs(docx_structure.read_structure(doc))]


def collect_paragraphs(structure, section: Optional[str] = None) -> List[Tuple[Optional[str], str]]:
    """(раздел, текст) для вопросов. section=None — введение, главы и заключение;
    для нераспознанной структуры — весь текст без служебных областей."""
    if section is not None:
        name = section.strip().casefold()
        texts = [(name, text) for text in structure.section_texts(name)]
    elif structure.has_body():
        texts = list(structure.body_items())
    else:
        texts = [
            (None, text) for text in structure.fallback_texts(include_tables=False)
            if not text.startswith(("Глава ", "Приложение ", "Таблица ", "Рисунок "))
            and not BIBLIOGRAPHY_ENTRY_RE.match(text)
        ]
    return [
        (name, text) for name, text in texts
        if len(text) >= MIN_PARAGRAPH_CHARS and not is_service_paragraph(text)
    ]


def detect_applicable_patterns(paragraph: str, section: str = None, allowlist: Optional[Sequence[str]] = None) -> list:
    """Найти, какие паттерны срабатывают на абзаце.

    Если указан section ('introduction' | 'chapter1-3' | 'conclusion'),
    фильтрует паттерны по полю applicable_sections — например, вопросы про
    пилотирование не задаются для аннотации/введения, вопросы про выбор
    технологии не задаются для заключения.
    Если секция не указана (обзор всей работы) — применяются все паттерны.
    Фраза, совпавшая с персональным cliche_allowlist (например, «следует
    отметить»), маркером не считается — так же, как в ai_detection_heuristic.py.
    """
    normalized = _norm(paragraph)
    allowed = {docx_structure.normalize_phrase(item) for item in (allowlist or [])}
    applicable = []
    for pattern in QUESTION_PATTERNS:
        # Фильтрация по секции
        if section is not None:
            applicable_sections = pattern.get("applicable_sections")
            if applicable_sections is not None and section not in applicable_sections:
                continue
        matched = False
        for marker in pattern["markers"]:
            for match in _marker_re(marker).finditer(normalized):
                if allowed and docx_structure.normalize_phrase(match.group(0)) in allowed:
                    continue
                matched = True
                break
            if matched:
                break
        if matched:
            applicable.append(pattern)
    return applicable


def generate_questions_for_paragraph(paragraph: str, max_questions: int = 3,
                                     section: str = None, rng: Optional[random.Random] = None,
                                     allowlist: Optional[Sequence[str]] = None) -> list:
    """Сгенерировать 2-3 вопроса для абзаца с учётом секции."""
    chooser = rng or random
    applicable = detect_applicable_patterns(paragraph, section=section, allowlist=allowlist)

    questions = []
    categories_used = set()

    # Специфические вопросы по паттернам
    for pattern in applicable:
        if pattern["category"] in categories_used:
            continue
        q = PLACEHOLDER_RE.sub("это решение", chooser.choice(pattern["questions"]))
        questions.append({
            "category": pattern["category"],
            "question": q,
        })
        categories_used.add(pattern["category"])
        if len(questions) >= max_questions:
            break

    # Добавить один генерик-вопрос если не хватает
    if len(questions) < max_questions:
        q = PLACEHOLDER_RE.sub("этот абзац", chooser.choice(GENERIC_QUESTIONS))
        questions.append({
            "category": "generic",
            "question": q,
        })

    return questions[:max_questions]


def extract_section(doc, section_name: str) -> list:
    """Абзацы раздела (≥ 80 символов) по заголовкам, без строк оглавления."""
    structure = docx_structure.read_structure(doc)
    return [text for _, text in collect_paragraphs(structure, section=section_name)]


def generate_report(paragraphs: list, verbose: bool = True, section: str = None,
                    available_count=None, rng: Optional[random.Random] = None,
                    allowlist: Optional[Sequence[str]] = None,
                    sections: Optional[Sequence[Optional[str]]] = None,
                    positions: Optional[Sequence[int]] = None) -> dict:
    """Полный отчёт — для каждого абзаца 2-3 вопроса.

    section: если указан — паттерны фильтруются по applicable_sections,
    и генерируются только релевантные вопросы для данной секции.
    positions: номера абзацев в полном списке (при выборке), sections — их разделы.
    """
    items = []
    for i, p in enumerate(paragraphs):
        questions = generate_questions_for_paragraph(p, max_questions=3, section=section, rng=rng, allowlist=allowlist)
        item = {
            "paragraph_id": positions[i] if positions is not None else i + 1,
            "preview": p[:150] + ("..." if len(p) > 150 else ""),
            "full_text": p,
            "questions": questions,
        }
        if sections is not None:
            item["section"] = sections[i]
        items.append(item)

    # Агрегированная статистика по категориям
    category_counts: Dict[str, int] = {}
    for item in items:
        for q in item["questions"]:
            category_counts[q["category"]] = category_counts.get(q["category"], 0) + 1

    return {
        "available_paragraphs": available_count if available_count is not None else len(paragraphs),
        "sampling": "all" if available_count in (None, len(paragraphs)) else "sample",
        "total_paragraphs": len(paragraphs),
        "total_questions": sum(len(i["questions"]) for i in items),
        "category_distribution": category_counts,
        "items": items,
    }


def print_report(report: dict, show_full: bool = False):
    """Текстовый отчёт."""
    print("\n" + "=" * 72)
    print("  ПРОВЕРКА ВЛАДЕНИЯ МАТЕРИАЛОМ (content ownership check)")
    print("  Симуляция вопросов, которые может задать комиссия")
    print("=" * 72)
    sampled = report.get("sampling") == "sample"
    if sampled:
        print(f"\nВЫБОРКА: {report['total_paragraphs']} из {report.get('available_paragraphs')} абзацев "
              f"(seed {report.get('seed')}). Полный прогон: --all")
    else:
        print(f"\nПолный прогон: {report['total_paragraphs']} абзацев из {report.get('available_paragraphs', report['total_paragraphs'])}")
    if report.get("structure") == "unrecognized":
        print("ВНИМАНИЕ: разделы ВКР не распознаны — абзацы взяты из всего текста без служебных частей")
    if (report.get("cliche_allowlist") or {}).get("phrases"):
        info = report["cliche_allowlist"]
        print(f"Персональный cliche_allowlist: {len(info['phrases'])} фраз ({info['state_file']})")
    print(f"Сгенерировано вопросов: {report['total_questions']}")

    if report.get("category_distribution"):
        print("\nРаспределение по типам вопросов:")
        for cat, count in sorted(report["category_distribution"].items(),
                                 key=lambda x: -x[1]):
            print(f"  • {cat}: {count}")

    print("\n" + "-" * 72)
    print("ТРЕНИРОВОЧНЫЕ ВОПРОСЫ (ответь на каждый за 60 секунд устно):")
    print("-" * 72)

    for item in report["items"]:
        where = f", {item['section']}" if item.get("section") else ""
        print(f"\n[Абзац #{item['paragraph_id']}{where}]")
        if show_full:
            print(f"\n{item['full_text']}\n")
        else:
            print(f"  {item['preview']}\n")
        for j, q in enumerate(item["questions"], 1):
            cat_label = f"[{q['category']}]"
            print(f"  {j}. {q['question']} {cat_label}")

    print("\n" + "=" * 72)
    print("  РЕКОМЕНДАЦИИ")
    print("=" * 72)
    if sampled:
        first = ("1. Это выборка абзацев. Перед защитой прогони всю работу: запусти с --all\n"
                 "   и пройди вопросы по КАЖДОМУ абзацу.")
    else:
        first = "1. Прогони через эти вопросы КАЖДЫЙ абзац своей работы."
    print(f"""
{first}
2. Если на какой-то вопрос не можешь ответить за 60 секунд — этот абзац
   надо либо переписать своими словами, либо изучить глубже.
3. Особое внимание:
   - Абзацы с цифрами/процентами — комиссия их любит проверять
   - Архитектурные решения — «нарисуйте на доске» классический вопрос
   - Ссылки на источники — «перескажите своими словами»
4. Пройди по работе минимум два раза:
   - За 2 недели до защиты (выявить слабые места)
   - За 2 дня до защиты (финальная репетиция)
5. Запиши себя на телефон, отвечая на вопросы. Послушай — если звучит
   как заученное, комиссия это тоже услышит.
""")


def main(argv: Optional[Sequence[str]] = None) -> int:
    common.reconfigure_stdio()
    parser = argparse.ArgumentParser(
        description="Генератор вопросов комиссии по тексту ВКР",
        epilog="Коды выхода: 0 — вопросы сформированы; 1 — раздел не найден или нет абзацев ≥ 80 символов "
               "(INSUFFICIENT_DATA); 2 — нет файла, повреждённый DOCX, нет python-docx, ошибка записи -o.",
    )
    parser.add_argument("docx_path", help="Путь к .docx ВКР")
    parser.add_argument("--section", help="Раздел: introduction | conclusion | chapterN (chapter1, chapter2, …)")
    parser.add_argument(
        "--sample", type=int, default=25,
        help="Случайная выборка N абзацев (по умолчанию 25; 0 — без выборки). Выборка помечается в отчёте.",
    )
    parser.add_argument(
        "--all", action="store_true",
        help="Прогнать ВСЕ абзацы без выборки (для финальной подготовки к защите).",
    )
    parser.add_argument("--full", action="store_true", help="Показывать полный текст абзацев")
    parser.add_argument("--json", action="store_true", help="Только JSON в stdout")
    parser.add_argument("-o", "--output", type=Path, help="Записать JSON-отчёт в файл (UTF-8, каталоги создаются)")
    parser.add_argument("--seed", type=int, default=42, help="Seed выборки и выбора вопросов (воспроизводимость)")
    args = parser.parse_args(argv)

    path = Path(args.docx_path)
    meta: Dict[str, Any] = {"tool": "content_ownership_check", "version": TOOL_VERSION, "docx": str(path)}

    def finish(payload: dict, code: int, text_printer=None) -> int:
        for key, value in meta.items():
            payload.setdefault(key, value)
        payload["exit_code"] = code
        if args.output:
            if docx_structure.is_inside_skill_dir(args.output):
                print(f"ОШИБКА: {args.output} внутри каталога скилла: укажи файл в каталоге проекта", file=sys.stderr)
                return 2
            try:
                common.atomic_write_text(args.output, common.dump_json(payload))
            except OSError as error:
                print(f"ОШИБКА: не удалось записать {args.output}: {error}", file=sys.stderr)
                return 2
        if args.json:
            sys.stdout.write(common.dump_json(payload))
        elif text_printer is not None:
            text_printer()
        return code

    def fail(code: int, error_code: str, message: str) -> int:
        payload = {"status": "error", "error": {"code": error_code, "message": message}}
        return finish(payload, code, lambda: print(f"ОШИБКА: {message}", file=sys.stderr))

    section = args.section.strip().casefold() if args.section else None
    if section is not None and not docx_structure.SECTION_NAME_RE.match(section):
        return fail(2, "SECTION_INVALID", f"неизвестный раздел «{args.section}»: introduction, conclusion, chapter1, chapter2, …")
    if args.sample < 0:
        return fail(2, "SAMPLE_INVALID", "--sample должен быть неотрицательным")
    if not path.is_file():
        return fail(2, "FILE_NOT_FOUND", f"файл не найден: {path}")
    try:
        structure = docx_structure.read_structure(path)
    except docx_structure.DependencyError as error:
        return fail(2, "DEPENDENCY_MISSING", str(error))
    except Exception as error:  # повреждённый пакет, не DOCX
        return fail(2, "DOCX_INVALID", f"не удалось открыть DOCX {path}: {type(error).__name__}: {error}")

    phrases, state = docx_structure.load_cliche_allowlist(path)
    summary = structure.summary()
    meta.update({
        "section": section,
        "structure": summary["structure"],
        "sections_found": summary["sections_found"],
        "cliche_allowlist": docx_structure.allowlist_info(phrases, state),
        "seed": args.seed,
    })

    if section is not None and section not in structure.section_names():
        found = ", ".join(structure.section_names()) or "нет"
        message = f"раздел {section} не найден; найдены: {found}"
        return finish({"status": "section_not_found", "error": {"code": "SECTION_NOT_FOUND", "message": message}},
                      1, lambda: print(f"ОШИБКА: {message}", file=sys.stderr))

    items = collect_paragraphs(structure, section=section)
    available = len(items)
    if available == 0:
        message = f"нет абзацев не короче {MIN_PARAGRAPH_CHARS} символов для вопросов"
        payload = {"status": "insufficient_data", "risk_level": "INSUFFICIENT_DATA",
                   "error": {"code": "INSUFFICIENT_DATA", "message": message},
                   "available_paragraphs": 0, "total_paragraphs": 0, "total_questions": 0, "items": []}
        return finish(payload, 1, lambda: print(f"INSUFFICIENT_DATA: {message}. Код выхода 1.", file=sys.stderr))

    rng = random.Random(args.seed)
    if not args.all and args.sample and args.sample < available:
        chosen = sorted(rng.sample(range(available), args.sample))
    else:
        chosen = list(range(available))
    selected = [items[index] for index in chosen]
    report = generate_report(
        [text for _, text in selected],
        section=section,
        available_count=available,
        rng=rng,
        allowlist=phrases,
        sections=[name for name, _ in selected],
        positions=[index + 1 for index in chosen],
    )
    report["status"] = "ok"
    report["sample_size"] = len(selected)
    if report["sampling"] == "sample":
        report["sampling_note"] = (f"ВЫБОРКА: {len(selected)} из {available} абзацев (seed {args.seed}); "
                                   "полный прогон — --all")
    else:
        report["sampling_note"] = f"полный прогон: {available} абзацев"

    def show() -> None:
        merged = dict(meta)
        merged.update(report)
        if section:
            print(f"\n[Раздел: {section}]")
        print_report(merged, show_full=args.full)

    return finish(report, 0, show)


if __name__ == "__main__":
    raise SystemExit(main())
