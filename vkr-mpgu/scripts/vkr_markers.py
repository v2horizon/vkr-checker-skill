#!/usr/bin/env python3
"""Единый список маркеров незавершённого текста для doctor, валидатора и сборки.

Модуль без внешних зависимостей. Маркер — это след черновика, который нельзя
сдавать: пометка о непроверенном источнике, незаполненное значение, заглушка
рисунка или оглавления. Код внутри fenced-блоков Markdown и листингов DOCX
маркером не считается: индексы вида ``arr[x]`` и шаблоны ``{{ name }}``
законны в программном коде.
"""

from __future__ import annotations

import re
from typing import List, Tuple

# (код маркера, регулярное выражение)
MARKER_PATTERNS: Tuple[Tuple[str, "re.Pattern[str]"], ...] = (
    ("CHECK_SOURCE", re.compile(r"\[ПРОВЕРИТЬ(?:[:\s][^\]\n]*)?\]", re.IGNORECASE)),
    ("NEEDS_CLARIFICATION", re.compile(r"\[ТРЕБУЕТ\s+УТОЧНЕНИЯ(?:[:\s][^\]\n]*)?\]", re.IGNORECASE)),
    ("FILL_IN", re.compile(r"\[(?:ВСТАВИТЬ|ЗАПОЛНИТЬ|ДОПОЛНИТЬ)(?:[:\s][^\]\n]*)?\]", re.IGNORECASE)),
    ("VALUE_PLACEHOLDER", re.compile(r"\[(?:N|X)\]")),
    ("VALUE_PLACEHOLDER", re.compile(r"\[список\]", re.IGNORECASE)),
    ("TEMPLATE_VARIABLE", re.compile(r"\{\{[^{}\n]{1,80}\}\}")),
    ("TITLE_PLACEHOLDER", re.compile(r"\[Название[^\]\n]*\]", re.IGNORECASE)),
    ("DRAFT_COMMENT", re.compile(r"<!--\s*(?:Черновик|Заполняется)", re.IGNORECASE)),
    ("TODO", re.compile(r"(?<!\w)(?:TODO|FIXME|TBD)(?!\w)")),
    ("TODO", re.compile(r"\[XXX\]")),
    ("FIGURE_PLACEHOLDER", re.compile(r"\[\s*Здесь\s+вставить[^\]\n]*\]", re.IGNORECASE)),
)

TOC_PLACEHOLDER = ("TOC_PLACEHOLDER", re.compile(r"Для обновления содержания", re.IGNORECASE))

_FENCE = re.compile(r"^[ \t]*(`{3,}|~{3,})")
_INLINE_CODE = re.compile(r"`[^`\n]+`")


def strip_markdown_code(text: str) -> str:
    """Удаляет fenced-блоки и inline-код, сохраняя число строк.

    Незакрытый fenced-блок считается кодом до конца текста.
    """
    lines = text.splitlines()
    result: List[str] = []
    fence: str | None = None
    for line in lines:
        match = _FENCE.match(line)
        if fence is None:
            if match:
                fence = match.group(1)[0] * 3
                result.append("")
                continue
            result.append(_INLINE_CODE.sub(" ", line))
        else:
            if match and match.group(1).startswith(fence):
                fence = None
            result.append("")
    return "\n".join(result)


def find_markers(text: str, *, docx_toc_placeholder: bool = False) -> List[str]:
    """Возвращает найденные маркеры в порядке шаблонов (каждое вхождение отдельно)."""
    if not text:
        return []
    found: List[str] = []
    patterns = list(MARKER_PATTERNS)
    if docx_toc_placeholder:
        patterns.append(TOC_PLACEHOLDER)
    for _code, pattern in patterns:
        found.extend(match.group(0) for match in pattern.finditer(text))
    return found


def find_marker_details(text: str, *, docx_toc_placeholder: bool = False) -> List[Tuple[str, str]]:
    """Как find_markers, но возвращает пары (код, фрагмент)."""
    if not text:
        return []
    found: List[Tuple[str, str]] = []
    patterns = list(MARKER_PATTERNS)
    if docx_toc_placeholder:
        patterns.append(TOC_PLACEHOLDER)
    for code, pattern in patterns:
        found.extend((code, match.group(0)) for match in pattern.finditer(text))
    return found


def is_soft_marker(marker: str) -> bool:
    """Мягкий маркер допустим в версии для научного руководителя (prefinal)."""
    return marker.strip().upper().startswith("[ПРОВЕРИТЬ")


if __name__ == "__main__":  # pragma: no cover - ручная проверка
    sample = (
        "Текст [ПРОВЕРИТЬ] и [ТРЕБУЕТ УТОЧНЕНИЯ: выборка] и [N]%.\n"
        "```python\nitems[x] = {{ value }}  # TODO\n```\n"
        "Приложение ToDo, XXX Олимпиада, arr[n]. TBD\n"
    )
    print(find_markers(strip_markdown_code(sample)))
