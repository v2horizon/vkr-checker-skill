#!/usr/bin/env python3
"""Создаёт рабочую структуру проекта ВКР.

Скрипт не перезаписывает существующие файлы: повторный запуск досоздаёт
недостающее. Проект нельзя создать внутри установленного скилла.

Коды возврата: 0 — создано или dry-run; 2 — ошибка конфигурации, конфликт с
существующим проектом или ошибка ввода-вывода.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import shutil
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

sys.dont_write_bytecode = True
_SCRIPTS = str(Path(__file__).resolve().parent)
if _SCRIPTS not in sys.path:
    sys.path.insert(0, _SCRIPTS)

import vkr_common as common  # noqa: E402
import vkr_memory as memory  # noqa: E402

SKILL_VERSION = common.SKILL_VERSION
PROTOCOL_VERSION = common.PROTOCOL_VERSION
AUDIT_PROTOCOL_VERSION = PROTOCOL_VERSION
SKILL_ROOT = common.SKILL_DIR

TITLE_PAGE_KEYS = common.TITLE_PAGE_KEYS
TOP_LEVEL_KEYS = ("vkr_type", "profile", "mode", "audit_intensity", "topic", "deadline", "collective", "members", "ai_rules", "title_page")
LEGACY_TITLE_KEYS = dict(common.LEGACY_FLAT_TITLE_KEYS)
DIRECTORIES = (
    "drafts",
    "sources/materials",
    "evidence/methodology",
    "evidence/product",
    "evidence/pilot",
    "evidence/figures",
    "evidence/defense",
    "evidence/approvals",
    "audit/reports",
    "audit/briefs",
    "audit/snapshots",
    "memory",
    "logs",
    "final",
    "exports",
    "backups/checkpoints",
)


def clean_text(value: Any) -> str:
    return " ".join(str(value if value is not None else "").replace("\x00", "").split())


def default_title_page() -> Dict[str, str]:
    return common.default_title_page()


def load_config(path: Optional[Path]) -> Dict[str, Any]:
    """Читает intake-конфигурацию. Возвращает значения и множество явно заданных ключей."""
    raw: Dict[str, Any] = {}
    if path is not None:
        loaded = common.read_json(Path(path), label=str(path))
        if not isinstance(loaded, dict):
            raise ValueError("Конфигурация intake должна быть JSON-объектом")
        raw = dict(loaded)
    if "validation_profile" in raw:
        if "profile" in raw and raw["profile"] != raw["validation_profile"]:
            raise ValueError("profile и validation_profile противоречат друг другу")
        raw["profile"] = raw.pop("validation_profile")
    allowed = set(TOP_LEVEL_KEYS) | set(LEGACY_TITLE_KEYS)
    unknown = sorted(set(raw) - allowed)
    if unknown:
        raise ValueError("Неизвестные ключи конфигурации: " + ", ".join(unknown) + ". Допустимы: " + ", ".join(sorted(allowed)))
    config: Dict[str, Any] = {"title_page": {}}
    explicit: Set[str] = set()
    title_page = raw.get("title_page", {})
    if not isinstance(title_page, dict):
        raise ValueError("title_page должен быть объектом")
    bad_title = sorted(set(title_page) - set(TITLE_PAGE_KEYS))
    if bad_title:
        raise ValueError("Неизвестные ключи title_page: " + ", ".join(bad_title) + ". Допустимы: " + ", ".join(TITLE_PAGE_KEYS))
    for key, value in title_page.items():
        if not isinstance(value, (str, int)) or isinstance(value, bool):
            raise ValueError(f"title_page.{key} должен быть строкой")
        config["title_page"][key] = clean_text(value)
    for legacy, target in LEGACY_TITLE_KEYS.items():
        if legacy in raw:
            value = clean_text(raw[legacy])
            if target in config["title_page"] and config["title_page"][target] != value:
                raise ValueError(f"{legacy} противоречит title_page.{target}")
            config["title_page"][target] = value
    for key in TOP_LEVEL_KEYS:
        if key == "title_page" or key not in raw:
            continue
        config[key] = raw[key]
        explicit.add(key)
    config["_explicit"] = explicit
    return config


def apply_overrides(config: Dict[str, Any], args: argparse.Namespace) -> None:
    explicit: Set[str] = config.setdefault("_explicit", set())
    for key in ("topic", "profile", "mode", "audit_intensity", "deadline", "vkr_type"):
        value = getattr(args, key, None)
        if value is not None:
            config[key] = value
            explicit.add(key)
    if getattr(args, "collective", False):
        config["collective"] = True
        explicit.add("collective")
    for option, target in (("student_name", "author"), ("institution", "university"), ("program_code", "program_code")):
        value = getattr(args, option, None)
        if value is not None:
            config["title_page"][target] = clean_text(value)


def normalize_config(config: Dict[str, Any]) -> Dict[str, Any]:
    explicit: Set[str] = set(config.get("_explicit", set()))
    result: Dict[str, Any] = {}
    profile = clean_text(config.get("profile") or "generic")
    if profile not in common.PROFILES:
        raise ValueError("profile должен быть одним из: " + ", ".join(common.PROFILES))
    result["profile"] = profile
    required_type = common.PROFILE_TYPES.get(profile)
    if "vkr_type" in explicit:
        vkr_type = clean_text(config.get("vkr_type")).lower()
        if vkr_type not in common.VKR_TYPES:
            raise ValueError("vkr_type должен быть project или regular")
        if required_type and vkr_type != required_type:
            raise ValueError(f"Профиль {profile} требует vkr_type={required_type}, указан {vkr_type}")
    else:
        vkr_type = required_type or "project"
    result["vkr_type"] = vkr_type
    mode = clean_text(config.get("mode") or "standard")
    if mode not in common.MODES:
        raise ValueError("mode должен быть одним из: " + ", ".join(common.MODES))
    result["mode"] = mode
    if "audit_intensity" in explicit:
        intensity = clean_text(config.get("audit_intensity"))
        if intensity not in common.AUDIT_INTENSITIES:
            raise ValueError("audit_intensity должен быть balanced, strict или maximum")
    else:
        intensity = "balanced" if mode.startswith("express") else "strict"
    result["audit_intensity"] = intensity
    result["topic"] = clean_text(config.get("topic"))
    deadline = clean_text(config.get("deadline"))
    if deadline and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", deadline):
        raise ValueError("deadline — дата в формате YYYY-MM-DD или пустая строка")
    if deadline and common.parse_iso8601(deadline) is None:
        raise ValueError(f"deadline: некорректная дата {deadline}")
    result["deadline"] = deadline
    collective = config.get("collective", False)
    if not isinstance(collective, bool):
        raise ValueError("collective должен быть JSON-значением true или false")
    result["collective"] = collective
    members = config.get("members", [])
    if not isinstance(members, list) or any(not isinstance(item, str) for item in members):
        raise ValueError("members должен быть массивом строк")
    result["members"] = [clean_text(item) for item in members if clean_text(item)]
    ai_rules = config.get("ai_rules")
    if ai_rules is not None and not isinstance(ai_rules, str):
        raise ValueError("ai_rules — строка с правилами вуза и кафедры об использовании ИИ или пустая строка")
    result["ai_rules"] = clean_text(ai_rules)
    title_page = default_title_page()
    title_page.update({key: value for key, value in config.get("title_page", {}).items() if value != "" or key not in ("work_type", "city", "year")})
    result["title_page"] = title_page
    result["_explicit"] = explicit
    result["_normalized"] = True
    return result


def validate_root(root: Path) -> Path:
    absolute = Path(os.path.abspath(os.path.expanduser(str(root))))
    resolved = absolute.resolve()
    skill_root = Path(os.path.abspath(str(SKILL_ROOT))).resolve()
    resolved_key = os.path.normcase(str(resolved))
    skill_key = os.path.normcase(str(skill_root))
    if resolved_key == skill_key or resolved_key.startswith(skill_key.rstrip("\\/") + os.sep):
        raise ValueError("Проект нельзя создавать внутри установленного скилла: выбери отдельную папку пользователя")
    if resolved.exists() and not resolved.is_dir():
        raise ValueError(f"Путь проекта не является каталогом: {resolved}")
    return resolved


def read_state_settings(text: str) -> Dict[str, str]:
    """Настройки из vkr-state.md проекта прежней версии."""
    settings: Dict[str, str] = {}

    def value(pattern: str) -> Optional[str]:
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(1).strip().lower() if match else None

    profile = value(r"Профиль требований[^\n:]*:\**\s*`?([a-z0-9-]+)")
    if profile in common.PROFILES:
        settings["profile"] = profile
    mode = value(r"\bРежим[^\n:]*:\**\s*`?(standard|express-14d|express-7d|express-4d)")
    if mode:
        settings["mode"] = mode
    intensity = value(r"audit_intensity`?[^\n:]*:\**\s*`?(balanced|strict|maximum)")
    if intensity:
        settings["audit_intensity"] = intensity
    vkr_type = value(r"(?:Тип ВКР|Формат ВКР)[^\n:]*:\**\s*`?(project|regular|проект|обычная)")
    if vkr_type:
        settings["vkr_type"] = {"проект": "project", "обычная": "regular"}.get(vkr_type, vkr_type)
    return settings


def reconcile_existing(root: Path, config: Dict[str, Any]) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """Сверяет конфигурацию с существующим проектом: неявные значения берутся из него."""
    explicit: Set[str] = config["_explicit"]
    conflicts: List[Dict[str, Any]] = []
    adopted: Dict[str, Any] = {}
    existing: Dict[str, Any] = {}
    source = ""
    project_path = root / "vkr-project.json"
    state_path = root / "vkr-state.md"
    if project_path.is_file():
        loaded = common.read_json(project_path, label="vkr-project.json")
        if not isinstance(loaded, dict):
            raise ValueError("Существующий vkr-project.json должен быть JSON-объектом")
        existing = {key: loaded[key] for key in ("vkr_type", "profile", "mode", "audit_intensity", "collective") if key in loaded}
        source = "vkr-project.json"
    elif state_path.is_file():
        existing = read_state_settings(common.read_text_utf8(state_path, "vkr-state.md"))
        source = "vkr-state.md"
    for key, value in existing.items():
        if key in explicit:
            if value != config.get(key):
                conflicts.append({"field": key, "existing": value, "requested": config.get(key), "source": source})
        elif config.get(key) != value:
            adopted[key] = value
            config[key] = value
    if "profile" in adopted or "vkr_type" in adopted:
        required = common.PROFILE_TYPES.get(config["profile"])
        if required and config["vkr_type"] != required:
            if "vkr_type" in explicit:
                conflicts.append({"field": "vkr_type", "existing": required, "requested": config["vkr_type"], "source": source})
            else:
                config["vkr_type"] = required
                adopted["vkr_type"] = required
    if "mode" in adopted and "audit_intensity" not in explicit and "audit_intensity" not in existing:
        config["audit_intensity"] = "balanced" if config["mode"].startswith("express") else "strict"
    return conflicts, adopted


def project_json(config: Dict[str, Any], created_at: str) -> Dict[str, Any]:
    return {
        "skill_version": SKILL_VERSION,
        "protocol_version": PROTOCOL_VERSION,
        "created_at": created_at,
        "vkr_type": config["vkr_type"],
        "profile": config["profile"],
        "mode": config["mode"],
        "audit_intensity": config["audit_intensity"],
        "topic": config["topic"],
        "deadline": config["deadline"],
        "collective": config["collective"],
        "members": config["members"],
        "title_page": config["title_page"],
    }


def state_markdown(config: Dict[str, Any], today: str) -> str:
    page = config["title_page"]

    def show(value: str) -> str:
        """Обязательное поле (doctor проверяет его на prefinal/final)."""
        return value or "не указано"

    def optional(value: str) -> str:
        return value or "—"

    def title(key: str) -> str:
        """Поле титула: «не указано», если его требует сборка титульного листа, иначе «—»."""
        return show(page[key]) if key in common.TITLE_PAGE_REQUIRED else optional(page[key])

    members = ", ".join(config["members"]) or ("не указаны" if config["collective"] else "индивидуальная работа")
    chapters = ["Глава I", "Глава II"] + (["Глава III"] if config["vkr_type"] == "project" else [])
    drafts = ["drafts/annotation.md", "drafts/introduction.md"] + [f"drafts/chapter-{n}.md" for n in range(1, len(chapters) + 1)] + ["drafts/conclusion.md"]
    methodology = common.METHODOLOGY_PROJECT_PATH if config["profile"].startswith("mpgu-") else "по профилю программы (добавь в evidence/index.json → methodology)"
    return f"""# Состояние проекта ВКР

> Создано `vkr-mpgu {SKILL_VERSION}`. Схема полей: `references/state-file-pattern.md` установленного скилла.
> Разделы «Журнал текущей работы», «Независимый аудит», «Handoff и продолжение», «История сессий» (точные названия) и строка «Последнее обновление» не входят в снимок аудита. Остальное входит: после финального аудита меняй только их.
> Последнее обновление: {today}

## Паспорт

- **Тема:** {show(config['topic'])}
- **Тип ВКР:** {config['vkr_type']}
- **Профиль требований:** {config['profile']}
- **Режим:** {config['mode']}
- **Дедлайн:** {config['deadline'] or 'не задан'}
- **Студент:** {title('author')}
- **Коллективная работа:** {'да' if config['collective'] else 'нет'}; участники: {members}
- **Вуз:** {title('university')}
- **Институт/факультет:** {title('institute')}
- **Кафедра:** {title('department')}
- **Направление:** {title('program_code')} {title('program_name')}
- **Научный руководитель:** {title('supervisor')}

Каноничные значения профиля, типа, режима и титульного листа хранятся в `vkr-project.json`.

## Требования и структура

- **Методичка:** {methodology}
- **Правила кафедры об использовании ИИ:** {optional(config.get('ai_rules', ''))}
- **`requirements_source` (документ кафедры, год):** уточнить
- **`language`:** ru
- **`chapter_plan`:** не утверждён (названия глав — по plan.md после согласования)
- **`page_target`:** уточнить (число и основание или «эвристический ориентир»)
- **`source_target`:** уточнить (число и основание или «эвристический ориентир»)
- **`educational_context_required`:** уточнить
- **`pilot_min`:** not_set
- **`implementation_certificate_required`:** уточнить
- **Разделы:** аннотация, введение, {', '.join(chapters)}, заключение, список литературы, приложения
- **Черновики:** {', '.join(drafts)}
- **Сдаваемый файл:** final/vkr.docx

## Ключевые решения

- Рабочая структура создана автоматически; существующие файлы не перезаписываются.

## Журнал текущей работы

- **Статус intake:** {'captured' if config['topic'] and page['program_code'] and page['university'] else 'in_progress'}
- **Этап:** intake / plan
- **Следующее действие:** завершить intake и сформировать план

## Независимый аудит

- **`audit_autorun`:** true
- **`audit_intensity`:** {config['audit_intensity']}
- **`protocol_version`:** {PROTOCOL_VERSION}
- **Независимость:** isolated_contexts
- **Итог:** NOT_READY — обновляется по выводу `vkr_project_doctor.py`

## Handoff и продолжение

- **Handoff:** `memory/handoff.json`, `memory/handoff.md`
- **Roadmap:** `memory/roadmap.md`
- **Индекс артефактов:** `memory/artifact-index.json`
- **Реестр замечаний:** `audit/findings.json`

## История сессий

- {today}: инициализирован проект ВКР.
"""


def plan_markdown(config: Dict[str, Any]) -> str:
    chapters = ["## Глава I. [Название]", "## Глава II. [Название]"]
    if config["vkr_type"] == "project":
        chapters.append("## Глава III. [Название проектной главы]")
    body = "\n\n".join(chapters)
    return f"""# Рабочий план ВКР

> План уточняет основной агент после intake и сверки с методичкой программы.

## Введение

{body}

## Заключение

## Список использованной литературы

## Приложения
"""


def draft_files(config: Dict[str, Any]) -> Dict[str, str]:
    files = {
        "drafts/annotation.md": "# Аннотация\n\n<!-- Заполняется после основного текста. -->\n",
        "drafts/introduction.md": "# Введение\n\n<!-- Черновик создаёт основной агент после утверждения плана. -->\n",
        "drafts/chapter-1.md": "# Глава I. [Название]\n\n<!-- Черновик -->\n",
        "drafts/chapter-2.md": "# Глава II. [Название]\n\n<!-- Черновик -->\n",
        "drafts/conclusion.md": "# Заключение\n\n<!-- Черновик -->\n",
    }
    if config["vkr_type"] == "project":
        files["drafts/chapter-3.md"] = "# Глава III. [Название]\n\n<!-- Черновик проектной главы -->\n"
    return files


def initial_files(config: Dict[str, Any], created_at: str, today: str, init_event_id: str) -> Dict[str, str]:
    methodology = []
    if config["profile"].startswith("mpgu-"):
        methodology = [{"id": common.METHODOLOGY_ID, "path": common.METHODOLOGY_PROJECT_PATH, "title": "Методические рекомендации по ВКР 09.03.02 (2024)"}]
    skill_scripts = common.SCRIPTS_DIR
    handoff = {
        "schema_version": memory.SCHEMA_VERSION,
        "updated_at": common.utc_now_iso(),
        "current_phase": "intake",
        "current_task": "Завершить intake и сформировать план",
        "last_completed": "Инициализация рабочей структуры",
        "last_event_id": init_event_id,
        "next_actions": ["завершить intake", "сверить требования", "сформировать plan.md"],
        "open_questions": [],
        "open_blockers": [],
        "continuation_files": list(memory.CONTINUATION_FILES),
        "audit_seq": 0,  # счётчик audit/manifest.json на момент handoff: status видит run после него
    }
    return {
        "vkr-project.json": common.dump_json(project_json(config, created_at)),
        "vkr-state.md": state_markdown(config, today),
        "plan.md": plan_markdown(config),
        "sources.json": "[]\n",
        "evidence/index.json": common.dump_json({
            "methodology": methodology, "sources": [], "product": [], "pilot": [], "figures": [], "defense": [], "approvals": [],
        }),
        "audit/snapshot-inputs.json": common.dump_json({
            "protocol_version": PROTOCOL_VERSION, "validation_profile": config["profile"], "inputs": [], "tool_versions": {},
        }),
        "audit/manifest.json": common.dump_json(common.manifest_skeleton(config["profile"])),
        "audit/findings.json": common.dump_json(common.findings_skeleton()),
        "memory/handoff.json": common.dump_json(handoff),
        "memory/roadmap.md": """# Roadmap ВКР

| Этап | Статус | Условие завершения |
|---|---|---|
| Intake и профиль требований | IN_PROGRESS | state и vkr-project.json заполнены, требования сверены |
| План | NOT_STARTED | план согласован и прошёл аудит plan |
| Черновик глав | NOT_STARTED | все главы в drafts/*.md имеют законченный текст |
| Источники и доказательства | NOT_STARTED | источники confirmed, claims confirmed с доказательствами |
| Полировка и независимый аудит | NOT_STARTED | BLOCKER/MAJOR закрыты адресной перепроверкой на новом снимке |
| DOCX и нормоконтроль | NOT_STARTED | final/vkr.docx собран, валидатор без ошибок |
| Подготовка к защите | NOT_STARTED | ответы Test-30 записаны в evidence/defense/ |
| Финальный gate | NOT_STARTED | doctor --stage final: PASS и READY_TO_SUBMIT |
""" + ("| Пилот (проектная ВКР) | NOT_STARTED | дизайн согласован, данные в evidence/pilot/; результаты в тексте только после данных |\n" if config["vkr_type"] == "project" else ""),
        "memory/decisions.jsonl": "",
        "memory/claims-register.json": "[]\n",
        "logs/activity.jsonl": json.dumps(
            {
                "event_id": init_event_id,
                "event_type": "project_initialized",
                "timestamp": created_at,
                "actor": "primary_agent",
                "summary": "Создана рабочая структура проекта ВКР",
                "skill_version": SKILL_VERSION,
                "skill_scripts": str(skill_scripts),
            },
            ensure_ascii=False,
            separators=(",", ":"),
        ) + "\n",
        "logs/tool-runs.jsonl": "",
        **draft_files(config),
    }


def initialize(root: Path, config: Dict[str, Any], dry_run: bool = False) -> Dict[str, Any]:
    root = validate_root(root)
    if not config.get("_normalized"):
        config = normalize_config(config)
    conflicts, adopted = reconcile_existing(root, config) if root.is_dir() else ([], {})
    result: Dict[str, Any] = {
        "skill_version": SKILL_VERSION,
        "protocol_version": PROTOCOL_VERSION,
        "project_root": str(root),
        "created": [],
        "skipped_existing": [],
        "conflicts": conflicts,
        "adopted": adopted,
        "warnings": [],
    }
    if conflicts:
        result.update(status="conflict", next_action="resolve_config_conflict")
        return result
    created_at = dt.datetime.now().astimezone().replace(microsecond=0).isoformat()
    today = dt.date.today().isoformat()
    init_event_id = "evt-init-" + common.sha256_bytes(f"{root}|{created_at}".encode("utf-8"))[:16]
    index_existed = (root / "memory" / "artifact-index.json").is_file()
    if not dry_run:
        root.mkdir(parents=True, exist_ok=True)
    for relative in DIRECTORIES:
        path = root / relative
        if path.is_dir():
            result["skipped_existing"].append(relative + "/")
        else:
            result["created"].append(relative + "/")
            if not dry_run:
                path.mkdir(parents=True, exist_ok=True)
    files = initial_files(config, created_at, today, init_event_id)
    files["memory/handoff.md"] = memory.render_handoff(root, json.loads(files["memory/handoff.json"]))
    for relative, content in files.items():
        path = root / relative
        if path.exists():
            result["skipped_existing"].append(relative)
            continue
        result["created"].append(relative)
        if not dry_run:
            path.parent.mkdir(parents=True, exist_ok=True)
            with path.open("x", encoding="utf-8", newline="\n") as stream:
                stream.write(content)
    if config["profile"].startswith("mpgu-"):
        target = root / common.METHODOLOGY_PROJECT_PATH
        asset = SKILL_ROOT / "assets" / common.METHODOLOGY_ASSET
        if target.exists():
            result["skipped_existing"].append(common.METHODOLOGY_PROJECT_PATH)
        elif not asset.is_file():
            result["warnings"].append(f"Нет файла методички в скилле: {asset}")
        else:
            result["created"].append(common.METHODOLOGY_PROJECT_PATH)
            if not dry_run:
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(str(asset), str(target))
        if "evidence/index.json" in result["skipped_existing"]:
            try:
                index = common.read_json(root / "evidence" / "index.json", default=None)
            except common.ProjectDataError:
                index = None
            listed = isinstance(index, dict) and any(
                isinstance(item, dict) and item.get("id") == common.METHODOLOGY_ID
                for item in common.as_list(index.get("methodology"))
            )
            if not listed:
                result["warnings"].append(
                    "evidence/index.json уже существовал: добавь в methodology запись "
                    + json.dumps({"id": common.METHODOLOGY_ID, "path": common.METHODOLOGY_PROJECT_PATH}, ensure_ascii=False)
                )
    if not dry_run and not index_existed:
        timestamp = common.utc_now_iso()
        index = {
            "schema_version": memory.SCHEMA_VERSION,
            "skill_version": SKILL_VERSION,
            "updated_at": timestamp,
            "baseline_at": timestamp,
            "artifacts": memory.build_baseline(root, timestamp),
        }
        common.atomic_write_json(root / "memory" / "artifact-index.json", index)
    elif index_existed:
        result["warnings"].append("memory/artifact-index.json уже существовал: проверь vkr_memory.py status и при необходимости record --rebaseline")
    result.update(
        status="dry_run" if dry_run else "initialized",
        next_action="complete_intake_and_plan",
        next_commands=[
            f'python "{common.SCRIPTS_DIR / "vkr_project_doctor.py"}" "{root}" --stage draft --json',
            f'python "{common.SCRIPTS_DIR / "vkr_memory.py"}" "{root}" status --json',
        ],
    )
    return result


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("project_dir", type=Path, help="Каталог отдельного проекта ВКР (не внутри скилла)")
    parser.add_argument("--config", type=Path, help="JSON с ответами intake (схема в references/project-initialization.md)")
    parser.add_argument("--topic")
    parser.add_argument("--vkr-type", choices=common.VKR_TYPES)
    parser.add_argument("--profile", choices=common.PROFILES)
    parser.add_argument("--mode", choices=common.MODES)
    parser.add_argument("--audit-intensity", choices=common.AUDIT_INTENSITIES)
    parser.add_argument("--deadline")
    parser.add_argument("--collective", action="store_true", help="Коллективная работа")
    parser.add_argument("--student-name", help="Автор для титульного листа (title_page.author)")
    parser.add_argument("--institution", help="Вуз (title_page.university)")
    parser.add_argument("--program-code", help="Код направления (title_page.program_code)")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--json", action="store_true", dest="json_output")
    parser.add_argument("-o", "--output", type=Path, help="Записать JSON-результат в файл UTF-8")
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    common.reconfigure_stdio()
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        config = load_config(args.config)
        apply_overrides(config, args)
        result = initialize(args.project_dir, normalize_config(config), args.dry_run)
        code = 2 if result.get("status") == "conflict" else 0
    except (common.ProjectDataError, ValueError, OSError) as error:
        result = {"status": "error", "error": str(error)}
        code = 2
    except Exception as error:  # без трассировки
        result = {"status": "error", "code": "INTERNAL_ERROR", "error": f"{type(error).__name__}: {error}"}
        code = 2
    if args.output:
        common.emit_json(result, args.output)
    elif args.json_output:
        common.emit_json(result)
    elif code == 2 and result.get("status") == "error":
        print(f"ERROR: {result['error']}", file=sys.stderr)
    else:
        print(f"Status: {result['status']}")
        print(f"Project: {result.get('project_root')}")
        print(f"Created: {len(result.get('created', []))}; skipped existing: {len(result.get('skipped_existing', []))}")
        for conflict in result.get("conflicts", []):
            print(f"Conflict: {conflict['field']}: {conflict['existing']} != {conflict['requested']}")
        for warning in result.get("warnings", []):
            print(f"Warning: {warning}")
    return code


if __name__ == "__main__":
    raise SystemExit(main())
