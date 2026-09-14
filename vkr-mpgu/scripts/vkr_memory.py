#!/usr/bin/env python3
"""Долговременная память проекта ВКР: status и транзакционный checkpoint.

``status`` сверяет файлы проекта с ``memory/artifact-index.json``: видит
изменённые, исчезнувшие и новые файлы в drafts/, final/, evidence/, sources/,
audit/reports/ и в корне проекта. Служебные файлы памяти и журналов внешними
изменениями не считаются. Индекс старой схемы получает ``needs_rebaseline``.

``status`` также сообщает ``handoff_stale: true``, если ``memory/handoff.json``
устарел: отслеживаемые файлы изменены, добавлены или удалены после checkpoint,
который его записал, или после него в ``audit/manifest.json`` появились новые
действия аудита (счётчик ``seq`` больше записанного в handoff ``audit_seq``).
Тогда первый пункт ``next_actions`` — не продолжать по next_actions handoff, а
прочитать state и audit/manifest, записать ``record --all-changed`` и обновить
handoff (current_task, next_actions).

``record`` сначала записывает данные (индекс, handoff, решения), затем событие
в ``logs/activity.jsonl``. Повтор с тем же событием не дублирует запись.
Файлы checkpoint — поле ``files``/``removed_files`` события и (или) флаг
``--all-changed``: изменённые, новые и удалённые отслеживаемые файлы по тому же
расчёту, что ``status``, включая отчёты ``audit/reports/``. Event-файл внутри
проекта — служебный вход ``record``: он не индексируется и не считается внешним
изменением (путь хранится в ``event_files`` индекса).

Коды возврата: 0 — ok/recorded; 1 — external_changes_detected, needs_rebaseline
или recovery_required; 2 — ошибка входных данных.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

sys.dont_write_bytecode = True
_SCRIPTS = str(Path(__file__).resolve().parent)
if _SCRIPTS not in sys.path:
    sys.path.insert(0, _SCRIPTS)

import vkr_common as common  # noqa: E402

SCHEMA_VERSION = 2
SERVICE_PATHS = {
    "memory/handoff.json",
    "memory/handoff.md",
    "memory/artifact-index.json",
    "memory/decisions.jsonl",
}
DISCOVERY_ROOTS = ("drafts", "final", "evidence", "sources", "audit/reports")
CONTENT_MEMORY_FILES = ("memory/roadmap.md", "memory/claims-register.json")
CONTINUATION_FILES = (
    "vkr-state.md",
    "vkr-project.json",
    "plan.md",
    "sources.json",
    "evidence/index.json",
    "memory/handoff.json",
    "memory/handoff.md",
    "memory/roadmap.md",
    "memory/claims-register.json",
    "memory/artifact-index.json",
    "audit/manifest.json",
    "audit/findings.json",
    "audit/snapshot-inputs.json",
)
EVENT_ID_RE = re.compile(r"^evt-[A-Za-z0-9_.:-]{6,80}$")
HANDOFF_KEYS = ("current_phase", "current_task", "last_completed", "next_actions", "open_questions", "open_blockers")


def is_service_path(logical_id: str) -> bool:
    return logical_id in SERVICE_PATHS or logical_id.startswith("logs/")


def project_root(path: Path) -> Path:
    root = Path(os.path.abspath(os.path.expanduser(str(path))))
    if not root.is_dir():
        raise ValueError(f"Каталог проекта не найден: {root}")
    root = root.resolve()
    if not (root / "vkr-state.md").is_file():
        raise ValueError(f"В каталоге проекта нет vkr-state.md: {root}")
    return root


def resolve_project_file(root: Path, value: str) -> Tuple[str, Path]:
    path = common.resolve_inside(root, value)
    if path is None:
        raise ValueError(f"Путь вне корня проекта или не относительный: {value}")
    return common.logical_id_for(root, path), path


def read_index(root: Path) -> Dict[str, Any]:
    index = common.read_json(
        root / "memory" / "artifact-index.json",
        default={"schema_version": SCHEMA_VERSION, "updated_at": None, "artifacts": {}},
        label="memory/artifact-index.json",
    )
    if not isinstance(index, dict) or not isinstance(index.get("artifacts", {}), dict):
        raise ValueError("memory/artifact-index.json: нужен объект с полем artifacts-объектом")
    for logical_id, record in index.get("artifacts", {}).items():
        if not isinstance(record, dict):
            raise ValueError(f"memory/artifact-index.json: запись {logical_id!r} должна быть объектом с sha256")
    index.setdefault("artifacts", {})
    return index


def discover_content_files(root: Path) -> List[Path]:
    """Файлы, изменения которых отслеживает память."""
    found: List[Path] = []
    for path in sorted(root.iterdir()):
        if path.is_file() and not common.is_ignored_file_name(path.name):
            found.append(path)
    for relative in CONTENT_MEMORY_FILES:
        path = root / relative
        if path.is_file():
            found.append(path)
    for relative_root in DISCOVERY_ROOTS:
        directory = root / relative_root
        if not directory.is_dir():
            continue
        for path in sorted(item for item in directory.rglob("*") if item.is_file()):
            if common.is_ignored_file_name(path.name) or "__pycache__" in path.parts:
                continue
            found.append(path)
    return found


def artifact_record(path: Path, timestamp: str) -> Dict[str, Any]:
    return {"sha256": common.sha256_file(path), "size": path.stat().st_size, "recorded_at": timestamp}


def build_baseline(root: Path, timestamp: str, exclude: Iterable[str] = ()) -> Dict[str, Dict[str, Any]]:
    """Все отслеживаемые файлы; ``exclude`` — ключи ``dedupe_key`` путей, которые не индексируются (event-файлы)."""
    skip = set(exclude)
    baseline: Dict[str, Dict[str, Any]] = {}
    for path in discover_content_files(root):
        logical_id = common.logical_id_for(root, path)
        if common.dedupe_key(logical_id) not in skip:
            baseline[logical_id] = artifact_record(path, timestamp)
    return baseline


def index_event_files(index: Dict[str, Any]) -> List[str]:
    """Event-файлы внутри проекта, уже использованные в record: служебные, не артефакты."""
    values = index.get("event_files")
    if not isinstance(values, list):
        return []
    return sorted({common.normalize_logical_id(value) for value in values if isinstance(value, str) and value.strip()})


def event_file_id(root: Path, event_path: Optional[Path]) -> Optional[str]:
    """Логический путь event-файла, если он лежит внутри проекта, иначе None."""
    if event_path is None:
        return None
    try:
        resolved = Path(os.path.abspath(os.path.expanduser(str(event_path)))).resolve()
        return common.logical_id_for(root, resolved)
    except (OSError, RuntimeError, ValueError):
        return None


def artifact_changes(root: Path, index: Dict[str, Any], extra_event_files: Iterable[str] = ()) -> Dict[str, List[Any]]:
    """Изменённые, исчезнувшие и новые файлы относительно индекса.

    Один расчёт для ``status`` и ``record --all-changed``. Служебные записи индекса
    и event-файлы (``event_files`` индекса и ``extra_event_files``) изменениями не считаются.
    """
    event_files = sorted(set(index_event_files(index)) | {common.normalize_logical_id(value) for value in extra_event_files if value})
    event_keys = {common.dedupe_key(value) for value in event_files}
    changed: List[Dict[str, str]] = []
    missing: List[Dict[str, str]] = []
    ignored: List[str] = []
    tracked = set()
    for logical_id, previous in sorted(index["artifacts"].items()):
        normalized = common.normalize_logical_id(logical_id)
        key = common.dedupe_key(normalized)
        tracked.add(key)
        if is_service_path(normalized):
            ignored.append(normalized)
            continue
        if key in event_keys:
            continue
        path = common.resolve_inside(root, normalized)
        if path is None or not path.is_file():
            missing.append({"logical_id": normalized, "previous_sha256": str(previous.get("sha256") or "")})
            continue
        current = common.sha256_file(path)
        if current != previous.get("sha256"):
            changed.append({"logical_id": normalized, "previous_sha256": str(previous.get("sha256") or ""), "current_sha256": current})
    untracked: List[Dict[str, str]] = []
    for path in discover_content_files(root):
        logical_id = common.logical_id_for(root, path)
        key = common.dedupe_key(logical_id)
        if key not in tracked and key not in event_keys:
            untracked.append({"logical_id": logical_id, "current_sha256": common.sha256_file(path)})
    return {"changed": changed, "missing": missing, "untracked": untracked,
            "ignored_service_entries": ignored, "ignored_event_files": event_files}


def activity_event_ids(root: Path) -> set:
    rows = common.read_jsonl(root / "logs" / "activity.jsonl", label="logs/activity.jsonl")
    return {str(value.get("event_id")) for _number, value, error in rows if not error and isinstance(value, dict) and value.get("event_id")}


def manifest_audit_seq(root: Path) -> Optional[int]:
    """Счётчик действий аудита ``seq`` из audit/manifest.json (None — файла нет или он не читается)."""
    try:
        manifest = common.read_json(root / "audit" / "manifest.json", default=None, label="audit/manifest.json")
    except (common.ProjectDataError, ValueError, OSError):
        return None
    return common.safe_int(manifest.get("seq")) if isinstance(manifest, dict) else None


def memory_command(root: Path, tail: str) -> str:
    return f'python "{common.SCRIPTS_DIR / "vkr_memory.py"}" "{root}" {tail}'


def handoff_staleness(handoff: Dict[str, Any], changes: Dict[str, List[Any]], manifest: Dict[str, Any]) -> Tuple[List[str], List[str]]:
    """Почему handoff устарел: (подробные причины, краткие); пусто — не устарел.

    Handoff пишет ``record`` вместе с индексом файлов, поэтому любое расхождение файлов с
    индексом появилось после записи handoff; действия аудита сверяются по счётчику ``seq``.
    """
    if not handoff:
        return [], []
    reasons: List[str] = []
    short: List[str] = []
    paths = [item["logical_id"] for key in ("changed", "untracked", "missing") for item in changes.get(key, [])]
    if paths:
        shown = ", ".join(paths[:8]) + (f" и ещё {len(paths) - 8}" if len(paths) > 8 else "")
        reasons.append(f"после checkpoint, записавшего handoff, изменены, добавлены или удалены файлы: {shown}")
        short.append(f"файлов изменено после checkpoint: {len(paths)}")
    recorded = common.safe_int(handoff.get("audit_seq"))
    current = common.safe_int(manifest.get("seq"))
    if recorded is not None and current is not None and current > recorded:
        runs = [
            str(run.get("run_id")) for run in manifest.get("runs", [])
            if isinstance(run, dict) and max(common.safe_int(run.get("planned_seq")) or 0, common.safe_int(run.get("closed_seq")) or 0) > recorded
        ]
        listed = ", ".join(runs[-8:])
        reasons.append(f"после handoff записаны действия аудита (audit/manifest.json seq {recorded} → {current})" + (f": {listed}" if runs else ""))
        short.append(f"run аудита после handoff: {listed}" if runs else "новые действия аудита после handoff")
    return reasons, short


def status(root: Path, tail: int) -> Dict[str, Any]:
    required = ("memory/handoff.json", "memory/artifact-index.json", "audit/manifest.json")
    missing_state = [relative for relative in required if not (root / relative).is_file()]
    handoff = common.read_json(root / "memory" / "handoff.json", default={}, label="memory/handoff.json")
    if not isinstance(handoff, dict):
        raise ValueError("memory/handoff.json должен быть JSON-объектом")
    index = read_index(root)
    legacy = index.get("schema_version") != SCHEMA_VERSION
    changes = artifact_changes(root, index)
    changed, missing, untracked = changes["changed"], changes["missing"], changes["untracked"]
    manifest = common.read_json(root / "audit" / "manifest.json", default={}, label="audit/manifest.json")
    if not isinstance(manifest, dict) or not isinstance(manifest.get("runs", []), list):
        raise ValueError("audit/manifest.json: нужен объект с массивом runs")
    manifest_copy = json.loads(json.dumps(manifest))
    if not common.manifest_is_native(manifest_copy):
        manifest_copy = common.migrate_legacy_manifest(manifest_copy, "generic")
    registry, registry_problems = common.replay_findings(root, manifest_copy)
    last_event = str(handoff.get("last_event_id") or "")
    interrupted = bool(last_event) and last_event not in activity_event_ids(root)
    if missing_state or interrupted:
        state = "recovery_required"
    elif legacy:
        state = "needs_rebaseline"
    elif changed or missing or untracked:
        state = "external_changes_detected"
    else:
        state = "ok"
    recent = [value for _number, value, error in common.read_jsonl(root / "logs" / "activity.jsonl", label="logs/activity.jsonl")[-tail:] if not error] if tail else []
    stale_reasons, stale_short = ([], []) if "memory/handoff.json" in missing_state else handoff_staleness(handoff, changes, manifest)
    next_action = {
        "recovery_required": "повтори прерванный record с тем же событием или восстанови недостающие файлы памяти",
        "needs_rebaseline": "после проверки файлов выполни record --rebaseline",
        "external_changes_detected": "сверь изменения, инвалидируй зависимые проверки и запиши checkpoint (record --event <EVENT_JSON> --all-changed)",
        "ok": "продолжай по handoff",
    }[state]
    next_actions: List[str] = []
    if stale_reasons:
        record_tail = "record --rebaseline --json, затем " + memory_command(root, "record --event <event.json> --json") if legacy else "record --event <event.json> --all-changed --json"
        stale_action = (
            "Handoff устарел (" + "; ".join(stale_short) + "; подробности — handoff_stale_reasons): не продолжай по его "
            "next_actions. Прочитай vkr-state.md "
            "(«Журнал текущей работы»), audit/manifest.json и audit/findings.json (последние run и замечания), plan.md и "
            "изменённые файлы, затем запиши checkpoint " + memory_command(root, record_tail)
            + " и обнови handoff в событии: current_phase, current_task, last_completed, next_actions"
        )
        if state == "ok":
            next_action = "handoff устарел: сверь state и audit/manifest, запиши checkpoint с актуальными current_task и next_actions"
        next_actions = [next_action, stale_action] if state == "recovery_required" else [stale_action, next_action]
    else:
        next_actions = [next_action]
    return {
        "status": state,
        "project_root": str(root),
        "schema_version": index.get("schema_version"),
        "missing_state_files": missing_state,
        "interrupted_record": last_event if interrupted else None,
        "handoff": handoff,
        "handoff_stale": bool(stale_reasons),
        "handoff_stale_reasons": stale_reasons,
        "artifact_changes": changes,
        "audit_summary": {
            "snapshot_manifest_sha256": manifest.get("snapshot_manifest_sha256"),
            "runs": [
                {"run_id": item.get("run_id"), "kind": item.get("kind"), "checkpoint_id": item.get("checkpoint_id"), "status": item.get("status")}
                for item in manifest.get("runs", [])[-5:] if isinstance(item, dict)
            ],
            "open_findings": sorted(key for key, item in registry.items() if item.get("status") == "open"),
            "registry_problems": len(registry_problems),
        },
        "continuation_files": [relative for relative in CONTINUATION_FILES if (root / relative).is_file()],
        "recent_events": recent,
        "next_action": next_action,
        "next_actions": next_actions,
    }


def render_handoff(root: Path, handoff: Dict[str, Any]) -> str:
    def bullets(values: Any, empty: str = "нет") -> str:
        if not isinstance(values, list) or not values:
            return f"- {empty}"
        return "\n".join(f"- {value}" for value in values)

    continuation = "\n".join(f"- `{relative}`" for relative in CONTINUATION_FILES)
    memory_cmd = f'python "{common.SCRIPTS_DIR / "vkr_memory.py"}" "{root}" status --json'
    doctor_cmd = f'python "{common.SCRIPTS_DIR / "vkr_project_doctor.py"}" "{root}" --stage draft --json'
    return f"""# Передача контекста следующему чату

- **Обновлено:** {handoff.get('updated_at', 'не указано')}
- **Текущая фаза:** {handoff.get('current_phase', 'не указано')}
- **Текущее задание:** {handoff.get('current_task', 'не указано')}
- **Последнее завершённое:** {handoff.get('last_completed', 'не указано')}
- **Последнее событие:** {handoff.get('last_event_id', 'не указано')}

## Следующие действия

{bullets(handoff.get('next_actions'))}

## Открытые вопросы

{bullets(handoff.get('open_questions'))}

## Блокеры

{bullets(handoff.get('open_blockers'))}

## Файлы для продолжения в новом чате

{continuation}

Перед продолжением выполни:

```bash
{memory_cmd}
{doctor_cmd}
```
"""


def append_jsonl(path: Path, item: Dict[str, Any]) -> None:
    """Дописывает строку целиком; при сбое откатывает файл к прежнему размеру."""
    path.parent.mkdir(parents=True, exist_ok=True)
    line = (json.dumps(item, ensure_ascii=False, separators=(",", ":")) + "\n").encode("utf-8")
    with path.open("ab") as stream:
        stream.seek(0, os.SEEK_END)
        size = stream.tell()
        try:
            stream.write(line)
            stream.flush()
            os.fsync(stream.fileno())
        except BaseException:
            try:
                stream.truncate(size)
            except OSError:
                pass
            raise


def jsonl_values(path: Path, key: str, label: str) -> set:
    return {
        str(value.get(key)) for _number, value, error in common.read_jsonl(path, label=label)
        if not error and isinstance(value, dict) and value.get(key)
    }


def validate_event(event: Any) -> Dict[str, Any]:
    if not isinstance(event, dict):
        raise ValueError("Событие должно быть JSON-объектом")
    for key in ("event_type", "summary"):
        if not isinstance(event.get(key), str) or not event[key].strip():
            raise ValueError(f"В событии нужно непустое поле {key}")
    for key in ("files", "removed_files", "decisions"):
        if key in event and not isinstance(event[key], list):
            raise ValueError(f"{key} должен быть массивом")
    for decision in event.get("decisions", []):
        if not isinstance(decision, dict) or not str(decision.get("decision") or "").strip():
            raise ValueError("Каждое решение — объект с полем decision")
    for key in ("files", "removed_files"):
        for value in event.get(key, []):
            if not isinstance(value, str) or not value.strip():
                raise ValueError(f"{key} должен содержать непустые относительные пути")
    if "event_id" in event and not (isinstance(event["event_id"], str) and EVENT_ID_RE.match(event["event_id"])):
        raise ValueError("event_id должен иметь вид evt-<идентификатор>")
    return event


def event_content_sha(event: Dict[str, Any], records: Dict[str, Dict[str, Any]], removed: List[str], rebaseline: bool) -> str:
    material = {
        "event": {key: value for key, value in event.items() if key != "event_id"},
        "artifacts": {key: value["sha256"] for key, value in sorted(records.items())},
        "removed": sorted(removed),
        "rebaseline": rebaseline,
    }
    return hashlib.sha256(json.dumps(material, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()


def committed_events(root: Path) -> List[Dict[str, Any]]:
    rows = common.read_jsonl(root / "logs" / "activity.jsonl", label="logs/activity.jsonl")
    return [value for _number, value, error in rows if not error and isinstance(value, dict)]


def derive_event_id(event: Dict[str, Any], content_sha: str, previous_event_id: str) -> str:
    """Идентификатор не меняется при повторе прерванной записи.

    Он зависит от содержания события и от последнего подтверждённого события
    журнала: повтор после сбоя получает тот же id, а новое событие с тем же
    текстом после успешной записи — другой.
    """
    if event.get("event_id"):
        return str(event["event_id"])
    digest = hashlib.sha256(f"{previous_event_id}|{content_sha}".encode("utf-8")).hexdigest()
    return f"evt-{digest[:32]}"


def last_record(index: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Состав последнего checkpoint: record пишет его в индекс до события в журнале."""
    value = index.get("last_record")
    if not isinstance(value, dict) or not isinstance(value.get("event_id"), str):
        return None
    lists = {}
    for key in ("files", "removed_files"):
        items = value.get(key)
        if not isinstance(items, list):
            return None
        lists[key] = [common.normalize_logical_id(item) for item in items if isinstance(item, str) and item.strip()]
    return {"event_id": value["event_id"], **lists}


def add_paths(root: Path, records: Dict[str, Dict[str, Any]], removed: List[str], paths: Iterable[str],
              timestamp: str, skip_keys: set) -> None:
    """Дополняет состав checkpoint: существующий файл — в records, исчезнувший — в removed."""
    for value in paths:
        logical_id = common.normalize_logical_id(value)
        if not logical_id or common.dedupe_key(logical_id) in skip_keys or is_service_path(logical_id):
            continue
        path = common.resolve_inside(root, logical_id)
        if path is not None and path.is_file():
            records[logical_id] = artifact_record(path, timestamp)
        elif logical_id not in removed:
            removed.append(logical_id)


def record(root: Path, event_path: Optional[Path], rebaseline: bool = False, all_changed: bool = False) -> Dict[str, Any]:
    if all_changed and rebaseline:
        raise ValueError("--all-changed не сочетается с --rebaseline: переиндексация и так учитывает все файлы")
    if event_path is not None:
        event = validate_event(common.read_json(Path(event_path), label=str(event_path)))
    elif rebaseline:
        event = {"event_type": "memory_rebaseline", "summary": "Переиндексация памяти после обновления скилла или внешних изменений"}
    else:
        raise ValueError("record требует --event (файлы — в поле files или флагом --all-changed) или --rebaseline")
    timestamp = common.utc_now_iso()
    index = read_index(root)
    if all_changed and index.get("schema_version") != SCHEMA_VERSION:
        raise ValueError("memory/artifact-index.json старой схемы: проверь изменения и выполни record --rebaseline")
    # Event-файл внутри проекта — вход record, а не артефакт: не индексируется и не считается внешним изменением.
    own_event = event_file_id(root, event_path)
    event_files = set(index_event_files(index)) | ({own_event} if own_event else set())
    event_keys = {common.dedupe_key(value) for value in event_files}

    records: Dict[str, Dict[str, Any]] = {}
    for value in event.get("files", []):
        logical_id, path = resolve_project_file(root, value)
        if is_service_path(logical_id):
            raise ValueError(f"files не может содержать служебный файл памяти или журнала: {logical_id}")
        if common.dedupe_key(logical_id) in event_keys:
            continue
        if not path.is_file():
            raise ValueError(f"Файл не найден: {logical_id}")
        records[logical_id] = artifact_record(path, timestamp)
    removed: List[str] = []
    for value in event.get("removed_files", []):
        logical_id, _path = resolve_project_file(root, value)
        removed.append(logical_id)
    if rebaseline:
        records = build_baseline(root, timestamp, exclude=event_keys)

    events = committed_events(root)
    last = events[-1] if events else {}
    if all_changed:
        # Тот же расчёт, что status: изменённые и новые файлы — в records, исчезнувшие — в removed.
        changes = artifact_changes(root, index, [own_event] if own_event else [])
        add_paths(root, records, removed, [item["logical_id"] for key in ("changed", "untracked", "missing") for item in changes[key]],
                  timestamp, event_keys)
        previous = last_record(index)
        committed_ids = {str(item.get("event_id")) for item in events if item.get("event_id")}
        if previous and previous["event_id"] not in committed_ids:
            # Прерванный checkpoint: индекс уже обновлён, событие не записано. Его файлы входят в повтор,
            # поэтому id события и решения не задваиваются.
            add_paths(root, records, removed, previous["files"] + previous["removed_files"], timestamp, event_keys)
        elif previous and previous["event_id"] == last.get("event_id"):
            # Повтор уже записанной команды сверяется с составом последнего события.
            candidate, candidate_removed = dict(records), list(removed)
            add_paths(root, candidate, candidate_removed, previous["files"] + previous["removed_files"], timestamp, event_keys)
            if event_content_sha(event, candidate, candidate_removed, rebaseline) == last.get("content_sha256"):
                return {"status": "already_recorded", "event_id": last.get("event_id"), "tracked_artifacts": candidate, "removed": candidate_removed}

    content_sha = event_content_sha(event, records, removed, rebaseline)
    if last.get("content_sha256") == content_sha or (event.get("event_id") and any(item.get("event_id") == event["event_id"] for item in events)):
        return {"status": "already_recorded", "event_id": last.get("event_id") if last.get("content_sha256") == content_sha else event["event_id"], "tracked_artifacts": records, "removed": removed}
    event_id = derive_event_id(event, content_sha, str(last.get("event_id") or ""))
    activity_path = root / "logs" / "activity.jsonl"

    # 1. Данные: индекс, handoff, решения. Повтор записывает то же самое.
    if rebaseline:
        artifacts = dict(records)
    else:
        artifacts = {
            key: value for key, value in index.get("artifacts", {}).items()
            if not is_service_path(common.normalize_logical_id(key))
            and common.dedupe_key(common.normalize_logical_id(key)) not in event_keys
        }
        artifacts.update(records)
    for logical_id in removed:
        artifacts.pop(logical_id, None)
    new_index = {
        "schema_version": SCHEMA_VERSION if (rebaseline or index.get("schema_version") == SCHEMA_VERSION) else index.get("schema_version", 1),
        "skill_version": common.SKILL_VERSION,
        "updated_at": timestamp,
        "artifacts": artifacts,
        "last_record": {"event_id": event_id, "files": [] if rebaseline else sorted(records), "removed_files": sorted(removed)},
    }
    if event_files:
        new_index["event_files"] = sorted(event_files)
    if rebaseline:
        new_index["baseline_at"] = timestamp
    common.atomic_write_json(root / "memory" / "artifact-index.json", new_index)

    handoff = common.read_json(root / "memory" / "handoff.json", default={}, label="memory/handoff.json")
    if not isinstance(handoff, dict):
        raise ValueError("memory/handoff.json должен быть JSON-объектом")
    for key in HANDOFF_KEYS:
        if key in event:
            handoff[key] = event[key]
    handoff.update({
        "schema_version": SCHEMA_VERSION,
        "updated_at": timestamp,
        "last_event_id": event_id,
        "last_event_type": event["event_type"],
        "last_summary": event["summary"],
        "continuation_files": list(CONTINUATION_FILES),
    })
    # Счётчик аудита на момент handoff: status по нему видит run, записанные после checkpoint.
    audit_seq = manifest_audit_seq(root)
    if audit_seq is None:
        handoff.pop("audit_seq", None)
    else:
        handoff["audit_seq"] = audit_seq
    handoff.pop("artifacts", None)
    common.atomic_write_json(root / "memory" / "handoff.json", handoff)
    common.atomic_write_text(root / "memory" / "handoff.md", render_handoff(root, handoff))

    decisions_path = root / "memory" / "decisions.jsonl"
    existing_decisions = jsonl_values(decisions_path, "decision_id", "memory/decisions.jsonl")
    for number, decision in enumerate(event.get("decisions", []), start=1):
        decision_id = f"dec-{event_id[4:]}-{number}"
        if decision_id in existing_decisions:
            continue
        append_jsonl(decisions_path, {**decision, "decision_id": decision_id, "event_id": event_id, "timestamp": timestamp, "actor": "primary_agent"})

    logged = {
        **{key: value for key, value in event.items() if key != "event_id"},
        "event_id": event_id,
        "content_sha256": content_sha,
        "timestamp": timestamp,
        "actor": "primary_agent",
        "artifacts": records,
        "removed_files": removed,
    }
    if all_changed:
        logged["all_changed"] = True
    if event["event_type"] == "tool_run":
        tool_path = root / "logs" / "tool-runs.jsonl"
        if event_id not in jsonl_values(tool_path, "event_id", "logs/tool-runs.jsonl"):
            append_jsonl(tool_path, logged)
    # 2. Событие — последним: оно подтверждает, что checkpoint записан целиком.
    append_jsonl(activity_path, logged)
    return {
        "status": "recorded",
        "event_id": event_id,
        "timestamp": timestamp,
        "rebaseline": rebaseline,
        "tracked_artifacts": records,
        "removed": removed,
        "handoff": handoff,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("project_dir", type=Path)
    sub = parser.add_subparsers(dest="command")
    sub.required = True
    status_parser = sub.add_parser("status", help="Сверить файлы проекта с памятью")
    status_parser.add_argument("--tail", type=int, default=10)
    status_parser.add_argument("--json", action="store_true", dest="json_output")
    status_parser.add_argument("-o", "--output", type=Path, help="Записать JSON-отчёт в файл UTF-8")
    record_parser = sub.add_parser("record", help="Записать checkpoint основного агента")
    record_parser.add_argument("--event", type=Path, help="JSON события: event_type, summary, поля handoff, files, removed_files, decisions")
    record_parser.add_argument(
        "--all-changed", action="store_true",
        help="Добавить в checkpoint все изменённые, новые и удалённые отслеживаемые файлы (тот же расчёт, что status); "
             "поля files и removed_files события учитываются вместе с ними",
    )
    record_parser.add_argument("--rebaseline", action="store_true", help="Переиндексировать все отслеживаемые файлы")
    record_parser.add_argument("--json", action="store_true", dest="json_output")
    record_parser.add_argument("-o", "--output", type=Path, help="Записать JSON-отчёт в файл UTF-8")
    return parser


def render_text(result: Dict[str, Any]) -> str:
    lines = [f"Status: {result.get('status')}"]
    if result.get("error"):
        lines.append(f"Error: {result['error']}")
    if result.get("event_id"):
        lines.append(f"Event: {result['event_id']}")
    changes = result.get("artifact_changes") or {}
    for key in ("changed", "missing", "untracked"):
        for item in changes.get(key, []):
            lines.append(f"{key}: {item.get('logical_id')}")
    if result.get("handoff_stale"):
        lines.append("Handoff: устарел — " + "; ".join(result.get("handoff_stale_reasons") or []))
    if result.get("next_actions"):
        lines.extend(f"Next: {action}" for action in result["next_actions"])
    elif result.get("next_action"):
        lines.append(f"Next: {result['next_action']}")
    return "\n".join(lines) + "\n"


def main(argv: Optional[List[str]] = None) -> int:
    common.reconfigure_stdio()
    args = build_parser().parse_args(argv)
    try:
        root = project_root(args.project_dir)
        if args.command == "status":
            result = status(root, max(0, args.tail))
        else:
            result = record(root, args.event, rebaseline=args.rebaseline, all_changed=args.all_changed)
        code = 1 if result.get("status") in ("external_changes_detected", "needs_rebaseline", "recovery_required") else 0
    except (common.ProjectDataError, ValueError, OSError) as error:
        result = {"status": "error", "error": str(error)}
        code = 2
    except Exception as error:  # без трассировки в выводе агента
        result = {"status": "error", "code": "INTERNAL_ERROR", "error": f"{type(error).__name__}: {error}"}
        code = 2
    if args.output:
        common.emit_json(result, args.output)
    elif args.json_output:
        common.emit_json(result)
    else:
        sys.stdout.write(render_text(result))
    return code


if __name__ == "__main__":
    raise SystemExit(main())
