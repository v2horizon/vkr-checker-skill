#!/usr/bin/env python3
"""Диагностика проекта ВКР на стадиях draft, prefinal и final (только чтение).

Doctor не доверяет записанным файлам там, где может проверить сам: пересчитывает
входы снимка, перезапускает валидатор по сдаваемому DOCX, восстанавливает реестр
замечаний из канонических отчётов и сверяет независимость волн аудита.

Коды возврата: 0 — PASS, 1 — FAIL, 2 — ERROR (проект недоступен).
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

sys.dont_write_bytecode = True
_SCRIPTS = str(Path(__file__).resolve().parent)
if _SCRIPTS not in sys.path:
    sys.path.insert(0, _SCRIPTS)

import vkr_common as common  # noqa: E402
import vkr_markers as markers  # noqa: E402

DOCTOR_VERSION = common.PROTOCOL_VERSION
canonical_snapshot_sha = common.canonical_snapshot_sha  # совместимость импорта для прежних тестов и скриптов

REQUIRED_FILES = (
    "vkr-project.json",
    "vkr-state.md",
    "plan.md",
    "sources.json",
    "evidence/index.json",
    "audit/snapshot-inputs.json",
    "audit/manifest.json",
    "memory/handoff.json",
    "memory/handoff.md",
    "memory/roadmap.md",
    "memory/decisions.jsonl",
    "memory/claims-register.json",
    "memory/artifact-index.json",
    "logs/activity.jsonl",
    "logs/tool-runs.jsonl",
)

REQUIRED_DIRS = (
    "drafts",
    "sources/materials",
    "evidence/methodology",
    "evidence/product",
    "evidence/pilot",
    "audit/reports",
    "memory",
    "logs",
    "final",
    "exports",
    "backups/checkpoints",
)

JSON_FILES = (
    "vkr-project.json",
    "sources.json",
    "evidence/index.json",
    "audit/snapshot-inputs.json",
    "audit/manifest.json",
    "audit/findings.json",
    "memory/handoff.json",
    "memory/claims-register.json",
    "memory/artifact-index.json",
)
JSON_OBJECT_FILES = (
    "vkr-project.json",
    "evidence/index.json",
    "audit/snapshot-inputs.json",
    "audit/manifest.json",
    "audit/findings.json",
    "memory/handoff.json",
    "memory/artifact-index.json",
)
JSONL_FILES = ("memory/decisions.jsonl", "logs/activity.jsonl", "logs/tool-runs.jsonl")
MEMORY_SERVICE_PATHS = {"memory/handoff.json", "memory/handoff.md", "memory/artifact-index.json", "memory/decisions.jsonl"}
REQUIRED_DRAFTS = ("annotation.md", "introduction.md", "chapter-1.md", "chapter-2.md", "conclusion.md")
MIN_SIGNIFICANT_CHARS = 300
TITLE_PAGE_REQUIRED = common.TITLE_PAGE_REQUIRED  # совпадает с полями, которые сборка помечает [ЗАПОЛНИТЬ: …]
SOURCE_ID_RE = re.compile(r"^[A-Za-z0-9_.:-]+$")
STAGE_ORDER = {"draft": 0, "prefinal": 1, "final": 2}


def finding(code: str, severity: str, path: str, message: str, details: Any = None) -> Dict[str, Any]:
    item: Dict[str, Any] = {"code": code, "severity": severity, "path": path, "message": message}
    if details:
        item["details"] = details
    return item


class Context:
    def __init__(self, root: Path, stage: str) -> None:
        self.root = root
        self.stage = stage
        self.findings: List[Dict[str, Any]] = []
        self.data: Dict[str, Any] = {}
        self.project: Dict[str, Any] = {}
        self.profile = "generic"
        self.vkr_type = ""
        self.evidence_entries: List[Dict[str, Any]] = []
        self.snapshot: Optional[Dict[str, Any]] = None
        self.current_sha: Optional[str] = None
        self.manifest: Optional[Dict[str, Any]] = None
        self.registry: Dict[str, Dict[str, Any]] = {}
        self.registry_ok = False
        self.waves: Optional[Tuple[Dict[str, Any], Dict[str, Any]]] = None
        self.degraded = False
        self.recorded_sha: Optional[str] = None
        self.manifest_tampered = False
        self.final_matrix: Optional[Dict[str, Any]] = None
        self.pages: Optional[Dict[str, Any]] = None  # кэш источников числа страниц (page_sources)

    def add(self, code: str, severity: str, path: str, message: str, details: Any = None) -> None:
        self.findings.append(finding(code, severity, path, message, details))

    def by_stage(self, draft: str, prefinal: str, final: str) -> str:
        return {"draft": draft, "prefinal": prefinal, "final": final}[self.stage]

    @property
    def final(self) -> bool:
        return self.stage == "final"

    @property
    def late(self) -> bool:
        return self.stage in ("prefinal", "final")


# ---------------------------------------------------------------------------
# Структура, JSON, конфигурация
# ---------------------------------------------------------------------------


def check_structure(ctx: Context) -> None:
    for relative in REQUIRED_FILES:
        if not (ctx.root / relative).is_file():
            ctx.add("FILE_MISSING", "ERROR", relative, "Обязательный файл отсутствует")
    for relative in REQUIRED_DIRS:
        if not (ctx.root / relative).is_dir():
            ctx.add("DIR_MISSING", "ERROR", relative, "Обязательный каталог отсутствует")


def load_json_files(ctx: Context) -> None:
    for relative in JSON_FILES:
        path = ctx.root / relative
        if not path.is_file():
            ctx.data[relative] = None
            continue
        try:
            ctx.data[relative] = common.read_json(path, label=relative)
        except common.ProjectDataError as error:
            ctx.data[relative] = None
            if relative == "audit/findings.json":
                ctx.add("AUDIT_FINDINGS_VIEW_INVALID", "WARNING", relative, f"Представление реестра повреждено ({error.detail}); doctor пересчитывает реестр из отчётов")
            else:
                ctx.add(error.code, "ERROR", relative, error.detail)
            continue
        if relative in JSON_OBJECT_FILES and not isinstance(ctx.data[relative], dict):
            code = "PROJECT_CONFIG_SHAPE" if relative == "vkr-project.json" else "JSON_OBJECT_REQUIRED"
            ctx.add(code, "ERROR", relative, "Ожидается JSON-объект")
            ctx.data[relative] = None
    for relative, label in (("sources.json", "SOURCES_SHAPE"), ("memory/claims-register.json", "CLAIMS_SHAPE")):
        if ctx.data.get(relative) is not None and not isinstance(ctx.data[relative], list):
            ctx.add(label, "ERROR", relative, "Ожидается JSON-массив")
            ctx.data[relative] = None


def check_project_config(ctx: Context) -> None:
    project = ctx.data.get("vkr-project.json")
    if not isinstance(project, dict):
        return
    ctx.project = project
    for key in ("skill_version", "profile", "mode", "vkr_type"):
        if not str(project.get(key) or "").strip():
            ctx.add("PROJECT_FIELD_MISSING", "ERROR", "vkr-project.json", f"Нет поля {key}")
    vkr_type = project.get("vkr_type")
    if vkr_type not in common.VKR_TYPES:
        ctx.add("PROJECT_TYPE_INVALID", "ERROR", "vkr-project.json", "vkr_type должен быть project или regular")
    else:
        ctx.vkr_type = str(vkr_type)
    profile = project.get("profile")
    if profile not in common.PROFILES:
        ctx.add("PROJECT_PROFILE_INVALID", "ERROR", "vkr-project.json", "Неизвестный профиль требований: " + ", ".join(common.PROFILES))
    else:
        ctx.profile = str(profile)
        expected_type = common.PROFILE_TYPES.get(ctx.profile)
        if expected_type and vkr_type in common.VKR_TYPES and vkr_type != expected_type:
            ctx.add("PROJECT_PROFILE_TYPE_MISMATCH", "ERROR", "vkr-project.json", f"Профиль {ctx.profile} требует vkr_type={expected_type}")
    if project.get("mode") not in common.MODES:
        ctx.add("PROJECT_MODE_INVALID", "ERROR", "vkr-project.json", "Неизвестный режим работы: " + ", ".join(common.MODES))
    intensity = project.get("audit_intensity")
    if intensity is not None and intensity not in common.AUDIT_INTENSITIES:
        ctx.add("PROJECT_AUDIT_INTENSITY_INVALID", "ERROR", "vkr-project.json", "audit_intensity: balanced, strict или maximum")
    if "collective" in project and not isinstance(project.get("collective"), bool):
        ctx.add("PROJECT_COLLECTIVE_INVALID", "ERROR", "vkr-project.json", "collective должен быть true или false")
    if "title_page" in project and not isinstance(project.get("title_page"), dict):
        ctx.add("PROJECT_TITLE_PAGE_INVALID", "ERROR", "vkr-project.json", "title_page должен быть объектом")
    if str(project.get("protocol_version") or "") != common.PROTOCOL_VERSION:
        ctx.add(
            "PROJECT_PROTOCOL_OLD", "WARNING", "vkr-project.json",
            f"Проект создан протоколом {project.get('protocol_version') or 'до 6.33'}; прежние записи аудита не учитываются, "
            "выполни vkr_memory.py record --rebaseline и новые волны через vkr_audit.py",
        )


def check_jsonl(ctx: Context) -> None:
    for relative in JSONL_FILES:
        path = ctx.root / relative
        try:
            rows = common.read_jsonl(path, label=relative)
        except common.ProjectDataError as error:
            ctx.add(error.code, "ERROR", relative, error.detail)
            continue
        for number, value, error in rows:
            if error:
                ctx.add("JSONL_INVALID", "ERROR", f"{relative}:{number}", error)
            elif not isinstance(value, dict):
                ctx.add("JSONL_SHAPE", "ERROR", f"{relative}:{number}", "Каждая строка должна быть JSON-объектом")


def _state_value(text: str, label: str) -> Optional[str]:
    match = re.search(rf"{label}[^\n:]*:\**\s*`?([A-Za-z0-9_.-]+)", text)
    return match.group(1) if match else None


def check_state(ctx: Context) -> None:
    path = ctx.root / "vkr-state.md"
    if not path.is_file():
        return
    try:
        text = common.read_text_utf8(path, "vkr-state.md")
    except common.ProjectDataError as error:
        ctx.add(error.code, "ERROR", "vkr-state.md", error.detail)
        return
    for marker in ("Тема", "Профиль требований", "Режим", "audit_intensity", "protocol_version"):
        if marker not in text:
            ctx.add("STATE_FIELD_MISSING", "ERROR", "vkr-state.md", f"Не найдено поле: {marker}")
    projection = common.state_projection(text)
    missing_lines = [line.strip(" -*") for line in projection.splitlines() if "не указано" in line.casefold()]
    if missing_lines:
        ctx.add(
            "STATE_INCOMPLETE", ctx.by_stage("WARNING", "ERROR", "ERROR"), "vkr-state.md",
            "Остались незаполненные паспортные данные («не указано»): " + "; ".join(missing_lines[:5]), missing_lines[:10],
        )
    state_profile = _state_value(text, "Профиль требований")
    if ctx.project and state_profile in common.PROFILES and state_profile != ctx.project.get("profile"):
        ctx.add(
            "STATE_PROFILE_MISMATCH", "WARNING", "vkr-state.md",
            f"В state профиль {state_profile}, в vkr-project.json — {ctx.project.get('profile')}; каноничен vkr-project.json",
        )


# ---------------------------------------------------------------------------
# Доказательства, источники, черновики
# ---------------------------------------------------------------------------


def check_evidence_index(ctx: Context) -> None:
    index = ctx.data.get("evidence/index.json")
    if index is None:
        return
    entries, problems = common.evidence_entries(ctx.root, index)
    ctx.evidence_entries = entries
    for problem in problems:
        severity = problem["severity"]
        if severity == "ERROR" and ctx.stage == "draft":
            severity = "WARNING"
        ctx.add(problem["code"], severity, problem["path"], problem["message"])
    if ctx.profile.startswith("mpgu-") and ctx.late:
        methodology = [entry for entry in entries if entry["group"] == "methodology" and entry.get("file")]
        if not methodology:
            ctx.add(
                "METHODOLOGY_EVIDENCE_MISSING", ctx.by_stage("WARNING", "WARNING", "ERROR"), "evidence/index.json",
                f"Для профиля {ctx.profile} нужна методичка в evidence/index.json → methodology ({common.METHODOLOGY_PROJECT_PATH})",
            )


def check_sources(ctx: Context) -> None:
    sources = ctx.data.get("sources.json")
    if not isinstance(sources, list):
        return
    if not sources and ctx.late:
        ctx.add("SOURCES_EMPTY", "ERROR", "sources.json", "Список источников пуст")
    seen_keys: Dict[str, int] = {}
    seen_ids: Dict[str, int] = {}
    unconfirmed: List[str] = []
    unverified: List[str] = []
    unstructured: List[str] = []
    for index, source in enumerate(sources, start=1):
        path = f"sources.json[{index}]"
        if not isinstance(source, dict):
            ctx.add("SOURCE_SHAPE", "ERROR", path, "Источник должен быть JSON-объектом")
            continue
        source_id = source.get("id")
        label = str(source_id) if source_id is not None else f"#{index}"
        if source_id is None or isinstance(source_id, bool) or not (
            isinstance(source_id, int) or (isinstance(source_id, str) and SOURCE_ID_RE.match(source_id))
        ):
            ctx.add("SOURCE_ID_INVALID", ctx.by_stage("WARNING", "ERROR", "ERROR"), path, "id — строка [A-Za-z0-9_.:-]+ или целое число")
        else:
            key = str(source_id).casefold()
            if key in seen_ids:
                ctx.add("SOURCE_ID_DUPLICATE", ctx.by_stage("WARNING", "ERROR", "ERROR"), path, f"id {source_id} уже есть в позиции {seen_ids[key]}")
            seen_ids[key] = index
        title = source.get("title")
        raw_text = source.get("raw_text")
        if not title and not raw_text:
            ctx.add("SOURCE_TITLE_MISSING", ctx.by_stage("WARNING", "ERROR", "ERROR"), path, "Нет title или raw_text")
            continue
        if not title:
            unstructured.append(label)
        doi = str(source.get("doi") or "").strip().casefold()
        dedupe = "doi:" + doi if doi else "title:" + "".join(ch for ch in str(title or raw_text).casefold() if ch.isalnum())
        if dedupe in seen_keys:
            ctx.add("SOURCE_DUPLICATE", "WARNING", path, f"Возможный дубликат позиции {seen_keys[dedupe]}")
        else:
            seen_keys[dedupe] = index
        status = common.normalize_source_status(source.get("status"))
        if status not in common.SOURCE_STATUSES:
            ctx.add("SOURCE_STATUS_INVALID", ctx.by_stage("WARNING", "ERROR", "ERROR"), path, "status: " + ", ".join(common.SOURCE_STATUSES))
        if status != "confirmed":
            unconfirmed.append(label)
            continue
        verification = source.get("verification")
        method_ok = isinstance(verification, dict) and verification.get("method") in common.VERIFICATION_METHODS
        checked_ok = isinstance(verification, dict) and common.parse_iso8601(verification.get("checked_at")) is not None
        trace_ok = isinstance(verification, dict) and (str(verification.get("url") or "").strip() or str(verification.get("note") or "").strip())
        if not (method_ok and checked_ok and trace_ok):
            unverified.append(label)
    if unconfirmed and ctx.late:
        ctx.add(
            "SOURCE_NOT_CONFIRMED", ctx.by_stage("INFO", "WARNING", "ERROR"), "sources.json",
            f"Не подтверждены источники ({len(unconfirmed)}): " + ", ".join(unconfirmed[:30]), unconfirmed,
        )
    if unverified and ctx.late:
        ctx.add(
            "SOURCE_VERIFICATION_MISSING", ctx.by_stage("INFO", "WARNING", "ERROR"), "sources.json",
            "У confirmed-источников нет verification {method, checked_at, url|note}: " + ", ".join(unverified[:30]), unverified,
        )
    if unstructured and ctx.late:
        ctx.add(
            "SOURCE_UNSTRUCTURED", ctx.by_stage("INFO", "WARNING", "ERROR"), "sources.json",
            "Записи только с raw_text нельзя оформить по ГОСТ: " + ", ".join(unstructured[:30]), unstructured,
        )


def _read_markdown(ctx: Context, path: Path) -> Optional[str]:
    try:
        return common.read_text_utf8(path, common.logical_id_for(ctx.root, path))
    except common.ProjectDataError as error:
        ctx.add(error.code, "ERROR", error.path, error.detail)
        return None


def markdown_targets(ctx: Context) -> List[Path]:
    targets = [ctx.root / "plan.md"]
    drafts = ctx.root / "drafts"
    if drafts.is_dir():
        targets.extend(sorted(path for path in drafts.glob("*.md") if path.is_file()))
    return [path for path in targets if path.is_file()]


def check_placeholders(ctx: Context) -> None:
    for path in markdown_targets(ctx):
        text = _read_markdown(ctx, path)
        if text is None:
            continue
        found = markers.find_marker_details(markers.strip_markdown_code(text))
        if not found:
            continue
        relative = common.logical_id_for(ctx.root, path)
        hard = [fragment for _code, fragment in found if not markers.is_soft_marker(fragment)]
        soft = [fragment for _code, fragment in found if markers.is_soft_marker(fragment)]
        if ctx.stage == "draft":
            ctx.add("PLACEHOLDER_FOUND", "INFO", relative, f"Найдено маркеров: {len(found)}", [fragment for _c, fragment in found][:10])
            continue
        if ctx.stage == "prefinal":
            if hard:
                ctx.add("PLACEHOLDER_FOUND", "ERROR", relative, f"Незакрытые маркеры: {len(hard)}", hard[:10])
            if soft:
                ctx.add("PLACEHOLDER_CHECK_PENDING", "WARNING", relative, f"Непроверенные места [ПРОВЕРИТЬ…]: {len(soft)}", soft[:10])
        else:
            ctx.add("PLACEHOLDER_FOUND", "ERROR", relative, f"Незакрытые маркеры: {len(found)}", [fragment for _c, fragment in found][:10])


def significant_chars(text: str) -> int:
    cleaned = markers.strip_markdown_code(text)
    cleaned = re.sub(r"<!--.*?-->", " ", cleaned, flags=re.DOTALL)
    lines = [line for line in cleaned.splitlines() if not re.match(r"^\s{0,3}#{1,6}\s", line)]
    cleaned = "\n".join(lines)
    for fragment in markers.find_markers(cleaned):
        cleaned = cleaned.replace(fragment, " ")
    return sum(1 for character in cleaned if character.isalnum())


def required_drafts(ctx: Context) -> List[str]:
    names = list(REQUIRED_DRAFTS)
    if ctx.profile == "mpgu-09-project":
        names.append("chapter-3.md")
    return names


def check_drafts(ctx: Context) -> None:
    if not ctx.late:
        return
    plan = ctx.root / "plan.md"
    if plan.is_file():
        text = _read_markdown(ctx, plan)
        if text is not None and not text.strip():
            ctx.add("PLAN_EMPTY", "ERROR", "plan.md", "План работы пуст")
    drafts = ctx.root / "drafts"
    if drafts.is_dir():
        for path in sorted(drafts.glob("*.md")):
            text = _read_markdown(ctx, path)
            if text is not None and not text.strip():
                ctx.add("DRAFT_EMPTY", "ERROR", common.logical_id_for(ctx.root, path), "Файл раздела пуст")
    for name in required_drafts(ctx):
        relative = f"drafts/{name}"
        path = ctx.root / relative
        if not path.is_file():
            ctx.add("DRAFT_REQUIRED_MISSING", "ERROR", relative, "Обязательный раздел отсутствует; для доработки готового DOCX выполни import_docx.py")
            continue
        text = _read_markdown(ctx, path)
        if text is None or not text.strip():
            continue
        count = significant_chars(text)
        if count < MIN_SIGNIFICANT_CHARS:
            ctx.add(
                "DRAFT_TOO_SHORT", "ERROR", relative,
                f"В разделе {count} значимых символов (без заголовков, комментариев и маркеров); нужно не меньше {MIN_SIGNIFICANT_CHARS}",
            )


def check_title_page(ctx: Context) -> None:
    if not ctx.late or not ctx.project:
        return
    title_page = ctx.project.get("title_page")
    missing = []
    if not str(ctx.project.get("topic") or "").strip():
        missing.append("topic")
    missing.extend(common.title_page_missing(title_page))
    if missing:
        ctx.add(
            "TITLE_PAGE_INCOMPLETE", ctx.by_stage("INFO", "WARNING", "ERROR"), "vkr-project.json",
            "Не заполнены данные титульного листа: " + ", ".join(missing), missing,
        )


# ---------------------------------------------------------------------------
# Память
# ---------------------------------------------------------------------------


def check_memory(ctx: Context) -> None:
    index = ctx.data.get("memory/artifact-index.json")
    if not isinstance(index, dict):
        return
    artifacts = index.get("artifacts", {})
    if not isinstance(artifacts, dict):
        ctx.add("ARTIFACT_INDEX_SHAPE", "ERROR", "memory/artifact-index.json", "artifacts должен быть объектом")
        return
    if index.get("schema_version") != 2:
        ctx.add(
            "MEMORY_NEEDS_REBASELINE", "WARNING", "memory/artifact-index.json",
            "Индекс памяти старой схемы: выполни vkr_memory.py <проект> record --rebaseline",
        )
    changed: List[str] = []
    missing: List[str] = []
    for logical_id, recorded in artifacts.items():
        if not isinstance(logical_id, str) or not isinstance(recorded, dict):
            ctx.add("ARTIFACT_RECORD_SHAPE", "ERROR", "memory/artifact-index.json", f"Некорректная запись артефакта {logical_id!r}")
            continue
        normalized = common.normalize_logical_id(logical_id)
        if normalized in MEMORY_SERVICE_PATHS or normalized.startswith("logs/"):
            continue
        path = common.resolve_inside(ctx.root, normalized)
        if path is None:
            ctx.add("MEMORY_PATH_OUTSIDE", "ERROR", logical_id, "Путь выходит за корень проекта")
            continue
        if not path.is_file():
            missing.append(normalized)
        elif common.sha256_file(path) != str(recorded.get("sha256") or ""):
            changed.append(normalized)
    if missing:
        ctx.add("ARTIFACT_MISSING", "WARNING", "memory/artifact-index.json", "Зафиксированные файлы исчезли; запиши checkpoint vkr_memory.py record --all-changed (или укажи их в removed_files)", missing)
    if changed:
        ctx.add("ARTIFACT_CHANGED", "WARNING", "memory/artifact-index.json", "Файлы изменены после последнего checkpoint памяти; запиши vkr_memory.py record --all-changed", changed)


# ---------------------------------------------------------------------------
# Снимок и реестр аудита
# ---------------------------------------------------------------------------


def check_snapshot(ctx: Context) -> None:
    snapshot, _problems = common.build_snapshot(ctx.root, ctx.profile)
    ctx.snapshot = snapshot
    ctx.current_sha = common.canonical_snapshot_sha(snapshot)
    recorded = ctx.data.get("audit/snapshot-inputs.json")
    manifest = ctx.data.get("audit/manifest.json")
    if not isinstance(recorded, dict) or not common.as_list(recorded.get("inputs")):
        ctx.add(
            "AUDIT_SNAPSHOT_MISSING", ctx.by_stage("INFO", "WARNING", "ERROR"), "audit/snapshot-inputs.json",
            "Снимок аудита ещё не создан: vkr_audit.py <проект> snapshot",
        )
        return
    recorded_sha = common.canonical_snapshot_sha(recorded)
    manifest_sha = manifest.get("snapshot_manifest_sha256") if isinstance(manifest, dict) else None
    problems = []
    if recorded_sha is None:
        problems.append("snapshot-inputs.json повреждён")
    elif manifest_sha != recorded_sha:
        problems.append("snapshot_manifest_sha256 в manifest не совпадает с каноническим хешем snapshot-inputs.json")
    if recorded_sha:
        stored_path = ctx.root / "audit" / "snapshots" / f"{recorded_sha}.json"
        try:
            stored = common.read_json(stored_path, default=None)
        except common.ProjectDataError:
            stored = None
        if common.canonical_snapshot_sha(stored) != recorded_sha:
            problems.append(f"нет неизменяемой копии audit/snapshots/{recorded_sha}.json")
    if problems:
        ctx.add("AUDIT_SNAPSHOT_HASH_MISMATCH", ctx.by_stage("WARNING", "ERROR", "ERROR"), "audit/snapshot-inputs.json", "; ".join(problems))
    if recorded_sha != ctx.current_sha:
        diff = common.diff_inputs(recorded.get("inputs"), snapshot["inputs"])
        changed_ids = diff["changed"] + diff["added"] + diff["removed"]
        extra = []
        if recorded.get("validation_profile") != snapshot["validation_profile"]:
            extra.append("validation_profile")
        if recorded.get("protocol_version") != snapshot["protocol_version"] or recorded.get("tool_versions") != snapshot["tool_versions"]:
            extra.append("protocol/tool_versions")
        details = {"changed": diff["changed"], "added": diff["added"], "removed": diff["removed"], "other": extra}
        ctx.add(
            "AUDIT_SNAPSHOT_STALE", ctx.by_stage("INFO", "WARNING", "ERROR"), "audit/snapshot-inputs.json",
            "Файлы изменились после снимка: " + ", ".join((changed_ids + extra)[:20]), details,
        )


def _runs(manifest: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [run for run in common.as_list(manifest.get("runs")) if isinstance(run, dict)]


def check_audit_registry(ctx: Context) -> None:
    manifest = ctx.data.get("audit/manifest.json")
    if not isinstance(manifest, dict):
        return
    manifest = json.loads(json.dumps(manifest))
    recorded = manifest.get("snapshot_manifest_sha256")
    ctx.recorded_sha = recorded if isinstance(recorded, str) else None
    integrity = common.manifest_integrity_problems(manifest)
    for problem in integrity:
        ctx.add(problem["code"], ctx.by_stage("WARNING", "ERROR", "ERROR"), problem["path"], problem["message"])
    ctx.manifest_tampered = bool(integrity)
    if common.manifest_is_native(manifest):
        legacy_runs = common.as_list(manifest.get("legacy_runs"))
        analysed = manifest
    else:
        legacy_runs = common.as_list(manifest.get("runs"))
        analysed = common.migrate_legacy_manifest(manifest, ctx.profile) if not integrity else {"runs": [], "schema": common.MANIFEST_SCHEMA}
    if legacy_runs and not integrity:
        ctx.add(
            "AUDIT_LEGACY_RUNS", "WARNING", "audit/manifest.json",
            f"{len(legacy_runs)} run записаны по протоколу до 6.33 и не учитываются; проведи проверки через vkr_audit.py plan",
        )
    orphans, legacy_reports = common.report_files_status(ctx.root, manifest)
    if orphans:
        ctx.add(
            "AUDIT_REPORT_ORPHAN", ctx.by_stage("WARNING", "ERROR", "ERROR"), "audit/reports",
            "Отчёты не связаны с записями manifest (удалённые или подложенные run, прерванная запись)", orphans[:20],
        )
    if legacy_reports:
        ctx.add(
            "AUDIT_LEGACY_REPORTS", "WARNING", "audit/reports",
            "Отчёты протокола до 6.33 лежат среди текущих: перенеси их командой vkr_audit.py <проект> migrate-legacy",
            legacy_reports[:20],
        )
    archive = ctx.root / "audit" / "legacy"
    if archive.is_dir() and any(path.is_file() for path in archive.rglob("*")):
        ctx.add(
            "AUDIT_HISTORY_ARCHIVED", "INFO", "audit/legacy",
            "Прежние записи аудита перенесены в audit/legacy и не учитываются; их замечания проверяются заново новыми волнами",
        )
    ctx.manifest = analysed
    registry, problems = common.replay_findings(ctx.root, analysed)
    ctx.registry = registry
    ctx.registry_ok = not problems
    for problem in problems:
        ctx.add(problem["code"], ctx.by_stage("WARNING", "ERROR", "ERROR"), problem["path"], problem["message"])

    open_high = [item for item in registry.values() if item.get("status") == "open" and item.get("severity") in common.HIGH_SEVERITIES]
    if open_high:
        ctx.add(
            "AUDIT_FINDINGS_OPEN", ctx.by_stage("WARNING", "ERROR", "ERROR"), "audit/findings.json",
            "Открыты замечания BLOCKER/MAJOR: " + ", ".join(item["finding_id"] for item in open_high),
            [
                {"finding_id": item["finding_id"], "severity": item["severity"], "gate": item.get("base_gate"),
                 "run_id": (item.get("source") or {}).get("run_id"), "observed": item.get("observed")}
                for item in open_high
            ],
        )

    focus = {sha for sha in (ctx.current_sha, ctx.recorded_sha) if sha}
    incomplete = []
    unclosed = []
    for run in _runs(analysed):
        status = common.normalize_status(run.get("status"))
        blocking = common.run_blocking_reasons(run, registry)
        if status in ("planned", "running") or (status == "reported" and blocking["pending"]):
            incomplete.append({
                "run_id": run.get("run_id"), "status": run.get("status"), "snapshot_manifest_sha256": run.get("snapshot_manifest_sha256"),
                "pending_tasks": blocking["pending"],
            })
        if run.get("snapshot_manifest_sha256") in focus and blocking["unclosed"]:
            unclosed.extend(f"{run.get('run_id')}/{item}" for item in blocking["unclosed"])
    if incomplete:
        ctx.add(
            "AUDIT_RUNS_INCOMPLETE", ctx.by_stage("INFO", "ERROR", "ERROR"), "audit/manifest.json",
            "Есть незавершённые run: " + ", ".join(f"{item['run_id']} ({item['status']})" for item in incomplete),
            incomplete,
        )
    if unclosed:
        ctx.add(
            "AUDIT_FAIL_UNCLOSED", ctx.by_stage("WARNING", "ERROR", "ERROR"), "audit/manifest.json",
            "На текущем или записанном снимке есть незакрытый fail; повторные волны на том же снимке его не закрывают", unclosed,
        )

    view = ctx.data.get("audit/findings.json")
    view_has_findings = isinstance(view, dict) and bool(common.as_list(view.get("findings")))
    if (registry or view_has_findings) and view is not None and not common.views_equal(view, common.findings_view(registry)):
        ctx.add("AUDIT_FINDINGS_VIEW_OUTDATED", "INFO", "audit/findings.json", "Представление реестра устарело; обновится командой vkr_audit.py findings")

    if ctx.stage == "prefinal" and common.project_intensity(ctx.project) in ("strict", "maximum"):
        current = [run for run in _runs(analysed) if run.get("snapshot_manifest_sha256") == ctx.current_sha]
        if not any(run.get("kind") == "primary" and wave_problems(ctx, run, "primary", final=False) == [] for run in current):
            ctx.add(
                "AUDIT_PREFINAL_WAVE_MISSING", "WARNING", "audit/manifest.json",
                "Матрица strict/maximum перед передачей руководителю предусматривает волну девяти ролей: "
                "vkr_audit.py plan --kind primary --preset prefinal",
            )


# ---------------------------------------------------------------------------
# Финальный DOCX и валидатор
# ---------------------------------------------------------------------------


def check_final_docx_build(ctx: Context) -> None:
    if not ctx.late or not (ctx.root / common.FINAL_DOCX_PATH).is_file():
        return
    problems = common.build_manifest_problems(ctx.root)
    if problems:
        # prefinal тоже ERROR: руководителю передаётся final/vkr.docx, и он должен совпадать с черновиками.
        ctx.add(
            "FINAL_DOCX_OUTDATED", "ERROR", common.FINAL_DOCX_PATH,
            "final/vkr.docx не соответствует черновикам: " + "; ".join(problems[:3]), problems[:20],
        )


DOCX_FIELDS_PATH = common.DOCX_FIELDS_PATH
ANNOTATION_DRAFT = "drafts/annotation.md"
# «52 с.», «52 стр.», «52 страницы», «на 52 страницах» (пробел может быть неразрывным).
ANNOTATION_PAGES_RE = re.compile(r"(?<![\d.,])(\d{1,4})\s*(?:с\.|стр\.?(?![а-яё])|страниц[а-яё]*)", re.IGNORECASE)
# Объём назван не целиком, а по части работы: методичка (structure-project.md) говорит об объёме
# текстовой части, а Word считает страницы всего файла — расхождение здесь не ошибка автора.
ANNOTATION_PARTIAL_RE = re.compile(
    r"текстов\w*\s+част|основн\w+\s+(?:текст|част)|без\s+(?:учёта\s+|учета\s+)?приложени"
    r"|без\s+списка\s+литератур|не\s+считая\s+приложени",
    re.IGNORECASE,
)


def _annotation_clean(text: str) -> str:
    return re.sub(r"<!--.*?-->", " ", markers.strip_markdown_code(text), flags=re.DOTALL)


def annotation_page_counts(text: str) -> List[int]:
    """Числа страниц, указанные в аннотации (без кода и HTML-комментариев)."""
    cleaned = _annotation_clean(text)
    return [int(match.group(1)) for match in ANNOTATION_PAGES_RE.finditer(cleaned) if int(match.group(1)) > 0]


def annotation_partial_markers(text: str) -> List[str]:
    """Оговорки «текстовой части / основного текста / без приложений» рядом с числом страниц."""
    cleaned = _annotation_clean(text)
    found: List[str] = []
    for match in ANNOTATION_PAGES_RE.finditer(cleaned):
        if int(match.group(1)) <= 0:
            continue
        # Предложение вокруг числа: от предыдущей точки/переноса строки до следующей.
        left = max(cleaned.rfind(".", 0, match.start()), cleaned.rfind("\n", 0, match.start()),
                   cleaned.rfind(";", 0, match.start()))
        right = min((position for position in (cleaned.find(".", match.end()), cleaned.find("\n", match.end()),
                                               cleaned.find(";", match.end())) if position != -1), default=len(cleaned))
        marker = ANNOTATION_PARTIAL_RE.search(cleaned[left + 1:right])
        if marker and marker.group(0) not in found:
            found.append(marker.group(0))
    return found


def docx_fields_for(ctx: Context, docx: Path) -> Tuple[Optional[Dict[str, Any]], str]:
    """exports/docx-fields.json, если он относится к этому DOCX; иначе (None, причина).

    Запись относится к DOCX по ``text_fingerprint`` (отпечаток сдаваемого текста, как в
    build-manifest): обязательная очистка метаданных после update_docx_fields.py меняет байты и
    ``docx_sha256``, но не текст. ``docx_sha256`` сверяется только у записи без отпечатка.
    """
    path = ctx.root / DOCX_FIELDS_PATH
    if not path.is_file():
        return None, f"нет {DOCX_FIELDS_PATH} (его пишет update_docx_fields.py после обновления полей в Word)"
    try:
        record = common.read_json(path, label=DOCX_FIELDS_PATH)
    except common.ProjectDataError as error:
        return None, f"{DOCX_FIELDS_PATH} не читается: {error.detail}"
    if not isinstance(record, dict):
        return None, f"{DOCX_FIELDS_PATH} должен быть JSON-объектом"
    pages = record.get("pages")
    if isinstance(pages, bool) or not isinstance(pages, int) or pages <= 0:
        return None, f"в {DOCX_FIELDS_PATH} нет числа страниц (pages)"
    other_version = f"{DOCX_FIELDS_PATH} относится к другой версии final/vkr.docx (текст пересобран или изменён после update_docx_fields.py)"
    fingerprint = record.get("text_fingerprint")
    if fingerprint:
        if record.get("fingerprint_schema") not in (None, common.FINGERPRINT_SCHEMA):
            return None, f"{DOCX_FIELDS_PATH} записан с другой схемой отпечатка текста: повтори update_docx_fields.py"
        current, error = common.docx_text_fingerprint(docx)
        if error:
            return None, f"не удалось прочитать текст final/vkr.docx: {error}"
        return (record, "") if current == fingerprint else (None, other_version)
    if str(record.get("docx_sha256") or "").casefold() == common.sha256_file(docx):
        return record, ""
    return None, other_version


def page_sources(ctx: Context, docx: Path) -> Dict[str, Any]:
    """Откуда известно число страниц сдаваемого DOCX: сам файл и exports/docx-fields.json.

    ``app_pages`` — ``docProps/app.xml → <Pages>`` (его пишет Word при сохранении, очистка
    метаданных его не трогает); ``record`` — запись update_docx_fields.py, относящаяся к этому
    тексту. Надёжным считается ``app_pages``: его нельзя подменить, не изменив сам DOCX (а это
    меняет sha и требует нового снимка). Кэшируется: проверок объёма две.
    """
    if ctx.pages is not None:
        return ctx.pages
    app_pages, app_reason = common.docx_app_pages(docx)
    record, record_reason = docx_fields_for(ctx, docx)
    result = {"app_pages": app_pages, "app_reason": app_reason,
              "record": record, "record_reason": record_reason,
              "pages": app_pages if app_pages is not None else (record["pages"] if record else None)}
    result["source"] = "DOCX (docProps/app.xml, Word)" if app_pages is not None else (
        DOCX_FIELDS_PATH if record else "")
    ctx.pages = result
    return result


def check_docx_pages(ctx: Context) -> None:
    """Число страниц в exports/docx-fields.json сверяется с самим DOCX; отметка word_opened."""
    if not ctx.late:
        return
    docx = ctx.root / common.FINAL_DOCX_PATH
    if not docx.is_file():
        return
    sources = page_sources(ctx, docx)
    record, app_pages = sources["record"], sources["app_pages"]
    if record is None:
        return
    if app_pages is not None and int(record["pages"]) != int(app_pages):
        ctx.add(
            "DOCX_PAGES_MISMATCH", "ERROR", DOCX_FIELDS_PATH,
            f"В {DOCX_FIELDS_PATH} записано {record['pages']} с., а в самом final/vkr.docx "
            f"{app_pages} с. (docProps/app.xml, число страниц пишет Word). Файл с числом страниц "
            "правили вручную или он остался от другой версии: повтори update_docx_fields.py",
            {"record_pages": record["pages"], "docx_pages": app_pages,
             "docx_fields_updated_at": record.get("updated_at")},
        )
        return
    if ctx.final and not record.get("word_opened"):
        ctx.add(
            "DOCX_FIELDS_UNCONFIRMED", "WARNING", DOCX_FIELDS_PATH,
            f"В {DOCX_FIELDS_PATH} нет отметки word_opened: число страниц записано не после успешного "
            "открытия документа в Word (запись прежней версии скрипта или правка вручную). "
            "Повтори update_docx_fields.py — он проверит, что Word открывает сдаваемый файл",
            {"pages": record["pages"], "updated_at": record.get("updated_at")},
        )


def check_annotation_pages(ctx: Context) -> None:
    """Объём в аннотации совпадает с числом страниц final/vkr.docx по Word (сам DOCX, затем docx-fields.json)."""
    if not ctx.late:
        return
    docx = ctx.root / common.FINAL_DOCX_PATH
    annotation = ctx.root / ANNOTATION_DRAFT
    if not docx.is_file() or not annotation.is_file():
        return
    text = _read_markdown(ctx, annotation)
    if text is None or not text.strip():
        return
    stated = annotation_page_counts(text)
    if not stated:
        ctx.add(
            "ANNOTATION_PAGES_UNVERIFIED", "INFO", ANNOTATION_DRAFT,
            "В аннотации не найден объём в страницах («N с.», «N стр.», «N страниц»): методичка требует сведения об объёме",
        )
        return
    sources = page_sources(ctx, docx)
    actual = sources["pages"]
    if actual is None:
        reason = "; ".join(item for item in (sources["record_reason"], sources["app_reason"]) if item)
        # На final нет ни одного источника числа страниц — это не справка, а работа для человека.
        ctx.add(
            "ANNOTATION_PAGES_UNVERIFIED", "WARNING" if ctx.final else "INFO", ANNOTATION_DRAFT,
            f"Объём в аннотации ({', '.join(map(str, stated))} с.) не сверен с DOCX, проверь число страниц вручную "
            f"(открой final/vkr.docx в Word): {reason}",
            {"stated": stated, "reason": reason},
        )
        return
    if actual in stated or (len(stated) > 1 and sum(stated) == actual):
        return
    details = {"stated": stated, "actual": actual, "source": sources["source"]}
    record = sources["record"]
    if record is not None:
        details["docx_fields_updated_at"] = record.get("updated_at")
    partial = annotation_partial_markers(text)
    if partial:
        # «Объём текстовой части — 10 страниц без учёта приложений» (structure-project.md): законная формулировка.
        details["partial_markers"] = partial
        ctx.add(
            "ANNOTATION_PAGES_PARTIAL", "WARNING", ANNOTATION_DRAFT,
            f"В аннотации указан объём части работы ({', '.join(map(str, stated))} с., оговорка «{partial[0]}»), "
            f"а doctor знает только полный счётчик Word — в final/vkr.docx {actual} с. Проверь число вручную: "
            "в аннотации надёжнее указать полный объём по Word (при необходимости — «в том числе N с. приложений»)",
            details,
        )
        return
    ctx.add(
        "ANNOTATION_PAGES_MISMATCH", "ERROR", ANNOTATION_DRAFT,
        f"В аннотации указано {', '.join(map(str, stated))} с., а в final/vkr.docx {actual} с. "
        f"(Word, {sources['source']})",
        details,
    )


def check_final_docx(ctx: Context) -> None:
    if not ctx.final:
        return
    docs, extras = common.final_listing(ctx.root)
    if extras:
        ctx.add("FINAL_EXTRA_FILES", "ERROR", "final/", "В final/ кроме сдаваемого vkr.docx есть другие документы или папки", extras)
    if docs and not (ctx.root / common.FINAL_DOCX_PATH).is_file():
        ctx.add("FINAL_DOCX_NAME", "ERROR", "final/", "Сдаваемый файл должен называться final/vkr.docx (так его собирает build_vkr.py)")
    if not docs:
        ctx.add("FINAL_DOCX_MISSING", "ERROR", "final/", "Нет сдаваемого final/vkr.docx")
        return
    if len(docs) > 1:
        ctx.add(
            "FINAL_DOCX_MULTIPLE", "ERROR", "final/",
            "В final/ должен быть ровно один сдаваемый DOCX; промежуточные сборки храни в exports/",
            [common.logical_id_for(ctx.root, path) for path in docs],
        )
    for path in docs:
        relative = common.logical_id_for(ctx.root, path)
        if not zipfile.is_zipfile(path):
            ctx.add("FINAL_DOCX_INVALID", "ERROR", relative, "Файл не является корректным DOCX/ZIP")
            continue
        try:
            with zipfile.ZipFile(path) as archive:
                broken = archive.testzip()
                missing = {"[Content_Types].xml", "word/document.xml"} - set(archive.namelist())
        except (OSError, zipfile.BadZipFile) as error:
            ctx.add("FINAL_DOCX_CORRUPT", "ERROR", relative, str(error))
            continue
        if broken:
            ctx.add("FINAL_DOCX_CORRUPT", "ERROR", relative, f"Повреждённый элемент архива: {broken}")
        if missing:
            ctx.add("FINAL_DOCX_INVALID", "ERROR", relative, "Нет обязательных частей DOCX: " + ", ".join(sorted(missing)))
        if broken or missing:
            continue
        package = common.docx_package_problems(path)
        if package:
            ctx.add(
                "FINAL_DOCX_CORRUPT", "ERROR", relative,
                "Word не откроет файл («Файл поврежден»): " + "; ".join(package[:3]), package[:20],
            )


def add_validation_warnings(ctx: Context, report: Dict[str, Any], path: str) -> None:
    """Предупреждения валидатора не блокируют gate, но должны быть видны в сводке doctor."""
    warnings = common.validator_warning_count(report)
    if warnings:
        ctx.add(
            "VALIDATION_WARNINGS", "WARNING", path,
            f"Предупреждений валидатора: {warnings} — прочитай audit/automated-validation.json и реши по каждому (исправить или обосновать)",
            common.validator_warning_messages(report),
        )


def report_recorded_validation_warnings(ctx: Context) -> None:
    """prefinal: предупреждения записанного отчёта по текущему final DOCX, без повторного запуска валидатора."""
    docs = common.final_docx_files(ctx.root)
    if len(docs) != 1 or not zipfile.is_zipfile(docs[0]):
        return
    report_path = "audit/automated-validation.json"
    if not (ctx.root / report_path).is_file():
        return
    try:
        recorded = common.read_json(ctx.root / report_path, label=report_path)
    except common.ProjectDataError:
        return
    if isinstance(recorded, dict) and recorded.get("document_sha256") == common.sha256_file(docs[0]):
        add_validation_warnings(ctx, recorded, report_path)


def check_validation(ctx: Context) -> None:
    if not ctx.final:
        report_recorded_validation_warnings(ctx)
        return
    docs = common.final_docx_files(ctx.root)
    if len(docs) != 1 or not zipfile.is_zipfile(docs[0]):
        return
    docx = docs[0]
    docx_sha = common.sha256_file(docx)
    report_path = "audit/automated-validation.json"
    recorded: Any = None
    if not (ctx.root / report_path).is_file():
        ctx.add("AUTOMATED_VALIDATION_MISSING", "ERROR", report_path, "Нет отчёта валидатора: vkr_audit.py <проект> validate")
    else:
        try:
            recorded = common.read_json(ctx.root / report_path, label=report_path)
        except common.ProjectDataError as error:
            ctx.add("AUTOMATED_VALIDATION_INVALID", "ERROR", report_path, error.detail)
        else:
            if not isinstance(recorded, dict):
                ctx.add("AUTOMATED_VALIDATION_INVALID", "ERROR", report_path, "Отчёт валидатора должен быть JSON-объектом")
                recorded = None
            elif recorded.get("document_sha256") != docx_sha or recorded.get("profile") != ctx.profile:
                ctx.add("AUTOMATED_VALIDATION_STALE", "ERROR", report_path, "Отчёт валидатора относится к другому DOCX или профилю: vkr_audit.py <проект> validate")

    result = common.run_validator(docx, ctx.profile)
    if not result.get("available"):
        ctx.add("VALIDATOR_UNAVAILABLE", "ERROR", common.logical_id_for(ctx.root, docx), f"Повторный запуск валидатора невозможен: {result.get('error')}")
        return
    report = result.get("report")
    if not isinstance(report, dict) or report.get("error") or common.normalize_status(report.get("status")) == "error":
        message = report.get("error") if isinstance(report, dict) else "нет отчёта"
        ctx.add("VALIDATION_FAILED_TO_RUN", "ERROR", common.logical_id_for(ctx.root, docx), f"Валидатор не смог проверить DOCX: {message}")
        return
    errors = common.validator_error_count(report)
    if errors is None:
        ctx.add("VALIDATION_FAILED_TO_RUN", "ERROR", common.logical_id_for(ctx.root, docx), "В отчёте валидатора нет summary.errors")
        return
    if errors > 0:
        ctx.add(
            "VALIDATION_ERRORS", "ERROR", common.logical_id_for(ctx.root, docx),
            f"Повторный запуск валидатора: ошибок {errors}", common.validator_error_messages(report),
        )
    add_validation_warnings(ctx, report, common.logical_id_for(ctx.root, docx))
    if isinstance(recorded, dict) and recorded.get("document_sha256") == docx_sha and common.validator_error_count(recorded) != errors:
        ctx.add(
            "AUTOMATED_VALIDATION_MISMATCH", "ERROR", report_path,
            f"Записанный отчёт (errors={common.validator_error_count(recorded)}) не совпадает с повторным запуском (errors={errors})",
        )


# ---------------------------------------------------------------------------
# Финальные волны
# ---------------------------------------------------------------------------


def wave_problems(ctx: Context, run: Dict[str, Any], kind: str, final: bool = True) -> List[str]:
    """Причины, по которым run не засчитывается как полная волна ``kind``.

    Для финальной волны нужен пресет final и матрица реплик continuous-audit.md
    для audit_intensity и vkr_type проекта.
    """
    reasons: List[str] = []
    if run.get("kind") != kind:
        return [f"kind={run.get('kind')}"]
    if run.get("snapshot_manifest_sha256") != ctx.current_sha:
        return ["другой снимок"]
    if final and run.get("preset") != "final":
        reasons.append(f"пресет {run.get('preset')!r}: финальная волна планируется с --preset final")
    if common.normalize_status(run.get("status")) != "complete":
        reasons.append(f"status={run.get('status')} (нужен complete после close)")
    if common.safe_int(run.get("closed_seq")) is None:
        reasons.append("run не закрыт командой close")
    snapshot_kinds = {str(item.get("kind")) for item in common.as_list((ctx.snapshot or {}).get("inputs")) if isinstance(item, dict)}
    has_product = "evidence_product" in snapshot_kinds
    requirements = common.final_wave_requirements(kind, common.project_intensity(ctx.project), ctx.vkr_type) if final else {gate: 1 for gate in common.BASE_GATES}
    covered: Dict[str, set] = {gate: set() for gate in common.BASE_GATES}
    auditors_by_gate: Dict[str, set] = {}
    auditors_by_role: Dict[str, List[str]] = {}
    degraded = run.get("independence") == common.DEGRADED
    for task in common.as_list(run.get("planned_tasks")):
        if not isinstance(task, dict):
            reasons.append("повреждённая задача")
            continue
        task_id = task.get("task_id")
        gate = str(task.get("base_gate") or "").upper()
        role = str(task.get("auditor_role") or "").upper()
        lens = task.get("audit_lens")
        expected = common.role_gate_lens(role)
        if expected is None or expected != (gate, lens):
            reasons.append(f"{task_id}: роль {role} и линза {lens} не согласованы (базовая роль — audit_lens=full_gate)")
            continue
        status = common.normalize_status(task.get("status"))
        attempts = [attempt for attempt in common.as_list(task.get("attempts")) if isinstance(attempt, dict)]
        if not attempts:
            reasons.append(f"{task_id}: нет записанного отчёта")
            continue
        last = attempts[-1]
        if task.get("report_sha256") != last.get("report_sha256") or task.get("report_path") != last.get("report_path"):
            reasons.append(f"{task_id}: report sha mismatch")
        failed_attempt = False
        for attempt in attempts:
            report, error = common.load_attempt_report(ctx.root, run, task, attempt)
            if error:
                reasons.append(f"{task_id}: {error}")
                failed_attempt = True
                continue
            assert report is not None
            if report.get("independence") == common.DEGRADED:
                degraded = True
        if failed_attempt:
            continue
        if status == "not_applicable":
            allowed, why = common.not_applicable_allowed(gate, run.get("preset"), ctx.vkr_type, "docx" in snapshot_kinds, has_product)
            if not allowed:
                reasons.append(f"{task_id}: {why}")
                continue
        elif status != "pass":
            reasons.append(f"{task_id}: status={status}")
            continue
        last_key = common.auditor_key(last.get("auditor_id"))
        auditors_by_role.setdefault(role, []).append(last_key)
        for attempt in attempts:
            key = common.auditor_key(attempt.get("auditor_id"))
            if not degraded and common.is_forbidden_auditor(attempt.get("auditor_id")):
                reasons.append(f"{task_id}: auditor_id {attempt.get('auditor_id')} не является независимым")
            auditors_by_gate.setdefault(key, set()).add(gate)
        if role == gate and lens == common.FULL_GATE_LENS:
            covered[gate].add(str(task.get("replica_id")))
    short = [f"{gate}: {len(covered[gate])} из {count}" for gate, count in requirements.items() if len(covered[gate]) < count]
    if short:
        reasons.append("не хватает пройденных реплик full_gate по матрице: " + ", ".join(short))
    if not degraded:
        shared = sorted(key for key, gates in auditors_by_gate.items() if len(gates) > 1)
        if shared:
            reasons.append("один auditor_id закрывает несколько gate: " + ", ".join(shared))
        doubled = sorted(role for role, keys in auditors_by_role.items() if len(set(keys)) < len(keys))
        if doubled:
            reasons.append("реплики одной роли выполнены одним аудитором: " + ", ".join(doubled))
    run_high = [
        item["finding_id"] for item in ctx.registry.values()
        if (item.get("source") or {}).get("run_id") == run.get("run_id")
        and item.get("status") == "open" and item.get("severity") in common.HIGH_SEVERITIES
    ]
    if run_high:
        reasons.append("открыты замечания run: " + ", ".join(run_high))
    run["_degraded"] = degraded
    return reasons


def run_participants(run: Dict[str, Any]) -> Tuple[set, set, set]:
    auditors, tasks, executions = set(), set(), set()
    for task in common.as_list(run.get("planned_tasks")):
        if not isinstance(task, dict):
            continue
        tasks.add(str(task.get("task_id")))
        for attempt in common.as_list(task.get("attempts")):
            if isinstance(attempt, dict):
                auditors.add(common.auditor_key(attempt.get("auditor_id")))
                executions.add(str(attempt.get("execution_id")))
    return auditors, tasks, executions


def check_intensity(ctx: Context, primary: Dict[str, Any]) -> None:
    """WARNING, если финальные волны проведены с audit_intensity ниже, чем более ранние run проекта."""
    rank = common.INTENSITY_RANK
    applied = primary.get("audit_intensity")
    if applied not in rank:
        applied = common.project_intensity(ctx.project)
    if ctx.final_matrix is not None:
        ctx.final_matrix["applied_audit_intensity"] = applied
    planned = common.safe_int(primary.get("planned_seq")) or 0
    earlier = [
        run for run in _runs(ctx.manifest or {})
        if (common.safe_int(run.get("planned_seq")) or 0) < planned and run.get("audit_intensity") in rank
    ]
    if not earlier:
        return
    top = max(earlier, key=lambda run: rank[run["audit_intensity"]])
    if rank[top["audit_intensity"]] > rank[applied]:
        ctx.add(
            "AUDIT_INTENSITY_LOWERED", "WARNING", "vkr-project.json",
            f"Финальные волны проведены с audit_intensity={applied}, а более ранний {top.get('run_id')} — с "
            f"{top['audit_intensity']}: матрица финального аудита урезана",
            {"applied": applied, "earlier_max": top["audit_intensity"], "earlier_run": top.get("run_id")},
        )


def check_final_waves(ctx: Context) -> None:
    if not ctx.final or ctx.manifest is None:
        return
    intensity = common.project_intensity(ctx.project)
    ctx.final_matrix = {
        "audit_intensity": intensity,
        "vkr_type": ctx.vkr_type,
        "primary": common.final_wave_requirements("primary", intensity, ctx.vkr_type),
        "blind_regression": common.final_wave_requirements("blind_regression", intensity, ctx.vkr_type),
    }
    current = [run for run in _runs(ctx.manifest) if run.get("snapshot_manifest_sha256") == ctx.current_sha]
    primary_details: Dict[str, List[str]] = {}
    valid_primary = []
    for run in current:
        if run.get("kind") != "primary":
            continue
        problems = wave_problems(ctx, run, "primary")
        if problems:
            primary_details[str(run.get("run_id"))] = problems
        else:
            valid_primary.append(run)
    if not valid_primary:
        ctx.add(
            "AUDIT_PRIMARY_WAVE_MISSING", "ERROR", "audit/manifest.json",
            "Нет завершённой primary-волны --preset final с полной матрицей реплик на текущем снимке",
            primary_details or {"current_snapshot": ["на текущем снимке нет primary run — vkr_audit.py plan --kind primary --preset final"]},
        )
        return
    blind_details: Dict[str, List[str]] = {}
    for blind in current:
        if blind.get("kind") != "blind_regression":
            continue
        problems = wave_problems(ctx, blind, "blind_regression")
        if problems:
            blind_details[str(blind.get("run_id"))] = problems
            continue
        blind_auditors, blind_tasks, blind_executions = run_participants(blind)
        pair_problems: List[str] = []
        for primary in valid_primary:
            reasons = []
            if (common.safe_int(blind.get("planned_seq")) or 0) <= (common.safe_int(primary.get("closed_seq")) or 0):
                reasons.append(f"запланирована до закрытия {primary.get('run_id')}")
            auditors, tasks, executions = run_participants(primary)
            degraded_pair = bool(primary.get("_degraded") or blind.get("_degraded"))
            if not degraded_pair and auditors & blind_auditors:
                reasons.append(f"пересекаются auditor_id с {primary.get('run_id')}: " + ", ".join(sorted(auditors & blind_auditors)))
            if tasks & blind_tasks:
                reasons.append(f"пересекаются task_id с {primary.get('run_id')}")
            if executions & blind_executions:
                reasons.append(f"пересекаются execution_id с {primary.get('run_id')}")
            if not reasons:
                ctx.waves = (primary, blind)
                ctx.degraded = degraded_pair
                check_intensity(ctx, primary)
                return
            pair_problems.extend(reasons)
        blind_details[str(blind.get("run_id"))] = pair_problems
    ctx.add(
        "AUDIT_BLIND_WAVE_MISSING", "ERROR", "audit/manifest.json",
        "Нет более поздней завершённой blind_regression-волны --preset final новыми аудиторами на текущем снимке",
        blind_details or {"current_snapshot": ["vkr_audit.py plan --kind blind_regression --preset final"]},
    )


def check_defense_evidence(ctx: Context) -> None:
    if not ctx.final:
        return
    for entry in ctx.evidence_entries:
        path = entry.get("file")
        if entry.get("group") == "defense" and path is not None and path.stat().st_size > 0:
            return
    ctx.add(
        "USER_EVIDENCE_REQUIRED", "ERROR", "evidence/index.json",
        "Нет записанных ответов пользователя (Test-30, вопросы комиссии) в evidence/defense/ с записью в evidence/index.json → defense",
    )


# ---------------------------------------------------------------------------
# Утверждения
# ---------------------------------------------------------------------------


def check_claims(ctx: Context) -> None:
    claims = ctx.data.get("memory/claims-register.json")
    if not isinstance(claims, list):
        return
    if not claims and ctx.final:
        ctx.add("CLAIMS_EMPTY", "ERROR", "memory/claims-register.json", "Нет реестра ключевых утверждений")
    sources = ctx.data.get("sources.json") if isinstance(ctx.data.get("sources.json"), list) else []
    source_status = {
        str(item.get("id")).casefold(): common.normalize_source_status(item.get("status"))
        for item in sources if isinstance(item, dict) and item.get("id") is not None
    }
    evidence_by_id = {entry["id"].casefold(): entry for entry in ctx.evidence_entries}
    snapshot_ids = {
        common.dedupe_key(str(item.get("logical_id"))) for item in common.as_list((ctx.snapshot or {}).get("inputs")) if isinstance(item, dict)
    }
    for index, claim in enumerate(claims, start=1):
        label = f"memory/claims-register.json[{index}]"
        if not isinstance(claim, dict):
            ctx.add("CLAIM_SHAPE", "ERROR", label, "Утверждение должно быть объектом")
            continue
        name = str(claim.get("claim_id") or f"#{index}")
        status = common.normalize_claim_status(claim.get("status"))
        if status not in common.CLAIM_STATUSES:
            ctx.add("CLAIM_STATUS_INVALID", ctx.by_stage("WARNING", "ERROR", "ERROR"), label, f"{name}: status — " + ", ".join(common.CLAIM_STATUSES))
            continue
        if status in ("unsupported", "invalidated"):
            ctx.add("CLAIM_UNSUPPORTED", ctx.by_stage("WARNING", "ERROR", "ERROR"), label, f"{name}: статус {status}")
        elif status in ("pending", "partial"):
            ctx.add("CLAIM_NOT_CONFIRMED", ctx.by_stage("INFO", "WARNING", "ERROR"), label, f"{name}: статус {status}")
        for key in ("source_ids", "evidence_ids", "evidence"):
            if key in claim and not isinstance(claim.get(key), list):
                ctx.add("CLAIM_REFS_SHAPE", "ERROR", label, f"{name}: {key} должен быть массивом")
        if not ctx.late:
            continue
        missing_fields = [key for key in ("claim_id", "text", "location") if not str(claim.get(key) or "").strip()]
        if missing_fields:
            ctx.add("CLAIM_FIELDS_MISSING", ctx.by_stage("INFO", "WARNING", "ERROR"), label, f"{name}: нет полей " + ", ".join(missing_fields))
        if not ctx.final or status != "confirmed":
            continue
        source_ids = [str(value) for value in common.as_list(claim.get("source_ids"))]
        evidence_ids = [str(value) for value in common.as_list(claim.get("evidence_ids"))]
        direct = [str(value) for value in common.as_list(claim.get("evidence"))]
        if not (source_ids or evidence_ids or direct):
            ctx.add("CLAIM_EVIDENCE_MISSING", "ERROR", label, f"{name}: нет source_ids, evidence_ids или evidence")
        confirmed_source = False
        local_evidence = False
        url_only: List[str] = []
        for source_id in source_ids:
            status_value = source_status.get(source_id.casefold())
            if status_value is None:
                ctx.add("CLAIM_SOURCE_REF_MISSING", "ERROR", label, f"{name}: source id {source_id} нет в sources.json")
            elif status_value != "confirmed":
                ctx.add("CLAIM_SOURCE_NOT_CONFIRMED", "ERROR", label, f"{name}: источник {source_id} не подтверждён")
            else:
                confirmed_source = True
        for evidence_id in evidence_ids:
            entry = evidence_by_id.get(evidence_id.casefold())
            if entry is None:
                hint = " (это id источника: перенеси в source_ids)" if evidence_id.casefold() in source_status else ""
                ctx.add("CLAIM_EVIDENCE_REF_MISSING", "ERROR", label, f"{name}: evidence id {evidence_id} нет в evidence/index.json{hint}")
            elif entry.get("url"):
                url_only.append(evidence_id)
            elif entry.get("file") is not None and entry["file"].stat().st_size == 0:
                ctx.add("CLAIM_EVIDENCE_FILE_EMPTY", "ERROR", label, f"{name}: файл {entry['path']} пуст")
            elif entry.get("file") is not None:
                local_evidence = True
        for value in direct:
            if common.is_url(value):
                url_only.append(value)
                continue
            path = common.resolve_inside(ctx.root, value)
            logical = common.logical_id_for(ctx.root, path) if path is not None else common.normalize_logical_id(value)
            if path is None or not logical.casefold().startswith(common.EVIDENCE_LOCAL_PREFIXES):
                ctx.add("CLAIM_EVIDENCE_PATH_INVALID", "ERROR", label, f"{name}: {value} — доказательство должно лежать в evidence/ или sources/materials/")
            elif not path.is_file() or path.stat().st_size == 0:
                ctx.add("CLAIM_EVIDENCE_LOCATOR_MISSING", "ERROR", label, f"{name}: файл {logical} отсутствует или пуст")
            elif common.dedupe_key(logical) not in snapshot_ids:
                ctx.add("CLAIM_EVIDENCE_NOT_IN_SNAPSHOT", "ERROR", label, f"{name}: {logical} не входит в снимок — добавь его в evidence/index.json")
            else:
                local_evidence = True
        if url_only and not (confirmed_source or local_evidence):
            ctx.add(
                "CLAIM_EVIDENCE_URL_ONLY", "ERROR", label,
                f"{name}: доказательства только по URL ({', '.join(url_only[:3])}) — нужен подтверждённый источник или локальная копия в снимке",
            )
        elif url_only:
            ctx.add("CLAIM_EVIDENCE_URL_ONLY", "WARNING", label, f"{name}: часть доказательств только по URL, машинно не проверяется")


# ---------------------------------------------------------------------------
# Итог
# ---------------------------------------------------------------------------

CHECKS: Tuple[Tuple[str, Callable[[Context], None]], ...] = (
    ("structure", check_structure),
    ("json", load_json_files),
    ("project_config", check_project_config),
    ("jsonl", check_jsonl),
    ("state", check_state),
    ("evidence_index", check_evidence_index),
    ("sources", check_sources),
    ("placeholders", check_placeholders),
    ("drafts", check_drafts),
    ("title_page", check_title_page),
    ("memory", check_memory),
    ("snapshot", check_snapshot),
    ("audit_registry", check_audit_registry),
    ("final_docx", check_final_docx),
    ("final_docx_build", check_final_docx_build),
    ("docx_pages", check_docx_pages),
    ("annotation_pages", check_annotation_pages),
    ("validation", check_validation),
    ("final_waves", check_final_waves),
    ("defense_evidence", check_defense_evidence),
    ("claims", check_claims),
)


def _cmd(script: str, root: Path, tail: str) -> str:
    return f'python "{common.SCRIPTS_DIR / script}" "{root}" {tail}'.rstrip()


def build_next_actions(ctx: Context) -> List[str]:
    root = ctx.root
    errors = {item["code"]: item for item in ctx.findings if item["severity"] == "ERROR"}
    warnings = {item["code"]: item for item in ctx.findings if item["severity"] in ("WARNING", "INFO")}
    everything = {**warnings, **errors}
    actions: List[str] = []

    def add(text: str) -> None:
        if text not in actions:
            actions.append(text)

    manifest_broken = any(item["code"] in ("JSON_INVALID", "ENCODING_INVALID", "JSON_OBJECT_REQUIRED") and item["path"] == "audit/manifest.json" for item in ctx.findings)
    tampered = bool({"AUDIT_RUN_TAMPERED", "AUDIT_SEQ_GAP", "AUDIT_REPORT_ORPHAN", "AUDIT_REPORT_MISMATCH", "AUDIT_ATTEMPT_INVALID"} & set(errors))
    if "INTERNAL_ERROR" in errors:
        add("Doctor не смог выполнить часть проверок (INTERNAL_ERROR): сообщи об ошибке инструмента, остальные находки достоверны")
    if {"FILE_MISSING", "DIR_MISSING"} & set(errors):
        add("Восстанови структуру проекта (существующие файлы не перезаписываются): " + _cmd("init_vkr_project.py", root, "--json"))
    if manifest_broken:
        add("audit/manifest.json повреждён: восстанови его из audit/manifest.json.bak (копия перед последней записью инструмента); snapshot и plan до восстановления не выполняй")
    for item in ctx.findings:
        if item["severity"] == "ERROR" and item["code"] in ("JSON_INVALID", "ENCODING_INVALID", "JSONL_INVALID", "PROJECT_CONFIG_SHAPE", "JSON_OBJECT_REQUIRED") and item["path"] != "audit/manifest.json":
            add(f"Исправь {item['path']}: {item['message']}")
    if "AUDIT_FINDINGS_VIEW_INVALID" in everything:
        add("Представление реестра audit/findings.json повреждено: пересобери его из отчётов командой " + _cmd("vkr_audit.py", root, "findings --json"))
    if tampered:
        add(
            "Записи аудита изменены вручную или удалены: восстанови audit/manifest.json из audit/manifest.json.bak; если копии нет — "
            "перенеси audit/manifest.json, audit/findings.json, audit/reports и audit/briefs в audit/legacy/<дата>/ и проведи волны заново"
        )
    if "AUDIT_LEGACY_REPORTS" in everything:
        add("Перенеси отчёты протокола до 6.33 в audit/legacy: " + _cmd("vkr_audit.py", root, "migrate-legacy --json"))
    if any(code.startswith("PROJECT_") and code != "PROJECT_PROTOCOL_OLD" for code in errors):
        add("Исправь vkr-project.json (профиль, тип ВКР, режим, collective) — см. references/project-initialization.md")
    if "STATE_INCOMPLETE" in errors or "STATE_INCOMPLETE" in warnings:
        lines = (everything.get("STATE_INCOMPLETE") or {}).get("details") or []
        add("Заполни в vkr-state.md поля со значением «не указано»" + (": " + "; ".join(lines[:5]) if lines else "") + " (данные титульного листа — в vkr-project.json → title_page)")
    if {"DRAFT_REQUIRED_MISSING", "DRAFT_TOO_SHORT", "DRAFT_EMPTY", "PLAN_EMPTY"} & set(errors):
        add("Допиши обязательные разделы в drafts/*.md; если есть готовый DOCX — перенеси его в черновики через import_docx.py")
    if "PLACEHOLDER_FOUND" in errors:
        add("Закрой маркеры незавершённости в plan.md и drafts/*.md (подробности в details)")
    if "PLACEHOLDER_CHECK_PENDING" in warnings:
        add("Проверь источники с пометкой [ПРОВЕРИТЬ] (verify_sources.py) и обнови их status в sources.json")
    if {"SOURCE_NOT_CONFIRMED", "SOURCE_VERIFICATION_MISSING", "SOURCE_UNSTRUCTURED", "SOURCES_EMPTY"} & set(errors):
        add("Подтверди источники: status=confirmed и verification {method, checked_at, url|note}; неподтверждённые удали из текста и реестра")
    if "USER_EVIDENCE_REQUIRED" in errors:
        add("Проведи с пользователем Test-30 и вопросы комиссии, запиши ответы в evidence/defense/<файл>.md и добавь запись в evidence/index.json → defense")
    if "TITLE_PAGE_INCOMPLETE" in everything:
        if isinstance(ctx.project, dict) and not isinstance(ctx.project.get("title_page"), dict) and any(key in ctx.project for key in common.LEGACY_FLAT_TITLE_KEYS):
            add("Перенеси плоские поля титула 6.32 в title_page: " + _cmd("vkr_audit.py", root, "migrate-legacy --json"))
        missing = (everything["TITLE_PAGE_INCOMPLETE"].get("details") or [])
        if "TITLE_PAGE_INCOMPLETE" in errors:
            add("Заполни в vkr-project.json topic и title_page: " + ", ".join(missing) + " — пустые поля сборка помечает [ЗАПОЛНИТЬ: …]")
    if any(code.startswith("CLAIM") for code in errors):
        add("Обнови memory/claims-register.json: у каждого утверждения claim_id, text, location, status=confirmed и ссылки на подтверждённые источники или локальные доказательства в evidence/")
    if any(code.startswith("EVIDENCE_") or code == "METHODOLOGY_EVIDENCE_MISSING" for code in errors):
        add("Исправь evidence/index.json: записи {id, path|url}, файлы только в evidence/ или sources/materials/")
    if "FINAL_EXTRA_FILES" in errors:
        add("Оставь в final/ только vkr.docx: другие документы и папки перенеси в exports/")
    if "FINAL_DOCX_MISSING" in errors:
        add("Собери сдаваемый файл: " + _cmd("build_vkr.py", root, "") + "; затем update_docx_fields.py, clean_docx_metadata.py --in-place, validate и snapshot")
    if "FINAL_DOCX_CORRUPT" in errors:
        add(
            "final/vkr.docx повреждён для Word: пересобери " + _cmd("build_vkr.py", root, "")
            + " → update_docx_fields.py → clean_docx_metadata.py --in-place (версия 6.33 с исправлением) → "
            + _cmd("vkr_audit.py", root, "validate") + " и snapshot"
        )
    if "FINAL_DOCX_NAME" in errors:
        add("Пересобери " + _cmd("build_vkr.py", root, "") + " (итог — final/vkr.docx) и убери прежний файл из final/ в exports/; затем update_docx_fields.py, clean_docx_metadata.py --in-place, validate и snapshot")
    if "FINAL_DOCX_MULTIPLE" in errors:
        add("Оставь в final/ только vkr.docx; промежуточные DOCX перенеси в exports/")
    if "FINAL_DOCX_OUTDATED" in everything:
        add(
            "final/vkr.docx не соответствует черновикам: пересобери " + _cmd("build_vkr.py", root, "")
            + "; если правили DOCX в Word — сначала import_docx.py --update; затем update_docx_fields.py, clean_docx_metadata.py --in-place, "
            + _cmd("vkr_audit.py", root, "validate") + " и snapshot"
        )
    if "ANNOTATION_PAGES_MISMATCH" in errors:
        details = errors["ANNOTATION_PAGES_MISMATCH"].get("details") or {}
        add(
            f"Исправь объём в drafts/annotation.md на {details.get('actual')} с. (число страниц final/vkr.docx по Word), затем "
            + _cmd("build_vkr.py", root, "") + " → update_docx_fields.py → clean_docx_metadata.py --in-place → "
            + _cmd("vkr_audit.py", root, "validate") + " и snapshot; если после пересборки число страниц изменилось — повтори "
            "(writing-workflow.md → «Этап 6: Аннотация»)"
        )
    if "DOCX_PAGES_MISMATCH" in errors:
        details = errors["DOCX_PAGES_MISMATCH"].get("details") or {}
        add(
            f"Число страниц в {DOCX_FIELDS_PATH} ({details.get('record_pages')}) не совпадает с самим DOCX "
            f"({details.get('docx_pages')} по docProps/app.xml): перезапиши запись — python \""
            + str(common.SCRIPTS_DIR / "update_docx_fields.py") + "\" \"" + str(root / common.FINAL_DOCX_PATH)
            + "\", затем clean_docx_metadata.py --in-place, " + _cmd("vkr_audit.py", root, "validate") + " и snapshot"
        )
    if "DOCX_FIELDS_UNCONFIRMED" in warnings:
        add(
            "Подтверди число страниц запуском Word: python \"" + str(common.SCRIPTS_DIR / "update_docx_fields.py")
            + "\" \"" + str(root / common.FINAL_DOCX_PATH) + "\" (он же проверит, что Word открывает сдаваемый файл)"
        )
    if "ANNOTATION_PAGES_PARTIAL" in warnings:
        details = warnings["ANNOTATION_PAGES_PARTIAL"].get("details") or {}
        add(
            f"Сверь объём в drafts/annotation.md вручную: указан объём части работы, полный счётчик Word — "
            f"{details.get('actual')} с. (методичка задаёт объём текстовой части, аннотация надёжнее с полным объёмом)"
        )
    elif "ANNOTATION_PAGES_UNVERIFIED" in warnings and ((warnings["ANNOTATION_PAGES_UNVERIFIED"].get("details") or {}).get("reason")):
        add(
            "Сверь объём в аннотации с DOCX: python \"" + str(common.SCRIPTS_DIR / "update_docx_fields.py") + "\" \""
            + str(root / common.FINAL_DOCX_PATH) + "\" (пишет exports/docx-fields.json) и повтори doctor; без Word — сверь число страниц вручную"
        )
    if "VALIDATOR_UNAVAILABLE" in errors:
        add("Установи python-docx (pip install python-docx) и повтори doctor")
    if "VALIDATION_ERRORS" in errors:
        add("Исправь ошибки валидатора в черновиках, пересобери DOCX, затем " + _cmd("vkr_audit.py", root, "validate"))
    if "VALIDATION_WARNINGS" in everything and "VALIDATION_ERRORS" not in errors:
        add("Прочитай предупреждения валидатора в audit/automated-validation.json и по каждому реши: исправить в черновиках (затем пересборка и validate) или обосновать в state")
    if {"AUTOMATED_VALIDATION_MISSING", "AUTOMATED_VALIDATION_STALE", "AUTOMATED_VALIDATION_MISMATCH", "AUTOMATED_VALIDATION_INVALID"} & set(errors):
        add("Перезапиши отчёт валидатора: " + _cmd("vkr_audit.py", root, "validate"))
    if not manifest_broken and not tampered and {"AUDIT_SNAPSHOT_MISSING", "AUDIT_SNAPSHOT_STALE", "AUDIT_SNAPSHOT_HASH_MISMATCH"} & set(everything):
        add("Зафиксируй текущие файлы: " + _cmd("vkr_audit.py", root, "snapshot") + " (волны прежнего снимка не засчитываются)")
    if "AUDIT_RUNS_INCOMPLETE" in errors:
        for run in errors["AUDIT_RUNS_INCOMPLETE"].get("details") or []:
            run_id = run.get("run_id")
            pending = ", ".join(run.get("pending_tasks") or [])
            add(
                f"Заверши {run_id}: запиши отчёты задач {pending} и " + _cmd("vkr_audit.py", root, f"close --run {run_id}")
                + f", или прекрати его: " + _cmd("vkr_audit.py", root, f"close --run {run_id} --abandon --reason \"…\"")
            )
    if "AUDIT_FINDINGS_OPEN" in errors:
        ids = [item.get("finding_id") for item in errors["AUDIT_FINDINGS_OPEN"].get("details") or []]
        add("Исправь замечания " + ", ".join(ids) + " (" + _cmd("vkr_audit.py", root, "findings --open") + "), пересобери DOCX, validate, snapshot и "
            + _cmd("vkr_audit.py", root, "plan --kind targeted_recheck --findings " + ",".join(ids)))
    if "AUDIT_FAIL_UNCLOSED" in errors:
        add(
            "На снимке остался незакрытый fail: исправь его замечания и перепроверь на новом снимке; MINOR без исправления — "
            "accept-risk с обоснованием, ошибочное замечание — invalid_finding с E3/E4 другим аудитором. Повторная волна на том же снимке fail не закрывает"
        )
    prerequisites = {
        "FINAL_DOCX_MISSING", "FINAL_DOCX_MULTIPLE", "FINAL_DOCX_OUTDATED", "FINAL_DOCX_CORRUPT", "FINAL_DOCX_NAME", "FINAL_EXTRA_FILES",
        "VALIDATION_ERRORS", "VALIDATOR_UNAVAILABLE", "ANNOTATION_PAGES_MISMATCH", "DOCX_PAGES_MISMATCH",
        "USER_EVIDENCE_REQUIRED", "PLACEHOLDER_FOUND", "DRAFT_REQUIRED_MISSING", "DRAFT_TOO_SHORT", "AUDIT_FINDINGS_OPEN",
        "AUDIT_FAIL_UNCLOSED", "AUDIT_RUN_TAMPERED", "AUDIT_SEQ_GAP", "AUDIT_REPORT_ORPHAN", "AUDIT_REPORT_MISMATCH",
    } & set(errors)
    if {"AUDIT_PRIMARY_WAVE_MISSING", "AUDIT_BLIND_WAVE_MISSING"} & set(errors) and prerequisites:
        add("После исправлений выше — финальные волны аудита на новом снимке (plan --kind primary --preset final, затем blind_regression)")
    elif "AUDIT_PRIMARY_WAVE_MISSING" in errors:
        add("Проведи финальную primary-волну: " + _cmd("vkr_audit.py", root, "plan --kind primary --preset final") + " → brief → record → close")
    elif "AUDIT_BLIND_WAVE_MISSING" in errors:
        add("Проведи слепую регрессию новыми аудиторами: " + _cmd("vkr_audit.py", root, "plan --kind blind_regression --preset final") + " → brief → record → close")
    if "MEMORY_NEEDS_REBASELINE" in warnings:
        add("Переиндексируй память: " + _cmd("vkr_memory.py", root, "record --rebaseline --json"))
    elif {"ARTIFACT_CHANGED", "ARTIFACT_MISSING"} & set(warnings):
        add("Запиши checkpoint памяти: " + _cmd("vkr_memory.py", root, "record --event <event.json> --all-changed --json"))
    if "AUDIT_INTENSITY_LOWERED" in warnings:
        details = warnings["AUDIT_INTENSITY_LOWERED"].get("details") or {}
        add(
            f"Финальный аудит проведён с audit_intensity={details.get('applied')} ниже прежней ({details.get('earlier_max')}): "
            "если снижение не согласовано с пользователем, верни прежнее значение в vkr-project.json, пересобери DOCX и повтори финальные волны"
        )
    if ctx.final and ctx.degraded and not errors:
        add("Для READY_TO_SUBMIT повтори обе финальные волны в отдельных чистых контекстах (--independence isolated_contexts) или получи внешнюю проверку руководителя")
    if not errors:
        if ctx.stage == "draft":
            add("Структура в порядке: продолжай по плану; перед передачей руководителю запусти doctor --stage prefinal")
        elif ctx.stage == "prefinal" and "AUDIT_PREFINAL_WAVE_MISSING" in warnings:
            add(
                "Ошибок нет, но матрица strict/maximum перед передачей руководителю предусматривает волну девяти ролей: "
                + _cmd("vkr_audit.py", root, "plan --kind primary --preset prefinal --json")
                + " → brief → record → close, затем снова doctor --stage prefinal. Передавать без неё — только по "
                "осознанному решению пользователя, записанному в state"
            )
        elif ctx.stage == "prefinal":
            add("Можно передавать руководителю; перед сдачей — финальные волны и doctor --stage final")
        elif not ctx.degraded:
            add("Готово к сдаче: final/vkr.docx прошёл финальный gate")
    return actions


def diagnose(root: Path, stage: str) -> Dict[str, Any]:
    if stage not in STAGE_ORDER:
        raise ValueError(f"Неизвестная стадия: {stage}")
    resolved = Path(os.path.abspath(os.path.expanduser(str(root))))
    if not resolved.is_dir():
        raise FileNotFoundError(f"Каталог проекта не найден: {resolved}")
    ctx = Context(resolved.resolve(), stage)
    for name, check in CHECKS:
        try:
            check(ctx)
        except Exception as error:  # находка вместо трассировки; остальные проверки продолжаются
            ctx.add("INTERNAL_ERROR", "ERROR", name, f"Проверка {name} прервана: {type(error).__name__}: {error}")
    counts = {severity: sum(1 for item in ctx.findings if item["severity"] == severity) for severity in ("ERROR", "WARNING", "INFO")}
    status = "FAIL" if counts["ERROR"] else "PASS"
    if status == "PASS" and stage == "final":
        readiness = "READY_FOR_SUPERVISOR_REVIEW" if ctx.degraded else "READY_TO_SUBMIT"
    elif status == "PASS" and stage == "prefinal":
        readiness = "READY_FOR_SUPERVISOR_REVIEW"
    else:
        readiness = "NOT_READY"
    registry = ctx.registry
    return {
        "status": status,
        "stage": stage,
        "readiness": readiness,
        "doctor_version": DOCTOR_VERSION,
        "project_root": str(ctx.root),
        "counts": counts,
        "audit": {
            "current_snapshot_sha256": ctx.current_sha,
            "recorded_snapshot_sha256": (ctx.data.get("audit/snapshot-inputs.json") or {}).get("snapshot_manifest_sha256")
            if isinstance(ctx.data.get("audit/snapshot-inputs.json"), dict) else None,
            "primary_run": ctx.waves[0].get("run_id") if ctx.waves else None,
            "blind_run": ctx.waves[1].get("run_id") if ctx.waves else None,
            "degraded_independence": ctx.degraded,
            "final_wave_matrix": ctx.final_matrix,
            "open_findings": sorted(item["finding_id"] for item in registry.values() if item.get("status") == "open"),
        },
        "findings": ctx.findings,
        "next_actions": build_next_actions(ctx),
    }


def main(argv: Optional[List[str]] = None) -> int:
    common.reconfigure_stdio()
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("project_dir", type=Path)
    parser.add_argument("--stage", choices=tuple(STAGE_ORDER), default="draft")
    parser.add_argument("--json", action="store_true", dest="json_output")
    parser.add_argument("-o", "--output", type=Path, help="Записать JSON-отчёт в файл UTF-8")
    args = parser.parse_args(argv)
    try:
        result = diagnose(args.project_dir, args.stage)
        code = 1 if result["status"] == "FAIL" else 0
    except Exception as error:
        result = {"status": "ERROR", "stage": args.stage, "readiness": "NOT_READY", "error": f"{type(error).__name__}: {error}", "findings": [], "next_actions": []}
        code = 2
    if args.output:
        common.emit_json(result, args.output)
    elif args.json_output:
        common.emit_json(result)
    else:
        print(f"{result['status']} ({result.get('readiness')}): {result.get('counts', result.get('error'))}")
        for item in result.get("findings", []):
            print(f"{item['severity']} {item['code']} {item['path']}: {item['message']}")
        if result.get("next_actions"):
            print("Дальше:")
            for action in result["next_actions"]:
                print(f"  - {action}")
    return code


if __name__ == "__main__":
    raise SystemExit(main())
