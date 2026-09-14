"""Ядро аудита 6.33 (WP1): init → черновики → build_vkr → vkr_audit → doctor final.

Тесты не ходят в сеть и не пишут в каталог скилла. Повторный запуск валидатора
в doctor подменяется in-process (``vkr_common.run_validator``): сам валидатор
проверяется своими тестами. Черновики берутся из ``tests/fixtures/build/project``
и собираются настоящим ``build_vkr.py``.

Запуск: ``python -m unittest discover -s tests -t .`` или ``python tests/test_core_audit_633.py``.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unittest
import zipfile
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Dict, Iterator, List, Optional, Tuple
from unittest import mock

os.environ["PYTHONDONTWRITEBYTECODE"] = "1"
sys.dont_write_bytecode = True

ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "vkr-mpgu"
SCRIPTS = SKILL / "scripts"
FIXTURE = ROOT / "tests" / "fixtures" / "build" / "project"
WORD_SAVED = ROOT / "tests" / "fixtures" / "build" / "vkr-word-saved.docx"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

import vkr_common as common  # noqa: E402
import vkr_audit as audit  # noqa: E402
import vkr_project_doctor as doctor  # noqa: E402
import vkr_memory as memory  # noqa: E402
import init_vkr_project as init_project  # noqa: E402

try:
    import docx  # noqa: F401
    HAVE_DOCX = True
except ImportError:  # pragma: no cover - окружение без python-docx
    HAVE_DOCX = False

BOM = chr(0xFEFF)
ZWSP = chr(0x200B)
CYRILLIC_A = chr(0x0410)
EXTRA = (
    "\n\nПолученные результаты согласуются с поставленной целью: учащиеся, работавшие с тренажёром, "
    "увереннее выполняли построение сечений и чаще доводили решение до конца. Ограничения исследования "
    "связаны с небольшой выборкой и коротким сроком пилотирования.\n"
)
MAJOR_FINDING = {
    "severity": "MAJOR",
    "location": {"section": "1.1", "quote": "Пространственное мышление рассматривается"},
    "criterion": "Источник должен подтверждать тезис",
    "rule_source": "ГОСТ Р 7.0.5-2008; методичка МПГУ",
    "observed": "Ссылка на с. 5 не подтверждает определение",
    "expected": "Тезис опирается на страницу, где дано определение",
    "evidence_level": "E3",
    "evidence": ["Каталог РГБ: запись 42, с. 5 — оглавление"],
    "proposed_fix": "Уточнить страницу или заменить источник",
    "fix_class": "AUTO_CONTENT",
    "needs_user_fact": False,
}
MINOR_FINDING = dict(MAJOR_FINDING, severity="MINOR", evidence_level="E2", fix_class="MANUAL_WORD")
PASS_REPORT = {"status": "pass", "summary": "Замечаний по области не найдено", "not_applicable_reason": "", "findings": []}
FAIL_REPORT = {"status": "fail", "summary": "Тезис не подтверждён", "findings": [MAJOR_FINDING]}


def write_json(path: Path, value: Any, bom: bool = False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text((BOM if bom else "") + json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def run_cli(script: str, *args: Any, cwd: Optional[Path] = None) -> Tuple[int, Any, str, str]:
    env = os.environ.copy()
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    result = subprocess.run(
        [sys.executable, str(SCRIPTS / script), *map(str, args)],
        cwd=str(cwd or ROOT), env=env, capture_output=True, text=True, encoding="utf-8",
    )
    try:
        payload = json.loads(result.stdout) if result.stdout.strip() else None
    except json.JSONDecodeError:
        payload = None
    return result.returncode, payload, result.stdout, result.stderr


def fake_validator(errors: int = 0, available: bool = True, warnings: int = 0):
    def run(docx_path: Path, profile: str) -> Dict[str, Any]:
        if not available:
            return {"available": False, "report": None, "error": "python-docx не установлен"}
        checks = {"Структура работы": [{"ok": errors == 0, "severity": "error", "message": "Нет раздела «Введение»" if errors else "Все разделы на месте"}]}
        if warnings:
            checks["Содержимое аннотации"] = [
                {"ok": False, "severity": "warning", "message": f"В аннотации нет поля {index}"} for index in range(1, warnings + 1)
            ]
        return {
            "available": True,
            "error": None,
            "report": {
                "path": str(docx_path), "profile": profile, "validator_version": "6.33",
                "document_sha256": common.sha256_file(Path(docx_path)),
                "summary": {"passed": 20, "warnings": warnings, "errors": errors}, "checks": checks,
            },
        }

    return run


@contextmanager
def patched_validator(errors: int = 0, available: bool = True) -> Iterator[None]:
    with mock.patch.object(common, "run_validator", fake_validator(errors, available)):
        yield


def audit_cmd(project: Path, *args: Any) -> Tuple[int, Dict[str, Any]]:
    code, payload, _args = audit.execute([str(project), *map(str, args)])
    return code, payload


def codes(report: Dict[str, Any], severity: Optional[str] = None) -> List[str]:
    return [item["code"] for item in report["findings"] if severity is None or item["severity"] == severity]


def finding_of(report: Dict[str, Any], code: str) -> Dict[str, Any]:
    return next(item for item in report["findings"] if item["code"] == code)


def diagnose(project: Path, stage: str = "final", errors: int = 0, available: bool = True) -> Dict[str, Any]:
    with patched_validator(errors, available):
        return doctor.diagnose(project, stage)


def intake(profile: str = "mpgu-09-project", intensity: str = "strict", **extra: Any) -> Dict[str, Any]:
    source = read_json(FIXTURE / "vkr-project.json")
    config = {key: source[key] for key in ("topic", "mode", "deadline", "collective", "members", "title_page")}
    config.update({"profile": profile, "audit_intensity": intensity})
    config.update(extra)
    return config


def create_project(base: Path, name: str = "project", **config: Any) -> Path:
    project = base / name
    config_path = base / f"{name}-intake.json"
    write_json(config_path, intake(**config))
    result = init_project.initialize(project, init_project.normalize_config(init_project.load_config(config_path)))
    assert result["status"] == "initialized", result
    return project


def rebuild(project: Path) -> None:
    import build_vkr  # noqa: WPS433

    code, result = build_vkr.run_build(project)
    assert code == 0, result


def fill_project(project: Path, *, product: bool = True, defense: bool = True, regular: bool = False, build: bool = True) -> None:
    for path in (FIXTURE / "drafts").glob("*.md"):
        shutil.copyfile(str(path), str(project / "drafts" / path.name))
    if regular:
        (project / "drafts" / "chapter-3.md").unlink()
    else:
        chapter = project / "drafts" / "chapter-3.md"
        chapter.write_text(
            chapter.read_text(encoding="utf-8").replace(
                "Результаты пилотирования будут показаны на рисунке 2.\n\n![Экран результатов пилотирования]()\n",
                "Результаты пилотирования приведены в разделе 3.2.\n",
            ),
            encoding="utf-8",
        )
    conclusion = project / "drafts" / "conclusion.md"
    conclusion.write_text(conclusion.read_text(encoding="utf-8").rstrip("\n") + EXTRA, encoding="utf-8")
    (project / "evidence" / "figures").mkdir(parents=True, exist_ok=True)
    shutil.copyfile(str(FIXTURE / "evidence" / "figures" / "architecture.png"), str(project / "evidence" / "figures" / "architecture.png"))
    (project / "plan.md").write_text(
        "# Рабочий план ВКР\n\n## Введение\n\nАктуальность, цель, задачи, методы.\n\n## Глава I. Теория\n\n## Глава II. Проектирование\n\n"
        "## Заключение\n\n## Список использованной литературы\n", encoding="utf-8",
    )
    sources = read_json(FIXTURE / "sources.json")
    for position, source in enumerate(sources, start=1):
        source["status"] = "confirmed"
        source["verification"] = {"checked_at": "2026-09-01T10:00:00+03:00", "method": "catalog", "url": f"https://search.rsl.ru/ru/record/{position:05d}", "note": "сверено"}
    write_json(project / "sources.json", sources)
    index = read_json(project / "evidence" / "index.json")
    index["figures"] = [{"id": "architecture", "path": "evidence/figures/architecture.png"}]
    if product:
        (project / "evidence" / "product" / "repo-readme.md").write_text("Описание функций тренажёра и скриншоты.\n", encoding="utf-8")
        index["product"] = [{"id": "product-readme", "path": "evidence/product/repo-readme.md"}]
    if defense:
        (project / "evidence" / "defense" / "test30.md").write_text("Test-30: ответы студента на вопросы по работе.\n", encoding="utf-8")
        index["defense"] = [{"id": "defense-test30", "path": "evidence/defense/test30.md"}]
    write_json(project / "evidence" / "index.json", index)
    write_json(project / "memory" / "claims-register.json", [{
        "claim_id": "CLM-001", "text": "Пилотирование с участием 20 учащихся", "location": "Глава III, 3.2",
        "status": "confirmed", "source_ids": ["yolkin2021"], "evidence_ids": ["product-readme"] if product else [],
    }])
    if build:
        rebuild(project)


def validate_and_snapshot(project: Path, errors: int = 0) -> str:
    with patched_validator(errors):
        code, payload = audit_cmd(project, "validate", "--json")
    assert code == (0 if errors == 0 else 1), payload
    code, payload = audit_cmd(project, "snapshot", "--json")
    assert code == 0, payload
    return payload["snapshot_manifest_sha256"]


def report_file(base: Path, name: str, report: Dict[str, Any], bom: bool = False) -> Path:
    path = base / "auditor-output" / f"{name}.json"
    write_json(path, report, bom=bom)
    return path


def run_wave(
    project: Path, kind: str, preset: str, prefix: str, *,
    reports: Optional[Dict[str, Dict[str, Any]]] = None, extra: Tuple[str, ...] = (),
    auditor: Optional[str] = None, expect_close: Optional[int] = None,
) -> Dict[str, Any]:
    code, plan = audit_cmd(project, "plan", "--kind", kind, "--preset", preset, "--json", *extra)
    assert code == 0, plan
    code, brief = audit_cmd(project, "brief", "--run", plan["run_id"], "--json")
    assert code == 0, brief
    reports = reports or {}
    for task in plan["tasks"]:
        role = task["auditor_role"]
        report = reports.get(f"{role}-{task['replica_id']}", reports.get(role, PASS_REPORT))
        path = report_file(project.parent, f"{project.name}-{plan['run_id']}-{task['task_id']}", report)
        auditor_id = auditor or f"{prefix}-{role}-{task['replica_id']}"
        code, recorded = audit_cmd(project, "record", "--run", plan["run_id"], "--task", task["task_id"], "--auditor-id", auditor_id, "--report", path, "--json")
        assert code == 0, recorded
    code, closed = audit_cmd(project, "close", "--run", plan["run_id"], "--json")
    if expect_close is not None:
        assert code == expect_close, closed
    plan["close"] = closed
    return plan


def fix_and_rebuild(project: Path, marker: str = "Исправлено") -> str:
    chapter = project / "drafts" / "chapter-1.md"
    chapter.write_text(chapter.read_text(encoding="utf-8").replace("[@egorov2020, с. 5]", "[@egorov2020, с. 17]") + f"\n{marker}.\n", encoding="utf-8")
    rebuild(project)
    return validate_and_snapshot(project)


def recheck(project: Path, auditor_id: str, verdict: str, *, level: Optional[str] = None, expect: int = 0, retry: Optional[str] = None, finding: str = "F-SRC-001") -> Dict[str, Any]:
    code, plan = audit_cmd(project, "plan", "--kind", "targeted_recheck", "--findings", finding, "--json")
    assert code == 0, plan
    return record_recheck(project, plan, auditor_id, verdict, level=level, expect=expect, retry=retry, finding=finding)


def record_recheck(project: Path, plan: Dict[str, Any], auditor_id: str, verdict: str, *, level: Optional[str] = None, expect: int = 0, retry: Optional[str] = None, finding: str = "F-SRC-001") -> Dict[str, Any]:
    entry = {"finding_id": finding, "verdict": verdict, "evidence": ["Глава I: проверено по оригиналу"]}
    if level:
        entry["evidence_level"] = level
    status = "pass" if verdict in ("pass", "invalid_finding") else verdict
    report = {"status": status, "summary": "перепроверка", "findings": [], "rechecks": [entry]}
    path = report_file(project.parent, f"{project.name}-{plan['run_id']}-{auditor_id}-{verdict}", report)
    args = ["record", "--run", plan["run_id"], "--task", plan["tasks"][0]["task_id"], "--auditor-id", auditor_id, "--report", path, "--json"]
    if retry:
        args += ["--retry-reason", retry]
    code, payload = audit_cmd(project, *args)
    assert code == expect, payload
    payload["plan"] = plan
    return payload


def final_waves(project: Path, prefix: str, **kwargs: Any) -> None:
    run_wave(project, "primary", "final", f"{prefix}P", expect_close=0, **kwargs)
    run_wave(project, "blind_regression", "final", f"{prefix}B", expect_close=0)


def rewrite_docx(source: Path, target: Path, edits: Optional[Dict[str, Any]] = None, extra: Optional[Dict[str, str]] = None) -> Path:
    """Копия DOCX с правками частей: edits — {часть: функция(текст XML или bytes)}, extra — новые части."""
    edits = edits or {}
    with zipfile.ZipFile(str(source)) as src, zipfile.ZipFile(str(target), "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            edit = edits.get(item.filename)
            if edit is not None:
                is_xml = item.filename.endswith((".xml", ".rels"))
                result = edit(data.decode("utf-8") if is_xml else data)
                data = result.encode("utf-8") if isinstance(result, str) else result
            dst.writestr(item, data)
        for name, text in (extra or {}).items():
            dst.writestr(name, text)
    return target


def before_body_end(xml: str, fragment: str) -> str:
    position = xml.rindex("<w:sectPr")
    return xml[:position] + fragment + xml[position:]


def snapshot_only(project: Path) -> str:
    code, payload = audit_cmd(project, "snapshot", "--json")
    assert code == 0, payload
    return payload["snapshot_manifest_sha256"]


class ProjectCase(unittest.TestCase):
    """Базовые состояния строятся один раз на класс и копируются в каждый тест."""

    state = "done"
    config: Dict[str, Any] = {}

    @classmethod
    def build_state(cls, project: Path) -> None:
        fill_project(project)
        validate_and_snapshot(project)
        if cls.state == "done":
            final_waves(project, "B", extra=("--roles", "OWN-DEFENSE"))
        elif cls.state == "failed":
            run_wave(project, "primary", "final", "P1", reports={"SRC-A": FAIL_REPORT}, expect_close=1)

    @classmethod
    def setUpClass(cls) -> None:
        if not HAVE_DOCX:
            raise unittest.SkipTest("python-docx не установлен")
        cls.class_dir = tempfile.TemporaryDirectory(prefix=f"vkr633-{cls.state}-")
        cls.template = create_project(Path(cls.class_dir.name), **cls.config)
        cls.build_state(cls.template)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.class_dir.cleanup()

    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="vkr633-case-")
        self.base = Path(self.temp.name)
        self.project = self.base / "project"
        shutil.copytree(str(self.template), str(self.project))

    def tearDown(self) -> None:
        self.temp.cleanup()

    def manifest(self) -> Dict[str, Any]:
        return read_json(self.project / "audit" / "manifest.json")

    def save_manifest(self, manifest: Dict[str, Any]) -> None:
        write_json(self.project / "audit" / "manifest.json", manifest)


@unittest.skipUnless(HAVE_DOCX, "python-docx не установлен")
class EndToEndCliTest(unittest.TestCase):
    def test_full_cycle_through_cli_reaches_ready_to_submit(self) -> None:
        with tempfile.TemporaryDirectory(prefix="vkr633-e2e-") as temporary:
            base = Path(temporary)
            project = base / "project"
            write_json(base / "intake.json", intake(intensity="balanced"))
            code, init, _out, err = run_cli("init_vkr_project.py", project, "--config", base / "intake.json", "--json")
            self.assertEqual(0, code, err)
            self.assertTrue((project / common.METHODOLOGY_PROJECT_PATH).is_file())
            code, draft, _out, _err = run_cli("vkr_project_doctor.py", project, "--stage", "draft", "--json")
            self.assertEqual("PASS", draft["status"], draft)

            fill_project(project, build=False)
            code, built, _out, err = run_cli("build_vkr.py", project, "--json")
            self.assertEqual(0, code, (built, err))
            self.assertEqual("exports/build-manifest.json", built["outputs"]["build_manifest"])
            with patched_validator():
                self.assertEqual(0, audit_cmd(project, "validate", "--json")[0])
            code, snap, _out, err = run_cli("vkr_audit.py", project, "snapshot", "--json")
            self.assertEqual(0, code, err)
            self.assertIn("exports/build-manifest.json", [item["logical_id"] for item in read_json(project / "audit" / "snapshot-inputs.json")["inputs"]])

            code, plan, _out, err = run_cli("vkr_audit.py", project, "plan", "--kind", "primary", "--preset", "final", "--json")
            self.assertEqual(0, code, (plan, err))
            self.assertEqual(9, len(plan["tasks"]))
            run_id = plan["run_id"]
            code, brief, _out, _err = run_cli("vkr_audit.py", project, "brief", "--run", run_id, "--json")
            own_brief = next(Path(item["absolute_path"]) for item in brief["briefs"] if item["auditor_role"] == "OWN").read_text(encoding="utf-8")
            self.assertIn("evidence/defense/", own_brief)
            for task in plan["tasks"]:
                report = FAIL_REPORT if task["auditor_role"] == "SRC" else PASS_REPORT
                path = report_file(base, task["task_id"], report, bom=task["auditor_role"] == "LNG")
                code, recorded, _out, err = run_cli(
                    "vkr_audit.py", project, "record", "--run", run_id, "--task", task["task_id"],
                    "--auditor-id", f"P1-{task['auditor_role']}", "--report", path, "--json",
                )
                self.assertEqual(0, code, (recorded, err))
            code, closed, _out, _err = run_cli("vkr_audit.py", project, "close", "--run", run_id, "--json")
            self.assertEqual(1, code, closed)
            self.assertEqual(["F-SRC-001"], closed["open_high_findings"])
            self.assertIn("AUDIT_FINDINGS_OPEN", codes(diagnose(project), "ERROR"))

            chapter = project / "drafts" / "chapter-1.md"
            chapter.write_text(chapter.read_text(encoding="utf-8").replace("[@egorov2020, с. 5]", "[@egorov2020, с. 17]"), encoding="utf-8")
            self.assertIn("FINAL_DOCX_OUTDATED", codes(diagnose(project, "prefinal"), "ERROR"))  # B1: и на prefinal
            code, built, _out, err = run_cli("build_vkr.py", project, "--json")
            self.assertEqual(0, code, err)
            with patched_validator():
                self.assertEqual(0, audit_cmd(project, "validate", "--json")[0])
            code, snap2, _out, _err = run_cli("vkr_audit.py", project, "snapshot", "--json")
            self.assertTrue(snap2["created"])
            self.assertIn("drafts/chapter-1.md", snap2["changes"]["changed"])

            code, rplan, _out, err = run_cli("vkr_audit.py", project, "plan", "--kind", "targeted_recheck", "--findings", "F-SRC-001", "--json")
            self.assertEqual(0, code, (rplan, err))
            code, rbrief, _out, _err = run_cli("vkr_audit.py", project, "brief", "--run", rplan["run_id"], "--json")
            self.assertIn(MAJOR_FINDING["observed"], Path(rbrief["briefs"][0]["absolute_path"]).read_text(encoding="utf-8"))
            rreport = report_file(base, "recheck", {
                "status": "pass", "summary": "Замечание устранено", "findings": [],
                "rechecks": [{"finding_id": "F-SRC-001", "verdict": "pass", "evidence": ["Глава I: [@egorov2020, с. 17]"]}],
            })
            task_id = rplan["tasks"][0]["task_id"]
            code, denied, _out, _err = run_cli("vkr_audit.py", project, "record", "--run", rplan["run_id"], "--task", task_id, "--auditor-id", "p1-src", "--report", rreport, "--json")
            self.assertEqual("AUDITOR_NOT_INDEPENDENT", denied["code"])
            code, recorded, _out, err = run_cli("vkr_audit.py", project, "record", "--run", rplan["run_id"], "--task", task_id, "--auditor-id", "TR-SRC", "--report", rreport, "--json")
            self.assertEqual("resolved", recorded["rechecks"][0]["finding_status"])
            self.assertEqual(0, run_cli("vkr_audit.py", project, "close", "--run", rplan["run_id"], "--json")[0])

            for kind, prefix in (("primary", "P2"), ("blind_regression", "B2")):
                code, wave, _out, err = run_cli("vkr_audit.py", project, "plan", "--kind", kind, "--preset", "final", "--json")
                self.assertEqual(0, code, (wave, err))
                code, briefs, _out, _err = run_cli("vkr_audit.py", project, "brief", "--run", wave["run_id"], "--json")
                if kind == "blind_regression":
                    for item in briefs["briefs"]:
                        text = Path(item["absolute_path"]).read_text(encoding="utf-8")
                        self.assertFalse("F-SRC-" in text or "rechecks" in text, item["task_id"])
                for task in wave["tasks"]:
                    path = report_file(base, task["task_id"], PASS_REPORT)
                    code, recorded, _out, err = run_cli(
                        "vkr_audit.py", project, "record", "--run", wave["run_id"], "--task", task["task_id"],
                        "--auditor-id", f"{prefix}-{task['auditor_role']}", "--report", path, "--json",
                    )
                    self.assertEqual(0, code, (recorded, err))
                self.assertEqual(0, run_cli("vkr_audit.py", project, "close", "--run", wave["run_id"], "--json")[0])

            final = diagnose(project, "final")
            self.assertEqual("PASS", final["status"], [item for item in final["findings"] if item["severity"] == "ERROR"])
            self.assertEqual("READY_TO_SUBMIT", final["readiness"])
            self.assertTrue((project / "audit" / "manifest.json.bak").is_file())
            self.assertFalse(list(SKILL.rglob("__pycache__")))


class FinalGateTest(ProjectCase):
    state = "done"

    def test_baseline_passes_and_copy_stays_valid(self) -> None:
        report = diagnose(self.project)
        self.assertEqual("PASS", report["status"], [item for item in report["findings"] if item["severity"] == "ERROR"])
        self.assertEqual("READY_TO_SUBMIT", report["readiness"])

    def test_strict_matrix_with_b_replicas_and_specialized_role(self) -> None:
        primary = next(run for run in self.manifest()["runs"] if run["kind"] == "primary")
        roles = {(task["auditor_role"], task["replica_id"]) for task in primary["planned_tasks"]}
        self.assertIn(("LOG", "B"), roles)
        self.assertIn(("OWN-DEFENSE", "A"), roles)
        self.assertEqual(16, len(primary["planned_tasks"]))

    def test_doctor_requires_matrix_replicas(self) -> None:
        with mock.patch.object(common, "project_intensity", return_value="maximum"):
            report = diagnose(self.project)
        details = finding_of(report, "AUDIT_PRIMARY_WAVE_MISSING")["details"]
        self.assertTrue(any("не хватает пройденных реплик" in reason for reasons in details.values() for reason in reasons), details)

    def test_docx_changed_after_snapshot_fails(self) -> None:
        shutil.copyfile(str(WORD_SAVED), str(self.project / "final" / "vkr.docx"))
        report = diagnose(self.project)
        errors = codes(report, "ERROR")
        self.assertIn("AUDIT_SNAPSHOT_STALE", errors)
        self.assertIn("AUDIT_PRIMARY_WAVE_MISSING", errors)
        self.assertIn("final/vkr.docx", finding_of(report, "AUDIT_SNAPSHOT_STALE")["details"]["changed"])

    def test_draft_edit_after_build_makes_docx_outdated(self) -> None:
        chapter = self.project / "drafts" / "chapter-2.md"
        chapter.write_text(chapter.read_text(encoding="utf-8") + "\nНовый абзац после сборки.\n", encoding="utf-8")
        report = diagnose(self.project)
        self.assertIn("FINAL_DOCX_OUTDATED", codes(report, "ERROR"))
        self.assertTrue(any("build_vkr.py" in action and "import_docx.py --update" in action for action in report["next_actions"]))

    def test_state_change_only_in_service_sections_keeps_snapshot(self) -> None:
        state = self.project / "vkr-state.md"
        text = state.read_text(encoding="utf-8")
        text = text.replace("- **Итог:** NOT_READY", "- **Итог:** READY_TO_SUBMIT")
        text = text.replace("## История сессий\n", "## История сессий\n\n- финальный аудит пройден\n")
        text = text.replace("> Последнее обновление:", "> Последнее обновление: 2026-09-20 (было")
        state.write_text(text, encoding="utf-8")
        report = diagnose(self.project)
        self.assertEqual("PASS", report["status"], codes(report, "ERROR"))

    def test_section_title_containing_audit_word_is_content(self) -> None:
        state = self.project / "vkr-state.md"
        state.write_text(state.read_text(encoding="utf-8").replace("## Ключевые решения\n", "## Целевая аудитория продукта\n\n- 10–11 классы.\n\n## Ключевые решения\n"), encoding="utf-8")
        self.assertIn("AUDIT_SNAPSHOT_STALE", codes(diagnose(self.project), "ERROR"))

    def test_claims_change_makes_snapshot_stale(self) -> None:
        claims_path = self.project / "memory" / "claims-register.json"
        claims = read_json(claims_path)
        claims[0]["text"] += " (уточнено)"
        write_json(claims_path, claims)
        self.assertIn("memory/claims-register.json", finding_of(diagnose(self.project), "AUDIT_SNAPSHOT_STALE")["details"]["changed"])

    def test_forged_validation_report_with_real_errors_fails(self) -> None:
        report = diagnose(self.project, errors=3)
        self.assertIn("VALIDATION_ERRORS", codes(report, "ERROR"))
        self.assertIn("AUTOMATED_VALIDATION_MISMATCH", codes(report, "ERROR"))
        self.assertIn("VALIDATOR_UNAVAILABLE", codes(diagnose(self.project, available=False), "ERROR"))

    def test_final_folder_contents(self) -> None:
        (self.project / "final" / "~$vkr.docx").write_bytes(b"lock")
        self.assertEqual("PASS", diagnose(self.project)["status"])
        shutil.copyfile(str(self.project / "final" / "vkr.docx"), str(self.project / "final" / "~$vkr-real.docx"))
        self.assertIn("FINAL_DOCX_MULTIPLE", codes(diagnose(self.project), "ERROR"))
        (self.project / "final" / "~$vkr-real.docx").unlink()
        (self.project / "final" / "old").mkdir()
        shutil.copyfile(str(self.project / "final" / "vkr.docx"), str(self.project / "final" / "vkr.pdf"))
        report = diagnose(self.project)
        self.assertEqual(["final/old/", "final/vkr.pdf"], sorted(finding_of(report, "FINAL_EXTRA_FILES")["details"]))
        shutil.copyfile(str(self.project / "final" / "vkr.docx"), str(self.project / "final" / "vkr-copy.docx"))
        with patched_validator():
            code, payload = audit_cmd(self.project, "validate", "--json")
        self.assertEqual("FINAL_DOCX_MULTIPLE", payload["code"])

    def test_report_tampering_with_recomputed_sha_is_detected(self) -> None:
        manifest = self.manifest()
        task = manifest["runs"][0]["planned_tasks"][0]
        attempt = task["attempts"][-1]
        path = self.project / task["report_path"]
        canonical = read_json(path)
        canonical["task_status"] = "fail"
        write_json(path, canonical)
        sha = common.sha256_file(path)
        attempt["report_sha256"] = sha
        task["report_sha256"] = sha
        self.save_manifest(manifest)
        report = diagnose(self.project)
        self.assertIn("AUDIT_REPORT_MISMATCH", codes(report, "ERROR"))

    def test_independence_flip_in_manifest_is_detected(self) -> None:
        manifest = self.manifest()
        for run in manifest["runs"]:
            run["independence"] = "degraded_independence"
        self.save_manifest(manifest)
        report = diagnose(self.project)
        self.assertIn("AUDIT_REPORT_MISMATCH", codes(report, "ERROR"))
        self.assertNotEqual("READY_TO_SUBMIT", report["readiness"])

    def test_kind_swap_and_seq_edit_are_detected(self) -> None:
        manifest = self.manifest()
        blind = next(run for run in manifest["runs"] if run["kind"] == "blind_regression")
        blind["planned_seq"] = 1
        self.save_manifest(manifest)
        errors = codes(diagnose(self.project), "ERROR")
        self.assertIn("AUDIT_SEQ_GAP", errors)
        self.assertIn("AUDIT_RUN_TAMPERED", errors)

    def test_orphan_report_and_null_run_are_errors(self) -> None:
        write_json(self.project / "audit" / "reports" / "R099-primary-final" / "R099-SRC-A.json", {"report_schema": "vkr-audit-report/6.33", "status": "fail"})
        manifest = self.manifest()
        manifest["runs"].append(None)
        self.save_manifest(manifest)
        errors = codes(diagnose(self.project), "ERROR")
        self.assertIn("AUDIT_REPORT_ORPHAN", errors)
        self.assertIn("AUDIT_RUN_TAMPERED", errors)

    def test_manual_waive_decision_is_validated(self) -> None:
        manifest = self.manifest()
        seq = manifest["seq"] + 1
        manifest["seq"] = seq
        manifest["decisions"].append({"decision_id": f"D-{seq:05d}", "finding_id": "F-XXX-001", "action": "waive", "approval": "vkr-project.json", "approval_sha256": "0" * 64, "at": "x", "seq": seq})
        self.save_manifest(manifest)
        self.assertIn("AUDIT_DECISION_INVALID", codes(diagnose(self.project), "ERROR"))

    def test_locator_to_frozen_copy_does_not_hide_changes(self) -> None:
        inputs_path = self.project / "audit" / "snapshot-inputs.json"
        snapshot = read_json(inputs_path)
        for item in snapshot["inputs"]:
            if item["logical_id"] == "plan.md":
                item["locator"] = "exports/plan-frozen.md"
        write_json(inputs_path, snapshot)
        (self.project / "plan.md").write_text("# План изменён после аудита\n", encoding="utf-8")
        self.assertIn("plan.md", finding_of(diagnose(self.project), "AUDIT_SNAPSHOT_STALE")["details"]["changed"])

    def test_iso_timestamps_and_bom(self) -> None:
        self.assertEqual(123456, common.parse_iso8601("2026-09-13T12:00:00.1234567+03:00", require_timezone=True).microsecond)
        self.assertIsNotNone(common.parse_iso8601("2026-09-13T12:00:00+0300", require_timezone=True))
        self.assertIsNone(common.parse_iso8601("2026-09-13T12:00:00", require_timezone=True))
        for relative in ("sources.json", "memory/claims-register.json", "vkr-project.json"):
            path = self.project / relative
            path.write_bytes(BOM.encode("utf-8") + path.read_bytes())
        report = diagnose(self.project, "prefinal")
        self.assertNotIn("JSON_INVALID", codes(report))
        self.assertNotIn("INTERNAL_ERROR", codes(report))

    def test_weak_claim_evidence_is_rejected(self) -> None:
        index = read_json(self.project / "evidence" / "index.json")
        index["pilot"] = [{"id": "pilot-survey", "url": "https://forms.example.org/survey"}]
        write_json(self.project / "evidence" / "index.json", index)
        claims = read_json(self.project / "memory" / "claims-register.json")
        claims.append({"claim_id": "CLM-002", "text": "Успеваемость выросла", "location": "3.2", "status": "confirmed", "evidence_ids": ["pilot-survey"]})
        claims.append({"claim_id": "CLM-003", "text": "Самоссылка", "location": "3.3", "status": "confirmed", "evidence": ["vkr-state.md"]})
        claims.append({"claim_id": "CLM-004", "text": "Источник вместо доказательства", "location": "3.3", "status": "confirmed", "evidence_ids": ["yolkin2021"]})
        write_json(self.project / "memory" / "claims-register.json", claims)
        report = diagnose(self.project)
        errors = codes(report, "ERROR")
        self.assertIn("CLAIM_EVIDENCE_URL_ONLY", errors)
        self.assertIn("CLAIM_EVIDENCE_PATH_INVALID", errors)
        self.assertIn("source_ids", finding_of(report, "CLAIM_EVIDENCE_REF_MISSING")["message"])

    def test_defense_answers_are_required_on_final(self) -> None:
        index = read_json(self.project / "evidence" / "index.json")
        index["defense"] = []
        write_json(self.project / "evidence" / "index.json", index)
        self.assertIn("USER_EVIDENCE_REQUIRED", codes(diagnose(self.project), "ERROR"))
        code, payload = audit_cmd(self.project, "plan", "--kind", "primary", "--preset", "final", "--json")
        self.assertEqual("USER_EVIDENCE_REQUIRED", payload["code"])

    def test_json_formatting_and_key_order_do_not_make_snapshot_stale(self) -> None:
        def reordered(value: Any) -> Any:
            if isinstance(value, dict):
                return {key: reordered(value[key]) for key in reversed(list(value))}
            if isinstance(value, list):
                return [reordered(item) for item in value]
            return value

        for relative in ("sources.json", "vkr-project.json", "evidence/index.json", "memory/claims-register.json"):
            path = self.project / relative
            path.write_text(BOM + json.dumps(reordered(read_json(path)), ensure_ascii=False, indent=4), encoding="utf-8")
        report = diagnose(self.project)
        self.assertEqual("READY_TO_SUBMIT", report["readiness"], codes(report))
        self.assertNotIn("AUDIT_SNAPSHOT_STALE", codes(report))
        self.assertNotIn("FINAL_DOCX_OUTDATED", codes(report))
        self.assertNotEqual(common.content_sha256(self.project / "sources.json"), common.sha256_file(self.project / "sources.json"))
        sources = read_json(self.project / "sources.json")
        sources.reverse()
        write_json(self.project / "sources.json", sources)
        report = diagnose(self.project)
        self.assertIn("AUDIT_SNAPSHOT_STALE", codes(report, "ERROR"))
        self.assertIn("FINAL_DOCX_OUTDATED", codes(report, "ERROR"))

    def test_word_corrupt_package_is_reported(self) -> None:
        final = self.project / "final" / "vkr.docx"

        def undeclared_prefix(xml: str) -> str:
            if 'mc:Ignorable="' in xml:
                return xml.replace('mc:Ignorable="', 'mc:Ignorable="w99x ', 1)
            head = xml.index("<w:document") + len("<w:document")
            declaration = "" if "xmlns:mc=" in xml else ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
            return xml[:head] + declaration + ' mc:Ignorable="w99x"' + xml[head:]

        broken = rewrite_docx(final, self.base / "broken.docx", edits={"word/document.xml": undeclared_prefix})
        shutil.copyfile(str(broken), str(final))
        report = diagnose(self.project)
        item = finding_of(report, "FINAL_DOCX_CORRUPT")
        self.assertEqual("ERROR", item["severity"])
        self.assertIn("w99x", " ".join(item["details"]))
        self.assertTrue(
            any("build_vkr.py" in action and "clean_docx_metadata.py --in-place" in action for action in report["next_actions"] if "повреждён" in action),
            report["next_actions"],
        )

    def test_final_docx_name_has_rebuild_action(self) -> None:
        (self.project / "final" / "vkr.docx").rename(self.project / "final" / "vkr-submit.docx")
        report = diagnose(self.project)
        self.assertIn("FINAL_DOCX_NAME", codes(report, "ERROR"))
        self.assertTrue(
            any("build_vkr.py" in action and "final/vkr.docx" in action and "убери прежний файл из final/" in action for action in report["next_actions"]),
            report["next_actions"],
        )

    def test_title_page_required_matches_build_markers(self) -> None:
        import create_vkr_docx as generator  # noqa: WPS433
        from docx import Document  # noqa: WPS433

        document = Document()
        generator.setup_styles(document)
        missing = generator.Renderer(document, None, []).title_page({}, "Тема")
        self.assertEqual(set(common.TITLE_PAGE_BUILD_REQUIRED), set(missing))
        self.assertEqual(("author",) + tuple(common.TITLE_PAGE_BUILD_REQUIRED), tuple(doctor.TITLE_PAGE_REQUIRED))
        config = read_json(self.project / "vkr-project.json")
        config["title_page"]["program_profile"] = ""
        config["title_page"]["head_of_department"] = ""
        write_json(self.project / "vkr-project.json", config)
        report = diagnose(self.project)
        self.assertEqual(["program_profile", "head_of_department"], finding_of(report, "TITLE_PAGE_INCOMPLETE")["details"])
        self.assertTrue(any("program_profile, head_of_department" in action for action in report["next_actions"]), report["next_actions"])

    def test_broken_findings_view_and_manifest_have_recovery_actions(self) -> None:
        (self.project / "audit" / "findings.json").write_text("{broken", encoding="utf-8")
        report = diagnose(self.project)
        self.assertEqual("PASS", report["status"])
        self.assertTrue(any("findings --json" in action for action in report["next_actions"]))
        manifest_path = self.project / "audit" / "manifest.json"
        manifest_path.write_bytes(manifest_path.read_bytes()[:-200])
        report = diagnose(self.project)
        self.assertTrue(any("manifest.json.bak" in action for action in report["next_actions"]))
        self.assertFalse(any(" snapshot" in action and "Зафиксируй" in action for action in report["next_actions"]))
        code, payload = audit_cmd(self.project, "next", "--stage", "final", "--json")
        self.assertEqual(0, code)
        self.assertIn("manifest_error", payload)


class FailedRunTest(ProjectCase):
    state = "failed"

    def test_hiding_failed_run_by_manifest_edits_is_detected(self) -> None:
        variants = {
            "planned_seq_removed": lambda m: m["runs"][0].pop("planned_seq"),
            "planned_seq_string": lambda m: m["runs"][0].__setitem__("planned_seq", str(m["runs"][0]["planned_seq"])),
            "run_deleted": lambda m: m.__setitem__("runs", m["runs"][1:]),
            "schema_removed": lambda m: m.pop("schema"),
        }
        pristine = self.manifest()
        for name, mutate in variants.items():
            with self.subTest(name):
                manifest = json.loads(json.dumps(pristine))
                mutate(manifest)
                self.save_manifest(manifest)
                errors = codes(diagnose(self.project), "ERROR")
                self.assertTrue({"AUDIT_RUN_TAMPERED", "AUDIT_SEQ_GAP", "AUDIT_REPORT_ORPHAN"} & set(errors), (name, errors))
                code, payload = audit_cmd(self.project, "plan", "--kind", "primary", "--preset", "final", "--checkpoint", "final-2", "--json")
                self.assertEqual("AUDIT_MANIFEST_TAMPERED", payload.get("code"), (name, payload))

    def test_deleting_failed_run_with_reports_leaves_seq_gap(self) -> None:
        manifest = self.manifest()
        deleted = manifest["runs"].pop(0)
        self.save_manifest(manifest)
        shutil.rmtree(self.project / "audit" / "reports" / deleted["run_id"])
        shutil.rmtree(self.project / "audit" / "briefs" / deleted["run_id"], ignore_errors=True)
        self.assertIn("AUDIT_SEQ_GAP", codes(diagnose(self.project), "ERROR"))

    def test_run_numbers_and_reports_are_never_reused(self) -> None:
        manifest = self.manifest()
        self.assertEqual(2, common.next_run_number(self.project, manifest))
        self.assertEqual(2, common.next_run_number(self.project, dict(manifest, runs=[])))
        code, plan = audit_cmd(self.project, "plan", "--kind", "primary", "--preset", "final", "--checkpoint", "final-2", "--json")
        target = self.project / "audit" / "reports" / plan["run_id"] / f"{plan['tasks'][0]['task_id']}.json"
        write_json(target, {"placed": "before record"})
        path = report_file(self.base, "pass", PASS_REPORT)
        code, payload = audit_cmd(self.project, "record", "--run", plan["run_id"], "--task", plan["tasks"][0]["task_id"], "--auditor-id", "N1", "--report", path, "--json")
        self.assertEqual("REPORT_EXISTS", payload["code"])
        self.assertEqual({"placed": "before record"}, read_json(target))

    def test_rerun_until_green_on_same_snapshot_does_not_pass(self) -> None:
        code, payload = audit_cmd(self.project, "plan", "--kind", "primary", "--preset", "final", "--json")
        self.assertEqual("RUN_FAILED_ON_SNAPSHOT", payload["code"])
        run_wave(self.project, "primary", "final", "P2", extra=("--checkpoint", "final-2"), expect_close=0)
        run_wave(self.project, "blind_regression", "final", "B2", expect_close=0)
        errors = codes(diagnose(self.project), "ERROR")
        self.assertIn("AUDIT_FINDINGS_OPEN", errors)
        self.assertIn("AUDIT_FAIL_UNCLOSED", errors)

    def test_recheck_retry_does_not_cancel_negative_verdict(self) -> None:
        plan_md = self.project / "plan.md"
        plan_md.write_text(plan_md.read_text(encoding="utf-8") + "\n", encoding="utf-8")
        validate_and_snapshot(self.project)
        first = recheck(self.project, "TR1-SRC", "fail")
        plan = first["plan"]
        denied = record_recheck(self.project, plan, "TR2-SRC", "pass", expect=1)
        self.assertEqual("RETRY_REASON_REQUIRED", denied["code"])
        again = record_recheck(self.project, plan, "TR2-SRC", "pass", retry="первый аудитор ошибся")
        self.assertEqual("open", again["rechecks"][0]["finding_status"])
        audit_cmd(self.project, "close", "--run", plan["run_id"])
        final_waves(self.project, "X")
        errors = codes(diagnose(self.project), "ERROR")
        self.assertIn("AUDIT_FINDINGS_OPEN", errors)

    def test_draft_fixed_but_docx_not_rebuilt_fails(self) -> None:
        chapter = self.project / "drafts" / "chapter-1.md"
        chapter.write_text(chapter.read_text(encoding="utf-8").replace("[@egorov2020, с. 5]", "[@egorov2020, с. 17]"), encoding="utf-8")
        audit_cmd(self.project, "snapshot", "--json")
        result = recheck(self.project, "TR-SRC", "pass")
        self.assertEqual("resolved", result["rechecks"][0]["finding_status"])
        audit_cmd(self.project, "close", "--run", result["plan"]["run_id"])
        final_waves(self.project, "D")
        report = diagnose(self.project)
        self.assertIn("FINAL_DOCX_OUTDATED", codes(report, "ERROR"))
        self.assertEqual("NOT_READY", report["readiness"])

    def test_not_recheckable_then_pass_on_new_snapshot_resolves(self) -> None:
        fix_and_rebuild(self.project)
        first = recheck(self.project, "TR-SRC", "not_recheckable")
        self.assertEqual(1, audit_cmd(self.project, "close", "--run", first["plan"]["run_id"])[0])
        second = recheck(self.project, "TR2-SRC", "pass")
        self.assertEqual("resolved", second["rechecks"][0]["finding_status"])
        self.assertEqual(0, audit_cmd(self.project, "close", "--run", second["plan"]["run_id"])[0])
        final_waves(self.project, "F2")
        report = diagnose(self.project)
        self.assertEqual("PASS", report["status"], codes(report, "ERROR"))

    def test_invalid_finding_on_same_snapshot_then_waves_pass(self) -> None:
        weak = recheck(self.project, "TR-SRC", "invalid_finding", level="E1", expect=1)
        self.assertEqual("REPORT_SCHEMA_INVALID", weak["code"])
        code, plans = audit_cmd(self.project, "plan", "--kind", "targeted_recheck", "--findings", "F-SRC-001", "--json")
        self.assertEqual("RUN_ALREADY_OPEN", plans["code"])
        manifest = self.manifest()
        run_id = manifest["runs"][-1]["run_id"]
        plan = {"run_id": run_id, "tasks": [{"task_id": manifest["runs"][-1]["planned_tasks"][0]["task_id"]}]}
        strong = record_recheck(self.project, plan, "TR-SRC", "invalid_finding", level="E4")
        self.assertEqual("rejected", strong["rechecks"][0]["finding_status"])
        audit_cmd(self.project, "close", "--run", run_id)
        final_waves(self.project, "F4", extra=("--checkpoint", "final-2"))
        report = diagnose(self.project)
        self.assertEqual("PASS", report["status"], codes(report, "ERROR"))

    def test_prefinal_sees_failed_run_after_file_touch(self) -> None:
        plan_md = self.project / "plan.md"
        plan_md.write_text(plan_md.read_text(encoding="utf-8") + " ", encoding="utf-8")
        errors = codes(diagnose(self.project, "prefinal"), "ERROR")
        self.assertIn("AUDIT_FINDINGS_OPEN", errors)
        self.assertIn("AUDIT_FAIL_UNCLOSED", errors)


class ReadyProjectTest(ProjectCase):
    state = "ready"

    def test_final_waves_need_final_preset_and_full_matrix(self) -> None:
        run_wave(self.project, "primary", "prefinal", "X4aP", expect_close=0)
        code, payload = audit_cmd(self.project, "plan", "--kind", "blind_regression", "--preset", "final", "--json")
        self.assertEqual("PRIMARY_WAVE_REQUIRED", payload["code"])
        roles = "MET,SRC,LOG,LNG,STY,EVD,TEC,DOC,OWN"
        run_wave(self.project, "primary", "custom", "X4cP", extra=("--roles", roles), expect_close=0)
        run_wave(self.project, "blind_regression", "custom", "X4cB", extra=("--roles", roles), expect_close=0)
        report = diagnose(self.project)
        details = finding_of(report, "AUDIT_PRIMARY_WAVE_MISSING")["details"]
        self.assertTrue(all(any("--preset final" in reason for reason in reasons) for reasons in details.values()), details)
        code, payload = audit_cmd(self.project, "plan", "--kind", "primary", "--preset", "final", "--replicas", "MET=A", "--json")
        self.assertEqual(2, code)
        self.assertEqual("REPLICAS_BELOW_MATRIX", payload["code"])

    def test_degraded_independence_caps_readiness_and_mixed_ids(self) -> None:
        run_wave(self.project, "primary", "final", "x", extra=("--independence", "degraded_independence"), auditor="main", expect_close=0)
        run_wave(self.project, "blind_regression", "final", "x", extra=("--independence", "degraded_independence"), auditor="main", expect_close=0)
        report = diagnose(self.project)
        self.assertEqual("PASS", report["status"], codes(report, "ERROR"))
        self.assertEqual("READY_FOR_SUPERVISOR_REVIEW", report["readiness"])

    def test_auditor_ids_are_strict(self) -> None:
        code, plan = audit_cmd(self.project, "plan", "--kind", "primary", "--preset", "final", "--json")
        path = report_file(self.base, "id", PASS_REPORT)
        task = plan["tasks"][0]["task_id"]
        for name in ("auditor" + ZWSP, "P1-SRC-" + CYRILLIC_A, " Main ", "Main Agent", "студент"):
            code, payload = audit_cmd(self.project, "record", "--run", plan["run_id"], "--task", task, "--auditor-id", name, "--report", path, "--json")
            self.assertEqual("AUDITOR_ID_INVALID", payload["code"], name)
        for name in ("main", "MAIN", "main-1", "main2", "claude-main", "Author-Bot", "Student1", "user", "self", "coordinator", "assistant_7",
                     "ClaudeMain", "MainAgent", "SelfCheck", "StudentBot", "authorReview", "GPTMain"):
            code, payload = audit_cmd(self.project, "record", "--run", plan["run_id"], "--task", task, "--auditor-id", name, "--report", path, "--json")
            self.assertEqual("AUDITOR_NOT_INDEPENDENT", payload["code"], name)
        # T3: склейка без разделителей — «M.A.I.N», «ma-in», «s.e.l.f» — тоже запрещённые лексемы.
        for name in ("M.A.I.N", "ma-in", "m-a-i-n", "s.e.l.f", "C.l.a.u.d.e", "a.u.t.h.o.r"):
            code, payload = audit_cmd(self.project, "record", "--run", plan["run_id"], "--task", task, "--auditor-id", name, "--report", path, "--json")
            self.assertEqual("AUDITOR_NOT_INDEPENDENT", payload["code"], name)
        for name in ("Mainn", "ma-in-1", "m.a.i.n.e", "domain-expert"):  # похожие, но не запрещённые
            self.assertFalse(common.is_forbidden_auditor(name), name)
        self.assertEqual(0, audit_cmd(self.project, "record", "--run", plan["run_id"], "--task", task, "--auditor-id", "mainframe-auditor", "--report", path)[0])

    def test_auditor_case_variant_is_rejected_between_waves(self) -> None:
        run_wave(self.project, "primary", "final", "P", expect_close=0)
        code, plan = audit_cmd(self.project, "plan", "--kind", "blind_regression", "--preset", "final", "--json")
        log_task = next(task for task in plan["tasks"] if task["auditor_role"] == "LOG")
        path = report_file(self.base, "blind-log", PASS_REPORT)
        code, payload = audit_cmd(self.project, "record", "--run", plan["run_id"], "--task", log_task["task_id"], "--auditor-id", "p-log-a", "--report", path, "--json")
        self.assertEqual("AUDITOR_NOT_INDEPENDENT", payload["code"])

    def test_one_auditor_cannot_close_two_gates(self) -> None:
        code, plan = audit_cmd(self.project, "plan", "--kind", "primary", "--preset", "final", "--json")
        path = report_file(self.base, "same", PASS_REPORT)
        self.assertEqual(0, audit_cmd(self.project, "record", "--run", plan["run_id"], "--task", plan["tasks"][0]["task_id"], "--auditor-id", "agent-x", "--report", path)[0])
        other = next(task for task in plan["tasks"] if task["base_gate"] != plan["tasks"][0]["base_gate"])
        code, payload = audit_cmd(self.project, "record", "--run", plan["run_id"], "--task", other["task_id"], "--auditor-id", "AGENT-X", "--report", path, "--json")
        self.assertEqual("AUDITOR_NOT_INDEPENDENT", payload["code"])

    def test_minor_fail_with_accepted_risk_does_not_block_new_waves(self) -> None:
        run_wave(self.project, "primary", "final", "F1P", reports={"DOC-A": {"status": "fail", "summary": "сдвиг таблицы", "findings": [MINOR_FINDING]}}, expect_close=1)
        self.assertEqual(0, audit_cmd(self.project, "accept-risk", "F-DOC-001", "--reason", "Сдвиг 1 мм от Word, согласовано")[0])
        # Подсказка doctor («plan --kind primary --preset final») выполнима буквально: замечания прежнего run
        # закрыты, поэтому новая волна получает новый checkpoint автоматически.
        final_waves(self.project, "F1Q")
        report = diagnose(self.project)
        self.assertEqual("PASS", report["status"], codes(report, "ERROR"))
        checkpoints = [run.get("checkpoint_id") for run in json.loads((self.project / "audit" / "manifest.json").read_text(encoding="utf-8-sig"))["runs"]]
        self.assertIn("final-2", checkpoints)

    def test_premature_close_requires_abandon(self) -> None:
        code, plan = audit_cmd(self.project, "plan", "--kind", "primary", "--preset", "final", "--json")
        path = report_file(self.base, "first", PASS_REPORT)
        audit_cmd(self.project, "record", "--run", plan["run_id"], "--task", plan["tasks"][0]["task_id"], "--auditor-id", "F3-first", "--report", path)
        code, payload = audit_cmd(self.project, "close", "--run", plan["run_id"], "--json")
        self.assertEqual(("RUN_TASKS_PENDING", 1), (payload["code"], code))
        self.assertIn("AUDIT_RUNS_INCOMPLETE", codes(diagnose(self.project, "prefinal"), "ERROR"))
        plan_md = self.project / "plan.md"
        plan_md.write_text(plan_md.read_text(encoding="utf-8") + " ", encoding="utf-8")
        self.assertIn("AUDIT_RUNS_INCOMPLETE", codes(diagnose(self.project, "prefinal"), "ERROR"))
        plan_md.write_text(plan_md.read_text(encoding="utf-8")[:-1], encoding="utf-8")
        code, payload = audit_cmd(self.project, "close", "--run", plan["run_id"], "--abandon", "--reason", "начали заново", "--json")
        self.assertEqual("abandoned", payload["status"])
        final_waves(self.project, "F3", extra=("--checkpoint", "final-2"))
        report = diagnose(self.project)
        self.assertEqual("PASS", report["status"], codes(report, "ERROR"))

    def test_accept_risk_rules(self) -> None:
        minor_sty = dict(MINOR_FINDING, fix_class="AUTO_SAFE", criterion="Повтор начала абзацев")
        run_wave(self.project, "primary", "custom", "S1", extra=("--roles", "STY", "--checkpoint", "sty-1"), reports={"STY": {"status": "pass", "summary": "мелочь", "findings": [minor_sty]}}, expect_close=0)
        for number in (2, 3, 4):
            run_wave(self.project, "primary", "custom", f"S{number}", extra=("--roles", "STY", "--checkpoint", f"sty-{number}"), expect_close=0)
        code, payload = audit_cmd(self.project, "accept-risk", "F-STY-001", "--reason", "стиль автора", "--json")
        self.assertEqual("ACCEPT_RISK_DENIED", payload["code"])
        self.assertIn("снимках", payload["error"])

    def test_waive_codes_and_rules(self) -> None:
        supervisor = dict(MAJOR_FINDING, fix_class="SUPERVISOR_APPROVAL")
        run_wave(self.project, "primary", "custom", "W", extra=("--roles", "MET,SRC", "--checkpoint", "chapter-1"), reports={"MET": {"status": "fail", "summary": "отклонение", "findings": [supervisor]}, "SRC": FAIL_REPORT}, expect_close=1)
        (self.project / "evidence" / "empty.md").write_text("", encoding="utf-8")
        self.assertEqual((2, "APPROVAL_INVALID"), (lambda r: (r[0], r[1]["code"]))(audit_cmd(self.project, "waive", "F-MET-001", "--approval", "plan.md", "--json")))
        self.assertEqual((1, "APPROVAL_EMPTY"), (lambda r: (r[0], r[1]["code"]))(audit_cmd(self.project, "waive", "F-MET-001", "--approval", "evidence/empty.md", "--json")))
        approval = self.project / "evidence" / "approvals" / "approval.md"
        approval.parent.mkdir(parents=True, exist_ok=True)
        approval.write_text("Руководитель одобрил отклонение структуры.\n", encoding="utf-8")
        waive = lambda finding: (lambda r: (r[0], r[1].get("code")))(audit_cmd(self.project, "waive", finding, "--approval", "evidence/approvals/approval.md", "--json"))  # noqa: E731
        # V3b п. 4: файл одобрения вне evidence/index.json → approvals не входит в снимок — отказ
        self.assertEqual((1, "APPROVAL_NOT_REGISTERED"), waive("F-MET-001"))
        index_path = self.project / "evidence" / "index.json"
        index = read_json(index_path)
        index["approvals"] = [{"id": "approval-structure", "path": "evidence/approvals/approval.md"}]
        write_json(index_path, index)
        self.assertEqual((1, "APPROVAL_NOT_IN_SNAPSHOT"), waive("F-MET-001"))
        snapshot_only(self.project)
        inputs = read_json(self.project / "audit" / "snapshot-inputs.json")["inputs"]
        self.assertIn(("evidence/approvals/approval.md", "evidence_approvals"), {(item["logical_id"], item["kind"]) for item in inputs})
        self.assertEqual((1, "WAIVE_DENIED"), waive("F-SRC-001"))
        self.assertEqual((0, None), waive("F-MET-001"))
        self.assertNotIn("AUDIT_WAIVER_INVALID", codes(diagnose(self.project, "prefinal")))
        index["approvals"] = []
        write_json(index_path, index)
        self.assertIn("AUDIT_WAIVER_INVALID", codes(diagnose(self.project, "prefinal"), "ERROR"))
        index["approvals"] = [{"id": "approval-structure", "path": "evidence/approvals/approval.md"}]
        write_json(index_path, index)
        approval.write_text("Подменено\n", encoding="utf-8")
        self.assertIn("AUDIT_WAIVER_INVALID", codes(diagnose(self.project, "prefinal"), "ERROR"))

    def test_accept_risk_counts_only_snapshots_with_changed_text(self) -> None:
        minor_sty = dict(MINOR_FINDING, fix_class="AUTO_SAFE", criterion="Повтор начала абзацев")
        run_wave(self.project, "primary", "custom", "S0", extra=("--roles", "STY", "--checkpoint", "sty-0"), reports={"STY": {"status": "pass", "summary": "мелочь", "findings": [minor_sty]}}, expect_close=0)
        plan = self.project / "plan.md"
        chapter = self.project / "drafts" / "chapter-2.md"
        for number in (1, 2, 3):
            if number == 2:
                chapter.write_text(chapter.read_text(encoding="utf-8").replace("\n", "   \n", 1) + "\n\n", encoding="utf-8")
            else:
                plan.write_text(plan.read_text(encoding="utf-8") + f"\nУточнение плана {number}.\n", encoding="utf-8")
            snapshot_only(self.project)
            run_wave(self.project, "primary", "custom", f"S{number}", extra=("--roles", "STY", "--checkpoint", f"sty-{number}"), expect_close=0)
        code, payload = audit_cmd(self.project, "accept-risk", "F-STY-001", "--reason", "стиль автора", "--json")
        self.assertEqual((1, "ACCEPT_RISK_DENIED"), (code, payload.get("code")), payload)
        self.assertIn("изменённым текстом", payload["error"])
        for number in (4, 5, 6):
            chapter.write_text(chapter.read_text(encoding="utf-8") + f"\nДополнение {number}: уточнены выводы по разделу о пилотировании.\n", encoding="utf-8")
            rebuild(self.project)
            snapshot_only(self.project)
            run_wave(self.project, "primary", "custom", f"S{number}", extra=("--roles", "STY", "--checkpoint", f"sty-{number}"), expect_close=0)
        code, payload = audit_cmd(self.project, "accept-risk", "F-STY-001", "--reason", "стиль автора", "--json")
        self.assertEqual(0, code, payload)

    def test_lowered_intensity_before_final_waves_is_reported(self) -> None:
        run_wave(self.project, "primary", "custom", "C", extra=("--roles", "SRC", "--checkpoint", "sources-1"), expect_close=0)
        config = read_json(self.project / "vkr-project.json")
        config["audit_intensity"] = "balanced"
        write_json(self.project / "vkr-project.json", config)
        rebuild(self.project)
        validate_and_snapshot(self.project)
        final_waves(self.project, "L", extra=("--roles", "OWN-DEFENSE"))
        report = diagnose(self.project)
        self.assertEqual("READY_TO_SUBMIT", report["readiness"], codes(report, "ERROR"))
        warning = finding_of(report, "AUDIT_INTENSITY_LOWERED")
        self.assertEqual(("WARNING", "balanced", "strict"), (warning["severity"], warning["details"]["applied"], warning["details"]["earlier_max"]))
        matrix = report["audit"]["final_wave_matrix"]
        self.assertEqual(("balanced", "balanced"), (matrix["audit_intensity"], matrix["applied_audit_intensity"]))
        self.assertEqual(common.final_wave_requirements("primary", "balanced", "project"), matrix["primary"])
        self.assertTrue(any("audit_intensity=balanced" in action for action in report["next_actions"]), report["next_actions"])

    def test_report_schema_and_not_applicable(self) -> None:
        code, plan = audit_cmd(self.project, "plan", "--kind", "primary", "--preset", "final", "--json")
        task = next(item for item in plan["tasks"] if item["auditor_role"] == "LOG")

        def attempt(report: Dict[str, Any], auditor: str = "S") -> Dict[str, Any]:
            path = report_file(self.base, "schema", report)
            return audit_cmd(self.project, "record", "--run", plan["run_id"], "--task", task["task_id"], "--auditor-id", auditor, "--report", path, "--json")[1]

        for severity in ("CRITICAL", "high", "MAJ" + chr(0x041E) + "R"):
            self.assertEqual("REPORT_SCHEMA_INVALID", attempt({"status": "fail", "summary": "x", "findings": [dict(MAJOR_FINDING, severity=severity)]}).get("code"), severity)
        self.assertIn("E2", attempt({"status": "fail", "summary": "x", "findings": [dict(MAJOR_FINDING, evidence_level="E1")]})["error"])
        self.assertIn("служебные", attempt(dict(PASS_REPORT, task_id="forged"))["error"])
        recorded = attempt({"status": "pass", "summary": "x", "findings": [dict(MAJOR_FINDING, severity="major")]})
        self.assertEqual("fail", recorded["task_status"])
        na = {"status": "not_applicable", "summary": "нет предмета", "not_applicable_reason": "нет", "findings": []}
        for role in ("TEC", "OWN", "DOC"):
            other = next(item for item in plan["tasks"] if item["auditor_role"] == role)
            path = report_file(self.base, f"na-{role}", na)
            code, payload = audit_cmd(self.project, "record", "--run", plan["run_id"], "--task", other["task_id"], "--auditor-id", f"NA-{role}", "--report", path, "--json")
            self.assertEqual("REPORT_SCHEMA_INVALID", payload["code"], role)

    def test_record_refuses_stale_snapshot(self) -> None:
        code, plan = audit_cmd(self.project, "plan", "--kind", "primary", "--preset", "final", "--json")
        (self.project / "plan.md").write_text("# План изменён во время аудита\n", encoding="utf-8")
        path = report_file(self.base, "stale", PASS_REPORT)
        code, payload = audit_cmd(self.project, "record", "--run", plan["run_id"], "--task", plan["tasks"][0]["task_id"], "--auditor-id", "Z", "--report", path, "--json")
        self.assertEqual("SNAPSHOT_STALE", payload["code"])

    def test_next_lists_concrete_commands(self) -> None:
        with patched_validator():
            code, payload = audit_cmd(self.project, "next", "--stage", "final", "--json")
        self.assertTrue(any("plan --kind primary --preset final" in action for action in payload["next_actions"]), payload["next_actions"])

    def test_stale_lock_of_dead_process_is_taken_over(self) -> None:
        (self.project / "audit" / ".vkr_audit.lock").write_text("99999999", encoding="ascii")
        self.assertTrue(common.pid_alive(os.getpid()))
        self.assertFalse(common.pid_alive(99999999))
        started = time.monotonic()
        code, payload = audit_cmd(self.project, "snapshot", "--json")
        self.assertEqual(0, code, payload)
        # Блокировка мёртвого процесса снимается сразу, без выдержки «неизвестного владельца» (5 с).
        self.assertLess(time.monotonic() - started, 4.0)


class FixBDoctorTest(ProjectCase):
    """Fix-B: doctor на prefinal, подсказка checkpoint памяти и объём в аннотации (B1, B2, B9, B13)."""

    state = "validated"

    def write_fields(self, pages: int, sha: Optional[str] = None, fingerprint: Optional[str] = None, legacy: bool = False) -> None:
        """Запись как у update_docx_fields.py (Fix-A); legacy — прежний формат без text_fingerprint."""
        docx = self.project / "final" / "vkr.docx"
        record: Dict[str, Any] = {
            "schema": "vkr-docx-fields", "docx": "final/vkr.docx",
            "docx_sha256": sha if sha is not None else common.sha256_file(docx),
            "pages": pages, "updated_at": "2026-09-14T10:00:00Z",
        }
        if not legacy:
            record.update({
                "fingerprint_schema": common.FINGERPRINT_SCHEMA,
                "text_fingerprint": fingerprint if fingerprint is not None else common.docx_text_fingerprint(docx)[0],
            })
        write_json(self.project / "exports" / "docx-fields.json", record)

    def test_prefinal_outdated_docx_is_error_and_not_ready(self) -> None:  # B1
        self.assertNotIn("FINAL_DOCX_OUTDATED", codes(diagnose(self.project, "prefinal")))
        chapter = self.project / "drafts" / "chapter-1.md"
        chapter.write_text(chapter.read_text(encoding="utf-8") + "\nАбзац, добавленный после сборки.\n", encoding="utf-8")
        report = diagnose(self.project, "prefinal")
        self.assertIn("FINAL_DOCX_OUTDATED", codes(report, "ERROR"))
        self.assertEqual(("FAIL", "NOT_READY"), (report["status"], report["readiness"]))
        self.assertFalse(any("Можно передавать" in action for action in report["next_actions"]), report["next_actions"])
        self.assertTrue(any("build_vkr.py" in action for action in report["next_actions"]))

    def test_validator_warnings_are_visible_in_validate_and_doctor(self) -> None:  # финальная приёмка, находка 1
        with mock.patch.object(common, "run_validator", fake_validator(0, True, warnings=2)):
            code, payload = audit_cmd(self.project, "validate", "--json")
            self.assertEqual(0, code, payload)
            self.assertEqual(2, len(payload["warnings_preview"]), payload)
            self.assertTrue(any("предупреждени" in step for step in payload["next"]), payload["next"])
            final = doctor.diagnose(self.project, "final")
        self.assertIn("VALIDATION_WARNINGS", codes(final, "WARNING"))
        self.assertTrue(any("automated-validation.json" in action for action in final["next_actions"]), final["next_actions"])
        # prefinal читает записанный отчёт без повторного запуска валидатора
        with patched_validator(0, True):
            prefinal = doctor.diagnose(self.project, "prefinal")
        self.assertIn("VALIDATION_WARNINGS", codes(prefinal, "WARNING"))

    def test_memory_checkpoint_hint_run_literally_clears_artifact_changed(self) -> None:  # B2
        report = diagnose(self.project, "draft")
        self.assertIn("ARTIFACT_CHANGED", codes(report, "WARNING"))
        hint = next(action for action in report["next_actions"] if action.startswith("Запиши checkpoint памяти"))
        self.assertIn("record --event <event.json> --all-changed --json", hint)
        event = self.base / "event.json"
        write_json(event, {"event_type": "checkpoint", "summary": "Черновики заполнены"})
        parts = [part.strip('"') for part in re.findall(r'"[^"]*"|\S+', hint.split(": ", 1)[1])]
        self.assertEqual("python", parts[0])
        command = [sys.executable] + [str(event) if part == "<event.json>" else part for part in parts[1:]]
        env = dict(os.environ, PYTHONDONTWRITEBYTECODE="1", PYTHONIOENCODING="utf-8")
        result = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", env=env)
        self.assertEqual(0, result.returncode, result.stdout + result.stderr)
        self.assertNotIn("ARTIFACT_CHANGED", codes(diagnose(self.project, "draft")))

    def test_prefinal_without_prefinal_wave_does_not_say_ready_to_hand_over(self) -> None:  # B13 (F23)
        report = diagnose(self.project, "prefinal")
        self.assertEqual("PASS", report["status"], codes(report, "ERROR"))
        self.assertIn("AUDIT_PREFINAL_WAVE_MISSING", codes(report, "WARNING"))
        self.assertFalse(any("Можно передавать" in action for action in report["next_actions"]), report["next_actions"])
        self.assertTrue(any("plan --kind primary --preset prefinal" in action for action in report["next_actions"]), report["next_actions"])
        run_wave(self.project, "primary", "prefinal", "PF", expect_close=0)
        report = diagnose(self.project, "prefinal")
        self.assertNotIn("AUDIT_PREFINAL_WAVE_MISSING", codes(report))
        self.assertTrue(any("Можно передавать" in action for action in report["next_actions"]), report["next_actions"])

    def test_annotation_page_count_forms(self) -> None:  # B9
        self.assertEqual([52], doctor.annotation_page_counts("Работа содержит 52 с., 12 рис., 3 табл., 25 источников, 2 прил."))
        self.assertEqual([32], doctor.annotation_page_counts("Объём — 32 страницы, 2 рисунка, 4 таблицы; использовано 12 источников."))
        self.assertEqual([45], doctor.annotation_page_counts("Работа изложена на 45 страницах [@egorov2020, с. 5]."))
        self.assertEqual([60], doctor.annotation_page_counts("Объём 60 стр. <!-- 58 с. --> `70 с.`"))
        self.assertEqual([], doctor.annotation_page_counts("Использовано 12 источников; пилот в 8 классе (22 учащихся)."))

    def test_annotation_pages_are_checked_against_docx_fields(self) -> None:  # B9 (F11, F16)
        report = diagnose(self.project, "prefinal")
        unverified = finding_of(report, "ANNOTATION_PAGES_UNVERIFIED")
        self.assertEqual("INFO", unverified["severity"])
        self.assertTrue(any("update_docx_fields.py" in action for action in report["next_actions"]), report["next_actions"])
        self.write_fields(45)  # fixture: «Объём работы — 45 страниц»
        report = diagnose(self.project, "prefinal")
        self.assertFalse({"ANNOTATION_PAGES_MISMATCH", "ANNOTATION_PAGES_UNVERIFIED"} & set(codes(report)))
        self.write_fields(47)
        for stage in ("prefinal", "final"):
            report = diagnose(self.project, stage)
            mismatch = finding_of(report, "ANNOTATION_PAGES_MISMATCH")
            self.assertEqual("ERROR", mismatch["severity"], stage)
            self.assertEqual(47, mismatch["details"]["actual"])
            self.assertIn("47", mismatch["message"])
            self.assertEqual("NOT_READY", report["readiness"])
            self.assertTrue(any("drafts/annotation.md на 47 с." in action for action in report["next_actions"]), report["next_actions"])
        # Запись о другой версии текста — нет данных, а не ошибка, даже если docx_sha256 совпадает.
        self.write_fields(47, fingerprint="f" * 64)
        report = diagnose(self.project, "prefinal")
        self.assertNotIn("ANNOTATION_PAGES_MISMATCH", codes(report))
        self.assertIn("ANNOTATION_PAGES_UNVERIFIED", codes(report, "INFO"))
        self.assertIn("другой версии", finding_of(report, "ANNOTATION_PAGES_UNVERIFIED")["message"])
        self.write_fields(47, legacy=True)  # прежний формат без отпечатка сверяется по docx_sha256
        self.assertIn("ANNOTATION_PAGES_MISMATCH", codes(diagnose(self.project, "prefinal"), "ERROR"))
        self.write_fields(47, sha="0" * 64, legacy=True)
        self.assertIn("ANNOTATION_PAGES_UNVERIFIED", codes(diagnose(self.project, "prefinal"), "INFO"))

    def test_docx_fields_relate_after_metadata_cleaning_by_text_fingerprint(self) -> None:  # B9
        import clean_docx_metadata as cleaner  # noqa: WPS433

        docx = self.project / "final" / "vkr.docx"
        self.write_fields(47)
        cleaner.clean_metadata(docx, docx, seed="b9")
        self.assertNotEqual(read_json(self.project / "exports" / "docx-fields.json")["docx_sha256"], common.sha256_file(docx))
        self.assertIn("ANNOTATION_PAGES_MISMATCH", codes(diagnose(self.project, "prefinal"), "ERROR"))
        self.write_fields(45, sha="0" * 64)  # sha после очистки другой, текст тот же — объём сверен
        self.assertFalse({"ANNOTATION_PAGES_MISMATCH", "ANNOTATION_PAGES_UNVERIFIED"} & set(codes(diagnose(self.project, "final"))))


class FixCPagesTest(ProjectCase):
    """Fix-C (T1, T2в, T5): число страниц берётся из самого DOCX, запись о полях — во входах снимка."""

    state = "validated"

    ORIGINAL_VOLUME = "Объём работы — 45 страниц, использовано 9 источников."

    def annotation(self, sentence: str) -> None:
        """Заменяет предложение об объёме в аннотации (каждый раз — от исходного текста)."""
        path = self.project / "drafts" / "annotation.md"
        original = getattr(self, "_annotation", None)
        if original is None:
            original = path.read_text(encoding="utf-8")
            assert self.ORIGINAL_VOLUME in original
            self._annotation = original
        path.write_text(original.replace(self.ORIGINAL_VOLUME, sentence), encoding="utf-8")

    def set_word_pages(self, pages: int) -> None:
        """Пишет в DOCX app.xml так, как его пишет Word при сохранении (текст не меняется)."""
        docx = self.project / "final" / "vkr.docx"
        characters = len("\n".join(common.docx_text_lines(docx)))
        app = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
            f"<Template>Normal.dotm</Template><TotalTime>0</TotalTime><Pages>{pages}</Pages>"
            f"<Words>{max(1, characters // 8)}</Words><Characters>{characters}</Characters>"
            "<Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity>"
            f"<CharactersWithSpaces>{characters}</CharactersWithSpaces><AppVersion>16.0000</AppVersion></Properties>"
        )
        with zipfile.ZipFile(str(docx)) as archive:
            parts = [(item, archive.read(item.filename)) for item in archive.infolist()]
        with zipfile.ZipFile(str(docx), "w", zipfile.ZIP_DEFLATED) as out:
            for item, data in parts:
                out.writestr(item, app.encode("utf-8") if item.filename == "docProps/app.xml" else data)

    def write_fields(self, pages: int, word_opened: Optional[bool] = True) -> None:
        docx = self.project / "final" / "vkr.docx"
        record: Dict[str, Any] = {
            "schema": "vkr-docx-fields", "docx": "final/vkr.docx",
            "docx_sha256": common.sha256_file(docx), "pages": pages,
            "updated_at": "2026-09-14T10:00:00Z",
            "fingerprint_schema": common.FINGERPRINT_SCHEMA,
            "text_fingerprint": common.docx_text_fingerprint(docx)[0],
        }
        if word_opened is not None:
            record["word_opened"] = word_opened
        write_json(self.project / "exports" / "docx-fields.json", record)

    def test_pages_of_word_saved_docx_are_read_from_the_file(self) -> None:  # T1
        self.assertEqual((14, None), common.docx_app_pages(WORD_SAVED))
        # сборка python-docx переносит app.xml из шаблона: <Pages>1</Pages> при нулевой статистике
        pages, reason = common.docx_app_pages(SKILL / "assets" / "template-mpgu.docx")
        self.assertIsNone(pages)
        self.assertIn("статистика", str(reason))
        pages, reason = common.docx_app_pages(self.project / "final" / "vkr.docx")
        self.assertIsNone(pages, "у собранного, но не сохранённого Word файла числа страниц нет")
        self.assertIn("app.xml", str(reason))

    def test_forged_pages_in_docx_fields_are_an_error(self) -> None:  # T1a (атака A1)
        self.set_word_pages(14)
        self.annotation("Объём работы — 14 страниц, использовано 9 источников.")
        self.write_fields(14)
        report = diagnose(self.project, "prefinal")
        self.assertFalse({"DOCX_PAGES_MISMATCH", "ANNOTATION_PAGES_MISMATCH"} & set(codes(report)), codes(report))
        # подделка числа страниц под аннотацию (отпечаток текста оставлен настоящим)
        self.annotation("Объём работы — 45 страниц, использовано 9 источников.")
        self.write_fields(45)
        for stage in ("prefinal", "final"):
            report = diagnose(self.project, stage)
            self.assertIn("DOCX_PAGES_MISMATCH", codes(report, "ERROR"), stage)
            self.assertIn("ANNOTATION_PAGES_MISMATCH", codes(report, "ERROR"), stage)
            self.assertEqual("NOT_READY", report["readiness"], stage)
            mismatch = finding_of(report, "DOCX_PAGES_MISMATCH")
            self.assertEqual((45, 14), (mismatch["details"]["record_pages"], mismatch["details"]["docx_pages"]))
            self.assertTrue(any("update_docx_fields.py" in action for action in report["next_actions"]))
        # без записи о полях объём сверяется с самим DOCX
        (self.project / "exports" / "docx-fields.json").unlink()
        report = diagnose(self.project, "final")
        self.assertIn("ANNOTATION_PAGES_MISMATCH", codes(report, "ERROR"))
        self.assertIn("14 с.", finding_of(report, "ANNOTATION_PAGES_MISMATCH")["message"])

    def test_docx_fields_are_part_of_the_snapshot(self) -> None:  # T1b
        self.write_fields(45)
        code, snapshot = audit_cmd(self.project, "snapshot", "--json")
        self.assertEqual(0, code, snapshot)
        inputs = read_json(self.project / "audit" / "snapshot-inputs.json")["inputs"]
        self.assertIn("exports/docx-fields.json", [item["logical_id"] for item in inputs])
        self.write_fields(46)  # правка доверенного файла после снимка видна аудиту
        report = diagnose(self.project, "final")
        self.assertIn("AUDIT_SNAPSHOT_STALE", codes(report, "ERROR"))
        self.assertIn("exports/docx-fields.json", finding_of(report, "AUDIT_SNAPSHOT_STALE")["details"]["changed"])

    def test_final_without_any_page_source_warns(self) -> None:  # T1c
        # Сборка без Word: <Pages> в app.xml остался от шаблона, записи о полях нет.
        self.assertIsNone(common.docx_app_pages(self.project / "final" / "vkr.docx")[0])
        report = diagnose(self.project, "prefinal")
        self.assertIn("ANNOTATION_PAGES_UNVERIFIED", codes(report, "INFO"))
        report = diagnose(self.project, "final")
        finding = finding_of(report, "ANNOTATION_PAGES_UNVERIFIED")
        self.assertEqual("WARNING", finding["severity"])
        self.assertIn("вручную", finding["message"])
        self.assertNotIn("ANNOTATION_PAGES_UNVERIFIED", codes(report, "ERROR"))  # предупреждение, не блокировка

    def test_record_without_word_opened_is_warned_on_final(self) -> None:  # T2в
        self.write_fields(45, word_opened=None)
        report = diagnose(self.project, "final")
        self.assertIn("DOCX_FIELDS_UNCONFIRMED", codes(report, "WARNING"))
        self.assertTrue(any("update_docx_fields.py" in action for action in report["next_actions"]))
        self.assertNotIn("DOCX_FIELDS_UNCONFIRMED", codes(diagnose(self.project, "prefinal")))
        self.write_fields(45, word_opened=True)
        self.assertNotIn("DOCX_FIELDS_UNCONFIRMED", codes(diagnose(self.project, "final")))

    def test_partial_volume_wording_is_a_warning(self) -> None:  # T5 (ложная блокировка)
        self.set_word_pages(14)
        self.write_fields(14)
        self.annotation("Объём текстовой части — 10 страниц без учёта приложений и списка литературы.")
        report = diagnose(self.project, "prefinal")
        self.assertNotIn("ANNOTATION_PAGES_MISMATCH", codes(report))
        partial = finding_of(report, "ANNOTATION_PAGES_PARTIAL")
        self.assertEqual("WARNING", partial["severity"])
        self.assertEqual((14, [10]), (partial["details"]["actual"], partial["details"]["stated"]))
        self.assertIn("полный счётчик Word", partial["message"])
        self.assertTrue(any("вручную" in action for action in report["next_actions"]))
        # правка аннотации делает DOCX устаревшим, но сама формулировка объёма ошибкой не считается
        self.assertEqual(["FINAL_DOCX_OUTDATED"], codes(report, "ERROR"))
        for sentence in ("Объём основного текста — 10 с. без приложений; приложения занимают 4 с.",
                         "Объём работы — 40 стр. основного текста."):
            self.annotation(sentence)
            self.assertNotIn("ANNOTATION_PAGES_MISMATCH", codes(diagnose(self.project, "prefinal")), sentence)
        # честные формулировки полного объёма расхождений не дают
        for sentence in ("Работа изложена на 14 страницах, содержит 5 таблиц и 3 рисунка.",
                         "Объём работы — 14 страниц, это на 5 с. больше первоначального плана.",
                         "Объём работы — 14 страниц при ориентире методички 40–50 страниц.",
                         "Объём работы — 14 с., из них 4 с. приложений.",
                         "Объём работы — 14 страниц. Таблица 2 приведена на с. 45."):
            self.annotation(sentence)
            codes_found = [c for c in codes(diagnose(self.project, "prefinal")) if c.startswith("ANNOTATION_PAGES")]
            self.assertEqual([], codes_found, sentence)


class RegularProjectTest(ProjectCase):
    state = "regular"
    config = {"profile": "mpgu-09-regular", "intensity": "balanced"}

    @classmethod
    def build_state(cls, project: Path) -> None:
        fill_project(project, product=False, regular=True)
        validate_and_snapshot(project)

    def test_regular_without_product_accepts_evd_tec_not_applicable(self) -> None:
        na = {"status": "not_applicable", "summary": "нет продукта", "not_applicable_reason": "обычная ВКР без программного продукта", "findings": []}
        run_wave(self.project, "primary", "final", "P", reports={"EVD": na, "TEC": na}, expect_close=0)
        run_wave(self.project, "blind_regression", "final", "B", reports={"EVD": na, "TEC": na}, expect_close=0)
        report = diagnose(self.project)
        self.assertEqual("PASS", report["status"], codes(report, "ERROR"))
        self.assertEqual("READY_TO_SUBMIT", report["readiness"])


@unittest.skipUnless(HAVE_DOCX, "python-docx не установлен")
class FingerprintTest(unittest.TestCase):
    def test_fingerprint_survives_word_field_update_and_metadata_cleaning(self) -> None:
        with tempfile.TemporaryDirectory(prefix="vkr633-fp-") as temporary:
            project = Path(temporary) / "project"
            shutil.copytree(str(FIXTURE), str(project))
            rebuild(project)
            manifest = read_json(project / common.BUILD_MANIFEST_PATH)
            built, error = common.docx_text_fingerprint(project / "final" / "vkr.docx")
            self.assertIsNone(error)
            self.assertEqual(built, manifest["docx"]["text_fingerprint"])
            self.assertEqual(built, common.docx_text_fingerprint(WORD_SAVED)[0])
            cleaned = Path(temporary) / "cleaned.docx"
            shutil.copyfile(str(WORD_SAVED), str(cleaned))
            code, _payload, _out, err = run_cli("clean_docx_metadata.py", cleaned, "--in-place", "--json")
            self.assertEqual(0, code, err)
            self.assertEqual(built, common.docx_text_fingerprint(cleaned)[0])
            shutil.copyfile(str(WORD_SAVED), str(project / "final" / "vkr.docx"))
            self.assertEqual([], common.build_manifest_problems(project))
            paths = [item["path"] for item in manifest["inputs"]]
            self.assertIn("evidence/figures/architecture.png", paths)
            self.assertIn("vkr-project.json", paths)
            (project / "drafts" / "appendix-3.md").write_text("# Приложение\n\nНовое приложение.\n", encoding="utf-8")
            self.assertTrue(any("появился после сборки" in item for item in common.build_manifest_problems(project)))


class FingerprintAreasTest(unittest.TestCase):
    """V3b п. 2: правки в Word вне тела абзацев меняют отпечаток и дают FINAL_DOCX_OUTDATED."""

    NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

    @classmethod
    def setUpClass(cls) -> None:
        if not HAVE_DOCX:
            raise unittest.SkipTest("python-docx не установлен")
        cls.temp = tempfile.TemporaryDirectory(prefix="vkr633-fparea-")
        cls.project = Path(cls.temp.name) / "project"
        shutil.copytree(str(FIXTURE), str(cls.project))
        rebuild(cls.project)
        shutil.copyfile(str(WORD_SAVED), str(cls.project / "final" / "vkr.docx"))
        cls.baseline = read_json(cls.project / common.BUILD_MANIFEST_PATH)["docx"]["text_fingerprint"]

    @classmethod
    def tearDownClass(cls) -> None:
        cls.temp.cleanup()

    def fingerprint(self, name: str, **changes: Any) -> str:
        target = rewrite_docx(WORD_SAVED, Path(self.temp.name) / f"{name}.docx", **changes)
        value, error = common.docx_text_fingerprint(target)
        self.assertIsNone(error, name)
        return str(value)

    def assert_changed(self, name: str, **changes: Any) -> None:
        self.assertNotEqual(self.baseline, self.fingerprint(name, **changes), name)

    def assert_same(self, name: str, **changes: Any) -> None:
        self.assertEqual(self.baseline, self.fingerprint(name, **changes), name)

    def test_baseline_matches_build(self) -> None:
        self.assertEqual([], common.build_manifest_problems(self.project))

    def test_hidden_existing_text(self) -> None:
        def hide(xml: str) -> str:
            start = xml.rindex("<w:rPr>", 0, xml.index("(Ф.И.О.)")) + len("<w:rPr>")
            return xml[:start] + "<w:vanish/>" + xml[start:]

        self.assert_changed("hidden", edits={"word/document.xml": hide})

    def test_textbox_text_gives_outdated_docx(self) -> None:
        box = (
            '<w:p><w:r><w:pict><v:shape id="box1" style="width:200pt;height:40pt"><v:textbox><w:txbxContent>'
            "<w:p><w:r><w:t>Пространственное мышление — врождённая способность</w:t></w:r></w:p>"
            "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>"
        )
        self.assert_changed("textbox", edits={"word/document.xml": lambda xml: before_body_end(xml, box)})
        shutil.copyfile(str(Path(self.temp.name) / "textbox.docx"), str(self.project / "final" / "vkr.docx"))
        try:
            self.assertTrue(any("не совпадает" in item for item in common.build_manifest_problems(self.project)))
        finally:
            shutil.copyfile(str(WORD_SAVED), str(self.project / "final" / "vkr.docx"))

    def test_footnote_and_endnote_text(self) -> None:
        note = '<w:footnote w:id="9"><w:p><w:r><w:t>Определение уточнено: см. с. 5 первоисточника.</w:t></w:r></w:p></w:footnote>'
        self.assert_changed("footnote", edits={"word/footnotes.xml": lambda xml: xml.replace("</w:footnotes>", note + "</w:footnotes>")})
        endnote = note.replace("w:footnote", "w:endnote")
        self.assert_changed("endnote", edits={"word/endnotes.xml": lambda xml: xml.replace("</w:endnotes>", endnote + "</w:endnotes>")})

    def test_header_footer_and_page_number_results(self) -> None:
        self.assert_changed("footer", edits={"word/footer1.xml": lambda xml: xml.replace("</w:p></w:ftr>", "<w:r><w:t> Черновик</w:t></w:r></w:p></w:ftr>")})
        self.assert_same("footer-page-number", edits={"word/footer1.xml": lambda xml: xml.replace("<w:t>2</w:t>", "<w:t>17</w:t>")})
        self.assert_changed("footer-page-text", edits={"word/footer1.xml": lambda xml: xml.replace("<w:t>2</w:t>", "<w:t>Черновик</w:t>")})
        header = f"<w:hdr {self.NS}><w:p><w:r><w:t>Лебедева А.С. Выпускная работа</w:t></w:r></w:p></w:hdr>"
        relation = '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header9.xml"/>'
        self.assert_changed(
            "header",
            edits={
                "word/_rels/document.xml.rels": lambda xml: xml.replace("</Relationships>", relation + "</Relationships>"),
                "word/document.xml": lambda xml: xml.replace('<w:footerReference w:type="default"', '<w:headerReference w:type="default" r:id="rId99"/><w:footerReference w:type="default"'),
            },
            extra={"word/header9.xml": header},
        )

    def test_comment_text(self) -> None:
        comments = f'<w:comments {self.NS}><w:comment w:id="0" w:author="Студент"><w:p><w:r><w:t>уточнить у руководителя</w:t></w:r></w:p></w:comment></w:comments>'
        relation = '<Relationship Id="rId98" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>'
        self.assert_changed("comment", edits={"word/_rels/document.xml.rels": lambda xml: xml.replace("</Relationships>", relation + "</Relationships>")}, extra={"word/comments.xml": comments})

    def test_replaced_picture(self) -> None:
        self.assert_changed("picture", edits={"word/media/image1.png": lambda data: data + bytes([0])})

    def test_toc_styled_paragraph_and_page_field_text(self) -> None:
        toc = '<w:p><w:pPr><w:pStyle w:val="14"/></w:pPr><w:r><w:t>Новый абзац в стиле оглавления: пилот показал рост на 40 %.</w:t></w:r></w:p>'
        self.assert_changed("toc-style", edits={"word/document.xml": lambda xml: before_body_end(xml, toc)})
        field = (
            '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>{}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
        )
        self.assert_changed("page-field-text", edits={"word/document.xml": lambda xml: before_body_end(xml, field.format("Пилот показал рост на 40 %."))})
        self.assert_same("page-field-number", edits={"word/document.xml": lambda xml: before_body_end(xml, field.format("12"))})


class DoctorStageTest(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="vkr633-doctor-")
        self.base = Path(self.temp.name)
        self.project = create_project(self.base, profile="generic", intensity="balanced")

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_fresh_init_passes_draft(self) -> None:
        report = doctor.diagnose(self.project, "draft")
        self.assertEqual("PASS", report["status"], codes(report, "ERROR"))

    def test_project_config_not_object_fails_without_traceback(self) -> None:
        (self.project / "vkr-project.json").write_text("[]", encoding="utf-8")
        code, payload, _out, err = run_cli("vkr_project_doctor.py", self.project, "--stage", "prefinal", "--json")
        self.assertEqual(1, code, err)
        self.assertNotIn("Traceback", err)
        self.assertIn("PROJECT_CONFIG_SHAPE", codes(payload))

    def test_planned_tasks_null_does_not_crash(self) -> None:
        manifest = read_json(self.project / "audit" / "manifest.json")
        manifest["runs"] = [{"run_id": "x", "planned_seq": 1, "planned_tasks": None}]
        write_json(self.project / "audit" / "manifest.json", manifest)
        code, payload, _out, err = run_cli("vkr_project_doctor.py", self.project, "--stage", "final", "--json")
        self.assertEqual(1, code, err)
        self.assertNotIn("Traceback", err)
        self.assertIn("AUDIT_RUN_TAMPERED", codes(payload))
        self.assertNotIn("INTERNAL_ERROR", codes(payload))

    def test_non_utf8_draft_is_reported_with_path(self) -> None:
        (self.project / "drafts" / "chapter-1.md").write_bytes("# Глава\n\nТекст в cp1251".encode("cp1251"))
        report = doctor.diagnose(self.project, "prefinal")
        self.assertEqual("drafts/chapter-1.md", finding_of(report, "ENCODING_INVALID")["path"])

    def test_pointer_drafts_and_chapter_three_rule(self) -> None:
        for path in (self.project / "drafts").glob("*.md"):
            path.write_text("Текст в vkr.docx\n", encoding="utf-8")
        (self.project / "drafts" / "chapter-3.md").unlink()
        report = doctor.diagnose(self.project, "prefinal")
        self.assertIn("drafts/introduction.md", [item["path"] for item in report["findings"] if item["code"] == "DRAFT_TOO_SHORT"])
        self.assertNotIn("drafts/chapter-3.md", [item["path"] for item in report["findings"] if item["code"] == "DRAFT_REQUIRED_MISSING"])

    def test_markers_use_shared_module_and_skip_code(self) -> None:
        body = (
            "Работа посвящена проектированию информационной системы сопровождения практики студентов. " * 5
            + "\n\nВКР о ToDo-приложении и XXX Олимпиаде; индекс `arr[x]`.\n\n```vue\n<template>{{ user.name }} TODO</template>\n```\n"
        )
        for name in ("introduction", "chapter-1", "chapter-2", "conclusion", "annotation"):
            (self.project / "drafts" / f"{name}.md").write_text(f"# {name}\n\n{body}", encoding="utf-8")
        (self.project / "drafts" / "chapter-3.md").unlink()
        (self.project / "plan.md").write_text("# План\n\nЦель и задачи.\n", encoding="utf-8")
        self.assertNotIn("PLACEHOLDER_FOUND", codes(doctor.diagnose(self.project, "prefinal")))
        chapter = self.project / "drafts" / "chapter-1.md"
        chapter.write_text(chapter.read_text(encoding="utf-8") + "\nИсточник [ПРОВЕРИТЬ: год].\n", encoding="utf-8")
        self.assertIn("PLACEHOLDER_CHECK_PENDING", codes(doctor.diagnose(self.project, "prefinal"), "WARNING"))
        self.assertIn("PLACEHOLDER_FOUND", codes(doctor.diagnose(self.project, "final"), "ERROR"))

    def test_state_incomplete_has_concrete_action_and_optional_fields_are_dash(self) -> None:
        state = (self.project / "vkr-state.md").read_text(encoding="utf-8")
        self.assertIn("- **Институт/факультет:** Институт математики и информатики", state)
        write_json(self.base / "sparse.json", {"profile": "generic", "title_page": {"author": "Иванов И.И."}})
        sparse = self.base / "sparse"
        init_project.initialize(sparse, init_project.normalize_config(init_project.load_config(self.base / "sparse.json")))
        text = (sparse / "vkr-state.md").read_text(encoding="utf-8")
        self.assertIn("- **Кафедра:** не указано", text)  # кафедру сборка ставит на титул (V3b п. 8)
        self.assertIn("- **Вуз:** —", text)  # вуз сборка подставляет по умолчанию
        self.assertIn("- **Правила кафедры об использовании ИИ:** —", text)  # ключа ai_rules нет в intake
        self.assertIn("- **Научный руководитель:** не указано", text)
        report = doctor.diagnose(sparse, "prefinal")
        self.assertTrue(any("Научный руководитель" in action for action in report["next_actions"]), report["next_actions"])


class LegacyMigrationTest(unittest.TestCase):
    def test_632_manifest_migrates_and_legacy_reports_are_warnings(self) -> None:
        with tempfile.TemporaryDirectory(prefix="vkr633-legacy-") as temporary:
            project = create_project(Path(temporary), profile="generic", intensity="balanced")
            legacy = {"protocol_version": "6.30", "snapshot_manifest_sha256": "x", "runs": [
                {"run_id": "final-primary", "status": "complete", "planned_tasks": [{"task_id": "primary-A-SRC", "base_gate": "SRC", "status": "fail", "report_path": "audit/reports/primary-A-SRC.json"}]},
            ]}
            write_json(project / "audit" / "manifest.json", legacy)
            write_json(project / "audit" / "reports" / "primary-A-SRC.json", {"status": "fail", "findings": [{"severity": "BLOCKER"}]})
            report = doctor.diagnose(project, "prefinal")
            self.assertIn("AUDIT_LEGACY_RUNS", codes(report, "WARNING"))
            self.assertIn("AUDIT_LEGACY_REPORTS", codes(report, "WARNING"))
            self.assertNotIn("AUDIT_REPORT_ORPHAN", codes(report))
            self.assertNotIn("AUDIT_RUN_TAMPERED", codes(report))
            code, payload = audit_cmd(project, "migrate-legacy", "--json")
            self.assertEqual(0, code, payload)
            self.assertEqual("audit/legacy/reports/primary-A-SRC.json", payload["moved"][0]["to"])
            manifest = read_json(project / "audit" / "manifest.json")
            self.assertEqual("vkr-audit-manifest", manifest["schema"])
            self.assertEqual(1, len(manifest["legacy_runs"]))
            report = doctor.diagnose(project, "prefinal")
            self.assertNotIn("AUDIT_LEGACY_REPORTS", codes(report))
            self.assertIn("AUDIT_LEGACY_RUNS", codes(report, "WARNING"))
            manifest["legacy_runs"][0]["status"] = "pass"
            write_json(project / "audit" / "manifest.json", manifest)
            self.assertIn("AUDIT_RUN_TAMPERED", codes(doctor.diagnose(project, "prefinal"), "ERROR"))


    def test_632_flat_title_fields_move_to_title_page(self) -> None:
        with tempfile.TemporaryDirectory(prefix="vkr633-legacy-title-") as temporary:
            project = create_project(Path(temporary), profile="generic", intensity="balanced")
            config = read_json(project / "vkr-project.json")
            config.pop("title_page")
            config.update({
                "student_name": "Иванов Иван Иванович", "institution": "Московский педагогический государственный университет",
                "institute": "Институт математики и информатики", "program_code": "09.03.02",
                "program_name": "Информационные системы и технологии", "supervisor": "доцент, канд. пед. наук П.П. Петров",
            })
            write_json(project / "vkr-project.json", config)
            report = doctor.diagnose(project, "prefinal")
            self.assertIn("title_page", finding_of(report, "TITLE_PAGE_INCOMPLETE")["details"])
            self.assertTrue(any("migrate-legacy" in action for action in report["next_actions"]), report["next_actions"])
            code, payload = audit_cmd(project, "migrate-legacy", "--json")
            self.assertEqual(0, code, payload)
            self.assertTrue(payload["title_page_migrated"])
            migrated = read_json(project / "vkr-project.json")
            self.assertNotIn("student_name", migrated)
            self.assertEqual("Иванов Иван Иванович", migrated["title_page"]["author"])
            self.assertEqual("Московский педагогический государственный университет", migrated["title_page"]["university"])
            expected = ["program_profile", "supervisor_title", "department", "head_title", "head_of_department"]
            self.assertEqual(expected, payload["title_page_missing"])
            self.assertEqual(expected, finding_of(doctor.diagnose(project, "prefinal"), "TITLE_PAGE_INCOMPLETE")["details"])
            code, payload = audit_cmd(project, "migrate-legacy", "--json")
            self.assertEqual((0, False), (code, payload["title_page_migrated"]))


class InitTest(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="vkr633-init-")
        self.base = Path(self.temp.name)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_relative_path_inside_skill_is_refused(self) -> None:
        name = "_wp1-relative-project-must-be-refused"
        code, payload, _out, err = run_cli("init_vkr_project.py", name, "--json", cwd=SKILL)
        self.assertEqual(2, code, (payload, err))
        self.assertFalse((SKILL / name).exists())

    def test_relative_path_outside_skill_gives_absolute_root(self) -> None:
        code, payload, _out, err = run_cli("init_vkr_project.py", "relative-project", "--json", cwd=self.base)
        self.assertEqual(0, code, err)
        self.assertTrue(Path(payload["project_root"]).is_absolute())

    def test_profile_type_mismatch_is_refused(self) -> None:
        write_json(self.base / "bad.json", {"profile": "mpgu-09-project", "vkr_type": "regular"})
        code, payload, _out, _err = run_cli("init_vkr_project.py", self.base / "p", "--config", self.base / "bad.json", "--json")
        self.assertEqual(2, code)
        self.assertFalse((self.base / "p").exists())

    def test_config_validation(self) -> None:
        for bad in ({"collective": "false"}, {"collective": 2}, {"typo_profile": "x"}, {"title_page": {"autor": "x"}}, {"deadline": "13.09.2026"}):
            path = self.base / "bad.json"
            write_json(path, bad)
            with self.assertRaises(ValueError, msg=bad):
                init_project.normalize_config(init_project.load_config(path))
        path = self.base / "express.json"
        write_json(path, {"mode": "express-4d", "collective": False}, bom=True)
        self.assertEqual("balanced", init_project.normalize_config(init_project.load_config(path))["audit_intensity"])
        write_json(path, {"profile": "mpgu-09-regular"})
        self.assertEqual("regular", init_project.normalize_config(init_project.load_config(path))["vkr_type"])

    def test_mpgu_project_state_template_and_snapshot_inputs(self) -> None:
        project = create_project(self.base, profile="mpgu-09-project")
        index = read_json(project / "evidence" / "index.json")
        self.assertEqual(common.METHODOLOGY_ID, index["methodology"][0]["id"])
        self.assertEqual("PASS", doctor.diagnose(project, "draft")["status"])
        self.assertEqual("ok", memory.status(project, 5)["status"])
        snapshot, _problems = common.build_snapshot(project, "mpgu-09-project")
        kinds = {item["logical_id"]: item["kind"] for item in snapshot["inputs"]}
        self.assertEqual("methodology", kinds[common.METHODOLOGY_PROJECT_PATH])
        self.assertEqual("state_projection", kinds[common.STATE_LOGICAL_ID])
        self.assertEqual("vkr-audit-manifest", read_json(project / "audit" / "manifest.json")["schema"])

    def test_ai_rules_in_state_template_doctor_and_snapshot(self) -> None:
        rules = "ИИ — только для проверки орфографии;  использование указать\nво введении"
        project = create_project(self.base, profile="mpgu-09-project", ai_rules=rules)
        state = (project / "vkr-state.md").read_text(encoding="utf-8")
        line = "- **Правила кафедры об использовании ИИ:** ИИ — только для проверки орфографии; использование указать во введении"
        requirements = state.split("## Требования и структура", 1)[1].split("\n## ", 1)[0]
        self.assertIn(line, requirements)
        note = state.split("\n## ", 1)[0]  # подсказка шаблона называет исключённые разделы точно (V3), без ключевых слов
        for title in common.STATE_EXCLUDED_TITLES[:4]:
            self.assertIn(f"«{title}»", note)
        self.assertNotIn("«сесси»", note)
        self.assertEqual("PASS", doctor.diagnose(project, "draft")["status"])
        self.assertEqual("ok", memory.status(project, 5)["status"])
        self.assertIn(line, common.state_projection(state))

        def state_sha() -> str:
            snapshot, _problems = common.build_snapshot(project, "mpgu-09-project")
            return {item["logical_id"]: item["sha256"] for item in snapshot["inputs"]}[common.STATE_LOGICAL_ID]

        before = state_sha()
        (project / "vkr-state.md").write_text(state.replace(line, line + "; уточнено у руководителя"), encoding="utf-8")
        self.assertNotEqual(before, state_sha())  # правила ИИ — содержание state и входят в снимок
        empty = create_project(self.base, name="empty-rules", profile="generic", ai_rules="")
        text = (empty / "vkr-state.md").read_text(encoding="utf-8")
        self.assertIn("- **Правила кафедры об использовании ИИ:** —", text)
        report = doctor.diagnose(empty, "prefinal")
        self.assertEqual([], [item["code"] for item in report["findings"] if "использовании ИИ" in json.dumps(item, ensure_ascii=False)])
        for bad in (["запрет"], 1, True, {"rule": "x"}):
            write_json(self.base / "bad.json", {"ai_rules": bad})
            with self.assertRaises(ValueError, msg=bad):
                init_project.normalize_config(init_project.load_config(self.base / "bad.json"))

    def test_state_template_has_requirements_profile_fields(self) -> None:  # B13 (F02)
        precedence = (SKILL / "references" / "requirements-precedence.md").read_text(encoding="utf-8")
        section = precedence.split("## Как использовать этот профиль", 1)[1]
        fields = [name for pair in re.findall(r"^- `([a-z_]+)`(?: и `([a-z_]+)`)?:", section, re.MULTILINE) for name in pair if name]
        self.assertTrue({"requirements_source", "pilot_min", "page_target", "source_target", "implementation_certificate_required"} <= set(fields), fields)
        labels = {"profile": "**Профиль требований:**", "vkr_type": "**Тип ВКР:**", "ai_rules": "**Правила кафедры об использовании ИИ:**"}
        project = create_project(self.base, profile="mpgu-09-project")
        state = (project / "vkr-state.md").read_text(encoding="utf-8")
        pattern = (SKILL / "references" / "state-file-pattern.md").read_text(encoding="utf-8").split("## Шаблон state-файла", 1)[1]
        for text in (state, pattern):
            for field in fields:
                self.assertIn(labels.get(field, f"`{field}`"), text, field)
        requirements = state.split("## Требования и структура", 1)[1].split("\n## ", 1)[0]
        self.assertIn("- **`pilot_min`:** not_set", requirements)
        self.assertNotIn("не указано", requirements)
        report = doctor.diagnose(project, "prefinal")
        self.assertFalse(any("requirements_source" in json.dumps(item, ensure_ascii=False) for item in report["findings"]))

    def test_reinit_adopts_legacy_state_and_reports_conflict(self) -> None:
        project = self.base / "legacy"
        project.mkdir()
        (project / "vkr-state.md").write_text("# VKR State\n\n- **Профиль требований:** `mpgu-09-regular`\n- **Режим:** express-7d\n", encoding="utf-8")
        result = init_project.initialize(project, init_project.normalize_config(init_project.load_config(None)))
        self.assertEqual("initialized", result["status"], result)
        self.assertEqual("regular", read_json(project / "vkr-project.json")["vkr_type"])
        write_json(self.base / "conflict.json", {"profile": "generic"})
        result = init_project.initialize(project, init_project.normalize_config(init_project.load_config(self.base / "conflict.json")))
        self.assertEqual("conflict", result["status"])


class MemoryTest(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="vkr633-memory-")
        self.base = Path(self.temp.name)
        self.project = create_project(self.base, profile="generic")

    def tearDown(self) -> None:
        self.temp.cleanup()

    def event(self, **values: Any) -> Path:
        path = self.base / "event.json"
        write_json(path, {"event_type": "checkpoint", "summary": "Завершён подраздел", **values}, bom=True)
        return path

    def test_new_files_in_audit_reports_and_root_are_detected(self) -> None:
        self.assertEqual("ok", memory.status(self.project, 5)["status"])
        write_json(self.project / "audit" / "reports" / "R001" / "R001-LOG-A.json", {"x": 1})
        (self.project / "notes.txt").write_text("заметка", encoding="utf-8")
        status = memory.status(self.project, 5)
        self.assertEqual({"audit/reports/R001/R001-LOG-A.json", "notes.txt"}, {item["logical_id"] for item in status["artifact_changes"]["untracked"]})
        self.assertIn("audit/findings.json", status["continuation_files"])

    def test_legacy_index_needs_rebaseline(self) -> None:
        write_json(self.project / "memory" / "artifact-index.json", {"schema_version": 1, "artifacts": {"memory/handoff.md": {"sha256": "0" * 64}}})
        status = memory.status(self.project, 5)
        self.assertEqual("needs_rebaseline", status["status"])
        code, payload, _out, err = run_cli("vkr_memory.py", self.project, "record", "--rebaseline", "--json")
        self.assertEqual(0, code, err)
        self.assertEqual("ok", memory.status(self.project, 5)["status"])

    def test_record_is_transactional_and_retry_does_not_duplicate(self) -> None:
        (self.project / "drafts" / "chapter-1.md").write_text("# Глава\n\nНовый текст\n", encoding="utf-8")
        event = self.event(files=["drafts/chapter-1.md"], decisions=[{"decision": "Оставить две метрики"}])
        original = memory.append_jsonl

        def failing(path: Path, item: Dict[str, Any]) -> None:
            if path.name == "activity.jsonl":
                raise OSError("диск заполнен")
            original(path, item)

        with mock.patch.object(memory, "append_jsonl", failing):
            with self.assertRaises(OSError):
                memory.record(self.project, event)
        self.assertEqual("recovery_required", memory.status(self.project, 5)["status"])
        self.assertEqual("recorded", memory.record(self.project, event)["status"])
        self.assertEqual("already_recorded", memory.record(self.project, event)["status"])
        self.assertEqual(1, len((self.project / "memory" / "decisions.jsonl").read_text(encoding="utf-8").splitlines()))

    def test_all_changed_records_status_changes_and_ignores_event_file_in_project(self) -> None:
        # Отчёты audit/reports/* пишет vkr_audit.py: агент их не перечисляет.
        code, payload = audit_cmd(self.project, "snapshot", "--json")
        self.assertEqual(0, code, payload)
        plan = run_wave(self.project, "primary", "plan", "M", expect_close=0)
        chapter = self.project / "drafts" / "chapter-1.md"
        chapter.write_text(chapter.read_text(encoding="utf-8") + "\nНовый абзац.\n", encoding="utf-8")
        (self.project / "drafts" / "chapter-2.md").unlink()
        (self.project / "notes.txt").write_text("заметка", encoding="utf-8")
        before = memory.status(self.project, 5)
        self.assertEqual("external_changes_detected", before["status"])
        expected = {item["logical_id"] for key in ("changed", "untracked") for item in before["artifact_changes"][key]}
        self.assertTrue({"drafts/chapter-1.md", "notes.txt"} <= expected, expected)
        self.assertEqual(5, len([name for name in expected if name.startswith(f"audit/reports/{plan['run_id']}/")]), expected)
        # Event-файл в корне проекта — вход record, а не внешнее изменение.
        event = self.project / "event.json"
        write_json(event, {"event_type": "audit_run_closed", "summary": "Аудит плана закрыт", "files": ["plan.md"]}, bom=True)
        code, payload, _out, err = run_cli("vkr_memory.py", self.project, "record", "--event", event, "--all-changed", "--json")
        self.assertEqual(0, code, err)
        self.assertEqual("recorded", payload["status"])
        self.assertEqual(expected | {"plan.md"}, set(payload["tracked_artifacts"]))  # поле files события учитывается вместе с флагом
        self.assertEqual(["drafts/chapter-2.md"], payload["removed"])
        after = memory.status(self.project, 5)
        self.assertEqual("ok", after["status"], after["artifact_changes"])
        self.assertEqual(["event.json"], after["artifact_changes"]["ignored_event_files"])
        self.assertNotIn("event.json", read_json(self.project / "memory" / "artifact-index.json")["artifacts"])
        code, again, _out, err = run_cli("vkr_memory.py", self.project, "record", "--event", event, "--all-changed", "--json")
        self.assertEqual((0, "already_recorded", payload["event_id"]), (code, again["status"], again["event_id"]), err)
        # Следующее событие в том же файле тоже не внешнее изменение; повтор команды не дублирует запись.
        write_json(event, {"event_type": "checkpoint", "summary": "Глава I начата", "decisions": [{"decision": "Начать с 1.1"}]})
        self.assertEqual("ok", memory.status(self.project, 5)["status"])
        self.assertEqual("recorded", memory.record(self.project, event, all_changed=True)["status"])
        self.assertEqual("already_recorded", memory.record(self.project, event, all_changed=True)["status"])
        self.assertEqual(3, len((self.project / "logs" / "activity.jsonl").read_text(encoding="utf-8").splitlines()))
        self.assertEqual(1, len((self.project / "memory" / "decisions.jsonl").read_text(encoding="utf-8").splitlines()))
        # Без --all-changed event-файл, перечисленный в files, тоже не индексируется.
        (self.project / "plan.md").write_text("# План\n\nОбновлён.\n", encoding="utf-8")
        write_json(event, {"event_type": "checkpoint", "summary": "План обновлён", "files": ["plan.md", "event.json"]})
        recorded = memory.record(self.project, event)
        self.assertEqual(["plan.md"], sorted(recorded["tracked_artifacts"]))
        self.assertEqual("ok", memory.status(self.project, 5)["status"])

    def test_all_changed_retry_after_interruption_does_not_duplicate(self) -> None:
        (self.project / "drafts" / "chapter-1.md").write_text("# Глава\n\nНовый текст\n", encoding="utf-8")
        write_json(self.project / "audit" / "reports" / "R009" / "R009-LOG-A.json", {"status": "pass"})
        event = self.event(decisions=[{"decision": "Оставить две метрики"}])
        original = memory.append_jsonl

        def failing(path: Path, item: Dict[str, Any]) -> None:
            if path.name == "activity.jsonl":
                raise OSError("диск заполнен")
            original(path, item)

        with mock.patch.object(memory, "append_jsonl", failing):
            with self.assertRaises(OSError):
                memory.record(self.project, event, all_changed=True)
        interrupted = read_json(self.project / "memory" / "handoff.json")["last_event_id"]
        self.assertEqual("recovery_required", memory.status(self.project, 5)["status"])
        retried = memory.record(self.project, event, all_changed=True)
        self.assertEqual("recorded", retried["status"])
        self.assertEqual(interrupted, retried["event_id"])
        self.assertEqual(["audit/reports/R009/R009-LOG-A.json", "drafts/chapter-1.md"], sorted(retried["tracked_artifacts"]))
        self.assertEqual("already_recorded", memory.record(self.project, event, all_changed=True)["status"])
        self.assertEqual(1, len((self.project / "memory" / "decisions.jsonl").read_text(encoding="utf-8").splitlines()))
        self.assertEqual("ok", memory.status(self.project, 5)["status"])

    def test_all_changed_usage_errors(self) -> None:
        event = self.event()
        for args in (("--all-changed", "--rebaseline", "--event", event), ("--all-changed",)):
            code, payload, _out, err = run_cli("vkr_memory.py", self.project, "record", *args, "--json")
            self.assertEqual(2, code, (args, err))
            self.assertEqual("error", payload["status"])
        write_json(self.project / "memory" / "artifact-index.json", {"schema_version": 1, "artifacts": {}})
        code, payload, _out, err = run_cli("vkr_memory.py", self.project, "record", "--event", event, "--all-changed", "--json")
        self.assertEqual(2, code, err)
        self.assertIn("--rebaseline", payload["error"])

    def test_bad_index_record_gives_json_error_without_traceback(self) -> None:
        write_json(self.project / "memory" / "artifact-index.json", {"schema_version": 2, "artifacts": {"plan.md": "abc"}})
        code, payload, _out, err = run_cli("vkr_memory.py", self.project, "status", "--json")
        self.assertEqual(2, code)
        self.assertNotIn("Traceback", err)


class OwnQuestionsTest(unittest.TestCase):
    """B12 (F08, F17): вопросы комиссии в отчёте OWN и задание OWN до финала."""

    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="vkr633-own-")
        self.base = Path(self.temp.name)
        self.project = create_project(self.base, profile="generic", intensity="strict")

    def tearDown(self) -> None:
        self.temp.cleanup()

    def planned(self, *args: str) -> Tuple[Dict[str, Any], Dict[str, str]]:
        code, plan = audit_cmd(self.project, "plan", "--kind", "primary", *args, "--json")
        self.assertEqual(0, code, plan)
        code, brief = audit_cmd(self.project, "brief", "--run", plan["run_id"], "--json")
        self.assertEqual(0, code, brief)
        texts = {f"{item['auditor_role']}-{item['replica_id']}": Path(item["absolute_path"]).read_text(encoding="utf-8") for item in brief["briefs"]}
        return plan, texts

    def record(self, plan: Dict[str, Any], key: str, report: Dict[str, Any]) -> Tuple[int, Dict[str, Any]]:
        task = next(task for task in plan["tasks"] if f"{task['auditor_role']}-{task['replica_id']}" == key)
        self.reports = getattr(self, "reports", 0) + 1
        path = report_file(self.base, f"{plan['run_id']}-{key}-{self.reports}", report)
        return audit_cmd(self.project, "record", "--run", plan["run_id"], "--task", task["task_id"], "--auditor-id", f"aud-{key}", "--report", path, "--json")

    def test_questions_field_is_validated_recorded_and_requested_by_brief(self) -> None:
        plan, texts = self.planned("--preset", "chapter", "--checkpoint", "chapter-1")
        self.assertIn('"questions"', texts["OWN-QUESTIONS-A"])
        self.assertIn("Test-30", texts["OWN-QUESTIONS-A"])
        self.assertNotIn('"questions"', texts["LOG-A"])
        base = {"status": "pass", "summary": "Вопросы подготовлены", "findings": []}
        for bad in ({"questions": "Почему Unity?"}, {"questions": [{"question": ""}]}, {"questions": [{"question": "Почему?", "answer": "x"}]}, {"questions": [7]}):
            code, payload = self.record(plan, "OWN-QUESTIONS-A", dict(base, **bad))
            self.assertEqual((1, "REPORT_SCHEMA_INVALID"), (code, payload.get("code")), bad)
        code, payload = self.record(plan, "LOG-A", dict(base, questions=["Почему выбран метод?"]))
        self.assertEqual(1, code, payload)
        self.assertIn("только в задачах роли OWN", payload["error"])
        questions = ["Почему выбран Unity?", {"question": "Как оценивалась эффективность?", "criterion": "называет метрику, выборку и ограничения"}]
        code, payload = self.record(plan, "OWN-QUESTIONS-A", dict(base, questions=questions))
        self.assertEqual(0, code, payload)
        self.assertEqual(2, payload["questions_recorded"])
        canonical = read_json(self.project / payload["report_path"])
        self.assertEqual(
            [{"question": "Почему выбран Unity?"}, {"question": "Как оценивалась эффективность?", "criterion": "называет метрику, выборку и ограничения"}],
            canonical["questions"],
        )
        self.assertEqual(0, audit_cmd(self.project, "findings", "--json")[0])
        self.assertNotIn("questions", read_json(self.project / self.record(plan, "LNG-A", base)[1]["report_path"]))

    def test_own_full_gate_before_final_checks_test30_readiness_without_fail(self) -> None:
        plan, texts = self.planned("--preset", "prefinal")
        own = texts["OWN-A"]
        self.assertIn("отсутствие само по себе не `fail`", own)
        self.assertIn('"questions"', own)
        self.assertNotIn("верни `fail` с замечанием BLOCKER", own)
        code, payload = self.record(plan, "OWN-A", {
            "status": "pass", "summary": "Материалы для ответов есть, вопросы подготовлены", "findings": [],
            "questions": [{"question": "Откуда взяты 22 учащихся?", "criterion": "ссылается на журнал пилота"}],
        })
        self.assertEqual(0, code, payload)
        self.assertEqual("pass", payload["task_status"])


class SubsectionStyParityTest(unittest.TestCase):
    """B11 (F10): STY «в каждом втором подразделе» считает plan по завершённым подразделам той же главы."""

    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="vkr633-sty-")
        self.project = create_project(Path(self.temp.name), profile="generic", intensity="strict")

    def tearDown(self) -> None:
        self.temp.cleanup()

    def roles(self, checkpoint: str, preset: str = "subsection") -> List[str]:
        code, plan = audit_cmd(self.project, "plan", "--kind", "primary", "--preset", preset, "--checkpoint", checkpoint, "--json")
        self.assertEqual(0, code, plan)
        self.plan = plan
        return sorted({task["auditor_role"] for task in plan["tasks"]})

    def abandon(self) -> None:
        code, closed = audit_cmd(self.project, "close", "--run", self.plan["run_id"], "--abandon", "--reason", "перепланирование", "--json")
        self.assertEqual(0, code, closed)

    def test_parity_by_completed_subsections_of_the_same_chapter(self) -> None:
        # 3.1 проверен пресетом product (подраздел с продуктом) — это первый подраздел главы III.
        run_wave(self.project, "primary", "product", "P31", extra=("--checkpoint", "subsection-3.1"), expect_close=0)
        self.assertIn("STY", self.roles("subsection-3.2"))  # второй подраздел главы III
        self.abandon()
        run_wave(self.project, "primary", "subsection", "S32", extra=("--checkpoint", "subsection-3.2"), expect_close=0)
        self.assertNotIn("STY", self.roles("subsection-2.1"))  # первый подраздел главы II
        self.abandon()
        run_wave(self.project, "primary", "subsection", "S21", extra=("--checkpoint", "subsection-2.1"), expect_close=0)
        self.assertIn("STY", self.roles("subsection-2.2"))
        self.abandon()  # прерванный run подраздела не засчитывается
        self.assertNotIn("STY", self.roles("subsection-1.1"))
        self.abandon()
        self.assertNotIn("STY", self.roles("subsection-1.2"))
        self.assertEqual([], self.plan["warnings"])
        self.abandon()
        roles = self.roles("section-without-number")
        self.assertTrue(any("subsection-<глава>.<номер>" in warning for warning in self.plan["warnings"]), self.plan["warnings"])
        self.assertNotIn("STY", roles)  # без номера главы: завершённых run пресета subsection два (3.2, 2.1) — чётный


class HandoffStaleTest(unittest.TestCase):
    """B10 (NC-1): чат оборвался после главы II — handoff отстаёт от файлов и журнала аудита."""

    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="vkr633-handoff-")
        self.base = Path(self.temp.name)
        self.project = create_project(self.base, profile="generic")

    def tearDown(self) -> None:
        self.temp.cleanup()

    def checkpoint(self, name: str, **values: Any) -> Dict[str, Any]:
        path = self.base / f"{name}.json"
        write_json(path, {"event_type": "checkpoint", "summary": name, **values})
        return memory.record(self.project, path, all_changed=True)

    def test_break_after_chapter_two(self) -> None:
        self.assertFalse(memory.status(self.project, 0)["handoff_stale"])
        (self.project / "drafts" / "chapter-3.md").write_text("# Глава III. Проект\n\n" + "Описание тренажёра. " * 40, encoding="utf-8")
        self.assertEqual("recorded", self.checkpoint(
            "chapter-3", current_phase="chapter-2", current_task="Написание главы II",
            next_actions=["написать 2.1 «Характеристика среды внедрения»", "аудит subsection-2.1"],
        )["status"])
        fresh = memory.status(self.project, 0)
        self.assertEqual(("ok", False, ["продолжай по handoff"]), (fresh["status"], fresh["handoff_stale"], fresh["next_actions"]))

        # Глава II написана и прошла аудит, checkpoint памяти не записан: чат оборвался.
        (self.project / "drafts" / "chapter-2.md").write_text(
            "# Глава II. Анализ\n\n## 2.1. Среда внедрения\n\n" + "Характеристика школы и учащихся. " * 40, encoding="utf-8",
        )
        (self.project / "evidence" / "pilot" / "school-context.md").write_text("Контекст школы.\n", encoding="utf-8")
        wave = run_wave(self.project, "primary", "subsection", "S21", extra=("--checkpoint", "subsection-2.1"), expect_close=0)

        status = memory.status(self.project, 0)
        self.assertEqual("external_changes_detected", status["status"])
        self.assertTrue(status["handoff_stale"])
        self.assertIn("написать 2.1 «Характеристика среды внедрения»", status["handoff"]["next_actions"])
        reasons = " ".join(status["handoff_stale_reasons"])
        self.assertIn("drafts/chapter-2.md", reasons)
        self.assertIn(wave["run_id"], reasons)
        first = status["next_actions"][0]
        for needle in ("Handoff устарел", "не продолжай по его next_actions", "vkr-state.md", "audit/manifest.json",
                       "record --event <event.json> --all-changed", "current_task", "next_actions"):
            self.assertIn(needle, first)
        code, payload, out, err = run_cli("vkr_memory.py", self.project, "status", "--json")
        self.assertEqual((1, True), (code, payload["handoff_stale"]), err)
        self.assertEqual(first, payload["next_actions"][0])
        code, _payload, text, err = run_cli("vkr_memory.py", self.project, "status")
        self.assertIn("Handoff: устарел", text, err)

        code, nxt = audit_cmd(self.project, "next", "--stage", "draft", "--json")
        self.assertEqual(0, code, nxt)
        self.assertTrue(nxt["handoff_stale"])
        self.assertTrue(nxt["next_actions"][0].startswith("Handoff устарел"), nxt["next_actions"])
        self.assertFalse(any(action.startswith("Запиши checkpoint памяти") for action in nxt["next_actions"]), nxt["next_actions"])

        # Восстановление: checkpoint --all-changed с актуальным handoff снимает флаг.
        self.checkpoint("chapter-2", current_phase="chapter-1", current_task="Написание главы I", next_actions=["написать 1.1"])
        status = memory.status(self.project, 0)
        self.assertEqual(("ok", False), (status["status"], status["handoff_stale"]))
        self.assertEqual(read_json(self.project / "audit" / "manifest.json")["seq"], read_json(self.project / "memory" / "handoff.json")["audit_seq"])
        self.assertFalse(audit_cmd(self.project, "next", "--stage", "draft", "--json")[1]["handoff_stale"])

    def test_audit_actions_after_handoff_make_it_stale_without_file_changes(self) -> None:
        self.checkpoint("start", next_actions=["аудит intake"])
        self.assertFalse(memory.status(self.project, 0)["handoff_stale"])
        code, plan = audit_cmd(self.project, "plan", "--kind", "primary", "--preset", "intake", "--json")
        self.assertEqual(0, code, plan)
        status = memory.status(self.project, 0)
        self.assertEqual("ok", status["status"])  # задания и снимки аудита не отслеживаются как файлы проекта
        self.assertTrue(status["handoff_stale"])
        self.assertIn(plan["run_id"], " ".join(status["handoff_stale_reasons"]))
        self.assertTrue(status["next_action"].startswith("handoff устарел"))


class StateProjectionTest(unittest.TestCase):
    def test_projection_excludes_exact_service_sections_only(self) -> None:
        fence = "`" * 3
        text = (
            BOM + "# State\r\n> Последнее обновление: 01.09.2026\r\n\r\n## Паспорт\r\n- Тема: X\r\n\r\n"
            "## Независимый аудит\r\n- Итог: FAIL\r\n### Подраздел\r\n- run\r\n\r\n"
            "## Handoff и продолжение\r\n- next\r\n\r\n## Целевая аудитория продукта\r\n- школьники\r\n\r\n"
            "## Стилевой профиль\r\n" + fence + "\r\n## История сессий\r\n" + fence + "\r\n"
        )
        projection = common.state_projection(text)
        self.assertIn("## Паспорт", projection)
        self.assertIn("## Целевая аудитория продукта", projection)
        self.assertIn("## История сессий", projection)
        self.assertNotIn("Итог: FAIL", projection)
        self.assertNotIn("next", projection)
        self.assertNotIn("Последнее обновление", projection)
        changed = text.replace("Итог: FAIL", "Итог: READY_TO_SUBMIT").replace("01.09.2026", "20.09.2026")
        self.assertEqual(projection, common.state_projection(changed))
        self.assertNotEqual(projection, common.state_projection(text.replace("школьники", "студенты")))


if __name__ == "__main__":
    unittest.main(verbosity=2)
