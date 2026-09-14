"""Рабочая структура проекта: init, память и doctor через CLI.

Сквозной путь до финального gate — в ``test_core_audit_633.py``.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any, Dict, Tuple

os.environ["PYTHONDONTWRITEBYTECODE"] = "1"
sys.dont_write_bytecode = True

ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "vkr-mpgu"
SCRIPTS = SKILL / "scripts"
INIT = SCRIPTS / "init_vkr_project.py"
MEMORY = SCRIPTS / "vkr_memory.py"
DOCTOR = SCRIPTS / "vkr_project_doctor.py"
AUDIT = SCRIPTS / "vkr_audit.py"


def run(script: Path, *args: Any, expected: int = 0, cwd: Path = ROOT) -> Dict[str, Any]:
    env = os.environ.copy()
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    result = subprocess.run(
        [sys.executable, str(script), *map(str, args)],
        cwd=str(cwd), env=env, text=True, encoding="utf-8", capture_output=True,
    )
    assert result.returncode == expected, (script.name, result.returncode, result.stdout, result.stderr)
    assert "Traceback" not in result.stderr, result.stderr
    value = json.loads(result.stdout)
    assert isinstance(value, dict), value
    return value


class AutoWorkspaceTest(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="vkr-auto-workspace-")
        self.base = Path(self.temp.name)
        self.project = self.base / "project"
        self.config = self.base / "intake.json"
        self.config.write_text(json.dumps({
            "student_name": "Иванов Иван Иванович",
            "institution": "МПГУ",
            "institute": "ИМО",
            "program_code": "09.03.02",
            "program_name": "Информационные системы и технологии",
            "supervisor": "Петрова А.Б.",
            "topic": "Информационная система сопровождения практики",
            "vkr_type": "project",
            "profile": "mpgu-09-project",
            "mode": "standard",
            "audit_intensity": "strict",
            "deadline": "2027-06-01",
        }, ensure_ascii=False), encoding="utf-8")

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_initialization_creates_complete_workspace(self) -> None:
        initialized = run(INIT, self.project, "--config", self.config, "--json")
        self.assertEqual("initialized", initialized["status"])
        self.assertEqual("6.33-production", initialized["skill_version"])
        for relative in (
            "vkr-state.md", "vkr-project.json", "plan.md", "sources.json", "evidence/index.json",
            "audit/snapshot-inputs.json", "audit/manifest.json", "audit/findings.json",
            "memory/handoff.json", "memory/handoff.md", "memory/roadmap.md", "memory/decisions.jsonl",
            "memory/claims-register.json", "memory/artifact-index.json", "logs/activity.jsonl",
            "logs/tool-runs.jsonl", "drafts/chapter-3.md", "sources/materials/methodology-09-03-02-2024.docx",
        ):
            self.assertTrue((self.project / relative).is_file(), relative)
        for relative in ("evidence/defense", "evidence/figures", "audit/briefs", "audit/snapshots", "backups/checkpoints"):
            self.assertTrue((self.project / relative).is_dir(), relative)
        project = json.loads((self.project / "vkr-project.json").read_text(encoding="utf-8"))
        self.assertEqual("Иванов Иван Иванович", project["title_page"]["author"])
        self.assertEqual("МПГУ", project["title_page"]["university"])
        index = json.loads((self.project / "memory/artifact-index.json").read_text(encoding="utf-8"))
        self.assertEqual(2, index["schema_version"])
        self.assertIn("vkr-state.md", index["artifacts"])
        self.assertIn("drafts/chapter-1.md", index["artifacts"])

        draft = run(DOCTOR, self.project, "--stage", "draft", "--json")
        self.assertEqual("PASS", draft["status"], draft)
        self.assertEqual("ok", run(MEMORY, self.project, "status", "--json")["status"])

    def test_rerun_preserves_user_files(self) -> None:
        run(INIT, self.project, "--config", self.config, "--json")
        sources = '[{"id": "s1", "title": "Реальный источник", "status": "pending"}]\n'
        (self.project / "sources.json").write_text(sources, encoding="utf-8")
        rerun = run(INIT, self.project, "--config", self.config, "--json")
        self.assertEqual("initialized", rerun["status"])
        self.assertIn("sources.json", rerun["skipped_existing"])
        self.assertEqual(sources, (self.project / "sources.json").read_text(encoding="utf-8"))
        conflict = run(INIT, self.project, "--profile", "generic", "--json", expected=2)
        self.assertEqual("conflict", conflict["status"])

    def test_memory_checkpoint_and_external_changes(self) -> None:
        run(INIT, self.project, "--config", self.config, "--json")
        event_path = self.base / "event.json"
        event_path.write_text(json.dumps({
            "event_type": "subsection_completed",
            "summary": "Завершён подраздел и проверены исходные данные",
            "current_phase": "chapter-1",
            "current_task": "Исправить замечания аудитора",
            "last_completed": "План подраздела 1.1",
            "next_actions": ["Запустить профильный аудит"],
            "files": ["vkr-state.md", "sources.json"],
            "decisions": [{"decision": "Сохранять источники с отдельным статусом проверки"}],
        }, ensure_ascii=False), encoding="utf-8")
        recorded = run(MEMORY, self.project, "record", "--event", event_path, "--json")
        self.assertEqual("recorded", recorded["status"])
        self.assertTrue(recorded["tracked_artifacts"]["sources.json"]["sha256"])
        handoff = json.loads((self.project / "memory/handoff.json").read_text(encoding="utf-8"))
        self.assertTrue(handoff["last_event_id"].startswith("evt-"))
        self.assertEqual(2, len((self.project / "logs/activity.jsonl").read_text(encoding="utf-8").splitlines()))
        self.assertEqual(1, len((self.project / "memory/decisions.jsonl").read_text(encoding="utf-8").splitlines()))

        state = self.project / "vkr-state.md"
        state.write_text(state.read_text(encoding="utf-8") + "\nВнешнее изменение\n", encoding="utf-8")
        status = run(MEMORY, self.project, "status", "--json", expected=1)
        self.assertEqual("external_changes_detected", status["status"])
        before = (self.project / "logs/activity.jsonl").read_text(encoding="utf-8")
        outside = self.base / "outside.json"
        outside.write_text(json.dumps({"event_type": "bad", "summary": "bad", "files": ["../outside.txt"]}), encoding="utf-8")
        error = run(MEMORY, self.project, "record", "--event", outside, "--json", expected=2)
        self.assertEqual("error", error["status"])
        self.assertEqual(before, (self.project / "logs/activity.jsonl").read_text(encoding="utf-8"))

    def test_doctor_reports_broken_json_and_unfinished_drafts(self) -> None:
        run(INIT, self.project, "--config", self.config, "--json")
        (self.project / "sources.json").write_text("{", encoding="utf-8")
        broken = run(DOCTOR, self.project, "--stage", "draft", "--json", expected=1)
        self.assertEqual("FAIL", broken["status"])
        self.assertIn("JSON_INVALID", {item["code"] for item in broken["findings"]})
        (self.project / "sources.json").write_text("[]", encoding="utf-8")
        prefinal = run(DOCTOR, self.project, "--stage", "prefinal", "--json", expected=1)
        codes = {item["code"] for item in prefinal["findings"]}
        self.assertIn("PLACEHOLDER_FOUND", codes)
        self.assertIn("DRAFT_TOO_SHORT", codes)
        self.assertEqual("NOT_READY", prefinal["readiness"])
        self.assertTrue(prefinal["next_actions"])
        final = run(DOCTOR, self.project, "--stage", "final", "--json", expected=1)
        self.assertIn("FINAL_DOCX_MISSING", {item["code"] for item in final["findings"]})
        nxt = run(AUDIT, self.project, "next", "--stage", "draft", "--json")
        self.assertEqual("ok", nxt["status"])

    def test_regular_type_and_skill_directory_protection(self) -> None:
        regular = self.base / "regular"
        run(INIT, regular, "--profile", "mpgu-09-regular", "--json")
        self.assertFalse((regular / "drafts/chapter-3.md").exists())
        self.assertEqual("regular", json.loads((regular / "vkr-project.json").read_text(encoding="utf-8"))["vkr_type"])
        refused = SKILL / "_test-project-must-be-refused"
        run(INIT, refused, "--dry-run", "--json", expected=2)
        self.assertFalse(refused.exists())
        run(INIT, "_relative-project-must-be-refused", "--json", expected=2, cwd=SKILL)
        self.assertFalse((SKILL / "_relative-project-must-be-refused").exists())


if __name__ == "__main__":
    unittest.main(verbosity=2)
