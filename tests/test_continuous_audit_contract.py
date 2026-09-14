"""Контракт документации аудита: роли, матрица, команды и финальный gate.

Проверяет только документы и код WP1 (continuous-audit, multi-agent-audit,
quality-control-loop, project-memory, project-initialization, state-file-pattern)
и не зависит от архивов прежних версий.
"""

from __future__ import annotations

import os
import re
import sys
import unittest
from pathlib import Path

os.environ["PYTHONDONTWRITEBYTECODE"] = "1"
sys.dont_write_bytecode = True

ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "vkr-mpgu"
SCRIPTS = SKILL / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

import vkr_audit as audit  # noqa: E402
import vkr_common as common  # noqa: E402

ROLES = set(common.BASE_GATES)
DOCS = (
    "references/continuous-audit.md",
    "references/multi-agent-audit.md",
    "references/quality-control-loop.md",
    "references/project-memory.md",
    "references/project-initialization.md",
    "references/state-file-pattern.md",
)


def read(relative: str) -> str:
    return (SKILL / relative).read_text(encoding="utf-8")


class ContinuousAuditContractTest(unittest.TestCase):
    def setUp(self) -> None:
        self.cadence = read("references/continuous-audit.md")
        self.protocol = read("references/multi-agent-audit.md")
        self.quality = read("references/quality-control-loop.md")
        self.memory = read("references/project-memory.md")
        self.initialization = read("references/project-initialization.md")
        self.state = read("references/state-file-pattern.md")

    def require(self, text: str, needles, label: str) -> None:
        flat = " ".join(text.split())
        missing = [needle for needle in needles if " ".join(needle.split()) not in flat]
        self.assertFalse(missing, f"{label}: нет {missing}")

    def test_versions_and_scripts(self) -> None:
        self.assertEqual("6.33-production", common.SKILL_VERSION)
        self.assertEqual("6.33", common.PROTOCOL_VERSION)
        for name in ("vkr_common.py", "vkr_audit.py", "vkr_project_doctor.py", "vkr_memory.py", "init_vkr_project.py", "vkr_markers.py"):
            self.assertTrue((SCRIPTS / name).is_file(), name)

    def test_nine_roles_and_lenses(self) -> None:
        protocol_roles = set(re.findall(r"^\| `([A-Z]{3})` \|", self.protocol, re.MULTILINE))
        self.assertEqual(ROLES, protocol_roles)
        for role in ROLES:
            self.assertIn(f"`{role}`", self.cadence)
        for role, (gate, lens) in common.SPECIALIZED_ROLES.items():
            self.assertIn(f"| `{role}` | `{gate}` | `{lens}` |", self.cadence)
        self.assertIn("| базовая роль (`MET` … `OWN`) | та же | `full_gate` |", self.cadence)

    def test_matrix_and_presets_are_documented(self) -> None:
        self.require(
            self.cadence,
            [
                "Intake завершён", "План сформирован или изменён", "Подраздел завершён", "Новая или заменённая ссылка",
                "Закончена глава", "Закончено описание пилота", "Собран DOCX", "Основной агент внёс пакет правок",
                "Перед передачей научруку", "Финальный DOCX", "balanced", "strict", "maximum",
            ],
            "матрица",
        )
        for preset in audit.PRESETS:
            self.assertIn(f"--preset {preset}", self.cadence, preset)
        for kind in common.AUDIT_KINDS:
            self.assertIn(f"--kind {kind}", self.cadence, kind)

    def test_documents_describe_commands_not_manual_bookkeeping(self) -> None:
        for command in ("validate", "snapshot", "plan", "brief", "record", "close", "findings", "accept-risk", "waive", "next"):
            self.assertRegex(self.cadence + self.protocol, rf"\$A {re.escape(command)}\b|vkr_audit\.py[^\n]*\b{re.escape(command)}\b|`{re.escape(command)}`", command)
        combined = "\n".join(read(relative) for relative in DOCS)
        for manual in ("обязательны положительный номер `attempt`", '"execution_id": "..."', "last_heartbeat", "замороженную копию"):
            self.assertNotIn(manual, combined)
        self.assertNotRegex(combined, r"\b[\w./-]+\.(?:md|py):\d+")

    def test_final_gate_rules(self) -> None:
        self.require(
            self.cadence,
            [
                "full_gate", "final/vkr.docx", "USER_EVIDENCE_REQUIRED", "READY_FOR_SUPERVISOR_REVIEW", "READY_TO_SUBMIT",
                "degraded_independence", "не пишется в `memory/claims-register.json`", "**до** снимка",
                "AUDIT_PRIMARY_WAVE_MISSING", "AUDIT_BLIND_WAVE_MISSING", "AUDIT_SNAPSHOT_STALE", "VALIDATION_ERRORS",
                "Когда субъективные замечания не сходятся", "Сколько это стоит", "evidence/defense/",
            ],
            "финальный gate",
        )
        self.require(self.quality, ["Сохранить законченный материал", "Запустить doctor повторно", "final/vkr.docx", "**до** снимка"], "quality loop")

    def test_read_only_auditor_boundary(self) -> None:
        self.require(
            self.protocol,
            ["они никогда не редактируют ВКР", "Только основной агент", "не возвращает\nпереписанный абзац, патч или изменённый файл"],
            "read-only",
        )

    def test_memory_and_initialization_contract(self) -> None:
        self.require(
            self.initialization + self.memory,
            [
                "init_vkr_project.py", "vkr_memory.py", "Пользователь не должен вручную", "handoff.json", "artifact-index.json",
                "audit/findings.json", "needs_rebaseline", "record --rebaseline", "title_page", "express-*",
                "methodology-09-03-02-2024", "evidence_ids", "source_ids",
            ],
            "память и init",
        )

    def test_state_template_keeps_cliche_check_enabled(self) -> None:
        self.assertIn("cliche_allowlist: []", self.state)
        # Тот же разбор, что у validate_vkr: список под незакомментированным ключом.
        active = re.search(r"^\s*cliche_allowlist\s*:\s*\n((?:\s+-\s+.+\n)+)", self.state, re.MULTILINE)
        self.assertIsNone(active)
        inline = re.findall(r"^\s*cliche_allowlist\s*:\s*\[(.*?)\]\s*$", self.state, re.MULTILINE)
        self.assertTrue(all(not item.strip() for item in inline), inline)
        for title in common.STATE_EXCLUDED_TITLES[:4]:
            self.assertIn(title, self.state)
        self.assertNotIn("воспроизводимости тайминга", self.state)

    def test_no_bytecode_in_skill(self) -> None:
        self.assertFalse([path for path in SKILL.rglob("*.pyc")])


if __name__ == "__main__":
    unittest.main(verbosity=2)
