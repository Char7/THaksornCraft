# -*- coding: utf-8 -*-
"""THaksornCraft optional OfficeCLI adapter QA.

All generated inputs/outputs are inside this single project. Test files are
retained as evidence; this suite never deletes or overwrites existing files.
"""
from __future__ import annotations
import hashlib
import json
import pathlib
import subprocess
import sys
import unittest
import uuid

ROOT = pathlib.Path(__file__).resolve().parents[1]
SKILL = ROOT / "skills" / "thaksorn-craft"
ADAPTER = SKILL / "scripts" / "officecli_adapter.py"
GENERATOR = SKILL / "scripts" / "thaksorn.py"
OUT = ROOT / "tests" / "output"
SOURCE = ROOT / "tests" / "fixtures" / "science.md"
MISSING_BINARY = ROOT / "tests" / "no-such-officecli.exe"


def run_adapter(*args: str, binary: str | None = None):
    cmd = [sys.executable, str(ADAPTER), "--workspace", str(ROOT)]
    if binary is not None:
        cmd += ["--binary", binary]
    return subprocess.run([*cmd, *map(str, args)], cwd=ROOT, capture_output=True,
                          text=True, encoding="utf-8", errors="replace", timeout=120)


def generated_path(label: str) -> pathlib.Path:
    OUT.mkdir(parents=True, exist_ok=True)
    return OUT / f"bridge_{label}_{uuid.uuid4().hex[:12]}.docx"


def prepare() -> pathlib.Path:
    target = generated_path("source")
    cmd = [sys.executable, str(GENERATOR), "render", str(SOURCE),
           "--output", str(target), "--no-embed"]
    run = subprocess.run(cmd, cwd=ROOT, capture_output=True,
                         text=True, encoding="utf-8", errors="replace", timeout=90)
    if run.returncode:
        raise RuntimeError(run.stderr)
    return target


def sha(path: pathlib.Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class TestOfficeCLIBridge(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.status = run_adapter("status")
        if cls.status.returncode == 0:
            cls.detect = json.loads(cls.status.stdout)
        else:
            cls.detect = {"available": False}

    def test_missing_binary_graceful_standalone(self):
        r = run_adapter("status", binary=str(MISSING_BINARY))
        self.assertEqual(r.returncode, 0, r.stderr)
        state = json.loads(r.stdout)
        self.assertFalse(state["available"])
        self.assertFalse(state["officecli_required"])

    def test_officecli_installed_or_report_optional(self):
        r = self.status
        self.assertEqual(r.returncode, 0, r.stderr)
        state = json.loads(r.stdout)
        self.assertIn("available", state)
        self.assertFalse(state["officecli_required"])

    def test_outside_workspace_blocked(self):
        r = run_adapter("inspect", str(ROOT.parent / "outside-project.txt"))
        self.assertNotEqual(r.returncode, 0)
        self.assertIn("OUTSIDE_WORKSPACE", r.stderr)

    def test_no_officecli_basic_read_fallback(self):
        source = prepare()
        r = run_adapter("inspect", str(source), "--mode", "text",
                        "--allow-basic-fallback", binary=str(MISSING_BINARY))
        self.assertEqual(r.returncode, 0, r.stderr)
        state = json.loads(r.stdout)
        self.assertFalse(state["officecli_used"])
        self.assertEqual(state["status"], "BASIC_READ_ONLY")
        self.assertIn("วิทยาศาสตร์", state["text"])

    def test_missing_officecli_does_not_create_output(self):
        source = prepare()
        output = generated_path("unavailable")
        r = run_adapter("replace-copy", str(source), "--output", str(output),
                        "--find", "วิทยาศาสตร์", "--replace", "วิทยาศาสตร์พื้นฐาน",
                        binary=str(MISSING_BINARY))
        self.assertNotEqual(r.returncode, 0)
        self.assertFalse(output.exists())
        self.assertTrue(source.exists())

    def test_missing_text_does_not_create_output(self):
        if not self.detect.get("available"):
            self.skipTest("OfficeCLI not installed")
        source = prepare()
        output = generated_path("nohit")
        r = run_adapter("replace-copy", str(source), "--output", str(output),
                        "--find", "คำที่ไม่มีทางพบ_xyz", "--replace", "ข้อความ")
        self.assertNotEqual(r.returncode, 0)
        self.assertFalse(output.exists())

    def test_dry_run_never_writes(self):
        if not self.detect.get("available"):
            self.skipTest("OfficeCLI not installed")
        source = prepare()
        target = generated_path("dry")
        r = run_adapter("replace-copy", str(source), "--output", str(target),
                        "--find", "วิทยาศาสตร์", "--replace", "วิทยาศาสตร์พื้นฐาน",
                        "--dry-run")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertEqual(json.loads(r.stdout)["expected_matches"], 1)
        self.assertFalse(target.exists())

    def test_live_read(self):
        if not self.detect.get("available"):
            self.skipTest("OfficeCLI not installed")
        source = prepare()
        r = run_adapter("inspect", str(source), "--mode", "text")
        self.assertEqual(r.returncode, 0, r.stderr)
        state = json.loads(r.stdout)
        self.assertTrue(state["officecli_used"])
        self.assertTrue(state["result"]["success"])

    def test_live_validate_reports_verdict(self):
        if not self.detect.get("available"):
            self.skipTest("OfficeCLI not installed")
        source = prepare()
        r = run_adapter("validate", str(source))
        self.assertIn(r.returncode, (0, 3), r.stderr)
        state = json.loads(r.stdout)
        self.assertTrue(state["officecli_used"])
        self.assertIn(state["status"], ("PASS", "SCHEMA_ISSUES"))

    def test_repair_generated_word_schema_copy(self):
        if not self.detect.get("available"):
            self.skipTest("OfficeCLI not installed")
        source = prepare()
        original = sha(source)
        normalized = generated_path("normalized")
        r = run_adapter("repair-schema-copy", str(source), "--output", str(normalized))
        self.assertEqual(r.returncode, 0, r.stderr)
        state = json.loads(r.stdout)
        self.assertEqual(state["status"], "PASS", r.stdout)
        self.assertTrue(state["officecli_validation"]["passed"])
        self.assertTrue(state["source_unchanged"])
        self.assertTrue(state["structure_unchanged"])
        self.assertTrue(state["text_unchanged"])
        self.assertEqual(original, sha(source))
        self.assertTrue(normalized.is_file())
        again = run_adapter("validate", str(normalized))
        self.assertEqual(again.returncode, 0, again.stdout + again.stderr)

    def test_repair_schema_without_officecli(self):
        source = prepare()
        target = generated_path("normalized_offline")
        original = sha(source)
        r = run_adapter("repair-schema-copy", str(source), "--output", str(target),
                        binary=str(MISSING_BINARY))
        self.assertEqual(r.returncode, 0, r.stderr)
        state = json.loads(r.stdout)
        self.assertEqual(state["status"], "NORMALIZED_NEEDS_SCHEMA_VALIDATION")
        self.assertFalse(state["officecli_used"])
        self.assertTrue(state["text_unchanged"])
        self.assertEqual(original, sha(source))

    def test_main_entrypoint_officecli_status(self):
        r = subprocess.run(
            [sys.executable, str(GENERATOR), "officecli", "--workspace", str(ROOT), "status"],
            cwd=ROOT, capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=60,
        )
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn("officecli_required", json.loads(r.stdout))

    def test_live_copy_edit_preserves_source(self):
        if not self.detect.get("available"):
            self.skipTest("OfficeCLI not installed")
        source = prepare()
        before = sha(source)
        target = generated_path("edited")
        r = run_adapter("replace-copy", str(source), "--output", str(target),
                        "--find", "วิทยาศาสตร์", "--replace", "วิทยาศาสตร์พื้นฐาน")
        self.assertEqual(r.returncode, 0, r.stderr)
        state = json.loads(r.stdout)
        self.assertEqual(state["status"], "PASS")
        self.assertEqual(state["expected_matches"], 1)
        self.assertTrue(state["source_unchanged"])
        self.assertTrue(state["structure_unchanged"])
        self.assertTrue(state["text_changed_exactly_as_requested"])
        self.assertEqual(sha(source), before)
        self.assertTrue(target.is_file())
        # Attempting the exact same output twice must never overwrite.
        second = run_adapter("replace-copy", str(source), "--output", str(target),
                             "--find", "วิทยาศาสตร์", "--replace", "วิทยาศาสตร์พื้นฐาน")
        self.assertNotEqual(second.returncode, 0)
        self.assertIn("already exists", second.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
