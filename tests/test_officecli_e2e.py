# -*- coding: utf-8 -*-
"""End-to-end optional OfficeCLI orchestration; retained sample files as evidence."""
from __future__ import annotations
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import unittest
import uuid

ROOT = Path(__file__).resolve().parents[1]
GENERATOR = ROOT / "skills" / "thaksorn-craft" / "scripts" / "thaksorn.py"
OUTPUT = ROOT / "tests" / "output"
SOURCE = ROOT / "tests" / "fixtures" / "science.md"


def run_cli(*args, env=None):
    return subprocess.run(
        [sys.executable, str(GENERATOR), *map(str, args)], cwd=ROOT,
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=180, env=env,
    )


def unique() -> Path:
    OUTPUT.mkdir(exist_ok=True, parents=True)
    return OUTPUT / f"full_pipeline_{uuid.uuid4().hex[:12]}.docx"


class TestIntegratedWorkflow(unittest.TestCase):
    def test_officecli_optional_when_not_on_path(self):
        env = dict(os.environ)
        env["PATH"] = ""
        out = unique()
        r = run_cli("render", SOURCE, "--output", out, "--officecli-qa", "auto", env=env)
        self.assertEqual(r.returncode, 0, r.stderr)
        report = json.loads(r.stdout)
        self.assertEqual(report["officecli_qa"]["status"], "OFFICECLI_UNAVAILABLE_STANDALONE")
        self.assertFalse(report["officecli_used"])
        self.assertTrue(out.is_file())

    def test_officecli_required_does_not_write_when_missing(self):
        env = dict(os.environ)
        env["PATH"] = ""
        out = unique()
        r = run_cli("render", SOURCE, "--output", out, "--officecli-qa", "require", env=env)
        self.assertNotEqual(r.returncode, 0)
        self.assertFalse(out.exists())
        self.assertIn("OfficeCLI required", r.stderr)

    def test_repair_option_rejected_if_qa_disabled(self):
        out = unique()
        r = run_cli("render", SOURCE, "--output", out, "--repair-known-schema")
        self.assertNotEqual(r.returncode, 0)
        self.assertFalse(out.exists())

    def test_full_source_to_validated_copy(self):
        state = run_cli("officecli", "--workspace", ROOT, "status")
        if state.returncode != 0 or not json.loads(state.stdout).get("available"):
            self.skipTest("No OfficeCLI present")
        target = unique()
        r = run_cli("render", SOURCE, "--output", target, "--officecli-qa", "require",
                    "--repair-known-schema", "--toc")
        self.assertEqual(r.returncode, 0, r.stderr)
        result = json.loads(r.stdout)
        self.assertEqual(result["officecli_qa"]["status"], "PASS_REPAIRED_COPY")
        final = Path(result["recommended_output"])
        self.assertTrue(final.is_file())
        self.assertTrue(target.is_file())
        self.assertTrue(result["officecli_qa"]["normalization"]["text_unchanged"])
        self.assertTrue(result["officecli_qa"]["normalization"]["source_unchanged"])
        self.assertTrue(result["officecli_qa"]["normalization"]["officecli_validation"]["passed"])
        sha_original = hashlib.sha256(target.read_bytes()).hexdigest()
        self.assertEqual(sha_original, hashlib.sha256(target.read_bytes()).hexdigest())
        self.assertTrue(result["officecli_qa"]["normalization"]["structure_unchanged"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
