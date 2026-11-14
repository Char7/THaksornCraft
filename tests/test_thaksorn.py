# -*- coding: utf-8 -*-
"""Standalone, no-OfficeCLI smoke tests. Outputs remain inside tests/output."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import unittest
import uuid
import zipfile

ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "skills" / "thaksorn-craft"
CLI = SKILL / "scripts" / "thaksorn.py"
FIXTURES = ROOT / "tests" / "fixtures"
OUTPUT = ROOT / "tests" / "output"


def call(*args: str):
    return subprocess.run(
        [sys.executable, str(CLI), *[str(s) for s in args]],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=120,
    )


def unique_out(ext: str) -> Path:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    return OUTPUT / f"thaksorn-{uuid.uuid4().hex[:12]}.{ext}"


class TestSkillBundle(unittest.TestCase):
    def test_skill_metadata(self):
        text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
        self.assertTrue(text.startswith("---\nname: thaksorn-craft\n"))
        self.assertIn("OfficeCLI", text)
        for item in ("workflow.md", "compatibility.md", "quality-gates.md", "licenses.md"):
            self.assertTrue((SKILL / "references" / item).is_file(), item)

    def test_vendor_and_licenses(self):
        for path in (
            "vendor/kien-thai/LICENSE",
            "vendor/kien-thai/skills/kien-thai/SKILL.md",
            "vendor/kien-thai/skills/kode-thai/SKILL.md",
            "vendor/thai-doc-skills/LICENSE",
            "vendor/thai-doc-skills/thai-docx/SKILL.md",
            "vendor/thai-doc-skills/thai-pdf/SKILL.md",
            "vendor/thai-doc-skills/thai-docx/fonts/LICENSE-NATIONAL-FONTS.txt",
        ):
            self.assertTrue((SKILL / path).is_file(), path)

    def test_doctor_without_officecli(self):
        r = call("doctor", "--json")
        self.assertEqual(r.returncode, 0, r.stderr)
        state = json.loads(r.stdout)
        self.assertEqual(state["product"], "THaksornCraft")
        self.assertFalse(state["officecli_required"])
        self.assertTrue(all(state["components"].values()))
        self.assertGreaterEqual(state["bundled_national_font_files"], 0)

    def test_reject_complex_markdown_instead_of_corrupting(self):
        source = unique_out("md")
        source.write_text("| หัวข้อ | คำตอบ |\n|---|---|\n| ก | ข |\n", encoding="utf-8")
        target = unique_out("docx")
        r = call("render", str(source), "--output", str(target))
        if importlib.util.find_spec("docx") is None:
            self.skipTest("python-docx is not installed")
        self.assertNotEqual(r.returncode, 0)
        self.assertIn("Markdown tables", r.stderr)
        self.assertFalse(target.exists())

    def test_reject_out_of_scope_image_reference(self):
        source = unique_out("json")
        source.write_text(json.dumps([
            {"type": "image", "path": "../outside.png", "width": 5}
        ], ensure_ascii=False), encoding="utf-8")
        target = unique_out("docx")
        if importlib.util.find_spec("docx") is None:
            self.skipTest("python-docx is not installed")
        r = call("render", str(source), "--output", str(target))
        self.assertNotEqual(r.returncode, 0)
        self.assertIn("image must be within", r.stderr)
        self.assertFalse(target.exists())

    @unittest.skipUnless(importlib.util.find_spec("docx") is not None, "python-docx not installed")
    def test_render_markdown_word_and_guard_overwrite(self):
        target = unique_out("docx")
        cmd = ("render", str(FIXTURES / "science.md"), "--output", str(target),
               "--toc", "--page-number", "footer-center")
        r = call(*cmd)
        self.assertEqual(r.returncode, 0, r.stderr)
        meta = json.loads(r.stdout)
        self.assertFalse(meta["officecli_used"])
        self.assertTrue(target.is_file())
        with zipfile.ZipFile(target) as z:
            names = z.namelist()
            xml = z.read("word/document.xml").decode("utf-8")
            self.assertIn("w:cs=", xml)
            self.assertIn("สารบัญ", xml)
            self.assertIn("word/fontTable.xml", names)
            if list((SKILL / "vendor/thai-doc-skills/thai-docx/fonts/national").glob("*.ttf")):
                self.assertTrue(meta["font_embedded"])
        second = call(*cmd)
        self.assertNotEqual(second.returncode, 0)
        self.assertIn("refusing to overwrite", second.stderr)

    @unittest.skipUnless(importlib.util.find_spec("docx") is not None, "python-docx not installed")
    def test_render_structured_word_tables(self):
        target = unique_out("docx")
        r = call("render", str(FIXTURES / "structured.json"), "--output", str(target))
        self.assertEqual(r.returncode, 0, r.stderr)
        with zipfile.ZipFile(target) as z:
            xml = z.read("word/document.xml").decode("utf-8")
            self.assertIn("<w:tbl>", xml)
            self.assertIn("w:cs=", xml)
            self.assertIn("Type=", z.read("word/_rels/document.xml.rels").decode("utf-8"))

    def test_pdf_if_runtime_supports_it(self):
        doctor = call("doctor", "--probe-pdf")
        state = json.loads(doctor.stdout)
        if state["pdf_ready"] is not True:
            self.skipTest("WeasyPrint/Pango not available in current environment")
        target = unique_out("pdf")
        r = call("render", str(FIXTURES / "science.md"), "--output", str(target))
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertTrue(target.read_bytes().startswith(b"%PDF"))
        state = json.loads(r.stdout)
        self.assertTrue(state["font_embedded"], "Bundled Thai font should embed")


    @unittest.skipUnless(sys.platform == "win32" and importlib.util.find_spec("win32com") is not None,
                         "Microsoft Word COM PDF backend unavailable")
    def test_pdf_using_word_backend_without_officecli(self):
        pdf = unique_out("pdf")
        r = call("render", str(FIXTURES / "structured.json"), "--output", str(pdf),
                 "--pdf-backend", "word", "--page-number", "footer-center")
        self.assertEqual(r.returncode, 0, r.stderr)
        state = json.loads(r.stdout)
        self.assertFalse(state["officecli_used"])
        self.assertTrue(pdf.read_bytes().startswith(b"%PDF"))
        self.assertTrue(Path(state["source_docx_companion"]).is_file())

    @unittest.skipUnless(sys.platform == "win32" and importlib.util.find_spec("win32com") is not None
                         and importlib.util.find_spec("docx") is not None,
                         "Microsoft Word and python-docx unavailable")
    def test_export_existing_word_keeps_source_unchanged(self):
        import hashlib
        docx = unique_out("docx")
        created = call("render", str(FIXTURES / "structured.json"), "--output", str(docx))
        self.assertEqual(created.returncode, 0, created.stderr)
        previous_hash = hashlib.sha256(docx.read_bytes()).hexdigest()
        pdf = unique_out("pdf")
        result = call("export-word", str(docx), "--output", str(pdf))
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(previous_hash, hashlib.sha256(docx.read_bytes()).hexdigest())
        self.assertTrue(pdf.read_bytes().startswith(b"%PDF"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
