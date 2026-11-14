#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Optional OfficeCLI bridge for THaksornCraft.

No OfficeCLI import, installation or modification. All document paths must stay
inside the explicitly supplied workspace (--workspace, default current directory).
Mutation is only performed on a NEW output copy; originals are untouched.
Agent preflight: read the relevant thai-docx/thai-pdf SKILL.md before invoking.
"""
from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile

SKILL_ROOT = Path(__file__).resolve().parents[1]
NS_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
SUPPORTED = {".docx", ".xlsx", ".pptx"}


def emit(data: dict) -> None:
    print(json.dumps(data, ensure_ascii=False, indent=2))


class BridgeError(Exception):
    pass


def workspace_root(args: argparse.Namespace) -> Path:
    root = Path(args.workspace).expanduser().resolve(strict=True)
    if not root.is_dir():
        raise BridgeError("Workspace is not an existing folder")
    return root


def scoped_path(value: str, root: Path, must_exist: bool, suffix: str | None = None) -> Path:
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = root / path
    resolved = path.resolve(strict=False)
    if not resolved.is_relative_to(root):
        raise BridgeError("OUTSIDE_WORKSPACE: document path is outside the selected job")
    if suffix and resolved.suffix.lower() != suffix:
        raise BridgeError(f"Expected output suffix {suffix}")
    if must_exist and not resolved.is_file():
        raise BridgeError("Source file is missing")
    if not must_exist and (not resolved.parent.is_dir() or path.is_symlink()):
        raise BridgeError("Output parent must exist and output must not be symlink")
    if resolved.suffix.lower() not in SUPPORTED and resolved.suffix.lower() != ".pdf":
        raise BridgeError("Unsupported file type")
    return resolved


def binary(args: argparse.Namespace) -> str | None:
    # Never accept an unverified path via a project document or untrusted text.
    if getattr(args, "binary", None):
        p = Path(args.binary).expanduser().resolve(strict=False)
        return str(p) if p.is_file() and p.suffix.lower() == ".exe" else None
    return shutil.which("officecli") or shutil.which("officecli.exe")


def probe(args: argparse.Namespace) -> dict:
    exe = binary(args)
    if not exe:
        return {
            "available": False, "mode": "standalone",
            "reason": "OfficeCLI is not installed or not on PATH",
            "officecli_required": False,
            "thai_docx_available": (SKILL_ROOT / "vendor/thai-doc-skills/thai-docx/SKILL.md").is_file(),
            "thai_pdf_available": (SKILL_ROOT / "vendor/thai-doc-skills/thai-pdf/SKILL.md").is_file(),
        }
    try:
        run = subprocess.run([exe, "--version"], text=True, encoding="utf-8",
                             errors="replace", capture_output=True, timeout=15,
                             env=non_resident_env())
    except (OSError, subprocess.TimeoutExpired) as err:
        return {"available": False, "mode": "standalone",
                "reason": f"OfficeCLI invocation failed: {type(err).__name__}", "officecli_required": False}
    version = run.stdout.strip().splitlines()[0] if run.stdout.strip() else ""
    return {
        "available": run.returncode == 0 and bool(re.fullmatch(r"\d+\.\d+\.\d+(?:[-+][A-Za-z0-9.-]+)?", version)),
        "mode": "optional-integration" if run.returncode == 0 and version else "standalone",
        "version": version, "binary": exe,
        "officecli_required": False,
        "supported_operations": ["inspect", "validate", "replace-copy"] if run.returncode == 0 else [],
    }


def non_resident_env() -> dict[str, str]:
    env = dict(os.environ)
    # Per OfficeCLI's own skill documentation, avoid persistent sessions,
    # ensuring reads/writes are flushed before another tool validates the file.
    env["OFFICECLI_NO_AUTO_RESIDENT"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def run_cli(exe: str, arguments: list[str], timeout: float = 45) -> tuple[int, object, str]:
    try:
        p = subprocess.run([exe, *arguments], capture_output=True, text=True,
                           encoding="utf-8", errors="replace",
                           timeout=timeout, env=non_resident_env())
    except subprocess.TimeoutExpired as error:
        raise BridgeError(f"OFFICECLI_TIMEOUT after {timeout}s") from error
    except OSError as error:
        raise BridgeError(f"OFFICECLI_EXEC_ERROR: {type(error).__name__}: {error}") from error
    if len(p.stdout) > 1_000_000:
        raise BridgeError("OfficeCLI output unexpectedly large")
    try:
        parsed = json.loads(p.stdout)
    except json.JSONDecodeError:
        parsed = {"raw_output": p.stdout[-4000:]}
    return p.returncode, parsed, p.stderr[-1200:]


def check_preflight() -> None:
    for rel in (
        "vendor/thai-doc-skills/thai-docx/SKILL.md",
        "vendor/thai-doc-skills/thai-pdf/SKILL.md",
    ):
        if not (SKILL_ROOT / rel).is_file():
            raise BridgeError(f"Missing mandatory bundled Thai Document Skills: {rel}")


def status(args: argparse.Namespace) -> int:
    result = probe(args)
    emit({"status": "AVAILABLE" if result.get("available") else "STANDALONE_ONLY", **result})
    return 0


def inspect(args: argparse.Namespace) -> int:
    root = workspace_root(args)
    source = scoped_path(args.source, root, must_exist=True)
    if source.suffix.lower() not in SUPPORTED:
        raise BridgeError("Only DOCX/XLSX/PPTX can be inspected by this bridge")
    p = probe(args)
    if not p.get("available"):
        if source.suffix.lower() == ".docx" and args.allow_basic_fallback:
            return fallback_docx(source)
        emit({"status": "OFFICECLI_UNAVAILABLE", "officecli_used": False,
              "source": str(source), "reason": p.get("reason")})
        return 2
    check_preflight()
    code, result, err = run_cli(p["binary"], ["view", str(source), args.mode, "--json"], timeout=75)
    emit({"status": "PASS" if code == 0 else "BLOCKED", "source": str(source),
          "mode": args.mode, "officecli_used": True, "exit_code": code,
          "result": result, "stderr": err})
    return 0 if code == 0 else 2


def validate(args: argparse.Namespace) -> int:
    root = workspace_root(args)
    source = scoped_path(args.source, root, must_exist=True)
    if source.suffix.lower() not in SUPPORTED:
        raise BridgeError("Validation is only supported for DOCX/XLSX/PPTX")
    p = probe(args)
    if not p.get("available"):
        emit({"status": "OFFICECLI_UNAVAILABLE", "officecli_used": False, "source": str(source)})
        return 2
    check_preflight()
    code, result, err = run_cli(p["binary"], ["validate", str(source), "--json"], timeout=90)
    success = code == 0 and isinstance(result, dict) and result.get("success") is True
    emit({"status": "PASS" if success else "SCHEMA_ISSUES",
          "officecli_used": True, "source": str(source),
          "exit_code": code, "result": result, "stderr": err})
    return 0 if success else 3


def fallback_docx(path: Path) -> int:
    # This is NOT an OpenXML schema validator. It is a basic ZIP+text fallback.
    try:
        with zipfile.ZipFile(path) as z:
            if z.testzip():
                raise BridgeError("Corrupted ZIP member")
            xml = z.read("word/document.xml")
        root = ET.fromstring(xml)
        text = "".join((e.text or "") for e in root.iter(f"{NS_W}t"))
        tables = sum(1 for _ in root.iter(f"{NS_W}tbl"))
    except (KeyError, zipfile.BadZipFile, ET.ParseError, OSError) as error:
        raise BridgeError(f"DOCX basic read failed: {error}") from error
    emit({"status": "BASIC_READ_ONLY", "officecli_used": False,
          "source": str(path), "text": text[:8000], "tables": tables,
          "warning": "Basic OOXML read only; NOT equivalent to OfficeCLI inspect/validate"})
    return 0


def paragraph_texts(path: Path) -> list[str]:
    """Plain text of every OOXML paragraph (including those inside tables)."""
    with zipfile.ZipFile(path) as z:
        document = ET.fromstring(z.read("word/document.xml"))
    return [
        "".join((t.text or "") for t in paragraph.iter(f"{NS_W}t"))
        for paragraph in document.iter(f"{NS_W}p")
    ]


def file_sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def structure(path: Path) -> dict:
    with zipfile.ZipFile(path) as z:
        document = ET.fromstring(z.read("word/document.xml"))
        parts = set(z.namelist())
    return {
        "tables": len(list(document.iter(f"{NS_W}tbl"))),
        "sections": len(list(document.iter(f"{NS_W}sectPr"))),
        "images": len([x for x in parts if x.startswith("word/media/")]),
        "headers": len([x for x in parts if re.fullmatch(r"word/header\d+\.xml", x)]),
        "footers": len([x for x in parts if re.fullmatch(r"word/footer\d+\.xml", x)]),
    }


def replace_copy(args: argparse.Namespace) -> int:
    """Replace ONE exact text string, only in a NEW copy, preserving original."""
    root = workspace_root(args)
    source = scoped_path(args.source, root, must_exist=True, suffix=".docx")
    output = scoped_path(args.output, root, must_exist=False, suffix=".docx")
    if output == source:
        raise BridgeError("Refusing in-place mutation")
    if output.exists():
        raise BridgeError("Output already exists: no overwrite permitted")
    if not args.find or args.find == args.replace:
        raise BridgeError("Nonempty --find and a different --replace value required")
    if len(args.find) > 4096 or len(args.replace) > 4096:
        raise BridgeError("Replacement text exceeds 4096 characters")
    p = probe(args)
    if not p.get("available"):
        emit({"status": "OFFICECLI_UNAVAILABLE", "officecli_used": False,
              "source_sha256": file_sha(source), "source": str(source)})
        return 2
    check_preflight()
    before_hash = file_sha(source)
    before_structure = structure(source)
    before_text = paragraph_texts(source)
    expected_text = [p.replace(args.find, args.replace) for p in before_text]
    expected_matches = sum(p.count(args.find) for p in before_text)
    if expected_matches == 0:
        raise BridgeError("Text to replace was not found as an exact match; no output was created")
    if args.dry_run:
        emit({"status": "DRY_RUN", "officecli_used": False,
              "source": str(source), "output": str(output),
              "find": args.find, "replace": args.replace,
              "source_sha256": before_hash, "source_structure": before_structure,
              "expected_matches": expected_matches})
        return 0
    # Never write to the original; output is explicitly requested and unique.
    shutil.copy2(source, output)
    code, result, err = run_cli(p["binary"],
                                ["set", str(output), "/", "--find", args.find,
                                 "--replace", args.replace, "--json"], timeout=100)
    after_structure = structure(output)
    sha_unchanged = before_hash == file_sha(source)
    structure_unchanged = before_structure == after_structure
    data = result.get("data") if isinstance(result, dict) else None
    matched = result.get("matched") if isinstance(result, dict) else None
    if matched is None and isinstance(data, dict):
        matched = data.get("matched")
    try:
        actual_text = paragraph_texts(output)
    except (zipfile.BadZipFile, KeyError, ET.ParseError):
        actual_text = []
    text_exact = actual_text == expected_text
    ok = code == 0 and sha_unchanged and structure_unchanged and text_exact and (
        not isinstance(result, dict) or result.get("success", True)
    )
    emit({
        "status": "PASS" if ok else "REVIEW_REQUIRED",
        "officecli_used": True, "source": str(source), "output": str(output),
        "source_unchanged": sha_unchanged,
        "structure_unchanged": structure_unchanged,
        "before_structure": before_structure, "after_structure": after_structure,
        "matches_reported": matched, "expected_matches": expected_matches,
        "text_changed_exactly_as_requested": text_exact, "exit_code": code,
        "result": result, "stderr": err,
        "needs_visual_qa": True,
    })
    return 0 if ok else 3



def repair_schema_copy(args: argparse.Namespace) -> int:
    """Repair narrow OOXML element-order defects from a bundled DOCX generator.

    Only emits a new copy. This is not general-purpose repair and must not be
    used for signed documents. Does not depend on OfficeCLI; if present, calls
    validation on the output as an optional QA gate.
    """
    root = workspace_root(args)
    source = scoped_path(args.source, root, must_exist=True, suffix=".docx")
    output = scoped_path(args.output, root, must_exist=False, suffix=".docx")
    if source == output or output.exists():
        raise BridgeError("Refusing in-place edit or overwrite of output")
    try:
        from lxml import etree
    except ImportError as exc:
        raise BridgeError("Schema repair requires lxml (included with python-docx)") from exc
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    w = lambda local: f"{{{ns}}}{local}"
    rank_names = (
        "rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline "
        "shadow emboss imprint noProof snapToGrid vanish webHidden color spacing "
        "w kern position sz szCs highlight u effect bdr shd fitText vertAlign "
        "rtl cs em lang eastAsianLayout specVanish oMath"
    ).split()
    rank = {w(name): i for i, name in enumerate(rank_names)}
    src_hash = file_sha(source)
    src_structure = structure(source)
    edited = {}
    reorders = 0
    with zipfile.ZipFile(source) as source_zip:
        files = source_zip.namelist()
        if any("_xmlsignatures" in n.lower() for n in files):
            raise BridgeError("Signed DOCX is not supported for normalization")
        document_name = "word/document.xml"
        settings_name = "word/settings.xml"
        doc = etree.fromstring(source_zip.read(document_name))
        for element in doc.iter(w("rPr")):
            before = list(element)
            ordered = sorted(before, key=lambda element: rank.get(element.tag, 10000))
            if ordered != before:
                element[:] = ordered
                reorders += 1
        if reorders:
            edited[document_name] = etree.tostring(
                doc, encoding="UTF-8", xml_declaration=True, standalone=True
            )
        if settings_name in files:
            settings = etree.fromstring(source_zip.read(settings_name))
            original_xml = etree.tostring(settings)
            # w:embedTrueTypeFonts / saveSubsetFonts must precede proofState.
            # The upstream font-embed utility can append them after proofState.
            for local in ("embedTrueTypeFonts", "saveSubsetFonts"):
                node = settings.find(w(local))
                if node is None:
                    continue
                settings.remove(node)
                proof = settings.find(w("proofState"))
                if proof is not None:
                    settings.insert(list(settings).index(proof), node)
                else:
                    zoom = settings.find(w("zoom"))
                    settings.insert(list(settings).index(zoom) + 1 if zoom is not None else 0, node)
            if etree.tostring(settings) != original_xml:
                edited[settings_name] = etree.tostring(
                    settings, encoding="UTF-8", xml_declaration=True, standalone=True
                )
        with zipfile.ZipFile(output, "w") as result_zip:
            for info in source_zip.infolist():
                result_zip.writestr(info, edited.get(info.filename, source_zip.read(info.filename)))
    unchanged_original = file_sha(source) == src_hash
    unchanged_structure = structure(output) == src_structure
    unchanged_text = paragraph_texts(output) == paragraph_texts(source)
    state = {
        "status": "CREATED_NOT_VISUALLY_REVIEWED",
        "source": str(source), "output": str(output),
        "officecli_used": False, "modified_parts": sorted(edited),
        "reordered_run_properties": reorders,
        "source_unchanged": unchanged_original,
        "structure_unchanged": unchanged_structure,
        "text_unchanged": unchanged_text,
    }
    p = probe(args)
    if p.get("available"):
        check_preflight()
        code, result, stderr = run_cli(p["binary"], ["validate", str(output), "--json"], timeout=90)
        state["officecli_validation"] = {
            "passed": code == 0 and isinstance(result, dict) and result.get("success") is True,
            "result": result, "stderr": stderr,
        }
        state["officecli_used"] = True
        if not state["officecli_validation"]["passed"]:
            state["status"] = "SCHEMA_ISSUES"
    if not unchanged_original or not unchanged_structure or not unchanged_text:
        state["status"] = "REVIEW_REQUIRED"
    elif state["status"] != "SCHEMA_ISSUES":
        state["status"] = "PASS" if p.get("available") else "NORMALIZED_NEEDS_SCHEMA_VALIDATION"
    emit(state)
    return 0 if state["status"] in {"PASS", "NORMALIZED_NEEDS_SCHEMA_VALIDATION"} else 3

def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    cli = argparse.ArgumentParser(description="Optional THaksornCraft ↔ OfficeCLI adapter")
    cli.add_argument("--workspace", default=".", help="Exact authorized workspace/project directory")
    cli.add_argument("--binary", help="Optional explicit path to existing officecli.exe")
    sub = cli.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("status", help="Detect OfficeCLI without modifying it")
    s.set_defaults(func=status)
    v = sub.add_parser("inspect", help="Read OfficeCLI view (read-only)")
    v.add_argument("source")
    v.add_argument("--mode", choices=("text", "issues", "stats", "outline", "annotated"), default="issues")
    v.add_argument("--allow-basic-fallback", action="store_true",
                   help="DOCX only: basic ZIP/text read if OfficeCLI absent")
    v.set_defaults(func=inspect)
    val = sub.add_parser("validate", help="Read-only OpenXML schema validation with OfficeCLI")
    val.add_argument("source")
    val.set_defaults(func=validate)
    edit = sub.add_parser("replace-copy", help="Replace exact text ONLY in a new DOCX copy")
    edit.add_argument("source")
    edit.add_argument("--output", "-o", required=True)
    edit.add_argument("--find", required=True)
    edit.add_argument("--replace", required=True)
    edit.add_argument("--dry-run", action="store_true")
    edit.set_defaults(func=replace_copy)

    fix = sub.add_parser("repair-schema-copy",
                         help="Narrow OOXML element-order normalization on a NEW DOCX copy")
    fix.add_argument("source")
    fix.add_argument("--output", "-o", required=True)
    fix.set_defaults(func=repair_schema_copy)
    args = cli.parse_args(argv)
    try:
        return args.func(args)
    except (BridgeError, OSError, ValueError, KeyError, zipfile.BadZipFile,
            ET.ParseError) as e:
        print(f"THaksornCraft OfficeCLI adapter BLOCKED: {type(e).__name__}: {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
