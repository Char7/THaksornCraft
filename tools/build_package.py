#!/usr/bin/env python3
"""Build a self-contained THaksornCraft skill ZIP within the project workspace.

Never edits bundled source or global skills. Preserves original third-party licenses.
"""
from __future__ import annotations

import argparse
import hashlib
from pathlib import Path
import zipfile

ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "skills" / "thaksorn-craft"
DIST = ROOT / "dist"
VERSION = "0.2.0"
ESSENTIAL = [
    "SKILL.md",
    "SOURCES.lock.json",
    "agents/openai.yaml",
    "scripts/thaksorn.py",
    "scripts/officecli_adapter.py",
    "references/officecli-integration.md",
    "vendor/kien-thai/LICENSE",
    "vendor/kien-thai/skills/kien-thai/SKILL.md",
    "vendor/kien-thai/skills/kode-thai/SKILL.md",
    "vendor/thai-doc-skills/LICENSE",
    "vendor/thai-doc-skills/thai-docx/SKILL.md",
    "vendor/thai-doc-skills/thai-pdf/SKILL.md",
    "vendor/thai-doc-skills/thai-docx/fonts/LICENSE-NATIONAL-FONTS.txt",
]


def build(overwrite: bool = False) -> Path:
    missing = [name for name in ESSENTIAL if not (SKILL / name).is_file()]
    if missing:
        raise RuntimeError("Missing required skill assets: " + ", ".join(missing))
    if not DIST.is_dir():
        DIST.mkdir(parents=True)
    archive = DIST / f"THaksornCraft_v{VERSION}.zip"
    if archive.exists() and not overwrite:
        raise FileExistsError(f"Archive exists (no implicit overwrite): {archive}")
    files = [
        path for path in SKILL.rglob("*")
        if path.is_file()
        and not any(part in {".git", "__pycache__", ".pytest_cache"} for part in path.parts)
        and path.suffix not in {".pyc", ".pyo"}
    ]
    if not files:
        raise RuntimeError("No skill files found")
    for path in files:
        if path.is_symlink():
            raise RuntimeError(f"Refusing symlink in distribution: {path}")
    with zipfile.ZipFile(archive, mode="w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as output:
        for file_path in sorted(files):
            output.write(file_path, arcname=f"thaksorn-craft/{file_path.relative_to(SKILL).as_posix()}")
    with zipfile.ZipFile(archive) as check:
        if check.testzip() is not None:
            raise RuntimeError("Archive ZIP failed integrity test")
    digest = hashlib.sha256(archive.read_bytes()).hexdigest()
    checksum = DIST / f"THaksornCraft_v{VERSION}.sha256.txt"
    checksum.write_text(f"{digest}  {archive.name}\n", encoding="utf-8")
    print(f"ZIP: {archive}\nFILES: {len(files)}\nSIZE: {archive.stat().st_size}\nSHA256: {digest}")
    return archive


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()
    build(overwrite=args.overwrite)
