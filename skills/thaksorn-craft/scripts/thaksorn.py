#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""THaksornCraft standalone document renderer and environment doctor.

Creates NEW Word/PDF documents from UTF-8 Markdown/TXT or structured JSON.
It never imports, installs, starts or calls OfficeCLI. It never overwrites by default.
Copyright 2026 THaksornCraft integration contributors.
Original engines retain their own licenses in vendor/.
"""
from __future__ import annotations

import argparse
import contextlib
import importlib.util
import json
from pathlib import Path
import sys
import zipfile

SKILL_ROOT = Path(__file__).resolve().parents[1]
VENDOR = SKILL_ROOT / "vendor" / "thai-doc-skills"
DOCX_ENGINE = VENDOR / "thai-docx" / "scripts" / "thai_docx.py"
PDF_ENGINE = VENDOR / "thai-pdf" / "scripts" / "thai_pdf.py"
FONT_DIR = VENDOR / "thai-docx" / "fonts" / "national"
VALID_TYPES = {
    "body", "title", "subtitle", "heading1", "heading2", "heading3",
    "bullet", "number", "quote", "caption", "table", "image", "pagebreak",
}


def make_utf8_console() -> None:
    """Prevent cp874 Windows terminals from crashing on Thai status text."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def emit(value: dict) -> None:
    print(json.dumps(value, ensure_ascii=False, indent=2))


def load_engine(file_path: Path, name: str):
    """Import one selected engine with a unique module name."""
    if not file_path.is_file():
        raise FileNotFoundError(f"Vendored renderer missing: {file_path}")
    # thai-docx's font_embed module is in the same scripts folder.
    sys.path.insert(0, str(file_path.parent))
    spec = importlib.util.spec_from_file_location(name, file_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load renderer: {file_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def validate_blocks(value: object, source: Path) -> list[dict]:
    if isinstance(value, dict):
        value = value.get("paragraphs")
    if not isinstance(value, list):
        raise ValueError("JSON input must be an array or object with a paragraphs array.")
    validated = []
    for index, block in enumerate(value, 1):
        if not isinstance(block, dict):
            raise ValueError(f"Paragraph {index} must be an object.")
        kind = block.get("type", "body")
        if kind not in VALID_TYPES:
            raise ValueError(f"Unsupported paragraph type at block {index}: {kind!r}")
        item = dict(block)
        item["type"] = kind
        if kind == "table":
            rows = item.get("rows")
            if not isinstance(rows, list) or not rows or not all(
                isinstance(row, list) and len(row) >= 1 for row in rows
            ):
                raise ValueError(f"Block {index}: table.rows must be a nonempty array of nonempty arrays.")
            width = len(rows[0])
            if any(len(row) != width for row in rows):
                raise ValueError(f"Block {index}: table rows must have equal column counts.")
            item["rows"] = [
                [str(cell) if cell is not None else "" for cell in row] for row in rows
            ]
        elif kind == "image":
            name = item.get("path")
            if not isinstance(name, str) or not name:
                raise ValueError(f"Block {index}: image.path required.")
            image = Path(name)
            if not image.is_absolute():
                image = source.parent / image
            image = image.resolve(strict=False)
            parent = source.parent.resolve(strict=True)
            if not image.is_relative_to(parent):
                raise ValueError(f"Block {index}: image must be within the input directory: {image}")
            if not image.is_file():
                raise FileNotFoundError(f"Block {index}: image not found: {image}")
            item["path"] = str(image)
        elif kind != "pagebreak" and not isinstance(item.get("text"), str):
            raise ValueError(f"Block {index}: text must be a string.")
        validated.append(item)
    if not validated:
        raise ValueError("Input has no document blocks.")
    return validated


def load_blocks(source: Path, engine) -> tuple[list[dict], str]:
    suffix = source.suffix.lower()
    raw = source.read_text(encoding="utf-8-sig")
    if suffix == ".json":
        return validate_blocks(json.loads(raw), source), "structured-json"
    if suffix not in {".txt", ".md"}:
        raise ValueError("Only .md, .txt and .json inputs are supported (not existing Office files).")
    import re
    if re.search(r"(?m)^\s*\|.+\|\s*$", raw):
        raise ValueError("Markdown tables are not supported; use structured JSON blocks instead.")
    if (chr(96) * 3) in raw or re.search(r"!\[[^\]]*\]\(", raw):
        raise ValueError("Markdown code fences/images are unsupported; use structured JSON blocks.")
    blocks = engine.parse_markdown(raw)
    return validate_blocks(blocks, source), "basic-markdown"


def word_embedded_fonts(output: Path) -> bool:
    with zipfile.ZipFile(output) as z:
        return any(
            n.startswith("word/fonts/") and n.lower().endswith(".odttf")
            for n in z.namelist()
        )


def word_export_pdf(source_docx: Path, output_pdf: Path, *, refresh_generated_toc: bool = False) -> None:
    """Export through installed Microsoft Word COM, not OfficeCLI.

    Requires Windows, Microsoft Word and optional pywin32; never saves changes
    to the source document. A separate Word instance is closed in all cases.
    """
    if sys.platform != "win32":
        raise RuntimeError("Microsoft Word COM PDF export is Windows-only.")
    try:
        import win32com.client
    except ImportError as exc:
        raise ImportError(
            "Optional Word PDF backend requires pywin32 and installed Microsoft Word."
        ) from exc
    app = None
    doc = None
    try:
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Open(str(source_docx), ReadOnly=True, AddToRecentFiles=False, Visible=False)
        if refresh_generated_toc:
            # Word can update the TOC field in-memory without saving the source DOCX.
            # The upstream DOCX renderer also inserts an instruction-only hint paragraph.
            # Remove that hint from the exported PDF; preserve the original DOCX unchanged.
            for paragraph in list(doc.Paragraphs):
                paragraph_text = str(paragraph.Range.Text)
                if "คลิกขวาที่สารบัญ" in paragraph_text and "Update Field" in paragraph_text:
                    paragraph.Range.Delete()
            doc.Fields.Update()
            for i in range(1, int(doc.TablesOfContents.Count) + 1):
                doc.TablesOfContents(i).Update()
        doc.ExportAsFixedFormat(str(output_pdf), 17)
    finally:
        if doc is not None:
            doc.Close(SaveChanges=False)
        if app is not None:
            app.Quit(SaveChanges=False)
    if not output_pdf.is_file() or output_pdf.read_bytes()[:4] != b"%PDF":
        raise RuntimeError("Microsoft Word export did not produce a valid PDF header.")


def cmd_doctor(args: argparse.Namespace) -> int:
    deps = {}
    for package in ("docx", "pythainlp", "weasyprint", "win32com"):
        deps[package] = importlib.util.find_spec(package) is not None
    fonts = list(FONT_DIR.glob("*.ttf")) if FONT_DIR.is_dir() else []
    report = {
        "product": "THaksornCraft",
        "python": sys.version.split()[0],
        "officecli_required": False,
        "components": {
            "kien_thai": (SKILL_ROOT / "vendor/kien-thai/skills/kien-thai/SKILL.md").is_file(),
            "kode_thai": (SKILL_ROOT / "vendor/kien-thai/skills/kode-thai/SKILL.md").is_file(),
            "thai_docx": DOCX_ENGINE.is_file(),
            "thai_pdf": PDF_ENGINE.is_file(),
        },
        "dependencies": deps,
        "bundled_national_font_files": len(fonts),
        "docx_ready": bool(deps["docx"] and DOCX_ENGINE.is_file()),
        "word_pdf_export_module_installed": bool(sys.platform == "win32" and deps["win32com"]),
        "pdf_ready": "not_probed",
    }
    if args.probe_pdf:
        import contextlib
        import io
        notice = io.StringIO()
        try:
            with contextlib.redirect_stdout(notice):
                from weasyprint import HTML
                data = HTML(string="<html><body><p>ทดสอบภาษาไทย</p></body></html>").write_pdf()
            report["pdf_ready"] = data.startswith(b"%PDF")
        except Exception as exc:
            report["pdf_ready"] = False
            report["pdf_error"] = f"{type(exc).__name__}: {exc}"
        if notice.getvalue().strip():
            report["pdf_import_notice"] = notice.getvalue().strip()
    emit(report)
    return 0 if all(report["components"].values()) else 2


def check_officecli_qa_readiness(args: argparse.Namespace, output: Path, fmt: str) -> dict | None:
    """Preflight optional OfficeCLI before creating document output."""
    if args.officecli_qa == "off":
        if args.repair_known_schema:
            raise ValueError("--repair-known-schema requires --officecli-qa auto or require")
        return None
    if fmt != ".docx":
        raise ValueError("--officecli-qa is for new DOCX output only; PDF has separate validation")
    from officecli_adapter import probe
    state = probe(argparse.Namespace(binary=None))
    if not state.get("available") and args.officecli_qa == "require":
        raise RuntimeError("OfficeCLI required by --officecli-qa=require, but not available")
    if args.repair_known_schema:
        normalized = output.with_name(output.stem + ".normalized.docx")
        if normalized.exists() or normalized.is_symlink():
            raise FileExistsError(f"Normalized output already exists: {normalized}")
    return state


def run_officecli_qa(
    output: Path, args: argparse.Namespace, availability: dict | None,
) -> dict:
    """Validate, and optionally fix a *separate copy*, using the bounded bridge."""
    if not availability or not availability.get("available"):
        return {
            "status": "OFFICECLI_UNAVAILABLE_STANDALONE",
            "officecli_used": False,
            "warning": "New DOCX generated; no OfficeCLI schema validation performed",
        }
    from officecli_adapter import check_preflight, repair_schema_copy, run_cli
    check_preflight()
    exit_code, result, stderr = run_cli(
        availability["binary"], ["validate", str(output), "--json"], timeout=90,
    )
    passed = exit_code == 0 and isinstance(result, dict) and result.get("success") is True
    qa = {
        "status": "PASS" if passed else "SCHEMA_ISSUES",
        "officecli_used": True,
        "checked_output": str(output),
        "validation_result": result,
        "stderr": stderr,
    }
    if not passed and args.repair_known_schema:
        import contextlib
        import io
        normalized = output.with_name(output.stem + ".normalized.docx")
        bridge_args = argparse.Namespace(
            workspace=str(output.parent), binary=None,
            source=str(output), output=str(normalized),
        )
        captured = io.StringIO()
        with contextlib.redirect_stdout(captured):
            repair_result = repair_schema_copy(bridge_args)
        repaired_report = json.loads(captured.getvalue())
        qa["normalization"] = repaired_report
        if repair_result == 0 and repaired_report.get("status") == "PASS":
            qa["status"] = "PASS_REPAIRED_COPY"
            qa["recommended_output"] = str(normalized)
    return qa


def cmd_render(args: argparse.Namespace) -> int:
    source = Path(args.source).expanduser().resolve(strict=True)
    output = Path(args.output).expanduser().resolve(strict=False)
    if not source.is_file():
        raise ValueError("Source must be an existing UTF-8 file.")
    if source.suffix.lower() not in {".md", ".txt", ".json"}:
        raise ValueError("Only .md, .txt and .json source files are supported.")
    fmt = output.suffix.lower()
    if fmt not in {".docx", ".pdf"}:
        raise ValueError("Output suffix must be .docx or .pdf.")
    if source == output:
        raise ValueError("Refusing in-place replacement of source input.")
    if not output.parent.is_dir():
        raise FileNotFoundError(f"Output directory does not exist: {output.parent}")
    if output.exists() and not args.overwrite:
        raise FileExistsError(f"Output exists; refusing to overwrite: {output}")
    if output.is_symlink():
        raise ValueError("Refusing to render to a symbolic-link output path.")
    if fmt == ".pdf" and args.no_embed:
        raise ValueError("--no-embed only applies to DOCX, not PDF.")
    availability = check_officecli_qa_readiness(args, output, fmt)

    use_word_pdf = fmt == ".pdf" and args.pdf_backend == "word"
    selected = DOCX_ENGINE if fmt == ".docx" or use_word_pdf else PDF_ENGINE
    module = load_engine(selected, f"thaksorn_{fmt[1:]}_engine")
    blocks, input_kind = load_blocks(source, module)
    kwargs = {
        "preset": args.preset,
        "font_name": args.font,
        "toc": args.toc,
        "title": args.title,
        "header_text": args.header,
        "footer_text": args.footer,
        "page_number": args.page_number,
        "page_number_format": args.page_number_format,
    }
    sidecar = None
    if fmt == ".docx":
        with contextlib.redirect_stdout(sys.stderr):
            module.create_docx(blocks, str(output), embed_fonts=not args.no_embed, **kwargs)
        embedded = word_embedded_fonts(output)
        renderer = "vendored python-docx (thai-docx)"
    elif use_word_pdf:
        sidecar = output.with_name(output.stem + ".source.docx")
        if sidecar.exists() and not args.overwrite:
            raise FileExistsError(f"Source DOCX companion exists; refusing overwrite: {sidecar}")
        if sidecar.is_symlink():
            raise ValueError("Refusing symbolic-link source DOCX companion.")
        with contextlib.redirect_stdout(sys.stderr):
            module.create_docx(blocks, str(sidecar), embed_fonts=True, **kwargs)
        word_export_pdf(sidecar, output, refresh_generated_toc=args.toc)
        embedded = None  # PDF font embedding must be independently verified.
        renderer = "vendored thai-docx + Microsoft Word COM PDF export"
    else:
        with contextlib.redirect_stdout(sys.stderr):
            _, embedded = module.create_pdf(blocks, str(output), **kwargs)
        renderer = "vendored WeasyPrint (thai-pdf)"

    qa = run_officecli_qa(output, args, availability) if args.officecli_qa != "off" else None
    status = "CREATED_NOT_VISUALLY_REVIEWED"
    if qa and qa["status"] == "SCHEMA_ISSUES":
        status = "SCHEMA_ISSUES"
    emit({
        "status": status,
        "output": str(output),
        "bytes": output.stat().st_size,
        "engine": renderer,
        "source_format": input_kind,
        "blocks": len(blocks),
        "font": args.font,
        "font_embedded": embedded if embedded is None else bool(embedded),
        "officecli_used": bool(qa and qa.get("officecli_used")),
        "officecli_qa": qa,
        "recommended_output": qa.get("recommended_output", str(output)) if qa else str(output),
        "source_docx_companion": str(sidecar) if sidecar else None,
        "needs_visual_qa": True,
    })
    return 3 if status == "SCHEMA_ISSUES" else 0


def cmd_export_word(args: argparse.Namespace) -> int:
    """Convert a pre-existing DOCX without altering it."""
    source = Path(args.source).expanduser().resolve(strict=True)
    output = Path(args.output).expanduser().resolve(strict=False)
    if not source.is_file() or source.suffix.lower() != ".docx":
        raise ValueError("Source must be an existing .docx")
    if output.suffix.lower() != ".pdf":
        raise ValueError("Target must have .pdf suffix")
    if not output.parent.is_dir():
        raise FileNotFoundError(f"Target directory is missing: {output.parent}")
    if output.exists() and not args.overwrite:
        raise FileExistsError(f"PDF output exists; refusing overwrite: {output}")
    if output.is_symlink():
        raise ValueError("Refusing symbolic-link target.")
    word_export_pdf(source, output)
    emit({
        "status": "CREATED_NOT_VISUALLY_REVIEWED",
        "output": str(output),
        "bytes": output.stat().st_size,
        "engine": "Microsoft Word COM (not OfficeCLI)",
        "source_docx": str(source),
        "officecli_used": False,
        "needs_visual_qa": True,
    })
    return 0


def cmd_officecli(args: argparse.Namespace) -> int:
    # Lazy import: no OfficeCLI dependency for standalone use.
    from officecli_adapter import main as bridge_main
    return bridge_main(args.adapter_args)


def main(argv: list[str] | None = None) -> int:
    make_utf8_console()
    raw_args = list(sys.argv[1:] if argv is None else argv)
    if raw_args and raw_args[0] == "officecli":
        from officecli_adapter import main as bridge_main
        return bridge_main(raw_args[1:])
    parser = argparse.ArgumentParser(description="THaksornCraft — standalone Thai document generator")
    sub = parser.add_subparsers(dest="command", required=True)

    doctor = sub.add_parser("doctor", help="Check optional local renderer prerequisites")
    doctor.add_argument("--json", action="store_true", help="Output is always JSON; kept for compatibility")
    doctor.add_argument("--probe-pdf", action="store_true", help="Render a small in-memory PDF to verify Pango")
    doctor.set_defaults(func=cmd_doctor)

    render = sub.add_parser("render", help="Create a NEW DOCX or PDF from Markdown/TXT/JSON")
    render.add_argument("source", help="UTF-8 .md/.txt or structured .json")
    render.add_argument("--output", "-o", required=True, help="New .docx or .pdf file")
    render.add_argument("--preset", choices=("saraban", "default", "book"), default="saraban")
    render.add_argument("--font", default="TH SarabunPSK", help="Thai font family")
    render.add_argument("--toc", action="store_true", help="Include automatic table of contents")
    render.add_argument("--title")
    render.add_argument("--header")
    render.add_argument("--footer")
    render.add_argument("--page-number", choices=("footer-center", "footer-right", "header-right"))
    render.add_argument("--page-number-format", default="{n}")
    render.add_argument("--no-embed", action="store_true", help="Disable DOCX font embedding")
    render.add_argument("--officecli-qa", choices=("off", "auto", "require"), default="off",
                        help="Optional OfficeCLI schema QA on DOCX (off by default)")
    render.add_argument("--repair-known-schema", action="store_true",
                        help="If schema fails, normalize known upstream OOXML ordering in a NEW copy")
    render.add_argument("--pdf-backend", choices=("weasyprint", "word"), default="weasyprint",
                        help="Use Pango/WeasyPrint (default) or optional Microsoft Word without OfficeCLI")
    render.add_argument("--overwrite", action="store_true", help="Explicitly permit replacing existing output")
    render.set_defaults(func=cmd_render)

    exporter = sub.add_parser("export-word", help="Convert existing DOCX to PDF with Word COM; no OfficeCLI")
    exporter.add_argument("source", help="Existing .docx file")
    exporter.add_argument("--output", "-o", required=True, help="Target .pdf file")
    exporter.add_argument("--overwrite", action="store_true")
    exporter.set_defaults(func=cmd_export_word)

    bridge = sub.add_parser("officecli", help="Optional OfficeCLI integration (never required)")
    bridge.add_argument("adapter_args", nargs=argparse.REMAINDER)
    bridge.set_defaults(func=cmd_officecli)

    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except Exception as exc:
        print(f"THaksornCraft BLOCKED: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
