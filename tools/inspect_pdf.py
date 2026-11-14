#!/usr/bin/env python3
"""PDF text, page and font verification for development QA only (requires pypdf)."""
from __future__ import annotations
import argparse
import json
import sys
from pathlib import Path
from pypdf import PdfReader

def inspect(path: Path) -> dict:
    reader = PdfReader(str(path))
    texts = [(p.extract_text() or "") for p in reader.pages]
    embedded, unembedded = set(), set()
    for page in reader.pages:
        if hasattr(page, "get_fonts"):
            emb, unemb = page.get_fonts()
            embedded.update(str(x) for x in emb)
            unembedded.update(str(x) for x in unemb)
    first_box = reader.pages[0].mediabox if reader.pages else None
    all_text = "\n".join(texts)
    return {
        "file": str(path),
        "pages": len(reader.pages),
        "page_first_points": ([float(first_box.width), float(first_box.height)] if first_box else []),
        "thai_chars_extracted": sum("\u0e00" <= ch <= "\u0e7f" for ch in all_text),
        "text_start": all_text[:900],
        "fonts_embedded": sorted(embedded),
        "fonts_not_embedded": sorted(unembedded),
        "a4_portrait": bool(first_box and 580 <= float(first_box.width) <= 610
                            and 820 <= float(first_box.height) <= 855),
    }

if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass
    parser = argparse.ArgumentParser()
    parser.add_argument("file")
    args = parser.parse_args()
    data = inspect(Path(args.file).resolve(strict=True))
    print(json.dumps(data, ensure_ascii=False, indent=2))
    raise SystemExit(0 if data["pages"] > 0 and data["thai_chars_extracted"] > 0 else 2)
