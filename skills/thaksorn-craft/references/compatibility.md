# Compatibility and dependencies

THaksornCraft is an **Agent Skill**, not a mandatory OfficeCLI plugin. The bundled renderer accepts NEW Markdown/TXT or JSON documents; it is not a full Word editor.

## Dependency map

| Capability | Runtime requirement |
|---|---|
| Read/use language rules | Any compatible agent/LLM; no Python or OfficeCLI |
| Create DOCX | Python + python-docx >=1.0; pythainlp >=4.0 recommended for line-break quality |
| Create PDF using bundled thai-pdf | Python + weasyprint >=60 + pythainlp; native Pango/font libraries must work |
| Alternative PDF via Microsoft Word (Windows) | Installed Microsoft Word + optional pywin32; no OfficeCLI |
| High-fidelity edit existing Word | External document editor (e.g. OfficeCLI) OPTIONAL |
| Matching DOCX/PDF pagination | Word-compatible PDF export, verify exact pages |
| Optional native Thai draft | Explicit user permission + local Ollama + suitable model; no auto-download |

Run doctor before rendering. On Windows, WeasyPrint can fail even with pip packages installed because native Pango/GObject libraries are missing. A successful import does NOT establish successful PDF rendering: use doctor --probe-pdf (renders only an in-memory test PDF).

Install packages into an isolated venv if user authorizes dependency installation:

    python -m venv .venv
    .venv\Scripts\python -m pip install -r skills\thaksorn-craft\requirements-docx.txt
    .venv\Scripts\python -m pip install -r skills\thaksorn-craft\requirements-pdf.txt
    .venv\Scripts\python -m pip install -r skills\thaksorn-craft\requirements-wordpdf-optional.txt

No OfficeCLI is required or installed by these requirements. When WeasyPrint lacks its system libraries, select --pdf-backend word on Windows *only if Word is already installed*. That path creates and retains a companion .source.docx, avoids Pango, and can also convert a previously approved DOCX with export-word. It is not portable to machines without Word. Do not change system Python or existing OfficeCLI.

## Two PDF implementations in upstream

thai-doc-skills has both thai-docx/scripts/thai_pdf.py and thai-pdf/scripts/thai_pdf.py with different source code. THaksornCraft explicitly routes independent PDF creation to the STANDALONE thai-pdf/scripts/thai_pdf.py. Never import by an ambiguous loose filename without controlling sys.path.

## Thai typography and font family

Use TH SarabunPSK when selected. The upstream presets may use TH Sarabun New instead; the CLI passes the selected font explicitly so the user-approved font takes priority. Verify font tables, optional embedded fonts and actual displayed Thai text. Bundled national fonts are **NOT MIT** (see references/licenses.md).

DOCX uses python-docx and Zero-Width Spaces (U+200B) from PyThaiNLP to improve Thai line breaks. PDF uses WeasyPrint/Pango plus word segmentation. Different engines can produce different page counts even with the same source.

## Inputs/outputs and limitations

- CLI renders NEW files. It rejects input extensions other than .md, .txt, .json, and output other than .docx/.pdf.
- Does not overwrite by default.
- Does not install system packages, mutate OfficeCLI, download models, or start network services.
- DOCX generated table/image/TOC handling derives from vendored thai-docx; visual inspection remains necessary.
- Not an automatic lossless conversion tool for complex existing Office files.
