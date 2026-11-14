---
name: thaksorn-craft
description: "THaksornCraft — draft, rewrite, audit, and format Thai-language content and create Thai DOCX/PDF documents. Use for substantial Thai prose, academic or official Thai writing, Thai proofreading, Thai typography, new Word documents, new PDFs, and document-quality validation. Standalone: never requires OfficeCLI. Optionally coordinate with OfficeCLI or other document tools when available and appropriate, without modifying their installation."
---

# THaksornCraft

Produce natural, accurate Thai writing and professionally formatted Thai documents. Compose four independently maintained upstream skills: **kien-thai**, **kode-thai**, **thai-docx**, **thai-pdf**. This skill functions without OfficeCLI; OfficeCLI is an optional external executor, not a dependency.

## Locate resources (paths relative to this SKILL.md)

- Thai language rules: vendor/kien-thai/skills/kien-thai/SKILL.md; load only relevant files in vendor/kien-thai/skills/kien-thai/references/.
- Iterative Thai audit: vendor/kien-thai/skills/kode-thai/SKILL.md.
- Word renderer: vendor/thai-doc-skills/thai-docx/SKILL.md and scripts/.
- PDF renderer: vendor/thai-doc-skills/thai-pdf/SKILL.md and scripts/.
- Operational workflow: references/workflow.md.
- Platform and compatibility: references/compatibility.md.
- Optional OfficeCLI bridge, preflight and schemas: references/officecli-integration.md.
- Acceptance checks: references/quality-gates.md.
- Licensing/source attribution: references/licenses.md and SOURCES.lock.json.
- Standalone helper: scripts/thaksorn.py.

## Mandatory dispatch

1. Identify request: prose only, rewrite/audit, new DOCX, new PDF, edit existing document, or exact DOCX-to-PDF export.
2. Select one register **before** drafting. Choose academic/official for scholarly reports and official documents; do not default to informal blog wording. Read kien-thai SKILL.md and its register reference as needed. Keep source quotations, names, dates, numbers, citations, science terminology, equations, legal wording, and meaning intact.
3. For prose-only work, follow kien-thai and answer directly. If the user explicitly requests deep repeated review, use kode-thai principles with **up to 3 full passes by default**. Converge early when clean; if still unstable, report unresolved items rather than looping indefinitely. Never run unattended infinite loops.
4. For **new** Thai Word/PDF documents, finish and approve the Thai text first, then call the standalone renderer below. Word and PDF can be generated with **no OfficeCLI**. Render from one clean source or structured paragraph JSON; inspect results.
5. For **existing** Office documents, default to content-preserving targeted edits, not full reconstruction from Markdown. If no safe editor is available, report the limit. Use OfficeCLI only if already available, permitted and suitable; never install it automatically. For an existing DOCX-to-PDF export on Windows, the separate export-word command uses Microsoft Word COM (not OfficeCLI).
6. If the user specifically needs **identical pagination between DOCX and PDF**, generate or modify a final DOCX first, then export THAT DOCX through an available Word-compatible exporter and verify pages/TOC/figures. Direct WeasyPrint PDF and python-docx/Word are different renderers and may paginate differently.
7. Perform appropriate QA from references/quality-gates.md. Do not claim a file/preview passed unless actually tested.

## Standalone rendering (does NOT invoke OfficeCLI)

Read references/compatibility.md before running. Python with the relevant optional dependencies is needed for file rendering. The skill can always perform language review without these packages.

- Diagnose: python scripts/thaksorn.py doctor --json --probe-pdf
- Create Word: python scripts/thaksorn.py render input.md --output result.docx --preset saraban --font "TH SarabunPSK"
- Create Word with optional OfficeCLI schema QA (no overwrite): python scripts/thaksorn.py render input.md --output result.docx --officecli-qa auto --repair-known-schema; if upstream XML needs repair, deliver the reported recommended_output .normalized.docx, not the unvalidated original
- Require OfficeCLI QA before creating DOCX: use --officecli-qa require (will block before writing if OfficeCLI is unavailable)
- Create PDF (WeasyPrint/Pango): python scripts/thaksorn.py render input.md --output result.pdf --preset saraban --font "TH SarabunPSK"
- Create PDF on Windows with installed Word when Pango is unavailable: python scripts/thaksorn.py render input.md --output result.pdf --pdf-backend word (requires optional pywin32, retains source DOCX companion)
- Export an existing DOCX to PDF without OfficeCLI: python scripts/thaksorn.py export-word input.docx --output result.pdf (Windows + Word + pywin32)
- Structured input: supply a JSON object with a paragraphs array of document blocks, including table/image/pagebreak. The Markdown parser supports only a simple subset.

Use A4 portrait, Arabic numerals, and TH SarabunPSK as defaults **when there is no stricter user/project specification**. Never silently replace requested fonts or change document geometry. The runner does not overwrite an existing output unless given explicit --overwrite. It does not edit existing DOCX/PDF files.

## OfficeCLI optional interface

Treat OfficeCLI as an independent tool. If available and authorized, use it only for tasks it handles better (existing Word edits, layout-preserving operations, document validation). Load the installed thai-docx skill and, for PDF, thai-pdf before **every** OfficeCLI invocation on PC2. Never modify the OfficeCLI installation, settings, other projects, credentials, or global skills.

The optional bridge is `scripts/officecli_adapter.py` and can also be invoked as `python scripts/thaksorn.py officecli ...`. It must be given the **exact authorized job directory** with `--workspace`. Read `references/officecli-integration.md` for details.

```text
python scripts/thaksorn.py officecli --workspace "C:\path\to\authorized\job" status
python scripts/thaksorn.py officecli --workspace "C:\path\to\authorized\job" inspect "draft.docx" --mode issues
python scripts/thaksorn.py officecli --workspace "C:\path\to\authorized\job" validate "draft.docx"
python scripts/thaksorn.py officecli --workspace "C:\path\to\authorized\job" replace-copy "draft.docx" --output "revised.docx" --find "ข้อความเดิม" --replace "ข้อความใหม่"
python scripts/thaksorn.py officecli --workspace "C:\path\to\authorized\job" repair-schema-copy "generated.docx" --output "validated.docx"
```

If OfficeCLI is missing, `status` reports `STANDALONE_ONLY`; new Word/PDF creation and Thai-language editing remain operational via their own backends. Never auto-install OfficeCLI. The optional `inspect --allow-basic-fallback` provides **basic ZIP/text inspection**, not full OfficeCLI validation. Exact text replacement is performed only on a new copy; no in-place mutation. If schema validation fails, report the actual errors; for known bundled-DOCX XML element-order issues, `repair-schema-copy` generates a new copy and revalidates through OfficeCLI when available.

## Safe execution

- Honor user/project workspace boundaries. Keep temporary files and outputs inside the authorized job directory.
- Read/inspect before editing. Never access unrelated sibling project folders or write outside the selected scope without explicit approval.
- Never overwrite original documents, change upstream vendored source files, install system packages, start a local model, or fetch external resources automatically.
- Ollama/Typhoon draft route is **optional** and disabled by default; only invoke on explicit request after inspecting the model and its local endpoint. The native-language advantage is not proven and does not guarantee accuracy.
- Preserve upstream MIT notices and independent national-font licensing in the distribution.

## Expected result

Return the final Thai prose or document path, chosen register, engine, dependencies used, meaningful QA results, and explicit limitations. If delivery is blocked, provide the completed work and exact remaining blocker instead of stating that everything passed.
