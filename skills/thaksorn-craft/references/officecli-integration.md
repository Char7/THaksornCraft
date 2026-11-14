# THaksornCraft × OfficeCLI — Optional integration

Edition: 0.2.0. THaksornCraft works without OfficeCLI. This guidance applies only when a compatible OfficeCLI executable is already installed and document work would benefit from it. The OfficeCLI installation and its global skills/configuration are NEVER modified.

## Architecture and policy

- THaksornCraft handles Thai writing, register-aware review, and standalone DOCX/PDF generation.
- The bounded optional bridge is scripts/officecli_adapter.py (also accessible through scripts/thaksorn.py officecli).
- OfficeCLI is an independent process; the bridge calls its documented CLI, not its internals.
- All document paths must resolve inside the exact selected --workspace. No sibling-project access.
- No OfficeCLI installation, resident process, or arbitrary command pass-through. Never overwrite existing documents.
- On PC2, before EVERY OfficeCLI invocation, read installed thai-docx/SKILL.md and, if PDF is involved, thai-pdf/SKILL.md. Review project/machine policies. The bridge checks bundled guidance exists but cannot attest that the agent read global instructions.
- For protected fields/formulas/URLs, never insert word-segmentation ZWSP or alter values to make Thai look prettier.

## Usage

From an explicitly authorized job project, call:

    python skills/thaksorn-craft/scripts/thaksorn.py officecli --workspace . status
    python skills/thaksorn-craft/scripts/thaksorn.py officecli --workspace . inspect draft.docx --mode issues
    python skills/thaksorn-craft/scripts/thaksorn.py officecli --workspace . inspect draft.docx --mode text
    python skills/thaksorn-craft/scripts/thaksorn.py officecli --workspace . validate draft.docx
    python skills/thaksorn-craft/scripts/thaksorn.py officecli --workspace . replace-copy draft.docx --output revised.docx --find "ข้อความเดิม" --replace "ข้อความใหม่"
    python skills/thaksorn-craft/scripts/thaksorn.py officecli --workspace . replace-copy draft.docx --output revised.docx --find "ข้อความเดิม" --replace "ข้อความใหม่" --dry-run
    python skills/thaksorn-craft/scripts/thaksorn.py officecli --workspace . repair-schema-copy generated.docx --output normalized.docx

The last command repairs a KNOWN narrow class of ordering problems in DOCX produced by the bundled Thai renderer. Never normalize a digitally signed or arbitrary user file without a separate case-specific review. It always writes a new copy, never the source file.

## One-command generation + OfficeCLI QA

The standalone document CLI supports optional QA on NEW DOCX, while retaining
a fully standalone default:

    python skills/thaksorn-craft/scripts/thaksorn.py render input.md --output result.docx --officecli-qa auto --repair-known-schema
    python skills/thaksorn-craft/scripts/thaksorn.py render input.md --output result.docx --officecli-qa require --repair-known-schema

When OfficeCLI is installed, the renderer builds a Word file, the bridge checks
it, and (only with --repair-known-schema) creates a separate result.normalized.docx
if the known upstream OOXML tag-order issue is found. The command reports
recommended_output pointing at the schema-validated copy. Without OfficeCLI,
--officecli-qa auto creates a standalone Word file and reports the QA was not
performed; --officecli-qa require fails BEFORE writing. The options are DOCX-only.

## Operations

| Operation | OfficeCLI required? | Existing input altered? |
|---|---|---|
| status | No | No |
| inspect | Usually; fallback DOCX ZIP text with --allow-basic-fallback | No |
| validate | Yes | No |
| replace-copy | Yes | No, produces new DOCX |
| repair-schema-copy | No; OfficeCLI needed for definitive schema QA | No, produces new DOCX |

The bridge only supports those operations. Other OfficeCLI functionality remains available separately to the Agent through the OfficeCLI skill under its policies; do not claim those untested tasks are covered by the bridge.

Read-only inspect supports DOCX/XLSX/PPTX and text/issues/stats/outline/annotated views. Exact-text replace-copy supports DOCX only. It checks an expected exact match and compares text, source SHA-256, and section/table/image/header/footer counts. It does not establish pixel-perfect layout; run visual QA separately.

An OfficeCLI view issues report with 0 issues is NOT identical to OpenXML schema validation. The bundled DOCX renderer can generate misordered WordprocessingML elements bCs/iCs and font-embedding settings; these are exposed as SCHEMA_ISSUES by OfficeCLI validate. An explicit repair-schema-copy of the tested source produced 0 schema errors, unchanged text and structure, and source SHA256 unchanged.

## Standalone and failure modes

- No OfficeCLI: status reports STANDALONE_ONLY; the language and standalone document features still work.
- inspect --allow-basic-fallback provides a limited DOCX text/ZIP read, NOT OfficeCLI schema validation.
- replace-copy with no OfficeCLI: BLOCKED; does not create a file.
- No exact text match: BLOCKED before writing.
- Existing output: BLOCKED (no automatic overwrite).
- Any failure during an OfficeCLI write: preserve the original; an output copy may remain as unapproved evidence. Never claim the copy is validated.
- repair-schema-copy with no OfficeCLI: NORMALIZED_NEEDS_SCHEMA_VALIDATION, not PASS.

## PC2 evidence

Tested on 2026-11-14 using OfficeCLI 1.0.161 found at officecli.exe (from PATH). Verified detection, document inspect, truthful schema failure, copy-edit with unchanged original and exact text, workspace isolation, graceful missing OfficeCLI behavior, and narrow schema repair producing a valid OfficeCLI OpenXML result. No OfficeCLI installation/settings or other projects touched.

This is an optional bounded adapter, not a full replacement for the OfficeCLI Skill. Standalone WeasyPrint PDF remains blocked on this PC2 by missing system Pango/GObject runtime; the separate Word COM PDF path does not use OfficeCLI.
