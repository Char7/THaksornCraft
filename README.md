# THaksornCraft

**Thai writing, proofreading and Word/PDF document production for AI Agents.**

THaksornCraft combines kien-thai (https://github.com/chakrit/kien-thai)
and thai-doc-skills (https://github.com/thaigqsoft/thai-doc-skills), preserving
their respective code and font license notices.

**Standalone first:** OfficeCLI is NOT required. If available, an optional
OfficeCLI adapter handles discovery, read-only inspection, schema validation,
exact-text replacement on a new copy, and known OOXML element-order repairs.

## Install and use

Clone this repository; copy the complete skills/thaksorn-craft directory into
your Agent's skills location, or load its SKILL.md directly.

Install the Word runtime (OfficeCLI is not required):

    python -m pip install -r skills/thaksorn-craft/requirements-docx.txt

Create Word:

    python skills/thaksorn-craft/scripts/thaksorn.py render tests/fixtures/science.md --output report.docx

Optional OfficeCLI status:

    python skills/thaksorn-craft/scripts/thaksorn.py officecli --workspace . status

Create Word and optionally validate + normalize schema via OfficeCLI:

    python skills/thaksorn-craft/scripts/thaksorn.py render tests/fixtures/science.md --output report.docx --officecli-qa auto --repair-known-schema

When OfficeCLI isn't installed, standalone DOCX continues to work. For stricter
preflight use --officecli-qa require. The command outputs recommended_output if
a separate normalized copy was produced. Never overwrite originals.

PDF: install optional requirements-pdf.txt plus native Pango/GObject libraries,
then render to a .pdf extension. On Windows with Microsoft Word installed,
--pdf-backend word is a separate optional PDF path using pywin32, not OfficeCLI.

## Font distribution note

Font binaries are NOT included in this GitHub source repository.
For optional embedded Thai national fonts, obtain original assets from
https://github.com/thaigqsoft/thai-doc-skills and follow the separate DIP/SIPA
National Font License. Place optional font files under the appropriate
vendor/thai-doc-skills/thai-docx/fonts/national directory and the equivalent
thai-pdf/fonts/national directory. Without font binaries the renderer cannot
guarantee font embedding, and display may depend on installed system fonts.

Source code from kien-thai and thai-doc-skills is MIT licensed (see LICENSE
files under vendor). National fonts have their own non-MIT license. The
THaksornCraft-specific integration contributions have no separate open-source
license granted yet; seek approval before redistributing those parts.

## Tests and limits

    python -m unittest discover -s tests -p "test_*.py" -v

Developer QA on PC2 using OfficeCLI 1.0.161: 27 tests, 26 passed and 1 skipped.
This does not guarantee pixel-perfect Word/PDF rendering, cross-platform
compatibility, or accuracy of all Thai official/scientific writing. WeasyPrint
on the tested Windows PC required missing Pango/GObject native DLLs.

Full integration documentation:
skills/thaksorn-craft/references/officecli-integration.md

Pinned upstream commits:
skills/thaksorn-craft/SOURCES.lock.json

Status: v0.2.0 developer preview, not production-certified.
