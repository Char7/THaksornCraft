# Workflow and task routing

## Choose the work path

| User goal | Language stage | Renderer/tool stage |
|---|---|---|
| Draft prose | kien-thai | No renderer required |
| Edit Thai prose | kien-thai | No renderer required |
| Rigorous iterative polish | kien-thai + bounded kode-thai | No renderer required |
| New Word (.docx) | Review/approve prose first | Standalone bundled thai-docx |
| New PDF (.pdf) | Review/approve prose first | Standalone bundled thai-pdf, subject to platform dependencies |
| Preserve existing Word structure | Review only explicitly selected text | A non-destructive targeted document editor; OfficeCLI optional |
| Match Word and PDF pages | No semantic changes after approval | Export final DOCX with Word-compatible engine and validate |
| Windows PDF with no Pango | Review/approve prose first | Optional Microsoft Word COM backend (does not require OfficeCLI) |

## Register protocol

Determine document audience, purpose, and voice before writing. Load vendor/kien-thai/skills/kien-thai/references/register.md only when needed; it contains academic, official/minutes, news/reference, explainer, marketing and personal-blog registers. For an official form or regulation, official language may appropriately contain formal terminology that is not an AI style defect. For scientific learning material, readability must not sacrifice scientific meaning.

Protect content in drafting and auditing:
- Exact names, titles, addresses, dates, amounts, percentages, tables, references and quotations
- Formulas, symbols, units and scientific terms
- Legal citations and fixed terms
- The document's intended register, headings, numbered sections and all explicit user instructions

Compare before/after at semantic level; reject any edit that changes a protected item or statement of fact. Do not infer missing sources.

## Iterative review

Use kode-thai rules when requested, but default to a maximum of 3 audit/fix passes. Log the most important change classes; if another pass would change wording without increasing factual or stylistic correctness, stop. For a large document, scope review to sections and maintain a change summary. Do not endlessly loop.

## Standalone rendering

Use Python 3 and required dependencies within an isolated environment. Work from an approved UTF-8 Markdown/TXT document or a JSON block list. CLI:

    python scripts/thaksorn.py doctor --json
    python scripts/thaksorn.py render content.md --output output.docx --preset saraban --font "TH SarabunPSK"
    python scripts/thaksorn.py render content.json --output output.pdf --preset saraban --font "TH SarabunPSK" --toc
    python scripts/thaksorn.py render content.md --output output.pdf --pdf-backend word
    python scripts/thaksorn.py export-word finished.docx --output finished.pdf

The --pdf-backend word option requires Windows + Microsoft Word + optional pywin32. It **does not call OfficeCLI**; it generates and retains a companion *.source.docx for traceability. The export-word command leaves the source DOCX unchanged.

Available presets: saraban, default, book; explicit options override preset. Supports --page-number (footer-center, footer-right, header-right), --page-number-format, --header, --footer, --title, --toc, --overwrite. Do not overwrite unless user authorized it.

Input JSON:
- Direct array of paragraph items, or object with "paragraphs": [...]
- Blocks: body, title, subtitle, heading1, heading2, heading3, bullet, number, quote, caption, table, image, pagebreak
- Table sample: {"type":"table","header":true,"rows":[["หัวข้อ","ผล"],["ก","100"]]}
- Image sample: {"type":"image","path":"local-figure.png","width":12,"caption":"ภาพที่ 1"}
- Image paths must be inside the authorized project. Do not fetch remote images or trust user-supplied script inputs.

Markdown input is intentionally basic; complex tables/images require JSON. If a user supplies a complex Word file, do not flatten and recreate it just to use this CLI.

## QA boundary

Open the produced file with an appropriate renderer when possible. For DOCX inspect XML fonts/sections/tables/images, refresh TOC in Word-compatible engine if available, and view printed pages. For PDF inspect copy/paste Thai, glyph completeness, embedded fonts, page break positions and table flow. No success claim without checks. If no renderer is available, state "generated, visual QA not run". See quality-gates.md.
