# Quality gates for THaksornCraft

## Language checks (all outputs)

1. Correct register: official/academic versus public posts or informal prose
2. Preserve exact facts, quotations, citations, scientific terms, entities, numbers, units, formulas and headings
3. Avoid grammar/voice drift, unnecessary repeated connectives and machine-translated sentence structure
4. Do not use a mechanical forbidden-word list to replace legitimate formal expressions in official content
5. Use bounded iterative review (3 passes default); report unresolved issues rather than falsely asserting CLEAN
6. Proofread final text, not just the source prompt

## DOCX checks

- Correct A4 portrait dimensions (unless requested otherwise) and margins
- Correct Word styles, complex script font w:cs, and font naming
- No split Thai words, broken diacritics, unintentional ZWSP in URLs, or invented content
- Tables fit within margins; no stray red underlining; captions and figures keep their positions and aspect ratio
- Fields/TOC/page numbers present as requested; refresh field results when the editor supports it
- If editing existing Word, confirm the original tables/images/sections/styles remain; do not silently recreate the file
- Inspect render/print preview where possible, not only ZIP/XML signatures

## PDF checks

- PDF opens, has the expected number of pages and embedded fonts
- Thai diacritics, glyphs, and copy/paste text are valid
- No clipped tables, blank pages, orphan headings, misplaced images or footer overlaps
- If requested to MATCH a final DOCX, compare page count, TOC page references, figures and layout page-by-page. The standalone PDF engine alone does not guarantee a match.

## Testing categories

- Prose: press release, official memo, science lesson and academic abstract; compare meaning before/after
- Document blocks: headings, numbered list, bold/italic, table, caption, image, pagebreak
- Fonts: TH SarabunPSK; English+Thai mixed; punctuation, math notation and Arabic digits
- Windows: DOCX generation, PDF native Pango dependencies and export fallback
- Standalone: OfficeCLI absent; no OfficeCLI invocation or import
- Optional integration: OfficeCLI present — still works when unavailable
- Licensing: bundled MIT notices and national font license intact

A test that has not been run should be reported NOT TESTED. An error should be reported BLOCKED. Do not equate a valid ZIP or successful Python return with visually approved document.
