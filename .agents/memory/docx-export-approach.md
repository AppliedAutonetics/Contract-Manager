---
name: DOCX export approach
description: How the Word export is implemented and why — replaces old custom python-docx HTML parser
---

## Rule
Word export uses **html2docx** (pip) + python-docx post-processing, NOT a custom HTML→DOCX parser.

## Why
The previous approach hand-mapped every HTML tag to python-docx calls. This was ~500 lines of buggy code that silently dropped tables, mis-read CSS, and diverged from the PDF. LibreOffice headless was attempted but crashes with SIGSEGV in the Nix/Replit environment (font/library incompatibility).

## How it works (generate_docx)
1. `_build_contract_html(body_html)` — single source of truth for the CSS used by BOTH PDF and Word exports.
2. Pass the body HTML to `html2docx.html2docx(full_html, title=...)` — handles tables, headings, bold/italic, lists, blockquotes.
3. Open result with `python-docx`, call `_apply_docx_styles(doc)` which:
   - Sets Arial/sizes/colours on Named styles (Normal, Heading 1–6)
   - Iterates every run in paragraphs AND table cells and sets `run.font.name = 'Arial'` to prevent Word theme-font override
   - Sets docDefaults at XML level as last-resort fallback
4. Set 1-inch page margins on all sections.

## Key constraints
- `html2docx` does NOT read inline CSS for font properties — styling must be applied post-conversion via python-docx.
- LibreOffice from Nix (7.0.6.2) crashes (rc=-11/SIGSEGV) — do NOT attempt again unless the Nix environment changes.
- On the production Ubuntu server, `html2docx` is also the right approach (same package, same behaviour).
- `_build_contract_html` is shared by both `_generate_pdf_html` (xhtml2pdf) and `generate_docx` — keep them in sync.

## How to apply
Any future change to the CSS/layout must be made in `_build_contract_html()` and verified in both PDF and Word exports.
