# Reconstruction Guide

## Choosing the right output mode

Use a hybrid deck for most requests: full-slide image background for visual fidelity plus editable text overlays for practical editing. Use an editable-only deck only when the user accepts visual differences and wants the main content rebuilt manually.

## Choosing OCR/local/API routing

Prefer this decision tree:

1. If the user cannot send slide images outside the machine, use local OCR.
2. If local OCR is installed, use `--ocr auto` or `--ocr local`.
3. If local OCR is missing and the environment is interactive, let the script ask the user to choose:
   - automatic local OCR installation,
   - OpenAI API OCR configuration,
   - or background-only PPT generation.
4. If the user chooses local installation, run with `--ocr local --auto-install-local` or choose option 1 in `--ocr auto`.
5. If the user chooses API OCR, run with `--ocr api`, read or prompt for `OPENAI_API_KEY`, and remind them that slide images are sent to the API.
6. If neither local nor API OCR is available, run with `--ocr none` and produce a visual-reference deck.

## Recommended phases

1. Background pass: one slide image per slide, exact aspect ratio.
2. Text pass: local OCR, API OCR, or user-provided sidecar JSON creates editable text boxes.
3. Layout pass: manually recreate titles, major body blocks, simple callouts, and obvious rectangles/lines.
4. Cleanup pass: remove duplicate OCR artifacts, normalize fonts, check line breaks, and keep image references only where useful.

## What to recreate as native PowerPoint objects

Good candidates:
- Titles and subtitles
- Body paragraphs and bullet lists
- Simple rectangles, lines, arrows, and divider rules
- Large background color panels
- Simple tables with clearly readable text

Poor candidates:
- Photos
- Logos/icons without vector originals
- Dense screenshots
- Complex charts without source data
- Decorative gradients and shadows

## Communication template

When reporting results to the user, use this structure:

- Created: output file path
- Source: number of input images
- Editable elements: text boxes and any recreated shapes
- OCR route: local, API, JSON, or none
- Reference layer: visible, faded, or absent
- Limitations: items left as image content
- Next improvement: one suggested action, such as providing OCR JSON, choosing API OCR, or providing source chart data
