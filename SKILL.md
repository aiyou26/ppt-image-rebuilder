---
name: ppt-image-rebuilder
description: rebuild editable powerpoint files from ppt-exported slide images such as png, jpg, or jpeg screenshots. use this when the user has images exported from slides and wants a best-effort editable .pptx, OCR text overlay, OpenAI API vision text extraction, OpenAI-compatible third-party API vision text extraction, local Tesseract OCR setup, layout reconstruction, visual reference backgrounds, or a repeatable workflow for converting static slide images back into powerpoint decks.
---

# PPT Image Rebuilder

## Goal

Turn a folder of PPT-exported slide images into a best-effort editable `.pptx`. Be explicit that image-to-PPT reconstruction cannot perfectly recover the original object tree. The preferred deliverable is a visually faithful PowerPoint with the original image as a reference/background plus editable text boxes from local OCR, OpenAI or OpenAI-compatible third-party API vision OCR, or a sidecar JSON file.

## Default Workflow

1. Collect inputs:
   - A folder of slide images in page order, or a ZIP/archive containing them.
   - Desired output `.pptx` path.
   - Whether the user wants maximum visual fidelity, maximum editability, or a hybrid.
2. Inspect slide images:
   - Sort images naturally, such as `slide1.png`, `slide2.png`, `slide10.png`.
   - Infer aspect ratio from the first image unless the user specifies 16:9 or 4:3.
3. Choose the text extraction route:
   - Start with `--ocr auto` unless the user explicitly chooses another route.
   - If local Tesseract OCR is already installed, `--ocr auto` uses it.
   - If local OCR is missing and the script is interactive, it asks the user to choose automatic local installation, OpenAI API configuration, or no OCR.
   - If the user chooses local installation, attempt to install `pytesseract` and the Tesseract binary using the available system package manager.
   - If the user chooses API OCR, run the API configuration wizard. Support the official OpenAI API and third-party OpenAI-compatible APIs by collecting provider, API base URL, API key, model name, and endpoint mode, optionally saving reusable settings to `.env`, installing the `openai` Python package if needed, and calling the configured vision model.
4. Generate a first-pass editable deck:
   - Run `scripts/rebuild_pptx_from_images.py`.
   - Use `--background-mode full` for visually faithful output.
5. Review the output:
   - Check slide count, page order, aspect ratio, and visible alignment.
   - Confirm text boxes are selectable/editable in PowerPoint.
   - If OCR missed text, request or create a sidecar OCR JSON and rerun.
6. Improve editability only where it matters:
   - Rebuild important title/body text as editable text boxes.
   - Recreate simple rectangles, lines, and large color blocks as editable shapes when visually obvious.
   - Leave complex charts, photos, icons, and screenshots as images unless the user explicitly asks for manual reconstruction.

## Commands

Run from the skill directory or adjust script paths as needed.

Default interactive route. This uses local OCR if available; otherwise it asks whether to install local OCR, configure API OCR, or skip OCR:

```bash
python scripts/rebuild_pptx_from_images.py --input /path/to/slide-images --output /path/to/editable.pptx --ocr auto --background-mode full
```

Force local OCR and automatically attempt installation if missing:

```bash
python scripts/rebuild_pptx_from_images.py --input /path/to/slide-images --output /path/to/editable.pptx --ocr local --auto-install-local --background-mode full
```

Force API OCR. In interactive mode, the script opens a beginner-friendly wizard and asks whether to use the official OpenAI API or a third-party OpenAI-compatible API:

```bash
python scripts/rebuild_pptx_from_images.py --input /path/to/slide-images --output /path/to/editable.pptx --ocr api --api-config-wizard --background-mode full
```


Use a third-party OpenAI-compatible API non-interactively after configuration is known:

```bash
python scripts/rebuild_pptx_from_images.py --input /path/to/slide-images --output /path/to/editable.pptx --ocr api --api-provider openai-compatible --api-base-url https://provider.example.com/v1 --api-key-env PPT_REBUILDER_API_KEY --api-model your-vision-model --api-endpoint-mode chat-completions --background-mode full
```

Use a reference-only editable text layer over a faded image:

```bash
python scripts/rebuild_pptx_from_images.py --input /path/to/slide-images --output /path/to/editable.pptx --ocr auto --background-mode full --reference-opacity 0.35
```

Use sidecar OCR data:

```bash
python scripts/rebuild_pptx_from_images.py --input /path/to/slide-images --output /path/to/editable.pptx --ocr json --ocr-json /path/to/ocr.json
```

Skip OCR and create a visually faithful background deck:

```bash
python scripts/rebuild_pptx_from_images.py --input /path/to/slide-images --output /path/to/editable.pptx --ocr none --background-mode full
```

## OCR Mode Rules

- `--ocr auto`: Prefer installed local OCR. If local OCR is missing and interactive input is available, ask the user to choose local installation, API configuration, or no OCR. In non-interactive mode, continue with no OCR.
- `--ocr local`: Use local Tesseract OCR. With `--auto-install-local`, attempt to install missing Python and system dependencies. If installation cannot complete because package managers or privileges are unavailable, stop with a clear error.
- `--ocr api`: Use official OpenAI API or OpenAI-compatible third-party API vision OCR. Read provider settings from command-line flags, environment variables, or `.env`; if missing in interactive mode, run the API configuration wizard. Never hard-code API keys in the skill.
- `--ocr json`: Use sidecar OCR JSON and do not call OCR engines.
- `--ocr none`: Do not create editable text boxes.

## API Configuration Notes

- Treat API keys as secrets. Prefer environment variables or a local `.env` file with restricted permissions.
- The script supports two API provider modes:
  - `openai`: official OpenAI API. Uses `OPENAI_API_KEY` by default. A custom `OPENAI_BASE_URL` may be used only when explicitly configured.
  - `openai-compatible`: third-party API, gateway, or proxy that accepts OpenAI-style image requests. Uses `PPT_REBUILDER_API_KEY`, `PPT_REBUILDER_API_BASE_URL`, `PPT_REBUILDER_API_MODEL`, and `PPT_REBUILDER_API_ENDPOINT_MODE` by default.
- For third-party APIs, ask the user for the provider's API Base URL, API Key, and vision-capable model name. Most compatible providers should use `--api-endpoint-mode chat-completions`; `auto` tries Responses first and then Chat Completions.
- The script may save non-secret API settings to `.env` after asking. It may save API keys only after explicit user agreement or `--save-api-key`.
- API OCR sends slide images to the configured provider. Tell the user before using this mode if the slides contain sensitive material.
- API OCR can improve recognition for complex layouts, but coordinates are still estimates and should be reviewed.
- Dedicated OCR services with non-OpenAI request formats, such as vendor-specific OCR APIs, require a separate adapter before use. Do not imply they are supported by the generic `openai-compatible` mode.

## OCR JSON Format

Accept this flexible sidecar format:

```json
{
  "slides": [
    {
      "image": "slide001.png",
      "items": [
        {"text": "Quarterly Review", "x": 0.08, "y": 0.06, "w": 0.52, "h": 0.08, "font_size": 28, "confidence": 96}
      ]
    }
  ]
}
```

Coordinates may be normalized floats from 0 to 1, or absolute pixels relative to the source image. Prefer normalized coordinates in generated JSON.

## Reconstruction Rules

- Always preserve page order and slide count.
- Use the slide image as a full-slide reference unless the user asks for an editable-only rebuild.
- Do not claim the reconstructed deck is fully editable unless every major visual element has been recreated as PowerPoint objects.
- Prefer editable text boxes for OCR/API text; keep text placement close to the original, even if font family is approximate.
- Use PowerPoint-native shapes only for simple, obvious geometry. Do not waste time tracing complex graphics that are better left as images.
- Keep generated scripts and intermediate files in a predictable output folder when doing a multi-pass reconstruction.
- When the user asks for high fidelity, keep the original image visible. When they ask for editability, fade or remove the original image only after recreating the main content.

## Quality Checklist

Before returning the final deck, verify:

- The output `.pptx` opens without errors.
- Slide count matches the number of input images.
- Slide aspect ratio matches the source images or the user's requested ratio.
- Text boxes are selectable/editable in PowerPoint.
- The original image layer is present when fidelity is required.
- Limitations are clearly stated, especially for charts, complex diagrams, icons, and screenshots.

## Troubleshooting

- If `python-pptx` is missing, install it in the working environment or add it to the project dependencies.
- If local OCR produces no text, check whether `pytesseract`, the Tesseract binary, and the requested language pack are installed. Rerun with `--ocr api` or `--ocr none` if needed.
- If automatic local installation fails, the environment may not have `brew`, `apt-get`, `conda`, `winget`, or sufficient permissions. Use API OCR or install Tesseract manually.
- If API OCR fails, confirm the API key, internet access, the `openai` Python package, the selected provider, the base URL, the endpoint mode, and whether the selected model supports image input.
- If coordinates look wrong with sidecar JSON, confirm whether OCR boxes are normalized or pixel-based and that they correspond to the same source image dimensions.
- If output is blurry, use the highest-resolution slide exports available and avoid re-encoding the source images.
