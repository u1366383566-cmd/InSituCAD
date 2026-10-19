# InSituCAD · AI Translation Workbench for CAD Drawings, Documents & Spreadsheets

A locally-deployed translation tool for architectural interior detailing and custom furniture industries: a **vision LLM reads drawings and translates them directly**, with translations **reflowed in-place** onto the CAD drawings. Regular PDF documents (brochures, manuals, contracts) are also supported with layout-preserving translation, Excel workbooks (.xlsx/.xlsm) are translated cell-by-cell with the full layout preserved, and Word documents (.docx) are translated in place into editable docx files. Accessed via browser; the entire workflow runs on your machine or LAN.

> Note: The web UI is currently in Chinese. A future i18n release is planned.

## Features

- **CAD drawing mode** — the vision LLM reads and translates in a single step, without depending on OCR text accuracy. Text-box detection uses a remote PaddleOCR service (optional) with local RapidOCR as fallback. 1:1 in-place erase-and-refill preserves the original geometry, and rotated annotation directions automatically follow the source drawing (0/90/180/270° pages).
- **Brochure/manual mode** — powered by the optional [BabelDoc](https://github.com/funstory-ai/BabelDoc) flow-layout engine; outputs Chinese-only or English–Chinese bilingual versions.
- **Excel spreadsheet mode** — cell texts are translated in place (no OCR needed: spreadsheet text is exact). Embedded images, merged cells, fonts/colors, row heights & column widths, formulas and macros are preserved untouched; on machines with Excel installed, formulas are recalculated via COM automation so totals also display in read-only previewers, and a PNG preview of the translated sheet is generated. Multi-line cells keep their line structure. Chinese text baked inside embedded images (material cards, dimension callouts) can be translated too via OCR + in-image erase-and-refill, with photos kept byte-identical.
- **Word document mode** — paragraph texts are translated in place into an editable .docx: heading styles, numbering lists, tables and images are preserved untouched; mixed-format paragraphs keep their bold "label:" prefixes; CJK literals in auto-numbering definitions (第N节) are neutralized so translated headings stay language-consistent. A page preview is rendered via Word COM when available. Old binary .doc files must be saved as .docx first.
- **Any source language** — annotations in English, Vietnamese, Thai, or any language supported by your chosen vision model are translated as-is, preserving diacritics and casing. The source language is auto-detected per task.
- **Switch between any OpenAI-compatible engines** — plug in any OpenAI-compatible multimodal model service (DeepSeek, Zhipu, Qwen, SiliconFlow, local models, …). Each engine is verified with a live vision test before activation (text-only models are rejected). Switching engines never affects running tasks, and costs are tracked per engine.
- **Terminology base** — categorized terms with automatic usage counters, JSON/CSV import & export. Approved corporate translations always override model output.
- **Project history** — every task (including failures) is persisted to local `task_history.json`, filterable and searchable; survives restarts.
- **LAN sharing** — teammates on the same network use it from a browser with zero installation (the server listens on `0.0.0.0:7860`).

## Quick Start

### Option 1: one-click launch (Windows)

1. Create a virtualenv in the project folder and install dependencies (see below);
2. Double-click `启动翻译工作台.bat` — the browser opens automatically at `http://127.0.0.1:7860` once the service is ready.

### Option 2: manual launch

```bash
pip install fastapi uvicorn pymupdf numpy opencv-python requests rapidocr_onnxruntime
# Optional: BabelDoc engine for the brochure/manual mode
pip install babeldoc
# Optional: Excel/Word modes (openpyxl + python-docx required; pywin32 enables COM preview/recalc when Office is installed)
pip install openpyxl python-docx pywin32

python server.py
```

Then open `http://127.0.0.1:7860`. LAN colleagues can use `http://<your-lan-ip>:7860` (the actual address is printed on startup).

> CJK fonts: system fonts (SimHei / Microsoft YaHei) are used automatically — no extra downloads.

## Configuration

1. Copy `config.example.json` to `config.json`;
2. Fill in your engine's `api_key` under `providers` (or add engines online via *Settings → Translation Engines*, which auto-verifies vision capability);
3. `ocr_api_url` is an **optional** remote PaddleOCR endpoint that improves text-box recall. Without it, local RapidOCR takes over automatically and everything still works.

`config.json` structure (identical fields to `config.example.json`):

| Field | Description |
| --- | --- |
| `providers` | Engine list (OpenAI-compatible: base_url / api_key / vision_model / pricing / balance endpoint) |
| `active_provider` | id of the currently active engine |
| `ocr_api_url` | optional, remote PaddleOCR detection endpoint |
| `deepseek_api_key` | legacy single-key field, kept for compatibility |

## Project Layout

```
├── server.py                 # FastAPI backend (tasks, terms, engines, history)
├── cad_pdf_translator.py     # CAD translation pipeline (detect → read+translate → in-place refill) + Excel mode
├── web/index.html            # Single-page workbench frontend
├── design-system/            # UI design spec
├── 启动翻译工作台.bat          # Windows one-click launcher
├── 使用说明.txt               # End-user manual (Chinese)
├── config.example.json       # Config template (copy to config.json)
├── config.json               # Local config & keys (not committed, see .gitignore)
├── output/                   # Translation artifacts (not committed)
├── uploads/                  # Upload scratch dir (not committed)
└── 翻译缓存_deepseek.json     # Terminology cache (not committed)
```

## Security

- API keys live only in the local `config.json`, which is excluded via `.gitignore` and never committed;
- Frontend and API are same-origin (no CORS); download/preview endpoints validate against path traversal; uploaded filenames are sanitized;
- No hardcoded keys anywhere in the code. To use remote OCR detection, deploy your own PaddleOCR service and put the endpoint in `config.json`.
