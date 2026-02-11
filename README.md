# Legal Document Formatter

Format raw legal text to match a DOCX template using an LLM. Upload a template (summons, complaint, motion, etc.), paste your text, and get a formatted document that preserves the template’s styles, numbering, alignment, and structure.

## Features

- **Template-driven formatting** — Extracts styles and structure from your DOCX (headings, list styles, captions, spacing).
- **LLM segmentation** — Uses Azure OpenAI (gpt-4o-mini) or Google Gemini to split and label text into blocks (block type + text).
- **Visual reference** — Optionally sends template page images to the LLM so layout and structure match the template.
- **Editor** — Edit the formatted result in the app (Quill/Lexical or plain text) and download as DOCX.
- **Legal formatting** — Numbered allegations with hanging indent, justified body text, centered court/section headers, spacing between paragraphs.

## Requirements

- Python 3.10+
- An LLM provider: **Azure OpenAI** (recommended) or **Google Gemini**

## Setup

1. **Clone and create a virtual environment**

   ```bash
   cd formatter
   python -m venv .venv
   source .venv/bin/activate   # Windows: .venv\Scripts\Activate.ps1
   ```

2. **Install dependencies**

   ```bash
   pip install -r requirements.txt
   ```

3. **Configure environment**

   Copy `.env.example` to `.env` and set your API keys.

   **Azure OpenAI (recommended)**

   ```env
   AZURE_OPENAI_API_KEY=your-azure-openai-key
   AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com
   AZURE_OPENAI_DEPLOYMENT=gpt-4o-mini
   ```

   If these are set, the app uses Azure. Optionally set `AZURE_OPENAI_API_VERSION` (default: `2024-02-15-preview`).

   **Google Gemini (alternative)**

   ```env
   GEMINI_API_KEY=your-gemini-key
   FORMATTER_LLM_MODEL=gemini-2.5-flash
   ```

   Used only when Azure env vars are not set. Requires `pip install google-genai`.

## Run the app

**Streamlit (default)**

```bash
streamlit run app.py
```

Then open the URL shown (e.g. http://localhost:8501).

1. Upload a DOCX template.
2. Paste your raw legal text.
3. Click **Format with LLM**.
4. Edit in the editor if needed, then **Download document (.docx)**.

## Project structure

| Path | Description |
|------|-------------|
| `app.py` | Streamlit UI: upload template, paste text, format, editor, download. |
| `backend.py` | `process_document()`, `extract_and_store_styles()`; builds DOCX from template + LLM blocks. |
| `utils/llm_formatter.py` | LLM calls: Azure OpenAI or Gemini; format and slot-fill prompts; JSON parsing. |
| `utils/style_extractor.py` | Extract styles, template structure, style guide, line samples from DOCX. |
| `utils/formatter.py` | `inject_blocks()`: inject (block_type, text) into template; alignment, numbering, spacing. |
| `utils/html_to_docx.py` | Convert editor HTML to DOCX (download-from-editor path). |
| `utils/docx_to_images.py` | Convert template DOCX pages to images for LLM vision. |

## Optional: FastAPI + React frontend

If present in the repo:

- **Backend:** From project root, `uvicorn api.main:app --reload --port 8000`
- **Frontend:** `cd frontend && npm install && npm run dev`

The React app proxies `/api` to the FastAPI backend. Use the same `.env` for API keys.

## Output

- **Download document (.docx)** — Template-formatted DOCX (same styles, numbering, alignment as template).
- **Download from editor** — DOCX built from current editor content (includes your edits; generic styling).

## Documentation

- `ARCHITECTURE.md` — Pipeline and design (slot-fill, style-only injection).
- `TEMPLATE_GUIDE.md` — How to prepare DOCX templates.

## License

See repository license file.
