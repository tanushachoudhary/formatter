"""Use an LLM to split and label text into styled blocks (block_type + text)."""

import json
import os
import re


def _sanitize_json_control_chars(raw: str) -> str:
    """Replace unescaped control characters inside JSON string values so json.loads succeeds."""
    # Match double-quoted string contents (handles \" inside)
    result = []
    i = 0
    while i < len(raw):
        if raw[i] == '"' and (i == 0 or raw[i - 1] != "\\"):
            result.append(raw[i])
            i += 1
            while i < len(raw):
                c = raw[i]
                if c == "\\" and i + 1 < len(raw):
                    result.append(c)
                    result.append(raw[i + 1])
                    i += 2
                    continue
                if c == '"':
                    result.append(c)
                    i += 1
                    break
                # JSON disallows unescaped control chars in strings; replace with space
                if ord(c) < 32:
                    result.append(" ")
                else:
                    result.append(c)
                i += 1
            continue
        result.append(raw[i])
        i += 1
    return "".join(result)

# Optional: use OpenAI or Azure OpenAI
try:
    from openai import AzureOpenAI, OpenAI
except ImportError:
    AzureOpenAI = None
    OpenAI = None

# Logical block types (fallbacks when block_type is not a template style name)
LOGICAL_BLOCK_TYPES = (
    "heading",
    "section_header",
    "paragraph",
    "numbered",
    "wherefore",
    "line",
    "signature_line",
    "page_break",
)

SYSTEM_PROMPT = """You format raw legal text so it has the same styling, structure, and formatting as the uploaded template document. Match the template's use of styles for titles, section headings, body paragraphs, numbered lists, and signature blocks.

Task: (1) Use the template structure and style guide below to label every part of the raw text. (2) Output one block per logical segment with the correct block_type (exact template style name or line/signature_line/page_break).

Rules:
- Document title / main heading -> use the template's title style (e.g. Heading 1) for the main document title (court name, case caption, MOTION TO RESTORE, etc.).
- Section headers -> use the template's section style (e.g. Heading 2) for each section heading (VERIFIED COMPLAINT, MEMORANDUM, FACTS, ARGUMENT, CONCLUSION, RELIEF REQUESTED, party captions, etc.).
- Case number / date / caption -> use the template style used for Index No., Date Filed, or similar (often right-aligned).
- Body paragraphs -> use the template's Normal (or body) style for narrative, argument, facts, verification text. Output plain text only (no markdown).
- Numbered or bulleted lists -> use the template's list style (e.g. List Number) for each list item wherever the template uses that style (allegations, motion grounds, relief items, etc.); one block per item.
- Addresses -> same style as in template (often Normal; one block per line if the template breaks them).
- Signature block -> use the template style for that block; use block_type signature_line for the underline line.
- Separator lines (dashes/dots ending in X) -> block_type line, exact line characters in text.
- Page breaks -> output page_break (empty text) before each new major section, matching the template (e.g. before MEMORANDUM, CONCLUSION, RELIEF REQUESTED, VERIFIED COMPLAINT, ATTORNEY'S VERIFICATION, NOTICE OF ENTRY, etc.).
- Checkboxes: Use [ ] and [x] in text; they render as checkbox symbols.

Preserve exact wording. Output plain text in the text field only (no markdown). One block per logical segment.

Reply with a JSON array only. Each element: {"block_type": "<exact style name from template or line/signature_line/page_break>", "text": "<content>"}."""


def _call_openai(text: str, style_schema: dict) -> list[tuple[str, str]]:
    """Call OpenAI or Azure OpenAI API; returns list of (block_type, text)."""
    if not OpenAI and not AzureOpenAI:
        raise RuntimeError("openai package not installed. pip install openai")

    style_guide = (style_schema.get("style_guide") or style_schema.get("style_guide_markdown") or "").strip()
    if not style_guide:
        style_list = style_schema.get("paragraph_style_names", []) or list(style_schema.get("style_map", {}).values())
        style_guide = "Style names: " + ", ".join(style_list)

    line_samples = style_schema.get("line_samples", [])
    line_note = ""
    if line_samples:
        examples = [s.get("text", "")[:60] + ("..." if len(s.get("text", "")) > 60 else "") for s in line_samples[:5]]
        line_note = f"\nLine/separator samples (block_type 'line' or 'signature_line'): {examples}\n"

    # Template content: how each paragraph was styled in the uploaded DOCX (so LLM can extract and apply formatting)
    template_content = style_schema.get("template_content", [])
    template_section = ""
    if template_content:
        lines = []
        for item in template_content:
            style_name = item.get("style") or "Normal"
            para_text = (item.get("text") or "").strip()
            lines.append(f"[{style_name}]: {para_text}" if para_text else f"[{style_name}]:")
        template_section = "Template document (each paragraph with its style name):\n" + "\n".join(lines) + "\n\n"

    user_content = f"""{template_section}Extracted style guide (use these exact style names as block_type):

{style_guide}
{line_note}

---

Raw legal text to format. Match the template structure above: use the same styles for titles, section headings, body paragraphs, and numbered lists as in the template. Insert page_break before each major section (as in the template). Use template style for case number/date and addresses/signature block. Output plain text only in the text field. Output a JSON array of {{"block_type": "<style name or line/signature_line/page_break>", "text": "<content>"}}.

---
{text}
---"""

    # Prefer Azure OpenAI if endpoint and key are set
    azure_key = os.environ.get("AZURE_OPENAI_API_KEY") or os.environ.get("AZURE_OPENAI_KEY")
    azure_endpoint = os.environ.get("AZURE_OPENAI_ENDPOINT")

    if azure_key and azure_endpoint:
        if not AzureOpenAI:
            raise RuntimeError("Azure OpenAI requested but openai package may be too old. pip install openai>=1.0.0")
        client = AzureOpenAI(
            api_key=azure_key,
            api_version=os.environ.get("AZURE_OPENAI_API_VERSION", "2024-02-15-preview"),
            azure_endpoint=azure_endpoint.rstrip("/"),
        )
        model = os.environ.get("AZURE_OPENAI_DEPLOYMENT") or os.environ.get("FORMATTER_LLM_MODEL", "gpt-4o-mini")
    else:
        api_key = os.environ.get("OPENAI_API_KEY")
        if not api_key:
            raise ValueError(
                "Set OPENAI_API_KEY for OpenAI, or AZURE_OPENAI_API_KEY and AZURE_OPENAI_ENDPOINT for Azure OpenAI"
            )
        client = OpenAI(api_key=api_key)
        model = os.environ.get("FORMATTER_LLM_MODEL", "gpt-4o-mini")

    resp = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_content},
        ],
        temperature=0.1,
    )
    raw = resp.choices[0].message.content.strip()
    # Strip markdown code fence if present
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
    # Remove unescaped control characters inside JSON strings (LLM sometimes emits literal newlines/tabs)
    raw = _sanitize_json_control_chars(raw)
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        # Fallback: replace all control chars in whole string then retry
        raw_fallback = re.sub(r"[\x00-\x1f]", " ", raw)
        data = json.loads(raw_fallback)
    out = []
    for item in data:
        bt = (item.get("block_type") or "paragraph").strip()
        if not bt:
            bt = "paragraph"
        # Accept any block_type: template style name or logical type (heading, paragraph, line, etc.)
        out.append((bt, item.get("text", "").strip()))
    return out


def format_text_with_llm(text: str, style_schema: dict) -> list[tuple[str, str]]:
    """Use LLM to convert raw text into list of (block_type, text). Uses OpenAI if available."""
    return _call_openai(text, style_schema)
