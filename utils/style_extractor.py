"""Extract styles and formatting from a DOCX and store as JSON."""

import json
import os

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Length

STORE_DIR = "output"
EXTRACTED_STYLES_FILE = "extracted_styles.json"
EXTRACTED_STYLE_GUIDE_FILE = "extracted_style_guide.txt"

PREFERRED_HEADING_1 = ("Heading 1", "Title", "Titre 1")
PREFERRED_HEADING_2 = ("Heading 2", "Subtitle", "Titre 2", "Section")
PREFERRED_NORMAL = ("Normal", "Body Text", "Paragraphe")
PREFERRED_LIST = ("List Number", "List Paragraph", "List")


def _get_paragraph_style_names(doc):
    return [
        s.name for s in doc.styles
        if s.type == WD_STYLE_TYPE.PARAGRAPH
    ]


def _pick_style(available, preferred_names, fallback_names=None):
    for name in preferred_names:
        if name in available:
            return name
    if fallback_names:
        for name in fallback_names:
            if name in available:
                return name
    return available[0] if available else None


def _length_pt(val):
    """Serialize a Length or None to points (float) for JSON."""
    if val is None:
        return None
    if isinstance(val, Length):
        return getattr(val, "pt", None)
    return None


def _enum_name(val):
    """Get enum member name for JSON (e.g. WD_UNDERLINE.DOTTED -> 'DOTTED')."""
    if val is None:
        return None
    if hasattr(val, "name"):
        return val.name
    return str(val)


def _extract_paragraph_format(pf):
    """Extract paragraph format to a JSON-serializable dict (including line_spacing and line_spacing_rule)."""
    if pf is None:
        return {}
    out = {}
    for attr in ("alignment", "space_before", "space_after", "left_indent", "right_indent", "first_line_indent", "line_spacing"):
        try:
            val = getattr(pf, attr, None)
            if val is None:
                continue
            if attr == "alignment":
                out[attr] = _enum_name(val)
            elif attr == "line_spacing":
                if isinstance(val, Length):
                    out[attr] = val.pt
                elif isinstance(val, (int, float)):
                    out[attr] = val
                else:
                    out[attr] = None
            else:
                out[attr] = _length_pt(val) if hasattr(val, "pt") else val
        except Exception:
            pass
    try:
        rule = getattr(pf, "line_spacing_rule", None)
        if rule is not None:
            out["line_spacing_rule"] = _enum_name(rule)
    except Exception:
        pass
    return out


def _extract_run_format(run):
    """Extract run/font format to a JSON-serializable dict (underline, bold, font name/size, etc.)."""
    if run is None:
        return {}
    font = run.font
    out = {}
    try:
        if font.bold is not None:
            out["bold"] = font.bold
    except Exception:
        pass
    try:
        if font.italic is not None:
            out["italic"] = font.italic
    except Exception:
        pass
    try:
        u = font.underline
        if u is not None:
            out["underline"] = _enum_name(u) if hasattr(u, "name") else (True if u else False)
    except Exception:
        pass
    try:
        if font.name is not None:
            out["name"] = font.name
    except Exception:
        pass
    try:
        if font.size is not None and hasattr(font.size, "pt"):
            out["size_pt"] = font.size.pt
    except Exception:
        pass
    return out


def _format_from_style_definition(doc: Document, style_name: str) -> dict | None:
    """Extract paragraph_format (and optionally font) from the style definition when no paragraph uses it."""
    try:
        style = doc.styles[style_name]
    except KeyError:
        return None
    if style.type != WD_STYLE_TYPE.PARAGRAPH:
        return None
    pf = {}
    run_fmt = {}
    try:
        pf = _extract_paragraph_format(style.paragraph_format)
    except Exception:
        pass
    try:
        # Style's font is on the style element (rPr); get first run property set
        if hasattr(style, "font") and style.font:
            run_fmt = _extract_run_format_from_font(style.font)
    except Exception:
        pass
    if not pf and not run_fmt:
        return None
    return {"paragraph_format": pf, "run_format": run_fmt}


def _extract_run_format_from_font(font) -> dict:
    """Extract run format from a Font object (e.g. from a style)."""
    if font is None:
        return {}
    out = {}
    try:
        if font.bold is not None:
            out["bold"] = font.bold
    except Exception:
        pass
    try:
        if font.italic is not None:
            out["italic"] = font.italic
    except Exception:
        pass
    try:
        u = font.underline
        if u is not None:
            out["underline"] = _enum_name(u) if hasattr(u, "name") else (True if u else False)
    except Exception:
        pass
    try:
        if font.name is not None:
            out["name"] = font.name
    except Exception:
        pass
    try:
        if font.size is not None and hasattr(font.size, "pt"):
            out["size_pt"] = font.size.pt
    except Exception:
        pass
    return out


def _merge_format(base: dict, override: dict) -> dict:
    """Merge override into base; override wins when both have a key. For nested dicts, merge recursively."""
    out = dict(base)
    for k, v in override.items():
        if k not in out or v is not None and v != "" and (not isinstance(v, dict) or v):
            if isinstance(v, dict) and isinstance(out.get(k), dict):
                out[k] = _merge_format(out[k], v)
            else:
                out[k] = v
    return out


# Default spacing/indent when template has none. Use 0 so output matches tight,
# single-spaced templates; template-defined spacing is always preserved.
DEFAULT_SPACE_AFTER_PT = 0.0
DEFAULT_SPACE_BEFORE_HEADING_PT = 0.0
DEFAULT_FIRST_LINE_INDENT_PT = 0.0
DEFAULT_NUMBERED_LEFT_INDENT_PT = 0.0
DEFAULT_NUMBERED_FIRST_LINE_INDENT_PT = 0.0


def _sample_formatting_per_style(doc: Document) -> dict:
    """For each paragraph style used in the doc, sample one paragraph and extract its formatting."""
    style_formatting = {}
    seen_styles = set()
    for para in doc.paragraphs:
        try:
            style_name = para.style.name if para.style else None
        except Exception:
            style_name = None
        if not style_name or style_name in seen_styles:
            continue
        seen_styles.add(style_name)
        pf = _extract_paragraph_format(para.paragraph_format)
        run_fmt = {}
        if para.runs:
            run_fmt = _extract_run_format(para.runs[0])
        if pf or run_fmt:
            style_formatting[style_name] = {
                "paragraph_format": pf,
                "run_format": run_fmt,
            }
    return style_formatting


def _enrich_style_formatting_from_definitions(doc: Document, style_map: dict, style_formatting: dict) -> dict:
    """Fill in formatting from style definitions; prefer paragraph sample for spacing so output matches template density."""
    result = dict(style_formatting)
    for block_type, style_name in style_map.items():
        if not style_name:
            continue
        from_para = result.get(style_name, {})
        from_def = _format_from_style_definition(doc, style_name)
        if from_def:
            merged = _merge_format(from_def, from_para)
            # Prefer paragraph sample for vertical spacing: if the template's paragraphs had no explicit
            # space_before/space_after/line_spacing, use 0 / single so we don't pull in the style's large values.
            para_fmt_para = (from_para.get("paragraph_format") or {})
            para_fmt_merged = merged.get("paragraph_format") or {}
            for key in ("space_before", "space_after"):
                if key not in para_fmt_para or para_fmt_para.get(key) is None:
                    para_fmt_merged = dict(para_fmt_merged)
                    para_fmt_merged[key] = 0.0
            for key in ("line_spacing", "line_spacing_rule"):
                if key not in para_fmt_para or para_fmt_para.get(key) is None:
                    para_fmt_merged = dict(para_fmt_merged)
                    if key == "line_spacing":
                        para_fmt_merged["line_spacing"] = 1.0
                    if key == "line_spacing_rule":
                        para_fmt_merged["line_spacing_rule"] = "MULTIPLE"
            if para_fmt_merged != (merged.get("paragraph_format") or {}):
                merged = dict(merged)
                merged["paragraph_format"] = para_fmt_merged
            result[style_name] = merged
        elif style_name not in result:
            result[style_name] = {"paragraph_format": {}, "run_format": {}}
    return result


def _apply_default_spacing_and_indent(style_map: dict, style_formatting: dict) -> dict:
    """Ensure paragraph and numbered styles have sensible spacing and indent when missing."""
    result = dict(style_formatting)
    normal_style = style_map.get("paragraph")
    list_style = style_map.get("numbered")
    if normal_style and normal_style in result:
        pf = result[normal_style].get("paragraph_format") or {}
        pf = dict(pf)
        if pf.get("space_after") is None:
            pf["space_after"] = DEFAULT_SPACE_AFTER_PT
        if pf.get("first_line_indent") is None:
            pf["first_line_indent"] = DEFAULT_FIRST_LINE_INDENT_PT
        result[normal_style] = {**result[normal_style], "paragraph_format": pf}
    if list_style and list_style in result:
        pf = result[list_style].get("paragraph_format") or {}
        pf = dict(pf)
        if pf.get("left_indent") is None:
            pf["left_indent"] = DEFAULT_NUMBERED_LEFT_INDENT_PT
        if pf.get("first_line_indent") is None:
            pf["first_line_indent"] = DEFAULT_NUMBERED_FIRST_LINE_INDENT_PT
        if pf.get("space_after") is None:
            pf["space_after"] = DEFAULT_SPACE_AFTER_PT
        result[list_style] = {**result[list_style], "paragraph_format": pf}
    return result


def _format_spec_to_lines(style_name: str, fmt: dict) -> list[str]:
    """Turn one style's paragraph_format + run_format into plain-text bullet lines."""
    lines = [f"Style: {style_name}", ""]
    pf = fmt.get("paragraph_format") or {}
    rf = fmt.get("run_format") or {}
    # Run/font
    run_parts = []
    if rf.get("bold"):
        run_parts.append("bold")
    if rf.get("italic"):
        run_parts.append("italic")
    if rf.get("underline"):
        u = rf["underline"]
        run_parts.append(f"underline ({u})" if isinstance(u, str) else "underline")
    if rf.get("name"):
        run_parts.append(f"font: {rf['name']}")
    if rf.get("size_pt") is not None:
        run_parts.append(f"{rf['size_pt']}pt")
    if run_parts:
        lines.append("- " + ", ".join(run_parts))
    # Paragraph
    para_parts = []
    if pf.get("alignment"):
        para_parts.append(f"align: {pf['alignment'].lower()}")
    if pf.get("space_before") is not None:
        para_parts.append(f"space before: {pf['space_before']}pt")
    if pf.get("space_after") is not None:
        para_parts.append(f"space after: {pf['space_after']}pt")
    if pf.get("left_indent") is not None:
        para_parts.append(f"left indent: {pf['left_indent']}pt")
    if pf.get("first_line_indent") is not None:
        para_parts.append(f"first line indent: {pf['first_line_indent']}pt")
    if pf.get("line_spacing") is not None:
        para_parts.append(f"line spacing: {pf['line_spacing']}pt")
    if para_parts:
        lines.append("- " + ", ".join(para_parts))
    if not run_parts and not para_parts:
        lines.append("- (default paragraph)")
    lines.append("")
    return lines


def _is_line_paragraph(para) -> bool:
    """True if paragraph text looks like a separator line (dots, dashes, underscores, optional X)."""
    if not para or not para.text:
        return False
    t = para.text.strip()
    if not t or len(t) < 3:
        return False
    # Allow letters only for trailing X or similar; rest should be symbols/spaces
    allowed = set(" ._-= \t\u00A0")
    has_symbol = False
    for c in t:
        if c in allowed or c in ".,":
            if c not in " \t":
                has_symbol = True
        elif c.isalpha() and c.upper() == "X" and t.rstrip().endswith("X"):
            continue
        elif not c.isalnum():
            has_symbol = True
        else:
            return False
    return has_symbol


def _extract_line_samples(doc: Document) -> list[dict]:
    """Extract paragraphs that look like separator/signature lines from the template."""
    samples = []
    seen_text = set()
    for para in doc.paragraphs:
        if not _is_line_paragraph(para):
            continue
        text = para.text.strip()
        if not text or text in seen_text:
            continue
        seen_text.add(text)
        alignment = None
        try:
            pf = para.paragraph_format
            if pf and pf.alignment is not None:
                alignment = _enum_name(pf.alignment)
        except Exception:
            pass
        samples.append({"text": text, "alignment": alignment})
    return samples


def build_style_guide(style_map: dict, style_formatting: dict, all_style_names: list = None) -> str:
    """Build plain-text listing of template styles and their formatting (font, bold, alignment, indent, etc.)."""
    lines = [
        "Extracted styles and formatting (from uploaded DOCX template)",
        "",
        "Goal: Apply these styles to raw text so the output has the same structure and formatting as the template. Match how the template uses each style (title, section headings, body, numbered lists, etc.). Use the exact style name as block_type.",
        "",
        "Application rules (match template structure)",
        "",
    ]
    h1 = style_map.get("heading")
    h2 = style_map.get("section_header")
    normal = style_map.get("paragraph")
    list_style = style_map.get("numbered")
    if h1:
        lines.append(f"- Document title / main heading (court name, case caption, or document title like MOTION TO RESTORE) -> use style: {h1}")
    if h2:
        lines.append(f"- Section headers / subheadings (e.g. VERIFIED COMPLAINT, MEMORANDUM, FACTS, ARGUMENT, CONCLUSION, RELIEF REQUESTED, party captions) -> use style: {h2}")
    if normal:
        lines.append(f"- Body paragraphs (narrative, legal argument, facts, verification text, general content) -> use style: {normal}")
    if list_style:
        lines.append(f"- Numbered or bulleted list items (allegations, motion grounds, relief items, any list where the template uses this style) -> use style: {list_style} (numbers 1., 2., 3. added automatically when applicable).")
    names_for_caption = all_style_names or list(style_formatting.keys()) if style_formatting else []
    caption_like = [n for n in names_for_caption if n and ("caption" in n.lower() or "right" in n.lower() or "index" in n.lower())]
    if caption_like:
        lines.append(f"- Case number / date / right-aligned caption -> use style {caption_like[0]} where the template uses it.")
    lines.append("- Addresses (TO:, court or party address lines): use the same style as in the template (often Normal; one block per line if the template breaks them).")
    lines.append("- Signature block (attorney name, firm, address, phone): use the template style for that block; use block_type signature_line for the underline line.")
    lines.append("- Separator lines (dashes/dots ending in X) -> use block_type line and put the line characters in text.")
    lines.append("- Signature underlines -> use block_type signature_line (optional label in text).")
    lines.append("- Page break -> use block_type page_break with empty text before each new major section (as in the template).")
    lines.append("")
    lines.append("Style definitions (font, alignment, indent, etc.)")
    lines.append("")
    names = list(all_style_names) if all_style_names else list(style_formatting.keys()) if style_formatting else []
    if not names and style_map:
        names = list(set(style_map.values()))
    for style_name in names:
        fmt = style_formatting.get(style_name, {})
        lines.extend(_format_spec_to_lines(style_name, fmt))
    if not names:
        lines.append("No styles in template.")
    return "\n".join(lines).strip()


def get_template_content_with_styles(doc: Document, max_paragraphs: int = 120, max_text_per_para: int = 400) -> list[dict]:
    """Extract each paragraph's style name and text so the LLM can see how the template was formatted."""
    out = []
    for i, para in enumerate(doc.paragraphs):
        if i >= max_paragraphs:
            break
        try:
            style_name = para.style.name if para.style else None
        except Exception:
            style_name = None
        text = (para.text or "").strip()
        if max_text_per_para and len(text) > max_text_per_para:
            text = text[: max_text_per_para] + "..."
        out.append({"style": style_name or "Normal", "text": text})
    return out


def extract_styles(doc: Document) -> dict:
    """Extract paragraph style names, style map, per-style formatting, and template content for LLM."""
    para_names = _get_paragraph_style_names(doc)
    if not para_names:
        return {"paragraph_style_names": [], "style_map": {}, "style_formatting": {}}

    heading_like = [
        n for n in para_names
        if any(kw in n.lower() for kw in ("heading", "title", "titre", "section"))
    ]
    list_like = [n for n in para_names if "list" in n.lower() or "number" in n.lower()]

    h1 = _pick_style(para_names, PREFERRED_HEADING_1, heading_like)
    h2 = _pick_style(para_names, PREFERRED_HEADING_2, [n for n in heading_like if n != h1])
    normal = _pick_style(para_names, PREFERRED_NORMAL, para_names)
    list_style = _pick_style(para_names, PREFERRED_LIST, list_like) if list_like else normal

    style_map = {
        "heading": h1 or normal,
        "section_header": h2 or h1 or normal,
        "paragraph": normal,
        "numbered": list_style,
        "wherefore": h2 or h1 or normal,
    }

    style_formatting = _sample_formatting_per_style(doc)
    style_formatting = _enrich_style_formatting_from_definitions(doc, style_map, style_formatting)
    style_formatting = _apply_default_spacing_and_indent(style_map, style_formatting)
    line_samples = _extract_line_samples(doc)
    style_guide = build_style_guide(
        style_map=style_map,
        style_formatting=style_formatting,
        all_style_names=para_names,
    )

    template_content = get_template_content_with_styles(doc)

    return {
        "paragraph_style_names": para_names,
        "style_map": style_map,
        "style_formatting": style_formatting,
        "style_guide": style_guide,
        "line_samples": line_samples,
        "template_content": template_content,
    }


def save_extracted_styles(schema: dict, base_dir: str = None) -> str:
    """Save extracted style schema to JSON and style guide to plain text. Returns path to JSON file."""
    base_dir = base_dir or os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    store_path = os.path.join(base_dir, STORE_DIR)
    os.makedirs(store_path, exist_ok=True)
    filepath = os.path.join(store_path, EXTRACTED_STYLES_FILE)
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(schema, f, indent=2, ensure_ascii=False)
    guide = schema.get("style_guide") or schema.get("style_guide_markdown")
    if guide:
        guide_path = os.path.join(store_path, EXTRACTED_STYLE_GUIDE_FILE)
        with open(guide_path, "w", encoding="utf-8") as f:
            f.write(guide)
    return filepath


def load_extracted_styles(base_dir: str = None) -> dict | None:
    """Load extracted style schema from JSON. Returns None if file missing."""
    base_dir = base_dir or os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    filepath = os.path.join(base_dir, STORE_DIR, EXTRACTED_STYLES_FILE)
    if not os.path.isfile(filepath):
        return None
    with open(filepath, encoding="utf-8") as f:
        return json.load(f)
