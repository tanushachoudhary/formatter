import re

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_UNDERLINE
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt

# Preferred style names to look for in the uploaded document (in order of preference)
PREFERRED_HEADING_1 = ("Heading 1", "Title", "Titre 1")
PREFERRED_HEADING_2 = ("Heading 2", "Subtitle", "Titre 2", "Section")
PREFERRED_NORMAL = ("Normal", "Body Text", "Paragraphe")
PREFERRED_LIST = ("List Number", "List Paragraph", "List")

# Unicode checkbox characters for rendering
CHECKBOX_UNCHECKED = "\u2610"  # ☐
CHECKBOX_CHECKED = "\u2611"    # ☑

# Prefixes that indicate a negligence/cause-of-action allegation (only these get auto-numbered when using list style)
NEGLIGENCE_ALLEGATION_PREFIXES = (
    "that ",
    "at all times ",
    "by reason of ",
    "at the time ",
    "on or about ",
)

# Phrases that indicate non-allegation content (address, signature block)—do not number even if list style
NOT_ALLEGATION_PHRASES = (
    "attorneys for plaintiff",
    "attorneys for defendant",
    "park avenue",
    "floor",
    "new york, new york",
    "street",
    "avenue",
    "road",
    "drive",
    "court",
    "pllc",
    "esq.",
    "tel.",
    "fax",
)


def _looks_like_negligence_allegation(text: str) -> bool:
    """True only if text looks like a cause-of-action allegation (That..., At all times...), not address/signature."""
    if not text or len(text.strip()) < 10:
        return False
    t = text.strip().lower()
    for phrase in NOT_ALLEGATION_PHRASES:
        if phrase in t and not t.startswith(("that ", "at all times ", "by reason of ")):
            return False
    if re.match(r"^\(\d{3}\)\s*\d{3}-\d{4}", t):
        return False
    for prefix in NEGLIGENCE_ALLEGATION_PREFIXES:
        if t.startswith(prefix):
            return True
    if re.match(r"^\d+\.\s*", t) and any(t.startswith(p) or p in t[:80] for p in NEGLIGENCE_ALLEGATION_PREFIXES):
        return True
    return False


def _looks_like_list_item(text: str) -> bool:
    """True if text looks like a numbered/bulleted list item (motion grounds, allegations, relief items, etc.), not address/signature."""
    if not text or len(text.strip()) < 3:
        return False
    t = text.strip().lower()
    for phrase in NOT_ALLEGATION_PHRASES:
        if phrase in t and not re.match(r"^[\dai]+[\.\)]\s*", t):
            return False
    if re.match(r"^\(\d{3}\)\s*\d{3}-\d{4}", t):
        return False
    # Numbered: "1. ...", "a. ...", "i. ..."
    if re.match(r"^\d+[\.\)]\s+", t) or re.match(r"^[a-z][\.\)]\s+", t) or re.match(r"^[ivx]+[\.\)]\s+", t):
        return True
    # Common list starters (legal)
    list_starts = ("that ", "at all times ", "by reason of ", "first,", "second,", "plaintiff moves", "defendant moves", "the court should", "the court must")
    for start in list_starts:
        if t.startswith(start):
            return True
    return False


# Phrases that typically start a new section (page break before these when they appear as headings)
SECTION_START_PHRASES = (
    "summons and verified complaint",
    "verified complaint",
    "attorney's verification",
    "verification",
    "notice of entry",
    "notice of settlement",
    "please take notice",
    "certification",
    "wherefore",
    "as and for a second cause",
    "as and for a first cause of action",
    # Motions, memoranda, and other document types
    "motion to restore",
    "motion to dismiss",
    "memorandum of law",
    "memorandum in support",
    "memorandum in opposition",
    "affirmation in support",
    "affirmation in opposition",
    "facts",
    "argument",
    "conclusion",
    "relief requested",
    "preliminary statement",
    "statement of facts",
    "legal standard",
    "application",
)


def _get_paragraph_style_names(doc):
    """Return list of paragraph style names defined in the document."""
    return [
        s.name for s in doc.styles
        if s.type == WD_STYLE_TYPE.PARAGRAPH
    ]


def _pick_style(doc, preferred_names, fallback_names=None):
    """Return the first preferred style name that exists in the document, else first fallback."""
    available = _get_paragraph_style_names(doc)
    for name in preferred_names:
        if name in available:
            return name
    if fallback_names:
        for name in fallback_names:
            if name in available:
                return name
    return available[0] if available else None


def _build_style_map_from_doc(doc):
    """Build a block_type -> style_name map using only styles present in the document."""
    para_names = _get_paragraph_style_names(doc)
    if not para_names:
        return None

    # Heading-like: names containing "heading", "title", "titre" (case-insensitive), in stable order
    heading_like = [n for n in para_names if any(kw in n.lower() for kw in ("heading", "title", "titre", "section"))]
    # List-like
    list_like = [n for n in para_names if "list" in n.lower() or "number" in n.lower()]

    h1 = _pick_style(doc, PREFERRED_HEADING_1, heading_like)
    h2 = _pick_style(doc, PREFERRED_HEADING_2, [n for n in heading_like if n != h1])
    normal = _pick_style(doc, PREFERRED_NORMAL, para_names)
    list_style = _pick_style(doc, PREFERRED_LIST, list_like) if list_like else normal

    return {
        "heading": h1 or normal,
        "section_header": h2 or h1 or normal,
        "paragraph": normal,
        "numbered": list_style,
        "wherefore": h2 or h1 or normal,
    }


def _apply_paragraph_format(paragraph, fmt: dict):
    """Apply stored paragraph format dict to a paragraph (alignment, spacing, indent)."""
    if not fmt or not paragraph:
        return
    pf = paragraph.paragraph_format
    try:
        if "alignment" in fmt and fmt["alignment"]:
            alignment = getattr(WD_ALIGN_PARAGRAPH, fmt["alignment"], None)
            if alignment is not None:
                pf.alignment = alignment
    except Exception:
        pass
    for attr, key in (
        ("space_before", "space_before"),
        ("space_after", "space_after"),
        ("left_indent", "left_indent"),
        ("right_indent", "right_indent"),
        ("first_line_indent", "first_line_indent"),
    ):
        try:
            val = fmt.get(key)
            if val is not None and isinstance(val, (int, float)):
                setattr(pf, attr, Pt(val))
        except Exception:
            pass
    try:
        if "line_spacing" in fmt and fmt["line_spacing"] is not None:
            val = fmt["line_spacing"]
            rule_name = fmt.get("line_spacing_rule")
            rule = getattr(WD_LINE_SPACING, rule_name, None) if isinstance(rule_name, str) else None
            # EXACTLY or AT_LEAST: use fixed height in points
            if rule in (WD_LINE_SPACING.EXACTLY, WD_LINE_SPACING.AT_LEAST):
                pf.line_spacing = Pt(val) if isinstance(val, (int, float)) else val
                pf.line_spacing_rule = rule
            # MULTIPLE, SINGLE, DOUBLE, ONE_POINT_FIVE: use multiplier (float)
            else:
                num = float(val) if isinstance(val, (int, float)) else None
                if num is not None and 0.25 <= num <= 3.0:
                    pf.line_spacing = num
    except Exception:
        pass


def _render_checkboxes(text: str) -> str:
    """Replace [ ], [x], [X] with Unicode checkbox characters so they render in the document."""
    if not text:
        return text
    text = re.sub(r"\[\s*[xX]\s*\]", CHECKBOX_CHECKED + " ", text)
    text = re.sub(r"\[\s*\]", CHECKBOX_UNCHECKED + " ", text)
    return text


def _is_section_start(text: str, block_type: str, style_map: dict, valid_style_names: set) -> bool:
    """True if this block looks like the start of a new document section (e.g. new heading that starts a page)."""
    if not text or not text.strip():
        return False
    t = text.strip().lower()
    # Must look like a heading: either logical type or template heading/section style
    is_heading = (
        block_type in ("heading", "section_header")
        or (style_map.get("heading") and block_type == style_map["heading"])
        or (style_map.get("section_header") and block_type == style_map["section_header"])
    )
    if not is_heading:
        return False
    for phrase in SECTION_START_PHRASES:
        if phrase in t or t.startswith(phrase):
            return True
    return False


def _add_paragraph_with_inline_formatting(doc, segments: list[tuple[str, bool, bool]], style, run_fmt_base: dict):
    """Add a paragraph with multiple runs for bold/italic segments. style and run_fmt_base from template."""
    p = doc.add_paragraph(style=style)
    for seg_text, bold, italic in segments:
        if not seg_text:
            continue
        run = p.add_run(seg_text)
        fmt = dict(run_fmt_base)
        if bold:
            fmt["bold"] = True
        if italic:
            fmt["italic"] = True
        _apply_run_format(run, fmt)
    return p


def _apply_run_format(run, fmt: dict):
    """Apply stored run/font format (bold, italic, underline, font name/size)."""
    if not fmt or not run:
        return
    font = run.font
    try:
        if "bold" in fmt:
            font.bold = fmt["bold"]
    except Exception:
        pass
    try:
        if "italic" in fmt:
            font.italic = fmt["italic"]
    except Exception:
        pass
    try:
        if "underline" in fmt:
            u = fmt["underline"]
            if u is True or u == "True":
                font.underline = True
            elif u is False or u == "False":
                font.underline = False
            elif isinstance(u, str) and hasattr(WD_UNDERLINE, u):
                font.underline = getattr(WD_UNDERLINE, u)
            else:
                font.underline = u
    except Exception:
        pass
    try:
        if "name" in fmt and fmt["name"]:
            font.name = fmt["name"]
    except Exception:
        pass
    try:
        if "size_pt" in fmt and fmt["size_pt"] is not None:
            font.size = Pt(fmt["size_pt"])
    except Exception:
        pass


# Fallback when template has no line samples
DEFAULT_SIGNATURE_LINE = "_________________________"
# Default separator line (dashes ending in X) so it always renders
DEFAULT_LINE = "----------------------------------------------------------------------X"


def _resolve_style(block_type: str, style_map: dict, style_formatting: dict):
    """Resolve block_type to a style name: use template style name if present, else logical style_map."""
    if block_type in style_formatting:
        return block_type
    return style_map.get(block_type, style_map.get("paragraph"))


def inject_blocks(doc, blocks, style_map=None, style_formatting=None, line_samples=None):
    """Add paragraphs using only template-driven styles and line samples. No hardcoded layout."""
    if style_map is None:
        style_map = _build_style_map_from_doc(doc)
    if not style_map:
        style_map = {"heading": None, "section_header": None, "paragraph": None, "numbered": None, "wherefore": None}
    style_formatting = style_formatting or {}
    line_samples = line_samples or []
    valid_style_names = set(style_formatting.keys())

    numbered_counter = 0
    for block_type, text in blocks:
        text = (text or "").strip()

        if block_type == "page_break":
            doc.add_page_break()
            continue

        if block_type == "signature_line":
            label = (text.strip() if text and text.strip() and text.strip() not in ("---", "—", "-") else None)
            line_text = None
            if line_samples:
                for s in line_samples:
                    t = s.get("text", "")
                    if "_" in t and t.strip().replace("_", "").replace(" ", "") == "":
                        line_text = t
                        break
                if line_text is None:
                    line_text = line_samples[0].get("text", DEFAULT_SIGNATURE_LINE)
            if not line_text:
                line_text = DEFAULT_SIGNATURE_LINE
            if label:
                line_text = f"{line_text}  {label}"
            style = _resolve_style("paragraph", style_map, style_formatting)
            p = doc.add_paragraph(line_text, style=style)
            if style and style in style_formatting:
                _apply_paragraph_format(p, style_formatting[style].get("paragraph_format", {}))
            continue

        if block_type == "line":
            line_text = (text or "").strip()
            # If LLM put instruction-like prose in text, ignore it and use template or default
            if line_text and ("block_type" in line_text or "text field" in line_text):
                line_text = ""
            if not line_text and line_samples:
                # Prefer a sample ending in X (legal separator), else first sample
                for s in line_samples:
                    t = s.get("text", "")
                    if t.rstrip().endswith("X") or t.rstrip().endswith("x"):
                        line_text = t
                        break
                if not line_text:
                    line_text = line_samples[0].get("text", DEFAULT_LINE)
            if not line_text:
                line_text = DEFAULT_LINE
            style = _resolve_style("paragraph", style_map, style_formatting)
            p = doc.add_paragraph(line_text, style=style)
            sample = line_samples[0] if line_samples else {}
            if sample.get("alignment"):
                try:
                    p.alignment = getattr(WD_ALIGN_PARAGRAPH, sample["alignment"], None)
                except Exception:
                    pass
            if style and style in style_formatting:
                _apply_paragraph_format(p, style_formatting[style].get("paragraph_format", {}))
            continue

        if not text:
            continue

        # Use block_type as style name if it exists in template; else map logical type via style_map
        style = block_type if block_type in valid_style_names else style_map.get(block_type, style_map.get("paragraph"))
        # Insert page break before new sections (e.g. VERIFIED COMPLAINT, ATTORNEY'S VERIFICATION, NOTICE OF ENTRY)
        if doc.paragraphs and _is_section_start(text, block_type, style_map, valid_style_names):
            doc.add_page_break()
        # Reset numbering after a section/heading so each cause of action can start at 1.
        if block_type in ("heading", "section_header") or (style_map.get("heading") and block_type == style_map["heading"]) or (style_map.get("section_header") and block_type == style_map["section_header"]):
            numbered_counter = 0
        # Auto-number list-style blocks that look like list items (motion grounds, allegations, relief items); skip address/signature
        list_style_name = style_map.get("numbered")
        is_list_style_block = block_type == "numbered" or (list_style_name and block_type == list_style_name)
        if is_list_style_block and _looks_like_list_item(text):
            numbered_counter += 1
            text = re.sub(r"^\d+[\.\)]\s*", "", text).strip()
            text = re.sub(r"^[a-z][\.\)]\s*", "", text, count=1).strip()
            text = re.sub(r"^[ivx]+[\.\)]\s*", "", text, count=1, flags=re.IGNORECASE).strip()
            text = f"{numbered_counter}. {text}"

        text = _render_checkboxes(text)
        # Plain text only (no markdown ** or * parsing)
        segments = [(text, False, False)]
        run_fmt = style_formatting.get(style, {}).get("run_format", {}) if style else {}
        _add_paragraph_with_inline_formatting(doc, segments, style, run_fmt)
        para = doc.paragraphs[-1]
        if style and style in style_formatting:
            _apply_paragraph_format(para, style_formatting[style].get("paragraph_format", {}))


def clear_document_body(doc):
    """Remove all paragraphs and tables from the document body, keeping section properties."""
    for para in list(doc.paragraphs):
        p_el = para._element
        p_el.getparent().remove(p_el)
    for table in list(doc.tables):
        tbl_el = table._element
        tbl_el.getparent().remove(tbl_el)
