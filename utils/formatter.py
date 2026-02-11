import re

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER, WD_UNDERLINE
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

# Section underline (thin bottom border) for headings
try:
    from utils.html_to_docx import _paragraph_border_bottom
except Exception:
    _paragraph_border_bottom = None

# Preferred style names to look for in the uploaded document (in order of preference)
PREFERRED_HEADING_1 = ("Heading 1", "Title", "Titre 1")
PREFERRED_HEADING_2 = ("Heading 2", "Subtitle", "Titre 2", "Section")
PREFERRED_NORMAL = ("Normal", "Body Text", "Paragraphe")
PREFERRED_LIST = ("List Number", "List Paragraph", "List")

# Unicode checkbox characters for rendering
CHECKBOX_UNCHECKED = "\u2610"  # ☐
CHECKBOX_CHECKED = "\u2611"    # ☑

# Phrases that indicate address/signature block—do not auto-number these even when using list style
NOT_LIST_CONTENT_PHRASES = (
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


def _looks_like_list_item(text: str) -> bool:
    """True if text looks like a list item (numbered, lettered, or common list starters); False for address/signature."""
    if not text or len(text.strip()) < 3:
        return False
    t = text.strip().lower()
    for phrase in NOT_LIST_CONTENT_PHRASES:
        if phrase in t and not re.match(r"^[\dai]+[\.\)]\s*", t):
            return False
    if re.match(r"^\(\d{3}\)\s*\d{3}-\d{4}", t):
        return False
    # Numbered or lettered: "1. ...", "a. ...", "i. ..."
    if re.match(r"^\d+[\.\)]\s+", t) or re.match(r"^[a-z][\.\)]\s+", t) or re.match(r"^[ivx]+[\.\)]\s+", t):
        return True
    # Common list starters (any document type)
    list_starts = (
        "that ", "first,", "second,", "third,", "plaintiff ", "defendant ", "the court ",
        "movant ", "respondent ", "applicant ", "petitioner ", "1.", "2.", "a.", "b.",
    )
    for start in list_starts:
        if t.startswith(start):
            return True
    return False


# Phrases that start numbered allegations in complaints (so we can prepend 1., 2., 3. when missing)
NUMBERED_ALLEGATION_STARTERS = (
    "that on ",
    "that the ",
    "that as ",
    "that said ",
    "that at ",
    "that plaintiff ",
    "that defendant ",
    "that respondent ",
    "that movant ",
    "pursuant to ",
    "at the time of ",
    "on or about ",
)


def _looks_like_numbered_allegation(text: str) -> bool:
    """True if paragraph is the kind of allegation that should be numbered (e.g. 'That on December...', 'Pursuant to CPLR...')."""
    if not text or len(text.strip()) < 10:
        return False
    t = text.strip().lower()
    if re.match(r"^\d+[\.\)]\s+", t):
        return True
    for start in NUMBERED_ALLEGATION_STARTERS:
        if t.startswith(start):
            return True
    return False


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


# Style names that are body text — always justify, never center; never inherit template italic
BODY_STYLE_NAMES = ("normal", "body text", "list paragraph", "list number", "list")

def _block_type_for_alignment(block_kind: str, section_type: str, style_name: str = "") -> str:
    """Map block_kind + section_type to alignment block_type for enforce_legal_alignment."""
    if block_kind == "line":
        return "line"
    if block_kind == "signature_line":
        return "signature"
    if block_kind == "section_underline":
        return "paragraph"
    # Body styles: always justify (prevent template center/italic from leaking)
    if (style_name or "").strip().lower() in BODY_STYLE_NAMES:
        return "paragraph"
    # Content slots: use section_type
    if section_type in ("caption",):
        return "section_header"
    if section_type in ("attorney_signature", "notary"):
        return "signature"
    if section_type == "to_section":
        return "to_section"
    return "paragraph"


# Legal header phrases that should be centered (court name, document/section titles)
_LEGAL_HEADER_PATTERNS = (
    "SUPREME COURT", "SUPERIOR COURT", "COUNTY OF", "COURT OF",
    "TO THE ABOVE NAMED DEFENDANT", "AS AND FOR A FIRST CAUSE OF ACTION",
    "AS AND FOR A SECOND CAUSE OF ACTION", "NEGLIGENCE", "SUMMONS",
    "VERIFIED COMPLAINT", "JURY TRIAL DEMANDED", "ATTORNEY'S VERIFICATION",
)


def _looks_like_legal_header(text: str) -> bool:
    """True if text is a short legal header that should be centered."""
    if not text or len(text) > 120:
        return False
    t = text.strip().upper()
    return any(p in t for p in _LEGAL_HEADER_PATTERNS)


def _ensure_hanging_indent_if_numbered(paragraph):
    """Apply hanging indent (number left, text indented) when paragraph starts with 'N. '."""
    if not paragraph:
        return
    try:
        text = (paragraph.text or "").strip()
        if not re.match(r"^\d+[\.\)]\s+", text):
            return
        pf = paragraph.paragraph_format
        left_pt = getattr(pf.left_indent, "pt", None) if pf.left_indent is not None else None
        if left_pt is None or left_pt == 0:
            pf.left_indent = Pt(36)
            pf.first_line_indent = Pt(-36)
    except Exception:
        pass


def enforce_legal_alignment(block_type: str, paragraph):
    """Override alignment by block type — renderer controls semantics; template style controls appearance."""
    if not paragraph:
        return
    try:
        # Center court/section headers by content so they match legal format even when LLM returns line/paragraph
        if block_type in ("line", "paragraph") and _looks_like_legal_header(paragraph.text or ""):
            block_type = "section_header"
        if block_type in ("heading", "section_header"):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif block_type in ("paragraph", "numbered", "body"):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        elif block_type in ("signature", "address", "to_section"):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif block_type == "line":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    except Exception:
        pass


def clear_body_italic(paragraph):
    """Remove italic from all runs so body text is not italicised by template style."""
    if not paragraph:
        return
    try:
        for run in paragraph.runs:
            run.italic = False
    except Exception:
        pass


def _apply_paragraph_format(paragraph, fmt: dict):
    """Apply stored paragraph format dict (exact Word features: alignment, spacing, indent, line_spacing, keep_*, page_break_before)."""
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
    for attr in ("page_break_before", "keep_with_next", "keep_together"):
        try:
            if attr in fmt and fmt[attr] is not None:
                setattr(pf, attr, bool(fmt[attr]))
        except Exception:
            pass
    try:
        tab_stops = fmt.get("tab_stops")
        if tab_stops and isinstance(tab_stops, list):
            pf.tab_stops.clear_all()
            for ts in tab_stops:
                pos_pt = ts.get("position_pt") if isinstance(ts, dict) else None
                if pos_pt is None:
                    continue
                align_name = (ts.get("alignment") or "LEFT") if isinstance(ts, dict) else "LEFT"
                leader_name = (ts.get("leader") or "SPACES") if isinstance(ts, dict) else "SPACES"
                align = getattr(WD_TAB_ALIGNMENT, align_name, WD_TAB_ALIGNMENT.LEFT)
                leader = getattr(WD_TAB_LEADER, leader_name, WD_TAB_LEADER.SPACES)
                pf.tab_stops.add_tab_stop(Pt(pos_pt), align, leader)
    except Exception:
        pass


def _render_checkboxes(text: str) -> str:
    """Replace [ ], [x], [X] with Unicode checkbox characters so they render in the document."""
    if not text:
        return text
    text = re.sub(r"\[\s*[xX]\s*\]", CHECKBOX_CHECKED + " ", text)
    text = re.sub(r"\[\s*\]", CHECKBOX_UNCHECKED + " ", text)
    return text


def _is_section_start(
    text: str, block_type: str, style_map: dict, valid_style_names: set,
    section_heading_samples: list = None,
) -> bool:
    """True if this heading should get a page break (template-driven: only when template had page break before this text)."""
    if not text or not text.strip():
        return False
    is_heading = (
        block_type in ("heading", "section_header")
        or (style_map.get("heading") and block_type == style_map["heading"])
        or (style_map.get("section_header") and block_type == style_map["section_header"])
    )
    if not is_heading:
        return False
    if not section_heading_samples:
        return False
    t = text.strip().lower()
    for sample in section_heading_samples:
        if sample in t or t in sample or t.startswith(sample) or sample.startswith(t):
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
    """Apply stored run/font format (bold, italic, underline, font name/size, color) — exact Word document features."""
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
    try:
        hex_color = fmt.get("color_rgb_hex")
        if hex_color and isinstance(hex_color, str) and len(hex_color) >= 6:
            font.color.rgb = RGBColor.from_string(hex_color.strip()[:6])
    except Exception:
        pass


# Fallback when template has no line samples
DEFAULT_SIGNATURE_LINE = "_________________________"
# Default separator line (dashes ending in X) so it always renders
DEFAULT_LINE = "----------------------------------------------------------------------X"


def _is_separator_noise(text: str) -> bool:
    """True if text is only underscores, dashes, equals, spaces, dots, or ends with X (stray separator noise)."""
    if not text or not text.strip():
        return True
    t = text.strip()
    if not t:
        return True
    # Allow trailing X (legal separator style e.g. "------------------------------------------------------------------X")
    if t and t[-1] in ("X", "x"):
        t = t[:-1].strip()
    # Only these chars: space, underscore, hyphen, dot, equals
    allowed = set(" _-.=\u00A0\t")
    if all(c in allowed for c in t):
        return True
    if re.match(r"^[\s\-\._=]+$", t):
        return True
    return False

# Phrases that start the main body (after caption); caption = everything before this
BODY_START_PHRASES = (
    "please take notice",
    "take further notice",
    "dated:",
    "affirms the following",
    "under the penalties of perjury",
    "being duly sworn",
    "duly sworn, says",
)
# Right-column caption: index number and motion/document title
RIGHT_CAPTION_PHRASES = (
    "index no",
    "index number",
    "notice of motion",
    "to restore",
    "affirmation in support",
    "affidavit of service",
    "memorandum of law",
)
# Block text that starts a new document (repeated caption) when not at start of paste
NEW_DOCUMENT_START_PHRASES = (
    "supreme court of the state of new york",
    "supreme court of new york",
)


def _split_into_document_segments(blocks: list) -> list[list]:
    """Split blocks into one list per document. New document starts at repeated court heading (not at index 0)."""
    if not blocks:
        return []
    segment_starts = [0]
    for i in range(1, len(blocks)):
        bt, text = blocks[i]
        t = (text or "").strip().lower()
        if not t:
            continue
        for phrase in NEW_DOCUMENT_START_PHRASES:
            if t.startswith(phrase) or t == phrase.strip() or phrase in t[:80]:
                segment_starts.append(i)
                break
    out = []
    for j in range(len(segment_starts)):
        start = segment_starts[j]
        end = segment_starts[j + 1] if j + 1 < len(segment_starts) else len(blocks)
        out.append(blocks[start:end])
    return out


def _split_caption_body(blocks: list) -> tuple[list, list, list]:
    """Split blocks into caption_left, caption_right, body. Returns (caption_left, caption_right, body_blocks)."""
    if not blocks:
        return [], [], []
    body_start_idx = None
    for i, (bt, text) in enumerate(blocks):
        t = (text or "").strip().lower()
        if not t:
            continue
        for phrase in BODY_START_PHRASES:
            if phrase in t or t.startswith(phrase):
                body_start_idx = i
                break
        if body_start_idx is not None:
            break
    if body_start_idx is None:
        return [], [], blocks
    caption_blocks = blocks[:body_start_idx]
    body_blocks = blocks[body_start_idx:]
    left, right = [], []
    for b in caption_blocks:
        bt, text = b
        t = (text or "").strip().lower()
        is_right = any(p in t for p in RIGHT_CAPTION_PHRASES) or (t == "to restore")
        if is_right:
            right.append(b)
        else:
            left.append(b)
    return left, right, body_blocks


def _resolve_style(block_type: str, style_map: dict, style_formatting: dict):
    """Resolve block_type to a style name: use template style name if present, else logical style_map."""
    if block_type in style_formatting:
        return block_type
    return style_map.get(block_type, style_map.get("paragraph"))


def inject_blocks(doc, blocks, style_map=None, style_formatting=None, line_samples=None, section_heading_samples=None, template_structure=None):
    """Inject text into template structure. When template_structure is provided (slot-fill):
    assign paragraph.style = template style only — no manual formatting. Word handles layout,
    numbering, spacing from style definitions. Renderer never invents formatting."""
    if style_map is None:
        style_map = _build_style_map_from_doc(doc)
    if not style_map:
        style_map = {"heading": None, "section_header": None, "paragraph": None, "numbered": None, "wherefore": None}
    style_formatting = style_formatting or {}
    line_samples = line_samples or []
    section_heading_samples = section_heading_samples or []
    valid_style_names = set(style_formatting.keys())

    # Structure-driven slot-fill: parser only — assign existing text to slots; never invent or fallback.
    if template_structure and len(blocks) == len(template_structure):
        seen = set()  # Caption deduplication: do not render the same block text twice (stops repeated court headers)
        for i in range(len(template_structure)):
            spec = template_structure[i]
            style = (blocks[i][0] if isinstance(blocks[i], (list, tuple)) else spec.get("style", "Normal"))
            slot_text = (blocks[i][1] if isinstance(blocks[i], (list, tuple)) and len(blocks[i]) > 1 else (blocks[i] if isinstance(blocks[i], str) else ""))
            slot_text = (slot_text or "").strip()
            if style not in valid_style_names:
                style = style_map.get("paragraph") or (list(valid_style_names)[0] if valid_style_names else "Normal")
            block_kind = spec.get("block_kind", "paragraph")
            section_type = spec.get("section_type", "body")
            template_text = (spec.get("template_text") or "").strip()
            if spec.get("page_break_before"):
                doc.add_page_break()
            if block_kind == "line":
                if not (slot_text or template_text).strip():
                    continue
                if template_text:
                    p = doc.add_paragraph(template_text, style=style)
                    fmt = (style_formatting.get(style) or {}).get("paragraph_format") or {}
                    _apply_paragraph_format(p, fmt)
                    enforce_legal_alignment(_block_type_for_alignment(block_kind, section_type, style), p)
                continue
            if block_kind == "section_underline":
                p = doc.add_paragraph(style=style)
                if _paragraph_border_bottom:
                    _paragraph_border_bottom(p, pt=0.5)
                fmt = (style_formatting.get(style) or {}).get("paragraph_format") or {}
                _apply_paragraph_format(p, fmt)
                enforce_legal_alignment(_block_type_for_alignment(block_kind, section_type, style), p)
                continue
            if block_kind == "signature_line":
                if template_text:
                    p = doc.add_paragraph(template_text, style=style)
                    fmt = (style_formatting.get(style) or {}).get("paragraph_format") or {}
                    _apply_paragraph_format(p, fmt)
                    enforce_legal_alignment(_block_type_for_alignment(block_kind, section_type, style), p)
                continue
            # Content slots: empty → skip; dedupe then render
            if not slot_text:
                continue
            if slot_text in seen:
                continue
            seen.add(slot_text)
            p = doc.add_paragraph(style=style)
            p.add_run(_render_checkboxes(slot_text))
            fmt = (style_formatting.get(style) or {}).get("paragraph_format") or {}
            _apply_paragraph_format(p, fmt)
            _ensure_hanging_indent_if_numbered(p)
            align_type = _block_type_for_alignment(block_kind, section_type, style)
            if align_type in ("paragraph", "numbered", "body"):
                try:
                    sa = p.paragraph_format.space_after
                    if sa is None or (getattr(sa, "pt", 0) or 0) == 0:
                        p.paragraph_format.space_after = Pt(6)
                except Exception:
                    pass
            enforce_legal_alignment(align_type, p)
            if align_type == "paragraph":
                clear_body_italic(p)
        trim_trailing_separators(doc)
        return

    # Fallback path when no template_structure: still use style only; no fake numbering (Word handles via style).
    segments = _split_into_document_segments(blocks)
    for seg_idx, segment in enumerate(segments):
        if seg_idx > 0:
            doc.add_page_break()
        caption_left, caption_right, body_blocks = _split_caption_body(segment)
        blocks_to_render = caption_left + caption_right + body_blocks if (caption_left or caption_right) else segment

        allegation_run_counter = 0
        for block_type, text in blocks_to_render:
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
                fmt = (style_formatting.get(style) or {}).get("paragraph_format") or {}
                _apply_paragraph_format(p, fmt)
                enforce_legal_alignment("signature", p)
                continue

            if block_type == "section_underline":
                style = _resolve_style("paragraph", style_map, style_formatting)
                p = doc.add_paragraph(style=style)
                if _paragraph_border_bottom:
                    _paragraph_border_bottom(p, pt=0.5)
                fmt = (style_formatting.get(style) or {}).get("paragraph_format") or {}
                _apply_paragraph_format(p, fmt)
                enforce_legal_alignment("paragraph", p)
                continue

            if block_type == "line":
                line_text = (text or "").strip()
                if line_text and ("block_type" in line_text or "text field" in line_text):
                    line_text = ""
                if not line_text and line_samples:
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
                fmt = (style_formatting.get(style) or {}).get("paragraph_format") or {}
                _apply_paragraph_format(p, fmt)
                enforce_legal_alignment("line", p)
                continue

            if not text:
                continue

            style = block_type if block_type in valid_style_names else style_map.get(block_type, style_map.get("paragraph"))
            list_style_name = style_map.get("numbered")
            if block_type in ("heading", "section_header") or style in (style_map.get("heading"), style_map.get("section_header")):
                allegation_run_counter = 0
            if doc.paragraphs and _is_section_start(text, block_type, style_map, valid_style_names, section_heading_samples):
                doc.add_page_break()
            # Only strip leading "1.", "a.", "i." when using the template's list style (Word will add numbers).
            # When using Normal/paragraph style, keep the literal number so points still show as 1., 2., etc.
            if _looks_like_list_item(text) and style == list_style_name and list_style_name:
                text = re.sub(r"^\d+[\.\)]\s*", "", text).strip()
                text = re.sub(r"^[a-z][\.\)]\s*", "", text, count=1).strip()
                text = re.sub(r"^[ivx]+[\.\)]\s*", "", text, count=1, flags=re.IGNORECASE).strip()
            # If this looks like a numbered allegation but has no number (LLM didn't add one), prepend 1., 2., 3....
            elif _looks_like_numbered_allegation(text) and style != list_style_name and not re.match(r"^\d+[\.\)]\s+", text.strip()):
                allegation_run_counter += 1
                text = f"{allegation_run_counter}. " + text
            elif not _looks_like_numbered_allegation(text):
                allegation_run_counter = 0
            text = _render_checkboxes(text)
            segments = [(text, False, False)]
            _add_paragraph_with_inline_formatting(doc, segments, style, {})
            p = doc.paragraphs[-1] if doc.paragraphs else None
            if p:
                fmt = (style_formatting.get(style) or {}).get("paragraph_format") or {}
                _apply_paragraph_format(p, fmt)
                _ensure_hanging_indent_if_numbered(p)
                # Ensure spacing between body/numbered paragraphs when template has none
                if style not in (style_map.get("heading"), style_map.get("section_header")):
                    try:
                        sa = p.paragraph_format.space_after
                        if sa is None or (getattr(sa, "pt", 0) or 0) == 0:
                            p.paragraph_format.space_after = Pt(6)
                    except Exception:
                        pass
                align_type = "section_header" if style in (style_map.get("heading"), style_map.get("section_header")) else "paragraph"
                enforce_legal_alignment(align_type, p)
                if align_type == "paragraph":
                    clear_body_italic(p)
    trim_trailing_separators(doc)


def _is_empty_or_noise_paragraph(para) -> bool:
    """True if paragraph has no content or only separator noise (whitespace, underscores, '- - -')."""
    if not para:
        return True
    text = (para.text or "").strip()
    if not text:
        return True
    return _is_separator_noise(text)


def trim_trailing_separators(doc):
    """Remove trailing paragraphs that look like separators (----, ====, ______). Call after rendering, before save."""
    def is_separator(text):
        t = (text or "").strip()
        return t.startswith("-") or t.startswith("=") or t.startswith("_")

    while doc.paragraphs:
        last = doc.paragraphs[-1]
        if is_separator(last.text):
            try:
                p = last._element
                p.getparent().remove(p)
            except Exception:
                break
        else:
            break


def remove_trailing_empty_and_noise(doc):
    """Remove trailing paragraphs that are empty or only separator noise (underscores, '- - -')."""
    paras = list(doc.paragraphs)
    if not paras:
        return
    removed = 0
    for para in reversed(paras):
        if _is_empty_or_noise_paragraph(para):
            try:
                p_el = para._element
                p_el.getparent().remove(p_el)
                removed += 1
            except Exception:
                break
        else:
            break


def force_single_column(doc):
    """Force all sections to single-column layout so the document renders as one column per page, not multi-column.
    Handles sectPr as direct children of body and sectPr inside paragraph properties (section breaks)."""
    try:
        body = doc.element.body
        for sect_pr in body.iter(qn("w:sectPr")):
            cols = None
            for c in sect_pr:
                if c.tag == qn("w:cols"):
                    cols = c
                    break
            if cols is not None:
                cols.set(qn("w:num"), "1")
            else:
                cols_el = OxmlElement("w:cols")
                cols_el.set(qn("w:num"), "1")
                sect_pr.insert(0, cols_el)
    except Exception:
        pass


def clear_document_body(doc):
    """Remove all paragraphs and tables from the document body, keeping section properties."""
    for para in list(doc.paragraphs):
        p_el = para._element
        p_el.getparent().remove(p_el)
    for table in list(doc.tables):
        tbl_el = table._element
        tbl_el.getparent().remove(tbl_el)
