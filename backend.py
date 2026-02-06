import os

from docx import Document
from docx.shared import Inches

from utils.formatter import clear_document_body, inject_blocks
from utils.llm_formatter import format_text_with_llm
from utils.style_extractor import extract_styles, load_extracted_styles, save_extracted_styles

# Summons-style page margins (generous, like formal legal documents)
DEFAULT_TOP_MARGIN_IN = 1.25
DEFAULT_BOTTOM_MARGIN_IN = 1.25
DEFAULT_LEFT_MARGIN_IN = 1.25
DEFAULT_RIGHT_MARGIN_IN = 1.25



def _project_dir():
    return os.path.dirname(os.path.abspath(__file__))


def get_document_preview_text(docx_path: str) -> str:
    """Build a plain-text preview of the formatted DOCX for display before download."""
    doc = Document(docx_path)
    lines = []
    for para in doc.paragraphs:
        text = (para.text or "").strip()
        lines.append(text if text else "")
    return "\n\n".join(lines).strip()


def extract_and_store_styles(template_file) -> dict:
    """Extract styles from the uploaded DOCX and save to JSON. Returns the style schema."""
    doc = Document(template_file)
    schema = extract_styles(doc)
    save_extracted_styles(schema, base_dir=_project_dir())
    return schema


def process_document(generated_text, template_file):
    """
    Input 1: Uploaded DOCX template (desired styles and formatting).
    Input 2: Raw legal text (unformatted).
    Goal: Extract styles and formatting from the DOCX and apply them to the raw text.

    Steps:
    1. Extract styles from DOCX: paragraph styles (Heading 1, Heading 2, Normal, List Number, etc.),
       font, bold/italic/underline, alignment, indentation, spacing, line samples.
    2. LLM maps each section of the raw text to the appropriate style and outputs styled blocks.
    3. Build a new DOCX with the template's styles applied to the generated blocks.
    """
    project_dir = _project_dir()
    doc = Document(template_file)

    # Extract styles from this template and store
    schema = extract_styles(doc)
    save_extracted_styles(schema, base_dir=project_dir)

    # LLM splits and labels the text into blocks
    blocks = format_text_with_llm(generated_text, schema)

    # Clear template body and fill with LLM output using stored style map and formatting
    clear_document_body(doc)
    # Set generous page margins (summons-style)
    try:
        section = doc.sections[-1]
        section.top_margin = Inches(DEFAULT_TOP_MARGIN_IN)
        section.bottom_margin = Inches(DEFAULT_BOTTOM_MARGIN_IN)
        section.left_margin = Inches(DEFAULT_LEFT_MARGIN_IN)
        section.right_margin = Inches(DEFAULT_RIGHT_MARGIN_IN)
    except Exception:
        pass
    inject_blocks(
        doc,
        blocks,
        style_map=schema["style_map"],
        style_formatting=schema.get("style_formatting", {}),
        line_samples=schema.get("line_samples", []),
    )

    output_path = os.path.join(project_dir, "output", "formatted_output.docx")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    preview_text = get_document_preview_text(output_path)
    return output_path, preview_text
