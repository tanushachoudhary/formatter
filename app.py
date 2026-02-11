from dotenv import load_dotenv
load_dotenv()

import os
import re
import streamlit as st
from backend import extract_and_store_styles, process_document
from utils.style_extractor import load_document_blueprint
from utils.html_to_docx import html_to_docx_bytes, normalize_editor_html, plain_text_to_simple_html, simple_html_to_plain_text

# Custom CSS for cleaner UI
st.markdown("""
<style>
  .stApp { max-width: 720px; margin: 0 auto; padding: 2rem 1.5rem; }
  .main-header { font-size: 1.75rem; font-weight: 600; color: #1a365d; margin-bottom: 0.25rem; }
  .main-desc { color: #5c5c5c; font-size: 0.9375rem; margin-bottom: 1.5rem; }
  div[data-testid="stVerticalBlock"] > div:has(> div[data-testid="stFileUploader"]) {
    background: #fff; border: 1px solid #e2dfd9; border-radius: 8px; padding: 1.25rem 1.5rem; margin-bottom: 1rem;
  }
  .stTextArea textarea { border-radius: 6px; border-color: #e2dfd9; }
  .stButton > button {
    background: #1a365d; color: white; border: none; border-radius: 6px;
    padding: 0.625rem 1.25rem; font-weight: 500; transition: background 0.15s;
  }
  .stButton > button:hover { background: #2c5282; color: white; }
  .stDownloadButton > button {
    background: #1a365d; color: white; border: none; border-radius: 6px;
    padding: 0.625rem 1.25rem; font-weight: 500;
  }
  .stDownloadButton > button:hover { background: #2c5282; color: white; }
  .stExpander { border: 1px solid #e2dfd9; border-radius: 8px; margin-bottom: 0.5rem; }
  .stSuccess { padding: 0.5rem 0.75rem; border-radius: 6px; background: #e6ffed; color: #276749; }
  .stError { padding: 0.5rem 0.75rem; border-radius: 6px; }
</style>
""", unsafe_allow_html=True)

try:
    from streamlit_quill import st_quill
    HAS_QUILL = True
except ImportError:
    HAS_QUILL = False

try:
    from streamlit_lexical import streamlit_lexical
    HAS_LEXICAL = True
except ImportError:
    HAS_LEXICAL = False

def _markdown_to_html(md: str) -> str:
    """Convert markdown to HTML for DOCX pipeline (requires markdown package)."""
    if not md or not md.strip():
        return "<p><br></p>"
    try:
        import markdown
        return markdown.markdown(md, extensions=["nl2br"])
    except Exception:
        return plain_text_to_simple_html(md)


def add_space_paragraph(html: str) -> str:
    """Append a blank paragraph for padding/space."""
    if not html or not html.strip():
        return "<p><br></p><p>&nbsp;</p>"
    html = html.rstrip()
    if not html.endswith("</p>"):
        html += "</p>"
    return html + "<p>&nbsp;</p>"


st.markdown('<p class="main-header">Legal Document Formatter</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="main-desc">Upload a DOCX template to extract its styles. When you click Format, each page is sent to the LLM '
    'as a visual reference so your text is structured to match the template. Enter your text and click Format to build the document.</p>',
    unsafe_allow_html=True,
)

template_file = st.file_uploader("Template (.docx)", type=["docx"], label_visibility="collapsed")

if template_file:
    schema = extract_and_store_styles(template_file)
    num_styles = len(schema.get("paragraph_style_names", []))
    num_tables = len(schema.get("tables", []))
    msg = f"✓ {num_styles} styles extracted"
    if num_tables:
        msg += f", {num_tables} table(s)"
    st.success(msg)
    guide = schema.get("style_guide") or schema.get("style_guide_markdown")
    if guide:
        with st.expander("Style guide — applied to input text"):
            st.text(guide)
    if schema.get("style_map"):
        with st.expander("Style mapping (block type → style name)"):
            st.json(schema["style_map"])
    tables = schema.get("tables") or []
    if tables:
        with st.expander("Extracted tables"):
            for t in tables:
                idx = t.get("table_index", 0)
                rows, cols = t.get("rows", 0), t.get("cols", 0)
                st.markdown(f"**Table {idx}**: {rows} × {cols}")
                preview = t.get("cell_preview") or []
                if preview:
                    for ri, row_cells in enumerate(preview):
                        for ci, cell_text in enumerate(row_cells):
                            if cell_text:
                                st.caption(f"  Row {ri}, Col {ci}: {cell_text[:80]}{'…' if len(cell_text) > 80 else ''}")
                    st.markdown("---")
    blueprint = load_document_blueprint()
    if blueprint:
        with st.expander("Document blueprint (formatting metadata)"):
            st.caption("Saved to output/document_blueprint.json")
            st.json(blueprint)

st.markdown("**Raw text to format**")
generated_text = st.text_area("Paste your legal text", height=280, label_visibility="collapsed", placeholder="Paste or type your document content here…")

if generated_text and template_file:
    if st.button("Format with LLM", type="primary"):
        with st.spinner("Calling LLM and building document…"):
            try:
                output_path, preview_text = process_document(generated_text, template_file)
                st.session_state["formatted_output_path"] = output_path
                st.session_state["formatted_editor"] = preview_text
                st.session_state["formatted_editor_html"] = plain_text_to_simple_html(preview_text)
                st.success("Document formatted successfully. Edit below with alignment and formatting, then download.")
            except Exception as e:
                st.error(str(e))

if st.session_state.get("formatted_output_path") or st.session_state.get("formatted_editor_html"):
    st.markdown("---")
    st.subheader("Editor")
    initial_html = st.session_state.get("formatted_editor_html", "<p><br></p>")
    # HTML to use for download: prefer current widget value so edits are always included
    html_for_download = st.session_state.get("formatted_editor_html") or initial_html

    # Single editor: Quill if available, else Lexical, else plain text
    if HAS_QUILL:
        _col, _cap = st.columns([1, 5])
        with _col:
            if st.button("Add space", key="add_space_quill", help="Append a blank paragraph."):
                current = st.session_state.get("formatted_editor_html") or initial_html
                st.session_state["formatted_editor_html"] = add_space_paragraph(current)
                try:
                    st.rerun()
                except AttributeError:
                    st.experimental_rerun()
        with _cap:
            st.caption("Use the toolbar for bold, italic, underline, alignment, lists.")
        editor_content = st_quill(
            value=initial_html,
            html=True,
            key="formatted_editor",
            toolbar=[
                ["bold", "italic", "underline", "strike"],
                [{"align": []}],
                [{"list": "ordered"}, {"list": "bullet"}],
                ["clean"],
            ],
        )
        if editor_content is not None:
            st.session_state["formatted_editor_html"] = editor_content
            html_for_download = editor_content
        else:
            html_for_download = st.session_state.get("formatted_editor_html") or initial_html
        editor_html = html_for_download
    elif HAS_LEXICAL:
        _col, _cap = st.columns([1, 5])
        with _col:
            if st.button("Add space", key="add_space_lexical", help="Append a blank paragraph."):
                current = st.session_state.get("formatted_editor_html") or initial_html
                st.session_state["formatted_editor_html"] = add_space_paragraph(current)
                try:
                    st.rerun()
                except AttributeError:
                    st.experimental_rerun()
        with _cap:
            st.caption("Use **bold**, *italic*, lists. Output is converted to DOCX.")
        initial_value = simple_html_to_plain_text(initial_html)
        md_content = streamlit_lexical(
            value=initial_value,
            placeholder="Edit document (markdown supported)",
            height=400,
            debounce=500,
            key="formatted_editor_lexical",
        )
        if md_content is not None:
            st.session_state["formatted_editor_html"] = _markdown_to_html(md_content)
            html_for_download = _markdown_to_html(md_content)
        else:
            html_for_download = st.session_state.get("formatted_editor_html") or _markdown_to_html(initial_value)
        editor_html = html_for_download
    else:
        _col, _cap = st.columns([1, 5])
        with _col:
            if st.button("Add space", key="add_space_plain", help="Append a blank paragraph."):
                current = st.session_state.get("formatted_editor_html") or initial_html
                st.session_state["formatted_editor_html"] = add_space_paragraph(current)
                try:
                    st.rerun()
                except AttributeError:
                    st.experimental_rerun()
        with _cap:
            st.caption("Install streamlit-quill or streamlit-lexical for rich editing.")
        editor_content = st.text_area(
            "Edit formatted document",
            value=simple_html_to_plain_text(initial_html),
            height=400,
            key="formatted_editor_plain",
        )
        if editor_content is not None:
            st.session_state["formatted_editor_html"] = plain_text_to_simple_html(editor_content)
            html_for_download = plain_text_to_simple_html(editor_content)
        else:
            html_for_download = st.session_state.get("formatted_editor_html") or initial_html
        editor_html = html_for_download

    try:
        # Build DOCX from the editor's current content (html_for_download) so edits are always included
        editor_html_norm = normalize_editor_html(html_for_download)
        docx_bytes = html_to_docx_bytes(editor_html_norm)
        if not docx_bytes or len(docx_bytes) == 0:
            st.warning("Document is empty. Format with LLM first, or add content in the editor above, then download.")
        else:
            clicked = st.download_button(
                "Download document (.docx)",
                data=docx_bytes,
                file_name="formatted_output.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_docx",
                type="primary",
            )
            if clicked:
                st.caption("Download started — check your browser’s downloads folder if the file didn’t open.")
        st.caption("Download includes your edits from the editor above.")
    except Exception as e:
        st.error(f"Could not build DOCX: {e}")
