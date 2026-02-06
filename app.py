from dotenv import load_dotenv
load_dotenv()

import streamlit as st
from backend import extract_and_store_styles, process_document

st.title("Legal Document Formatter")
st.write(
    "Upload a DOCX template to extract its styles. Then enter your text and click Format to have an LLM "
    "apply those styles (no merging with template content)."
)

template_file = st.file_uploader("Upload the DOCX template", type=["docx"])

if template_file:
    schema = extract_and_store_styles(template_file)
    st.success(f"Styles extracted and stored: {len(schema.get('paragraph_style_names', []))} paragraph styles.")
    guide = schema.get("style_guide") or schema.get("style_guide_markdown")
    if guide:
        with st.expander("Style guide — applied to input text"):
            st.text(guide)
    if schema.get("style_map"):
        with st.expander("Style mapping (block type → style name)"):
            st.json(schema["style_map"])

generated_text = st.text_area("Enter the generated text", height=300)

if generated_text and template_file:
    if st.button("Format with LLM"):
        with st.spinner("Calling LLM and building document…"):
            try:
                output_path, preview_text = process_document(generated_text, template_file)
                st.success("Document formatted successfully. Preview below, then download when ready.")

                with st.expander("Preview formatted document", expanded=True):
                    st.text(preview_text)

                with open(output_path, "rb") as f:
                    data = f.read()
                st.download_button(
                    "Download Formatted Document",
                    data=data,
                    file_name="formatted_output.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            except Exception as e:
                st.error(str(e))
