import { useState } from "react";
import { extractStyles, formatDocument, getDocxFromHtml } from "./api";
import { Editor } from "./components/Editor";

export default function App() {
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [rawText, setRawText] = useState("");
  const [schema, setSchema] = useState<Record<string, unknown> | null>(null);
  const [previewHtml, setPreviewHtml] = useState("<p><br></p>");
  const [editorHtml, setEditorHtml] = useState("<p><br></p>");
  const [docxBase64, setDocxBase64] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const onTemplateChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setTemplateFile(file);
    setSchema(null);
    setError(null);
    setLoading(true);
    extractStyles(file)
      .then((s) => {
        setSchema(s);
        setError(null);
      })
      .catch((err) => setError(err.message))
      .finally(() => setLoading(false));
  };

  const onFormat = () => {
    if (!templateFile || !rawText.trim()) {
      setError("Upload a template and enter text.");
      return;
    }
    setError(null);
    setLoading(true);
    formatDocument(templateFile, rawText)
      .then(({ preview_html, docx_base64 }) => {
        setPreviewHtml(preview_html);
        setEditorHtml(preview_html);
        setDocxBase64(docx_base64);
      })
      .catch((err) => setError(err.message))
      .finally(() => setLoading(false));
  };

  const onDownloadFormatted = () => {
    if (!docxBase64) return;
    setError(null);
    try {
      const bin = atob(docxBase64);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      const blob = new Blob([bytes], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "formatted_output.docx";
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Download failed");
    }
  };

  const onDownloadFromEditor = async () => {
    setError(null);
    try {
      const blob = await getDocxFromHtml(editorHtml);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "formatted_output.docx";
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Download failed");
    }
  };

  const hasFormattedContent = previewHtml !== "<p><br></p>" || editorHtml !== "<p><br></p>";

  return (
    <div className="app">
      <h1 className="app-header">Legal Document Formatter</h1>
      <p className="app-description">
        Upload a DOCX template, paste your text, then format. Edit in the editor and download as DOCX.
      </p>

      <section className="card">
        <span className="section-title">Template</span>
        <label className="label">
          Choose a .docx file
          <input
            type="file"
            accept=".docx"
            onChange={onTemplateChange}
            className="input-file"
          />
        </label>
        <div style={{ marginTop: "0.75rem", display: "flex", alignItems: "center", flexWrap: "wrap", gap: "0.5rem" }}>
          {loading && <span className="status-loading">Extracting styles…</span>}
          {schema && !loading && (
            <span className="status-success">
              ✓ {String((schema.paragraph_style_names as string[])?.length ?? 0)} styles extracted
            </span>
          )}
        </div>
      </section>

      <section className="card">
        <span className="section-title">Raw text to format</span>
        <label className="label">
          Paste your legal text
          <textarea
            value={rawText}
            onChange={(e) => setRawText(e.target.value)}
            placeholder="Paste or type your document content here…"
            className="textarea-input"
            rows={10}
          />
        </label>
        <button
          type="button"
          className="btn btn-primary"
          onClick={onFormat}
          disabled={!templateFile || !rawText.trim() || loading}
        >
          Format with LLM
        </button>
      </section>

      {error && (
        <div className="alert-error" role="alert">
          {error}
        </div>
      )}

      {hasFormattedContent && (
        <section className="card editor-section">
          <h2 className="section-title">Editor</h2>
          <p className="editor-hint">
            Edit below. Use <strong>Download document</strong> for template formatting (recommended), or{" "}
            <strong>Download from editor</strong> to keep your edits with generic styling.
          </p>
          <Editor
            key={previewHtml}
            initialHtml={editorHtml}
            onHtmlChange={setEditorHtml}
            placeholder="Formatted content…"
            editable
          />
          <div className="download-actions">
            <button
              type="button"
              className="btn btn-primary"
              onClick={onDownloadFormatted}
              disabled={!docxBase64}
              title="Uses template styles, alignment, numbering (same as original document)"
            >
              Download document (.docx)
            </button>
            <button
              type="button"
              className="btn btn-secondary"
              onClick={onDownloadFromEditor}
              title="Uses current editor content; layout may differ from template"
            >
              Download from editor
            </button>
          </div>
        </section>
      )}
    </div>
  );
}
