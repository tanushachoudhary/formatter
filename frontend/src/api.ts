const API_BASE = "/api";

export async function extractStyles(file: File): Promise<Record<string, unknown>> {
  const form = new FormData();
  form.append("file", file);
  const r = await fetch(`${API_BASE}/extract-styles`, {
    method: "POST",
    body: form,
  });
  if (!r.ok) {
    const err = await r.json().catch(() => ({ detail: r.statusText }));
    throw new Error((err as { detail?: string }).detail || "Extract failed");
  }
  return r.json();
}

export async function formatDocument(file: File, text: string): Promise<{
  preview_text: string;
  preview_html: string;
  docx_base64: string;
}> {
  const form = new FormData();
  form.append("file", file);
  form.append("text", text);
  const r = await fetch(`${API_BASE}/format`, {
    method: "POST",
    body: form,
  });
  if (!r.ok) {
    const err = await r.json().catch(() => ({ detail: r.statusText }));
    throw new Error((err as { detail?: string }).detail || "Format failed");
  }
  return r.json();
}

export async function getDocxFromHtml(html: string): Promise<Blob> {
  const r = await fetch(`${API_BASE}/docx-from-html`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ html }),
  });
  if (!r.ok) {
    const err = await r.json().catch(() => ({ detail: r.statusText }));
    throw new Error((err as { detail?: string }).detail || "DOCX build failed");
  }
  return r.blob();
}
