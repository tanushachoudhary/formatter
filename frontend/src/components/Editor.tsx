import { useEditor, EditorContent, BubbleMenu } from "@tiptap/react";
import StarterKit from "@tiptap/starter-kit";
import Underline from "@tiptap/extension-underline";
import TextAlign from "@tiptap/extension-text-align";
import TextStyle from "@tiptap/extension-text-style";
import FontFamily from "@tiptap/extension-font-family";
import { useEffect } from "react";

const FONT_OPTIONS = [
  { label: "Font", value: "" },
  { label: "Times New Roman", value: "Times New Roman" },
  { label: "Arial", value: "Arial" },
  { label: "Georgia", value: "Georgia" },
  { label: "Calibri", value: "Calibri" },
  { label: "Verdana", value: "Verdana" },
  { label: "Garamond", value: "Garamond, serif" },
  { label: "Cambria", value: "Cambria, serif" },
];

const extensions = [
  StarterKit.configure({ heading: { levels: [1, 2, 3] } }),
  Underline,
  TextAlign.configure({ types: ["heading", "paragraph"] }),
  TextStyle,
  FontFamily.configure({ types: ["textStyle"] }),
];

interface EditorProps {
  initialHtml: string;
  onHtmlChange: (html: string) => void;
  placeholder?: string;
  editable?: boolean;
}

export function Editor({
  initialHtml,
  onHtmlChange,
  placeholder = "Paste or type…",
  editable = true,
}: EditorProps) {
  const editor = useEditor({
    extensions,
    content: initialHtml,
    editable,
    editorProps: {
      attributes: {
        "data-placeholder": placeholder,
      },
    },
    onUpdate: ({ editor }) => {
      onHtmlChange(editor.getHTML());
    },
  });

  useEffect(() => {
    if (!editor) return;
    const current = editor.getHTML();
    const empty = "<p></p>" === current || "<p><br></p>" === current;
    if (initialHtml && (empty || !current.trim())) {
      editor.commands.setContent(initialHtml, false);
    }
  }, [initialHtml, editor]);

  if (!editor) return null;

  const addSpace = () => {
    editor.chain().focus().insertContent("<p>&nbsp;</p>").run();
    onHtmlChange(editor.getHTML());
  };

  return (
    <div className="editor-wrapper">
      {editable && (
        <div className="tiptap-toolbar" style={{ borderRadius: "6px 6px 0 0", marginBottom: 0 }}>
          <select
            title="Font"
            value={editor.getAttributes("textStyle").fontFamily ?? ""}
            onChange={(e) => {
              const v = e.target.value;
              if (v) editor.chain().focus().setFontFamily(v).run();
              else editor.chain().focus().unsetFontFamily().run();
              onHtmlChange(editor.getHTML());
            }}
            style={{ padding: "6px 8px", borderRadius: 4, border: "1px solid #ccc", minWidth: 140 }}
          >
            {FONT_OPTIONS.map((opt) => (
              <option key={opt.value || "default"} value={opt.value}>
                {opt.label}
              </option>
            ))}
          </select>
          <button type="button" onClick={addSpace} title="Insert blank paragraph">
            Add space
          </button>
          <span className="toolbar-sep" aria-hidden />
          <button
            type="button"
            onClick={() => editor.chain().focus().toggleBold().run()}
            className={editor.isActive("bold") ? "is-active" : ""}
          >
            Bold
          </button>
          <button
            type="button"
            onClick={() => editor.chain().focus().toggleItalic().run()}
            className={editor.isActive("italic") ? "is-active" : ""}
          >
            Italic
          </button>
          <button
            type="button"
            onClick={() => editor.chain().focus().toggleUnderline().run()}
            className={editor.isActive("underline") ? "is-active" : ""}
          >
            Underline
          </button>
          <span className="toolbar-sep" aria-hidden />
          <button
            type="button"
            onClick={() => editor.chain().focus().setTextAlign("left").run()}
            className={editor.isActive({ textAlign: "left" }) ? "is-active" : ""}
          >
            Left
          </button>
          <button
            type="button"
            onClick={() => editor.chain().focus().setTextAlign("center").run()}
            className={editor.isActive({ textAlign: "center" }) ? "is-active" : ""}
          >
            Center
          </button>
          <button
            type="button"
            onClick={() => editor.chain().focus().setTextAlign("justify").run()}
            className={editor.isActive({ textAlign: "justify" }) ? "is-active" : ""}
          >
            Justify
          </button>
        </div>
      )}
      {editable && (
        <BubbleMenu editor={editor} tippyOptions={{ duration: 100 }}>
          <div className="tiptap-toolbar">
            <button
              type="button"
              onClick={() => editor.chain().focus().toggleBold().run()}
              className={editor.isActive("bold") ? "is-active" : ""}
            >
              Bold
            </button>
            <button
              type="button"
              onClick={() => editor.chain().focus().toggleItalic().run()}
              className={editor.isActive("italic") ? "is-active" : ""}
            >
              Italic
            </button>
            <button
              type="button"
              onClick={() => editor.chain().focus().toggleUnderline().run()}
              className={editor.isActive("underline") ? "is-active" : ""}
            >
              Underline
            </button>
            <button
              type="button"
              onClick={() => editor.chain().focus().setTextAlign("left").run()}
              className={editor.isActive({ textAlign: "left" }) ? "is-active" : ""}
            >
              Left
            </button>
            <button
              type="button"
              onClick={() => editor.chain().focus().setTextAlign("center").run()}
              className={editor.isActive({ textAlign: "center" }) ? "is-active" : ""}
            >
              Center
            </button>
            <button
              type="button"
              onClick={() => editor.chain().focus().setTextAlign("justify").run()}
              className={editor.isActive({ textAlign: "justify" }) ? "is-active" : ""}
            >
              Justify
            </button>
          </div>
        </BubbleMenu>
      )}
      <EditorContent editor={editor} />
    </div>
  );
}
