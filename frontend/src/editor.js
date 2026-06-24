// SuperDoc editor entry — single-user load/edit/download (no collaboration).
import { SuperDoc } from "superdoc";
import "superdoc/style.css";

const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const params = new URLSearchParams(location.search);
const docId = params.get("doc");
const userName = params.get("user") || "Guest";
const el = (id) => document.getElementById(id);

if (!docId) {
  el("editor").innerHTML =
    '<div class="msg">Missing <code>?doc=&lt;id&gt;</code>.</div>';
  throw new Error("no doc id");
}

fetch(`/api/documents/${docId}`)
  .then((r) => r.json())
  .then((m) => {
    if (m && m.title) {
      el("docTitle").textContent = m.title;
      document.title = m.title;
    }
  })
  .catch(() => {});

let superdoc;
try {
  superdoc = new SuperDoc({
    selector: "#editor",
    toolbar: "#toolbar",
    documents: [
      {
        id: docId,
        type: DOCX_MIME,
        url: `/api/documents/${docId}/base.docx`,
        name: "proposal.docx",
      },
    ],
    documentMode: "editing",
    pagination: true,
    user: { name: userName },
    onReady: () => {
      el("download").disabled = false;
    },
    onException: ({ error }) => console.error("SuperDoc exception:", error),
  });
  window.superdoc = superdoc;
} catch (e) {
  console.error("Failed to init SuperDoc:", e);
  el("editor").innerHTML =
    '<div class="msg">Editor failed to load. See console.</div>';
}

el("download").addEventListener("click", async () => {
  try {
    let out = await superdoc.exportEditorsToDOCX();
    let b = Array.isArray(out) ? out[0] : out;
    if (b && b.blob) b = b.blob;
    if (b instanceof ArrayBuffer) b = new Blob([b], { type: DOCX_MIME });
    const url = URL.createObjectURL(b);
    const a = document.createElement("a");
    a.href = url;
    a.download = (document.title || "proposal") + ".docx";
    a.click();
    URL.revokeObjectURL(url);
  } catch (e) {
    console.error(e);
  }
});
