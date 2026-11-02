
export function downloadTextFile(
  content: string,
  filename: string,
  mimeType = "text/markdown;charset=utf-8"
): void {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = filename;
  link.style.display = "none";

  document.body.appendChild(link);
  link.click();
  link.remove();

  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function markdownFilename(originalName: string): string {
  const base = originalName
    .replace(/\.docx$/i, "")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
    .trim()
    .replace(/[. ]+$/, "");

  return `${base || "converted-document"}.md`;
}
export function pdfMarkdownFilename(originalName: string): string {
  const base = originalName
    .replace(/\.pdf$/i, "")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
    .trim()
    .replace(/[. ]+$/, "");

  return `${base || "converted-document"}.md`;
}
export function powerpointMarkdownFilename(
  originalName: string
): string {
  const base = originalName
    .replace(/\.pptx$/i, "")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
    .trim()
    .replace(/[. ]+$/, "");

  return `$DIL1.md`;
}

export async function downloadWordFile(
  blob: Blob,
  originalName: string
): Promise<void> {
  const base = originalName
    .replace(/\.(md|markdown|txt)$/i, "")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
    .trim()
    .replace(/[. ]+$/, "");

  const filename = `${base || "converted-document"}.docx`;
  const url = URL.createObjectURL(blob);

  try {
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    link.remove();
  } finally {
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
}
