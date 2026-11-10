
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

function sanitizeFilename(originalName: string): string {
  return originalName
    .replace(/\.[^.]+$/, "")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
    .trim()
    .replace(/[. ]+$/, "");
}

export function markdownFilename(originalName: string): string {
  const base = originalName
    .replace(/\.(docx?)$/i, "")
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
    .replace(/\.pptx?$/i, "")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
    .trim()
    .replace(/[. ]+$/, "");

  return `${base || "converted-document"}.md`;
}

export async function downloadWordFile(
  blob: Blob,
  originalName: string
): Promise<void> {
  if (!(blob instanceof Blob)) {
    throw new Error("Invalid Word document Blob.");
  }

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

export function downloadConvertedFile(
  blob: Blob,
  originalName: string,
  extension: ".pdf" | ".pptx"
): void {
  // Validate the generated file
  if (!(blob instanceof Blob)) {
    console.error("Invalid download value:", {
      blob,
      type: typeof blob,
      constructor: (blob as unknown as { constructor?: { name?: string } })
        ?.constructor?.name,
    });
    
    throw new Error(
      "Download failed: The generated file is not a valid Blob."
    );
  }

  if (blob.size === 0) {
    throw new Error("Download failed: The generated file is empty.");
  }

  const base = originalName
    .replace(/\.[^.]+$/, "")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
    .trim()
    .replace(/[. ]+$/, "");

  const filename = `${base || "converted-document"}${extension}`;

  let url: string;

  try {
    url = URL.createObjectURL(blob);
  } catch (error) {
    console.error("Blob URL creation failed:", error);
    throw new Error("Unable to create a download URL.");
  }

  try {
    const link = document.createElement("a");

    link.href = url;
    link.download = filename;
    link.style.display = "none";

    document.body.appendChild(link);
    link.click();
    link.remove();
  } finally {
    window.setTimeout(() => {
      URL.revokeObjectURL(url);
    }, 1000);
  }
}
