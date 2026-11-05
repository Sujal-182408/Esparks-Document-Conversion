import * as pdfjsLib from "pdfjs-dist";
pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  "pdfjs-dist/build/pdf.worker.min.mjs",
  import.meta.url
).toString();

type ExtractedLine = {
  text: string;
  fontSize: number;
  y: number;
  endX: number;
};

function escapeMarkdown(text: string): string {
  return text.replace(/([\\`*_])/g, "\\$1");
}

function formatPdfLine(line: ExtractedLine, bodyFontSize: number): string {
  const text = line.text.trim();
  if (!text) return "";

  const bullet = text.match(/^[•●▪◦‣]\s*(.+)$/);
  if (bullet) return `- ${bullet[1]}`;

  const numberedItem = text.match(/^(\d+)[.)]\s+(.+)$/);
  if (numberedItem) return `${numberedItem[1]}. ${numberedItem[2]}`;

  if (line.fontSize >= bodyFontSize * 1.2) return `## ${text}`;

  return text;
}

export async function pdfToMarkdown(file: File): Promise<string> {
  if (!file.name.toLowerCase().endsWith(".pdf"))
    throw new Error("Please select a PDF file.");
  if (file.size === 0)
    throw new Error("The selected file is empty.");
  if (file.size > 20 * 1024 * 1024)
    throw new Error("The maximum file size is 20 MB.");
  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  // Check the PDF file signature.
  const signature = new TextDecoder().decode(bytes.slice(0,5));
  if(signature !== "%PDF-")
    throw new Error("this file does not appear to be a valid PDF.");
  let loadingTask:
  | ReturnType<typeof pdfjsLib.getDocument>
  | undefined;
  try {
     loadingTask = pdfjsLib.getDocument({
      data: new Uint8Array(buffer),
    });

    const pdf = await loadingTask.promise;

    const pages: string[] = [];

    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
      const page = await pdf.getPage(pageNumber);
      const content = await page.getTextContent();
      const lines: ExtractedLine[] = [];

      for (const item of content.items) {
        if (!("str" in item)) continue;

        const text = escapeMarkdown(item.str);
        if (!text) continue;

        const y = item.transform[5];
        const x = item.transform[4];
        const fontSize = Math.hypot(item.transform[0], item.transform[1]);
        const currentLine = lines[lines.length - 1];

        if (!currentLine || Math.abs(y - currentLine.y) > 3) {
          lines.push({ text, fontSize, y, endX: x + item.width });
        } else {
          const gap = x - currentLine.endX;
          if (
            gap > fontSize * 0.2 &&
            currentLine.text &&
            !/\s$/.test(currentLine.text) &&
            !/^\s/.test(text)
          ) {
            currentLine.text += " ";
          }
          currentLine.text += text;
          currentLine.fontSize = Math.max(currentLine.fontSize, fontSize);
          currentLine.endX = x + item.width;
        }
      }

      const fontSizes = lines
        .map((line) => line.fontSize)
        .filter((fontSize) => fontSize > 0)
        .sort((left, right) => left - right);
      const bodyFontSize =
        fontSizes[Math.floor(fontSizes.length / 2)] ?? 12;
      const titleIndex =
        pageNumber === 1 && lines.length > 0
          ? lines.reduce(
              (largestIndex, line, index, allLines) =>
                line.fontSize > allLines[largestIndex].fontSize
                  ? index
                  : largestIndex,
              0
            )
          : -1;
      const hasProminentTitle =
        titleIndex >= 0 &&
        lines[titleIndex].fontSize >= bodyFontSize * 1.15;
      const pageText = lines
        .map((line, index) =>
          hasProminentTitle && index === titleIndex
            ? `# ${line.text.trim()}`
            : formatPdfLine(line, bodyFontSize)
        )
        .filter(Boolean)
        .join("\n\n")
        .trim();
      if (pageText) {
        pages.push(
          pdf.numPages > 1
            ? `<!-- Page ${pageNumber} -->\n\n${pageText}`
            : pageText
        );
      }
    }

    const markdown = pages.join("\n\n---\n\n").trim();

    if (!markdown) {
      throw new Error(
        "No selectable text was found. This may be a scanned PDF; OCR is not supported yet."
      );
    }

    return markdown + "\n";
  } catch (error) {
    console.error("PDF to Markdown conversion failed:", error);
  
    if (
      error instanceof Error &&
      (
        error.message.startsWith("No selectable text") ||
        error.message.startsWith("This file")
      )
    ) {
      throw error;
    }
  
    throw new Error(
      "PDF conversion failed. Check the browser console for details."
    );
  } finally {
    if (loadingTask) {
      await loadingTask.destroy();
    }
  }
}