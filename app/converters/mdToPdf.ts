
import { marked } from "marked";
import DOMPurify from "dompurify";

type Html2PdfWorker = {
  set: (options: Record<string, unknown>) => Html2PdfWorker;
  from: (element: HTMLElement) => Html2PdfWorker;
  toCanvas: () => Html2PdfWorker;
  toPdf: () => Html2PdfWorker;
  get: {
    (key: "canvas"): Promise<HTMLCanvasElement>;
    (key: "pdf"): Promise<{
      output: (type: "blob") => Blob;
    }>;
  };
};

type Html2PdfFactory = () => Html2PdfWorker;

export async function convertMarkdownToPdf(
  markdown: string,
  fontSize = 14
): Promise<Blob> {
  if (!markdown.trim()) {
    throw new Error("The Markdown file is empty.");
  }

  if (new Blob([markdown]).size > 28 * 1024 * 1024) {
    throw new Error("The maximum Markdown file size is 28 MB.");
  }

  const html = await marked.parse(markdown, { breaks: true });
  const safeHtml = DOMPurify.sanitize(html);

  const container = document.createElement("div");
  container.style.cssText = [
    "width:100%",
    "box-sizing:border-box",
    "padding:32px",
    "background:#ffffff",
    "color:#222222",
    "font-family:Arial,Helvetica,sans-serif",
    `font-size:${fontSize}pt`,
    "line-height:1.65",
    "overflow-wrap:break-word",
    "pointer-events:none",
  ].join(";");

  container.innerHTML = `
    <style>
      * { box-sizing: border-box; }
      article { width: 100%; }
      h1, h2, h3, h4, h5, h6 {
        color: #000000;
        line-height: 1.25;
        margin: 1.4em 0 0.55em;
        page-break-after: avoid;
      }
      h1 { font-size: ${fontSize * (30 / 14)}pt; border-bottom: 2px solid #888888; padding-bottom: 0.55em; }
      h2 { font-size: ${fontSize * (24 / 14)}pt; }
      h3 { font-size: ${fontSize * (20 / 14)}pt; }
      h4, h5, h6 { font-size: ${fontSize * (17 / 14)}pt; }
      p, ul, ol { margin: 0 0 1em; }
      ul, ol { padding-left: 1.85em; }
      li { margin: 0 0 0.35em; page-break-inside: avoid; break-inside: avoid; }
      p, blockquote, pre, img, tr {
        page-break-inside: avoid;
        break-inside: avoid;
      }
      strong { color: #000000; }
      del { color: #555555; }
      code {
        font-family: "Courier New", monospace;
        background: #eeeeee;
        border-radius: 3px;
        padding: 0.07em 0.28em;
      }
      pre {
        white-space: pre-wrap;
        overflow-wrap: anywhere;
        background: #eeeeee;
        border: 1px solid #888888;
        border-radius: 5px;
        padding: 0.85em;
        page-break-inside: avoid;
      }
      pre code { background: transparent; padding: 0; }
      blockquote {
        border-left: 4px solid #888888;
        background: #f5f5f5;
        margin: 1.15em 0;
        padding: 0.7em 1.15em;
        color: #333333;
      }
      table { border-collapse: collapse; width: 100%; margin: 1.15em 0; }
      th, td {
        border: 1px solid #888888;
        padding: 0.5em 0.65em;
        text-align: left;
        vertical-align: top;
        overflow-wrap: anywhere;
      }
      th { background: #eeeeee; color: #000000; }
      a { color: #000000; overflow-wrap: anywhere; }
      img { display: block; max-width: 100%; height: auto; margin: 0.85em auto; }
      hr { border: 0; border-top: 1px solid #888888; margin: 1.6em 0; }
    </style>
    <article>${safeHtml}</article>
  `;

  try {
    // Load the browser-only library when the conversion is requested.
    const html2pdfModule = await import("html2pdf.js");
    const html2pdf = (
      html2pdfModule as unknown as { default: Html2PdfFactory }
    ).default;

    const worker = html2pdf().set({
      margin: [10, 10, 10, 10],
      filename: "converted-document.pdf",
      image: { type: "jpeg", quality: 0.95 },
      html2canvas: {
        scale: 2,
        useCORS: true,
        allowTaint: false,
        logging: false,
        backgroundColor: "#ffffff",
        windowWidth: 794,
      },
      jsPDF: {
        unit: "mm",
        format: "a4",
        orientation: "portrait",
      },
      pagebreak: {
        mode: ["css", "legacy"],
      },
    });

    await worker.from(container).toCanvas();

    const canvas = await worker.get("canvas");
    if (canvas.width === 0 || canvas.height === 0) {
      throw new Error(
        `The PDF page could not be rendered (${canvas.width} × ${canvas.height}).`
      );
    }

    const context = canvas.getContext("2d", { willReadFrequently: true });
    if (!context) {
      throw new Error("The browser could not inspect the rendered PDF page.");
    }

    const sampleCanvas = document.createElement("canvas");
    sampleCanvas.width = 64;
    sampleCanvas.height = Math.min(
      256,
      Math.max(64, Math.ceil((canvas.height / canvas.width) * 64))
    );
    const sampleContext = sampleCanvas.getContext("2d", {
      willReadFrequently: true,
    });
    if (!sampleContext) {
      throw new Error("The browser could not inspect the rendered PDF page.");
    }
    sampleContext.drawImage(
      canvas,
      0,
      0,
      sampleCanvas.width,
      sampleCanvas.height
    );
    const pixels = sampleContext.getImageData(
      0,
      0,
      sampleCanvas.width,
      sampleCanvas.height
    ).data;
    let visiblePixels = 0;
    for (let index = 0; index < pixels.length; index += 4) {
      if (
        pixels[index] < 245 ||
        pixels[index + 1] < 245 ||
        pixels[index + 2] < 245
      ) {
        visiblePixels += 1;
        if (visiblePixels >= 5) break;
      }
    }

    if (visiblePixels < 5) {
      throw new Error(
        "The PDF page rendered blank. Please check the Markdown content and try again."
      );
    }

    await worker.toPdf();

    const pdf = await worker.get("pdf");
    const blob = pdf.output("blob");

    if (blob.size < 1000) {
      throw new Error("The PDF output was empty. Please try another Markdown file.");
    }

    return blob;
  } catch (error) {
    if (
      error instanceof Error &&
      (error.message.includes("PDF") || error.message.includes("28 MB"))
    ) {
      throw error;
    }

    throw new Error(
      "Could not generate the PDF. Please check the Markdown content and try again."
    );
  } finally {
    container.remove();
  }
}
