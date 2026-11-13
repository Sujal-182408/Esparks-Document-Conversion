import { marked } from "marked";
import DOMPurify from "dompurify";

let mermaidRenderCounter = 0;

/**
 * Convert Mermaid fenced code blocks into SVG before Markdown is rendered.
 * This function runs only in the browser because Mermaid needs the DOM.
 */
async function renderMermaidBlocks(markdown: string): Promise<string> {
  const matches = Array.from(
    markdown.matchAll(/```mermaid\s*\n([\s\S]*?)\n```/gi)
  );

  if (matches.length === 0) return markdown;

  const mermaidModule = await import("mermaid");
  const mermaid = mermaidModule.default;

  mermaid.initialize({
    startOnLoad: false,
    securityLevel: "strict",
    theme: "default",
  });

  let renderedMarkdown = markdown;

  for (const match of matches) {
    const fullBlock = match[0];
    const diagramSource = match[1]?.trim();

    if (!diagramSource) continue;

    const id = `pdf-mermaid-${++mermaidRenderCounter}`;
    const { svg } = await mermaid.render(id, diagramSource);
    const replacement = `<div class="pdf-mermaid">${svg}</div>`;

    renderedMarkdown = renderedMarkdown.replace(fullBlock, replacement);
  }

  return renderedMarkdown;
}

function waitForImages(container: HTMLElement): Promise<void> {
  const images = Array.from(container.querySelectorAll("img"));

  return Promise.all(
    images.map(async (image) => {
      if (image.complete && image.naturalWidth > 0) return;

      try {
        await image.decode();
      } catch {
        // A failed external image should not prevent the rest of the document
        // from being converted. html2canvas will omit images it cannot load.
      }
    })
  ).then(() => undefined);
}

export async function convertMarkdownToPdf(
  markdown: string,
  fontSize = 14
): Promise<Blob> {
  if (typeof window === "undefined" || typeof document === "undefined") {
    throw new Error("PDF conversion is available only in the browser.");
  }

  if (!markdown.trim()) {
    throw new Error("The Markdown file is empty.");
  }

  if (new Blob([markdown]).size > 28 * 1024 * 1024) {
    throw new Error("The maximum Markdown file size is 28 MB.");
  }

  if (!Number.isFinite(fontSize) || fontSize < 8 || fontSize > 32) {
    throw new Error("The PDF font size must be between 8 and 32 pt.");
  }

  let container: HTMLDivElement | null = null;

  try {
    const markdownWithDiagrams = await renderMermaidBlocks(markdown);
    const html = await marked.parse(markdownWithDiagrams, {
      breaks: true,
      gfm: true,
    });

    // Permit sanitized SVG so Mermaid diagrams survive sanitization.
    const safeHtml = DOMPurify.sanitize(html, {
      USE_PROFILES: { html: true, svg: true, svgFilters: true },
    });

    container = document.createElement("div");
    container.setAttribute("aria-hidden", "true");
    container.style.cssText = [
      // Keep the capture element inside the viewport. html2pdf clones the
      // source element before html2canvas renders it; placing the source far
      // outside the viewport can make the cloned canvas measure as 0px high.
      "position:absolute",
      "left:0",
      "top:0",
      "z-index:-1",
      "display:block",
      "width:718px",
      "min-height:1px",
      "box-sizing:border-box",
      "padding:28px",
      "background:#ffffff",
      "color:#222222",
      "font-family:Arial,Helvetica,sans-serif",
      `font-size:${fontSize}pt`,
      "line-height:1.55",
      "overflow-wrap:anywhere",
      "pointer-events:none",
    ].join(";");

    container.innerHTML = `
      <style>
        *, *::before, *::after { box-sizing: border-box; }
        article { width: 100%; }
        h1, h2, h3, h4, h5, h6 {
          color: #111827;
          line-height: 1.25;
          margin: 1.25em 0 0.5em;
          break-after: avoid;
          page-break-after: avoid;
        }
        .pdf-heading-section { break-inside: avoid; page-break-inside: avoid; }
        h1 {
          font-size: ${fontSize * (30 / 14)}pt;
          border-bottom: 2px solid #9ca3af;
          padding-bottom: 0.45em;
        }
        h2 { font-size: ${fontSize * (24 / 14)}pt; }
        h3 { font-size: ${fontSize * (20 / 14)}pt; }
        h4, h5, h6 { font-size: ${fontSize * (17 / 14)}pt; }
        p { margin: 0 0 0.9em; }
        ul, ol { margin: 0 0 1em; padding-left: 1.8em; }
        li { margin: 0 0 0.3em; }
        strong { color: #111827; }
        del { color: #6b7280; }
        a { color: #1d4ed8; overflow-wrap: anywhere; }
        hr { border: 0; border-top: 1px solid #9ca3af; margin: 1.4em 0; }
        blockquote {
          border-left: 4px solid #9ca3af;
          background: #f3f4f6;
          margin: 1em 0;
          padding: 0.65em 1em;
          color: #374151;
        }
        pre {
          white-space: pre-wrap;
          overflow-wrap: anywhere;
          word-break: break-word;
          background: #f3f4f6;
          border: 1px solid #d1d5db;
          border-radius: 5px;
          padding: 0.8em;
          margin: 0.8em 0 1em;
          break-inside: auto;
          page-break-inside: auto;
        }
        pre code {
          background: transparent;
          border: 0;
          padding: 0;
          white-space: inherit;
        }
        pre code {
          color: #1f2937;
          background: transparent !important;
          border: 0;
          border-radius: 0;
          padding: 0;
        }
        table {
          display: table;
          width: 100%;
          max-width: 100%;
          border: 1px solid #9ca3af !important;
          border-collapse: separate;
          border-spacing: 0;
          table-layout: auto;
          margin: 1em 0;
          font-size: 0.88em;
          background: #ffffff;
        }
        thead { display: table-header-group; }
        tbody { display: table-row-group; }
        tfoot { display: table-footer-group; }
        tr { display: table-row; break-inside: avoid; page-break-inside: avoid; }
        th, td {
          display: table-cell;
          border: 0 !important;
          border-right: 1px solid #9ca3af !important;
          border-bottom: 1px solid #9ca3af !important;
          padding: 0.55em 0.65em;
          text-align: left;
          vertical-align: top;
          overflow-wrap: break-word;
          word-break: normal;
          line-height: 1.45;
        }
        th:last-child, td:last-child { border-right: 0 !important; }
        tbody tr:last-child td { border-bottom: 0 !important; }
        th {
          background: #e5e7eb !important;
          color: #111827;
          font-weight: 700;
        }
        tbody tr:nth-child(even) td { background: #f8fafc; }
        img, svg { max-width: 100%; height: auto; }
        img { display: block; margin: 0.8em auto; break-inside: avoid; page-break-inside: avoid; }
        .pdf-mermaid {
          width: 100%;
          margin: 1em 0;
          text-align: center;
          break-inside: avoid;
          page-break-inside: avoid;
        }
        .pdf-mermaid svg {
          display: inline-block;
          max-width: 100%;
          height: auto;
        }
        .contains-task-list { list-style: disc; padding-left: 1.8em; }
        li.task-list-item { list-style: disc; }
        .pdf-task-row { display: flex; align-items: baseline; gap: 0.3em; }
        .pdf-task-label { display: block; flex: 1 1 auto; min-width: 0; }
        .pdf-task-checkbox {
          display: inline-block;
          flex: 0 0 auto;
          width: 0.72em;
          height: 0.72em;
          margin: 0;
          border: 1px solid #64748b;
          border-radius: 2px;
          vertical-align: baseline;
          position: relative;
          background: #ffffff;
        }
        .pdf-task-checkbox.is-checked {
          background: #6b7280;
          border-color: #6b7280;
        }
        .pdf-task-checkbox.is-checked::after {
          content: "";
          position: absolute;
          left: 28%;
          top: 12%;
          width: 27%;
          height: 52%;
          border: solid #1f2937;
          border-width: 0 1.5px 1.5px 0;
          transform: rotate(45deg);
        }
        .pdf-highlight {
          position: relative;
          display: inline-block;
          line-height: 1.25;
          padding: 0 0.08em;
        }
        .pdf-highlight__background {
          position: absolute;
          z-index: 0;
          top: 0.22em;
          right: 0;
          bottom: 0.08em;
          left: 0;
          border-radius: 2px;
          background: #fff176;
        }
        .pdf-highlight__text { position: relative; z-index: 1; }
        .pdf-inline-code {
          position: relative;
          display: inline-block;
          font-family: "Courier New", monospace;
          font-size: 0.9em;
          line-height: 1.25;
          color: #1e293b;
          vertical-align: baseline;
          border-radius: 4px;
          padding: 0 0.32em;
          white-space: normal;
          overflow-wrap: anywhere;
        }
        .pdf-inline-code__background {
          position: absolute;
          z-index: 0;
          /* Keep the tint behind the glyph area instead of the line-box top. */
          top: 0.22em;
          right: 0;
          left: 0;
          bottom: 0.08em;
          border-radius: inherit;
          background: #e2e8f0;
        }
        .pdf-inline-code__text {
          position: relative;
          z-index: 1;
        }
      </style>
      <article>${safeHtml}</article>
    `;

    // Convert task-list checkboxes to stable visual squares.
    container
      .querySelectorAll<HTMLInputElement>('input[type="checkbox"]')
      .forEach((input) => {
        const checkbox = document.createElement("span");
        const checked = input.checked || input.hasAttribute("checked");
        checkbox.className = `pdf-task-checkbox${checked ? " is-checked" : ""}`;
        checkbox.setAttribute("aria-hidden", "true");
        input.replaceWith(checkbox);

        const parent = checkbox.parentElement;
        if (!parent) return;

        const row = document.createElement("span");
        row.className = "pdf-task-row";
        const label = document.createElement("span");
        label.className = "pdf-task-label";

        Array.from(parent.childNodes).forEach((node) => {
          if (node === checkbox) {
            row.appendChild(checkbox);
            return;
          }

          if (
            node instanceof HTMLElement &&
            ["UL", "OL", "PRE", "TABLE", "BLOCKQUOTE"].includes(node.tagName)
          ) {
            return;
          }

          label.appendChild(node);
        });

        row.appendChild(label);
        parent.insertBefore(row, parent.firstChild);
      });

    container.querySelectorAll<HTMLElement>("code").forEach((code) => {
      if (code.closest("pre")) return;

      const inlineCode = document.createElement("span");
      inlineCode.className = "pdf-inline-code";
      const background = document.createElement("span");
      background.className = "pdf-inline-code__background";
      background.setAttribute("aria-hidden", "true");
      const text = document.createElement("span");
      text.className = "pdf-inline-code__text";
      text.textContent = code.textContent;
      inlineCode.append(background, text);
      code.replaceWith(inlineCode);
    });

    // Marked's <mark> output is kept but normalized for html2canvas.
    container.querySelectorAll<HTMLElement>("mark").forEach((mark) => {
      const highlight = document.createElement("span");
      highlight.className = "pdf-highlight";
      const background = document.createElement("span");
      background.className = "pdf-highlight__background";
      background.setAttribute("aria-hidden", "true");
      const text = document.createElement("span");
      text.className = "pdf-highlight__text";
      while (mark.firstChild) text.appendChild(mark.firstChild);
      highlight.append(background, text);
      mark.replaceWith(highlight);
    });

    const markdownArticle = container.querySelector("article");
    if (markdownArticle) {
      const blocks = Array.from(markdownArticle.children);
      for (let index = 0; index < blocks.length; index += 1) {
        const heading = blocks[index];
        if (!/^H[1-6]$/.test(heading.tagName)) continue;

        const followingBlock = blocks[index + 1];
        const section = document.createElement("div");
        section.className = "pdf-heading-section";
        markdownArticle.insertBefore(section, heading);
        section.appendChild(heading);

        if (followingBlock && !/^H[1-6]$/.test(followingBlock.tagName)) {
          section.appendChild(followingBlock);
          index += 1;
        }
      }
    }

    document.body.appendChild(container);

    // Force browser layout before html2pdf clones and measures the element.
    // Two animation frames allow styles and asynchronously inserted SVG to settle.
    await new Promise<void>((resolve) => {
      requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
    });

    const containerRect = container.getBoundingClientRect();
    const article = container.querySelector("article");
    const articleRect = article?.getBoundingClientRect();
    const measuredHeight = Math.max(
      containerRect.height,
      container.scrollHeight,
      articleRect?.height ?? 0,
      article?.scrollHeight ?? 0
    );

    if (containerRect.width <= 0 || measuredHeight <= 0) {
      throw new Error(
        `The Markdown content has no measurable layout (container: ${containerRect.width} × ${containerRect.height}px; article: ${articleRect?.width ?? 0} × ${articleRect?.height ?? 0}px).`
      );
    }

    if ("fonts" in document) {
      await document.fonts.ready;
    }
    await waitForImages(container);

    // Render the attached source element directly. html2pdf.js adds another
    // clone/container layer; in this case its internal clone was producing a
    // 0px-high canvas even though the original DOM had a valid measured height.
    // Direct html2canvas rendering avoids that broken intermediate measurement.
    const [{ default: html2canvas }, jspdfModule] = await Promise.all([
      import("html2canvas"),
      import("jspdf"),
    ]);

    const captureWidth = Math.ceil(container.scrollWidth);
    const captureHeight = Math.ceil(container.scrollHeight);

    if (captureWidth <= 0 || captureHeight <= 0) {
      throw new Error(
        `Cannot capture the Markdown layout (${captureWidth} × ${captureHeight}px).`
      );
    }

    const canvas = await html2canvas(container, {
      backgroundColor: "#ffffff",
      scale: 2,
      useCORS: true,
      allowTaint: false,
      logging: false,
      width: captureWidth,
      height: captureHeight,
      windowWidth: Math.max(794, captureWidth),
      windowHeight: Math.max(window.innerHeight, captureHeight),
      scrollX: 0,
      scrollY: 0,
    });

    if (canvas.width <= 0 || canvas.height <= 0) {
      throw new Error(
        `HTML rendering returned an empty canvas (${canvas.width} × ${canvas.height}px).`
      );
    }

    const context = canvas.getContext("2d", { willReadFrequently: true });
    if (!context) {
      throw new Error("The browser could not inspect the rendered document.");
    }

    // Check a small sample for visible content rather than accepting a blank
    // white canvas as a successful conversion.
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
      throw new Error("The browser could not inspect the rendered document.");
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
        "The document rendered blank. Please check the Markdown content and try again."
      );
    }

    const jsPDF = jspdfModule.jsPDF;
    const pdf = new jsPDF({
      orientation: "portrait",
      unit: "mm",
      format: "a4",
      compress: true,
    });

    // A4 with 10mm margins on each side.
    const margin = 10;
    const pageWidthMm = 210;
    const pageHeightMm = 297;
    const contentWidthMm = pageWidthMm - margin * 2;
    const contentHeightMm = pageHeightMm - margin * 2;

    // Slice the tall source canvas into page-sized images. This avoids asking
    // jsPDF to place one enormous image across multiple pages.
    const pageSliceHeight = Math.max(
      1,
      Math.floor((canvas.width * contentHeightMm) / contentWidthMm)
    );

    const layoutScale = canvas.height / captureHeight;
    const containerTop = container.getBoundingClientRect().top;
    const pageFlowNodes = Array.from(
      article?.querySelectorAll<HTMLElement>(
        ".pdf-heading-section, p, li, pre, blockquote, tr, img, .pdf-mermaid"
      ) ?? []
    ).filter((node) => {
      const enclosingBlock = node.parentElement?.closest(
        ".pdf-heading-section, li, pre, blockquote, tr, .pdf-mermaid"
      );
      return !enclosingBlock;
    });
    const pageFlowBounds = pageFlowNodes.map((node) => {
      const rect = node.getBoundingClientRect();
      return {
        top: (rect.top - containerTop) * layoutScale,
        bottom: (rect.bottom - containerTop) * layoutScale,
      };
    });

    let pageIndex = 0;
    let sourceY = 0;

    while (sourceY < canvas.height) {
      const targetY = Math.min(sourceY + pageSliceHeight, canvas.height);
      let nextY = targetY;

      if (targetY < canvas.height) {
        const crossingBlock = pageFlowBounds
          .filter(
            (bounds) =>
              bounds.top > sourceY + 2 &&
              bounds.top < targetY &&
              bounds.bottom > targetY &&
              bounds.bottom - bounds.top < pageSliceHeight
          )
          .sort((first, second) => first.top - second.top)[0];

        if (crossingBlock) {
          nextY = Math.floor(crossingBlock.top);
        }
      }

      nextY = Math.max(sourceY + 1, nextY);
      const sliceHeight = Math.min(nextY - sourceY, canvas.height - sourceY);

      const pageCanvas = document.createElement("canvas");
      pageCanvas.width = canvas.width;
      pageCanvas.height = sliceHeight;

      const pageContext = pageCanvas.getContext("2d");
      if (!pageContext) {
        throw new Error("The browser could not prepare a PDF page.");
      }

      pageContext.fillStyle = "#ffffff";
      pageContext.fillRect(0, 0, pageCanvas.width, pageCanvas.height);
      pageContext.drawImage(
        canvas,
        0,
        sourceY,
        canvas.width,
        sliceHeight,
        0,
        0,
        canvas.width,
        sliceHeight
      );

      const imageData = pageCanvas.toDataURL("image/jpeg", 0.96);
      const imageHeightMm =
        (sliceHeight / canvas.width) * contentWidthMm;

      if (pageIndex > 0) {
        pdf.addPage();
      }

      pdf.addImage(
        imageData,
        "JPEG",
        margin,
        margin,
        contentWidthMm,
        Math.min(contentHeightMm, imageHeightMm),
        undefined,
        "FAST"
      );

      pageCanvas.width = 0;
      pageCanvas.height = 0;
      pageIndex += 1;
      sourceY += sliceHeight;
    }

    const blob = pdf.output("blob");

    if (blob.size < 1000) {
      throw new Error(
        "The PDF output was empty. Please try another Markdown file."
      );
    }

    return blob;
  } catch (error) {
    if (error instanceof Error) {
      throw error;
    }

    throw new Error("Could not generate the PDF. Please try again.");
  } finally {
    container?.remove();
  }
}
