
import { marked, type Token, type Tokens } from "marked";
import type PptxGenJS from "pptxgenjs";

export type MdToPptxOptions = {
  fontSize?: number;
  title?: string;
  author?: string;
};

type TextRun = {
  text: string;
  options?: Record<string, any>;
};

type PptxSlide = PptxGenJS.Slide;

type HtmlTableCell = {
  runs: TextRun[];
  text: string;
  colspan: number;
  rowspan: number;
  header: boolean;
  align: "left" | "center" | "right";
};

const SW = 13.333;
const MARGIN_X = 0.65;
const CONTENT_W = SW - MARGIN_X * 2;
const CONTENT_TOP = 1.35;
const CONTENT_BOTTOM = 6.98;
const FOOTER_Y = 7.17;
const GAP = 0.12;

const FONT = "Aptos";
const MONO = "Consolas";

const MAX_MARKDOWN_SIZE = 28 * 1024 * 1024;
const MAX_IMAGE_SIZE = 8 * 1024 * 1024;
const MAX_IMAGE_DIMENSION = 3000;

const COLORS = {
  primary: "087E8B",
  accent: "16B8C4",
  body: "263238",
  muted: "718096",
  border: "DCE5E9",
  background: "FFFFFF",
  codeBackground: "F1F5F9",
  quoteBackground: "F0FAFB",
  tableAlternate: "F5FAFB",
};

const FONT_SIZE = {
  title: 38,
  h1: 30,
  h2: 25,
  h3: 22,
  h4: 20,
  h5: 18,
  h6: 16,
  body: 18,
  code: 10,
  small: 9,
};

const LINE_HEIGHT = 1.25;

function cleanText(value: unknown): string {
  return String(value ?? "")
    .replace(/\u0000/g, "")
    .replace(/\u00a0/g, " ")
    .replace(/\u200b/g, "")
    .replace(/\uFEFF/g, "")
    .replace(/\r\n?/g, "\n")
    .replace(/\t/g, "    ");
}

function mergeRuns(runs: TextRun[]): TextRun[] {
  const result: TextRun[] = [];

  for (const run of runs) {
    if (!run.text) continue;

    const previous = result[result.length - 1];

    if (
      previous &&
      JSON.stringify(previous.options ?? {}) ===
        JSON.stringify(run.options ?? {})
    ) {
      previous.text += run.text;
    } else {
      result.push({
        text: run.text,
        options: { ...run.options },
      });
    }
  }

  return result;
}

function safeUrl(value: string): string | undefined {
  try {
    const url = new URL(
      value,
      typeof window !== "undefined"
        ? window.location.href
        : "https://localhost",
    );

    if (["http:", "https:", "mailto:"].includes(url.protocol)) {
      return url.href;
    }
  } catch {
    return undefined;
  }

  return undefined;
}

function getPlainText(tokens: Token[] = []): string {
  return tokens
    .map((token) => {
      if (token.type === "text") {
        const item = token as Tokens.Text;

        if (item.tokens?.length) {
          return getPlainText(item.tokens as Token[]);
        }

        return item.text;
      }

      if (token.type === "codespan") {
        return (token as Tokens.Codespan).text;
      }

      if (token.type === "image") {
        return (token as Tokens.Image).text;
      }

      if (token.type === "html") {
        return (token as Tokens.HTML).text.replace(/<[^>]*>/g, "");
      }

      if ("tokens" in token && Array.isArray(token.tokens)) {
        return getPlainText(token.tokens as Token[]);
      }

      if ("text" in token && typeof token.text === "string") {
        return token.text;
      }

      return "";
    })
    .join("");
}

function inlineRuns(
  tokens: Token[] = [],
  inherited: Record<string, any> = {},
): TextRun[] {
  const output: TextRun[] = [];

  const add = (
    text: string,
    options: Record<string, any> = inherited,
  ) => {
    if (text) {
      output.push({
        text,
        options: { ...options },
      });
    }
  };

  for (const token of tokens) {
    switch (token.type) {
      case "escape":
        add((token as Tokens.Escape).text);
        break;

      case "text": {
        const item = token as Tokens.Text;

        if (item.tokens?.length) {
          output.push(
            ...inlineRuns(item.tokens as Token[], inherited),
          );
        } else {
          add(item.text);
        }

        break;
      }

      case "strong":
        output.push(
          ...inlineRuns(token.tokens as Token[], {
            ...inherited,
            bold: true,
          }),
        );
        break;

      case "em":
        output.push(
          ...inlineRuns(token.tokens as Token[], {
            ...inherited,
            italic: true,
          }),
        );
        break;

      case "del":
        output.push(
          ...inlineRuns(token.tokens as Token[], {
            ...inherited,
            strike: true,
          }),
        );
        break;

      case "codespan":
        add((token as Tokens.Codespan).text, {
          ...inherited,
          fontFace: MONO,
          color: "B42318",
          fill: { color: "F1F5F9" },
        });
        break;

      case "link": {
        const link = token as Tokens.Link;
        const href = safeUrl(link.href);

        output.push(
          ...inlineRuns(link.tokens as Token[], {
            ...inherited,
            color: COLORS.primary,
            underline: { color: COLORS.primary },
            ...(href ? { hyperlink: { url: href } } : {}),
          }),
        );

        break;
      }

      case "image": {
        const image = token as Tokens.Image;
        add(image.text || "[Image]");
        break;
      }

      case "br":
        add("\n");
        break;

      case "html": {
        const value = (token as Tokens.HTML).text;

        add(
          value
            .replace(/<br\s*\/?>/gi, "\n")
            .replace(/<[^>]*>/g, ""),
        );

        break;
      }

      default: {
        if ("tokens" in token && Array.isArray(token.tokens)) {
          output.push(
            ...inlineRuns(token.tokens as Token[], inherited),
          );
        } else if ("text" in token && typeof token.text === "string") {
          add(token.text);
        }
      }
    }
  }

  return mergeRuns(output);
}

function measureText(
  text: string,
  fontSize: number,
  bold = false,
  fontFace = FONT,
): number {
  if (typeof document === "undefined") {
    return (text.length * fontSize * 0.55) / 72;
  }

  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");

  if (!context) {
    return (text.length * fontSize * 0.55) / 72;
  }

  context.font = `${bold ? "700" : "400"} ${fontSize}px "${fontFace}"`;

  return context.measureText(text).width / 96;
}

function wrapPlainText(
  text: string,
  maxWidth: number,
  fontSize: number,
  fontFace = FONT,
): string[] {
  const output: string[] = [];

  for (const paragraph of cleanText(text).split("\n")) {
    if (!paragraph) {
      output.push("");
      continue;
    }

    let current = "";

    for (const word of paragraph.split(/(\s+)/)) {
      if (
        !current ||
        measureText(current + word, fontSize, false, fontFace) <= maxWidth
      ) {
        current += word;
        continue;
      }

      output.push(current.trimEnd());
      current = word.trimStart();

      while (
        current &&
        measureText(current, fontSize, false, fontFace) > maxWidth
      ) {
        let cut = 0;

        for (let i = 1; i <= current.length; i++) {
          if (
            measureText(current.slice(0, i), fontSize, false, fontFace) >
            maxWidth
          ) {
            break;
          }

          cut = i;
        }

        if (!cut) cut = 1;

        output.push(current.slice(0, cut));
        current = current.slice(cut);
      }
    }

    if (current.trim()) output.push(current.trimEnd());
  }

  return output.length ? output : [""];
}

function wrapRuns(
  runs: TextRun[],
  maxWidth: number,
  defaultSize: number,
): TextRun[][] {
  const lines: TextRun[][] = [[]];
  let currentWidth = 0;

  const push = (text: string, options: Record<string, any>) => {
    if (!text) return;

    const currentLine = lines[lines.length - 1];
    const previous = currentLine[currentLine.length - 1];

    if (
      previous &&
      JSON.stringify(previous.options ?? {}) === JSON.stringify(options)
    ) {
      previous.text += text;
    } else {
      currentLine.push({
        text,
        options: { ...options },
      });
    }
  };

  for (const run of runs) {
    const style = run.options ?? {};
    const fontSize = Number(style.fontSize ?? defaultSize);
    const fontFace = String(style.fontFace ?? FONT);
    const bold = Boolean(style.bold);
    const parts = cleanText(run.text).split(/(\n|\s+)/);

    for (const part of parts) {
      if (!part) continue;

      if (part === "\n") {
        lines.push([]);
        currentWidth = 0;
        continue;
      }

      let remaining = part;

      while (remaining) {
        const isWhitespace = /^\s+$/.test(remaining);

        if (
          currentWidth > 0 &&
          currentWidth +
            measureText(remaining, fontSize, bold, fontFace) >
            maxWidth &&
          !isWhitespace
        ) {
          lines.push([]);
          currentWidth = 0;
          remaining = remaining.trimStart();

          if (!remaining) break;
        }

        const available = maxWidth - currentWidth;

        if (
          measureText(remaining, fontSize, bold, fontFace) <= available
        ) {
          push(remaining, style);
          currentWidth += measureText(
            remaining,
            fontSize,
            bold,
            fontFace,
          );
          break;
        }

        let cut = 0;

        for (let i = 1; i <= remaining.length; i++) {
          if (
            measureText(
              remaining.slice(0, i),
              fontSize,
              bold,
              fontFace,
            ) > available
          ) {
            break;
          }

          cut = i;
        }

        if (!cut) {
          if (currentWidth > 0) {
            lines.push([]);
            currentWidth = 0;
            continue;
          }

          cut = 1;
        }

        const piece = remaining.slice(0, cut);

        push(piece, style);

        currentWidth += measureText(
          piece,
          fontSize,
          bold,
          fontFace,
        );

        remaining = remaining.slice(cut);

        if (remaining) {
          lines.push([]);
          currentWidth = 0;
        }
      }
    }
  }

  return lines.length ? lines : [[]];
}

function decodeSvgDataUrl(src: string): string {
  const comma = src.indexOf(",");

  if (comma < 0) {
    throw new Error("Invalid SVG data URL.");
  }

  const metadata = src.slice(0, comma);
  const payload = src.slice(comma + 1);

  try {
    if (/;base64/i.test(metadata)) {
      const binary = atob(payload);
      const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));

      return new TextDecoder().decode(bytes);
    }

    return decodeURIComponent(payload);
  } catch {
    throw new Error("Unable to decode SVG data.");
  }
}

async function svgToPng(svg: string): Promise<string> {
  const parser = new DOMParser();
  const parsed = parser.parseFromString(svg, "image/svg+xml");

  if (parsed.querySelector("parsererror")) {
    throw new Error("Invalid SVG markup.");
  }

  const svgElement = parsed.documentElement;

  if (svgElement.tagName.toLowerCase() !== "svg") {
    throw new Error("SVG root element is missing.");
  }

  parsed.querySelectorAll("script, foreignObject").forEach((element) => {
    element.remove();
  });

  parsed.querySelectorAll("*").forEach((element) => {
    for (const attribute of Array.from(element.attributes)) {
      if (/^on/i.test(attribute.name)) {
        element.removeAttribute(attribute.name);
      }

      if (
        ["href", "xlink:href"].includes(attribute.name.toLowerCase()) &&
        /^\s*javascript:/i.test(attribute.value)
      ) {
        element.removeAttribute(attribute.name);
      }
    }
  });

  if (!svgElement.getAttribute("xmlns")) {
    svgElement.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  }

  let width = Number.parseFloat(svgElement.getAttribute("width") ?? "");
  let height = Number.parseFloat(svgElement.getAttribute("height") ?? "");

  const viewBox = svgElement.getAttribute("viewBox");

  if ((!width || !height) && viewBox) {
    const values = viewBox.split(/[\s,]+/).map(Number);

    if (values.length === 4) {
      width ||= values[2];
      height ||= values[3];
    }
  }

  width = Math.max(1, Math.min(width || 1200, MAX_IMAGE_DIMENSION));
  height = Math.max(1, Math.min(height || 800, MAX_IMAGE_DIMENSION));

  svgElement.setAttribute("width", String(width));
  svgElement.setAttribute("height", String(height));

  const svgString = new XMLSerializer().serializeToString(parsed);
  const blob = new Blob([svgString], {
    type: "image/svg+xml;charset=utf-8",
  });

  const objectUrl = URL.createObjectURL(blob);

  try {
    const image = new Image();

    await new Promise<void>((resolve, reject) => {
      image.onload = () => resolve();
      image.onerror = () => reject(new Error("Unable to render SVG."));
      image.src = objectUrl;
    });

    const canvas = document.createElement("canvas");

    const scale = Math.min(
      2,
      MAX_IMAGE_DIMENSION / Math.max(image.naturalWidth, image.naturalHeight),
    );

    canvas.width = Math.max(1, Math.round(image.naturalWidth * scale));
    canvas.height = Math.max(1, Math.round(image.naturalHeight * scale));

    const context = canvas.getContext("2d");

    if (!context) {
      throw new Error("Canvas is unavailable.");
    }

    context.drawImage(image, 0, 0, canvas.width, canvas.height);

    return canvas.toDataURL("image/png");
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

async function resolveImage(src: string): Promise<string> {
  if (/^data:image\/svg\+xml/i.test(src)) {
    return svgToPng(decodeSvgDataUrl(src));
  }

  if (/^data:image\/(png|jpe?g|gif|webp);base64,/i.test(src)) {
    if ((src.length * 3) / 4 > MAX_IMAGE_SIZE) {
      throw new Error("Image exceeds the 8 MB limit.");
    }

    if (/^data:image\/webp/i.test(src)) {
      const blob = await fetch(src).then((response) => response.blob());
      return convertBlobToPng(blob);
    }

    return src;
  }

  if (src.startsWith("data:image/")) {
    throw new Error("Unsupported embedded image format.");
  }

  let url: URL;

  try {
    url = new URL(
      src,
      typeof window !== "undefined"
        ? window.location.href
        : "https://localhost",
    );
  } catch {
    throw new Error(`Invalid image URL: ${src}`);
  }

  if (!["http:", "https:"].includes(url.protocol)) {
    throw new Error("Unsupported image URL.");
  }

  const response = await fetch(url.href, {
    mode: "cors",
    credentials: "omit",
  });

  if (!response.ok) {
    throw new Error(`Image request failed: ${response.status}`);
  }

  const blob = await response.blob();

  if (blob.size > MAX_IMAGE_SIZE) {
    throw new Error("Image exceeds the 8 MB limit.");
  }

  if (!blob.type.startsWith("image/")) {
    throw new Error("The URL does not contain an image.");
  }

  if (blob.type === "image/svg+xml") {
    return svgToPng(await blob.text());
  }

  if (
    !["image/png", "image/jpeg", "image/gif"].includes(blob.type)
  ) {
    return convertBlobToPng(blob);
  }

  return blobToDataUrl(blob);
}

function blobToDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = () => {
      if (typeof reader.result === "string") {
        resolve(reader.result);
      } else {
        reject(new Error("Unable to read image."));
      }
    };

    reader.onerror = () => reject(new Error("Unable to read image."));
    reader.readAsDataURL(blob);
  });
}

async function convertBlobToPng(blob: Blob): Promise<string> {
  const objectUrl = URL.createObjectURL(blob);

  try {
    const image = new Image();

    await new Promise<void>((resolve, reject) => {
      image.onload = () => resolve();
      image.onerror = () => reject(new Error("Unsupported image format."));
      image.src = objectUrl;
    });

    const canvas = document.createElement("canvas");

    const scale = Math.min(
      1,
      MAX_IMAGE_DIMENSION /
        Math.max(image.naturalWidth, image.naturalHeight),
    );

    canvas.width = Math.max(1, Math.round(image.naturalWidth * scale));
    canvas.height = Math.max(1, Math.round(image.naturalHeight * scale));

    const context = canvas.getContext("2d");

    if (!context) throw new Error("Canvas is unavailable.");

    context.drawImage(image, 0, 0, canvas.width, canvas.height);

    return canvas.toDataURL("image/png");
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function parseJson(code: string): Record<string, any> | null {
  try {
    const result = JSON.parse(code);

    return result && typeof result === "object" && !Array.isArray(result)
      ? result
      : null;
  } catch {
    return null;
  }
}

export async function convertMarkdownToPowerPoint(
  markdown: string,
  opts: MdToPptxOptions = {},
): Promise<Blob> {
  if (!markdown?.trim()) {
    throw new Error("The Markdown file is empty.");
  }

  if (new Blob([markdown]).size > MAX_MARKDOWN_SIZE) {
    throw new Error("The maximum Markdown file size is 28 MB.");
  }

  const source = cleanText(markdown)
    .replace(/^---\r?\n[\s\S]*?\r?\n---\r?\n/, "")
    .trim();

  if (!source) {
    throw new Error("The Markdown file contains no readable content.");
  }

  const { default: PptxGenJS } = await import("pptxgenjs");
  const pptx = new PptxGenJS();

  pptx.layout = "LAYOUT_WIDE";
  pptx.author = opts.author ?? "eSparks Document Conversion";
  pptx.title = opts.title ?? "Markdown Presentation";
  pptx.subject = "Markdown to PowerPoint conversion";
  pptx.company = "eSparks IT Solutions";

  // PptxGenJS does not support pptx.lang or theme.lang.
  pptx.theme = {
    headFontFace: FONT,
    bodyFontFace: FONT,
  };

  const bodySize = Math.min(
    24,
    Math.max(10, opts.fontSize ?? FONT_SIZE.body),
  );

  const tokens = marked.lexer(source, {
    gfm: true,
    breaks: true,
  }) as Token[];

  let slide!: PptxSlide;
  let slideNumber = 0;
  let y = CONTENT_TOP;
  let hasContent = false;
  let currentTitle = opts.title ?? "Markdown Presentation";

  const warnings: string[] = [];

  function addFooter() {
    slide.addText("eSparks Document Conversion", {
      x: MARGIN_X,
      y: FOOTER_Y,
      w: 3.5,
      h: 0.18,
      fontFace: FONT,
      fontSize: 8,
      color: COLORS.muted,
      margin: 0,
    });

    slide.addText(String(slideNumber), {
      x: SW - 1.05,
      y: FOOTER_Y,
      w: 0.4,
      h: 0.18,
      fontFace: FONT,
      fontSize: 8,
      color: COLORS.muted,
      align: "right",
      margin: 0,
    });
  }

  function addContentSlide(title: string) {
    slideNumber++;
    currentTitle = title || currentTitle;

    slide = pptx.addSlide();
    slide.background = { color: COLORS.background };

    slide.addText(currentTitle, {
      x: MARGIN_X,
      y: 0.35,
      w: CONTENT_W,
      h: 0.72,
      fontFace: FONT,
      fontSize: FONT_SIZE.h2,
      bold: true,
      color: COLORS.primary,
      margin: 0,
      valign: "middle",
      fit: "shrink",
    });

    slide.addShape(pptx.ShapeType.line, {
      x: MARGIN_X,
      y: 1.16,
      w: CONTENT_W,
      h: 0,
      line: { color: COLORS.border, width: 1 },
    });

    addFooter();

    y = CONTENT_TOP;
    hasContent = false;
  }

  function addTitleSlide(title: string) {
    slideNumber++;
    currentTitle = title || "Markdown Presentation";

    slide = pptx.addSlide();
    slide.background = { color: COLORS.background };

    slide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: SW,
      h: 0.16,
      line: { color: COLORS.primary, transparency: 100 },
      fill: { color: COLORS.primary },
    });

    slide.addText(currentTitle, {
      x: 1.1,
      y: 2.1,
      w: 11.1,
      h: 1.6,
      fontFace: FONT,
      fontSize: FONT_SIZE.title,
      bold: true,
      color: COLORS.primary,
      align: "center",
      valign: "middle",
      margin: 0.1,
      fit: "shrink",
    });

    slide.addText("Converted from Markdown", {
      x: 2,
      y: 4.05,
      w: 9.3,
      h: 0.45,
      fontFace: FONT,
      fontSize: 17,
      color: COLORS.muted,
      align: "center",
      margin: 0,
    });

    addFooter();

    y = CONTENT_TOP;
    hasContent = true;
  }

  function ensureSlide() {
    if (!slide) {
      addContentSlide(currentTitle);
    }
  }

  function nextSlide() {
    addContentSlide(`${currentTitle} (continued)`);
  }

  function ensureSpace(height: number) {
    ensureSlide();

    if (y + height <= CONTENT_BOTTOM) return;

    if (hasContent) {
      nextSlide();
    }
  }

  function drawRuns(
    inputRuns: TextRun[],
    options: {
      fontSize?: number;
      color?: string;
      bold?: boolean;
      italic?: boolean;
      align?: "left" | "center" | "right" | "justify";
      indent?: number;
      quote?: boolean;
    } = {},
  ) {
    const fontSize = options.fontSize ?? bodySize;
    const indent = options.indent ?? 0;
    const x = MARGIN_X + indent + (options.quote ? 0.2 : 0);
    const width = CONTENT_W - indent - (options.quote ? 0.2 : 0);

    const runs = mergeRuns(inputRuns);

    if (!runs.map((run) => run.text).join("").trim()) return;

    const lines = wrapRuns(runs, width - 0.08, fontSize);
    const lineHeight = (fontSize / 72) * LINE_HEIGHT;

    let index = 0;

    while (index < lines.length) {
      ensureSlide();

      if (y + lineHeight > CONTENT_BOTTOM && hasContent) {
        nextSlide();
        continue;
      }

      const availableLines = Math.max(
        1,
        Math.floor((CONTENT_BOTTOM - y - 0.12) / lineHeight),
      );

      const pageLines = lines.slice(index, index + availableLines);

      const height = Math.max(
        0.18,
        pageLines.length * lineHeight + 0.1,
      );

      ensureSpace(height);

      const richRuns: TextRun[] = [];

      pageLines.forEach((line, lineIndex) => {
        if (lineIndex) {
          richRuns.push({ text: "\n" });
        }

        richRuns.push(...line);
      });

      if (options.quote) {
        slide.addShape(pptx.ShapeType.roundRect, {
          x: MARGIN_X,
          y: y - 0.02,
          w: CONTENT_W,
          h: height + 0.06,
          rectRadius: 0.06,
          line: {
            color: COLORS.quoteBackground,
            transparency: 100,
          },
          fill: { color: COLORS.quoteBackground },
        });

        slide.addShape(pptx.ShapeType.rect, {
          x: MARGIN_X,
          y,
          w: 0.055,
          h: height,
          line: {
            color: COLORS.primary,
            transparency: 100,
          },
          fill: { color: COLORS.primary },
        });
      }

      slide.addText(
        richRuns.map((run) => ({
          text: run.text,
          options: {
            fontFace: FONT,
            fontSize,
            color: options.color ?? COLORS.body,
            ...run.options,
            bold: options.bold || run.options?.bold,
            italic: options.italic || run.options?.italic,
          },
        })),
        {
          x,
          y,
          w: width,
          h: height,
          fontFace: FONT,
          fontSize,
          color: options.color ?? COLORS.body,
          margin: 0.03,
          valign: "top",
          align: options.align ?? "left",
          breakLine: false,
          fit: "shrink",
        },
      );

      y += height + GAP;
      hasContent = true;
      index += pageLines.length;

      if (index < lines.length) {
        nextSlide();
      }
    }
  }

  function drawText(
    text: string,
    options: {
      fontSize?: number;
      color?: string;
      bold?: boolean;
      italic?: boolean;
      align?: "left" | "center" | "right" | "justify";
      indent?: number;
      quote?: boolean;
    } = {},
  ) {
    const value = cleanText(text);

    if (!value.trim()) return;

    drawRuns([{ text: value }], options);
  }

  function addHeading(text: string, level: number) {
    const value = cleanText(text).trim();

    if (!value) return;

    if (level === 1 && slideNumber === 0) {
      addTitleSlide(value);
      return;
    }

    if (level <= 2) {
      addContentSlide(value);
      return;
    }

    const sizes: Record<number, number> = {
      3: FONT_SIZE.h3,
      4: FONT_SIZE.h4,
      5: FONT_SIZE.h5,
      6: FONT_SIZE.h6,
    };

    drawText(value, {
      fontSize: sizes[level] ?? FONT_SIZE.h6,
      color: COLORS.primary,
      bold: true,
    });
  }

  function addCodeBlock(code: string, language = "") {
    const fontSize = FONT_SIZE.code;
    const padding = 0.22;
    const maxWidth = CONTENT_W - padding * 2;

    const lines = cleanText(code)
      .split("\n")
      .flatMap((line) =>
        line
          ? wrapPlainText(line, maxWidth, fontSize, MONO)
          : [""],
      );

    const lineHeight = 0.19;
    let index = 0;

    while (index < lines.length) {
      ensureSlide();

      if (y + 0.5 > CONTENT_BOTTOM && hasContent) {
        nextSlide();
        continue;
      }

      const available = Math.max(
        1,
        Math.floor((CONTENT_BOTTOM - y - 0.35) / lineHeight),
      );

      const pageLines = lines.slice(index, index + available);
      const height = pageLines.length * lineHeight + 0.36;

      ensureSpace(height);

      slide.addShape(pptx.ShapeType.roundRect, {
        x: MARGIN_X,
        y,
        w: CONTENT_W,
        h: height,
        rectRadius: 0.05,
        line: { color: COLORS.border, width: 0.6 },
        fill: { color: COLORS.codeBackground },
      });

      if (language) {
        slide.addText(language.toUpperCase(), {
          x: MARGIN_X + padding,
          y: y + 0.07,
          w: CONTENT_W - padding * 2,
          h: 0.15,
          fontFace: FONT,
          fontSize: 7,
          bold: true,
          color: COLORS.muted,
          margin: 0,
        });
      }

      slide.addText(pageLines.join("\n"), {
        x: MARGIN_X + padding,
        y: y + 0.23,
        w: CONTENT_W - padding * 2,
        h: Math.max(0.16, height - 0.28),
        fontFace: MONO,
        fontSize,
        color: "334155",
        margin: 0,
        valign: "top",
        breakLine: false,
        fit: "shrink",
      });

      y += height + GAP;
      hasContent = true;
      index += pageLines.length;

      if (index < lines.length) {
        nextSlide();
      }
    }
  }

  async function addImage(src: string, alt = "Image") {
    try {
      const data = await resolveImage(src);

      const dimensions = await new Promise<{
        width: number;
        height: number;
      }>((resolve, reject) => {
        const image = new Image();

        image.onload = () => {
          resolve({
            width: image.naturalWidth,
            height: image.naturalHeight,
          });
        };

        image.onerror = () => reject(new Error("Unable to decode image."));
        image.src = data;
      });

      const scale = Math.min(
        CONTENT_W / dimensions.width,
        (CONTENT_BOTTOM - CONTENT_TOP) / dimensions.height,
        1,
      );

      const width = dimensions.width * scale;
      const height = dimensions.height * scale;

      ensureSlide();
      ensureSpace(height + GAP);

      if (y + height > CONTENT_BOTTOM && hasContent) {
        nextSlide();
      }

      slide.addImage({
        data,
        x: MARGIN_X + (CONTENT_W - width) / 2,
        y,
        w: width,
        h: height,
        altText: alt,
      });

      y += height + GAP;
      hasContent = true;
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Unable to load image.";

      warnings.push(`${alt}: ${message}`);

      drawText(`[Image unavailable: ${alt}]`, {
        color: COLORS.muted,
        italic: true,
        fontSize: 11,
      });
    }
  }

  function getCellHeight(text: string, width: number): number {
    const lines = wrapPlainText(
      text,
      Math.max(0.3, width - 0.16),
      10,
    );

    return Math.max(0.34, lines.length * 0.18 + 0.16);
  }

  function addMarkdownTable(token: Tokens.Table) {
    const headers = token.header;
    const rows = token.rows;
    const columnCount = Math.max(1, headers.length);
    const allRows = [headers, ...rows];

    const maxLengths = Array.from(
      { length: columnCount },
      (_, index) =>
        Math.max(
          5,
          ...allRows.map((row) =>
            getPlainText(row[index]?.tokens as Token[] ?? []).length,
          ),
        ),
    );

    const totalLength = maxLengths.reduce((sum, value) => sum + value, 0);

    const colWidths = maxLengths.map(
      (value) => (value / totalLength) * CONTENT_W,
    );

    const normalizedRows = allRows.map((row) =>
      Array.from({ length: columnCount }, (_, index) => row[index]),
    );

    const rowHeights = normalizedRows.map((row) =>
      Math.max(
        0.34,
        ...row.map((cell, index) =>
          getCellHeight(
            getPlainText(cell?.tokens as Token[] ?? []),
            colWidths[index],
          ),
        ),
      ),
    );

    const availableHeight = CONTENT_BOTTOM - CONTENT_TOP;
    let start = 0;

    while (start < normalizedRows.length) {
      const repeatedHeader = start > 0;
      const chunk: number[] = [];
      let usedHeight = 0;

      if (repeatedHeader) {
        chunk.push(0);
        usedHeight += rowHeights[0];
      }

      let cursor = start === 0 ? 0 : 1;

      while (
        cursor < normalizedRows.length &&
        usedHeight + rowHeights[cursor] <= availableHeight
      ) {
        chunk.push(cursor);
        usedHeight += rowHeights[cursor];
        cursor++;
      }

      if (!chunk.length) {
        chunk.push(start);
        cursor = start + 1;
        usedHeight = rowHeights[start];
      }

      ensureSlide();

      if (y + usedHeight > CONTENT_BOTTOM && hasContent) {
        nextSlide();
      }

      const tableRows = chunk.map((rowIndex) =>
        normalizedRows[rowIndex].map((cell, columnIndex) => {
          const isHeader = rowIndex === 0;
          const runs = cell
            ? inlineRuns(cell.tokens as Token[])
            : [];

          const richText = runs.length
            ? runs.map((run) => ({
                text: run.text,
                options: {
                  ...run.options,
                  fontFace: FONT,
                  fontSize: 10,
                  color: isHeader ? "FFFFFF" : COLORS.body,
                  bold: isHeader || Boolean(run.options?.bold),
                },
              }))
            : [{ text: "", options: {} }];

          return {
            text: richText,
            options: {
              fill: {
                color: isHeader
                  ? COLORS.primary
                  : rowIndex % 2 === 0
                    ? COLORS.tableAlternate
                    : COLORS.background,
              },
              fontFace: FONT,
              fontSize: 10,
              color: isHeader ? "FFFFFF" : COLORS.body,
              bold: isHeader,
              align: token.align?.[columnIndex] ?? "left",
              valign: "middle",
              margin: 0.07,
            },
          };
        }),
      );

      const actualHeight = chunk.reduce(
        (sum, rowIndex) => sum + rowHeights[rowIndex],
        0,
      );

      slide.addTable(tableRows as any, {
        x: MARGIN_X,
        y,
        w: CONTENT_W,
        colW: colWidths,
        rowH: chunk.map((rowIndex) => rowHeights[rowIndex]),
        fontFace: FONT,
        fontSize: 10,
        color: COLORS.body,
        border: {
          type: "solid",
          color: COLORS.border,
          pt: 0.7,
        },
        margin: 0.07,
        valign: "middle",
        autoPage: false,
      } as any);

      y += actualHeight + GAP;
      hasContent = true;
      start = cursor;

      if (start < normalizedRows.length) {
        nextSlide();
      }
    }
  }

  function readHtmlCell(cell: HTMLTableCellElement): HtmlTableCell {
    const runs = htmlInlineRuns(cell);
    const alignValue =
      cell.getAttribute("align") ||
      cell.style.textAlign ||
      "left";

    const align = ["left", "center", "right"].includes(alignValue)
      ? (alignValue as "left" | "center" | "right")
      : "left";

    return {
      runs,
      text: runs.map((run) => run.text).join(""),
      colspan: Math.max(1, cell.colSpan || 1),
      rowspan: Math.max(1, cell.rowSpan || 1),
      header:
        cell.tagName.toLowerCase() === "th" ||
        cell.closest("thead") !== null,
      align,
    };
  }

  function addHtmlTable(table: HTMLTableElement) {
    const sourceRows = Array.from(table.rows);

    if (!sourceRows.length) return;

    const grid: Array<Array<HtmlTableCell | null>> = [];
    const rowCount = sourceRows.length;
    let columnCount = 0;

    sourceRows.forEach((row, rowIndex) => {
      grid[rowIndex] ??= [];
      let columnIndex = 0;

      for (const sourceCell of Array.from(row.cells)) {
        while (grid[rowIndex][columnIndex] !== undefined) {
          columnIndex++;
        }

        const cell = readHtmlCell(sourceCell);
        const colspan = cell.colspan;
        const rowspan = Math.min(cell.rowspan, rowCount - rowIndex);

        grid[rowIndex][columnIndex] = cell;

        for (let r = rowIndex; r < rowIndex + rowspan; r++) {
          grid[r] ??= [];

          for (let c = columnIndex; c < columnIndex + colspan; c++) {
            if (r === rowIndex && c === columnIndex) continue;
            grid[r][c] = null;
          }
        }

        columnIndex += colspan;
        columnCount = Math.max(columnCount, columnIndex);
      }
    });

    columnCount = Math.max(1, columnCount);

    const widths = Array(columnCount).fill(CONTENT_W / columnCount);

    const rowHeights = sourceRows.map((row, rowIndex) => {
      const cells = Array.from(row.cells);

      return Math.max(
        0.34,
        ...cells.map((sourceCell) => {
          const value = readHtmlCell(sourceCell);
          const cellWidth =
            widths.reduce((sum, item) => sum + item, 0) /
            Math.max(1, cells.length);

          return getCellHeight(value.text, cellWidth);
        }),
      );
    });

    const rowsPerPage = Math.max(
      1,
      Math.floor((CONTENT_BOTTOM - CONTENT_TOP) / 0.55),
    );

    let start = 0;

    while (start < rowCount) {
      ensureSlide();

      const chunkEnd = Math.min(rowCount, start + rowsPerPage);
      const chunkHeight = rowHeights
        .slice(start, chunkEnd)
        .reduce((a, b) => a + b, 0);

      if (y + chunkHeight > CONTENT_BOTTOM && hasContent) {
        nextSlide();
      }

      const tableRows: any[] = [];

      for (let r = start; r < chunkEnd; r++) {
        const row: any[] = [];

        for (let c = 0; c < columnCount; c++) {
          const cell = grid[r]?.[c];

          // Keep every logical column in every row.
          // Merged cells are flattened into their visible grid slots.
          if (cell === null || !cell) {
            row.push({
              text: "",
              options: {
                fill: {
                  color: r % 2 === 0
                    ? COLORS.tableAlternate
                    : COLORS.background,
                },
                fontFace: FONT,
                fontSize: 9,
                color: COLORS.body,
                margin: 0.06,
                valign: "middle",
              },
            });

            continue;
          }

          const isHeader = cell.header;

          const richText = cell.runs.length
            ? cell.runs.map((run) => ({
                text: run.text,
                options: {
                  ...run.options,
                  fontFace: FONT,
                  fontSize: 9,
                  color: isHeader ? "FFFFFF" : COLORS.body,
                  bold: isHeader || Boolean(run.options?.bold),
                },
              }))
            : [{ text: "", options: {} }];

          row.push({
            text: richText,
            options: {
              fill: {
                color: isHeader
                  ? COLORS.primary
                  : r % 2 === 0
                    ? COLORS.tableAlternate
                    : COLORS.background,
              },
              fontFace: FONT,
              fontSize: 9,
              color: isHeader ? "FFFFFF" : COLORS.body,
              bold: isHeader,
              align: cell.align,
              valign: "middle",
              margin: 0.06,
            },
          });
        }

        tableRows.push(row);
      }

      slide.addTable(tableRows, {
        x: MARGIN_X,
        y,
        w: CONTENT_W,
        colW: widths,
        rowH: rowHeights.slice(start, chunkEnd),
        fontFace: FONT,
        fontSize: 9,
        border: {
          type: "solid",
          color: COLORS.border,
          pt: 0.7,
        },
        margin: 0.06,
        valign: "middle",
        autoPage: false,
      } as any);

      y += chunkHeight + GAP;
      hasContent = true;
      start = chunkEnd;

      if (start < rowCount) {
        nextSlide();
      }
    }
  }

  function htmlInlineRuns(element: Element): TextRun[] {
    const output: TextRun[] = [];

    function visit(node: Node, inherited: Record<string, any> = {}) {
      if (node.nodeType === Node.TEXT_NODE) {
        if (node.textContent) {
          output.push({
            text: node.textContent,
            options: { ...inherited },
          });
        }

        return;
      }

      if (node.nodeType !== Node.ELEMENT_NODE) return;

      const child = node as HTMLElement;
      const tag = child.tagName.toLowerCase();
      const style = { ...inherited };

      if (["strong", "b"].includes(tag)) style.bold = true;
      if (["em", "i"].includes(tag)) style.italic = true;
      if (tag === "u") style.underline = true;
      if (["s", "del", "strike"].includes(tag)) style.strike = true;

      if (tag === "code") {
        style.fontFace = MONO;
        style.color = "B42318";
      }

      if (tag === "a") {
        const href = safeUrl(child.getAttribute("href") ?? "");

        if (href) {
          style.hyperlink = { url: href };
          style.color = COLORS.primary;
          style.underline = { color: COLORS.primary };
        }
      }

      if (tag === "span") {
        const color = child.style.color;

        if (color && /^#[0-9a-f]{3,8}$/i.test(color)) {
          style.color = color.slice(1);
        }

        if (child.style.fontWeight === "bold") style.bold = true;
        if (child.style.fontStyle === "italic") style.italic = true;

        if (child.style.textDecoration.includes("underline")) {
          style.underline = true;
        }
      }

      if (tag === "br") {
        output.push({ text: "\n", options: style });
        return;
      }

      for (const nested of Array.from(child.childNodes)) {
        visit(nested, style);
      }
    }

    for (const node of Array.from(element.childNodes)) {
      visit(node);
    }

    return mergeRuns(output);
  }

  async function renderHtmlNode(element: Element): Promise<void> {
    const tag = element.tagName.toLowerCase();

    if (/^h[1-6]$/.test(tag)) {
      addHeading(element.textContent ?? "", Number(tag[1]));
      return;
    }

    if (tag === "table") {
      addHtmlTable(element as HTMLTableElement);
      return;
    }

    if (tag === "svg") {
      const svg = element.cloneNode(true) as SVGElement;

      if (!svg.getAttribute("xmlns")) {
        svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
      }

      const markup = new XMLSerializer().serializeToString(svg);

      await addImage(
        `data:image/svg+xml,${encodeURIComponent(markup)}`,
        svg.getAttribute("aria-label") ?? "SVG graphic",
      );

      return;
    }

    if (tag === "img") {
      const image = element as HTMLImageElement;
      const src = image.getAttribute("src");

      if (src) {
        await addImage(src, image.alt || "HTML image");
      }

      return;
    }

    if (tag === "pre") {
      const code = element.querySelector("code");
      const language =
        code?.className.match(/language-([\w-]+)/)?.[1] ?? "";

      const value = code?.textContent ?? element.textContent ?? "";

      await renderCode(value, language);
      return;
    }

    if (tag === "p") {
      const images = Array.from(element.querySelectorAll("img"));
      const clone = element.cloneNode(true) as HTMLElement;

      clone.querySelectorAll("img").forEach((image) => image.remove());

      drawRuns(htmlInlineRuns(clone));

      for (const image of images) {
        const src = image.getAttribute("src");

        if (src) {
          await addImage(
            src,
            image.getAttribute("alt") ?? "HTML image",
          );
        }
      }

      return;
    }

    if (tag === "blockquote") {
      for (const child of Array.from(element.children)) {
        if (child.tagName.toLowerCase() === "p") {
          drawRuns(htmlInlineRuns(child), { quote: true });
        } else {
          await renderHtmlNode(child);
        }
      }

      return;
    }

    if (tag === "ul" || tag === "ol") {
      const items = Array.from(element.children).filter(
        (child) => child.tagName.toLowerCase() === "li",
      );

      for (const [index, item] of items.entries()) {
        const prefix = tag === "ol" ? `${index + 1}. ` : "• ";

        drawRuns([
          {
            text: prefix,
            options: {
              bold: true,
              color: COLORS.primary,
            },
          },
          ...htmlInlineRuns(item),
        ]);
      }

      return;
    }

    if (tag === "hr") {
      ensureSlide();
      ensureSpace(0.22);

      slide.addShape(pptx.ShapeType.line, {
        x: MARGIN_X,
        y: y + 0.03,
        w: CONTENT_W,
        h: 0,
        line: { color: COLORS.border, width: 1 },
      });

      y += 0.22;
      hasContent = true;
      return;
    }

    if (tag === "br") {
      drawRuns([{ text: "\n" }]);
      return;
    }

    if (["script", "style", "input", "meta", "link"].includes(tag)) {
      return;
    }

    if (
      ["div", "section", "article", "main", "figure", "header", "footer"].includes(tag)
    ) {
      for (const child of Array.from(element.children)) {
        await renderHtmlNode(child);
      }

      if (!element.children.length && element.textContent?.trim()) {
        drawRuns(htmlInlineRuns(element));
      }

      return;
    }

    if (element.children.length) {
      drawRuns(htmlInlineRuns(element));
    } else if (element.textContent?.trim()) {
      drawRuns(htmlInlineRuns(element));
    }
  }

  async function renderHtml(html: string) {
    const parser = new DOMParser();
    const parsed = parser.parseFromString(html, "text/html");

    for (const node of Array.from(parsed.body.childNodes)) {
      if (node.nodeType === Node.TEXT_NODE) {
        const value = node.textContent?.trim();

        if (value) drawText(value);
      } else if (node instanceof Element) {
        await renderHtmlNode(node);
      }
    }
  }

  function addChart(config: Record<string, any>) {
    const typeName = String(
      config.type ?? config.chartType ?? "bar",
    ).toLowerCase();

    const chartTypes: Record<string, any> = {
      bar: pptx.ChartType.bar,
      column: pptx.ChartType.bar,
      line: pptx.ChartType.line,
      pie: pptx.ChartType.pie,
      doughnut: pptx.ChartType.doughnut,
      area: pptx.ChartType.area,
    };

    const chartType = chartTypes[typeName];

    if (!chartType) {
      warnings.push(`Unsupported chart type: ${typeName}`);
      addCodeBlock(JSON.stringify(config, null, 2), "json");
      return;
    }

    const categories = config.categories ?? config.labels;
    let sourceSeries = config.series;

    if (!Array.isArray(sourceSeries) && Array.isArray(config.data)) {
      sourceSeries = [
        {
          name: config.seriesName ?? "Series",
          values: config.data,
        },
      ];
    }

    if (!Array.isArray(categories) || !Array.isArray(sourceSeries)) {
      warnings.push("Chart data must include categories and series.");
      addCodeBlock(JSON.stringify(config, null, 2), "json");
      return;
    }

    const series = sourceSeries.map((item: any) => ({
      name: String(item.name ?? "Series"),
      labels: categories.map(String),
      values: (item.values ?? item.data ?? []).map(Number),
    }));

    const chartHeight = Math.min(4.6, CONTENT_BOTTOM - CONTENT_TOP);

    ensureSlide();
    ensureSpace(chartHeight + GAP);

    slide.addChart(chartType, series, {
      x: MARGIN_X + 0.1,
      y,
      w: CONTENT_W - 0.2,
      h: chartHeight,
      catAxisLabelFontFace: FONT,
      catAxisLabelFontSize: 10,
      valAxisLabelFontFace: FONT,
      valAxisLabelFontSize: 9,
      showTitle: Boolean(config.title),
      title: String(config.title ?? ""),
      showLegend:
        series.length > 1 ||
        typeName === "pie" ||
        typeName === "doughnut",
      showValue: true,
      showPercent: typeName === "pie" || typeName === "doughnut",
      showCatName: false,
      chartColors: [
        "087E8B",
        "16B8C4",
        "F59E0B",
        "8B5CF6",
        "0F766E",
        "E76F51",
      ],
      showBorder: false,
      ...(typeName === "column"
        ? {
            catAxisLabelRotate: 0,
            valAxisMinVal: 0,
            catAxisLabelPos: "low",
            valAxisLabelPos: "nextTo",
          }
        : {}),
      ...(typeName === "bar" || typeName === "column"
        ? { barDir: typeName === "column" ? "col" : "bar" }
        : {}),
    } as any);

    y += chartHeight + GAP;
    hasContent = true;
  }

  function addCustomShape(config: Record<string, any>) {
    const shapeTypes: Record<string, any> = {
      rect: pptx.ShapeType.rect,
      rectangle: pptx.ShapeType.rect,
      roundRect: pptx.ShapeType.roundRect,
      roundedRectangle: pptx.ShapeType.roundRect,
      ellipse: pptx.ShapeType.ellipse,
      circle: pptx.ShapeType.ellipse,
      diamond: pptx.ShapeType.diamond,
      triangle: pptx.ShapeType.triangle,
      line: pptx.ShapeType.line,
      chevron: pptx.ShapeType.chevron,
      hexagon: pptx.ShapeType.hexagon,
      pentagon: pptx.ShapeType.pentagon,
      star5: pptx.ShapeType.star5,
      arrow: pptx.ShapeType.rightArrow,
      rightArrow: pptx.ShapeType.rightArrow,
      leftArrow: pptx.ShapeType.leftArrow,
      upArrow: pptx.ShapeType.upArrow,
      downArrow: pptx.ShapeType.downArrow,
    };

    const type = shapeTypes[String(config.type ?? "rect")];

    if (!type) {
      warnings.push(`Unsupported custom shape: ${String(config.type)}`);
      return;
    }

    const x = MARGIN_X + Number(config.x ?? 0);
    const shapeY = y + Number(config.y ?? 0);
    const w = Math.max(0.05, Number(config.w ?? 1));
    const h = Math.max(0.05, Number(config.h ?? 0.7));

    ensureSlide();

    if (y + h + GAP > CONTENT_BOTTOM && hasContent) {
      nextSlide();
    }

    const actualY = y + Number(config.y ?? 0);

    slide.addShape(type, {
      x,
      y: actualY,
      w,
      h,
      fill: {
        color: String(config.fill ?? COLORS.accent).replace("#", ""),
        transparency: Number(config.transparency ?? 0),
      },
      line: {
        color: String(config.line ?? COLORS.primary).replace("#", ""),
        width: Number(config.lineWidth ?? 1.2),
      },
      rotate: Number(config.rotate ?? 0),
    });

    const label = config.label ?? config.text;

    if (label) {
      slide.addText(String(label), {
        x,
        y: actualY + h / 2 - 0.15,
        w,
        h: 0.3,
        fontFace: FONT,
        fontSize: Number(config.fontSize ?? 11),
        bold: true,
        color: String(config.textColor ?? "FFFFFF").replace("#", ""),
        align: "center",
        valign: "middle",
        margin: 0.02,
        fit: "shrink",
      });
    }

    y = Math.max(y, actualY + h + GAP);
    hasContent = true;
  }

  async function renderMermaid(sourceText: string) {
    try {
      const mermaidModule = await import("mermaid");
      const mermaid = mermaidModule.default;

      mermaid.initialize({
        startOnLoad: false,
        securityLevel: "strict",
        theme: "neutral",
      });

      const id = `mermaid-${Math.random().toString(36).slice(2)}`;
      const result = await mermaid.render(id, sourceText);

      await addImage(result.svg, "Mermaid diagram");
    } catch (error) {
      warnings.push(
        `Mermaid rendering failed: ${
          error instanceof Error ? error.message : "Unknown error"
        }`,
      );

      addCodeBlock(sourceText, "mermaid");
    }
  }

  async function renderCode(code: string, language = "") {
    const normalizedLanguage = language.toLowerCase();

    if (normalizedLanguage === "mermaid") {
      await renderMermaid(code);
      return;
    }

    if (normalizedLanguage === "pptx-chart") {
      const config = parseJson(code);

      if (config) {
        addChart(config);
      } else {
        addCodeBlock(code, language);
      }

      return;
    }

    if (normalizedLanguage === "pptx-shape") {
      const config = parseJson(code);

      if (config) {
        addCustomShape(config);
      } else {
        addCodeBlock(code, language);
      }

      return;
    }

    addCodeBlock(code, language);
  }

  async function addList(token: Tokens.List, level = 0) {
    for (const [index, item] of token.items.entries()) {
      const prefix = item.task
        ? item.checked
          ? "☑ "
          : "☐ "
        : token.ordered
          ? `${Number(token.start || 1) + index}. `
          : "• ";

      const children = (item.tokens ?? []) as Token[];

      const nestedLists = children.filter(
        (child) => child.type === "list",
      ) as Tokens.List[];

      const contentTokens = children.filter(
        (child) => child.type !== "list",
      );

      const content = inlineRuns(contentTokens);

      if (content.map((run) => run.text).join("").trim()) {
        drawRuns(
          [
            {
              text: prefix,
              options: {
                bold: true,
                color: COLORS.primary,
              },
            },
            ...content,
          ],
          {
            indent: Math.min(level * 0.28, 1.3),
          },
        );
      }

      for (const child of contentTokens) {
        if (child.type === "image") {
          const image = child as Tokens.Image;

          await addImage(image.href, image.text || "Markdown image");
        }

        if (child.type === "code") {
          const code = child as Tokens.Code;

          await renderCode(code.text, code.lang ?? "");
        }
      }

      for (const nested of nestedLists) {
        await addList(nested, level + 1);
      }
    }
  }

  async function renderToken(token: Token, quote = false): Promise<void> {
    switch (token.type) {
      case "space":
        y += 0.04;
        return;

      case "heading": {
        const heading = token as Tokens.Heading;

        addHeading(
          getPlainText(heading.tokens as Token[]),
          heading.depth,
        );

        return;
      }

      case "paragraph": {
        const paragraph = token as Tokens.Paragraph;
        const runs = inlineRuns(paragraph.tokens as Token[]);

        drawRuns(runs, { quote });
        return;
      }

      case "text": {
        const item = token as Tokens.Text;

        if (item.tokens?.length) {
          drawRuns(inlineRuns(item.tokens as Token[]), { quote });
        } else {
          drawText(item.text, { quote });
        }

        return;
      }

      case "list":
        await addList(token as Tokens.List);
        return;

      case "blockquote": {
        const blockquote = token as Tokens.Blockquote;

        for (const child of blockquote.tokens as Token[]) {
          await renderToken(child, true);
        }

        return;
      }

      case "code": {
        const code = token as Tokens.Code;

        await renderCode(code.text, code.lang ?? "");
        return;
      }

      case "table":
        addMarkdownTable(token as Tokens.Table);
        return;

      case "hr":
        ensureSlide();
        ensureSpace(0.22);

        slide.addShape(pptx.ShapeType.line, {
          x: MARGIN_X,
          y: y + 0.03,
          w: CONTENT_W,
          h: 0,
          line: { color: COLORS.border, width: 1 },
        });

        y += 0.22;
        hasContent = true;
        return;

      case "image": {
        const image = token as Tokens.Image;

        await addImage(image.href, image.text || "Markdown image");
        return;
      }

      case "html": {
        const html = token as Tokens.HTML;

        await renderHtml(html.text);
        return;
      }

      case "def":
        return;

      default: {
        if ("tokens" in token && Array.isArray(token.tokens)) {
          for (const child of token.tokens as Token[]) {
            await renderToken(child, quote);
          }
        } else if ("text" in token && typeof token.text === "string") {
          drawText(token.text, { quote });
        }
      }
    }
  }

  try {
    for (const token of tokens) {
      await renderToken(token);
    }

    if (slideNumber === 0) {
      addTitleSlide(opts.title ?? "Markdown Presentation");

      drawText("No readable Markdown content was found.", {
        color: COLORS.muted,
      });
    }

    if (warnings.length && slide) {
      slide.addNotes(
        `Conversion warnings:\n${warnings
          .map((warning) => `- ${warning}`)
          .join("\n")}`,
      );
    }

    const output = await pptx.write({
      outputType: "blob",
    });

    if (output instanceof Blob) {
      return output;
    }

    return new Blob([output as BlobPart], {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
  } catch (error) {
    console.error("Markdown to PowerPoint conversion failed:", error);

    throw new Error(
      error instanceof Error
        ? `PowerPoint conversion failed: ${error.message}`
        : "PowerPoint conversion failed.",
    );
  }
} 
