import { marked, type Token, type Tokens } from "marked";
import hljs from "highlight.js/lib/common";
import mermaid from "mermaid";

export type MdToPptxOptions = {
  fontSize?: number;
  title?: string;
  author?: string;
};

type TextRun = {
  text: string;
  options?: Record<string, any>;
};

type PptxSlide = ReturnType<
  InstanceType<typeof import("pptxgenjs").default>["addSlide"]
>;

type Align = "left" | "center" | "right";

type HtmlTableCell = {
  runs: TextRun[];
  text: string;
  colspan: number;
  rowspan: number;
  header: boolean;
  align: Align;
};

const SW = 13.333;
const SH = 7.5;
const MARGIN_X = 0.65;
const CONTENT_W = SW - MARGIN_X * 2;
const CONTENT_TOP = 1.35;
const CONTENT_BOTTOM = 6.98;
const FOOTER_Y = 7.17;
const GAP = 0.12;

const FONT = "Arial";
const MONO = "Consolas";

const MAX_MARKDOWN_SIZE = 28 * 1024 * 1024;
const MAX_IMAGE_SIZE = 8 * 1024 * 1024;
const MAX_IMAGE_DIMENSION = 3000;
const MAX_TABLE_COLUMNS = 30;
const MAX_TABLE_ROWS = 5000;
const MAX_CODE_LINES_PER_SLIDE = 25;

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
  codeText: "334155",
  keyword: "7C3AED",
  string: "0F766E",
  number: "B45309",
  comment: "64748B",
  function: "0369A1",
  variable: "BE185D",
  literal: "B45309",
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

// Vertical spacing between rendered text lines.
const LINE_HEIGHT = 1.4;

// Additional gap between consecutive PowerPoint text lines (in inches).
const LINE_GAP = 0.035;

// Additional spacing after paragraphs.
const PARAGRAPH_GAP = 0.14;

// Additional spacing after headings.
const HEADING_GAP = 0.18;

/* -------------------------------------------------------------------------- */
/*                               TEXT HELPERS                                 */
/* -------------------------------------------------------------------------- */

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
    const base =
      typeof window !== "undefined"
        ? window.location.href
        : "https://localhost/";

    const url = new URL(value, base);

    if (["http:", "https:", "mailto:"].includes(url.protocol)) {
      return url.href;
    }
  } catch {
    // Invalid URL.
  }

  return undefined;
}

function getPlainText(tokens: Token[] = []): string {
  return tokens
    .map((token) => {
      if (token.type === "text") {
        const item = token as Tokens.Text;

        return item.tokens?.length
          ? getPlainText(item.tokens as Token[])
          : item.text;
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

/* -------------------------------------------------------------------------- */
/*                            INLINE MARKDOWN                                 */
/* -------------------------------------------------------------------------- */

function inlineRuns(
  tokens: Token[] = [],
  inherited: Record<string, any> = {},
): TextRun[] {
  const output: TextRun[] = [];

  const add = (
    text: string,
    options = inherited,
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
            ...inlineRuns(
              item.tokens as Token[],
              inherited,
            ),
          );
        } else {
          add(item.text);
        }

        break;
      }

      case "strong":
        output.push(
          ...inlineRuns(
            token.tokens as Token[],
            {
              ...inherited,
              bold: true,
            },
          ),
        );
        break;

      case "em":
        output.push(
          ...inlineRuns(
            token.tokens as Token[],
            {
              ...inherited,
              italic: true,
            },
          ),
        );
        break;

      case "del":
        output.push(
          ...inlineRuns(
            token.tokens as Token[],
            {
              ...inherited,
              strike: true,
            },
          ),
        );
        break;

      case "codespan":
        add(
          (token as Tokens.Codespan).text,
          {
            ...inherited,
            fontFace: MONO,
            color: "B42318",
            fill: {
              color: "F1F5F9",
            },
          },
        );
        break;

      case "link": {
        const link = token as Tokens.Link;
        const href = safeUrl(link.href);

        output.push(
          ...inlineRuns(
            link.tokens as Token[],
            {
              ...inherited,
              color: COLORS.primary,
              underline: {
                color: COLORS.primary,
              },
              ...(href
                ? {
                    hyperlink: {
                      url: href,
                    },
                  }
                : {}),
            },
          ),
        );

        break;
      }

      case "image":
        add(
          (token as Tokens.Image).text || "[Image]",
        );
        break;

      case "br":
        add("\n");
        break;

      case "html":
        add(
          (token as Tokens.HTML).text
            .replace(/<br\s*\/?>/gi, "\n")
            .replace(/<[^>]*>/g, ""),
        );
        break;

      default:
        if (
          "tokens" in token &&
          Array.isArray(token.tokens)
        ) {
          output.push(
            ...inlineRuns(
              token.tokens as Token[],
              inherited,
            ),
          );
        } else if (
          "text" in token &&
          typeof token.text === "string"
        ) {
          add(token.text);
        }
    }
  }

  return mergeRuns(output);
}

/* -------------------------------------------------------------------------- */
/*                            TEXT MEASUREMENT                                */
/* -------------------------------------------------------------------------- */

function measureText(
  text: string,
  fontSize: number,
  bold = false,
  fontFace = FONT,
): number {
  if (typeof document === "undefined") {
    return (text.length * fontSize * 0.6) / 72;
  }

  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");

  if (!context) {
    return (text.length * fontSize * 0.6) / 72;
  }

  // PowerPoint font size is in points.
  // Canvas font size is in pixels.
  const fontSizePx = fontSize * (96 / 72);

  context.font = `${bold ? "700" : "400"} ${fontSizePx}px "${fontFace}"`;

  // Convert canvas pixels into PowerPoint inches.
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
        measureText(
          current + word,
          fontSize,
          false,
          fontFace,
        ) <= maxWidth
      ) {
        current += word;
        continue;
      }

      output.push(current.trimEnd());
      current = word.trimStart();

      while (
        current &&
        measureText(
          current,
          fontSize,
          false,
          fontFace,
        ) > maxWidth
      ) {
        let cut = 0;

        for (let i = 1; i <= current.length; i++) {
          if (
            measureText(
              current.slice(0, i),
              fontSize,
              false,
              fontFace,
            ) > maxWidth
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

    if (current.trim()) {
      output.push(current.trimEnd());
    }
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

  const push = (
    text: string,
    options: Record<string, any>,
  ) => {
    if (!text) return;

    const line = lines[lines.length - 1];
    const previous = line[line.length - 1];

    if (
      previous &&
      JSON.stringify(previous.options ?? {}) ===
        JSON.stringify(options)
    ) {
      previous.text += text;
    } else {
      line.push({
        text,
        options: { ...options },
      });
    }
  };

  for (const run of runs) {
    const style = run.options ?? {};

    const fontSize = Number(
      style.fontSize ?? defaultSize,
    );

    const fontFace = String(
      style.fontFace ?? FONT,
    );

    const bold = Boolean(style.bold);

    for (const part of cleanText(run.text).split(/(\n|\s+)/)) {
      if (!part) continue;

      if (part === "\n") {
        lines.push([]);
        currentWidth = 0;
        continue;
      }

      let remaining = part;

      while (remaining) {
        const whitespace = /^\s+$/.test(remaining);

        const partWidth = measureText(
          remaining,
          fontSize,
          bold,
          fontFace,
        );

        if (
          currentWidth > 0 &&
          currentWidth + partWidth > maxWidth &&
          !whitespace
        ) {
          lines.push([]);
          currentWidth = 0;
          remaining = remaining.trimStart();

          if (!remaining) break;
        }

        const available = Math.max(
          0.01,
          maxWidth - currentWidth,
        );

        if (
          measureText(
            remaining,
            fontSize,
            bold,
            fontFace,
          ) <= available
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

/* -------------------------------------------------------------------------- */
/*                              IMAGE HELPERS                                 */
/* -------------------------------------------------------------------------- */

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

      const bytes = Uint8Array.from(
        binary,
        (char) => char.charCodeAt(0),
      );

      return new TextDecoder().decode(bytes);
    }

    return decodeURIComponent(payload);
  } catch {
    throw new Error("Unable to decode SVG data.");
  }
}

function sanitizeSvg(
  svg: string,
): {
  markup: string;
  width: number;
  height: number;
} {
  const parser = new DOMParser();

  const parsed = parser.parseFromString(
    svg,
    "image/svg+xml",
  );

  if (parsed.querySelector("parsererror")) {
    throw new Error("Invalid SVG markup.");
  }

  const root = parsed.documentElement;

  if (root.tagName.toLowerCase() !== "svg") {
    throw new Error("SVG root element is missing.");
  }

  // Remove unsafe or unsupported elements.
  parsed
    .querySelectorAll(
      "script, foreignObject, iframe, object, embed",
    )
    .forEach((node) => node.remove());

  // Remove event handlers and unsafe external references.
  parsed.querySelectorAll("*").forEach((element) => {
    for (const attribute of Array.from(element.attributes)) {
      const name = attribute.name.toLowerCase();
      const value = attribute.value.trim();

      if (name.startsWith("on")) {
        element.removeAttribute(attribute.name);
      }

      if (
        ["href", "xlink:href"].includes(name) &&
        /^(javascript:|https?:\/\/|\/\/)/i.test(value)
      ) {
        element.removeAttribute(attribute.name);
      }
    }
  });

  root.setAttribute(
    "xmlns",
    "http://www.w3.org/2000/svg",
  );

  // Always prioritize the actual SVG viewBox.
  const viewBox = root.getAttribute("viewBox");

  const viewBoxValues = viewBox
    ?.trim()
    .split(/[\s,]+/)
    .map(Number);

  let width = 0;
  let height = 0;

  if (
    viewBoxValues?.length === 4 &&
    viewBoxValues.every(Number.isFinite) &&
    viewBoxValues[2] > 0 &&
    viewBoxValues[3] > 0
  ) {
    width = viewBoxValues[2];
    height = viewBoxValues[3];
  } else {
    width = Number.parseFloat(
      root.getAttribute("width") ?? "",
    );

    height = Number.parseFloat(
      root.getAttribute("height") ?? "",
    );
  }

  if (!Number.isFinite(width) || width <= 0) {
    width = 1200;
  }

  if (!Number.isFinite(height) || height <= 0) {
    height = 800;
  }

  // Preserve the original aspect ratio.
  const ratio = width / height;

  if (width > MAX_IMAGE_DIMENSION) {
    width = MAX_IMAGE_DIMENSION;
    height = width / ratio;
  }

  if (height > MAX_IMAGE_DIMENSION) {
    height = MAX_IMAGE_DIMENSION;
    width = height * ratio;
  }

  width = Math.max(1, Math.round(width));
  height = Math.max(1, Math.round(height));

  // Create a viewBox if Mermaid did not provide one.
  if (!viewBox) {
    root.setAttribute(
      "viewBox",
      `0 0 ${width} ${height}`,
    );
  }

  root.setAttribute("width", String(width));
  root.setAttribute("height", String(height));

  root.setAttribute(
    "preserveAspectRatio",
    "xMidYMid meet",
  );

  return {
    markup: new XMLSerializer().serializeToString(parsed),
    width,
    height,
  };
}

async function svgToPng(
  svg: string,
): Promise<string> {
  const safe = sanitizeSvg(svg);

  const blob = new Blob(
    [safe.markup],
    {
      type: "image/svg+xml;charset=utf-8",
    },
  );

  const objectUrl = URL.createObjectURL(blob);

  try {
    const image = new Image();

    await new Promise<void>((resolve, reject) => {
      image.onload = () => resolve();

      image.onerror = () => {
        reject(
          new Error("Unable to render Mermaid SVG."),
        );
      };

      image.src = objectUrl;
    });

    // Use the sanitized dimensions as the source of truth.
    const sourceWidth = safe.width;
    const sourceHeight = safe.height;

    // Render at higher resolution for better PowerPoint quality.
    const scale = Math.min(
      2,
      MAX_IMAGE_DIMENSION /
        Math.max(sourceWidth, sourceHeight),
    );

    const canvas = document.createElement("canvas");

    canvas.width = Math.max(
      1,
      Math.round(sourceWidth * scale),
    );

    canvas.height = Math.max(
      1,
      Math.round(sourceHeight * scale),
    );

    const context = canvas.getContext("2d");

    if (!context) {
      throw new Error("Canvas is unavailable.");
    }

    context.clearRect(
      0,
      0,
      canvas.width,
      canvas.height,
    );

    context.drawImage(
      image,
      0,
      0,
      canvas.width,
      canvas.height,
    );

    return canvas.toDataURL("image/png");
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function blobToDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = () =>
      typeof reader.result === "string"
        ? resolve(reader.result)
        : reject(new Error("Unable to read image."));

    reader.onerror = () => {
      reject(new Error("Unable to read image."));
    };

    reader.readAsDataURL(blob);
  });
}

async function convertBlobToPng(blob: Blob): Promise<string> {
  const objectUrl = URL.createObjectURL(blob);

  try {
    const image = new Image();

    await new Promise<void>((resolve, reject) => {
      image.onload = () => resolve();

      image.onerror = () => {
        reject(new Error("Unsupported image format."));
      };

      image.src = objectUrl;
    });

    const scale = Math.min(
      1,
      MAX_IMAGE_DIMENSION /
        Math.max(image.naturalWidth, image.naturalHeight),
    );

    const canvas = document.createElement("canvas");

    canvas.width = Math.max(
      1,
      Math.round(image.naturalWidth * scale),
    );

    canvas.height = Math.max(
      1,
      Math.round(image.naturalHeight * scale),
    );

    const context = canvas.getContext("2d");

    if (!context) {
      throw new Error("Canvas is unavailable.");
    }

    context.drawImage(
      image,
      0,
      0,
      canvas.width,
      canvas.height,
    );

    return canvas.toDataURL("image/png");
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

async function resolveImage(src: string): Promise<string> {
  const value = src.trim();

  if (/^data:image\/svg\+xml/i.test(value)) {
    return svgToPng(decodeSvgDataUrl(value));
  }

  if (
    /^data:image\/(png|jpe?g|gif|webp);base64,/i.test(value)
  ) {
    if ((value.length * 3) / 4 > MAX_IMAGE_SIZE) {
      throw new Error("Image exceeds the 8 MB limit.");
    }

    if (/^data:image\/webp/i.test(value)) {
      const response = await fetch(value);

      return convertBlobToPng(await response.blob());
    }

    return value;
  }

  if (value.startsWith("data:image/")) {
    throw new Error("Unsupported embedded image format.");
  }

  let url: URL;

  try {
    url = new URL(
      value,
      typeof window !== "undefined"
        ? window.location.href
        : "https://localhost/",
    );
  } catch {
    throw new Error("Invalid image URL.");
  }

  if (!["http:", "https:"].includes(url.protocol)) {
    throw new Error("Unsupported image URL.");
  }

  const response = await fetch(url.href, {
    mode: "cors",
    credentials: "omit",
  });

  if (!response.ok) {
    throw new Error(
      `Image request failed: ${response.status}`,
    );
  }

  const blob = await response.blob();

  if (blob.size > MAX_IMAGE_SIZE) {
    throw new Error("Image exceeds the 8 MB limit.");
  }

  if (!blob.type.startsWith("image/")) {
    throw new Error("URL does not contain an image.");
  }

  if (blob.type === "image/svg+xml") {
    return svgToPng(await blob.text());
  }

  if (
    ["image/png", "image/jpeg", "image/gif"].includes(blob.type)
  ) {
    return blobToDataUrl(blob);
  }

  return convertBlobToPng(blob);
}

/* -------------------------------------------------------------------------- */
/*                         MARKDOWN / CODE HELPERS                            */
/* -------------------------------------------------------------------------- */

function parseJson(code: string): Record<string, any> | null {
  try {
    const result = JSON.parse(code);

    return result &&
      typeof result === "object" &&
      !Array.isArray(result)
      ? result
      : null;
  } catch {
    return null;
  }
}

function normalizeColor(
  value: unknown,
  fallback: string,
): string {
  const color = String(value ?? fallback)
    .replace(/^#/, "")
    .toUpperCase();

  return /^[0-9A-F]{6}$/.test(color)
    ? color
    : fallback;
}

function highlightedRuns(
  code: string,
  language: string,
): TextRun[] {
  const source = cleanText(code);

  if (!source) {
    return [{ text: " " }];
  }

  let highlightedHtml = "";

  try {
    const lang = language.trim().toLowerCase();

    if (lang && hljs.getLanguage(lang)) {
      highlightedHtml = hljs.highlight(source, {
        language: lang,
        ignoreIllegals: true,
      }).value;
    } else {
      highlightedHtml = hljs.highlightAuto(source).value;
    }
  } catch {
    return [
      {
        text: source,
        options: {
          color: COLORS.codeText,
        },
      },
    ];
  }

  const output: TextRun[] = [];

  const tokenColor: Record<string, string> = {
    "hljs-keyword": COLORS.keyword,
    "hljs-selector-tag": COLORS.keyword,
    "hljs-literal": COLORS.literal,
    "hljs-number": COLORS.number,
    "hljs-string": COLORS.string,
    "hljs-doctag": COLORS.string,
    "hljs-regexp": COLORS.string,
    "hljs-title": COLORS.function,
    "hljs-title.function_": COLORS.function,
    "hljs-function": COLORS.function,
    "hljs-variable": COLORS.variable,
    "hljs-attr": COLORS.variable,
    "hljs-attribute": COLORS.variable,
    "hljs-built_in": COLORS.function,
    "hljs-type": COLORS.accent,
    "hljs-comment": COLORS.comment,
    "hljs-quote": COLORS.comment,
    "hljs-meta": COLORS.muted,
    "hljs-params": COLORS.codeText,
    "hljs-property": COLORS.variable,
    "hljs-operator": COLORS.codeText,
    "hljs-punctuation": COLORS.muted,
  };

  const parser = new DOMParser();

  const parsed = parser.parseFromString(
    `<div>${highlightedHtml}</div>`,
    "text/html",
  );

  const walk = (
    node: Node,
    inheritedColor = COLORS.codeText,
  ) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent ?? "";

      if (text) {
        output.push({
          text,
          options: {
            color: inheritedColor,
          },
        });
      }

      return;
    }

    if (node.nodeType !== Node.ELEMENT_NODE) return;

    const element = node as HTMLElement;
    const classes = Array.from(element.classList);

    let color = inheritedColor;

    for (const className of classes) {
      if (tokenColor[className]) {
        color = tokenColor[className];
        break;
      }
    }

    element.childNodes.forEach((child) => {
      walk(child, color);
    });
  };

  parsed.body.firstElementChild?.childNodes.forEach((node) => {
    walk(node);
  });

  return mergeRuns(output);
}

function decodeHtmlEntities(value: string): string {
  const element = document.createElement("textarea");

  element.innerHTML = value;

  return element.value;
}

function getTokenText(token: Token): string {
  if (token.type === "text") {
    const item = token as Tokens.Text;

    return item.tokens?.length
      ? getPlainText(item.tokens as Token[])
      : item.text;
  }

  if (token.type === "codespan") {
    return (token as Tokens.Codespan).text;
  }

  if (token.type === "image") {
    return (token as Tokens.Image).text;
  }

  if ("text" in token && typeof token.text === "string") {
    return token.text;
  }

  if ("tokens" in token && Array.isArray(token.tokens)) {
    return getPlainText(token.tokens as Token[]);
  }

  return "";
}

function getBlockText(token: Token): string {
  if (token.type === "heading") {
    return getPlainText(
      (token as Tokens.Heading).tokens,
    );
  }

  if (token.type === "paragraph") {
    return getPlainText(
      (token as Tokens.Paragraph).tokens,
    );
  }

  if (token.type === "blockquote") {
    return (token as Tokens.Blockquote).tokens
      .map(getBlockText)
      .join("\n");
  }

  return getTokenText(token);
}

function getFontSize(
  level: number,
  baseSize: number,
): number {
  const sizes: Record<number, number> = {
    1: FONT_SIZE.h1,
    2: FONT_SIZE.h2,
    3: FONT_SIZE.h3,
    4: FONT_SIZE.h4,
    5: FONT_SIZE.h5,
    6: FONT_SIZE.h6,
  };

  return sizes[level] ?? baseSize;
}

function estimateTextHeight(
  text: string,
  width: number,
  fontSize: number,
  lineHeight = LINE_HEIGHT,
): number {
  const lines = wrapPlainText(
    text,
    width,
    fontSize,
  );

  return Math.max(
    fontSize / 72,
    lines.length * (fontSize / 72) * lineHeight,
  );
}

function getImageDimensions(
  width: number,
  height: number,
  maxWidth: number,
  maxHeight: number,
): {
  width: number;
  height: number;
} {
  const ratio = Math.min(
    maxWidth / width,
    maxHeight / height,
    1,
  );

  return {
    width: width * ratio,
    height: height * ratio,
  };
}

function getListIndent(depth: number): number {
  return Math.min(depth * 0.28, 1.2);
}

function getBulletCharacter(ordered: boolean): string {
  return ordered ? "•" : "•";
}

function getTableAlignment(
  align?: string | null,
): Align {
  if (align === "center") return "center";
  if (align === "right") return "right";

  return "left";
}

function getTableCellText(
  cell: Tokens.TableCell,
): string {
  return getPlainText(cell.tokens as Token[]);
}

function createParagraphRuns(
  token: Token,
  fontSize: number,
): TextRun[] {
  if (
    "tokens" in token &&
    Array.isArray(token.tokens)
  ) {
    return inlineRuns(
      token.tokens as Token[],
    ).map((run) => ({
      ...run,
      options: {
        fontFace: FONT,
        fontSize,
        color: COLORS.body,
        ...run.options,
      },
    }));
  }

  return [
    {
      text: getTokenText(token),
      options: {
        fontFace: FONT,
        fontSize,
        color: COLORS.body,
      },
    },
  ];
}

function createCodeRuns(
  code: string,
  language: string,
): TextRun[] {
  return highlightedRuns(code, language);
}

function getCodeLanguage(info?: string): string {
  if (!info) return "";

  return info.trim().split(/\s+/)[0].toLowerCase();
}

/* -------------------------------------------------------------------------- */
/*                            MERMAID DIAGRAMS                                */
/* -------------------------------------------------------------------------- */

let mermaidInitialized = false;
let mermaidCounter = 0;

async function renderMermaidSvg(
  code: string,
): Promise<string> {
  if (!mermaidInitialized) {
    mermaid.initialize({
      startOnLoad: false,
      securityLevel: "strict",

      theme: "base",
      fontFamily: "Arial",

      // Consistent colors and typography.
      themeVariables: {
        fontFamily: "Arial",

        primaryColor: "#E0F2FE",
        primaryTextColor: "#263238",
        primaryBorderColor: "#087E8B",

        lineColor: "#64748B",

        secondaryColor: "#CCFBF1",
        tertiaryColor: "#F1F5F9",

        background: "#FFFFFF",
        mainBkg: "#E0F2FE",

        nodeBorder: "#087E8B",

        clusterBkg: "#F1F5F9",
        clusterBorder: "#CBD5E1",

        edgeLabelBackground: "#FFFFFF",
      },

      // Flowchart configuration.
      flowchart: {
        htmlLabels: false,
        curve: "basis",
        padding: 15,
        useMaxWidth: false,
        nodeSpacing: 40,
        rankSpacing: 50,
      },

      // Sequence diagram configuration.
      sequence: {
        useMaxWidth: false,
        diagramMarginX: 30,
        diagramMarginY: 20,
        actorMargin: 60,
        messageMargin: 35,
      },

      // Gantt chart configuration.
      gantt: {
        barHeight: 16,
        barGap: 4,
        topPadding: 30,
        leftPadding: 75,
        rightPadding: 25,
        gridLineStartPadding: 35,
        fontSize: 12,
        sectionFontSize: 13,
        numberSectionStyles: 4,
      },

      // Pie chart configuration.
      pie: {
        useWidth: 450,
      },
    });

    mermaidInitialized = true;
  }

  const id = `mermaid-diagram-${++mermaidCounter}`;

  try {
    const result = await mermaid.render(
      id,
      code.trim(),
    );

    return result.svg;
  } catch (error) {
    console.error(
      "Mermaid diagram parsing/rendering failed:",
      error,
    );

    throw error;
  }
}

function getHeadingColor(level: number): string {
  return level <= 2 ? COLORS.primary : COLORS.body;
}

function getHeadingSpacing(level: number): number {
  if (level === 1) return 0.28;
  if (level === 2) return 0.24;

  return 0.16;
}

function isHorizontalRule(token: Token): boolean {
  return token.type === "hr";
}

function isBlockToken(token: Token): boolean {
  return [
    "heading",
    "paragraph",
    "blockquote",
    "list",
    "table",
    "code",
    "hr",
    "image",
    "html",
  ].includes(token.type);
}

function getSafeFilename(value: string): string {
  return cleanText(value)
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 150);
}

function getTitleFromMarkdown(
  tokens: Token[],
  fallback: string,
): string {
  const heading = tokens.find(
    (token) => token.type === "heading",
  ) as Tokens.Heading | undefined;

  if (heading) {
    return (
      getPlainText(heading.tokens as Token[]).trim() ||
      fallback
    );
  }

  return fallback;
}

/* -------------------------------------------------------------------------- */
/*                       POWERPOINT GENERATION START                          */
/* -------------------------------------------------------------------------- */

async function convertMarkdownToPowerPoint(
  markdown: string,
  opts: MdToPptxOptions = {},
): Promise<Blob> {
  const pptxModule = await import("pptxgenjs");
  const pptx = new pptxModule.default();

  const source = cleanText(markdown);

  if (!source.trim()) {
    throw new Error("Markdown content cannot be empty.");
  }

  if (new Blob([source]).size > MAX_MARKDOWN_SIZE) {
    throw new Error(
      "Markdown content exceeds the 28 MB limit.",
    );
  }

  const baseFontSize = Math.max(
    10,
    Math.min(
      opts.fontSize ?? FONT_SIZE.body,
      32,
    ),
  );

  const parsedTokens = marked.lexer(source);

  const documentTitle =
    opts.title?.trim() ||
    getTitleFromMarkdown(
      parsedTokens,
      "Markdown Presentation",
    );

  pptx.layout = "LAYOUT_WIDE";
  pptx.author =
    opts.author || "Markdown PowerPoint Converter";
  pptx.subject = "Generated from Markdown";
  pptx.title = documentTitle;
  pptx.company = "Markdown Converter";

  pptx.theme = {
    headFontFace: FONT,
    bodyFontFace: FONT,
  };

  pptx.defineSlideMaster({
    title: "MASTER",
    background: {
      color: COLORS.background,
    },
    objects: [
      {
        line: {
          x: MARGIN_X,
          y: 0.82,
          w: CONTENT_W,
          h: 0,
          line: {
            color: COLORS.border,
            width: 0.8,
          },
        },
      },
      {
        text: {
          text: "MARKDOWN PRESENTATION",
          options: {
            x: MARGIN_X,
            y: 0.35,
            w: 5,
            h: 0.2,
            fontFace: FONT,
            fontSize: 8,
            bold: true,
            charSpacing: 1.5,
            color: COLORS.muted,
            margin: 0,
          },
        },
      },
    ],
    slideNumber: {
      x: SW - MARGIN_X - 0.5,
      y: FOOTER_Y,
      w: 0.5,
      h: 0.15,
      color: COLORS.muted,
      fontFace: FONT,
      fontSize: 8,
      align: "right",
      margin: 0,
    },
  });

  let slide: PptxSlide | null = null;
  let currentY = CONTENT_TOP;
  let slideCount = 0;
  let blockCount = 0;

  const createSlide = (): PptxSlide => {
    const newSlide = pptx.addSlide("MASTER");

    newSlide.background = {
      color: COLORS.background,
    };

    slideCount++;

    return newSlide;
  };

  const addFooter = (
    target: PptxSlide,
    text = documentTitle,
  ) => {
    target.addText(text, {
      x: MARGIN_X,
      y: FOOTER_Y,
      w: CONTENT_W - 0.8,
      h: 0.15,
      fontFace: FONT,
      fontSize: 8,
      color: COLORS.muted,
      breakLine: false,
    });
  };

  const addSectionTitle = (
    target: PptxSlide,
    title: string,
    subtitle?: string,
  ) => {
    target.addText(title, {
      x: MARGIN_X,
      y: 0.96,
      w: CONTENT_W,
      h: 0.32,
      fontFace: FONT,
      fontSize: 15,
      bold: true,
      color: COLORS.primary,
      margin: 0,
      breakLine: false,
    });

    if (subtitle) {
      target.addText(subtitle, {
        x: MARGIN_X,
        y: 1.3,
        w: CONTENT_W,
        h: 0.2,
        fontFace: FONT,
        fontSize: 9,
        color: COLORS.muted,
        margin: 0,
      });
    }
  };

  const ensureSlide = (): PptxSlide => {
    if (!slide) {
      slide = createSlide();
      currentY = CONTENT_TOP;
      addFooter(slide);
    }

    return slide;
  };

  const nextSlide = (
    title?: string,
    subtitle?: string,
  ): PptxSlide => {
    slide = createSlide();
    currentY = CONTENT_TOP;

    addFooter(slide);

    if (title) {
      addSectionTitle(slide, title, subtitle);

      currentY = subtitle ? 1.62 : 1.42;
    }

    return slide;
  };

  const ensureSpace = (
    requiredHeight: number,
    title?: string,
  ): PptxSlide => {
    const target = ensureSlide();

    if (
      currentY + requiredHeight >
      CONTENT_BOTTOM
    ) {
      return nextSlide(title);
    }

    return target;
  };

 const getRunsHeight = (
  runs: TextRun[],
  width: number,
  defaultSize: number,
): number => {
  const lines = wrapRuns(
    runs,
    width,
    defaultSize,
  );

  return lines.reduce((total, line) => {
    const lineFontSize = Math.max(
      defaultSize,
      ...line.map((run) =>
        Number(run.options?.fontSize ?? defaultSize),
      ),
    );

    return (
      total +
      (lineFontSize / 72) * LINE_HEIGHT +
      LINE_GAP
    );
  }, 0) - LINE_GAP;
};

const drawRuns = (
  runs: TextRun[],
  x: number,
  y: number,
  width: number,
  fontSize: number,
  options: Record<string, any> = {},
): number => {
  const target = ensureSlide();

  const normalizedRuns = runs.map((run) => ({
    text: run.text,
    options: {
      fontFace: FONT,
      fontSize,
      color: COLORS.body,
      ...run.options,
    },
  }));

  const lines = wrapRuns(
    normalizedRuns,
    width,
    fontSize,
  );

  let drawY = y;

  for (const line of lines) {
    const nonEmpty = line.filter(
      (run) => run.text,
    );

    const lineFontSize = Math.max(
      fontSize,
      ...nonEmpty.map((run) =>
        Number(run.options?.fontSize ?? fontSize),
      ),
    );

    const lineHeight =
      (lineFontSize / 72) * LINE_HEIGHT;

    const boxHeight = lineHeight + 0.025;

    if (!nonEmpty.length) {
      drawY += lineHeight + LINE_GAP;
      continue;
    }

    target.addText(nonEmpty, {
      fontFace: FONT,
      fontSize,
      color: COLORS.body,

      margin: 0,
      breakLine: false,
      valign: "middle",

      paraSpaceAfter: 0,

      ...options,

      // Keep text dimensions consistent in PowerPoint.
      fit: "none",
      wrap: false,

      // Preserve calculated shape dimensions.
      x,
      y: drawY,
      w: width,
      h: boxHeight,
    });

    drawY += lineHeight + LINE_GAP;
  }

  return Math.max(0, drawY - y - LINE_GAP);
};

const drawHeading = (
  token: Tokens.Heading,
) => {
  const level = token.depth;

  const size = getFontSize(
    level,
    baseFontSize,
  );

  const runs = inlineRuns(
    token.tokens as Token[],
  );

  const spacing = getHeadingSpacing(level);

  const lines = wrapRuns(
    runs,
    CONTENT_W,
    size,
  );

  const estimatedHeight =
    Math.max(1, lines.length) *
    ((size / 72) * LINE_HEIGHT + LINE_GAP);

  const bottomSpacing = 0.12;

  if (
    currentY + spacing + estimatedHeight + bottomSpacing >
    CONTENT_BOTTOM
  ) {
    nextSlide(
      getPlainText(token.tokens as Token[]),
    );
  }

  currentY += spacing;

  const height = drawRuns(
    runs,
    MARGIN_X,
    currentY,
    CONTENT_W,
    size,
    {
      bold: true,
      color: getHeadingColor(level),
      valign: "middle",
    },
  );

  currentY += Math.max(
    height,
    estimatedHeight,
  ) + bottomSpacing;
};

 const drawParagraph = (
  token: Tokens.Paragraph,
  extraIndent = 0,
) => {
  const runs = createParagraphRuns(
    token,
    baseFontSize,
  );

  const text = getPlainText(
    token.tokens as Token[],
  );

  const width = CONTENT_W - extraIndent;

  const lines = wrapRuns(
    runs,
    width,
    baseFontSize,
  );

  const estimatedHeight =
    Math.max(1, lines.length) *
    ((baseFontSize / 72) * LINE_HEIGHT + LINE_GAP);

  if (
    currentY + estimatedHeight + GAP >
    CONTENT_BOTTOM
  ) {
    nextSlide();
  }

  const height = drawRuns(
    runs,
    MARGIN_X + extraIndent,
    currentY,
    width,
    baseFontSize,
  );

  currentY += Math.max(
    height,
    estimatedHeight,
  ) + GAP;

  return text;
};

  const drawHorizontalRule = () => {
    const target = ensureSpace(0.3);

    target.addShape("line", {
      x: MARGIN_X,
      y: currentY + 0.08,
      w: CONTENT_W,
      h: 0,
      line: {
        color: COLORS.border,
        width: 1.2,
      },
    });

    currentY += 0.3;
  };

const drawCodeBlock = async (
  token: Tokens.Code,
) => {
  const code = cleanText(token.text);
  const language = getCodeLanguage(token.lang);

  /* ---------------------------------------------------------------------- */
  /*                         MERMAID DIAGRAM RENDERING                       */
  /* ---------------------------------------------------------------------- */

 if (language === "mermaid") {
  try {
    const svg = await renderMermaidSvg(code);

    const imageData = await svgToPng(svg);

    const image = new Image();

    await new Promise<void>((resolve, reject) => {
      image.onload = () => resolve();

      image.onerror = () => {
        reject(
          new Error("Unable to load Mermaid diagram image."),
        );
      };

      image.src = imageData;
    });

    const originalWidth = image.naturalWidth;
    const originalHeight = image.naturalHeight;

    if (!originalWidth || !originalHeight) {
      throw new Error("Invalid Mermaid image dimensions.");
    }

    // Calculate proportional dimensions.
    // Never independently stretch width and height.
    const aspectRatio = originalWidth / originalHeight;

    const maxWidth = CONTENT_W;
    const maxHeight = 4.8;

    const scale = Math.min(
  maxWidth / originalWidth,
  maxHeight / originalHeight,
  1,
);

    const diagramWidth = originalWidth * scale;
    const diagramHeight = originalHeight * scale;

    if (
      currentY + diagramHeight + GAP >
      CONTENT_BOTTOM
    ) {
      nextSlide("Diagram");
    }

    const target = ensureSlide();

    target.addImage({
      data: imageData,

      x:
        MARGIN_X +
        (CONTENT_W - diagramWidth) / 2,

      y: currentY,

      w: diagramWidth,
      h: diagramHeight,

      altText: "Mermaid diagram",
    });

    currentY += diagramHeight + GAP;

    return;
  } catch (error) {
    console.error(
      "Mermaid rendering failed:",
      error,
    );

    // Fall back to normal code rendering.
  }
}

  /* ---------------------------------------------------------------------- */
  /*                          NORMAL CODE BLOCKS                             */
  /* ---------------------------------------------------------------------- */

  const lines = code.split("\n");
  const chunks: string[][] = [];

  for (
    let index = 0;
    index < lines.length;
    index += MAX_CODE_LINES_PER_SLIDE
  ) {
    chunks.push(
      lines.slice(
        index,
        index + MAX_CODE_LINES_PER_SLIDE,
      ),
    );
  }

  for (
    let chunkIndex = 0;
    chunkIndex < chunks.length;
    chunkIndex++
  ) {
    const chunk = chunks[chunkIndex];

    const codeText = chunk.join("\n");
    const fontSize = FONT_SIZE.code;

    const lineHeight =
      (fontSize / 72) * 1.4;

    const headerHeight = language ? 0.29 : 0.14;
    const bottomPadding = 0.14;

    const boxHeight = Math.max(
      0.45,
      chunk.length * lineHeight +
        headerHeight +
        bottomPadding,
    );

    if (
      currentY + boxHeight + GAP >
      CONTENT_BOTTOM
    ) {
      nextSlide(
        chunkIndex > 0
          ? `${language || "Code"} (continued)`
          : undefined,
      );
    }

    const target = ensureSlide();

    target.addShape("roundRect", {
      x: MARGIN_X,
      y: currentY,
      w: CONTENT_W,
      h: boxHeight,

      fill: {
        color: COLORS.codeBackground,
      },

      line: {
        color: COLORS.border,
        width: 0.7,
      },
    });

    if (language) {
      target.addText(
        language.toUpperCase(),
        {
          x: MARGIN_X + 0.16,
          y: currentY + 0.08,
          w: CONTENT_W - 0.32,
          h: 0.14,

          fontFace: FONT,
          fontSize: 7,
          bold: true,
          color: COLORS.muted,
          charSpacing: 1,

          margin: 0,
          fit: "none",
        },
      );
    }

    const highlighted = createCodeRuns(
      codeText,
      language,
    );

    target.addText(highlighted, {
      x: MARGIN_X + 0.16,
      y: currentY + headerHeight,
      w: CONTENT_W - 0.32,
      h: boxHeight - headerHeight - bottomPadding,

      fontFace: MONO,
      fontSize,
      color: COLORS.codeText,

      margin: 0,
      valign: "top",
      breakLine: false,

      lineSpacingMultiple: 1.4,

      fit: "none",
      wrap: false,
    });

    currentY += boxHeight + GAP;
  }
};

    /* ------------------------------------------------------------------------ */
  /*                              BLOCKQUOTES                                 */
  /* ------------------------------------------------------------------------ */

 const drawBlockquote = (
  token: Tokens.Blockquote,
) => {
  const text = getBlockText(token);

  if (!text.trim()) return;

  const quoteWidth = CONTENT_W - 0.45;
  const quoteX = MARGIN_X + 0.22;

  const quoteRuns: TextRun[] = [
    {
      text,
      options: {
        italic: true,
        color: COLORS.body,
        fontFace: FONT,
        fontSize: baseFontSize,
      },
    },
  ];

  const lines = wrapRuns(
    quoteRuns,
    quoteWidth,
    baseFontSize,
  );

  const estimatedHeight =
    Math.max(1, lines.length) *
      ((baseFontSize / 72) * LINE_HEIGHT + LINE_GAP) +
    0.28;

  if (
    currentY + estimatedHeight + GAP >
    CONTENT_BOTTOM
  ) {
    nextSlide();
  }

  const target = ensureSlide();

  target.addShape("roundRect", {
    x: MARGIN_X,
    y: currentY,
    w: CONTENT_W,
    h: estimatedHeight,
    fill: {
      color: COLORS.quoteBackground,
    },
    line: {
      color: COLORS.border,
      width: 0.6,
    },
  });

  target.addShape("rect", {
    x: MARGIN_X,
    y: currentY,
    w: 0.055,
    h: estimatedHeight,
    line: {
      color: COLORS.primary,
      transparency: 100,
    },
    fill: {
      color: COLORS.primary,
    },
  });

  const textHeight = drawRuns(
    quoteRuns,
    quoteX,
    currentY + 0.12,
    quoteWidth,
    baseFontSize,
    {
      italic: true,
      color: COLORS.body,
      valign: "middle",
    },
  );

  currentY += Math.max(
    estimatedHeight,
    textHeight + 0.28,
  ) + GAP;
};

  /* ------------------------------------------------------------------------ */
  /*                              LISTS                                       */
  /* ------------------------------------------------------------------------ */

 const drawList = async (
  token: Tokens.List,
  depth = 0,
): Promise<void> => {
  const indent = getListIndent(depth);

  for (
    let index = 0;
    index < token.items.length;
    index++
  ) {
    const item = token.items[index];

    const prefix = token.ordered
      ? `${Number(token.start ?? 1) + index}.`
      : getBulletCharacter(false);

    const itemTokens = item.tokens as Token[];

    const firstParagraph = itemTokens.find(
      (child) =>
        child.type === "text" ||
        child.type === "paragraph",
    );

    const itemRuns = firstParagraph
      ? inlineRuns(
          (firstParagraph as any).tokens ??
            itemTokens,
        )
      : [
          {
            text: getPlainText(itemTokens),
          },
        ];

    const itemText = getPlainText(itemTokens);

    const textWidth =
      CONTENT_W - indent - 0.35;

    const lines = wrapRuns(
      itemRuns,
      textWidth,
      baseFontSize,
    );

    const estimatedHeight =
      Math.max(1, lines.length) *
      (
        (baseFontSize / 72) * LINE_HEIGHT +
        LINE_GAP
      );

    if (
      currentY + estimatedHeight + 0.06 >
      CONTENT_BOTTOM
    ) {
      nextSlide();
    }

    const target = ensureSlide();

    const itemY = currentY;

    // Bullet or numbering.
    target.addText(prefix, {
      x: MARGIN_X + indent,
      y: itemY,
      w: 0.3,
      h: Math.max(estimatedHeight, 0.25),
      fontFace: FONT,
      fontSize: baseFontSize,
      bold: true,
      color: COLORS.primary,
      margin: 0,
      valign: "top",
      fit: "none",
      wrap: false,
    });

    // Main list item.
    const renderedHeight = drawRuns(
      itemRuns,
      MARGIN_X + indent + 0.35,
      itemY,
      textWidth,
      baseFontSize,
      {
        valign: "top",
      },
    );

    currentY += Math.max(
      renderedHeight,
      estimatedHeight,
    ) + 0.06;

    // Render nested content.
    let firstParagraphRendered = false;

    for (const child of itemTokens) {
      if (
        child.type === "text" ||
        child.type === "paragraph"
      ) {
        if (!firstParagraphRendered) {
          firstParagraphRendered = true;
          continue;
        }

        if (child.type === "paragraph") {
          drawParagraph(
            child as Tokens.Paragraph,
            indent + 0.35,
          );
        } else {
          drawTextBlock(
            getTokenText(child),
            {
              fontSize: baseFontSize,
            },
          );
        }
      } else if (child.type === "list") {
        await drawList(
          child as Tokens.List,
          depth + 1,
        );
      } else {
        await renderToken(child);
      }
    }

    if (!itemText.trim()) {
      currentY += 0.04;
    }
  }

  currentY += 0.08;
};

  /* ------------------------------------------------------------------------ */
  /*                              TABLES                                      */
  /* ------------------------------------------------------------------------ */

 const drawTable = async (
  token: Tokens.Table,
): Promise<void> => {
  const headers = token.header;
  const rows = token.rows;

  const columnCount = Math.min(
    headers.length,
    MAX_TABLE_COLUMNS,
  );

  if (!columnCount) return;

  const tableRows = [
    headers.slice(0, columnCount),
    ...rows
      .slice(0, MAX_TABLE_ROWS)
      .map((row) => row.slice(0, columnCount)),
  ];

  const columnWidths = Array(columnCount).fill(
    CONTENT_W / columnCount,
  );

  const rowHeights: number[] = [];

  const pptxRows: any[] = [];

  for (
    let rowIndex = 0;
    rowIndex < tableRows.length;
    rowIndex++
  ) {
    const row = tableRows[rowIndex];

    let maxLines = 1;

    const cells = row.map((cell, columnIndex) => {
      const text = getTableCellText(cell);

      const fontSize = rowIndex === 0 ? 10 : 9;

      const cellWidth =
        columnWidths[columnIndex] - 0.2;

      const cellLines = wrapPlainText(
        text || " ",
        cellWidth,
        fontSize,
      );

      maxLines = Math.max(
        maxLines,
        cellLines.length,
      );

      return {
        text: text || " ",
        options: {
          fontFace: FONT,
          fontSize,
          bold: rowIndex === 0,
          color:
            rowIndex === 0
              ? "FFFFFF"
              : COLORS.body,
          align: getTableAlignment(cell.align),
          valign: "middle",
          margin: 0.08,
          breakLine: false,
          fill:
            rowIndex === 0
              ? {
                  color: COLORS.primary,
                }
              : rowIndex % 2 === 0
                ? {
                    color: COLORS.tableAlternate,
                  }
                : {
                    color: COLORS.background,
                  },
        },
      };
    });

    const rowHeight = Math.max(
      0.42,
      maxLines * (9 / 72) * 1.4 + 0.2,
    );

    rowHeights.push(rowHeight);
    pptxRows.push(cells);
  }

  const tableHeight = rowHeights.reduce(
    (total, height) => total + height,
    0,
  );

  const availableHeight =
    CONTENT_BOTTOM - currentY;

  // Move the complete table to a new slide when necessary.
  if (tableHeight > availableHeight) {
    nextSlide("Table");
  }

  const target = ensureSlide();

  target.addTable(pptxRows, {
    x: MARGIN_X,
    y: currentY,
    w: CONTENT_W,
    colW: columnWidths,

    border: {
      type: "solid",
      color: COLORS.border,
      pt: 0.6,
    },

    fill: {
      color: COLORS.background,
    },

    color: COLORS.body,
    fontFace: FONT,
    fontSize: 9,
    margin: 0.08,
    valign: "middle",

    rowH: rowHeights,

    // Prevent PptxGenJS from independently creating slides.
    autoPage: false,
  });

  currentY += tableHeight + GAP;
};

  /* ------------------------------------------------------------------------ */
  /*                              IMAGES                                      */
  /* ------------------------------------------------------------------------ */

  const drawImage = async (
    token: Tokens.Image,
  ): Promise<void> => {
    const source = token.href;

    if (!source) return;

    try {
      const imageData = await resolveImage(source);

      const image = new Image();

      await new Promise<void>((resolve, reject) => {
        image.onload = () => resolve();

        image.onerror = () => {
          reject(
            new Error(
              "Unable to determine image dimensions.",
            ),
          );
        };

        image.src = imageData;
      });

      const dimensions = getImageDimensions(
        image.naturalWidth,
        image.naturalHeight,
        CONTENT_W,
        4.5,
      );

      if (
        currentY + dimensions.height + 0.35 >
        CONTENT_BOTTOM
      ) {
        nextSlide(token.text || "Image");
      }

      const target = ensureSlide();

      target.addImage({
        data: imageData,
        x:
          MARGIN_X +
          (CONTENT_W - dimensions.width) / 2,
        y: currentY,
        w: dimensions.width,
        h: dimensions.height,
        altText: token.text || "Markdown image",
      });

      currentY += dimensions.height + 0.18;

      if (token.text) {
        target.addText(token.text, {
          x: MARGIN_X,
          y: currentY,
          w: CONTENT_W,
          h: 0.2,
          fontFace: FONT,
          fontSize: FONT_SIZE.small,
          italic: true,
          align: "center",
          color: COLORS.muted,
          margin: 0,
        });

        currentY += 0.25;
      }
    } catch (error) {
      console.warn(
        "Unable to add Markdown image:",
        error,
      );

      const target = ensureSlide();

      target.addText(
        `[Image unavailable: ${token.text || source}]`,
        {
          x: MARGIN_X,
          y: currentY,
          w: CONTENT_W,
          h: 0.3,
          fontFace: FONT,
          fontSize: 10,
          italic: true,
          color: COLORS.muted,
          margin: 0,
        },
      );

      currentY += 0.4;
    }
  };

  /* ------------------------------------------------------------------------ */
  /*                             HTML BLOCKS                                  */
  /* ------------------------------------------------------------------------ */

  const drawHtmlBlock = (
    token: Tokens.HTML,
  ) => {
    const text = cleanText(
      token.text.replace(/<[^>]*>/g, " "),
    ).trim();

    if (!text) return;

    const runs = [
      {
        text: decodeHtmlEntities(text),
        options: {
          fontFace: FONT,
          fontSize: baseFontSize,
          color: COLORS.body,
        },
      },
    ];

    const height = estimateTextHeight(
      text,
      CONTENT_W,
      baseFontSize,
    );

    if (currentY + height > CONTENT_BOTTOM) {
      nextSlide();
    }

    const drawnHeight = drawRuns(
      runs,
      MARGIN_X,
      currentY,
      CONTENT_W,
      baseFontSize,
    );

    currentY += drawnHeight + GAP;
  };

  /* ------------------------------------------------------------------------ */
  /*                            TEXT BLOCKS                                   */
  /* ------------------------------------------------------------------------ */

const drawTextBlock = (
  text: string,
  options: Record<string, any> = {},
) => {
  if (!text.trim()) return;

  const size = Number(
    options.fontSize ?? baseFontSize,
  );

  const runs: TextRun[] = [
    {
      text,
      options: {
        fontFace: FONT,
        fontSize: size,
        color: COLORS.body,
        ...options,
      },
    },
  ];

  const lines = wrapRuns(
    runs,
    CONTENT_W,
    size,
  );

  const estimatedHeight =
    Math.max(1, lines.length) *
    (
      (size / 72) * LINE_HEIGHT +
      LINE_GAP
    );

  if (
    currentY + estimatedHeight + GAP >
    CONTENT_BOTTOM
  ) {
    nextSlide();
  }

  const renderedHeight = drawRuns(
    runs,
    MARGIN_X,
    currentY,
    CONTENT_W,
    size,
    {
      ...options,
      valign: "top",
      fit: "none",
      wrap: false,
    },
  );

  currentY += Math.max(
    renderedHeight,
    estimatedHeight,
  ) + GAP;
};
  /* ------------------------------------------------------------------------ */
  /*                         MARKDOWN TOKEN RENDERER                          */
  /* ------------------------------------------------------------------------ */

  const renderToken = async (
    token: Token,
  ): Promise<void> => {
    blockCount++;

    switch (token.type) {
      case "space":
        currentY += 0.08;
        break;

      case "heading":
        drawHeading(token as Tokens.Heading);
        break;

      case "paragraph":
        drawParagraph(token as Tokens.Paragraph);
        break;

      case "text":
        drawTextBlock(getTokenText(token));
        break;

      case "code":
  await drawCodeBlock(token as Tokens.Code);
  break;

      case "blockquote":
        drawBlockquote(token as Tokens.Blockquote);
        break;

      case "list":
        await drawList(token as Tokens.List);
        break;

      case "table":
        await drawTable(token as Tokens.Table);
        break;

      case "image":
        await drawImage(token as Tokens.Image);
        break;

      case "hr":
        drawHorizontalRule();
        break;

      case "html":
        drawHtmlBlock(token as Tokens.HTML);
        break;

      case "def":
        break;

      default:
        if (
          "tokens" in token &&
          Array.isArray(token.tokens)
        ) {
          for (const child of token.tokens as Token[]) {
            await renderToken(child);
          }
        } else {
          const text = getTokenText(token);

          if (text.trim()) {
            drawTextBlock(text);
          }
        }
    }
  };

  /* ------------------------------------------------------------------------ */
  /*                         RENDER MARKDOWN DOCUMENT                         */
  /* ------------------------------------------------------------------------ */

  for (const token of parsedTokens) {
    await renderToken(token);
  }

  // Ensure at least one slide exists.
  if (!slide || slideCount === 0) {
    slide = createSlide();

    addFooter(slide);

    drawTextBlock(
      "No readable Markdown content was found.",
      {
        fontSize: baseFontSize,
        color: COLORS.muted,
      },
    );
  }

  // Document metadata.
  pptx.title = documentTitle;
  pptx.author =
    opts.author || "Markdown PowerPoint Converter";
  pptx.subject = "Generated from Markdown";
  pptx.company = "Markdown Converter";

  /* ------------------------------------------------------------------------ */
  /*                          GENERATE POWERPOINT BLOB                        */
  /* ------------------------------------------------------------------------ */

  const output = await pptx.write({
    outputType: "blob",
  });

  if (!(output instanceof Blob)) {
    throw new Error(
      "PowerPoint generation failed: Invalid output.",
    );
  }

  if (output.size === 0) {
    throw new Error(
      "PowerPoint generation failed: Empty file.",
    );
  }

  return output;
};

/* -------------------------------------------------------------------------- */
/*                             PUBLIC EXPORTS                                 */
/* -------------------------------------------------------------------------- */

/**
 * Generates a PowerPoint Blob from Markdown.
 *
 * Use this function when another component handles downloading.
 */
export async function mdToPptx(
  markdown: string,
  options: MdToPptxOptions = {},
): Promise<Blob> {
  try {
    return await convertMarkdownToPowerPoint(
      markdown,
      options,
    );
  } catch (error) {
    console.error(
      "Markdown to PowerPoint conversion failed:",
      error,
    );

    throw new Error(
      error instanceof Error
        ? error.message
        : "Unable to convert Markdown to PowerPoint.",
    );
  }
}

/**
 * Generates and directly downloads a PowerPoint file.
 *
 * Use this only when you want this function to handle
 * the download itself.
 */
export async function downloadMarkdownAsPptx(
  markdown: string,
  options: MdToPptxOptions = {},
): Promise<void> {
  const blob = await mdToPptx(
    markdown,
    options,
  );

  const filename = getSafeFilename(
    options.title || "markdown-presentation",
  );

  const url = URL.createObjectURL(blob);

  try {
    const anchor = document.createElement("a");

    anchor.href = url;

    anchor.download = `${
      filename || "markdown-presentation"
    }.pptx`;

    document.body.appendChild(anchor);

    anchor.click();

    anchor.remove();
  } finally {
    setTimeout(() => {
      URL.revokeObjectURL(url);
    }, 1000);
  }
}