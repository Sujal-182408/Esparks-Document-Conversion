import { marked, type Token, type Tokens } from "marked";
import type PptxGenJS from "pptxgenjs";

const MAX_FILE_SIZE = 20 * 1024 * 1024;
const SLIDE_WIDTH = 13.333;
const CONTENT_BOTTOM = 6.85;
const BODY_X = 0.78;
const BODY_WIDTH = 11.8;

type TextRun = PptxGenJS.TextProps;
type TextRunOptions = NonNullable<TextRun["options"]>;

function isListToken(token: Token): token is Tokens.List {
  return (
    token.type === "list" &&
    "items" in token &&
    Array.isArray(token.items) &&
    "ordered" in token &&
    typeof token.ordered === "boolean"
  );
}

function isTableToken(token: Token): token is Tokens.Table {
  return (
    token.type === "table" &&
    "header" in token &&
    Array.isArray(token.header) &&
    "rows" in token &&
    Array.isArray(token.rows)
  );
}

function isImageToken(token: Token): token is Tokens.Image {
  return (
    token.type === "image" &&
    "href" in token &&
    typeof token.href === "string" &&
    "text" in token &&
    typeof token.text === "string"
  );
}

function plainText(tokens: Token[]): string {
  return tokens
    .map((token) => {
      if ("tokens" in token && Array.isArray(token.tokens)) {
        return plainText(token.tokens);
      }

      if ("text" in token && typeof token.text === "string") {
        return token.text;
      }

      return "";
    })
    .join("");
}

function inlineRuns(
  tokens: Token[],
  inherited: TextRunOptions = {}
): TextRun[] {
  const runs: TextRun[] = [];

  const add = (text: string, options: TextRunOptions = {}) => {
    if (text) {
      runs.push({
        text,
        options: { ...inherited, ...options },
      });
    }
  };

  for (const token of tokens) {
    switch (token.type) {
      case "text":
      case "escape":
        if ("tokens" in token && Array.isArray(token.tokens)) {
          runs.push(...inlineRuns(token.tokens, inherited));
        } else if ("text" in token) {
          add(token.text);
        }
        break;

      case "strong":
        runs.push(
          ...inlineRuns(token.tokens ?? [], { ...inherited, bold: true })
        );
        break;

      case "em":
        runs.push(
          ...inlineRuns(token.tokens ?? [], { ...inherited, italic: true })
        );
        break;

      case "del":
        runs.push(
          ...inlineRuns(token.tokens ?? [], { ...inherited, strike: true })
        );
        break;

      case "codespan":
        add(token.text, {
          fontFace: "Courier New",
          color: "B45309",
          breakLine: false,
        });
        break;

      case "link":
        runs.push(
          ...inlineRuns(token.tokens ?? [], {
            ...inherited,
            color: "0369A1",
            underline: { color: "0369A1" },
            hyperlink: { url: token.href },
          })
        );
        break;

      case "image":
        add(`[Image: ${token.text || token.href}]`, {
          italic: true,
          color: "0369A1",
          underline: { color: "0369A1" },
          hyperlink: { url: token.href },
        });
        break;

      case "br":
        runs.push({ text: "", options: { ...inherited, breakLine: true } });
        break;

      default:
        if ("tokens" in token && Array.isArray(token.tokens)) {
          runs.push(...inlineRuns(token.tokens, inherited));
        } else if ("text" in token && typeof token.text === "string") {
          add(token.text);
        }
    }
  }

  return runs;
}

function splitRuns(runs: TextRun[], maxCharacters: number): TextRun[][] {
  const chunks: TextRun[][] = [[]];
  let chunkLength = 0;

  for (const run of runs) {
    const text = run.text ?? "";
    if (!text && run.options?.breakLine) {
      chunks[chunks.length - 1].push(run);
      continue;
    }

    let offset = 0;

    while (offset < text.length) {
      const available = maxCharacters - chunkLength;
      if (available === 0) {
        chunks.push([]);
        chunkLength = 0;
      }

      const length = Math.min(maxCharacters - chunkLength, text.length - offset);
      chunks[chunks.length - 1].push({
        text: text.slice(offset, offset + length),
        options: run.options,
      });
      offset += length;
      chunkLength += length;
    }
  }

  return chunks.filter((chunk) =>
    chunk.some((run) => run.text || run.options?.breakLine)
  );
}

export async function convertMarkdownToPowerPoint(
  markdown: string
): Promise<Blob> {
  if (!markdown.trim()) {
    throw new Error("The Markdown file is empty.");
  }

  if (new Blob([markdown]).size > MAX_FILE_SIZE) {
    throw new Error("The maximum file size is 20 MB.");
  }

  const { default: PptxGenJS } = await import("pptxgenjs");
  const presentation = new PptxGenJS();
  presentation.layout = "LAYOUT_WIDE";
  presentation.author = "eSparks Document Conversion";
  presentation.subject = "Converted from Markdown";
  presentation.title = "Markdown Presentation";
  presentation.theme = {
    headFontFace: "Arial",
    bodyFontFace: "Arial",
  };

  let slideNumber = 0;
  let slideTitle = "";
  let slide = presentation.addSlide();
  let y = 1.55;
  let hasContent = false;

  const addSlide = (title: string) => {
    if (slideNumber > 0) {
      slide = presentation.addSlide();
    }

    slideNumber += 1;
    slideTitle = title || "Untitled section";
    slide.background = { color: "FFFFFF" };
    slide.addShape(presentation.ShapeType.rect, {
      x: 0,
      y: 0,
      w: SLIDE_WIDTH,
      h: 0.14,
      line: { color: "06B6D4", transparency: 100 },
      fill: { color: "06B6D4" },
    });
    slide.addText(slideTitle, {
      x: 0.72,
      y: 0.48,
      w: 11.9,
      h: 0.74,
      fontFace: "Arial",
      fontSize: 28,
      bold: true,
      color: "0F172A",
      fit: "shrink",
      margin: 0,
      valign: "middle",
    });
    slide.addShape(presentation.ShapeType.line, {
      x: 0.72,
      y: 1.32,
      w: 11.9,
      h: 0,
      line: { color: "E2E8F0", width: 1 },
    });
    slide.addText(String(slideNumber), {
      x: 12.05,
      y: 7.08,
      w: 0.55,
      h: 0.2,
      fontFace: "Arial",
      fontSize: 9,
      color: "64748B",
      align: "right",
      margin: 0,
    });
    y = 1.55;
    hasContent = false;
  };

  const addRichContent = (
    runs: TextRun[],
    options: {
      fontSize?: number;
      color?: string;
      fontFace?: string;
      bold?: boolean;
      italic?: boolean;
      bulletLevel?: number;
      quote?: boolean;
    } = {}
  ) => {
    const allText = runs.map((run) => run.text ?? "").join("");
    if (!allText.trim()) return;

    if (slideNumber === 0) addSlide("Markdown Presentation");

    const fontSize = options.fontSize ?? 17;
    const chunks = splitRuns(runs, 420);

    for (const chunk of chunks) {
      const chunkText = chunk.map((run) => run.text ?? "").join("");
      const lineCount = Math.max(
        1,
        chunkText.split("\n").reduce(
          (total, line) => total + Math.max(1, Math.ceil(line.length / 102)),
          0
        )
      );
      const height = Math.min(lineCount * (fontSize / 72) * 1.45 + 0.12, 2.4);
      const indent = (options.bulletLevel ?? 0) * 0.28;

      if (y + height > CONTENT_BOTTOM && hasContent) {
        addSlide(`${slideTitle} (continued)`);
      }

      if (options.quote) {
        slide.addShape(presentation.ShapeType.rect, {
          x: BODY_X + indent,
          y: y + 0.02,
          w: 0.05,
          h: Math.max(0.25, height - 0.04),
          line: { color: "06B6D4", transparency: 100 },
          fill: { color: "06B6D4" },
        });
      }

      slide.addText(chunk, {
        x: BODY_X + indent + (options.quote ? 0.18 : 0),
        y,
        w: BODY_WIDTH - indent - (options.quote ? 0.18 : 0),
        h: height,
        fontFace: options.fontFace ?? "Arial",
        fontSize,
        color: options.color ?? "334155",
        bold: options.bold,
        italic: options.italic,
        fit: "shrink",
        valign: "top",
        breakLine: false,
        margin: 0.04,
        paraSpaceAfter: 6,
        bullet: options.bulletLevel !== undefined
          ? { indent: fontSize * 1.05 }
          : undefined,
      });
      y += height + 0.14;
      hasContent = true;
    }
  };

  const addImageBlock = async (token: Tokens.Image) => {
    if (!/^data:image\/(png|jpe?g|gif);base64,/i.test(token.href)) {
      throw new Error(
        "PowerPoint conversion requires images to be embedded as PNG, JPEG, or GIF data URIs."
      );
    }

    const dimensions = await new Promise<{ width: number; height: number } | null>(
      (resolve) => {
        const image = new Image();
        image.onload = () =>
          resolve({ width: image.naturalWidth, height: image.naturalHeight });
        image.onerror = () => resolve(null);
        image.src = token.href;
      }
    );

    if (!dimensions || dimensions.width === 0 || dimensions.height === 0) {
      throw new Error("An embedded Markdown image could not be decoded.");
    }

    const scale = Math.min(
      BODY_WIDTH / dimensions.width,
      3.8 / dimensions.height,
      1
    );
    const width = dimensions.width * scale;
    const height = dimensions.height * scale;

    if (y + height > CONTENT_BOTTOM && hasContent) {
      addSlide(`${slideTitle} (continued)`);
    }

    slide.addImage({
      data: token.href,
      x: BODY_X + (BODY_WIDTH - width) / 2,
      y,
      w: width,
      h: height,
      altText: token.text || "Markdown image",
    });
    y += height + 0.18;
    hasContent = true;
  };

  const addTextBlock = async (
    tokens: Token[],
    options: Parameters<typeof addRichContent>[1] = {}
  ) => {
    let textTokens: Token[] = [];

    const flushText = () => {
      if (textTokens.length) addRichContent(inlineRuns(textTokens), options);
      textTokens = [];
    };

    for (const token of tokens) {
      if (isImageToken(token)) {
        flushText();
        await addImageBlock(token);
      } else {
        textTokens.push(token);
      }
    }

    flushText();
  };

  const addList = async (token: Tokens.List, level = 0) => {
    for (const [index, item] of token.items.entries()) {
      const prefix = token.ordered
        ? `${Number(token.start || 1) + index}. `
        : "";

      if (!item.tokens.some(isImageToken)) {
        const itemRuns = inlineRuns(item.tokens);
        if (prefix) itemRuns.unshift({ text: prefix, options: { bold: true } });
        addRichContent(itemRuns, {
          bulletLevel: token.ordered ? undefined : level,
        });
      } else {
        let textTokens: Token[] = [];
        let firstBlock = true;
        const flushText = () => {
          if (!textTokens.length) return;
          const runs = inlineRuns(textTokens);
          if (prefix && firstBlock) {
            runs.unshift({ text: prefix, options: { bold: true } });
          }
          addRichContent(runs, {
            bulletLevel: token.ordered ? undefined : level,
          });
          firstBlock = false;
          textTokens = [];
        };

        for (const child of item.tokens) {
          if (isImageToken(child)) {
            flushText();
            await addImageBlock(child);
          } else {
            textTokens.push(child);
          }
        }
        flushText();
        if (prefix && firstBlock) {
          addRichContent([{ text: prefix, options: { bold: true } }]);
        }
      }

      for (const child of item.tokens) {
        if (isListToken(child)) {
          await addList(child, level + 1);
        }
      }
    }
  };

  const addTable = (token: Tokens.Table) => {
    const tableCells = [token.header, ...token.rows].flat();
    if (tableCells.some((cell) => cell.tokens.some(isImageToken))) {
      throw new Error(
        "PowerPoint conversion does not support images inside Markdown tables."
      );
    }

    const columnCount = Math.max(1, token.header.length);
    const columnWidths = Array.from(
      { length: columnCount },
      () => BODY_WIDTH / columnCount
    );
    const header: PptxGenJS.TableRow = token.header.map((cell) => ({
      text: cell.text,
      options: {
        bold: true,
        color: "0F172A",
        fill: { color: "CFFAFE" },
        align: cell.align ?? "left",
        valign: "middle" as const,
      },
    }));
    const rows: PptxGenJS.TableRow[] = token.rows.map((row) =>
      row.map((cell, index) => ({
        text: cell.text,
        options: {
          color: "334155",
          fill: { color: "FFFFFF" },
          align: token.align[index] ?? "left",
          valign: "middle" as const,
        },
      }))
    );
    const rowHeight = (row: PptxGenJS.TableRow) =>
      Math.min(
        2.4,
        Math.max(
          0.44,
          ...row.map((cell, index) => {
            const charactersPerLine = Math.max(
              12,
              Math.floor(columnWidths[index] * 10)
            );
            const text = typeof cell.text === "string" ? cell.text : "";
            return (
              Math.ceil(text.length / charactersPerLine) * 0.2 + 0.18
            );
          })
        )
      );
    const headerHeight = rowHeight(header);
    const rowHeights = rows.map(rowHeight);
    let rowIndex = 0;
    let renderedHeader = false;

    while (!renderedHeader || rowIndex < rows.length) {
      const pageRows: PptxGenJS.TableRow[] = [header];
      const pageRowHeights = [headerHeight];
      let tableHeight = headerHeight;
      const startRowIndex = rowIndex;

      while (
        rowIndex < rows.length &&
        tableHeight + rowHeights[rowIndex] <= CONTENT_BOTTOM - y
      ) {
        pageRows.push(rows[rowIndex]);
        pageRowHeights.push(rowHeights[rowIndex]);
        tableHeight += rowHeights[rowIndex];
        rowIndex += 1;
      }

      if (rowIndex === startRowIndex && rows.length > 0 && hasContent) {
        addSlide(`${slideTitle} (continued)`);
        continue;
      }

      if (rowIndex === startRowIndex && rows.length > 0) {
        pageRows.push(rows[rowIndex]);
        pageRowHeights.push(rowHeights[rowIndex]);
        tableHeight += rowHeights[rowIndex];
        rowIndex += 1;
      }

      slide.addTable(pageRows, {
        x: BODY_X,
        y,
        w: BODY_WIDTH,
        h: tableHeight,
        colW: columnWidths,
        rowH: pageRowHeights,
        border: { type: "solid", color: "CBD5E1", pt: 0.7 },
        fontFace: "Arial",
        fontSize: 12,
        margin: 0.08,
        color: "334155",
      });
      y += tableHeight + 0.18;
      hasContent = true;
      renderedHeader = true;

      if (rowIndex < rows.length) {
        addSlide(`${slideTitle} (continued)`);
      }
    }
  };

  for (const token of marked.lexer(markdown)) {
    switch (token.type) {
      case "heading":
        if (token.depth <= 2 || slideNumber === 0) {
          addSlide(token.text);
        } else {
          addRichContent(inlineRuns(token.tokens ?? []), {
            fontSize: token.depth === 3 ? 21 : 18,
            bold: true,
            color: "0F172A",
          });
        }
        break;

      case "paragraph":
      case "text":
        if ("tokens" in token && token.tokens) {
          await addTextBlock(token.tokens);
        } else {
          addRichContent([{ text: plainText([token]) }]);
        }
        break;

      case "list":
        if (isListToken(token)) await addList(token);
        break;

      case "blockquote":
        await addTextBlock(token.tokens ?? [], {
          quote: true,
          italic: true,
          color: "475569",
        });
        break;

      case "code":
        addRichContent([{ text: token.text, options: { breakLine: false } }], {
          fontFace: "Courier New",
          fontSize: 13,
          color: "0F172A",
        });
        break;

      case "table":
        if (isTableToken(token)) addTable(token);
        break;

      case "hr":
        if (slideNumber === 0) addSlide("Markdown Presentation");
        if (y + 0.2 > CONTENT_BOTTOM) {
          addSlide(`${slideTitle} (continued)`);
        }
        slide.addShape(presentation.ShapeType.line, {
          x: BODY_X,
          y: y + 0.08,
          w: BODY_WIDTH,
          h: 0,
          line: { color: "CBD5E1", width: 1 },
        });
        y += 0.28;
        hasContent = true;
        break;

      case "space":
      case "def":
        break;

      default:
        if ("tokens" in token && Array.isArray(token.tokens)) {
          await addTextBlock(token.tokens);
        } else if ("text" in token && typeof token.text === "string") {
          addRichContent([{ text: token.text }]);
        }
    }
  }

  if (slideNumber === 0) {
    addSlide("Markdown Presentation");
  }

  if (!hasContent) {
    addRichContent([{ text: "No readable content was found in the Markdown file." }]);
  }

  const output = await presentation.write({ outputType: "blob" });
  if (!(output instanceof Blob) || output.size === 0) {
    throw new Error("Could not generate the PowerPoint presentation.");
  }

  return output;
}
