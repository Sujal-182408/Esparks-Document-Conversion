
import {
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  TextRun,
  ExternalHyperlink,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  ImageRun,
  ShadingType,
} from "docx";
import { marked, type Token, type Tokens } from "marked";

type InlineToken =
  | Tokens.Text
  | Tokens.Strong
  | Tokens.Em
  | Tokens.Codespan
  | Tokens.Link
  | Tokens.Del
  | Tokens.Br
  | Tokens.Escape
  | Tokens.Image;

function imageRunFromDataUri(src: string, alt: string): ImageRun | null {
  const match = src.match(
    /^data:image\/(png|jpe?g|gif|bmp);base64,([A-Za-z0-9+/=]+)$/i
  );
  if (!match) return null;

  const imageType = match[1].toLowerCase();
  const type = imageType === "jpeg" ? "jpg" : imageType;
  const binary = atob(match[2]);
  const data = Uint8Array.from(binary, (character) => character.charCodeAt(0));
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  let width = 400;
  let height = 300;

  if (type === "png" && data.length >= 24) {
    width = view.getUint32(16);
    height = view.getUint32(20);
  } else if (type === "gif" && data.length >= 10) {
    width = view.getUint16(6, true);
    height = view.getUint16(8, true);
  } else if (type === "bmp" && data.length >= 26) {
    width = Math.abs(view.getInt32(18, true));
    height = Math.abs(view.getInt32(22, true));
  } else if (type === "jpg") {
    let offset = 2;
    while (offset + 9 < data.length) {
      if (data[offset] !== 0xff) {
        offset += 1;
        continue;
      }

      const marker = data[offset + 1];
      const segmentLength = view.getUint16(offset + 2);
      if (
        [0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf]
          .includes(marker)
      ) {
        height = view.getUint16(offset + 5);
        width = view.getUint16(offset + 7);
        break;
      }
      if (segmentLength < 2) break;
      offset += 2 + segmentLength;
    }
  }

  const scale = Math.min(
    1,
    450 / Math.max(1, width),
    600 / Math.max(1, height)
  );
  return new ImageRun({
    data,
    type: type as "png" | "jpg" | "gif" | "bmp",
    transformation: {
      width: Math.max(1, Math.round(width * scale)),
      height: Math.max(1, Math.round(height * scale)),
    },
    altText: { name: alt || "Markdown image" },
  });
}

function inlineRuns(
  tokens: InlineToken[]
): (TextRun | ExternalHyperlink | ImageRun)[] {
  const runs: (TextRun | ExternalHyperlink | ImageRun)[] = [];

  for (const token of tokens) {
    switch (token.type) {
      case "text":
      case "escape":
        runs.push(new TextRun(token.text));
        break;

      case "strong":
        runs.push(
          new TextRun({
            text: token.text,
            bold: true,
          })
        );
        break;

      case "em":
        runs.push(
          new TextRun({
            text: token.text,
            italics: true,
          })
        );
        break;

      case "del":
        runs.push(
          new TextRun({
            text: token.text,
            strike: true,
          })
        );
        break;

      case "codespan":
        runs.push(
          new TextRun({
            text: token.text,
            font: "Courier New",
          })
        );
        break;

      case "br":
        runs.push(new TextRun({ break: 1 }));
        break;

      case "link":
        runs.push(
          new ExternalHyperlink({
            link: token.href,
            children: [
              new TextRun({
                text: token.text,
                color: "000000",
                underline: { color: "000000" },
              }),
            ],
          })
        );
        break;

      case "image":
        {
          const imageRun = imageRunFromDataUri(token.href, token.text);
          if (!imageRun) {
            throw new Error(
              "Word conversion requires images to be embedded as PNG, JPEG, GIF, or BMP data URIs."
            );
          }
          runs.push(imageRun);
        }
        break;
    }
  }

  return runs;
}

function getInlineTokens(token: Token): InlineToken[] {
  if ("tokens" in token && Array.isArray(token.tokens)) {
    return token.tokens as InlineToken[];
  }

  return [];
}

function makeParagraph(
  token: Token,
  options: {
    heading?: (typeof HeadingLevel)[keyof typeof HeadingLevel];
    bullet?: boolean;
    numbered?: boolean;
    listLevel?: number;
    quote?: boolean;
  } = {},
  fontSize = 14
): Paragraph {
  const runs = inlineRuns(getInlineTokens(token));

  return new Paragraph({
    children: runs.length ? runs : [new TextRun("")],
    heading: options.heading,
    numbering: options.bullet
      ? { reference: "markdown-bullets", level: options.listLevel ?? 0 }
      : options.numbered
        ? { reference: "markdown-numbering", level: options.listLevel ?? 0 }
        : undefined,
    indent: options.quote ? { left: 600 } : undefined,
    spacing: { after: Math.round(160 * (fontSize / 14)) },
  });
}

function convertTokens(
  tokens: Token[],
  quote = false,
  listLevel = 0,
  fontSize = 14
): (Paragraph | Table)[] {
  const output: (Paragraph | Table)[] = [];

  for (const token of tokens) {
    switch (token.type) {
      case "heading": {
        const headingLevels = [
          HeadingLevel.HEADING_1,
          HeadingLevel.HEADING_2,
          HeadingLevel.HEADING_3,
          HeadingLevel.HEADING_4,
          HeadingLevel.HEADING_5,
          HeadingLevel.HEADING_6,
        ];

        output.push(
          makeParagraph(token, {
            heading: headingLevels[Math.min(token.depth - 1, 5)],
          }, fontSize)
        );
        break;
      }

      case "paragraph":
        output.push(makeParagraph(token, { quote }, fontSize));
        break;

      case "text":
        if (token.tokens) {
          output.push(makeParagraph(token, { quote }, fontSize));
        }
        break;

      case "blockquote":
        output.push(...convertTokens(token.tokens ?? [], true, 0, fontSize));
        break;

      case "list":
        for (const item of token.items) {
          const itemToken = {
            type: "text",
            text: item.text,
            tokens: item.tokens,
          } as Token;

          output.push(
            makeParagraph(itemToken, {
              bullet: !token.ordered,
              numbered: token.ordered,
              listLevel,
            }, fontSize)
          );

          for (const child of item.tokens) {
            if (child.type === "list") {
              output.push(...convertTokens([child], quote, listLevel + 1, fontSize));
            }
          }
        }
        break;

      case "code":
        {
          const code = token.text.replace(/\n$/, "");
          const lines: string[] = code.split("\n");
          output.push(
            new Paragraph({
              children: lines.flatMap((line, index) => [
                new TextRun({
                  text: line || " ",
                  font: "Courier New",
                }),
                ...(index < lines.length - 1
                  ? [new TextRun({ break: 1 })]
                  : []),
              ]),
              shading: { type: ShadingType.CLEAR, fill: "F2F4F7" },
              indent: { left: 180, right: 180 },
              spacing: {
                before: Math.round(100 * (fontSize / 14)),
                after: Math.round(200 * (fontSize / 14)),
                line: Math.round(240 * (fontSize / 14)),
              },
            })
          );
        }
        break;

      case "hr":
        output.push(
          new Paragraph({
            border: {
              bottom: {
                style: BorderStyle.SINGLE,
                size: 6,
                color: "000000",
              },
            },
            spacing: { after: Math.round(200 * (fontSize / 14)) },
          })
        );
        break;

      case "table": {
        const tableToken = token as Tokens.Table;
        const rows = [
          tableToken.header,
          ...tableToken.rows,
        ];

        const tableRows: TableRow[] = rows.map((row, rowIndex) => {
          return new TableRow({
            children: row.map((cell) => {
              return new TableCell({
                children: [
                  new Paragraph({
                    children:
                      rowIndex === 0
                        ? [new TextRun({ text: cell.text, bold: true })]
                        : inlineRuns(cell.tokens as InlineToken[]),
                  }),
                ],
                width: {
                  size: 100 / row.length,
                  type: WidthType.PERCENTAGE,
                },
              });
            }),
          });
        });

        output.push(
          new Table({
            rows: tableRows,
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1, color: "888888" },
              bottom: { style: BorderStyle.SINGLE, size: 1, color: "888888" },
              left: { style: BorderStyle.SINGLE, size: 1, color: "888888" },
              right: { style: BorderStyle.SINGLE, size: 1, color: "888888" },
              insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "888888" },
              insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "888888" },
            },
          })
        );
        break;
      }

      case "space":
        break;

      default:
        // Preserve readable text for otherwise unsupported token types.
        if ("text" in token && typeof token.text === "string" && token.text.trim()) {
          output.push(new Paragraph({ text: token.text }));
        }
    }
  }

  return output;
}

export async function convertMarkdownToWord(
  markdown: string,
  fontSize = 14
): Promise<Blob> {
  if (!markdown.trim()) {
    throw new Error("Please enter or upload Markdown content first.");
  }

  try {
    const tokens = marked.lexer(markdown, { breaks: true });

    const document = new Document({
      numbering: {
        config: [
          {
            reference: "markdown-numbering",
            levels: Array.from({ length: 6 }, (_, level) => ({
              level,
              format: "decimal" as const,
              text: `%${level + 1}.`,
              alignment: "left" as const,
              style: {
                paragraph: {
                  indent: { left: 720 * (level + 1), hanging: 360 },
                },
              },
            })),
          },
          {
            reference: "markdown-bullets",
            levels: Array.from({ length: 6 }, (_, level) => ({
              level,
              format: "bullet" as const,
              text: "•",
              alignment: "left" as const,
              style: {
                paragraph: {
                  indent: { left: 720 * (level + 1), hanging: 360 },
                  run: {
                    font: "Arial",
                  },
                },
              },
            })),
          },
        ],
      },
      styles: {
        default: {
          document: {
            run: { font: "Arial", size: fontSize * 2, color: "000000" },
            paragraph: {
              spacing: {
                after: Math.round(160 * (fontSize / 14)),
                line: Math.round(276 * (fontSize / 14)),
              },
            },
          },
          title: {
            run: { font: "Arial", size: Math.round(fontSize * 2 * (36 / 22)), bold: true, color: "000000" },
            paragraph: { spacing: { before: Math.round(240 * (fontSize / 14)), after: Math.round(240 * (fontSize / 14)) } },
          },
          heading1: {
            run: { font: "Arial", size: Math.round(fontSize * 2 * (32 / 22)), bold: true, color: "000000" },
            paragraph: { spacing: { before: Math.round(320 * (fontSize / 14)), after: Math.round(160 * (fontSize / 14)) }, keepNext: true },
          },
          heading2: {
            run: { font: "Arial", size: Math.round(fontSize * 2 * (28 / 22)), bold: true, color: "000000" },
            paragraph: { spacing: { before: Math.round(280 * (fontSize / 14)), after: Math.round(140 * (fontSize / 14)) }, keepNext: true },
          },
          heading3: {
            run: { font: "Arial", size: Math.round(fontSize * 2 * (24 / 22)), bold: true, color: "000000" },
            paragraph: { spacing: { before: Math.round(240 * (fontSize / 14)), after: Math.round(120 * (fontSize / 14)) }, keepNext: true },
          },
          heading4: {
            run: { font: "Arial", size: fontSize * 2, bold: true, color: "000000" },
          },
          heading5: {
            run: { font: "Arial", size: fontSize * 2, bold: true, color: "000000" },
          },
          heading6: {
            run: { font: "Arial", size: fontSize * 2, bold: true, color: "000000" },
          },
          hyperlink: {
            run: { color: "000000", underline: { color: "000000" } },
          },
        },
      },
      sections: [
        {
          properties: {
            page: {
              margin: {
                top: 1440,
                right: 1440,
                bottom: 1440,
                left: 1440,
              },
            },
          },
          children: convertTokens(tokens, false, 0, fontSize),
        },
      ],
    });

    return await Packer.toBlob(document);
  } catch (error) {
    if (
      error instanceof Error &&
      error.message.startsWith("Word conversion requires images")
    ) {
      throw error;
    }

    throw new Error(
      "Unable to create the Word document. Check the Markdown content and try again."
    );
  }
}
