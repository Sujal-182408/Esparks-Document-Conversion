
// @ts-ignore -- `docx` declarations may be missing in some environments.
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

function inlineRuns(tokens: InlineToken[]): (TextRun | ExternalHyperlink)[] {
  const runs: (TextRun | ExternalHyperlink)[] = [];

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
                style: "Hyperlink",
              }),
            ],
          })
        );
        break;

      case "image":
        // Preserve the image reference as readable text.
        // Embedding image files requires a separate image-handling strategy.
        runs.push(
          new TextRun({
            text: `[Image: ${token.text || token.href}]`,
            italics: true,
          })
        );
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
    quote?: boolean;
  } = {}
): Paragraph {
  const runs = inlineRuns(getInlineTokens(token));

  return new Paragraph({
    children: runs.length ? runs : [new TextRun("")],
    heading: options.heading,
    bullet: options.bullet ? { level: 0 } : undefined,
    numbering: options.numbered
      ? { reference: "markdown-numbering", level: 0 }
      : undefined,
    indent: options.quote ? { left: 600 } : undefined,
    spacing: { after: 160 },
  });
}

function convertTokens(tokens: Token[]): (Paragraph | Table)[] {
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
          })
        );
        break;
      }

      case "paragraph":
        output.push(makeParagraph(token));
        break;

      case "text":
        if (token.tokens) {
          output.push(makeParagraph(token));
        }
        break;

      case "blockquote":
        output.push(
          new Paragraph({
            children: [new TextRun({ text: "Quote", bold: true })],
            indent: { left: 600 },
          })
        );
        output.push(...convertTokens(token.tokens ?? []));
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
            })
          );

          // Include nested lists where present.
          for (const child of item.tokens) {
            if (child.type === "list") {
              output.push(...convertTokens([child]));
            }
          }
        }
        break;

      case "code":
        output.push(
          new Paragraph({
            children: [
              new TextRun({
                text: token.text.replace(/\n$/, ""),
                font: "Courier New",
              }),
            ],
            spacing: { before: 100, after: 200 },
          })
        );
        break;

      case "hr":
        output.push(
          new Paragraph({
            border: {
              bottom: {
                style: BorderStyle.SINGLE,
                size: 6,
                color: "808080",
              },
            },
            spacing: { after: 200 },
          })
        );
        break;

      case "table": {
        const rows = [
          token.header,
          ...token.rows.flat(),
        ];

        const tableRows: TableRow[] = rows.map((row, rowIndex) => {
          const cells = Array.isArray(row) ? row : [row];

          return new TableRow({
            children: cells.map((cell) => {
              const cellTokens =
                typeof cell === "string"
                  ? (marked.lexer(cell)[0] as Token | undefined)
                  : undefined;

              const text =
                typeof cell === "string"
                  ? cell
                  : cell.text;

              return new TableCell({
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text,
                        bold: rowIndex === 0,
                      }),
                    ],
                  }),
                ],
                width: { size: 100 / cells.length, type: WidthType.PERCENTAGE },
              });
            }),
          });
        });

        output.push(
          new Table({
            rows: tableRows,
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
              bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
              left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
              right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
              insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
              insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
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
  markdown: string
): Promise<Blob> {
  if (!markdown.trim()) {
    throw new Error("Please enter or upload Markdown content first.");
  }

  try {
    const tokens = marked.lexer(markdown);

    const document = new Document({
      numbering: {
        config: [
          {
            reference: "markdown-numbering",
            levels: [
              {
                level: 0,
                format: "decimal",
                text: "%1.",
                alignment: "left",
                style: {
                  paragraph: {
                    indent: { left: 720, hanging: 360 },
                  },
                },
              },
            ],
          },
        ],
      },
      sections: [
        {
          properties: {},
          children: convertTokens(tokens),
        },
      ],
    });

    return await Packer.toBlob(document);
  } catch {
    throw new Error(
      "Unable to create the Word document. Check the Markdown content and try again."
    );
  }
}
