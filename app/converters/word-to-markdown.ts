import { convertToHtml } from "mammoth/mammoth.browser";
import TurndownService from "turndown";

const turnDown = new TurndownService({
  headingStyle: "atx",
  bulletListMarker: "-",
  codeBlockStyle: "fenced",
  emDelimiter: "*",
});

turnDown.addRule("table", {
  filter: "table",
  replacement: (_content, table) => {
    const rows = Array.from(table.querySelectorAll("tr")).map((row) =>
      Array.from(row.querySelectorAll("th, td")).map((cell) =>
        turnDown
          .turndown(cell.innerHTML)
          .replace(/\s*\n\s*/g, " ")
          .replace(/\|/g, "\\|")
          .trim()
      )
    );

    if (!rows.length) return "";

    const columnCount = Math.max(...rows.map((row) => row.length));
    const normalizedRows = rows.map((row) =>
      Array.from({ length: columnCount }, (_, index) => row[index] ?? "")
    );
    const header = normalizedRows[0];
    const separator = header.map(() => "---");

    return `\n\n| ${header.join(" | ")} |\n| ${separator.join(" | ")} |\n${normalizedRows
      .slice(1)
      .map((row) => `| ${row.join(" | ")} |`)
      .join("\n")}\n\n`;
  },
});

export async function wordToMarkdown(file: File): Promise<string> {
  if (!file.name.toLowerCase().endsWith(".docx"))
    throw new Error("Please select a Word .docx file");
  if (file.size === 0)
    throw new Error("The selected file is empty.");
  if (file.size > 20 * 1024 * 1024)
    throw new Error("The maximum file size is 20 MB.");
  // DOCX is a ZIP-based format. Check its signature before parsing.
  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b)
    throw new Error("This file does not appear to be a valid DOCX document.");
  try {
    const result = await convertToHtml({
      arrayBuffer: buffer,
    }, {
      styleMap: [
        "p[style-name='Title'] => h1:fresh",
        "p[style-name='Document Title'] => h1:fresh",
        "p[style-name='Subtitle'] => h2:fresh",
        "p.Heading1 => h1:fresh",
        "p[style-name='Heading 1'] => h1:fresh",
        "p[style-name='heading 1'] => h1:fresh",
        "p.Heading2 => h2:fresh",
        "p[style-name='Heading 2'] => h2:fresh",
        "p[style-name='heading 2'] => h2:fresh",
        "p.Heading3 => h3:fresh",
        "p[style-name='Heading 3'] => h3:fresh",
        "p[style-name='heading 3'] => h3:fresh",
        "p.Heading4 => h4:fresh",
        "p[style-name='Heading 4'] => h4:fresh",
        "p[style-name='heading 4'] => h4:fresh",
        "p.Heading5 => h5:fresh",
        "p[style-name='Heading 5'] => h5:fresh",
        "p[style-name='heading 5'] => h5:fresh",
        "p.Heading6 => h6:fresh",
        "p[style-name='Heading 6'] => h6:fresh",
        "p[style-name='heading 6'] => h6:fresh",
      ],
    });
    const markdown = turnDown.turndown(result.value).trim();
    if (!markdown) {
      throw new Error(
        "No convertible text was found in this Word document."
      );
    }

    return markdown + "\n";
  } catch (error) {
    if (
      error instanceof Error &&
      error.message.startsWith("No convertible text")
    )
      throw error;
    throw new Error(
      "Unable to convert this Word document. It may be corrupt, unsupported, or protected."
    );
  }
}