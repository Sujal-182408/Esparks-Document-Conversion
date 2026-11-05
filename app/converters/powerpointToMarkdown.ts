
import JSZip from "jszip";

const MAX_FILE_SIZE = 20 * 1024 * 1024;

function escapeMarkdown(text: string): string {
  return text
    .replace(/\\/g, "\\\\")
    .replace(/([*_`])/g, "\\$1");
}

function getParagraphMarkdown(paragraph: Element): {
  text: string;
  bullet: boolean;
  level: number;
} {
  const paragraphProperties =
    paragraph.getElementsByTagName("a:pPr")[0];
  const level = Math.min(
    Math.max(Number(paragraphProperties?.getAttribute("lvl") ?? 0), 0),
    5
  );
  const bullet =
    !!paragraphProperties?.getElementsByTagName("a:buChar").length ||
    !!paragraphProperties?.getElementsByTagName("a:buAutoNum").length;
  const runs = Array.from(paragraph.getElementsByTagName("a:r"));
  const text = runs.length
    ? runs
        .map((run) => {
          const value = Array.from(run.getElementsByTagName("a:t"))
            .map((node) => node.textContent ?? "")
            .join("");
          if (!value) return "";

          const properties = run.getElementsByTagName("a:rPr")[0];
          const bold = properties?.getAttribute("b") === "1";
          const italic = properties?.getAttribute("i") === "1";
          const escaped = escapeMarkdown(value);

          if (bold && italic) return `***${escaped}***`;
          if (bold) return `**${escaped}**`;
          if (italic) return `*${escaped}*`;
          return escaped;
        })
        .join("")
    : Array.from(paragraph.getElementsByTagName("a:t"))
        .map((node) => escapeMarkdown(node.textContent ?? ""))
        .join("");

  return { text: text.trim(), bullet, level };
}

function getPlaceholderType(shape: Element): string | null {
  return shape.getElementsByTagName("p:ph")[0]?.getAttribute("type") ?? null;
}

function getTableMarkdown(table: Element): string {
  const rows = Array.from(table.getElementsByTagName("a:tr")).map((row) =>
    Array.from(row.getElementsByTagName("a:tc")).map((cell) =>
      Array.from(cell.getElementsByTagName("a:t"))
        .map((node) => escapeMarkdown(node.textContent ?? "").trim())
        .join("")
        .replace(/\|/g, "\\|")
    )
  );

  if (!rows.length) return "";

  const columnCount = Math.max(...rows.map((row) => row.length));
  const normalizedRows = rows.map((row) =>
    Array.from({ length: columnCount }, (_, index) => row[index] ?? "")
  );
  const separator = normalizedRows[0].map(() => "---");

  return [
    `| ${normalizedRows[0].join(" | ")} |`,
    `| ${separator.join(" | ")} |`,
    ...normalizedRows.slice(1).map((row) => `| ${row.join(" | ")} |`),
  ].join("\n");
}

export async function convertPowerPointToMarkdown(
  file: File
): Promise<string> {
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    throw new Error("Please select a PowerPoint .pptx file.");
  }

  if (file.size === 0) {
    throw new Error("The selected PowerPoint file is empty.");
  }

  if (file.size > MAX_FILE_SIZE) {
    throw new Error("The maximum file size is 20 MB.");
  }

  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);

  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
    throw new Error("This does not appear to be a valid PPTX file.");
  }

  try {
    const zip = await JSZip.loadAsync(buffer);

    const parser = new DOMParser();
    const presentationFile = zip.file("ppt/presentation.xml");
    const relationshipsFile = zip.file(
      "ppt/_rels/presentation.xml.rels"
    );

    if (!presentationFile || !relationshipsFile) {
      throw new Error("Required PowerPoint presentation files are missing.");
    }

    const presentationXml = parser.parseFromString(
      await presentationFile.async("text"),
      "application/xml"
    );

    const relationshipsXml = parser.parseFromString(
      await relationshipsFile.async("text"),
      "application/xml"
    );

    const relationshipMap = new Map<string, string>();

    Array.from(
      relationshipsXml.getElementsByTagName("Relationship")
    ).forEach((relationship) => {
      const id = relationship.getAttribute("Id");
      const target = relationship.getAttribute("Target");

      if (id && target) {
        relationshipMap.set(id, target);
      }
    });

    const slideIds = Array.from(
      presentationXml.getElementsByTagName("p:sldId")
    );

    if (slideIds.length === 0) {
      throw new Error("No slides were found in this presentation.");
    }

    const markdownSlides: string[] = [];

    for (let index = 0; index < slideIds.length; index++) {
      const relationshipId =
        slideIds[index].getAttribute("r:id");

      if (!relationshipId) continue;

      const target = relationshipMap.get(relationshipId);

      if (!target) continue;

      const slidePath = target.startsWith("/")
        ? target.slice(1)
        : `ppt/${target.replace(/^(\.\/)+/, "")}`;

      const slideFile = zip.file(slidePath);

      if (!slideFile) continue;

      const slideXml = parser.parseFromString(
        await slideFile.async("text"),
        "application/xml"
      );

      const shapes = Array.from(
        slideXml.getElementsByTagName("p:sp")
      );
      const hasTitlePlaceholder = shapes.some((shape) => {
        const type = getPlaceholderType(shape);
        return type === "title" || type === "ctrTitle";
      });

      const slideLines: string[] = [];
      let headingAdded = false;

      for (const shape of shapes) {
        const placeholderType = getPlaceholderType(shape);
        const isTitlePlaceholder =
          placeholderType === "title" || placeholderType === "ctrTitle";
        const isSubtitlePlaceholder = placeholderType === "subTitle";
        const paragraphs = Array.from(
          shape.getElementsByTagName("a:p")
        );

        for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
          const { text, bullet, level } = getParagraphMarkdown(paragraph);

          if (!text) continue;

          if (
            (isTitlePlaceholder && paragraphIndex === 0) ||
            (isSubtitlePlaceholder && paragraphIndex === 0)
          ) {
            const headingLevel =
              isTitlePlaceholder && paragraphIndex === 0 && index === 0
                ? "#"
                : "##";
            slideLines.push(`${headingLevel} ${text}`);
            headingAdded = true;
          } else if (!hasTitlePlaceholder && !headingAdded && !bullet) {
            slideLines.push(`${index === 0 ? "#" : "##"} ${text}`);
            headingAdded = true;
          } else {
            slideLines.push(
              bullet ? `${"  ".repeat(level)}- ${text}` : text
            );
          }
        }
      }

      const tables = Array.from(slideXml.getElementsByTagName("a:tbl"))
        .map(getTableMarkdown)
        .filter(Boolean);
      if (tables.length) slideLines.push(...tables);

      if (slideLines.length === 0) {
        slideLines.push("_No extractable text on this slide._");
      }

      markdownSlides.push(
        headingAdded
          ? slideLines.join("\n\n")
          : `${index === 0 ? "#" : "##"} Slide ${index + 1}\n\n${slideLines.join("\n\n")}`
      );
    }

    if (markdownSlides.length === 0) {
      throw new Error("No readable slide content was found.");
    }

    return markdownSlides.join("\n\n---\n\n") + "\n";
  } catch (error) {
    if (
      error instanceof Error &&
      (
        error.message.startsWith("Required PowerPoint") ||
        error.message.startsWith("No slides") ||
        error.message.startsWith("No readable")
      )
    ) {
      throw error;
    }

    console.error("PowerPoint conversion failed:", error);

    throw new Error(
      "Unable to read this PowerPoint file. It may be corrupt, encrypted, or unsupported."
    );
  }
}
