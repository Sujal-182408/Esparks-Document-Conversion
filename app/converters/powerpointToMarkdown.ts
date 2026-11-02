
import JSZip from "jszip";

const MAX_FILE_SIZE = 20 * 1024 * 1024;

function escapeMarkdown(text: string): string {
  return text
    .replace(/\\/g, "\\\\")
    .replace(/([*_`])/g, "\\$1")
    .trim();
}

function getTextFromParagraph(paragraph: Element): string {
  return Array.from(paragraph.getElementsByTagName("a:t"))
    .map((node) => node.textContent ?? "")
    .join("")
    .trim();
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

      const slideLines: string[] = [];
      let headingAdded = false;

      for (const shape of shapes) {
        const paragraphs = Array.from(
          shape.getElementsByTagName("a:p")
        );

        for (const paragraph of paragraphs) {
          const text = getTextFromParagraph(paragraph);

          if (!text) continue;

          const escapedText = escapeMarkdown(text);

          if (!headingAdded) {
            slideLines.push(`## ${escapedText}`);
            headingAdded = true;
          } else {
            const level =
              paragraph.getElementsByTagName("a:pPr")[0]
                ?.getAttribute("lvl");

            const indent = Math.min(
              Math.max(Number(level ?? 0), 0),
              5
            );

            slideLines.push(
              `${"  ".repeat(indent)}- ${escapedText}`
            );
          }
        }
      }

      if (slideLines.length === 0) {
        slideLines.push("_No extractable text on this slide._");
      }

      markdownSlides.push(
        `# Slide ${index + 1}\n\n${slideLines.join("\n")}`
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
