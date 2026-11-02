import {convertToHtml} from "mammoth/mammoth.browser";
import TurndownService from "turndown";

const turnDown = new TurndownService({
  headingStyle: "atx",
  bulletListMarker: "-",
  codeBlockStyle: "fenced",
  emDelimiter: "*"
});

export async function wordToMarkdown(file: File):Promise<string>{
  if(!file.name.toLocaleLowerCase().endsWith(".docx"))
    throw new Error("Please select a Word .docx file");
  if(file.size === 0)
    throw new Error("The selected file is empty.");
  if(file.size > 20 * 1024 * 1024)
    throw new Error("The maximum file size is 20 MB.");
  // DOCX is a ZIP-based format. Check its signature before parsing.
  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  if(bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b)
    throw new Error("This file does not appear to be a valid DOCX document.");
  try {
    const result = await convertToHtml({
      arrayBuffer: buffer
    });
    const markdown = turnDown.turndown(result.value).trim();
    if (!markdown) {
      throw new Error(
        "No convertible text was found in this Word document."
      );
    }

    return markdown + "\n";
  } catch (error) {
    if(error instanceof Error && error.message.startsWith("No convertible text."))
      throw error;
    throw new Error(
      "Unable to convert this Word document. It may be corrupt, unsupported, or protected."
    );
  }

}