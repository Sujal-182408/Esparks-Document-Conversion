import { ConversionOption } from "../types/conversion";

export const conversionOptions: ConversionOption[] = [
  {
    id: "word-to-markdown",
    title: "Word to Markdown",
    description: "Convert Word document into readable Markdown",
    inputExtensions: ".docx",
    outputExtensions: ".md"
  },
  {
    id: "pdf-to-markdown",
    title: "PDF to Markdown",
    description: "Extract text and structure from PDF documents.",
    inputExtensions: ".pdf",
    outputExtensions: ".md"
  },
  {
    id: "powerpoint-to-markdown",
    title: "PowerPoint to Markdown",
    description: "Extract slide content into Markdown.",
    inputExtensions: ".pptx",
    outputExtensions: ".md"
  },
  {
    id: "markdown-to-word",
    title: "MarkDown to Word",
    description: "Create an editable Word document from Markdown.",
    inputExtensions: ".md",
    outputExtensions: ".docx"
  },
  {
    id: "markdown-to-pdf",
    title: "MarkDown to PDF",
    description: "Generate a PDF from Markdown content.",
    inputExtensions: ".md",
    outputExtensions: ".pdf"
  },
  {
    id: "markdown-to-powerpoint",
    title: "MarkDown to PowerPoint",
    description: "Turn Markdown headings and content into slides.",
    inputExtensions: ".md",
    outputExtensions: ".pptx"
  },
];