
"use client";

import { useRef, useState, useEffect } from "react";
import { useSearchParams } from "next/navigation";
import {
  Upload, FileText, Download, RefreshCw, X,
  CheckCircle2, AlertCircle, ChevronDown, Eye, Maximize2, Minimize2,
} from "lucide-react";
import ReactMarkdown, { type Components } from "react-markdown";
import remarkGfm from "remark-gfm";
import remarkBreaks from "remark-breaks";

import {
  downloadTextFile,
  downloadConvertedFile,
  downloadWordFile,
  markdownFilename,
  pdfMarkdownFilename,
  powerpointMarkdownFilename
} from "../services/downloadservice";

import { wordToMarkdown } from "../converters/word-to-markdown";
import { pdfToMarkdown } from "../converters/pdfToMarkdown";
import { convertPowerPointToMarkdown } from "../converters/powerpointToMarkdown";
import { convertMarkdownToWord } from "../converters/markdownToWord";
import { convertMarkdownToPdf } from "../converters/mdToPdf";
import { convertMarkdownToPowerPoint } from "../converters/markdownToPowerPoint";
import { conversionOptions, type Format } from "../converters/conversionOptions";

type DocumentConverterProps = {
  initialInputFormat?: Format;
  initialOutputFormat?: Format;
};

type GeneratedDocument = {
  blob: Blob;
  format: "docx" | "pdf" | "pptx";
};

const validFormats: Format[] = ["docx", "pdf", "pptx", "md"];
const MAX_FILE_SIZE = 20 * 1024 * 1024;
const MAX_MARKDOWN_SIZE = 28 * 1024 * 1024;

const EXAMPLE_MARKDOWN = `# Markdown conversion example

This example shows **bold text**, *italic text*, and inline code such as \`const answer = 42;\`.

## Lists

- Write headings and paragraphs
- Keep nested lists
  - Preserve the hierarchy

1. Convert Markdown to Word
2. Convert Markdown to PDF or PowerPoint

## Code snippet

\`\`\`javascript
function greet(name) {
  const message = \`Hello, \${name}!\`;
  console.log(message);
}

greet("Markdown");
\`\`\`

## Table

| Format | Extension |
| --- | --- |
| Word | .docx |
| PDF | .pdf |
| PowerPoint | .pptx |
`;

const formatNames: Record<Format, string> = {
  docx: "Word Document",
  pdf: "PDF Document",
  pptx: "PowerPoint Presentation",
  md: "Markdown File"
};

const extensions: Record<Format, string> = {
  docx: ".docx",
  pdf: ".pdf",
  pptx: ".pptx",
  md: ".md"
};

/**
 * Validate URLs before rendering them in Markdown previews.
 * Empty URLs and unsafe protocols are rejected.
 */
const safeMarkdownUrl = (value?: string): string | undefined => {
  const url = value?.trim();

  if (!url) return undefined;

  if (/^(https?:\/\/|mailto:|tel:)/i.test(url)) {
    return url;
  }

  if (/^(data:image\/(?:png|jpe?g|gif|webp|svg\+xml);base64,)/i.test(url)) {
    return url;
  }

  if (/^(blob:|\/|\.\/|\.\.\/|#)/.test(url)) {
    return url;
  }

  return undefined;
};

const markdownPreviewComponents = (fontSize: number): Components => ({
  h1: ({ children }) => (
    <h1
      style={{ fontSize: `${fontSize * (30 / 14)}pt`, lineHeight: 1.25, paddingBottom: "0.35em" }}
      className="mb-[0.55em] mt-[1.4em] border-b border-slate-200 font-bold text-slate-900 first:mt-0"
    >
      {children}
    </h1>
  ),
  h2: ({ children }) => (
    <h2
      style={{ fontSize: `${fontSize * (24 / 14)}pt`, lineHeight: 1.25 }}
      className="mb-[0.55em] mt-[1.4em] font-bold text-slate-900"
    >
      {children}
    </h2>
  ),
  h3: ({ children }) => (
    <h3
      style={{ fontSize: `${fontSize * (20 / 14)}pt`, lineHeight: 1.25 }}
      className="mb-[0.5em] mt-[1.2em] font-semibold text-slate-900"
    >
      {children}
    </h3>
  ),
  h4: ({ children }) => (
    <h4
      style={{ fontSize: `${fontSize * (17 / 14)}pt`, lineHeight: 1.25 }}
      className="mb-[0.5em] mt-[1.2em] font-semibold text-slate-900"
    >
      {children}
    </h4>
  ),
  h5: ({ children }) => (
    <h5
      style={{ fontSize: `${fontSize * (17 / 14)}pt`, lineHeight: 1.25 }}
      className="mb-[0.5em] mt-[1.2em] font-semibold text-slate-900"
    >
      {children}
    </h5>
  ),
  h6: ({ children }) => (
    <h6
      style={{ fontSize: `${fontSize * (17 / 14)}pt`, lineHeight: 1.25 }}
      className="mb-[0.5em] mt-[1.2em] font-semibold text-slate-900"
    >
      {children}
    </h6>
  ),
  p: ({ children }) => (
    <p style={{ marginBottom: "0.8em" }}>{children}</p>
  ),
  ul: ({ children }) => (
    <ul className="mb-[1em] list-disc space-y-[0.25em] pl-[1.85em]">
      {children}
    </ul>
  ),
  ol: ({ children }) => (
    <ol className="mb-[1em] list-decimal space-y-[0.25em] pl-[1.85em]">
      {children}
    </ol>
  ),
  li: ({ children }) => (
    <li className="pl-1">{children}</li>
  ),
  blockquote: ({ children }) => (
    <blockquote className="my-[1em] border-l-4 border-cyan-500 bg-slate-100 py-[0.7em] pl-[1em] text-slate-600">
      {children}
    </blockquote>
  ),
  a: ({ children, href }) => {
    const safeHref = safeMarkdownUrl(href);

    if (!safeHref) {
      return <>{children}</>;
    }

    return (
      <a
        href={safeHref}
        className="text-cyan-700 underline decoration-cyan-300 underline-offset-2"
        target="_blank"
        rel="noopener noreferrer"
      >
        {children}
      </a>
    );
  },
  img: ({ src, alt, title }) => {
    const safeSrc = safeMarkdownUrl(src);

    if (!safeSrc) return null;

    return (
      <span className="my-4 block">
        {/* eslint-disable-next-line @next/next/no-img-element */}
        <img
          src={safeSrc}
          alt={alt || "Markdown image"}
          title={title}
          loading="lazy"
          decoding="async"
          className="mx-auto h-auto max-w-full rounded-lg object-contain"
          onError={(event) => {
            event.currentTarget.style.display = "none";
          }}
        />
        {alt && (
          <span className="mt-1 block text-center text-xs text-slate-500">
            {alt}
          </span>
        )}
      </span>
    );
  },
  strong: ({ children }) => (
    <strong className="font-bold text-slate-900">{children}</strong>
  ),
  del: ({ children }) => (
    <del className="text-slate-500">{children}</del>
  ),
  code: ({ children, className }) => (
    <code className={`${className ?? ""} rounded bg-slate-200 px-1.5 py-0.5 font-mono text-[0.9em] text-slate-800`}>
      {children}
    </code>
  ),
  pre: ({ children }) => (
    <pre
      style={{ fontSize: `${fontSize}pt`, lineHeight: 1.5 }}
      className="mb-[1em] overflow-x-auto rounded-lg bg-slate-900 p-[0.85em] font-mono text-slate-100 [&_code]:bg-transparent [&_code]:p-0 [&_code]:text-inherit"
    >
      {children}
    </pre>
  ),
  table: ({ children }) => (
    <div className="mb-[1em] overflow-x-auto">
      <table
        style={{ fontSize: `${fontSize}pt` }}
        className="w-full border-collapse text-left"
      >
        {children}
      </table>
    </div>
  ),
  thead: ({ children }) => (
    <thead className="bg-cyan-50 text-slate-900">
      {children}
    </thead>
  ),
  th: ({ children }) => (
    <th className="border border-slate-300 px-3 py-2 font-semibold">
      {children}
    </th>
  ),
  td: ({ children }) => (
    <td className="border border-slate-300 px-3 py-2 align-top">
      {children}
    </td>
  ),
  hr: () => (
    <hr className="my-7 border-slate-300" />
  ),
});

const detectFormat = (filename: string): Format | null => {
  const extension = filename.toLowerCase().split(".").pop();

  return validFormats.includes(extension as Format)
    ? extension as Format
    : null;
};

const markdownOutputFilename = (filename: string, format: Format): string => {
  if (format === "docx") return markdownFilename(filename);
  if (format === "pdf") return pdfMarkdownFilename(filename);
  if (format === "pptx") return powerpointMarkdownFilename(filename);
  return `${filename.replace(/\.[^.]+$/, "")}.md`;
};

async function createMarkdownOutput(
  markdown: string,
  format: Format,
  fontSize: number
): Promise<GeneratedDocument> {
  if (format === "docx") {
    return {
      blob: await convertMarkdownToWord(markdown, fontSize),
      format: "docx",
    };
  }

  if (format === "pdf") {
    return {
      blob: await convertMarkdownToPdf(markdown, fontSize),
      format: "pdf",
    };
  }

  if (format === "pptx") {
    return {
      blob: await convertMarkdownToPowerPoint(markdown, { fontSize }),
      format: "pptx",
    };
  }

  throw new Error("Select a document format to convert the Markdown file.");
}

async function extractMarkdown(file: File, format: Format): Promise<string> {
  if (format === "md") return file.text();
  if (format === "docx") return wordToMarkdown(file);
  if (format === "pdf") return pdfToMarkdown(file);
  if (format === "pptx") return convertPowerPointToMarkdown(file);

  throw new Error("This input format is not supported.");
}

const formatFileSize = (size: number): string =>
  size < 1024
    ? `${size} B`
    : size < 1024 * 1024
      ? `${(size / 1024).toFixed(1)} KB`
      : `${(size / 1024 / 1024).toFixed(2)} MB`;

const markdownForPreview = (markdown: string): string =>
  markdown
    .replace(
      /<span\b[^>]*\bdata-pptx-(?:color|font-size)="[^"]+"[^>]*>([\s\S]*?)<\/span>/gi,
      "$1",
    )
    .replace(/<\/?u>/gi, "")
    .replace(
      /<!--\s*pptx-align:(?:left|center|right|justify)\s*-->/gi,
      "",
    )
    .replace(/<!--\s*slide\s*-->/gi, "\n\n---\n\n");

export default function DocumentConverter({
  initialInputFormat,
  initialOutputFormat
}: DocumentConverterProps = {}) {
  const searchParams = useSearchParams();

  const preferredInput = searchParams.get("from") ?? initialInputFormat;
  const preferredOutput = searchParams.get("to") ?? initialOutputFormat;

  const validInput = validFormats.includes(preferredInput as Format)
    ? preferredInput as Format
    : null;

  const validOutput = validFormats.includes(preferredOutput as Format)
    ? preferredOutput as Format
    : null;

  const selectedOutput =
    validInput && validOutput && conversionOptions[validInput].includes(validOutput)
      ? validOutput
      : validInput
        ? conversionOptions[validInput][0]
        : undefined;

  return (
    <DocumentConverterContent
      key={JSON.stringify([preferredInput, preferredOutput])}
      initialInputFormat={validInput ?? undefined}
      initialOutputFormat={selectedOutput}
    />
  );
}

function DocumentConverterContent({
  initialInputFormat,
  initialOutputFormat
}: DocumentConverterProps) {
  const fileRef = useRef<HTMLInputElement>(null);
  const fileReadIdRef = useRef(0);
  const fontUpdateIdRef = useRef(0);

  const [file, setFile] = useState<File | null>(null);
  const [customMarkdownMode, setCustomMarkdownMode] = useState(
    initialInputFormat === "md",
  );
  const [inputFormat, setInputFormat] = useState<Format | null>(
    initialInputFormat ?? null
  );
  const [outputFormat, setOutputFormat] = useState<Format | "">(
    initialOutputFormat ?? ""
  );
  const [previewMarkdown, setPreviewMarkdown] = useState<string | null>(null);
  const [previewLoading, setPreviewLoading] = useState(false);
  const [fontSize, setFontSize] = useState(14);
  const [fontSizeUpdating, setFontSizeUpdating] = useState(false);
  const [result, setResult] = useState("");
  const [generatedFile, setGeneratedFile] = useState<GeneratedDocument | null>(null);
  const [busy, setBusy] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState("");
  const [previewOpen, setPreviewOpen] = useState(false);
  const [previewMode, setPreviewMode] = useState<"rendered" | "source">("rendered");
  const [previewFullscreen, setPreviewFullscreen] = useState(false);
  const [outputPreviewFullscreen, setOutputPreviewFullscreen] = useState(false);

  const outputOptions = inputFormat
    ? conversionOptions[inputFormat]
    : [];

  const selectFile = (selectedFile?: File) => {
    const readId = ++fileReadIdRef.current;

    fontUpdateIdRef.current += 1;
    setFontSizeUpdating(false);
    setError("");
    setSuccess("");
    setResult("");
    setPreviewMarkdown(null);
    setPreviewLoading(false);
    setGeneratedFile(null);
    setPreviewOpen(false);
    setPreviewFullscreen(false);
    setOutputPreviewFullscreen(false);

    if (!selectedFile) return;

    const detectedFormat = detectFormat(selectedFile.name);

    if (!detectedFormat) {
      setFile(null);
      setError("Unsupported file format. Please upload DOCX, PDF, PPTX or MD.");
      return;
    }

    if (selectedFile.size === 0) {
      setFile(null);
      setError("The selected file is empty.");
      return;
    }

    const maximumSize =
      detectedFormat === "md" ? MAX_MARKDOWN_SIZE : MAX_FILE_SIZE;

    if (selectedFile.size > maximumSize) {
      setFile(null);
      setError(
        detectedFormat === "md"
          ? "Markdown files must not exceed 28 MB."
          : "File size must not exceed 20 MB.",
      );
      return;
    }

    if (initialInputFormat && detectedFormat !== initialInputFormat) {
      setFile(null);
      setError(
        `Please upload a ${formatNames[initialInputFormat]} (${extensions[initialInputFormat]}) file.`
      );
      return;
    }

    setCustomMarkdownMode(false);
    setFile(selectedFile);
    setInputFormat(detectedFormat);
    setPreviewLoading(true);

    extractMarkdown(selectedFile, detectedFormat)
      .then((markdown) => {
        if (readId === fileReadIdRef.current) {
          setPreviewMarkdown(markdown);
        }
      })
      .catch((readError: unknown) => {
        if (readId !== fileReadIdRef.current) return;

        setError(
          readError instanceof Error
            ? `Could not prepare the preview: ${readError.message}`
            : "Could not prepare the preview."
        );
      })
      .finally(() => {
        if (readId === fileReadIdRef.current) {
          setPreviewLoading(false);
        }
      });

    setOutputFormat((currentOutput) =>
      currentOutput && conversionOptions[detectedFormat].includes(currentOutput)
        ? currentOutput
        : conversionOptions[detectedFormat][0] || ""
    );
  };

  const startCustomMarkdown = () => {
    fileReadIdRef.current += 1;
    fontUpdateIdRef.current += 1;

    setFontSizeUpdating(false);
    setFile(null);
    setInputFormat("md");
    setCustomMarkdownMode(true);
    setPreviewMarkdown("");
    setPreviewLoading(false);
    setResult("");
    setGeneratedFile(null);
    setError("");
    setSuccess("");
    setPreviewOpen(false);
    setPreviewFullscreen(false);
    setOutputPreviewFullscreen(false);

    setOutputFormat((current) =>
      current && conversionOptions.md.includes(current)
        ? current
        : conversionOptions.md[0],
    );

    if (fileRef.current) fileRef.current.value = "";
  };

  const loadMarkdownExample = () => {
    fontUpdateIdRef.current += 1;
    setFontSizeUpdating(false);
    setPreviewMarkdown(EXAMPLE_MARKDOWN);
    setResult("");
    setGeneratedFile(null);
    setError("");
    setSuccess("");
    setPreviewOpen(false);
  };

  const editMarkdown = (markdown: string) => {
    fontUpdateIdRef.current += 1;
    setFontSizeUpdating(false);
    setPreviewMarkdown(markdown);
    setResult("");
    setGeneratedFile(null);
    setError("");
    setSuccess("");
  };

  const handleConvert = async () => {
    if ((!file && !customMarkdownMode) || !inputFormat || !outputFormat) {
      setError("Please provide Markdown content or upload a file and select an output format.");
      return;
    }

    if (!conversionOptions[inputFormat].includes(outputFormat)) {
      setError("This conversion is not supported.");
      return;
    }

    if (inputFormat === "md" && !previewMarkdown?.trim()) {
      setError("Enter Markdown content or load the example before converting.");
      return;
    }

    setBusy(true);
    fontUpdateIdRef.current += 1;
    setFontSizeUpdating(false);
    setError("");
    setSuccess("");
    setResult("");
    setGeneratedFile(null);
    setPreviewOpen(false);
    setPreviewFullscreen(false);

    try {
      const markdown =
        previewMarkdown ??
        (file ? await extractMarkdown(file, inputFormat) : "");

      setPreviewMarkdown(markdown);
      setResult(markdown);

      if (inputFormat === "md") {
        const document = await createMarkdownOutput(
          markdown,
          outputFormat,
          fontSize
        );
        setGeneratedFile(document);
      } else {
        setGeneratedFile(null);
      }

      setSuccess("Your document has been converted successfully. Preview or download it below.");
    } catch (err) {
      console.error("Conversion error:", err);

      setError(
        err instanceof Error
          ? err.message
          : "Conversion failed. Please try again."
      );
    } finally {
      setBusy(false);
    }
  };

  const updateFontSize = async (nextFontSize: number) => {
    const minFontSize = outputFormat === "pptx" ? 10 : 8;
    const maxFontSize = outputFormat === "pptx" ? 24 : 32;

    const clampedFontSize = Math.min(
      maxFontSize,
      Math.max(minFontSize, nextFontSize)
    );

    const updateId = ++fontUpdateIdRef.current;
    setFontSize(clampedFontSize);

    if (
      !generatedFile ||
      (!file && !customMarkdownMode) ||
      !inputFormat ||
      !outputFormat ||
      outputFormat === "md"
    ) {
      setFontSizeUpdating(false);
      return;
    }

    setFontSizeUpdating(true);
    setError("");
    setSuccess("");

    try {
      const markdown =
        previewMarkdown ??
        (file ? await extractMarkdown(file, inputFormat) : "");

      const converted = await createMarkdownOutput(
        markdown,
        outputFormat,
        clampedFontSize
      );

      if (updateId !== fontUpdateIdRef.current) return;

      setResult(markdown);
      setGeneratedFile(converted);
      setSuccess(`Output regenerated at ${clampedFontSize} pt.`);
    } catch (err) {
      if (updateId !== fontUpdateIdRef.current) return;

      console.error("Font-size update error:", err);
      setGeneratedFile(null);

      setError(
        err instanceof Error
          ? err.message
          : "Could not regenerate the document at the selected font size."
      );
    } finally {
      if (updateId === fontUpdateIdRef.current) {
        setFontSizeUpdating(false);
      }
    }
  };

  const handleDownload = async () => {
    if ((!file && !customMarkdownMode) || fontSizeUpdating) return;

    const sourceFilename = file?.name ?? "custom-markdown.md";

    if (generatedFile) {
      if (generatedFile.format === "docx") {
        await downloadWordFile(generatedFile.blob, sourceFilename);
      } else {
        downloadConvertedFile(
          generatedFile.blob,
          sourceFilename,
          generatedFile.format === "pdf" ? ".pdf" : ".pptx"
        );
      }
      return;
    }

    if (outputFormat === "md" && result && inputFormat) {
      downloadTextFile(
        result,
        markdownOutputFilename(sourceFilename, inputFormat),
        "text/markdown;charset=utf-8"
      );
    }
  };

  const reset = () => {
    fileReadIdRef.current += 1;
    fontUpdateIdRef.current += 1;

    setFontSizeUpdating(false);
    setFile(null);
    setCustomMarkdownMode(initialInputFormat === "md");
    setPreviewMarkdown(null);
    setPreviewLoading(false);
    setResult("");
    setGeneratedFile(null);
    setError("");
    setSuccess("");
    setPreviewOpen(false);
    setPreviewFullscreen(false);
    setOutputPreviewFullscreen(false);

    setInputFormat(initialInputFormat ?? null);
    setOutputFormat(
      initialOutputFormat ??
      (initialInputFormat ? conversionOptions[initialInputFormat][0] : "")
    );

    if (fileRef.current) fileRef.current.value = "";

    if (initialInputFormat) {
      const validOutput =
        initialOutputFormat &&
        conversionOptions[initialInputFormat].includes(initialOutputFormat);

      if (!validOutput) {
        setOutputFormat(conversionOptions[initialInputFormat][0] || "");
      }
    }
  };

  useEffect(() => {
    if (!previewOpen && !outputPreviewFullscreen) return;

    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        if (previewFullscreen) {
          setPreviewFullscreen(false);
        } else if (previewOpen) {
          setPreviewOpen(false);
        } else {
          setOutputPreviewFullscreen(false);
        }
      }
    };

    window.addEventListener("keydown", handleKeyDown);

    return () => {
      document.body.style.overflow = previousOverflow;
      window.removeEventListener("keydown", handleKeyDown);
    };
  }, [previewOpen, previewFullscreen, outputPreviewFullscreen]);

  const renderMarkdown = (markdown: string) => (
    <article
      style={{ fontSize: `${fontSize}pt`, lineHeight: 1.5 }}
      className="break-words font-sans text-slate-700"
    >
      <ReactMarkdown
        remarkPlugins={[remarkGfm, remarkBreaks]}
        components={markdownPreviewComponents(fontSize)}
      >
        {markdownForPreview(markdown)}
      </ReactMarkdown>
    </article>
  );

  const renderFontSizeControls = () => (
    <div className="flex items-center gap-2 rounded-lg border border-slate-200 bg-white px-2 py-1 text-slate-700">
      <label className="whitespace-nowrap text-xs font-semibold">
        Font size
        <select
          aria-label="Output font size"
          value={fontSize}
          onChange={(event) => void updateFontSize(Number(event.target.value))}
          disabled={busy}
          className="ml-2 rounded border-0 bg-transparent py-1 pr-1 text-sm font-semibold text-slate-800 outline-none focus:ring-2 focus:ring-cyan-500 disabled:opacity-50"
        >
          {Array.from(
            {
              length:
                (outputFormat === "pptx" ? 24 : 32) -
                (outputFormat === "pptx" ? 10 : 8) + 1
            },
            (_, index) => index + (outputFormat === "pptx" ? 10 : 8),
          ).map((size) => (
            <option key={size} value={size}>
              {size} pt
            </option>
          ))}
        </select>
      </label>
    </div>
  );

  return (
    <section id="converter" className="w-full">
      <div className="mx-auto w-full max-w-none">
        {/* CONVERTER CARD */}
        <div className="glass-card glow-cyan rounded-2xl p-5 sm:p-6">
          {/* UPLOAD */}
          <div className="mb-6">
            <div className="mb-3 flex items-center justify-between">
              <p className="text-sm font-semibold text-white">
                1. Upload your document
              </p>
              <span className="text-xs text-slate-500">Step 01</span>
            </div>

            {!file ? (
              <button
                type="button"
                onClick={() => fileRef.current?.click()}
                onDragOver={(e) => {
                  e.preventDefault();
                  setDragging(true);
                }}
                onDragLeave={() => setDragging(false)}
                onDrop={(e) => {
                  e.preventDefault();
                  setDragging(false);
                  selectFile(e.dataTransfer.files[0]);
                }}
                className={`flex min-h-56 w-full flex-col items-center justify-center rounded-2xl border-2 border-dashed p-6 text-center transition-all ${
                  dragging
                    ? "border-cyan-400 bg-cyan-400/10"
                    : "border-slate-700 bg-slate-950/40 hover:border-cyan-400/60"
                }`}
              >
                <div className="mb-4 flex h-14 w-14 items-center justify-center rounded-2xl bg-cyan-400/10 text-cyan-400">
                  <Upload size={27} />
                </div>

                <p className="text-base font-semibold text-white">
                  Drop your file here
                </p>

                <p className="mt-2 text-sm text-slate-400">
                  or click to browse your device
                </p>

                <span className="mt-4 rounded-lg border border-white/10 px-3 py-2 text-[11px] text-slate-300">
                  DOCX · PDF · PPTX · MD · Maximum 20 MB (28 MB for Markdown)
                </span>
              </button>
            ) : (
              <div className="flex items-center gap-3 rounded-xl border border-cyan-400/20 bg-cyan-400/5 p-4">
                <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-xl bg-cyan-400/10 text-cyan-400">
                  <FileText size={23} />
                </div>

                <div className="min-w-0 flex-1">
                  <p className="truncate text-sm font-medium text-white">
                    {file.name}
                  </p>

                  <p className="mt-1 text-xs text-slate-400">
                    {formatNames[inputFormat!]} · {formatFileSize(file.size)}
                  </p>
                </div>

                <button
                  type="button"
                  onClick={reset}
                  aria-label="Remove file"
                  title="Remove file"
                  className="rounded-lg p-2 text-slate-400 transition hover:bg-white/10 hover:text-white"
                >
                  <X size={19} />
                </button>
              </div>
            )}

            {file && previewMarkdown !== null && (
              <div className="mt-3 grid grid-cols-2 gap-2 sm:grid-cols-4">
                <div className="min-w-0 rounded-lg border border-white/10 bg-slate-950/50 p-3 sm:col-span-2">
                  <p className="text-[11px] uppercase tracking-wide text-slate-500">
                    Loaded file
                  </p>
                  <p
                    className="mt-1 truncate text-sm font-medium text-slate-200"
                    title={file.name}
                  >
                    {file.name}
                  </p>
                </div>

                <div className="rounded-lg border border-white/10 bg-slate-950/50 p-3">
                  <p className="text-[11px] uppercase tracking-wide text-slate-500">
                    File size
                  </p>
                  <p className="mt-1 text-sm font-medium text-slate-200">
                    {formatFileSize(file.size)}
                  </p>
                </div>

                <div className="rounded-lg border border-white/10 bg-slate-950/50 p-3">
                  <p className="text-[11px] uppercase tracking-wide text-slate-500">
                    Text characters
                  </p>
                  <p className="mt-1 text-sm font-medium text-slate-200">
                    {Array.from(previewMarkdown).length.toLocaleString()}
                  </p>
                </div>

                <div className="rounded-lg border border-white/10 bg-slate-950/50 p-3">
                  <p className="text-[11px] uppercase tracking-wide text-slate-500">
                    Lines
                  </p>
                  <p className="mt-1 text-sm font-medium text-slate-200">
                    {previewMarkdown.split(/\r\n|\r|\n/).length.toLocaleString()}
                  </p>
                </div>
              </div>
            )}

            {file && previewLoading && (
              <p className="mt-3 flex items-center gap-2 text-xs text-cyan-300" role="status">
                <RefreshCw size={14} className="animate-spin" />
                Preparing file preview and text information...
              </p>
            )}

            {!file && !customMarkdownMode && !initialInputFormat && (
              <button
                type="button"
                onClick={startCustomMarkdown}
                className="mt-3 rounded-lg border border-cyan-400/30 bg-cyan-400/10 px-4 py-2.5 text-sm font-semibold text-cyan-300 transition hover:border-cyan-400 hover:bg-cyan-400/20"
              >
                Or write custom Markdown
              </button>
            )}
          </div>

          <input
            ref={fileRef}
            type="file"
            accept=".docx,.pdf,.pptx,.md"
            hidden
            onChange={(e) => {
              selectFile(e.target.files?.[0]);
              e.target.value = "";
            }}
          />

          {inputFormat === "md" && (
            <section className="mb-6" aria-label="Markdown content editor">
              <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
                <div>
                  <h2 className="text-sm font-semibold text-white">
                    2. Markdown content
                  </h2>
                  <p className="mt-1 text-xs text-slate-400">
                    Edit the uploaded Markdown or write your own content to convert.
                  </p>
                </div>

                <button
                  type="button"
                  onClick={loadMarkdownExample}
                  disabled={previewLoading || busy}
                  className="rounded-lg border border-cyan-400/30 bg-cyan-400/10 px-4 py-2 text-xs font-semibold text-cyan-300 transition hover:border-cyan-400 hover:bg-cyan-400/20 disabled:cursor-not-allowed disabled:opacity-50"
                >
                  Load example
                </button>
              </div>

              <textarea
                aria-label="Markdown source content"
                value={previewMarkdown ?? ""}
                onChange={(event) => editMarkdown(event.target.value)}
                disabled={previewLoading || busy}
                placeholder={"# Your document title\n\nWrite Markdown here. Use fenced code blocks for code snippets."}
                spellCheck={false}
                rows={14}
                className="w-full resize-y rounded-xl border border-white/10 bg-slate-950 px-4 py-3 font-mono text-sm leading-6 text-slate-200 outline-none transition placeholder:text-slate-600 focus:border-cyan-400 disabled:cursor-not-allowed disabled:opacity-60"
              />

              <div className="mt-2 flex flex-wrap gap-x-5 gap-y-1 text-xs text-slate-400" aria-live="polite">
                <span>
                  Text characters:{" "}
                  <strong className="font-medium text-slate-200">
                    {Array.from(previewMarkdown ?? "").length.toLocaleString()}
                  </strong>
                </span>

                <span>
                  Lines:{" "}
                  <strong className="font-medium text-slate-200">
                    {(previewMarkdown ?? "").split(/\r\n|\r|\n/).length.toLocaleString()}
                  </strong>
                </span>
              </div>
            </section>
          )}

          {/* OUTPUT FORMAT */}
          <div className="mb-6">
            <div className="mb-3 flex items-center justify-between">
              <p className="text-sm font-semibold text-white">
                {inputFormat === "md" ? "3." : "2."} Choose output format
              </p>
              <span className="text-xs text-slate-500">Step 02</span>
            </div>

            <div className="relative">
              <select
                value={outputFormat}
                onChange={(e) => {
                  fontUpdateIdRef.current += 1;
                  setFontSizeUpdating(false);

                  const nextFormat = e.target.value as Format;
                  setOutputFormat(nextFormat);

                  if (nextFormat === "pptx") {
                    setFontSize((current) =>
                      Math.min(24, Math.max(10, current))
                    );
                  }

                  setResult("");
                  setGeneratedFile(null);
                  setSuccess("");
                  setError("");
                  setPreviewOpen(false);
                  setOutputPreviewFullscreen(false);
                }}
                disabled={!inputFormat || busy || outputOptions.length === 0}
                className="w-full appearance-none rounded-xl border border-white/10 bg-slate-950 px-4 py-3.5 pr-12 text-sm text-white outline-none transition focus:border-cyan-400 disabled:cursor-not-allowed disabled:opacity-50"
              >
                {!inputFormat ? (
                  <option value="">Upload a document first</option>
                ) : (
                  outputOptions.map((format) => (
                    <option key={format} value={format}>
                      {formatNames[format]} ({extensions[format]})
                    </option>
                  ))
                )}
              </select>

              <ChevronDown
                className="pointer-events-none absolute right-4 top-1/2 -translate-y-1/2 text-cyan-400"
                size={19}
              />
            </div>

            {inputFormat && outputFormat && (
              <div className="mt-3 flex items-center justify-between rounded-lg bg-slate-950/60 px-3 py-2.5">
                <span className="text-xs font-medium text-slate-400">
                  Conversion
                </span>
                <span className="text-xs font-semibold text-cyan-400">
                  {extensions[inputFormat]} → {extensions[outputFormat]}
                </span>
              </div>
            )}
          </div>

          {(file || customMarkdownMode) && inputFormat && outputFormat && previewMarkdown !== null && (
            <section
              aria-label="Source and output preview"
              className={`mb-6 ${outputPreviewFullscreen ? "fixed inset-0 z-[90] flex h-dvh flex-col overflow-hidden bg-[#070b14] p-3 sm:p-5" : ""}`}
            >
              <div className={`grid gap-4 ${outputPreviewFullscreen ? "min-h-0 flex-1 lg:grid-cols-2" : "lg:grid-cols-2"}`}>
                <div className={`flex min-w-0 flex-col overflow-hidden rounded-xl border border-white/10 bg-slate-950/60 ${outputPreviewFullscreen ? "min-h-0" : ""}`}>
                  <div className="flex items-center justify-between border-b border-white/10 px-4 py-3">
                    <div>
                      <h3 className="text-sm font-semibold text-white">
                        {inputFormat === "md" ? "Markdown source" : "Extracted content"}
                      </h3>
                      <p className="mt-1 text-xs text-slate-500">
                        {file?.name ?? "Custom Markdown"}
                      </p>
                    </div>

                    <span className="rounded-md bg-white/5 px-2 py-1 text-[11px] font-semibold text-slate-400">
                      {extensions[inputFormat]}
                    </span>
                  </div>

                  <pre
                    style={{ fontSize: `${fontSize}pt`, lineHeight: 1.5 }}
                    className={`overflow-auto whitespace-pre-wrap break-words p-4 font-mono text-slate-300 ${outputPreviewFullscreen ? "min-h-0 flex-1" : "max-h-[34rem]"}`}
                  >
                    {previewMarkdown}
                  </pre>
                </div>

                <div className="relative flex min-h-0 min-w-0 flex-col overflow-hidden rounded-xl border border-cyan-400/20 bg-[#111c2d]">
                  <div className="flex shrink-0 flex-wrap items-center justify-between gap-3 border-b border-white/10 px-4 py-3">
                    <div>
                      <h3 className="text-sm font-semibold text-white">
                        {formatNames[outputFormat]} preview
                      </h3>
                      <p className="mt-1 text-xs text-slate-400">
                        {extensions[outputFormat]} ·{" "}
                        {fontSizeUpdating
                          ? "Updating download…"
                          : outputFormat === "md"
                            ? "Font size adjusts this preview only"
                            : "Font size is applied to the downloaded file"}
                      </p>
                    </div>

                    <div className="flex items-center gap-2">
                      {renderFontSizeControls()}

                      <button
                        type="button"
                        onClick={() => setOutputPreviewFullscreen((current) => !current)}
                        aria-label={outputPreviewFullscreen ? "Exit full screen output preview" : "View output preview full screen"}
                        title={outputPreviewFullscreen ? "Exit full screen" : "Full screen"}
                        className="rounded-lg border border-white/10 p-2 text-slate-300 transition hover:border-cyan-400/50 hover:text-cyan-300"
                      >
                        {outputPreviewFullscreen
                          ? <Minimize2 size={17} />
                          : <Maximize2 size={17} />}
                      </button>
                    </div>
                  </div>

                  <div className={`min-h-0 flex-1 overflow-auto bg-[#080d17] p-3 sm:p-5 ${outputPreviewFullscreen ? "" : "max-h-[38rem]"}`}>
                    {outputFormat === "md" ? (
                      <div className="mx-auto min-h-64 max-w-3xl overflow-auto rounded-lg border border-slate-200 bg-white px-6 py-8 shadow-xl sm:px-10 sm:py-12">
                        {renderMarkdown(previewMarkdown)}
                      </div>
                    ) : outputFormat === "pptx" ? (
                      <div className="mx-auto min-h-64 max-w-4xl overflow-auto rounded-lg border border-slate-300 bg-white p-5 shadow-xl sm:p-8">
                        <div className="mb-5 h-1.5 rounded-full bg-slate-500" />
                        {renderMarkdown(previewMarkdown)}
                      </div>
                    ) : (
                      <div className="mx-auto min-h-64 max-w-3xl overflow-auto rounded-lg border border-slate-200 bg-white px-6 py-8 shadow-xl sm:px-10 sm:py-12">
                        {renderMarkdown(previewMarkdown)}
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </section>
          )}

          {/* ERROR */}
          {error && (
            <div className="mb-5 flex items-start gap-2 rounded-lg border border-red-500/20 bg-red-500/10 p-3 text-sm text-red-300">
              <AlertCircle size={18} className="shrink-0" />
              {error}
            </div>
          )}

          {/* SUCCESS */}
          {success && (
            <div className="mb-5 flex items-start gap-2 rounded-lg border border-emerald-500/20 bg-emerald-500/10 p-3 text-sm text-emerald-300">
              <CheckCircle2 size={18} className="shrink-0" />
              {success}
            </div>
          )}

          {/* CONVERT BUTTON */}
          <button
            type="button"
            onClick={handleConvert}
            disabled={
              (!file && !customMarkdownMode) ||
              !outputFormat ||
              busy ||
              previewLoading ||
              outputOptions.length === 0 ||
              (inputFormat === "md" && !previewMarkdown?.trim())
            }
            className="flex w-full items-center justify-center gap-2 rounded-xl bg-cyan-400 px-6 py-4 text-sm font-bold text-slate-950 transition-all hover:bg-cyan-300 disabled:cursor-not-allowed disabled:opacity-40"
          >
            {busy ? (
              <>
                <RefreshCw size={19} className="animate-spin" />
                Converting...
              </>
            ) : (
              <>
                Convert Document <span>→</span>
              </>
            )}
          </button>

          {/* RESULT */}
          {result && (
            <div className="mt-6">
              <div className="flex flex-col gap-4 rounded-xl border border-cyan-400/20 bg-cyan-400/4 p-4 sm:flex-row sm:items-center sm:justify-between">
                <div className="flex items-center gap-3">
                  <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-emerald-400/10 text-emerald-400">
                    <CheckCircle2 size={21} />
                  </div>

                  <div>
                    <h3 className="text-sm font-semibold text-white">
                      Conversion completed
                    </h3>
                    <p className="mt-1 text-xs text-slate-400">
                      {formatNames[outputFormat as Format]} is ready to download.
                    </p>
                  </div>
                </div>

                <div className="flex flex-wrap gap-2">
                  <button
                    type="button"
                    onClick={() => {
                      setPreviewMode("source");
                      setPreviewFullscreen(false);
                      setPreviewOpen(true);
                    }}
                    className="flex items-center justify-center gap-2 rounded-lg border border-cyan-400/30 bg-cyan-400/10 px-4 py-2.5 text-sm font-semibold text-cyan-300 transition hover:border-cyan-400 hover:bg-cyan-400/20"
                  >
                    <Eye size={17} /> Preview
                  </button>

                  <button
                    type="button"
                    onClick={handleDownload}
                    disabled={busy || fontSizeUpdating || (inputFormat === "md" && !generatedFile)}
                    className="flex items-center justify-center gap-2 rounded-lg bg-cyan-400 px-4 py-2.5 text-sm font-semibold text-slate-950 transition hover:bg-cyan-300 disabled:cursor-not-allowed disabled:opacity-50"
                  >
                    <Download size={17} />
                    Download
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* PRIVACY NOTE */}
        <div className="mt-4 flex items-center justify-center gap-2 text-center text-xs text-slate-500">
          <CheckCircle2 size={14} className="text-cyan-400" />
          Your files are processed in your browser and are not uploaded to our servers.
        </div>
      </div>

      {/* FULL SCREEN PREVIEW */}
      {previewOpen && result && (
        <div className={`fixed inset-0 z-[100] flex items-center justify-center bg-black/80 ${previewFullscreen ? "p-0" : "p-3 sm:p-8"}`}>
          <div className={`flex h-full w-full flex-col overflow-hidden bg-[#070b14] ${previewFullscreen ? "max-w-none rounded-none" : "max-h-[92vh] max-w-6xl rounded-xl border border-white/10 shadow-2xl"}`}>
            <div className="flex shrink-0 items-center justify-between border-b border-white/10 bg-[#111c2d] px-5 py-4 sm:px-8">
              <div className="flex min-w-0 items-center gap-3">
                <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-cyan-400/10 text-cyan-400">
                  <FileText size={21} />
                </div>

                <div className="min-w-0">
                  <h2 className="font-semibold text-white">Document output</h2>
                  <p className="truncate text-xs text-slate-400">
                    {outputFormat
                      ? `${(file?.name ?? "custom-markdown.md").replace(/\.[^.]+$/, "")}${extensions[outputFormat]}`
                      : ""}
                  </p>
                </div>
              </div>

              <div className="flex items-center gap-2">
                {renderFontSizeControls()}

                <button
                  type="button"
                  onClick={() => setPreviewFullscreen((current) => !current)}
                  aria-label={previewFullscreen ? "Exit full screen preview" : "View preview full screen"}
                  title={previewFullscreen ? "Exit full screen" : "Full screen"}
                  className="rounded-lg border border-white/10 p-2 text-slate-300 transition hover:border-cyan-400/50 hover:text-cyan-300"
                >
                  {previewFullscreen
                    ? <Minimize2 size={18} />
                    : <Maximize2 size={18} />}
                </button>

                <button
                  type="button"
                  onClick={() => setPreviewOpen(false)}
                  aria-label="Close preview"
                  className="rounded-lg border border-white/10 p-2 text-slate-400 transition hover:border-red-400/30 hover:bg-red-400/10 hover:text-red-400"
                >
                  <X size={20} />
                </button>
              </div>
            </div>

            {outputFormat === "md" && (
              <div className="flex shrink-0 items-center justify-between gap-3 border-b border-white/5 px-5 py-3 sm:px-8">
                <div className="flex items-center gap-1 rounded-lg bg-white/5 p-1">
                  <button
                    type="button"
                    onClick={() => setPreviewMode("source")}
                    aria-pressed={previewMode === "source"}
                    className={`rounded-md px-3 py-1.5 text-xs font-semibold transition ${
                      previewMode === "source"
                        ? "bg-cyan-400 text-slate-950"
                        : "text-slate-400 hover:text-white"
                    }`}
                  >
                    .md source
                  </button>

                  <button
                    type="button"
                    onClick={() => setPreviewMode("rendered")}
                    aria-pressed={previewMode === "rendered"}
                    className={`rounded-md px-3 py-1.5 text-xs font-semibold transition ${
                      previewMode === "rendered"
                        ? "bg-cyan-400 text-slate-950"
                        : "text-slate-400 hover:text-white"
                    }`}
                  >
                    Rendered
                  </button>
                </div>

                <span className="text-xs text-slate-500">
                  Download contains the Markdown source
                </span>
              </div>
            )}

            <div className="min-h-0 flex-1 overflow-y-auto bg-[#080d17] p-4 sm:p-8 lg:p-12">
              <div className="mx-auto min-h-full w-full max-w-5xl rounded-xl border border-slate-200/10 bg-[#f8fafc] px-6 py-8 text-slate-800 shadow-2xl sm:px-12 sm:py-12">
                {outputFormat === "md" ? (
                  previewMode === "source" ? (
                    <pre
                      style={{ fontSize: `${fontSize}pt`, lineHeight: 1.5 }}
                      className="whitespace-pre-wrap break-words font-mono text-slate-700"
                    >
                      {result}
                    </pre>
                  ) : (
                    renderMarkdown(result)
                  )
                ) : (
                  renderMarkdown(result)
                )}
              </div>
            </div>

            <div className="flex shrink-0 items-center justify-between gap-3 border-t border-white/10 bg-[#111c2d] px-5 py-4 sm:px-8">
              <span className="truncate text-xs text-slate-400">
                {formatNames[outputFormat as Format]} ·{" "}
                {`${(file?.name ?? "custom-markdown.md").replace(/\.[^.]+$/, "")}${extensions[outputFormat as Format]}`}
              </span>

              <button
                type="button"
                onClick={handleDownload}
                disabled={busy || fontSizeUpdating || (inputFormat === "md" && !generatedFile)}
                className="flex items-center justify-center gap-2 rounded-lg bg-cyan-400 px-5 py-2.5 text-sm font-bold text-slate-950 transition hover:bg-cyan-300 disabled:cursor-not-allowed disabled:opacity-50"
              >
                {busy
                  ? <RefreshCw size={17} className="animate-spin" />
                  : <Download size={17} />}
                {busy ? "Updating..." : `Download ${extensions[outputFormat as Format]}`}
              </button>
            </div>
          </div>
        </div>
      )}
    </section>
  );
}
