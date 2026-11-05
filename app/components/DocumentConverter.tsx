
"use client";

import { useRef, useState, useEffect } from "react";
import { useSearchParams } from "next/navigation";
import {
  Upload, FileText, Download, RefreshCw, X,
  CheckCircle2, AlertCircle, ChevronDown, Eye
} from "lucide-react";
import ReactMarkdown, { type Components } from "react-markdown";
import remarkGfm from "remark-gfm";

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

const markdownPreviewComponents: Components = {
  h1: ({ children }) => (
    <h1 className="mb-5 mt-8 border-b border-slate-200 pb-3 text-3xl font-bold text-slate-900 first:mt-0">
      {children}
    </h1>
  ),
  h2: ({ children }) => (
    <h2 className="mb-4 mt-7 text-2xl font-bold text-slate-900">
      {children}
    </h2>
  ),
  h3: ({ children }) => (
    <h3 className="mb-3 mt-6 text-xl font-semibold text-slate-900">
      {children}
    </h3>
  ),
  h4: ({ children }) => (
    <h4 className="mb-3 mt-5 text-lg font-semibold text-slate-900">
      {children}
    </h4>
  ),
  h5: ({ children }) => (
    <h5 className="mb-2 mt-4 text-base font-semibold text-slate-900">
      {children}
    </h5>
  ),
  h6: ({ children }) => (
    <h6 className="mb-2 mt-4 text-sm font-semibold text-slate-900">
      {children}
    </h6>
  ),
  p: ({ children }) => <p className="mb-4 leading-7">{children}</p>,
  ul: ({ children }) => (
    <ul className="mb-5 list-disc space-y-1 pl-7">{children}</ul>
  ),
  ol: ({ children }) => (
    <ol className="mb-5 list-decimal space-y-1 pl-7">{children}</ol>
  ),
  li: ({ children }) => <li className="pl-1">{children}</li>,
  blockquote: ({ children }) => (
    <blockquote className="my-5 border-l-4 border-cyan-500 bg-slate-100 py-2 pl-4 text-slate-600">
      {children}
    </blockquote>
  ),
  a: ({ children, href }) => (
    <a
      href={href}
      className="text-cyan-700 underline decoration-cyan-300 underline-offset-2"
      target="_blank"
      rel="noreferrer"
    >
      {children}
    </a>
  ),
  strong: ({ children }) => (
    <strong className="font-bold text-slate-900">{children}</strong>
  ),
  del: ({ children }) => <del className="text-slate-500">{children}</del>,
  code: ({ children, className }) => (
    <code
      className={`${className ?? ""} rounded bg-slate-200 px-1.5 py-0.5 font-mono text-[0.9em] text-slate-800`}
    >
      {children}
    </code>
  ),
  pre: ({ children }) => (
    <pre className="mb-5 overflow-x-auto rounded-lg bg-slate-900 p-4 font-mono text-sm leading-6 text-slate-100 [&_code]:bg-transparent [&_code]:p-0 [&_code]:text-inherit">
      {children}
    </pre>
  ),
  table: ({ children }) => (
    <div className="mb-5 overflow-x-auto">
      <table className="w-full border-collapse text-left text-sm">
        {children}
      </table>
    </div>
  ),
  thead: ({ children }) => (
    <thead className="bg-cyan-50 text-slate-900">{children}</thead>
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
  hr: () => <hr className="my-7 border-slate-300" />,
};

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

export default function DocumentConverter({
  initialInputFormat,
  initialOutputFormat
}: DocumentConverterProps = {}) {
  const searchParams = useSearchParams();
  const fileRef = useRef<HTMLInputElement>(null);

  const [file, setFile] = useState<File | null>(null);
  const [inputFormat, setInputFormat] = useState<Format | null>(null);
  const [outputFormat, setOutputFormat] = useState<Format | "">("");
  const [result, setResult] = useState("");
  const [generatedFile, setGeneratedFile] = useState<GeneratedDocument | null>(null);
  const [busy, setBusy] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState("");
  const [previewOpen, setPreviewOpen] = useState(false);
  const [previewMode, setPreviewMode] = useState<"rendered" | "source">("rendered");

  const preferredInput = searchParams.get("from") ?? initialInputFormat;
  const preferredOutput = searchParams.get("to") ?? initialOutputFormat;

  useEffect(() => {
    const validInput = validFormats.includes(preferredInput as Format)
      ? preferredInput as Format
      : null;

    const validOutput = validFormats.includes(preferredOutput as Format)
      ? preferredOutput as Format
      : null;

    if (!validInput) {
      setInputFormat(null);
      setOutputFormat("");
      return;
    }

    setInputFormat(validInput);

    const selectedOutput =
      validOutput && conversionOptions[validInput].includes(validOutput)
        ? validOutput
        : conversionOptions[validInput][0];

    setOutputFormat(selectedOutput || "");
    setFile(null);
    setResult("");
    setGeneratedFile(null);
    setSuccess("");
    setError("");
    setPreviewOpen(false);

    if (fileRef.current) fileRef.current.value = "";
  }, [preferredInput, preferredOutput]);

  const outputOptions = inputFormat
    ? conversionOptions[inputFormat]
    : [];

  const selectFile = (selectedFile?: File) => {
    setError("");
    setSuccess("");
    setResult("");
    setGeneratedFile(null);
    setPreviewOpen(false);

    if (!selectedFile) return;

    const detectedFormat = detectFormat(selectedFile.name);

    if (!detectedFormat) {
      setError("Unsupported file format. Please upload DOCX, PDF, PPTX or MD.");
      return;
    }

    if (selectedFile.size === 0) {
      setError("The selected file is empty.");
      return;
    }

    if (selectedFile.size > 20 * 1024 * 1024) {
      setError("File size must not exceed 20 MB.");
      return;
    }

    if (inputFormat && detectedFormat !== inputFormat) {
      setError(
        `Please upload a ${formatNames[inputFormat]} (${extensions[inputFormat]}) file.`
      );
      return;
    }

    setFile(selectedFile);
    setInputFormat(detectedFormat);

    setOutputFormat((currentOutput) =>
      currentOutput && conversionOptions[detectedFormat].includes(currentOutput)
        ? currentOutput
        : conversionOptions[detectedFormat][0] || ""
    );
  };

  const handleConvert = async () => {
    if (!file || !inputFormat || !outputFormat) {
      setError("Please upload a file and select an output format.");
      return;
    }

    if (!conversionOptions[inputFormat].includes(outputFormat)) {
      setError("This conversion is not supported.");
      return;
    }

    setBusy(true);
    setError("");
    setSuccess("");
    setResult("");
    setGeneratedFile(null);
    setPreviewOpen(false);

    try {
      if (inputFormat === "docx" && outputFormat === "md") {
        const markdown = await wordToMarkdown(file);
        setResult(markdown);
      } else if (inputFormat === "pdf" && outputFormat === "md") {
        const markdown = await pdfToMarkdown(file);
        setResult(markdown);
      } else if (inputFormat === "pptx" && outputFormat === "md") {
        const markdown = await convertPowerPointToMarkdown(file);
        setResult(markdown);
      } else if (inputFormat === "md" && outputFormat === "docx") {
        const markdown = await file.text();
        const wordFile = await convertMarkdownToWord(markdown);

        setResult(markdown);
        setGeneratedFile({ blob: wordFile, format: "docx" });
      } else if (inputFormat === "md" && outputFormat === "pdf") {
        const markdown = await file.text();
        const pdfFile = await convertMarkdownToPdf(markdown);

        setResult(markdown);
        setGeneratedFile({ blob: pdfFile, format: "pdf" });
      } else if (inputFormat === "md" && outputFormat === "pptx") {
        const markdown = await file.text();
        const powerpointFile = await convertMarkdownToPowerPoint(markdown);

        setResult(markdown);
        setGeneratedFile({ blob: powerpointFile, format: "pptx" });
      } else {
        throw new Error("This conversion has not been implemented yet.");
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

  const handleDownload = async () => {
    if (!file) return;

    if (generatedFile) {
      if (generatedFile.format === "docx") {
        await downloadWordFile(generatedFile.blob, file.name);
      } else {
        downloadConvertedFile(
          generatedFile.blob,
          file.name,
          generatedFile.format === "pdf" ? ".pdf" : ".pptx"
        );
      }
      return;
    }

    if (outputFormat === "md" && result && inputFormat) {
      downloadTextFile(
        result,
        markdownOutputFilename(file.name, inputFormat),
        "text/markdown;charset=utf-8"
      );
    }
  };

  const reset = () => {
    setFile(null);
    setResult("");
    setGeneratedFile(null);
    setError("");
    setSuccess("");
    setPreviewOpen(false);

    if (inputFormat) {
      const validOutput =
        outputFormat && conversionOptions[inputFormat].includes(outputFormat);

      if (!validOutput) {
        setOutputFormat(conversionOptions[inputFormat][0] || "");
      }
    }

    if (fileRef.current) fileRef.current.value = "";
  };

  useEffect(() => {
    if (!previewOpen) return;

    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        setPreviewOpen(false);
      }
    };

    window.addEventListener("keydown", handleKeyDown);

    return () => {
      document.body.style.overflow = previousOverflow;
      window.removeEventListener("keydown", handleKeyDown);
    };
  }, [previewOpen]);

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
              <span className="text-xs text-slate-500">
                Step 01
              </span>
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
                  DOCX · PDF · PPTX · MD · Maximum 20 MB
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
                    {formatNames[inputFormat!]} ·{" "}
                    {(file.size / 1024 / 1024).toFixed(2)} MB
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

          {/* OUTPUT FORMAT */}
          <div className="mb-6">
            <div className="mb-3 flex items-center justify-between">
              <p className="text-sm font-semibold text-white">
                2. Choose output format
              </p>
              <span className="text-xs text-slate-500">
                Step 02
              </span>
            </div>

            <div className="relative">
              <select
                value={outputFormat}
                onChange={(e) => {
                  setOutputFormat(e.target.value as Format);
                  setResult("");
                  setGeneratedFile(null);
                  setSuccess("");
                  setError("");
                  setPreviewOpen(false);
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
            disabled={!file || !outputFormat || busy || outputOptions.length === 0}
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
                      setPreviewOpen(true);
                    }}
                    className="flex items-center justify-center gap-2 rounded-lg border border-cyan-400/30 bg-cyan-400/10 px-4 py-2.5 text-sm font-semibold text-cyan-300 transition hover:border-cyan-400 hover:bg-cyan-400/20"
                  >
                    <Eye size={17} /> Preview
                  </button>

                  <button
                    type="button"
                    onClick={handleDownload}
                    className="flex items-center justify-center gap-2 rounded-lg bg-cyan-400 px-4 py-2.5 text-sm font-semibold text-slate-950 transition hover:bg-cyan-300"
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
        <div className="fixed inset-0 z-[100] flex h-screen w-screen flex-col bg-[#070b14]">
          <div className="flex shrink-0 items-center justify-between border-b border-white/10 bg-[#111c2d] px-5 py-4 sm:px-8">
            <div className="flex min-w-0 items-center gap-3">
              <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-cyan-400/10 text-cyan-400">
                <FileText size={21} />
              </div>

              <div className="min-w-0">
                <h2 className="font-semibold text-white">Document output</h2>
                <p className="truncate text-xs text-slate-400">
                  {file && outputFormat
                    ? `${file.name.replace(/\.[^.]+$/, "")}${extensions[outputFormat]}`
                    : ""}
                </p>
              </div>
            </div>
            <button
              type="button"
              onClick={() => setPreviewOpen(false)}
              aria-label="Close preview"
              className="rounded-lg border border-white/10 p-2 text-slate-400 transition hover:border-red-400/30 hover:bg-red-400/10 hover:text-red-400"
            >
              <X size={20} />
            </button>
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
                  <pre className="whitespace-pre-wrap break-words font-mono text-sm leading-7 text-slate-700">
                    {result}
                  </pre>
                ) : (
                  <article className="break-words font-serif text-sm leading-7 text-slate-700">
                    <ReactMarkdown
                      remarkPlugins={[remarkGfm]}
                      components={markdownPreviewComponents}
                    >
                      {result}
                    </ReactMarkdown>
                  </article>
                )
              ) : (
                <div className="flex min-h-64 flex-col items-center justify-center text-center">
                  <FileText size={42} className="mb-4 text-cyan-600" />
                  <h3 className="text-xl font-semibold text-slate-900">
                    {formatNames[outputFormat as Format]} created
                  </h3>
                  <p className="mt-2 max-w-md text-sm text-slate-600">
                    The converted file is ready. Download it to open the
                    {` ${extensions[outputFormat as Format]}`} document.
                  </p>
                </div>
              )}
            </div>
          </div>

          <div className="flex shrink-0 items-center justify-between gap-3 border-t border-white/10 bg-[#111c2d] px-5 py-4 sm:px-8">
            <span className="truncate text-xs text-slate-400">
              {formatNames[outputFormat as Format]} ·{" "}
              {file
                ? `${file.name.replace(/\.[^.]+$/, "")}${extensions[outputFormat as Format]}`
                : ""}
            </span>
            <button
              type="button"
              onClick={handleDownload}
              className="flex items-center justify-center gap-2 rounded-lg bg-cyan-400 px-5 py-2.5 text-sm font-bold text-slate-950 transition hover:bg-cyan-300"
            >
              <Download size={17} />
              Download {extensions[outputFormat as Format]}
            </button>
          </div>
        </div>
      )}
    </section>
  );
}
