
"use client";

import { useRef, useState, useEffect } from "react";
import { useSearchParams } from "next/navigation";
import {
  Upload, FileText, Download, RefreshCw, X,
  CheckCircle2, AlertCircle, ChevronDown, Eye, Maximize2
} from "lucide-react";

import {
  downloadTextFile,
  downloadWordFile,
  markdownFilename,
  pdfMarkdownFilename,
  powerpointMarkdownFilename
} from "../services/downloadservice";

import { wordToMarkdown } from "../converters/word-to-markdown";
import { pdfToMarkdown } from "../converters/pdfToMarkdown";
import { convertPowerPointToMarkdown } from "../converters/powerpointToMarkdown";
import { convertMarkdownToWord } from "../converters/markdownToWord";

type Format = "docx" | "pdf" | "pptx" | "md";

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

// These are the conversions currently implemented.
const conversionOptions: Record<Format, Format[]> = {
  docx: ["md"],
  pdf: ["md"],
  pptx: ["md"],
  md: ["docx"]
};

const detectFormat = (filename: string): Format | null => {
  const extension = filename.toLowerCase().split(".").pop();
  return validFormats.includes(extension as Format)
    ? extension as Format
    : null;
};

export default function DocumentConverter() {
  const searchParams = useSearchParams();
  const fileRef = useRef<HTMLInputElement>(null);

  const [file, setFile] = useState<File | null>(null);
  const [inputFormat, setInputFormat] = useState<Format | null>(null);
  const [outputFormat, setOutputFormat] = useState<Format | "">("");
  const [result, setResult] = useState("");
  const [busy, setBusy] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState("");
  const [previewOpen, setPreviewOpen] = useState(false);

  // Read the converter selected from the navbar.
  const preferredInput = searchParams.get("from");
  const preferredOutput = searchParams.get("to");

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

    // Check whether the uploaded file matches the navbar selection.
    if (inputFormat && detectedFormat !== inputFormat) {
      setError(
        `Please upload a ${formatNames[inputFormat]} (${extensions[inputFormat]}) file.`
      );
      return;
    }

    setFile(selectedFile);
    setInputFormat(detectedFormat);

    // Preserve navbar output selection when valid.
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
        const filename = file.name.replace(/\.md$/i, ".docx");
        await downloadWordFile(wordFile, filename);
        setSuccess("Your Word document has been generated and downloaded.");
      } else {
        throw new Error("This conversion has not been implemented yet.");
      }
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

  const handleDownload = () => {
    if (!result || !file || !inputFormat || !outputFormat) return;

    let filename = file.name.replace(/\.[^.]+$/, ".md");

    if (inputFormat === "docx") filename = markdownFilename(file.name);
    if (inputFormat === "pdf") filename = pdfMarkdownFilename(file.name);
    if (inputFormat === "pptx") filename = powerpointMarkdownFilename(file.name);

    downloadTextFile(result, filename);
  };

  const reset = () => {
    setFile(null);
    setResult("");
    setError("");
    setSuccess("");
    setPreviewOpen(false);

    // Keep the navbar-selected formats after removing a file.
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
      if (event.key === "Escape") setPreviewOpen(false);
    };

    window.addEventListener("keydown", handleKeyDown);

    return () => {
      document.body.style.overflow = previousOverflow;
      window.removeEventListener("keydown", handleKeyDown);
    };
  }, [previewOpen]);

  return (
    <section id="converter" className="container-main py-16">
      <div className="mx-auto max-w-3xl">
        <div className="mb-10 text-center">
          <span className="mb-4 inline-block rounded-full border border-cyan-400/20 bg-cyan-400/10 px-4 py-2 text-xs font-semibold uppercase tracking-widest text-cyan-400">
            Dynamic Document Converter
          </span>
          <h2 className="text-3xl font-bold tracking-tight text-white sm:text-5xl">
            Convert Your <span className="text-gradient">Documents</span>
          </h2>
          <p className="mx-auto mt-4 max-w-xl text-sm leading-7 text-slate-400 sm:text-base">
            Upload any supported document. We detect its format automatically and show the available conversions.
          </p>
        </div>

        <div className="glass-card glow-cyan p-6 sm:p-9">
          <div className="mb-6">
            <p className="mb-3 text-sm font-semibold text-white">
              1. Upload your document
            </p>

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
                className={`flex min-h-64 w-full flex-col items-center justify-center rounded-2xl border-2 border-dashed p-8 text-center transition-all ${
                  dragging
                    ? "border-cyan-400 bg-cyan-400/10"
                    : "border-slate-700 bg-slate-950/40 hover:border-cyan-400/60"
                }`}
              >
                <div className="mb-5 flex h-16 w-16 items-center justify-center rounded-2xl bg-cyan-400/10 text-cyan-400">
                  <Upload size={30} />
                </div>
                <p className="text-lg font-semibold text-white">
                  Drop your file here
                </p>
                <p className="mt-2 text-sm text-slate-400">
                  or click to browse your device
                </p>
                <span className="mt-5 rounded-lg border border-white/10 px-4 py-2 text-xs text-slate-300">
                  DOCX · PDF · PPTX · MD · Maximum 20 MB
                </span>
              </button>
            ) : (
              <div className="flex items-center gap-4 rounded-xl border border-cyan-400/20 bg-cyan-400/[0.05] p-5">
                <div className="flex h-12 w-12 shrink-0 items-center justify-center rounded-xl bg-cyan-400/10 text-cyan-400">
                  <FileText size={25} />
                </div>
                <div className="min-w-0 flex-1">
                  <p className="truncate font-medium text-white">{file.name}</p>
                  <p className="mt-1 text-xs text-slate-400">
                    Detected: {formatNames[inputFormat!]} ({extensions[inputFormat!]}) ·{" "}
                    {(file.size / 1024 / 1024).toFixed(2)} MB
                  </p>
                </div>
                <button
                  type="button"
                  onClick={reset}
                  aria-label="Remove file"
                  className="rounded-lg p-2 text-slate-400 hover:bg-white/10 hover:text-white"
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

          <div className="mb-6">
            <p className="mb-3 text-sm font-semibold text-white">
              2. Choose output format
            </p>

            <div className="relative">
              <select
                value={outputFormat}
                onChange={(e) => {
                  setOutputFormat(e.target.value as Format);
                  setResult("");
                  setSuccess("");
                  setError("");
                }}
                disabled={!inputFormat || busy || outputOptions.length === 0}
                className="w-full appearance-none rounded-xl border border-white/10 bg-slate-950 px-4 py-4 pr-12 text-sm text-white outline-none transition focus:border-cyan-400 disabled:cursor-not-allowed disabled:opacity-50"
              >
                {!inputFormat ? (
                  <option value="">Choose a converter first</option>
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
              <p className="mt-2 text-xs text-slate-500">
                {extensions[inputFormat]} → {extensions[outputFormat]}
              </p>
            )}
          </div>

          {error && (
            <div className="mb-5 flex items-start gap-2 rounded-lg border border-red-500/20 bg-red-500/10 p-3 text-sm text-red-300">
              <AlertCircle size={18} className="shrink-0" />
              {error}
            </div>
          )}

          {success && (
            <div className="mb-5 flex items-start gap-2 rounded-lg border border-emerald-500/20 bg-emerald-500/10 p-3 text-sm text-emerald-300">
              <CheckCircle2 size={18} className="shrink-0" />
              {success}
            </div>
          )}

          <button
            type="button"
            onClick={handleConvert}
            disabled={!file || !outputFormat || busy || outputOptions.length === 0}
            className="flex w-full items-center justify-center gap-2 rounded-xl bg-cyan-400 px-6 py-4 font-bold text-slate-950 transition-all hover:bg-cyan-300 disabled:cursor-not-allowed disabled:opacity-40"
          >
            {busy ? (
              <>
                <RefreshCw size={19} className="animate-spin" />
                Converting...
              </>
            ) : (
              <>Convert Document <span>→</span></>
            )}
          </button>

          {result && (
            <div className="mt-7">
              <div className="flex flex-col gap-4 rounded-xl border border-cyan-400/20 bg-cyan-400/[0.04] p-5 sm:flex-row sm:items-center sm:justify-between">
                <div className="flex items-center gap-3">
                  <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-emerald-400/10 text-emerald-400">
                    <CheckCircle2 size={22} />
                  </div>
                  <div>
                    <h3 className="font-semibold text-white">
                      Conversion completed
                    </h3>
                    <p className="mt-1 text-xs text-slate-400">
                      Your document is ready to preview or download.
                    </p>
                  </div>
                </div>

                <div className="flex flex-wrap gap-2">
                  <button
                    type="button"
                    onClick={() => setPreviewOpen(true)}
                    className="flex items-center justify-center gap-2 rounded-lg border border-cyan-400/30 bg-cyan-400/10 px-4 py-2.5 text-sm font-semibold text-cyan-300 transition hover:border-cyan-400 hover:bg-cyan-400/20"
                  >
                    <Eye size={17} /> Preview
                  </button>
                  <button
                    type="button"
                    onClick={handleDownload}
                    className="flex items-center justify-center gap-2 rounded-lg bg-cyan-400 px-4 py-2.5 text-sm font-semibold text-slate-950 transition hover:bg-cyan-300"
                  >
                    <Download size={17} /> Download
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>

        <p className="mt-5 text-center text-xs text-slate-500">
          Your files are processed in your browser and are not uploaded to our servers.
        </p>
      </div>

      {previewOpen && result && (
        <div className="fixed inset-0 z-[100] flex h-screen w-screen flex-col bg-[#070b14]">
          <div className="flex shrink-0 items-center justify-between border-b border-white/10 bg-[#111c2d] px-5 py-4 sm:px-8">
            <div className="flex min-w-0 items-center gap-3">
              <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-cyan-400/10 text-cyan-400">
                <FileText size={21} />
              </div>
              <div className="min-w-0">
                <h2 className="font-semibold text-white">Document Preview</h2>
                <p className="truncate text-xs text-slate-400">{file?.name}</p>
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

          <div className="flex shrink-0 items-center justify-between border-b border-white/5 px-5 py-3 sm:px-8">
            <div className="flex items-center gap-2 text-xs text-slate-400">
              <span className="h-2 w-2 rounded-full bg-emerald-400" />
              Preview ready
            </div>
            <div className="flex items-center gap-2 text-xs font-medium text-cyan-400">
              <Maximize2 size={14} />
              Full Screen Preview
            </div>
          </div>

          <div className="min-h-0 flex-1 overflow-y-auto bg-[#080d17] p-4 sm:p-8 lg:p-12">
            <div className="mx-auto min-h-full w-full max-w-5xl rounded-lg border border-slate-200/10 bg-[#f8fafc] px-6 py-10 text-slate-800 shadow-2xl sm:px-12 sm:py-14 lg:px-20 lg:py-16">
              <div className="mb-8 flex items-center justify-between border-b border-slate-200 pb-5">
                <div>
                  <p className="text-xs font-bold uppercase tracking-[2px] text-cyan-700">
                    eSparks Document Conversion
                  </p>
                  <p className="mt-1 text-xs text-slate-400">
                    Converted Document
                  </p>
                </div>
                <span className="rounded-full bg-cyan-50 px-3 py-1 text-xs font-semibold text-cyan-700">
                  {extensions[outputFormat as Format]}
                </span>
              </div>

              <article className="break-words font-serif text-sm leading-8 text-slate-700">
                {result.split("\n").map((line, index) => {
                  const heading = line.match(/^(#{1,6})\s+(.+)$/);
                  const listItem = line.match(/^\s*[-*+]\s+(.+)$/);
                  const numberedItem = line.match(/^\s*\d+\.\s+(.+)$/);

                  if (heading) {
                    const level = heading[1].length;
                    const headingClass = level === 1
                      ? "mt-8 mb-4 text-3xl font-bold text-slate-900"
                      : level === 2
                      ? "mt-7 mb-3 text-2xl font-bold text-slate-900"
                      : "mt-5 mb-2 text-lg font-bold text-slate-900";

                    return (
                      <div key={index} className={headingClass}>
                        {heading[2]}
                      </div>
                    );
                  }

                  if (listItem || numberedItem) {
                    return (
                      <div key={index} className="ml-5 list-item list-disc pl-1">
                        {listItem?.[1] || numberedItem?.[1]}
                      </div>
                    );
                  }

                  if (!line.trim()) return <div key={index} className="h-3" />;

                  return <p key={index} className="mb-2">{line}</p>;
                })}
              </article>

              <div className="mt-16 border-t border-slate-200 pt-5 text-center text-xs text-slate-400">
                Generated with eSparks Document Conversion
              </div>
            </div>
          </div>

          <div className="flex shrink-0 items-center justify-between gap-3 border-t border-white/10 bg-[#111c2d] px-5 py-4 sm:px-8">
            <span className="text-xs text-slate-400">
              eSparks Document Conversion · Preview
            </span>
            <button
              type="button"
              onClick={handleDownload}
              className="flex items-center justify-center gap-2 rounded-lg bg-cyan-400 px-5 py-2.5 text-sm font-bold text-slate-950 transition hover:bg-cyan-300"
            >
              <Download size={17} />
              Download Document
            </button>
          </div>
        </div>
      )}
    </section>
  );
}
