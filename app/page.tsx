"use client";

import { useRef, useState } from "react";

import Link from "next/link";
import {
  downloadTextFile,
  downloadWordFile,
  markdownFilename,
  pdfMarkdownFilename,
  powerpointMarkdownFilename,
} from "./services/downloadservice";
import { wordToMarkdown } from "./converters/word-to-markdown";
import { pdfToMarkdown } from "./converters/pdfToMarkdown";
import { convertPowerPointToMarkdown } from "./converters/powerpointToMarkdown";
import { convertMarkdownToWord } from "./converters/markdownToWord";

const formats = [
  { value: "docx", label: "Word document", ext: ".docx", icon: "W" },
  { value: "pdf", label: "PDF document", ext: ".pdf", icon: "P" },
  { value: "pptx", label: "PowerPoint", ext: ".pptx", icon: "S" },
  { value: "md", label: "Markdown", ext: ".md", icon: "M" },
] as const;

type Format = (typeof formats)[number]["value"];
const acceptByFormat: Record<Format, string> = {
  docx: ".docx",
  pdf: ".pdf",
  pptx: ".pptx",
  md: ".md,.markdown,.txt",
};

const implemented = ["docx-md", "pdf-md", "pptx-md", "md-docx"];

export default function Home() {
  const [inputFormat, setInputFormat] = useState<Format>("docx");
  const [outputFormat, setOutputFormat] = useState<Format>("md");
  const [file, setFile] = useState<File | null>(null);
  const [markdown, setMarkdown] = useState("");
  const [busy, setBusy] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  const workflowKey = `${inputFormat}-${outputFormat}`;
  const canConvert = implemented.includes(workflowKey);

  function resetResult() {
    setMarkdown("");
    setError("");
    setSuccess(false);
  }

  function selectFile(nextFile?: File) {
    resetResult();
    if (!nextFile) return;

    const allowedExtensions = acceptByFormat[inputFormat]
      .split(",")
      .map((extension) => extension.trim());

    if (
      !allowedExtensions.some((extension) =>
        nextFile.name.toLowerCase().endsWith(extension),
      )
    ) {
      setFile(null);
      setError(
        `Choose a valid ${formats.find((f) => f.value === inputFormat)?.label} file.`,
      );
      return;
    }

    if (nextFile.size === 0) {
      setFile(null);
      setError("This file is empty. Please choose another file.");
      return;
    }

    if (nextFile.size > 20 * 1024 * 1024) {
      setFile(null);
      setError("The maximum file size is 20 MB.");
      return;
    }

    setFile(nextFile);
    setError("");
  }

  const handleConvert = async () => {
    setError("");
    setSuccess(false);
    setMarkdown("");
    if (!file) {
      setError("Please select a file before converting.");
      return;
    }

    resetResult();

    if (!canConvert) {
      setError(
        `${formats.find((f) => f.value === inputFormat)?.label} to ${
          formats.find((f) => f.value === outputFormat)?.label
        } is not connected yet. This workflow needs its converter module before it can generate a real downloadable file.`,
      );
      return;
    }

    setBusy(true);
    try {
      let result: string;
      if (inputFormat === "docx" && outputFormat === "md") {
        result = await wordToMarkdown(file);
      } else if (inputFormat === "pdf" && outputFormat === "md") {
        result = await pdfToMarkdown(file);
      } else if (inputFormat === "pptx" && outputFormat === "md") {
        result = await convertPowerPointToMarkdown(file);
      } else if (inputFormat === "md" && outputFormat === "docx") {
        const markdown = await file.text();
        const docxBlob = await convertMarkdownToWord(markdown);
        await downloadWordFile(docxBlob, file.name);
        setSuccess(true);
        return;
      } else {
        setError("This conversion is not implemented yet.");
        return;
      }
      setSuccess(true);
      setMarkdown(result);
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : "Conversion failed. Please try another file.",
      );
    } finally {
      setBusy(false);
    }
  };

  function handleDownload() {
    if (!markdown.trim() || !file) {
      setError("Convert a document before downloading.");
      return;
    }
    const filename =
      inputFormat === "pdf"
        ? pdfMarkdownFilename(file.name)
        : inputFormat === "pptx"
          ? powerpointMarkdownFilename(file.name)
          : markdownFilename(file.name);

    downloadTextFile(markdown, filename);
  }

  const inputInfo = formats.find((f) => f.value === inputFormat)!;
  const outputInfo = formats.find((f) => f.value === outputFormat)!;

  return (
    <main className="min-h-screen overflow-hidden bg-[#090b13] text-white">
      <div className="pointer-events-none fixed inset-0 overflow-hidden">
        <div className="absolute -top-48 left-1/2 h-[460px] w-[760px] -translate-x-1/2 rounded-full bg-violet-600/15 blur-[130px]" />
        <div className="absolute right-[-180px] top-[550px] h-[380px] w-[380px] rounded-full bg-indigo-500/10 blur-[110px]" />
      </div>

      <div className="relative">
        <header className="mx-auto flex max-w-7xl items-center justify-between px-5 py-6 sm:px-8">
          <Link href="/" className="flex items-center gap-3">
            <span className="flex h-10 w-10 items-center justify-center rounded-xl bg-violet-500 text-xl font-black shadow-lg shadow-violet-500/20">
              ↗
            </span>
            <span className="text-lg font-bold tracking-tight">
              Convert<span className="text-violet-400">Flow</span>
            </span>
          </Link>

          <a
            href="#converter"
            className="rounded-full border border-white/10 bg-white/[0.04] px-4 py-2 text-sm font-medium text-slate-200 transition hover:border-violet-400/50 hover:bg-violet-500/10"
          >
            Open converter <span className="ml-1">↘</span>
          </a>
        </header>

        <section className="mx-auto grid max-w-7xl items-center gap-12 px-5 pb-16 pt-12 sm:px-8 md:pb-24 md:pt-20 lg:grid-cols-[1.05fr_0.95fr] lg:gap-16">
          <div>
            <div className="mb-7 inline-flex items-center gap-2 rounded-full border border-violet-400/20 bg-violet-400/[0.08] px-3.5 py-2 text-xs font-semibold tracking-wide text-violet-200">
              <span className="h-2 w-2 rounded-full bg-emerald-400" />
              DOCUMENT CONVERSION, SIMPLIFIED
            </div>

            <h1 className="max-w-2xl text-4xl font-semibold leading-[1.1] tracking-tight sm:text-5xl md:text-6xl">
              Your files.
              <br />
              <span className="bg-gradient-to-r from-violet-300 via-fuchsia-300 to-indigo-300 bg-clip-text text-transparent">
                A better format.
              </span>
            </h1>

            <p className="mt-6 max-w-xl text-base leading-8 text-slate-400 sm:text-lg">
              Convert documents into the formats you need. A simple workspace
              for Word, PDF, PowerPoint and Markdown—without the clutter.
            </p>

            <div className="mt-8 flex flex-wrap gap-3">
              {formats.map((format) => (
                <div
                  key={format.value}
                  className="flex items-center gap-2 rounded-xl border border-white/[0.08] bg-white/[0.035] px-3 py-2.5 text-sm text-slate-300"
                >
                  <span className="flex h-7 w-7 items-center justify-center rounded-lg bg-white/[0.07] text-xs font-bold text-violet-200">
                    {format.icon}
                  </span>
                  {format.ext}
                </div>
              ))}
            </div>

            <div className="mt-10 flex flex-wrap items-center gap-x-6 gap-y-3 text-sm text-slate-400">
              <span className="flex items-center gap-2">
                <span className="text-emerald-400">✓</span> No account required
              </span>
              <span className="flex items-center gap-2">
                <span className="text-emerald-400">✓</span> Up to 20 MB per file
              </span>
            </div>
          </div>

          <div id="converter" className="scroll-mt-8">
            <div className="rounded-[28px] border border-white/10 bg-[#111420]/90 p-4 shadow-2xl shadow-black/30 backdrop-blur-xl sm:p-6">
              <div className="flex items-start justify-between gap-4 px-1 pb-5">
                <div>
                  <p className="text-lg font-semibold">File converter</p>
                  <p className="mt-1 text-sm text-slate-400">
                    Choose your formats and get started.
                  </p>
                </div>
                <span className="rounded-lg border border-emerald-400/20 bg-emerald-400/[0.08] px-2.5 py-1.5 text-xs font-medium text-emerald-300">
                  Free to use
                </span>
              </div>

              <div className="grid grid-cols-1 gap-3 sm:grid-cols-[1fr_auto_1fr] sm:items-end">
                <div>
                  <label
                    htmlFor="input-format"
                    className="mb-2 block text-xs font-semibold uppercase tracking-wider text-slate-400"
                  >
                    Convert from
                  </label>
                  <select
                    id="input-format"
                    value={inputFormat}
                    onChange={(event) => {
                      const nextFormat = event.target.value as Format;
                      setInputFormat(nextFormat);
                      setFile(null);
                      resetResult();
                      if (nextFormat === outputFormat) {
                        setOutputFormat(nextFormat === "md" ? "docx" : "md");
                      }
                    }}
                    className="w-full rounded-xl border border-white/10 bg-[#191d2b] px-3 py-3.5 text-sm text-white outline-none transition focus:border-violet-400"
                  >
                    {formats.map((format) => (
                      <option key={format.value} value={format.value}>
                        {format.icon}　{format.label} ({format.ext})
                      </option>
                    ))}
                  </select>
                </div>

                <div className="hidden h-12 w-10 items-center justify-center text-xl text-violet-300 sm:flex">
                  →
                </div>

                <div>
                  <label
                    htmlFor="output-format"
                    className="mb-2 block text-xs font-semibold uppercase tracking-wider text-slate-400"
                  >
                    Convert to
                  </label>
                  <select
                    id="output-format"
                    value={outputFormat}
                    onChange={(event) => {
                      setOutputFormat(event.target.value as Format);
                      resetResult();
                    }}
                    className="w-full rounded-xl border border-white/10 bg-[#191d2b] px-3 py-3.5 text-sm text-white outline-none transition focus:border-violet-400"
                  >
                    {formats
                      .filter((format) => format.value !== inputFormat)
                      .map((format) => (
                        <option key={format.value} value={format.value}>
                          {format.icon}　{format.label} ({format.ext})
                        </option>
                      ))}
                  </select>
                </div>
              </div>

              <div className="my-5 flex items-center gap-3 text-xs text-slate-500">
                <span className="h-px flex-1 bg-white/[0.08]" />
                YOUR DOCUMENT
                <span className="h-px flex-1 bg-white/[0.08]" />
              </div>

              <input
                ref={fileRef}
                type="file"
                accept={acceptByFormat[inputFormat]}
                className="hidden"
                onChange={(event) => selectFile(event.target.files?.[0])}
              />

              <button
                type="button"
                onClick={() => fileRef.current?.click()}
                onDragOver={(event) => {
                  event.preventDefault();
                  setDragging(true);
                }}
                onDragLeave={() => setDragging(false)}
                onDrop={(event) => {
                  event.preventDefault();
                  setDragging(false);
                  selectFile(event.dataTransfer.files?.[0]);
                }}
                className={`flex min-h-[190px] w-full flex-col items-center justify-center rounded-2xl border border-dashed px-4 py-7 text-center transition ${
                  dragging
                    ? "border-violet-300 bg-violet-500/10"
                    : file
                      ? "border-emerald-400/40 bg-emerald-400/[0.04]"
                      : "border-white/15 bg-white/[0.025] hover:border-violet-400/60 hover:bg-violet-500/[0.04]"
                }`}
              >
                <span className="mb-4 flex h-14 w-14 items-center justify-center rounded-2xl border border-violet-300/20 bg-violet-400/10 text-2xl text-violet-200">
                  {file ? "✓" : "↑"}
                </span>
                {file ? (
                  <>
                    <span className="max-w-full break-all text-sm font-semibold text-white">
                      {file.name}
                    </span>
                    <span className="mt-2 text-xs text-slate-400">
                      {(file.size / 1024 / 1024).toFixed(2)} MB · Ready to
                      convert
                    </span>
                    <span className="mt-3 text-xs font-medium text-violet-300">
                      Choose another file
                    </span>
                  </>
                ) : (
                  <>
                    <span className="text-sm font-semibold text-slate-100">
                      Drop your file here
                    </span>
                    <span className="mt-2 text-sm text-slate-400">
                      or <span className="text-violet-300">browse files</span>{" "}
                      on your device
                    </span>
                    <span className="mt-3 text-xs text-slate-500">
                      {inputInfo.ext} · Maximum 20 MB
                    </span>
                  </>
                )}
              </button>

              {error && (
                <div
                  role="alert"
                  className="mt-4 rounded-xl border border-rose-400/20 bg-rose-400/[0.07] px-4 py-3 text-sm leading-6 text-rose-200"
                >
                  {error}
                </div>
              )}

              <button
                type="button"
                onClick={handleConvert}
                disabled={!file || busy}
                className="mt-5 flex w-full items-center justify-center gap-2 rounded-xl bg-linear-to-r from-violet-600 to-indigo-600 px-5 py-4 text-sm font-semibold text-white shadow-lg shadow-violet-900/25 transition hover:brightness-110 disabled:cursor-not-allowed disabled:opacity-40"
              >
                {busy ? (
                  <>
                    <span className="animate-spin">◌</span> Converting...
                  </>
                ) : (
                  <>
                    Convert to {outputInfo.ext} <span>↗</span>
                  </>
                )}
              </button>

              {success && (
                <div className="mt-4 rounded-xl border border-emerald-400/20 bg-emerald-400/[0.06] p-4">
                  <div className="flex items-center gap-2 text-sm font-semibold text-emerald-300">
                    <span>✓</span> Conversion completed
                  </div>
                  <p className="mt-1 text-xs leading-5 text-slate-400">
                    {
                      inputFormat === "md" && outputFormat === "docx"?
                      "Your Word document has been downloaded." : "Your Markdown file is ready. Preview it below or download it."
                    }
                  </p>
                  <textarea
                    aria-label="Converted Markdown preview"
                    value={markdown}
                    readOnly
                    className="mt-3 min-h-32 w-full resize-y rounded-lg border border-white/10 bg-[#0a0d16] p-3 font-mono text-xs leading-6 text-slate-300 outline-none"
                  />
                  <button
                    type="button"
                    onClick={handleDownload}
                    className="mt-3 w-full rounded-lg bg-emerald-400 px-4 py-3 text-sm font-semibold text-[#07120d] transition hover:bg-emerald-300"
                  >
                    ↓ Download {markdownFilename(file?.name ?? "document.docx")}
                  </button>
                </div>
              )}

              <p className="mt-4 text-center text-xs leading-5 text-slate-500">
                Your document is processed in this browser for the connected
                workflow. Avoid uploading confidential files until the complete
                privacy and processing setup has been verified.
              </p>
            </div>
          </div>
        </section>

        <section className="border-t border-white/[0.07] bg-white/[0.015]">
          <div className="mx-auto grid max-w-7xl gap-8 px-5 py-12 sm:px-8 md:grid-cols-3 md:py-16">
            {[
              {
                number: "01",
                title: "Choose your formats",
                description:
                  "Select the document you have and the format you need.",
              },
              {
                number: "02",
                title: "Upload your document",
                description:
                  "Pick a file from your device and check it before conversion.",
              },
              {
                number: "03",
                title: "Download your result",
                description:
                  "Preview supported conversions and download the generated file.",
              },
            ].map((step) => (
              <div key={step.number} className="flex gap-4">
                <span className="flex h-11 w-11 shrink-0 items-center justify-center rounded-xl border border-violet-400/20 bg-violet-400/[0.07] text-sm font-bold text-violet-300">
                  {step.number}
                </span>
                <div>
                  <h2 className="font-semibold text-slate-100">{step.title}</h2>
                  <p className="mt-2 text-sm leading-6 text-slate-400">
                    {step.description}
                  </p>
                </div>
              </div>
            ))}
          </div>
        </section>

        <footer className="mx-auto flex max-w-7xl flex-col gap-3 px-5 py-7 text-xs text-slate-500 sm:flex-row sm:items-center sm:justify-between sm:px-8">
          <p>
            © {new Date().getFullYear()} ConvertFlow. Built for simpler
            documents.
          </p>
          <p>Word · PDF · PowerPoint · Markdown</p>
        </footer>
      </div>
    </main>
  );
}
