
"use client";

import dynamic from "next/dynamic";
import { Suspense } from "react";
import { useSearchParams } from "next/navigation";
import Navbar from "./components/Navbar";
import Hero from "./components/Hero";
import Footer from "./components/Footer";

const HowItWorks = dynamic(
  () => import("./components/HowItWorks"),
  { loading: () => <div className="min-h-[400px]" /> }
);

const About = dynamic(
  () => import("./components/About"),
  { loading: () => <div className="min-h-[300px]" /> }
);

const converters = {
  "word-to-markdown": {
    title: "Word to Markdown Converter",
    description: "Convert DOCX Word documents into clean Markdown files.",
    inputFormat: "docx",
    outputFormat: "md"
  },
  "pdf-to-markdown": {
    title: "PDF to Markdown Converter",
    description: "Convert PDF documents into Markdown.",
    inputFormat: "pdf",
    outputFormat: "md"
  },
  "powerpoint-to-markdown": {
    title: "PowerPoint to Markdown Converter",
    description: "Convert PowerPoint presentations into Markdown.",
    inputFormat: "pptx",
    outputFormat: "md"
  },
  "markdown-to-word": {
    title: "Markdown to Word Converter",
    description: "Convert Markdown into Word documents.",
    inputFormat: "md",
    outputFormat: "docx"
  },
  "markdown-to-pdf": {
    title: "Markdown to PDF Converter",
    description: "Convert Markdown into PDF documents.",
    inputFormat: "md",
    outputFormat: "pdf"
  },
  "markdown-to-powerpoint": {
    title: "Markdown to PowerPoint Converter",
    description: "Convert Markdown into PowerPoint presentations.",
    inputFormat: "md",
    outputFormat: "pptx"
  }
} as const;

function HomeContent() {
  const searchParams = useSearchParams();
  const converter = searchParams.get("converter");

  const selected = converter
    ? converters[converter as keyof typeof converters]
    : undefined;

  return (
    <div className="home-page flex min-h-dvh flex-col overflow-x-clip">
      <Navbar />

      <main className="flex-1">
        {selected ? (
          <section className="grid-background border-b border-white/5 px-5 py-16 text-center sm:py-20">
            <div className="container-main">
              <p className="mb-4 text-xs font-semibold uppercase tracking-[3px] text-cyan-400">
                eSparks Document Conversion
              </p>

              <h1 className="text-3xl font-bold sm:text-5xl">
                {selected.title}
              </h1>

              <p className="mx-auto mt-5 max-w-2xl text-sm leading-7 text-slate-400 sm:text-base">
                {selected.description}
              </p>
            </div>
          </section>
        ) : converter ? (
          <section className="container-main py-24 text-center">
            <h1 className="text-3xl font-bold">Converter Not Found</h1>
          </section>
        ) : (
          <>
            <Hero />
            {/* <HowItWorks /> */}
            {/* <About /> */}
          </>
        )}
      </main>

      <Footer />
    </div>
  );
}

export default function Home() {
  return (
    <Suspense
      fallback={
        <div className="min-h-screen min-h-dvh bg-[var(--background)]" />
      }
    >
      <HomeContent />
    </Suspense>
  );
}
