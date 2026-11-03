
"use client";

import { Suspense } from "react";
import { useSearchParams } from "next/navigation";
import Navbar from "./components/Navbar";
import Hero from "./components/Hero";
import ConverterSection from "./components/ConverterSection";
import HowItWorks from "./components/HowItWorks";
import About from "./components/About";
import Footer from "./components/Footer";
import DocumentConverter from "./components/DocumentConverter";

const converters = {
  "word-to-markdown": {
    title: "Word to Markdown Converter",
    description: "Convert DOCX Word documents into clean Markdown files.",
    inputFormat: "docx",
    outputFormat: "md"
  },
  "pdf-to-markdown": {
    title: "PDF to Markdown Converter",
    description: "Convert PDF documents into Markdown format.",
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
    description: "Convert Markdown files into downloadable Word documents.",
    inputFormat: "md",
    outputFormat: "docx"
  }
} as const;

function HomeContent() {
  const searchParams = useSearchParams();
  const type = searchParams.get("type");
  const selected = type
    ? converters[type as keyof typeof converters]
    : undefined;

  if (selected) {
    return (
      <main className="min-h-screen bg-[#070b14] text-white">
        <Navbar />
        <section className="grid-background border-b border-white/5 px-5 py-16 text-center sm:py-20">
          <div className="container-main">
            <p className="mb-4 text-xs font-semibold uppercase tracking-[3px] text-cyan-400">
              eSparks Document Conversion
            </p>
            <h1 className="text-3xl font-bold text-white sm:text-5xl">
              {selected.title}
            </h1>
            <p className="mx-auto mt-5 max-w-2xl text-sm leading-7 text-slate-400 sm:text-base">
              {selected.description}
            </p>
          </div>
        </section>
       
        <Footer />
      </main>
    );
  }

  return (
    <main className="min-h-screen overflow-hidden bg-[#070b14] text-white">
      <div className="pointer-events-none fixed inset-0 z-0 overflow-hidden">
        <div className="absolute -top-48 left-1/2 h-[460px] w-[760px] -translate-x-1/2 rounded-full bg-cyan-500/[0.08] blur-[130px]" />
        <div className="absolute right-[-180px] top-[550px] h-[380px] w-[380px] rounded-full bg-blue-500/[0.07] blur-[110px]" />
      </div>
      <div className="relative z-10">
        <Navbar />
        <Hero />
      <DocumentConverter />
        <HowItWorks />
        <About />
        <Footer />
      </div>
    </main>
  );
}

export default function Home() {
  return (
    <Suspense fallback={<div className="min-h-screen bg-[#070b14]" />}>
      <HomeContent />
    </Suspense>
  );
}
