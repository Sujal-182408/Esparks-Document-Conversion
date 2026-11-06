"use client";

import { Suspense } from "react";
import DocumentConverter from "./DocumentConverter";

export default function Hero() {
  return (
    <section className="relative overflow-hidden border-b border-white/5">
      <div className="grid-background absolute inset-0 opacity-40" />
      <div className="absolute -left-40 top-20 h-80 w-80 rounded-full bg-cyan-500/10 blur-[120px]" />
      <div className="absolute -right-40 top-20 h-80 w-80 rounded-full bg-blue-500/10 blur-[120px]" />

      <div className="container-main relative py-12 sm:py-16 lg:py-20">
        <div className="mx-auto w-full max-w-4xl text-center">
          <div className="mb-7 inline-flex items-center gap-2 rounded-full border border-cyan-400/20 bg-cyan-400/5 px-4 py-2">
            <span className="h-2 w-2 animate-pulse rounded-full bg-cyan-400" />
            <span className="text-xs font-medium tracking-wide text-cyan-300">
              SIMPLE · FAST · SECURE
            </span>
          </div>

          <h1 className="mx-auto max-w-2xl text-4xl font-extrabold leading-[1.12] tracking-tight text-white sm:text-5xl lg:text-6xl">
            Convert Documents.
            <br />
            <span className="text-gradient">
              Simplify Your Workflow.
            </span>
          </h1>

          <p className="mx-auto mt-6 max-w-2xl text-sm leading-7 text-slate-400 sm:mt-7 sm:text-lg sm:leading-8">
            Convert Word, PDF, PowerPoint and Markdown files with ease.
            Fast browser-based processing, simple downloads, and no account
            required.
          </p>

          <div className="mt-7 flex flex-wrap items-center justify-center gap-x-7 gap-y-3 text-xs text-slate-400 sm:mt-9 sm:text-sm">
            <span className="flex items-center gap-2">
              <span className="text-cyan-400">✓</span> No registration
            </span>
            <span className="flex items-center gap-2">
              <span className="text-cyan-400">✓</span> Up to 20 MB
            </span>
            <span className="flex items-center gap-2">
              <span className="text-cyan-400">✓</span> Browser processing
            </span>
          </div>
        </div>

        <div className="mx-auto mt-10 w-full min-w-0 max-w-6xl sm:mt-12">
          <Suspense
            fallback={
              <div className="flex min-h-72 items-center justify-center rounded-2xl border border-white/10 bg-white/5 px-4 text-center text-sm text-slate-400 sm:min-h-80 sm:text-base">
                Loading document converter...
              </div>
            }
          >
            <DocumentConverter />
          </Suspense>
        </div>

      </div>
    </section>
  );
}