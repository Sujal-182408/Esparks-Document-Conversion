
import Link from "next/link";

const formats = [
  { name: "DOCX", label: "Word Document", color: "text-blue-400", icon: "W" },
  { name: "PDF", label: "PDF Document", color: "text-red-400", icon: "P" },
  { name: "PPTX", label: "PowerPoint", color: "text-orange-400", icon: "P" },
  { name: "MD", label: "Markdown File", color: "text-cyan-400", icon: "M" }
];

export default function Hero() {
  return (
    <section className="relative overflow-hidden border-b border-white/5 mt-10">
      <div className="grid-background absolute inset-0 opacity-40" />
      <div className="absolute -left-40 top-20 h-80 w-80 rounded-full bg-cyan-500/10 blur-[120px]" />
      <div className="absolute -right-40 top-20 h-80 w-80 rounded-full bg-blue-500/10 blur-[120px]" />

      <div className="container-main relative grid min-h-162.5 items-center gap-16 py-20 lg:grid-cols-2 lg:py-24">
        <div>
          <div className="mb-7 inline-flex items-center gap-2 rounded-full border border-cyan-400/20 bg-cyan-400/5 px-4 py-2">
            <span className="h-2 w-2 animate-pulse rounded-full bg-cyan-400" />
            <span className="text-xs font-medium tracking-wide text-cyan-300">
              SIMPLE · FAST · SECURE
            </span>
          </div>

          <h1 className="max-w-2xl text-4xl font-extrabold leading-[1.12] tracking-tight text-white sm:text-5xl lg:text-6xl">
            Convert Documents.
            <br />
            <span className="text-gradient">Simplify Your Workflow.</span>
          </h1>

          <p className="mt-7 max-w-xl text-base leading-8 text-slate-400 sm:text-lg">
            Convert Word, PDF, PowerPoint and Markdown files with ease.
            Fast browser-based processing, simple downloads, and no account
            required.
          </p>

          <div className="mt-9 flex flex-wrap gap-4">
            <Link
          href="/#converter"
          className="hidden items-center gap-2 rounded-lg bg-cyan-400 px-5 py-2.5 text-sm font-bold text-slate-950 transition hover:bg-cyan-300 md:flex"
        >
              Explore Converters <span className="ml-2">→</span>
            </Link>
            <Link
              href="/#how-it-works"
              className="rounded-xl border border-white/10 bg-white/5 px-7 py-4 text-sm font-semibold text-white transition-all hover:border-cyan-400/40 hover:bg-white/10"
            >
              How It Works
            </Link>
          </div>

          <div className="mt-10 flex flex-wrap items-center gap-x-7 gap-y-3 text-sm text-slate-400">
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

        <div className="relative mx-auto w-full max-w-xl">
          <div className="absolute -inset-5 rounded-[32px] bg-cyan-400/5 blur-2xl" />
          <div className="glass-card glow-cyan relative overflow-hidden p-6 sm:p-8">
            <div className="mb-8 flex items-center justify-between">
              <div>
                <p className="text-sm font-semibold text-white">Document Workspace</p>
                <p className="mt-1 text-xs text-slate-500">Your files, simplified</p>
              </div>
              <div className="rounded-lg border border-emerald-400/20 bg-emerald-400/10 px-3 py-1.5 text-xs text-emerald-300">
                Ready
              </div>
            </div>

            <div className="grid grid-cols-2 gap-4">
              {formats.map((format) => (
                <div
                  key={format.name}
                  className="rounded-2xl border border-white/10 bg-white/[0.03] p-5 transition-all hover:-translate-y-1 hover:border-cyan-400/30"
                >
                  <div className={`mb-4 flex h-12 w-12 items-center justify-center rounded-xl bg-white/5 text-xl font-extrabold ${format.color}`}>
                    {format.icon}
                  </div>
                  <p className="font-bold text-white">{format.name}</p>
                  <p className="mt-1 text-xs text-slate-500">{format.label}</p>
                </div>
              ))}
            </div>

            <div className="my-7 flex items-center gap-3">
              <div className="h-px flex-1 bg-white/10" />
              <span className="text-xs text-slate-500">CONVERT BETWEEN FORMATS</span>
              <div className="h-px flex-1 bg-white/10" />
            </div>

            <div className="flex items-center justify-between rounded-xl border border-cyan-400/20 bg-cyan-400/5 p-4">
              <div className="flex items-center gap-3">
                <div className="flex h-10 w-10 items-center justify-center rounded-lg bg-cyan-400/10 text-cyan-300">
                  ↗
                </div>
                <div>
                  <p className="text-sm font-semibold text-white">Simple conversion</p>
                  <p className="mt-1 text-xs text-slate-500">Upload · Convert · Download</p>
                </div>
              </div>
              <span className="text-xl text-cyan-400">→</span>
            </div>
          </div>

          <div className="absolute -bottom-6 -left-5 hidden rounded-xl border border-white/10 bg-[#111c2d] px-4 py-3 shadow-xl sm:block">
            <p className="text-xs text-slate-400">Maximum file size</p>
            <p className="mt-1 text-lg font-bold text-white">20 MB</p>
          </div>
          <div className="absolute -right-5 -top-5 hidden rounded-xl border border-white/10 bg-[#111c2d] px-4 py-3 shadow-xl sm:block">
            <p className="text-xs text-slate-400">Processing</p>
            <p className="mt-1 text-sm font-bold text-cyan-400">In your browser</p>
          </div>
        </div>
      </div>
    </section>
  );
}
