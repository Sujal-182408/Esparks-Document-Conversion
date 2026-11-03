
import { ShieldCheck, Zap, FileCheck2, Globe } from "lucide-react";

const features = [
  {
    icon: ShieldCheck,
    title: "Privacy Focused",
    description: "Files are processed directly in your browser without uploading them to our servers."
  },
  {
    icon: Zap,
    title: "Fast Conversion",
    description: "Convert supported documents with a simple and efficient workflow."
  },
  {
    icon: FileCheck2,
    title: "Multiple Formats",
    description: "Work with DOCX, PDF, PPTX and Markdown formats."
  },
  {
    icon: Globe,
    title: "Accessible Anywhere",
    description: "Use the conversion tools from any modern browser."
  }
];

export default function About() {
  return (
    <section id="about" className="container-main py-24">
      <div className="grid items-center gap-14 lg:grid-cols-2">
        <div>
          <span className="mb-4 inline-block rounded-full border border-cyan-400/20 bg-cyan-400/10 px-4 py-2 text-xs font-semibold uppercase tracking-widest text-cyan-400">
            About Our Platform
          </span>

          <h2 className="text-3xl font-bold leading-tight text-white md:text-5xl">
            Simplifying Your
            <span className="text-gradient"> Document Workflow</span>
          </h2>

          <p className="mt-6 text-sm leading-8 text-slate-400 md:text-base">
            eSparks Document Conversion provides simple tools to transform documents
            between supported formats. Our goal is to make document conversion
            convenient, accessible and straightforward.
          </p>

          <p className="mt-4 text-sm leading-8 text-slate-400 md:text-base">
            With browser-based processing, your selected files are handled locally
            without requiring an account or server upload.
          </p>

          <div className="mt-8 flex flex-wrap gap-3">
            {["DOCX", "PDF", "PPTX", "MARKDOWN"].map((format) => (
              <span key={format} className="rounded-lg border border-white/10 bg-white/5 px-4 py-2 text-xs font-semibold tracking-wider text-cyan-300">
                {format}
              </span>
            ))}
          </div>
        </div>

        <div className="grid gap-4 sm:grid-cols-2">
          {features.map((feature) => {
            const Icon = feature.icon;

            return (
              <div key={feature.title} className="glass-card p-6 transition-all hover:border-cyan-400/40">
                <div className="mb-5 flex h-12 w-12 items-center justify-center rounded-xl bg-cyan-400/10 text-cyan-400">
                  <Icon size={24} />
                </div>
                <h3 className="mb-2 text-lg font-semibold text-white">
                  {feature.title}
                </h3>
                <p className="text-sm leading-6 text-slate-400">
                  {feature.description}
                </p>
              </div>
            );
          })}
        </div>
      </div>
    </section>
  );
}
