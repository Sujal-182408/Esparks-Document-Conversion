
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
    <section id="about" className="about-section container-main py-24">
      <div className="grid items-center gap-14 lg:grid-cols-2">
        <div>
          <span className="about-badge mb-4 inline-block rounded-full border px-4 py-2 text-xs font-semibold uppercase tracking-widest">
            About Our Platform
          </span>

          <h2 className="text-3xl font-bold leading-tight md:text-5xl">
            Simplifying Your
            <span className="text-gradient"> Document Workflow</span>
          </h2>

          <p className="about-muted mt-6 text-sm leading-8 md:text-base">
            eSparks Document Conversion provides simple tools to transform documents
            between supported formats. Our goal is to make document conversion
            convenient, accessible and straightforward.
          </p>

          <p className="about-muted mt-4 text-sm leading-8 md:text-base">
            With browser-based processing, your selected files are handled locally
            without requiring an account or server upload.
          </p>

          <div className="mt-8 flex flex-wrap gap-3">
            {["DOCX", "PDF", "PPTX", "MARKDOWN"].map((format) => (
              <span key={format} className="about-format rounded-lg border px-4 py-2 text-xs font-semibold tracking-wider">
                {format}
              </span>
            ))}
          </div>
        </div>

        <div className="grid gap-4 sm:grid-cols-2">
          {features.map((feature) => {
            const Icon = feature.icon;

            return (
              <div key={feature.title} className="about-feature glass-card p-6 transition-all hover:border-cyan-400/40">
                <div className="mb-5 flex h-12 w-12 items-center justify-center rounded-xl bg-cyan-400/10 text-cyan-500">
                  <Icon size={24} />
                </div>
                <h3 className="about-heading mb-2 text-lg font-semibold">
                  {feature.title}
                </h3>
                <p className="about-muted text-sm leading-6">
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
