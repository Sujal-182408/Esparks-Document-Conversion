
import { Upload, RefreshCw, Download } from "lucide-react";

const steps = [
  {
    number: "01",
    title: "Upload Your Document",
    description: "Choose your DOCX, PDF or PPTX file, or enter Markdown content.",
    icon: Upload
  },
  {
    number: "02",
    title: "Convert Instantly",
    description: "Our browser-based tools process your document and generate the converted content.",
    icon: RefreshCw
  },
  {
    number: "03",
    title: "Download Your File",
    description: "Preview your converted content and download the result in your preferred format.",
    icon: Download
  }
];

export default function HowItWorks() {
  return (
    <section id="how-it-works" className="relative overflow-hidden py-24">
      <div className="container-main">
        <div className="mb-16 text-center">
          <span className="mb-4 inline-block rounded-full border border-cyan-400/20 bg-cyan-400/10 px-4 py-2 text-xs font-semibold uppercase tracking-widest text-cyan-400">
            Simple Process
          </span>
          <h2 className="text-3xl font-bold text-white md:text-5xl">
            Convert in <span className="text-gradient">Three Easy Steps</span>
          </h2>
          <p className="mx-auto mt-4 max-w-xl text-sm leading-7 text-slate-400 md:text-base">
            No complicated setup. Just upload, convert and download.
          </p>
        </div>

        <div className="grid gap-8 md:grid-cols-3">
          {steps.map((step) => {
            const Icon = step.icon;

            return (
              <div key={step.number} className="glass-card relative p-8 text-center transition-all duration-300 hover:-translate-y-2 hover:border-cyan-400/40">
                <span className="absolute right-6 top-5 text-4xl font-black text-white/5">
                  {step.number}
                </span>

                <div className="mx-auto mb-6 flex h-16 w-16 items-center justify-center rounded-2xl bg-cyan-400/10 text-cyan-400 ring-1 ring-cyan-400/20">
                  <Icon size={28} />
                </div>

                <h3 className="mb-3 text-xl font-semibold text-white">
                  {step.title}
                </h3>
                <p className="text-sm leading-7 text-slate-400">
                  {step.description}
                </p>
              </div>
            );
          })}
        </div>
      </div>
    </section>
  );
}
