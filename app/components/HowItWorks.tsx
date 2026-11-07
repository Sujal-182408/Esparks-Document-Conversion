
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
    <section id="how-it-works" className="how-section relative overflow-hidden py-24">
      <div className="container-main">
        <div className="mb-16 text-center">
          <span className="how-badge mb-4 inline-block rounded-full border px-4 py-2 text-xs font-semibold uppercase tracking-widest">
            Simple Process
          </span>

          <h2 className="text-3xl font-bold md:text-5xl">
            Convert in <span className="text-gradient">Three Easy Steps</span>
          </h2>

          <p className="how-muted mx-auto mt-4 max-w-xl text-sm leading-7 md:text-base">
            No complicated setup. Just upload, convert and download.
          </p>
        </div>

        <div className="grid gap-8 md:grid-cols-3">
          {steps.map((step) => {
            const Icon = step.icon;

            return (
              <div
                key={step.number}
                className="how-card glass-card relative p-8 text-center transition-all duration-300 hover:-translate-y-2"
              >
                <span className="how-number absolute right-6 top-5 text-4xl font-black">
                  {step.number}
                </span>

                <div className="mx-auto mb-6 flex h-16 w-16 items-center justify-center rounded-2xl bg-cyan-400/10 text-cyan-500 ring-1 ring-cyan-400/20">
                  <Icon size={28} />
                </div>

                <h3 className="how-heading mb-3 text-xl font-semibold">
                  {step.title}
                </h3>

                <p className="how-muted text-sm leading-7">
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
