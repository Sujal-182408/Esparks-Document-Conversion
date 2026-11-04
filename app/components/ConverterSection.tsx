
import Link from "next/link";
import { FileText, File, Presentation, FileCode2, ArrowUpRight } from "lucide-react";

const converters = [
  {
    title: "Word to Markdown",
    href: "/?converter=word-to-markdown",
    icon: FileText,
    from: "DOCX",
    to: "MD",
    description: "Turn Word documents into clean, easy-to-edit Markdown.",
  },
  {
    title: "PDF to Markdown",
    href: "/?converter=pdf-to-markdown",
    icon: File,
    from: "PDF",
    to: "MD",
    description: "Extract readable text from PDFs and convert it to Markdown.",
  },
  {
    title: "PowerPoint to Markdown",
    href: "/?converter=powerpoint-to-markdown",
    icon: Presentation,
    from: "PPTX",
    to: "MD",
    description: "Convert presentation content into structured Markdown.",
  },
  {
    title: "Markdown to Word",
    href: "/?converter=markdown-to-word",
    icon: FileCode2,
    from: "MD",
    to: "DOCX",
    description: "Create a formatted Word document from your Markdown files.",
  },
  {
    title: "Markdown to PDF",
    href: "/?converter=markdown-to-pdf",
    icon: File,
    from: "MD",
    to: "PDF",
    description: "Create a formatted PDF document from your Markdown files.",
  },
  {
    title: "Markdown to PowerPoint",
    href: "/?converter=markdown-to-powerpoint",
    icon: Presentation,
    from: "MD",
    to: "PPTX",
    description: "Turn Markdown sections and content into presentation slides.",
  },
];

export default function ConverterSection() {
  return (
    <section id="converters" className="container-main py-24">
      <div className="mb-12 text-center">
        <span className="mb-4 inline-block rounded-full border border-cyan-400/20 bg-cyan-400/10 px-4 py-2 text-xs font-semibold uppercase tracking-widest text-cyan-400">
          Our Tools
        </span>
        <h2 className="text-3xl font-bold tracking-tight text-white md:text-5xl">
          Powerful <span className="text-gradient">Document Converters</span>
        </h2>
        <p className="mx-auto mt-4 max-w-2xl text-sm leading-7 text-slate-400 md:text-base">
          Convert your documents quickly with simple, secure and browser-based tools.
        </p>
      </div>

      <div className="grid gap-6 sm:grid-cols-2 lg:grid-cols-4">
        {converters.map((item) => {
          const Icon = item.icon;

          return (
            <Link
              key={item.title}
              href={item.href}
              className="group glass-card p-6 transition-all duration-300 hover:-translate-y-2 hover:border-cyan-400/40 hover:shadow-[0_15px_45px_rgba(34,211,238,0.08)]"
            >
              <div className="mb-6 flex items-center justify-between">
                <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-cyan-400/10 text-cyan-400 transition-colors group-hover:bg-cyan-400 group-hover:text-slate-950">
                  <Icon size={24} />
                </div>
                <ArrowUpRight className="text-slate-500 transition-all group-hover:-translate-y-1 group-hover:translate-x-1 group-hover:text-cyan-400" size={20} />
              </div>

              <div className="mb-4 flex items-center gap-2 text-xs font-bold">
                <span className="rounded-md bg-white/5 px-2 py-1 text-slate-300">{item.from}</span>
                <span className="text-cyan-400">→</span>
                <span className="rounded-md bg-cyan-400/10 px-2 py-1 text-cyan-400">{item.to}</span>
              </div>

              <h3 className="mb-3 text-lg font-semibold text-white">{item.title}</h3>
              <p className="text-sm leading-6 text-slate-400">{item.description}</p>
            </Link>
          );
        })}
      </div>
    </section>
  );
}
