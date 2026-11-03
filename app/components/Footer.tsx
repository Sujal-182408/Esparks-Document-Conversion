
import Link from "next/link";
import { ArrowUpRight, Mail } from "lucide-react";

const quickLinks = [
  { name: "Home", href: "/" },
  { name: "Converters", href: "/#converter" },
  { name: "How It Works", href: "/#how-it-works" },
  { name: "About", href: "/#about" }
];

const formats = ["Word to Markdown", "PDF to Markdown", "PowerPoint to Markdown", "Markdown to Word"];

export default function Footer() {
  return (
    <footer className="border-t border-white/10 bg-[#050912]">
      <div className="container-main py-16">
        <div className="grid gap-12 sm:grid-cols-2 lg:grid-cols-4">
          <div className="lg:col-span-2">
            <Link href="/" className="inline-flex items-center gap-3">
              <img src="logo.svg" alt="Esparks IT Conversion" width={300} height={50} />

              
            </Link>

            <p className="mt-5 max-w-sm text-sm leading-7 text-slate-400">
              Simple, secure and convenient document conversion tools for your everyday workflow.
            </p>

            <p className="mt-5 text-xs font-medium uppercase tracking-widest text-cyan-400">
              Document Conversion Platform
            </p>
          </div>

          <div>
            <h3 className="mb-5 text-sm font-semibold uppercase tracking-wider text-white">
              Quick Links
            </h3>
            <ul className="space-y-4">
              {quickLinks.map((link) => (
                <li key={link.name}>
                  <Link href={link.href} className="text-sm text-slate-400 transition-colors hover:text-cyan-400">
                    {link.name}
                  </Link>
                </li>
              ))}
            </ul>
          </div>

          <div>
            <h3 className="mb-5 text-sm font-semibold uppercase tracking-wider text-white">
              Supported Tools
            </h3>
            <ul className="space-y-4">
              {formats.map((format) => (
                <li key={format} className="text-sm text-slate-400">
                  {format}
                </li>
              ))}
            </ul>
          </div>
        </div>

        <div className="mt-14 flex flex-col justify-between gap-5 border-t border-white/10 pt-7 sm:flex-row sm:items-center">
          <p className="text-xs text-slate-500">
            © {new Date().getFullYear()} eSparks IT Solutions. All rights reserved.
          </p>

          <Link
            href="https://esparksit.com"
            target="_blank"
            rel="noopener noreferrer"
            className="inline-flex items-center gap-2 text-xs font-medium text-slate-400 transition-colors hover:text-cyan-400"
          >
            Visit eSparks IT Solutions
            <ArrowUpRight size={14} />
          </Link>
        </div>
      </div>
    </footer>
  );
}
