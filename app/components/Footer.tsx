
"use client";

import Link from "next/link";
import { ArrowUpRight } from "lucide-react";
import { useTheme } from "./ThemeProvider";
import Image from "next/image";

const quickLinks = [
  { name: "Home", href: "/" },
  { name: "Converters", href: "/#converter" },
  { name: "How It Works", href: "/how-it-works" },
  { name: "About", href: "/about" }
];

const formats = [
  "Word to Markdown",
  "PDF to Markdown",
  "PowerPoint to Markdown",
  "Markdown to Word",
  "Markdown to PDF",
  "Markdown to PowerPoint"
];

export default function Footer() {
  const { theme } = useTheme();

  return (
    <footer className="site-footer mt-auto w-full shrink-0 border-t">
      <div className="container-main py-6 sm:py-8 lg:py-10">
        <div className="grid grid-cols-1 gap-10 sm:grid-cols-2 sm:gap-x-8 sm:gap-y-10 lg:grid-cols-[minmax(0,1.3fr)_minmax(10rem,0.7fr)_minmax(14rem,1fr)] lg:gap-12">
          <div className="min-w-0">
            <Link
              href="/"
              aria-label="eSparks Document Conversion home"
              className="inline-flex max-w-full items-center"
            >
              <Image
            
            src={theme === "light" ? "/brand-light.png" : "/brand.png"}
                alt="eSparks Document Conversion"
                width={280}
                height={56}
                className="h-7 w-70 object-contain"
              />
            </Link>

            <p className="footer-muted mt-5 max-w-md text-sm leading-7">
              Simple, secure and convenient document conversion tools for your everyday workflow.
            </p>

            <p className="mt-4 text-xs font-medium uppercase tracking-[0.18em] text-cyan-500">
              Document Conversion Platform
            </p>
          </div>

          <nav aria-label="Footer quick links">
            <h3 className="footer-heading mb-5 text-sm font-semibold uppercase tracking-wider">
              Quick Links
            </h3>

            <ul className="space-y-3">
              {quickLinks.map((link) => (
                <li key={link.name}>
                  <Link
                    href={link.href}
                    className="footer-link inline-flex min-h-8 items-center text-sm transition-colors hover:text-cyan-500 focus-visible:rounded-sm focus-visible:outline focus-visible:outline-2 focus-visible:outline-cyan-400"
                  >
                    {link.name}
                  </Link>
                </li>
              ))}
            </ul>
          </nav>

          <div>
            <h3 className="footer-heading mb-5 text-sm font-semibold uppercase tracking-wider">
              Supported Tools
            </h3>

            <ul className="grid grid-cols-1 gap-y-3 sm:grid-cols-2 lg:grid-cols-1">
              {formats.map((format) => (
                <li key={format} className="footer-muted text-sm leading-6">
                  {format}
                </li>
              ))}
            </ul>
          </div>
        </div>

        <div className="footer-bottom mt-8 flex flex-col gap-4 border-t pt-5 sm:mt-9 sm:flex-row sm:items-center sm:justify-between sm:gap-6 sm:pt-6">
          <p className="footer-muted text-xs leading-5">
            © {new Date().getFullYear()} eSparks IT Solutions. All rights reserved.
          </p>

          <Link
            href="https://esparksit.com"
            target="_blank"
            rel="noopener noreferrer"
            className="footer-link inline-flex min-h-8 w-fit items-center gap-2 text-xs font-medium transition-colors hover:text-cyan-500 focus-visible:rounded-sm focus-visible:outline focus-visible:outline-2 focus-visible:outline-cyan-400"
          >
            Visit eSparks IT Solutions
            <ArrowUpRight size={14} />
          </Link>
        </div>
      </div>
    </footer>
  );
}
