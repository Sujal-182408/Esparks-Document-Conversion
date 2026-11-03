
"use client";

import Link from "next/link";
import { useState } from "react";
import {
  Menu, X, ChevronDown, FileText, File,
  Presentation, Code2, ArrowRight
} from "lucide-react";

const navLinks = [
  { name: "Home", href: "/" },
  { name: "How It Works", href: "/#how-it-works" },
  { name: "About", href: "/#about" }
];

const converterLinks = [
  {
    name: "Word to Markdown",
    href: "/?from=docx&to=md#converter",
    icon: FileText
  },
  {
    name: "PDF to Markdown",
    href: "/?from=pdf&to=md#converter",
    icon: File
  },
  {
    name: "PowerPoint to Markdown",
    href: "/?from=pptx&to=md#converter",
    icon: Presentation
  },
  {
    name: "Markdown to Word",
    href: "/?from=md&to=docx#converter",
    icon: Code2
  }
];

export default function Navbar() {
  const [menuOpen, setMenuOpen] = useState(false);
  const [converterOpen, setConverterOpen] = useState(false);

  const closeMenu = () => {
    setMenuOpen(false);
    setConverterOpen(false);
  };

  return (
    <header className="sticky top-0 z-50 border-b border-white/8 bg-[#070b14]/90 backdrop-blur-xl">
      <nav className="container-main flex h-19 items-center justify-between">
        {/* Logo */}
        <Link href="/" onClick={closeMenu} className="flex items-center gap-3">
        <img src="brand.png" alt="Esparks IT Conversion" width={200} height={50} />
        </Link>

        {/* Desktop Navigation */}
        <div className="hidden items-center gap-8 md:flex">
          <Link
            href="/"
            className="text-sm text-slate-300 transition hover:text-cyan-400"
          >
            Home
          </Link>

          {/* Converter Dropdown */}
          <div className="group relative">
            <button
              type="button"
              onClick={() => setConverterOpen(!converterOpen)}
              aria-expanded={converterOpen}
              className="flex items-center gap-1.5 text-sm text-slate-300 transition hover:text-cyan-400"
            >
              Converters
              <ChevronDown
                size={15}
                className={`transition-transform ${
                  converterOpen ? "rotate-180" : ""
                }`}
              />
            </button>

            <div
              className={`absolute left-1/2 top-full w-64 -translate-x-1/2 pt-5 ${
                converterOpen
                  ? "block"
                  : "hidden group-hover:block group-focus-within:block"
              }`}
            >
              <div className="rounded-xl border border-white/10 bg-[#0d1422] p-2 shadow-2xl shadow-black/40">
                {converterLinks.map(({ name, href, icon: Icon }) => (
                  <Link
                    key={href}
                    href={href}
                    onClick={() => setConverterOpen(false)}
                    className="flex items-center gap-3 rounded-lg px-3 py-3 text-sm text-slate-300 transition hover:bg-cyan-400/10 hover:text-cyan-400"
                  >
                    <Icon size={17} />
                    {name}
                  </Link>
                ))}
              </div>
            </div>
          </div>

          {navLinks.slice(1).map((link) => (
            <Link
              key={link.name}
              href={link.href}
              className="text-sm text-slate-300 transition hover:text-cyan-400"
            >
              {link.name}
            </Link>
          ))}
        </div>

        {/* Desktop CTA */}
        <Link
          href="/#converter"
          className="hidden items-center gap-2 rounded-lg bg-cyan-400 px-5 py-2.5 text-sm font-bold text-slate-950 transition hover:bg-cyan-300 md:flex"
        >
          Start Converting <ArrowRight size={16} />
        </Link>

        {/* Mobile Menu Button */}
        <button
          type="button"
          onClick={() => setMenuOpen(!menuOpen)}
          aria-label={menuOpen ? "Close navigation menu" : "Open navigation menu"}
          aria-expanded={menuOpen}
          className="rounded-lg border border-white/10 p-2 text-white md:hidden"
        >
          {menuOpen ? <X size={22} /> : <Menu size={22} />}
        </button>
      </nav>

      {/* Mobile Navigation */}
      {menuOpen && (
        <div className="border-t border-white/10 bg-[#0a101d] px-5 py-5 md:hidden">
          <div className="container-main flex flex-col gap-2">
            <Link
              href="/"
              onClick={closeMenu}
              className="rounded-lg px-3 py-3 text-sm text-slate-300 hover:bg-white/5"
            >
              Home
            </Link>

            {/* Mobile Converter Dropdown */}
            <div>
              <button
                type="button"
                onClick={() => setConverterOpen(!converterOpen)}
                aria-expanded={converterOpen}
                className="flex w-full items-center justify-between rounded-lg px-3 py-3 text-sm text-slate-300 hover:bg-white/5"
              >
                Converters
                <ChevronDown
                  size={16}
                  className={`transition-transform ${
                    converterOpen ? "rotate-180" : ""
                  }`}
                />
              </button>

              {converterOpen && (
                <div className="ml-3 border-l border-cyan-400/20 pl-3">
                  {converterLinks.map(({ name, href, icon: Icon }) => (
                    <Link
                      key={href}
                      href={href}
                      onClick={closeMenu}
                      className="flex items-center gap-3 rounded-lg px-3 py-3 text-sm text-slate-400 transition hover:text-cyan-400"
                    >
                      <Icon size={16} />
                      {name}
                    </Link>
                  ))}
                </div>
              )}
            </div>

            {navLinks.slice(1).map((link) => (
              <Link
                key={link.name}
                href={link.href}
                onClick={closeMenu}
                className="rounded-lg px-3 py-3 text-sm text-slate-300 hover:bg-white/5"
              >
                {link.name}
              </Link>
            ))}

            <Link
              href="/#converter"
              onClick={closeMenu}
              className="mt-2 flex items-center justify-center gap-2 rounded-lg bg-cyan-400 px-5 py-3 text-sm font-bold text-slate-950"
            >
              Start Converting <ArrowRight size={16} />
            </Link>
          </div>
        </div>
      )}
    </header>
  );
}
