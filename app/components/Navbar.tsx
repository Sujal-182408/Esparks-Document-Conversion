"use client";

import Link from "next/link";
import { useState, useEffect } from "react";
import {
  Menu,
  X,
  ChevronDown,
  FileText,
  File,
  Presentation,
  Code2,
  ArrowRight,
  ArrowUp
} from "lucide-react";

const navLinks = [
  { name: "Home", href: "/" },
  { name: "How It Works", href: "/how-it-works" },
  { name: "About", href: "/about" }
];

const converterLinks = [
  { name: "Word to Markdown", href: "/?converter=word-to-markdown", icon: FileText },
  { name: "PDF to Markdown", href: "/?converter=pdf-to-markdown", icon: File },
  { name: "PowerPoint to Markdown", href: "/?converter=powerpoint-to-markdown", icon: Presentation },
  { name: "Markdown to Word", href: "/?converter=markdown-to-word", icon: Code2 },
  { name: "Markdown to PDF", href: "/?converter=markdown-to-pdf", icon: File },
  { name: "Markdown to PowerPoint", href: "/?converter=markdown-to-powerpoint", icon: Presentation }
];

export default function Navbar() {
  const [mobileOpen, setMobileOpen] = useState(false);
  const [converterOpen, setConverterOpen] = useState(false);
  const [showBackToTop, setShowBackToTop] = useState(false);

  useEffect(() => {
    const handleScroll = () => {
      setShowBackToTop(window.scrollY > 300);
    };

    window.addEventListener("scroll", handleScroll, { passive: true });
    handleScroll();

    return () => window.removeEventListener("scroll", handleScroll);
  }, []);

  const closeMenu = () => {
    setMobileOpen(false);
    setConverterOpen(false);
  };

  const scrollToTop = () => {
    window.scrollTo({
      top: 0,
      behavior: "smooth"
    });
  };

  return (
    <>
     <header className="fixed top-0 left-0 z-50 w-full border-b border-white/10 bg-[#070b14]/95 shadow-lg shadow-black/10 backdrop-blur-xl">
        <nav className="container-main flex h-20 items-center justify-between">

          {/* Logo */}
          <Link
            href="/"
            onClick={closeMenu}
            className="flex items-center gap-2"
          >
            <img
              src="/brand.png"
              alt="eSparksDocs"
              className="h-7 w-70 object-contain"
            />
          </Link>

          {/* Desktop Navigation */}
          <div className="hidden items-center gap-8 lg:flex">

            <Link
              href="/"
              className="text-sm font-medium text-slate-300 transition hover:text-cyan-400"
            >
              Home
            </Link>

            {/* Desktop Converter Dropdown */}
            <div className="group relative">
              <button
                type="button"
                className="flex items-center gap-1.5 text-sm font-medium text-slate-300 transition hover:text-cyan-400"
              >
                Converters
                <ChevronDown
                  size={16}
                  className="transition-transform duration-200 group-hover:rotate-180"
                />
              </button>

              <div className="invisible absolute left-1/2 top-full z-50 w-72 -translate-x-1/2 pt-5 opacity-0 transition-all duration-200 group-hover:visible group-hover:opacity-100">
                <div className="rounded-xl border border-white/10 bg-[#101827] p-2 shadow-2xl">
                  {converterLinks.map(({ name, href, icon: Icon }) => (
                    <Link
                      key={href}
                      href={href}
                      className="flex items-center gap-3 rounded-lg px-3 py-3 text-sm text-slate-300 transition hover:bg-cyan-400/10 hover:text-cyan-400"
                    >
                      <Icon size={17} />
                      {name}
                    </Link>
                  ))}
                </div>
              </div>
            </div>

            {/* Other Navigation Links */}
            {navLinks.slice(1).map((link) => (
              <Link
                key={link.name}
                href={link.href}
                className="text-sm font-medium text-slate-300 transition hover:text-cyan-400"
              >
                {link.name}
              </Link>
            ))}
          </div>

          {/* Desktop CTA */}
          <Link
            href="/#converter"
            className="hidden items-center gap-2 rounded-lg bg-cyan-400 px-5 py-2.5 text-sm font-semibold text-[#070b14] transition hover:bg-cyan-300 lg:flex"
          >
            Start Converting
            <ArrowRight size={16} />
          </Link>

          {/* Mobile Menu Button */}
          <button
            type="button"
            onClick={() => setMobileOpen(!mobileOpen)}
            className="rounded-lg p-2 text-white transition hover:bg-white/10 lg:hidden"
            aria-label={mobileOpen ? "Close menu" : "Open menu"}
          >
            {mobileOpen ? <X size={24} /> : <Menu size={24} />}
          </button>
        </nav>

        {/* Mobile Navigation */}
        {mobileOpen && (
          <div className="border-t border-white/10 bg-[#070b14] px-5 py-5 lg:hidden">
            <div className="flex flex-col gap-4">

              <Link
                href="/"
                onClick={closeMenu}
                className="text-sm font-medium text-slate-300 transition hover:text-cyan-400"
              >
                Home
              </Link>

              {/* Mobile Converter Dropdown */}
              <div>
                <button
                  type="button"
                  onClick={() => setConverterOpen(!converterOpen)}
                  className="flex w-full items-center justify-between text-sm font-medium text-slate-300 transition hover:text-cyan-400"
                >
                  Converters
                  <ChevronDown
                    size={16}
                    className={`transition-transform duration-200 ${
                      converterOpen ? "rotate-180" : ""
                    }`}
                  />
                </button>

                {converterOpen && (
                  <div className="mt-3 flex flex-col gap-1 border-l border-cyan-400/30 pl-3">
                    {converterLinks.map(({ name, href, icon: Icon }) => (
                      <Link
                        key={href}
                        href={href}
                        onClick={closeMenu}
                        className="flex items-center gap-3 rounded-lg px-3 py-2.5 text-sm text-slate-400 transition hover:bg-cyan-400/10 hover:text-cyan-400"
                      >
                        <Icon size={16} />
                        {name}
                      </Link>
                    ))}
                  </div>
                )}
              </div>

              {/* Mobile Other Links */}
              {navLinks.slice(1).map((link) => (
                <Link
                  key={link.name}
                  href={link.href}
                  onClick={closeMenu}
                  className="text-sm font-medium text-slate-300 transition hover:text-cyan-400"
                >
                  {link.name}
                </Link>
              ))}

              {/* Mobile CTA */}
              <Link
                href="/#converter"
                onClick={closeMenu}
                className="mt-2 flex items-center justify-center gap-2 rounded-lg bg-cyan-400 px-5 py-3 text-sm font-semibold text-[#070b14] transition hover:bg-cyan-300"
              >
                Start Converting
                <ArrowRight size={16} />
              </Link>
            </div>
          </div>
        )}
      </header>

      {/* Floating Back to Top */}
      <button
        type="button"
        onClick={scrollToTop}
        aria-label="Back to top"
        title="Back to top"
        className={`fixed right-5 top-24 z-[60] flex h-11 w-11 items-center justify-center rounded-xl border border-cyan-400/30 bg-[#101827]/95 text-cyan-400 shadow-lg shadow-cyan-500/10 backdrop-blur-xl transition-all duration-300 hover:-translate-y-1 hover:bg-cyan-400 hover:text-[#070b14] sm:right-8 ${
          showBackToTop
            ? "visible translate-y-0 opacity-100"
            : "invisible translate-y-3 opacity-0"
        }`}
      >
        <ArrowUp size={20} />
      </button>
    </>
  );
}