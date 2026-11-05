
"use client";

import Link from "next/link";
import { useState } from "react";
import { Menu, X } from "lucide-react";

const navLinks = [
  { name: "Home", href: "/" },
  { name: "How It Works", href: "/how-it-works" },
  { name: "About", href: "/about" }
];



export default function Navbar() {
  const [mobileOpen, setMobileOpen] = useState(false);

  const closeMenu = () => setMobileOpen(false);

  return (
    <header className="fixed top-0 left-0 z-50 w-full border-b border-white/10 bg-[#070b14]/95 shadow-lg shadow-black/10 backdrop-blur-xl">
      <nav className="container-main flex h-20 items-center justify-between">
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

        <div className="hidden items-center gap-8 lg:flex">
          {navLinks.map((link) => (
            <Link
              key={link.name}
              href={link.href}
              className="text-sm font-medium text-slate-300 transition hover:text-cyan-400"
            >
              {link.name}
            </Link>
          ))}
        </div>

        <button
          type="button"
          onClick={() => setMobileOpen(!mobileOpen)}
          className="rounded-lg p-2 text-white transition hover:bg-white/10 lg:hidden"
          aria-label={mobileOpen ? "Close menu" : "Open menu"}
        >
          {mobileOpen ? <X size={24} /> : <Menu size={24} />}
        </button>
      </nav>

      {mobileOpen && (
        <div className="border-t border-white/10 bg-[#070b14] px-5 py-5 lg:hidden">
          <div className="flex flex-col gap-5">
            {navLinks.map((link) => (
              <Link
                key={link.name}
                href={link.href}
                onClick={closeMenu}
                className="text-sm font-medium text-slate-300 transition hover:text-cyan-400"
              >
                {link.name}
              </Link>
            ))}
          </div>
        </div>
      )}
    </header>
  );
}
