
"use client";

import Link from "next/link";
import { useEffect, useState } from "react";
import { Menu, X, Sun, Moon } from "lucide-react";

const navLinks = [
  { name: "Home", href: "/" },
  { name: "How It Works", href: "/how-it-works" },
  { name: "About", href: "/about" }
];

type Theme = "dark" | "light";

export default function Navbar() {
  const [mobileOpen, setMobileOpen] = useState(false);
  const [theme, setTheme] = useState<Theme>("dark");
  const [mounted, setMounted] = useState(false);

  useEffect(() => {
    const savedTheme = localStorage.getItem("esparks-theme");
    const initialTheme: Theme = savedTheme === "light" ? "light" : "dark";

    setTheme(initialTheme);
    document.documentElement.setAttribute("data-theme", initialTheme);
    setMounted(true);
  }, []);

  const toggleTheme = () => {
    const nextTheme: Theme = theme === "dark" ? "light" : "dark";

    setTheme(nextTheme);
    document.documentElement.setAttribute("data-theme", nextTheme);
    localStorage.setItem("esparks-theme", nextTheme);
  };

  const closeMenu = () => setMobileOpen(false);

  return (
    <header className="site-navbar fixed left-0 top-0 z-50 w-full border-b shadow-lg backdrop-blur-xl">
      <nav className="container-main flex h-20 items-center justify-between">
        <Link
          href="/"
          onClick={closeMenu}
          className="flex items-center gap-2"
        >
          
<img
  src={theme === "light" ? "/brand-light.png" : "/brand.png"}
  alt="eSparks Document Conversion"
  className="h-7 w-70 object-contain"
/>

        </Link>

        <div className="hidden items-center gap-8 lg:flex">
          {navLinks.map((link) => (
            <Link
              key={link.name}
              href={link.href}
              className="nav-link text-sm font-medium transition"
            >
              {link.name}
            </Link>
          ))}
        </div>

        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={toggleTheme}
            className="theme-toggle rounded-xl p-2.5 transition"
            aria-label={
              theme === "dark" ? "Switch to light theme" : "Switch to dark theme"
            }
            title={theme === "dark" ? "Light mode" : "Dark mode"}
          >
            {mounted && theme === "light" ? (
              <Moon size={20} />
            ) : (
              <Sun size={20} />
            )}
          </button>

          <button
            type="button"
            onClick={() => setMobileOpen(!mobileOpen)}
            className="mobile-menu-button rounded-lg p-2 transition lg:hidden"
            aria-label={mobileOpen ? "Close menu" : "Open menu"}
          >
            {mobileOpen ? <X size={24} /> : <Menu size={24} />}
          </button>
        </div>
      </nav>

      {mobileOpen && (
        <div className="mobile-nav border-t px-5 py-5 lg:hidden">
          <div className="flex flex-col gap-5">
            {navLinks.map((link) => (
              <Link
                key={link.name}
                href={link.href}
                onClick={closeMenu}
                className="nav-link text-sm font-medium transition"
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
