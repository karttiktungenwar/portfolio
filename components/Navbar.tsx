"use client";

import { useEffect, useState } from "react";

const links = [
  { href: "#hero", label: "~/home" },
  { href: "#about", label: "about" },
  { href: "#experience", label: "log" },
  { href: "#projects", label: "builds" },
  { href: "#contact", label: "contact" }
];

export default function Navbar() {
  const [scrolled, setScrolled] = useState(false);

  useEffect(() => {
    const onScroll = () => setScrolled(window.scrollY > 12);
    onScroll();
    window.addEventListener("scroll", onScroll, { passive: true });
    return () => window.removeEventListener("scroll", onScroll);
  }, []);

  return (
    <header
      className={`fixed top-0 z-50 w-full transition-colors duration-300 ${
        scrolled ? "bg-ink/90 backdrop-blur border-b border-line" : "bg-transparent"
      }`}
    >
      <nav className="mx-auto flex max-w-6xl items-center justify-between px-6 py-4 font-mono text-sm">
        <a href="#hero" className="flex items-center gap-2 text-paper focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-signal rounded" aria-label="Meet Kartik, Mobile App Developer and Designer">
          <span className="h-2 w-2 rounded-full bg-signal animate-pulse" aria-hidden="true" />
          <span className="font-bold tracking-tight">Meet Kartik.</span> 
          <span className="hidden md:inline text-muted text-sm border-l border-white/10 pl-2">
            Mobile App Developer & Designer
          </span>
        </a>
        <ul className="hidden items-center gap-8 md:flex">
          {links.map((link) => (
            <li key={link.href}>
              <a
                href={link.href}
                className="text-muted transition-colors hover:text-signal"
              >
                {link.label}
              </a>
            </li>
          ))}
        </ul>
        <a
          href="#contact"
          className="rounded border border-signal/40 px-3 py-1.5 text-signal transition-colors hover:bg-signal/10"
        >
          say hi
        </a>
      </nav>
    </header>
  );
}
