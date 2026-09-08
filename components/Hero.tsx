"use client";

import { motion } from "framer-motion";
import { profile } from "@/data/portfolio";
import {FaWhatsapp, FaMedium, FaMicrosoft} from "react-icons/fa";
import { FiGithub, FiLinkedin, FiInstagram, FiFacebook } from "react-icons/fi";

function ContactRow({
  icon,
  label,
  href
}: {
  icon: React.ReactNode;
  label: string;
  href?: string;
}) {
  const content = (
    <span className="flex items-center gap-3 text-sm text-muted transition-colors hover:text-signal">
      <span className="text-signal">{icon}</span>
      {label}
    </span>
  );

  // Reusable animation and gesture props
  const motionProps = {
    whileHover: { scale: 1.05, y: -2 },
    whileTap: { scale: 0.95 },
    transition: { type: "spring", stiffness: 400, damping: 15 }
  };

  return href ? (
    <motion.a 
      href={href} 
      target="_blank" 
      rel="noopener noreferrer"
      className="inline-block" // Ensures the hit box fits the content neatly
      {...motionProps}
    >
      {content}
    </motion.a>
  ) : (
    <motion.div className="inline-block" {...motionProps}>
      {content}
    </motion.div>
  );
}


const bootLines = [
  "$ whoami",
  `> ${profile.name}`,
  "$ role --current",
  `> ${profile.role}`,
  "$ status",
  "> production-ready, accepting new builds"
];

export default function Hero() {
  return (
    <section
      id="hero"
      className="relative flex min-h-screen items-center overflow-hidden bg-ink pt-24"
    >
      <div className="absolute inset-0 bg-grid bg-grid-cell opacity-[0.15]" aria-hidden />
      <div className="pointer-events-none absolute -top-32 right-0 h-96 w-96 rounded-full bg-signal/10 blur-3xl" />

      <div className="relative mx-auto grid w-full max-w-6xl gap-12 px-6 md:grid-cols-2 md:items-center">
        <div>
          <motion.p
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5 }}
            className="mb-4 font-mono text-sm text-signal"
          >
            v{new Date().getFullYear()}.release — available for work
          </motion.p>

          <motion.h1
            initial={{ opacity: 0, y: 16 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6, delay: 0.1 }}
            className="font-display text-4xl font-bold leading-[1.05] text-paper sm:text-5xl lg:text-6xl"
          >
            {profile.name}
          </motion.h1>

          <motion.p
            initial={{ opacity: 0, y: 16 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6, delay: 0.2 }}
            className="mt-5 max-w-lg text-lg text-muted"
          >
            {profile.tagline}
          </motion.p>
          <motion.div
            initial={{ opacity: 0, y: 16 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6, delay: 0.3 }}
            className="mt-8 flex flex-wrap gap-4"
          >
            <a
              href={profile.cv}
              target="_blank"
              rel="noopener noreferrer"
              className="rounded bg-signal px-5 py-3 font-mono text-sm font-medium text-ink transition-transform hover:-translate-y-0.5"
            >
              Download CV
            </a>
            <a
              href="#projects"
              className="rounded border border-line px-5 py-3 font-mono text-sm text-paper transition-colors hover:border-signal hover:text-signal"
            >
              View build log
            </a>
          </motion.div>
          <motion.br
            initial={{ opacity: 0, y: 16 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6, delay: 0.4 }}
          />  
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6, ease: "easeOut" }}
            className="space-y-3" // Adds clean spacing between the text and the icons
          >
            <p className="text-sm text-muted">Let's Connect Here:</p>
            
            <div className="flex flex-wrap gap-4"> 
              <ContactRow
                icon={<FiGithub />}
                label=""
                href={profile.social.github}
              />
              <ContactRow
                icon={<FiLinkedin />}
                label=""
                href={profile.social.linkedin}
              />
              <ContactRow
                icon={<FaMedium />}
                label=""
                href={profile.social.medium}
              />  
              <ContactRow
                icon={<FaMicrosoft />}
                label=""
                href={profile.social.microsoftteams}
              />  
              <ContactRow
                icon={<FaWhatsapp />}
                label=""
                href={profile.social.whatsapp}
              />
              <ContactRow
                icon={<FiInstagram />}
                label=""
                href={profile.social.instagram}
              />
              <ContactRow
                icon={<FiFacebook />}
                label=""
                href={profile.social.facebook}
              />  
            </div>
          </motion.div>
        </div>

        <motion.div
          initial={{ opacity: 0, scale: 0.97 }}
          animate={{ opacity: 1, scale: 1 }}
          transition={{ duration: 0.6, delay: 0.2 }}
          className="rounded-lg border border-line bg-panel/80 shadow-2xl shadow-black/40"
        >
          <div className="flex items-center gap-1.5 border-b border-line px-4 py-3">
            <span className="h-2.5 w-2.5 rounded-full bg-red-400/70" />
            <span className="h-2.5 w-2.5 rounded-full bg-yellow-400/70" />
            <span className="h-2.5 w-2.5 rounded-full bg-signal/70" />
            <span className="ml-3 font-mono text-xs text-muted">boot.log</span>
          </div>
          <div className="space-y-2 p-6 font-mono text-sm">
            {bootLines.map((line, i) => (
              <motion.p
                key={line + i}
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                transition={{ duration: 0.4, delay: 0.5 + i * 0.18 }}
                className={line.startsWith(">") ? "pl-4 text-signal" : "text-paper"}
              >
                {line}
              </motion.p>
            ))}
            <motion.span
              initial={{ opacity: 0 }}
              animate={{ opacity: [0, 1, 0] }}
              transition={{ duration: 1, repeat: Infinity, delay: 1.6 }}
              className="inline-block h-4 w-2 translate-y-0.5 bg-signal"
              aria-hidden
            />
          </div>
        </motion.div>
      </div>
    </section>
  );
}
