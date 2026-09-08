"use client";

import { motion } from "framer-motion";
import { about, skills } from "@/data/portfolio";

export default function About() {
  return (
    <section id="about" className="bg-ink px-6 py-24">
      <div className="mx-auto max-w-6xl">
        <SectionEyebrow index="01" label="about" />

        <div className="mt-10 grid gap-12 md:grid-cols-2">
          <div>
            <h2 className="font-display text-3xl font-bold text-paper sm:text-4xl">
              A developer who ships, not just codes.
            </h2>
            <p className="mt-6 text-muted">{about.summary}</p>
            <p className="mt-4 text-muted">{about.detail}</p>
          </div>

          <div className="space-y-5">
            {skills.map((skill, i) => (
              <motion.div
                key={skill.name}
                initial={{ opacity: 0, x: 20 }}
                whileInView={{ opacity: 1, x: 0 }}
                viewport={{ once: true, margin: "-60px" }}
                transition={{ duration: 0.5, delay: i * 0.06 }}
              >
                <div className="mb-1.5 flex items-center justify-between font-mono text-sm">
                  <span className="text-paper">{skill.name}</span>
                  <span className="text-signal">{skill.level}%</span>
                </div>
                <div className="h-1.5 w-full overflow-hidden rounded-full bg-line">
                  <motion.div
                    initial={{ width: 0 }}
                    whileInView={{ width: `${skill.level}%` }}
                    viewport={{ once: true }}
                    transition={{ duration: 0.9, delay: i * 0.06 + 0.1, ease: "easeOut" }}
                    className="h-full rounded-full bg-gradient-to-r from-signal2 to-signal"
                  />
                </div>
              </motion.div>
            ))}
          </div>
        </div>
      </div>
    </section>
  );
}

export function SectionEyebrow({ index, label }: { index: string; label: string }) {
  return (
    <div className="flex items-center gap-3 font-mono text-sm text-muted">
      <span className="text-signal">{index}</span>
      <span className="h-px flex-1 max-w-[40px] bg-line" />
      <span className="uppercase tracking-widest">{label}</span>
    </div>
  );
}
