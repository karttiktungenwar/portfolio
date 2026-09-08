"use client";

import { motion } from "framer-motion";
import { education, experience } from "@/data/portfolio";
import { SectionEyebrow } from "./About";

function statusClass(status: string) {
  if (status === "current") return "tag-current";
  if (status === "archived") return "tag-archived";
  return "tag-stable";
}

export default function Experience() {
  return (
    <section id="experience" className="bg-panel/40 px-6 py-24">
      <div className="mx-auto max-w-4xl">
        <SectionEyebrow index="02" label="release log — experience and education" />
        <h2 className="mt-6 font-display text-3xl font-bold text-paper sm:text-4xl">
          Every role, tagged like a build.
        </h2>


      {/* Grid Container: Stacks on mobile (grid-cols-1), side-by-side on large screens (lg:grid-cols-2) */}
<div className="mt-16 grid grid-cols-1 gap-12 lg:grid-cols-2">
  
  {/* Left Side: Experience */}
        <div className="mt-16">
          <p className="font-mono text-xs uppercase tracking-widest text-muted">experience</p>
        <ol className="relative mt-12 space-y-8 border-l border-line pl-8">
          {experience.map((job, i) => (
            <motion.li
              key={job.role + job.period}
              initial={{ opacity: 0, y: 16 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true, margin: "-60px" }}
              transition={{ duration: 0.5, delay: i * 0.08 }}
              className="relative"
            >
              <span className="absolute -left-[2.35rem] top-1.5 h-3 w-3 rounded-full border-2 border-ink bg-signal" />
              <div className="flex flex-wrap items-center gap-3">
                <span className="font-mono text-xs text-muted">{job.version}</span>
                <span
                  className={`rounded border px-2 py-0.5 font-mono text-xs uppercase ${statusClass(
                    job.status
                  )}`}
                >
                  {job.status}
                </span>
              </div>
              <h3 className="mt-2 font-display text-xl font-semibold text-paper">
                {job.role}
              </h3>
              <p className="text-muted">{job.company}</p>
              <p className="mt-1 font-mono text-xs text-muted">{job.period}</p>
            </motion.li>
          ))}
        </ol>
        </div>

        <div className="mt-16">
          <p className="font-mono text-xs uppercase tracking-widest text-muted">education</p>
          <br className="mt-2 border-t border-line" />
          <ol className="mt-6 space-y-6 border-l border-line pl-8">
            {education.map((edu, i) => (
              <motion.li
                key={edu.degree}
                initial={{ opacity: 0, y: 12 }}
                whileInView={{ opacity: 1, y: 0 }}
                viewport={{ once: true, margin: "-40px" }}
                transition={{ duration: 0.45, delay: i * 0.08 }}
                className="relative"
              >
                <span className="absolute -left-[2.35rem] top-1.5 h-2.5 w-2.5 rounded-full border-2 border-ink bg-signal2" />
                <h4 className="font-display text-lg font-semibold text-paper">
                  {edu.degree}
                </h4>
                <p className="text-muted">{edu.school}</p>
                <p className="mt-1 font-mono text-xs text-muted">{edu.period}</p>
              </motion.li>
            ))}
          </ol>
        </div>
        </div>
      </div>
    </section>
  );
}
