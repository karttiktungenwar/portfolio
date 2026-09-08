"use client";

import { motion } from "framer-motion";
import { FiExternalLink } from "react-icons/fi";
import { projects } from "@/data/portfolio";
import { SectionEyebrow } from "./About";

export default function Projects() {
  return (
    <section id="projects" className="bg-ink px-6 py-24">
      <div className="mx-auto max-w-6xl">
        <SectionEyebrow index="03" label="build log — projects" />
        <h2 className="mt-6 font-display text-3xl font-bold text-paper sm:text-4xl">
          Shipped builds, not just ideas.
        </h2>

        <div className="mt-12 grid gap-6 md:grid-cols-2">
          {projects.map((project, i) => (
            <motion.article
              key={project.name}
              initial={{ opacity: 0, y: 20 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true, margin: "-60px" }}
              transition={{ duration: 0.5, delay: (i % 2) * 0.1 }}
              className="group rounded-lg border border-line bg-panel/60 p-6 transition-colors hover:border-signal/40"
            >
              <div className="flex items-center justify-between font-mono text-xs">
                <span className="text-muted">{project.version}</span>
                <span
                  className={`rounded border px-2 py-0.5 uppercase ${
                    project.tag === "stable" ? "tag-stable" : "tag-archived"
                  }`}
                >
                  {project.tag}
                </span>
              </div>

              <h3 className="mt-3 font-display text-xl font-semibold text-paper">
                {project.name}
              </h3>
              <p className="font-mono text-xs text-muted">
                {project.company} · {project.period}
              </p>
              <p className="mt-3 text-sm text-muted">{project.description}</p>

              {project.changes.length > 0 && (
                <ul className="mt-4 space-y-1 font-mono text-xs">
                  {project.changes.map((change) => (
                    <li key={change} className="text-signal/90">
                      {change}
                    </li>
                  ))}
                </ul>
              )}

              <a
                href={project.link}
                target="_blank"
                rel="noopener noreferrer"
                className="mt-5 inline-flex items-center gap-1.5 font-mono text-sm text-signal2 transition-colors group-hover:text-signal"
              >
                View release <FiExternalLink aria-hidden />
              </a>
            </motion.article>
          ))}
        </div>
      </div>
    </section>
  );
}
