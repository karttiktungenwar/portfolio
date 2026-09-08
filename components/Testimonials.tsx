"use client";

import { motion } from "framer-motion";
import { testimonials } from "@/data/portfolio";
import { SectionEyebrow } from "./About";

export default function Testimonials() {
  return (
    <section id="testimonials" className="bg-panel/40 px-6 py-24">
      <div className="mx-auto max-w-6xl">
        <SectionEyebrow index="04" label="reviews" />
        <h2 className="mt-6 font-display text-3xl font-bold text-paper sm:text-4xl">
          Notes from the field.
        </h2>

        <div className="mt-12 grid gap-6 md:grid-cols-3">
          {testimonials.map((t, i) => (
            <motion.figure
              key={t.quote.slice(0, 20)}
              initial={{ opacity: 0, y: 20 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true, margin: "-60px" }}
              transition={{ duration: 0.5, delay: i * 0.1 }}
              className="rounded-lg border border-line bg-panel/60 p-6"
            >
              <span className="font-mono text-3xl text-signal/60">&ldquo;</span>
              <blockquote className="text-sm text-muted">{t.quote}</blockquote>
              <figcaption className="mt-4 font-mono text-xs text-paper">
                {t.name} <span className="text-muted">— {t.title}</span>
              </figcaption>
            </motion.figure>
          ))}
        </div>
      </div>
    </section>
  );
}
