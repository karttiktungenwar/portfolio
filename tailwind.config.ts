import type { Config } from "tailwindcss";

const config: Config = {
  content: [
    "./app/**/*.{ts,tsx}",
    "./components/**/*.{ts,tsx}",
    "./data/**/*.{ts,tsx}"
  ],
  theme: {
    extend: {
      colors: {
        ink: "#0B0F14",
        panel: "#10161D",
        line: "#1E2731",
        signal: "#4CE0B3",
        signal2: "#3B82F6",
        paper: "#E9EEF3",
        muted: "#8593A2"
      },
      fontFamily: {
        display: ["var(--font-display)"],
        body: ["var(--font-body)"],
        mono: ["var(--font-mono)"]
      },
      backgroundImage: {
        grid: "linear-gradient(to right, #1E2731 1px, transparent 1px), linear-gradient(to bottom, #1E2731 1px, transparent 1px)"
      },
      backgroundSize: {
        "grid-cell": "40px 40px"
      }
    }
  },
  plugins: []
};

export default config;
