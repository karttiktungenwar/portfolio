import type { Metadata } from "next";
import { Space_Grotesk, Inter, JetBrains_Mono } from "next/font/google";
import "./globals.css";

const display = Space_Grotesk({
  subsets: ["latin"],
  variable: "--font-display",
  weight: ["500", "700"]
});

const body = Inter({
  subsets: ["latin"],
  variable: "--font-body",
  weight: ["400", "500", "600"]
});

const mono = JetBrains_Mono({
  subsets: ["latin"],
  variable: "--font-mono",
  weight: ["400", "500"]
});

export const metadata: Metadata = {
 icons: {
    icon: "/favicon.ico", // Standard favicon
    // shortcut: "/shortcut-icon.png", // Optional alternate formats
    // apple: "/apple-touch-icon.png",  // Optional iOS icon
  },
  title: "Kartik Tungenwar — Mobile App Developer & Designer",
  description:
    "Career portfolio of Kartik Tungenwar, Mobile App Developer & Designer specializing in secure, production-grade mobile applications.",
  metadataBase: new URL("https://karttiktungenwar.github.io"),
  openGraph: {
    title: "Kartik Tungenwar — Mobile App Developer & Designer",
    description:
      "Career portfolio of Kartik Tungenwar, Mobile App Developer & Designer specializing in secure, production-grade mobile applications.",
    url: "https://karttiktungenwar.github.io/portfolio/",
    type: "website"
  }
};

export default function RootLayout({
  children
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en" className={`${display.variable} ${body.variable} ${mono.variable}`}>
      <body className="bg-ink text-paper font-body antialiased selection:bg-signal selection:text-ink">
        {children}
      </body>
    </html>
  );
}
