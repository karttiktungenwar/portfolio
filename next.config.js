/** @type {import('next').NextConfig} */
const nextConfig = {
  output: "export", // Enforces a static HTML build
  basePath: "/portfolio-nextjs", // Matches your GitHub repository subfolder
  images: {
    unoptimized: true, // Required for static exports
  },
};

module.exports = nextConfig;