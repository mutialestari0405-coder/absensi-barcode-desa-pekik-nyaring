import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  output: "standalone",
  /* config options here */
  typescript: {
    ignoreBuildErrors: true,
  },
  reactStrictMode: false,
  // Ikutkan berkas pendukung saat deploy: template database (bootstrap-db.ts)
  // dan skrip export Excel (dipanggil /api/export)
  outputFileTracingIncludes: {
    "/**": ["./prisma/template.db", "./scripts/export_excel.py"],
  },
};

export default nextConfig;
