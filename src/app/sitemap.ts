import type { MetadataRoute } from "next";
import { URL_SITUS } from "@/lib/site-config";

// Peta situs untuk membantu Google menemukan halaman
export default function sitemap(): MetadataRoute.Sitemap {
  return [
    {
      url: URL_SITUS,
      lastModified: new Date(),
      changeFrequency: "weekly",
      priority: 1,
    },
  ];
}
