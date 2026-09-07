import type { MetadataRoute } from "next";
import { URL_SITUS } from "@/lib/site-config";

// Minta mesin pencari (Google dkk.) mengindeks situs ini
export default function robots(): MetadataRoute.Robots {
  return {
    rules: {
      userAgent: "*",
      allow: "/",
    },
    sitemap: `${URL_SITUS}/sitemap.xml`,
  };
}
