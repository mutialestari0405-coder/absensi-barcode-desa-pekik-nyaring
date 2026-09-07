import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";
import { Toaster } from "@/components/ui/toaster";
import { SITE, URL_SITUS } from "@/lib/site-config";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  metadataBase: new URL(URL_SITUS),
  title: `Absensi Perangkat Desa ${SITE.namaDesa} — Sistem Absensi QR`,
  description: `Absensi Perangkat Desa ${SITE.namaDesa}, Kecamatan ${SITE.kecamatan}, Kabupaten ${SITE.kabupaten}. Sistem absensi online berbasis QR Code dengan kamera scanner — praktis, cepat, tanpa aplikasi tambahan, langsung dari HP.`,
  keywords: [
    `absensi perangkat desa ${SITE.namaDesa}`,
    `absensi desa ${SITE.namaDesa}`,
    "absensi qr code desa",
    "absensi online perangkat desa",
    "sistem absensi qr",
    SITE.namaDesa,
    `kecamatan ${SITE.kecamatan}`,
    SITE.kabupaten,
    "bengkulu tengah",
  ],
  authors: [{ name: SITE.pemerintahDesa }],
  icons: {
    icon: SITE.logo,
  },
  openGraph: {
    title: `Absensi Perangkat Desa ${SITE.namaDesa}`,
    description: `Sistem absensi online berbasis QR Code untuk Perangkat Desa ${SITE.namaDesa}, Kecamatan ${SITE.kecamatan}, Kabupaten ${SITE.kabupaten}.`,
    type: "website",
    locale: "id_ID",
    siteName: `Absensi Perangkat Desa ${SITE.namaDesa}`,
  },
  robots: {
    index: true,
    follow: true,
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="id" suppressHydrationWarning>
      <body
        className={`${geistSans.variable} ${geistMono.variable} antialiased bg-background text-foreground`}
      >
        {children}
        <Toaster />
      </body>
    </html>
  );
}
