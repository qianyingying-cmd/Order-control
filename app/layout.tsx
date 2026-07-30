import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Wilson Order Control Tower",
  description: "Transparent inventory, cash flow and order review dashboard.",
  openGraph: {
    title: "Wilson Order Control Tower",
    description: "Inventory · OIH · Sales · Cash · Decision",
    images: [{ url: "/og.png", width: 1745, height: 909 }],
  },
  twitter: {
    card: "summary_large_image",
    title: "Wilson Order Control Tower",
    description: "Inventory · OIH · Sales · Cash · Decision",
    images: ["/og.png"],
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
