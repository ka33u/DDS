import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "直驱选型",
  description: "依据拉拔工艺自动计算力、转矩、转速与功率，并生成可导出的直驱电机选型参数表。",
  applicationName: "DirectDrive Selector",
  openGraph: {
    title: "直驱选型",
    description: "从道次工艺到直驱电机规格的一站式工程选型工作台。",
    type: "website",
    locale: "zh_CN",
    images: [{ url: "/og.png", width: 1200, height: 630, alt: "直驱选型" }],
  },
  twitter: {
    card: "summary_large_image",
    title: "直驱选型",
    description: "从拉拔工艺到直驱电机规格",
    images: ["/og.png"],
  },
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="zh-CN"><body>{children}</body></html>;
}
