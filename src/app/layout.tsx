import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "FBA比价查询系统",
  description: "多供应商 FBA 物流价格查询与对比",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="zh-CN">
      <body className="min-h-screen bg-gray-50 text-zinc-900 antialiased">
        {children}
      </body>
    </html>
  );
}
