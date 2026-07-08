import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "会议记录辅助工具",
  description: "提取议题内容并生成会议记录 Word 文档"
};

export default function RootLayout({
  children
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
