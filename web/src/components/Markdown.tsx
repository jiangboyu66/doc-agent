import { Marked } from "marked";
import { useMemo } from "react";

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/**
 * 模型输出可能夹带文档里的内容（包括恶意构造的 HTML），所以：
 * 原始 HTML 一律转义；链接只允许 http(s) 与站内 /api 路径。
 */
const md = new Marked({
  gfm: true,
  breaks: true,
  renderer: {
    html({ text }) {
      return esc(text);
    },
    link({ href, text }) {
      const ok = /^(https?:\/\/|\/api\/)/i.test(href ?? "");
      if (!ok) return esc(text);
      return `<a href="${esc(href)}" target="_blank" rel="noopener noreferrer">${esc(text)}</a>`;
    },
    image({ text }) {
      return esc(text ? `[图片：${text}]` : "[图片]");
    },
  },
});

export function Markdown({ text, className }: { text: string; className?: string }) {
  const html = useMemo(() => md.parse(text, { async: false }) as string, [text]);
  return <div className={`md ${className ?? ""}`} dangerouslySetInnerHTML={{ __html: html }} />;
}
