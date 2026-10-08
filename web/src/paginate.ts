/**
 * 快速预览的分页
 *
 * docx-preview 不会自己分页：只在分页符、分节符，以及文档里记录的"上次分页位置"处分页。
 * 文档内容增减后（例如扩写摘要），那一页会被撑长，两栏的底部也对不齐。这里按 Word 的规则重新分页：
 *   - 每页正文区高度固定；多栏节先填满左栏再填右栏（column-fill: auto），节的最后一页两栏平衡；
 *   - 放不下的段落按行拆到下一页：前半段末行两端对齐，续段不再首行缩进、没有段前距；
 *   - 表格、图片等整体移到下一页；比一整页还高的内容单独占一页（页面随之变长）；
 *   - 页眉页脚取原渲染中对应页（首页 / 偶数页 / 奇数页）的那一份。
 * 内容本来就放得下的页面（例如逐页保留原排版的转换结果）保持不变。
 */

import JSZip from "jszip";

const MAX_PAGES = 3000;

/**
 * 去掉文档里记录的"上次分页位置"（lastRenderedPageBreak）后再渲染正文：由 paginate() 重新分页。
 * （docx-preview 的 ignoreLastRenderedPageBreak 选项会把每个分节都拆成单独一页，不能用。）
 */
export async function stripRenderedBreaks(blob: Blob): Promise<Blob> {
  return (await prepareDocx(blob)).blob;
}

/** 页眉页脚中的页码域占位符：docx-preview 不渲染域，渲染后由 numberPages() 换成实际页码 */
const PH_PAGE = "\uE000PAGE\uE000";
const PH_NUMPAGES = "\uE000NUMPAGES\uE000";

function markFields(xml: string): string {
  // 简单域 <w:fldSimple w:instr=" PAGE ">…</w:fldSimple>
  xml = xml.replace(/<(\w+):fldSimple\b([^>]*?)\binstr="([^"]*)"([^>]*?)(?:\/>|>([\s\S]*?)<\/\1:fldSimple>)/g, (m, p: string, _a: string, instr: string, _b: string, inner?: string) => {
    const ph = /\bNUMPAGES\b/i.test(instr) ? PH_NUMPAGES : /\bPAGE\b/i.test(instr) ? PH_PAGE : null;
    if (!ph) return m;
    const rPr = inner?.match(new RegExp(`<${p}:rPr>[\\s\\S]*?</${p}:rPr>`))?.[0] ?? "";
    return `<${p}:r>${rPr}<${p}:t>${ph}</${p}:t></${p}:r>`;
  });
  // 复杂域：PAGE / NUMPAGES 指令之后、separate 之后的第一段结果文字
  xml = xml.replace(/(<(\w+):instrText[^>]*>\s*(PAGE|NUMPAGES)\b[^<]*<\/\2:instrText>(?:(?!fldCharType="end")[\s\S])*?fldCharType="separate"(?:(?!fldCharType="end")[\s\S])*?<\2:t(?:\s[^>]*)?>)[^<]*(<\/\2:t>)/g,
    (_m, pre: string, _p: string, kind: string, post: string) => pre + (kind.toUpperCase() === "NUMPAGES" ? PH_NUMPAGES : PH_PAGE) + post);
  return xml;
}

/**
 * 渲染前的准备：
 *   - 正文渲染用的副本去掉旧分页位置（由 paginate 重新分页）；
 *   - 页眉页脚中的页码域换成占位符（两份渲染都需要），渲染后填入实际页码；
 *   - 读取"连续分节前不平衡分栏"兼容选项（期刊排版会打开）与起始页码。
 */
export async function prepareDocx(blob: Blob): Promise<{ blob: Blob; plain: Blob; balance: boolean; pageStart: number }> {
  const zip = await JSZip.loadAsync(blob);
  const settings = await zip.file("word/settings.xml")?.async("string");
  const balance = !(settings && /<\w+:noColumnBalance(\s|\/|>)/.test(settings) && !/<\w+:noColumnBalance\s+\w+:val="(0|false)"/.test(settings));
  const mime = blob.type || "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  let changed = false;
  for (const name of Object.keys(zip.files)) {
    if (!/^word\/(header|footer)\d*\.xml$/.test(name)) continue;
    const x = await zip.file(name)!.async("string");
    const y = markFields(x);
    if (y !== x) { zip.file(name, y); changed = true; }
  }
  const f = zip.file("word/document.xml");
  const xml = f ? await f.async("string") : "";
  const pageStart = Number(/<\w+:pgNumType\b[^>]*\w+:start="(\d+)"/.exec(xml)?.[1] ?? 1) || 1;
  const plain = changed ? await zip.generateAsync({ type: "blob", mimeType: mime }) : blob;
  if (!xml.includes("lastRenderedPageBreak")) return { blob: plain, plain, balance, pageStart };
  zip.file("word/document.xml", xml.replace(/<w:lastRenderedPageBreak\s*\/>/g, ""));
  return { blob: await zip.generateAsync({ type: "blob", mimeType: mime }), plain, balance, pageStart };
}

/**
 * "最小值"行距（w:lineRule="atLeast"）：docx-preview 把它渲染成 line-height: calc(100% + N pt)，
 * 即字号再加 N pt——约为 Word 里实际行距的 1.9 倍，和固定行距 / 单倍行距的段落排在一起时行距忽松忽紧。
 * Word 的语义是"至少 N pt，字体自然行高更大时取自然行高"，这里改写成 max(N pt, 1.15em)。
 * 同时处理样式表（<style>）和段落上的内联样式。
 */
const AT_LEAST = /calc\(\s*100%\s*\+\s*(-?[\d.]+)(pt|px)\s*\)/g; // 内联样式经浏览器规范化后单位是 px
const atLeastValue = (_m: string, n: string, unit: string) => `max(${n}${unit}, 1.15em)`;
export function fixAtLeastSpacing(root: HTMLElement): void {
  root.querySelectorAll("style").forEach((st) => {
    const t = st.textContent ?? "";
    if (t.includes("calc(")) st.textContent = t.replace(/line-height\s*:\s*calc\([^;}]*\)/g, (m) => m.replace(AT_LEAST, atLeastValue));
  });
  root.querySelectorAll<HTMLElement>('[style*="calc("]').forEach((el) => {
    const v = el.style.lineHeight;
    if (v && v.startsWith("calc(")) el.style.lineHeight = v.replace(AT_LEAST, atLeastValue);
  });
}

/** 填入页眉页脚中的页码，并在每页下方加"第 n 页 / 共 N 页"标注；返回总页数 */
export function numberPages(root: HTMLElement, start = 1): number {
  const wrapper = root.querySelector<HTMLElement>(".docx-wrapper") ?? root;
  wrapper.querySelectorAll(".page-label").forEach((x) => x.remove());
  const pages = Array.from(wrapper.children).filter((x): x is HTMLElement => x instanceof HTMLElement && x.tagName === "SECTION");
  pages.forEach((pg, i) => {
    const walker = document.createTreeWalker(pg, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const t = n.textContent ?? "";
      if (t.includes("\uE000")) n.textContent = t.split(PH_PAGE).join(String(start + i)).split(PH_NUMPAGES).join(String(pages.length));
    }
    pg.dataset.page = String(i + 1);
    const label = document.createElement("div");
    label.className = "page-label";
    label.textContent = `第 ${i + 1} 页 / 共 ${pages.length} 页`;
    pg.after(label);
  });
  return pages.length;
}

function pxHeight(el: HTMLElement): number {
  const v = parseFloat(getComputedStyle(el).minHeight);
  return Number.isFinite(v) && v > 0 ? v : el.offsetHeight;
}

/** 文本偏移 → (文本节点, 节点内偏移) */
function locate(root: Node, k: number): { node: Node; offset: number } | null {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let pos = 0;
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const len = n.textContent?.length ?? 0;
    if (pos + len >= k) return { node: n, offset: k - pos };
    pos += len;
  }
  return null;
}

/** 复制段落，只保留前 k 个字符（head）或第 k 个字符之后的部分（tail） */
function cut(el: HTMLElement, k: number, part: "head" | "tail"): HTMLElement {
  const cl = el.cloneNode(true) as HTMLElement;
  const at = locate(cl, k);
  if (!at) return cl;
  const r = document.createRange();
  if (part === "head") { r.setStart(at.node, at.offset); r.setEnd(cl, cl.childNodes.length); }
  else { r.setStart(cl, 0); r.setEnd(at.node, at.offset); }
  r.deleteContents();
  return cl;
}

const CJK = /[⺀-鿿豈-﫿＀-￯　-〿]/;

export function paginate(content: HTMLElement, hf: HTMLElement | null, opts: { balance?: boolean } = {}): number {
  const balance = opts.balance !== false;
  const wrapper = content.querySelector<HTMLElement>(".docx-wrapper") ?? content;
  const groups = Array.from(wrapper.children).filter((x): x is HTMLElement => x instanceof HTMLElement && x.tagName === "SECTION");
  const hfPages = hf ? Array.from(hf.querySelectorAll<HTMLElement>("section")) : [];
  const pickHF = (i: number): HTMLElement | null => {
    if (!hfPages.length) return null;
    if (i < hfPages.length) return hfPages[i];
    const same = hfPages.filter((_, j) => j >= 1 && j % 2 === i % 2);
    return same[same.length - 1] ?? hfPages[hfPages.length - 1];
  };
  let pageNo = 0, made = 0;

  for (const src of groups) {
    const pageH = pxHeight(src);
    // 本来就是一页的内容不动。文档自己逐页分页时（逐页保留原排版的转换结果、手动分页的文档），
    // 每一段在浏览器里可能因字宽差异略长于一页，只有明显超过两页的才重新分页
    const limit = groups.length > 1 ? 2 : 1.02;
    if (src.scrollHeight <= pageH * limit + 2) { pageNo++; continue; }

    let page!: HTMLElement, body!: HTMLElement;
    const remaining = (after: HTMLElement | null) =>
      body.getBoundingClientRect().bottom - (after ? after.getBoundingClientRect().bottom : body.getBoundingClientRect().top);
    const newPage = () => {
      if (made++ > MAX_PAGES) throw new Error("分页过多");
      page = src.cloneNode(false) as HTMLElement;
      page.style.height = `${pageH}px`;
      const ref = pickHF(pageNo) ?? src;
      const header = ref.querySelector<HTMLElement>(":scope > header") ?? src.querySelector<HTMLElement>(":scope > header");
      const footer = ref.querySelector<HTMLElement>(":scope > footer") ?? src.querySelector<HTMLElement>(":scope > footer");
      body = document.createElement("div");
      body.className = "pg-body";
      body.style.cssText = "flex:1 1 auto;min-height:0;overflow:hidden;position:relative;z-index:1;";
      if (header) page.appendChild(header.cloneNode(true));
      page.appendChild(body);
      if (footer) page.appendChild(footer.cloneNode(true));
      wrapper.insertBefore(page, src);
      pageNo++;
    };
    newPage();

    // 放不放得下：看最后放入的元素——多栏时它的任何一段落进了第 cols+1 栏（容器右侧之外）即溢出；
    // 单栏时看它的底边。不用 scrollWidth：略宽于栏的公式图片会让 scrollWidth 变大，却并没有溢出到下一栏
    const fits = (c: HTMLElement, cols: number) => {
      const el = c.lastElementChild as HTMLElement | null;
      if (!el) return true;
      const box = c.getBoundingClientRect();
      const rects = Array.from(el.getClientRects());
      if (!rects.length) return true;
      if (cols >= 2) {
        const colW = box.width / cols;
        return rects.every((r) => r.left < box.right - 0.5 * colW && r.bottom <= box.bottom + 1);
      }
      return rects.every((r) => r.bottom <= box.bottom + 1);
    };

    // 按行拆段落：二分查找放得下的最长前缀（只在词边界处断开）
    const split = (el: HTMLElement, c: HTMLElement, cols: number): [HTMLElement, HTMLElement] | null => {
      if (el.tagName !== "P") return null;
      const text = el.textContent ?? "";
      if (text.trim().length < 2) return null;
      const cuts: number[] = [];
      for (let i = 1; i < text.length; i++) {
        const a = text[i - 1];
        if (/\s/.test(a) || a === "-" || a === "‐" || CJK.test(a) || CJK.test(text[i])) cuts.push(i);
      }
      let lo = 0, hi = cuts.length - 1, best = -1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        const head = cut(el, cuts[mid], "head");
        c.appendChild(head);
        const ok = fits(c, cols);
        c.removeChild(head);
        if (ok) { best = mid; lo = mid + 1; } else hi = mid - 1;
      }
      if (best < 0) return null;
      const k = cuts[best];
      const head = cut(el, k, "head"), tail = cut(el, k, "tail");
      if (getComputedStyle(el).textAlign === "justify") head.style.textAlignLast = "justify";
      head.style.marginBottom = "0";
      tail.style.textIndent = "0";
      tail.style.marginTop = "0";
      return [head, tail];
    };

    const children = Array.from(src.children).filter((x): x is HTMLElement => x instanceof HTMLElement && x.tagName !== "HEADER" && x.tagName !== "FOOTER");
    let last: HTMLElement | null = null; // 本页最后一个容器
    let lastCols = 0;
    /** 推迟到下一页顶部的通栏图表（与 LaTeX 的页顶浮动体一致：文字继续排满本页，不留空白） */
    const deferred: HTMLElement[] = [];
    const isFloatBand = (a: HTMLElement) => {
      if (a.tagName !== "ARTICLE" || (parseInt(a.style.columnCount || "1", 10) || 1) > 1) return false;
      if (!a.querySelector("img,svg,table")) return false;
      // 题注不长、没有成段的正文（表格内的文字不算）
      return Array.from(a.children).every((k) => k.tagName === "TABLE" || (k.textContent ?? "").length < 700);
    };
    const flushDeferred = () => {
      while (deferred.length) placeArticle(deferred.shift()!, null);
    };

    const placeArticle = (child: HTMLElement, reuse: HTMLElement | null) => {
      const isArticle = child.tagName === "ARTICLE";
      const cols = isArticle ? parseInt(child.style.columnCount || "1", 10) || 1 : 1;
      const open = () => {
        const c = (isArticle ? child.cloneNode(false) : document.createElement("div")) as HTMLElement;
        c.style.columnFill = "auto";
        c.style.height = `${Math.max(1, remaining(last))}px`;
        c.style.overflow = "hidden";
        c.style.marginBottom = "0";
        body.appendChild(c);
        return c;
      };
      // 节结束（或换页前的最后一段）：多栏平衡、高度收紧到内容
      const close = (c: HTMLElement, full: boolean) => {
        if (!full) {
          if (balance || cols < 2) {
            c.style.height = "auto";
            c.style.columnFill = "balance";
          } else {
            // 不平衡：保持整栏高度、先排满左栏再排右栏；内容全在左栏时，高度收紧到左栏内容（后面的节紧接其下）
            const box = c.getBoundingClientRect();
            const colW = box.width / cols;
            let maxBottom = box.top, inLeftOnly = true;
            for (const el of Array.from(c.children)) {
              for (const r of Array.from(el.getClientRects())) {
                if (r.left >= box.left + colW * 0.75) inLeftOnly = false;
                maxBottom = Math.max(maxBottom, r.bottom);
              }
            }
            if (inLeftOnly) c.style.height = `${Math.max(1, maxBottom - box.top)}px`;
          }
        }
        last = c;
        lastCols = cols;
      };
      let c: HTMLElement;
      if (reuse) {
        // 接着上一段同栏数的正文排（中间的通栏图表推迟到了下一页）
        reuse.style.columnFill = "auto";
        reuse.style.height = `${Math.max(1, body.getBoundingClientRect().bottom - reuse.getBoundingClientRect().top)}px`;
        c = reuse;
      } else c = open();
      const queue: HTMLElement[] = isArticle ? Array.from(child.children) as HTMLElement[] : [child];
      while (queue.length) {
        const item = queue.shift()!;
        c.appendChild(item);
        if (fits(c, cols)) continue;
        c.removeChild(item);
        const parts = split(item, c, cols);
        if (parts) {
          c.appendChild(parts[0]);
          queue.unshift(parts[1]);
        } else if (!c.childElementCount && !last) {
          // 一整页都放不下（大表格、大图）：单独占一页，页面随之变长
          c.appendChild(item);
          c.style.height = "auto";
          c.style.overflow = "visible";
          page.style.height = "";
          body.style.overflow = "visible";
        } else {
          queue.unshift(item);
        }
        close(c, true);
        newPage();
        last = null;
        lastCols = 0;
        flushDeferred();
        c = open();
      }
      close(c, false);
    };

    for (const child of children) {
      const cols = child.tagName === "ARTICLE" ? parseInt(child.style.columnCount || "1", 10) || 1 : 1;
      // 通栏图表：页面上已经有内容时推迟到下一页顶部
      if (isFloatBand(child) && last) { deferred.push(child); continue; }
      const reuse = deferred.length && cols > 1 && last && lastCols === cols && body.contains(last) ? last : null;
      placeArticle(child, reuse);
    }
    // 文末还有推迟的图表：放得下就放在本页，否则下一页
    flushDeferred();
    // 空的末页（恰好排满时）去掉
    if (body && !body.textContent?.trim() && !body.querySelector("img,svg,table")) page.remove();
    src.remove();
  }
  return pageNo;
}
