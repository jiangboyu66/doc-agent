/**
 * 外科手术式 XML 解析 / 序列化
 *
 * 为什么不用普通 DOM 库：DOM 解析再序列化会改写属性引号、命名空间声明顺序、实体写法、空白、
 * standalone 声明……哪怕一个字都没改，输出的字节也和原文件不同。要做到"未修改的内容字节级保真"，
 * 就必须保留每个节点在源文本中的原始片段。
 *
 * 做法：解析时每个节点都记住自己在源文本中的 [start, end) 区间（src）。修改某个节点时把它和
 * 所有祖先标记为 dirty。序列化时：未 dirty 的节点直接输出源文本切片（字节级一致）；dirty 的节点
 * 只重新拼接自身的开/闭标签，子节点继续按同样规则递归——所以一个段落被改，只有这个段落内被
 * 碰到的 run 会重新生成，同一段落里其它 run 以及文档其余部分都原样输出。
 */

export type XNode = XElement | XText | XRaw;

export class XText {
  readonly type = "text" as const;
  parent: XElement | XDocument | null = null;
  dirty = false;
  constructor(
    public raw: string,
    public src: [number, number] | null = null
  ) {}
  get text(): string {
    return decodeEntities(this.raw);
  }
  set text(v: string) {
    this.raw = encodeText(v);
    markDirty(this);
  }
}

/** 注释、处理指令、CDATA、DOCTYPE —— 原样保留 */
export class XRaw {
  readonly type = "raw" as const;
  parent: XElement | XDocument | null = null;
  dirty = false;
  constructor(
    public raw: string,
    public src: [number, number] | null = null
  ) {}
}

export class XElement {
  readonly type = "el" as const;
  parent: XElement | XDocument | null = null;
  children: XNode[] = [];
  dirty = false;
  /** 属性被改过时才重新生成开标签，否则沿用 openRaw */
  attrsDirty = false;
  src: [number, number] | null = null;
  openRaw: string | null = null;
  selfClosing = false;
  constructor(
    public name: string,
    public attrs: Array<[string, string]> = []
  ) {}

  get local(): string {
    const i = this.name.indexOf(":");
    return i === -1 ? this.name : this.name.slice(i + 1);
  }
  get prefix(): string {
    const i = this.name.indexOf(":");
    return i === -1 ? "" : this.name.slice(0, i);
  }
  getAttr(name: string): string | undefined {
    const a = this.attrs.find(([k]) => k === name);
    return a ? decodeEntities(a[1]) : undefined;
  }
  setAttr(name: string, value: string): void {
    const enc = encodeAttr(value);
    const a = this.attrs.find(([k]) => k === name);
    if (a) {
      if (a[1] === enc) return;
      a[1] = enc;
    } else this.attrs.push([name, enc]);
    this.attrsDirty = true;
    markDirty(this);
  }
  removeAttr(name: string): void {
    const i = this.attrs.findIndex(([k]) => k === name);
    if (i === -1) return;
    this.attrs.splice(i, 1);
    this.attrsDirty = true;
    markDirty(this);
  }
  elements(): XElement[] {
    return this.children.filter((c): c is XElement => c.type === "el");
  }
  child(name: string): XElement | undefined {
    return this.children.find((c): c is XElement => c.type === "el" && c.name === name);
  }
  childrenNamed(name: string): XElement[] {
    return this.children.filter((c): c is XElement => c.type === "el" && c.name === name);
  }
  /** 深度优先遍历所有后代元素 */
  *descendants(): Generator<XElement> {
    for (const c of this.children) {
      if (c.type === "el") {
        yield c;
        yield* c.descendants();
      }
    }
  }
  closest(name: string): XElement | null {
    let p = this.parent;
    while (p && p.type === "el") {
      if (p.name === name) return p;
      p = p.parent;
    }
    return null;
  }
}

export class XDocument {
  readonly type = "doc" as const;
  children: XNode[] = [];
  dirty = false;
  constructor(public source: string) {}
  get root(): XElement {
    const r = this.children.find((c): c is XElement => c.type === "el");
    if (!r) throw new Error("XML 没有根元素");
    return r;
  }
}

export function markDirty(n: XNode | XDocument): void {
  let cur: XNode | XDocument | null = n;
  while (cur) {
    if (cur.dirty) {
      // 祖先已 dirty 时仍需继续向上，因为祖先链可能在 dirty 之后被重新挂载
    }
    cur.dirty = true;
    cur = cur.type === "doc" ? null : cur.parent;
  }
}

// ---------------- 实体编解码 ----------------

const NAMED: Record<string, string> = { lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" };

export function decodeEntities(s: string): string {
  if (s.indexOf("&") === -1) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|lt|gt|amp|quot|apos);/g, (_m, e: string) => {
    if (e[0] === "#") {
      const code = e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return String.fromCodePoint(code);
    }
    return NAMED[e];
  });
}

export function encodeText(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

export function encodeAttr(s: string): string {
  return encodeText(s).replace(/"/g, "&quot;");
}

// ---------------- 解析 ----------------

const NAME_END = /[\s/>]/;

export function parseXml(source: string): XDocument {
  const doc = new XDocument(source);
  const stack: Array<XElement | XDocument> = [doc];
  const top = () => stack[stack.length - 1];
  const append = (n: XNode) => {
    const p = top();
    n.parent = p;
    p.children.push(n);
  };

  let i = 0;
  const len = source.length;
  while (i < len) {
    const lt = source.indexOf("<", i);
    if (lt === -1) {
      if (i < len) append(new XText(source.slice(i), [i, len]));
      break;
    }
    if (lt > i) append(new XText(source.slice(i, lt), [i, lt]));

    if (source.startsWith("<?", lt)) {
      const end = source.indexOf("?>", lt);
      if (end === -1) throw new Error(`XML 解析失败：处理指令未闭合 @${lt}`);
      append(new XRaw(source.slice(lt, end + 2), [lt, end + 2]));
      i = end + 2;
      continue;
    }
    if (source.startsWith("<!--", lt)) {
      const end = source.indexOf("-->", lt);
      if (end === -1) throw new Error(`XML 解析失败：注释未闭合 @${lt}`);
      append(new XRaw(source.slice(lt, end + 3), [lt, end + 3]));
      i = end + 3;
      continue;
    }
    if (source.startsWith("<![CDATA[", lt)) {
      const end = source.indexOf("]]>", lt);
      if (end === -1) throw new Error(`XML 解析失败：CDATA 未闭合 @${lt}`);
      append(new XRaw(source.slice(lt, end + 3), [lt, end + 3]));
      i = end + 3;
      continue;
    }
    if (source.startsWith("<!", lt)) {
      const end = source.indexOf(">", lt);
      append(new XRaw(source.slice(lt, end + 1), [lt, end + 1]));
      i = end + 1;
      continue;
    }
    if (source[lt + 1] === "/") {
      const end = source.indexOf(">", lt);
      if (end === -1) throw new Error(`XML 解析失败：闭合标签未结束 @${lt}`);
      const name = source.slice(lt + 2, end).trim();
      const cur = top();
      if (cur.type !== "el" || cur.name !== name) {
        throw new Error(`XML 解析失败：闭合标签 </${name}> 与 <${cur.type === "el" ? cur.name : "文档"}> 不匹配 @${lt}`);
      }
      cur.src = [cur.src![0], end + 1];
      stack.pop();
      i = end + 1;
      continue;
    }

    // 开始标签：手动扫描以正确处理属性值中的 '>'
    let j = lt + 1;
    while (j < len && !NAME_END.test(source[j])) j++;
    const name = source.slice(lt + 1, j);
    if (!name) throw new Error(`XML 解析失败：空标签名 @${lt}`);
    const attrs: Array<[string, string]> = [];
    let selfClosing = false;
    while (j < len) {
      const ch = source[j];
      if (ch === ">") {
        j++;
        break;
      }
      if (ch === "/" && source[j + 1] === ">") {
        selfClosing = true;
        j += 2;
        break;
      }
      if (/\s/.test(ch)) {
        j++;
        continue;
      }
      const eq = source.indexOf("=", j);
      if (eq === -1) throw new Error(`XML 解析失败：属性缺少 '=' @${j}`);
      const an = source.slice(j, eq).trim();
      let q = eq + 1;
      while (/\s/.test(source[q])) q++;
      const quote = source[q];
      if (quote !== '"' && quote !== "'") throw new Error(`XML 解析失败：属性值缺少引号 @${q}`);
      const close = source.indexOf(quote, q + 1);
      if (close === -1) throw new Error(`XML 解析失败：属性值未闭合 @${q}`);
      // 统一以双引号语义存储已编码值
      let v = source.slice(q + 1, close);
      if (quote === "'") v = v.replace(/"/g, "&quot;");
      attrs.push([an, v]);
      j = close + 1;
    }
    const el = new XElement(name, attrs);
    el.openRaw = source.slice(lt, j);
    el.selfClosing = selfClosing;
    el.src = [lt, j];
    append(el);
    if (!selfClosing) stack.push(el);
    i = j;
  }
  if (stack.length !== 1) {
    const unclosed = stack[stack.length - 1];
    throw new Error(`XML 解析失败：元素 <${unclosed.type === "el" ? unclosed.name : "?"}> 未闭合`);
  }
  return doc;
}

/** 解析一个 XML 片段（可含多个顶层节点），节点不带 src，全部视为新节点 */
export function parseFragment(xml: string): XNode[] {
  const doc = parseXml(`<__f>${xml}</__f>`);
  const wrapper = doc.root;
  const nodes = wrapper.children;
  const strip = (n: XNode) => {
    n.src = null;
    n.dirty = true;
    if (n.type === "el") {
      n.openRaw = null;
      n.attrsDirty = true;
      n.children.forEach(strip);
    }
  };
  nodes.forEach((n) => {
    strip(n);
    n.parent = null;
  });
  return nodes;
}

// ---------------- 序列化 ----------------

function buildOpen(el: XElement, selfClose: boolean): string {
  let s = "<" + el.name;
  for (const [k, v] of el.attrs) s += ` ${k}="${v}"`;
  return s + (selfClose ? "/>" : ">");
}

export function serializeNode(n: XNode, source: string): string {
  if (!n.dirty && n.src) return source.slice(n.src[0], n.src[1]);
  if (n.type === "text" || n.type === "raw") return n.raw;
  const el = n;
  if (el.children.length === 0) {
    if (el.selfClosing && !el.attrsDirty && el.openRaw) return el.openRaw;
    if (!el.selfClosing && !el.attrsDirty && el.openRaw) return el.openRaw + `</${el.name}>`;
    return buildOpen(el, true);
  }
  const open = el.attrsDirty || !el.openRaw || el.selfClosing ? buildOpen(el, false) : el.openRaw;
  let inner = "";
  for (const c of el.children) inner += serializeNode(c, source);
  return open + inner + `</${el.name}>`;
}

export function serializeDocument(doc: XDocument): string {
  if (!doc.dirty) return doc.source;
  let out = "";
  for (const c of doc.children) out += serializeNode(c, doc.source);
  return out;
}

/** 深拷贝一个节点（以序列化再解析的方式，保证拷贝体与原节点字节内容一致） */
export function cloneNode(n: XNode, source: string): XNode {
  return parseFragment(serializeNode(n, source))[0];
}

// ---------------- 树操作 ----------------

function siblingsOf(n: XNode): XNode[] {
  if (!n.parent) throw new Error("节点没有父节点");
  return n.parent.children;
}

export function removeNode(n: XNode): void {
  const sibs = siblingsOf(n);
  const i = sibs.indexOf(n);
  if (i === -1) throw new Error("节点不在父节点中");
  sibs.splice(i, 1);
  markDirty(n.parent!);
  n.parent = null;
}

export function insertBefore(ref: XNode, nodes: XNode[]): void {
  const parent = ref.parent!;
  const sibs = parent.children;
  const i = sibs.indexOf(ref);
  sibs.splice(i, 0, ...nodes);
  nodes.forEach((x) => (x.parent = parent));
  markDirty(parent);
}

export function insertAfter(ref: XNode, nodes: XNode[]): void {
  const parent = ref.parent!;
  const sibs = parent.children;
  const i = sibs.indexOf(ref);
  sibs.splice(i + 1, 0, ...nodes);
  nodes.forEach((x) => (x.parent = parent));
  markDirty(parent);
}

export function appendChild(parent: XElement, nodes: XNode[]): void {
  parent.children.push(...nodes);
  nodes.forEach((x) => (x.parent = parent));
  if (parent.selfClosing) parent.attrsDirty = true; // 自闭合元素要改写成成对标签
  markDirty(parent);
}

export function replaceNode(oldNode: XNode, nodes: XNode[]): void {
  insertBefore(oldNode, nodes);
  removeNode(oldNode);
}

/** 读取元素内所有文本（解码后） */
export function innerText(el: XElement): string {
  let s = "";
  for (const c of el.children) {
    if (c.type === "text") s += c.text;
    else if (c.type === "el") s += innerText(c);
  }
  return s;
}
