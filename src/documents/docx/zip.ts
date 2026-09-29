/**
 * 保真 ZIP 读写
 *
 * docx 是一个 ZIP 包。常见 ZIP 库在写回时会把所有条目重新压缩一遍——解压后的内容虽然相同，
 * 但压缩数据、条目时间戳、额外字段都会变。这里自己实现：
 *   - 未修改的条目：本地文件头 + 压缩数据 + 数据描述符 原封不动地按字节拷贝；
 *   - 修改过的条目：只重新压缩这一个条目；
 *   - 中央目录保留原有顺序与全部字段，只修正偏移量和被修改条目的 CRC/大小。
 * 这样 word/media、styles.xml、theme、字体等所有没被碰到的部件都与原文件逐字节一致。
 */

import zlib from "node:zlib";

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

export function crc32(buf: Buffer): number {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

export interface ZipEntry {
  name: string;
  /** 中央目录头原始字节（含文件名/额外字段/注释） */
  centralRaw: Buffer;
  method: number;
  flags: number;
  crc: number;
  compSize: number;
  size: number;
  localOffset: number;
  /** 条目在原包中占据的完整字节区间（本地头 + 数据 + 可能的数据描述符） */
  extent: [number, number];
}

export class ZipPackage {
  readonly entries: ZipEntry[] = [];
  private readonly comment: Buffer;
  private readonly replaced = new Map<string, Buffer>();
  private readonly added: Array<{ name: string; data: Buffer }> = [];

  constructor(private readonly buf: Buffer) {
    const eocd = findEocd(buf);
    const total = buf.readUInt16LE(eocd + 10);
    const cdSize = buf.readUInt32LE(eocd + 12);
    const cdOffset = buf.readUInt32LE(eocd + 16);
    const commentLen = buf.readUInt16LE(eocd + 20);
    this.comment = buf.subarray(eocd + 22, eocd + 22 + commentLen);
    if (cdOffset === 0xffffffff || total === 0xffff) throw new Error("暂不支持 ZIP64 格式的文档包");

    let p = cdOffset;
    for (let i = 0; i < total; i++) {
      if (buf.readUInt32LE(p) !== SIG_CENTRAL) throw new Error("文档包中央目录损坏");
      const flags = buf.readUInt16LE(p + 8);
      const method = buf.readUInt16LE(p + 10);
      const crc = buf.readUInt32LE(p + 16);
      const compSize = buf.readUInt32LE(p + 20);
      const size = buf.readUInt32LE(p + 24);
      const nameLen = buf.readUInt16LE(p + 28);
      const extraLen = buf.readUInt16LE(p + 30);
      const cLen = buf.readUInt16LE(p + 32);
      const localOffset = buf.readUInt32LE(p + 42);
      const name = buf.subarray(p + 46, p + 46 + nameLen).toString(flags & 0x800 ? "utf8" : "latin1");
      const end = p + 46 + nameLen + extraLen + cLen;
      this.entries.push({
        name,
        centralRaw: Buffer.from(buf.subarray(p, end)),
        method,
        flags,
        crc,
        compSize,
        size,
        localOffset,
        extent: [localOffset, 0],
      });
      p = end;
    }
    void cdSize;
    // 计算每个条目的字节区间：到下一个条目（按偏移排序）或中央目录起点为止
    const sorted = [...this.entries].sort((a, b) => a.localOffset - b.localOffset);
    for (let i = 0; i < sorted.length; i++) {
      sorted[i].extent[1] = i + 1 < sorted.length ? sorted[i + 1].localOffset : cdOffset;
    }
  }

  has(name: string): boolean {
    return this.entries.some((e) => e.name === name) || this.added.some((a) => a.name === name);
  }

  names(): string[] {
    return [...this.entries.map((e) => e.name), ...this.added.map((a) => a.name)];
  }

  /** 读取原始（未替换前）条目内容 */
  readOriginal(name: string): Buffer {
    const e = this.entries.find((x) => x.name === name);
    if (!e) throw new Error(`文档包中不存在部件 ${name}`);
    const lo = e.localOffset;
    if (this.buf.readUInt32LE(lo) !== SIG_LOCAL) throw new Error(`部件 ${name} 本地文件头损坏`);
    const nLen = this.buf.readUInt16LE(lo + 26);
    const xLen = this.buf.readUInt16LE(lo + 28);
    const start = lo + 30 + nLen + xLen;
    const data = this.buf.subarray(start, start + e.compSize);
    if (e.method === 0) return Buffer.from(data);
    if (e.method === 8) return zlib.inflateRawSync(data);
    throw new Error(`部件 ${name} 使用了不支持的压缩方式 ${e.method}`);
  }

  read(name: string): Buffer {
    const r = this.replaced.get(name);
    if (r) return r;
    const a = this.added.find((x) => x.name === name);
    if (a) return a.data;
    return this.readOriginal(name);
  }

  readText(name: string): string {
    return this.read(name).toString("utf8");
  }

  /** 写入部件内容；若内容与原始字节完全相同则视为未修改（保持原压缩数据） */
  write(name: string, data: Buffer): void {
    const existing = this.entries.find((e) => e.name === name);
    if (existing) {
      if (crc32(data) === existing.crc && data.length === existing.size && data.equals(this.readOriginal(name))) {
        this.replaced.delete(name);
      } else this.replaced.set(name, data);
      return;
    }
    const a = this.added.find((x) => x.name === name);
    if (a) a.data = data;
    else this.added.push({ name, data });
  }

  modifiedParts(): string[] {
    return [...this.replaced.keys(), ...this.added.map((a) => a.name)];
  }

  toBuffer(): Buffer {
    const out: Buffer[] = [];
    const central: Buffer[] = [];
    let offset = 0;

    for (const e of this.entries) {
      const newData = this.replaced.get(e.name);
      if (!newData) {
        const chunk = this.buf.subarray(e.extent[0], e.extent[1]);
        out.push(chunk);
        const c = Buffer.from(e.centralRaw);
        c.writeUInt32LE(offset, 42);
        central.push(c);
        offset += chunk.length;
      } else {
        const { local, cen } = buildEntry(e.name, newData, e.flags, e.centralRaw, offset);
        out.push(local);
        central.push(cen);
        offset += local.length;
      }
    }
    for (const a of this.added) {
      const { local, cen } = buildEntry(a.name, a.data, 0x800, null, offset);
      out.push(local);
      central.push(cen);
      offset += local.length;
    }
    const cdBuf = Buffer.concat(central);
    const eocd = Buffer.alloc(22);
    eocd.writeUInt32LE(SIG_EOCD, 0);
    eocd.writeUInt16LE(0, 4);
    eocd.writeUInt16LE(0, 6);
    eocd.writeUInt16LE(central.length, 8);
    eocd.writeUInt16LE(central.length, 10);
    eocd.writeUInt32LE(cdBuf.length, 12);
    eocd.writeUInt32LE(offset, 16);
    eocd.writeUInt16LE(this.comment.length, 20);
    return Buffer.concat([...out, cdBuf, eocd, this.comment]);
  }
}

function findEocd(buf: Buffer): number {
  const min = Math.max(0, buf.length - 65557);
  for (let i = buf.length - 22; i >= min; i--) {
    if (buf.readUInt32LE(i) === SIG_EOCD) return i;
  }
  throw new Error("不是有效的 docx 文件（找不到 ZIP 目录结束标记）");
}

function dosDateTime(d = new Date()): { time: number; date: number } {
  return {
    time: (d.getHours() << 11) | (d.getMinutes() << 5) | Math.floor(d.getSeconds() / 2),
    date: ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate(),
  };
}

function buildEntry(name: string, data: Buffer, origFlags: number, origCentral: Buffer | null, offset: number) {
  const nameBuf = Buffer.from(name, origFlags & 0x800 ? "utf8" : "latin1");
  const comp = zlib.deflateRawSync(data);
  const crc = crc32(data);
  const { time, date } = dosDateTime();
  const flags = origFlags & ~0x0008 & ~0x0001; // 清除数据描述符/加密标志

  const local = Buffer.alloc(30);
  local.writeUInt32LE(SIG_LOCAL, 0);
  local.writeUInt16LE(20, 4);
  local.writeUInt16LE(flags, 6);
  local.writeUInt16LE(8, 8);
  local.writeUInt16LE(time, 10);
  local.writeUInt16LE(date, 12);
  local.writeUInt32LE(crc, 14);
  local.writeUInt32LE(comp.length, 18);
  local.writeUInt32LE(data.length, 22);
  local.writeUInt16LE(nameBuf.length, 26);
  local.writeUInt16LE(0, 28);

  let cen: Buffer;
  if (origCentral) {
    // 沿用原中央目录头的版本、属性、额外字段、注释，只更新 CRC/大小/偏移/时间/方法
    cen = Buffer.from(origCentral);
    cen.writeUInt16LE(flags, 8);
    cen.writeUInt16LE(8, 10);
    cen.writeUInt16LE(time, 12);
    cen.writeUInt16LE(date, 14);
    cen.writeUInt32LE(crc, 16);
    cen.writeUInt32LE(comp.length, 20);
    cen.writeUInt32LE(data.length, 24);
    cen.writeUInt32LE(offset, 42);
  } else {
    cen = Buffer.alloc(46 + nameBuf.length);
    cen.writeUInt32LE(SIG_CENTRAL, 0);
    cen.writeUInt16LE(20, 4);
    cen.writeUInt16LE(20, 6);
    cen.writeUInt16LE(flags, 8);
    cen.writeUInt16LE(8, 10);
    cen.writeUInt16LE(time, 12);
    cen.writeUInt16LE(date, 14);
    cen.writeUInt32LE(crc, 16);
    cen.writeUInt32LE(comp.length, 20);
    cen.writeUInt32LE(data.length, 24);
    cen.writeUInt16LE(nameBuf.length, 28);
    cen.writeUInt32LE(offset, 42);
    nameBuf.copy(cen, 46);
  }
  return { local: Buffer.concat([local, nameBuf, comp]), cen };
}
