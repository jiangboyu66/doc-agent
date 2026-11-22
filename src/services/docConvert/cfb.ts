/**
 * OLE2 复合文档（Compound File Binary，[MS-CFB]）读取：.doc 的外层容器。
 *
 * 一个 .doc 文件是一个小型"文件系统"：按扇区（512 / 4096 字节）存放若干流（WordDocument、0Table / 1Table、Data……），
 * 扇区之间用 FAT 链接；小于 4096 字节的流放在"迷你流"里，用迷你 FAT 链接。这里只读不写。
 * 损坏或构造异常的文件（链表成环、越界）一律抛出 DocFileError，不会死循环。
 */

export class DocFileError extends Error {
  constructor(message: string, readonly code: "not-ole" | "corrupt" | "encrypted" | "old-version" | "unsupported" = "corrupt") {
    super(message);
  }
}

const SIG = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;

interface DirEntry {
  name: string;
  type: number; // 1 存储，2 流，5 根
  left: number;
  right: number;
  child: number;
  start: number;
  size: number;
}

export class CompoundFile {
  private readonly sectorSize: number;
  private readonly miniSectorSize: number;
  private readonly miniCutoff: number;
  private readonly fat: Uint32Array;
  private readonly miniFat: Uint32Array;
  private readonly dirs: DirEntry[];
  private readonly miniStream: Buffer;
  /** 根存储下的流：名称 → 目录项 */
  private readonly streams = new Map<string, DirEntry>();

  constructor(private readonly buf: Buffer) {
    if (buf.length < 512 || !buf.subarray(0, 8).equals(SIG)) throw new DocFileError("不是 Word 97-2003 文档（缺少 OLE2 复合文档签名）。", "not-ole");
    const shift = buf.readUInt16LE(0x1e);
    if (shift !== 9 && shift !== 12) throw new DocFileError("复合文档头已损坏（扇区大小异常）。");
    this.sectorSize = 1 << shift;
    this.miniSectorSize = 1 << buf.readUInt16LE(0x20);
    this.miniCutoff = buf.readUInt32LE(0x38) || 4096;
    const nFat = buf.readUInt32LE(0x2c);
    const firstDir = buf.readUInt32LE(0x30);
    const firstMiniFat = buf.readUInt32LE(0x3c);
    const nMiniFat = buf.readUInt32LE(0x40);
    let difatSect = buf.readUInt32LE(0x44);
    const nDifat = buf.readUInt32LE(0x48);
    const maxSectors = Math.ceil(buf.length / this.sectorSize) + 1;
    if (nFat > maxSectors) throw new DocFileError("复合文档头已损坏（FAT 扇区数异常）。");

    // DIFAT：头里的 109 项 + DIFAT 扇区链
    const fatSectors: number[] = [];
    for (let i = 0; i < 109 && fatSectors.length < nFat; i++) {
      const s = buf.readUInt32LE(0x4c + i * 4);
      if (s !== FREESECT) fatSectors.push(s);
    }
    const perSect = this.sectorSize / 4;
    for (let k = 0; k < nDifat && difatSect !== ENDOFCHAIN && difatSect !== FREESECT && fatSectors.length < nFat; k++) {
      const off = this.sectorOffset(difatSect);
      for (let i = 0; i < perSect - 1 && fatSectors.length < nFat; i++) {
        const s = buf.readUInt32LE(off + i * 4);
        if (s !== FREESECT) fatSectors.push(s);
      }
      difatSect = buf.readUInt32LE(off + (perSect - 1) * 4);
    }
    const fat = new Uint32Array(fatSectors.length * perSect);
    fatSectors.forEach((s, i) => {
      const off = this.sectorOffset(s);
      for (let j = 0; j < perSect; j++) fat[i * perSect + j] = off + j * 4 + 4 <= buf.length ? buf.readUInt32LE(off + j * 4) : FREESECT;
    });
    this.fat = fat;

    // 目录
    const dirBuf = this.readChain(firstDir, this.fat, this.sectorSize, (s) => this.sector(s));
    this.dirs = [];
    for (let off = 0; off + 128 <= dirBuf.length; off += 128) {
      const nameLen = dirBuf.readUInt16LE(off + 64);
      const name = dirBuf.toString("utf16le", off, off + Math.max(0, Math.min(64, nameLen) - 2));
      this.dirs.push({
        name, type: dirBuf[off + 66],
        left: dirBuf.readUInt32LE(off + 68), right: dirBuf.readUInt32LE(off + 72), child: dirBuf.readUInt32LE(off + 76),
        start: dirBuf.readUInt32LE(off + 116), size: dirBuf.readUInt32LE(off + 120),
      });
    }
    const root = this.dirs[0];
    if (!root || root.type !== 5) throw new DocFileError("复合文档的根目录已损坏。");

    // 迷你 FAT 与迷你流
    const miniFatBuf = nMiniFat && firstMiniFat !== ENDOFCHAIN ? this.readChain(firstMiniFat, this.fat, this.sectorSize, (s) => this.sector(s)) : Buffer.alloc(0);
    this.miniFat = new Uint32Array(miniFatBuf.length / 4);
    for (let i = 0; i < this.miniFat.length; i++) this.miniFat[i] = miniFatBuf.readUInt32LE(i * 4);
    this.miniStream = root.start !== ENDOFCHAIN && root.size ? this.readChain(root.start, this.fat, this.sectorSize, (s) => this.sector(s), root.size) : Buffer.alloc(0);

    // 根存储下的流（红黑树，中序遍历；防环）
    const seen = new Set<number>();
    const walk = (i: number) => {
      if (i === FREESECT || i >= this.dirs.length || seen.has(i)) return;
      seen.add(i);
      const d = this.dirs[i];
      walk(d.left);
      if (d.type === 2) this.streams.set(d.name.toUpperCase(), d);
      walk(d.right);
    };
    walk(root.child);
  }

  private sectorOffset(s: number): number {
    const off = (s + 1) * this.sectorSize;
    if (off >= this.buf.length) throw new DocFileError("复合文档已损坏（扇区越界）。");
    return off;
  }

  private sector(s: number): Buffer {
    const off = this.sectorOffset(s);
    return this.buf.subarray(off, Math.min(off + this.sectorSize, this.buf.length));
  }

  private readChain(start: number, fat: Uint32Array, size: number, get: (s: number) => Buffer, limit?: number): Buffer {
    const parts: Buffer[] = [];
    let s = start, total = 0;
    const seen = new Set<number>();
    while (s !== ENDOFCHAIN && s !== FREESECT) {
      if (seen.has(s) || s >= fat.length) throw new DocFileError("复合文档已损坏（扇区链异常）。");
      seen.add(s);
      parts.push(get(s));
      total += size;
      if (limit !== undefined && total >= limit) break;
      s = fat[s] ?? ENDOFCHAIN;
    }
    const out = Buffer.concat(parts);
    return limit !== undefined ? out.subarray(0, Math.min(limit, out.length)) : out;
  }

  /** 流名不区分大小写（[MS-CFB] 2.6.4） */
  has(name: string): boolean {
    return this.streams.has(name.toUpperCase());
  }

  /** 读取根存储下的一个流 */
  stream(name: string): Buffer {
    const d = this.streams.get(name.toUpperCase());
    if (!d) throw new DocFileError(`复合文档里缺少 ${name} 流。`);
    if (d.size === 0) return Buffer.alloc(0);
    if (d.size < this.miniCutoff) {
      return this.readChain(d.start, this.miniFat, this.miniSectorSize, (s) => {
        const off = s * this.miniSectorSize;
        if (off >= this.miniStream.length) throw new DocFileError("复合文档已损坏（迷你扇区越界）。");
        return this.miniStream.subarray(off, off + this.miniSectorSize);
      }, d.size);
    }
    return this.readChain(d.start, this.fat, this.sectorSize, (s) => this.sector(s), d.size);
  }
}
