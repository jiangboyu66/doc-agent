/**
 * SPRM（Single Property Modifier，[MS-DOC] 2.2.5）：.doc 里所有格式都是一串"属性修改"指令。
 * 每条指令 2 字节操作码 + 操作数；操作数长度由操作码的 spra 位决定（变长的以长度字节开头）。
 */

export interface Sprm {
  op: number;
  arg: Buffer;
}

export function parseGrpprl(buf: Buffer, start = 0, end = buf.length): Sprm[] {
  const out: Sprm[] = [];
  let i = start;
  end = Math.min(end, buf.length);
  while (i + 2 <= end) {
    const op = buf.readUInt16LE(i);
    i += 2;
    let len: number, skip = 0;
    switch (op >> 13) {
      case 0: case 1: len = 1; break;
      case 2: case 4: case 5: len = 2; break;
      case 3: len = 4; break;
      case 7: len = 3; break;
      default: {
        if (op === 0xd608 || op === 0xd606) {
          // sprmTDefTable：2 字节长度（比余下部分多 1）
          if (i + 2 > end) return out;
          len = buf.readUInt16LE(i) - 1;
          skip = 2;
        } else if (op === 0xc615) {
          // sprmPChgTabs：长度字节为 255 时按内容计算
          if (i + 1 > end) return out;
          const cb = buf[i];
          if (cb === 255 && i + 2 <= end) {
            const del = buf[i + 1];
            const addAt = i + 2 + del * 4;
            const add = addAt < end ? buf[addAt] : 0;
            len = 1 + del * 4 + 1 + add * 3;
          } else len = cb;
          skip = 1;
        } else {
          if (i + 1 > end) return out;
          len = buf[i];
          skip = 1;
        }
      }
    }
    if (len < 0 || i + skip + len > end) return out;
    out.push({ op, arg: buf.subarray(i + skip, i + skip + len) });
    i += skip + len;
  }
  return out;
}

/** 操作码常量（只列出用到的） */
export const S = {
  // 字符
  CFRMarkDel: 0x0800, CFRMarkIns: 0x0801, CFData: 0x0806, CFOle2: 0x080a, CIbstRMark: 0x4804, CDttmRMark: 0x6805,
  CIbstRMarkDel: 0x4863, CDttmRMarkDel: 0x6864, CPicLocation: 0x6a03, CFSpec: 0x0855, CFObj: 0x0856, CSymbol: 0x6a09,
  CIstd: 0x4a30, CFBold: 0x0835, CFItalic: 0x0836, CFStrike: 0x0837, CFOutline: 0x0838, CFShadow: 0x0839,
  CFSmallCaps: 0x083a, CFCaps: 0x083b, CFVanish: 0x083c, CKul: 0x2a3e, CDxaSpace: 0x8840, CIco: 0x2a42, CHps: 0x4a43,
  CHpsPos: 0x4845, CIss: 0x2a48, CHighlight: 0x2a0c, CRgFtc0: 0x4a4f, CRgFtc1: 0x4a50, CRgFtc2: 0x4a51, CFtcBi: 0x4a5e,
  CCharScale: 0x4852, CFDStrike: 0x2a53, CFImprint: 0x0854, CFEmboss: 0x0858, CFBoldBi: 0x085c, CFItalicBi: 0x085d,
  CIdctHint: 0x286f, CHpsKern: 0x484b, CFNoProof: 0x0875,
  CHpsBi: 0x4a61, CCv: 0x6870, CShd80: 0x4866, CShd: 0xca71, CKcd: 0x2a34, CBrc80: 0x6865, CBrc: 0xca72, CRgLid0_80: 0x486d, CRgLid1_80: 0x486e, CRgLid0: 0x4873, CRgLid1: 0x4874, CLidBi: 0x485f,
  // 段落
  PIstd: 0x4600, PJc80: 0x2403, PJc: 0x2461, PFKeep: 0x2405, PFKeepFollow: 0x2406, PFPageBreakBefore: 0x2407,
  PIlvl: 0x260a, PIlfo: 0x460b, PChgTabsPapx: 0xc60d, PChgTabs: 0xc615, PDxaRight80: 0x840e, PDxaLeft80: 0x840f, PDxaLeft180: 0x8411,
  PDxaRight: 0x845d, PDxaLeft: 0x845e, PDxaLeft1: 0x8460, PDxcRight: 0x4455, PDxcLeft: 0x4456, PDxcLeft1: 0x4457,
  PDyaLine: 0x6412, PDyaBefore: 0xa413, PDyaAfter: 0xa414, PDylBefore: 0x4458, PDylAfter: 0x4459,
  PFDyaBeforeAuto: 0x245b, PFDyaAfterAuto: 0x245c, PFInTable: 0x2416, PFTtp: 0x2417, PFWidowControl: 0x2431,
  PFKinsoku: 0x2433, PFAutoSpaceDE: 0x2437, PFAutoSpaceDN: 0x2438, POutLvl: 0x2640, PFBiDi: 0x2441, PFContextualSpacing: 0x246d,
  PBrcTop80: 0x6424, PBrcLeft80: 0x6425, PBrcBottom80: 0x6426, PBrcRight80: 0x6427, PBrcBetween80: 0x6428,
  PBrcTop: 0xc64e, PBrcLeft: 0xc64f, PBrcBottom: 0xc650, PBrcRight: 0xc651, PBrcBetween: 0xc652,
  PShd80: 0x442d, PShd: 0xc64d, PItap: 0x6649, PFInnerTableCell: 0x244b, PFInnerTtp: 0x244c, PHugePapx: 0x6646,
  PFNoAutoHyph: 0x242a,
  // 图文框（定位段落）与东亚排版
  PDxaAbs: 0x8418, PDyaAbs: 0x8419, PDxaWidth: 0x841a, PPc: 0x261b, PWr: 0x2423, PWHeightAbs: 0x442b, PDcs: 0x442c,
  PDyaFromText: 0x842e, PDxaFromText: 0x842f, PFLocked: 0x2430, PFNoLineNumb: 0x240c,
  PFWordWrap: 0x2434, PFOverflowPunct: 0x2435, PFTopLinePunct: 0x2436, PWAlignFont: 0x4439, PFUsePgsuSettings: 0x2447, PFAdjustRight: 0x2448,
  // 表格
  TDefTable: 0xd608, TDxaGapHalf: 0x9602, TDyaRowHeight: 0x9407, TTableHeader: 0x3404, TFCantSplit: 0x3403, TFCantSplit90: 0x3466,
  TJc90: 0x5400, TJc: 0x548a, TDxaLeft: 0x9601, TTableBorders80: 0xd605, TTableBorders: 0xd613,
  TDefTableShd80: 0xd609, TDefTableShd: 0xd612, TDefTableShd2nd: 0xd616, TDefTableShd3rd: 0xd60c,
  TSetBrc80: 0xd620, TSetBrc: 0xd62f, TVertMerge: 0xd62b, TVertAlign: 0xd62c, TTableWidth: 0xf614, TCellPadding: 0xd632,
  TCellPaddingDefault: 0xd634, TWidthBefore: 0xf617, TFAutofit: 0x3615, TIstd: 0x563a,
  // 节
  SBkc: 0x3009, SFTitlePage: 0x300a, SCcolumns: 0x500b, SDxaColumns: 0x900c, SNfcPgn: 0x300e, SFPgnRestart: 0x3011,
  SPgnStart97: 0x501c, SPgnStart: 0x7044, SBOrientation: 0x301d, SXaPage: 0xb01f, SYaPage: 0xb020, SDxaLeft: 0xb021, SDxaRight: 0xb022,
  SDyaTop: 0x9023, SDyaBottom: 0x9024, SDzaGutter: 0xb025, SDyaHdrTop: 0xb017, SDyaHdrBottom: 0xb018, SLBetween: 0x3019,
  SVjc: 0x301a, SDyaLinePitch: 0x9031, SClm: 0x5032, SDxtCharSpace: 0x7030, SFEvenlySpaced: 0x3005, SDxaColWidth: 0xf203, SDxaColSpacing: 0xf204,
} as const;

/** 在一组 sprm 里找最后一条指定操作码的（后面的覆盖前面的） */
export function last(sprms: Sprm[], ...ops: number[]): Sprm | undefined {
  for (let i = sprms.length - 1; i >= 0; i--) if (ops.includes(sprms[i].op)) return sprms[i];
  return undefined;
}
