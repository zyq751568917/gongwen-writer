#!/usr/bin/env node
/**
 * ooxml-format.mjs — 无 Word 环境兜底引擎（Engine B）
 *
 * 用 spec.json 生成符合 GB/T 9704-2012 版式的 docx：
 *   页面 A4、边距 37/35/28/26mm、文档网格每页 22 行每行 28 字、
 *   中文槽方正 GBK 四款、西文槽 Times New Roman、字号/行距/缩进/对齐、
 *   红色分隔线、版记分隔线、奇偶页页码（宋体 4 号一字线）。
 *
 * 只用 Node 内置模块（fs + zlib），无第三方依赖、无网络访问。
 * 网格参数按 Word 自身写出的基准值标定（docGrid linesAndChars）；
 * 若正文并非 3 号字或版心被改动，网格需在 Word 中复核。
 *
 * 用法：
 *   node ooxml-format.mjs --spec spec.json --out 输出.docx [--report 报告.json]
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { inflateRawSync, deflateRawSync } from 'node:zlib';

/* ---------------- 最小 ZIP 读写 ---------------- */

const CRC_TABLE = (() => {
  const t = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c;
  }
  return t;
})();

function crc32(buf) {
  let c = -1;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

function readZip(buf) {
  let eocd = -1;
  const min = Math.max(0, buf.length - 66000);
  for (let i = buf.length - 22; i >= min; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('不是有效的 zip（docx）文件');
  const count = buf.readUInt16LE(eocd + 10);
  let off = buf.readUInt32LE(eocd + 16);
  const entries = new Map();
  for (let n = 0; n < count; n++) {
    if (buf.readUInt32LE(off) !== 0x02014b50) break;
    const method = buf.readUInt16LE(off + 10);
    const compSize = buf.readUInt32LE(off + 20);
    const nameLen = buf.readUInt16LE(off + 28);
    const extraLen = buf.readUInt16LE(off + 30);
    const commentLen = buf.readUInt16LE(off + 32);
    const localOff = buf.readUInt32LE(off + 42);
    const name = buf.toString('utf8', off + 46, off + 46 + nameLen);
    const lNameLen = buf.readUInt16LE(localOff + 26);
    const lExtraLen = buf.readUInt16LE(localOff + 28);
    const dataStart = localOff + 30 + lNameLen + lExtraLen;
    const raw = buf.subarray(dataStart, dataStart + compSize);
    entries.set(name, method === 0 ? Buffer.from(raw) : inflateRawSync(raw));
    off += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

function writeZip(entries) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const [name, data] of entries) {
    const nameBuf = Buffer.from(name, 'utf8');
    const deflated = deflateRawSync(data, { level: 9 });
    const useDeflate = deflated.length < data.length;
    const body = useDeflate ? deflated : data;
    const method = useDeflate ? 8 : 0;
    const crc = crc32(data);
    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0);
    lh.writeUInt16LE(20, 4);
    lh.writeUInt16LE(0x0800, 6);
    lh.writeUInt16LE(method, 8);
    lh.writeUInt16LE(0, 10);
    lh.writeUInt16LE(0x21, 12);
    lh.writeUInt32LE(crc, 14);
    lh.writeUInt32LE(body.length, 18);
    lh.writeUInt32LE(data.length, 22);
    lh.writeUInt16LE(nameBuf.length, 26);
    locals.push(lh, nameBuf, body);
    const ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50, 0);
    ch.writeUInt16LE(20, 4);
    ch.writeUInt16LE(20, 6);
    ch.writeUInt16LE(0x0800, 8);
    ch.writeUInt16LE(method, 10);
    ch.writeUInt16LE(0, 12);
    ch.writeUInt16LE(0x21, 14);
    ch.writeUInt32LE(crc, 16);
    ch.writeUInt32LE(body.length, 20);
    ch.writeUInt32LE(data.length, 24);
    ch.writeUInt16LE(nameBuf.length, 28);
    ch.writeUInt32LE(0, 38);
    ch.writeUInt32LE(offset, 42);
    centrals.push(ch, nameBuf);
    offset += 30 + nameBuf.length + body.length;
  }
  const centralBuf = Buffer.concat(centrals);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(entries.size, 8);
  eocd.writeUInt16LE(entries.size, 10);
  eocd.writeUInt32LE(centralBuf.length, 12);
  eocd.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, centralBuf, eocd]);
}

/* ---------------- 角色默认值（与 Engine A 保持一致） ---------------- */

const FONT_TITLE = '方正小标宋_GBK';
const FONT_H1 = '方正黑体_GBK';
const FONT_H2 = '方正楷体_GBK';
const FONT_BODY = '方正仿宋_GBK';
const FONT_WEST = 'Times New Roman';
const FONT_PAGE = '宋体';

const DEFAULTS = {
  redHeader: { font: FONT_TITLE, size: 22, align: 'center', first: 0, left: 0, right: 0, line: 29, before: 0, after: 0 },
  redLine: { font: FONT_BODY, size: 6, align: 'left', first: 0, left: 0, right: 0, line: 1, before: 0, after: 0 },
  docNumber: { font: FONT_BODY, size: 16, align: 'center', first: 0, left: 0, right: 0, line: 29, before: 2, after: 0 },
  signer: { font: FONT_BODY, size: 16, align: 'right', first: 0, left: 0, right: 1, line: 29, before: 0, after: 0 },
  title: { font: FONT_TITLE, size: 22, align: 'center', first: 0, left: 0, right: 0, line: 29, before: 2, after: 0 },
  mainRecipient: { font: FONT_BODY, size: 16, align: 'left', first: 0, left: 0, right: 0, line: 29, before: 1, after: 0 },
  body: { font: FONT_BODY, size: 16, align: 'both', first: 2, left: 0, right: 0, line: 29, before: 0, after: 0 },
  h1: { font: FONT_H1, size: 16, align: 'left', first: 2, left: 0, right: 0, line: 29, before: 0, after: 0 },
  h2: { font: FONT_H2, size: 16, align: 'left', first: 2, left: 0, right: 0, line: 29, before: 0, after: 0 },
  h3: { font: FONT_BODY, size: 16, align: 'left', first: 2, left: 0, right: 0, line: 29, before: 0, after: 0 },
  h4: { font: FONT_BODY, size: 16, align: 'left', first: 2, left: 0, right: 0, line: 29, before: 0, after: 0 },
  attachmentNote: { font: FONT_BODY, size: 16, align: 'left', first: 0, left: 2, right: 0, line: 29, before: 1, after: 0 },
  issuer: { font: FONT_BODY, size: 16, align: 'right', first: 0, left: 0, right: 0, line: 29, before: 3, after: 0 },
  date: { font: FONT_BODY, size: 16, align: 'right', first: 0, left: 0, right: 4, line: 29, before: 0, after: 0 },
  note: { font: FONT_BODY, size: 16, align: 'left', first: 0, left: 2, right: 0, line: 29, before: 1, after: 0 },
  cc: { font: FONT_BODY, size: 14, align: 'left', first: 0, left: 1, right: 1, line: 21, before: 1, after: 0 },
  printer: { font: FONT_BODY, size: 14, align: 'left', first: 0, left: 1, right: 1, line: 21, before: 0, after: 0 },
  attachmentTitle: { font: FONT_TITLE, size: 22, align: 'center', first: 0, left: 0, right: 0, line: 29, before: 0, after: 0 },
  attachmentMarker: { font: FONT_H1, size: 16, align: 'left', first: 0, left: 0, right: 0, line: 29, before: 0, after: 0 },
};

const ALIGN = { left: 'left', center: 'center', right: 'right', both: 'both', justify: 'both' };

/* ---------------- 单位换算与常量 ---------------- */

const mm2twip = (mm) => Math.round((mm / 25.4) * 1440);
const A4_W = 11907;
const A4_H = 16839;
const REF_CHAR_SPACE = -834; // Word 对 3号字/每行28字 写出的字符网格值

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/* ---------------- XML 构造 ---------------- */

function runXml(text, { font, size, color }) {
  const rPr =
    '<w:rPr>' +
    '<w:rFonts w:ascii="' + FONT_WEST + '" w:hAnsi="' + FONT_WEST + '" w:eastAsia="' + font + '" w:cs="' + FONT_WEST + '"/>' +
    (color === 'red' ? '<w:color w:val="FF0000"/>' : '') +
    '<w:sz w:val="' + Math.round(size * 2) + '"/><w:szCs w:val="' + Math.round(size * 2) + '"/>' +
    '</w:rPr>';
  const parts = String(text).split('\n');
  let body = '';
  parts.forEach((seg, i) => {
    if (i > 0) body += '<w:r>' + rPr + '<w:br/></w:r>';
    if (seg.length > 0) body += '<w:r>' + rPr + '<w:t xml:space="preserve">' + esc(seg) + '</w:t></w:r>';
  });
  return body || '<w:r>' + rPr + '<w:t xml:space="preserve"></w:t></w:r>';
}

function pPrXml(o) {
  let x = '<w:pPr>';
  if (o.pageBreakBefore) x += '<w:pageBreakBefore/>';
  if (o.borderTop || o.borderBottom) {
    x += '<w:pBdr>';
    if (o.borderTop) x += '<w:top w:val="single" w:sz="' + o.borderTop.sz + '" w:space="1" w:color="' + o.borderTop.color + '"/>';
    if (o.borderBottom) x += '<w:bottom w:val="single" w:sz="' + o.borderBottom.sz + '" w:space="1" w:color="' + o.borderBottom.color + '"/>';
    x += '</w:pBdr>';
  }
  x += '<w:snapToGrid/>';
  x += '<w:spacing w:before="0" w:beforeLines="' + Math.round(o.before * 100) + '" w:after="0" w:afterLines="' + Math.round(o.after * 100) +
       '" w:line="' + Math.round(o.line * 20) + '" w:lineRule="exact"/>';
  x += '<w:ind w:firstLineChars="' + Math.round(o.first * 100) + '" w:leftChars="' + Math.round(o.left * 100) +
       '" w:rightChars="' + Math.round(o.right * 100) + '"/>';
  x += '<w:jc w:val="' + (ALIGN[o.align] || 'left') + '"/>';
  x += '</w:pPr>';
  return x;
}

function footerXml(align, indentSide) {
  const rPr = '<w:rPr><w:rFonts w:ascii="' + FONT_PAGE + '" w:hAnsi="' + FONT_PAGE + '" w:eastAsia="' + FONT_PAGE + '"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>';
  const ind = indentSide === 'right' ? '<w:ind w:rightChars="100"/>' : '<w:ind w:leftChars="100"/>';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:p><w:pPr>' + ind + '<w:jc w:val="' + align + '"/></w:pPr>' +
    '<w:r>' + rPr + '<w:t xml:space="preserve">— </w:t></w:r>' +
    '<w:r>' + rPr + '<w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r>' + rPr + '<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
    '<w:r>' + rPr + '<w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r>' + rPr + '<w:t>1</w:t></w:r>' +
    '<w:r>' + rPr + '<w:fldChar w:fldCharType="end"/></w:r>' +
    '<w:r>' + rPr + '<w:t xml:space="preserve"> —</w:t></w:r>' +
    '</w:p></w:ftr>';
}

/* ---------------- 主流程 ---------------- */

function parseArgs(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--spec') out.spec = argv[++i];
    else if (a === '--out') out.out = argv[++i];
    else if (a === '--report') out.report = argv[++i];
  }
  return out;
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.spec || !args.out) {
    console.error('用法: node ooxml-format.mjs --spec spec.json --out 输出.docx [--report 报告.json]');
    process.exit(2);
  }

  const spec = JSON.parse(readFileSync(args.spec, 'utf8').replace(/^\uFEFF/, ''));
  const blocks = Array.isArray(spec.blocks) ? spec.blocks : [];
  if (blocks.length === 0) {
    console.error('spec.blocks 为空');
    process.exit(2);
  }
  const unknown = blocks.filter((b) => !Object.prototype.hasOwnProperty.call(DEFAULTS, b.role)).map((b) => b.role);
  if (unknown.length) {
    console.error('未知角色: ' + unknown.join(', '));
    process.exit(2);
  }
  for (let i = 0; i < blocks.length; i++) {
    if (blocks[i].text === undefined || blocks[i].text === null) {
      console.error('block #' + i + ' 缺少 text');
      process.exit(2);
    }
  }

  const page = spec.page || {};
  const mTop = page.marginTopMm ?? 37;
  const mBottom = page.marginBottomMm ?? 35;
  const mLeft = page.marginLeftMm ?? 28;
  const mRight = page.marginRightMm ?? 26;
  const gridLines = page.gridLines ?? 22;
  const gridChars = page.gridChars ?? 28;

  const titleIdx = blocks.findIndex((b) => b.role === 'title');
  const hasRedLine = blocks.some((b) => b.role === 'redLine');
  const bjIdx = blocks.map((b, i) => (b.role === 'cc' || b.role === 'printer' ? i : -1)).filter((i) => i >= 0);

  const bodyXml = blocks
    .map((b, i) => {
      const d = DEFAULTS[b.role];
      const o = {
        font: b.fontFarEast || d.font,
        size: b.sizePt ?? d.size,
        align: b.align || d.align,
        first: b.indentFirstChars ?? d.first,
        left: b.indentLeftChars ?? d.left,
        right: b.indentRightChars ?? d.right,
        line: b.lineSpacingPt ?? d.line,
        before: b.spaceBeforeLines ?? d.before,
        after: b.spaceAfterLines ?? d.after,
        pageBreakBefore: !!b.pageBreakBefore,
        borderTop: null,
        borderBottom: null,
      };
      const color = b.color === 'red' ? 'red' : null;
      if (!hasRedLine && titleIdx > 0 && i === titleIdx - 1 &&
          ['redHeader', 'docNumber', 'signer'].includes(blocks[i].role)) {
        o.borderBottom = { sz: 8, color: 'FF0000' };
      }
      if (b.role === 'redLine') o.borderBottom = { sz: 8, color: 'FF0000' };
      const bjPos = bjIdx.indexOf(i);
      if (bjPos >= 0) {
        if (bjPos === 0) o.borderTop = { sz: 8, color: '000000' };
        else o.borderBottom = { sz: 6, color: '000000' };
        if (bjPos === bjIdx.length - 1) o.borderBottom = { sz: 8, color: '000000' };
      }
      return '<w:p>' + pPrXml(o) + runXml(b.text, o) + '</w:p>';
    })
    .join('');

  const top = mm2twip(mTop);
  const bottom = mm2twip(mBottom);
  const left = mm2twip(mLeft);
  const right = mm2twip(mRight);
  const contentH = A4_H - top - bottom;
  const linePitch = Math.max(1, Math.floor(contentH / gridLines));
  const charSpace = gridChars === 28 ? REF_CHAR_SPACE : Math.round((REF_CHAR_SPACE * 28) / gridChars);

  const sectPr =
    '<w:sectPr>' +
    '<w:footerReference w:type="default" r:id="rId3"/>' +
    '<w:footerReference w:type="even" r:id="rId4"/>' +
    '<w:pgSz w:w="' + A4_W + '" w:h="' + A4_H + '"/>' +
    '<w:pgMar w:top="' + top + '" w:right="' + right + '" w:bottom="' + bottom + '" w:left="' + left +
    '" w:header="1134" w:footer="1587" w:gutter="0"/>' +
    '<w:cols w:space="425"/>' +
    '<w:docGrid w:type="linesAndChars" w:linePitch="' + linePitch + '" w:charSpace="' + charSpace + '"/>' +
    '</w:sectPr>';

  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<w:body>' + bodyXml + sectPr + '</w:body></w:document>';

  const stylesXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:docDefaults><w:rPrDefault><w:rPr>' +
    '<w:rFonts w:ascii="' + FONT_WEST + '" w:hAnsi="' + FONT_WEST + '" w:eastAsia="' + FONT_BODY + '" w:cs="' + FONT_WEST + '"/>' +
    '<w:sz w:val="32"/><w:szCs w:val="32"/>' +
    '</w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:line="580" w:lineRule="exact"/></w:pPr></w:pPrDefault>' +
    '</w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>' +
    '<w:rPr><w:rFonts w:ascii="' + FONT_WEST + '" w:hAnsi="' + FONT_WEST + '" w:eastAsia="' + FONT_BODY + '"/>' +
    '<w:sz w:val="32"/></w:rPr></w:style>' +
    '</w:styles>';

  const settingsXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:defaultTabStop w:val="420"/><w:evenAndOddHeaders/><w:compat/></w:settings>';

  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
    '<Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
    '</Types>';

  const rels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';

  const docRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>' +
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' +
    '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>' +
    '</Relationships>';

  const entries = new Map();
  entries.set('[Content_Types].xml', Buffer.from(contentTypes, 'utf8'));
  entries.set('_rels/.rels', Buffer.from(rels, 'utf8'));
  entries.set('word/document.xml', Buffer.from(documentXml, 'utf8'));
  entries.set('word/styles.xml', Buffer.from(stylesXml, 'utf8'));
  entries.set('word/settings.xml', Buffer.from(settingsXml, 'utf8'));
  entries.set('word/footer1.xml', Buffer.from(footerXml('right', 'right'), 'utf8'));
  entries.set('word/footer2.xml', Buffer.from(footerXml('left', 'left'), 'utf8'));
  entries.set('word/_rels/document.xml.rels', Buffer.from(docRels, 'utf8'));

  writeFileSync(args.out, writeZip(entries));

  const report = {
    engine: 'B(ooxml-node)',
    output: args.out,
    page: {
      size: 'A4',
      marginTopMm: mTop, marginBottomMm: mBottom, marginLeftMm: mLeft, marginRightMm: mRight,
      gridLines, gridChars,
      docGrid: { type: 'linesAndChars', linePitch, charSpace },
    },
    fonts: { title: FONT_TITLE, h1: FONT_H1, h2: FONT_H2, body: FONT_BODY, western: FONT_WEST, pageNumber: FONT_PAGE },
    blocks: blocks.map((b, i) => ({ index: i, role: b.role })),
    needsWord: [
      '字符网格按 3 号字标定；若正文改用其他字号，需在 Word 中用 verify-gongwen.ps1 复核',
      '印章、附注分隔等手工要素需按单位要求补',
    ],
  };
  if (args.report) writeFileSync(args.report, JSON.stringify(report, null, 2), 'utf8');
  console.log(JSON.stringify(report, null, 2));
}

main();
