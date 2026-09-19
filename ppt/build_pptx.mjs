// 用 Node 直接生成 .pptx（OOXML），不依赖任何外部库。
// 结构与命名空间取自真实 PPTX 参考文件。
//
// 关键修正（v2）—— 解决「PowerPoint 显示内容有问题」：
//  1. 文本内边距全部归零（lIns/rIns/tIns/bIns=0）。
//     原来圆角矩形（代码块）用默认 <a:bodyPr/>，左右各 0.1 英寸内边距
//     把可用宽度吃掉 0.2 英寸，换行位置与估算完全不符。
//  2. 按真实字体度量计算每块文字所需高度（含自动换行）。
//     原来是固定行高硬估（一行≈0.38"），而微软雅黑行高系数约 1.42、
//     Consolas 约 1.28，估算严重偏小 → 文字超出文本框被裁。
//  3. 开启 <a:normAutofit/> 兜底：万一字体或换行与预期不符，
//     PowerPoint 会缩小字号而不是直接裁掉文字。
//  4. 标题改为底部对齐并给足高度；各行预留余量；超界时自动压缩间距。
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';

const ROOT = process.cwd();
const IMG = path.join(ROOT, 'ppt', 'img');
const deck = JSON.parse(fs.readFileSync(path.join(ROOT, 'ppt', 'deck.json'), 'utf8'));
const OUT = process.env.PROBE_OUT
  ? path.resolve(process.env.PROBE_OUT)
  : path.join(ROOT, 'ppt', deck.meta.out);

// ---- 探针模式：生成不同复杂度的最小文件，用于定位 PowerPoint 修复提示 ----
// PROBE_MODE = blank | text | image | mixed，PROBE_SLIDES = 页数
if (process.env.PROBE_MODE) {
  const mode = process.env.PROBE_MODE;
  const n = Number(process.env.PROBE_SLIDES || 1);
  const mk = (o) => Object.assign({ layout: 'content', title: '', lines: [] }, o);
  const sets = {
    blank: [mk({ layout: 'image', image: '', title: '', lines: [] })],
    text: [mk({
      title: '纯文字页',
      lines: ['H|标题行', 'B|普通要点一行', 'B|再来一条要点', 'C|git status', 'Q|这是一句引用', 'N|这是注释']
    })],
    image: [mk({ layout: 'image', image: '02_three_areas.png', title: '带图片页', caption: '示意图' })],
    mixed: [
      mk({ title: '混合页 A', lines: ['B|第一条', 'B|第二条'] }),
      mk({ layout: 'image', image: '03_push_pull.png', title: '混合页 B', caption: '图' }),
      mk({ title: '混合页 C', lines: ['H|小标题', 'C|git add -A'] })
    ]
  };
  deck.slides = (sets[mode] || sets.text).slice(0, n);
  deck.meta = Object.assign({}, deck.meta, { out: path.basename(OUT), title: 'probe-' + mode });
  console.log(`[probe] mode=${mode} slides=${deck.slides.length} -> ${path.basename(OUT)}`);
}

// ---- 特性开关：用于逐项排查 PowerPoint 修复提示的来源 ----
// 设为 "0" 可关闭对应特性，用来做对照实验
const F = {
  theme: process.env.FEAT_THEME !== '0',                 // 是否打包 theme1.xml
  exProps: process.env.FEAT_EXPROPS !== '0',             // presProps/viewProps/tableStyles
  autofit: process.env.FEAT_AUTOFIT !== '0',             // normAutofit 兜底
  zeropad: process.env.FEAT_ZEROPAD !== '0'              // 文本内边距归零
};
if (process.env.FEAT_THEME || process.env.FEAT_EXPROPS || process.env.FEAT_AUTOFIT || process.env.FEAT_ZEROPAD) {
  console.log('[feat]', JSON.stringify(F));
}


const EMU = 914400;                            // EMU per inch
const PX = 96;                                 // CSS px per inch
const SLIDE_W = 12192000, SLIDE_H = 6858000;   // 13.333" x 7.5"

// ---------------------------------------------------------------- 颜色
const C = {
  bg: '0B1020', panel: '111A2C', line: '23304A',
  accent: '4DD4FF', title: '9FE8FF', heading: 'FFD166', bullet: 'DCE6FA',
  code: '7FE9FF', codeBg: '060A14', quote: '63F5A0', note: '93A3C4',
  muted: '6C7A99', cover1: '12325A', cover2: '3A1A2C'
};

// ---------------------------------------------------------------- 字体度量
// 宽度按 em 估算（1em = 字号）；行高按字体的真实行高系数
function charEm(ch, font) {
  const code = ch.codePointAt(0);
  const isCJK = (code >= 0x2E80 && code <= 0x9FFF) ||
    (code >= 0xF900 && code <= 0xFAFF) ||
    (code >= 0xFF00 && code <= 0xFF60) ||
    (code >= 0x3000 && code <= 0x303F);
  if (font === 'Consolas') return isCJK ? 1.0 : 0.55;
  if (isCJK) return 1.0;
  if (/[iIljt.,:;'!|]/.test(ch)) return 0.28;
  if (/[A-Z0-9]/.test(ch)) return 0.62;
  if (ch === ' ') return 0.30;
  return 0.52;
}
function textEm(s, font) {
  let em = 0;
  for (const ch of s) em += charEm(ch, font);
  return em;
}
const LINE_FACTOR = { 'Microsoft YaHei': 1.42, 'Consolas': 1.28 };

// 计算文本在给定宽度（英寸）下需要的高度（英寸）
// ⚠ 单位陷阱：本文件的 size 参数是「1/100 磅」（OOXML 的 sz 规定），
//   所以要用 size/100 换算成磅再参与几何计算。
function measureText(text, size, boxWidthIn, font, lineSpacingPct = 100) {
  const pt = size / 100;                       // 1/100 磅 → 磅
  const emPerIn = 72 / pt;                     // 每英寸可容纳的 em 数
  const boxEm = boxWidthIn * emPerIn;
  const paras = String(text).split('\n');
  let totalLines = 0, maxEm = 0;
  for (const p of paras) {
    if (p === '') { totalLines += 1; continue; }
    const pem = textEm(p, font);
    maxEm = Math.max(maxEm, pem);
    let lines = 1, cur = 0;
    for (const ch of p) {
      const w = charEm(ch, font);
      if (cur + w > boxEm && cur > 0) { lines++; cur = 0; }
      cur += w;
    }
    totalLines += lines;
  }
  const lineFactor = LINE_FACTOR[font] || 1.42;
  const lineH = pt * lineFactor * (lineSpacingPct / 100) / 72;   // 英寸
  return { lines: totalLines, height: totalLines * lineH, maxWidthEm: maxEm, boxEm };
}

// ---------------------------------------------------------------- xml utils
const esc = (s) => String(s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
const inch = (v) => Math.round(v * EMU);

let shapeIdSeq = 1;
const nextId = () => ++shapeIdSeq;

function rPr({ size = 1400, color = C.bullet, bold = 0, font = 'Microsoft YaHei', italic = 0 }) {
  const latin = font === 'Consolas' ? 'Consolas' : font;
  const ea = font === 'Consolas' ? 'Microsoft YaHei' : font;
  return `<a:rPr lang="zh-CN" altLang="en-US" sz="${size}" b="${bold}" i="${italic}" dirty="0">` +
    `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` +
    `<a:latin typeface="${latin}"/><a:ea typeface="${ea}"/><a:cs typeface="${latin}"/></a:rPr>`;
}

// 统一 bodyPr：内边距归零 + normAutofit 兜底（内边距可影响换行位置）
function bodyPr({ anchor = 't', autofit = true } = {}) {
  const pad = F.zeropad ? ' lIns="0" rIns="0" tIns="0" bIns="0"' : '';
  const fit = (autofit && F.autofit) ? `<a:normAutofit/>` : `<a:noAutofit/>`;
  return `<a:bodyPr wrap="square"${pad} anchor="${anchor}">${fit}</a:bodyPr>`;
}

function para(runs, opts = {}) {
  const { algn = 'l', bullet = false, marL = 0, indent = 0, spaceAfter = 0, lineSpacing = 100 } = opts;
  let pPr = `<a:pPr algn="${algn}" marL="${marL}" indent="${indent}">`;
  pPr += `<a:lnSpc><a:spcPct val="${lineSpacing * 1000}"/></a:lnSpc>`;
  pPr += `<a:spcBef><a:spcPts val="0"/></a:spcBef>`;
  pPr += `<a:spcAft><a:spcPts val="${spaceAfter}"/></a:spcAft>`;
  if (bullet) {
    pPr += `<a:buClr><a:srgbClr val="${C.accent}"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="\u25C6"/>`;
  } else {
    pPr += `<a:buNone/>`;
  }
  pPr += `</a:pPr>`;
  const body = runs.map(r => `<a:r>${rPr(r)}<a:t>${esc(r.text)}</a:t></a:r>`).join('');
  return `<a:p>${pPr}${body}</a:p>`;
}

function txBox(x, y, w, h, paras, opts = {}) {
  const id = nextId();
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="tb${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${inch(x)}" y="${inch(y)}"/><a:ext cx="${inch(w)}" cy="${inch(h)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
    `<p:txBody>${bodyPr({ anchor: opts.anchor || 't' })}<a:lstStyle/>` +
    paras.join('') + `</p:txBody></p:sp>`;
}

function rrect(x, y, w, h, fill, opts = {}) {
  const id = nextId();
  const adj = opts.adj !== undefined ? opts.adj : 12000;
  const alpha = opts.alpha !== undefined ? `<a:alpha val="${Math.round((1 - opts.alpha) * 100000)}"/>` : '';
  const line = opts.line
    ? `<a:ln w="${opts.lineW || 9525}"><a:solidFill><a:srgbClr val="${opts.line}"/></a:solidFill></a:ln>`
    : `<a:ln><a:noFill/></a:ln>`;
  const rot = opts.rot ? ` rot="${Math.round(opts.rot * 60000)}"` : '';
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="rr${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm${rot}><a:off x="${inch(x)}" y="${inch(y)}"/><a:ext cx="${inch(w)}" cy="${inch(h)}"/></a:xfrm>` +
    `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ${adj}"/></a:avLst></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="${fill}">${alpha}</a:srgbClr></a:solidFill>${line}</p:spPr>` +
    `<p:txBody>${bodyPr()}<a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>`;
}

function rect(x, y, w, h, fill, opts = {}) {
  const id = nextId();
  const line = opts.line
    ? `<a:ln w="${opts.lineW || 9525}"><a:solidFill><a:srgbClr val="${opts.line}"/></a:solidFill></a:ln>`
    : `<a:ln><a:noFill/></a:ln>`;
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="rc${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${inch(x)}" y="${inch(y)}"/><a:ext cx="${inch(w)}" cy="${inch(h)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>${line}</p:spPr>` +
    `<p:txBody>${bodyPr()}<a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>`;
}

function pic(x, y, w, h, relId, name) {
  const id = nextId();
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="${inch(x)}" y="${inch(y)}"/><a:ext cx="${inch(w)}" cy="${inch(h)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln><a:noFill/></a:ln></p:spPr></p:pic>`;
}

// spTree 的子元素是 <xsd:choice maxOccurs="unbounded">，schema 要求按
// sp → grpSp → graphicFrame → cxnSp → pic 的顺序排列（同类可重复）。
// 我们构建时是按语义顺序 push 的（标题、图片、正文…），因此必须在这里排序，
// 否则 PowerPoint 会判定「内容有问题，需要修复」。
const SHAPE_ORDER = { 'p:sp': 0, 'p:grpSp': 1, 'p:graphicFrame': 2, 'p:cxnSp': 3, 'p:pic': 4 };
function spTree(shapes) {
  const sorted = shapes.slice().sort((a, b) => {
    const ra = a.match(/^<p:([A-Za-z]+)/);
    const rb = b.match(/^<p:([A-Za-z]+)/);
    const ka = SHAPE_ORDER['p:' + (ra ? ra[1] : 'sp')] ?? 0;
    const kb = SHAPE_ORDER['p:' + (rb ? rb[1] : 'sp')] ?? 0;
    return ka - kb;
  });
  return `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    sorted.join('') + `</p:spTree>`;
}

// ---------------------------------------------------------------- 行内标记
function runsFromMarkup(txt, baseSize, baseColor) {
  const runs = [];
  let i = 0, buf = '', mode = 'plain';
  const flush = () => { if (buf) { runs.push({ text: buf, mode }); buf = ''; } };
  while (i < txt.length) {
    if (txt.substr(i, 2) === '**') { flush(); mode = mode === 'bold' ? 'plain' : 'bold'; i += 2; continue; }
    if (txt[i] === '`') { flush(); mode = mode === 'code' ? 'plain' : 'code'; i += 1; continue; }
    buf += txt[i]; i += 1;
  }
  flush();
  if (!runs.length) runs.push({ text: '', mode: 'plain' });
  return runs.map(r => {
    if (r.mode === 'bold') return { text: r.text, size: baseSize, color: C.heading, bold: 1, font: 'Microsoft YaHei' };
    if (r.mode === 'code') return { text: r.text, size: baseSize - 100, color: C.code, bold: 0, font: 'Consolas' };
    return { text: r.text, size: baseSize, color: baseColor, bold: 0, font: 'Microsoft YaHei' };
  });
}
const plain = (s) => s.replace(/\*\*/g, '').replace(/`/g, '');

// ---------------------------------------------------------------- 版心
const L = {
  marginL: 0.85, marginR: 0.85,
  headerTop: 0.52, headerH: 0.66,
  bodyTop: 1.50, bodyBottom: 7.5 - 0.55,
  contentW: 13.333 - 0.85 * 2
};

function pngSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}
function fitPic(file, boxX, boxY, boxW, boxH) {
  const { w: pw, h: ph } = pngSize(file);
  const ar = pw / ph;
  let w = boxW, h = w / ar;
  if (h > boxH) { h = boxH; w = h * ar; }
  return { x: boxX + (boxW - w) / 2, y: boxY + (boxH - h) / 2, w, h };
}

// ---------------------------------------------------------------- 构建
const media = [];
const slideXmls = [];

function addImage(slide, file, boxX, boxY, boxW, boxH) {
  let m = media.find(x => x.file === file);
  if (!m) { m = { name: 'image' + (media.length + 1) + '.png', file }; media.push(m); }
  const relId = 'rIdImg' + (media.indexOf(m) + 1);
  slide.rels.push({ id: relId, type: 'image', target: '../media/' + m.name });
  const p = fitPic(path.join(IMG, file), boxX, boxY, boxW, boxH);
  slide.shapes.push(pic(p.x, p.y, p.w, p.h, relId, file));
  return p;
}

function header(sh, title) {
  sh.push(rect(L.marginL, L.headerTop + 0.30, 0.78, 0.065, C.accent));
  sh.push(txBox(L.marginL, L.headerTop, L.contentW, L.headerH,
    [para([{ text: title, size: 2700, color: C.title, bold: 1 }])], { anchor: 'b' }));
}

let overflowCount = 0;

for (const sd of deck.slides) {
  const slide = { shapes: [], rels: [] };
  const sh = slide.shapes;
  const pageNo = slideXmls.length + 1;

  if (sd.layout === 'cover') {
    sh.push(rrect(-2.0, -1.6, 8.6, 11.2, C.cover1, { alpha: 0.55, rot: 18 }));
    sh.push(rrect(7.6, 2.2, 8.4, 9.0, C.cover2, { alpha: 0.45, rot: -14 }));
    sh.push(rect(1.0, 2.02, 1.5, 0.075, C.accent));
    sh.push(txBox(1.0, 2.22, 11.3, 1.0,
      [para([{ text: sd.title, size: 4000, color: C.title, bold: 1 }])]));
    sh.push(txBox(1.0, 3.40, 11.0, 0.5,
      [para([{ text: sd.subtitle, size: 1900, color: C.bullet }])]));
    let y = 4.32;
    for (const ln of sd.lines) {
      const txt = ln.replace(/^N\|/, '');
      sh.push(txBox(1.0, y, 11.0, 0.36, [para([{ text: txt, size: 1400, color: C.note }])]));
      y += 0.42;
    }
  } else if (sd.layout === 'toc') {
    header(sh, sd.title);
    const items = sd.lines.map(l => l.replace(/^B\|/, ''));
    const s = 1900;
    const lineH = s * LINE_FACTOR['Microsoft YaHei'] / 72;
    const need = items.length * (lineH + 0.14) + 0.7;
    const cardH = Math.min(L.bodyBottom - L.bodyTop, Math.max(2.0, need));
    sh.push(rrect(L.marginL, L.bodyTop, L.contentW, cardH, C.panel, { adj: 3000, line: C.line }));
    const paras = items.map(t => para([{ text: t, size: s, color: C.bullet }],
      { bullet: true, marL: 228600, indent: -228600, spaceAfter: 1600 }));
    sh.push(txBox(L.marginL + 0.5, L.bodyTop + 0.36, L.contentW - 1.0, cardH - 0.7, paras));
  } else if (sd.layout === 'image') {
    header(sh, sd.title);
    const capH = sd.caption ? 0.40 : 0;
    const maxH = L.bodyBottom - L.bodyTop - capH - 0.10;
    addImage(slide, sd.image, L.marginL, L.bodyTop, L.contentW, maxH);
    if (sd.caption) {
      sh.push(txBox(L.marginL, L.bodyTop + maxH + 0.10, L.contentW, 0.32,
        [para([{ text: sd.caption, size: 1300, color: C.muted }])]));
    }
  } else {
    header(sh, sd.title);
    const withImg = sd.image && sd.image !== '';
    let textL = L.marginL, textW = L.contentW;
    if (withImg) {
      const colW = 5.00;
      addImage(slide, sd.image, 13.333 - L.marginR - colW, L.bodyTop, colW, L.bodyBottom - L.bodyTop);
      textW = 13.333 - L.marginL - L.marginR - colW - 0.40;
    }

    // 逐行度量 → 生成块
    const BULLET_MAR = 0.30;
    const blocks = [];
    for (const ln of sd.lines) {
      if (!ln) { blocks.push({ empty: true, height: 0.09, gapBefore: 0 }); continue; }
      const kind = ln.slice(0, 2), txt = ln.slice(2);
      if (kind === 'H|') {
        const size = 1700;
        const m = measureText(plain(txt), size, textW, 'Microsoft YaHei');
        blocks.push({
          paras: [para(runsFromMarkup(txt, size, C.heading))],
          height: m.height + 0.05, gapBefore: 0.11
        });
      } else if (kind === 'C|') {
        const size = 1250;
        const inner = textW - 0.32;
        const m = measureText(txt, size, inner, 'Consolas');
        blocks.push({
          code: txt, codeSize: size, inner,
          height: Math.max(0.33, m.height + 0.13), gapBefore: 0.04
        });
      } else if (kind === 'Q|') {
        const size = 1400;
        const inner = textW - 0.34;
        const m = measureText(plain(txt), size, inner, 'Microsoft YaHei');
        blocks.push({
          paras: [para(runsFromMarkup(txt, size, C.quote), { lineSpacing: 104 })],
          height: m.height + 0.05, gapBefore: 0.075, quote: true
        });
      } else if (kind === 'N|') {
        const size = 1300;
        const inner = textW - 0.34;
        const m = measureText(plain(txt), size, inner, 'Microsoft YaHei');
        blocks.push({
          paras: [para(runsFromMarkup(txt, size, C.note), { lineSpacing: 104 })],
          height: m.height + 0.03, gapBefore: 0.025, indent: 0.30
        });
      } else {
        const size = withImg ? 1350 : 1450;
        const inner = textW - BULLET_MAR - 0.06;
        const m = measureText(plain(txt), size, inner, 'Microsoft YaHei');
        blocks.push({
          paras: [para(runsFromMarkup(txt, size, C.bullet), { lineSpacing: 104 })],
          height: m.height + 0.04, gapBefore: 0.08, bullet: true, indent: BULLET_MAR
        });
      }
    }

    // 放置：先尝试自然排布，放不下则压缩间距（不改字号）
    const natural = blocks.reduce((a, b) => a + b.height + (b.gapBefore || 0), 0);
    const avail = L.bodyBottom - L.bodyTop;
    const gaps = blocks.reduce((a, b) => a + (b.gapBefore || 0), 0);
    const fixed = natural - gaps;
    let squeeze = 1;
    if (natural > avail && gaps > 0) {
      // 只压缩空白间距，尽量保留文字大小
      squeeze = Math.max(0, (avail - fixed) / gaps);
    }
    let y = L.bodyTop;
    for (const b of blocks) {
      y += (b.gapBefore || 0) * squeeze;
      if (b.empty) { y += b.height; continue; }
      const ind = b.indent || 0;
      if (b.code !== undefined) {
        sh.push(rrect(textL + ind, y, textW - ind, b.height, C.codeBg, { adj: 22000 }));
        sh.push(txBox(textL + ind + 0.16, y + 0.065, b.inner, b.height - 0.11,
          [para([{ text: b.code, size: b.codeSize, color: C.code, font: 'Consolas' }])]));
      } else if (b.quote) {
        sh.push(rect(textL + ind, y + 0.015, 0.045, Math.max(0.18, b.height - 0.03), C.quote));
        sh.push(txBox(textL + ind + 0.20, y, textW - ind - 0.24, b.height, b.paras));
      } else {
        if (b.bullet) {
          sh.push(txBox(textL, y + 0.03, 0.24, 0.28,
            [para([{ text: '\u25C6', size: 1000, color: C.accent }])]));
        }
        sh.push(txBox(textL + ind, y, textW - ind, b.height, b.paras));
      }
      y += b.height;
    }
    if (y > L.bodyBottom + 0.02) {
      overflowCount++;
      console.log(`  ! p${pageNo} 内容到 ${y.toFixed(2)}"，超出下边界 ${(y - L.bodyBottom).toFixed(2)}"（normAutofit 兜底）`);
    }
  }
  slideXmls.push(slide);
}

slideXmls.forEach((s, i) => {
  if (i === 0) return;
  s.shapes.push(txBox(11.9, L.bodyBottom + 0.08, 1.0, 0.3,
    [para([{ text: String(i + 1), size: 1100, color: C.muted, font: 'Consolas' }], { algn: 'r' })]));
});

// ---------------------------------------------------------------- 部件
function slideXml(slide) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
    `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
    `<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${C.bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` +
    spTree(slide.shapes) + `</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}
function slideRels(slide) {
  const rels = [{ id: 'rId1', type: 'slideLayout', target: '../slideLayouts/slideLayout1.xml' }];
  for (const r of slide.rels) rels.push(r);
  const body = rels.map(r => {
    const t = r.type === 'image'
      ? 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
      : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
    return `<Relationship Id="${r.id}" Type="${t}" Target="${r.target}"/>`;
  }).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
}

const slideMaster = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<p:sldMaster preserve="1" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
  `<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${C.bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` +
  spTree([]) + `</p:cSld>` +
  `<p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
  `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>` +
  `<p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="2700" b="1"><a:solidFill><a:srgbClr val="${C.title}"/></a:solidFill><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:defRPr></a:lvl1pPr></p:titleStyle>` +
  `<p:bodyStyle><a:lvl1pPr algn="l"><a:defRPr sz="1400"><a:solidFill><a:srgbClr val="${C.bullet}"/></a:solidFill><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:defRPr></a:lvl1pPr></p:bodyStyle>` +
  `<p:otherStyle><a:lvl1pPr algn="l"><a:defRPr sz="1400"><a:solidFill><a:srgbClr val="${C.bullet}"/></a:solidFill></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>`;

const slideMasterRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
  (F.theme
    ? `<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>`
    : '') +
  `</Relationships>`;

// ---------------------------------------------------------------- 主题部件
// 主题（theme1.xml）是 PowerPoint 期望的必需部件：它提供配色方案、字体方案
// 与格式方案（fillStyleLst / lnStyleLst / effectStyleLst / bgFillStyleLst）。
// 缺失时 PowerPoint 会提示「内容需要修复」。
// 下面这份是按标准 Office 主题结构自己生成的（深色配色 + 微软雅黑）。
const themeFonts = [
  ['zh-CN', '宋体'], ['ja-JP', 'ＭＳ Ｐゴシック'], ['ko-KR', '맑은 고딕'],
  ['zh-TW', '新細明體'], ['ru-RU', 'Arial'], ['en-US', 'Microsoft YaHei'], ['ar-SA', 'Arial']
];
const fontSub = (fonts) => themeFonts.map(([s, t]) =>
  `<a:font script="${s}" typeface="${t}"/>`).join('');

const theme = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="GitTutorial">` +
  `<a:themeElements>` +
  // 配色方案
  `<a:clrScheme name="GitTutorial">` +
  `<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>` +
  `<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>` +
  `<a:dk2><a:srgbClr val="${C.bg}"/></a:dk2>` +
  `<a:lt2><a:srgbClr val="${C.panel}"/></a:lt2>` +
  `<a:accent1><a:srgbClr val="${C.accent}"/></a:accent1>` +
  `<a:accent2><a:srgbClr val="${C.heading}"/></a:accent2>` +
  `<a:accent3><a:srgbClr val="${C.quote}"/></a:accent3>` +
  `<a:accent4><a:srgbClr val="${C.title}"/></a:accent4>` +
  `<a:accent5><a:srgbClr val="${C.code}"/></a:accent5>` +
  `<a:accent6><a:srgbClr val="${C.muted}"/></a:accent6>` +
  `<a:hlink><a:srgbClr val="${C.accent}"/></a:hlink>` +
  `<a:folHlink><a:srgbClr val="${C.quote}"/></a:folHlink>` +
  `</a:clrScheme>` +
  // 字体方案
  `<a:fontScheme name="GitTutorial">` +
  `<a:majorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/>` +
  `<a:cs typeface=""/>${fontSub()}</a:majorFont>` +
  `<a:minorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/>` +
  `<a:cs typeface=""/>${fontSub()}</a:minorFont>` +
  `</a:fontScheme>` +
  // 格式方案（必须四个列表、每类 3 项）
  `<a:fmtScheme name="GitTutorial">` +
  `<a:fillStyleLst>` +
  `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
  `<a:gradFill rotWithShape="1"><a:gsLst>` +
  `<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>` +
  `<a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>` +
  `<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>` +
  `</a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>` +
  `<a:gradFill rotWithShape="1"><a:gsLst>` +
  `<a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs>` +
  `<a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs>` +
  `<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs>` +
  `</a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill>` +
  `</a:fillStyleLst>` +
  `<a:lnStyleLst>` +
  `<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln>` +
  `<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>` +
  `<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>` +
  `</a:lnStyleLst>` +
  `<a:effectStyleLst>` +
  `<a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>` +
  `<a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>` +
  `<a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle>` +
  `</a:effectStyleLst>` +
  `<a:bgFillStyleLst>` +
  `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
  `<a:gradFill rotWithShape="1"><a:gsLst>` +
  `<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>` +
  `<a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs>` +
  `<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>` +
  `</a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill>` +
  `<a:gradFill rotWithShape="1"><a:gsLst>` +
  `<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs>` +
  `<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs>` +
  `</a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>` +
  `</a:bgFillStyleLst>` +
  `</a:fmtScheme>` +
  `</a:themeElements>` +
  `<a:objectDefaults/>` +
  `<a:extraClrSchemeLst/>` +
  `</a:theme>`;

const slideLayout = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">` +
  `<p:cSld name="Blank">` + spTree([]) + `</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

const slideLayoutRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`;

const nSlides = slideXmls.length;
let presRelsBody = `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>`;
for (let i = 0; i < nSlides; i++) {
  presRelsBody += `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`;
}
const presentationRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${presRelsBody}</Relationships>`;

let sldIdLst = '';
for (let i = 0; i < nSlides; i++) sldIdLst += `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`;
const presentation = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
  `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
  `<p:sldIdLst>${sldIdLst}</p:sldIdLst>` +
  `<p:sldSz cx="${SLIDE_W}" cy="${SLIDE_H}"/>` +
  `<p:notesSz cx="6858000" cy="9144000"/>` +
  `<p:defaultTextStyle><a:defPPr><a:defRPr lang="zh-CN"/></a:defPPr></p:defaultTextStyle>` +
  `</p:presentation>`;

let overrides = '';
for (let i = 0; i < nSlides; i++) {
  overrides += `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
}
let imgDefaults = '';
const exts = new Set(media.map(m => path.extname(m.name).slice(1).toLowerCase()));
for (const e of exts) imgDefaults += `<Default Extension="${e}" ContentType="image/${e === 'jpg' ? 'jpeg' : e}"/>`;
const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` + imgDefaults +
  `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
  `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
  `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
  overrides +
  `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
  `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>` +
  `</Types>`;

const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
  `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
  `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>` +
  `</Relationships>`;

const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
const core = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
  `<dc:title>${esc(deck.meta.title)}</dc:title><dc:creator>rrafiya</dc:creator>` +
  `<cp:lastModifiedBy>rrafiya</cp:lastModifiedBy>` +
  `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>` +
  `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;

const app = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
  `<Application>Microsoft Office PowerPoint</Application><Slides>${nSlides}</Slides>` +
  `<PresentationFormat>宽屏</PresentationFormat></Properties>`;

// ---------------------------------------------------------------- 补充标准部件
// PowerPoint 期望演示文稿包里有这几份设置部件；缺失时它常会提示
// 「内容有问题，需要修复」。参考真实 PPTX 后逐一补齐。
const presProps = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
  `<p:showPr useTimings="0" showNarration="0" loop="0">` +
  `<p:present/><p:sldAll/><p:penClr><a:srgbClr val="FF0000"/></p:penClr>` +
  `</p:showPr><p:clrMru><a:srgbClr val="${C.accent}"/><a:srgbClr val="${C.title}"/>` +
  `<a:srgbClr val="${C.heading}"/><a:srgbClr val="${C.bg}"/><a:srgbClr val="FFFFFF"/></p:clrMru></p:presentationPr>`;

const viewProps = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" lastView="sldView">` +
  `<p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>` +
  `<p:slideViewPr><p:cSldViewPr snapToGrid="0" snapToObjects="1">` +
  `<p:cViewPr varScale="1" zoomScale="100"><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr>` +
  `<p:guideLst/></p:cSldViewPr></p:slideViewPr>` +
  `<p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr>` +
  `<p:gridSpacing cx="72008" cy="72008"/></p:viewPr>`;

const tableStyles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`;

// presentation.xml.rels 需要额外挂上 presProps / viewProps / tableStyles
// （这三个部件通过关系类型与 presentation 关联，XML 内部无需额外引用）
const presentationRels2 = presentationRels.replace(
  '</Relationships>',
  (F.exProps
    ? `<Relationship Id="rIdProps" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>` +
      `<Relationship Id="rIdView" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>` +
      `<Relationship Id="rIdTbl" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>`
    : '') +
  `</Relationships>`
);

const contentTypes2 = contentTypes.replace(
  `<Override PartName="/docProps/core.xml"`,
  (F.exProps
    ? `<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>` +
      `<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>` +
      `<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>`
    : '') +
  (F.theme
    ? `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`
    : '') +
  `<Override PartName="/docProps/core.xml"`
);

// ---------------------------------------------------------------- zip
function zip(files) {
  const chunks = [], central = [];
  let offset = 0;
  const t = new Int32Array(256);
  for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c; }
  const crc32 = (buf) => { let c = -1; for (let i = 0; i < buf.length; i++) c = t[(c ^ buf[i]) & 0xFF] ^ (c >>> 8); return (c ^ -1) >>> 0; };
  for (const f of files) {
    const nameBuf = Buffer.from(f.name, 'utf8');
    const data = Buffer.isBuffer(f.data) ? f.data : Buffer.from(f.data, 'utf8');
    const comp = zlib.deflateRawSync(data, { level: 6 });
    const useStore = comp.length >= data.length;
    const body = useStore ? data : comp;
    const method = useStore ? 0 : 8;
    const crc = crc32(data);
    const local = Buffer.alloc(30 + nameBuf.length);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(method, 8); local.writeUInt16LE(0, 10); local.writeUInt16LE(0, 12);
    local.writeUInt32LE(crc, 14); local.writeUInt32LE(body.length, 18); local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26); local.writeUInt16LE(0, 28);
    nameBuf.copy(local, 30);
    chunks.push(local, body);
    const cd = Buffer.alloc(46 + nameBuf.length);
    cd.writeUInt32LE(0x02014b50, 0); cd.writeUInt16LE(20, 4); cd.writeUInt16LE(20, 6);
    cd.writeUInt16LE(0x0800, 8); cd.writeUInt16LE(method, 10); cd.writeUInt16LE(0, 12); cd.writeUInt16LE(0, 14);
    cd.writeUInt32LE(crc, 16); cd.writeUInt32LE(body.length, 20); cd.writeUInt32LE(data.length, 24);
    cd.writeUInt16LE(nameBuf.length, 28); cd.writeUInt16LE(0, 30); cd.writeUInt16LE(0, 32);
    cd.writeUInt16LE(0, 34); cd.writeUInt16LE(0, 36); cd.writeUInt32LE(0, 38); cd.writeUInt32LE(offset, 42);
    nameBuf.copy(cd, 46);
    central.push(cd);
    offset += local.length + body.length;
  }
  const cdBuf = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(0, 4); end.writeUInt16LE(0, 6);
  end.writeUInt16LE(files.length, 8); end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(cdBuf.length, 12); end.writeUInt32LE(offset, 16); end.writeUInt16LE(0, 20);
  return Buffer.concat([...chunks, cdBuf, end]);
}

const files = [
  { name: '[Content_Types].xml', data: contentTypes2 },
  { name: '_rels/.rels', data: rootRels },
  { name: 'docProps/core.xml', data: core },
  { name: 'docProps/app.xml', data: app },
  { name: 'ppt/presentation.xml', data: presentation },
  { name: 'ppt/_rels/presentation.xml.rels', data: presentationRels2 },
  { name: 'ppt/slideMasters/slideMaster1.xml', data: slideMaster },
  { name: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', data: slideMasterRels },
  ...(F.theme ? [{ name: 'ppt/theme/theme1.xml', data: theme }] : []),
  ...(F.exProps ? [
    { name: 'ppt/presProps.xml', data: presProps },
    { name: 'ppt/viewProps.xml', data: viewProps },
    { name: 'ppt/tableStyles.xml', data: tableStyles }
  ] : []),
  { name: 'ppt/slideLayouts/slideLayout1.xml', data: slideLayout },
  { name: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', data: slideLayoutRels }
];
slideXmls.forEach((s, i) => {
  files.push({ name: `ppt/slides/slide${i + 1}.xml`, data: slideXml(s) });
  files.push({ name: `ppt/slides/_rels/slide${i + 1}.xml.rels`, data: slideRels(s) });
});
for (const m of media) files.push({ name: 'ppt/media/' + m.name, data: fs.readFileSync(path.join(IMG, m.file)) });

const buf = zip(files);
fs.writeFileSync(OUT, buf);
console.log('slides   :', nSlides);
console.log('images   :', media.length);
console.log('parts    :', files.length);
console.log('size     :', (buf.length / 1024).toFixed(0), 'KB');
console.log('overflow :', overflowCount, '页（已由 normAutofit 兜底）');
console.log('SAVED    :', OUT);
