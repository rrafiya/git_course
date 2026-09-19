// 用 Node 直接生成 .pptx（OOXML），不依赖任何外部库。
// 结构与命名空间取自真实 PPTX 参考文件（ppt/ref）。
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';

const ROOT = process.cwd();
const IMG = path.join(ROOT, 'ppt', 'img');
const deck = JSON.parse(fs.readFileSync(path.join(ROOT, 'ppt', 'deck.json'), 'utf8'));
const OUT = path.join(ROOT, 'ppt', deck.meta.out);

const EMU = 914400;                       // EMU per inch
const SLIDE_W = 12192000, SLIDE_H = 6858000;   // 13.333" x 7.5"

// ---------------------------------------------------------------- colours
const C = {
  bg: '0B1020',
  panel: '111A2C',
  line: '23304A',
  accent: '4DD4FF',
  title: '9FE8FF',
  heading: 'FFD166',
  bullet: 'DCE6FA',
  code: '7FE9FF',
  codeBg: '060A14',
  quote: '63F5A0',
  note: '93A3C4',
  muted: '6C7A99',
  cover1: '12325A',
  cover2: '3A1A2C',
  white: 'FFFFFF'
};

// ---------------------------------------------------------------- xml utils
const esc = (s) => String(s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;').replace(/'/g, '&apos;');

const inch = (v) => Math.round(v * EMU);

// rPr for a run
function rPr({ size = 1400, color = C.bullet, bold = 0, font = 'Microsoft YaHei', italic = 0 }) {
  const latin = font === 'Consolas' ? 'Consolas' : font;
  const ea = font === 'Consolas' ? 'Microsoft YaHei' : font;
  return `<a:rPr lang="zh-CN" altLang="en-US" sz="${size}" b="${bold}" i="${italic}" dirty="0">` +
    `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` +
    `<a:latin typeface="${latin}"/><a:ea typeface="${ea}"/><a:cs typeface="${latin}"/></a:rPr>`;
}

// para with runs: [{text, size, color, bold, font}]
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

let shapeIdSeq = 1;
const nextId = () => ++shapeIdSeq;

// simple textbox
function txBox(x, y, w, h, paras, opts = {}) {
  const id = nextId();
  const anchor = opts.anchor || 't';
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="tb${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${inch(x)}" y="${inch(y)}"/><a:ext cx="${inch(w)}" cy="${inch(h)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
    `<p:txBody><a:bodyPr wrap="square" lIns="0" rIns="0" tIns="0" bIns="0" anchor="${anchor}"><a:noAutofit/></a:bodyPr><a:lstStyle/>` +
    paras.join('') + `</p:txBody></p:sp>`;
}

// rounded rectangle (filled, optional border, optional transparency)
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
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`;
}

// plain rectangle
function rect(x, y, w, h, fill, opts = {}) {
  const id = nextId();
  const line = opts.line
    ? `<a:ln w="${opts.lineW || 9525}"><a:solidFill><a:srgbClr val="${opts.line}"/></a:solidFill></a:ln>`
    : `<a:ln><a:noFill/></a:ln>`;
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="rc${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${inch(x)}" y="${inch(y)}"/><a:ext cx="${inch(w)}" cy="${inch(h)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>${line}</p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`;
}

// picture
function pic(x, y, w, h, relId, name) {
  const id = nextId();
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="${inch(x)}" y="${inch(y)}"/><a:ext cx="${inch(w)}" cy="${inch(h)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln><a:noFill/></a:ln></p:spPr></p:pic>`;
}

function spTree(shapes) {
  return `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    shapes.join('') + `</p:spTree>`;
}

// ---------------------------------------------------------------- inline markup
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

// PNG size reader
function pngSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}

// fit a picture into a box, centred
function fitPic(file, boxX, boxY, boxW, boxH) {
  const { w: pw, h: ph } = pngSize(file);
  const ar = pw / ph;
  let w = boxW, h = w / ar;
  if (h > boxH) { h = boxH; w = h * ar; }
  return { x: boxX + (boxW - w) / 2, y: boxY + (boxH - h) / 2, w, h };
}

// ---------------------------------------------------------------- slides
const media = [];         // {name, file}
const slideXmls = [];     // {xml, rels:[{id,type,target}], pics:[...]}

function newSlide() { return { shapes: [], rels: [] }; }

function header(sh, title, size) {
  sh.push(rect(0.95, 0.6, 0.9, 0.07, C.accent));
  sh.push(txBox(0.95, 0.75, 11.4, 0.75, [para([{ text: title, size, color: C.title, bold: 1 }])]));
}

function addImage(slide, file, boxX, boxY, boxW, boxH) {
  const key = 'ppt/img/' + file;
  let m = media.find(x => x.file === file);
  if (!m) { m = { name: 'image' + (media.length + 1) + '.png', file }; media.push(m); }
  const relId = 'rIdImg' + (media.indexOf(m) + 1);
  slide.rels.push({ id: relId, type: 'image', target: '../media/' + m.name });
  const f = path.join(IMG, file);
  const p = fitPic(f, boxX, boxY, boxW, boxH);
  slide.shapes.push(pic(p.x, p.y, p.w, p.h, relId, file));
  return p;
}

for (const sd of deck.slides) {
  const s = newSlide();
  const sh = s.shapes;

  if (sd.layout === 'cover') {
    sh.push(rrect(-2.0, -1.6, 8.6, 11.2, C.cover1, { alpha: 0.55, rot: 18 }));
    sh.push(rrect(7.6, 2.2, 8.4, 9.0, C.cover2, { alpha: 0.45, rot: -14 }));
    sh.push(rect(1.0, 2.02, 1.5, 0.075, C.accent));
    sh.push(txBox(1.0, 2.25, 11.3, 1.4, [para([{ text: sd.title, size: 4000, color: C.title, bold: 1 }])]));
    sh.push(txBox(1.0, 3.55, 11.0, 0.7, [para([{ text: sd.subtitle, size: 1900, color: C.bullet }])]));
    let y = 4.55;
    for (const ln of sd.lines) {
      const txt = ln.replace(/^N\|/, '');
      sh.push(txBox(1.0, y, 11.0, 0.4, [para([{ text: txt, size: 1400, color: C.note }])]));
      y += 0.44;
    }
  } else if (sd.layout === 'toc') {
    header(sh, sd.title, 3000);
    sh.push(rrect(0.95, 2.0, 11.45, 4.7, C.panel, { adj: 3000, line: C.line }));
    const paras = sd.lines.map(ln =>
      para([{ text: ln.replace(/^B\|/, ''), size: 1900, color: C.bullet }],
        { bullet: true, marL: 228600, indent: -228600, spaceAfter: 1200 })
    );
    sh.push(txBox(1.45, 2.35, 10.5, 4.1, paras));
  } else if (sd.layout === 'image') {
    header(sh, sd.title, 2800);
    const capH = 0.5;
    const maxH = 7.5 - 1.68 - capH;
    addImage(s, sd.image, 0.62, 1.68, 12.1, maxH);
    if (sd.caption) {
      sh.push(txBox(0.95, 1.68 + maxH + 0.04, 11.45, 0.42,
        [para([{ text: sd.caption, size: 1400, color: C.muted }])]));
    }
  } else {
    header(sh, sd.title, 2800);
    const withImg = sd.image && sd.image !== '';
    let textL = 0.95, textW = 11.45;
    if (withImg) {
      const colW = 5.2, colH = 5.2;
      addImage(s, sd.image, 13.333 - 0.95 - colW, 1.78, colW, colH);
      textW = 13.333 - 0.95 * 2 - colW - 0.35;
    }
    let y = 1.78;
    for (const ln of sd.lines) {
      if (!ln) { y += 0.1; continue; }
      const kind = ln.slice(0, 2);
      const txt = ln.slice(2);
      if (kind === 'H|') {
        sh.push(txBox(textL, y, textW, 0.44, [para([{ text: txt, size: 1700, color: C.heading, bold: 1 }])]));
        y += 0.54;
      } else if (kind === 'C|') {
        const bh = 0.37;
        sh.push(rrect(textL, y, textW, bh, C.codeBg, { adj: 22000 }));
        sh.push(txBox(textL + 0.16, y + 0.045, textW - 0.3, bh - 0.09,
          [para([{ text: txt, size: 1250, color: C.code, font: 'Consolas' }])]));
        y += 0.43;
      } else if (kind === 'Q|') {
        sh.push(rect(textL, y, 0.045, 0.5, C.quote));
        sh.push(txBox(textL + 0.22, y, textW - 0.25, 0.5,
          [para(runsFromMarkup(txt, 1400, C.quote), { lineSpacing: 105 })]));
        y += 0.58;
      } else if (kind === 'N|') {
        sh.push(txBox(textL + 0.28, y, textW - 0.3, 0.34,
          [para([{ text: txt, size: 1300, color: C.note }])]));
        y += 0.38;
      } else {
        const size = withImg ? 1400 : 1450;
        // diamond bullet
        sh.push(txBox(textL, y + 0.035, 0.22, 0.3,
          [para([{ text: '\u25C6', size: 1000, color: C.accent }])]));
        sh.push(txBox(textL + 0.3, y, textW - 0.32, 0.7,
          [para(runsFromMarkup(txt, size, C.bullet), { lineSpacing: 105 })]));
        y += 0.6;
      }
    }
  }
  slideXmls.push(s);
}

// page numbers
slideXmls.forEach((s, i) => {
  if (i === 0) return;   // 封面不显示页码
  s.shapes.push(txBox(11.9, 6.95, 1.0, 0.35,
    [para([{ text: String(i + 1), size: 1100, color: C.muted, font: 'Consolas' }], { algn: 'r' })]));
});

// ---------------------------------------------------------------- parts
function slideXml(slide, idx) {
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
  `<p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="2800" b="1"><a:solidFill><a:srgbClr val="${C.title}"/></a:solidFill><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:defRPr></a:lvl1pPr></p:titleStyle>` +
  `<p:bodyStyle><a:lvl1pPr algn="l"><a:defRPr sz="1400"><a:solidFill><a:srgbClr val="${C.bullet}"/></a:solidFill><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:defRPr></a:lvl1pPr></p:bodyStyle>` +
  `<p:otherStyle><a:lvl1pPr algn="l"><a:defRPr sz="1400"><a:solidFill><a:srgbClr val="${C.bullet}"/></a:solidFill></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>`;

const slideMasterRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
  `</Relationships>`;

const slideLayout = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">` +
  `<p:cSld name="Blank">` + spTree([]) + `</p:cSld>` +
  `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

const slideLayoutRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>` +
  `</Relationships>`;

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
for (const e of exts) {
  imgDefaults += `<Default Extension="${e}" ContentType="image/${e === 'jpg' ? 'jpeg' : e}"/>`;
}
const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  imgDefaults +
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
  `<dc:title>${esc(deck.meta.title)}</dc:title>` +
  `<dc:creator>rrafiya</dc:creator>` +
  `<cp:lastModifiedBy>rrafiya</cp:lastModifiedBy>` +
  `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>` +
  `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
  `</cp:coreProperties>`;

const app = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
  `<Application>Microsoft Office PowerPoint</Application><Slides>${nSlides}</Slides>` +
  `<PresentationFormat>宽屏</PresentationFormat></Properties>`;

// ---------------------------------------------------------------- zip writer
function zip(files) {
  const chunks = [];
  const central = [];
  let offset = 0;

  const crcTable = (() => {
    const t = new Int32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c;
    }
    return t;
  })();
  function crc32(buf) {
    let c = -1;
    for (let i = 0; i < buf.length; i++) c = crcTable[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ -1) >>> 0;
  }

  for (const f of files) {
    const nameBuf = Buffer.from(f.name, 'utf8');
    const data = Buffer.isBuffer(f.data) ? f.data : Buffer.from(f.data, 'utf8');
    const comp = zlib.deflateRawSync(data, { level: 6 });
    const useStore = comp.length >= data.length;
    const body = useStore ? data : comp;
    const method = useStore ? 0 : 8;
    const crc = crc32(data);

    const local = Buffer.alloc(30 + nameBuf.length);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0800, 6);      // UTF-8 flag
    local.writeUInt16LE(method, 8);
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(body.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);
    nameBuf.copy(local, 30);

    chunks.push(local, body);

    const cd = Buffer.alloc(46 + nameBuf.length);
    cd.writeUInt32LE(0x02014b50, 0);
    cd.writeUInt16LE(20, 4);
    cd.writeUInt16LE(20, 6);
    cd.writeUInt16LE(0x0800, 8);
    cd.writeUInt16LE(method, 10);
    cd.writeUInt16LE(0, 12);
    cd.writeUInt16LE(0, 14);
    cd.writeUInt32LE(crc, 16);
    cd.writeUInt32LE(body.length, 20);
    cd.writeUInt32LE(data.length, 24);
    cd.writeUInt16LE(nameBuf.length, 28);
    cd.writeUInt16LE(0, 30);
    cd.writeUInt16LE(0, 32);
    cd.writeUInt16LE(0, 34);
    cd.writeUInt16LE(0, 36);
    cd.writeUInt32LE(0, 38);
    cd.writeUInt32LE(offset, 42);
    nameBuf.copy(cd, 46);
    central.push(cd);

    offset += local.length + body.length;
  }

  const cdBuf = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(files.length, 8);
  end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(cdBuf.length, 12);
  end.writeUInt32LE(offset, 16);
  end.writeUInt16LE(0, 20);

  return Buffer.concat([...chunks, cdBuf, end]);
}

// ---------------------------------------------------------------- assemble
const files = [
  { name: '[Content_Types].xml', data: contentTypes },
  { name: '_rels/.rels', data: rootRels },
  { name: 'docProps/core.xml', data: core },
  { name: 'docProps/app.xml', data: app },
  { name: 'ppt/presentation.xml', data: presentation },
  { name: 'ppt/_rels/presentation.xml.rels', data: presentationRels },
  { name: 'ppt/slideMasters/slideMaster1.xml', data: slideMaster },
  { name: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', data: slideMasterRels },
  { name: 'ppt/slideLayouts/slideLayout1.xml', data: slideLayout },
  { name: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', data: slideLayoutRels }
];
slideXmls.forEach((s, i) => {
  files.push({ name: `ppt/slides/slide${i + 1}.xml`, data: slideXml(s, i) });
  files.push({ name: `ppt/slides/_rels/slide${i + 1}.xml.rels`, data: slideRels(s) });
});
for (const m of media) {
  files.push({ name: 'ppt/media/' + m.name, data: fs.readFileSync(path.join(IMG, m.file)) });
}

const buf = zip(files);
fs.writeFileSync(OUT, buf);

console.log('slides   :', nSlides);
console.log('images   :', media.length);
console.log('parts    :', files.length);
console.log('size     :', (buf.length / 1024).toFixed(0), 'KB');
console.log('SAVED    :', OUT);
