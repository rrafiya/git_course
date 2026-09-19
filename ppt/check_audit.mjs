// 全面 schema 体检：坐标范围、必需属性、重复 id、内容类型一致性等
import fs from 'node:fs';
import zlib from 'node:zlib';

const file = process.argv[2] || 'ppt/Git与GitHub从入门到实战.pptx';
const b = fs.readFileSync(file);
let e = -1;
for (let i = b.length - 22; i >= 0; i--) if (b.readUInt32LE(i) === 0x06054b50) { e = i; break; }
const n = b.readUInt16LE(e + 10), cdOff = b.readUInt32LE(e + 16);
const parts = new Map();
let p = cdOff;
for (let i = 0; i < n; i++) {
  const nl = b.readUInt16LE(p + 28), el = b.readUInt16LE(p + 30), cl = b.readUInt16LE(p + 32);
  const lho = b.readUInt32LE(p + 42), csize = b.readUInt32LE(p + 20);
  const name = b.toString('utf8', p + 46, p + 46 + nl);
  const lnl = b.readUInt16LE(lho + 26), lel = b.readUInt16LE(lho + 28);
  parts.set(name, zlib.inflateRawSync(b.subarray(lho + 30 + lnl + lel, lho + 30 + lnl + lel + csize)).toString('utf8'));
  p += 46 + nl + el + cl;
}

const EMU = 914400;
const SLIDE_W = 12192000, SLIDE_H = 6858000;
const slides = [...parts.keys()].filter(k => /^ppt\/slides\/slide\d+\.xml$/.test(k))
  .sort((a, c) => (+a.match(/(\d+)/)[1]) - (+c.match(/(\d+)/)[1]));

let problems = 0;
const P = (m) => { console.log('  ❌ ' + m); problems++; };

console.log('=== 1. 形状 id 唯一性 ===');
for (const s of slides) {
  const xml = parts.get(s);
  const ids = [...xml.matchAll(/<p:cNvPr id="(\d+)"/g)].map(m => +m[1]);
  const dup = ids.filter((v, i) => ids.indexOf(v) !== i);
  if (dup.length) P(`${s} 形状 id 重复: ${[...new Set(dup)].join(',')}`);
  if (ids.some(v => v <= 1 && ids.length > 1 && ids.filter(x => x === 1).length > 1)) P(`${s} id=1 被重复使用`);
}
console.log('  已检查', slides.length, '页');

console.log('');
console.log('=== 2. 坐标是否在画布范围内 ===');
for (const s of slides) {
  const xml = parts.get(s);
  const offs = [...xml.matchAll(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/g)];
  const exts = [...xml.matchAll(/<a:ext cx="(\d+)" cy="(\d+)"\/>/g)];
  // 装饰性光晕允许出血，只查明显离谱的值
  for (let i = 0; i < Math.min(offs.length, exts.length); i++) {
    const x = +offs[i][1], y = +offs[i][2], cx = +exts[i][1], cy = +exts[i][2];
    if (Math.abs(x) > SLIDE_W * 2 || Math.abs(y) > SLIDE_H * 2) P(`${s} 第${i}个形状坐标离谱 x=${x / EMU}" y=${y / EMU}"`);
    if (cx > SLIDE_W * 2 || cy > SLIDE_H * 2) P(`${s} 第${i}个形状尺寸离谱 cx=${(cx / EMU).toFixed(1)}" cy=${(cy / EMU).toFixed(1)}"`);
  }
}
console.log('  已检查坐标范围');

console.log('');
console.log('=== 3. 必需属性 ===');
for (const s of slides) {
  const xml = parts.get(s);
  for (const m of xml.matchAll(/<a:srgbClr(?![^>]*val=)[^>]*>/g)) P(`${s} 有 srgbClr 缺少 val 属性`);
  for (const m of xml.matchAll(/<a:rPr(?![^>]*lang=)[^>]*>/g)) P(`${s} 有 rPr 缺少 lang 属性`);
  for (const m of xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)) {
    if (/[<>&](?!amp;|lt;|gt;|quot;|apos;|#)/.test(m[1])) P(`${s} 文本含未转义字符: ${m[1].slice(0, 30)}`);
  }
  for (const m of xml.matchAll(/<p:cNvPr id="\d+"(?![^>]*name=)[^>]*>/g)) P(`${s} cNvPr 缺少 name 属性`);
}
console.log('  已检查必需属性');

console.log('');
console.log('=== 4. spTree 子元素顺序（逐页） ===');
const ORDER = { sp: 0, grpSp: 1, graphicFrame: 2, cxnSp: 3, pic: 4 };
for (const s of slides) {
  const xml = parts.get(s);
  const i0 = xml.indexOf('<p:spTree>'), i1 = xml.indexOf('</p:spTree>');
  const tree = xml.slice(i0 + 11, i1);
  // 只取顶层子元素（深度 1）
  const tags = [];
  let depth = 0;
  const re = /<(\/?)(p:[A-Za-z]+)([^>]*?)(\/?)>/g;
  let m;
  while ((m = re.exec(tree))) {
    const closing = m[1] === '/', tag = m[2].slice(2), self = m[4] === '/';
    if (depth === 0 && !closing) tags.push(tag);
    if (!self) depth += closing ? -1 : 1;
    else if (!closing) { /* self-closing 不改变深度 */ }
  }
  const keys = tags.filter(t => ORDER[t] !== undefined).map(t => ORDER[t]);
  for (let i = 1; i < keys.length; i++) {
    if (keys[i] < keys[i - 1]) P(`${s} spTree 顺序: ${tags.filter(t => ORDER[t] !== undefined).join(' -> ')}`);
  }
}
console.log('  已检查', slides.length, '页的 spTree 顺序');

console.log('');
console.log('=== 5. 图片关系与 media 一一对应 ===');
const relImgs = new Set();
for (const k of [...parts.keys()].filter(x => x.endsWith('.rels'))) {
  for (const m of parts.get(k).matchAll(/Target="\.\.\/media\/([^"]+)"/g)) relImgs.add(m[1]);
}
const mediaFiles = new Set([...parts.keys()].filter(k => k.startsWith('ppt/media/')).map(k => k.replace('ppt/media/', '')));
for (const f of relImgs) if (!mediaFiles.has(f)) P(`关系引用了不存在的图片 ${f}`);
console.log(`  被引用的图片 ${relImgs.size} 个，media 目录 ${mediaFiles.size} 个`);

console.log('');
console.log(problems ? `❌ 共发现 ${problems} 处问题` : '✅ 全部体检项目通过');
process.exit(problems ? 1 : 0);
