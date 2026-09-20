// 体检任意 pptx：部件清单 + 关键部件 + spTree 顺序 + 是否可直接用
import fs from 'node:fs';
import zlib from 'node:zlib';

const file = process.argv[2];
const b = fs.readFileSync(file);
let e = -1;
for (let i = b.length - 22; i >= 0; i--) if (b.readUInt32LE(i) === 0x06054b50) { e = i; break; }
if (e < 0) { console.log('❌ 不是有效的 ZIP/OOXML 文件'); process.exit(1); }
const n = b.readUInt16LE(e + 10), cdOff = b.readUInt32LE(e + 16);

const parts = new Map();
let p = cdOff;
let dirs = 0;
for (let i = 0; i < n; i++) {
  const nl = b.readUInt16LE(p + 28), el = b.readUInt16LE(p + 30), cl = b.readUInt16LE(p + 32);
  const lho = b.readUInt32LE(p + 42), csize = b.readUInt32LE(p + 20), method = b.readUInt16LE(p + 10);
  const name = b.toString('utf8', p + 46, p + 46 + nl);
  const lnl = b.readUInt16LE(lho + 26), lel = b.readUInt16LE(lho + 28);
  if (name.endsWith('/') || csize === 0) { dirs++; p += 46 + nl + el + cl; continue; }
  const raw = b.subarray(lho + 30 + lnl + lel, lho + 30 + lnl + lel + csize);
  parts.set(name, method === 8 ? zlib.inflateRawSync(raw) : Buffer.from(raw));
  p += 46 + nl + el + cl;
}

console.log('文件:', file.split(/[\\/]/).pop(), (b.length / 1024).toFixed(0), 'KB');
console.log('ZIP 条目:', n, '（其中目录项', dirs, '个已忽略）');
console.log('实际部件:', parts.size);
console.log('');

console.log('=== 目录分布 ===');
const dd = {};
for (const k of parts.keys()) { const d = k.includes('/') ? k.split('/').slice(0, -1).join('/') : '(root)'; dd[d] = (dd[d] || 0) + 1; }
for (const [d, c] of Object.entries(dd).sort()) console.log('  ' + d.padEnd(28) + c);

console.log('');
console.log('=== 关键部件（PowerPoint 是否可能报修复）===');
const need = [
  ['ppt/theme/theme1.xml', '主题（缺它必报修复）'],
  ['ppt/presProps.xml', '放映设置'],
  ['ppt/viewProps.xml', '视图设置'],
  ['ppt/tableStyles.xml', '表格样式'],
  ['ppt/presentation.xml', '主部件'],
  ['ppt/slideMasters/slideMaster1.xml', '母版'],
  ['ppt/slideLayouts/slideLayout1.xml', '版式']
];
for (const [k, desc] of need) {
  console.log('  ' + (parts.has(k) ? '有' : '无') + '  ' + k.padEnd(40) + desc);
}

console.log('');
const slides = [...parts.keys()].filter(k => /^ppt\/slides\/slide\d+\.xml$/.test(k))
  .sort((a, c) => (+a.match(/\d+/)[0]) - (+c.match(/\d+/)[0]));
console.log('幻灯片:', slides.length, '张');
const media = [...parts.keys()].filter(k => k.startsWith('ppt/media/'));
console.log('媒体文件:', media.length, '个');

// spTree 顺序检查（前 5 页 + 最后 1 页）
const ORDER = { sp: 0, grpSp: 1, graphicFrame: 2, cxnSp: 3, pic: 4 };
console.log('');
console.log('=== spTree 子元素顺序（抽样）===');
const sample = slides.length <= 6 ? slides : [...slides.slice(0, 5), slides[slides.length - 1]];
let badSeq = 0;
for (const s of sample) {
  const xml = parts.get(s).toString('utf8');
  const i0 = xml.indexOf('<p:spTree>'), i1 = xml.indexOf('</p:spTree>');
  if (i0 < 0) { console.log('  ' + s.split('/').pop().padEnd(14) + ' ⚠ 无 spTree'); continue; }
  const tree = xml.slice(i0 + 11, i1);
  const tags = [];
  let depth = 0;
  const re = /<(\/?)(p:[A-Za-z]+)([^>]*?)(\/?)>/g;
  let m;
  while ((m = re.exec(tree))) {
    const closing = m[1] === '/', tag = m[2].slice(2), self = m[4] === '/';
    if (depth === 0 && !closing) tags.push(tag);
    if (!self) depth += closing ? -1 : 1;
  }
  const seq = tags.filter(t => ORDER[t] !== undefined);
  const keys = seq.map(t => ORDER[t]);
  let ok = true;
  for (let i = 1; i < keys.length; i++) if (keys[i] < keys[i - 1]) ok = false;
  if (!ok) badSeq++;
  console.log('  ' + s.split('/').pop().padEnd(14) + (ok ? '✅ 顺序正确' : '❌ 顺序错误') + '   ' + seq.slice(0, 8).join(' -> '));
}

console.log('');
console.log('=== 结论 ===');
const hasTheme = parts.has('ppt/theme/theme1.xml');
if (hasTheme && !badSeq) console.log('  ✅ 含主题、spTree 顺序正确 → 大概率能正常打开');
else {
  if (!hasTheme) console.log('  ⚠ 缺少主题部件 ppt/theme/theme1.xml —— 可能在 PowerPoint 里提示「需要修复」');
  if (badSeq) console.log('  ⚠ 有 ' + badSeq + ' 页 spTree 顺序违规 —— 可能在 PowerPoint 里提示「需要修复」');
}
