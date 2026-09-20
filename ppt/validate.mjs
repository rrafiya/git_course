// 严格校验 pptx：ZIP 完整性 + XML 良构 + 关系/内容类型一致性 + 图片有效性
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';

const file = process.argv[2];
const b = fs.readFileSync(file);
let fails = 0, warns = 0;
const ok = (m) => console.log('  ✅ ' + m);
const bad = (m) => { console.log('  ❌ ' + m); fails++; };
const warn = (m) => { console.log('  ⚠ ' + m); warns++; };

console.log('文件:', path.basename(file), (b.length / 1024).toFixed(0), 'KB');

// ---------- 1. ZIP ----------
let e = -1;
for (let i = b.length - 22; i >= 0; i--) if (b.readUInt32LE(i) === 0x06054b50) { e = i; break; }
if (e < 0) { bad('找不到 EOCD'); process.exit(1); }
const n = b.readUInt16LE(e + 10), cdOff = b.readUInt32LE(e + 16);
const entries = [];
let dirEntries = 0;
let p = cdOff;
for (let i = 0; i < n; i++) {
  if (b.readUInt32LE(p) !== 0x02014b50) { bad('中央目录签名错误'); break; }
  const method = b.readUInt16LE(p + 10), crc = b.readUInt32LE(p + 16);
  const csize = b.readUInt32LE(p + 20), usize = b.readUInt32LE(p + 24);
  const nl = b.readUInt16LE(p + 28), el = b.readUInt16LE(p + 30), cl = b.readUInt16LE(p + 32);
  const lho = b.readUInt32LE(p + 42);
  const nm = b.toString('utf8', p + 46, p + 46 + nl);
  // 跳过显式目录条目：PowerPoint 保存时会写入 "_rels/"、"ppt/media/" 这类
  // 目录项，它们不是文件，不应参与「内容类型覆盖」与「图片有效性」检查。
  if (nm.endsWith('/')) { p += 46 + nl + el + cl; dirEntries++; continue; }
  entries.push({ name: nm, method, crc, csize, usize, lho });
  p += 46 + nl + el + cl;
}
if (entries.length + dirEntries !== n) bad('中央目录条目数不符');
else ok(`ZIP 中央目录完整（${n} 个条目，其中 ${dirEntries} 个目录项已忽略）`);

const tbl = (() => { const t = new Int32Array(256); for (let k = 0; k < 256; k++) { let c = k; for (let j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[k] = c; } return t; })();
const crc32 = (buf) => { let c = -1; for (let i = 0; i < buf.length; i++) c = tbl[(c ^ buf[i]) & 0xFF] ^ (c >>> 8); return (c ^ -1) >>> 0; };

const data = new Map();
let badCrc = 0;
for (const en of entries) {
  const lh = en.lho;
  if (b.readUInt32LE(lh) !== 0x04034b50) { bad('本地头签名错误 ' + en.name); continue; }
  const nl = b.readUInt16LE(lh + 26), el = b.readUInt16LE(lh + 28);
  const start = lh + 30 + nl + el;
  const raw = b.subarray(start, start + en.csize);
  let d;
  try { d = en.method === 8 ? zlib.inflateRawSync(raw) : Buffer.from(raw); }
  catch (err) { bad('解压失败 ' + en.name); continue; }
  if (d.length !== en.usize) { bad('长度不符 ' + en.name); continue; }
  if (crc32(d) !== en.crc) { bad('CRC 不符 ' + en.name); badCrc++; continue; }
  data.set(en.name, d);
}
if (!badCrc) ok(`全部 ${entries.length} 个条目 CRC 与解压长度校验通过`);

// ---------- 2. XML ----------
function wellFormed(s) {
  const stack = [];
  const re = /<(\/?)([A-Za-z_][\w:.-]*)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/g;
  let m;
  while ((m = re.exec(s))) {
    const closing = m[1] === '/', name = m[2], self = m[4] === '/';
    if (closing) {
      if (!stack.length || stack[stack.length - 1] !== name) return `标签不匹配 </${name}>`;
      stack.pop();
    } else if (!self && !/^(br|hr|img|meta|link|input)$/i.test(name)) stack.push(name);
  }
  if (stack.length) return '未闭合: ' + stack.join(',');
  return null;
}
const xmlNames = [...data.keys()].filter(k => k.endsWith('.xml') || k.endsWith('.rels'));
let xmlBad = 0;
for (const k of xmlNames) {
  const err = wellFormed(data.get(k).toString('utf8'));
  if (err) { bad('XML 问题 ' + k + ' -> ' + err); xmlBad++; }
}
if (!xmlBad) ok(`全部 ${xmlNames.length} 个 XML 部件良构`);

// ---------- 3. 关系与内容类型一致性 ----------
const ct = data.get('[Content_Types].xml').toString('utf8');
const defaults = [...ct.matchAll(/<Default Extension="([^"]+)" ContentType="([^"]+)"/g)].map(m => ({ ext: m[1].toLowerCase(), type: m[2] }));
const overrides = [...ct.matchAll(/<Override PartName="([^"]+)" ContentType="([^"]+)"/g)].map(m => ({ part: m[1], type: m[2] }));
ok(`[Content_Types].xml: ${defaults.length} 个 Default + ${overrides.length} 个 Override`);

// 每个部件都应被 Default 或 Override 覆盖
let uncovered = [];
for (const k of data.keys()) {
  if (k === '[Content_Types].xml' || k.endsWith('.rels')) continue;
  const ext = path.extname(k).slice(1).toLowerCase();
  const hasOverride = overrides.some(o => o.part === '/' + k);
  const hasDefault = defaults.some(d => d.ext === ext);
  if (!hasOverride && !hasDefault) uncovered.push(k);
}
if (uncovered.length) bad('以下部件未被内容类型覆盖: ' + uncovered.join(', '));
else ok('所有部件都有对应的内容类型声明');

// Override 指向的部件必须存在
let missingOverride = overrides.filter(o => !data.has(o.part.slice(1)));
if (missingOverride.length) bad('Override 指向不存在的部件: ' + missingOverride.map(o => o.part).join(', '));
else ok('所有 Override 指向的部件都存在');

// 关系目标必须存在
let relBad = 0, relTotal = 0;
for (const k of [...data.keys()].filter(x => x.endsWith('.rels'))) {
  const xml = data.get(k).toString('utf8');
  const baseDir = path.posix.dirname(path.posix.dirname(k));   // _rels 的上一级
  for (const m of xml.matchAll(/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*\/>/g)) {
    relTotal++;
    const target = m[2];
    if (/^https?:/.test(target)) continue;
    const resolved = path.posix.normalize(path.posix.join(baseDir, target));
    if (!data.has(resolved)) { bad(`关系目标缺失: ${k} -> ${target} (解析为 ${resolved})`); relBad++; }
  }
}
if (!relBad) ok(`全部 ${relTotal} 条关系目标都存在`);

// 每个 slide 必须被 presentation 关系引用
const presRels = data.get('ppt/_rels/presentation.xml.rels').toString('utf8');
const presXml = data.get('ppt/presentation.xml').toString('utf8');
const relIds = [...presRels.matchAll(/Id="([^"]+)"/g)].map(m => m[1]);
const usedIds = [...presXml.matchAll(/r:id="([^"]+)"/g)].map(m => m[1]);
const danglingIds = usedIds.filter(id => !relIds.includes(id));
if (danglingIds.length) bad('presentation.xml 引用了不存在的关系: ' + danglingIds.join(', '));
else ok(`presentation.xml 的 ${usedIds.length} 个 r:id 引用全部有效`);

// 每张 slide 都有对应关系文件
const slideFiles = [...data.keys()].filter(k => /^ppt\/slides\/slide\d+\.xml$/.test(k));
let noRels = [];
for (const s of slideFiles) {
  const id = s.match(/slide(\d+)\.xml$/)[1];
  if (!data.has(`ppt/slides/_rels/slide${id}.xml.rels`)) noRels.push(s);
}
if (noRels.length) bad('缺少关系文件的幻灯片: ' + noRels.join(', '));
else ok(`${slideFiles.length} 张幻灯片都有关系文件`);

// ---------- 4. 图片有效性 ----------
const pngs = [...data.keys()].filter(k => k.startsWith('ppt/media/'));
let pngBad = 0;
for (const k of pngs) {
  const d = data.get(k);
  const sig = d.subarray(0, 8).toString('hex');
  if (sig !== '89504e470d0a1a0a') { bad('不是有效 PNG: ' + k); pngBad++; continue; }
  const w = d.readUInt32BE(16), h = d.readUInt32BE(20);
  if (w < 10 || h < 10) { bad('PNG 尺寸异常: ' + k); pngBad++; }
}
if (!pngBad) ok(`${pngs.length} 个图片均为有效 PNG`);

// 图片关系都指向 media
const imgRels = [];
for (const k of [...data.keys()].filter(x => x.endsWith('.rels'))) {
  const xml = data.get(k).toString('utf8');
  for (const m of xml.matchAll(/Type="[^"]*\/image"[^>]*Target="([^"]+)"/g)) imgRels.push(m[1]);
}
ok(`图片关系 ${imgRels.length} 条`);

console.log('');
console.log(`幻灯片 ${slideFiles.length} 张 · 图片 ${pngs.length} 个 · 部件 ${data.size} 个`);
console.log(fails ? `❌ 失败 ${fails} 项` + (warns ? `，警告 ${warns} 项` : '') : '✅ 全部校验通过');
process.exit(fails ? 1 : 0);
