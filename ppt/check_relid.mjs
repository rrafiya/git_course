// 校验关系 Id 是否符合 OPC 要求的 xsd:ID (NCName) 形式
import fs from 'node:fs';
import zlib from 'node:zlib';

function relIds(file) {
  const b = fs.readFileSync(file);
  let e = -1;
  for (let i = b.length - 22; i >= 0; i--) if (b.readUInt32LE(i) === 0x06054b50) { e = i; break; }
  const n = b.readUInt16LE(e + 10), cdOff = b.readUInt32LE(e + 16);
  const ids = new Set();
  const byFile = new Map();
  let p = cdOff;
  for (let i = 0; i < n; i++) {
    const nl = b.readUInt16LE(p + 28), el = b.readUInt16LE(p + 30), cl = b.readUInt16LE(p + 32);
    const lho = b.readUInt32LE(p + 42), csize = b.readUInt32LE(p + 20), method = b.readUInt16LE(p + 10);
    const name = b.toString('utf8', p + 46, p + 46 + nl);
    if (name.endsWith('.rels')) {
      const lnl = b.readUInt16LE(lho + 26), lel = b.readUInt16LE(lho + 28);
      const raw = b.subarray(lho + 30 + lnl + lel, lho + 30 + lnl + lel + csize);
      const s = (method === 8 ? zlib.inflateRawSync(raw) : Buffer.from(raw)).toString('utf8');
      const list = [];
      for (const m of s.matchAll(/Id="([^"]+)"/g)) { ids.add(m[1]); list.push(m[1]); }
      byFile.set(name, list);
    }
    p += 46 + nl + el + cl;
  }
  return { ids: [...ids], byFile };
}

const NCNAME = /^[A-Za-z_][A-Za-z0-9._-]*$/;

for (const file of process.argv.slice(2)) {
  const { ids, byFile } = relIds(file);
  console.log('\n### ' + file.split(/[\\/]/).pop());
  console.log('  关系 Id 总数:', ids.length);
  console.log('  样例:', ids.slice(0, 12).join(', '));
  const bad = ids.filter(id => !NCNAME.test(id));
  console.log('  不符合 NCName 的 Id:', bad.length ? bad.join(', ') + '  ❌' : '无  ✅');
  // 检查每个 rels 文件内是否有重复 Id
  let dup = 0;
  for (const [f, list] of byFile) {
    const s = new Set();
    for (const id of list) { if (s.has(id)) { console.log('  ❌ 重复 Id in ' + f + ': ' + id); dup++; } s.add(id); }
  }
  if (!dup) console.log('  同文件内无重复 Id  ✅');
  // 检查 Id 是否与目标类型约定的前缀冲突
  const weird = ids.filter(id => /^(rIdImg|rIdProps|rIdView|rIdTbl)$/.test(id));
  if (weird.length) console.log('  非标准 Id 命名（可能被严格校验拒绝）:', weird.join(', '), ' ⚠');
}
