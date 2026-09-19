// 按 OOXML schema 的「子元素顺序」规则校验生成的 pptx
// DrawingML / PresentationML 是 sequence 类型，子元素顺序错了 PowerPoint 会提示修复
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

const ORDER = {
  'p:sld': ['p:cSld', 'p:clrMapOvr', 'p:transition', 'p:timing', 'p:extLst'],
  'p:cSld': ['p:bg', 'p:spTree', 'p:custDataLst', 'p:controls', 'p:extLst'],
  'p:bgPr': ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill', 'a:effectLst', 'a:effectDag', 'a:extLst'],
  'p:spTree': ['p:nvGrpSpPr', 'p:grpSpPr', 'p:sp', 'p:grpSp', 'p:graphicFrame', 'p:cxnSp', 'p:pic', 'p:contentPart'],
  'p:sp': ['p:nvSpPr', 'p:spPr', 'p:style', 'p:txBody'],
  'p:nvSpPr': ['p:cNvPr', 'p:cNvSpPr', 'p:nvPr'],
  'p:cNvSpPr': ['a:spLocks', 'a:extLst'],
  'p:spPr': ['a:xfrm', 'a:custGeom', 'a:prstGeom', 'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill', 'a:ln', 'a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst'],
  'a:xfrm': ['a:off', 'a:ext', 'a:chOff', 'a:chExt'],
  'a:prstGeom': ['a:avLst', 'a:gd'],
  'a:ln': ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:pattFill', 'a:prstDash', 'a:custDash', 'a:round', 'a:bevel', 'a:miter', 'a:headEnd', 'a:tailEnd'],
  'p:txBody': ['a:bodyPr', 'a:lstStyle', 'a:p'],
  'a:p': ['a:pPr', 'a:r', 'a:br', 'a:fld', 'a:endParaRPr'],
  'a:pPr': ['a:lnSpc', 'a:spcBef', 'a:spcAft', 'a:buClrTx', 'a:buClr', 'a:buSzTx', 'a:buSzPct', 'a:buSzPts', 'a:buFontTx', 'a:buFont', 'a:buNone', 'a:buAutoNum', 'a:buChar', 'a:tabLst', 'a:defRPr', 'a:extLst'],
  'a:r': ['a:rPr', 'a:t'],
  'a:rPr': ['a:ln', 'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill', 'a:effectLst', 'a:effectDag', 'a:highlight', 'a:uLnTx', 'a:uLn', 'a:uFillTx', 'a:uFill', 'a:latin', 'a:ea', 'a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst'],
  'a:bodyPr': ['a:noAutofit', 'a:normAutofit', 'a:spAutoFit', 'a:scene3d', 'a:sp3d', 'a:flatTx', 'a:extLst'],
  'p:pic': ['p:nvPicPr', 'p:blipFill', 'p:spPr', 'p:style'],
  'p:nvPicPr': ['p:cNvPr', 'p:cNvPicPr', 'p:nvPr'],
  'p:cNvPicPr': ['a:picLocks', 'a:extLst'],
  'p:blipFill': ['a:blip', 'a:srcRect', 'a:tile', 'a:stretch'],
  'p:sldMaster': ['p:cSld', 'p:clrMap', 'p:sldLayoutIdLst', 'p:transition', 'p:timing', 'p:hf', 'p:txStyles', 'p:extLst'],
  'p:sldLayout': ['p:cSld', 'p:clrMapOvr', 'p:transition', 'p:timing', 'p:hf', 'p:extLst'],
  'p:presentation': ['p:sldMasterIdLst', 'p:notesMasterIdLst', 'p:handoutMasterIdLst', 'p:sldIdLst', 'p:sldSz', 'p:notesSz', 'p:smartTags', 'p:embeddedFontLst', 'p:custShowLst', 'p:photoAlbum', 'p:custDataLst', 'p:kinsoku', 'p:defaultTextStyle', 'p:modifyVerifier', 'p:extLst'],
  'p:presentationPr': ['p:showPr', 'p:clrMru', 'p:extLst'],
  'p:showPr': ['p:present', 'p:browse', 'p:kiosk', 'p:sldAll', 'p:sldRg', 'p:custShow', 'p:penClr', 'p:extLst'],
  'p:txStyles': ['p:titleStyle', 'p:bodyStyle', 'p:otherStyle', 'p:extLst'],
  'a:tblStyleLst': []
};

function validate(xml, name) {
  const errors = [];
  const stack = [];                 // 只放标签名
  const lastIdx = [];               // 与栈同深度：该父元素最近一次子元素索引
  const re = /<(\/?)([A-Za-z_][\w:.-]*)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/g;
  let m;
  while ((m = re.exec(xml))) {
    const closing = m[1] === '/', tag = m[2], self = m[4] === '/';
    if (closing) { stack.pop(); lastIdx.pop(); continue; }

    const parent = stack[stack.length - 1];
    if (parent) {
      const spec = ORDER[parent];
      if (spec) {
        let idx = -1;
        for (let i = 0; i < spec.length; i++) {
          if (spec[i] === tag) { idx = i; break; }
        }
        if (idx < 0) {
          errors.push(`${name}: <${parent}> 下出现 schema 未列出的 <${tag}>`);
        } else {
          const prev = lastIdx[stack.length - 1];
          if (prev !== undefined && idx < prev) {
            errors.push(`${name}: <${parent}> 子元素顺序错误 —— <${tag}> 出现在索引更靠后的元素之后`);
          }
          lastIdx[stack.length - 1] = idx;
        }
      } else {
        lastIdx[stack.length - 1] = undefined;
      }
    }
    if (!self) {
      stack.push(tag);
      lastIdx.push(undefined);
    }
  }
  return errors;
}

const targets = [...parts.keys()].filter(k => k.endsWith('.xml') && !k.startsWith('docProps'));
let total = 0;
const byPart = new Map();
for (const t of targets) {
  const errs = validate(parts.get(t), t);
  if (errs.length) { byPart.set(t, errs); total += errs.length; }
}
for (const [t, errs] of byPart) {
  console.log(`\n--- ${t} （${errs.length} 处）---`);
  [...new Set(errs)].slice(0, 8).forEach(x => console.log('  ❌ ' + x));
}
console.log('');
console.log('检查部件数:', targets.length);
console.log(total ? `❌ 发现 ${total} 处结构问题` : '✅ 未发现子元素顺序问题');
