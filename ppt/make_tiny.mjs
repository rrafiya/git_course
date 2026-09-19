// 生成一个"绝对最简"的 pptx（1 页纯文字，无主题、无扩展部件、无自动缩放）
// 用途：如果连它也提示修复，说明问题在我的基础包结构；否则问题在某个附加特性
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';

const OUT = process.argv[2] || 'ppt/_probe/Z_tiny.pptx';
const EMU = 914400;
const inch = (v) => Math.round(v * EMU);
const HDR = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

const emptyTree = `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>`;

const slide = HDR + `<p:sld ${NS}><p:cSld><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox 1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="${inch(1)}" y="${inch(1)}"/><a:ext cx="${inch(8)}" cy="${inch(1.5)}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
  `<p:txBody><a:bodyPr wrap="square" lIns="0" rIns="0" tIns="0" bIns="0" anchor="t"><a:noAutofit/></a:bodyPr><a:lstStyle/>` +
  `<a:p><a:pPr algn="l"><a:buNone/></a:pPr>` +
  `<a:r><a:rPr lang="en-US" sz="2400" b="1" dirty="0"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill>` +
  `<a:latin typeface="Arial"/><a:ea typeface="Microsoft YaHei"/></a:rPr><a:t>Hello Git</a:t></a:r></a:p>` +
  `</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;

const layout = HDR + `<p:sldLayout ${NS} type="blank" preserve="1"><p:cSld name="Blank">` +
  emptyTree + `</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

const master = HDR + `<p:sldMaster ${NS}><p:cSld><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
  `</p:spTree></p:cSld>` +
  `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
  `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>` +
  `</p:sldMaster>`;

const masterRels = HDR + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
  `</Relationships>`;

const layoutRels = HDR + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>` +
  `</Relationships>`;

const presentation = HDR + `<p:presentation ${NS}>` +
  `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
  `<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>` +
  `<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>` +
  `</p:presentation>`;

const presRels = HDR + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
  `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
  `</Relationships>`;

const slideRels = HDR + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
  `</Relationships>`;

const rootRels = HDR + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
  `</Relationships>`;

const contentTypes = HDR + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
  `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
  `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
  `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
  `</Types>`;

// ---- ZIP ----
function zip(files) {
  const chunks = [], central = [];
  let offset = 0;
  const t = new Int32Array(256);
  for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c; }
  const crc32 = (buf) => { let c = -1; for (let i = 0; i < buf.length; i++) c = t[(c ^ buf[i]) & 0xFF] ^ (c >>> 8); return (c ^ -1) >>> 0; };
  for (const f of files) {
    const nb = Buffer.from(f.name, 'utf8');
    const data = Buffer.from(f.data, 'utf8');
    const comp = zlib.deflateRawSync(data, { level: 6 });
    const useStore = comp.length >= data.length;
    const body = useStore ? data : comp;
    const method = useStore ? 0 : 8;
    const crc = crc32(data);
    const local = Buffer.alloc(30 + nb.length);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(method, 8); local.writeUInt16LE(0, 10); local.writeUInt16LE(0, 12);
    local.writeUInt32LE(crc, 14); local.writeUInt32LE(body.length, 18); local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nb.length, 26); local.writeUInt16LE(0, 28);
    nb.copy(local, 30);
    chunks.push(local, body);
    const cd = Buffer.alloc(46 + nb.length);
    cd.writeUInt32LE(0x02014b50, 0); cd.writeUInt16LE(20, 4); cd.writeUInt16LE(20, 6);
    cd.writeUInt16LE(0x0800, 8); cd.writeUInt16LE(method, 10); cd.writeUInt16LE(0, 12); cd.writeUInt16LE(0, 14);
    cd.writeUInt32LE(crc, 16); cd.writeUInt32LE(body.length, 20); cd.writeUInt32LE(data.length, 24);
    cd.writeUInt16LE(nb.length, 28); cd.writeUInt16LE(0, 30); cd.writeUInt16LE(0, 32);
    cd.writeUInt16LE(0, 34); cd.writeUInt16LE(0, 36); cd.writeUInt32LE(0, 38); cd.writeUInt32LE(offset, 42);
    nb.copy(cd, 46);
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
  { name: '[Content_Types].xml', data: contentTypes },
  { name: '_rels/.rels', data: rootRels },
  { name: 'ppt/presentation.xml', data: presentation },
  { name: 'ppt/_rels/presentation.xml.rels', data: presRels },
  { name: 'ppt/slideMasters/slideMaster1.xml', data: master },
  { name: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', data: masterRels },
  { name: 'ppt/slideLayouts/slideLayout1.xml', data: layout },
  { name: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', data: layoutRels },
  { name: 'ppt/slides/slide1.xml', data: slide },
  { name: 'ppt/slides/_rels/slide1.xml.rels', data: slideRels }
];
fs.mkdirSync(path.dirname(path.resolve(OUT)), { recursive: true });
fs.writeFileSync(OUT, zip(files));
console.log('已生成最简文件:', OUT, ((fs.statSync(OUT).size) / 1024).toFixed(1), 'KB');
console.log('包含部件:', files.length, '个（无主题、无扩展部件、无图片）');
