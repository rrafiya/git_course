// 生成 HTML 预览页（与 pptx 同一套内容与版式），便于在浏览器直接查看
import fs from 'node:fs';
import path from 'node:path';

const ROOT = process.cwd();
const IMGDIR = path.join(ROOT, 'ppt', 'img');
const deck = JSON.parse(fs.readFileSync(path.join(ROOT, 'ppt', 'deck.json'), 'utf8'));
const OUT = path.join(ROOT, 'ppt', '预览.html');

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

// 行内标记 -> HTML
function rich(txt) {
  return esc(txt)
    .replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>')
    .replace(/`([^`]+)`/g, '<code>$1</code>');
}

function pngSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}

const slides = deck.slides.map((sd, i) => {
  const num = i + 1;
  if (sd.layout === 'cover') {
    const lines = sd.lines.map(l => `<div class="cl">${esc(l.replace(/^N\|/, ''))}</div>`).join('');
    return `<section class="slide cover">
      <div class="g1"></div><div class="g2"></div>
      <div class="accent"></div>
      <h1>${esc(sd.title)}</h1>
      <p class="sub">${esc(sd.subtitle)}</p>
      <div class="clines">${lines}</div>
    </section>`;
  }
  if (sd.layout === 'toc') {
    const items = sd.lines.map(l => `<li>${esc(l.replace(/^B\|/, ''))}</li>`).join('');
    return `<section class="slide">
      <div class="accent small"></div><h2>${esc(sd.title)}</h2>
      <div class="tocCard"><ul>${items}</ul></div>
    </section>`;
  }
  if (sd.layout === 'image') {
    const img = `img/${sd.image}`;
    return `<section class="slide">
      <div class="accent small"></div><h2>${esc(sd.title)}</h2>
      <div class="imgFull"><img src="${img}" alt=""></div>
      ${sd.caption ? `<div class="cap">${esc(sd.caption)}</div>` : ''}
    </section>`;
  }
  // content
  const withImg = sd.image && sd.image !== '';
  const body = sd.lines.map(ln => {
    if (!ln) return '<div class="gap"></div>';
    const kind = ln.slice(0, 2), txt = ln.slice(2);
    if (kind === 'H|') return `<div class="h">${rich(txt)}</div>`;
    if (kind === 'C|') return `<div class="code">${esc(txt)}</div>`;
    if (kind === 'Q|') return `<div class="quote">${rich(txt)}</div>`;
    if (kind === 'N|') return `<div class="note">${rich(txt)}</div>`;
    return `<div class="li"><span class="dot">◆</span><span>${rich(txt)}</span></div>`;
  }).join('');
  return `<section class="slide">
    <div class="accent small"></div><h2>${esc(sd.title)}</h2>
    <div class="cols${withImg ? ' two' : ''}">
      <div class="text">${body}</div>
      ${withImg ? `<div class="imgs"><img src="img/${sd.image}" alt=""></div>` : ''}
    </div>
  </section>`;
});

const html = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>${esc(deck.meta.title)} — 预览</title>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { background:#05070e; font-family:"Microsoft YaHei","Segoe UI",sans-serif;
    padding:26px 0 60px; }
  .hintbar { max-width:1120px; margin:0 auto 18px; color:#7d8bab; font-size:13px;
    display:flex; align-items:center; gap:14px; flex-wrap:wrap; }
  .hintbar b { color:#9fe8ff; }
  .hintbar .pill { background:#152039; border:1px solid #2a3a5c; border-radius:20px;
    padding:3px 12px; color:#9db4de; }
  .slide { position:relative; width:1120px; height:630px; margin:0 auto 26px;
    background:#0B1020; border-radius:10px; overflow:hidden;
    box-shadow:0 10px 34px rgba(0,0,0,.6), 0 0 0 1px #1b2536; }
  .slide .pgnum { position:absolute; right:26px; bottom:16px; color:#6C7A99;
    font-family:Consolas,monospace; font-size:13px; }
  h1 { position:absolute; left:84px; top:188px; font-size:44px; color:#9FE8FF; }
  h2 { position:absolute; left:80px; top:60px; font-size:29px; color:#9FE8FF; }
  .accent { position:absolute; left:84px; top:168px; width:126px; height:6px;
    background:#4DD4FF; border-radius:3px; }
  .accent.small { left:80px; top:48px; width:76px; height:6px; }
  .cover .sub { position:absolute; left:84px; top:298px; font-size:21px; color:#C4D4F0; }
  .cover .clines { position:absolute; left:84px; top:382px; }
  .cover .cl { color:#93A3C4; font-size:15px; line-height:1.9; }
  .cover .g1 { position:absolute; left:-170px; top:-140px; width:720px; height:940px;
    background:#12325A; opacity:.45; border-radius:80px; transform:rotate(18deg); }
  .cover .g2 { position:absolute; right:-200px; top:180px; width:700px; height:760px;
    background:#3A1A2C; opacity:.55; border-radius:80px; transform:rotate(-14deg); }
  .tocCard { position:absolute; left:80px; top:168px; width:960px; height:394px;
    background:#111A2C; border:1px solid #23304A; border-radius:12px; padding:26px 34px; }
  .tocCard ul { list-style:none; }
  .tocCard li { color:#DCE6FA; font-size:20px; line-height:1.5; padding:9px 0 9px 26px;
    position:relative; }
  .tocCard li:before { content:"◆"; position:absolute; left:0; color:#4DD4FF; font-size:11px; top:14px; }
  .imgFull { position:absolute; left:52px; top:141px; right:52px; height:398px;
    display:flex; align-items:center; justify-content:center; }
  .imgFull img { max-width:100%; max-height:100%; border-radius:6px; }
  .cap { position:absolute; left:80px; bottom:44px; color:#6C7A99; font-size:14px; }
  .cols { position:absolute; left:80px; top:150px; right:80px; bottom:44px; }
  .cols.two { display:flex; gap:30px; }
  .cols.two .text { flex:1; min-width:0; }
  .cols.two .imgs { width:436px; flex:0 0 436px; display:flex; align-items:center; justify-content:center; }
  .cols.two .imgs img { max-width:100%; max-height:436px; border-radius:6px; }
  .h { color:#FFD166; font-size:18px; font-weight:700; margin:8px 0 8px; }
  .code { background:#060A14; color:#7FE9FF; font-family:Consolas,monospace;
    font-size:13px; padding:6px 12px; border-radius:5px; margin:4px 0; }
  .quote { border-left:4px solid #63F5A0; padding-left:14px; color:#63F5A0;
    font-size:15px; line-height:1.55; margin:6px 0; }
  .note { color:#93A3C4; font-size:14px; line-height:1.6; padding-left:18px; margin:3px 0; }
  .li { display:flex; gap:10px; color:#DCE6FA; font-size:15.5px; line-height:1.5; margin:7px 0; }
  .li .dot { color:#4DD4FF; font-size:10px; padding-top:7px; }
  .li b { color:#FFD166; }
  .li code, .note code { background:#0a0f1c; color:#7FE9FF; font-family:Consolas,monospace;
    font-size:13px; padding:1px 6px; border-radius:4px; }
  .gap { height:8px; }
</style>
</head>
<body>
<div class="hintbar">
  <b>${esc(deck.meta.title)}</b>
  <span class="pill">共 ${deck.slides.length} 页</span>
  <span class="pill">16:9 · 1120×630</span>
  <span>此页用于预览版式；正式文件为 <b>${esc(deck.meta.out)}</b>（用 PowerPoint 打开）</span>
</div>
${slides.map((s, i) => s.replace('</section>', `<div class="pgnum">${i + 1}</div></section>`)).join('\n')}
</body>
</html>`;

fs.writeFileSync(OUT, html, 'utf8');
console.log('slides:', deck.slides.length);
console.log('SAVED :', OUT);
console.log('size  :', (fs.statSync(OUT).size / 1024).toFixed(0), 'KB');
