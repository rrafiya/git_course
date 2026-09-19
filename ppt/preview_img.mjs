// 生成图片版预览页（引用 slides/*.png，所见即所得）
import fs from 'node:fs';
import path from 'node:path';

const ROOT = process.cwd();
const DIR = path.join(ROOT, 'ppt', 'slides');
const deck = JSON.parse(fs.readFileSync(path.join(ROOT, 'ppt', 'deck.json'), 'utf8'));
const OUT = path.join(ROOT, 'ppt', '预览-图片版.html');

const files = fs.readdirSync(DIR).filter(f => /^slide\d+\.png$/.test(f))
  .sort((a, b) => (+a.match(/\d+/)[0]) - (+b.match(/\d+/)[0]));

const titles = deck.slides.map(s => s.title || deck.meta.subtitle);

const blocks = files.map((f, i) => {
  const n = i + 1;
  return `<figure class="pg">
  <figcaption><span class="no">${n}</span>${titles[i] || ''}</figcaption>
  <img src="slides/${f}" alt="第 ${n} 页">
</figure>`;
}).join('\n');

const html = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>${deck.meta.title} — 图片版预览</title>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { background:#05070e; color:#c8d6f5; padding:28px 20px 70px;
    font-family:"Microsoft YaHei","Segoe UI",system-ui,sans-serif; }
  header { max-width:1120px; margin:0 auto 26px; }
  header h1 { font-size:24px; color:#9fe8ff; margin-bottom:6px; }
  header p { font-size:13px; color:#6c7a99; line-height:1.7; }
  header .pills { margin-top:10px; display:flex; gap:10px; flex-wrap:wrap; }
  header .pill { background:#152039; border:1px solid #2a3a5c; border-radius:20px;
    padding:4px 12px; font-size:12px; color:#9db4de; }
  .pg { max-width:1120px; margin:0 auto 30px; }
  .pg figcaption { font-size:13px; color:#8ea2cc; padding:0 2px 8px; display:flex;
    align-items:center; gap:9px; }
  .pg .no { background:#1b2950; border:1px solid #33518f; color:#9fe8ff;
    border-radius:5px; padding:1px 8px; font-family:Consolas,monospace; font-size:12px; }
  .pg img { display:block; width:100%; height:auto; border-radius:10px;
    box-shadow:0 10px 30px rgba(0,0,0,.6), 0 0 0 1px #1b2536; background:#0B1020; }
  footer { max-width:1120px; margin:10px auto 0; font-size:12px; color:#5f7199;
    border-top:1px solid #1b2536; padding-top:14px; line-height:1.8; }
  footer code { background:#111a2c; padding:1px 6px; border-radius:4px; color:#9fe8ff;
    font-family:Consolas,monospace; }
</style>
</head>
<body>
<header>
  <h1>${deck.meta.title}</h1>
  <p>图片版预览 —— 每页都是按 PPT 版式渲染出的静态图片，所见即所得。<br>
     若这里的版式正常，说明 PPT 内容无误；PowerPoint 打开时的差异就出在播放器渲染上。</p>
  <div class="pills">
    <span class="pill">共 ${files.length} 页</span>
    <span class="pill">16:9</span>
    <span class="pill">正式文件：${deck.meta.out}</span>
  </div>
</header>
${blocks}
<footer>
  对比用：<code>预览.html</code> 是实时 HTML 版（可选中文字），
  本页是图片版（与 PPT 版式一致）。<br>
  仓库：${deck.meta.repo}
</footer>
</body>
</html>`;

fs.writeFileSync(OUT, html, 'utf8');
console.log('slides:', files.length);
console.log('SAVED :', OUT);
console.log('size  :', (fs.statSync(OUT).size / 1024).toFixed(0), 'KB');
