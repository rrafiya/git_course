// 通用截图工具：把 HTML 文件渲染成 PNG（供 PPT 使用）
// 用法: node _shot.mjs <port> <outDir> <html1:name:WxH> [html2:name:WxH ...]
import fs from 'node:fs';
import path from 'node:path';

const PORT = Number(process.argv[2]);
const OUT = process.argv[3];
const specs = process.argv.slice(4);
fs.mkdirSync(OUT, { recursive: true });

const sleep = (ms) => new Promise(r => setTimeout(r, ms));

async function session() {
  const tgt = await (await fetch(`http://127.0.0.1:${PORT}/json/new?${encodeURIComponent('about:blank')}`, { method: 'PUT' })).json();
  const ws = new WebSocket(tgt.webSocketDebuggerUrl);
  await new Promise((res, rej) => { ws.onopen = res; ws.onerror = () => rej(new Error('ws')); });
  let id = 0; const pend = new Map();
  ws.onmessage = e => {
    const m = JSON.parse(e.data);
    if (m.id && pend.has(m.id)) { const p = pend.get(m.id); pend.delete(m.id); m.error ? p.rej(new Error(JSON.stringify(m.error))) : p.res(m.result); }
  };
  const send = (method, params = {}) => new Promise((res, rej) => { const i = ++id; pend.set(i, { res, rej }); ws.send(JSON.stringify({ id: i, method, params })); });
  return { ws, send };
}

for (const spec of specs) {
  const parts = spec.split('|');
  const file = parts[0];
  const name = parts[1];
  const dim = (parts[2] || '1280x720').split('x').map(Number);
  const dpr = parts[3] ? Number(parts[3]) : 2;

  const s = await session();
  await s.send('Runtime.enable'); await s.send('Page.enable');
  await s.send('Emulation.setDeviceMetricsOverride', { width: dim[0], height: dim[1], deviceScaleFactor: dpr, mobile: false });
  await s.send('Page.navigate', { url: 'file:///' + path.resolve(file).replace(/\\/g, '/') });
  await sleep(1800);
  const r = await s.send('Page.captureScreenshot', { format: 'png', captureBeyondViewport: false });
  const out = path.join(OUT, name + '.png');
  fs.writeFileSync(out, Buffer.from(r.data, 'base64'));
  const b = fs.readFileSync(out);
  console.log(`${name}.png  ${b.readUInt32BE(16)}x${b.readUInt32BE(20)}  ${(b.length / 1024).toFixed(0)} KB`);
  s.ws.close();
}
process.exit(0);
