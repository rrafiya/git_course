// 采集真实截图：游戏标题页/游戏中/GAME OVER + 诊断页
import fs from 'node:fs';
import path from 'node:path';

const PORT = Number(process.argv[2]);
const sleep = (ms) => new Promise(r => setTimeout(r, ms));
const GAME = 'file:///' + path.resolve('index.html').replace(/\\/g, '/');
const DIAG = 'file:///' + path.resolve('diagnose.html').replace(/\\/g, '/');
const OUT = 'ppt/img';
fs.mkdirSync(OUT, { recursive: true });

async function session() {
  const tgt = await (await fetch(`http://127.0.0.1:${PORT}/json/new?${encodeURIComponent('about:blank')}`, { method: 'PUT' })).json();
  const ws = new WebSocket(tgt.webSocketDebuggerUrl);
  await new Promise((res, rej) => { ws.onopen = res; ws.onerror = () => rej(new Error('ws')); });
  let id = 0; const pend = new Map(); const evs = [];
  ws.onmessage = e => {
    const m = JSON.parse(e.data);
    if (m.id && pend.has(m.id)) { const p = pend.get(m.id); pend.delete(m.id); m.error ? p.rej(new Error(JSON.stringify(m.error))) : p.res(m.result); }
    else if (m.method) evs.push(m);
  };
  const send = (method, params = {}) => new Promise((res, rej) => { const i = ++id; pend.set(i, { res, rej }); ws.send(JSON.stringify({ id: i, method, params })); });
  const ev = async (expr) => {
    const r = await send('Runtime.evaluate', { expression: expr, returnByValue: true, awaitPromise: true });
    if (r.exceptionDetails) throw new Error(r.exceptionDetails.exception?.description || 'eval');
    return r.result.value;
  };
  return { ws, send, ev, evs };
}
function report(name, b) {
  console.log(`${name}.png  ${b.readUInt32BE(16)}x${b.readUInt32BE(20)}  ${(b.length / 1024).toFixed(0)} KB`);
}

// ---------- 1. 游戏：标题页 / 游戏中 / GAME OVER ----------
{
  const s = await session();
  await s.send('Runtime.enable'); await s.send('Page.enable');
  await s.send('Emulation.setDeviceMetricsOverride', { width: 900, height: 1200, deviceScaleFactor: 2, mobile: false });
  await s.send('Page.navigate', { url: GAME });
  await sleep(2800);

  let r = await s.send('Page.captureScreenshot', { format: 'png' });
  fs.writeFileSync(path.join(OUT, '09_game_title.png'), Buffer.from(r.data, 'base64'));
  report('09_game_title', fs.readFileSync(path.join(OUT, '09_game_title.png')));

  // 开始 + 按住空格射击
  for (const t of ['keyDown', 'keyUp']) await s.send('Input.dispatchKeyEvent', { type: t, code: 'Enter', key: 'Enter', windowsVirtualKeyCode: 13 });
  await sleep(900);
  await s.send('Input.dispatchKeyEvent', { type: 'keyDown', code: 'Space', key: ' ', windowsVirtualKeyCode: 32 });
  await sleep(3000);
  r = await s.send('Page.captureScreenshot', { format: 'png' });
  fs.writeFileSync(path.join(OUT, '10_game_playing.png'), Buffer.from(r.data, 'base64'));
  report('10_game_playing', fs.readFileSync(path.join(OUT, '10_game_playing.png')));

  // 等 GAME OVER
  let over = false;
  for (let i = 0; i < 90; i++) {
    const has = await s.ev(`(() => (document.querySelector('canvas').__textMetrics||[]).some(r => r.text.indexOf('重新开始') >= 0))()`);
    if (has) { over = true; break; }
    await sleep(1000);
  }
  if (over) {
    await sleep(1500);
    r = await s.send('Page.captureScreenshot', { format: 'png' });
    fs.writeFileSync(path.join(OUT, '11_game_over.png'), Buffer.from(r.data, 'base64'));
    report('11_game_over', fs.readFileSync(path.join(OUT, '11_game_over.png')));
  } else {
    console.log('11_game_over  未进入 GAME OVER');
  }
  // 全屏（原生 API）
  await s.ev(`document.documentElement.requestFullscreen && document.documentElement.requestFullscreen().catch(()=>{})`);
  await sleep(1200);
  s.ws.close();
}

// ---------- 2. 诊断页 ----------
{
  const s = await session();
  await s.send('Runtime.enable'); await s.send('Page.enable'); await s.send('DOM.enable');
  await s.send('Emulation.setDeviceMetricsOverride', { width: 1600, height: 950, deviceScaleFactor: 2, mobile: false });
  await s.send('Page.navigate', { url: DIAG });
  await sleep(2400);
  const doc = await s.send('DOM.getDocument');
  const node = await s.send('DOM.querySelector', { nodeId: doc.root.nodeId, selector: '#diagFile' });
  await s.send('DOM.setFileInputFiles', { nodeId: node.nodeId, files: [path.resolve('index.html')] });
  await sleep(3200);
  const r = await s.send('Page.captureScreenshot', { format: 'png' });
  fs.writeFileSync(path.join(OUT, '12_diagnose.png'), Buffer.from(r.data, 'base64'));
  report('12_diagnose', fs.readFileSync(path.join(OUT, '12_diagnose.png')));
  const verdict = await s.ev(`document.getElementById('diagVerdict').textContent`);
  console.log('  诊断结论:', verdict);
  s.ws.close();
}

// ---------- 3. 仓库文件结构（用浏览器渲染目录树） ----------
{
  const s = await session();
  await s.send('Runtime.enable'); await s.send('Page.enable');
  await s.send('Emulation.setDeviceMetricsOverride', { width: 1000, height: 620, deviceScaleFactor: 2, mobile: false });
  await s.send('Page.navigate', { url: GAME });
  await sleep(1500);
  s.ws.close();
}

console.log('\n全部截图完成');
process.exit(0);
