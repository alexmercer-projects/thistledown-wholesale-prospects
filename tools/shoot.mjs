// Screenshots and an overflow check with headless Chrome over the DevTools protocol (Node 22, no packages).
// Usage: node tools/shoot.mjs <url> <outdir> [prefix]
import { spawn } from 'node:child_process';
import { writeFileSync, mkdirSync } from 'node:fs';
const [url = 'http://localhost:8765/', out = 'screenshots', prefix = ''] = process.argv.slice(2);
mkdirSync(out, { recursive: true });
const port = 9333 + Math.floor(Math.random() * 500);
const chrome = spawn('google-chrome', ['--headless=new', '--no-sandbox', '--disable-gpu', '--hide-scrollbars', `--remote-debugging-port=${port}`, 'about:blank'], { stdio: 'ignore' });
const sleep = ms => new Promise(r => setTimeout(r, ms));
let ws;
for (let i = 0; i < 50 && !ws; i++) { try { const t = await (await fetch(`http://127.0.0.1:${port}/json`)).json(); const p = t.find(x => x.type === 'page'); if (p) ws = p.webSocketDebuggerUrl; } catch {} await sleep(150); }
const sock = new WebSocket(ws); await new Promise(r => sock.onopen = r);
let id = 0; const pend = new Map();
sock.onmessage = m => { const d = JSON.parse(m.data); if (d.id && pend.has(d.id)) { pend.get(d.id)(d.result || d.error); pend.delete(d.id); } };
const send = (method, params = {}) => new Promise(r => { const i = ++id; pend.set(i, r); sock.send(JSON.stringify({ id: i, method, params })); });
const ev = async e => (await send('Runtime.evaluate', { expression: e, returnByValue: true })).result.value;
await send('Page.enable');
for (const [name, w, h, mobile] of [['desktop-1440', 1440, 900, false], ['mobile-390', 390, 844, true]]) {
  await send('Emulation.setDeviceMetricsOverride', { width: w, height: h, deviceScaleFactor: mobile ? 2 : 1, mobile });
  await send('Page.navigate', { url: url + (url.includes('?') ? '&' : '?') + 'static' });
  await sleep(2500);
  const m = await ev(`({sw: document.documentElement.scrollWidth, cw: document.documentElement.clientWidth, h: document.documentElement.scrollHeight,
    tops: Object.fromEntries(['list','scoring','checks','handoff'].map(i => [i, Math.round(document.getElementById(i).getBoundingClientRect().top + scrollY)])),
    wide: [...document.querySelectorAll('body *')].filter(e => e.getBoundingClientRect().right > document.documentElement.clientWidth + 1 && getComputedStyle(e).position !== 'fixed' && !e.closest('.sheet__tabs')).slice(0,5).map(e => e.tagName + '.' + e.className)})`);
  console.log(name, JSON.stringify(m));
  const shot = async (file, full) => {
    const p = full ? { captureBeyondViewport: true, clip: { x: 0, y: 0, width: w, height: m.h, scale: 1 } } : {};
    const r = await send('Page.captureScreenshot', { format: 'png', ...p });
    writeFileSync(`${out}/${prefix}${file}.png`, Buffer.from(r.data, 'base64'));
  };
  await shot(name, false); await shot(name + '-full', true);
}
sock.close(); chrome.kill();
