// node frames.mjs <plan A|B> <outdir> <fps|list:t1,t2,..> [w] [h]
import { chromium } from 'playwright-core';
import fs from 'fs';
const [plan, out, mode, w = 1280, h = 720] = process.argv.slice(2);
fs.mkdirSync(out, { recursive: true });
const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium_headless_shell-1194/chrome-linux/headless_shell', args: ['--use-gl=angle', '--use-angle=swiftshade', '--enable-unsafe-swiftshader'] });
const page = await browser.newPage({ viewport: { width: +w, height: +h } });
page.on('pageerror', e => console.log('PAGEERR', e.message));
await page.goto(`http://localhost:8123/process.html?plan=${plan}&w=${w}&h=${h}`);
await page.waitForFunction(() => window.__ready === true, null, { timeout: 120000 });
let times;
if (mode.startsWith('list:')) times = mode.slice(5).split(',').map(Number);
else { const total = await page.evaluate(() => window.renderAt(0)); const fps = +mode; times = []; for (let i = 0; i <= total * fps; i++) times.push(i / fps); }
const t0 = Date.now();
for (let i = 0; i < times.length; i++) {
  await page.evaluate(t => window.renderAt(t), times[i]);
  await page.screenshot({ path: `${out}/f_${String(i).padStart(5, '0')}.png` });
  if (i % 50 === 0) console.log(i, '/', times.length, ((Date.now() - t0) / 1000).toFixed(0) + 's');
}
await browser.close();
