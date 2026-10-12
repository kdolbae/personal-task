import { chromium } from 'playwright-core';
const jobs = JSON.parse(process.argv[2]);   // [[view, units, outfile], ...]
const browser = await chromium.launch({
  executablePath: '/opt/pw-browsers/chromium_headless_shell-1194/chrome-linux/headless_shell',
  args: ['--use-gl=angle', '--use-angle=swiftshader', '--enable-unsafe-swiftshader', '--ignore-gpu-blocklist', '--no-sandbox'],
});
for (const [view, units, out] of jobs) {
  const page = await browser.newPage({ viewport: { width: 1920, height: 1080 } });
  page.on('console', m => console.log('console:', m.text()));
  page.on('pageerror', e => console.log('pageerror:', e.message));
  await page.goto(`http://localhost:8123/scene.html?view=${view}&units=${units}`);
  await page.waitForFunction('window.__ready === true', null, { timeout: 100000 });
  await page.screenshot({ path: out });
  console.log('saved', out);
  await page.close();
}
await browser.close();
