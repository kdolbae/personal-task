import { chromium } from '/home/claude/personal-task/experiments/2026-10-12-kkt-drying-furnace-base/render/node_modules/playwright-core/index.mjs';
const b=await chromium.launch({executablePath:'/opt/pw-browsers/chromium_headless_shell-1194/chrome-linux/headless_shell'});
const p=await b.newPage();
await p.goto('file://'+process.cwd()+'/plan.html',{waitUntil:'load'});
await p.pdf({path:process.argv[2],format:'A4',printBackground:true,displayHeaderFooter:true,
 headerTemplate:'<div style="font-size:7pt;color:#666;width:100%;padding:0 16mm;font-family:sans-serif">진공건조로 회전·안착 및 접합 용접 작업계획서 · 엔엠솔루션</div>',
 footerTemplate:'<div style="font-size:7pt;color:#666;width:100%;text-align:center;font-family:sans-serif">- <span class="pageNumber"></span> / <span class="totalPages"></span> -</div>',
 margin:{top:'16mm',bottom:'16mm',left:'16mm',right:'16mm'}});
await b.close();
