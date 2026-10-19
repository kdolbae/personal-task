import { chromium } from '/home/claude/personal-task/experiments/2026-10-12-kkt-drying-furnace-base/render/node_modules/playwright-core/index.mjs';
const b=await chromium.launch({executablePath:'/opt/pw-browsers/chromium_headless_shell-1194/chrome-linux/headless_shell'});
const p=await b.newPage();
await p.goto('file://'+process.cwd()+'/plan.html',{waitUntil:'load'});
await p.pdf({path:process.argv[2],format:'A4',printBackground:true,displayHeaderFooter:true,
 headerTemplate:'<div style="font-size:8px;color:#5D6878;width:100%;padding:0 16mm;font-family:Pretendard,sans-serif;display:flex;justify-content:space-between"><span>금강변압기 진공건조로 회전·안착·용접·이동 작업계획서 · NM-2610-01</span><span>엔엠솔루션</span></div>',
 footerTemplate:'<div style="font-size:8px;color:#5D6878;width:100%;padding:0 16mm;font-family:Pretendard,sans-serif;display:flex;justify-content:space-between"><span>엔엠솔루션  |  경기도 평택시 고덕동 도시지원로 121 지식공장아이타워 906호  |  대표전화 1800-5901  |  직통 010-2369-3691  |  nm-solution.co.kr</span><span><span class="pageNumber"></span> / <span class="totalPages"></span></span></div>',
 margin:{top:'16mm',bottom:'16mm',left:'16mm',right:'16mm'}});
await b.close();
