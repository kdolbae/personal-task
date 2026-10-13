import { chromium } from '/home/claude/personal-task/experiments/2026-10-12-kkt-drying-furnace-base/render/node_modules/playwright-core/index.mjs';
const b=await chromium.launch({executablePath:'/opt/pw-browsers/chromium_headless_shell-1194/chrome-linux/headless_shell'});
const p=await b.newPage();
await p.goto('file://'+process.cwd()+'/cover.html',{waitUntil:'load'});
await p.pdf({path:process.argv[2],format:'A4',printBackground:true,preferCSSPageSize:true,margin:{top:0,bottom:0,left:0,right:0}});
await b.close();
