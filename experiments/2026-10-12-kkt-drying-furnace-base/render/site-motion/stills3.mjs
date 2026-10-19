import { chromium } from '/home/claude/personal-task/experiments/2026-10-12-kkt-drying-furnace-base/render/node_modules/playwright-core/index.mjs';
import fs from 'fs';
const [out,view,list]=process.argv.slice(2);fs.mkdirSync(out,{recursive:true});
const b=await chromium.launch({executablePath:'/opt/pw-browsers/chromium_headless_shell-1194/chrome-linux/headless_shell',args:['--use-gl=angle','--use-angle=swiftshade','--enable-unsafe-swiftshader']});
const p=await b.newPage({viewport:{width:1920,height:1080}});p.on('pageerror',e=>console.log('ERR',e.message));p.on('console',m=>{if(m.type()==='error')console.log('CON',m.text())});
await p.goto('http://127.0.0.1:8766/site-motion3.html?view='+view);await p.waitForFunction(()=>window.ready,null,{timeout:120000});
for(const t of list.split(',').map(Number)){const d=await p.evaluate(t=>{const s=window.renderFrame(t);return document.getElementById('film').toDataURL('image/jpeg',.85).split(',')[1]},t);fs.writeFileSync(`${out}/${view}_${String(t).padStart(5,'0')}.jpg`,Buffer.from(d,'base64'))}
await b.close();
