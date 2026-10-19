import { chromium } from '/home/claude/personal-task/experiments/2026-10-12-kkt-drying-furnace-base/render/node_modules/playwright-core/index.mjs';
import fs from 'fs';
const [out,mode,query]=process.argv.slice(2);fs.mkdirSync(out,{recursive:true});
const b=await chromium.launch({executablePath:'/opt/pw-browsers/chromium_headless_shell-1194/chrome-linux/headless_shell',args:['--use-gl=angle','--use-angle=swiftshade','--enable-unsafe-swiftshader']});
const p=await b.newPage({viewport:{width:1920,height:1080}});p.on('pageerror',e=>console.log('ERR',e.message));
await p.goto('http://127.0.0.1:8766/site-motion4.html?'+query);await p.waitForFunction(()=>window.ready,null,{timeout:120000});
const total=await p.evaluate(()=>window.TOTAL);
let times;if(mode.startsWith('list:'))times=mode.slice(5).split(',').map(Number);else{const fps=+mode;times=[];for(let i=0;i<total*fps;i++)times.push(i/fps)}
const S=+(process.env.START||0),E=+(process.env.END||1e9),t0=Date.now();
for(let i=S;i<Math.min(times.length,E);i++){const d=await p.evaluate(t=>{window.renderFrame(t);return document.getElementById('film').toDataURL('image/jpeg',.93).split(',')[1]},times[i]);fs.writeFileSync(`${out}/f_${String(i).padStart(5,'0')}.jpg`,Buffer.from(d,'base64'));if(i%50==0)console.log(i,times.length,((Date.now()-t0)/1000).toFixed(0))}
console.log('done',times.length,((Date.now()-t0)/1000).toFixed(0));await b.close();
