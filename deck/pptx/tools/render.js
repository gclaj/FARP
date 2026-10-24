const { chromium } = require('playwright');const http=require('http'),fs=require('fs'),path=require('path');
const [slide,startS,endS,fps,port]=[+process.argv[2],+process.argv[3],+process.argv[4],+process.argv[5],+process.argv[6]];
const out=`vid/s${String(slide).padStart(2,'0')}`;fs.mkdirSync(out,{recursive:true});
const root=__dirname+'/srv';const srv=http.createServer((q,r)=>{let p=decodeURIComponent(q.url.split('?')[0]);if(p==='/')p='/index.html';const f=path.join(root,p);if(!fs.existsSync(f)){r.writeHead(404);return r.end()}r.writeHead(200,{'Content-Type':p.endsWith('.html')?'text/html':p.endsWith('.js')?'text/javascript':'application/octet-stream'});fs.createReadStream(f).pipe(r)}).listen(port);
(async()=>{const b=await chromium.launch({executablePath:'/opt/pw-browsers/chromium',args:['--no-sandbox','--use-gl=angle','--use-angle=swiftshader','--enable-unsafe-swiftshader']});
const p=await b.newPage({viewport:{width:1600,height:900}});
await p.addInitScript(()=>{window.__VT=0;const nat=window.requestAnimationFrame.bind(window);performance.now=()=>window.__VT;window.requestAnimationFrame=cb=>nat(()=>cb(window.__VT));});
await p.goto('http://127.0.0.1:'+port+'/');await p.waitForTimeout(9000);
await p.addStyleTag({content:'#bar,#lbt,#lbb,#grain,#cap{display:none!important}#wrap{bottom:0!important}'});await p.evaluate(()=>dispatchEvent(new Event('resize')));
await p.evaluate(k=>{document.querySelectorAll('#dots button')[k-1].click()},slide);
const tick=async dt=>{await p.evaluate(d=>{window.__VT+=d},dt);await p.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))))};
// warm two frames, then skip to start
await tick(50);await tick(50);if(startS>0){await tick(startS*1000)}
const n=Math.round((endS-startS)*fps);const t0=Date.now();
for(let i=0;i<n;i++){await tick(1000/fps);await p.screenshot({path:`${out}/f${String(i).padStart(5,'0')}.jpg`,type:'jpeg',quality:92,clip:{x:0,y:0,width:1600,height:900}});if(i%50===0)console.log(slide,i,'/',n,Math.round((Date.now()-t0)/1000)+'s')}
console.log('done',slide,n);await b.close();srv.close();})();
