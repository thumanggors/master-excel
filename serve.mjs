// static preview server for dist/ + preview.html; POST /save?name=x.png writes dist/pages/x.png
import http from 'node:http'; import fs from 'node:fs'; import path from 'node:path';
const T={'.html':'text/html','.pdf':'application/pdf','.png':'image/png','.css':'text/css','.js':'text/javascript'};
const out=path.join(import.meta.dirname,'dist','pages'); fs.mkdirSync(out,{recursive:true});
http.createServer((q,r)=>{const u=new URL(q.url,'http://x');
if(q.method==='POST'&&u.pathname==='/save'){const b=[];q.on('data',c=>b.push(c));q.on('end',()=>{fs.writeFileSync(path.join(out,path.basename(u.searchParams.get('name'))),Buffer.concat(b));r.end('ok')});return}
// /book serves the PDF without a .pdf URL (some browser setups intercept .pdf fetches)
const f=u.pathname==='/book'?path.join(import.meta.dirname,'dist','MASTER-RUMUS-EXCEL.pdf'):path.join(import.meta.dirname,decodeURIComponent(u.pathname));
fs.readFile(f,(e,d)=>{if(e){r.writeHead(404);return r.end('404')}r.writeHead(200,{'Content-Type':T[path.extname(f)]||'application/octet-stream','Cache-Control':'no-store'});r.end(d)})}).listen(5177);
