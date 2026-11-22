const V='sat-v23';
const FILES=['./','index.html','Acces.html','FormUsagers.html','FormPersonnels.html','FormCogecs.html','FormUsagersHopital.html','FormPersonnelHopital.html','FormCommunauteHopital.html','hform.js','AdminDashboard.html','AdminQR.html','AdminAttestations.html','AdminSuivi.html','Facture.html','Enqueteur.html','attestation.js','att/cadre.jpg','att/logo.png','att/tampon.png','att/narrow.woff','att/narrow-bold.woff','att/gelasio-bold.woff','att/gelasio-bolditalic.woff','config.js','manifest.json','icon-192.png','icon-512.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(V).then(c=>Promise.all(FILES.map(f=>c.add(f).catch(()=>{})))).then(()=>self.skipWaiting()));});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==V).map(x=>caches.delete(x)))).then(()=>self.clients.claim()));});
self.addEventListener('fetch',e=>{const r=e.request,u=new URL(r.url);
 if(r.method!=='GET'||u.origin!==location.origin)return;
 e.respondWith(fetch(r).then(res=>{if(res.ok){const cp=res.clone();caches.open(V).then(c=>c.put(r,cp));}return res;})
  .catch(()=>caches.match(r,{ignoreSearch:true}).then(m=>m||caches.match('index.html'))));});
