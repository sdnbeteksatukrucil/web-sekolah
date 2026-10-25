const sb=supabase.createClient(CFG.url,CFG.key),$=s=>document.querySelector(s),P={};let me;
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const ICON={dashboard:'🏠',siswa:'🎓',guru:'👩‍🏫',absensi:'🗓️',nilai:'📝',catatan:'⭐',tugas:'📚',pengaturan:'⚙️'};
const LBL={siswa:'Data Siswa',guru:'Data Guru'};
const MENU={admin:['dashboard','siswa','guru','absensi','nilai','tugas','catatan','pengaturan'],guru:['dashboard','siswa','guru','absensi','nilai','tugas','catatan'],siswa:['dashboard','tugas','nilai']};
const KLS=[1,2,3,4,5,6],today=()=>new Date().toISOString().slice(0,10);
const opts=(o,v='')=>o.map(([a,b])=>`<option value="${a}" ${a==v?'selected':''}>${b}</option>`).join('');
const kopt=all=>(all?[['','Semua kelas']]:[]).concat(KLS.map(k=>[k,'Kelas '+k]));
const sopt=S=>S.map(s=>[s.id,`${s.nama} (K${s.kelas})`]);
const inp=(n,l,v='',t='text',x='')=>`<label class="block text-sm mt-2">${l}<input name="${n}" type="${t}" value="${esc(v)}" ${x} class="inp"></label>`;
const ta=(n,l,v='',r=3)=>`<label class="block text-sm mt-2">${l}<textarea name="${n}" rows="${r}" class="inp">${esc(v)}</textarea></label>`;
const sel=(n,l,o,v='')=>`<label class="block text-sm mt-2">${l}<select name="${n}" class="inp">${opts(o,v)}</select></label>`;
const GR={'text-indigo-600':'from-indigo-500 to-violet-500','text-pink-600':'from-pink-500 to-rose-400','text-green-600':'from-emerald-500 to-teal-400','text-red-600':'from-rose-500 to-orange-400','text-amber-600':'from-amber-400 to-orange-400','text-sky-600':'from-sky-500 to-cyan-400'};
const stat=(l,v,c='text-indigo-600')=>`<div class="rounded-3xl p-4 text-white text-center shadow-lg bg-gradient-to-br ${GR[c]||GR['text-indigo-600']}"><div class="text-3xl font-extrabold">${v}</div><div class="text-xs opacity-90">${l}</div></div>`;
const wa=(hp,t)=>{let n=String(hp||'').replace(/\D/g,'');if(n.startsWith('0'))n='62'+n.slice(1);return `https://wa.me/${n}?text=${encodeURIComponent(t)}`};
const rng=m=>{const [y,mo]=m.split('-').map(Number);return[m+'-01',mo===12?`${y+1}-01-01`:`${y}-${String(mo+1).padStart(2,'0')}-01`]};
const up=async f=>{const p=Date.now()+'-'+f.name.replace(/[^\w.]/g,'_');const {error}=await sb.storage.from('media').upload(p,f);if(error)throw error;return sb.storage.from('media').getPublicUrl(p).data.publicUrl};
const xls=(rows,name)=>{const ws=XLSX.utils.aoa_to_sheet(rows),wb=XLSX.utils.book_new();ws['!cols']=rows[0].map((_,i)=>({wch:Math.min(40,Math.max(8,...rows.map(r=>String(r[i]??'').length)+0))}));XLSX.utils.book_append_sheet(wb,ws,'Data');XLSX.writeFile(wb,name+'.xlsx')};
const pager=(el,tot,pg,set)=>{const n=Math.max(1,Math.ceil(tot/20));el.innerHTML=`<div class="flex items-center justify-center gap-3 text-sm py-2"><button class="btn2" ${pg<1?'disabled':''}>‹</button><span>${pg+1} / ${n} · ${tot} data</span><button class="btn2" ${pg>=n-1?'disabled':''}>›</button></div>`;const [a,b]=el.querySelectorAll('button');a.onclick=()=>set(pg-1);b.onclick=()=>set(pg+1)};
const printQR=R=>{if(!R.length)return alert('Tidak ada siswa');const w=window.open('','_blank');if(!w)return alert('Izinkan pop-up untuk mencetak');
 const c=R.map(s=>{const d=document.createElement('div');new QRCode(d,{text:s.id,width:120,height:120});return `<div class="c"><img src="${d.querySelector('canvas').toDataURL()}"><b>${esc(s.nama)}</b><br>Kelas ${s.kelas}${s.nisn?' · NISN '+esc(s.nisn):''}<br><small>SD Negeri Betek I</small></div>`}).join('');
 w.document.write(`<title>Kartu QR</title><style>body{font-family:sans-serif;margin:8mm;display:grid;grid-template-columns:repeat(3,1fr);gap:6mm}.c{border:1px solid #333;border-radius:8px;padding:8px;text-align:center;font-size:12px;break-inside:avoid}img{width:120px;height:120px}</style>${c}`);w.document.close();setTimeout(()=>w.print(),300)};
const TPL=['Nama','NISN','NIS','NIK','TTL','Kelas','Ayah','Ibu','Pekerjaan','HP'];
const tpl=()=>{const ws=XLSX.utils.aoa_to_sheet([TPL,['Contoh Nama','0012345678','1234','3513000000000001','Probolinggo, 01-01-2016',1,'Nama Ayah','Nama Ibu','Petani','081234567890']]);
 for(let i=2;i<=300;i++)['B','C','D','J'].forEach(c=>{const a=c+i;if(!ws[a])ws[a]={t:'s',v:''};ws[a].t='s';ws[a].z='@'});
 ws['!ref']='A1:J300';ws['!cols']=TPL.map(()=>({wch:20}));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Siswa');XLSX.writeFile(wb,'Template-Data-Siswa.xlsx')};
async function importSiswa(ev){const f=ev.target.files[0];if(!f)return;try{const t=v=>String(v??'').trim(),wb=XLSX.read(await f.arrayBuffer()),R=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
 const D=R.map(r=>{let hp=t(r.HP).replace(/\D/g,'');if(hp&&!/^(0|62)/.test(hp))hp='0'+hp;return{nama:t(r.Nama),nisn:t(r.NISN),nis:t(r.NIS),nik:t(r.NIK),ttl:t(r.TTL),kelas:+r.Kelas,ayah:t(r.Ayah),ibu:t(r.Ibu),pekerjaan:t(r.Pekerjaan),hp}}).filter(d=>d.nama&&d.kelas>=1&&d.kelas<=6);
 if(!D.length)throw new Error('Tidak ada baris valid. Gunakan template; kolom Nama dan Kelas (1-6) wajib diisi.');
 if(!confirm(`Import ${D.length} siswa?`))return;
 for(let i=0;i<D.length;i+=100){const {error}=await sb.from('siswa').insert(D.slice(i,i+100));if(error)throw error}
 alert('Berhasil import '+D.length+' siswa');go()}catch(x){alert(x.message)}finally{ev.target.value=''}}
function modal(title,body,onOk){const d=document.createElement('div');d.className='fixed inset-0 z-50 bg-black/50 flex items-end sm:items-center justify-center p-2';
 d.innerHTML=`<div class="bg-white rounded-2xl w-full max-w-lg max-h-[90vh] overflow-auto p-5"><div class="flex justify-between mb-2"><b>${title}</b><button data-x>✕</button></div><form>${body}<button class="btn mt-4 w-full">Simpan</button></form></div>`;
 document.body.append(d);d.querySelector('[data-x]').onclick=()=>d.remove();
 d.querySelector('form').onsubmit=async ev=>{ev.preventDefault();const f=Object.fromEntries(new FormData(ev.target));try{await onOk(f);d.remove()}catch(x){alert(x.message||x)}}}
const mkAkun=(title,payload)=>modal(title,inp('username','Username',payload.u||'')+inp('password','Password (min. 6 karakter)','','password','minlength=6 required'),async f=>{
 const {data,error}=await sb.functions.invoke('create-user',{body:{...payload.b,username:f.username,password:f.password}});
 if(error){let m=error.message;try{const j=await error.context.json();m=j.error||j.message||m}catch{}throw new Error(m+' (pastikan fungsi create-user sudah di-deploy dan Verify JWT dimatikan)')}if(data?.error)throw new Error(data.error);alert('Akun berhasil dibuat')});

async function init(){
 const {data:{session}}=await sb.auth.getSession();if(!session)return location='index.html';
 const {data}=await sb.from('profiles').select('*').eq('id',session.user.id).single();
 if(!data){alert('Akun belum terdaftar di sistem');await sb.auth.signOut();return location='index.html'}
 me=data;$('#who').textContent=`${me.nama||''} · ${me.role}${me.kelas?' · kelas '+me.kelas:''}`;
 $('#nav').innerHTML=MENU[me.role].map(m=>`<a href="#${m}" data-m="${m}" class="flex gap-2 px-3 py-2 rounded-lg hover:bg-white/10 capitalize">${ICON[m]} ${LBL[m]||m}</a>`).join('');
 $('#out').onclick=async()=>{await sb.auth.signOut();location='index.html'};
 $('#burger').onclick=()=>$('#side').classList.toggle('-translate-x-full');
 addEventListener('hashchange',go);go()}
async function go(){const m=location.hash.slice(1)||'dashboard';if(!MENU[me.role].includes(m)){location.hash='dashboard';return}
 $('#side').classList.add('-translate-x-full');document.querySelectorAll('[data-m]').forEach(a=>a.classList.toggle('bg-white/20',a.dataset.m===m));
 $('#main').innerHTML='<p class="p-6 text-slate-400">Memuat…</p>';try{await P[m]()}catch(x){$('#main').innerHTML=`<p class="text-red-600">${esc(x.message)}</p>`}}

/* ===== DASHBOARD ===== */
P.dashboard=async()=>{const t=today(),[a,b]=rng(t.slice(0,7));
 if(me.role==='siswa'){
  const [A,C]=await Promise.all([sb.from('absensi').select('status').gte('tanggal',a).lt('tanggal',b),sb.from('catatan').select('*').order('tanggal',{ascending:false})]);
  const c=s=>A.data.filter(x=>x.status===s).length,pt=C.data.reduce((n,x)=>n+(x.jenis==='penghargaan'?x.poin:-x.poin),0);
  $('#main').innerHTML=`<h1 class="h mb-3">Halo, ${esc(me.nama)} 👋</h1><p class="text-sm text-slate-500 mb-2">Kehadiran bulan ini</p><div class="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">${stat('Hadir',c('H'),'text-green-600')}${stat('Sakit',c('S'),'text-amber-600')}${stat('Izin',c('I'),'text-sky-600')}${stat('Alpha',c('A'),'text-red-600')}</div>
  <div class="card"><b>Poin Catatan: <span class="${pt<0?'text-red-600':'text-green-600'}">${pt}</span></b><div class="divide-y mt-2">${C.data.map(x=>`<div class="py-2 text-sm flex gap-2"><span>${x.jenis==='penghargaan'?'🏅':'⚠️'}</span><div class="flex-1">${esc(x.deskripsi)}<div class="text-xs text-slate-400">${x.tanggal}</div></div><b>${x.jenis==='penghargaan'?'+':'-'}${x.poin}</b></div>`).join('')||'<p class="text-slate-400 text-sm">Belum ada catatan.</p>'}</div></div>`;return}
 const [S,G,H,A,V,B]=await Promise.all([sb.from('siswa').select('kelas'),me.role==='admin'?sb.from('guru').select('id'):{data:[]},sb.from('absensi').select('id',{count:'exact',head:true}).eq('tanggal',t).eq('status','H'),sb.from('absensi').select('id',{count:'exact',head:true}).eq('tanggal',t).eq('status','A'),sb.from('catatan').select('id',{count:'exact',head:true}).eq('jenis','pelanggaran'),sb.from('berita').select('judul,jenis,tanggal').order('tanggal',{ascending:false}).limit(5)]);
 const per=KLS.map(k=>[k,S.data.filter(x=>x.kelas===k).length]);
 $('#main').innerHTML=`<h1 class="h mb-3">Dashboard</h1><div class="grid grid-cols-2 md:grid-cols-5 gap-3 mb-4">${stat('Jumlah Siswa',S.data.length)}${me.role==='admin'?stat('Jumlah Guru',G.data.length,'text-pink-600'):''}${stat('Hadir Hari Ini',H.count||0,'text-green-600')}${stat('Alpha Hari Ini',A.count||0,'text-red-600')}${stat('Catatan Pelanggaran',V.count||0,'text-amber-600')}</div>
 <div class="grid md:grid-cols-2 gap-4"><div class="card"><b>Siswa per Kelas</b><div class="mt-2 space-y-1 text-sm">${per.map(([k,n])=>`<div class="flex justify-between"><span>Kelas ${k}</span><b>${n}</b></div>`).join('')}</div></div>
 <div class="card"><b>Berita & Pengumuman Terbaru</b><div class="mt-2 divide-y text-sm">${B.data.map(x=>`<div class="py-2"><span class="text-xs bg-slate-100 rounded px-1">${x.jenis}</span> ${esc(x.judul)}<div class="text-xs text-slate-400">${x.tanggal}</div></div>`).join('')||'<p class="text-slate-400">Belum ada.</p>'}</div></div></div>`};

/* ===== DATA SISWA ===== */
P.siswa=async()=>{const ad=me.role==='admin',{data:L}=await sb.from('siswa').select('*').order('kelas').order('nama');let k='',pg=0,q='';
 $('#main').innerHTML=`<div class="flex flex-wrap gap-2 items-center mb-3"><h1 class="h mr-auto">Data Siswa</h1><input id="q" placeholder="Cari nama / NISN…" class="inp !w-auto !mt-0"><button id="pq" class="btn2">🖨 Cetak Kartu QR</button>${ad?`<select id="kf" class="inp !w-auto !mt-0">${opts(kopt(1))}</select><button id="ad" class="btn">＋ Tambah Siswa</button><button id="tpl" class="btn2">⬇ Template</button><label class="btn2 cursor-pointer">⬆ Import<input id="imp" type="file" accept=".xlsx,.xls,.csv" hidden></label>`:''}</div>
 <div class="card overflow-x-auto"><table class="text-sm w-full"><thead class="bg-slate-100"><tr><th>QR</th><th>Nama</th><th>NISN</th><th>NIS</th><th>NIK</th><th>TTL</th><th>Kelas</th><th>Ayah</th><th>Ibu</th><th>Pekerjaan</th><th>No. HP</th>${ad?'<th>Aksi</th>':''}</tr></thead><tbody id="tb"></tbody></table></div><div id="pgn"></div>`;
 const form=s=>inp('nama','Nama Siswa',s?.nama,'text','required')+inp('nisn','NISN',s?.nisn)+inp('nis','NIS',s?.nis)+inp('nik','NIK',s?.nik)+inp('ttl','Tempat, Tanggal Lahir',s?.ttl)+sel('kelas','Kelas',kopt(),s?.kelas)+inp('ayah','Nama Ayah',s?.ayah)+inp('ibu','Nama Ibu',s?.ibu)+inp('pekerjaan','Pekerjaan Orang Tua',s?.pekerjaan)+inp('hp','No. HP (WhatsApp)',s?.hp,'tel');
 const save=id=>async f=>{f.kelas=+f.kelas;const {error}=id?await sb.from('siswa').update(f).eq('id',id):await sb.from('siswa').insert(f);if(error)throw error;go()};
 const flt=()=>L.filter(s=>(!k||s.kelas==k)&&(!q||(s.nama+' '+(s.nisn||'')).toLowerCase().includes(q)));const draw=()=>{const A=flt(),R=A.slice(pg*20,pg*20+20);pager($('#pgn'),A.length,pg,p=>{pg=p;draw()});
  $('#tb').innerHTML=R.map(s=>`<tr class="border-t"><td><div data-qr="${s.id}"></div></td><td class="whitespace-nowrap font-medium">${esc(s.nama)}</td><td>${esc(s.nisn)}</td><td>${esc(s.nis)}</td><td>${esc(s.nik)}</td><td>${esc(s.ttl)}</td><td>${s.kelas}</td><td>${esc(s.ayah)}</td><td>${esc(s.ibu)}</td><td>${esc(s.pekerjaan)}</td><td>${esc(s.hp)}</td>${ad?`<td class="whitespace-nowrap"><button data-e="${s.id}" class="text-indigo-600">Edit</button> · <button data-a="${s.id}" class="text-green-700">Akun</button> · <button data-d="${s.id}" class="text-red-600">Hapus</button></td>`:''}</tr>`).join('')||'<tr><td class="p-4 text-slate-400">Belum ada data</td></tr>';
  document.querySelectorAll('[data-qr]').forEach(el=>new QRCode(el,{text:el.dataset.qr,width:64,height:64}));
  document.querySelectorAll('[data-e]').forEach(b=>b.onclick=()=>{const s=L.find(x=>x.id===b.dataset.e);modal('Edit Siswa',form(s),save(s.id))});
  document.querySelectorAll('[data-a]').forEach(b=>b.onclick=()=>{const s=L.find(x=>x.id===b.dataset.a);mkAkun('Buat Akun Siswa: '+s.nama,{u:s.nisn||'',b:{role:'siswa',nama:s.nama,kelas:s.kelas,siswa_id:s.id}})});
  document.querySelectorAll('[data-d]').forEach(b=>b.onclick=async()=>{if(confirm('Hapus siswa beserta absensi, nilai, dan catatannya?')){await sb.from('siswa').delete().eq('id',b.dataset.d);go()}})};
 $('#q').oninput=ev=>{q=ev.target.value.toLowerCase();pg=0;draw()};$('#pq').onclick=()=>printQR(flt());if(ad){$('#kf').onchange=ev=>{k=ev.target.value;pg=0;draw()};$('#ad').onclick=()=>modal('Tambah Siswa',form(),save());$('#tpl').onclick=tpl;$('#imp').onchange=importSiswa}
 draw()};

/* ===== DATA GURU ===== */
P.guru=async()=>{const ad=me.role==='admin',{data:G}=await sb.from('guru').select('*').order('nama');
 const drive=u=>u?window.open(u,'_blank','noopener'):alert('Link Google Drive belum diisi admin');
 const form=g=>inp('nama','Nama Guru',g?.nama,'text','required')+inp('nip','NIP',g?.nip)+sel('kelas','Wali Kelas',[[0,'Semua Kelas (Kepala Sekolah / Guru Mapel)'],...kopt()],g?.kelas)+inp('hp','No. HP',g?.hp,'tel')+inp('drive_url','Link Google Drive (Data Saya)',g?.drive_url,'url');
 const save=id=>async f=>{f.kelas=+f.kelas;const {error}=id?await sb.from('guru').update(f).eq('id',id):await sb.from('guru').insert(f);if(error)throw error;go()};
 $('#main').innerHTML=`<div class="flex items-center mb-3"><h1 class="h mr-auto">Data Guru</h1>${ad?'<button id="ad" class="btn">＋ Tambah Guru</button>':''}</div><div class="card overflow-x-auto"><table class="text-sm w-full"><thead class="bg-slate-100"><tr><th>Nama</th><th>NIP</th><th>Wali Kelas</th><th>No. HP</th><th>Data Saya</th>${ad?'<th>Aksi</th>':''}</tr></thead><tbody>${G.map((g,i)=>`<tr class="border-t"><td class="font-medium">${esc(g.nama)}</td><td>${esc(g.nip)}</td><td>${g.kelas===0?'Semua kelas':g.kelas||'-'}</td><td>${esc(g.hp)}</td><td><button data-o="${i}" class="btn !py-1">📁 Data Saya</button></td>${ad?`<td class="whitespace-nowrap"><button data-e="${i}" class="text-indigo-600">Edit</button> · <button data-a="${i}" class="text-green-700">Akun</button> · <button data-d="${i}" class="text-red-600">Hapus</button></td>`:''}</tr>`).join('')||'<tr><td class="p-4 text-slate-400">Belum ada data</td></tr>'}</tbody></table></div>`;
 document.querySelectorAll('[data-o]').forEach(b=>b.onclick=()=>drive(G[b.dataset.o].drive_url));
 if(ad){$('#ad').onclick=()=>modal('Tambah Guru',form(),save());
  document.querySelectorAll('[data-e]').forEach(b=>b.onclick=()=>{const g=G[b.dataset.e];modal('Edit Guru',form(g),save(g.id))});
  document.querySelectorAll('[data-a]').forEach(b=>b.onclick=()=>{const g=G[b.dataset.a];mkAkun('Buat Akun Guru: '+g.nama,{b:{role:'guru',nama:g.nama,kelas:g.kelas,guru_id:g.id}})});
  document.querySelectorAll('[data-d]').forEach(b=>b.onclick=async()=>{if(confirm('Hapus data guru?')){await sb.from('guru').delete().eq('id',G[b.dataset.d].id);go()}})}};

/* ===== ABSENSI ===== */
P.absensi=async()=>{const g=me.role==='guru',t=today(),{data:S}=await sb.from('siswa').select('id,nama,kelas,hp').order('kelas').order('nama');let rows=[],cur='';
 $('#main').innerHTML=`<h1 class="h mb-3">Absensi</h1>${g?`<div class="card"><div class="flex flex-wrap gap-2"><input id="tg" type="date" value="${t}" class="inp !w-auto !mt-0"><button id="sc" class="btn">📷 Scan QR Absen</button></div><div id="ls" class="mt-3 divide-y"></div></div>`:''}
 <div class="card"><div class="flex flex-wrap gap-2 items-end"><input id="bl" type="month" value="${t.slice(0,7)}" class="inp !w-auto !mt-0">${g?'':`<select id="kf" class="inp !w-auto !mt-0">${opts(kopt(1))}</select>`}<button id="dl" class="btn">⬇ Unduh Excel</button><button id="wg" class="btn2">💬 Kirim ke WA Grup</button></div><div id="rk" class="overflow-x-auto mt-3"></div></div>`;
 const msg=r=>`Yth. Orang tua/wali ${r.s.nama}. Rekap kehadiran bulan ${cur}: Hadir ${r.c.H}, Sakit ${r.c.S}, Izin ${r.c.I}, Alpha ${r.c.A}. - SD Negeri Betek I`;
 const rek=async()=>{cur=$('#bl').value;const [a,b]=rng(cur),D=new Date(+cur.slice(0,4),+cur.slice(5),0).getDate();
  const {data}=await sb.from('absensi').select('siswa_id,tanggal,status').gte('tanggal',a).lt('tanggal',b),kf=$('#kf')?.value;
  rows=S.filter(s=>!kf||s.kelas==kf).map(s=>{const d=Array(D).fill(''),c={H:0,S:0,I:0,A:0};data.filter(x=>x.siswa_id===s.id).forEach(x=>{d[+x.tanggal.slice(8)-1]=x.status;c[x.status]++});return{s,d,c}});
  $('#rk').innerHTML=`<table class="text-xs"><thead class="bg-slate-100"><tr><th>Nama</th>${[...Array(D)].map((_,i)=>`<th class="!px-1">${i+1}</th>`).join('')}<th>H</th><th>S</th><th>I</th><th>A</th><th>WA</th></tr></thead><tbody>${rows.map(r=>`<tr class="border-t"><td class="whitespace-nowrap">${esc(r.s.nama)}</td>${r.d.map(x=>`<td class="!px-1 text-center">${x}</td>`).join('')}<td>${r.c.H}</td><td>${r.c.S}</td><td>${r.c.I}</td><td>${r.c.A}</td><td><a target="_blank" rel="noopener" href="${wa(r.s.hp,msg(r))}">📲</a></td></tr>`).join('')}</tbody></table>`};
 $('#bl').onchange=rek;if($('#kf'))$('#kf').onchange=rek;
 $('#dl').onclick=()=>xls([['Nama Siswa','Kelas',...rows[0]?.d.map((_,i)=>i+1)||[],'Hadir','S','I','A'],...rows.map(r=>[r.s.nama,r.s.kelas,...r.d,r.c.H,r.c.S,r.c.I,r.c.A])],'Rekap-Absensi-'+cur);
 $('#wg').onclick=()=>window.open('https://wa.me/?text='+encodeURIComponent(`Rekap kehadiran siswa bulan ${cur} - SD Negeri Betek I\n`+rows.map(r=>`${r.s.nama}: H${r.c.H} S${r.c.S} I${r.c.I} A${r.c.A}`).join('\n')),'_blank');
 if(g){const day=async()=>{const {data}=await sb.from('absensi').select('siswa_id,status').eq('tanggal',$('#tg').value),mp=Object.fromEntries(data.map(x=>[x.siswa_id,x.status]));
   $('#ls').innerHTML=S.map(s=>`<div class="flex items-center justify-between py-2"><span class="text-sm">${esc(s.nama)}</span><select data-s="${s.id}" class="inp !w-auto !mt-0">${opts([['','—'],['H','Hadir'],['S','Sakit'],['I','Izin'],['A','Alpha']],mp[s.id]||'')}</select></div>`).join('');
   $('#ls').querySelectorAll('select').forEach(el=>el.onchange=async()=>{const m={siswa_id:el.dataset.s,tanggal:$('#tg').value};
    const {error}=el.value?await sb.from('absensi').upsert({...m,status:el.value},{onConflict:'siswa_id,tanggal'}):await sb.from('absensi').delete().match(m);if(error)alert(error.message);rek()})};
  $('#tg').onchange=day;day();
  $('#sc').onclick=()=>{const o=document.createElement('div');o.className='fixed inset-0 z-50 bg-black/80 flex flex-col items-center justify-center p-4';
   o.innerHTML='<div id="rd" class="w-full max-w-sm bg-white rounded-xl overflow-hidden"></div><p id="rs" class="text-white mt-3 text-center">Arahkan kamera ke QR siswa</p><button class="btn mt-3">Tutup</button>';document.body.append(o);
   const q=new Html5Qrcode('rd');let last='';const close=async()=>{try{await q.stop()}catch{}o.remove();rek();day()};o.querySelector('button').onclick=close;
   q.start({facingMode:'environment'},{fps:10,qrbox:220},async tx=>{if(tx===last)return;last=tx;const s=S.find(x=>x.id===tx);if(!s){$('#rs').textContent='❌ QR tidak dikenal / bukan siswa kelas ini'}else{
     const {error}=await sb.from('absensi').upsert({siswa_id:s.id,tanggal:$('#tg').value,status:'H'},{onConflict:'siswa_id,tanggal'});$('#rs').textContent=error?'❌ '+error.message:`✅ ${s.nama} hadir`}setTimeout(()=>last='',3000)}).catch(x=>$('#rs').textContent='Kamera gagal dibuka: '+x)}}
 rek()};

/* ===== NILAI & CATATAN (halaman daftar generik) ===== */
const listPage=c=>async()=>{const ed=me.role!=='siswa',S=ed?(await sb.from('siswa').select('id,nama,kelas').order('kelas').order('nama')).data:[];
 const {data:L}=await sb.from(c.t).select('*,siswa(nama,kelas)').order('created_at',{ascending:false});let k='',pg=0,q='';
 $('#main').innerHTML=`<div class="flex flex-wrap gap-2 items-center mb-3"><h1 class="h mr-auto">${c.title}</h1><input id="q" placeholder="Cari nama siswa…" class="inp !w-auto !mt-0">${me.role==='admin'?`<select id="kf" class="inp !w-auto !mt-0">${opts(kopt(1))}</select>`:''}<button id="xl" class="btn2">⬇ Excel</button>${ed?'<button id="ad" class="btn">＋ Tambah</button>':''}</div><div class="card overflow-x-auto"><table class="text-sm w-full"><thead class="bg-slate-100"><tr><th>Kelas</th><th>Nama</th>${c.cols.map(x=>`<th>${x[1]}</th>`).join('')}${ed?'<th></th>':''}</tr></thead><tbody id="tb"></tbody></table></div><div id="pgn"></div>`;
 const R=()=>L.filter(x=>(!k||x.siswa?.kelas==k)&&(!q||(x.siswa?.nama||'').toLowerCase().includes(q)));
 const draw=()=>{const A=R(),V=A.slice(pg*20,pg*20+20);pager($('#pgn'),A.length,pg,p=>{pg=p;draw()});$('#tb').innerHTML=V.map(x=>`<tr class="border-t"><td>${x.siswa?.kelas??''}</td><td class="whitespace-nowrap">${esc(x.siswa?.nama)}</td>${c.cols.map(y=>`<td>${esc(x[y[0]])}</td>`).join('')}${ed?`<td class="whitespace-nowrap"><button data-e="${x.id}" class="text-indigo-600">Edit</button> · <button data-d="${x.id}" class="text-red-600">Hapus</button></td>`:''}</tr>`).join('')||'<tr><td class="p-4 text-slate-400">Belum ada data</td></tr>';
  document.querySelectorAll('[data-e]').forEach(b=>b.onclick=()=>{const x=L.find(y=>y.id===b.dataset.e);modal('Edit '+c.title,sel('siswa_id','Siswa',sopt(S),x.siswa_id)+c.form(x),async f=>{(c.num||[]).forEach(n=>f[n]=+f[n]);const {error}=await sb.from(c.t).update(f).eq('id',x.id);if(error)throw error;go()})});
  document.querySelectorAll('[data-d]').forEach(b=>b.onclick=async()=>{if(confirm('Hapus data ini?')){await sb.from(c.t).delete().eq('id',b.dataset.d);go()}})};
 $('#q').oninput=ev=>{q=ev.target.value.toLowerCase();pg=0;draw()};
 if($('#kf'))$('#kf').onchange=ev=>{k=ev.target.value;pg=0;draw()};
 $('#xl').onclick=()=>xls([['Kelas','Nama',...c.cols.map(x=>x[1])],...R().map(x=>[x.siswa?.kelas,x.siswa?.nama,...c.cols.map(y=>x[y[0]])])],c.title.replace(/\s/g,'-'));
 if(ed)$('#ad').onclick=()=>modal(c.title,sel('siswa_id','Siswa',sopt(S))+c.form(),async f=>{(c.num||[]).forEach(n=>f[n]=+f[n]);const {error}=await sb.from(c.t).insert(f);if(error)throw error;go()});
 draw()};
const MAPEL0=['Agama','Bhs. Indonesia','Matematika','Pendidikan Pancasila','IPAS','Seni Rupa','PJOK','Bahasa Inggris','Koding dan KA','Mulok'];
P.nilai=async()=>{const ed=me.role!=='siswa';
 const [a,b,c]=await Promise.all([sb.from('pengaturan').select('data').eq('key','mapel').maybeSingle(),sb.from('siswa').select('id,nama,kelas').order('kelas').order('nama'),sb.from('nilai').select('siswa_id,mapel,semester,nilai')]);
 const M=a.data?.data?.list?.length?a.data.data.list:MAPEL0,S=b.data,N=c.data;let sem=P._sem||'Ganjil',k='',q='',pg=0;
 const nv=(id,m)=>N.find(x=>x.siswa_id===id&&x.mapel===m&&x.semester===sem)?.nilai;
 $('#main').innerHTML=`<div class="flex flex-wrap gap-2 items-center mb-3"><h1 class="h mr-auto">Nilai Siswa</h1><select id="sm" class="inp !w-auto !mt-0">${opts([['Ganjil','Semester Ganjil'],['Genap','Semester Genap']],sem)}</select>${me.role==='admin'?`<select id="kf" class="inp !w-auto !mt-0">${opts(kopt(1))}</select>`:''}<input id="q" placeholder="Cari nama…" class="inp !w-auto !mt-0"><button id="xl" class="btn2">⬇ Excel</button>${ed?'<button id="am" class="btn2">⚙ Atur Mapel</button>':''}</div><div class="card overflow-x-auto"><table class="text-sm w-full"><thead class="bg-indigo-100"><tr><th>Nama Siswa</th><th>Kls</th>${M.map(m=>`<th>${esc(m)}</th>`).join('')}<th>Rata²</th>${ed?'<th></th>':''}</tr></thead><tbody id="tb"></tbody></table></div><div id="pgn"></div>`;
 const rows=()=>S.filter(s=>(!k||s.kelas==k)&&(!q||s.nama.toLowerCase().includes(q)));
 const avg=s=>{const v=M.map(m=>nv(s.id,m)).filter(x=>x!=null);return v.length?(v.reduce((p,x)=>p+ +x,0)/v.length).toFixed(1):''};
 const draw=()=>{const A=rows(),V=A.slice(pg*20,pg*20+20);pager($('#pgn'),A.length,pg,p=>{pg=p;draw()});
  $('#tb').innerHTML=V.map(s=>`<tr class="border-t"><td class="whitespace-nowrap font-medium">${esc(s.nama)}</td><td>${s.kelas}</td>${M.map(m=>`<td>${nv(s.id,m)??''}</td>`).join('')}<td class="font-semibold">${avg(s)}</td>${ed?`<td><button data-i="${s.id}" class="btn !py-1">✏️ Input</button></td>`:''}</tr>`).join('')||'<tr><td class="p-4 text-slate-400">Belum ada data</td></tr>';
  document.querySelectorAll('[data-i]').forEach(bt=>bt.onclick=()=>{const s=S.find(x=>x.id===bt.dataset.i);
   modal(`Nilai ${s.nama} · ${sem}`,`<div class="grid grid-cols-2 gap-x-3">${M.map((m,i)=>inp('m'+i,m,nv(s.id,m)??'','number','min=0 max=100 step=0.01')).join('')}</div>`,async f=>{
    const D=M.map((m,i)=>f['m'+i]===''?null:{siswa_id:s.id,mapel:m,semester:sem,nilai:+f['m'+i]}).filter(Boolean);
    if(!D.length)throw new Error('Isi minimal satu nilai');const {error}=await sb.from('nilai').upsert(D,{onConflict:'siswa_id,mapel,semester'});if(error)throw error;go()})})};
 $('#sm').onchange=ev=>{sem=P._sem=ev.target.value;draw()};$('#q').oninput=ev=>{q=ev.target.value.toLowerCase();pg=0;draw()};if($('#kf'))$('#kf').onchange=ev=>{k=ev.target.value;pg=0;draw()};
 $('#xl').onclick=()=>xls([['Nama Siswa','Kelas',...M,'Rata-rata'],...rows().map(s=>[s.nama,s.kelas,...M.map(m=>nv(s.id,m)??''),avg(s)])],'Nilai-'+sem);
 if(ed)$('#am').onclick=()=>modal('Atur Mata Pelajaran',ta('list','Satu mata pelajaran per baris',M.join('\n'),10),async f=>{const list=f.list.split('\n').map(x=>x.trim()).filter(Boolean);if(!list.length)throw new Error('Minimal satu mata pelajaran');const {error}=await sb.from('pengaturan').upsert({key:'mapel',data:{list}});if(error)throw error;go()});
 draw()};
P.tugas=async()=>{const ed=me.role!=='siswa',{data:L}=await sb.from('tugas').select('*').order('created_at',{ascending:false});let k='';
 const ko=me.role==='admin'||me.kelas===0?kopt():[[me.kelas,'Kelas '+me.kelas]];
 $('#main').innerHTML=`<div class="flex flex-wrap gap-2 items-center mb-3"><h1 class="h mr-auto">Tugas</h1>${ed?`<select id="kf" class="inp !w-auto !mt-0">${opts(kopt(1))}</select><button id="ad" class="btn">＋ Beri Tugas</button>`:''}</div><div id="tl" class="grid md:grid-cols-2 gap-3"></div>`;
 const draw=()=>{$('#tl').innerHTML=L.filter(x=>!k||x.kelas==k).map(x=>`<div class="card !mb-0 border-t-4 border-fuchsia-400"><div class="flex flex-wrap gap-2 text-xs mb-1"><span class="bg-indigo-100 text-indigo-700 rounded-full px-2 py-0.5">Kelas ${x.kelas}</span><span class="bg-amber-100 text-amber-700 rounded-full px-2 py-0.5">${esc(x.mapel)}</span>${x.deadline?`<span class="ml-auto text-rose-600">⏰ ${x.deadline}</span>`:''}</div><b>${esc(x.judul)}</b><p class="text-sm text-slate-600 whitespace-pre-line mt-1">${esc(x.deskripsi)}</p><div class="mt-2 flex gap-3 text-sm">${/^https?:/i.test(x.link||'')?`<a class="text-indigo-600 underline" target="_blank" rel="noopener" href="${esc(x.link)}">🔗 Buka materi</a>`:''}${ed?`<button data-d="${x.id}" class="ml-auto text-red-600">Hapus</button>`:''}</div></div>`).join('')||'<p class="text-slate-400">Belum ada tugas.</p>';
  document.querySelectorAll('[data-d]').forEach(b=>b.onclick=async()=>{if(confirm('Hapus tugas ini?')){await sb.from('tugas').delete().eq('id',b.dataset.d);go()}})};
 if(ed){$('#kf').onchange=ev=>{k=ev.target.value;draw()};$('#ad').onclick=()=>modal('Beri Tugas',sel('kelas','Kelas',ko)+inp('mapel','Mata Pelajaran','','text','required')+inp('judul','Judul Tugas','','text','required')+ta('deskripsi','Petunjuk / Deskripsi')+inp('deadline','Batas Pengumpulan','','date')+inp('link','Link materi / Google Drive (opsional)','','url'),async f=>{f.kelas=+f.kelas;if(!f.deadline)f.deadline=null;const {error}=await sb.from('tugas').insert(f);if(error)throw error;go()})}
 draw()};
P.catatan=listPage({t:'catatan',title:'Catatan Siswa',cols:[['jenis','Jenis'],['poin','Poin'],['deskripsi','Deskripsi'],['tanggal','Tanggal']],num:['poin'],
 form:v=>sel('jenis','Jenis',[['penghargaan','Penghargaan'],['pelanggaran','Pelanggaran']],v?.jenis)+inp('poin','Poin',v?.poin??5,'number','min=0 required')+ta('deskripsi','Deskripsi',v?.deskripsi)});

/* ===== PENGATURAN WEB PUBLIK (admin) ===== */
P.pengaturan=async()=>{const [a,b,c]=await Promise.all([sb.from('pengaturan').select('data').eq('key','site').maybeSingle(),sb.from('berita').select('*').order('tanggal',{ascending:false}),sb.from('galeri').select('*').order('created_at',{ascending:false})]);
 const s=a.data?.data||{};
 $('#main').innerHTML=`<h1 class="h mb-3">Pengaturan Web Publik</h1>
 <form id="fs" class="card"><b>Profil, Visi, Misi & Kontak</b>${inp('nama','Nama Sekolah',s.nama)}${inp('tagline','Slogan',s.tagline)}${ta('profil','Profil Sekolah',s.profil,5)}${ta('visi','Visi',s.visi)}${ta('misi','Misi (satu baris = satu poin)',s.misi,5)}${inp('alamat','Alamat',s.alamat)}${inp('telp','Telepon',s.telp)}${inp('email','Email',s.email,'email')}${inp('wa','No. WhatsApp (628xxx)',s.wa)}${inp('yt','Link YouTube',s.yt,'url')}${inp('tiktok','Link TikTok',s.tiktok,'url')}${inp('wachan','Link Channel WhatsApp',s.wachan,'url')}
  <label class="block text-sm mt-2">Logo ${s.logo?`<img src="${esc(s.logo)}" class="h-10 inline">`:''}<input type="file" name="logo" accept="image/*" class="inp"></label><label class="block text-sm mt-2">Foto Sekolah ${s.hero?`<img src="${esc(s.hero)}" class="h-10 inline">`:''}<input type="file" name="hero" accept="image/*" class="inp"></label><label class="block text-sm mt-2">Foto Latar Beranda ${s.bg?`<img src="${esc(s.bg)}" class="h-10 inline">`:''}<input type="file" name="bg" accept="image/*" class="inp"></label>${inp('bgop','Ketebalan foto latar (10 = sangat transparan, 90 = tebal)',s.bgop??40,'number','min=10 max=90')}<button class="btn mt-4">Simpan Pengaturan</button></form>
 <div class="card"><div class="flex items-center mb-2"><b class="mr-auto">Berita & Pengumuman</b><button id="ab" class="btn">＋ Tambah</button></div><div class="divide-y text-sm">${b.data.map(x=>`<div class="py-2 flex gap-2"><span class="text-xs bg-slate-100 rounded px-1 h-fit">${x.jenis}</span><span class="flex-1">${esc(x.judul)}</span><button data-bd="${x.id}" class="text-red-600">Hapus</button></div>`).join('')||'<p class="text-slate-400">Belum ada.</p>'}</div></div>
 <div class="card"><b>Galeri Kegiatan</b><label class="block text-sm mt-2">Unggah foto (boleh banyak)<input id="gu" type="file" accept="image/*" multiple class="inp"></label><div class="grid grid-cols-3 md:grid-cols-6 gap-2 mt-3">${c.data.map(x=>`<div class="relative"><img loading="lazy" src="${esc(x.url)}" class="aspect-square object-cover rounded-lg"><button data-gd="${x.id}" class="absolute top-1 right-1 bg-white/90 rounded px-1 text-xs text-red-600">✕</button></div>`).join('')}</div></div>`;
 $('#fs').onsubmit=async ev=>{ev.preventDefault();const f=Object.fromEntries(new FormData(ev.target)),d={...s,...f};
  try{for(const n of['logo','hero','bg']){if(f[n]?.size)d[n]=await up(f[n]);else d[n]=s[n]}const {error}=await sb.from('pengaturan').upsert({key:'site',data:d});if(error)throw error;alert('Tersimpan');go()}catch(x){alert(x.message)}};
 $('#ab').onclick=()=>modal('Tambah Berita / Pengumuman',sel('jenis','Jenis',[['berita','Berita'],['pengumuman','Pengumuman']])+inp('judul','Judul','','text','required')+ta('isi','Isi',' ',5)+`<label class="block text-sm mt-2">Gambar (opsional)<input type="file" name="gambar" accept="image/*" class="inp"></label>`,async f=>{
  const d={jenis:f.jenis,judul:f.judul,isi:f.isi.trim()};if(f.gambar?.size)d.gambar=await up(f.gambar);const {error}=await sb.from('berita').insert(d);if(error)throw error;go()});
 document.querySelectorAll('[data-bd]').forEach(x=>x.onclick=async()=>{if(confirm('Hapus?')){await sb.from('berita').delete().eq('id',x.dataset.bd);go()}});
 document.querySelectorAll('[data-gd]').forEach(x=>x.onclick=async()=>{if(confirm('Hapus foto?')){await sb.from('galeri').delete().eq('id',x.dataset.gd);go()}});
 $('#gu').onchange=async ev=>{try{for(const f of ev.target.files){await sb.from('galeri').insert({judul:f.name,url:await up(f)})}go()}catch(x){alert(x.message)}}};
init();
