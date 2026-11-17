/* Дашборд виконання Плану роботи — локальна версія.
   Джерела даних (за пріоритетом): 1) Google-таблиця (CONFIG.sheetId або csvUrl),
   2) файл xlsx кнопкою «Відкрити xlsx», 3) вбудовані дані з data.js. */
const CONFIG={
  // Ваша таблиця. Щоб сторінка могла її читати, відкрийте доступ:
  // Google Sheets → Поділитися → Загальний доступ → «Усі, хто має посилання» → роль «Читач».
  sheetId:'10VAw8cK2dC4mgH7VdKZDtNlByxFEHazcjz06r7hT3bg',
  sheetName:'Звіт 2026',
  // Необов'язково: готове CSV-посилання («Опублікувати в Інтернеті»). Якщо заповнено — використовується замість sheetId.
  csvUrl:'',
  refreshMs:20000   // як часто оновлювати, мс
};
const csvSource=()=>CONFIG.csvUrl||(CONFIG.sheetId?'https://docs.google.com/spreadsheets/d/'+CONFIG.sheetId+'/gviz/tq?tqx=out:csv&headers=0&sheet='+encodeURIComponent(CONFIG.sheetName):'');
const UN=[['Департамент розвитку сімейних','Деп. розвитку сімейних форм і усиновлення'],['Департамент надання соціальної','Деп. надання соціальної підтримки'],['Департамент моніторингу та захисту','Деп. моніторингу та захисту прав дітей'],['Управління координації та методичного','Упр. координації та методичного забезпечення'],['Управління стратегічного розвитку','Упр. стратегічного розвитку та аналізу процесів'],['Відділ запобігання дитячій','Відділ запобігання бездоглядності та насильству'],['Управління правового','Упр. правового забезпечення'],['Відділ з питань цифрового','Відділ цифрового розвитку та інновацій'],['Відділ управління персоналом','Відділ управління персоналом'],['Сектор з питань запобігання корупції','Сектор запобігання корупції'],['Сектор внутрішнього аудиту','Сектор внутрішнього аудиту'],['Управління фінансового','Упр. фінансового забезпечення та обліку']];
const $=i=>document.getElementById(i);
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const pct=(a,b)=>b?Math.round(a/b*100):0;
const norm=s=>String(s||'').replace(/\s+/g,' ').trim();
const cut=(s,n)=>s.length>n?s.slice(0,n)+'…':s;
const dirs=[['Захист прав дитини (розд. 1–6)',1,6],['Підтримка сімей з дітьми (розд. 7–10)',7,10],['Зростання в сімейному оточенні (розд. 11–12)',11,12],['Організаційний розвиток Служби (розд. 13–20)',13,20]];
const stat=a=>({n:a.length,ok:a.filter(t=>t.ok).length});
const NOW=new Date(),TD=new Date(NOW.getFullYear(),NOW.getMonth(),NOW.getDate()),DAY=864e5,YEAR=2026,EOY=new Date(YEAR,11,31,23,59,59);
const p2=n=>String(n).padStart(2,'0'),dstr=d=>d.getFullYear()+'-'+p2(d.getMonth()+1)+'-'+p2(d.getDate()),fmtD=s=>s.slice(8)+'.'+s.slice(5,7);
const fmtT=ms=>{const d=new Date(ms);return p2(d.getDate())+'.'+p2(d.getMonth()+1)+' '+p2(d.getHours())+':'+p2(d.getMinutes())};
const wip=t=>!t.ok&&/виконуєт/i.test(t.info||'');
let extra=null,extraLabel='',T=STATIC,src='static',at=0,open=new Set(),snaps=[],latest=null,A=null,G=null,Fv=null,err='',db=null,canW=false,dbA=false,dbB=false,persisted='',dlz=null,sampleFn=null,mcp=null;

function parse(A,G,F){const out=[];A.forEach((r,i)=>{const m=/^(\d+)\.(\d+)\.?$/.exec(norm(r&&r[0]));if(!m)return;
 const resp=norm(r[3]),us=UN.map(([k,s])=>[resp.indexOf(k),s]).filter(x=>x[0]>=0).sort((a,b)=>a[0]-b[0]).map(x=>x[1]),g=norm(G[i]&&G[i][0]);
 out.push({r:i+1,id:m[1]+'.'+m[2],s:+m[1],t:norm(r[1]),term:norm(r[2]),u:us[0]||'Керівники підрозділів',un:!us.length&&!/керівник/i.test(resp),co:us.slice(1),g,ok:g.toUpperCase()==='ВИКОНАНО',info:norm(F&&F[i]&&F[i][0]).slice(0,1500)})});return out}
const MS=[['січ',0],['лют',1],['берез',2],['квіт',3],['трав',4],['черв',5],['лип',6],['серп',7],['верес',8],['жовт',9],['листоп',10],['груд',11]];
function dlDate(term){const s=term.toLowerCase();let best=null,cnt=0;MS.forEach(([k,m])=>{const re=new RegExp('(^|[^а-яіїєґ])'+k,'g');let x,f=0;while((x=re.exec(s))){f=1;if(!best||x.index>best.i)best={i:x.index,m}}cnt+=f});
 if(!best)return null;const d=/(\d{1,2})\s*[а-яіїєґ]+/.exec(s);
 return(d&&cnt===1&&s.indexOf('до')===0)?new Date(YEAR,best.m,+d[1],23,59,59):new Date(YEAR,best.m+1,0,23,59,59)}
function dlList(){const L=T.filter(t=>!/^протягом/i.test(t.term)).map(t=>({...t,d:dlDate(t.term)})).sort((a,b)=>(a.d||0)-(b.d||0)||a.r-b.r);
 L.forEach(t=>{t.dd=t.d?Math.round((new Date(t.d.getFullYear(),t.d.getMonth(),t.d.getDate())-TD)/DAY):null;if(t.ok){t.k='ok';t.l='Виконано'}else if(t.d&&t.d<TD){t.k='bad';t.l='Прострочено на '+(-t.dd)+' дн.'}else{t.k='wait';t.l=t.d?(t.dd===0?'Строк — сьогодні':'У строк, залишилось '+t.dd+' дн.'):'У строк'}});return L}

const row=(u,x)=>`<div class="row" data-unit="${esc(u)}" title="${esc(u)}: ${x.ok} з ${x.n}. Натисніть, щоб відфільтрувати"><div class="lbl">${esc(u)}</div><div class="bar"><b class="ok" style="width:${x.ok/x.n*100}%"></b><b class="no" style="width:${(x.n-x.ok)/x.n*100}%"></b></div><div class="pc">${pct(x.ok,x.n)}%<small>${x.ok}/${x.n}</small></div></div>`;
function fill(sel,vals,lab){const c=sel.value;sel.length=1;vals.forEach(v=>sel.insertAdjacentHTML('beforeend',`<option value="${esc(v)}">${esc(lab?lab+v:v)}</option>`));sel.value=c}
const kc=(v,l,cls,act)=>`<div class="c kc ${cls||''}" ${act?`data-act="${act}" role="button" tabindex="0" title="Натисніть, щоб переглянути"`:''}><div class="kv">${v}</div><div class="kl">${l}</div></div>`;

function setData(t,s,a){T=t;src=s;at=a||Date.now();status();build()}
function status(){const e=err?' · '+err:'';const L={csv:'Джерело: Google-таблиця (CSV) · оновлено '+fmtT(at),file:'Джерело: файл «'+fname+'» · завантажено '+fmtT(at),static:'Показано вбудовані дані (data.js)'};$('src').textContent=L[src]+e;$('src').className=src==='static'?'warn':'live'}

function build(){
 const n=T.length,ok=T.filter(t=>t.ok).length,no=n-ok;
 const bySec={},units={};T.forEach(t=>{(bySec[t.s]=bySec[t.s]||[]).push(t);(units[t.u]=units[t.u]||[]).push(t)});
 const secs=Object.keys(bySec).sort((a,b)=>a-b),full=Object.values(units).filter(a=>a.every(t=>t.ok)).length,zero=secs.filter(s=>stat(bySec[s]).ok===0);
 $('hero').innerHTML=`<div class="ring" style="background:conic-gradient(var(--lime) 0 ${pct(ok,n)}%,#3a3f2c 0)"><i>${pct(ok,n)}%</i></div><div><div class="big">${ok} із ${n} завдань виконано</div><p>Не виконано ${no} із ${n} завдань.</p></div>`;
 $('kpi').innerHTML=kc(n,'завдань у плані','','all')+kc(no,'не виконано','hl','nd')+kc(full+' з '+Object.keys(units).length,'підрозділів на 100%','','u100')+kc(zero.length,'розділів без виконаних завдань','','z0');
 const ds=dirs.map(([l,a,b])=>[l,stat(T.filter(t=>t.s>=a&&t.s<=b))]);
 $('dir').innerHTML=ds.map(([l,x])=>`<div class="c dc"><div class="t">${esc(l)}</div><div class="n">${pct(x.ok,x.n)}%</div><div class="bar" style="margin:6px 0"><b class="ok" style="width:${x.n?x.ok/x.n*100:0}%"></b><b class="no" style="width:${x.n?(x.n-x.ok)/x.n*100:0}%"></b></div><div class="mu">${x.ok} з ${x.n} завдань</div></div>`).join('');
 const mx=Math.max(1,...secs.map(s=>bySec[s].length));
 $('sec').innerHTML=secs.map(s=>{const x=stat(bySec[s]);return `<div class="col" data-sec="${s}" title="Розділ ${s}: ${x.ok} з ${x.n} виконано. Натисніть, щоб відфільтрувати"><div class="st" style="height:${x.n/mx*88}%"><b class="ok" style="height:${x.ok/x.n*100}%"></b><b class="no" style="height:${(x.n-x.ok)/x.n*100}%"></b></div><span>${s}</span></div>`}).join('');
 $('unit').innerHTML=Object.entries(units).sort((a,b)=>pct(stat(a[1]).ok,a[1].length)-pct(stat(b[1]).ok,b[1].length)).map(([u,a])=>row(u,stat(a))).join('');
 fill($('fsec'),secs,'Розділ ');fill($('fu'),Object.keys(units).sort());
 renderAtt();renderTrend();renderDl();render();
}
function renderAtt(){
 const L=dlList(),bad=L.filter(t=>t.k==='bad'),soon=L.filter(t=>t.k==='wait'&&t.dd!==null&&t.dd<=60),ann=T.filter(t=>!t.ok&&/^протягом/i.test(t.term)),left=Math.round((new Date(YEAR,11,31)-TD)/DAY);
 const it=[...bad,...soon];
 $('att').innerHTML=`<h2>Потребує уваги</h2><div class="g k4">${kc(left,'днів до кінця '+YEAR+' року')}${kc(bad.length,'прострочених строкових завдань',bad.length?'bd':'')}${kc(soon.length,'строків у найближчі 60 днів')}${kc(ann.length,'річних завдань іще не виконано','', 'ann')}</div>`+
 (it.length?`<ul class="al">${it.map(t=>`<li><b>${esc(t.id)}</b><div>${esc(cut(t.t,150))}<div class="mu">${esc(t.u)} · строк: ${esc(t.term)}</div></div><span class="b ${t.k==='bad'?'bd2':'no'}">${esc(t.l)}</span></li>`).join('')}</ul>`:'<p class="mu" style="margin:12px 0 0">Прострочених строкових завдань і строків у найближчі 60 днів немає.</p>')+
 `<p class="mu" style="margin:12px 0 0;font-size:12px">Завдання «Протягом року» вважаються такими, що мають бути виконані до 31 грудня ${YEAR}.</p>`;
}
function renderTrend(){const P=snaps.filter(s=>s.n).sort((a,b)=>a.date<b.date?-1:1),el=$('trend');
 $('trendc').hidden=P.length<2;if(P.length<2){el.innerHTML='';return}
 const W=Math.max(300,el.clientWidth||600),H=170,l=34,r=10,t=12,b=24,x=i=>l+i/(P.length-1)*(W-l-r),y=v=>t+(100-v)/100*(H-t-b),pts=P.map((s,i)=>[x(i),y(pct(s.ok,s.n)),s]);
 const f=pct(P[0].ok,P[0].n),e=pct(P[P.length-1].ok,P[P.length-1].n);
 el.innerHTML=`<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" style="display:block;max-width:100%" role="img" aria-label="Динаміка відсотка виконання">${[0,50,100].map(v=>`<line x1="${l}" x2="${W-r}" y1="${y(v)}" y2="${y(v)}" stroke="var(--bd)"/><text x="${l-6}" y="${y(v)+4}" text-anchor="end" font-size="11" fill="var(--mu)">${v}%</text>`).join('')}<polyline fill="none" stroke="var(--ok)" stroke-width="3" stroke-linejoin="round" points="${pts.map(p=>p[0]+','+p[1]).join(' ')}"/>${pts.map(p=>`<circle cx="${p[0]}" cy="${p[1]}" r="4.5" fill="var(--ok)"><title>${fmtD(p[2].date)}: ${p[2].ok} з ${p[2].n} (${pct(p[2].ok,p[2].n)}%)</title></circle>`).join('')}<text x="${l}" y="${H-5}" font-size="11" fill="var(--mu)">${fmtD(P[0].date)}</text><text x="${W-r}" y="${H-5}" text-anchor="end" font-size="11" fill="var(--mu)">${fmtD(P[P.length-1].date)}</text></svg><p class="mu" style="margin:6px 0 0">З ${fmtD(P[0].date)} до ${fmtD(P[P.length-1].date)}: ${f}% → ${e}% (${e-f>=0?'+':''}${e-f} в.п.).</p>`}
function renderDl(){const L=dlList(),c=k=>L.filter(t=>t.k===k).length,ok=c('ok'),bad=c('bad'),w=c('wait');
 $('dlk').innerHTML=kc(L.length,'строкових завдань')+kc(ok,'виконано ('+pct(ok,L.length)+'%)')+kc(bad,'прострочено',bad?'bd':'')+kc(w,'не виконано, строк не минув');
 $('dlb').innerHTML=L.map(t=>`<tr><td>${t.r}</td><td>${esc(t.t)}</td><td class="nw">${esc(t.term)}</td><td>${esc(t.u)}${t.co.length?'<div class="mu">+ '+esc(t.co.join('; '))+'</div>':''}</td><td><span class="sb ${t.ok?'ok':'bad'}">${t.ok?'ВИКОНАНО':'НЕ ВИКОНАНО'}</span></td><td><span class="b ${t.k==='ok'?'ok':t.k==='bad'?'bd2':'no'}">${esc(t.l)}</span></td></tr>`).join('');
 $('dls').textContent=L.length?'— '+L.length+' пунктів із конкретним місяцем або датою':'— таких завдань не знайдено'}

function filtered(){const q=$('q').value.toLowerCase(),fs=$('fs').value,fc=$('fsec').value,fu=$('fu').value;
 return T.filter(t=>(!q||t.t.toLowerCase().includes(q))&&(fs===''||(+t.ok)==fs)&&(!fc||t.s==fc)&&(!fu||t.u===fu)&&(!extra||extra(t)))}
function render(){const L=filtered();$('cnt').textContent='— показано '+L.length+' з '+T.length+extraLabel;
 $('tb').innerHTML=L.map(t=>{const o=open.has(t.id),s=o?t.t:cut(t.t,110);
  return `<tr class="t" data-id="${esc(t.id)}"><td>${esc(t.id)}</td><td>${esc(s)}${o?(t.co.length?'<div class="mu">Співвиконавці: '+esc(t.co.join('; '))+'</div>':'')+'<div class="inf"><b>Інформація виконання:</b> '+(t.info?esc(t.info):'<span class="mu">не заповнено</span>')+'</div>':''}</td><td>${esc(t.term)}</td><td>${esc(t.u)}</td><td><span class="b ${t.ok?'ok':'no'}">${t.ok?'Виконано':'Не виконано'}</span></td></tr>`}).join('')}
function focusTable(){$('all').scrollIntoView({behavior:'smooth',block:'start'})}
function reset(){extra=null;extraLabel='';$('q').value='';$('fs').value='';$('fsec').value='';$('fu').value='';render()}
document.addEventListener('click',e=>{const x=e.target.closest('[data-sec],[data-unit],[data-act],tr.t');if(!x)return;
 if(x.dataset.sec){reset();$('fsec').value=x.dataset.sec;render();focusTable()}
 else if(x.dataset.unit){reset();$('fu').value=x.dataset.unit;render();focusTable()}
 else if(x.dataset.act==='nd'){reset();$('fs').value='0';render();focusTable()}
 else if(x.dataset.act==='all'){reset();render();focusTable()}
 else if(x.dataset.act==='ann'){reset();$('fs').value='0';render();focusTable()}
 else if(x.dataset.act==='z0'){reset();const z=[...new Set(T.map(t=>t.s))].filter(s=>!T.some(t=>t.s===s&&t.ok));extra=t=>z.includes(t.s);extraLabel=' · розділи без виконаних: '+z.join(', ');render();focusTable()}
 else if(x.dataset.act==='u100'){const f=new Set(Object.entries(T.reduce((m,t)=>((m[t.u]=m[t.u]||[]).push(t),m),{})).filter(([,a])=>a.every(t=>t.ok)).map(([u])=>u));$('unit').scrollIntoView({behavior:'smooth',block:'center'});document.querySelectorAll('#unit .row').forEach(e=>{if(f.has(e.dataset.unit)){e.classList.add('hl100');setTimeout(()=>e.classList.remove('hl100'),3000)}})}
 else if(x.matches('tr.t')){const i=x.dataset.id;open.has(i)?open.delete(i):open.add(i);render()}});
['q','fs','fsec','fu'].forEach(i=>$(i).addEventListener('input',()=>{extra=null;extraLabel='';render()}));document.addEventListener('keydown',e=>{if((e.key==='Enter'||e.key===' ')&&e.target.dataset&&e.target.dataset.act){e.preventDefault();e.target.click()}});$('e0').onclick=reset;

const q=v=>'"'+String(v==null?'':v).replace(/"/g,'""')+'"';
function csv(L){return '\uFEFF'+[['Рядок','№','Розділ','Завдання','Строк','Підрозділ','Співвиконавці','Статус (G)','Інформація виконання'],...L.map(t=>[t.r,t.id,t.s,t.t,t.term,t.u,t.co.join('; '),t.ok?'ВИКОНАНО':'НЕ ВИКОНАНО',t.info])].map(r=>r.map(q).join(';')).join('\r\n')}
const msg=m=>{$('msg').textContent=m;setTimeout(()=>$('msg').textContent='',5000)};
$('e1').onclick=()=>save(filtered(),'zvit_2026_vidbir');$('e2').onclick=()=>save(T.filter(t=>!t.ok),'zvit_2026_nevykonani');$('e3').onclick=()=>save(dlList(),'zvit_2026_strokovi');
$('pr').onclick=()=>{try{window.print()}catch(e){msg('Для друку натисніть Ctrl+P')}};

let fname='';const LS='zvit2026_snaps';
try{snaps=Object.values(JSON.parse(localStorage.getItem(LS)||'{}'))}catch(e){}
function saveSnap(){try{const o=JSON.parse(localStorage.getItem(LS)||'{}'),n=T.length,ok=T.filter(t=>t.ok).length,d=dstr(TD);o[d]={date:d,n,ok,no:n-ok};localStorage.setItem(LS,JSON.stringify(o));snaps=Object.values(o);renderTrend()}catch(e){}}
async function save(L,name){const b=new Blob([csv(L)],{type:'text/csv;charset=utf-8'}),a=document.createElement('a');a.href=URL.createObjectURL(b);a.download=name+'_'+dstr(TD)+'.csv';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}
function parseCsv(s){const rows=[];let r=[],c='',q=false;for(let i=0;i<s.length;i++){const ch=s[i];if(q){if(ch==='"'){if(s[i+1]==='"'){c+='"';i++}else q=false}else c+=ch}else if(ch==='"')q=true;else if(ch===','){r.push(c);c=''}else if(ch==='\n'||ch==='\r'){if(ch==='\r'&&s[i+1]==='\n')i++;r.push(c);c='';rows.push(r);r=[]}else c+=ch}if(c!==''||r.length){r.push(c);rows.push(r)}return rows}
function useRows(rows,kind,name){const A=rows.map(r=>[r[0],r[1],r[2],r[3]]),G=rows.map(r=>[r[6]]),F=rows.map(r=>[r[5]]),t=parse(A,G,F);
 if(!t.length){err='завдань не знайдено — перевірте аркуш і стовпці A–G';status();return}err='';fname=name||'';setData(t,kind,Date.now());saveSnap()}
async function loadCsv(){const url=csvSource();if(!url){msg('Вкажіть sheetId у app.js або відкрийте xlsx');return}
 try{const r=await fetch(url,{cache:'no-store'});if(!r.ok)throw new Error(r.status);useRows(parseCsv(await r.text()),'csv')}catch(e){err='не вдалося прочитати таблицю (перевірте, що доступ відкрито для всіх, хто має посилання)';status()}}
$('rf').onclick=loadCsv;$('up').onclick=()=>$('fi').click();
$('fi').onchange=e=>{const f=e.target.files[0];if(!f)return;if(!window.XLSX){err='бібліотека SheetJS не завантажилась (потрібен інтернет)';status();return}
 const rd=new FileReader();rd.onload=ev=>{try{const wb=XLSX.read(ev.target.result,{type:'array'}),ws=wb.Sheets['Звіт 2026']||wb.Sheets[wb.SheetNames[0]],rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:''});
  rows.forEach(r=>{if(typeof r[0]==='number'&&r[0]>=40000){const d=new Date(Math.round((r[0]-25569)*864e5));r[0]=d.getUTCDate()+'.'+(d.getUTCMonth()+1)+'.'}});useRows(rows,'file',f.name)}catch(x){err='не вдалося прочитати xlsx';status()}};
 rd.readAsArrayBuffer(f);e.target.value=''};
status();build();let rz;addEventListener('resize',()=>{clearTimeout(rz);rz=setTimeout(renderTrend,150)});

if(csvSource()){loadCsv();setInterval(loadCsv,CONFIG.refreshMs)}
