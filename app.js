/* Дашборд виконання Плану роботи — локальна версія.
   Джерела даних (за пріоритетом): 1) Google-таблиця (CONFIG.sheetId або csvUrl),
   2) файл xlsx кнопкою «Відкрити xlsx», 3) вбудовані дані з data.js. */
const CONFIG={
  // Щоб сторінка могла читати таблицю, відкрийте доступ:
  // Google Sheets → Поділитися → Загальний доступ → «Усі, хто має посилання» → роль «Читач».
  sheetId:'10VAw8cK2dC4mgH7VdKZDtNlByxFEHazcjz06r7hT3bg',
  sheetName:'Звіт 2026',
  // Необов'язково: готове CSV-посилання («Опублікувати в Інтернеті»). Якщо заповнено — використовується замість sheetId.
  csvUrl:'',
  refreshMs:60000   // як часто оновлювати, мс
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
let SECN=Object.assign({},SECN0),coMode=false,sortK='',sortD=1,shown=25,chg=[],lastFocus=null,extra=null,extraLabel='',T=STATIC,src='static',at=0,open=new Set(),snaps=[],latest=null,A=null,G=null,Fv=null,err='',db=null,canW=false,dbA=false,dbB=false,persisted='',dlz=null,sampleFn=null,mcp=null;

const GRP='Керівники підрозділів';
const sn=s=>SECN[s]?('Розділ '+s+'. '+SECN[s]):'Розділ '+s;
const idKey=t=>t.s*1000+(+String(t.id).split('.')[1]);
function parse(A,G,F){const out=[],nm={};A.forEach((r,i)=>{const a=norm(r&&r[0]),sm=/^(?:Розділ\s+)?(\d+)\.\s*([^\d\s].*)$/i.exec(a);if(sm){nm[+sm[1]]=sm[2];return}
 const m=/^(\d+)\.(\d+)\.?$/.exec(a);if(!m)return;
 const resp=norm(r[3]),us=UN.map(([k,s])=>[resp.indexOf(k),s]).filter(x=>x[0]>=0).sort((a,b)=>a[0]-b[0]).map(x=>x[1]),g=norm(G[i]&&G[i][0]);
 out.push({r:i+1,id:m[1]+'.'+m[2],s:+m[1],t:norm(r[1]),term:norm(r[2]),u:us[0]||GRP,un:!us.length&&!/керівник/i.test(resp),co:us.slice(1),g,ok:g.toUpperCase()==='ВИКОНАНО',info:norm(F&&F[i]&&F[i][0]).slice(0,1500)})});
 if(Object.keys(nm).length)SECN=Object.assign({},SECN,nm);return out}
function unitMap(){const m={};T.forEach(t=>{(coMode?[t.u,...t.co]:[t.u]).forEach(u=>{if(u===GRP)return;(m[u]=m[u]||[]).push(t)})});return m}
const grpTasks=()=>T.filter(t=>t.u===GRP);
const MS=[['січ',0],['лют',1],['берез',2],['квіт',3],['трав',4],['черв',5],['лип',6],['серп',7],['верес',8],['жовт',9],['листоп',10],['груд',11]];
function dlDate(term){const s=term.toLowerCase();let best=null,cnt=0;MS.forEach(([k,m])=>{const re=new RegExp('(^|[^а-яіїєґ])'+k,'g');let x,f=0;while((x=re.exec(s))){f=1;if(!best||x.index>best.i)best={i:x.index,m}}cnt+=f});
 if(!best)return null;const d=/(\d{1,2})\s*[а-яіїєґ]+/.exec(s);
 return(d&&cnt===1&&s.indexOf('до')===0)?new Date(YEAR,best.m,+d[1],23,59,59):new Date(YEAR,best.m+1,0,23,59,59)}
function dlList(){const L=T.filter(t=>!/^протягом/i.test(t.term)).map(t=>({...t,d:dlDate(t.term)})).sort((a,b)=>(a.d||0)-(b.d||0)||a.r-b.r);
 L.forEach(t=>{t.dd=t.d?Math.round((new Date(t.d.getFullYear(),t.d.getMonth(),t.d.getDate())-TD)/DAY):null;if(t.ok){t.k='ok';t.l='Виконано'}else if(t.d&&t.d<TD){t.k='bad';t.l='Прострочено на '+(-t.dd)+' дн.'}else{t.k='wait';t.l=t.d?(t.dd===0?'Строк — сьогодні':'У строк, залишилось '+t.dd+' дн.'):'У строк'}});return L}

const row=(u,x)=>`<div class="row" data-unit="${esc(u)}" role="button" tabindex="0" aria-label="${esc(u)}: виконано ${x.ok} з ${x.n}. Відкрити картку підрозділу" title="${esc(u)}: ${x.ok} з ${x.n}. Натисніть, щоб відкрити картку"><div class="lbl">${esc(u)}</div><div class="bar"><b class="ok" style="width:${x.ok/x.n*100}%"></b><b class="no" style="width:${(x.n-x.ok)/x.n*100}%"></b></div><div class="pc">${pct(x.ok,x.n)}%<small>${x.ok}/${x.n}</small></div></div>`;
function fill(sel,vals,lab){const c=sel.value;sel.length=1;vals.forEach(v=>sel.insertAdjacentHTML('beforeend',`<option value="${esc(v)}">${esc(typeof lab==='function'?lab(v):(lab?lab+v:v))}</option>`));sel.value=c}
const kc=(v,l,cls,act)=>`<div class="c kc ${cls||''}" ${act?`data-act="${act}" role="button" tabindex="0" title="Натисніть, щоб переглянути"`:''}><div class="kv">${v}</div><div class="kl">${l}</div></div>`;

function setData(t,s,a){T=t;src=s;at=a||Date.now();status();build()}
function status(){const e=err?' · '+err:'';const L={csv:'Джерело: Google-таблиця (CSV) · оновлено '+fmtT(at),file:'Джерело: файл «'+fname+'» · завантажено '+fmtT(at),static:'Показано вбудовані дані (data.js)'};$('src').textContent=L[src]+e;$('src').className=src==='static'?'warn':'live'}

function build(){
 const n=T.length,ok=T.filter(t=>t.ok).length,no=n-ok,bySec={},um=unitMap();T.forEach(t=>{(bySec[t.s]=bySec[t.s]||[]).push(t)});
 const secs=Object.keys(bySec).sort((a,b)=>a-b),ua=Object.entries(um),full=ua.filter(([,a])=>a.every(t=>t.ok)).length,zero=secs.filter(s=>stat(bySec[s]).ok===0);
 $('hero').innerHTML=`<div class="ring" style="background:conic-gradient(var(--lime) 0 ${pct(ok,n)}%,#3a3f2c 0)"><i>${pct(ok,n)}%</i></div><div><div class="big">${ok} із ${n} завдань виконано</div><p>Не виконано ${no} із ${n} завдань.</p></div>`;
 $('kpi').innerHTML=kc(n,'завдань у плані','','all')+kc(no,'не виконано','hl','nd')+kc(full+' з '+ua.length,'підрозділів на 100%','','u100')+kc(zero.length,'розділів без виконаних завдань','','z0');
 $('dir').innerHTML=dirs.map(([l,a,b],i)=>{const x=stat(T.filter(t=>t.s>=a&&t.s<=b));return `<div class="c dc" data-dir="${i}" role="button" tabindex="0" aria-label="${esc(l)}: ${pct(x.ok,x.n)}% виконано. Показати завдання" title="Натисніть, щоб показати завдання напряму"><div class="t">${esc(l)}</div><div class="n">${pct(x.ok,x.n)}%</div><div class="bar" style="margin:6px 0"><b class="ok" style="width:${x.n?x.ok/x.n*100:0}%"></b><b class="no" style="width:${x.n?(x.n-x.ok)/x.n*100:0}%"></b></div><div class="mu">${x.ok} з ${x.n} завдань</div></div>`}).join('');
 const mx=Math.max(1,...secs.map(s=>bySec[s].length));
 $('sec').innerHTML=secs.map(s=>{const x=stat(bySec[s]);return `<div class="col" data-sec="${s}" role="button" tabindex="0" aria-label="${esc(sn(s))}: ${x.ok} з ${x.n} виконано. Показати завдання" title="${esc(sn(s))}: ${x.ok} з ${x.n} виконано. Натисніть, щоб відфільтрувати"><div class="st" style="height:${x.n/mx*88}%"><b class="ok" style="height:${x.ok/x.n*100}%"></b><b class="no" style="height:${(x.n-x.ok)/x.n*100}%"></b></div><span>${s}</span></div>`}).join('');
 $('unit').innerHTML=ua.sort((a,b)=>pct(stat(a[1]).ok,a[1].length)-pct(stat(b[1]).ok,b[1].length)).map(([u,a])=>row(u,stat(a))).join('');
 const g=grpTasks();$('grp').textContent=g.length?'Окремо від рейтингу: завдання для керівників усіх підрозділів — виконано '+g.filter(t=>t.ok).length+' з '+g.length+'.':'';
 fill($('fsec'),secs,s=>s+'. '+cut(SECN[s]||'',55));fill($('fu'),[...Object.keys(um).sort(),...(g.length?[GRP]:[])]);
 renderAtt();renderChanges();renderTrend();renderDl();renderMx();render();if(document.body.classList.contains('rpt'))renderReport();
}
function renderAtt(){
 const L=dlList(),bad=L.filter(t=>t.k==='bad'),soon=L.filter(t=>t.k==='wait'&&t.dd!==null&&t.dd<=60),ann=T.filter(t=>!t.ok&&/^протягом/i.test(t.term)),left=Math.round((new Date(YEAR,11,31)-TD)/DAY);
 const it=[...bad,...soon];
 $('att').innerHTML=`<h2>Потребує уваги</h2><div class="g k4">${kc(left,'днів до кінця '+YEAR+' року (показати невиконані)','','nd')}${kc(bad.length,'прострочених строкових завдань',bad.length?'bd':'','over')}${kc(soon.length,'строків у найближчі 60 днів','','soon')}${kc(ann.length,'річних завдань іще не виконано','', 'ann')}</div>`+
 (it.length?`<ul class="al">${it.map(t=>`<li data-task="${esc(t.id)}" role="button" tabindex="0" title="Натисніть, щоб відкрити завдання в таблиці"><b>${esc(t.id)}</b><div>${esc(cut(t.t,150))}<div class="mu">${esc(t.u)} · строк: ${esc(t.term)}</div></div><span class="b ${t.k==='bad'?'bd2':'no'}">${esc(t.l)}</span></li>`).join('')}</ul>`:'<p class="mu" style="margin:12px 0 0">Прострочених строкових завдань і строків у найближчі 60 днів немає.</p>')+
 `<p class="mu" style="margin:12px 0 0;font-size:12px">Завдання «Протягом року» вважаються такими, що мають бути виконані до 31 грудня ${YEAR}.</p>`;
}
function renderTrend(){const P=snaps.filter(s=>s.n).sort((a,b)=>a.date<b.date?-1:1),el=$('trend');
 $('trendc').hidden=P.length<2;if(P.length<2){el.innerHTML='';return}
 if(!el.clientWidth)return;const W=Math.max(300,el.clientWidth),H=170,l=34,r=10,t=12,b=24,x=i=>l+i/(P.length-1)*(W-l-r),y=v=>t+(100-v)/100*(H-t-b),pts=P.map((s,i)=>[x(i),y(pct(s.ok,s.n)),s]);
 const f=pct(P[0].ok,P[0].n),e=pct(P[P.length-1].ok,P[P.length-1].n);
 el.innerHTML=`<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" style="display:block;max-width:100%" role="img" aria-label="Динаміка відсотка виконання">${[0,50,100].map(v=>`<line x1="${l}" x2="${W-r}" y1="${y(v)}" y2="${y(v)}" stroke="var(--bd)"/><text x="${l-6}" y="${y(v)+4}" text-anchor="end" font-size="11" fill="var(--mu)">${v}%</text>`).join('')}<polyline fill="none" stroke="var(--ok)" stroke-width="3" stroke-linejoin="round" points="${pts.map(p=>p[0]+','+p[1]).join(' ')}"/>${pts.map(p=>`<circle cx="${p[0]}" cy="${p[1]}" r="4.5" fill="var(--ok)"><title>${fmtD(p[2].date)}: ${p[2].ok} з ${p[2].n} (${pct(p[2].ok,p[2].n)}%)</title></circle>`).join('')}<text x="${l}" y="${H-5}" font-size="11" fill="var(--mu)">${fmtD(P[0].date)}</text><text x="${W-r}" y="${H-5}" text-anchor="end" font-size="11" fill="var(--mu)">${fmtD(P[P.length-1].date)}</text></svg><p class="mu" style="margin:6px 0 0">З ${fmtD(P[0].date)} до ${fmtD(P[P.length-1].date)}: ${f}% → ${e}% (${e-f>=0?'+':''}${e-f} в.п.).</p>`}
function renderDl(){const L=dlList(),c=k=>L.filter(t=>t.k===k).length,ok=c('ok'),bad=c('bad'),w=c('wait');
 $('dlk').innerHTML=kc(L.length,'строкових завдань','','dl')+kc(ok,'виконано ('+pct(ok,L.length)+'%)','','dlok')+kc(bad,'прострочено',bad?'bd':'','over')+kc(w,'не виконано, строк не минув','','wait');
 $('dlb').innerHTML=L.map(t=>`<tr class="tk" data-task="${esc(t.id)}" role="button" tabindex="0" title="Натисніть, щоб відкрити завдання в таблиці"><td>${t.r}</td><td>${esc(t.t)}</td><td class="nw">${esc(t.term)}</td><td>${esc(t.u)}${t.co.length?'<div class="mu">+ '+esc(t.co.join('; '))+'</div>':''}</td><td><span class="sb ${t.ok?'ok':'bad'}">${t.ok?'ВИКОНАНО':'НЕ ВИКОНАНО'}</span></td><td><span class="b ${t.k==='ok'?'ok':t.k==='bad'?'bd2':'no'}">${esc(t.l)}</span></td></tr>`).join('');
 $('dls').textContent=L.length?'— '+L.length+' пунктів із конкретним місяцем або датою':'— таких завдань не знайдено';renderCal(L)}

function renderCal(L){const M=['Січ','Лют','Бер','Кві','Тра','Чер','Лип','Сер','Вер','Жов','Лис','Гру'],cm=TD.getFullYear()===YEAR?TD.getMonth():-1,ann=T.filter(t=>/^протягом/i.test(t.term));
 $('cal').innerHTML=`<div class="cal">${M.map((m,i)=>`<div class="cm${i===cm?' now':''}"><div class="mh">${m}${i===cm?' · сьогодні':''}</div>${L.filter(t=>t.d&&t.d.getMonth()===i).map(t=>`<span class="chip ${t.k}" data-task="${esc(t.id)}" role="button" tabindex="0" title="${esc(t.id+': '+t.t+' — '+t.term+' ('+t.l+')')}">${esc(t.id)}</span>`).join('')}</div>`).join('')}</div><p class="mu" style="margin:8px 0 14px;font-size:12px">Завдання «Протягом року»: ${ann.length}, із них не виконано ${ann.filter(t=>!t.ok).length}. Колір позначки: зелений — виконано, помаранчевий — строк попереду, червоний — прострочено.</p>`}
function renderMx(){const um=unitMap(),secs=[...new Set(T.map(t=>t.s))].sort((a,b)=>a-b),rows=Object.entries(um).sort((a,b)=>pct(stat(a[1]).ok,a[1].length)-pct(stat(b[1]).ok,b[1].length)),g=grpTasks();if(g.length)rows.push([GRP,g]);
 $('mx').innerHTML=`<thead><tr><th class="mxu">Підрозділ</th>${secs.map(s=>`<th class="mxh" title="${esc(sn(s))}">${s}</th>`).join('')}</tr></thead><tbody>${rows.map(([u,a])=>`<tr><th class="mxu" scope="row">${esc(u)}</th>${secs.map(s=>{const c=a.filter(t=>t.s===s);if(!c.length)return '<td class="mx0"></td>';const nd=c.filter(t=>!t.ok).length,cl=!nd?'mxok':nd===c.length?'mxno':'mxpt',lab=u+', розділ '+s+': виконано '+(c.length-nd)+' з '+c.length;return `<td class="${cl}" data-cell="${esc(u)}|${s}" role="button" tabindex="0" title="${esc(lab)}" aria-label="${esc(lab)}">${nd?nd:'✓'}</td>`}).join('')}</tr>`).join('')}</tbody>`}
function renderChanges(){const ev=[];chg.slice().sort((a,b)=>b.at-a.at).forEach(c=>(c.items||[]).forEach(i=>ev.push({at:c.at,id:i.id,to:i.to})));$('chc').hidden=!ev.length;if(!ev.length)return;
 const tm=Object.fromEntries(T.map(t=>[t.id,t]));
 $('chg').innerHTML=`<ul class="al2">${ev.slice(0,12).map(i=>{const t=tm[i.id];return `<li data-task="${esc(i.id)}" role="button" tabindex="0" title="Натисніть, щоб відкрити завдання в таблиці"><b>${esc(i.id)}</b><div>${esc(cut(t?t.t:'',130))}<div class="mu">${fmtT(i.at)}${t?' · '+esc(t.u):''}</div></div><span class="b ${i.to?'ok':'no'}">${i.to?'Стало виконано':'Знято виконання'}</span></li>`}).join('')}</ul>`}
function openUnit(u,from){lastFocus=from||document.activeElement;const um=unitMap(),tasks=u===GRP?grpTasks():(um[u]||[]),x=stat(tasks),nd=tasks.filter(t=>!t.ok),dm=Object.fromEntries(dlList().map(t=>[t.id,t])),coN=T.filter(t=>t.co.includes(u)).length;
 const li=t=>{const d=dm[t.id];return `<li data-task="${esc(t.id)}" role="button" tabindex="0"><b>${esc(t.id)}</b><div>${esc(cut(t.t,160))}<div class="mu">${esc(t.term)}${d&&!t.ok?(d.k==='bad'?' · <span class="b bd2">'+esc(d.l)+'</span>':d.dd!==null?' · '+esc(d.l):''):''}</div></div></li>`};
 $('pn').innerHTML=`<button class="x" data-close aria-label="Закрити">✕</button><h2>${esc(u)}</h2><div class="pbig">${pct(x.ok,x.n)}%<span class="mu"> виконано · ${x.ok} з ${x.n}</span></div><div class="bar" style="margin:8px 0 14px"><b class="ok" style="width:${x.n?x.ok/x.n*100:0}%"></b><b class="no" style="width:${x.n?(x.n-x.ok)/x.n*100:0}%"></b></div>`
 +(coN&&u!==GRP?`<p class="mu">Також співвиконавець у ${coN} завданнях${coMode?' (враховано в цифрах вище)':''}.</p>`:'')
 +`<h3>Не виконано (${nd.length})</h3>`+(nd.length?`<ul class="al2">${nd.map(li).join('')}</ul>`:'<p class="mu">Невиконаних завдань немає.</p>')
 +`<details><summary>Виконано (${x.ok})</summary><ul class="al2">${tasks.filter(t=>t.ok).map(li).join('')}</ul></details><p><button data-ptable="${esc(u)}">Показати в таблиці</button></p>`;
 $('pov').hidden=false;document.body.classList.add('nsc');$('pn').focus()}
function closePanel(){$('pov').hidden=true;document.body.classList.remove('nsc');if(lastFocus&&lastFocus.focus)lastFocus.focus()}
function renderReport(){const n=T.length,ok=T.filter(t=>t.ok).length,nd=T.filter(t=>!t.ok),L=dlList(),bad=L.filter(t=>t.k==='bad'),left=Math.round((new Date(YEAR,11,31)-TD)/DAY),um=unitMap(),secs=[...new Set(T.map(t=>t.s))].sort((a,b)=>a-b),zero=secs.filter(s=>!T.some(t=>t.s===s&&t.ok)),weak=Object.entries(um).filter(([,a])=>a.some(t=>!t.ok)).sort((a,b)=>pct(stat(a[1]).ok,a[1].length)-pct(stat(b[1]).ok,b[1].length));
 const risks=['До кінця '+YEAR+' року залишилось '+left+' дн., не виконано '+nd.length+' із '+n+' завдань.',...bad.map(t=>'Прострочено: '+t.id+' — '+cut(t.t,110)+' ('+t.term+', '+t.l.toLowerCase()+').'),...(zero.length?['Розділи без виконаних завдань: '+zero.map(s=>s+'. '+cut(SECN[s]||'',50)).join('; ')+'.']:[]),...(weak.length?['Підрозділи з невиконаними завданнями: '+weak.map(([u,a])=>u+' ('+stat(a).ok+'/'+a.length+')').join('; ')+'.']:[])];
 $('report').innerHTML=`<div class="rb"><button id="rp">Друкувати / зберегти PDF</button> <button id="rc">Закрити звіт</button></div><h1>Звіт про виконання Плану роботи на ${YEAR} рік</h1><p class="mu">Державна служба України у справах дітей · станом на ${TD.toLocaleDateString('uk-UA')}</p>
 <div class="g k4">${kc(n,'завдань у плані')}${kc(ok+' ('+pct(ok,n)+'%)','виконано')}${kc(nd.length,'не виконано')}${kc(bad.length,'прострочено строкових',bad.length?'bd':'')}</div>
 <h2 style="margin-top:20px">За пріоритетними напрямами</h2><table class="rt"><thead><tr><th>Напрям</th><th>Виконано</th><th>Не виконано</th><th>%</th></tr></thead><tbody>${dirs.map(([l,a,b])=>{const x=stat(T.filter(t=>t.s>=a&&t.s<=b));return `<tr><td>${esc(l)}</td><td>${x.ok}</td><td>${x.n-x.ok}</td><td>${pct(x.ok,x.n)}%</td></tr>`}).join('')}</tbody></table>
 <h2 style="margin-top:20px">Ризики</h2><ul>${risks.map(x=>`<li>${esc(x)}</li>`).join('')}</ul>
 <h2 style="margin-top:20px">Невиконані завдання (${nd.length})</h2><table class="rt"><thead><tr><th>№</th><th>Завдання</th><th>Підрозділ</th><th>Строк</th></tr></thead><tbody>${nd.map(t=>`<tr><td>${esc(t.id)}</td><td>${esc(cut(t.t,170))}</td><td>${esc(t.u)}</td><td>${esc(t.term)}</td></tr>`).join('')}</tbody></table>`}
const SK={id:idKey,term:t=>(dlDate(t.term)||EOY).getTime(),u:t=>t.u,st:t=>+t.ok};
function filtered(){const q=$('q').value.toLowerCase(),fs=$('fs').value,fc=$('fsec').value,fu=$('fu').value;
 let L=T.filter(t=>(!q||t.t.toLowerCase().includes(q))&&(fs===''||(+t.ok)==fs)&&(!fc||t.s==fc)&&(!fu||t.u===fu||(coMode&&t.co.includes(fu)))&&(!extra||extra(t)));
 if(sortK){const f=SK[sortK];L=[...L].sort((a,b)=>{const x=f(a),y=f(b);return((typeof x==='string'?x.localeCompare(y,'uk'):x-y)*sortD)||idKey(a)-idKey(b)})}return L}
function render(){const L=filtered(),V=L.slice(0,shown);$('cnt').textContent='— показано '+V.length+' з '+L.length+(L.length!==T.length?' (усього '+T.length+')':'')+extraLabel;
 $('tb').innerHTML=V.map(t=>{const o=open.has(t.id),s=o?t.t:cut(t.t,110);
  return `<tr class="t" data-id="${esc(t.id)}" tabindex="0" role="button" aria-expanded="${o}"><td>${esc(t.id)}</td><td>${esc(s)}${o?'<div class="mu">'+esc(sn(t.s))+'</div>'+(t.co.length?'<div class="mu">Співвиконавці: '+esc(t.co.join('; '))+'</div>':'')+'<div class="inf"><b>Інформація виконання:</b> '+(t.info?esc(t.info):'<span class="mu">не заповнено</span>')+'</div>':''}</td><td>${esc(t.term)}</td><td>${esc(t.u)}</td><td><span class="b ${t.ok?'ok':'no'}">${t.ok?'Виконано':'Не виконано'}</span></td></tr>`}).join('');
 $('mr').innerHTML=L.length>shown?`<button data-more="25">Показати ще ${Math.min(25,L.length-shown)}</button> <button data-more="all">Показати всі (${L.length})</button>`:'';
 document.querySelectorAll('th[data-sort]').forEach(th=>{const a=th.dataset.sort===sortK;th.setAttribute('aria-sort',a?(sortD>0?'ascending':'descending'):'none');th.querySelector('i').textContent=a?(sortD>0?' ▲':' ▼'):''})}
function focusTable(){$('all').scrollIntoView({behavior:'smooth',block:'start'})}
function reset(){extra=null;extraLabel='';shown=25;$('q').value='';$('fs').value='';$('fsec').value='';$('fu').value='';render()}
const go=()=>{shown=25;render();focusTable()};
document.addEventListener('click',e=>{const tg=e.target;
 if(tg.id==='rc'){document.body.classList.remove('rpt');renderTrend();return}
 if(tg.id==='rp'){try{window.print()}catch(x){}msg('Якщо вікно друку не відкрилось, натисніть Ctrl+P (на Mac — Cmd+P)');return}
 const mo=tg.closest('[data-more]');if(mo){shown=mo.dataset.more==='all'?1e9:shown+25;render();return}
 const so=tg.closest('th[data-sort]');if(so){const k=so.dataset.sort;if(sortK===k)sortD=-sortD;else{sortK=k;sortD=1}shown=25;render();return}
 if(tg.closest('[data-close]')){closePanel();return}
 const pt=tg.closest('[data-ptable]');if(pt){const u=pt.dataset.ptable;closePanel();reset();$('fu').value=u;go();return}
 const x=tg.closest('[data-task],[data-sec],[data-unit],[data-dir],[data-cell],[data-act],tr.t');if(!x)return;
 if(x.dataset.task){const id=x.dataset.task;if(!$('pov').hidden)closePanel();reset();open.add(id);extra=t=>t.id===id;extraLabel=' · завдання '+id;go()}
 else if(x.dataset.sec){reset();$('fsec').value=x.dataset.sec;go()}
 else if(x.dataset.unit){openUnit(x.dataset.unit,x)}
 else if(x.dataset.dir){const [l,a,b]=dirs[+x.dataset.dir];reset();extra=t=>t.s>=a&&t.s<=b;extraLabel=' · напрям: '+l;go()}
 else if(x.dataset.cell){const [u,s]=x.dataset.cell.split('|');reset();$('fu').value=u;$('fsec').value=s;go()}
  else if(x.dataset.act==='nd'){reset();$('fs').value='0';go()}
 else if(x.dataset.act==='done'){reset();$('fs').value='1';go()}
 else if(['ann','over','soon','dl','dlok','wait'].includes(x.dataset.act)){const a=x.dataset.act,dl=dlList(),bad=new Set(dl.filter(t=>t.k==='bad').map(t=>t.id)),wt=new Set(dl.filter(t=>t.k==='wait').map(t=>t.id)),sn2=new Set(dl.filter(t=>t.k==='wait'&&t.dd!==null&&t.dd<=60).map(t=>t.id)),pr=t=>/^протягом/i.test(t.term);reset();
  const F={ann:[t=>!t.ok&&pr(t),'річні завдання без виконання'],over:[t=>bad.has(t.id),'прострочені строкові завдання'],soon:[t=>sn2.has(t.id),'строки в найближчі 60 днів'],dl:[t=>!pr(t),'строкові завдання'],dlok:[t=>!pr(t)&&t.ok,'виконані строкові завдання'],wait:[t=>wt.has(t.id),'не виконані, строк не минув']}[a];extra=F[0];extraLabel=' · '+F[1];go()}
 else if(x.dataset.act==='all'){reset();go()}
 else if(x.dataset.act==='z0'){reset();const z=[...new Set(T.map(t=>t.s))].filter(s=>!T.some(t=>t.s===s&&t.ok));extra=t=>z.includes(t.s);extraLabel=' · розділи без виконаних: '+z.join(', ');go()}
 else if(x.dataset.act==='u100'){const f=new Set(Object.entries(unitMap()).filter(([,a])=>a.every(t=>t.ok)).map(([u])=>u));$('unit').scrollIntoView({behavior:'smooth',block:'center'});document.querySelectorAll('#unit .row').forEach(el=>{if(f.has(el.dataset.unit)){el.classList.add('hl100');setTimeout(()=>el.classList.remove('hl100'),3000)}})}
 else if(x.matches('tr.t')){const i=x.dataset.id;open.has(i)?open.delete(i):open.add(i);render()}});
['q','fs','fsec','fu'].forEach(i=>$(i).addEventListener('input',()=>{extra=null;extraLabel='';shown=25;render()}));
document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!$('pov').hidden){closePanel();return}
 if((e.key==='Enter'||e.key===' ')&&e.target.getAttribute&&e.target.getAttribute('role')==='button'){e.preventDefault();e.target.click()}});
$('co').onchange=e=>{coMode=e.target.checked;build()};$('e0').onclick=reset;
const TH=['auto','light','dark'],THL={auto:'Тема: авто',light:'Тема: світла',dark:'Тема: темна'};
function applyTheme(m){if(m==='auto')document.documentElement.removeAttribute('data-theme');else document.documentElement.setAttribute('data-theme',m);$('th').textContent=THL[m]}
function initTheme(){let m='auto';try{m=localStorage.getItem('zvit_theme')||'auto'}catch(e){}if(!THL[m])m='auto';applyTheme(m);$('th').onclick=()=>{m=TH[(TH.indexOf(m)+1)%3];applyTheme(m);try{localStorage.setItem('zvit_theme',m)}catch(e){}}}

const q=v=>'"'+String(v==null?'':v).replace(/"/g,'""')+'"';
function csv(L){return '\uFEFF'+[['Рядок','№','Розділ','Завдання','Строк','Підрозділ','Співвиконавці','Статус (G)','Інформація виконання'],...L.map(t=>[t.r,t.id,t.s,t.t,t.term,t.u,t.co.join('; '),t.ok?'ВИКОНАНО':'НЕ ВИКОНАНО',t.info])].map(r=>r.map(q).join(';')).join('\r\n')}
const msg=m=>{$('msg').textContent=m;setTimeout(()=>$('msg').textContent='',5000)};
$('rf').onclick=()=>msg('Оновлення працює у вікні claude.ai з підключеним Google Sheets');
$('e1').onclick=()=>save(filtered(),'zvit_2026_vidbir');$('e2').onclick=()=>save(T.filter(t=>!t.ok),'zvit_2026_nevykonani');$('e3').onclick=()=>save(dlList(),'zvit_2026_strokovi');
$('pr').onclick=()=>{renderReport();document.body.classList.add('rpt');window.scrollTo(0,0)};

let fname='';const LS='zvit2026_snaps';
try{snaps=Object.values(JSON.parse(localStorage.getItem(LS)||'{}'))}catch(e){}
function saveSnap(){try{const o=JSON.parse(localStorage.getItem(LS)||'{}'),n=T.length,ok=T.filter(t=>t.ok).length,d=dstr(TD);o[d]={date:d,n,ok,no:n-ok};localStorage.setItem(LS,JSON.stringify(o));snaps=Object.values(o);renderTrend()}catch(e){}}
async function save(L,name){const b=new Blob([csv(L)],{type:'text/csv;charset=utf-8'}),a=document.createElement('a');a.href=URL.createObjectURL(b);a.download=name+'_'+dstr(TD)+'.csv';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}
function parseCsv(s){const rows=[];let r=[],c='',q=false;for(let i=0;i<s.length;i++){const ch=s[i];if(q){if(ch==='"'){if(s[i+1]==='"'){c+='"';i++}else q=false}else c+=ch}else if(ch==='"')q=true;else if(ch===','){r.push(c);c=''}else if(ch==='\n'||ch==='\r'){if(ch==='\r'&&s[i+1]==='\n')i++;r.push(c);c='';rows.push(r);r=[]}else c+=ch}if(c!==''||r.length){r.push(c);rows.push(r)}return rows}
function trackLocal(){try{const prev=JSON.parse(localStorage.getItem('zvit2026_prev')||'null'),cur={};T.forEach(t=>cur[t.id]=t.ok?1:0);
 if(prev){const items=T.filter(t=>t.id in prev&&prev[t.id]!==cur[t.id]).map(t=>({id:t.id,to:!!t.ok}));if(items.length){const l=JSON.parse(localStorage.getItem('zvit2026_chg')||'[]');l.unshift({at:Date.now(),items});localStorage.setItem('zvit2026_chg',JSON.stringify(l.slice(0,50)))}}
 localStorage.setItem('zvit2026_prev',JSON.stringify(cur));chg=JSON.parse(localStorage.getItem('zvit2026_chg')||'[]');renderChanges()}catch(e){}}
try{chg=JSON.parse(localStorage.getItem('zvit2026_chg')||'[]')}catch(e){}
function useRows(rows,kind,name){const A=rows.map(r=>[r[0],r[1],r[2],r[3]]),G=rows.map(r=>[r[6]]),F=rows.map(r=>[r[5]]),t=parse(A,G,F);
 if(!t.length){err='завдань не знайдено — перевірте аркуш і стовпці A–G';status();return}err='';fname=name||'';setData(t,kind,Date.now());saveSnap();trackLocal()}
async function loadCsv(){const url=csvSource();if(!url){msg('Вкажіть sheetId у app.js або відкрийте xlsx');return}
 try{const r=await fetch(url,{cache:'no-store'});if(!r.ok)throw new Error(r.status);useRows(parseCsv(await r.text()),'csv')}catch(e){err='не вдалося прочитати таблицю (перевірте, що доступ відкрито для всіх, хто має посилання)';status()}}
$('rf').onclick=loadCsv;$('up').onclick=()=>$('fi').click();
$('fi').onchange=e=>{const f=e.target.files[0];if(!f)return;if(!window.XLSX){err='бібліотека SheetJS не завантажилась (потрібен інтернет)';status();return}
 const rd=new FileReader();rd.onload=ev=>{try{const wb=XLSX.read(ev.target.result,{type:'array'}),ws=wb.Sheets['Звіт 2026']||wb.Sheets[wb.SheetNames[0]],rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:''});
  rows.forEach(r=>{if(typeof r[0]==='number'&&r[0]>=40000){const d=new Date(Math.round((r[0]-25569)*864e5));r[0]=d.getUTCDate()+'.'+(d.getUTCMonth()+1)+'.'}});useRows(rows,'file',f.name)}catch(x){err='не вдалося прочитати xlsx';status()}};
 rd.readAsArrayBuffer(f);e.target.value=''};
status();build();let rz;addEventListener('resize',()=>{clearTimeout(rz);rz=setTimeout(renderTrend,150)});initTheme();

if(csvSource()){loadCsv();setInterval(loadCsv,CONFIG.refreshMs)}
