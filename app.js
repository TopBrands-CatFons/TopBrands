/* ═══════════════════════════════════════════════
   1. CONFIGURACIÓ I DADES
═══════════════════════════════════════════════ */
let ADMIN_PW_MASTER='cuponsTB'; // valor per defecte; es sobreescriu des de Neon a init()
let ADMIN_PWS={}; // {nom: contrasenya} — un Admin, una contrasenya pròpia, assignada pel Master
let myAdminPw=''; // contrasenya amb la qual l'Admin actual ha entrat (per detectar revocacions)
let authedAs=''; // nom amb el qual s'ha superat el checkPw() actual
let adminLevel=''; // '' | 'admin' | 'master'
let A=['Admin'];
let pins={}; // {nom Presentador Fitxa/Tramitador dia-D: PIN}
let unlockedNames=new Set(); // noms desbloquejats amb PIN durant aquesta sessió
let G=['JE','JR','OG'];
let PF=['AC','AG','AM','AMO','CR','GP','JE','JP','JR','MO','OG','PL','PO','RL']; // Presentadors Fitxa: genera el botó de rol
let TD=[...PF]; // Tramitadors dia-D: genera el botó de rol (llista independent de PF)
let TF3=[...PF];
let TFH=['AG']; // Equip Hora: genera el botó del rol Hora
let QW=[...PF]; // Equip Web: genera el botó del rol Web
let WEBLIST=['SI','NO','SBY','PROB'];
let CUPS=['','CEXP','PI','EC','IA'];
let SITS=['','COMPLETAT','PROCÉS','POTENCIAL','HO DESCARTA','FA COMPE','NO COMPLEIX','NUL'];
let OTORGAT_OPTS=['','SI','NO','REQUERIT'];
let CONSULTORS=[]; // llista d'emails de consultors per al desplegable "Email consultor" a Requeriments
let CONS=['AMO','MO','PL']; // llista d'inicials de consultors per al desplegable "Consultor" a Comercial

const SIT_COLORS={
  'COMPLETAT':  {bg:'#00B050',fg:'#fff'},
  'PROCÉS':  {bg:'#DAF2D0',fg:'#222'},
  'POTENCIAL':  {bg:'#FFEB9C',fg:'#222'},
  'HO DESCARTA':{bg:'#44B3E1',fg:'#fff'},
  'FA COMPE':   {bg:'#C0E6F5',fg:'#222'},
  'NO COMPLEIX':{bg:'#F4CCCC',fg:'#000'},
  'NUL':        {bg:'#FF0000',fg:'#fff'},
};
const SIT_ORDER={'COMPLETAT':0,'PROCÉS':1,'POTENCIAL':2,'FA COMPE':3,'HO DESCARTA':4,'NO COMPLEIX':5,'NUL':6};
const ITA_COLORS={
  '':         {bg:'',fg:'#aaa'},
  'No aplica':{bg:'',fg:'#aaa'},
  'DEMANAR':  {bg:'#F4CCCC',fg:'#000'},
  'DEMANAT':  {bg:'#FFEB9C',fg:'#222'},
  'REBUT':    {bg:'#DAF2D0',fg:'#222'},
};
const WEB_COLORS={'NO':{bg:'#F4CCCC',fg:'#000'},'SBY':{bg:'#FFEB9C',fg:'#222'}};
const CUP_COLORS={CEXP:{bg:'#EEEDFE',fg:'#534AB7'},PI:{bg:'#E1F5EE',fg:'#0F6E56'},EC:{bg:'#E6F1FB',fg:'#185FA5'},IA:{bg:'#FAEEDA',fg:'#854F0B'}};
const OTORGAT_ESTAT={
  'SI':{label:'ATORGAT',bg:'#A02B93',fg:'#fff'},
  'NO':{label:'REFUSAT',bg:'#808080',fg:'#fff'},
  'REQUERIT':{label:'REQUERIT',bg:'#F2CEED',fg:'#000'},
};
const REQ_ESTAT_COLORS={
  'PENDENT PRESENTACIÓ':{bg:'#FFD966',fg:'#3d2b00'},
  'PRESENTAT':          {bg:'#C1F0C8',fg:'#163d1f'},
  'PERDUT':             {bg:'#EA9999',fg:'#4a0000'},
  'RESOLT':             {bg:'#00B050',fg:'#fff'},
  'ATORGAT':            {bg:'#A02B93',fg:'#fff'},
};
const RESOLUCIO_TO_ESTAT={'':'','CONCEDIT':'ATORGAT','RESOLT':'RESOLT','REFUSAT':'PERDUT','INADMISSIÓ':'PERDUT'};
const RESOLUCIO_COLORS={
  'CONCEDIT':  {bg:'#00B050',fg:'#fff'},
  'RESOLT':    {bg:'#00B050',fg:'#fff'},
  'INADMISSIÓ':{bg:'#EA9999',fg:'#4a0000'},
  'REFUSAT':   {bg:'#EA9999',fg:'#4a0000'},
};
let REQ_ESTAT_OPTS=['','PENDENT PRESENTACIÓ','PRESENTAT','PERDUT','RESOLT','ATORGAT','No aplica'];
let RESOLUCIO_OPTS=['','CONCEDIT','RESOLT','INADMISSIÓ','REFUSAT'];
const SCH_HORES=(()=>{const h=[];for(let i=9;i<=15;i++){h.push(`${String(i).padStart(2,'0')}:00`);if(i<15)h.push(`${String(i).padStart(2,'0')}:30`);}return h;})();

let D=[];
let nid=1;
let cU='Admin', cT='a';
let fCup='', fSit='', fQ='', fWeb=false, fHora=false, fNova=false;
let fComOtor='';
let fComPendent='';
let fComKam='', fComCup='', fComCons='';
let comSortField=null, comSortDir=1; // ordenació manual (KAM/Empresa/Cupó/Consultor) a Comercial
let notesExpanded=new Set();
let reqNotesExpanded=new Set();
let comNotesExpanded=new Set();
let collapsedGroups=new Set(['PROCÉS','POTENCIAL','HO DESCARTA','FA COMPE','NO COMPLEIX','NUL','']);
let REQ=[]; // requeriments carregats
let REQ_TIPUS=[]; // catàleg de tipus de requeriment
let lastBBDDRows=[]; // files (filtrades) mostrades a la taula BBDD, per exportar
let lastComRows=[]; // files (filtrades) mostrades a la taula Comercial, per exportar
let XLSXLib=null;
async function loadXLSXLib(){
  if(!XLSXLib)XLSXLib=await import('https://esm.sh/xlsx@0.18.5');
  return XLSXLib;
}
async function exportSheet(filename,sheetName,header,dataRows){
  const XLSX=await loadXLSXLib();
  const ws=XLSX.utils.aoa_to_sheet([header,...dataRows]);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,sheetName);
  XLSX.writeFile(wb,filename);
}
async function runExport(btn,fn){
  if(!btn)return fn();
  const orig=btn.textContent;
  btn.disabled=true;btn.textContent='Exportant…';
  try{await fn();}
  catch(e){console.error(e);alert('No s\'ha pogut exportar l\'Excel.');}
  finally{btn.disabled=false;btn.textContent=orig;}
}
function boolTxt(v){return v?'SI':'NO';}
async function exportBBDD(ev){
  await runExport(ev&&ev.target,async()=>{
    const header=['KAM','Empresa','Nova','Cupó anterior','Cupó','Contacte','Mòbil','eMail','Entrada','Estat LEAD','Ticket','Omplir Fitxa - Qui','Omplir Fitxa - Data','HOLDED','Verificar Fitxa - Qui','Verificar Fitxa - Data','ITA','Web - Qui','Web','URL web inicial','URL web final','Validar Fitxa - Qui','Validar Fitxa - Data','Validar Hora - Qui','Validar Hora - Hora','Presentador','Presentat','Resguard arxivat','Notes BBDD'];
    const data=lastBBDDRows.map(d=>[d.g,d.emp,boolTxt(d.nova),d.cup_ant,d.cup,d.cont,d.mob,d.email,d.reg,d.sit,d.tke,d.f1q,d.f1d,boolTxt(d.holded),d.f2q,d.f2d,d.ita,d.webq,d.web,d.url_web,d.url_web_check,d.f3q,d.f3d,d.fhq,d.hora,d.pres,boolTxt(d.presentat),boolTxt(d.resguard),d.notes]);
    await exportSheet('BBDD_cupons.xlsx','BBDD',header,data);
  });
}
async function exportComercial(ev){
  await runExport(ev&&ev.target,async()=>{
    const header=['KAM','Empresa','Nova','Cupó','Consultor','Atorgat','Requeriment','# Edicte S/N','Procés comercial previ','Primer contacte','Proposta preparada','Proposta presentada','Proposta enviada','Acceptada','Kick-off','TODO','Email consultor','Email consultor CC','Notes comercial'];
    const data=lastComRows.map(d=>{
      const reqs=REQ.filter(r=>r.cupo_id===d.id);
      const latest=reqs.length?reqs.reduce((a,b)=>b.id>a.id?b:a):null;
      return [d.g,d.emp,boolTxt(d.nova),d.cup,d.consultor,d.otorgat,latest?(latest.estat||''):'',d.edicte_od,boolTxt(d.proc_comercial_previ),d.primer_contacte,boolTxt(d.proposta_preparada),d.proposta_presentada,d.proposta_enviada,boolTxt(d.oferta_acceptada),d.kickoff_esperat,boolTxt(d.todo),d.email_consultor,d.email_consultor_cc,d.notes_comercial];
    });
    await exportSheet('Comercial_cupons.xlsx','Comercial',header,data);
  });
}
async function exportRequeriments(ev){
  await runExport(ev&&ev.target,async()=>{
    const header=['Empresa','Tiquet','Cupó','Expedient','# Edicte R','Tipus','Aclariment tècnic','Comentaris Back Office','Dead line','Estat requeriment','TODO','Email consultor','Email consultor CC','Data presentació','Comentaris KAM','Resolució final','Data resolució'];
    const data=REQ.map(r=>{
      const ids=parseTipusIds(r.tipus_ids);
      const tipusNames=ids.map(tid=>{const t=REQ_TIPUS.find(x=>x.id===tid);return t?t.alies.split(',')[0]:'';}).filter(Boolean).join(', ');
      return [r.c_emp,r.c_tke,r.c_cup,r.expedient,r.edicte_r,tipusNames,r.aclariment_tecnic,r.comentaris_backoffice,r.dead_line,r.estat,boolTxt(r.todo),r.email_consultor,r.email_consultor_cc,r.data_presentacio,r.comentaris_kam,r.resolucio_final,r.data_resolucio];
    });
    await exportSheet('Requeriments_cupons.xlsx','Requeriments',header,data);
  });
}


/* ═══════════════════════════════════════════════
   2. UTILITATS
═══════════════════════════════════════════════ */
function fmtDate(v){
  if(!v)return '';
  v=v.trim();
  // Formats: d/m, dd/mm, dd.mm.aa, dd.mm.aa hh:mm
  const m1=v.match(/^(\d{1,2})[\/\-\.](\d{1,2})$/);
  if(m1){
    const now=new Date();
    const d=String(m1[1]).padStart(2,'0');
    const mo=String(m1[2]).padStart(2,'0');
    const y=String(now.getFullYear()).slice(2);
    return `${d}.${mo}.${y}`;
  }
  const m2=v.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})/);
  if(m2){
    const d=String(m2[1]).padStart(2,'0');
    const mo=String(m2[2]).padStart(2,'0');
    const y=m2[3].length===4?m2[3].slice(2):m2[3];
    const rest=v.includes(' ')?v.slice(v.indexOf(' ')):'';
    return `${d}.${mo}.${y}${rest}`;
  }
  return v;
}
function fmtHora(v){
  if(!v)return '';
  v=v.trim();
  const m=v.match(/^(\d{1,2}):?(\d{2})?$/);
  if(m)return `${String(m[1]).padStart(2,'0')}:${m[2]||'00'}`;
  return v;
}
function addTimestamp(dateStr){
  if(!dateStr||dateStr.includes(' '))return dateStr;
  const now=new Date();
  return `${dateStr} ${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}`;
}
function dateToMs(s){
  if(!s)return 0;
  const [dt,tm='00:00']=s.split(' ');
  const [d,m,y]=dt.split('.');
  return new Date(`20${y}-${m}-${d}T${tm}`).getTime();
}

/* ═══════════════════════════════════════════════
   3. LÒGICA DE DADES
═══════════════════════════════════════════════ */
function isValidDate(s){
  return s&&/^\d{2}\.\d{2}\.\d{2}/.test(s);
}
function checkAutoComplete(d){
  if(!d.reg&&['COMPLETAT','PROCÉS'].includes(d.sit)){d.sit='POTENCIAL';return;}
  if(['HO DESCARTA','FA COMPE','NO COMPLEIX','NUL'].includes(d.sit))return;
  const f1ok=d.f1q&&isValidDate(d.f1d);
  const f2ok=d.f2q&&isValidDate(d.f2d);
  const f3ok=d.f3q&&isValidDate(d.f3d);
  const itaOk=d.cup!=='IA'||d.ita==='REBUT';
  const holdedOk=!!d.holded;
  const allDone=f1ok&&f2ok&&f3ok&&itaOk&&holdedOk;
  if(allDone&&d.reg){
    if(d.web!=='') d.sit='COMPLETAT';
    else if(d.sit!=='COMPLETAT') showWebDialog(d);
  } else if(!allDone&&d.sit==='COMPLETAT'&&YEAR_CONFIG.strictComplete){
    d.sit='PROCÉS';
  }
}
async function reorderTke(el){
  const id=+el.dataset.id;
  const newNum=parseInt(el.textContent.trim());
  const d=D.find(x=>x.id===id);
  if(!d||isNaN(newNum))return render();
  const actius=D.filter(x=>['COMPLETAT','PROCÉS'].includes(x.sit)&&x.reg&&x.tke)
    .sort((a,b)=>parseInt(a.tke)-parseInt(b.tke));
  const n=Math.max(1,Math.min(actius.length,newNum));
  const idx=actius.findIndex(x=>x.id===id);
  actius.splice(idx,1);
  actius.splice(n-1,0,d);
  actius.forEach((x,i)=>x.tke=String(i+1));
  showSaving();
  try{
    await Promise.all(actius.map(x=>sbUpdate(x.id,x)));
    showSaved();
  }catch(e){showError(e.message);}
  render();
}

function showWebDialog(d){
  if(document.getElementById('web-dialog')) return;
  const el=document.createElement('div');
  el.id='web-dialog';
  el.style='position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:99999;display:flex;align-items:center;justify-content:center';
  el.innerHTML=`<div style="background:#fff;border-radius:8px;padding:24px;max-width:400px;box-shadow:0 8px 32px rgba(0,0,0,.2)">
    <p style="font-size:11pt;font-weight:bold;margin-bottom:10px">⚠️ Verificació WEB requerida</p>
    <p style="font-size:10pt;margin-bottom:20px;line-height:1.5">No es pot marcar <strong>${d.emp||'aquesta empresa'}</strong> com a COMPLETAT sense haver verificat en el camp WEB que l'empresa té web en idioma estranger. Si us plau, verifica-ho i selecciona una opció al camp WEB.</p>
    <div style="text-align:right">
      <button onclick="document.getElementById('web-dialog').remove()" style="padding:6px 20px;border:1px solid #bbb;border-radius:4px;cursor:pointer;font-size:10pt;background:#2c5aa0;color:#fff">OK</button>
    </div>
  </div>`;
  document.body.appendChild(el);
}
function calcTickets(){
  D.forEach(d=>{if(!['COMPLETAT','PROCÉS'].includes(d.sit)||!d.reg)d.tke='';});
  D.filter(d=>['COMPLETAT','PROCÉS'].includes(d.sit)&&d.reg)
   .sort((a,b)=>dateToMs(a.reg)-dateToMs(b.reg))
   .forEach((d,i)=>d.tke=String(i+1));
}
function delRow_placeholder(){} // sobreescrita per Supabase




/* ═══════════════════════════════════════════════
   4. PERMISOS
═══════════════════════════════════════════════ */
const HELPER_FIELDS=['emp','cont','mob','email','reg','notes']; // camps que pot tocar un Admin ajudant (no Master)
function canEd(d,f){
  if(fHora){
    if(f==='hora')return (cT==='a'&&adminLevel==='master')?!!d.fhq:(cT==='h'&&d.sit==='COMPLETAT'&&!d.hora&&d.fhq===cU);
    return false;
  }
  if(cT==='a'){
    if(adminLevel==='master'){ if(f==='hora') return !!d.fhq; return true; }
    if(adminLevel==='admin') return HELPER_FIELDS.includes(f);
    return false; // no autenticat → només lectura
  }
  if(cT==='g')return false;
  if(cT==='p'){
    if(f==='notes')return d.pres===cU;
    return false; // presentat/resguard es gestionen per separat
  }
  if(cT==='w'){
    if(f==='url_web'||f==='url_web_check')return d.webq===cU;
    return false; // el camp "web" el gestiona webCell() directament
  }
  if(cT==='h'){
    if(f==='hora')return d.sit==='COMPLETAT'&&!d.hora&&d.fhq===cU;
    return false;
  }
  if(cT==='t'){
    if(f==='notes')return true;
    if(f==='f1d')return d.f1q===cU;
    if(f==='f2d')return d.f2q===cU;
    if(f==='f3d')return d.f3q===cU;
    return false;
  }
  return false;
}

/* ═══════════════════════════════════════════════
   5. RENDER
═══════════════════════════════════════════════ */
function ec(d,f,v,multi){
  if(!canEd(d,f))return `<span>${v||''}</span>`;
  return `<span class="ed" contenteditable="true" data-id="${d.id}" data-f="${f}" onblur="sv(this)" onkeydown="if(event.key==='Enter'&&!${!!multi}){event.preventDefault();this.blur()}">${v||''}</span>`;
}
function ovCell(v,html){
  return `<div class="ov-cell">${html}${v?`<div class="ov-pop">${v}</div>`:''}</div>`;
}
function sel(d,f,list,v){
  if(cT!=='a'||!adminLevel)return `<span>${v||''}</span>`;
  const emptyOpt=f==='g'?'<option value=""></option>':'';
  return `<select onchange="svs(${d.id},'${f}',this.value)">${emptyOpt}${list.map(x=>`<option value="${x}"${v===x?' selected':''}>${x}</option>`).join('')}</select>`;
}
function selFQ(d,n){
  const f='f'+n+'q', lst=n===1?PF:n===2?G:TF3, v=d[f];
  if(cT!=='a'||!adminLevel)return `<span>${v||''}</span>`;
  return `<select onchange="svs(${d.id},'${f}',this.value)"><option value=""></option>${lst.map(x=>`<option${v===x?' selected':''}>${x}</option>`).join('')}</select>`;
}
function selFH(d){
  if(cT!=='a'||adminLevel!=='master')return `<span>${d.fhq||''}</span>`;
  return `<select onchange="svs(${d.id},'fhq',this.value)"><option value=""></option>${TFH.map(x=>`<option${d.fhq===x?' selected':''}>${x}</option>`).join('')}</select>`;
}
function selWebQ(d){
  if(cT!=='a'||adminLevel!=='master')return `<span>${d.webq||''}</span>`;
  return `<select onchange="svs(${d.id},'webq',this.value)"><option value=""></option>${QW.map(x=>`<option${d.webq===x?' selected':''}>${x}</option>`).join('')}</select>`;
}
function sitCell(d){
  const sc=SIT_COLORS[d.sit]||{};
  const bg=sc.bg||'',fg=sc.fg||'#222';
  const label=d.sit||'';
  const canEditSit=(cT==='a'&&adminLevel==='master')||(cT==='t'&&d.sit!=='COMPLETAT');
  if(!canEditSit)return `<div style="background:${bg};color:${fg};padding:2px 4px;border-radius:3px">${label}</div>`;
  const opts=(cT==='t'?SITS.filter(s=>s&&s!=='COMPLETAT'&&s!=='PROCÉS'):SITS.filter(s=>s))
    .map(s=>{const c=SIT_COLORS[s]||{};return `<div class="sit-opt" style="background:${c.bg||'#fff'};color:${c.fg||'#222'}" onclick="svSit(${d.id},'${s}')">${s}</div>`;}).join('');
  return `<div class="sit-wrap" style="background:${bg};color:${fg}" onclick="toggleDrop(event,this,'sit-drop')">
    <span>${label}</span><span class="sit-arrow">▾</span>
    <div class="sit-drop">${opts}</div>
  </div>`;
}
function holdedCell(d){
  const canHold=(cT==='a'&&adminLevel)||(cT==='t'&&d.f1q===cU);
  return `<td class="chold webcel"><input type="checkbox"${d.holded?' checked':''}${canHold?'':' disabled'} onchange="svs(${d.id},'holded',this.checked)"></td>`;
}
function itaCell(d){
  const v=d.ita||'';
  const ic=ITA_COLORS[v]||ITA_COLORS[''];
  const bgStyle=ic.bg?`background:${ic.bg};`:'';
  const label=v||'No aplica';
  if(d.cup!=='IA')return `<td class="cita" style="${bgStyle}color:${ic.fg};text-align:center">${label}</td>`;
  const canIta=(cT==='a'&&adminLevel==='master')||(cT==='t'&&(d.f1q===cU||d.f2q===cU||d.f3q===cU));
  if(!canIta)return `<td class="cita" style="${bgStyle}color:${ic.fg};text-align:center">${label}</td>`;
  const opts=['No aplica','DEMANAR','DEMANAT','REBUT'].map(s=>{
    const key=s==='No aplica'?'':s;
    const c=ITA_COLORS[key]||ITA_COLORS[''];
    return `<div class="ita-opt" style="${c.bg?'background:'+c.bg+';':''}color:${c.fg}" onclick="svIta(${d.id},'${key}')">${s}</div>`;
  }).join('');
  return `<td class="cita" style="${bgStyle}color:${ic.fg}">
    <div class="ita-wrap" onclick="toggleDrop(event,this,'ita-drop')">
      <span>${label}</span><span class="sit-arrow">▾</span>
      <div class="ita-drop">${opts}</div>
    </div></td>`;
}
function webCell(d){
  const wc=WEB_COLORS[d.web]||{};
  const bgStyle=wc.bg?`background:${wc.bg};color:${wc.fg};`:'';
  const canW=(cT==='a'&&adminLevel==='master'&&!fHora)||(cT==='w'&&d.webq===cU);
  if(!canW)return `<td class="cweb webcel" style="${bgStyle}"><span>${d.web||''}</span></td>`;
  return `<td class="cweb webcel" style="${bgStyle}"><select style="${bgStyle}width:100%" onchange="svs(${d.id},'web',this.value)">
    <option value=""></option>${WEBLIST.map(x=>`<option value="${x}"${d.web===x?' selected':''}>${x}</option>`).join('')}
  </select></td>`;
}
function presCell(d){
  if(cT!=='a'||adminLevel!=='master')return `<td class="cpres">${d.pres||''}</td>`;
  return `<td class="cpres"><select onchange="svs(${d.id},'pres',this.value)"><option value=""></option>${TD.map(x=>`<option${d.pres===x?' selected':''}>${x}</option>`).join('')}</select></td>`;
}
function presTd(d){
  const canP=d.sit==='COMPLETAT'&&d.pres&&((cT==='a'&&adminLevel==='master')||(cT==='p'&&d.pres===cU));
  return `<td class="cpret webcel"><input type="checkbox"${d.presentat?' checked':''}${canP?'':' disabled'} onchange="svs(${d.id},'presentat',this.checked)"></td>`;
}
function resguardTd(d){
  const bg=d.presentat&&!d.resguard?'background:#F4CCCC;color:#000;':'';
  const canR=d.sit==='COMPLETAT'&&d.presentat&&((cT==='a'&&adminLevel==='master')||(cT==='p'&&d.pres===cU));
  return `<td class="crsgd webcel" style="${bg}"><input type="checkbox"${d.resguard?' checked':''}${canR?'':' disabled'} onchange="svs(${d.id},'resguard',this.checked)"></td>`;
}
function rowHtml(d,isMaster){
  return `<tr>
    <td style="text-align:center;padding:2px">${isMaster?`<button onclick="delRow(${d.id})" style="border:none;background:none;cursor:pointer;color:#c00;font-size:14pt;font-weight:bold" title="Eliminar">×</button>`:''}</td>
    <td class="cg">${sel(d,'g',G,d.g)}</td>
    <td class="cemp">${ec(d,'emp',d.emp)}</td>
    <td class="cnov webcel"><input type="checkbox"${d.nova?' checked':''}${(cT==='a'&&adminLevel)?'':' disabled'} onchange="svs(${d.id},'nova',this.checked)"></td>
    <td class="ccupant">${sel(d,'cup_ant',CUPS,d.cup_ant)}</td>
    <td class="ccup">${sel(d,'cup',CUPS,d.cup)}</td>
    <td class="ccon">${ec(d,'cont',d.cont)}</td>
    <td class="cmob">${ec(d,'mob',d.mob)}</td>
    <td class="ceml">${ovCell(d.email,ec(d,'email',d.email))}</td>
    <td class="creg">${ec(d,'reg',d.reg)}</td>
    <td class="csit">${sitCell(d)}</td>
    <td class="ctke" style="text-align:center;color:#555">${isMaster&&d.tke?`<span class="ed" contenteditable="true" data-id="${d.id}" data-f="tke" onblur="reorderTke(this)" onkeydown="if(event.key==='Enter'){event.preventDefault();this.blur()}" style="cursor:text">${d.tke}</span>`:d.tke||''}</td>
    <td class="cfq">${selFQ(d,1)}</td><td class="cfd">${ec(d,'f1d',d.f1d)}</td>
    ${holdedCell(d)}
    <td class="cfq">${selFQ(d,2)}</td><td class="cfd">${ec(d,'f2d',d.f2d)}</td>
    ${itaCell(d)}
    <td class="cfq">${selWebQ(d)}</td>
    ${webCell(d)}
    <td class="curlweb">${ovCell(d.url_web,ec(d,'url_web',d.url_web))}</td>
    <td class="curlwebchk">${ovCell(d.url_web_check,ec(d,'url_web_check',d.url_web_check))}</td>
    <td class="cfq">${selFQ(d,3)}</td><td class="cfd">${ec(d,'f3d',d.f3d)}</td>
    <td class="cfq">${selFH(d)}</td>
    <td class="chra">${ec(d,'hora',d.hora)}</td>
    ${presCell(d)}
    ${presTd(d)}
    ${resguardTd(d)}
    <td class="ncell cnot"><div class="notes-wrap${notesExpanded.has(d.id)?' expanded':''}" data-id="${d.id}">${ec(d,'notes',d.notes,true)}<button type="button" class="notes-toggle" onclick="toggleNotes(this)">+</button></div></td>
  </tr>`;
}
/* ═══════════════════════════════════════════════
   5b. PESTANYA COMERCIAL
═══════════════════════════════════════════════ */
function comCheckColorCell(d,f,colorObj){
  const bg=d[f]?colorObj.bg:'',fg=d[f]?colorObj.fg:'#222';
  const canEdit=cT==='a'&&adminLevel;
  const interact=canEdit?` onchange="svs(${d.id},'${f}',this.checked)"`:` onclick="return false;" style="cursor:default"`;
  return `<td class="com-check" style="background:${bg};color:${fg}"><input type="checkbox"${d[f]?' checked':''}${interact}></td>`;
}
function comReqEstatCell(d){
  const reqs=REQ.filter(r=>r.cupo_id===d.id);
  const latest=reqs.length?reqs.reduce((a,b)=>b.id>a.id?b:a):null;
  const estat=latest?(latest.estat||''):'';
  const c=REQ_ESTAT_COLORS[estat]||{};
  const bg=c.bg||'',fg=c.fg||'#222';
  return `<td class="com-req" style="background:${bg};color:${fg}"><span style="display:inline-block;width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${estat}</span></td>`;
}
function comCheckCell(d,f){
  const canEdit=cT==='a'&&adminLevel;
  const interact=canEdit?` onchange="svs(${d.id},'${f}',this.checked)"`:` onclick="return false;" style="cursor:default"`;
  return `<td class="com-check"><input type="checkbox"${d[f]?' checked':''}${interact}></td>`;
}
function comConsultorCell(d){
  const canEdit=cT==='a'&&adminLevel;
  if(!canEdit)return `<td class="ccons">${d.consultor||''}</td>`;
  return `<td class="ccons"><select onchange="svs(${d.id},'consultor',this.value)"><option value=""></option>${CONS.map(x=>`<option value="${x}"${d.consultor===x?' selected':''}>${x}</option>`).join('')}</select></td>`;
}
function comTodoCell(d){
  const canEdit=cT==='a'&&adminLevel;
  const bg=d.todo?'background:#FF3333;':'';
  const interact=canEdit?` onchange="svs(${d.id},'todo',this.checked)"`:` onclick="return false;" style="cursor:default"`;
  return `<td class="com-check" style="${bg}"><input type="checkbox"${d.todo?' checked':''}${interact}></td>`;
}
function comEmailSel(d,field,val,canEdit){
  if(!canEdit)return `<span style="max-width:190px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;display:inline-block">${val||''}</span>`;
  return `<select class="sel-small" style="max-width:190px;width:100%" onchange="svs(${d.id},'${field}',this.value)"><option value=""${val?'':' selected'}>—</option>${CONSULTORS.map(c=>`<option value="${c}"${val===c?' selected':''}>${c}</option>`).join('')}</select>`;
}
function comEmailConsultorCell(d){
  const canEdit=cT==='a'&&adminLevel;
  const hasEmail=!!(d.email_consultor&&d.email_consultor.trim());
  return `<td><div style="display:flex;flex-direction:column;gap:2px;min-width:215px">
    <div style="display:flex;align-items:center;gap:4px">${comEmailSel(d,'email_consultor',d.email_consultor,canEdit)}<button type="button" onclick="sendComEmail(${d.id})"${hasEmail?'':' disabled'} title="Enviar per correu al consultor" style="border:none;background:none;cursor:${hasEmail?'pointer':'not-allowed'};font-size:12pt;opacity:${hasEmail?'1':'.35'};flex-shrink:0">✉️</button></div>
    <div style="display:flex;align-items:center;gap:4px"><span style="font-size:7pt;color:#888;width:16px;flex-shrink:0">CC</span>${comEmailSel(d,'email_consultor_cc',d.email_consultor_cc,canEdit)}</div>
  </div></td>`;
}
function sendComEmail(id){
  const d=D.find(x=>x.id===id);
  if(!d||!d.email_consultor)return;
  const reqs=REQ.filter(r=>r.cupo_id===d.id);
  const latest=reqs.length?reqs.reduce((a,b)=>b.id>a.id?b:a):null;
  const reqEstat=latest?(latest.estat||''):'';
  const subject=`Seguiment ${d.emp||''} — Cupó ${d.cup||''}`;
  const body=[
    `KAM: ${d.g||''}`,
    `Empresa: ${d.emp||''}`,
    `Cupó: ${d.cup||''}`,
    `Consultor: ${d.consultor||''}`,
    `Atorgat: ${d.otorgat||''}`,
    `Requeriment: ${reqEstat}`,
    `# Edicte S/N: ${d.edicte_od||''}`,
    `Procés comercial previ: ${d.proc_comercial_previ?'SI':'PENDENT'}`,
    `Primer contacte: ${d.primer_contacte||''}`,
    `Proposta preparada: ${d.proposta_preparada?'SI':'PENDENT'}`,
    `Proposta presentada: ${d.proposta_presentada||''}`,
    `Proposta enviada: ${d.proposta_enviada||''}`,
    `Acceptada: ${d.oferta_acceptada?'SI':'PENDENT'}`,
    `Kick-off: ${d.kickoff_esperat||''}`,
    `Notes comercial: ${d.notes_comercial||''}`,
  ].join('\n');
  const cc=d.email_consultor_cc?`&cc=${encodeURIComponent(d.email_consultor_cc)}`:'';
  window.location.href=`mailto:${d.email_consultor}?subject=${encodeURIComponent(subject)}${cc}&body=${encodeURIComponent(body)}`;
}
function comOtorgatCell(d){
  const c=OTORGAT_ESTAT[d.otorgat]||{};
  const bg=c.bg||'',fg=c.fg||'#222';
  const canEdit=cT==='a'&&adminLevel;
  if(!canEdit)return `<td class="com-otor" style="background:${bg};color:${fg}"><span style="display:inline-block;width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${d.otorgat||''}</span></td>`;
  const style=bg?` style="background:${bg};color:${fg}"`:'';
  return `<td class="com-otor"><select${style} onchange="svs(${d.id},'otorgat',this.value)">${OTORGAT_OPTS.map(o=>`<option value="${o}"${d.otorgat===o?' selected':''}>${o||'—'}</option>`).join('')}</select></td>`;
}
function comRowHtml(d){
  return `<tr>
    <td class="cg">${d.g||''}</td>
    <td class="cemp">${d.emp||''}</td>
    <td class="cnov webcel"><input type="checkbox"${d.nova?' checked':''} onclick="return false;" style="cursor:default"></td>
    <td class="ccup">${d.cup||''}</td>
    ${comConsultorCell(d)}
    ${comOtorgatCell(d)}
    ${comReqEstatCell(d)}
    <td class="cedicte">${ec(d,'edicte_od',d.edicte_od)}</td>
    ${comCheckColorCell(d,'proc_comercial_previ',OTORGAT_ESTAT.SI)}
    <td class="com-date">${ec(d,'primer_contacte',d.primer_contacte)}</td>
    ${comCheckColorCell(d,'proposta_preparada',OTORGAT_ESTAT.SI)}
    <td class="com-date">${ec(d,'proposta_presentada',d.proposta_presentada)}</td>
    <td class="com-date">${ec(d,'proposta_enviada',d.proposta_enviada)}</td>
    ${comCheckCell(d,'oferta_acceptada')}
    <td class="com-date">${ec(d,'kickoff_esperat',d.kickoff_esperat)}</td>
    ${comTodoCell(d)}
    ${comEmailConsultorCell(d)}
    <td class="ncell com-notes"><div class="notes-wrap${comNotesExpanded.has(d.id)?' expanded':''}" data-id="${d.id}">${ec(d,'notes_comercial',d.notes_comercial,true)}<button type="button" class="notes-toggle" onclick="toggleComNotes(this)">+</button></div></td>
  </tr>`;
}
function comOtorgatCounts(){
  const completatsCount=D.filter(d=>d.sit==='COMPLETAT').length;
  const otorgatsCount=D.filter(d=>d.otorgat==='SI').length;
  const refusatsCount=D.filter(d=>d.otorgat==='NO').length;
  const requeritsCount=D.filter(d=>d.otorgat==='REQUERIT').length;
  const pendentsCount=completatsCount-otorgatsCount-refusatsCount;
  return [
    {l:'COMPLETATS',n:completatsCount,bg:'#00B050',fg:'#fff'},
    {l:'ATORGATS',n:otorgatsCount,bg:'#A02B93',fg:'#fff'},
    {l:'REFUSATS',n:refusatsCount,bg:'#808080',fg:'#fff'},
    {l:'PENDENTS',n:pendentsCount,bg:'#fff',fg:'#222'},
    {l:'REQUERITS',n:requeritsCount,bg:'#F2CEED',fg:'#000'},
  ];
}
function cupPresentatCounts(){
  return ['CEXP','PI','EC','IA'].map(c=>{
    const col=CUP_COLORS[c];
    return {l:c,n:D.filter(d=>d.presentat&&d.cup===c).length,bg:col.bg,fg:col.fg};
  });
}
function cupSiNoGroups(){
  return ['CEXP','PI','EC','IA'].map(c=>{
    const col=CUP_COLORS[c];
    const rows=D.filter(d=>d.presentat&&d.cup===c);
    const si=rows.filter(d=>d.otorgat==='SI').length;
    const no=rows.filter(d=>d.otorgat==='NO').length;
    return [
      {l:c,n:rows.length,bg:col.bg,fg:col.fg},
      {l:'SI',n:si,bg:OTORGAT_ESTAT.SI.bg,fg:OTORGAT_ESTAT.SI.fg},
      {l:'NO',n:no,bg:OTORGAT_ESTAT.NO.bg,fg:OTORGAT_ESTAT.NO.fg},
    ];
  });
}
function comSort(field){
  if(comSortField===field)comSortDir*=-1;else{comSortField=field;comSortDir=1;}
  renderComercial();
}
function updateComSortArrows(){
  ['g','emp','cup','consultor'].forEach(f=>{
    const el=document.getElementById('csort-'+f);
    if(el)el.textContent=comSortField===f?(comSortDir===1?' ▲':' ▼'):'';
  });
}
function renderComercial(){
  let rows=D.filter(d=>d.presentat);
  if(fComOtor)rows=rows.filter(d=>fComOtor==='_BLANK_'?!d.otorgat:d.otorgat===fComOtor);
  if(fComPendent){
    const [pField,pStatus]=fComPendent.split(':');
    rows=rows.filter(d=>pStatus==='complet'?!!d[pField]:!d[pField]);
  }
  if(fComKam)rows=rows.filter(d=>d.g===fComKam);
  if(fComCup)rows=rows.filter(d=>d.cup===fComCup);
  if(fComCons)rows=rows.filter(d=>d.consultor===fComCons);
  rows=rows.sort((a,b)=>{
    if(comSortField){
      const cmp=(a[comSortField]||'').toString().localeCompare((b[comSortField]||'').toString(),'ca',{sensitivity:'base',numeric:true})*comSortDir;
      if(cmp!==0)return cmp;
    }
    return (parseInt(a.tke)||0)-(parseInt(b.tke)||0);
  });
  lastComRows=rows;
  const tbody=document.getElementById('com-tbody');
  if(tbody)tbody.innerHTML=rows.map(comRowHtml).join('');
  const rcEl=document.getElementById('recomptes-comercial');
  if(rcEl)rcEl.innerHTML=comOtorgatCounts().map(r=>`<div class="rcomp" style="background:${r.bg};color:${r.fg}"><span class="rc-n">${r.n}</span><span class="rc-l">${r.l}</span></div>`).join('');
  const rcCupEl=document.getElementById('recomptes-cupo-comercial');
  if(rcCupEl)rcCupEl.innerHTML=cupSiNoGroups().map(g=>`<div class="cup-group">${g.map(r=>`<div class="rcomp" style="background:${r.bg};color:${r.fg}"><span class="rc-n">${r.n||'—'}</span><span class="rc-l">${r.l}</span></div>`).join('')}</div>`).join('');
  const novesComCount=D.filter(d=>d.presentat&&d.nova).length;
  const rcNovaEl=document.getElementById('recomptes-nova-comercial');
  if(rcNovaEl)rcNovaEl.innerHTML=`<div class="rcomp" style="background:#2c5aa0;color:#fff"><span class="rc-n">${novesComCount}</span><span class="rc-l">NOVES</span></div>`;
  const fckEl=document.getElementById('fComKamSel');
  if(fckEl)fckEl.innerHTML=`<option value="">Tots</option>${G.map(x=>`<option value="${x}"${fComKam===x?' selected':''}>${x}</option>`).join('')}`;
  const fccEl=document.getElementById('fComCupSel');
  if(fccEl)fccEl.innerHTML=`<option value="">Tots</option>${CUPS.filter(x=>x).map(x=>`<option value="${x}"${fComCup===x?' selected':''}>${x}</option>`).join('')}`;
  const fcnEl=document.getElementById('fComConsSel');
  if(fcnEl)fcnEl.innerHTML=`<option value="">Tots</option>${CONS.map(x=>`<option value="${x}"${fComCons===x?' selected':''}>${x}</option>`).join('')}`;
  requestAnimationFrame(()=>updateNotesToggles('#com-table-wrap'));
  updateComSortArrows();
}

function toggleGroup(sit){
  if(collapsedGroups.has(sit))collapsedGroups.delete(sit);else collapsedGroups.add(sit);
  render();
}

function render(){
  rebuildRoleBtns();

  const hints={a:adminLevel==='master'?'Master autenticat — edició completa i gestió d\'admins.':adminLevel==='admin'?`Admin ${cU} — creació de registres i dades bàsiques.`:'Vista general — seleccioneu el vostre rol o autentiqueu-vos via ☰ Admin.',g:`KAM ${cU} — els vostres registres.`,
    t:`Presentador Fitxa ${cU} — editeu les vostres fitxes i l'ITA.`,p:`Tramitador dia-D ${cU} — editeu Presentat i Resguard.`,
    w:`Gestor Web ${cU} — editeu Web, URL web inicial i URL web final.`,h:`Gestor Hora ${cU} — editeu la Validar Hora.`};
  document.getElementById('hint-txt').textContent=hints[cT]||'';

  document.getElementById('btn-nova').style.display=cT==='a'?'inline-block':'none';
  const canAddRow=cT==='a'&&!!adminLevel;
  const isMaster=cT==='a'&&adminLevel==='master';
  const webPendent=D.filter(d=>['COMPLETAT','PROCÉS'].includes(d.sit)&&d.web!=='SI').length;
  const btnWeb=document.getElementById('btnWeb');
  btnWeb.textContent=`WEB Pendent${webPendent>0?' ('+webPendent+')':''}`;
  btnWeb.style.fontWeight=webPendent>0?'bold':'normal';
  btnWeb.style.background=webPendent>0?'#F4CCCC':'';
  btnWeb.style.borderColor=webPendent>0?'#c00':'';
  document.getElementById('btn-nova').style.opacity=canAddRow?'1':'0.4';
  document.getElementById('btn-nova').title=canAddRow?'':'Requereix autenticació Admin';
  document.getElementById('btn-admin').style.display=cT==='a'?'inline-block':'none'; // sempre visible
  document.getElementById('btn-llistes').style.display=isMaster?'inline-block':'none';

  let rows=D.filter(d=>{
    if(cT==='g'&&d.g!==cU)return false;
    if(cT==='t'&&(d.f1q!==cU&&d.f2q!==cU&&d.f3q!==cU))return false;
    if(cT==='t'&&!['COMPLETAT','PROCÉS'].includes(d.sit))return false;
    if(cT==='p'&&(d.pres!==cU||d.sit!=='COMPLETAT'))return false;
    if(cT==='w'&&d.webq!==cU)return false;
    if(cT==='h'&&(d.fhq!==cU||d.sit!=='COMPLETAT'))return false;
    if(fWeb&&d.web==='SI')return false;
    if(fHora&&(d.hora||d.sit!=='COMPLETAT'))return false;
    if(fNova&&!d.nova)return false;
    if(fCup&&d.cup!==fCup)return false;
    if(fSit&&d.sit!==fSit)return false;
    if(fQ){const q=fQ.toLowerCase();if(!d.emp.toLowerCase().includes(q)&&!d.cont.toLowerCase().includes(q))return false;}
    return true;
  });

  rows.sort((a,b)=>{
    const aA=['COMPLETAT','PROCÉS'].includes(a.sit), bA=['COMPLETAT','PROCÉS'].includes(b.sit);
    if(aA&&bA)return (parseInt(a.tke)||0)-(parseInt(b.tke)||0);
    if(aA)return -1; if(bA)return 1;
    return (SIT_ORDER[a.sit]??99)-(SIT_ORDER[b.sit]??99);
  });

  // Recomptes
  const cnt={};
  SITS.filter(s=>s).forEach(s=>cnt[s]=0);
  D.forEach(d=>{if(cnt[d.sit]!==undefined)cnt[d.sit]++;});
  const total=D.length;
  const rc=[
    {l:'REGISTRES',n:total,bg:'#fff',fg:'#222'},
    {l:'PROCÉS',n:cnt['PROCÉS'],bg:'#DAF2D0',fg:'#222'},
    {l:'POTENCIALS',n:cnt['POTENCIAL'],bg:'#fff',fg:'#222'},
    {l:'HO DESCARTA',n:cnt['HO DESCARTA'],bg:'#fff',fg:'#222'},
    {l:'FA COMPE',n:cnt['FA COMPE'],bg:'#fff',fg:'#222'},
    {l:'NO COMPLEIX',n:cnt['NO COMPLEIX'],bg:'#fff',fg:'#222'},
    {l:'NUL',n:cnt['NUL'],bg:'#fff',fg:'#222'},
  ];
  document.getElementById('recomptes').innerHTML=rc.map(r=>`<div class="rcomp" style="background:${r.bg};color:${r.fg}"><span class="rc-n">${r.n}</span><span class="rc-l">${r.l}</span></div>`).join('');
  document.getElementById('recomptes-otorgats').innerHTML=comOtorgatCounts().map(r=>`<div class="rcomp" style="background:${r.bg};color:${r.fg}"><span class="rc-n">${r.n}</span><span class="rc-l">${r.l}</span></div>`).join('');
  document.getElementById('recomptes-cupo').innerHTML=cupPresentatCounts().map(r=>`<div class="rcomp" style="background:${r.bg};color:${r.fg}"><span class="rc-n">${r.n}</span><span class="rc-l">${r.l}</span></div>`).join('');

  const novesCount=D.filter(d=>d.nova).length;
  document.getElementById('recomptes-nova').innerHTML=`<div class="rcomp" style="background:#2c5aa0;color:#fff"><span class="rc-n">${novesCount}</span><span class="rc-l">NOVES</span></div>`;

  // Actualitzar selectors filtres
  const fcs=document.getElementById('fCupSel');
  const fss=document.getElementById('fSitSel');
  fcs.innerHTML=`<option value="">Tots</option>${CUPS.filter(x=>x).map(x=>`<option value="${x}"${fCup===x?' selected':''}>${x}</option>`).join('')}`;
  fss.innerHTML=`<option value="">Totes</option>${SITS.filter(x=>x).map(x=>`<option value="${x}"${fSit===x?' selected':''}>${x}</option>`).join('')}`;

  const GROUP_DEFS=[
    {key:'COMPLETAT',label:'COMPLETAT'},
    {key:'PROCÉS',label:'PROCÉS'},
    {key:'POTENCIAL',label:'POTENCIAL'},
    {key:'FA COMPE',label:'FA COMPE'},
    {key:'HO DESCARTA',label:'HO DESCARTA'},
    {key:'NO COMPLEIX',label:'NO COMPLEIX'},
    {key:'NUL',label:'NUL'},
    {key:'',label:'Sense estat'},
  ];
  let tbodyHtml='';
  lastBBDDRows=[];
  GROUP_DEFS.forEach(g=>{
    const groupRows=rows.filter(d=>d.sit===g.key);
    if(!groupRows.length)return;
    lastBBDDRows.push(...groupRows);
    const collapsed=collapsedGroups.has(g.key);
    const sc=SIT_COLORS[g.key]||{bg:'#eee',fg:'#333'};
    tbodyHtml+=`<tr class="group-row"><td colspan="28" style="background:${sc.bg||'#eee'};color:${sc.fg||'#333'};cursor:pointer;font-weight:bold;padding:5px 8px" onclick="toggleGroup('${g.key}')">
      <span style="display:inline-block;width:16px">${collapsed?'▶':'▼'}</span>${g.label} <span style="font-weight:normal;opacity:.8">(${groupRows.length})</span>
    </td></tr>`;
    if(!collapsed)tbodyHtml+=groupRows.map(d=>rowHtml(d,isMaster)).join('');
  });
  document.getElementById('tbody').innerHTML=tbodyHtml;
  renderComercial();

  // Ajust top de la 2a fila de capçalera
  requestAnimationFrame(()=>{
    const rows=document.querySelectorAll('#table-wrap thead tr');
    if(rows[0]){
      const h1=rows[0].offsetHeight;
      if(rows[1])rows[1].querySelectorAll('th').forEach(th=>th.style.top=h1+'px');
      if(rows[2])rows[2].querySelectorAll('th').forEach(th=>th.style.top=(h1+(rows[1]?rows[1].offsetHeight:0))+'px');
    }
    updateNotesToggles('#table-wrap');
    adjustTableHeight();
    updateStickyOffsets();
  });
}
function adjustTableHeight(){
  const wrap=document.getElementById('table-wrap');
  if(!wrap)return;
  const top=wrap.getBoundingClientRect().top;
  const avail=window.innerHeight-top-12;
  wrap.style.maxHeight=Math.max(200,avail)+'px';
}
function updateStickyOffsets(){
  const ths=document.querySelectorAll('#table-wrap thead tr:first-child th');
  if(ths.length<6)return;
  const root=document.documentElement.style;
  let sum=0;
  root.setProperty('--stick1','0px');
  for(let i=0;i<5;i++){
    sum+=ths[i].offsetWidth;
    root.setProperty('--stick'+(i+2),sum+'px');
  }
}
function updateReqStickyOffsets(){
  const ths=document.querySelectorAll('#req-table-wrap thead th');
  if(ths.length<5)return;
  const root=document.documentElement.style;
  let sum=0;
  root.setProperty('--rstick1','0px');
  for(let i=0;i<4;i++){
    sum+=ths[i].offsetWidth;
    root.setProperty('--rstick'+(i+2),sum+'px');
  }
}
function adjustLayout(){adjustTableHeight();updateStickyOffsets();updateReqStickyOffsets();}
window.addEventListener('resize',adjustLayout);
if(window.ResizeObserver){
  const headerObs=new ResizeObserver(()=>adjustLayout());
  const topRowEl=document.querySelector('.top-row');
  const hintEl=document.getElementById('hint');
  if(topRowEl)headerObs.observe(topRowEl);
  if(hintEl)headerObs.observe(hintEl);
  headerObs.observe(document.body);
}
function toggleNotes(btn){
  const wrap=btn.parentElement;
  const id=+wrap.dataset.id;
  const isExp=wrap.classList.toggle('expanded');
  if(isExp)notesExpanded.add(id);else notesExpanded.delete(id);
  btn.textContent=isExp?'−':'+';
}
function toggleReqNotes(btn){
  const wrap=btn.parentElement;
  const key=wrap.dataset.rid+':'+wrap.dataset.rf;
  const isExp=wrap.classList.toggle('expanded');
  if(isExp)reqNotesExpanded.add(key);else reqNotesExpanded.delete(key);
  btn.textContent=isExp?'−':'+';
}
function toggleComNotes(btn){
  const wrap=btn.parentElement;
  const id=+wrap.dataset.id;
  const isExp=wrap.classList.toggle('expanded');
  if(isExp)comNotesExpanded.add(id);else comNotesExpanded.delete(id);
  btn.textContent=isExp?'−':'+';
}
function updateNotesToggles(scope){
  document.querySelectorAll(scope+' .notes-wrap').forEach(w=>{
    const content=w.firstElementChild;
    const btn=w.querySelector('.notes-toggle');
    if(!content||!btn)return;
    if(w.classList.contains('expanded')){
      btn.textContent='−';
      btn.style.display='inline-block';
    } else {
      const overflow=content.scrollHeight>content.clientHeight+1;
      btn.textContent='+';
      btn.style.display=overflow?'inline-block':'none';
    }
  });
}

/* ═══════════════════════════════════════════════
   6. DROPDOWNS
═══════════════════════════════════════════════ */
function toggleDrop(e,wrap,cls){
  e.stopPropagation();
  const drop=wrap.querySelector('.'+cls);
  const isOpen=drop.classList.contains('open');
  document.querySelectorAll('.sit-drop.open,.ita-drop.open').forEach(x=>x.classList.remove('open'));
  if(!isOpen){
    const r=wrap.getBoundingClientRect();
    drop.style.top=r.bottom+'px';drop.style.left=r.left+'px';
    if(cls==='sit-drop')drop.style.minWidth=r.width+'px';
    drop.classList.add('open');
  }
}
document.addEventListener('click',()=>document.querySelectorAll('.sit-drop.open,.ita-drop.open').forEach(x=>x.classList.remove('open')));

/* ═══════════════════════════════════════════════
   7. PESTANYES
═══════════════════════════════════════════════ */
function showTab(id){
  document.querySelectorAll('.tab-btn').forEach((b,i)=>b.classList.toggle('active',['manual','bbdd','comercial','sch-tickets','sch-cupo','requeriments','tipus'][i]===id));
  document.querySelectorAll('.tab-content').forEach(d=>d.classList.remove('active'));
  document.getElementById('tab-'+id).classList.add('active');
  if(id==='comercial')renderComercial();
  if(id==='sch-tickets')renderSchTickets();
  if(id==='sch-cupo')renderSchCupo();
  if(id==='requeriments')renderRequeriments();
  if(id==='tipus')renderTipusTab();
}

/* ═══════════════════════════════════════════════
   8. SCHEDULES
═══════════════════════════════════════════════ */
function cupBadge(cup){
  const m={CEXP:'background:#EEEDFE;color:#534AB7',PI:'background:#E1F5EE;color:#0F6E56',EC:'background:#E6F1FB;color:#185FA5',IA:'background:#FAEEDA;color:#854F0B'};
  return cup?`<span class="sch-cup" style="${m[cup]||''}">${cup}</span>`:'';
}
function schCell(d){
  const canCheck=(cT==='a'&&adminLevel==='master')||(cT==='p'&&d.pres===cU);
  return `<div>
    <span class="sch-cell-emp">${d.tke?d.tke+'. ':''}${d.emp}</span>${cupBadge(d.cup)}
    <div class="sch-checks">
      <label><input type="checkbox"${d.presentat?' checked':''}${canCheck?'':' disabled'} onchange="svs(${d.id},'presentat',this.checked)"> Presentat</label>
      <label><input type="checkbox"${d.resguard?' checked':''}${canCheck&&d.presentat?'':' disabled'} onchange="svs(${d.id},'resguard',this.checked)"> Arxivat</label>
    </div>
  </div>`;
}
function renderSchTickets(){
  const presos=TD.filter(p=>D.some(d=>d.pres===p&&d.sit==='COMPLETAT'));
  document.getElementById('sch-tickets-head').innerHTML=`<tr><th>Hora</th>${presos.map(p=>`<th>${p}</th>`).join('')}</tr>`;
  document.getElementById('sch-tickets-body').innerHTML=SCH_HORES.map(h=>{
    const cells=presos.map(p=>{
      const m=D.filter(d=>d.hora===h&&d.pres===p&&d.sit==='COMPLETAT').sort((a,b)=>(parseInt(a.tke)||0)-(parseInt(b.tke)||0));
      return `<td>${m.map((d,i)=>(i>0?'<hr class="sch-sep">':'')+schCell(d)).join('')}</td>`;
    }).join('');
    return `<tr><td class="sch-hora">${h}</td>${cells}</tr>`;
  }).join('');
}
function renderSchCupo(){
  const CUPS_SCH=['CEXP','PI','EC','IA'];
  document.getElementById('sch-cupo-head').innerHTML=`<tr><th>Hora</th>${CUPS_SCH.map(c=>`<th>${cupBadge(c)}</th>`).join('')}</tr>`;
  document.getElementById('sch-cupo-body').innerHTML=SCH_HORES.map(h=>{
    const cells=CUPS_SCH.map(cup=>{
      const m=D.filter(d=>d.hora===h&&d.cup===cup&&d.sit==='COMPLETAT').sort((a,b)=>(parseInt(a.tke)||0)-(parseInt(b.tke)||0));
      return `<td>${m.map((d,i)=>(i>0?'<hr class="sch-sep">':'')+schCell(d)).join('')}</td>`;
    }).join('');
    return `<tr><td class="sch-hora">${h}</td>${cells}</tr>`;
  }).join('');
}

/* ═══════════════════════════════════════════════
   9. FILTRES
═══════════════════════════════════════════════ */
function setU(u,t){
  if(t==='a'&&u!==authedAs){adminLevel='';myAdminPw='';} // canviar a una altra identitat d'Admin obliga a re-autenticar-se
  if((t==='t'||t==='p'||t==='w'||t==='h')&&pins[u]&&!unlockedNames.has(u)){
    promptPin(u,t);
    return;
  }
  cU=u;cT=t;fWeb=false;fHora=false;
  document.getElementById('btnWeb').classList.remove('active');
  document.getElementById('btnHora').classList.remove('active');
  render();
}
function promptPin(u,t){
  if(document.getElementById('pin-dialog'))return;
  const el=document.createElement('div');
  el.id='pin-dialog';
  el.style='position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:99999;display:flex;align-items:center;justify-content:center';
  el.innerHTML=`<div style="background:#fff;border-radius:8px;padding:20px;min-width:240px;box-shadow:0 8px 32px rgba(0,0,0,.2)">
    <p style="font-size:10pt;font-weight:bold;margin-bottom:10px">PIN de ${u}</p>
    <input type="password" id="pin-inp" style="border:1px solid #ccc;padding:4px 6px;border-radius:3px;width:100px;margin-right:6px;">
    <button id="pin-ok-btn" style="padding:4px 12px;border:1px solid #bbb;border-radius:3px;cursor:pointer">Entrar</button>
    <span id="pin-err" style="color:red;font-size:9pt;display:none;margin-left:6px">Incorrecte</span>
    <div style="text-align:right;margin-top:10px"><span onclick="document.getElementById('pin-dialog').remove()" style="cursor:pointer;color:#999;font-size:9pt">Cancel·lar</span></div>
  </div>`;
  document.body.appendChild(el);
  const submit=()=>{
    const v=document.getElementById('pin-inp').value.trim();
    if(v&&v.toUpperCase()===String(pins[u]).toUpperCase()){
      unlockedNames.add(u);
      document.getElementById('pin-dialog').remove();
      setU(u,t);
    }else{
      document.getElementById('pin-err').style.display='inline';
      document.getElementById('pin-inp').value='';
    }
  };
  document.getElementById('pin-ok-btn').onclick=submit;
  document.getElementById('pin-inp').addEventListener('keydown',e=>{if(e.key==='Enter')submit();});
  setTimeout(()=>document.getElementById('pin-inp').focus(),50);
}
function toggleWeb(){
  fWeb=!fWeb;
  document.getElementById('btnWeb').classList.toggle('active',fWeb);
  render();
}
function toggleHora(){
  fHora=!fHora;
  document.getElementById('btnHora').classList.toggle('active',fHora);
  render();
}
function toggleNova(){
  fNova=!fNova;
  document.getElementById('btnNova').classList.toggle('active',fNova);
  render();
}
function scrollTable(px){
  const w=document.getElementById('table-wrap');
  if(!w)return;
  w.scrollLeft=w.scrollLeft+px;
}

/* ═══════════════════════════════════════════════
   10. BOTONS DE ROL
═══════════════════════════════════════════════ */
function rebuildRoleBtns(){
  document.getElementById('rbA').innerHTML=A.map(u=>`<button class="rbtn${u===cU&&cT==='a'?' active':''}" onclick="setU('${u}','a')">${u==='Admin'?'Menú':u}</button>`).join('');
  document.getElementById('rbG').innerHTML=G.map(u=>`<button class="rbtn${u===cU&&cT==='g'?' active':''}" onclick="setU('${u}','g')">${u}</button>`).join('');
  document.getElementById('rbT').innerHTML=PF.map(u=>`<button class="rbtn${u===cU&&cT==='t'?' active':''}" onclick="setU('${u}','t')">${u}</button>`).join('');
  document.getElementById('rbP').innerHTML=TD.map(u=>`<button class="rbtn${u===cU&&cT==='p'?' active':''}" onclick="setU('${u}','p')">${u}</button>`).join('');
  document.getElementById('rbW').innerHTML=QW.map(u=>`<button class="rbtn${u===cU&&cT==='w'?' active':''}" onclick="setU('${u}','w')">${u}</button>`).join('');
  document.getElementById('rbH').innerHTML=TFH.map(u=>`<button class="rbtn${u===cU&&cT==='h'?' active':''}" onclick="setU('${u}','h')">${u}</button>`).join('');
}
function rebuildAdminBtns(){
  document.getElementById('rbA').innerHTML=A.map(u=>`<button class="rbtn${u===cU&&cT==='a'?' active':''}" onclick="setU('${u}','a')">${u==='Admin'?'Menú':u}</button>`).join('');
}

/* ═══════════════════════════════════════════════
   11. PANELL ADMIN
═══════════════════════════════════════════════ */
function togglePanel(id){
  const p=document.getElementById('panel-'+id);
  const ov=document.getElementById('overlay');
  if(p.classList.contains('open')){closeAllPanels();return;}
  closeAllPanels();
  if(id==='admin'){
    if(adminLevel){
      document.getElementById('admin-pw-zone').style.display='none';
      document.getElementById('admin-list-wrap').style.display='block';
      renderAdminList();
    }else{
      document.getElementById('admin-pw-zone').style.display='block';
      document.getElementById('admin-list-wrap').style.display='none';
      document.getElementById('pw-err').style.display='none';
      document.getElementById('pw-inp').value='';
      setTimeout(()=>document.getElementById('pw-inp').focus(),50);
    }
  }else if(id==='llistes')renderLlistes();
  p.classList.add('open');ov.classList.add('open');
}
function closeAllPanels(){
  document.querySelectorAll('.panel').forEach(p=>p.classList.remove('open'));
  document.getElementById('overlay').classList.remove('open');
}
function checkPw(){
  const v=document.getElementById('pw-inp').value;
  if(cU==='Admin'){
    if(v!==ADMIN_PW_MASTER){
      document.getElementById('pw-err').style.display='inline';
      document.getElementById('pw-inp').value='';
      return;
    }
    adminLevel='master';myAdminPw=v;authedAs=cU;
  }else{
    if(!ADMIN_PWS[cU]||v!==ADMIN_PWS[cU]){
      document.getElementById('pw-err').style.display='inline';
      document.getElementById('pw-inp').value='';
      return;
    }
    adminLevel='admin';myAdminPw=v;authedAs=cU;
  }
  document.getElementById('pw-err').style.display='none';
  document.getElementById('admin-pw-zone').style.display='none';
  document.getElementById('admin-list-wrap').style.display='block';
  renderAdminList();render();
}
function adminLogout(){
  adminLevel='';myAdminPw='';authedAs='';
  closeAllPanels();
  render();
}
function renderAdminList(){
  const canEdit=adminLevel==='master';
  document.getElementById('admin-list').innerHTML=A.map((u,i)=>`<div class="admin-item">
    <span>${u}${canEdit&&i>0?` <span style="font-size:8pt;color:#888">(${ADMIN_PWS[u]||''})</span>`:''}</span>
    ${canEdit?(i>0?`<button class="llista-del" title="Canviar contrasenya" onclick="resetAdminPw('${u}')" style="border:none;background:none;cursor:pointer;font-size:11pt;margin-right:4px">🔑</button><button class="llista-del" title="Eliminar" onclick="removeAdmin('${u}')">×</button>`:'<span style="font-size:8pt;color:#aaa">principal (Master)</span>'):''}
  </div>`).join('');
  const az=document.getElementById('admin-add-zone');
  const pz=document.getElementById('admin-pw-change-zone');
  if(az)az.style.display=canEdit?'flex':'none';
  if(pz)pz.style.display=canEdit?'block':'none';
}
async function addAdmin(){
  if(adminLevel!=='master')return;
  const v=document.getElementById('new-admin').value.trim();
  const pw=document.getElementById('new-admin-pw').value.trim();
  if(!v||!pw||A.includes(v))return;
  A.push(v);ADMIN_PWS[v]=pw;
  document.getElementById('new-admin').value='';document.getElementById('new-admin-pw').value='';
  rebuildAdminBtns();renderAdminList();
  try{await sbAddAdmin(v,pw);}catch(e){showError(e.message);}
}
async function removeAdmin(u){
  if(adminLevel!=='master')return;
  if(A.length<=1)return;
  A=A.filter(x=>x!==u);delete ADMIN_PWS[u];rebuildAdminBtns();renderAdminList();
  try{await sbRemoveAdmin(u);}catch(e){showError(e.message);}
}
async function resetAdminPw(u){
  if(adminLevel!=='master')return;
  const v=prompt(`Nova contrasenya per a ${u}:`);
  if(!v)return;
  ADMIN_PWS[u]=v;
  try{await sbSetAdminPassword(u,v);showSaved();}catch(e){showError(e.message);}
}
async function changePw(){
  if(adminLevel!=='master')return;
  const v=document.getElementById('new-pw').value.trim();
  if(!v)return;
  ADMIN_PW_MASTER=v;document.getElementById('new-pw').value='';document.getElementById('pw-ok').style.display='inline';setTimeout(()=>document.getElementById('pw-ok').style.display='none',2000);
  try{await setMasterPw(v);}catch(e){console.warn('Error desant contrasenya Master:',e);}
}

/* ═══════════════════════════════════════════════
   12. PANELL LLISTES
═══════════════════════════════════════════════ */
function renderLlistes(){
  const llistes=[
    {title:'KAM',arr:G,key:'g',note:''},
    {title:'Presentadors Fitxa',arr:PF,key:'pf',note:'(= Qui Omplir)'},
    {title:'Tramitadors dia-D',arr:TD,key:'td',note:'(= Qui Presenta)'},
    {title:'Gestor Web',arr:QW,key:'qw',note:'(= Qui Web)'},
    {title:'Gestor Hora',arr:TFH,key:'tfh',note:'(= Qui Hora)'},
    {title:'Cupons',arr:CUPS.filter(x=>x),key:'c',note:''},
    {title:'Estat LEAD',arr:SITS.filter(x=>x),key:'s',note:''},
    {title:'Web',arr:WEBLIST,key:'web',note:''},
    {title:'Qui — Valid.',arr:TF3,key:'tf3',note:'(llista pròpia)'},
    {title:'Atorgat',arr:OTORGAT_OPTS.filter(x=>x),key:'otorgat',note:''},
    {title:'Estat requeriment',arr:REQ_ESTAT_OPTS.filter(x=>x),key:'reqestat',note:''},
    {title:'Resolució final',arr:RESOLUCIO_OPTS.filter(x=>x),key:'resolucio',note:''},
    {title:'Consultors (email)',arr:CONSULTORS,key:'consultors',note:''},
    {title:'Consultor',arr:CONS,key:'cons',note:'(desplegable a Comercial)'},
  ];
  document.getElementById('llistes-contingut').innerHTML=`<div class="llistes-grid">${llistes.map(l=>`
    <div class="llista-grp">
      <h4>${l.title}${l.note?` <span class="llista-note">${l.note}</span>`:''}</h4>
      ${l.arr.map(v=>{
        const pinField=(l.key==='pf'||l.key==='td'||l.key==='qw'||l.key==='tfh')?`<input type="text" value="${pins[v]||''}" placeholder="PIN" class="llista-pin" onchange="setPin('${v}',this.value.trim())">`:'';
        return `<div class="llista-item"><span>${v}</span><span style="display:flex;align-items:center">${pinField}<button class="llista-del" onclick="removeFromList('${l.key}','${v}')">×</button></span></div>`;
      }).join('')}
      <div class="llista-add">
        <input type="text" id="inp-${l.key}" placeholder="Afegir…">
        <button onclick="addToList('${l.key}')">+</button>
      </div>
    </div>`).join('')}</div>`;
}
function addToList(key){
  const el=document.getElementById('inp-'+key);
  const raw=el.value.trim();
  const v=key==='consultors'?raw.toLowerCase():raw.toUpperCase();
  if(!v)return;
  let arr=null;
  if(key==='g'&&!G.includes(v)){G.push(v);arr=G;}
  else if(key==='pf'&&!PF.includes(v)){PF.push(v);rebuildRoleBtns();arr=PF;}
  else if(key==='td'&&!TD.includes(v)){TD.push(v);rebuildRoleBtns();arr=TD;}
  else if(key==='qw'&&!QW.includes(v)){QW.push(v);rebuildRoleBtns();arr=QW;}
  else if(key==='c'&&!CUPS.includes(v)){CUPS.push(v);arr=CUPS;}
  else if(key==='s'&&!SITS.includes(v)){SITS.push(v);arr=SITS;}
  else if(key==='web'&&!WEBLIST.includes(v)){WEBLIST.push(v);arr=WEBLIST;}
  else if(key==='tf3'&&!TF3.includes(v)){TF3.push(v);arr=TF3;}
  else if(key==='tfh'&&!TFH.includes(v)){TFH.push(v);rebuildRoleBtns();arr=TFH;}
  else if(key==='otorgat'&&!OTORGAT_OPTS.includes(v)){OTORGAT_OPTS.push(v);arr=OTORGAT_OPTS;}
  else if(key==='reqestat'&&!REQ_ESTAT_OPTS.includes(v)){REQ_ESTAT_OPTS.push(v);arr=REQ_ESTAT_OPTS;}
  else if(key==='resolucio'&&!RESOLUCIO_OPTS.includes(v)){RESOLUCIO_OPTS.push(v);arr=RESOLUCIO_OPTS;}
  else if(key==='consultors'&&!CONSULTORS.includes(v)){CONSULTORS.push(v);arr=CONSULTORS;}
  else if(key==='cons'&&!CONS.includes(v)){CONS.push(v);arr=CONS;}
  el.value='';renderLlistes();render();renderRequerimentsTable();renderComercial();
  if(arr)setList(key,arr).catch(e=>console.warn('Error desant llista:',e));
}
function removeFromList(key,v){
  let arr=null;
  if(key==='g'){G=G.filter(x=>x!==v);arr=G;}
  else if(key==='pf'){PF=PF.filter(x=>x!==v);rebuildRoleBtns();arr=PF;}
  else if(key==='td'){TD=TD.filter(x=>x!==v);rebuildRoleBtns();arr=TD;}
  else if(key==='qw'){QW=QW.filter(x=>x!==v);rebuildRoleBtns();arr=QW;}
  else if(key==='c'){CUPS=CUPS.filter(x=>x!==v&&x!=='');if(!CUPS.includes(''))CUPS.unshift('');arr=CUPS;}
  else if(key==='s'){SITS=SITS.filter(x=>x!==v&&x!=='');if(!SITS.includes(''))SITS.unshift('');arr=SITS;}
  else if(key==='web'){WEBLIST=WEBLIST.filter(x=>x!==v);arr=WEBLIST;}
  else if(key==='tf3'){TF3=TF3.filter(x=>x!==v);arr=TF3;}
  else if(key==='tfh'){TFH=TFH.filter(x=>x!==v);rebuildRoleBtns();arr=TFH;}
  else if(key==='otorgat'){OTORGAT_OPTS=OTORGAT_OPTS.filter(x=>x!==v&&x!=='');if(!OTORGAT_OPTS.includes(''))OTORGAT_OPTS.unshift('');arr=OTORGAT_OPTS;}
  else if(key==='reqestat'){REQ_ESTAT_OPTS=REQ_ESTAT_OPTS.filter(x=>x!==v&&x!=='');if(!REQ_ESTAT_OPTS.includes(''))REQ_ESTAT_OPTS.unshift('');arr=REQ_ESTAT_OPTS;}
  else if(key==='resolucio'){RESOLUCIO_OPTS=RESOLUCIO_OPTS.filter(x=>x!==v&&x!=='');if(!RESOLUCIO_OPTS.includes(''))RESOLUCIO_OPTS.unshift('');arr=RESOLUCIO_OPTS;}
  else if(key==='consultors'){CONSULTORS=CONSULTORS.filter(x=>x!==v);arr=CONSULTORS;}
  else if(key==='cons'){CONS=CONS.filter(x=>x!==v);arr=CONS;}
  renderLlistes();render();renderRequerimentsTable();renderComercial();
  if(arr)setList(key,arr).catch(e=>console.warn('Error desant llista:',e));
}

/* ═══════════════════════════════════════════════
   13. NEON — CONNEXIÓ I PERSISTÈNCIA
═══════════════════════════════════════════════ */
const NEON_CONN='postgresql://neondb_owner:npg_wCyd8sYr0cBb@ep-rapid-butterfly-abxlbyao-pooler.eu-west-2.aws.neon.tech/neondb?sslmode=require&channel_binding=require';
let neonSql=null;
async function initNeon(){if(neonSql)return;const{neon}=await import('https://esm.sh/@neondatabase/serverless');neonSql=neon(NEON_CONN);}
async function neonQuery(sql,params=[]){await initNeon();const res=params.length?await neonSql.query(sql,params):await neonSql.query(sql);return{rows:res.rows||res};}

async function getMasterPw(){
  const r=await neonQuery('SELECT pw FROM admin_config WHERE id=1');
  return r.rows[0]?.pw||'';
}
async function setMasterPw(v){
  await neonQuery('INSERT INTO admin_config(id,pw) VALUES(1,$1) ON CONFLICT(id) DO UPDATE SET pw=$1',[v]);
}
async function getAllLists(){
  const r=await neonQuery('SELECT key,items FROM app_lists');
  const m={};
  r.rows.forEach(row=>{m[row.key]=typeof row.items==='string'?JSON.parse(row.items):row.items;});
  return m;
}
async function setList(key,items){
  await neonQuery(`INSERT INTO app_lists(key,items) VALUES($1,$2) ON CONFLICT(key) DO UPDATE SET items=$2`,[key,JSON.stringify(items)]);
}
async function getPins(){
  const r=await neonQuery('SELECT name,pin FROM user_pins');
  const m={};
  r.rows.forEach(row=>{if(row.pin)m[row.name]=row.pin;});
  return m;
}
async function setPin(name,pin){
  pins[name]=pin;
  try{await neonQuery(`INSERT INTO user_pins(name,pin) VALUES($1,$2) ON CONFLICT(name) DO UPDATE SET pin=$2`,[name,pin]);}
  catch(e){console.warn('Error desant PIN:',e);}
}

/* ═══════════════════════════════════════════════
   15. REQUERIMENTS (seguiment post-presentació)
═══════════════════════════════════════════════ */
const REQ_FIELDS=['expedient','edicte_r','tipus_ids','aclariment_tecnic','comentaris_backoffice','dead_line','estat',
  'data_presentacio','comentaris_kam','resolucio_final','data_resolucio','todo','email_consultor','email_consultor_cc'];
const TIPUS_FIELDS=['alies','pregunta','comentari','recurs_nom','recurs_link'];

function isMasterActive(){return cT==='a'&&adminLevel==='master';}

async function getRequeriments(){
  const res=await neonQuery(`SELECT r.*, c.emp AS c_emp, c.cup AS c_cup, c.tke AS c_tke FROM ${YEAR_CONFIG.reqTable} r LEFT JOIN ${YEAR_CONFIG.table} c ON c.id=r.cupo_id ORDER BY r.id ASC`);
  return res.rows||[];
}
async function getRequerimentTipus(){
  const res=await neonQuery(`SELECT * FROM requeriment_tipus ORDER BY id`);
  return res.rows||[];
}
async function getCompletatsForSelect(){
  const res=await neonQuery(`SELECT id, emp, cup, tke FROM ${YEAR_CONFIG.table} WHERE sit='COMPLETAT' ORDER BY emp ASC`);
  return res.rows||[];
}
async function insertRequerimentRow(cupoId){
  const res=await neonQuery(`INSERT INTO ${YEAR_CONFIG.reqTable} (cupo_id) VALUES ($1) RETURNING id`,[cupoId]);
  return res.rows?.[0]?.id;
}
async function updateRequerimentField(id,f,v){
  if(!REQ_FIELDS.includes(f))return;
  await neonQuery(`UPDATE ${YEAR_CONFIG.reqTable} SET ${f}=$1, updated_at=NOW() WHERE id=$2`,[v,id]);
}
async function deleteRequerimentRow(id){
  await neonQuery(`DELETE FROM ${YEAR_CONFIG.reqTable} WHERE id=$1`,[id]);
}
async function insertTipusRow(alies,pregunta,comentari,recursNom,recursLink){
  const res=await neonQuery(`INSERT INTO requeriment_tipus(alies,pregunta,comentari,recurs_nom,recurs_link) VALUES($1,$2,$3,$4,$5) RETURNING id`,[alies,pregunta,comentari,recursNom,recursLink]);
  return res.rows?.[0]?.id;
}
async function updateTipusField(id,f,v){
  if(!TIPUS_FIELDS.includes(f))return;
  await neonQuery(`UPDATE requeriment_tipus SET ${f}=$1 WHERE id=$2`,[v,id]);
}
async function deleteTipusRow(id){
  await neonQuery(`DELETE FROM requeriment_tipus WHERE id=$1`,[id]);
}

function parseTipusIds(v){
  if(Array.isArray(v))return v;
  if(typeof v==='string'&&v)return JSON.parse(v);
  return [];
}
function ecReq(r,f,v,multi){
  if(!isMasterActive())return `<span>${v||''}</span>`;
  return `<span class="ed" contenteditable="true" data-rid="${r.id}" data-rf="${f}" onblur="svReq(this)" onkeydown="if(event.key==='Enter'&&!${!!multi}){event.preventDefault();this.blur()}">${v||''}</span>`;
}
const REQ_DATE_FIELDS=['dead_line','data_presentacio','data_resolucio'];
async function svReq(el){
  const id=+el.dataset.rid, f=el.dataset.rf;
  let v=el.textContent.trim();
  if(REQ_DATE_FIELDS.includes(f))v=fmtDate(v).replace(/\./g,'/');
  const r=REQ.find(x=>x.id===id); if(r)r[f]=v;
  await updateRequerimentField(id,f,v);
  renderRequerimentsTable();
}
async function svsReq(id,f,v){
  const r=REQ.find(x=>x.id===id); if(r)r[f]=v;
  await updateRequerimentField(id,f,v);
  renderRequerimentsTable();
}
function selReqEstat(r){
  const c=REQ_ESTAT_COLORS[r.estat]||{};
  const bg=c.bg||'',fg=c.fg||'#222';
  const label=r.estat||'N/A';
  if(!isMasterActive())return `<div class="sel-small" style="background:${bg};color:${fg};padding:2px 4px;border-radius:3px">${label}</div>`;
  const style=bg?` style="background:${bg};color:${fg}"`:'';
  return `<select class="sel-small"${style} onchange="svsReq(${r.id},'estat',this.value)">${REQ_ESTAT_OPTS.map(o=>`<option value="${o}"${r.estat===o?' selected':''}>${o||'N/A'}</option>`).join('')}</select>`;
}
function selResolucioFinal(r){
  const c=RESOLUCIO_COLORS[r.resolucio_final]||{};
  const bg=c.bg||'',fg=c.fg||'#222';
  if(!isMasterActive())return `<span class="sel-small" style="background:${bg};color:${fg};padding:2px 4px;border-radius:3px">${r.resolucio_final||''}</span>`;
  const style=bg?` style="background:${bg};color:${fg}"`:'';
  return `<select class="sel-small"${style} onchange="svsReqResolucio(${r.id},this.value)">${RESOLUCIO_OPTS.map(o=>`<option value="${o}"${r.resolucio_final===o?' selected':''}>${o||'—'}</option>`).join('')}</select>`;
}
async function svsReqResolucio(id,val){
  const r=REQ.find(x=>x.id===id); if(!r)return;
  r.resolucio_final=val;
  await updateRequerimentField(id,'resolucio_final',val);
  if(val in RESOLUCIO_TO_ESTAT){
    r.estat=RESOLUCIO_TO_ESTAT[val];
    await updateRequerimentField(id,'estat',r.estat);
  }
  renderRequerimentsTable();
}
function reqTodoCell(r){
  const master=isMasterActive();
  const bg=r.todo?'background:#FF3333;':'';
  return `<td style="text-align:center;${bg}"><input type="checkbox"${r.todo?' checked':''}${master?'':' disabled'} onchange="svsReq(${r.id},'todo',this.checked)"></td>`;
}
function reqEmailSel(r,field,val,master){
  if(!master)return `<span style="max-width:190px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;display:inline-block">${val||''}</span>`;
  return `<select class="sel-small" style="max-width:190px;width:100%" onchange="svsReq(${r.id},'${field}',this.value)"><option value=""${val?'':' selected'}>—</option>${CONSULTORS.map(c=>`<option value="${c}"${val===c?' selected':''}>${c}</option>`).join('')}</select>`;
}
function reqConsultorCell(r){
  const master=isMasterActive();
  const hasEmail=!!(r.email_consultor&&r.email_consultor.trim());
  return `<td><div style="display:flex;flex-direction:column;gap:2px;min-width:215px">
    <div style="display:flex;align-items:center;gap:4px">${reqEmailSel(r,'email_consultor',r.email_consultor,master)}<button type="button" onclick="sendReqEmail(${r.id})"${hasEmail?'':' disabled'} title="Enviar TODO per correu al consultor" style="border:none;background:none;cursor:${hasEmail?'pointer':'not-allowed'};font-size:12pt;opacity:${hasEmail?'1':'.35'};flex-shrink:0">✉️</button></div>
    <div style="display:flex;align-items:center;gap:4px"><span style="font-size:7pt;color:#888;width:16px;flex-shrink:0">CC</span>${reqEmailSel(r,'email_consultor_cc',r.email_consultor_cc,master)}</div>
  </div></td>`;
}
function sendReqEmail(id){
  const r=REQ.find(x=>x.id===id);
  if(!r||!r.email_consultor)return;
  const tipus=parseTipusIds(r.tipus_ids).map(tid=>{const t=REQ_TIPUS.find(x=>x.id===tid);return t?t.alies.split(',')[0]:'';}).filter(Boolean).join(', ');
  const subject=`Requeriment ${r.c_emp||''} — Tiquet ${r.c_tke||''}`;
  const body=[
    `Empresa: ${r.c_emp||''}`,
    `Tiquet: ${r.c_tke||''}`,
    `Cupó: ${r.c_cup||''}`,
    `Expedient: ${r.expedient||''}`,
    `# Edicte R: ${r.edicte_r||''}`,
    `Tipus: ${tipus}`,
    `Aclariment tècnic: ${r.aclariment_tecnic||''}`,
    `Comentaris Back Office: ${r.comentaris_backoffice||''}`,
    `Dead line: ${r.dead_line||''}`,
    `Estat requeriment: ${r.estat||''}`,
  ].join('\n');
  const cc=r.email_consultor_cc?`&cc=${encodeURIComponent(r.email_consultor_cc)}`:'';
  window.location.href=`mailto:${r.email_consultor}?subject=${encodeURIComponent(subject)}${cc}&body=${encodeURIComponent(body)}`;
}
function reqWideCell(r,f,innerHtml){
  const key=r.id+':'+f;
  const exp=reqNotesExpanded.has(key)?' expanded':'';
  return `<td class="rq-wide"><div class="notes-wrap${exp}" data-rid="${r.id}" data-rf="${f}">${innerHtml}<button type="button" class="notes-toggle" onclick="toggleReqNotes(this)">+</button></div></td>`;
}
function tipusBadges(r){
  const ids=parseTipusIds(r.tipus_ids);
  const names=ids.map(tid=>{const t=REQ_TIPUS.find(x=>x.id===tid);return t?t.alies.split(',')[0]:'';}).filter(Boolean);
  const master=isMasterActive();
  const badges=names.map(n=>`<span style="display:inline-block;background:#EEEDFE;color:#534AB7;border-radius:3px;padding:1px 5px;margin:1px;font-size:8pt">${n}</span>`).join('');
  if(!master)return `<div style="min-width:80px">${badges}</div>`;
  return `<div style="cursor:pointer;min-width:80px" onclick="openTipusPicker(${r.id})">${badges||'<span style="color:#aaa;font-size:8pt">+ afegir</span>'}</div>`;
}
function openTipusPicker(rid){
  if(!isMasterActive())return;
  if(document.getElementById('tipus-picker-dialog'))return;
  const r=REQ.find(x=>x.id===rid); if(!r)return;
  const ids=parseTipusIds(r.tipus_ids);
  const el=document.createElement('div');
  el.id='tipus-picker-dialog';
  el.style='position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:99999;display:flex;align-items:center;justify-content:center';
  el.innerHTML=`<div style="background:#fff;border-radius:8px;padding:20px;max-width:500px;max-height:80vh;overflow-y:auto;box-shadow:0 8px 32px rgba(0,0,0,.2)">
    <p style="font-weight:bold;margin-bottom:10px">Tipus de requeriment</p>
    ${REQ_TIPUS.map(t=>`<label style="display:block;margin-bottom:6px;font-size:9.5pt"><input type="checkbox" value="${t.id}" ${ids.includes(t.id)?'checked':''}> ${t.alies}</label>`).join('')}
    <div style="text-align:right;margin-top:14px">
      <button onclick="saveTipusPicker(${rid})" style="padding:5px 14px;border:1px solid #2c5aa0;background:#2c5aa0;color:#fff;border-radius:4px;cursor:pointer">Desar</button>
      <button onclick="document.getElementById('tipus-picker-dialog').remove()" style="padding:5px 14px;border:1px solid #bbb;border-radius:4px;cursor:pointer;background:#f5f5f5;margin-left:6px">Cancel·lar</button>
    </div>
  </div>`;
  document.body.appendChild(el);
}
async function saveTipusPicker(rid){
  const dialog=document.getElementById('tipus-picker-dialog');
  const ids=[...dialog.querySelectorAll('input[type=checkbox]:checked')].map(x=>+x.value);
  const r=REQ.find(x=>x.id===rid); if(r)r.tipus_ids=ids;
  dialog.remove();
  renderRequerimentsTable();
  await updateRequerimentField(rid,'tipus_ids',JSON.stringify(ids));
}
async function addRequeriment(){
  if(!isMasterActive())return;
  if(document.getElementById('nou-req-dialog'))return;
  const completats=await getCompletatsForSelect();
  const el=document.createElement('div');
  el.id='nou-req-dialog';
  el.style='position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:99999;display:flex;align-items:center;justify-content:center';
  el.innerHTML=`<div style="background:#fff;border-radius:8px;padding:20px;min-width:340px;box-shadow:0 8px 32px rgba(0,0,0,.2)">
    <p style="font-weight:bold;margin-bottom:10px">Nou requeriment — tria l'empresa/cupó (només COMPLETATS)</p>
    <select id="nou-req-sel" style="width:100%;padding:5px">
      ${completats.map(c=>`<option value="${c.id}">${c.emp} — ${c.cup} (Tiquet ${c.tke||'—'})</option>`).join('')}
    </select>
    <div style="text-align:right;margin-top:14px">
      <button onclick="confirmNouReq()" style="padding:5px 14px;border:1px solid #2c5aa0;background:#2c5aa0;color:#fff;border-radius:4px;cursor:pointer">Crear</button>
      <button onclick="document.getElementById('nou-req-dialog').remove()" style="padding:5px 14px;border:1px solid #bbb;border-radius:4px;cursor:pointer;background:#f5f5f5;margin-left:6px">Cancel·lar</button>
    </div>
  </div>`;
  document.body.appendChild(el);
}
async function confirmNouReq(){
  const sel=document.getElementById('nou-req-sel');
  const cupoId=+sel.value;
  document.getElementById('nou-req-dialog').remove();
  await insertRequerimentRow(cupoId);
  const d=D.find(x=>x.id===cupoId);
  if(d&&d.otorgat!=='REQUERIT'){
    d.otorgat='REQUERIT';
    try{await sbUpdate(d.id,d);}catch(e){console.warn('No s\'ha pogut marcar Atorgat com a REQUERIT:',e);}
    renderComercial();
  }
  renderRequeriments();
}
async function delRequerimentRow(id){
  if(!confirm('Eliminar aquest requeriment?'))return;
  const req=REQ.find(x=>x.id===id);
  const cupoId=req?req.cupo_id:null;
  await deleteRequerimentRow(id);
  REQ=REQ.filter(x=>x.id!==id);
  if(cupoId&&!REQ.some(x=>x.cupo_id===cupoId)){
    const d=D.find(x=>x.id===cupoId);
    if(d&&d.otorgat==='REQUERIT'){
      d.otorgat='';
      try{await sbUpdate(d.id,d);}catch(e){console.warn('No s\'ha pogut buidar Atorgat:',e);}
      renderComercial();
    }
  }
  renderRequerimentsTable();
}
function reqRowHtml(r){
  const master=isMasterActive();
  return `<tr>
    <td style="text-align:center">${master?`<button onclick="delRequerimentRow(${r.id})" style="border:none;background:none;cursor:pointer;color:#c00;font-size:13pt" title="Eliminar">×</button>`:''}</td>
    <td>${r.c_emp||''}</td>
    <td style="text-align:center">${r.c_tke||''}</td>
    <td style="text-align:center">${r.c_cup||''}</td>
    <td>${ecReq(r,'expedient',r.expedient)}</td>
    <td class="cedicte">${ecReq(r,'edicte_r',r.edicte_r)}</td>
    ${reqWideCell(r,'tipus',tipusBadges(r))}
    ${reqWideCell(r,'aclariment_tecnic',ecReq(r,'aclariment_tecnic',r.aclariment_tecnic,true))}
    ${reqWideCell(r,'comentaris_backoffice',ecReq(r,'comentaris_backoffice',r.comentaris_backoffice,true))}
    <td class="rq-date">${ecReq(r,'dead_line',r.dead_line)}</td>
    <td>${selReqEstat(r)}</td>
    ${reqTodoCell(r)}
    ${reqConsultorCell(r)}
    <td class="rq-date">${ecReq(r,'data_presentacio',r.data_presentacio)}</td>
    ${reqWideCell(r,'comentaris_kam',ecReq(r,'comentaris_kam',r.comentaris_kam,true))}
    <td>${selResolucioFinal(r)}</td>
    <td class="rq-date">${ecReq(r,'data_resolucio',r.data_resolucio)}</td>
  </tr>`;
}
function renderRequerimentsTable(){
  document.getElementById('req-tbody').innerHTML=REQ.map(r=>reqRowHtml(r)).join('');
  requestAnimationFrame(()=>{
    updateReqStickyOffsets();
    updateNotesToggles('#req-table-wrap');
  });
}
async function renderRequeriments(){
  const master=isMasterActive();
  const hint=document.getElementById('req-hint');
  const actionsGrp=document.getElementById('req-actions-grp');
  if(actionsGrp)actionsGrp.style.display=master?'block':'none';
  hint.textContent=master?'Seguiment de requeriments post-presentació (només registres COMPLETAT).':'Seguiment de requeriments post-presentació — només lectura.';
  const[reqs,tipus]=await Promise.all([getRequeriments(),getRequerimentTipus()]);
  REQ=reqs;REQ_TIPUS=tipus;
  renderRequerimentsTable();
}
function ecTipus(t,f,v,multi){
  if(!isMasterActive())return `<span>${v||''}</span>`;
  return `<span class="ed" contenteditable="true" data-tid="${t.id}" data-tf="${f}" onblur="svTipus(this)" onkeydown="if(event.key==='Enter'&&!${!!multi}){event.preventDefault();this.blur()}">${v||''}</span>`;
}
async function svTipus(el){
  const id=+el.dataset.tid, f=el.dataset.tf;
  const v=el.textContent.trim();
  const t=REQ_TIPUS.find(x=>x.id===id); if(t)t[f]=v;
  await updateTipusField(id,f,v);
}
function tipusRowHtml(t){
  const master=isMasterActive();
  return `<tr>
    <td style="text-align:center">${master?`<button onclick="delTipusUI(${t.id})" style="border:none;background:none;cursor:pointer;color:#c00;font-size:13pt" title="Eliminar">×</button>`:''}</td>
    <td style="text-align:center;color:#555">${t.id}</td>
    <td>${ecTipus(t,'alies',t.alies)}</td>
    <td class="ncell">${ecTipus(t,'pregunta',t.pregunta,true)}</td>
    <td class="ncell">${ecTipus(t,'comentari',t.comentari,true)}</td>
    <td>${ecTipus(t,'recurs_nom',t.recurs_nom)}</td>
    <td>${ecTipus(t,'recurs_link',t.recurs_link)}</td>
  </tr>`;
}
function renderTipusTable(){
  document.getElementById('tipus-tbody').innerHTML=REQ_TIPUS.map(t=>tipusRowHtml(t)).join('');
}
async function renderTipusTab(){
  const master=isMasterActive();
  const actionsGrp=document.getElementById('tipus-actions-grp');
  if(actionsGrp)actionsGrp.style.display=master?'block':'none';
  document.getElementById('tipus-hint').textContent=master?'FAQs — catàleg de tipus de requeriment (compartit entre 2026 i 2027).':'FAQs — només lectura.';
  REQ_TIPUS=await getRequerimentTipus();
  renderTipusTable();
}
async function delTipusUI(id){
  if(!isMasterActive())return;
  if(!confirm('Eliminar aquest tipus del catàleg?'))return;
  await deleteTipusRow(id);
  REQ_TIPUS=REQ_TIPUS.filter(x=>x.id!==id);
  renderTipusTable();
}
async function addTipusForm(){
  if(!isMasterActive())return;
  await insertTipusRow('Nou tipus','','','','');
  REQ_TIPUS=await getRequerimentTipus();
  renderTipusTable();
}

async function sbGet(){
  const res=await neonQuery(`SELECT * FROM ${YEAR_CONFIG.table} ORDER BY id ASC`);
  return res.rows||[];
}
async function sbInsert(d){
  const s=rowToSb(d);
  const res=await neonQuery(
    `INSERT INTO ${YEAR_CONFIG.table} (g,emp,nova,cup,cup_ant,otorgat,cont,mob,email,url_web,url_web_check,reg,sit,tke,f1q,f1d,f2q,f2d,f3q,f3d,web,webq,ita,fhq,hora,pres,presentat,resguard,notes,proc_comercial_previ,te_requeriment,primer_contacte,proposta_preparada,proposta_presentada,proposta_enviada,oferta_acceptada,kickoff_esperat,holded,notes_comercial,consultor,todo,email_consultor,email_consultor_cc,edicte_od) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,$25,$26,$27,$28,$29,$30,$31,$32,$33,$34,$35,$36,$37,$38,$39,$40,$41,$42,$43,$44) RETURNING *`,
    [s.g,s.emp,s.nova,s.cup,s.cup_ant,s.otorgat,s.cont,s.mob,s.email,s.url_web,s.url_web_check,s.reg,s.sit,s.tke,s.f1q,s.f1d,s.f2q,s.f2d,s.f3q,s.f3d,s.web,s.webq,s.ita,s.fhq,s.hora,s.pres,s.presentat,s.resguard,s.notes,s.proc_comercial_previ,s.te_requeriment,s.primer_contacte,s.proposta_preparada,s.proposta_presentada,s.proposta_enviada,s.oferta_acceptada,s.kickoff_esperat,s.holded,s.notes_comercial,s.consultor,s.todo,s.email_consultor,s.email_consultor_cc,s.edicte_od]
  );
  return res.rows?.[0];
}
async function sbUpdate(id,d){
  const s=rowToSb(d);
  // Comprovació de conflicte: algú altre pot haver modificat aquest registre mentre l'editaves
  const chk=await neonQuery(`SELECT updated_at FROM ${YEAR_CONFIG.table} WHERE id=$1`,[id]);
  const serverTs=chk.rows?.[0]?.updated_at;
  if(d.updated_at&&serverTs&&new Date(serverTs).getTime()!==new Date(d.updated_at).getTime()){
    const overwrite=confirm("⚠️ Aquest registre l'ha modificat un altre usuari mentre l'editaves.\n\nAcceptar = sobreescriure amb els teus canvis\nCancel·lar = descartar els teus canvis i recarregar les dades actuals");
    if(!overwrite){
      const fresh=await neonQuery(`SELECT * FROM ${YEAR_CONFIG.table} WHERE id=$1`,[id]);
      if(fresh.rows?.[0]){
        const idx=D.findIndex(x=>x.id===id);
        if(idx>-1){D[idx]=sbToRow(fresh.rows[0]);calcTickets();render();}
      }
      throw new Error('CANCELLED_BY_USER');
    }
  }
  const res=await neonQuery(
    `UPDATE ${YEAR_CONFIG.table} SET g=$1,emp=$2,nova=$3,cup=$4,cup_ant=$5,otorgat=$6,cont=$7,mob=$8,email=$9,url_web=$10,url_web_check=$11,reg=$12,sit=$13,tke=$14,f1q=$15,f1d=$16,f2q=$17,f2d=$18,f3q=$19,f3d=$20,web=$21,webq=$22,ita=$23,fhq=$24,hora=$25,pres=$26,presentat=$27,resguard=$28,notes=$29,proc_comercial_previ=$30,te_requeriment=$31,primer_contacte=$32,proposta_preparada=$33,proposta_presentada=$34,proposta_enviada=$35,oferta_acceptada=$36,kickoff_esperat=$37,holded=$38,notes_comercial=$39,consultor=$40,todo=$41,email_consultor=$42,email_consultor_cc=$43,edicte_od=$44,updated_at=NOW() WHERE id=$45 RETURNING updated_at`,
    [s.g,s.emp,s.nova,s.cup,s.cup_ant,s.otorgat,s.cont,s.mob,s.email,s.url_web,s.url_web_check,s.reg,s.sit,s.tke,s.f1q,s.f1d,s.f2q,s.f2d,s.f3q,s.f3d,s.web,s.webq,s.ita,s.fhq,s.hora,s.pres,s.presentat,s.resguard,s.notes,s.proc_comercial_previ,s.te_requeriment,s.primer_contacte,s.proposta_preparada,s.proposta_presentada,s.proposta_enviada,s.oferta_acceptada,s.kickoff_esperat,s.holded,s.notes_comercial,s.consultor,s.todo,s.email_consultor,s.email_consultor_cc,s.edicte_od,id]
  );
  d.updated_at=res.rows?.[0]?.updated_at;
}
async function sbDelete(id){
  await neonQuery(`DELETE FROM ${YEAR_CONFIG.table} WHERE id=$1`,[id]);
}

async function sbGetAdmins(){
  const res=await neonQuery(`SELECT name,password FROM admins ORDER BY id ASC`);
  return res.rows;
}
async function sbAddAdmin(name,password){
  await neonQuery(`INSERT INTO admins(name,password) VALUES ($1,$2) ON CONFLICT (name) DO UPDATE SET password=$2`,[name,password]);
}
async function sbRemoveAdmin(name){
  await neonQuery(`DELETE FROM admins WHERE name=$1`,[name]);
}
async function sbSetAdminPassword(name,password){
  await neonQuery(`UPDATE admins SET password=$1 WHERE name=$2`,[password,name]);
}

function rowToSb(d){
  return {g:d.g,emp:d.emp,nova:d.nova,cup:d.cup,cup_ant:d.cup_ant,otorgat:d.otorgat,cont:d.cont,mob:d.mob,email:d.email,
    url_web:d.url_web,url_web_check:d.url_web_check,
    reg:d.reg,sit:d.sit,tke:d.tke,f1q:d.f1q,f1d:d.f1d,f2q:d.f2q,f2d:d.f2d,
    f3q:d.f3q,f3d:d.f3d,web:d.web,webq:d.webq,ita:d.ita,fhq:d.fhq,hora:d.hora,
    pres:d.pres,presentat:d.presentat,resguard:d.resguard,notes:d.notes,
    proc_comercial_previ:d.proc_comercial_previ,te_requeriment:d.te_requeriment,primer_contacte:d.primer_contacte,
    proposta_preparada:d.proposta_preparada,
    proposta_presentada:d.proposta_presentada,proposta_enviada:d.proposta_enviada,
    oferta_acceptada:d.oferta_acceptada,kickoff_esperat:d.kickoff_esperat,holded:d.holded,
    notes_comercial:d.notes_comercial,consultor:d.consultor,todo:d.todo,email_consultor:d.email_consultor,email_consultor_cc:d.email_consultor_cc,edicte_od:d.edicte_od};
}
function sbToRow(r){
  return {id:r.id,g:r.g||'',emp:r.emp||'',nova:!!r.nova,cup:r.cup||'',cup_ant:r.cup_ant||'',otorgat:r.otorgat||'',
    cont:r.cont||'',mob:r.mob||'',email:r.email||'',
    url_web:r.url_web||'',url_web_check:r.url_web_check||'',reg:r.reg||'',
    sit:r.sit||'POTENCIAL',tke:r.tke||'',f1q:r.f1q||'',f1d:r.f1d||'',
    f2q:r.f2q||'',f2d:r.f2d||'',f3q:r.f3q||'',f3d:r.f3d||'',web:r.web||'',webq:r.webq||'',
    ita:r.ita||'',fhq:r.fhq||'',hora:r.hora||'',pres:r.pres||'',
    presentat:!!r.presentat,resguard:!!r.resguard,notes:r.notes||'',updated_at:r.updated_at||null,
    proc_comercial_previ:!!r.proc_comercial_previ,te_requeriment:!!r.te_requeriment,primer_contacte:r.primer_contacte||'',
    proposta_preparada:!!r.proposta_preparada,
    proposta_presentada:r.proposta_presentada||'',proposta_enviada:r.proposta_enviada||'',
    oferta_acceptada:!!r.oferta_acceptada,kickoff_esperat:r.kickoff_esperat||'',holded:!!r.holded,
    notes_comercial:r.notes_comercial||'',consultor:r.consultor||'',todo:!!r.todo,email_consultor:r.email_consultor||'',email_consultor_cc:r.email_consultor_cc||'',edicte_od:r.edicte_od||''};
}

// Indicador de guardat
function showSaving(){
  let el=document.getElementById('sb-status');
  if(!el){el=document.createElement('div');el.id='sb-status';
    el.style='position:fixed;bottom:10px;right:10px;background:#2c5aa0;color:#fff;padding:5px 12px;border-radius:4px;font-size:9pt;z-index:9999';
    document.body.appendChild(el);}
  el.textContent='💾 Guardant...';el.style.display='block';
}
function showSaved(){
  const el=document.getElementById('sb-status');
  if(el){el.textContent='✓ Guardat';setTimeout(()=>el.style.display='none',1500);}
}
function showError(msg){
  const el=document.getElementById('sb-status');
  if(el){el.style.background='#c00';el.textContent='⚠ Error: '+msg;setTimeout(()=>el.style.display='none',3000);}
}

/* ═══════════════════════════════════════════════
   14. INICIALITZACIÓ — Carrega des de Neon
═══════════════════════════════════════════════ */
async function init(){
  try{
    const rows=await sbGet();
    D=rows.map(sbToRow);
    nid=D.length>0?Math.max(...D.map(x=>x.id))+1:1;
    calcTickets();render();
  }catch(e){
    console.error('Error carregant dades:',e);
    render(); // mostra buit si falla
  }
  try{
    REQ=await getRequeriments();
    renderComercial();
  }catch(e){
    console.error('Error carregant requeriments:',e);
  }
  try{
    const admins=await sbGetAdmins();
    A=['Admin',...admins.map(x=>x.name)];
    ADMIN_PWS={};admins.forEach(x=>ADMIN_PWS[x.name]=x.password);
    rebuildAdminBtns();
  }catch(e){
    console.error('Error carregant llista d\'admins:',e);
  }
  try{
    const[pw,lists,pinMap]=await Promise.all([getMasterPw(),getAllLists(),getPins()]);
    if(pw)ADMIN_PW_MASTER=pw;
    if(lists.g)G=lists.g; if(lists.pf)PF=lists.pf; if(lists.td)TD=lists.td; if(lists.qw)QW=lists.qw;
    if(lists.c)CUPS=lists.c; if(lists.s)SITS=lists.s; if(lists.web)WEBLIST=lists.web;
    if(lists.tf3)TF3=lists.tf3; if(lists.tfh)TFH=lists.tfh;
    if(lists.otorgat)OTORGAT_OPTS=lists.otorgat;
    if(lists.reqestat)REQ_ESTAT_OPTS=lists.reqestat;
    if(lists.resolucio)RESOLUCIO_OPTS=lists.resolucio;
    if(lists.consultors)CONSULTORS=lists.consultors;
    if(lists.cons)CONS=lists.cons;
    pins=pinMap;
    rebuildRoleBtns();render();
  }catch(e){
    console.error('Error carregant configuració (contrasenya/llistes/PINs):',e);
  }
}

// Sobreescriure addRow per guardar a Neon
const _addRow=addRow;
async function addRow(){
  if(!adminLevel){togglePanel('admin');return;} // demanar password si no autenticat
  const nou={id:nid++,g:'',emp:'',nova:false,cup:'',cup_ant:'',otorgat:'',cont:'',mob:'',email:'',url_web:'',url_web_check:'',
    reg:'',sit:'POTENCIAL',tke:'',f1q:'',f1d:'',f2q:'',f2d:'',f3q:'',f3d:'',
    web:'',webq:'',ita:'',fhq:'',hora:'',pres:'',presentat:false,resguard:false,notes:'',
    proc_comercial_previ:false,te_requeriment:false,primer_contacte:'',proposta_preparada:false,proposta_presentada:'',
    proposta_enviada:'',oferta_acceptada:false,kickoff_esperat:'',holded:false,notes_comercial:'',consultor:'',todo:false,email_consultor:'',email_consultor_cc:'',edicte_od:''};
  showSaving();
  try{
    const saved=await sbInsert(nou);
    if(saved){nou.id=saved.id;nou.updated_at=saved.updated_at;}
    D.push(nou);nid=nou.id+1;
    calcTickets();render();showSaved();
  }catch(e){showError(e.message);}
}

// Sobreescriure delRow per esborrar a Neon
async function delRow(id){
  if(!confirm('Eliminar aquest registre?'))return;
  showSaving();
  try{
    await sbDelete(id);
    const i=D.findIndex(x=>x.id===id);
    if(i>-1)D.splice(i,1);
    calcTickets();render();showSaved();
  }catch(e){showError(e.message);}
}

// Sobreescriure sv per guardar a Neon
const _sv=sv;
async function sv(el){
  const id=+el.dataset.id,f=el.dataset.f;
  const d=D.find(x=>x.id===id);
  if(!d)return;
  let v=el.textContent.trim();
  if(f==='emp'){
    v=v.toUpperCase();
  } else if(f==='reg'){
    v=addTimestamp(fmtDate(v));
    if(v&&(d.sit===''||d.sit==='POTENCIAL'))d.sit='PROCÉS';
    if(!v&&d.sit==='PROCÉS')d.sit='POTENCIAL';
  } else if(['f1d','f2d','f3d'].includes(f)){
    v=fmtDate(v);
  } else if(f==='hora'){
    v=fmtHora(v);
  } else if(['primer_contacte','proposta_presentada','proposta_enviada','kickoff_esperat'].includes(f)){
    v=fmtDate(v).replace(/\./g,'/');
  }
  d[f]=v;checkAutoComplete(d);calcTickets();
  showSaving();
  try{await sbUpdate(d.id,d);showSaved();}catch(e){if(e.message!=='CANCELLED_BY_USER')showError(e.message);}
  render();
}

// Sobreescriure svs per guardar a Neon
const _svs=svs;
async function svs(id,f,v){
  const d=D.find(x=>x.id===id);
  if(!d)return;
  if(f==='cup'&&v&&v===d.cup_ant){alert('El Cupó no pot coincidir amb el Cupó anterior.');render();return;}
  d[f]=v;
  if(f==='cup'){if(v==='IA'&&!d.ita)d.ita='DEMANAR';if(v!=='IA')d.ita='';}
  checkAutoComplete(d);calcTickets();
  showSaving();
  try{await sbUpdate(d.id,d);showSaved();}catch(e){if(e.message!=='CANCELLED_BY_USER')showError(e.message);}
  render();
}

// Sobreescriure svSit per guardar a Neon
const _svSit=svSit;
async function svSit(id,val){
  document.querySelectorAll('.sit-drop.open').forEach(x=>x.classList.remove('open'));
  const d=D.find(x=>x.id===id);
  if(!d)return;
  if(['COMPLETAT','PROCÉS'].includes(val)&&!d.reg){alert('Cal indicar la data d\'Entrada abans de poder marcar '+val+'.');return;}
  d.sit=val;calcTickets();
  showSaving();
  try{await sbUpdate(d.id,d);showSaved();}catch(e){if(e.message!=='CANCELLED_BY_USER')showError(e.message);}
  render();
}

// Sobreescriure svIta per guardar a Neon
async function svIta(id,val){
  document.querySelectorAll('.ita-drop.open').forEach(x=>x.classList.remove('open'));
  const d=D.find(x=>x.id===id);
  if(!d)return;
  d.ita=val;
  showSaving();
  try{await sbUpdate(d.id,d);showSaved();}catch(e){if(e.message!=='CANCELLED_BY_USER')showError(e.message);}
  render();
}

init();

// ── POLLING: actualitza dades cada 2 segons ──
setInterval(async()=>{
  // No actualitzar si l'usuari està editant activament o té un desplegable obert
  if(document.activeElement&&document.activeElement.classList.contains('ed'))return;
  if(document.querySelector('.sit-drop.open,.ita-drop.open'))return;
  try{
    const rows=await sbGet();
    const nouD=rows.map(sbToRow);
    if(JSON.stringify(nouD)!==JSON.stringify(D)){
      D=nouD;
      nid=D.length>0?Math.max(...D.map(x=>x.id))+1:1;
      calcTickets();render();
    }
  }catch(e){console.warn('Polling error:',e);}
  try{
    const nouReq=await getRequeriments();
    if(JSON.stringify(nouReq)!==JSON.stringify(REQ)){
      REQ=nouReq;
      renderComercial();
      if(document.getElementById('tab-requeriments').classList.contains('active'))renderRequerimentsTable();
    }
  }catch(e){console.warn('Polling requeriments error:',e);}
  try{
    const admins=await sbGetAdmins();
    const nouA=['Admin',...admins.map(x=>x.name)];
    const nouPws={};admins.forEach(x=>nouPws[x.name]=x.password);
    // Revocació en calent: si el Master ha eliminat aquest admin o li ha canviat la contrasenya, es tanca la sessió
    if(cT==='a'&&cU!=='Admin'&&adminLevel&&(!nouA.includes(cU)||nouPws[cU]!==myAdminPw)){
      adminLevel='';myAdminPw='';authedAs='';cU='Admin';cT='a';
      A=nouA;ADMIN_PWS=nouPws;
      rebuildAdminBtns();render();
      alert('El Master t\'ha revocat l\'accés d\'Admin.');
    }else if(JSON.stringify(nouA)!==JSON.stringify(A)||JSON.stringify(nouPws)!==JSON.stringify(ADMIN_PWS)){
      A=nouA;ADMIN_PWS=nouPws;
      rebuildAdminBtns();
      if(document.getElementById('panel-admin').classList.contains('open'))renderAdminList();
    }
  }catch(e){console.warn('Polling admins error:',e);}
  try{
    const[pw,lists,pinMap]=await Promise.all([getMasterPw(),getAllLists(),getPins()]);
    let changed=false;
    if(pw&&pw!==ADMIN_PW_MASTER){ADMIN_PW_MASTER=pw;}
    if(JSON.stringify(pinMap)!==JSON.stringify(pins)){pins=pinMap;changed=true;}
    const curLists={g:G,pf:PF,td:TD,qw:QW,c:CUPS,s:SITS,web:WEBLIST,tf3:TF3,tfh:TFH,otorgat:OTORGAT_OPTS,reqestat:REQ_ESTAT_OPTS,resolucio:RESOLUCIO_OPTS,consultors:CONSULTORS,cons:CONS};
    Object.keys(curLists).forEach(k=>{
      if(lists[k]&&JSON.stringify(lists[k])!==JSON.stringify(curLists[k]))changed=true;
    });
    if(lists.g)G=lists.g; if(lists.pf)PF=lists.pf; if(lists.td)TD=lists.td; if(lists.qw)QW=lists.qw;
    if(lists.c)CUPS=lists.c; if(lists.s)SITS=lists.s; if(lists.web)WEBLIST=lists.web;
    if(lists.tf3)TF3=lists.tf3; if(lists.tfh)TFH=lists.tfh;
    if(lists.otorgat)OTORGAT_OPTS=lists.otorgat;
    if(lists.reqestat)REQ_ESTAT_OPTS=lists.reqestat;
    if(lists.resolucio)RESOLUCIO_OPTS=lists.resolucio;
    if(lists.consultors)CONSULTORS=lists.consultors;
    if(lists.cons)CONS=lists.cons;
    if(changed){render();renderRequerimentsTable();renderComercial();}
  }catch(e){console.warn('Polling config error:',e);}
},2000);
