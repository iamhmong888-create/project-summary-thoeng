'use strict';
/* =====================================================================
   ระบบสรุปผลการดำเนินโครงการ — โค้ดร่วมของหน้าผู้เรียน (index.html)
   และหน้าครู/แอดมิน (admin.html)   ·   MODE กำหนดในไฟล์ HTML
   ===================================================================== */
const PAGE_MODE = window.APP_MODE || 'learner';
let MODE = PAGE_MODE;
const ADMIN_VIEWS = ['login','setup','home','new','edit','p','r','test','org','staff','rooms'];
const FIREBASE_CONFIG = window.FIREBASE_CONFIG || null;
const DOCX_URL = 'https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js';
const XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
const FB_VER = '10.12.2';
const STAFF_DOMAIN = '@staff.example.com';
const GENDERS = ['ชาย','หญิง'];
const AGES = ['ต่ำกว่า 15 ปี','16 - 39 ปี','40 - 59 ปี','60 ขึ้นไป'];
const EDU_LEVELS = ['ประถมศึกษา','มัธยมศึกษาตอนต้น','มัธยมศึกษาตอนปลาย'];
const PREFIXES = ['นาย','นาง','นางสาว'];
const PREFIX_GENDER = {'นาย':'ชาย','นาง':'หญิง','นางสาว':'หญิง'};
const DEFAULT_ADMIN_NAMES = ['วนศาสตร์'];
const DEFAULT_CLASSES = [
  ['ศกร.ระดับตำบลตับเต่า','ประสิทธิ วิชัย'],['ศกร.ระดับตำบลหงาว','เด่นนภา นิทะโน'],['ศกร.ระดับตำบลเวียง','รัชฎาภรณ์ สุริยนต์'],
  ['ศกร.ระดับตำบลสันทรายงาม','วีนัส สรรพสุข'],['ศกร.ระดับตำบลปล้อง','พรทิพย์ สุขบัญชาชัย'],['ศกร.ระดับตำบลงิ้ว','กัญญ์วรา สีหมนตรี'],
  ['ศกร.ระดับตำบลหนองแรด','วรินท์นิภา จันทวิมล'],['ศกร.ระดับตำบลแม่ลอย','ดวงกมล เข็มเพชร'],['ศกร.ระดับตำบลศรีดอนไชย','ศุภชัย แก้วกระจ่าง'],
  ['ศกร.ระดับตำบลเชียงเคี่ยน','ธันยพัต ปกครองบ้าน'],['ศกร.ระดับตำบลเวียง','อดิเรก ประยศ'],['ศกร.ระดับตำบลงิ้ว','ทับทิม พุทธตาล'],
  ['ศกร.ระดับตำบลเชียงเคี่ยน','ปราณี ใจกว้าง'],['ศศช.บ้านขุนต้า','สรนันท์ แสนยศ'],['ศศช.บ้านราฎภักดี','สายสุนีย์ โกฏฐา'],
  ['ศศช.บ้านขุนห้วยไคร้','สุนิสา จินดาจักร'],['ศศช.บ้านเล่าตาขาว','นรินทร์ลักษณ์ นักหล่อ'],['ศศช.บ้านราฎรักษา','วิมล สุทธิจันทร์'],
  ['ศศช.บ้านไทยสามัคคี','อนุพงษ์ สมจริงจงรักษ์'],['ศศช.บ้านราฎภักดี','วนศาสตร์ วัฒนาสกุลลี'],['ศรช.ระดับตำบลหงาว','ธวัชชัย เพ็ชรแสง'],
  ['ศรช.ระดับตำบลเวียง','ชินดนัย คำต๊ะ']].map(([center,teacher])=>({center,teacher}));
const roomLabel = c => c.teacher ? `${c.center} (ครู${c.teacher})` : c.center;
const OCCS = ['เกษตรกร','รับจ้าง','ค้าขาย/ธุรกิจส่วนตัว','ข้าราชการ/พนักงานรัฐ','นักเรียน/นักศึกษา','แม่บ้าน/ไม่ได้ประกอบอาชีพ','อื่น ๆ'];
const LEVELS = [[4.51,'มากที่สุด'],[3.51,'มาก'],[2.51,'ปานกลาง'],[1.51,'น้อย'],[0,'น้อยที่สุด']];
const SCALE_LABELS = {5:'มากที่สุด',4:'มาก',3:'ปานกลาง',2:'น้อย',1:'น้อยที่สุด'};
const OPT_LABELS = ['ก','ข','ค','ง','จ','ฉ'];
const DEFAULT_ITEMS = `# ด้านเนื้อหา
เนื้อหาตรงตามความต้องการของผู้เข้าร่วม
เนื้อหาเพียงพอและเหมาะสมกับเวลา
เนื้อหาเป็นประโยชน์ต่อการพัฒนาคุณภาพชีวิต
# ด้านกระบวนการจัดกิจกรรม
การเตรียมความพร้อมก่อนจัดกิจกรรม
กิจกรรมเหมาะสมกับวัตถุประสงค์ของโครงการ
ระยะเวลาในการจัดกิจกรรมมีความเหมาะสม
# ด้านวิทยากร
วิทยากรมีความรู้ความสามารถในเรื่องที่ถ่ายทอด
วิทยากรถ่ายทอดเข้าใจง่ายและใช้สื่อเหมาะสม
วิทยากรเปิดโอกาสให้ซักถามและมีส่วนร่วม
# ด้านสถานที่และสิ่งอำนวยความสะดวก
สถานที่ วัสดุ และอุปกรณ์มีความเหมาะสม
# ด้านการนำความรู้ไปใช้
สามารถนำความรู้ไปใช้ในชีวิตประจำวันได้
สามารถนำความรู้ไปต่อยอดอาชีพหรือสร้างรายได้ได้`;
const TEST_SAMPLE = `1. ข้อใดคือคำถามตัวอย่าง
ก. ตัวเลือกที่ 1
*ข. ตัวเลือกที่ถูก (ใส่ * หน้าข้อที่ถูก)
ค. ตัวเลือกที่ 3
ง. ตัวเลือกที่ 4`;
const DEFAULT_SETTINGS = {
  org:'ศูนย์ส่งเสริมการเรียนรู้ระดับอำเภอเทิง',
  parent:'สำนักงานส่งเสริมการเรียนรู้ประจำจังหวัดเชียงราย',
  director:'', directorPos:'ผู้อำนวยการศูนย์ส่งเสริมการเรียนรู้ระดับอำเภอเทิง'
};

/* ------------------------------------------------------------------ */
/* เครื่องมือทั่วไป                                                    */
/* ------------------------------------------------------------------ */
const $ = (s, el=document) => el.querySelector(s);
const $$ = (s, el=document) => [...el.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const uid = () => Math.random().toString(36).slice(2,10) + Date.now().toString(36);
const num = v => { const n = parseFloat(String(v ?? '').replace(/,/g,'')); return isNaN(n) ? 0 : n; };
const f2 = n => (Math.round(n*100)/100).toFixed(2);
const money = n => num(n).toLocaleString('th-TH',{minimumFractionDigits:2,maximumFractionDigits:2});
const pct = (a,b) => b ? (a/b*100) : 0;
const lines = s => String(s||'').split(/\n+/).map(x=>x.trim()).filter(Boolean);
const level = m => (LEVELS.find(([t]) => m >= t) || LEVELS[4])[1];
const app = $('#app');

const PREFIX_RE = /^(นางสาว|นาง|นาย|น\.ส\.|ด\.ช\.|ด\.ญ\.|เด็กชาย|เด็กหญิง)\s*(?=[ก-ฮ])/;
function normName(s){ return String(s||'').replace(/\s+/g,' ').trim().replace(PREFIX_RE,'').trim(); }
const nameKey = s => normName(s).replace(/\s+/g,'').toLowerCase();
function normPhone(s){ let d=String(s||'').replace(/\D/g,''); if(d.length===9 && d[0]!=='0') d='0'+d; return d; }
async function sha256(s){ const b=await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s)); return [...new Uint8Array(b)].map(x=>x.toString(16).padStart(2,'0')).join(''); }
async function staffEmail(name){ return 't'+(await sha256('psum-staff:'+nameKey(name))).slice(0,24)+STAFF_DOMAIN; }

function thDate(d){ if(!d) return ''; const x = new Date(d+'T00:00:00'); return isNaN(x) ? d : x.toLocaleDateString('th-TH',{day:'numeric',month:'long',year:'numeric'}); }
function dateRange(p){
  if(!p.startDate) return '';
  if(!p.endDate || p.endDate===p.startDate) return 'วันที่ ' + thDate(p.startDate);
  return 'วันที่ ' + thDate(p.startDate) + ' ถึงวันที่ ' + thDate(p.endDate);
}
function currentFY(){ const d=new Date(); return d.getFullYear()+543+(d.getMonth()>=9?1:0); }
function monthYearTH(){ return new Date().toLocaleDateString('th-TH',{month:'long',year:'numeric'}); }

function bahtText(n){
  n = Math.round(num(n)*100)/100;
  const [ip, sp] = n.toFixed(2).split('.');
  const D = ['','หนึ่ง','สอง','สาม','สี่','ห้า','หก','เจ็ด','แปด','เก้า'];
  const P = ['','สิบ','ร้อย','พัน','หมื่น','แสน'];
  function grp(s, eet){
    let out=''; const L=s.length;
    for(let i=0;i<L;i++){
      const d=+s[i], pos=L-1-i; if(!d) continue;
      if(pos===1 && d===1) out+='สิบ';
      else if(pos===1 && d===2) out+='ยี่สิบ';
      else if(pos===0 && d===1 && eet) out+='เอ็ด';
      else out+=D[d]+P[pos];
    }
    return out;
  }
  function read(s){
    s = s.replace(/^0+/,''); if(!s) return '';
    if(s.length>6) return read(s.slice(0,s.length-6))+'ล้าน'+grp(s.slice(-6).replace(/^0+/,''), true);
    return grp(s, s.length>1);
  }
  return (ip==='0' ? 'ศูนย์' : read(ip)) + 'บาท' + (sp==='00' ? 'ถ้วน' : read(sp)+'สตางค์');
}

function parseItems(text){
  return lines(text).map(l => l.startsWith('#') ? {t:'h',text:l.replace(/^#+\s*/,'')} : {t:'q',text:l.replace(/^[-•\d.)\s]+/,'')||l});
}
function itemsToText(items){ return (items||[]).map(i => (i.t==='h'?'# ':'') + i.text).join('\n'); }

/* แบบทดสอบ: บรรทัดคำถาม ตามด้วยตัวเลือก ก. ข. ค. ง. — ใส่ * หน้าข้อที่ถูก */
const OPT_RE = /^(\*?)\s*([กขคงจฉ]|[a-fA-F])\s*[.)]\s*(.*)$/;
function parseTest(text){
  const qs=[], errors=[]; let cur=null;
  for(const raw of String(text||'').split('\n')){
    const l=raw.trim(); if(!l) continue;
    const m=l.match(OPT_RE);
    if(m && cur){ if(m[1]) cur.ans=cur.opts.length; cur.opts.push(m[3].replace(/^\*\s*/,'').trim()); continue; }
    if(m && !cur){ errors.push(`ตัวเลือก "${l}" ไม่มีคำถามนำหน้า`); continue; }
    cur={q:l.replace(/^\d+\s*[.)]\s*/,''),opts:[],ans:-1}; qs.push(cur);
  }
  qs.forEach((q,i)=>{
    if(q.opts.length<2) errors.push(`ข้อ ${i+1}: ต้องมีตัวเลือกอย่างน้อย 2 ตัว`);
    if(q.ans<0) errors.push(`ข้อ ${i+1}: ยังไม่ได้ใส่ * หน้าตัวเลือกที่ถูก`);
  });
  return {questions:qs, errors};
}
/* การสลับลำดับข้อ: คำตอบยังบันทึกตามลำดับข้อเดิมเสมอ จึงคิดคะแนนและเปรียบเทียบได้ตามปกติ */
function seededRng(str){ let h=1779033703^str.length; for(let i=0;i<str.length;i++){ h=Math.imul(h^str.charCodeAt(i),3432918353); h=h<<13|h>>>19; }
  let s=h>>>0; return ()=>{ s|=0; s=s+0x6D2B79F5|0; let t=Math.imul(s^s>>>15,1|s); t=t+Math.imul(t^t>>>7,61|t)^t; return ((t^t>>>14)>>>0)/4294967296; }; }
function shuffledIdx(n, rnd){ const a=[...Array(n).keys()]; for(let i=n-1;i>0;i--){ const j=Math.floor(rnd()*(i+1)); [a[i],a[j]]=[a[j],a[i]]; } return a; }
const SHUFFLE_MODES = {none:'ไม่สลับ (ก่อนเรียนและหลังเรียนลำดับเดียวกัน)', post:'สลับข้อในแบบทดสอบหลังเรียน (ทุกคนได้ลำดับเดียวกัน ใช้พิมพ์ลงกระดาษได้)', random:'สุ่มลำดับข้อทุกครั้ง (ผู้เรียนแต่ละคนได้ลำดับต่างกัน)'};
function testLayout(id, kind, test, staff){
  const Q=test.questions||[]; const mode=test.shuffle||'none'; let rnd=null;
  if(mode==='post' && kind==='post') rnd=seededRng(id+':post');
  else if(mode==='random' && !staff) rnd=Math.random;
  const qOrder = rnd ? shuffledIdx(Q.length,rnd) : [...Q.keys()];
  const optOrders = Q.map(q=> rnd && test.shuffleOpt ? shuffledIdx(q.opts.length,rnd) : [...q.opts.keys()]);
  return {qOrder, optOrders};
}
function testToText(qs){ return (qs||[]).map((q,i)=>`${i+1}. ${q.q}\n`+q.opts.map((o,j)=>`${j===q.ans?'*':''}${OPT_LABELS[j]||j+1}. ${o}`).join('\n')).join('\n\n'); }

function toast(msg){ const t=$('#toast'); if(!t) return; t.textContent=msg; t.classList.add('on'); clearTimeout(toast._t); toast._t=setTimeout(()=>t.classList.remove('on'),2800); }
function ask(msg, ok='ยืนยัน', danger=true){
  return new Promise(res=>{
    const m=document.createElement('div'); m.className='modal';
    m.innerHTML=`<div class="box"><p style="margin-top:0">${esc(msg)}</p><div class="row end"><button class="btn ghost" data-v="0">ยกเลิก</button><button class="btn ${danger?'danger':''}" data-v="1">${esc(ok)}</button></div></div>`;
    m.addEventListener('click',e=>{ const v=e.target.dataset.v; if(v!=null){ m.remove(); res(v==='1'); } });
    document.body.appendChild(m);
  });
}
function loadScript(src){ return new Promise((res,rej)=>{ const s=document.createElement('script'); s.src=src; s.onload=res; s.onerror=()=>rej(new Error('โหลดไม่สำเร็จ: '+src)); document.head.appendChild(s); }); }
async function download(blob, name){
  try{
    if(window.claude && typeof window.claude.use==='function'){
      const dl=await window.claude.use('downloads');
      if(dl){
        try{ await dl.save({filename:name,data:blob}); toast('บันทึกไฟล์แล้ว'); }
        catch(e){ if(e?.code!=='declined') toast('ดาวน์โหลดไม่สำเร็จ ('+(e?.code||e?.message||e)+')'); }
        return;
      }
    }
  }catch(e){}
  try{
    const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(()=>URL.revokeObjectURL(a.href),8000);
    toast('กำลังดาวน์โหลด '+name);
  }catch(e){ toast('ดาวน์โหลดไม่ได้ในหน้านี้'); }
}
async function copyText(t){
  try{ await navigator.clipboard.writeText(t); toast('คัดลอกแล้ว'); }
  catch(e){ toast('คัดลอกไม่ได้ — กรุณาเลือกข้อความแล้วคัดลอกเอง'); }
}
function compressImage(file, maxSide){
  return new Promise((res,rej)=>{
    const url=URL.createObjectURL(file), img=new Image();
    img.onload=()=>{
      let w=img.naturalWidth, h=img.naturalHeight; const r=Math.min(1, maxSide/Math.max(w,h));
      w=Math.round(w*r); h=Math.round(h*r);
      const c=document.createElement('canvas'); c.width=w; c.height=h;
      const g=c.getContext('2d'); g.fillStyle='#fff'; g.fillRect(0,0,w,h); g.drawImage(img,0,0,w,h);
      let q=0.75, d=c.toDataURL('image/jpeg',q);
      while(d.length>650000 && q>0.35){ q-=0.1; d=c.toDataURL('image/jpeg',q); }
      URL.revokeObjectURL(url); res({data:d,w,h});
    };
    img.onerror=()=>{ URL.revokeObjectURL(url); rej(new Error('ไฟล์ภาพไม่ถูกต้อง')); };
    img.src=url;
  });
}
function learnerUrl(id, kind='j'){
  const u=new URL('index.html', location.href); u.hash=kind+'-'+id; return u.href;
}

/* อ่านรายชื่อบุคลากรจาก Excel: ต้องมีหัวตาราง ชื่อ และ เบอร์โทร (มี คำนำหน้า สกุล ตำแหน่ง ได้) */
async function readStaffExcel(file){
  if(!window.XLSX) await loadScript(XLSX_URL);
  const wb=XLSX.read(await file.arrayBuffer(),{type:'array'});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:false,defval:''}).map(r=>r.map(c=>String(c??'').trim()));
  const hi=rows.findIndex(r=>r.some(c=>c==='ชื่อ'||c==='ชื่อ-สกุล') && r.some(c=>/เบอร์|โทร/.test(c)));
  if(hi<0) throw new Error('ไม่พบหัวตาราง "ชื่อ" และ "เบอร์โทร" ในไฟล์');
  const H=rows[hi]; const col=f=>H.findIndex(f);
  const cName=col(c=>c==='ชื่อ'||c==='ชื่อ-สกุล'), cPre=col(c=>c.includes('คำนำหน้า')),
        cSur=col(c=>c==='สกุล'||c==='นามสกุล'), cPos=col(c=>c.includes('ตำแหน่ง')), cTel=col(c=>/เบอร์|โทร/.test(c));
  const out=[];
  for(const r of rows.slice(hi+1)){
    let name=r[cName]||''; if(!name.trim()) continue;
    let surname=cSur>=0?r[cSur]:'';
    if(H[cName]==='ชื่อ-สกุล' && !surname){ const p=normName(name).split(' '); name=p[0]; surname=p.slice(1).join(' '); }
    const position=cPos>=0?r[cPos]:'';
    out.push({ prefix:cPre>=0?r[cPre]:'', name:normName(name), surname:surname.trim(), position,
      phone:normPhone(r[cTel]), role:(/ผู้บริหาร|ผู้อำนวยการ|แอดมิน|admin/i.test(position)||DEFAULT_ADMIN_NAMES.includes(normName(name)))?'admin':'teacher' });
  }
  return out;
}
function checkStaffRows(rows){
  const errs=[]; const seen={};
  rows.forEach((r,i)=>{
    if(!r.name) errs.push(`แถว ${i+1}: ไม่มีชื่อ`);
    if(r.phone.length<6) errs.push(`${r.name||'แถว '+(i+1)}: เบอร์โทรไม่ถูกต้อง`);
    const k=nameKey(r.name); if(seen[k]) errs.push(`ชื่อ "${r.name}" ซ้ำกัน — ระบบล็อกอินด้วยชื่ออย่างเดียว จึงต้องไม่ซ้ำ`); seen[k]=1;
  });
  if(!rows.some(r=>r.role==='admin')) errs.push('ต้องมีผู้ดูแลระบบ (แอดมิน) อย่างน้อย 1 คน');
  return errs;
}

/* ------------------------------------------------------------------ */
/* ข้อมูลสาธารณะที่ผู้เรียนเห็น (ไม่มีข้อมูลส่วนตัว)                    */
/* ------------------------------------------------------------------ */
function publicDoc(p, settings){
  const t=p.test||{}; const qs=t.questions||[];
  const sat={open:p.surveyOpen!==false, items:p.items||[]};
  const test={ questions:qs.map(q=>({q:q.q,opts:q.opts})), key:t.showScore!==false?qs.map(q=>q.ans):null,
    preOpen:!!t.preOpen&&qs.length>0, postOpen:!!t.postOpen&&qs.length>0, passPct:num(t.passPct)||60,
    shuffle:t.shuffle||'none', shuffleOpt:!!t.shuffleOpt };
  return { title:p.name||'', org:settings?.org||'', dateText:dateRange(p), location:p.location||'',
    responsible:p.responsible||'', active:sat.open||test.preOpen||test.postOpen, sat, test, updatedAt:Date.now() };
}

/* ------------------------------------------------------------------ */
/* ชั้นข้อมูล: โหมดทดลอง (localStorage)                                */
/* ------------------------------------------------------------------ */
class LocalStore{
  constructor(){ this.key='psum-demo-v3'; this.mem=null; }
  async init(){
    try{ this.mem=JSON.parse(localStorage.getItem(this.key)||'null'); }catch(e){ this.mem=null; }
    if(!this.mem){ this.mem={projects:{},responses:{},tests:{},photos:{},staff:{},settings:{...DEFAULT_SETTINGS}}; seedDemo(this.mem); this.persist(); }
    this.mem.tests ||= {}; this.mem.staff ||= {};
  }
  persist(){
    try{ localStorage.setItem(this.key, JSON.stringify(this.mem)); }
    catch(e){ if(String(e.name).includes('Quota')) toast('พื้นที่เก็บข้อมูลในเบราว์เซอร์เต็ม — ลองลบภาพบางส่วน'); }
  }
  reset(){ try{ localStorage.removeItem(this.key); sessionStorage.removeItem('psum-session'); }catch(e){} }
  // บัญชีผู้ใช้
  hasStaff(){ return Object.keys(this.mem.staff).length>0; }
  async currentUser(){ let e=null; try{ e=sessionStorage.getItem('psum-session'); }catch(x){} return e&&this.mem.staff[e] ? {email:e,...this.mem.staff[e]} : null; }
  async login(name, phone){
    const e=await staffEmail(name); const m=this.mem.staff[e];
    if(!m || m.ph!==await sha256(normPhone(phone))) throw new Error('ชื่อหรือเบอร์โทรไม่ถูกต้อง');
    try{ sessionStorage.setItem('psum-session',e); }catch(x){}
    return {email:e,...m};
  }
  async logout(){ try{ sessionStorage.removeItem('psum-session'); }catch(x){} }
  async listStaff(){ return Object.entries(this.mem.staff).map(([email,m])=>({email,...m,ph:undefined})); }
  async addStaff(rows){
    const rep=[];
    for(const r of rows){ const e=await staffEmail(r.name);
      this.mem.staff[e]={prefix:r.prefix||'',name:r.name,surname:r.surname||'',position:r.position||'',role:r.role,ph:await sha256(normPhone(r.phone))};
      rep.push({name:r.name,status:'สร้างแล้ว'}); }
    this.persist(); return rep;
  }
  async setupStaff(rows){ return this.addStaff(rows); }
  async updateStaff(email, data){ Object.assign(this.mem.staff[email], data); this.persist(); }
  async removeStaff(email){ delete this.mem.staff[email]; this.persist(); }
  // โครงการ
  async listProjects(){ return Object.values(this.mem.projects).sort((a,b)=>(b.updatedAt||0)-(a.updatedAt||0)); }
  async getProject(id){ return this.mem.projects[id] ? JSON.parse(JSON.stringify(this.mem.projects[id])) : null; }
  async saveProject(p){ const id=p.id||uid(); this.mem.projects[id]={...p,id,updatedAt:Date.now()}; this.persist(); return id; }
  async deleteProject(id){ for(const k of ['projects','responses','tests','photos']) delete this.mem[k][id]; this.persist(); }
  async getPublic(id){ const p=this.mem.projects[id]; return p ? publicDoc(p, this.mem.settings) : null; }
  async listPublic(){ return Object.values(this.mem.projects).map(p=>({id:p.id,...publicDoc(p,this.mem.settings)})).filter(d=>d.active).sort((a,b)=>b.updatedAt-a.updatedAt); }
  async addResponse(pid,r){ (this.mem.responses[pid] ||= {})[uid()]=r; this.persist(); }
  async listResponses(pid){ return Object.entries(this.mem.responses[pid]||{}).map(([id,r])=>({id,...r})).sort((a,b)=>a.at-b.at); }
  async deleteResponse(pid,rid){ delete (this.mem.responses[pid]||{})[rid]; this.persist(); }
  async addTest(pid,r){ (this.mem.tests[pid] ||= {})[uid()]=r; this.persist(); }
  async listTests(pid){ return Object.entries(this.mem.tests[pid]||{}).map(([id,r])=>({id,...r})).sort((a,b)=>a.at-b.at); }
  async deleteTest(pid,rid){ delete (this.mem.tests[pid]||{})[rid]; this.persist(); }
  async listPhotos(pid){ return Object.entries(this.mem.photos[pid]||{}).map(([id,r])=>({id,...r})).sort((a,b)=>a.at-b.at); }
  async addPhoto(pid,ph){ (this.mem.photos[pid] ||= {})[uid()]={...ph,at:Date.now()}; this.persist(); }
  async updatePhoto(pid,id,data){ Object.assign(this.mem.photos[pid][id],data); this.persist(); }
  async deletePhoto(pid,id){ delete this.mem.photos[pid][id]; this.persist(); }
  async getSettings(){ return {...DEFAULT_SETTINGS, ...this.mem.settings}; }
  async getClasses(){ return this.mem.classes || DEFAULT_CLASSES; }
  async saveClasses(list){ this.mem.classes=list; this.persist(); }
  async saveSettings(s){ this.mem.settings=s; this.persist(); }
}

/* ------------------------------------------------------------------ */
/* ชั้นข้อมูล: Firebase (Firestore + Authentication)                   */
/* ------------------------------------------------------------------ */
class FireStore{
  async init(){
    const b=`https://www.gstatic.com/firebasejs/${FB_VER}/`;
    await loadScript(b+'firebase-app-compat.js');
    await loadScript(b+'firebase-auth-compat.js');
    await loadScript(b+'firebase-firestore-compat.js');
    firebase.initializeApp(FIREBASE_CONFIG);
    this.db=firebase.firestore(); this.auth=firebase.auth();
    this.members=null;
  }
  P(){ return this.db.collection('projects'); }
  staffRef(db=this.db){ return db.collection('meta').doc('staff'); }
  sec(){ return firebase.apps.find(a=>a.name==='sec') || firebase.initializeApp(FIREBASE_CONFIG,'sec'); }
  waitAuth(){ return new Promise(res=>{ const un=this.auth.onAuthStateChanged(u=>{ un(); res(u); }); }); }
  hasStaff(){ return true; }
  async currentUser(){
    const u=this.auth.currentUser || await this.waitAuth(); if(!u) return null;
    try{ const d=await this.staffRef().get(); this.members=d.exists?d.data().members||{}:{}; }catch(e){ this.members={}; }
    const m=this.members[u.email]; if(!m){ await this.auth.signOut(); return null; }
    return {email:u.email,...m};
  }
  async login(name, phone){
    const email=await staffEmail(name);
    try{ await this.auth.signInWithEmailAndPassword(email, normPhone(phone)); }
    catch(e){ throw new Error(/network/i.test(e.code||'')?'เชื่อมต่ออินเทอร์เน็ตไม่ได้':'ชื่อหรือเบอร์โทรไม่ถูกต้อง'); }
    const u=await this.currentUser(); if(!u) throw new Error('บัญชีนี้ไม่ได้รับสิทธิ์เข้าใช้ระบบ');
    return u;
  }
  async logout(){ await this.auth.signOut(); }
  async listStaff(){ const d=await this.staffRef().get(); const m=d.exists?d.data().members||{}:{}; return Object.entries(m).map(([email,v])=>({email,...v})); }
  async makeAccount(sa, email, phone){
    try{ await sa.createUserWithEmailAndPassword(email, phone); return 'สร้างแล้ว'; }
    catch(e){ if(e.code==='auth/email-already-in-use') return 'มีบัญชีอยู่แล้ว (ใช้รหัสเดิม)'; throw e; }
    finally{ try{ await sa.signOut(); }catch(x){} }
  }
  member(r){ return {prefix:r.prefix||'',name:r.name,surname:r.surname||'',position:r.position||'',role:r.role}; }
  /* ตั้งค่าครั้งแรก: สร้างบัญชีผู้ดูแลก่อน แล้วบันทึกรายชื่อ จากนั้นสร้างบัญชีที่เหลือ */
  async setupStaff(rows, onProgress=()=>{}){
    const s=this.sec(), sa=s.auth(), sdb=s.firestore();
    const members={}; for(const r of rows) members[await staffEmail(r.name)]=this.member(r);
    const admin=rows.find(r=>r.role==='admin'); const ae=await staffEmail(admin.name);
    try{ await sa.createUserWithEmailAndPassword(ae, admin.phone); }
    catch(e){ if(e.code==='auth/email-already-in-use') await sa.signInWithEmailAndPassword(ae, admin.phone); else throw e; }
    try{ await this.staffRef(sdb).set({members},{merge:true}); }
    catch(e){ await sa.signOut(); throw new Error('ระบบตั้งค่าไว้แล้ว ให้ผู้ดูแลระบบเพิ่มบุคลากรจากเมนู "บุคลากร" แทน'); }
    await sa.signOut();
    const rep=[{name:admin.name,status:'สร้างแล้ว (ผู้ดูแลระบบ)'}];
    for(const r of rows){ if(r===admin) continue; onProgress(r.name);
      try{ rep.push({name:r.name,status:await this.makeAccount(sa, await staffEmail(r.name), r.phone)}); }
      catch(e){ rep.push({name:r.name,status:'ผิดพลาด: '+(e.code||e.message)}); } }
    return rep;
  }
  /* แอดมินเพิ่มบุคลากร: บันทึกสิทธิ์ก่อน แล้วสร้างบัญชีด้วย app สำรอง (ไม่ทำให้แอดมินหลุดจากระบบ) */
  async addStaff(rows, onProgress=()=>{}){
    const add={}; for(const r of rows) add[await staffEmail(r.name)]=this.member(r);
    await this.staffRef().set({members:add},{merge:true});
    const sa=this.sec().auth(); const rep=[];
    for(const r of rows){ onProgress(r.name);
      try{ rep.push({name:r.name,status:await this.makeAccount(sa, await staffEmail(r.name), r.phone)}); }
      catch(e){ rep.push({name:r.name,status:'ผิดพลาด: '+(e.code||e.message)}); } }
    return rep;
  }
  async updateStaff(email, data){ const cur=(await this.listStaff()).find(m=>m.email===email)||{}; delete cur.email; await this.staffRef().set({members:{[email]:{...cur,...data}}},{merge:true}); }
  async removeStaff(email){ await this.staffRef().update(new firebase.firestore.FieldPath('members',email), firebase.firestore.FieldValue.delete()); }
  async listProjects(){ const s=await this.P().orderBy('updatedAt','desc').get(); return s.docs.map(d=>({id:d.id,...d.data()})); }
  async getProject(id){ const d=await this.P().doc(id).get(); return d.exists ? {id:d.id,...d.data()} : null; }
  async saveProject(p){
    const {id, ...data}=p; data.updatedAt=Date.now();
    const ref = id ? this.P().doc(id) : this.P().doc();
    await ref.set(data);
    await this.db.collection('surveys').doc(ref.id).set(publicDoc(data, S.settings));
    return ref.id;
  }
  async deleteProject(id){
    for(const sub of ['responses','tests','photos']){
      const s=await this.P().doc(id).collection(sub).get();
      await Promise.all(s.docs.map(d=>d.ref.delete()));
    }
    await this.P().doc(id).delete();
    await this.db.collection('surveys').doc(id).delete();
  }
  async getPublic(id){ const d=await this.db.collection('surveys').doc(id).get(); return d.exists ? d.data() : null; }
  async listPublic(){ const s=await this.db.collection('surveys').where('active','==',true).get(); return s.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>b.updatedAt-a.updatedAt); }
  async addResponse(pid,r){ await this.P().doc(pid).collection('responses').add(r); }
  async listResponses(pid){ const s=await this.P().doc(pid).collection('responses').orderBy('at').get(); return s.docs.map(d=>({id:d.id,...d.data()})); }
  async deleteResponse(pid,rid){ await this.P().doc(pid).collection('responses').doc(rid).delete(); }
  async addTest(pid,r){ await this.P().doc(pid).collection('tests').add(r); }
  async listTests(pid){ const s=await this.P().doc(pid).collection('tests').orderBy('at').get(); return s.docs.map(d=>({id:d.id,...d.data()})); }
  async deleteTest(pid,rid){ await this.P().doc(pid).collection('tests').doc(rid).delete(); }
  async listPhotos(pid){ const s=await this.P().doc(pid).collection('photos').orderBy('at').get(); return s.docs.map(d=>({id:d.id,...d.data()})); }
  async addPhoto(pid,ph){ await this.P().doc(pid).collection('photos').add({...ph,at:Date.now()}); }
  async updatePhoto(pid,id,data){ await this.P().doc(pid).collection('photos').doc(id).update(data); }
  async deletePhoto(pid,id){ await this.P().doc(pid).collection('photos').doc(id).delete(); }
  async getSettings(){ try{ const d=await this.db.collection('settings').doc('main').get(); return {...DEFAULT_SETTINGS, ...(d.exists?d.data():{})}; }catch(e){ return {...DEFAULT_SETTINGS}; } }
  async saveSettings(s){ await this.db.collection('settings').doc('main').set(s); }
  async getClasses(){ try{ const d=await this.db.collection('meta').doc('classes').get(); return d.exists ? d.data().list||[] : DEFAULT_CLASSES; }catch(e){ return DEFAULT_CLASSES; } }
  async saveClasses(list){ await this.db.collection('meta').doc('classes').set({list}); }
}

/* ------------------------------------------------------------------ */
/* ข้อมูลตัวอย่างสำหรับโหมดทดลอง                                       */
/* ------------------------------------------------------------------ */
function seedDemo(mem){
  let s=20261113; const rnd=()=>{ s|=0; s=s+0x6D2B79F5|0; let t=Math.imul(s^s>>>15,1|s); t=t+Math.imul(t^t>>>7,61|t)^t; return ((t^t>>>14)>>>0)/4294967296; };
  const pick=a=>a[Math.floor(rnd()*a.length)];
  const id='demo1'; const items=parseItems(DEFAULT_ITEMS); const nq=items.filter(i=>i.t==='q').length;
  const test=parseTest(`1. ดอกหญ้าที่เหมาะสำหรับทำไม้กวาดควรมีลักษณะอย่างไร
ก. สดและมีความชื้น
*ข. แห้งสนิท ไม่มีเชื้อรา
ค. ตัดมาใหม่ภายใน 1 วัน
ง. มีสีเขียวเข้ม
2. ก่อนนำดอกหญ้ามามัด ควรทำอย่างไรก่อน
*ก. ฟาดเอาเมล็ดและเศษออก
ข. แช่น้ำค้างคืน
ค. ทาสีให้สวยงาม
ง. ตัดให้สั้นที่สุด
3. วัสดุใดนิยมใช้ทำด้ามไม้กวาดดอกหญ้า
ก. เหล็กเส้น
ข. ท่อพีวีซีเท่านั้น
*ค. ไม้ไผ่หรือหวาย
ง. กระดาษแข็ง
4. การมัดดอกหญ้าให้แน่นควรใช้วัสดุใด
ก. หนังยาง
*ข. ลวดหรือเชือกไนลอน
ค. เทปใส
ง. ด้ายเย็บผ้า
5. ข้อใดช่วยยืดอายุการใช้งานของไม้กวาด
ก. วางไม้กวาดโดยให้ดอกหญ้ากดกับพื้น
*ข. แขวนไม้กวาดเมื่อไม่ใช้งาน
ค. ล้างน้ำทุกวัน
ง. ตากแดดจัดทั้งวัน
6. การกำหนดราคาขายไม้กวาดควรคำนึงถึงข้อใดมากที่สุด
*ก. ต้นทุนวัสดุและค่าแรง
ข. ราคาที่สูงที่สุดในตลาด
ค. ความชอบของผู้ผลิต
ง. สีของด้าม
7. ข้อใดเป็นช่องทางจำหน่ายที่เหมาะกับกลุ่มอาชีพในชุมชน
ก. ขายเฉพาะในครอบครัว
*ข. ตลาดนัดชุมชนและช่องทางออนไลน์
ค. ส่งออกต่างประเทศทันที
ง. ไม่ต้องจำหน่าย
8. การเก็บดอกหญ้าควรทำในช่วงใด
*ก. ช่วงที่ดอกหญ้าแก่เต็มที่
ข. ช่วงที่ดอกยังอ่อน
ค. ช่วงฝนตกหนัก
ง. เวลาใดก็ได้
9. ข้อใดคือความปลอดภัยในการใช้มีดตัดแต่งดอกหญ้า
ก. ตัดเข้าหาตัว
*ข. ตัดออกจากตัวและเก็บมีดให้เรียบร้อย
ค. ใช้มีดทื่อเพื่อความปลอดภัย
ง. วางมีดไว้บนพื้น
10. ประโยชน์ของการรวมกลุ่มผลิตไม้กวาดคือข้อใด
ก. ทำให้ต้นทุนสูงขึ้น
ข. แข่งขันกันเองในชุมชน
*ค. ลดต้นทุนและเพิ่มอำนาจต่อรอง
ง. ลดคุณภาพสินค้า`).questions;
  mem.projects[id]={ id, name:'โครงการอบรมอาชีพการทำไม้กวาดดอกหญ้า', fiscalYear:String(currentFY()),
    responsible:'นางสาวตัวอย่าง ใจดี', responsiblePos:'ครู ศกร.', owner:'', ownerName:'',
    rationale:'ตำบลมีดอกหญ้าเป็นวัตถุดิบในท้องถิ่นจำนวนมาก แต่ประชาชนส่วนใหญ่ยังไม่มีความรู้ในการแปรรูปเพื่อสร้างรายได้ ศูนย์ส่งเสริมการเรียนรู้ระดับอำเภอเทิงจึงจัดโครงการนี้ขึ้น เพื่อให้ประชาชนมีความรู้และทักษะในการทำไม้กวาดดอกหญ้า สามารถนำไปเป็นอาชีพเสริมเพิ่มรายได้ให้ครอบครัว',
    objectives:'เพื่อให้ผู้เข้าร่วมมีความรู้และทักษะในการทำไม้กวาดดอกหญ้า\nเพื่อให้ผู้เข้าร่วมสามารถนำความรู้ไปประกอบอาชีพเสริมได้',
    targetGroup:'ประชาชนทั่วไปในตำบลเวียง', targetQty:25,
    targetQual:'ผู้เข้าร่วมโครงการมีความรู้และทักษะในการทำไม้กวาดดอกหญ้า และมีความพึงพอใจในระดับมากขึ้นไป', targetQualPct:80,
    startDate:'2026-10-20', endDate:'2026-10-22', location:'ศาลาอเนกประสงค์ หมู่ 3 ตำบลเวียง อำเภอเทิง จังหวัดเชียงราย',
    budgetSource:'งบดำเนินงาน', budgetPlanned:9000, budgetSpent:8650,
    method:'ประชุมวางแผนและเสนอโครงการเพื่อขออนุมัติ\nประสานวิทยากรและกลุ่มเป้าหมาย\nทดสอบความรู้ก่อนเรียน\nจัดอบรมภาคทฤษฎีและภาคปฏิบัติ 3 วัน\nทดสอบความรู้หลังเรียนและประเมินความพึงพอใจ\nสรุปผลการดำเนินงาน',
    results:'ผู้เข้าร่วมสามารถทำไม้กวาดดอกหญ้าได้ด้วยตนเองคนละ 2–3 ด้าม และมีผู้เข้าร่วม 6 คนรวมกลุ่มกันผลิตเพื่อจำหน่าย',
    problems:'ดอกหญ้าบางส่วนยังไม่แห้งสนิท ทำให้ต้องเลื่อนการฝึกปฏิบัติบางขั้นตอน',
    suggestions:'ควรจัดอบรมต่อยอดด้านการออกแบบบรรจุภัณฑ์และการตลาดออนไลน์',
    actualQty:'', items, surveyOpen:true,
    test:{questions:test, preOpen:true, postOpen:true, showScore:true, passPct:60}, updatedAt:Date.now() };
  const fn=['สมศรี','บุญมา','อำไพ','ประยูร','จันทร์เพ็ญ','สมบัติ','แสงเดือน','คำปัน','บัวผัน','ทองใบ','นงลักษณ์','สุพัตรา','อินทร','ศรีนวล','วันเพ็ญ','จำปา','บุญศรี','คำแปง','ศรีไพร','มณี','สมพร','ดวงใจ','พิมพ์'];
  const ln=['ใจดี','คำมา','ปัญญาดี','แก้วมูล','สุขสม','ทองคำ','อินต๊ะ','ศรีวงค์','บุญเรือง','จันทร์แก้ว'];
  const comments=['อยากให้จัดอบรมอีก','ได้ความรู้มาก นำไปทำขายได้','อยากให้สอนเรื่องการขายออนไลน์','ระยะเวลาสั้นไปนิด','วิทยากรสอนเข้าใจง่าย'];
  const R={}, T={};
  for(let k=0;k<23;k++){
    const female=rnd()<0.75; const name=(female?pick(['นาง','นางสาว']):'นาย')+fn[k]+' '+pick(ln);
    const room=roomLabel(DEFAULT_CLASSES[[2,10,0][k%3]]); const level=pick([EDU_LEVELS[0],EDU_LEVELS[1],EDU_LEVELS[1],EDU_LEVELS[2],EDU_LEVELS[2],EDU_LEVELS[2]]);
    const sid=String(6710000000+k*137+11);
    const base=rnd()<0.6?5:4;
    const scores=Array.from({length:nq},()=>{ const r=rnd(); return Math.max(3,Math.min(5, r<0.65?base:(r<0.85?base-1:5))); });
    const t0=Date.now()-86400000*3;
    R['r'+k]={ name, gender:female?'หญิง':'ชาย',
      age:pick([AGES[1],AGES[1],AGES[1],AGES[2],AGES[2],AGES[3]]), level, room,
      scores, comment:rnd()<0.3?pick(comments):'', at:t0+86400000*2+k*60000 };
    const ans=p=>test.map(q=>rnd()<p?q.ans:(q.ans+1+Math.floor(rnd()*3))%4);
    T['a'+k]={kind:'pre',name,sid,level,room,answers:ans(0.42),at:t0+k*60000};
    if(k<21) T['b'+k]={kind:'post',name,sid,level,room,answers:ans(0.76),at:t0+86400000*2+k*60000};
  }
  mem.responses[id]=R; mem.tests[id]=T;
}

/* ------------------------------------------------------------------ */
/* สถิติ                                                               */
/* ------------------------------------------------------------------ */
function msd(vals){
  const n=vals.length; if(!n) return {n:0,mean:0,sd:0};
  const mean=vals.reduce((a,b)=>a+b,0)/n;
  const sd=n>1 ? Math.sqrt(vals.reduce((a,b)=>a+(b-mean)**2,0)/(n-1)) : 0;
  return {n,mean,sd};
}
function computeStats(p, resps){
  const items=p.items||[]; const rows=[]; let qi=0, cur=null; const all=[];
  const valid=v=>Number.isFinite(v)&&v>=1&&v<=5;
  for(const it of items){
    if(it.t==='h'){ cur={t:'h',text:it.text,vals:[]}; rows.push(cur); continue; }
    const vals=resps.map(r=>r.scores?.[qi]).filter(valid);
    const s=msd(vals); rows.push({t:'q',text:it.text,qi,...s,lvl:level(s.mean)});
    if(cur) cur.vals.push(...vals); all.push(...vals); qi++;
  }
  for(const r of rows) if(r.t==='h'){ Object.assign(r,msd(r.vals)); r.lvl=level(r.mean); delete r.vals; }
  const overall=msd(all); overall.lvl=level(overall.mean);
  const perPerson=resps.map(r=>msd((r.scores||[]).filter(valid)).mean).filter(x=>x>0);
  const good=perPerson.filter(m=>m>=3.51).length;
  const dist=(key,labels)=>{ const c={}; resps.forEach(r=>{ const v=r[key]||'ไม่ระบุ'; c[v]=(c[v]||0)+1; });
    const keys=[...labels.filter(l=>c[l]), ...Object.keys(c).filter(k=>!labels.includes(k))];
    return keys.map(k=>({label:k,n:c[k],pct:pct(c[k],resps.length)})); };
  const n=resps.length;
  const actual = p.actualQty!=='' && p.actualQty!=null ? num(p.actualQty) : n;
  return { n, rows, overall, pctGood:pct(good,perPerson.length), actual,
    gender:dist('gender',GENDERS), age:dist('age',AGES), level:dist('level',EDU_LEVELS), room:dist('room',[]),
    comments:resps.map(r=>(r.comment||'').trim()).filter(Boolean) };
}
function testStats(p, rows){
  const Q=p.test?.questions||[], N=Q.length, cut=num(p.test?.passPct)||60;
  const sc=r=>Q.reduce((s,q,i)=>s+((r.answers||[])[i]===q.ans?1:0),0);
  const lk=r=>r.sid?'id:'+String(r.sid).replace(/\s/g,''):'n:'+nameKey(r.name);
  const latest=k=>{ const m=new Map(); rows.filter(r=>r.kind===k).sort((a,b)=>a.at-b.at).forEach(r=>m.set(lk(r),{...r,score:sc(r)})); return [...m.values()]; };
  const pre=latest('pre'), post=latest('post');
  const sum=a=>{ const s=msd(a.map(x=>x.score)); const pass=a.filter(x=>N&&pct(x.score,N)>=cut).length; return {...s,pm:pct(s.mean,N),pass,passP:pct(pass,a.length)}; };
  const preM=new Map(pre.map(x=>[lk(x),x])), postM=new Map(post.map(x=>[lk(x),x]));
  const keys=[...new Set([...preM.keys(),...postM.keys()])];
  const learners=keys.map(k=>{ const a=preM.get(k), b=postM.get(k), x=b||a; return {name:x.name, sid:x.sid||'', room:x.room||'', level:x.level||'', pre:a?a.score:null, post:b?b.score:null}; })
    .sort((a,b)=>(a.room||'').localeCompare(b.room||'','th')||String(a.sid).localeCompare(String(b.sid)));
  const roomNames=[...new Set(learners.map(l=>l.room||'ไม่ระบุ'))];
  const rooms=roomNames.map(r=>{ const L=learners.filter(l=>(l.room||'ไม่ระบุ')===r); const a=msd(L.filter(l=>l.pre!=null).map(l=>l.pre)), b=msd(L.filter(l=>l.post!=null).map(l=>l.post));
    return {room:r, preN:a.n, preM:a.mean, postN:b.n, postM:b.mean}; });
  const pairs=learners.filter(l=>l.pre!=null&&l.post!=null);
  const improved=pairs.filter(l=>l.post>l.pre).length;
  const diff=msd(pairs.map(l=>l.post-l.pre));
  const items=Q.map((q,i)=>({q:q.q, pre:pct(pre.filter(r=>(r.answers||[])[i]===q.ans).length,pre.length), post:pct(post.filter(r=>(r.answers||[])[i]===q.ans).length,post.length)}));
  return {N,cut,pre,post,preS:sum(pre),postS:sum(post),learners,rooms,pairs,improved,diff,items,has:N>0&&(pre.length+post.length)>0};
}
function testText(ts){
  if(!ts.has) return '';
  const parts=[`ผลการทดสอบความรู้ (คะแนนเต็ม ${ts.N} คะแนน) พบว่า`];
  if(ts.pre.length) parts.push(`ก่อนเรียนมีผู้ทำแบบทดสอบ ${ts.pre.length} คน คะแนนเฉลี่ย ${f2(ts.preS.mean)} คะแนน (ร้อยละ ${f2(ts.preS.pm)})`);
  if(ts.post.length) parts.push(`${ts.pre.length?'และ':''}หลังเรียนมีผู้ทำแบบทดสอบ ${ts.post.length} คน คะแนนเฉลี่ย ${f2(ts.postS.mean)} คะแนน (ร้อยละ ${f2(ts.postS.pm)}) ผ่านเกณฑ์ร้อยละ ${ts.cut} จำนวน ${ts.postS.pass} คน คิดเป็นร้อยละ ${f2(ts.postS.passP)}`);
  let s=parts.join(' ');
  if(ts.pairs.length) s+=` เมื่อเปรียบเทียบผู้เรียนที่ทำแบบทดสอบทั้งสองครั้งจำนวน ${ts.pairs.length} คน พบว่าคะแนนหลังเรียน${ts.diff.mean>=0?'สูงกว่า':'ต่ำกว่า'}ก่อนเรียนเฉลี่ย ${f2(Math.abs(ts.diff.mean))} คะแนน และมีผู้เรียนที่มีคะแนนเพิ่มขึ้น ${ts.improved} คน คิดเป็นร้อยละ ${f2(pct(ts.improved,ts.pairs.length))}`;
  return s;
}
function satisfactionText(p, st){
  const tq=num(p.targetQualPct)||80; const parts=[];
  parts.push(`ผู้เข้าร่วมโครงการมีความพึงพอใจโดยรวมอยู่ในระดับ${st.overall.lvl} (x̄ = ${f2(st.overall.mean)}, S.D. = ${f2(st.overall.sd)}, คิดเป็นร้อยละ ${f2(st.overall.mean*20)}) และผู้เข้าร่วมร้อยละ ${f2(st.pctGood)} มีความพึงพอใจในระดับมากขึ้นไป ซึ่ง${st.pctGood>=tq?'บรรลุ':'ยังไม่บรรลุ'}เป้าหมายเชิงคุณภาพที่กำหนดไว้ร้อยละ ${tq}`);
  const secs=st.rows.filter(r=>r.t==='h'&&r.n); const qs=st.rows.filter(r=>r.t==='q'&&r.n);
  if(secs.length>1){ const s=[...secs].sort((a,b)=>b.mean-a.mean);
    parts.push(`เมื่อพิจารณาเป็นรายด้าน พบว่า ${s[0].text}มีค่าเฉลี่ยสูงสุด (x̄ = ${f2(s[0].mean)}) ส่วน${s.at(-1).text}มีค่าเฉลี่ยต่ำสุด (x̄ = ${f2(s.at(-1).mean)})`); }
  if(qs.length>1){ const s=[...qs].sort((a,b)=>b.mean-a.mean);
    parts.push(`เมื่อพิจารณาเป็นรายข้อ พบว่า ข้อที่มีค่าเฉลี่ยสูงสุด คือ ${s[0].text} (x̄ = ${f2(s[0].mean)}) และข้อที่มีค่าเฉลี่ยต่ำสุด คือ ${s.at(-1).text} (x̄ = ${f2(s.at(-1).mean)})`); }
  return parts.join(' ');
}
function autoSummary(p, st, ts){
  const t=num(p.targetQty); const parts=[];
  let s=`จากการดำเนิน${p.name||'โครงการ'}${p.startDate?' ระหว่าง'+dateRange(p):''}${p.location?' ณ '+p.location:''} มีผู้เข้าร่วมโครงการจำนวน ${st.actual} คน`;
  if(t) s+=` จากเป้าหมาย ${t} คน คิดเป็นร้อยละ ${f2(pct(st.actual,t))} ซึ่ง${st.actual>=t?'บรรลุ':'ยังไม่บรรลุ'}เป้าหมายเชิงปริมาณ`;
  parts.push(s);
  if(ts && ts.has) parts.push(testText(ts));
  if(st.n) parts.push(`จากการประเมินความพึงพอใจของผู้เข้าร่วมโครงการจำนวน ${st.n} คน พบว่า ` + satisfactionText(p,st));
  return parts.join(' ');
}

/* ------------------------------------------------------------------ */
/* สถานะแอปและการนำทาง                                                 */
/* ------------------------------------------------------------------ */
const S={ store:null, settings:{...DEFAULT_SETTINGS}, user:null, firebase:!!FIREBASE_CONFIG, ready:false };
const isAdmin = () => S.user?.role==='admin';
const inFrame = (()=>{ try{ return window.self!==window.top; }catch(e){ return true; } })();

async function boot(){
  try{
    S.store = S.firebase ? new FireStore() : new LocalStore();
    await S.store.init();
    S.user=await S.store.currentUser();
    if(!S.firebase || S.user) S.settings=await S.store.getSettings();
  }catch(e){ app.innerHTML=`<div class="card"><h2>เริ่มระบบไม่สำเร็จ</h2><p>${esc(e.message)}</p></div>`; return; }
  S.ready=true; drawHeader(); route();
}
window.addEventListener('hashchange',()=>S.ready&&route());

function drawHeader(){
  const h=$('#hdr'); if(!h) return;
  const org=S.settings.org;
  if(MODE==='learner'){
    h.innerHTML=`<img class="logo" src="logo.png" alt=""><div class="brand brand-lg"><span class="shine"><span class="shine-text">${esc(org)}</span></span><small>แบบทดสอบและแบบประเมินสำหรับผู้เรียน</small></div>`;
    return;
  }
  h.innerHTML=`<img class="logo" src="logo.png" alt=""><div class="brand">ระบบสรุปผลการดำเนินโครงการ<small>${esc(org)}</small></div>
    <nav>${S.user?`<a href="#home">โครงการ</a><a href="#new">+ โครงการใหม่</a>${isAdmin()?'<a href="#staff">บุคลากร</a><a href="#rooms">ห้องเรียน</a><a href="#org">ตั้งค่าหน่วยงาน</a>':''}
      <span class="who">👤 ${esc(S.user.name)}${isAdmin()?' (ผู้ดูแล)':''}</span><button id="logoutBtn">ออกจากระบบ</button>`:''}
      ${!S.firebase?'<span class="badge">โหมดทดลอง</span>':''}</nav>`;
  if($('#logoutBtn')) $('#logoutBtn').onclick=async()=>{ await S.store.logout(); S.user=null; drawHeader(); location.hash='login'; route(); };
}

async function route(){
  const h=decodeURIComponent(location.hash.slice(1));
  const i=h.indexOf('-'); const view=i<0?h:h.slice(0,i); const id=i<0?'':h.slice(i+1);
  window.scrollTo(0,0);
  const want = PAGE_MODE==='admin' || ADMIN_VIEWS.includes(view) ? 'admin' : (['pre','post','sat'].includes(view) && MODE==='admin' ? 'admin' : 'learner');
  if(want!==MODE){ MODE=want; if(MODE==='admin' && !S.user) S.user=await S.store.currentUser(); drawHeader(); }
  try{
    if(view==='pre'||view==='post') return await renderTest(id, view);
    if(view==='sat') return await renderSurvey(id);
    if(MODE==='learner'){
      if(view==='j') return await renderJoin(id);
      return await renderLearnerHome();
    }
    if(view==='setup') return await renderSetup();
    if(!S.user){ if(!S.firebase && !S.store.hasStaff()) return renderSetup(); return renderLogin(); }
    if(view==='new') return await renderEdit(null);
    if(view==='edit') return await renderEdit(id);
    if(view==='p') return await renderProject(id);
    if(view==='r') return await renderReport(id);
    if(view==='test') return await renderTestEditor(id);
    if(view==='org' && isAdmin()) return await renderSettings();
    if(view==='staff' && isAdmin()) return await renderStaff();
    if(view==='rooms' && isAdmin()) return await renderRooms();
    return await renderHome();
  }catch(e){
    console.error(e);
    app.innerHTML=`<div class="card"><h2>เกิดข้อผิดพลาด</h2><p>${esc(e.message||e)}</p><a class="btn ghost" href="#">กลับหน้าหลัก</a></div>`;
  }
}

/* ================================================================== */
/*                         หน้าสำหรับผู้เรียน                           */
/* ================================================================== */
async function renderLearnerHome(){
  app.innerHTML='<div class="empty">กำลังโหลด…</div>';
  const list=await S.store.listPublic();
  app.innerHTML=`
    <div class="lhero"><h1>สวัสดีครับ/ค่ะ ผู้เรียนทุกท่าน</h1><div>เลือกโครงการหรือกิจกรรมที่เข้าร่วม แล้วทำแบบทดสอบก่อนเรียน หลังเรียน และแบบประเมินความพึงพอใจ</div></div>
    ${!S.firebase?'<div class="note"><b>โหมดทดลอง:</b> ข้อมูลเก็บในเบราว์เซอร์เครื่องนี้เท่านั้น</div>':''}
    <input type="search" id="lq" placeholder="🔍 ค้นหาชื่อโครงการ" style="margin-bottom:14px">
    <div class="lcards" id="lc"></div>
`;
  const draw=()=>{
    const q=$('#lq').value.trim().toLowerCase();
    const rows=list.filter(d=>!q||d.title.toLowerCase().includes(q));
    $('#lc').innerHTML=rows.length?rows.map(d=>`<a class="lcard" href="#j-${esc(d.id)}"><b>${esc(d.title)}</b>
      <div class="muted small">${esc(d.dateText)}${d.location?' · '+esc(d.location):''}</div>
      <div class="row" style="margin-top:8px;gap:6px">${d.test.preOpen?'<span class="pill">ก่อนเรียน</span>':''}${d.test.postOpen?'<span class="pill">หลังเรียน</span>':''}${d.sat.open?'<span class="pill">ประเมินความพึงพอใจ</span>':''}</div></a>`).join('')
      :'<div class="empty" style="grid-column:1/-1">ยังไม่มีโครงการที่เปิดให้ทำแบบทดสอบหรือแบบประเมิน</div>';
  };
  $('#lq').oninput=draw; draw();
}
async function renderJoin(id){
  app.innerHTML='<div class="empty">กำลังโหลด…</div>';
  const d=await S.store.getPublic(id);
  if(!d){ app.innerHTML='<div class="card sv"><h2>ไม่พบโครงการ</h2><p>ลิงก์อาจไม่ถูกต้อง กรุณาสแกน QR Code อีกครั้ง</p><a href="#" class="btn ghost">ดูโครงการทั้งหมด</a></div>'; return; }
  const nq=(d.test.questions||[]).length;
  const btn=(href,ic,t,open,sub)=>`<a class="bigbtn ${open?'':'off'}" href="${href}"><span class="ic">${ic}</span><span>${t}<small>${open?sub:'ยังไม่เปิดให้ทำ'}</small></span></a>`;
  app.innerHTML=`<div class="sv">
    <div class="row" style="margin-bottom:8px"><a href="#" class="backbtn">← โครงการทั้งหมด</a></div>
    <div class="card"><div class="muted small">${esc(d.org)}</div><h1>${esc(d.title)}</h1><div class="muted small">${esc(d.dateText)}${d.location?' · '+esc(d.location):''}</div></div>
    <div class="big">
      ${btn('#pre-'+esc(id),'📝','แบบทดสอบก่อนเรียน',d.test.preOpen,`${nq} ข้อ · ทำก่อนเริ่มกิจกรรม`)}
      ${btn('#post-'+esc(id),'✅','แบบทดสอบหลังเรียน',d.test.postOpen,`${nq} ข้อ · ทำหลังจบกิจกรรม`)}
      ${btn('#sat-'+esc(id),'⭐','แบบประเมินความพึงพอใจ',d.sat.open,'ใช้เวลาประมาณ 3 นาที')}
    </div></div>`;
}
const savedInfo=()=>{ try{ return JSON.parse(localStorage.getItem('psum-learner2')||'{}'); }catch(e){ return {}; } };
const saveInfo=o=>{ try{ localStorage.setItem('psum-learner2',JSON.stringify({...savedInfo(),...o})); }catch(e){} };
const chipsHTML=(name,opts,sel)=>`<div class="chips">${opts.map((o,i)=>`<input type="radio" name="${name}" id="${name}${i}" value="${esc(o)}" ${o===sel?'checked':''}><label for="${name}${i}">${esc(o)}</label>`).join('')}</div>`;
function roomSelect(classes, sel){
  return `<select name="room"><option value="">-- เลือกศูนย์ / ครูประจำกลุ่ม --</option>${classes.map(c=>{ const l=roomLabel(c); return `<option ${l===sel?'selected':''}>${esc(l)}</option>`; }).join('')}</select>`;
}
/* ช่องข้อมูลผู้เรียน (ใช้ทั้งแบบทดสอบและแบบประเมิน) */
function learnerFields(classes, pre, opt){
  const v=pre||{};
  return `<div class="q" data-req="prefix"><div class="qt">คำนำหน้า</div>${chipsHTML('prefix',PREFIXES,v.prefix)}</div>
    <div class="q" data-req="name"><div class="qt">ชื่อ-สกุล <span class="muted small">(ไม่ต้องพิมพ์คำนำหน้า)</span></div><input type="text" name="name" maxlength="150" value="${esc(v.name||'')}" autocomplete="name" placeholder="เช่น สมชาย ใจดี"></div>
    ${opt.sid?`<div class="q" data-req="sid"><div class="qt">รหัสนักศึกษา</div><input type="text" name="sid" maxlength="30" inputmode="numeric" value="${esc(v.sid||'')}"></div>`:''}
    ${opt.gender?`<div class="q" data-req="gender"><div class="qt">เพศ</div>${chipsHTML('gender',GENDERS,v.gender)}</div>`:''}
    ${opt.age?`<div class="q" data-req="age"><div class="qt">อายุ</div>${chipsHTML('age',AGES,v.age)}</div>`:''}
    <div class="q" data-req="level"><div class="qt">ระดับชั้น</div>${chipsHTML('level',EDU_LEVELS,v.level)}</div>
    <div class="q" data-req="room"><div class="qt">เรียนอยู่ห้องเรียน</div>${roomSelect(classes,v.room)}</div>`;
}
function bindLearnerFields(f){
  const clr=e=>{ const q=e.target.closest('.q'); if(q) q.classList.remove('miss'); };
  f.addEventListener('change',clr); f.addEventListener('input',clr);
  $$('input[name=prefix]',f).forEach(r=>r.addEventListener('change',()=>{ const g=PREFIX_GENDER[r.value]; const gi=g&&$(`input[name=gender][value="${g}"]`,f); if(gi) gi.checked=true; }));
}
function checkRequired(f){
  let first=null;
  $$('.q[data-req]',f).forEach(q=>{ const k=q.dataset.req; const el=f.elements[k];
    const ok = el && (el.tagName==='SELECT'||el.type==='text') ? !!el.value.trim() : !!$(`input[name="${k}"]:checked`,f);
    q.classList.toggle('miss',!ok); if(!ok&&!first) first=q; });
  return first;
}
function backLink(id){ return `<a class="backbtn" href="#${MODE==='admin'&&S.user?'p':'j'}-${esc(id)}">กลับไปหน้าโครงการ</a>`; }

async function renderTest(id, kind){
  app.innerHTML='<div class="empty">กำลังโหลดแบบทดสอบ…</div>';
  const [d, classes]=await Promise.all([S.store.getPublic(id), S.store.getClasses()]);
  const label=kind==='pre'?'แบบทดสอบก่อนเรียน':'แบบทดสอบหลังเรียน';
  if(!d){ app.innerHTML='<div class="card sv"><h2>ไม่พบแบบทดสอบ</h2></div>'; return; }
  const staffEntry = MODE==='admin' && S.user;
  if(!d.test[kind+'Open'] && !staffEntry){ app.innerHTML=`<div class="card sv"><h2>${label}</h2><p>${esc(d.title)}</p><p>แบบทดสอบนี้ยังไม่เปิดหรือปิดรับแล้ว</p>${backLink(id)}</div>`; return; }
  const Q=d.test.questions||[];
  if(!Q.length){ app.innerHTML=`<div class="card sv"><h2>${label}</h2><p>ยังไม่มีข้อสอบ</p>${backLink(id)}</div>`; return; }
  app.innerHTML=`<form class="sv" id="tf" novalidate>
    <div class="row" style="margin-bottom:8px">${backLink(id).replace('กลับไปหน้าโครงการ','← กลับ')}</div>
    <div class="card"><div class="muted small">${label}${staffEntry?' · <b>ครูบันทึกแทนผู้เรียน</b>':''}</div><h1>${esc(d.title)}</h1><div class="muted small">${Q.length} ข้อ · เลือกคำตอบที่ถูกที่สุดเพียงข้อเดียว</div></div>
    <h3>ข้อมูลผู้เรียน</h3>
    ${learnerFields(classes, staffEntry?{}:savedInfo(), {sid:true})}
    <h3>คำถาม</h3>
    ${(()=>{ const L=testLayout(id,kind,d.test,staffEntry); return L.qOrder.map((i,k)=>{ const q=Q[i];
      return `<div class="q" data-req="a${i}"><div class="qt">${k+1}. ${esc(q.q)}</div><div class="opts">
      ${L.optOrders[i].map((j,m)=>`<input type="radio" name="a${i}" id="a${i}_${j}" value="${j}"><label for="a${i}_${j}"><i>${OPT_LABELS[m]||m+1}.</i><span>${esc(q.opts[j])}</span></label>`).join('')}</div></div>`; }).join(''); })()}
    <div class="row" style="margin:16px 0"><button class="btn" id="tbtn" style="font-size:17px;padding:10px 26px">ส่งคำตอบ</button><span class="err" id="terr"></span></div>
  </form>`;
  const f=$('#tf'); bindLearnerFields(f);
  f.onsubmit=async e=>{
    e.preventDefault();
    const first=checkRequired(f);
    if(first){ first.scrollIntoView({behavior:'smooth',block:'center'}); $('#terr').textContent='กรุณากรอกและตอบให้ครบทุกข้อ'; return; }
    const val=n=>$(`input[name="${n}"]:checked`,f)?.value||'';
    const info={prefix:val('prefix'), name:normName(f.name.value), sid:f.sid.value.trim(), level:val('level'), room:f.room.value};
    const answers=Q.map((_,i)=>+val('a'+i));
    $('#tbtn').disabled=true; $('#terr').textContent='กำลังส่ง…';
    try{
      await S.store.addTest(id,{kind,name:info.prefix+info.name,sid:info.sid,level:info.level,room:info.room,answers,at:Date.now()});
      if(!staffEntry) saveInfo(info);
      const key=d.test.key; let scoreHtml='';
      if(key){ const sc=answers.reduce((s,a,i)=>s+(a===key[i]?1:0),0); const pc=pct(sc,Q.length);
        scoreHtml=`<div class="scorebig">${sc} / ${Q.length}</div><p>คิดเป็นร้อยละ ${f2(pc)} ${kind==='post'?(pc>=d.test.passPct?'· <b class="opt-ok">ผ่านเกณฑ์</b>':'· ยังไม่ผ่านเกณฑ์'):''}</p>`; }
      app.innerHTML=`<div class="card sv" style="text-align:center;padding:36px 20px"><div style="font-size:44px">${kind==='pre'?'📝':'🎉'}</div>
        <h2>ส่ง${label}แล้ว</h2><p class="muted">${esc(info.prefix+info.name)} · รหัส ${esc(info.sid)}</p>${scoreHtml}
        <div class="row" style="justify-content:center;margin-top:10px">${staffEntry?`<button class="btn ghost" id="again">บันทึกของผู้เรียนคนถัดไป</button>`:''}${backLink(id)}</div></div>`;
      if($('#again')) $('#again').onclick=()=>renderTest(id,kind);
      window.scrollTo(0,0);
    }catch(err){ $('#tbtn').disabled=false; $('#terr').textContent='ส่งไม่สำเร็จ กรุณาลองใหม่ ('+err.message+')'; }
  };
}

async function renderSurvey(id){
  app.innerHTML='<div class="empty">กำลังโหลดแบบประเมิน…</div>';
  const [d, classes]=await Promise.all([S.store.getPublic(id), S.store.getClasses()]);
  if(!d){ app.innerHTML='<div class="card sv"><h2>ไม่พบแบบประเมิน</h2><p>ลิงก์อาจไม่ถูกต้อง กรุณาสแกน QR Code อีกครั้ง</p></div>'; return; }
  const staffEntry = MODE==='admin' && S.user;
  if(!d.sat.open && !staffEntry){ app.innerHTML=`<div class="card sv"><h2>${esc(d.title)}</h2><p>แบบประเมินนี้ปิดรับคำตอบแล้ว ขอบคุณที่ให้ความสนใจ</p>${backLink(id)}</div>`; return; }
  let qi=0;
  const saved=staffEntry?{}:savedInfo();
  if(!saved.gender && saved.prefix) saved.gender=PREFIX_GENDER[saved.prefix];
  app.innerHTML=`<form class="sv" id="sf" novalidate>
    <div class="row" style="margin-bottom:8px">${backLink(id).replace('กลับไปหน้าโครงการ','← กลับ')}</div>
    <div class="card"><div class="muted small">แบบประเมินความพึงพอใจ${staffEntry?' · <b>ครูบันทึกจากแบบประเมินกระดาษ</b>':''}</div><h1>${esc(d.title)}</h1>
      <div class="muted small">${esc(d.dateText)}${d.location?' · '+esc(d.location):''}</div></div>
    <h3>ตอนที่ 1 ข้อมูลผู้ตอบแบบประเมิน</h3>
    ${learnerFields(classes, saved, {gender:true, age:true})}
    <h3>ตอนที่ 2 ความพึงพอใจต่อโครงการ</h3>
    <p class="muted small" style="margin-top:0">5 = มากที่สุด · 4 = มาก · 3 = ปานกลาง · 2 = น้อย · 1 = น้อยที่สุด</p>
    ${d.sat.items.map(it=>{ if(it.t==='h') return `<div style="font-weight:700;margin:14px 0 6px">${esc(it.text)}</div>`;
      const k=qi++; return `<div class="q" data-req="s${k}"><div class="qt">${k+1}. ${esc(it.text)}</div><div class="scale">
        ${[5,4,3,2,1].map(v=>`<input type="radio" name="s${k}" id="s${k}_${v}" value="${v}"><label for="s${k}_${v}"><b>${v}</b>${SCALE_LABELS[v]}</label>`).join('')}</div></div>`; }).join('')}
    <h3>ตอนที่ 3 ข้อเสนอแนะ</h3>
    <div class="q"><textarea name="comment" maxlength="2000" placeholder="(ไม่บังคับ)"></textarea></div>
    <div class="row" style="margin:16px 0"><button class="btn" id="sbtn" style="font-size:17px;padding:10px 26px">ส่งแบบประเมิน</button><span class="err" id="serr"></span></div>
  </form>`;
  const nq=qi; const f=$('#sf'); bindLearnerFields(f);
  f.onsubmit=async e=>{
    e.preventDefault();
    const first=checkRequired(f);
    if(first){ first.scrollIntoView({behavior:'smooth',block:'center'}); $('#serr').textContent='กรุณาตอบให้ครบทุกข้อ'; return; }
    const val=n=>$(`input[name="${n}"]:checked`,f)?.value||'';
    const info={prefix:val('prefix'), name:normName(f.name.value), level:val('level'), room:f.room.value};
    const r={ name:info.prefix+info.name, gender:val('gender'), age:val('age'), level:info.level, room:info.room,
      scores:Array.from({length:nq},(_,k)=>+val('s'+k)), comment:f.comment.value.trim(), at:Date.now() };
    $('#sbtn').disabled=true; $('#serr').textContent='กำลังส่ง…';
    try{
      await S.store.addResponse(id,r); if(!staffEntry) saveInfo(info);
      app.innerHTML=`<div class="card sv" style="text-align:center;padding:40px 20px"><div style="font-size:48px">🙏</div><h2>ขอบคุณที่ตอบแบบประเมิน</h2>
        <p class="muted">${esc(d.title)}</p><div class="row" style="justify-content:center">${staffEntry?'<button class="btn ghost" id="again">บันทึกของผู้เข้าร่วมคนถัดไป</button>':''}${backLink(id)}</div></div>`;
      if($('#again')) $('#again').onclick=()=>renderSurvey(id); window.scrollTo(0,0);
    }catch(err){ $('#sbtn').disabled=false; $('#serr').textContent='ส่งไม่สำเร็จ กรุณาลองใหม่ ('+err.message+')'; }
  };
}

/* ================================================================== */
/*                       หน้าสำหรับครูและแอดมิน                         */
/* ================================================================== */
function renderLogin(){
  app.innerHTML=`<div class="card login">
    <div style="text-align:center"><img src="logo.png" alt="" style="width:110px;height:110px"></div>
    <h2 style="text-align:center">เข้าสู่ระบบ (ครูและเจ้าหน้าที่)</h2>
    ${!S.firebase?'<div class="note">โหมดทดลอง: ใช้ชื่อและเบอร์โทรตามไฟล์รายชื่อที่นำเข้า หรือบัญชีทดลอง ชื่อ <b>ทดลอง</b> รหัส <b>1234</b> (ถ้าสร้างไว้)</div>':''}
    <form id="lf"><label>ชื่อ <span class="hint">(พิมพ์เฉพาะชื่อ ไม่ต้องมีคำนำหน้าและนามสกุล)</span></label><input type="text" name="name" required autocomplete="username" placeholder="เช่น วนศาสตร์">
    <label>รหัสผ่าน <span class="hint">(เบอร์โทรศัพท์ 10 หลัก)</span></label><input type="password" name="pw" required inputmode="numeric" autocomplete="current-password" placeholder="0xxxxxxxxx">
    <div class="row" style="margin-top:16px"><button class="btn" id="lbtn">เข้าสู่ระบบ</button><span class="err" id="lerr"></span></div></form>
    <p class="muted small" style="margin-top:18px">ผู้เรียนไม่ต้องเข้าสู่ระบบ — <a href="${PAGE_MODE==='learner'?'#':'index.html'}">ไปหน้าผู้เรียน</a><br>
    ติดตั้งระบบครั้งแรก? <a href="#setup">นำเข้ารายชื่อบุคลากรจาก Excel</a></p></div>`;
  $('#lf').onsubmit=async e=>{
    e.preventDefault(); const f=e.target; $('#lerr').textContent='กำลังตรวจสอบ…'; $('#lbtn').disabled=true;
    try{ S.user=await S.store.login(f.name.value, f.pw.value); S.settings=await S.store.getSettings(); drawHeader(); location.hash='home'; route(); }
    catch(err){ $('#lerr').textContent=err.message; $('#lbtn').disabled=false; }
  };
}

function staffPreview(rows){
  return `<table class="t"><thead><tr><th>ชื่อ (ใช้ล็อกอิน)</th><th>สกุล</th><th>ตำแหน่ง</th><th>รหัสผ่าน</th><th>สิทธิ์</th></tr></thead><tbody>
    ${rows.map((r,i)=>`<tr><td><b>${esc(r.name)}</b></td><td>${esc(r.surname)}</td><td>${esc(r.position)}</td>
      <td class="small">${r.phone?esc(r.phone.slice(0,3))+'••••'+esc(r.phone.slice(-3)):'<span class="err">ไม่มี</span>'}</td>
      <td><select data-i="${i}" class="rolesel" style="width:auto;padding:4px 6px"><option value="teacher" ${r.role==='teacher'?'selected':''}>ครู</option><option value="admin" ${r.role==='admin'?'selected':''}>ผู้ดูแลระบบ</option></select></td></tr>`).join('')}
    </tbody></table>`;
}

async function renderSetup(){
  app.innerHTML=`<div class="card" style="max-width:860px;margin:0 auto">
    <h2>ตั้งค่าระบบครั้งแรก: นำเข้ารายชื่อบุคลากร</h2>
    <p>เลือกไฟล์ Excel รายชื่อบุคลากร ไฟล์ต้องมีหัวตาราง <b>ชื่อ</b> และ <b>เบอร์โทร</b> (มีคำนำหน้า สกุล และตำแหน่งได้) ระบบจะสร้างบัญชีให้ทุกคน
    <b>ล็อกอินด้วยชื่ออย่างเดียว และใช้เบอร์โทรเป็นรหัสผ่าน</b> ผู้ที่มีตำแหน่ง "ผู้บริหาร" จะได้สิทธิ์ผู้ดูแลระบบ (เปลี่ยนได้ในตาราง)</p>
    <p class="muted small">เบอร์โทรจะไม่ถูกบันทึกเป็นข้อความในระบบ${S.firebase?' แต่ใช้ตั้งเป็นรหัสผ่านในระบบล็อกอินของ Firebase':' โหมดทดลองจะเก็บไว้ในรูปที่เข้ารหัสแบบทางเดียว'}</p>
    <label class="btn" style="width:auto;display:inline-flex">📂 เลือกไฟล์ Excel<input type="file" id="xf" accept=".xlsx,.xls,.csv" hidden></label>
    ${!S.firebase?'<button class="btn ghost" id="demoAcc" style="margin-left:8px">ข้ามไปใช้บัญชีทดลอง</button>':''}
    <div id="pv" style="margin-top:16px;overflow-x:auto"></div></div>`;
  let rows=[];
  if($('#demoAcc')) $('#demoAcc').onclick=async()=>{
    await S.store.addStaff([{name:'ทดลอง',surname:'ระบบ',position:'ผู้ดูแลระบบ (ทดลอง)',phone:'1234',role:'admin'}]);
    toast('สร้างบัญชีทดลองแล้ว: ชื่อ ทดลอง รหัส 1234'); location.hash='login'; route();
  };
  $('#xf').onchange=async e=>{
    const file=e.target.files[0]; if(!file) return;
    $('#pv').innerHTML='<div class="muted">กำลังอ่านไฟล์…</div>';
    try{ rows=await readStaffExcel(file); }catch(err){ $('#pv').innerHTML=`<p class="err">${esc(err.message)}</p>`; return; }
    const draw=()=>{
      const errs=checkStaffRows(rows);
      $('#pv').innerHTML=`<p>พบ <b>${rows.length}</b> คน</p>${staffPreview(rows)}
        ${errs.length?`<div class="note" style="margin-top:12px">${errs.map(esc).join('<br>')}</div>`:''}
        <div class="row end" style="margin-top:14px"><span class="muted small" id="prog"></span><button class="btn" id="mk" ${errs.length?'disabled':''}>สร้างบัญชีทั้งหมด</button></div>`;
      $$('.rolesel').forEach(s=>s.onchange=()=>{ rows[+s.dataset.i].role=s.value; draw(); });
      $('#mk').onclick=async()=>{
        $('#mk').disabled=true;
        try{
          const rep=await S.store.setupStaff(rows, n=>{ $('#prog').textContent='กำลังสร้างบัญชี: '+n; });
          $('#pv').innerHTML=`<div class="note" style="background:#eaf5ef;border-color:#bfe0cc">สร้างบัญชีเรียบร้อย ${rep.length} คน</div>
            <table class="t"><tbody>${rep.map(r=>`<tr><td>${esc(r.name)}</td><td>${esc(r.status)}</td></tr>`).join('')}</tbody></table>
            <div class="row end" style="margin-top:12px"><a class="btn" href="#login">ไปหน้าเข้าสู่ระบบ</a></div>`;
        }catch(err){ $('#prog').textContent=''; toast(err.message); $('#mk').disabled=false; $('#pv').insertAdjacentHTML('afterbegin',`<p class="err">${esc(err.message)}</p>`); }
      };
    };
    draw();
  };
}

async function renderStaff(){
  app.innerHTML='<div class="empty">กำลังโหลด…</div>';
  const list=(await S.store.listStaff()).sort((a,b)=>(a.role===b.role?0:a.role==='admin'?-1:1)||a.name.localeCompare(b.name,'th'));
  app.innerHTML=`<div class="row" style="margin-bottom:12px"><h1>บุคลากร (${list.length} คน)</h1><span class="spacer"></span>
      <label class="btn ghost" style="margin:0;font-weight:400">📂 นำเข้าเพิ่มจาก Excel<input type="file" id="xf" accept=".xlsx,.xls,.csv" hidden></label></div>
    <div class="card" style="overflow-x:auto"><table class="t"><thead><tr><th>ชื่อ (ใช้ล็อกอิน)</th><th>สกุล</th><th>ตำแหน่ง</th><th>สิทธิ์</th><th></th></tr></thead><tbody>
      ${list.map(m=>`<tr><td><span class="muted small">${esc(m.prefix||'')}</span><b>${esc(m.name)}</b></td><td>${esc(m.surname)}</td><td>${esc(m.position)}</td>
        <td><select data-e="${esc(m.email)}" class="rolesel" style="width:auto;padding:4px 6px" ${m.email===S.user.email?'disabled':''}><option value="teacher" ${m.role!=='admin'?'selected':''}>ครู</option><option value="admin" ${m.role==='admin'?'selected':''}>ผู้ดูแลระบบ</option></select></td>
        <td class="num">${m.email===S.user.email?'<span class="muted small">(คุณ)</span>':`<button class="btn ghost danger sm del" data-e="${esc(m.email)}" data-n="${esc(m.name)}">ลบ</button>`}</td></tr>`).join('')}
    </tbody></table></div>
    <form class="card" id="af"><h2>เพิ่มบุคลากร</h2>
      <div class="grid g4"><div><label>คำนำหน้า</label><input type="text" name="prefix"></div><div><label>ชื่อ *</label><input type="text" name="name" required></div>
      <div><label>สกุล</label><input type="text" name="surname"></div><div><label>ตำแหน่ง</label><input type="text" name="position"></div>
      <div><label>เบอร์โทร (รหัสผ่าน) *</label><input type="text" name="phone" inputmode="numeric" required></div>
      <div><label>สิทธิ์</label><select name="role"><option value="teacher">ครู</option><option value="admin">ผู้ดูแลระบบ</option></select></div></div>
      <div class="row end" style="margin-top:12px"><span class="muted small" id="prog"></span><button class="btn">เพิ่ม</button></div>
      ${S.firebase?'<p class="muted small">หมายเหตุ: ถ้าบุคลากรเปลี่ยนเบอร์โทร ต้องลบบัญชีเดิมใน Firebase Console → Authentication ก่อน แล้วจึงเพิ่มใหม่</p>':''}</form>`;
  const doAdd=async rows=>{
    const all=[...list.map(m=>({...m,phone:'000000'})),...rows]; const errs=checkStaffRows(all).filter(x=>!x.includes('ผู้ดูแลระบบ'));
    if(errs.length){ toast(errs[0]); return false; }
    const rep=await S.store.addStaff(rows, n=>{ $('#prog').textContent='กำลังสร้างบัญชี: '+n; });
    toast(`เพิ่มแล้ว ${rep.length} คน`); renderStaff(); return true;
  };
  $('#af').onsubmit=async e=>{ e.preventDefault(); const f=e.target;
    await doAdd([{prefix:f.prefix.value.trim(),name:normName(f.name.value),surname:f.surname.value.trim(),position:f.position.value.trim(),phone:normPhone(f.phone.value),role:f.role.value}]); };
  $('#xf').onchange=async e=>{ const file=e.target.files[0]; if(!file) return;
    try{ const rows=(await readStaffExcel(file)).filter(r=>!list.some(m=>nameKey(m.name)===nameKey(r.name)));
      if(!rows.length){ toast('ไม่มีรายชื่อใหม่ในไฟล์'); return; }
      if(await ask(`พบรายชื่อใหม่ ${rows.length} คน: ${rows.map(r=>r.name).join(', ')} — เพิ่มทั้งหมด?`,'เพิ่ม',false)) await doAdd(rows);
    }catch(err){ toast(err.message); } };
  $$('.rolesel').forEach(s=>s.onchange=async()=>{ await S.store.updateStaff(s.dataset.e,{role:s.value}); toast('เปลี่ยนสิทธิ์แล้ว'); });
  $$('.del').forEach(b=>b.onclick=async()=>{ if(await ask(`ลบ ${b.dataset.n} ออกจากระบบ? (จะเข้าสู่ระบบไม่ได้อีก)`,'ลบ')){ await S.store.removeStaff(b.dataset.e); renderStaff(); } });
}

/* ------------------------------------------------------------------ */
/* ห้องเรียน / ครูประจำกลุ่ม (แอดมินแก้ไขได้)                          */
/* ------------------------------------------------------------------ */
async function renderRooms(){
  let list=(await S.store.getClasses()).map(c=>({...c}));
  app.innerHTML=`<div class="row" style="margin-bottom:6px"><h1>ห้องเรียน / ครูประจำกลุ่ม</h1></div>
    <p class="muted">รายการนี้ใช้เป็นตัวเลือก "เรียนอยู่ห้องเรียน" ในแบบทดสอบและแบบประเมินของผู้เรียน แก้ไขแล้วกดบันทึก</p>
    <div class="card" style="overflow-x:auto"><table class="t"><thead><tr><th class="c" style="width:40px">ที่</th><th>ศูนย์ / ห้องเรียน</th><th>ชื่อครู (ไม่ต้องใส่คำว่า "ครู")</th><th style="width:150px"></th></tr></thead><tbody id="rb"></tbody></table>
      <div class="row" style="margin-top:10px"><button class="btn ghost sm" id="addR">+ เพิ่มห้องเรียน</button><span class="spacer"></span><button class="btn ghost sm" id="defR">คืนค่าเริ่มต้น</button><button class="btn" id="saveR">บันทึก</button></div></div>
    <div class="card"><h2>ตัวอย่างที่ผู้เรียนเห็น</h2><div id="pvR"></div></div>`;
  const draw=()=>{
    $('#rb').innerHTML=list.map((c,i)=>`<tr><td class="c">${i+1}</td>
      <td><input type="text" data-i="${i}" data-k="center" value="${esc(c.center)}" placeholder="เช่น ศกร.ระดับตำบลเวียง"></td>
      <td><input type="text" data-i="${i}" data-k="teacher" value="${esc(c.teacher)}" placeholder="ชื่อ สกุล"></td>
      <td class="num" style="white-space:nowrap"><button class="btn ghost sm mv" data-i="${i}" data-d="-1" ${i?'':'disabled'}>↑</button><button class="btn ghost sm mv" data-i="${i}" data-d="1" ${i<list.length-1?'':'disabled'}>↓</button><button class="btn ghost danger sm rm" data-i="${i}">ลบ</button></td></tr>`).join('');
    $('#pvR').innerHTML=roomSelect(list.filter(c=>c.center.trim()),'');
    $$('#rb input').forEach(inp=>inp.oninput=()=>{ list[+inp.dataset.i][inp.dataset.k]=inp.value; $('#pvR').innerHTML=roomSelect(list.filter(c=>c.center.trim()),''); });
    $$('#rb .rm').forEach(b=>b.onclick=()=>{ list.splice(+b.dataset.i,1); draw(); });
    $$('#rb .mv').forEach(b=>b.onclick=()=>{ const i=+b.dataset.i, j=i+(+b.dataset.d); [list[i],list[j]]=[list[j],list[i]]; draw(); });
  };
  draw();
  $('#addR').onclick=()=>{ list.push({center:'',teacher:''}); draw(); $$('#rb input[data-k=center]').at(-1).focus(); };
  $('#defR').onclick=async()=>{ if(await ask('คืนค่ารายการห้องเรียนเป็นค่าเริ่มต้น? (ยังไม่บันทึกจนกว่าจะกดบันทึก)','คืนค่า',false)){ list=DEFAULT_CLASSES.map(c=>({...c})); draw(); } };
  $('#saveR').onclick=async()=>{
    const clean=list.map(c=>({center:c.center.trim(),teacher:c.teacher.trim().replace(/^ครู\s*/,'')})).filter(c=>c.center);
    await S.store.saveClasses(clean); list=clean; draw(); toast(`บันทึกแล้ว ${clean.length} ห้องเรียน`);
  };
}

/* ------------------------------------------------------------------ */
/* หน้าหลัก: รายการโครงการ                                             */
/* ------------------------------------------------------------------ */
async function renderHome(){
  app.innerHTML='<div class="empty">กำลังโหลด…</div>';
  const ps=await S.store.listProjects();
  const years=[...new Set(ps.map(p=>p.fiscalYear).filter(Boolean))].sort().reverse();
  app.innerHTML=`
    ${!S.firebase?`<div class="note"><b>โหมดทดลอง:</b> ข้อมูลเก็บในเบราว์เซอร์เครื่องนี้เท่านั้น และมีโครงการตัวอย่างให้ลองใช้${inFrame?' · หน้าตัวอย่างนี้ดาวน์โหลดรายงาน Word ได้ แต่ปุ่มพิมพ์จะใช้ได้เมื่อติดตั้งบน GitHub Pages':''}</div>`:''}
    <div class="row" style="margin-bottom:12px"><h1>โครงการ</h1><span class="spacer"></span><a class="btn ghost" href="${PAGE_MODE==='learner'?'#':'index.html'}">หน้าผู้เรียน ↗</a><a class="btn" href="#new">+ สร้างโครงการ</a></div>
    <div class="grid g4" id="sum" style="margin-bottom:16px"></div>
    <div class="card">
      <div class="row" style="margin-bottom:10px">
        <input type="search" id="q" placeholder="ค้นหาชื่อโครงการ / สถานที่ / ผู้รับผิดชอบ" style="flex:1;min-width:220px">
        <select id="fy" style="width:auto"><option value="">ทุกปีงบประมาณ</option>${years.map(y=>`<option ${y===String(currentFY())?'selected':''}>${esc(y)}</option>`).join('')}</select>
        <label style="margin:0;font-weight:400;display:flex;gap:6px;align-items:center"><input type="checkbox" id="mine" ${isAdmin()?'':'checked'}> เฉพาะโครงการของฉัน</label>
      </div>
      <div id="list" style="overflow-x:auto"></div>
    </div>`;
  const draw=()=>{
    const q=$('#q').value.trim().toLowerCase(), fy=$('#fy').value, mine=$('#mine').checked;
    const rows=ps.filter(p=>(!fy||p.fiscalYear===fy) && (!mine||!p.owner||p.owner===S.user.email) && (!q||[p.name,p.location,p.responsible,p.ownerName].join(' ').toLowerCase().includes(q)));
    const bp=rows.reduce((a,p)=>a+num(p.budgetPlanned),0), bs=rows.reduce((a,p)=>a+num(p.budgetSpent),0);
    const openCount=rows.filter(p=>p.surveyOpen!==false||p.test?.preOpen||p.test?.postOpen).length;
    $('#sum').innerHTML=`
      <div class="stat"><div class="k">จำนวนโครงการ</div><div class="v">${rows.length}</div><div class="s">${fy?'ปีงบประมาณ '+esc(fy):'ทุกปีงบประมาณ'}</div></div>
      <div class="stat"><div class="k">งบประมาณที่ได้รับ</div><div class="v">${money(bp)}</div><div class="s">บาท</div></div>
      <div class="stat"><div class="k">ใช้จ่ายจริง</div><div class="v">${money(bs)}</div><div class="s">ร้อยละ ${f2(pct(bs,bp))} ของงบที่ได้รับ</div></div>
      <div class="stat"><div class="k">เปิดให้ผู้เรียนทำแบบทดสอบ/ประเมิน</div><div class="v">${openCount}</div><div class="s">โครงการ</div></div>`;
    $('#list').innerHTML = rows.length ? `<table class="t"><thead><tr><th>โครงการ</th><th>วันที่</th><th>ผู้รับผิดชอบ</th><th class="num">งบประมาณ (บาท)</th><th class="c">เปิดรับ</th></tr></thead><tbody>
      ${rows.map(p=>`<tr class="link" data-id="${esc(p.id)}"><td><b>${esc(p.name)}</b><div class="muted small">${esc(p.location||'')}</div></td>
        <td class="small">${esc(thDate(p.startDate))}</td><td class="small">${esc(p.responsible||p.ownerName||'')}</td>
        <td class="num">${money(p.budgetPlanned)}</td><td class="c small">${[p.test?.preOpen&&'ก่อน',p.test?.postOpen&&'หลัง',p.surveyOpen!==false&&'ประเมิน'].filter(Boolean).map(x=>`<span class="pill">${x}</span>`).join(' ')||'<span class="pill off">ปิด</span>'}</td></tr>`).join('')}
      </tbody></table>` : `<div class="empty">ยังไม่มีโครงการ${q||fy||mine?'ที่ตรงกับเงื่อนไข':''}<br><br><a class="btn" href="#new">+ สร้างโครงการ</a></div>`;
    $$('#list tr.link').forEach(tr=>tr.onclick=()=>location.hash='p-'+tr.dataset.id);
  };
  $('#q').oninput=draw; $('#fy').onchange=draw; $('#mine').onchange=draw;
  if(!years.includes(String(currentFY()))) $('#fy').value='';
  draw();
}

/* ------------------------------------------------------------------ */
/* ฟอร์มสร้าง/แก้ไขโครงการ                                             */
/* ------------------------------------------------------------------ */
async function renderEdit(id){
  let p = id ? await S.store.getProject(id) : null;
  if(id && !p) throw new Error('ไม่พบโครงการนี้');
  let hasResp=false;
  if(p){ hasResp=(await S.store.listResponses(id)).length>0; }
  p = p || { fiscalYear:String(currentFY()), items:parseItems(DEFAULT_ITEMS), surveyOpen:true, targetQualPct:80, budgetSource:'งบดำเนินงาน', actualQty:'',
    responsible:[S.user.prefix,S.user.name,' ',S.user.surname].join('').trim(), responsiblePos:S.user.position||'',
    owner:S.user.email, ownerName:S.user.name, test:{questions:[],preOpen:true,postOpen:false,showScore:true,passPct:60} };
  const v=k=>esc(p[k]??'');
  app.innerHTML=`
    <div class="row" style="margin-bottom:12px"><h1>${id?'แก้ไขโครงการ':'สร้างโครงการใหม่'}</h1><span class="spacer"></span>
      <a class="btn ghost" href="${id?'#p-'+esc(id):'#home'}">ยกเลิก</a></div>
    <form id="pf">
    <fieldset><legend>ข้อมูลโครงการ</legend>
      <label>ชื่อโครงการ *</label><input type="text" name="name" value="${v('name')}" required placeholder="เช่น โครงการอบรมอาชีพการทำไม้กวาดดอกหญ้า">
      <div class="grid g2">
        <div><label>ปีงบประมาณ</label><input type="text" name="fiscalYear" value="${v('fiscalYear')}"></div>
        <div><label>ผู้รับผิดชอบโครงการ</label><input type="text" name="responsible" value="${v('responsible')}" placeholder="ชื่อ-สกุล"></div>
        <div><label>ตำแหน่งผู้รับผิดชอบ</label><input type="text" name="responsiblePos" value="${v('responsiblePos')}"></div>
      </div>
      <label>หลักการและเหตุผล</label><textarea name="rationale" rows="5">${v('rationale')}</textarea>
      <label>วัตถุประสงค์ <span class="hint">(บรรทัดละ 1 ข้อ)</span></label><textarea name="objectives">${v('objectives')}</textarea>
    </fieldset>
    <fieldset><legend>เป้าหมาย</legend>
      <div class="grid g2">
        <div><label>กลุ่มเป้าหมาย</label><input type="text" name="targetGroup" value="${v('targetGroup')}" placeholder="เช่น ประชาชนทั่วไปในตำบล…"></div>
        <div><label>จำนวนเป้าหมาย (คน)</label><input type="number" min="0" name="targetQty" value="${v('targetQty')}"></div>
      </div>
      <label>เป้าหมายเชิงคุณภาพ</label><textarea name="targetQual" rows="2">${v('targetQual')}</textarea>
      <label>เกณฑ์ความสำเร็จเชิงคุณภาพ <span class="hint">(ร้อยละของผู้ตอบที่พึงพอใจระดับมากขึ้นไป)</span></label>
      <input type="number" min="0" max="100" name="targetQualPct" value="${v('targetQualPct')}" style="max-width:160px">
    </fieldset>
    <fieldset><legend>ระยะเวลา สถานที่ และงบประมาณ</legend>
      <div class="grid g2">
        <div><label>วันเริ่ม</label><input type="date" name="startDate" value="${v('startDate')}"></div>
        <div><label>วันสิ้นสุด</label><input type="date" name="endDate" value="${v('endDate')}"></div>
      </div>
      <label>สถานที่ดำเนินการ</label><input type="text" name="location" value="${v('location')}">
      <div class="grid g2">
        <div><label>แหล่งงบประมาณ</label><input type="text" name="budgetSource" value="${v('budgetSource')}"></div>
        <div><label>งบประมาณที่ได้รับ (บาท)</label><input type="number" step="0.01" min="0" name="budgetPlanned" value="${v('budgetPlanned')}"></div>
        <div><label>ใช้จ่ายจริง (บาท)</label><input type="number" step="0.01" min="0" name="budgetSpent" value="${v('budgetSpent')}"></div>
      </div>
    </fieldset>
    <fieldset><legend>การดำเนินงานและผลที่ได้ <span class="muted small">(กรอกหลังจัดกิจกรรมได้)</span></legend>
      <label>ขั้นตอน/วิธีการดำเนินงาน <span class="hint">(บรรทัดละ 1 ขั้นตอน)</span></label><textarea name="method">${v('method')}</textarea>
      <label>จำนวนผู้เข้าร่วมจริง (คน) <span class="hint">— เว้นว่างไว้เพื่อใช้จำนวนผู้ตอบแบบประเมิน</span></label>
      <input type="number" min="0" name="actualQty" value="${v('actualQty')}" style="max-width:160px">
      <label>ผลการดำเนินงานเชิงคุณภาพ <span class="hint">(สิ่งที่เกิดขึ้นจริง เช่น ผลงาน การนำไปใช้)</span></label><textarea name="results">${v('results')}</textarea>
      <label>ปัญหาและอุปสรรค <span class="hint">(บรรทัดละ 1 ข้อ)</span></label><textarea name="problems" rows="3">${v('problems')}</textarea>
      <label>ข้อเสนอแนะ <span class="hint">(บรรทัดละ 1 ข้อ)</span></label><textarea name="suggestions" rows="3">${v('suggestions')}</textarea>
    </fieldset>
    <fieldset><legend>แบบประเมินความพึงพอใจ (มาตราส่วน 5 ระดับ)</legend>
      ${hasResp?'<div class="note">โครงการนี้มีผู้ตอบแบบประเมินแล้ว การเพิ่ม ลบ หรือสลับลำดับข้อคำถามจะทำให้คะแนนเดิมไม่ตรงข้อ ควรแก้แค่ถ้อยคำเท่านั้น</div>':''}
      <label>ข้อคำถาม <span class="hint">(บรรทัดละ 1 ข้อ · ขึ้นต้นด้วย # เพื่อตั้งชื่อด้าน)</span></label>
      <textarea name="items" rows="14" style="font-size:15px">${esc(itemsToText(p.items))}</textarea>
      <div class="row" style="margin-top:6px"><button type="button" class="btn ghost sm" id="resetItems">ใช้ชุดคำถามมาตรฐาน</button></div>
      <label style="display:flex;gap:8px;align-items:center;font-weight:400"><input type="checkbox" name="surveyOpen" ${p.surveyOpen!==false?'checked':''}> เปิดรับแบบประเมิน</label>
    </fieldset>
    <p class="muted small">แบบทดสอบก่อนเรียน-หลังเรียน ตั้งค่าได้ที่หน้าโครงการหลังบันทึก</p>
    <div class="row end"><a class="btn ghost" href="${id?'#p-'+esc(id):'#home'}">ยกเลิก</a><button class="btn" id="saveBtn">บันทึกโครงการ</button></div>
    </form>`;
  $('#resetItems').onclick=()=>{ $('[name=items]').value=DEFAULT_ITEMS; };
  $('#pf').onsubmit=async e=>{
    e.preventDefault(); const f=e.target; const fd=new FormData(f);
    const data={...p}; for(const [k,val] of fd.entries()) if(k!=='items'&&k!=='surveyOpen') data[k]=String(val).trim();
    data.items=parseItems(fd.get('items')); data.surveyOpen=f.surveyOpen.checked;
    if(!data.items.some(i=>i.t==='q')){ toast('ต้องมีข้อคำถามอย่างน้อย 1 ข้อ'); return; }
    $('#saveBtn').disabled=true;
    try{ const nid=await S.store.saveProject(data); toast('บันทึกแล้ว'); location.hash='p-'+nid; }
    catch(err){ toast('บันทึกไม่สำเร็จ: '+err.message); $('#saveBtn').disabled=false; }
  };
}

/* ------------------------------------------------------------------ */
/* ตัวแก้ไขแบบทดสอบก่อน-หลังเรียน                                      */
/* ------------------------------------------------------------------ */
async function renderTestEditor(id){
  const p=await S.store.getProject(id); if(!p) throw new Error('ไม่พบโครงการนี้');
  const t=p.test||{questions:[],preOpen:true,postOpen:false,showScore:true,passPct:60};
  const hasRows=(await S.store.listTests(id)).length>0;
  app.innerHTML=`<div class="row" style="margin-bottom:6px"><a href="#p-${esc(id)}" class="backbtn">← กลับไปหน้าโครงการ</a></div>
    <h1>แบบทดสอบก่อนเรียน-หลังเรียน</h1><p class="muted">ใส่ข้อสอบชุดเดียว ใช้ทั้งก่อนเรียนและหลังเรียน และเลือกให้ระบบสลับลำดับข้อในหลังเรียนได้ (ตั้งค่าด้านล่างช่องข้อสอบ)</p>
    ${hasRows?'<div class="note">มีผู้เรียนทำแบบทดสอบแล้ว ถ้าเพิ่ม ลบ หรือย้ายข้อในช่องข้อสอบ คะแนนเดิมจะคำนวณผิด ควรแก้เฉพาะถ้อยคำหรือเฉลยเท่านั้น (การเปลี่ยนตัวเลือก "สลับลำดับข้อ" ทำได้ ไม่กระทบคะแนน)</div>':''}
    <div class="grid g2" style="align-items:start">
      <div class="card"><h2>พิมพ์หรือวางข้อสอบ</h2>
        <p class="muted small" style="margin-top:-6px">บรรทัดคำถาม แล้วตามด้วยตัวเลือก ก. ข. ค. ง. บรรทัดละตัวเลือก <b>ใส่ * หน้าตัวเลือกที่ถูก</b> คัดลอกจาก Word มาวางได้</p>
        <textarea id="tx" rows="22" style="font-size:15px">${esc(t.questions?.length?testToText(t.questions):'')}</textarea>
        <div class="row" style="margin-top:6px"><button type="button" class="btn ghost sm" id="sample">ใส่ตัวอย่างรูปแบบ</button></div>
        <label>เกณฑ์ผ่าน (ร้อยละ)</label><input type="number" id="cut" min="0" max="100" value="${esc(t.passPct??60)}" style="max-width:140px">
        <label style="display:flex;gap:8px;align-items:center;font-weight:400"><input type="checkbox" id="pre" ${t.preOpen?'checked':''}> เปิดรับแบบทดสอบก่อนเรียน</label>
        <label style="display:flex;gap:8px;align-items:center;font-weight:400;margin-top:4px"><input type="checkbox" id="post" ${t.postOpen?'checked':''}> เปิดรับแบบทดสอบหลังเรียน</label>
        <label style="display:flex;gap:8px;align-items:center;font-weight:400;margin-top:4px"><input type="checkbox" id="show" ${t.showScore!==false?'checked':''}> แสดงคะแนนให้ผู้เรียนเห็นทันทีหลังส่ง</label>
        <label>การสลับลำดับข้อ</label>
        <select id="shf">${Object.entries(SHUFFLE_MODES).map(([k,v])=>`<option value="${k}" ${(t.shuffle||'none')===k?'selected':''}>${v}</option>`).join('')}</select>
        <label style="display:flex;gap:8px;align-items:center;font-weight:400;margin-top:6px"><input type="checkbox" id="shfo" ${t.shuffleOpt?'checked':''}> สลับตัวเลือก ก ข ค ง ด้วย</label>
        <p class="muted small">ระบบตรวจคำตอบตามข้อเดิมเสมอ คะแนนก่อนและหลังเรียนจึงเปรียบเทียบกันได้ถูกต้อง แม้ลำดับที่ผู้เรียนเห็นจะต่างกัน</p>
        <p class="muted small">ถ้าแสดงคะแนนทันที เฉลยจะถูกส่งไปที่เครื่องผู้เรียนด้วย ผู้ที่ชำนาญคอมพิวเตอร์อาจเปิดดูเฉลยได้</p>
        <div class="row end" style="margin-top:10px"><a class="btn ghost" href="#p-${esc(id)}">ยกเลิก</a><button class="btn" id="sv">บันทึกแบบทดสอบ</button></div>
      </div>
      <div class="card"><h2>ตัวอย่างที่ผู้เรียนจะเห็น <span class="muted small" id="cnt"></span></h2>
        <div class="row" style="margin-bottom:10px"><div class="tabs"><button type="button" class="btn sm" id="tbPre">ก่อนเรียน</button><button type="button" class="btn ghost sm" id="tbPost">หลังเรียน</button></div>
          <span class="spacer"></span><button type="button" class="btn ghost sm" id="cpTest">คัดลอกข้อสอบไปพิมพ์ (ลำดับนี้)</button></div>
        <div class="muted small" id="shNote" style="margin-bottom:8px"></div><div id="err"></div><div id="prev" class="sv"></div></div>
    </div>`;
  let view='pre', layout=null;
  const draw=()=>{
    const {questions,errors}=parseTest($('#tx').value);
    $('#cnt').textContent=`(${questions.length} ข้อ)`;
    $('#err').innerHTML=errors.length?`<div class="note">${errors.slice(0,8).map(esc).join('<br>')}${errors.length>8?'<br>…':''}</div>`:'';
    const mode=$('#shf').value;
    layout=testLayout(id, view, {questions, shuffle:mode, shuffleOpt:$('#shfo').checked}, true);
    $('#tbPre').className='btn sm'+(view==='pre'?'':' ghost'); $('#tbPost').className='btn sm'+(view==='post'?'':' ghost');
    $('#shNote').textContent = mode==='random' ? 'โหมดสุ่ม: ผู้เรียนแต่ละคนจะเห็นลำดับต่างจากนี้ (ตัวอย่างนี้คือลำดับเดิม ใช้สำหรับพิมพ์และบันทึกจากกระดาษ)'
      : mode==='post' && view==='post' ? 'ลำดับหลังเรียนที่สลับแล้ว ทุกคนเห็นเหมือนกัน (หน้าบันทึกจากกระดาษของครูก็ใช้ลำดับนี้)' : '';
    $('#prev').innerHTML=layout.qOrder.map((i,k)=>{ const q=questions[i];
      return `<div class="q"><div class="qt">${k+1}. ${esc(q.q)} ${i!==k?`<span class="muted small">(ข้อเดิม ${i+1})</span>`:''}</div>${layout.optOrders[i].map((j,m)=>`<div class="${j===q.ans?'opt-ok':''}">${OPT_LABELS[m]||m+1}. ${esc(q.opts[j])}${j===q.ans?' ✓':''}</div>`).join('')}</div>`; }).join('')||'<div class="empty">ยังไม่มีข้อสอบ</div>';
    return {questions,errors};
  };
  $('#tx').oninput=draw; $('#shf').onchange=draw; $('#shfo').onchange=draw; draw();
  $('#tbPre').onclick=()=>{ view='pre'; draw(); }; $('#tbPost').onclick=()=>{ view='post'; draw(); };
  $('#cpTest').onclick=()=>{
    const {questions}=draw(); if(!questions.length) return;
    const body=layout.qOrder.map((i,k)=>`${k+1}. ${questions[i].q}\n`+layout.optOrders[i].map((j,m)=>`    ${OPT_LABELS[m]||m+1}. ${questions[i].opts[j]}`).join('\n')).join('\n\n');
    const key=layout.qOrder.map((i,k)=>`${k+1}.${OPT_LABELS[layout.optOrders[i].indexOf(questions[i].ans)]}`).join('  ');
    copyText(`แบบทดสอบ${view==='pre'?'ก่อนเรียน':'หลังเรียน'}\n\n${body}\n\n----- เฉลย -----\n${key}`);
  };
  $('#sample').onclick=()=>{ $('#tx').value=($('#tx').value.trim()?$('#tx').value.trim()+'\n\n':'')+TEST_SAMPLE; draw(); };
  $('#sv').onclick=async()=>{
    const {questions,errors}=draw();
    if(errors.length){ toast('กรุณาแก้ไขข้อสอบให้ถูกต้องก่อนบันทึก'); return; }
    p.test={questions,preOpen:$('#pre').checked,postOpen:$('#post').checked,showScore:$('#show').checked,passPct:num($('#cut').value)||60,
      shuffle:$('#shf').value, shuffleOpt:$('#shfo').checked};
    await S.store.saveProject(p); toast('บันทึกแบบทดสอบแล้ว'); location.hash='p-'+id;
  };
}

/* ------------------------------------------------------------------ */
/* หน้ารายละเอียดโครงการ + ผลวิเคราะห์                                 */
/* ------------------------------------------------------------------ */
function distTable(title, rows, total){
  return `<div><h3 style="font-size:16px;margin:0 0 6px">${esc(title)}</h3><table class="t"><thead><tr><th>${esc(title)}</th><th class="num">จำนวน</th><th class="num">ร้อยละ</th></tr></thead><tbody>
    ${rows.map(r=>`<tr><td>${esc(r.label)}</td><td class="num">${r.n}</td><td class="num">${f2(r.pct)}</td></tr>`).join('')}
    <tr class="tot"><td>รวม</td><td class="num">${total}</td><td class="num">${total?'100.00':'-'}</td></tr></tbody></table></div>`;
}
async function renderProject(id){
  app.innerHTML='<div class="empty">กำลังโหลด…</div>';
  const p=await S.store.getProject(id); if(!p) throw new Error('ไม่พบโครงการนี้');
  const [resps, photos, trows]=await Promise.all([S.store.listResponses(id), S.store.listPhotos(id), S.store.listTests(id)]);
  const st=computeStats(p,resps), ts=testStats(p,trows); const t=num(p.targetQty);
  const joinUrl=learnerUrl(id,'j');
  const summary=autoSummary(p,st,ts);
  const T=p.test||{}; const nQ=(T.questions||[]).length;
  const tog=(k,on,label)=>`<button class="btn ${on?'':'ghost'} sm tog" data-k="${k}">${on?'● ':'○ '}${label}: ${on?'เปิด':'ปิด'}</button>`;
  app.innerHTML=`
    <div class="row" style="margin-bottom:6px"><a href="#home" class="backbtn">← โครงการทั้งหมด</a></div>
    <div class="row" style="margin-bottom:14px;align-items:flex-start">
      <div style="flex:1;min-width:260px"><h1>${esc(p.name)}</h1>
        <div class="muted">${esc(dateRange(p))}${p.location?' · '+esc(p.location):''}</div>
        <div class="muted small">ปีงบประมาณ ${esc(p.fiscalYear||'-')} · ผู้รับผิดชอบ ${esc(p.responsible||'-')}</div></div>
      <div class="row"><a class="btn" href="#r-${esc(id)}">📄 รายงานสรุปผล</a><a class="btn ghost" href="#edit-${esc(id)}">แก้ไข</a><button class="btn ghost danger" id="delP">ลบ</button></div>
    </div>
    <div class="grid g4" style="margin-bottom:16px">
      <div class="stat"><div class="k">ผู้เข้าร่วมโครงการ</div><div class="v">${st.actual}${t?` <span style="font-size:16px;font-weight:400">/ ${t} คน</span>`:''}</div>
        <div class="s">${t?`ร้อยละ ${f2(pct(st.actual,t))} · ${st.actual>=t?'บรรลุเป้าหมาย':'ยังไม่บรรลุเป้าหมาย'}`:'ยังไม่ได้กำหนดเป้าหมาย'}</div></div>
      <div class="stat"><div class="k">คะแนนเฉลี่ย ก่อน → หลังเรียน</div><div class="v">${ts.has?`${f2(ts.preS.mean)} → ${f2(ts.postS.mean)}`:'–'}</div><div class="s">${ts.has?`เต็ม ${ts.N} · ผ่านเกณฑ์หลังเรียน ร้อยละ ${f2(ts.postS.passP)}`:'ยังไม่มีผลการทดสอบ'}</div></div>
      <div class="stat"><div class="k">ความพึงพอใจโดยรวม</div><div class="v">${st.n?f2(st.overall.mean):'–'}</div><div class="s">${st.n?`ระดับ${st.overall.lvl} · ผู้ตอบ ${st.n} คน`:'ยังไม่มีผู้ตอบ'}</div></div>
      <div class="stat"><div class="k">งบประมาณใช้จริง</div><div class="v">${money(p.budgetSpent)}</div><div class="s">จาก ${money(p.budgetPlanned)} บาท (ร้อยละ ${f2(pct(num(p.budgetSpent),num(p.budgetPlanned)))})</div></div>
    </div>

    <div class="card">
      <h2>QR Code สำหรับผู้เรียน</h2>
      <div class="qr">
        <div id="qrbox"></div>
        <div style="flex:1;min-width:260px">
          <p style="margin-top:0">ผู้เรียนสแกนแล้วจะเข้าหน้าโครงการนี้ เลือกทำแบบทดสอบก่อนเรียน หลังเรียน หรือแบบประเมินความพึงพอใจได้ (เฉพาะส่วนที่เปิดไว้)${!S.firebase?' <span class="muted small">(โหมดทดลอง: ทำได้เฉพาะในเบราว์เซอร์เครื่องนี้)</span>':''}</p>
          <div class="linkbox">${esc(joinUrl)}</div>
          <div class="row" style="margin-top:8px"><button class="btn ghost sm" id="cpUrl">คัดลอกลิงก์</button><a class="btn ghost sm" href="${PAGE_MODE==='learner'?'#j-'+esc(id):esc(joinUrl)}" ${PAGE_MODE==='learner'?'':'target="_blank"'}>เปิดหน้าผู้เรียน ↗</a></div>
          <div class="row" style="margin-top:12px">${tog('pre',!!T.preOpen&&nQ>0,'ก่อนเรียน')}${tog('post',!!T.postOpen&&nQ>0,'หลังเรียน')}${tog('sat',p.surveyOpen!==false,'ประเมินความพึงพอใจ')}</div>
        </div>
      </div>
    </div>

    <div class="card">
      <div class="row"><h2 style="margin:0">สรุปผลอัตโนมัติ</h2><span class="spacer"></span><button class="btn ghost sm" id="cpSum">คัดลอกข้อความ</button></div>
      <p style="margin-bottom:0;text-indent:2em">${esc(summary)}</p>
    </div>

    <div class="card" style="overflow-x:auto">
      <div class="row"><h2 style="margin:0">แบบทดสอบก่อนเรียน-หลังเรียน ${nQ?`<span class="muted small">(${nQ} ข้อ · เกณฑ์ผ่านร้อยละ ${esc(T.passPct??60)})</span>`:''}</h2><span class="spacer"></span>
        <a class="btn ${nQ?'ghost':''} sm" href="#test-${esc(id)}">${nQ?'แก้ไขข้อสอบ':'+ สร้างแบบทดสอบ'}</a>
        ${nQ?`<a class="btn ghost sm" href="#pre-${esc(id)}">บันทึกก่อนเรียน (กระดาษ)</a><a class="btn ghost sm" href="#post-${esc(id)}">บันทึกหลังเรียน (กระดาษ)</a>`:''}</div>
      ${!nQ?'<div class="empty">ยังไม่ได้สร้างแบบทดสอบ</div>':!ts.has?'<div class="empty">ยังไม่มีผู้เรียนทำแบบทดสอบ</div>':`
      <table class="t" style="margin-top:12px"><thead><tr><th>การทดสอบ</th><th class="num">ผู้ทำ (คน)</th><th class="num">x̄</th><th class="num">S.D.</th><th class="num">ร้อยละ</th><th class="num">ผ่านเกณฑ์</th></tr></thead><tbody>
        <tr><td>ก่อนเรียน</td><td class="num">${ts.pre.length}</td><td class="num">${f2(ts.preS.mean)}</td><td class="num">${f2(ts.preS.sd)}</td><td class="num">${f2(ts.preS.pm)}</td><td class="num">${ts.preS.pass} (${f2(ts.preS.passP)}%)</td></tr>
        <tr><td>หลังเรียน</td><td class="num">${ts.post.length}</td><td class="num">${f2(ts.postS.mean)}</td><td class="num">${f2(ts.postS.sd)}</td><td class="num">${f2(ts.postS.pm)}</td><td class="num">${ts.postS.pass} (${f2(ts.postS.passP)}%)</td></tr>
        ${ts.pairs.length?`<tr class="tot"><td>ผลต่าง (ทำครบ 2 ครั้ง ${ts.pairs.length} คน)</td><td></td><td class="num">${ts.diff.mean>=0?'+':''}${f2(ts.diff.mean)}</td><td></td><td class="num">${ts.diff.mean>=0?'+':''}${f2(pct(ts.diff.mean,ts.N))}</td><td class="num">คะแนนเพิ่ม ${ts.improved} คน</td></tr>`:''}
      </tbody></table>
      <details style="margin-top:12px"><summary style="cursor:pointer;font-weight:600">คะแนนรายบุคคล (${ts.learners.length} คน)</summary>
        <table class="t"><thead><tr><th class="c">ที่</th><th>รหัส</th><th>ชื่อ-สกุล</th><th>ห้องเรียน</th><th class="num">ก่อนเรียน</th><th class="num">หลังเรียน</th><th class="num">ผลต่าง</th><th>ผล</th></tr></thead><tbody>
        ${ts.learners.map((l,i)=>{ const d=l.pre!=null&&l.post!=null?l.post-l.pre:null; const pass=l.post!=null&&pct(l.post,ts.N)>=ts.cut;
          return `<tr><td class="c">${i+1}</td><td class="small">${esc(l.sid)}</td><td>${esc(l.name)}</td><td class="small">${esc(l.room)}</td><td class="num">${l.pre??'-'}</td><td class="num">${l.post??'-'}</td><td class="num">${d==null?'-':(d>0?'+':'')+d}</td><td>${l.post==null?'<span class="muted">ยังไม่ทำหลังเรียน</span>':pass?'<span class="pill">ผ่าน</span>':'<span class="pill off">ไม่ผ่าน</span>'}</td></tr>`; }).join('')}
        </tbody></table></details>
      ${ts.rooms.length>1?`<details style="margin-top:8px"><summary style="cursor:pointer;font-weight:600">สรุปรายห้องเรียน (${ts.rooms.length})</summary>
        <table class="t"><thead><tr><th>ห้องเรียน</th><th class="num">ก่อนเรียน (คน)</th><th class="num">x̄ ก่อน</th><th class="num">หลังเรียน (คน)</th><th class="num">x̄ หลัง</th></tr></thead><tbody>
        ${ts.rooms.map(r=>`<tr><td>${esc(r.room)}</td><td class="num">${r.preN}</td><td class="num">${r.preN?f2(r.preM):'-'}</td><td class="num">${r.postN}</td><td class="num">${r.postN?f2(r.postM):'-'}</td></tr>`).join('')}</tbody></table></details>`:''}
      <details style="margin-top:8px"><summary style="cursor:pointer;font-weight:600">ร้อยละที่ตอบถูกรายข้อ</summary>
        <table class="t"><thead><tr><th>ข้อ</th><th class="num">ก่อนเรียน</th><th class="num">หลังเรียน</th><th style="width:25%"></th></tr></thead><tbody>
        ${ts.items.map((it,i)=>`<tr><td>${i+1}. ${esc(it.q)}</td><td class="num">${f2(it.pre)}</td><td class="num">${f2(it.post)}</td><td><div class="bar" style="margin-bottom:3px"><i style="width:${it.pre}%;background:#9bb8ad"></i></div><div class="bar"><i style="width:${it.post}%"></i></div></td></tr>`).join('')}
        </tbody></table></details>
      <details style="margin-top:8px"><summary style="cursor:pointer;font-weight:600">รายการส่งทั้งหมด (${trows.length})</summary>
        <table class="t"><tbody>${trows.map(r=>`<tr><td>${r.kind==='pre'?'ก่อนเรียน':'หลังเรียน'}</td><td class="small">${esc(r.sid||'')}</td><td>${esc(r.name)}</td><td class="small muted">${new Date(r.at).toLocaleString('th-TH')}</td><td class="num"><button class="btn ghost danger sm delt" data-id="${esc(r.id)}">ลบ</button></td></tr>`).join('')}</tbody></table></details>`}
    </div>

    <div class="card" style="overflow-x:auto">
      <div class="row"><h2 style="margin:0">ผลการประเมินความพึงพอใจ</h2><span class="spacer"></span><a class="btn ghost sm" href="#sat-${esc(id)}">บันทึกจากแบบประเมินกระดาษ</a></div>
      ${st.n?`<table class="t" style="margin-top:12px"><thead><tr><th>รายการประเมิน</th><th class="num">x̄</th><th class="num">S.D.</th><th class="num">ร้อยละ</th><th>ระดับ</th><th style="width:20%"></th></tr></thead><tbody>
      ${(()=>{ let k=0; return st.rows.map(r=>r.t==='h'
        ? `<tr class="sec"><td>${esc(r.text)}</td><td class="num">${r.n?f2(r.mean):'-'}</td><td class="num">${r.n?f2(r.sd):'-'}</td><td class="num">${r.n?f2(r.mean*20):'-'}</td><td>${r.n?r.lvl:''}</td><td></td></tr>`
        : `<tr><td>${++k}. ${esc(r.text)}</td><td class="num">${f2(r.mean)}</td><td class="num">${f2(r.sd)}</td><td class="num">${f2(r.mean*20)}</td><td>${r.lvl}</td><td><div class="bar"><i style="width:${r.mean/5*100}%"></i></div></td></tr>`).join(''); })()}
      <tr class="tot"><td>ภาพรวม</td><td class="num">${f2(st.overall.mean)}</td><td class="num">${f2(st.overall.sd)}</td><td class="num">${f2(st.overall.mean*20)}</td><td>${st.overall.lvl}</td><td><div class="bar"><i style="width:${st.overall.mean/5*100}%"></i></div></td></tr>
      </tbody></table>
      <p class="muted small">ร้อยละ = ค่าเฉลี่ย × 20 (เทียบกับคะแนนเต็ม 5) · เกณฑ์: 4.51–5.00 มากที่สุด · 3.51–4.50 มาก · 2.51–3.50 ปานกลาง · 1.51–2.50 น้อย · 1.00–1.50 น้อยที่สุด</p>`
      :'<div class="empty">ยังไม่มีผู้ตอบแบบประเมิน</div>'}
    </div>

    ${st.n?`<div class="card"><h2>ข้อมูลทั่วไปของผู้ตอบแบบประเมิน</h2><div class="grid g2" style="grid-template-columns:repeat(auto-fit,minmax(240px,1fr))">
      ${distTable('เพศ',st.gender,st.n)}${distTable('อายุ',st.age,st.n)}${distTable('ระดับชั้น',st.level,st.n)}${distTable('ห้องเรียน',st.room,st.n)}</div></div>`:''}

    ${st.comments.length?`<div class="card"><h2>ข้อเสนอแนะจากผู้ตอบ (${st.comments.length})</h2><ul style="margin:0">${st.comments.map(c=>`<li>${esc(c)}</li>`).join('')}</ul></div>`:''}

    <div class="card">
      <div class="row"><h2 style="margin:0">ภาพกิจกรรม (${photos.length})</h2><span class="spacer"></span>
        <label class="btn ghost sm" style="margin:0;font-weight:400">+ เพิ่มภาพ<input type="file" accept="image/*" multiple id="phIn" hidden></label></div>
      <p class="muted small">ระบบย่อขนาดภาพให้อัตโนมัติ และจะนำภาพไปใส่ในภาคผนวกของรายงาน แนะนำโครงการละ 4–8 ภาพ</p>
      <div class="photos" id="phs">${photos.map(ph=>`<figure data-id="${esc(ph.id)}"><img src="${ph.data}" alt="">
        <figcaption><input type="text" value="${esc(ph.caption||'')}" placeholder="คำบรรยายภาพ" class="cap">
        <div class="row end" style="margin-top:4px"><button class="btn ghost danger sm delph">ลบภาพ</button></div></figcaption></figure>`).join('')}</div>
    </div>

    <div class="card" style="overflow-x:auto">
      <h2>รายชื่อผู้ตอบแบบประเมินความพึงพอใจ (${resps.length})</h2>
      ${resps.length?`<table class="t"><thead><tr><th class="c">ที่</th><th>ชื่อ-สกุล</th><th>เพศ</th><th>อายุ</th><th>ระดับชั้น</th><th>ห้องเรียน</th><th class="num">ค่าเฉลี่ย</th><th></th></tr></thead><tbody>
      ${resps.map((r,i)=>`<tr><td class="c">${i+1}</td><td>${esc(r.name)}</td><td>${esc(r.gender)}</td><td>${esc(r.age)}</td><td class="small">${esc(r.level||'')}</td><td class="small">${esc(r.room||r.occ||'')}</td>
        <td class="num">${f2(msd((r.scores||[]).filter(x=>x>=1&&x<=5)).mean)}</td><td class="num"><button class="btn ghost danger sm delr" data-id="${esc(r.id)}">ลบ</button></td></tr>`).join('')}</tbody></table>`
      :'<div class="empty">ยังไม่มีข้อมูล</div>'}
    </div>`;

  try{ if(window.QRCode) new QRCode($('#qrbox'),{text:joinUrl,width:180,height:180,correctLevel:QRCode.CorrectLevel.M}); else $('#qrbox').textContent='โหลด QR ไม่ได้'; }catch(e){}
  $('#cpUrl').onclick=()=>copyText(joinUrl);
  $('#cpSum').onclick=()=>copyText(summary);
  $$('.tog').forEach(b=>b.onclick=async()=>{
    const k=b.dataset.k;
    if(k==='sat') p.surveyOpen=p.surveyOpen===false;
    else { if(!nQ){ toast('ยังไม่มีข้อสอบ — กด "สร้างแบบทดสอบ" ก่อน'); return; } p.test={...T,[k+'Open']:!T[k+'Open']}; }
    await S.store.saveProject(p); renderProject(id);
  });
  $('#delP').onclick=async()=>{ if(await ask(`ลบ "${p.name}" รวมทั้งแบบทดสอบ แบบประเมิน และภาพทั้งหมด? ลบแล้วกู้คืนไม่ได้`,'ลบโครงการ')){ await S.store.deleteProject(id); toast('ลบแล้ว'); location.hash='home'; } };
  $$('.delr').forEach(b=>b.onclick=async()=>{ if(await ask('ลบคำตอบของผู้ตอบรายนี้?','ลบ')){ await S.store.deleteResponse(id,b.dataset.id); renderProject(id); } });
  $$('.delt').forEach(b=>b.onclick=async()=>{ if(await ask('ลบผลการทดสอบรายการนี้?','ลบ')){ await S.store.deleteTest(id,b.dataset.id); renderProject(id); } });
  $('#phIn').onchange=async e=>{
    const files=[...e.target.files]; if(!files.length) return; toast('กำลังเพิ่มภาพ…');
    for(const f of files){ try{ const im=await compressImage(f, S.firebase?1200:900); await S.store.addPhoto(id,{...im,caption:''}); }catch(err){ toast(err.message); } }
    renderProject(id);
  };
  $$('#phs figure').forEach(fg=>{
    $('.cap',fg).onchange=e=>S.store.updatePhoto(id,fg.dataset.id,{caption:e.target.value.trim()}).then(()=>toast('บันทึกคำบรรยายแล้ว'));
    $('.delph',fg).onclick=async()=>{ if(await ask('ลบภาพนี้?','ลบ')){ await S.store.deletePhoto(id,fg.dataset.id); renderProject(id); } };
  });
}

/* ------------------------------------------------------------------ */
/* ตั้งค่าหน่วยงาน                                                     */
/* ------------------------------------------------------------------ */
async function renderSettings(){
  const s=S.settings;
  app.innerHTML=`<h1>ตั้งค่าหน่วยงาน</h1><p class="muted">ข้อมูลนี้ใช้บนหน้าปก คำนำ และส่วนลงนามของรายงาน</p>
    <form id="stf" class="card">
      <label>ชื่อหน่วยงาน</label><input type="text" name="org" value="${esc(s.org)}">
      <label>หน่วยงานต้นสังกัด</label><input type="text" name="parent" value="${esc(s.parent)}">
      <div class="grid g2"><div><label>ชื่อผู้อำนวยการ <span class="hint">(ผู้เห็นชอบรายงาน)</span></label><input type="text" name="director" value="${esc(s.director)}"></div>
      <div><label>ตำแหน่ง</label><input type="text" name="directorPos" value="${esc(s.directorPos)}"></div></div>
      <div class="row end" style="margin-top:16px"><button class="btn">บันทึก</button></div>
    </form>
    ${!S.firebase?`<div class="card"><h2>ข้อมูลโหมดทดลอง</h2><p class="muted">ลบข้อมูลทั้งหมดในเบราว์เซอร์เครื่องนี้ (รวมรายชื่อบุคลากร) แล้วเริ่มใหม่ด้วยโครงการตัวอย่าง</p><button class="btn ghost danger" id="rst">ล้างข้อมูลทดลอง</button></div>`:''}`;
  $('#stf').onsubmit=async e=>{ e.preventDefault(); const fd=new FormData(e.target); const ns={}; for(const [k,v] of fd.entries()) ns[k]=String(v).trim();
    await S.store.saveSettings(ns); S.settings={...DEFAULT_SETTINGS,...ns}; drawHeader(); toast('บันทึกแล้ว'); };
  if($('#rst')) $('#rst').onclick=async()=>{ if(await ask('ล้างข้อมูลทดลองทั้งหมด?','ล้างข้อมูล')){ S.store.reset(); await S.store.init(); S.user=null; S.settings=await S.store.getSettings(); drawHeader(); location.hash='setup'; route(); } };
}

/* ------------------------------------------------------------------ */
/* รายงานสรุปผล: สร้างโครงร่างครั้งเดียว ใช้ได้ทั้งพรีวิวและ Word       */
/* ------------------------------------------------------------------ */
function buildReport(p, st, ts, photos, resps, set){
  const B=[]; const add=(k,o={})=>B.push({k,...o});
  let tn=0; const cap=t=>add('cap',{text:`ตารางที่ ${++tn} ${t}`});
  const fy=p.fiscalYear||'';
  const t=num(p.targetQty); const tq=num(p.targetQualPct)||80;
  const textOr=(s,d='-')=>String(s||'').trim()||d;
  let sec=0; const H2=txt=>add('h2',{text:`${++sec}. ${txt}`});

  add('space',{n:1});
  add('logo',{});
  add('space',{n:2});
  add('title',{text:'รายงานผลการดำเนินโครงการ',size:26});
  add('title',{text:p.name||'',size:22});
  if(fy) add('title',{text:`ประจำปีงบประมาณ ${fy}`,size:20});
  add('space',{n:9});
  add('title',{text:set.org,size:20});
  if(set.parent) add('title',{text:set.parent,size:18});

  add('h1',{text:'คำนำ',pb:true});
  add('p',{text:`รายงานผลการดำเนิน${p.name} ฉบับนี้ จัดทำขึ้นเพื่อสรุปผลการดำเนินงานตามวัตถุประสงค์และเป้าหมายของโครงการ${p.startDate?' ซึ่งดำเนินการเมื่อ'+dateRange(p):''}${p.location?' ณ '+p.location:''} โดยนำเสนอหลักการและเหตุผล วิธีการดำเนินงาน ผลการดำเนินงาน${ts.has?' ผลการทดสอบความรู้ก่อนเรียนและหลังเรียน':''} ผลการประเมินความพึงพอใจของผู้เข้าร่วมโครงการ ตลอดจนปัญหาอุปสรรคและข้อเสนอแนะ`});
  add('p',{text:`${set.org} หวังเป็นอย่างยิ่งว่ารายงานฉบับนี้จะเป็นประโยชน์ต่อการวางแผนและพัฒนาการจัดกิจกรรมในครั้งต่อไป`});
  add('sign',{lines:[set.org, monthYearTH()]});

  add('h1',{text:'สารบัญ',pb:true});
  add('ol',{items:['คำนำ','บทที่ 1 บทนำ','บทที่ 2 วิธีการดำเนินงาน','บทที่ 3 ผลการดำเนินงาน','บทที่ 4 สรุปผล ปัญหาอุปสรรค และข้อเสนอแนะ','ภาคผนวก'],plain:true});

  add('h1',{text:'บทที่ 1\nบทนำ',pb:true}); sec=0;
  H2('หลักการและเหตุผล'); lines(p.rationale).forEach(l=>add('p',{text:l})); if(!lines(p.rationale).length) add('p',{text:'-'});
  H2('วัตถุประสงค์'); add('ol',{items:lines(p.objectives).length?lines(p.objectives):['-'],prefix:'2.'});
  H2('เป้าหมาย');
  add('li',{text:`3.1 เชิงปริมาณ: ${textOr(p.targetGroup,'กลุ่มเป้าหมาย')} จำนวน ${t||'-'} คน`});
  add('li',{text:`3.2 เชิงคุณภาพ: ${textOr(p.targetQual)}`});
  H2('ระยะเวลาดำเนินการ'); add('p',{text:textOr(dateRange(p))});
  H2('สถานที่ดำเนินการ'); add('p',{text:textOr(p.location)});
  H2('งบประมาณ'); add('p',{text:`${textOr(p.budgetSource,'งบประมาณ')} จำนวน ${money(p.budgetPlanned)} บาท (${bahtText(p.budgetPlanned)})`});
  H2('ผู้รับผิดชอบโครงการ'); add('p',{text:[p.responsible,p.responsiblePos].filter(Boolean).join(' ตำแหน่ง ')||'-'});

  add('h1',{text:'บทที่ 2\nวิธีการดำเนินงาน',pb:true}); sec=0;
  H2('ขั้นตอนการดำเนินงาน'); add('ol',{items:lines(p.method).length?lines(p.method):['-'],prefix:'1.'});
  H2('เครื่องมือที่ใช้ในการประเมิน');
  if(ts.has) add('li',{text:`2.1 แบบทดสอบวัดความรู้ก่อนเรียนและหลังเรียน แบบเลือกตอบ จำนวน ${ts.N} ข้อ (ข้อละ 1 คะแนน) เกณฑ์ผ่านร้อยละ ${ts.cut}`});
  add(ts.has?'li':'p',{text:`${ts.has?'2.2 ':''}แบบประเมินความพึงพอใจของผู้เข้าร่วมโครงการ แบ่งออกเป็น 3 ตอน ได้แก่ ตอนที่ 1 ข้อมูลทั่วไปของผู้ตอบแบบประเมิน ตอนที่ 2 ความพึงพอใจต่อการดำเนินโครงการ เป็นแบบมาตราส่วนประมาณค่า 5 ระดับ และตอนที่ 3 ข้อเสนอแนะเพิ่มเติม`});
  H2('การวิเคราะห์ข้อมูล');
  add('p',{text:'วิเคราะห์ข้อมูลด้วยค่าร้อยละ ค่าเฉลี่ย (x̄) และส่วนเบี่ยงเบนมาตรฐาน (S.D.)'+(ts.has?' เปรียบเทียบคะแนนก่อนเรียนและหลังเรียนด้วยผลต่างของคะแนนเฉลี่ย':'')+' โดยร้อยละของความพึงพอใจคำนวณจากค่าเฉลี่ยเทียบกับคะแนนเต็ม 5 (ค่าเฉลี่ย × 20) และใช้เกณฑ์การแปลความหมายค่าเฉลี่ย ดังนี้'});
  add('table',{head:['ค่าเฉลี่ย','ระดับความพึงพอใจ'],widths:[50,50],align:['c','c'],
    rows:[['4.51 – 5.00','มากที่สุด'],['3.51 – 4.50','มาก'],['2.51 – 3.50','ปานกลาง'],['1.51 – 2.50','น้อย'],['1.00 – 1.50','น้อยที่สุด']]});

  add('h1',{text:'บทที่ 3\nผลการดำเนินงาน',pb:true}); sec=0;
  H2('ผลการดำเนินงานเชิงปริมาณ');
  cap('ผลการดำเนินงานเชิงปริมาณ');
  add('table',{head:['เป้าหมาย (คน)','ดำเนินการจริง (คน)','ร้อยละ','ผลการประเมิน'],widths:[25,25,20,30],align:['c','c','c','c'],
    rows:[[t||'-',st.actual,t?f2(pct(st.actual,t)):'-',t?(st.actual>=t?'บรรลุเป้าหมาย':'ไม่บรรลุเป้าหมาย'):'-']]});
  H2('ผลการดำเนินงานเชิงคุณภาพ');
  lines(p.results).forEach(l=>add('p',{text:l}));
  if(st.n) add('p',{text:`ผู้เข้าร่วมโครงการร้อยละ ${f2(st.pctGood)} มีความพึงพอใจในระดับมากขึ้นไป ${st.pctGood>=tq?'บรรลุ':'ไม่บรรลุ'}เป้าหมายเชิงคุณภาพที่กำหนดไว้ร้อยละ ${tq}`});
  else if(!lines(p.results).length) add('p',{text:'-'});
  if(ts.has){
    H2('ผลการทดสอบความรู้ก่อนเรียนและหลังเรียน');
    cap(`ผลการทดสอบความรู้ก่อนเรียนและหลังเรียน (คะแนนเต็ม ${ts.N} คะแนน)`);
    const rows=[];
    if(ts.pre.length) rows.push(['ก่อนเรียน',ts.pre.length,f2(ts.preS.mean),f2(ts.preS.sd),f2(ts.preS.pm),`${ts.preS.pass} (${f2(ts.preS.passP)})`]);
    if(ts.post.length) rows.push(['หลังเรียน',ts.post.length,f2(ts.postS.mean),f2(ts.postS.sd),f2(ts.postS.pm),`${ts.postS.pass} (${f2(ts.postS.passP)})`]);
    add('table',{head:['การทดสอบ','จำนวน (คน)','x̄','S.D.','ร้อยละ','ผ่านเกณฑ์ คน (ร้อยละ)'],widths:[20,14,12,12,14,28],align:['l','c','c','c','c','c'],rows});
    add('p',{text:`จากตารางที่ ${tn} `+testText(ts).replace(/^ผลการทดสอบความรู้ \(คะแนนเต็ม \d+ คะแนน\) พบว่า /,'พบว่า ')});
    if(ts.learners.length){
      cap('คะแนนการทดสอบก่อนเรียนและหลังเรียนรายบุคคล');
      add('table',{head:['ที่','รหัส','ชื่อ-สกุล','ก่อน','หลัง','ผลต่าง','ผล'],widths:[6,21,37,9,9,9,9],align:['c','c','l','c','c','c','c'],
        rows:ts.learners.map((l,i)=>{ const d=l.pre!=null&&l.post!=null?l.post-l.pre:null;
          return [i+1,l.sid||'-',l.name,l.pre??'-',l.post??'-',d==null?'-':(d>0?'+':'')+d,l.post==null?'-':(pct(l.post,ts.N)>=ts.cut?'ผ่าน':'ไม่ผ่าน')]; })});
      if(ts.rooms.length>1){
        cap('ผลการทดสอบก่อนเรียนและหลังเรียน จำแนกตามห้องเรียน');
        add('table',{head:['ห้องเรียน','ก่อน (คน)','x̄ ก่อน','หลัง (คน)','x̄ หลัง'],widths:[48,13,13,13,13],align:['l','c','c','c','c'],
          rows:ts.rooms.map(r=>[r.room,r.preN,r.preN?f2(r.preM):'-',r.postN,r.postN?f2(r.postM):'-'])});
      }
    }
  }
  H2('ผลการประเมินความพึงพอใจ');
  if(st.n){
    add('p',{text:`ผู้ตอบแบบประเมินทั้งหมด ${st.n} คน วิเคราะห์ผลได้ดังนี้`});
    add('h3',{text:'ตอนที่ 1 ข้อมูลทั่วไปของผู้ตอบแบบประเมิน'});
    for(const [nm,arr] of [['เพศ',st.gender],['อายุ',st.age],['ระดับชั้น',st.level],['ห้องเรียน',st.room]]){
      cap(`จำนวนและร้อยละของผู้ตอบแบบประเมิน จำแนกตาม${nm}`);
      add('table',{head:[nm,'จำนวน (คน)','ร้อยละ'],widths:nm==='ห้องเรียน'?[64,18,18]:[50,25,25],align:['l','c','c'],
        rows:[...arr.map(r=>[r.label,r.n,f2(r.pct)]),['รวม',st.n,'100.00']],bold:[arr.length]});
      const top=[...arr].sort((a,b)=>b.n-a.n)[0];
      if(top) add('p',{text:`จากตารางที่ ${tn} ผู้ตอบแบบประเมินส่วนใหญ่${({'เพศ':'เป็นเพศ','อายุ':'มีอายุ ','ระดับชั้น':'เรียนระดับ','ห้องเรียน':'เรียนอยู่ที่ '})[nm]}${top.label} จำนวน ${top.n} คน คิดเป็นร้อยละ ${f2(top.pct)}`});
    }
    add('h3',{text:'ตอนที่ 2 ความพึงพอใจของผู้เข้าร่วมโครงการ'});
    cap('ค่าเฉลี่ยและส่วนเบี่ยงเบนมาตรฐานของความพึงพอใจของผู้เข้าร่วมโครงการ');
    let k=0; const bold=[]; const rows=[];
    st.rows.forEach(r=>{ if(r.t==='h'){ bold.push(rows.length); rows.push([r.text,r.n?f2(r.mean):'-',r.n?f2(r.sd):'-',r.n?f2(r.mean*20):'-',r.n?r.lvl:'-']); }
      else rows.push([`${++k}. ${r.text}`,f2(r.mean),f2(r.sd),f2(r.mean*20),r.lvl]); });
    bold.push(rows.length); rows.push(['ภาพรวม',f2(st.overall.mean),f2(st.overall.sd),f2(st.overall.mean*20),st.overall.lvl]);
    add('table',{head:['รายการประเมิน','x̄','S.D.','ร้อยละ','ระดับ'],widths:[50,11,11,12,16],align:['l','c','c','c','c'],rows,bold});
    add('p',{text:`จากตารางที่ ${tn} พบว่า ` + satisfactionText(p,st)});
    add('h3',{text:'ตอนที่ 3 ข้อเสนอแนะจากผู้ตอบแบบประเมิน'});
    if(st.comments.length) add('ol',{items:st.comments,prefix:''}); else add('p',{text:'- ไม่มี -'});
  }else add('p',{text:'ยังไม่มีข้อมูลการประเมินความพึงพอใจ'});

  add('h1',{text:'บทที่ 4\nสรุปผล ปัญหาอุปสรรค และข้อเสนอแนะ',pb:true}); sec=0;
  H2('สรุปผลการดำเนินงาน'); add('p',{text:autoSummary(p,st,ts)});
  H2('การใช้งบประมาณ');
  add('p',{text:`ได้รับงบประมาณ ${money(p.budgetPlanned)} บาท ใช้จ่ายจริง ${money(p.budgetSpent)} บาท (${bahtText(p.budgetSpent)}) คิดเป็นร้อยละ ${f2(pct(num(p.budgetSpent),num(p.budgetPlanned)))} ของงบประมาณที่ได้รับ${num(p.budgetPlanned)-num(p.budgetSpent)>0?` คงเหลือ ${money(num(p.budgetPlanned)-num(p.budgetSpent))} บาท`:''}`});
  H2('ปัญหาและอุปสรรค'); add('ol',{items:lines(p.problems).length?lines(p.problems):['-'],prefix:'3.'});
  H2('ข้อเสนอแนะ'); add('ol',{items:lines(p.suggestions).length?lines(p.suggestions):['-'],prefix:'4.'});
  add('sign',{lines:['ลงชื่อ .............................. ผู้รายงาน',`(${p.responsible||'..............................'})`,p.responsiblePos||'']});
  add('sign',{lines:['ลงชื่อ .............................. ผู้เห็นชอบ',`(${set.director||'..............................'})`,set.directorPos||'']});

  add('h1',{text:'ภาคผนวก',pb:true});
  if(photos.length){ add('h2',{text:'ภาพประกอบการดำเนินโครงการ'}); add('imgs',{items:photos}); }
  if(resps.length){
    add('h2',{text:'รายชื่อผู้ตอบแบบประเมินความพึงพอใจ',pb:photos.length>0});
    add('table',{head:['ที่','ชื่อ-สกุล','เพศ','ระดับชั้น','ห้องเรียน'],widths:[7,30,10,20,33],align:['c','l','c','l','l'],
      rows:resps.map((r,i)=>[i+1,r.name,r.gender,r.level||'',r.room||r.occ||''])});
  }
  if(ts.has && ts.items.length){
    add('h2',{text:'ร้อยละของผู้ตอบถูกในแบบทดสอบรายข้อ',pb:photos.length>0||resps.length>0});
    add('table',{head:['ข้อ','คำถาม','ก่อนเรียน','หลังเรียน'],widths:[8,60,16,16],align:['c','l','c','c'],
      rows:ts.items.map((it,i)=>[i+1,it.q,f2(it.pre),f2(it.post)])});
  }
  if(!photos.length && !resps.length && !ts.has) add('p',{text:'-'});
  return B;
}

const olLabel=(pre,i)=>pre?`${pre}${i+1}`:`${i+1}.`;
function reportToHTML(B){
  const o=[];
  for(const b of B){
    switch(b.k){
      case 'space': o.push(`<div style="height:${b.n*1.4}em"></div>`); break;
      case 'logo': o.push(`<div style="text-align:center"><img src="logo.png" alt="" style="width:150px;height:150px"></div>`); break;
      case 'title': o.push(`<div class="rp-title" style="font-size:${b.size}px">${esc(b.text)}</div>`); break;
      case 'h1': if(b.pb) o.push('<div class="rp-pb"></div>'); o.push(`<div class="rp-h1">${esc(b.text).replace(/\n/g,'<br>')}</div>`); break;
      case 'h2': if(b.pb) o.push('<div class="rp-pb"></div>'); o.push(`<div class="rp-h2">${esc(b.text)}</div>`); break;
      case 'h3': o.push(`<div class="rp-h2" style="padding-left:1.5em">${esc(b.text)}</div>`); break;
      case 'cap': o.push(`<div class="rp-cap">${esc(b.text)}</div>`); break;
      case 'p': o.push(`<p class="rp-p">${esc(b.text)}</p>`); break;
      case 'li': o.push(`<p class="rp-li">${esc(b.text)}</p>`); break;
      case 'ol': b.items.forEach((t,i)=>o.push(`<p class="rp-li">${b.plain?'':esc(olLabel(b.prefix,i)+' ')}${esc(t)}</p>`)); break;
      case 'table': o.push(`<table><thead><tr>${b.head.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${b.rows.map((r,ri)=>`<tr class="${(b.bold||[]).includes(ri)?'b':''}">${r.map((c,ci)=>`<td style="text-align:${({l:'left',c:'center',r:'right'})[b.align?.[ci]||'l']}">${esc(c)}</td>`).join('')}</tr>`).join('')}</tbody></table>`); break;
      case 'imgs': o.push(`<div class="rp-imgs">${b.items.map(ph=>`<figure><img src="${ph.data}" alt="">${ph.caption?`<figcaption>${esc(ph.caption)}</figcaption>`:''}</figure>`).join('')}</div>`); break;
      case 'sign': o.push(`<div class="rp-sign">${b.lines.map(l=>`<div>${esc(l)||'&nbsp;'}</div>`).join('')}</div>`); break;
    }
  }
  return o.join('\n');
}

async function reportToDocx(B, filename){
  if(!window.docx) await loadScript(DOCX_URL);
  const D=window.docx, F='TH Sarabun New';
  const R=(text,o={})=>new D.TextRun({text:String(text??''),bold:o.bold,size:o.size||32,font:F});
  const AL={l:D.AlignmentType.LEFT,c:D.AlignmentType.CENTER,r:D.AlignmentType.RIGHT};
  const P=(runs,o={})=>new D.Paragraph({children:runs,alignment:o.align,pageBreakBefore:o.pb,
    indent:o.first?{firstLine:720}:(o.left?{left:o.left}:undefined),spacing:{before:o.before||0,after:o.after||0,line:o.line}});
  const none={style:D.BorderStyle.NONE,size:0,color:'FFFFFF'};
  const thin={style:D.BorderStyle.SINGLE,size:4,color:'666666'};
  const out=[]; const CW=11906-1701-1134;
  for(const b of B){
    switch(b.k){
      case 'space': for(let i=0;i<b.n;i++) out.push(P([R('')])); break;
      case 'logo': try{ const buf=new Uint8Array(await (await fetch('logo.png')).arrayBuffer());
        out.push(P([new D.ImageRun({data:buf,transformation:{width:150,height:150}})],{align:D.AlignmentType.CENTER})); }catch(e){} break;
      case 'title': out.push(P([R(b.text,{bold:true,size:b.size*2})],{align:D.AlignmentType.CENTER,after:80})); break;
      case 'h1': out.push(new D.Paragraph({alignment:D.AlignmentType.CENTER,pageBreakBefore:!!b.pb,spacing:{after:240},
        children:b.text.split('\n').map((l,i)=>new D.TextRun({text:l,bold:true,size:40,font:F,break:i?1:0}))})); break;
      case 'h2': out.push(P([R(b.text,{bold:true})],{before:160,pb:b.pb})); break;
      case 'h3': out.push(P([R(b.text,{bold:true})],{before:120,left:360})); break;
      case 'cap': out.push(P([R(b.text,{bold:true})],{before:120,after:40})); break;
      case 'p': out.push(new D.Paragraph({children:[R(b.text)],alignment:D.AlignmentType.THAI_DISTRIBUTE||D.AlignmentType.JUSTIFIED,indent:{firstLine:720}})); break;
      case 'li': out.push(P([R(b.text)],{left:720})); break;
      case 'ol': b.items.forEach((t,i)=>out.push(P([R(b.plain?t:`${olLabel(b.prefix,i)} ${t}`)],{left:720}))); break;
      case 'table': {
        const cw=b.widths.map(w=>Math.round(CW*w/100));
        const mk=(cells,isHead,isBold)=>new D.TableRow({tableHeader:isHead,cantSplit:true,children:cells.map((c,ci)=>new D.TableCell({
          width:{size:cw[ci],type:D.WidthType.DXA},
          shading:isHead?{fill:'E7EFEC',type:D.ShadingType.CLEAR,color:'auto'}:undefined,
          borders:{top:thin,bottom:thin,left:thin,right:thin},
          margins:{left:80,right:80},
          children:[P([R(c,{bold:isHead||isBold,size:30})],{align:isHead?D.AlignmentType.CENTER:AL[b.align?.[ci]||'l']})]}))});
        out.push(new D.Table({width:{size:CW,type:D.WidthType.DXA},columnWidths:cw,
          rows:[mk(b.head,true,false), ...b.rows.map((r,ri)=>mk(r,false,(b.bold||[]).includes(ri)))]}));
        out.push(P([R('')]));
        break; }
      case 'imgs': {
        const W=290; const cells=b.items.map(ph=>{
          const bin=atob(ph.data.split(',')[1]); const u=new Uint8Array(bin.length); for(let i=0;i<bin.length;i++) u[i]=bin.charCodeAt(i);
          const h=Math.round(W*(ph.h||3)/(ph.w||4));
          return new D.TableCell({width:{size:CW/2,type:D.WidthType.DXA},borders:{top:none,bottom:none,left:none,right:none},
            children:[P([new D.ImageRun({data:u,transformation:{width:W,height:h}})],{align:D.AlignmentType.CENTER,before:80}),
                      P([R(ph.caption||'',{size:28})],{align:D.AlignmentType.CENTER,after:120})]});
        });
        if(cells.length%2) cells.push(new D.TableCell({width:{size:CW/2,type:D.WidthType.DXA},borders:{top:none,bottom:none,left:none,right:none},children:[P([R('')])]}));
        const rows=[]; for(let i=0;i<cells.length;i+=2) rows.push(new D.TableRow({cantSplit:true,children:[cells[i],cells[i+1]]}));
        out.push(new D.Table({width:{size:CW,type:D.WidthType.DXA},columnWidths:[CW/2,CW/2],rows,
          borders:{top:none,bottom:none,left:none,right:none,insideHorizontal:none,insideVertical:none}}));
        break; }
      case 'sign': out.push(P([R('')])); b.lines.forEach(l=>out.push(P([R(l)],{align:D.AlignmentType.CENTER,left:3400}))); break;
    }
  }
  const doc=new D.Document({
    creator:S.settings.org, title:filename,
    styles:{default:{document:{run:{font:F,size:32}}}},
    sections:[{
      properties:{titlePage:true,page:{size:{width:11906,height:16838},margin:{top:1418,bottom:1134,left:1701,right:1134}}},
      headers:{default:new D.Header({children:[new D.Paragraph({alignment:D.AlignmentType.RIGHT,children:[new D.TextRun({children:[D.PageNumber.CURRENT],font:F,size:32})]})]}),
               first:new D.Header({children:[P([])]})},
      children:out }]
  });
  const blob=await D.Packer.toBlob(doc);
  await download(blob, filename+'.docx');
}

async function renderReport(id){
  app.innerHTML='<div class="empty">กำลังสร้างรายงาน…</div>';
  const p=await S.store.getProject(id); if(!p) throw new Error('ไม่พบโครงการนี้');
  const [resps,photos,trows]=await Promise.all([S.store.listResponses(id),S.store.listPhotos(id),S.store.listTests(id)]);
  const st=computeStats(p,resps), ts=testStats(p,trows);
  const B=buildReport(p,st,ts,photos,resps,S.settings);
  const fname=('รายงานผล_'+(p.name||'โครงการ')).replace(/[\\/:*?"<>|]/g,'').slice(0,80);
  app.innerHTML=`<div class="row noprint" style="margin-bottom:14px"><a href="#p-${esc(id)}" class="backbtn">← กลับไปหน้าโครงการ</a><span class="spacer"></span>
      <button class="btn" id="dlW">⬇ ดาวน์โหลด Word (.docx)</button><button class="btn ghost" id="prt">พิมพ์ / บันทึก PDF</button></div>
    <div class="note noprint">ไฟล์ Word ใช้ฟอนต์ <b>TH Sarabun New</b> ขนาด 16 พอยต์ ถ้าเครื่องไม่มีฟอนต์นี้ Word จะแสดงด้วยฟอนต์อื่นแทน ปรับแก้ข้อความใน Word ได้ตามต้องการ</div>
    <div class="paper">${reportToHTML(B)}</div>`;
  $('#dlW').onclick=async e=>{ e.target.disabled=true; e.target.textContent='กำลังสร้างไฟล์…';
    try{ await reportToDocx(B,fname); }catch(err){ console.error(err); toast('สร้างไฟล์ไม่สำเร็จ: '+err.message); }
    e.target.disabled=false; e.target.textContent='⬇ ดาวน์โหลด Word (.docx)'; };
  $('#prt').onclick=()=>{ try{ window.print(); }catch(e){ toast('สั่งพิมพ์ไม่ได้ในหน้านี้'); } };
}

boot();
