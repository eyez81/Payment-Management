'use strict';
const SHEET_ID = '1HuCg5keyzs4BQIa3ZjzEyEVB54j-SwtGarFP5uq48Pk';
const TAB = 'מנויים';
const HEADERS = ['שם השירות','קטגוריה','סכום לתשלום','מטבע','תדירות','תאריך החיוב הבא','תאריך סיום','אמצעי תשלום','מצב','הערות'];
const SCOPE = 'https://www.googleapis.com/auth/spreadsheets';
const $ = id => document.getElementById(id);
const state = {token:null, rows:[], editing:null, busy:false};
const escapeHtml = s => String(s ?? '').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const dateValue = s => /^\d{4}-\d{2}-\d{2}$/.test(s||'') ? new Date(`${s}T12:00:00`) : null;
const dateText = s => {const d=dateValue(s);return d ? new Intl.DateTimeFormat('he-IL',{day:'numeric',month:'long',year:'numeric'}).format(d) : '—'};
const daysTo = s => {const d=dateValue(s);const now=new Date();now.setHours(12,0,0,0);return d?Math.round((d-now)/86400000):null};
const currencies = {'שקל':'ILS','דולר':'USD','אירו':'EUR'};
const money = (amount,currency) => new Intl.NumberFormat('he-IL',{style:'currency',currency:currencies[currency]||'ILS',maximumFractionDigits:2}).format(Number(amount)||0);
function notice(text,error=false){$('message').textContent=text;$('message').classList.toggle('error',error)}
function setBusy(value){state.busy=value;document.querySelectorAll('button').forEach(b=>b.disabled=value)}
function sheetUrl(range){return `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(TAB+'!'+range)}`}
async function api(url,options={}){
  const response=await fetch(url,{...options,headers:{Authorization:`Bearer ${state.token}`,...(options.body?{'Content-Type':'application/json'}:{})}});
  let body;try{body=await response.json()}catch{body={}}
  if(!response.ok){if(response.status===401)state.token=null;throw new Error(body.error?.message||`שגיאת Google (${response.status})`)}return body;
}
async function authorize(){
  for(let i=0;i<50&&!window.google?.accounts?.oauth2;i++)await new Promise(r=>setTimeout(r,100));
  if(!window.google?.accounts?.oauth2)throw new Error('לא ניתן לטעון את הכניסה של Google.');
  const clientId=$('clientId').value.trim();
  if(!/^\d+-[a-z0-9]+\.apps\.googleusercontent\.com$/.test(clientId))throw new Error('יש להזין OAuth Client ID תקין.');
  return new Promise((resolve,reject)=>{
    const client=google.accounts.oauth2.initTokenClient({client_id:clientId,scope:SCOPE,callback:r=>r.error?reject(new Error(r.error)):(state.token=r.access_token,resolve()),error_callback:e=>reject(new Error(e.message||e.type||'חלון הכניסה נסגר'))});
    client.requestAccessToken({prompt:''});
  });
}
// Google Sheets returns date cells as serial numbers. Preserve their type on every write.
function serialToDate(n){const d=new Date(Date.UTC(1899,11,30)+Math.round(n)*86400000);return d.toISOString().slice(0,10)}
function dateToSerial(s){if(!s)return '';const d=dateValue(s);if(!d)throw new Error('תאריך לא תקין');return Math.round((Date.UTC(+s.slice(0,4),+s.slice(5,7)-1,+s.slice(8,10))-Date.UTC(1899,11,30))/86400000)}
function normalizeDate(v){if(typeof v==='number')return serialToDate(v);if(/^\d{4}-\d{2}-\d{2}$/.test(String(v||'')))return String(v);const match=String(v||'').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);return match?`${match[3]}-${match[2].padStart(2,'0')}-${match[1].padStart(2,'0')}`:''}
function fromCells(cells,rowNumber){return {rowNumber,name:String(cells[0]??''),category:String(cells[1]??''),amount:String(cells[2]??''),currency:String(cells[3]??''),frequency:String(cells[4]??''),nextDate:normalizeDate(cells[5]),endDate:normalizeDate(cells[6]),payment:String(cells[7]??''),status:String(cells[8]??''),notes:String(cells[9]??'')}}
function toCells(r){return [r.name,r.category,Number(r.amount),r.currency,r.frequency,dateToSerial(r.nextDate),dateToSerial(r.endDate),r.payment,r.status,r.notes]}
async function readSheet(){const result=await api(sheetUrl('A1:J1000')+'?valueRenderOption=UNFORMATTED_VALUE');const values=result.values||[];
  if(HEADERS.some((h,i)=>values[0]?.[i]!==h))throw new Error('כותרות הגיליון אינן תואמות לקובץ ״מנויים ותשלומים״. בדוק גם ששם הלשונית הוא ״מנויים״.');
  return {values,rows:values.slice(1).map((r,i)=>fromCells(r,i+2)).filter(r=>r.name)};
}
async function load(){const fresh=await readSheet();state.rows=fresh.rows;$('setup').hidden=true;$('dashboard').hidden=false;$('connection').textContent='מחובר לגיליון';$('connection').classList.add('on');render();notice('הנתונים נטענו מהגיליון.');}
async function connect(){try{setBusy(true);localStorage.setItem('payment-client-id',$('clientId').value.trim());await authorize();await load()}catch(e){notice(e.message,true)}finally{setBusy(false)}}
async function ensureToken(){if(!state.token)await authorize()}
async function putRow(rowNumber,row){await api(sheetUrl(`A${rowNumber}:J${rowNumber}`)+'?valueInputOption=RAW',{method:'PUT',body:JSON.stringify({values:[toCells(row)]})})}
async function save(row){await ensureToken();const fresh=await readSheet();
  if(state.editing){const old=state.rows.find(r=>r.rowNumber===state.editing);const latest=fresh.rows.find(r=>r.rowNumber===state.editing);if(!old||!latest||JSON.stringify(old)!==JSON.stringify(latest))throw new Error('הרשומה השתנתה בגיליון. לחץ על רענון ונסה שוב.');await putRow(state.editing,row)}
  else {const blankIndex=fresh.values.findIndex((cells,i)=>i>0&&!cells?.[0]);const rowNumber=blankIndex>=1?blankIndex+1:Math.max(2,fresh.values.length+1);await putRow(rowNumber,row)}
  await load();notice('המנוי נשמר בגיליון.');
}
async function remove(row){if(!confirm(`למחוק את ${row.name} מהגיליון?`))return;try{setBusy(true);await ensureToken();const fresh=await readSheet();const latest=fresh.rows.find(r=>r.rowNumber===row.rowNumber);if(!latest||JSON.stringify(row)!==JSON.stringify(latest))throw new Error('הרשומה השתנתה בגיליון. רענן את הנתונים לפני מחיקה.');await api(sheetUrl(`A${row.rowNumber}:J${row.rowNumber}`)+':clear',{method:'POST',body:'{}'});await load();notice('המנוי נמחק מהגיליון.')}catch(e){notice(e.message,true)}finally{setBusy(false)}}
function ended(r){return !!r.endDate&&daysTo(r.endDate)<0}
function active(r){return r.status==='פעיל'&&!ended(r)}
function soon(r){const d=daysTo(r.nextDate);return active(r)&&d!==null&&d>=0&&d<=30&&(!r.endDate||r.nextDate<=r.endDate)}
function totals(factor){const sums={};state.rows.filter(active).forEach(r=>{if(r.frequency==='חד פעמי')return;const val=Number(r.amount);if(!Number.isFinite(val))return;sums[r.currency]=(sums[r.currency]||0)+val*(r.frequency==='חודשי'?(factor===12?12:1):(factor===12?1:1/12))});return Object.entries(sums).map(([c,n])=>money(n,c)).join(' + ')||money(0,'שקל')}
function render(){
  $('monthly').textContent=totals(1);$('yearly').textContent=totals(12);$('activeCount').textContent=state.rows.filter(active).length;$('upcomingCount').textContent=state.rows.filter(soon).length;
  const term=$('search').value.trim().toLowerCase(),filter=$('filter').value;
  const rows=state.rows.filter(r=>`${r.name} ${r.category}`.toLowerCase().includes(term)&&(filter==='all'||filter==='active'&&active(r)||filter==='soon'&&soon(r)||filter==='ended'&&ended(r))).sort((a,b)=>a.nextDate.localeCompare(b.nextDate));
  $('countLabel').textContent=`${rows.length} מתוך ${state.rows.length} מנויים`;$('empty').hidden=rows.length>0;$('cards').replaceChildren();
  for(const r of rows){const card=document.createElement('article');card.className='card';const status=ended(r)?'הסתיים':r.status==='מושהה'?'מושהה':r.status==='בוטל'?'בוטל':soon(r)?'חיוב קרוב':'פעיל';const pill=ended(r)||r.status==='בוטל'?'ended':soon(r)?'soon':'';
    card.innerHTML=`<div class="card-top"><h3>${escapeHtml(r.name)}</h3><span class="pill ${pill}">${status}</span></div><div class="price">${escapeHtml(money(r.amount,r.currency))}</div><div class="muted">${escapeHtml(r.frequency||'')} · ${escapeHtml(r.category||'ללא קטגוריה')}</div><dl><dt>חיוב הבא</dt><dd>${escapeHtml(dateText(r.nextDate))}</dd><dt>תאריך סיום</dt><dd>${escapeHtml(dateText(r.endDate))}</dd><dt>אמצעי תשלום</dt><dd>${escapeHtml(r.payment||'—')}</dd></dl><div class="card-actions"><button class="secondary edit">עריכה</button><button class="danger delete">מחיקה</button></div>`;
    card.querySelector('.edit').addEventListener('click',()=>openEditor(r));card.querySelector('.delete').addEventListener('click',()=>remove(r));$('cards').append(card);
  }
}
function openEditor(row=null){state.editing=row?.rowNumber||null;$('dialogTitle').textContent=row?'עריכת מנוי':'מנוי חדש';$('form').reset();if(row)for(const key of ['name','category','amount','currency','frequency','nextDate','endDate','payment','status','notes'])$('form').elements[key].value=row[key];else $('form').elements.nextDate.value=new Date().toLocaleDateString('en-CA');$('editor').showModal()}
async function submit(e){e.preventDefault();const row=Object.fromEntries(new FormData($('form')).entries());if(row.endDate&&row.endDate<row.nextDate){notice('תאריך הסיום מוקדם מתאריך החיוב הבא.',true);return}if(!Number.isFinite(Number(row.amount))||Number(row.amount)<0){notice('יש להזין סכום תקין.',true);return}try{setBusy(true);await save(row);$('editor').close()}catch(err){notice(err.message,true)}finally{setBusy(false)}}
$('connectBtn').addEventListener('click',connect);$('settingsBtn').addEventListener('click',()=>{$('setup').hidden=false;$('setup').scrollIntoView({behavior:'smooth'})});$('refreshBtn').addEventListener('click',async()=>{try{setBusy(true);await ensureToken();await load()}catch(e){notice(e.message,true)}finally{setBusy(false)}});
$('addBtn').addEventListener('click',()=>openEditor());$('closeBtn').addEventListener('click',()=>$('editor').close());$('cancelBtn').addEventListener('click',()=>$('editor').close());$('form').addEventListener('submit',submit);$('search').addEventListener('input',render);$('filter').addEventListener('change',render);
$('clientId').value=localStorage.getItem('payment-client-id')||'';
