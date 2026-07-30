(() => {
'use strict';

const STORE_KEY = 'fiat_ops_ai_v1';
const state = loadState();

const pageMeta = {
  dashboard:['總覽','今日風險、補幣與結算建議'],
  balance:['商戶餘額','貼上 Robot 文字並自動解析'],
  stats:['7日數據','上傳提款與代收報表'],
  pending:['Pending 分析','辨識延遲商戶、幣種與回報重點'],
  settlement:['結算話術','批量生成可直接傳送的結算內容'],
  report:['主管回報','整合所有資料生成回報'],
  backup:['備份','匯出、匯入或清除資料']
};

const $ = s => document.querySelector(s);
const $$ = s => [...document.querySelectorAll(s)];
const num = v => {
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  const n = Number(String(v ?? '').replace(/,/g,'').replace(/[^\d.-]/g,''));
  return Number.isFinite(n) ? n : 0;
};
const fmt = v => num(v).toLocaleString('en-US',{maximumFractionDigits:2});
const normCurrency = v => String(v ?? '').toUpperCase().replace(/FIAT/g,'').replace(/[^A-Z]/g,'').trim();
const esc = v => String(v ?? '').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
const today = () => new Date().toISOString().slice(0,10);

function defaultState(){
  return {balanceText:'',balances:[],withdraw:{rows:[],summary:{}},deposit:{rows:[],summary:{}},pendingText:'',pending:[],pendingSummary:'',settlements:[],report:''};
}
function loadState(){
  try { return Object.assign(defaultState(), JSON.parse(localStorage.getItem(STORE_KEY) || '{}')); }
  catch { return defaultState(); }
}
function save(){
  try { localStorage.setItem(STORE_KEY, JSON.stringify(state)); $('#saveStatus').textContent='剛剛已自動儲存'; }
  catch { $('#saveStatus').textContent='瀏覽器儲存失敗'; }
}
function toast(msg){ const el=$('#toast'); el.textContent=msg; el.classList.add('show'); clearTimeout(window.__toast); window.__toast=setTimeout(()=>el.classList.remove('show'),1800); }
function copyText(text){ navigator.clipboard?.writeText(text).then(()=>toast('已複製')).catch(()=>{ const t=document.createElement('textarea');t.value=text;document.body.appendChild(t);t.select();document.execCommand('copy');t.remove();toast('已複製'); }); }

function go(page){
  $$('.page').forEach(x=>x.classList.remove('active'));
  $$('.nav-item').forEach(x=>x.classList.toggle('active',x.dataset.page===page));
  $('#page-'+page).classList.add('active');
  $('#pageTitle').textContent=pageMeta[page][0];
  $('#pageSubtitle').textContent=pageMeta[page][1];
  window.scrollTo({top:0,behavior:'smooth'});
}
$$('.nav-item').forEach(b=>b.addEventListener('click',()=>go(b.dataset.page)));
$$('[data-go]').forEach(b=>b.addEventListener('click',()=>go(b.dataset.go)));

function parseBalanceText(text){
  const lines = text.split(/\r?\n/).map(x=>x.trim()).filter(Boolean);
  const out=[]; let merchant='未識別商戶', currency='';
  for(let i=0;i<lines.length;i++){
    const line=lines[i];
    const curMatch=line.match(/\b([A-Z]{3,5})(?:FIAT)?\b/i);
    if(curMatch && /FIAT|currency|幣種/i.test(line)) currency=normCurrency(curMatch[1]);
    if(!/available|balance|currency|fiat|pt\d|usd|^\d/i.test(line) && line.length<80) merchant=line.replace(/[:：]$/,'').trim();
    const amountMatch=line.match(/(-?[\d,]+(?:\.\d+)?)\s*\(\s*(-?[\d,]+(?:\.\d+)?)\s*USD\s*\)/i);
    if(amountMatch){
      if(!currency){
        for(let j=Math.max(0,i-4);j<=Math.min(lines.length-1,i+4);j++){
          const m=lines[j].match(/\b([A-Z]{3,5})FIAT\b/i); if(m){currency=normCurrency(m[1]);break;}
        }
      }
      out.push({merchant,currency:currency||'UNKNOWN',amount:num(amountMatch[1]),usd:num(amountMatch[2])});
    }
  }
  if(!out.length){
    const blocks=text.split(/PT\d+[HMS]/i);
    for(const block of blocks){
      const am=block.match(/(-?[\d,]+(?:\.\d+)?)\s*\(\s*(-?[\d,]+(?:\.\d+)?)\s*USD\s*\)/i);
      const cm=block.match(/\b([A-Z]{3,5})FIAT\b/i);
      if(am) out.push({merchant:(block.match(/(?:^|\n)([^\n]{2,50})\n/)||[])[1]||'未識別商戶',currency:cm?normCurrency(cm[1]):'UNKNOWN',amount:num(am[1]),usd:num(am[2])});
    }
  }
  return out;
}
function renderBalances(){
  $('#balanceInput').value=state.balanceText||'';
  $('#balanceCount').textContent=`${state.balances.length} 筆`;
  $('#balanceTable').innerHTML=state.balances.length?state.balances.map(r=>`<tr><td>${esc(r.merchant)}</td><td>${esc(r.currency)}</td><td>${fmt(r.amount)}</td><td>${fmt(r.usd)}</td></tr>`).join(''):'<tr><td colspan="4" class="muted center">尚未解析</td></tr>';
}

function parseDelimited(text){
  const lines=text.replace(/^\uFEFF/,'').split(/\r?\n/).filter(x=>x.trim());
  if(!lines.length) return [];
  const delim=lines[0].includes('\t')?'\t':',';
  const parseLine=line=>{ let out=[],cur='',q=false; for(let i=0;i<line.length;i++){const c=line[i];if(c==='"'){if(q&&line[i+1]==='"'){cur+='"';i++;}else q=!q;}else if(c===delim&&!q){out.push(cur);cur='';}else cur+=c;}out.push(cur);return out;};
  const heads=parseLine(lines[0]).map(x=>x.trim());
  return lines.slice(1).map(l=>{const vals=parseLine(l);const o={};heads.forEach((h,i)=>o[h]=vals[i]??'');return o;});
}
async function readSheet(file){
  const ext=file.name.split('.').pop().toLowerCase();
  if(ext==='csv'||ext==='tsv') return parseDelimited(await file.text());
  if(!window.XLSX) throw new Error('Excel 元件載入失敗，請確認網路後重新整理；也可另存 CSV 再上傳。');
  const buf=await file.arrayBuffer(); const wb=XLSX.read(buf,{type:'array'});
  const ws=wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws,{defval:''});
}
function pick(obj,names){
  const entries=Object.entries(obj); for(const n of names){ const f=entries.find(([k])=>k.toLowerCase().replace(/\s|_|-/g,'')===n.toLowerCase().replace(/\s|_|-/g,'')); if(f) return f[1]; }
  for(const n of names){ const f=entries.find(([k])=>k.toLowerCase().includes(n.toLowerCase())); if(f) return f[1]; }
  return '';
}
function summarizeStats(rows){
  const summary={};
  rows.forEach(r=>{
    const c=normCurrency(pick(r,['Currency','幣種']));
    const a=num(pick(r,['Total Amount','TotalAmount','Sum Amount','Amount','總金額']));
    if(!c||!a) return;
    summary[c]=(summary[c]||0)+a;
  });
  return summary;
}
async function handleStats(file,type){
  try{
    const rows=await readSheet(file); const summary=summarizeStats(rows);
    state[type]={rows,summary}; save(); renderStats(); renderDashboard();
    $('#'+type+'Info').textContent=`已讀取 ${rows.length} 列，辨識 ${Object.keys(summary).length} 個幣種`;
    toast('報表已讀取');
  }catch(e){alert(e.message);}
}
function renderStats(){
  const currencies=[...new Set([...Object.keys(state.withdraw.summary||{}),...Object.keys(state.deposit.summary||{})])].sort();
  $('#withdrawInfo').textContent=state.withdraw.rows.length?`已載入 ${state.withdraw.rows.length} 列，${Object.keys(state.withdraw.summary||{}).length} 個幣種`:'尚未上傳';
  $('#depositInfo').textContent=state.deposit.rows.length?`已載入 ${state.deposit.rows.length} 列，${Object.keys(state.deposit.summary||{}).length} 個幣種`:'尚未上傳';
  $('#statsTable').innerHTML=currencies.length?currencies.map(c=>{const w=state.withdraw.summary[c]||0,d=state.deposit.summary[c]||0;return `<tr><td>${c}</td><td>${fmt(w)}</td><td>${fmt(w/7)}</td><td>${fmt(d)}</td><td>${fmt(d/7)}</td><td>${fmt((d-w)/7)}</td></tr>`}).join(''):'<tr><td colspan="6" class="muted center">尚無資料</td></tr>';
}

function parsePendingText(text){
  const lines=text.trim().split(/\r?\n/).filter(x=>x.trim());
  if(!lines.length) return [];
  let rows=[];
  if(lines[0].includes('\t')){
    const heads=lines[0].split('\t').map(x=>x.trim());
    rows=lines.slice(1).map(line=>{const vals=line.split('\t');const o={};heads.forEach((h,i)=>o[h]=vals[i]??'');return o;});
  }else{
    const headers=['Provider','Channel','Currency','Sum Amount','Count','Avg. Pend Hour','Max Hour','Min Hour','Action'];
    if(/Provider\s+Channel\s+Currency/i.test(lines[0])){
      const body=lines.slice(1).join('\n');
      rows=body.split(/\n/).map(line=>{
        const p=line.trim().split(/\s{2,}|\t+/); const o={};headers.forEach((h,i)=>o[h]=p[i]??'');return o;
      });
    }else return [];
  }
  return rows.map(r=>({
    provider:String(pick(r,['Provider','商戶'])||'未識別'),
    channel:String(pick(r,['Channel','通道'])||''),
    currency:normCurrency(pick(r,['Currency','幣種'])),
    amount:num(pick(r,['Sum Amount','Total Amount','Amount','金額'])),
    count:num(pick(r,['Count','筆數'])),
    avg:num(pick(r,['Avg. Pend Hour','Avg Pend Hour','Avg Hour','平均時數'])),
    max:num(pick(r,['Max Hour','最長時數'])),
    min:num(pick(r,['Min Hour','最短時數']))
  })).filter(r=>r.currency&&r.provider);
}
function pendingJudgement(r){
  if(r.max>=24||r.avg>=12) return ['嚴重延遲','risk'];
  if(r.max>=12||r.avg>=6) return ['處理偏慢','warn'];
  return ['持續觀察','ok'];
}
function createPendingSummary(rows){
  if(!rows.length) return '';
  const byCur={}; rows.forEach(r=>(byCur[r.currency]??=[]).push(r));
  const parts=[];
  Object.entries(byCur).sort((a,b)=>Math.max(...b[1].map(x=>x.max))-Math.max(...a[1].map(x=>x.max))).forEach(([c,list])=>{
    const worst=[...list].sort((a,b)=>b.avg-a.avg)[0];
    const totalCount=list.reduce((s,x)=>s+x.count,0), totalAmount=list.reduce((s,x)=>s+x.amount,0);
    const severe=list.filter(x=>x.avg>=12||x.max>=24);
    let line=`${c}：共 ${fmt(totalCount)} 筆，金額 ${fmt(totalAmount)}，延遲主要集中於 ${worst.provider}`;
    line+=`（平均 ${fmt(worst.avg)} 小時，最長 ${fmt(worst.max)} 小時）`;
    if(severe.length>1) line+='，屬幣種多商戶普遍延遲';
    else line+='，較可能為個別商戶處理偏慢';
    parts.push(line+'。');
  });
  return parts.join('\n');
}
function renderPending(){
  $('#pendingInput').value=state.pendingText||'';
  $('#pendingCount').textContent=`${state.pending.length} 筆資料`;
  $('#pendingSummary').textContent=state.pendingSummary||'尚未分析';
  $('#pendingSummary').classList.toggle('muted',!state.pendingSummary);
  $('#pendingTable').innerHTML=state.pending.length?state.pending.map(r=>{const [t,c]=pendingJudgement(r);return `<tr><td>${r.currency}</td><td>${esc(r.provider)}</td><td>${esc(r.channel)}</td><td>${fmt(r.amount)}</td><td>${fmt(r.count)}</td><td>${fmt(r.avg)}</td><td>${fmt(r.max)}</td><td class="status ${c}">${t}</td></tr>`}).join(''):'<tr><td colspan="8" class="muted center">尚無資料</td></tr>';
}

function aggregateBalance(){
  const out={}; state.balances.forEach(r=>{out[r.currency]=(out[r.currency]||0)+num(r.amount)}); return out;
}
function dashboardRows(){
  const bal=aggregateBalance(), pend={};
  state.pending.forEach(r=>{pend[r.currency]=(pend[r.currency]||0)+r.count});
  const cs=[...new Set([...Object.keys(bal),...Object.keys(state.withdraw.summary||{}),...Object.keys(state.deposit.summary||{}),...Object.keys(pend)])].sort();
  return cs.map(c=>{
    const b=bal[c]||0,w=(state.withdraw.summary[c]||0)/7,d=(state.deposit.summary[c]||0)/7,p=pend[c]||0;
    const days=w?b/w:null;
    let status='正常',klass='ok',advice='持續觀察';
    if(w&&days<1){status='風險';klass='risk';advice='建議立即補幣';}
    else if((w&&days<1.5)||p>=20){status='注意';klass='warn';advice=p>=20?'追蹤 Pending 並切換備援':'餘額偏低，準備補幣';}
    else if(d>w&&w&&days>3){advice='可考慮結算';}
    else if(w>d&&w){advice='關注餘額下降';}
    return {currency:c,balance:b,w,d,days,p,status,klass,advice};
  });
}
function renderDashboard(){
  const rows=dashboardRows();
  $('#riskCount').textContent=rows.filter(x=>x.status==='風險').length;
  $('#warnCount').textContent=rows.filter(x=>x.status==='注意').length;
  $('#normalCount').textContent=rows.filter(x=>x.status==='正常').length;
  $('#loadedCount').textContent=[state.balances.length,state.withdraw.rows.length,state.deposit.rows.length,state.pending.length].filter(Boolean).length;
  $('#dashboardTable').innerHTML=rows.length?rows.map(r=>`<tr><td>${r.currency}</td><td>${fmt(r.balance)}</td><td>${fmt(r.w)}</td><td>${fmt(r.d)}</td><td>${r.days===null?'—':fmt(r.days)}</td><td>${fmt(r.p)}</td><td class="status ${r.klass}">${r.status}</td><td>${r.advice}</td></tr>`).join(''):'<tr><td colspan="8" class="muted center">尚無資料</td></tr>';
  const priorities=rows.filter(x=>x.status!=='正常'||x.p>0).slice(0,8);
  $('#priorityList').className=priorities.length?'':'empty';
  $('#priorityList').innerHTML=priorities.length?priorities.map(r=>`<div class="priority-item"><b>${r.currency}｜${r.status}</b><span>${r.advice}${r.p?`；Pending ${fmt(r.p)} 筆`:''}</span></div>`).join(''):'尚未有足夠資料。先貼上商戶餘額或 Pending 資料。';
  const funds=rows.filter(x=>/補幣|結算|餘額/.test(x.advice));
  $('#fundingList').className=funds.length?'':'empty';
  $('#fundingList').innerHTML=funds.length?funds.map(r=>`<div class="fund-item"><b>${r.currency}</b><span>${r.advice}${r.days!==null?`；可支撐 ${fmt(r.days)} 天`:''}</span></div>`).join(''):'等待餘額及 7 日提款資料。';
}

function settlementRow(data={}){
  const tr=document.createElement('tr');
  tr.innerHTML=['merchant','currency','amount','rate','fiatFee','usdFee','received','txid'].map(k=>`<td><input data-k="${k}" value="${esc(data[k]??'')}"></td>`).join('')+`<td><button class="btn small danger remove-row">刪除</button></td>`;
  tr.querySelector('.remove-row').onclick=()=>tr.remove();
  return tr;
}
function renderSettlementRows(){ const body=$('#settlementRows'); body.innerHTML=''; const rows=state.settlements.length?state.settlements:[{}]; rows.forEach(r=>body.appendChild(settlementRow(r))); }
function collectSettlements(){ return $$('#settlementRows tr').map(tr=>{const o={};tr.querySelectorAll('input').forEach(i=>o[i.dataset.k]=i.value.trim());return o;}).filter(o=>Object.values(o).some(Boolean)); }
function makeSettlementText(r){
  const merchant=r.merchant||'商戶',c=(r.currency||'').toUpperCase(),amount=num(r.amount),rate=num(r.rate),ff=num(r.fiatFee),uf=num(r.usdFee),received=num(r.received);
  if(!amount||!rate) return '';
  let calc='',usdt=0;
  if(ff){usdt=(amount-ff)/rate;calc=`(${fmt(amount)}-${fmt(ff)})/${fmt(rate)}=${fmt(received||usdt)} U`;}
  else if(uf){usdt=amount/rate-uf;calc=`(${fmt(amount)}/${fmt(rate)})-fee ${fmt(uf)}=${fmt(received||usdt)} U`;}
  else {usdt=amount/rate;calc=`${fmt(amount)}/${fmt(rate)}=${fmt(received||usdt)} U`;}
  return `${merchant}出U,結算${c} ${fmt(amount)},匯率${fmt(rate)}。${calc},請你確認,謝謝。${r.txid?`\nTXID：${r.txid}`:''}`;
}

function buildReport(tone='simple'){
  const rows=dashboardRows(), issues=rows.filter(x=>x.status!=='正常'||x.p>0), normals=rows.filter(x=>x.status==='正常'&&!x.p);
  const lines=[];
  lines.push(`${$('#reportDate').value||today()} 法幣營運回報`);
  lines.push('');
  lines.push('【重點事項】');
  if(!issues.length) lines.push('目前未發現明顯異常。');
  issues.forEach(r=>{
    let s=`${r.currency}：${r.advice}`;
    if(r.balance) s+=`，目前餘額 ${fmt(r.balance)}`;
    if(r.w) s+=`，7日提款日均 ${fmt(r.w)}`;
    if(r.p) s+=`，Pending ${fmt(r.p)} 筆`;
    if(tone==='detail'&&r.days!==null) s+=`，約可支撐 ${fmt(r.days)} 天`;
    lines.push(s+'。');
  });
  if(state.pendingSummary){ lines.push(''); lines.push('【Pending 分析】'); lines.push(state.pendingSummary); }
  lines.push(''); lines.push('【正常幣種】');
  lines.push(normals.length?normals.map(x=>x.currency).join('、')+' 目前無明顯異常。':'目前無可列入的正常幣種資料。');
  return lines.join('\n');
}

function exportCSV(rows,name){
  const csv='\uFEFF'+rows.map(r=>r.map(v=>`"${String(v??'').replace(/"/g,'""')}"`).join(',')).join('\n');
  download(new Blob([csv],{type:'text/csv;charset=utf-8'}),name);
}
function download(blob,name){const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},500);}
function renderAll(){renderBalances();renderStats();renderPending();renderSettlementRows();renderDashboard();$('#reportDate').value=$('#reportDate').value||today();$('#reportOutput').value=state.report||'';}

$('#parseBalance').onclick=()=>{state.balanceText=$('#balanceInput').value;state.balances=parseBalanceText(state.balanceText);save();renderBalances();renderDashboard();toast(`已解析 ${state.balances.length} 筆餘額`);};
$('#clearBalance').onclick=()=>{state.balanceText='';state.balances=[];save();renderBalances();renderDashboard();};
$('#sampleBalance').onclick=()=>{$('#balanceInput').value=`TopPayNew\nCurrency: IDRFIAT\nAvailable Balance\n1225299801.55 (67865.96 USD)\nPT3M\n\nEtpay\nCurrency: INRFIAT\nAvailable Balance\n358751.00 (4305.01 USD)\nPT3M`;};
$('#withdrawFile').onchange=e=>e.target.files[0]&&handleStats(e.target.files[0],'withdraw');
$('#depositFile').onchange=e=>e.target.files[0]&&handleStats(e.target.files[0],'deposit');
$('#clearStats').onclick=()=>{state.withdraw={rows:[],summary:{}};state.deposit={rows:[],summary:{}};save();renderStats();renderDashboard();};
$('#parsePending').onclick=()=>{state.pendingText=$('#pendingInput').value;state.pending=parsePendingText(state.pendingText);state.pendingSummary=createPendingSummary(state.pending);save();renderPending();renderDashboard();toast(`已分析 ${state.pending.length} 列`);};
$('#clearPending').onclick=()=>{state.pendingText='';state.pending=[];state.pendingSummary='';save();renderPending();renderDashboard();};
$('#samplePending').onclick=()=>{$('#pendingInput').value=`Provider\tChannel\tCurrency\tSum Amount\tCount\tAvg. Pend Hour\tMax Hour\tMin Hour\tAction\nTopPayNew\tTopPayNew_INR_UPI\tINRFIAT\t1185652.00\t15\t9.50\t23.00\t6.03\t\nEtpay\tEtpay_INR_BANKTRANSFER\tINRFIAT\t358751.00\t6\t21.93\t36.71\t8.75\t\nSimpaisa\tSimpaisa_Jazzcash\tPKRFIAT\t131739.00\t3\t10.13\t11.96\t6.24\t`;};
$('#copyPendingReport').onclick=()=>copyText(state.pendingSummary||'');
$('#addSettlementRow').onclick=()=>$('#settlementRows').appendChild(settlementRow());
$('#clearSettlement').onclick=()=>{state.settlements=[];$('#settlementOutput').value='';renderSettlementRows();save();};
$('#generateSettlement').onclick=()=>{state.settlements=collectSettlements();const out=state.settlements.map(makeSettlementText).filter(Boolean).join('\n\n');$('#settlementOutput').value=out;save();toast('結算話術已生成');};
$('#copySettlement').onclick=()=>copyText($('#settlementOutput').value);
$('#generateReport').onclick=()=>{state.report=buildReport($('#reportTone').value);$('#reportOutput').value=state.report;save();};
$('#generateReportTop').onclick=()=>{go('report');state.report=buildReport($('#reportTone').value);$('#reportOutput').value=state.report;save();};
$('#copyReport').onclick=()=>copyText($('#reportOutput').value);
$('#refreshBtn').onclick=()=>{renderAll();toast('已重新整理');};
$('#exportDashboard').onclick=()=>{const rows=dashboardRows();exportCSV([['Currency','Balance','7D Withdraw Avg','7D Deposit Avg','Support Days','Pending','Status','Advice'],...rows.map(r=>[r.currency,r.balance,r.w,r.d,r.days??'',r.p,r.status,r.advice])],`fiat-dashboard-${today()}.csv`);};
$('#exportBackup').onclick=()=>download(new Blob([JSON.stringify(state,null,2)],{type:'application/json'}),`fiat-operations-backup-${today()}.json`);
$('#importBackup').onchange=async e=>{try{const data=JSON.parse(await e.target.files[0].text());Object.assign(state,defaultState(),data);save();renderAll();toast('備份已還原');}catch{alert('備份檔案無法讀取');}};
$('#resetAll').onclick=()=>{if(confirm('確定清除全部資料？')){Object.assign(state,defaultState());save();renderAll();toast('資料已清除');}};
$('#reportOutput').addEventListener('input',e=>{state.report=e.target.value;save();});
$('#balanceInput').addEventListener('input',e=>{state.balanceText=e.target.value;});
$('#pendingInput').addEventListener('input',e=>{state.pendingText=e.target.value;});

renderAll();
})();
