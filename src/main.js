import './style.css';
import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { createWorker } from 'tesseract.js';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

const TEMPLATE_URL = '/품목내역(통합).xls';
const ACCEPTED = '.xlsx,.xls,.csv,.pdf,.png,.jpg,.jpeg,.webp';
let items = [];
let sourceName = '';
let confirmed = false;
let ocrWorker = null;

const $ = s => document.querySelector(s);
const clean = v => String(v ?? '').replace(/\s+/g, ' ').trim();
const num = v => {
  const n = Number(String(v ?? '').replace(/[^0-9.-]/g, ''));
  return Number.isFinite(n) ? n : 0;
};
const won = v => `${Math.round(num(v)).toLocaleString('ko-KR')}원`;
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const uuid = () => crypto.randomUUID?.() || `${Date.now()}-${Math.random()}`;

const app = $('#app');
app.innerHTML = `
<header class="hero"><div class="hero-inner">
  <span class="badge">API 키 없이 브라우저에서 처리</span>
  <h1>품의서 작성기</h1>
  <p>견적서에서 품목을 자동 추출하고, 사람이 확인·수정·확정한 값만 에듀파인 품목내역(통합) 원본 서식에 입력합니다.</p>
</div></header>
<main>
  <section class="card">
    <div class="step">1</div><h2>견적서 업로드</h2>
    <p class="sub">엑셀·PDF·사진을 지원합니다. 자동 인식 결과는 반드시 다음 단계에서 확인하세요.</p>
    <div class="actions">
      <button id="uploadBtn" class="btn primary">견적서 업로드</button>
      <button id="blankBtn" class="btn secondary">에듀파인 빈 서식 내려받기</button>
      <input id="fileInput" type="file" accept="${ACCEPTED}" hidden>
    </div>
    <div id="dropzone" class="dropzone"><b>파일을 끌어다 놓아도 됩니다</b><span>XLSX · XLS · CSV · PDF · PNG · JPG · WEBP</span></div>
    <div id="status" class="status"></div>
  </section>

  <section id="resultCard" class="card hidden">
    <div class="head"><div><div class="step">2</div><h2>품목 확인 및 수정</h2><p id="sourceInfo" class="sub"></p></div><button id="addBtn" class="btn soft">행 추가</button></div>
    <div class="notice">자동 인식이 애매한 곳은 직접 수정하세요. 합계·총계·공급가액·부가세 행은 품목에서 제외됩니다.</div>
    <div class="table-wrap"><table>
      <thead><tr><th>순번</th><th>내용</th><th>규격</th><th>수량</th><th>예상단가<br><small>부가세 포함</small></th><th>예상금액</th><th></th></tr></thead>
      <tbody id="rows"></tbody>
    </table></div>
    <div class="summary"><div><span>품목 수</span><strong id="count">0개</strong></div><div><span>총금액</span><strong id="total">0원</strong></div><div><span>확정 상태</span><strong id="confirmState">미확정</strong></div></div>
    <div class="actions top">
      <button id="confirmBtn" class="btn confirm">이 내용으로 최종 확정</button>
      <button id="downloadBtn" class="btn primary" disabled>품목내역(통합).xls 내려받기</button>
      <button id="resetBtn" class="btn secondary">다른 견적서 불러오기</button>
    </div>
    <p class="foot">다운로드 시 다시 분석하지 않습니다. 화면에서 확정한 내용·규격·수량·예상단가만 원본 서식의 A:D 열에 그대로 입력됩니다.</p>
  </section>
</main>`;

function setStatus(message, kind='info') {
  const box = $('#status'); box.textContent = message; box.className = `status show ${kind}`;
}
function saveBlob(blob, name) {
  const url = URL.createObjectURL(blob); const a = document.createElement('a');
  a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function isSummaryLabel(v) {
  const s = clean(v).toLowerCase().replace(/[\s·ㆍ:_()\[\]-]/g, '');
  return /^(합계|총계|소계|공급가액|부가세|세액|vat|합계금액|총금액|견적금액|금액합계|계)$/.test(s)
    || /^(합계|총계|소계|공급가액|부가세|세액|합계금액|총금액)/.test(s);
}
function headerType(v) {
  const s = clean(v).toLowerCase().replace(/\s/g, '');
  const dict = {
    name:['품명','품목명','상품명','제품명','내역','내용','물품명','자재명','description','item'],
    spec:['규격','사양','용량','포장','spec','size'],
    qty:['수량','주문수량','요청수량','qty','quantity'],
    price:['단가','예상단가','공급단가','판매가','unitprice','price'],
    amount:['금액','예상금액','합계금액','amount','total']
  };
  return Object.keys(dict).find(k => dict[k].some(x => s.includes(x))) || '';
}
function numericValue(v) {
  const s = String(v ?? '').trim();
  if (!/^-?[₩￦]?\s*[\d,]+(?:\.\d+)?\s*(원)?$/.test(s)) return null;
  return num(s);
}
function tableToItems(rows) {
  if (!rows?.length) return [];
  let headerRow = -1, map = {}, best = -1;
  for (let r=0; r<Math.min(rows.length, 50); r++) {
    const found = {};
    rows[r].forEach((cell,c) => { const t=headerType(cell); if(t && found[t]===undefined) found[t]=c; });
    const score = ['name','qty','price','amount'].filter(k => found[k]!==undefined).length + (found.spec!==undefined?.5:0);
    if (score > best) { best=score; headerRow=r; map=found; }
  }
  if (best < 2) return inferItems(rows);
  const out=[];
  for (let r=headerRow+1; r<rows.length; r++) {
    const row=rows[r]; const name=clean(row[map.name]);
    if (!name || isSummaryLabel(name) || /^(비고|특기사항|계좌번호|사업자|주소|연락처|대표자)/i.test(name)) continue;
    const quantity = map.qty!==undefined ? num(row[map.qty]) : 0;
    const price = map.price!==undefined ? num(row[map.price]) : 0;
    const amount = map.amount!==undefined ? num(row[map.amount]) : 0;
    if (!(quantity || price || amount)) continue;
    const q = quantity || 1;
    const p = price || (amount && q ? Math.round(amount/q) : 0);
    if (!p) continue;
    out.push({id:uuid(), name, spec:clean(row[map.spec]), quantity:q, unitPrice:p, amount:Math.round(q*p)});
  }
  return out;
}
function inferItems(rows) {
  const out=[];
  for (const row of rows) {
    const vals=row.map(clean).filter(Boolean);
    if(vals.length<3) continue;
    const nums=vals.map((v,i)=>({i,v:numericValue(v)})).filter(x=>x.v!==null);
    const texts=vals.filter(v=>numericValue(v)===null && !/^(no|번호|순번)$/i.test(v));
    if(nums.length<2 || !texts.length || isSummaryLabel(texts[0])) continue;
    let quantity=1, unitPrice=0, amount=0;
    if(nums.length>=3){ quantity=nums[nums.length-3].v; unitPrice=nums[nums.length-2].v; amount=nums[nums.length-1].v; }
    else { unitPrice=nums[nums.length-2].v; amount=nums[nums.length-1].v; quantity=amount>=unitPrice&&unitPrice>0?Math.max(1,Math.round(amount/unitPrice)):1; }
    if(unitPrice<=0 || amount<unitPrice) continue;
    out.push({id:uuid(), name:texts[0], spec:texts.slice(1).join(' '), quantity, unitPrice, amount:Math.round(quantity*unitPrice)});
  }
  return out;
}
function rowsFromText(textValue) {
  return String(textValue).split(/\r?\n/).map(x=>x.trim()).filter(Boolean).map(line=>line.split(/\t|\s{2,}/).filter(Boolean));
}
async function parseSpreadsheet(file) {
  const wb=XLSX.read(await file.arrayBuffer(),{type:'array',raw:false}); let best=[];
  for(const name of wb.SheetNames){ const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:'',raw:false}); const parsed=tableToItems(rows); if(parsed.length>best.length) best=parsed; }
  return best;
}
function pdfRows(items) {
  const lines=new Map();
  for(const item of items){ const y=Math.round((item.transform?.[5]||0)/4)*4; if(!lines.has(y))lines.set(y,[]); lines.get(y).push({x:item.transform?.[4]||0,t:item.str||''}); }
  return [...lines.entries()].sort((a,b)=>b[0]-a[0]).map(([,cells])=>cells.sort((a,b)=>a.x-b.x).map(c=>c.t));
}
async function getWorker(){ if(!ocrWorker) ocrWorker=await createWorker('kor+eng'); return ocrWorker; }
async function recognizeCanvas(canvas){ const worker=await getWorker(); const {data:{text}}=await worker.recognize(canvas); return text; }
async function parsePdf(file) {
  const pdf=await pdfjsLib.getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise; let rows=[]; let chars=0;
  for(let p=1;p<=Math.min(pdf.numPages,10);p++){ setStatus(`PDF ${p}/${Math.min(pdf.numPages,10)}쪽을 읽고 있습니다…`,'loading'); const page=await pdf.getPage(p); const content=await page.getTextContent(); chars+=content.items.reduce((s,i)=>s+(i.str||'').length,0); rows.push(...pdfRows(content.items)); }
  let parsed=tableToItems(rows); if(parsed.length) return parsed;
  if(chars<100 || !parsed.length){ rows=[]; for(let p=1;p<=Math.min(pdf.numPages,6);p++){ setStatus(`스캔 PDF ${p}/${Math.min(pdf.numPages,6)}쪽을 문자 인식 중입니다…`,'loading'); const page=await pdf.getPage(p); const vp=page.getViewport({scale:2}); const canvas=document.createElement('canvas'); canvas.width=vp.width; canvas.height=vp.height; await page.render({canvasContext:canvas.getContext('2d'),viewport:vp}).promise; rows.push(...rowsFromText(await recognizeCanvas(canvas))); } parsed=tableToItems(rows); }
  return parsed;
}
async function parseImage(file){ setStatus('사진을 문자 인식 중입니다…','loading'); const bitmap=await createImageBitmap(file); const scale=Math.min(2,2400/Math.max(bitmap.width,bitmap.height)); const canvas=document.createElement('canvas'); canvas.width=Math.round(bitmap.width*scale); canvas.height=Math.round(bitmap.height*scale); canvas.getContext('2d').drawImage(bitmap,0,0,canvas.width,canvas.height); return tableToItems(rowsFromText(await recognizeCanvas(canvas))); }
function normalize(list){ return list.filter(x=>x.name&&!isSummaryLabel(x.name)).map(x=>({id:x.id||uuid(),name:clean(x.name),spec:clean(x.spec),quantity:num(x.quantity)||1,unitPrice:Math.round(num(x.unitPrice)),amount:Math.round((num(x.quantity)||1)*num(x.unitPrice))})).filter(x=>x.unitPrice>0); }
async function handleFile(file){
  if(!file)return; confirmed=false; sourceName=file.name; setStatus(`${file.name}을 분석하고 있습니다…`,'loading');
  try{ const ext=file.name.split('.').pop().toLowerCase(); let parsed=[]; if(['xlsx','xls','csv'].includes(ext)) parsed=await parseSpreadsheet(file); else if(ext==='pdf') parsed=await parsePdf(file); else parsed=await parseImage(file); items=normalize(parsed); if(!items.length) throw new Error('품목을 자동으로 찾지 못했습니다. 행 추가 버튼으로 직접 입력하거나 표가 선명한 파일을 올려 주세요.'); render(); $('#resultCard').classList.remove('hidden'); $('#sourceInfo').textContent=`원본 파일: ${sourceName}`; setStatus(`${items.length}개 품목을 찾았습니다. 내용을 확인하고 수정한 뒤 최종 확정하세요.`,'success'); $('#resultCard').scrollIntoView({behavior:'smooth'}); }
  catch(e){console.error(e);setStatus(e.message||'파일 분석 중 오류가 발생했습니다.','error');}
}
function markDirty(){ confirmed=false; $('#downloadBtn').disabled=true; $('#confirmState').textContent='미확정'; $('#confirmState').className='bad'; }
function render(){
  $('#rows').innerHTML=items.map((it,i)=>`<tr data-id="${it.id}"><td class="center">${i+1}</td><td><input data-k="name" value="${esc(it.name)}"></td><td><input data-k="spec" value="${esc(it.spec)}"></td><td><input data-k="quantity" type="number" min="0" step="0.01" value="${it.quantity}"></td><td><input data-k="unitPrice" type="number" min="0" step="1" value="${it.unitPrice}"></td><td class="money">${won(it.amount)}</td><td><button class="delete" data-del="${it.id}">×</button></td></tr>`).join(''); updateSummary(); markDirty();
}
function updateSummary(){ items.forEach(x=>x.amount=Math.round(num(x.quantity)*num(x.unitPrice))); $('#count').textContent=`${items.length}개`; $('#total').textContent=won(items.reduce((s,x)=>s+x.amount,0)); }
async function exportTemplate(){
  if(!confirmed) throw new Error('먼저 “이 내용으로 최종 확정”을 눌러 주세요.');
  const finalItems=items.filter(x=>clean(x.name)&&num(x.quantity)>0&&num(x.unitPrice)>=0);
  if(!finalItems.length) throw new Error('다운로드할 품목이 없습니다.');
  const res=await fetch(encodeURI(TEMPLATE_URL),{cache:'no-store'}); if(!res.ok) throw new Error(`품목내역(통합) 원본 서식을 불러오지 못했습니다. (${res.status})`);
  const wb=XLSX.read(await res.arrayBuffer(),{type:'array',cellStyles:true,cellNF:true,bookVBA:true});
  const ws=wb.Sheets[wb.SheetNames[0]]; if(!ws) throw new Error('원본 서식 시트를 찾지 못했습니다.');
  // 원본 서식은 A1:D1 = 내용, 규격, 수량, 예상단가. 이 4개 열 외에는 만들지 않는다.
  const templateCells=['A1','B1','C1','D1'].map(a=>clean(ws[a]?.v));
  const expected=['내용','규격','수량','예상단가'];
  if(!expected.every((v,i)=>templateCells[i]===v)) throw new Error('등록된 파일이 지정된 품목내역(통합) 원본 서식과 다릅니다.');
  const oldRange=XLSX.utils.decode_range(ws['!ref']||'A1:D1');
  for(let r=1;r<=Math.max(oldRange.e.r,500);r++) for(let c=0;c<4;c++) delete ws[XLSX.utils.encode_cell({r,c})];
  finalItems.forEach((it,i)=>{
    const r=i+1; const vals=[clean(it.name),clean(it.spec),num(it.quantity),Math.round(num(it.unitPrice))];
    vals.forEach((v,c)=>{ const addr=XLSX.utils.encode_cell({r,c}); const style=ws[XLSX.utils.encode_cell({r:0,c})]?.s; ws[addr]={v,t:typeof v==='number'?'n':'s'}; if(style) ws[addr].s=style; if(c>=2) ws[addr].z='#,##0'; });
  });
  ws['!ref']=`A1:D${finalItems.length+1}`;
  const out=XLSX.write(wb,{bookType:'biff8',type:'array',cellStyles:true,bookVBA:true});
  saveBlob(new Blob([out],{type:'application/vnd.ms-excel'}),'품목내역(통합).xls');
}

$('#uploadBtn').onclick=()=>$('#fileInput').click();
$('#blankBtn').onclick=()=>downloadBlank().catch(e=>alert(e.message));
async function downloadBlank(){const r=await fetch(encodeURI(TEMPLATE_URL),{cache:'no-store'});if(!r.ok)throw new Error('빈 서식을 불러오지 못했습니다.');saveBlob(await r.blob(),'품목내역(통합).xls');}
$('#fileInput').onchange=()=>handleFile($('#fileInput').files[0]);
const dz=$('#dropzone'); ['dragenter','dragover'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add('over')})); ['dragleave','drop'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove('over')})); dz.addEventListener('drop',e=>handleFile(e.dataTransfer.files[0]));
$('#rows').addEventListener('input',e=>{const tr=e.target.closest('tr');if(!tr)return;const it=items.find(x=>x.id===tr.dataset.id);if(!it)return;const k=e.target.dataset.k;it[k]=['quantity','unitPrice'].includes(k)?num(e.target.value):e.target.value;it.amount=Math.round(num(it.quantity)*num(it.unitPrice));tr.children[5].textContent=won(it.amount);updateSummary();markDirty();});
$('#rows').addEventListener('click',e=>{const id=e.target.dataset.del;if(!id)return;items=items.filter(x=>x.id!==id);render();});
$('#addBtn').onclick=()=>{items.push({id:uuid(),name:'',spec:'',quantity:1,unitPrice:0,amount:0});render();};
$('#confirmBtn').onclick=()=>{const invalid=items.find(x=>!clean(x.name)||num(x.quantity)<=0||num(x.unitPrice)<0);if(invalid){alert('모든 품목의 내용·수량·예상단가를 확인해 주세요.');return;}confirmed=true;$('#downloadBtn').disabled=false;$('#confirmState').textContent='확정 완료';$('#confirmState').className='good';setStatus('최종 확정되었습니다. 이제 화면의 값 그대로 원본 서식에 내려받을 수 있습니다.','success');};
$('#downloadBtn').onclick=()=>exportTemplate().catch(e=>alert(e.message));
$('#resetBtn').onclick=()=>{items=[];confirmed=false;$('#fileInput').value='';$('#resultCard').classList.add('hidden');setStatus('새 견적서를 업로드해 주세요.','info');};
