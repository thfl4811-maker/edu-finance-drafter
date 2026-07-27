import './style.css';
import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { createWorker } from 'tesseract.js';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

const TEMPLATE_URL = '/품목내역(통합) (2).xls';
const supported = '.xlsx,.xls,.csv,.pdf,.png,.jpg,.jpeg,.webp';

let items = [];
let sourceFileName = '';
let ocrWorker = null;
let documentTotals = { supply: 0, tax: 0, total: 0 };

const money = (n) => Number(n || 0).toLocaleString('ko-KR');
const num = (v) => {
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  const cleaned = String(v ?? '').replace(/[^0-9.-]/g, '');
  return cleaned ? Number(cleaned) || 0 : 0;
};
const clean = (v) => String(v ?? '').replace(/\s+/g, ' ').trim();
const escapeHtml = (s) => String(s ?? '').replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));

const app = document.querySelector('#app');
app.innerHTML = `
<div class="app-shell">
  <header class="hero"><div class="hero-inner">
    <div class="badge">학교 행정업무 도우미 · 에듀파인 연계</div>
    <h1>품의서 작성기</h1>
    <p>견적서 한 번으로 에듀파인 업로드용 품목내역 엑셀과 물품 구입 품의서 초안을 함께 만들어요.</p>
  </div></header>
  <main>
    <div class="tabs">
      <button class="tab active" data-tab="purchase">① 물품 구입 품의</button>
      <button class="tab" data-tab="meeting">② 업무추진비·협의회 품의</button>
    </div>

    <section id="purchase" class="panel active">
      <div class="card">
        <h2>견적서를 업로드해 주세요</h2>
        <p class="sub">엑셀·PDF·사진 견적서에서 품목명, 규격, 수량, 단가를 찾아 표준 형식으로 변환합니다.</p>
        <div class="action-row">
          <button id="uploadBtn" class="btn btn-primary">견적서 업로드</button>
          <button id="blankBtn" class="btn btn-secondary">에듀파인 빈 서식 내려받기</button>
          <input id="fileInput" class="hidden" type="file" accept="${supported}">
        </div>
        <div id="dropzone" class="dropzone">
          <strong>파일을 이곳에 끌어다 놓아도 됩니다</strong>
          <small>지원 형식: XLSX, XLS, CSV, PDF, PNG, JPG, WEBP<br>스캔 PDF와 사진은 문자 인식으로 분석합니다.</small>
        </div>
        <div id="status" class="status"></div>
        <div id="progress" class="progress"><div></div></div>
      </div>

      <div id="resultCard" class="card hidden">
        <div class="section-head"><div><h2>품목 변환 결과</h2><p class="sub" id="sourceInfo"></p></div><button id="addRowBtn" class="btn btn-soft">행 추가</button></div>
        <div class="table-wrap"><table>
          <thead><tr><th>순번</th><th>품목명</th><th>규격</th><th>단위</th><th>수량</th><th>단가<br><small>(부가세 포함)</small></th><th>예상금액</th><th></th></tr></thead>
          <tbody id="itemBody"></tbody>
        </table></div>
        <div class="summary-grid">
          <div class="summary-box"><span>품목 수</span><strong id="itemCount">0개</strong></div>
          <div class="summary-box"><span>예상 합계</span><strong id="totalAmount">0원</strong></div>
          <div class="summary-box"><span>검산 결과</span><strong id="checkResult">정상</strong></div>
        </div>
        <div class="action-row" style="margin-top:16px">
          <button id="downloadBtn" class="btn btn-primary">에듀파인 품목내역.xls 내려받기</button>
          <button id="resetBtn" class="btn btn-soft">다른 견적서 불러오기</button>
        </div>
        <p class="note">최종 파일은 등록된 에듀파인 서식의 열 순서와 제목을 유지합니다. S2B물품번호는 비워 두고, 수수료는 0으로 입력합니다.</p>
      </div>

      <div id="draftCard" class="card hidden">
        <h2>물품 구입 품의서 작성</h2>
        <p class="sub">견적서 품목을 바탕으로 제목과 본문을 자동 작성합니다.</p>
        <div class="form-grid">
          <div class="field full"><label for="purpose">구입 목적</label><input id="purpose" placeholder="예: 치유텃밭 요리교실 운영"></div>
          <div class="field"><label for="budget">예산 과목</label><input id="budget" placeholder="예: 교육운영비"></div>
          <div class="field"><label for="purchaseMethod">구입 방법</label><select id="purchaseMethod"><option>카드 결제</option><option>학교장터(S2B)</option><option>계좌이체</option><option>기타</option></select></div>
          <div class="field"><label for="related">관련 문서</label><input id="related" placeholder="없으면 비워 두세요"></div>
          <div class="field"><label for="attachment">붙임</label><input id="attachment" value="견적서 1부"></div>
        </div>
        <div class="action-row" style="margin-top:16px"><button id="makeDraftBtn" class="btn btn-primary">품의서 초안 만들기</button></div>
      </div>

      <div id="draftResultCard" class="card hidden">
        <div class="section-head"><h2>품의서 초안</h2><div class="action-row"><button id="copyDraftBtn" class="btn btn-secondary">본문 복사</button></div></div>
        <div id="draftPreview" class="preview"></div>
      </div>
    </section>

    <section id="meeting" class="panel">
      <div class="card">
        <h2>업무추진비·협의회 품의</h2>
        <p class="sub">협의회 기본 정보와 참석 인원을 입력하면 금액을 계산하고 공문서 형식의 본문을 만듭니다.</p>
        <div class="form-grid">
          <div class="field full"><label for="meetingName">협의회명</label><input id="meetingName" placeholder="예: 학교급식 운영 협의회"></div>
          <div class="field"><label for="meetingDate">일시</label><input id="meetingDate" placeholder="예: 2026. 7. 28.(화) 16:00"></div>
          <div class="field"><label for="meetingPlace">장소</label><input id="meetingPlace" placeholder="예: 교무실"></div>
          <div class="field full"><label for="agenda">협의사항</label><textarea id="agenda" placeholder="예: 자율선택급식 운영 방안 협의 외 2건"></textarea></div>
          <div class="field"><label for="attendeeName">참석자 표기</label><input id="attendeeName" placeholder="예: 김소리 외 4명"></div>
          <div class="field"><label for="attendeeCount">참석 인원</label><input id="attendeeCount" type="number" min="1" value="1"></div>
          <div class="field"><label for="perPerson">1인당 단가</label><input id="perPerson" type="number" min="0" step="100" placeholder="15000"></div>
          <div class="field"><label for="meetingRelated">관련 문서</label><input id="meetingRelated" placeholder="없으면 비워 두세요"></div>
        </div>
        <div class="action-row" style="margin-top:16px"><button id="makeMeetingBtn" class="btn btn-primary">협의회 품의서 만들기</button></div>
      </div>
      <div id="meetingResultCard" class="card hidden">
        <div class="section-head"><h2>협의회 품의서 초안</h2><button id="copyMeetingBtn" class="btn btn-secondary">본문 복사</button></div>
        <div id="meetingPreview" class="preview"></div>
      </div>
    </section>
  </main>
</div>`;

const $ = (s) => document.querySelector(s);
const fileInput = $('#fileInput');
const statusBox = $('#status');
const progressBox = $('#progress');
const progressBar = $('#progress > div');

function setStatus(message, isError = false) {
  statusBox.textContent = message;
  statusBox.className = `status show${isError ? ' error' : ''}`;
}
function setProgress(pct) {
  progressBox.classList.toggle('show', pct >= 0 && pct < 100);
  progressBar.style.width = `${Math.max(0, Math.min(100, pct))}%`;
  if (pct >= 100) setTimeout(() => progressBox.classList.remove('show'), 400);
}

async function downloadTemplate() {
  const res = await fetch(TEMPLATE_URL);
  if (!res.ok) throw new Error('빈 서식 파일을 불러오지 못했습니다.');
  const blob = await res.blob();
  saveBlob(blob, '품목내역(통합).xls');
}

function saveBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

async function handleFile(file) {
  if (!file) return;
  sourceFileName = file.name;
  setStatus(`${file.name} 분석을 시작합니다.`);
  setProgress(5);
  try {
    const ext = file.name.split('.').pop().toLowerCase();
    let parsed = [];
    if (['xlsx','xls','csv'].includes(ext)) parsed = await parseSpreadsheet(file);
    else if (ext === 'pdf') parsed = await parsePdf(file);
    else if (['png','jpg','jpeg','webp'].includes(ext)) parsed = await parseImage(file);
    else throw new Error('지원하지 않는 파일 형식입니다.');

    items = normalizeItems(parsed, documentTotals).filter(i => i.name && !isSummaryLabel(i.name) && (i.quantity || i.unitPrice || i.amount));
    if (!items.length) throw new Error('품목을 자동으로 찾지 못했습니다. 표에서 품목명·규격·수량·단가가 보이도록 다시 저장하거나 선명한 파일을 올려 주세요.');
    renderItems();
    $('#resultCard').classList.remove('hidden');
    $('#draftCard').classList.remove('hidden');
    $('#sourceInfo').textContent = `원본 파일: ${sourceFileName}`;
    setProgress(100);
    setStatus(`${items.length}개 품목을 찾았습니다. 내용을 확인한 뒤 내려받으세요.`);
    $('#resultCard').scrollIntoView({behavior:'smooth', block:'start'});
  } catch (err) {
    console.error(err);
    setProgress(100);
    setStatus(err.message || '파일 분석 중 오류가 발생했습니다.', true);
  }
}

async function parseSpreadsheet(file) {
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, {type:'array', cellDates:false});
  let best = [];
  let bestScore = -1;
  const allSheets = [];
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], {header:1, defval:'', raw:false});
    allSheets.push(rows);
    const parsed = tableRowsToItems(rows);
    const score = parsed.length * 10 + parsed.reduce((s, i) => s + (i.name ? 2 : 0) + (i.quantity ? 1 : 0) + (i.unitPrice ? 1 : 0), 0);
    if (score > bestScore) { bestScore = score; best = parsed; }
  }
  documentTotals = detectDocumentTotals(allSheets);
  setProgress(78);
  return best;
}

function compactLabel(v) {
  return clean(v).toLowerCase().replace(/[\s·ㆍ:_()\[\]-]/g, '');
}

function isSummaryLabel(v) {
  const s = compactLabel(v);
  return /^(합계|총계|소계|공급가액|부가세|세액|vat|합계금액|총금액|견적금액|금액합계|계)$/.test(s)
    || /^(합계|총계|소계|공급가액|부가세|세액|합계금액|총금액)/.test(s);
}

function nextNumber(row, startIndex) {
  for (let i = startIndex + 1; i < row.length; i++) {
    const n = numericValue(row[i]);
    if (n !== null && n > 0) return n;
  }
  return 0;
}

function detectDocumentTotals(sheetRows) {
  let supply = 0, tax = 0, total = 0;
  const explicitTotals = [];
  const currencyTotals = [];
  for (const rows of sheetRows) {
    for (const row of rows) {
      const joined = row.map(clean).join(' ');
      row.forEach((cell, idx) => {
        const label = compactLabel(cell);
        if (label === '공급가액') supply = Math.max(supply, nextNumber(row, idx));
        if (label === '부가세' || label === '세액' || label === 'vat') tax = Math.max(tax, nextNumber(row, idx));
        if (['합계금액','총금액','견적금액','합계','총계'].includes(label)) {
          const n = nextNumber(row, idx);
          if (n) explicitTotals.push(n);
        }
      });
      if (/합\s*계\s*금\s*액|一金|원整|견적\s*금액/i.test(joined)) {
        const nums = joined.match(/[\d,]{4,}/g) || [];
        nums.map(num).filter(n => n > 0).forEach(n => currencyTotals.push(n));
      }
    }
  }
  if (explicitTotals.length) total = Math.max(...explicitTotals);
  if (!total && currencyTotals.length) total = Math.max(...currencyTotals);
  if (!total && supply && tax) total = supply + tax;
  if (!tax && total && supply && total > supply) tax = total - supply;
  return {supply, tax, total};
}

function headerScore(value, type) {
  const s = clean(value).toLowerCase().replace(/\s/g,'');
  const dict = {
    name:['품명','품목명','상품명','제품명','내역','내용','물품명','자재명','item','description'],
    spec:['규격','사양','용량','포장','포장단위','spec','size'],
    unit:['단위','unit'],
    qty:['수량','주문수량','요청수량','qty','quantity'],
    price:['단가','예상단가','공급단가','판매가','unitprice','price'],
    amount:['금액','예상금액','공급가액','합계','amount','total']
  };
  return dict[type].some(k => s.includes(k)) ? 1 : 0;
}

function tableRowsToItems(rows) {
  if (!rows?.length) return [];
  let bestHeader = -1, map = null, bestScore = 0;
  for (let r=0; r<Math.min(rows.length, 40); r++) {
    const row = rows[r];
    const m = {};
    ['name','spec','unit','qty','price','amount'].forEach(type => {
      const idx = row.findIndex(v => headerScore(v, type));
      if (idx >= 0) m[type] = idx;
    });
    const score = ['name','qty','price','amount'].filter(k => m[k] !== undefined).length + (m.spec !== undefined ? .5 : 0) + (m.unit !== undefined ? .25 : 0);
    if (score > bestScore) { bestScore = score; bestHeader = r; map = m; }
  }
  if (bestScore < 2) return inferByNumericPattern(rows);
  const out = [];
  for (let r=bestHeader+1; r<rows.length; r++) {
    const row = rows[r];
    const name = clean(row[map.name]);
    if (!name || isSummaryLabel(name)) continue;
    const q = map.qty !== undefined ? num(row[map.qty]) : 0;
    const p = map.price !== undefined ? num(row[map.price]) : 0;
    const a = map.amount !== undefined ? num(row[map.amount]) : 0;
    if (!(q || p || a)) continue;
    if (/^(비고|특기사항|계좌번호|사업자|주소|연락처)/i.test(name)) continue;
    out.push({name, spec:clean(row[map.spec]), unit:clean(row[map.unit]), quantity:q || 1, unitPrice:p, amount:a || q*p});
  }
  return out;
}

function inferByNumericPattern(rows) {
  const out = [];
  for (const row of rows) {
    const vals = row.map(clean).filter(Boolean);
    if (vals.length < 3) continue;
    const numeric = vals.map((v,i)=>({i,v:numericValue(v)})).filter(x=>x.v!==null);
    if (numeric.length < 2) continue;
    const text = vals.filter(v=>numericValue(v)===null && !/^(no|번호|순번)$/i.test(v));
    if (!text.length) continue;
    const quantity = numeric.length >= 3 ? numeric[numeric.length-3].v : 1;
    const unitPrice = numeric[numeric.length-2].v;
    const amount = numeric[numeric.length-1].v;
    if (amount < unitPrice || unitPrice <= 0) continue;
    out.push({name:text[0], spec:text.slice(1).join(' '), unit:'', quantity, unitPrice, amount});
  }
  return out;
}
function numericValue(v) {
  const s = String(v).trim();
  if (!/^-?[₩￦]?\s*[\d,]+(?:\.\d+)?\s*(원)?$/.test(s)) return null;
  return num(s);
}

async function parsePdf(file) {
  const data = new Uint8Array(await file.arrayBuffer());
  const pdf = await pdfjsLib.getDocument({data}).promise;
  const textRows = [];
  let textChars = 0;
  for (let p=1; p<=pdf.numPages; p++) {
    setProgress(10 + Math.round((p/pdf.numPages)*35));
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    textChars += content.items.reduce((s,i)=>s+(i.str||'').length,0);
    textRows.push(...pdfTextToRows(content.items));
  }
  let parsed = tableRowsToItems(textRows);
  if (parsed.length) return parsed;
  if (textChars < 80 || !parsed.length) {
    setStatus('PDF 표를 이미지 문자 인식으로 다시 분석하고 있습니다. 잠시 기다려 주세요.');
    const ocrRows = [];
    for (let p=1; p<=pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const viewport = page.getViewport({scale:2});
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width; canvas.height = viewport.height;
      await page.render({canvasContext:canvas.getContext('2d'), viewport}).promise;
      const text = await recognizeCanvas(canvas, 45 + Math.round((p/pdf.numPages)*45));
      ocrRows.push(...plainTextToRows(text));
    }
    parsed = tableRowsToItems(ocrRows);
  }
  return parsed;
}

function pdfTextToRows(items) {
  const lines = new Map();
  for (const item of items) {
    const y = Math.round((item.transform?.[5] || 0) / 4) * 4;
    if (!lines.has(y)) lines.set(y, []);
    lines.get(y).push({x:item.transform?.[4]||0, text:item.str||''});
  }
  return [...lines.entries()].sort((a,b)=>b[0]-a[0]).map(([,cells]) => cells.sort((a,b)=>a.x-b.x).map(c=>c.text));
}

async function parseImage(file) {
  const img = await createImageBitmap(file);
  const canvas = document.createElement('canvas');
  const scale = Math.min(2, 2400 / Math.max(img.width, img.height));
  canvas.width = Math.round(img.width*scale); canvas.height = Math.round(img.height*scale);
  canvas.getContext('2d').drawImage(img,0,0,canvas.width,canvas.height);
  const text = await recognizeCanvas(canvas, 80);
  return tableRowsToItems(plainTextToRows(text));
}

async function recognizeCanvas(canvas, targetProgress=90) {
  if (!ocrWorker) {
    ocrWorker = await createWorker('kor+eng', 1, {logger:m => {
      if (m.status === 'recognizing text') setProgress(45 + Math.round(m.progress * (targetProgress-45)));
    }});
  }
  const {data:{text}} = await ocrWorker.recognize(canvas);
  return text;
}

function plainTextToRows(text) {
  return String(text).split(/\r?\n/).map(line => line.trim()).filter(Boolean).map(line => line.split(/\t|\s{2,}/).filter(Boolean));
}

function normalizeItems(raw, totals = {}) {
  const prepared = raw
    .filter(r => !isSummaryLabel(r.name))
    .map((r,idx) => {
      const quantity = num(r.quantity) || 1;
      const supplyUnitPrice = num(r.unitPrice);
      let supplyAmount = num(r.amount);
      if (!supplyAmount && quantity && supplyUnitPrice) supplyAmount = quantity * supplyUnitPrice;
      let spec = clean(r.spec);
      let unit = clean(r.unit);
      if (!unit) {
        const m = spec.match(/(?:^|\s)(개|매|권|대|봉|병|통|팩|박스|상자|세트|kg|g|L|ml|m|롤|포|조|건)(?:$|\s)/i);
        if (m) unit = m[1];
      }
      return {id:crypto.randomUUID?.() || `${Date.now()}-${idx}`, name:clean(r.name), spec, unit, quantity, supplyUnitPrice, supplyAmount};
    });

  const supplySum = prepared.reduce((s,i)=>s+num(i.supplyAmount),0);
  const targetTotal = num(totals.total);
  let ratio = 1;
  if (targetTotal > supplySum && supplySum > 0) {
    const candidate = targetTotal / supplySum;
    if (candidate >= 1 && candidate <= 1.2) ratio = candidate;
  } else if (num(totals.tax) > 0 && num(totals.supply) > 0) {
    const candidate = (num(totals.supply) + num(totals.tax)) / num(totals.supply);
    if (candidate >= 1 && candidate <= 1.2) ratio = candidate;
  }

  let allocated = 0;
  return prepared.map((it, idx) => {
    let amount;
    if (ratio > 1 && targetTotal > 0 && idx === prepared.length - 1) {
      amount = targetTotal - allocated;
    } else {
      amount = Math.round(it.supplyAmount * ratio);
      allocated += amount;
    }
    const unitPrice = it.quantity ? Math.round(amount / it.quantity) : Math.round(it.supplyUnitPrice * ratio);
    amount = Math.round(unitPrice * it.quantity);
    if (ratio > 1 && targetTotal > 0 && idx === prepared.length - 1) {
      const diff = targetTotal - (allocated + amount);
      if (Math.abs(diff) <= Math.max(10, it.quantity)) amount += diff;
    }
    return {...it, unitPrice, amount, vat: Math.max(0, amount - it.supplyAmount)};
  });
}

function renderItems() {
  const body = $('#itemBody');
  body.innerHTML = items.map((it,i)=>`<tr data-id="${it.id}">
    <td class="money">${i+1}</td>
    <td><input data-key="name" value="${escapeHtml(it.name)}"></td>
    <td><input data-key="spec" value="${escapeHtml(it.spec)}"></td>
    <td><input data-key="unit" value="${escapeHtml(it.unit)}"></td>
    <td><input data-key="quantity" type="number" step="0.01" value="${it.quantity}"></td>
    <td><input data-key="unitPrice" type="number" step="1" value="${it.unitPrice}"></td>
    <td class="money">${money(it.amount)}원</td>
    <td class="row-actions"><button class="icon-btn" data-remove="${it.id}" title="삭제">×</button></td>
  </tr>`).join('');
  updateSummary();
}

function updateSummary() {
  let mismatch = 0;
  for (const it of items) {
    const calc = Math.round(num(it.quantity)*num(it.unitPrice));
    if (it.amount && Math.abs(calc-num(it.amount)) > 1) mismatch++;
    it.amount = calc || num(it.amount);
  }
  $('#itemCount').textContent = `${items.length}개`;
  $('#totalAmount').textContent = `${money(items.reduce((s,i)=>s+num(i.amount),0))}원`;
  $('#checkResult').textContent = mismatch ? `${mismatch}건 확인 필요` : '정상';
}

async function exportEdufine() {
  const filtered = items.filter(
    (it) => it.name && !isSummaryLabel(it.name)
  );

  if (!filtered.length) {
    throw new Error('내려받을 품목이 없습니다.');
  }

  const res = await fetch(encodeURI(TEMPLATE_URL), { cache: 'no-store' });
  if (!res.ok) {
    throw new Error(`에듀파인 원본 서식을 불러오지 못했습니다. (${res.status})`);
  }

  const buffer = await res.arrayBuffer();
  const wb = XLSX.read(buffer, {
    type: 'array',
    cellStyles: true,
    cellFormula: true,
    cellNF: true,
    cellDates: true,
    bookVBA: true
  });

  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) {
    throw new Error('에듀파인 원본 서식의 시트를 찾지 못했습니다.');
  }

  // 사용자가 제공한 에듀파인 원본 양식의 고정 입력 위치
  // 4행부터 품목을 입력하고 A~I열을 그대로 사용한다.
  const START_ROW = 4;

  // 기존 품목 입력값만 삭제한다.
  // 제목·안내문·셀 병합·열 너비·서식은 건드리지 않는다.
  const currentRange = XLSX.utils.decode_range(ws['!ref'] || 'A1:I4');
  const clearEndRow = Math.max(currentRange.e.r + 1, START_ROW + 300);

  for (let row = START_ROW; row <= clearEndRow; row++) {
    for (let col = 0; col <= 8; col++) {
      const addr = XLSX.utils.encode_cell({ r: row - 1, c: col });
      const cell = ws[addr];

      if (cell) {
        delete cell.v;
        delete cell.w;
        delete cell.f;
        cell.t = 'z';
      }
    }
  }

  // 4행의 서식을 각 품목 행에 복사한다.
  filtered.forEach((it, index) => {
    const rowNumber = START_ROW + index;
    const rowIndex = rowNumber - 1;

    const values = [
      index + 1,                              // A 순번
      clean(it.name),                         // B 내용
      '',                                     // C S2B물품번호
      clean(it.spec),                         // D 규격
      clean(it.unit) || '개',                 // E 단위
      num(it.quantity),                       // F 수량
      Math.round(num(it.unitPrice)),           // G 예상단가(부가세 포함)
      Math.round(num(it.quantity) * num(it.unitPrice)), // H 예상금액
      0                                       // I 수수료
    ];

    values.forEach((value, colIndex) => {
      const addr = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      const styleAddr = XLSX.utils.encode_cell({
        r: START_ROW - 1,
        c: colIndex
      });

      const styleCell = ws[styleAddr] || {};
      ws[addr] = {
        ...styleCell,
        v: value,
        t: typeof value === 'number' ? 'n' : 's'
      };

      delete ws[addr].w;
      delete ws[addr].f;

      if ([0, 5, 6, 7, 8].includes(colIndex)) {
        ws[addr].z = styleCell.z || '#,##0';
      }
    });
  });

  const lastRowIndex = START_ROW - 1 + filtered.length - 1;
  const outRange = XLSX.utils.decode_range(ws['!ref'] || 'A1:I4');
  outRange.e.r = Math.max(outRange.e.r, lastRowIndex);
  outRange.e.c = Math.max(outRange.e.c, 8);
  ws['!ref'] = XLSX.utils.encode_range(outRange);

  const out = XLSX.write(wb, {
    bookType: 'biff8',
    type: 'array',
    cellStyles: true,
    bookVBA: true
  });

  saveBlob(
    new Blob([out], { type: 'application/vnd.ms-excel' }),
    '품목내역(통합).xls'
  );
}
function buildPurchaseDraft() {
  if (!items.length) return;
  const purpose = clean($('#purpose').value) || '업무 운영';
  const budget = clean($('#budget').value);
  const method = clean($('#purchaseMethod').value);
  const related = clean($('#related').value);
  const attachment = clean($('#attachment').value) || '견적서 1부';
  const total = items.reduce((s,i)=>s+num(i.amount),0);
  const first = items[0]?.name || '물품';
  const extra = Math.max(0, items.length-1);
  const listText = extra ? `${first} 외 ${extra}종` : first;
  const title = `${purpose} 물품 구입`;
  const lines = [];
  lines.push(`제목: ${title}`, '');
  if (related) lines.push(`1. 관련: ${related}`, '');
  lines.push(`${related ? '2' : '1'}. ${purpose}에 필요한 물품을 다음과 같이 구입하고자 합니다.`, '');
  lines.push(`  가. 구입 목적: ${purpose}`);
  lines.push(`  나. 구입 품목: ${listText}`);
  lines.push(`  다. 소요예산: 금${money(total)}원`);
  lines.push(`  라. 산출내역: 붙임 견적서 참조`);
  if (budget) lines.push(`  마. 예산 과목: ${budget}`);
  lines.push(`  ${budget ? '바' : '마'}. 구입 방법: ${method}`, '');
  lines.push(`붙임  ${attachment}.  끝.`);
  const draft = lines.join('\n');
  $('#draftPreview').textContent = draft;
  $('#draftResultCard').classList.remove('hidden');
  $('#draftResultCard').scrollIntoView({behavior:'smooth', block:'start'});
}

function buildMeetingDraft() {
  const name = clean($('#meetingName').value) || '업무 협의회';
  const date = clean($('#meetingDate').value) || '(일시 입력)';
  const place = clean($('#meetingPlace').value) || '(장소 입력)';
  const agenda = clean($('#agenda').value) || '(협의사항 입력)';
  const attendee = clean($('#attendeeName').value) || '(참석자 입력)';
  const count = Math.max(1, num($('#attendeeCount').value));
  const per = num($('#perPerson').value);
  const related = clean($('#meetingRelated').value);
  const total = count*per;
  const lines = [`제목: ${name} 실시`, ''];
  if (related) lines.push(`1. 관련: ${related}`, '');
  lines.push(`${related ? '2' : '1'}. ${name}를 아래와 같이 실시하고자 합니다.`, '');
  lines.push(`  가. 일시: ${date}`);
  lines.push(`  나. 장소: ${place}`);
  lines.push(`  다. 협의사항: ${agenda}`);
  lines.push(`  라. 참석자: ${attendee}, 총 ${count}명`);
  lines.push(`  마. 소요예산: 금${money(total)}원`);
  lines.push(`  바. 산출내역: ${money(per)}원 × ${count}명`, '');
  lines.push('붙임  지출품의서 1부.  끝.');
  $('#meetingPreview').textContent = lines.join('\n');
  $('#meetingResultCard').classList.remove('hidden');
  $('#meetingResultCard').scrollIntoView({behavior:'smooth', block:'start'});
}

async function copyText(text, button) {
  await navigator.clipboard.writeText(text);
  const old = button.textContent; button.textContent = '복사 완료'; setTimeout(()=>button.textContent=old,1400);
}

// Events
$('.tabs').addEventListener('click', e => {
  const btn = e.target.closest('.tab'); if (!btn) return;
  document.querySelectorAll('.tab').forEach(x=>x.classList.toggle('active',x===btn));
  document.querySelectorAll('.panel').forEach(x=>x.classList.toggle('active',x.id===btn.dataset.tab));
});
$('#uploadBtn').onclick = () => fileInput.click();
$('#blankBtn').onclick = () => downloadTemplate().catch(e=>setStatus(e.message,true));
fileInput.onchange = () => handleFile(fileInput.files[0]);
const dropzone = $('#dropzone');
['dragenter','dragover'].forEach(ev=>dropzone.addEventListener(ev,e=>{e.preventDefault();dropzone.classList.add('dragover')}));
['dragleave','drop'].forEach(ev=>dropzone.addEventListener(ev,e=>{e.preventDefault();dropzone.classList.remove('dragover')}));
dropzone.addEventListener('drop',e=>handleFile(e.dataTransfer.files[0]));
$('#itemBody').addEventListener('input', e => {
  const tr=e.target.closest('tr'); if(!tr) return;
  const it=items.find(x=>x.id===tr.dataset.id); if(!it) return;
  const k=e.target.dataset.key; it[k]=['quantity','unitPrice'].includes(k)?num(e.target.value):e.target.value;
  it.amount=num(it.quantity)*num(it.unitPrice); tr.children[6].textContent=`${money(it.amount)}원`; updateSummary();
});
$('#itemBody').addEventListener('click', e => { const id=e.target.dataset.remove; if(!id)return; items=items.filter(x=>x.id!==id); renderItems(); });
$('#addRowBtn').onclick = () => { items.push(normalizeItems([{name:'',spec:'',unit:'',quantity:1,unitPrice:0,amount:0}])[0]); renderItems(); };
$('#downloadBtn').onclick = async () => {
  try {
    await exportEdufine();
  } catch (e) {
    console.error(e);
    setStatus(e.message || '에듀파인 파일을 만드는 중 오류가 발생했습니다.', true);
  }
};

$('#resetBtn').onclick = () => { fileInput.value=''; items=[]; $('#resultCard').classList.add('hidden'); $('#draftCard').classList.add('hidden'); $('#draftResultCard').classList.add('hidden'); setStatus('새 견적서를 업로드해 주세요.'); };
$('#makeDraftBtn').onclick = buildPurchaseDraft;
$('#copyDraftBtn').onclick = e => copyText($('#draftPreview').textContent,e.currentTarget);
$('#makeMeetingBtn').onclick = buildMeetingDraft;
$('#copyMeetingBtn').onclick = e => copyText($('#meetingPreview').textContent,e.currentTarget);
