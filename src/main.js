import './style.css';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
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
let documentTotals = { supply: 0, tax: 0, total: 0 };
let vatMode = 'auto';

const $ = (selector) => document.querySelector(selector);
const clean = (value) => String(value ?? '').replace(/\s+/g, ' ').trim();
const num = (value) => {
  const n = Number(String(value ?? '').replace(/[^0-9.-]/g, ''));
  return Number.isFinite(n) ? n : 0;
};
const won = (value) => `${Math.round(num(value)).toLocaleString('ko-KR')}원`;
const esc = (value) => String(value ?? '').replace(/[&<>"']/g, (char) => ({
  '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
}[char]));
const uuid = () => crypto.randomUUID?.() || `${Date.now()}-${Math.random()}`;

const app = $('#app');
app.innerHTML = `
<header class="hero">
  <div class="hero-inner">
    <span class="badge">API 키 없이 브라우저에서 처리</span>
    <h1>품의서 작성기</h1>
    <p>견적서를 확인 가능한 표로 바꾸고, 사용자가 확정한 값만 실제 에듀파인 품목내역(통합) 원본 서식에 입력합니다.</p>
  </div>
</header>

<main>
  <div class="tabs">
    <button class="tab active" data-tab="purchase">① 물품 구입 품의</button>
    <button class="tab" data-tab="meeting">② 업무추진비·협의회 품의</button>
  </div>

  <section id="purchase" class="panel active">
    <section class="card">
      <div class="step">1</div>
      <h2>견적서 업로드</h2>
      <p class="sub">엑셀·PDF·사진을 지원합니다. 자동 인식 결과는 다음 단계에서 직접 수정하고 확정할 수 있습니다.</p>
      <div class="actions">
        <button id="uploadBtn" class="btn primary">견적서 업로드</button>
        <button id="blankBtn" class="btn secondary">실제 에듀파인 빈 서식 내려받기</button>
        <input id="fileInput" type="file" accept="${ACCEPTED}" hidden>
      </div>
      <div id="dropzone" class="dropzone">
        <b>파일을 끌어다 놓아도 됩니다</b>
        <span>XLSX · XLS · CSV · PDF · PNG · JPG · WEBP</span>
      </div>
      <div id="status" class="status"></div>
    </section>

    <section id="resultCard" class="card hidden">
      <div class="head">
        <div>
          <div class="step">2</div>
          <h2>품목 확인 및 수정</h2>
          <p id="sourceInfo" class="sub"></p>
        </div>
        <button id="addBtn" class="btn soft">행 추가</button>
      </div>

      <div class="notice">
        합계·총계·공급가액·부가세 행은 품목에서 제외합니다. 자동 인식이 애매한 항목은 직접 고쳐 주세요.
      </div>

      <div class="vat-box">
        <div>
          <strong>부가세 처리</strong>
          <p id="vatHint">견적서의 공급가액·부가세·총액을 확인해 자동 판단합니다.</p>
        </div>
        <select id="vatMode">
          <option value="auto">자동 판단</option>
          <option value="included">현재 단가가 부가세 포함</option>
          <option value="excluded">현재 단가가 부가세 별도 → 10% 더하기</option>
        </select>
      </div>

      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>순번</th>
              <th>내용</th>
              <th>규격</th>
              <th>수량</th>
              <th>예상단가<br><small>부가세 포함 최종값</small></th>
              <th>예상금액</th>
              <th></th>
            </tr>
          </thead>
          <tbody id="rows"></tbody>
        </table>
      </div>

      <div class="summary">
        <div><span>품목 수</span><strong id="count">0개</strong></div>
        <div><span>확정 총금액</span><strong id="total">0원</strong></div>
        <div><span>확정 상태</span><strong id="confirmState" class="bad">미확정</strong></div>
      </div>

      <div class="actions top">
        <button id="confirmBtn" class="btn confirm">이 내용으로 최종 확정</button>
        <button id="downloadBtn" class="btn primary" disabled>품목내역(통합).xls 내려받기</button>
        <button id="resetBtn" class="btn secondary">다른 견적서 불러오기</button>
      </div>
      <p class="foot">다운로드할 때 견적서를 다시 분석하지 않습니다. 화면에서 최종 확정한 내용·규격·수량·예상단가만 실제 원본 서식에 입력됩니다.</p>
    </section>

    <section id="draftCard" class="card hidden">
      <div class="step">3</div>
      <h2>물품 구입 품의서 작성</h2>
      <p class="sub">확정한 품목과 총금액을 이용해 품의서 제목과 본문을 작성합니다.</p>
      <div class="form-grid">
        <label class="field full">
          <span>구입 목적</span>
          <input id="purpose" placeholder="예: 치유텃밭 요리교실 운영">
        </label>
        <label class="field">
          <span>예산 과목</span>
          <input id="budget" placeholder="예: 교육운영비">
        </label>
        <label class="field">
          <span>구입 방법</span>
          <select id="purchaseMethod">
            <option>카드 결제</option>
            <option>학교장터(S2B)</option>
            <option>계좌이체</option>
            <option>기타</option>
          </select>
        </label>
        <label class="field">
          <span>관련 문서</span>
          <input id="related" placeholder="없으면 비워 두세요">
        </label>
        <label class="field">
          <span>붙임</span>
          <input id="attachment" value="견적서 1부">
        </label>
      </div>
      <div class="actions top">
        <button id="makeDraftBtn" class="btn primary">물품 구입 품의서 초안 만들기</button>
      </div>
    </section>

    <section id="draftResultCard" class="card hidden">
      <div class="head">
        <h2>물품 구입 품의서 초안</h2>
        <button id="copyDraftBtn" class="btn secondary">본문 복사</button>
      </div>
      <pre id="draftPreview" class="preview"></pre>
    </section>
  </section>

  <section id="meeting" class="panel">
    <section class="card">
      <div class="step">1</div>
      <h2>업무추진비·협의회 품의</h2>
      <p class="sub">협의회 정보와 인원을 입력하면 금액을 계산하고 품의서 초안을 만듭니다.</p>
      <div class="form-grid">
        <label class="field full">
          <span>협의회명</span>
          <input id="meetingName" placeholder="예: 학교급식 운영 협의회">
        </label>
        <label class="field">
          <span>일시</span>
          <input id="meetingDate" placeholder="예: 2026. 7. 28.(화) 16:00">
        </label>
        <label class="field">
          <span>장소</span>
          <input id="meetingPlace" placeholder="예: 교무실">
        </label>
        <label class="field full">
          <span>협의사항</span>
          <textarea id="agenda" placeholder="예: 자율선택급식 운영 방안 협의 외 2건"></textarea>
        </label>
        <label class="field">
          <span>참석자 표기</span>
          <input id="attendeeName" placeholder="예: 김소리 외 4명">
        </label>
        <label class="field">
          <span>참석 인원</span>
          <input id="attendeeCount" type="number" min="1" value="1">
        </label>
        <label class="field">
          <span>1인당 단가</span>
          <input id="perPerson" type="number" min="0" step="100" placeholder="15000">
        </label>
        <label class="field">
          <span>관련 문서</span>
          <input id="meetingRelated" placeholder="없으면 비워 두세요">
        </label>
      </div>
      <div class="meeting-total">
        예상 총금액 <strong id="meetingTotal">0원</strong>
      </div>
      <div class="actions top">
        <button id="makeMeetingBtn" class="btn primary">협의회 품의서 만들기</button>
      </div>
    </section>

    <section id="meetingResultCard" class="card hidden">
      <div class="head">
        <h2>협의회 품의서 초안</h2>
        <button id="copyMeetingBtn" class="btn secondary">본문 복사</button>
      </div>
      <pre id="meetingPreview" class="preview"></pre>
    </section>
  </section>
</main>
`;

function setStatus(message, kind = 'info') {
  const box = $('#status');
  box.textContent = message;
  box.className = `status show ${kind}`;
}

function saveBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = name;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function compact(value) {
  return clean(value).toLowerCase().replace(/[\s·ㆍ:_()\[\]{}*※-]/g, '');
}

function isSummaryLabel(value) {
  const s = compact(value);
  return /^(합계|총계|소계|공급가액|부가세|세액|vat|합계금액|총금액|견적금액|금액합계|계)$/.test(s)
    || /^(합계|총계|소계|공급가액|부가세|세액|합계금액|총금액)/.test(s);
}

function headerType(value) {
  const s = compact(value);
  const dict = {
    name: ['품명', '품목명', '상품명', '제품명', '내역', '내용', '물품명', '자재명', 'description', 'item'],
    spec: ['규격', '사양', '용량', '포장', 'spec', 'size'],
    qty: ['수량', '주문수량', '요청수량', 'qty', 'quantity'],
    price: ['단가', '예상단가', '공급단가', '판매가', 'unitprice', 'price'],
    amount: ['금액', '예상금액', '공급가액', '합계금액', 'amount', 'total']
  };
  return Object.keys(dict).find((key) => dict[key].some((alias) => s.includes(compact(alias)))) || '';
}

function numericValue(value) {
  const s = String(value ?? '').trim();
  if (!/^-?[₩￦]?\s*[\d,]+(?:\.\d+)?\s*(원)?$/.test(s)) return null;
  return num(s);
}

function nextNumber(row, startIndex) {
  for (let index = startIndex + 1; index < row.length; index += 1) {
    const value = numericValue(row[index]);
    if (value !== null && value > 0) return value;
  }
  return 0;
}

function detectDocumentTotals(allSheetRows) {
  let supply = 0;
  let tax = 0;
  let total = 0;
  const explicitTotals = [];

  for (const rows of allSheetRows) {
    for (const row of rows) {
      row.forEach((cell, index) => {
        const label = compact(cell);
        if (label === '공급가액') supply = Math.max(supply, nextNumber(row, index));
        if (['부가세', '세액', 'vat'].includes(label)) tax = Math.max(tax, nextNumber(row, index));
        if (['합계', '총계', '합계금액', '총금액', '견적금액'].includes(label)) {
          const value = nextNumber(row, index);
          if (value > 0) explicitTotals.push(value);
        }
      });
    }
  }

  if (explicitTotals.length) total = Math.max(...explicitTotals);
  if (!total && supply && tax) total = supply + tax;
  if (!tax && total > supply && supply > 0) tax = total - supply;

  return { supply, tax, total };
}

function tableToItems(rows) {
  if (!rows?.length) return [];

  let headerRow = -1;
  let map = {};
  let bestScore = -1;

  for (let rowIndex = 0; rowIndex < Math.min(rows.length, 60); rowIndex += 1) {
    const found = {};
    rows[rowIndex].forEach((cell, columnIndex) => {
      const type = headerType(cell);
      if (type && found[type] === undefined) found[type] = columnIndex;
    });

    const score = ['name', 'qty', 'price', 'amount'].filter((key) => found[key] !== undefined).length
      + (found.spec !== undefined ? 0.5 : 0);

    if (score > bestScore) {
      bestScore = score;
      headerRow = rowIndex;
      map = found;
    }
  }

  if (bestScore < 2) return inferItems(rows);

  const result = [];
  for (let rowIndex = headerRow + 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    const name = clean(row[map.name]);

    if (!name || isSummaryLabel(name)) continue;
    if (/^(비고|특기사항|계좌번호|사업자|주소|연락처|대표자|상호|견적일자)/i.test(name)) continue;

    const quantity = map.qty !== undefined ? num(row[map.qty]) : 0;
    const price = map.price !== undefined ? num(row[map.price]) : 0;
    const amount = map.amount !== undefined ? num(row[map.amount]) : 0;

    if (!(quantity || price || amount)) continue;

    const q = quantity || 1;
    const p = price || (amount && q ? Math.round(amount / q) : 0);
    if (!p) continue;

    result.push({
      id: uuid(),
      name,
      spec: clean(row[map.spec]),
      quantity: q,
      sourceUnitPrice: p,
      unitPrice: p,
      amount: Math.round(q * p)
    });
  }
  return result;
}

function inferItems(rows) {
  const result = [];

  for (const row of rows) {
    const values = row.map(clean).filter(Boolean);
    if (values.length < 3) continue;

    const numeric = values
      .map((value, index) => ({ index, value: numericValue(value) }))
      .filter((entry) => entry.value !== null);

    const texts = values.filter((value) => numericValue(value) === null && !/^(no|번호|순번)$/i.test(value));
    if (numeric.length < 2 || !texts.length || isSummaryLabel(texts[0])) continue;

    let quantity = 1;
    let unitPrice = 0;
    let amount = 0;

    if (numeric.length >= 3) {
      quantity = numeric[numeric.length - 3].value;
      unitPrice = numeric[numeric.length - 2].value;
      amount = numeric[numeric.length - 1].value;
    } else {
      unitPrice = numeric[numeric.length - 2].value;
      amount = numeric[numeric.length - 1].value;
      quantity = amount >= unitPrice && unitPrice > 0
        ? Math.max(1, Math.round(amount / unitPrice))
        : 1;
    }

    if (unitPrice <= 0 || amount < unitPrice) continue;

    result.push({
      id: uuid(),
      name: texts[0],
      spec: texts.slice(1).join(' '),
      quantity,
      sourceUnitPrice: unitPrice,
      unitPrice,
      amount: Math.round(quantity * unitPrice)
    });
  }
  return result;
}

function rowsFromText(textValue) {
  return String(textValue)
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => line.split(/\t|\s{2,}/).filter(Boolean));
}

// 한셀(HCell)로 저장한 xlsx에는 hs: 전용 서식 태그가 들어 있어 SheetJS가 시트를 통째로 비워 버린다.
// 업로드 단계에서 그 태그만 지워 정상적인 xlsx로 만든 뒤 읽는다. (일반 엑셀 파일은 그대로 통과)
async function cleanHancellXlsx(arrayBuffer) {
  const head = new Uint8Array(arrayBuffer, 0, 2);
  if (head[0] !== 0x50 || head[1] !== 0x4b) return arrayBuffer; // zip(xlsx)이 아니면 건너뜀(.xls·.csv)
  try {
    const zip = await JSZip.loadAsync(arrayBuffer);
    let changed = false;
    const targets = Object.keys(zip.files).filter((p) => /^xl\/.+\.xml$/.test(p));
    for (const p of targets) {
      const xml = await zip.file(p).async('string');
      if (!xml.includes('<hs:')) continue;
      zip.file(p, xml
        .replace(/<hs:[^>]*\/>/g, '')
        .replace(/<hs:([\w-]+)[^>]*>[\s\S]*?<\/hs:\1>/g, ''));
      changed = true;
    }
    return changed ? await zip.generateAsync({ type: 'arraybuffer' }) : arrayBuffer;
  } catch (e) {
    return arrayBuffer;
  }
}

async function parseSpreadsheet(file) {
  const buffer = await cleanHancellXlsx(await file.arrayBuffer());
  const workbook = XLSX.read(buffer, { type: 'array', raw: false });
  let best = [];
  const allRows = [];

  for (const sheetName of workbook.SheetNames) {
    if (!workbook.Sheets[sheetName]) continue;
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      defval: '',
      raw: false
    });
    allRows.push(rows);
    const parsed = tableToItems(rows);
    if (parsed.length > best.length) best = parsed;
  }

  documentTotals = detectDocumentTotals(allRows);
  return best;
}

function pdfRows(contentItems) {
  const lines = new Map();
  for (const item of contentItems) {
    const y = Math.round((item.transform?.[5] || 0) / 4) * 4;
    if (!lines.has(y)) lines.set(y, []);
    lines.get(y).push({ x: item.transform?.[4] || 0, text: item.str || '' });
  }

  return [...lines.entries()]
    .sort((a, b) => b[0] - a[0])
    .map(([, cells]) => cells.sort((a, b) => a.x - b.x).map((cell) => cell.text));
}

async function getWorker() {
  if (!ocrWorker) ocrWorker = await createWorker('kor+eng');
  return ocrWorker;
}

async function recognizeCanvas(canvas) {
  const worker = await getWorker();
  const { data: { text } } = await worker.recognize(canvas);
  return text;
}

async function parsePdf(file) {
  const pdf = await pdfjsLib.getDocument({
    data: new Uint8Array(await file.arrayBuffer())
  }).promise;

  let rows = [];
  let characterCount = 0;
  const pageLimit = Math.min(pdf.numPages, 10);

  for (let pageNumber = 1; pageNumber <= pageLimit; pageNumber += 1) {
    setStatus(`PDF ${pageNumber}/${pageLimit}쪽을 읽고 있습니다…`, 'loading');
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    characterCount += content.items.reduce((sum, item) => sum + (item.str || '').length, 0);
    rows.push(...pdfRows(content.items));
  }

  documentTotals = detectDocumentTotals([rows]);
  let parsed = tableToItems(rows);
  if (parsed.length) return parsed;

  if (characterCount < 100 || !parsed.length) {
    rows = [];
    const ocrPageLimit = Math.min(pdf.numPages, 6);

    for (let pageNumber = 1; pageNumber <= ocrPageLimit; pageNumber += 1) {
      setStatus(`스캔 PDF ${pageNumber}/${ocrPageLimit}쪽을 문자 인식 중입니다…`, 'loading');
      const page = await pdf.getPage(pageNumber);
      const viewport = page.getViewport({ scale: 2 });
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
      rows.push(...rowsFromText(await recognizeCanvas(canvas)));
    }

    documentTotals = detectDocumentTotals([rows]);
    parsed = tableToItems(rows);
  }

  return parsed;
}

async function parseImage(file) {
  setStatus('사진을 문자 인식 중입니다…', 'loading');
  const bitmap = await createImageBitmap(file);
  const scale = Math.min(2, 2400 / Math.max(bitmap.width, bitmap.height));
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(bitmap.width * scale);
  canvas.height = Math.round(bitmap.height * scale);
  canvas.getContext('2d').drawImage(bitmap, 0, 0, canvas.width, canvas.height);

  const rows = rowsFromText(await recognizeCanvas(canvas));
  documentTotals = detectDocumentTotals([rows]);
  return tableToItems(rows);
}

function normalize(list) {
  return list
    .filter((item) => item.name && !isSummaryLabel(item.name))
    .map((item) => ({
      id: item.id || uuid(),
      name: clean(item.name),
      spec: clean(item.spec),
      quantity: num(item.quantity) || 1,
      sourceUnitPrice: Math.round(num(item.sourceUnitPrice || item.unitPrice)),
      unitPrice: Math.round(num(item.unitPrice)),
      amount: Math.round((num(item.quantity) || 1) * num(item.unitPrice))
    }))
    .filter((item) => item.unitPrice > 0);
}

function autoVatNeedsAdding() {
  const { supply, tax, total } = documentTotals;
  if (supply > 0 && tax > 0 && total === supply + tax) return true;

  const sourceSum = items.reduce(
    (sum, item) => sum + num(item.quantity) * num(item.sourceUnitPrice),
    0
  );

  if (total > sourceSum && sourceSum > 0) {
    const ratio = total / sourceSum;
    return ratio >= 1.08 && ratio <= 1.12;
  }

  return false;
}

function applyVatMode(mode) {
  vatMode = mode;
  const shouldAdd = mode === 'excluded' || (mode === 'auto' && autoVatNeedsAdding());

  items = items.map((item) => {
    const source = num(item.sourceUnitPrice || item.unitPrice);
    const unitPrice = shouldAdd ? Math.round(source * 1.1) : Math.round(source);
    return {
      ...item,
      unitPrice,
      amount: Math.round(num(item.quantity) * unitPrice)
    };
  });

  updateVatHint(shouldAdd);
  renderRowsOnly();
  markDirty();
}

function updateVatHint(addingVat) {
  const { supply, tax, total } = documentTotals;
  let detail = '';

  if (supply || tax || total) {
    detail = ` 인식값: 공급가액 ${won(supply)}, 부가세 ${won(tax)}, 총액 ${won(total)}.`;
  }

  $('#vatHint').textContent = addingVat
    ? `부가세 별도로 판단해 각 단가에 10%를 더했습니다.${detail}`
    : `현재 단가를 부가세 포함 금액으로 사용합니다.${detail}`;
}

async function handleFile(file) {
  if (!file) return;

  confirmed = false;
  sourceName = file.name;
  documentTotals = { supply: 0, tax: 0, total: 0 };
  setStatus(`${file.name}을 분석하고 있습니다…`, 'loading');

  try {
    const extension = file.name.split('.').pop().toLowerCase();
    let parsed = [];

    if (['xlsx', 'xls', 'csv'].includes(extension)) parsed = await parseSpreadsheet(file);
    else if (extension === 'pdf') parsed = await parsePdf(file);
    else parsed = await parseImage(file);

    items = normalize(parsed);
    if (!items.length) {
      throw new Error('품목을 자동으로 찾지 못했습니다. 행 추가 버튼으로 직접 입력하거나 표가 선명한 파일을 올려 주세요.');
    }

    $('#vatMode').value = 'auto';
    applyVatMode('auto');

    $('#resultCard').classList.remove('hidden');
    $('#draftCard').classList.remove('hidden');
    $('#sourceInfo').textContent = `원본 파일: ${sourceName}`;
    setStatus(`${items.length}개 품목을 찾았습니다. 부가세 처리와 품목 값을 확인한 뒤 최종 확정하세요.`, 'success');
    $('#resultCard').scrollIntoView({ behavior: 'smooth' });
  } catch (error) {
    console.error(error);
    setStatus(error.message || '파일 분석 중 오류가 발생했습니다.', 'error');
  }
}

function markDirty() {
  confirmed = false;
  $('#downloadBtn').disabled = true;
  $('#confirmState').textContent = '미확정';
  $('#confirmState').className = 'bad';
}

function renderRowsOnly() {
  $('#rows').innerHTML = items.map((item, index) => `
    <tr data-id="${item.id}">
      <td class="center">${index + 1}</td>
      <td><input data-key="name" value="${esc(item.name)}"></td>
      <td><input data-key="spec" value="${esc(item.spec)}"></td>
      <td><input data-key="quantity" type="number" min="0" step="0.01" value="${item.quantity}"></td>
      <td><input data-key="unitPrice" type="number" min="0" step="1" value="${item.unitPrice}"></td>
      <td class="money">${won(item.amount)}</td>
      <td><button class="delete" data-delete="${item.id}" title="삭제">×</button></td>
    </tr>
  `).join('');

  updateSummary();
}

function updateSummary() {
  items.forEach((item) => {
    item.amount = Math.round(num(item.quantity) * num(item.unitPrice));
  });
  $('#count').textContent = `${items.length}개`;
  $('#total').textContent = won(items.reduce((sum, item) => sum + item.amount, 0));
}

function normalizeHeader(value) {
  return compact(value).replace(/필수/g, '');
}

function findTemplateLayout(worksheet) {
  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:D20');
  const aliases = {
    name: ['내용', '품명', '품목명'],
    spec: ['규격'],
    qty: ['수량'],
    price: ['예상단가', '단가']
  };

  for (let row = range.s.r; row <= Math.min(range.e.r, range.s.r + 100); row += 1) {
    const found = {};
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      const label = normalizeHeader(worksheet[address]?.v);
      if (!label) continue;

      for (const [key, names] of Object.entries(aliases)) {
        if (names.some((name) => {
          const target = normalizeHeader(name);
          return label === target || label.includes(target);
        })) {
          found[key] = col;
        }
      }
    }

    if (found.name !== undefined
      && found.spec !== undefined
      && found.qty !== undefined
      && found.price !== undefined) {
      return { headerRow: row, columns: found };
    }
  }

  // 사용자가 제공한 품목내역(통합) 서식의 기본 배치가 A:D인 경우의 안전한 대체값
  return {
    headerRow: 0,
    columns: { name: 0, spec: 1, qty: 2, price: 3 }
  };
}

async function downloadBlank() {
  const response = await fetch(encodeURI(TEMPLATE_URL), { cache: 'no-store' });
  if (!response.ok) throw new Error('실제 에듀파인 원본 서식을 불러오지 못했습니다.');
  saveBlob(await response.blob(), '품목내역(통합).xls');
}

async function exportTemplate() {
  if (!confirmed) throw new Error('먼저 “이 내용으로 최종 확정”을 눌러 주세요.');

  const finalItems = items.filter(
    (item) => clean(item.name) && num(item.quantity) > 0 && num(item.unitPrice) >= 0
  );
  if (!finalItems.length) throw new Error('다운로드할 품목이 없습니다.');

  const response = await fetch(encodeURI(TEMPLATE_URL), { cache: 'no-store' });
  if (!response.ok) {
    throw new Error(`실제 품목내역(통합) 원본 서식을 불러오지 못했습니다. (${response.status})`);
  }

  const workbook = XLSX.read(await response.arrayBuffer(), {
    type: 'array',
    cellStyles: true,
    cellNF: true,
    cellFormula: true,
    bookVBA: true
  });

  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!worksheet) throw new Error('품목내역(통합) 원본 서식의 시트를 찾지 못했습니다.');

  const { headerRow, columns } = findTemplateLayout(worksheet);
  const startRow = headerRow + 1;
  const originalRange = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:D1');
  const targetColumns = Object.values(columns);

  // 원본의 제목·열·서식은 유지하고 기존 품목 값만 지운다.
  for (let row = startRow; row <= Math.max(originalRange.e.r, startRow + 500); row += 1) {
    for (const col of targetColumns) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = worksheet[address];
      if (cell) {
        delete cell.v;
        delete cell.w;
        delete cell.f;
        cell.t = 'z';
      }
    }
  }

  finalItems.forEach((item, index) => {
    const row = startRow + index;
    const values = {
      name: clean(item.name),
      spec: clean(item.spec),
      qty: num(item.quantity),
      price: Math.round(num(item.unitPrice))
    };

    for (const [key, col] of Object.entries(columns)) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      const styleAddress = XLSX.utils.encode_cell({ r: startRow, c: col });
      const styleCell = worksheet[styleAddress] || worksheet[
        XLSX.utils.encode_cell({ r: headerRow, c: col })
      ] || {};
      const value = values[key];

      worksheet[address] = {
        ...styleCell,
        v: value,
        t: typeof value === 'number' ? 'n' : 's'
      };
      delete worksheet[address].w;
      delete worksheet[address].f;

      if (['qty', 'price'].includes(key)) {
        worksheet[address].z = styleCell.z || '#,##0';
      }
    }
  });

  originalRange.e.r = Math.max(originalRange.e.r, startRow + finalItems.length - 1);
  originalRange.e.c = Math.max(originalRange.e.c, ...targetColumns);
  worksheet['!ref'] = XLSX.utils.encode_range(originalRange);

  const output = XLSX.write(workbook, {
    bookType: 'biff8',
    type: 'array',
    cellStyles: true,
    bookVBA: true
  });

  saveBlob(
    new Blob([output], { type: 'application/vnd.ms-excel' }),
    '품목내역(통합).xls'
  );
}

function buildPurchaseDraft() {
  const finalItems = items.filter((item) => clean(item.name));
  if (!finalItems.length) {
    alert('품목을 먼저 입력해 주세요.');
    return;
  }

  const purpose = clean($('#purpose').value) || '업무 운영';
  const budget = clean($('#budget').value);
  const method = clean($('#purchaseMethod').value);
  const related = clean($('#related').value);
  const attachment = clean($('#attachment').value) || '견적서 1부';
  const total = finalItems.reduce((sum, item) => sum + num(item.amount), 0);
  const first = finalItems[0]?.name || '물품';
  const extra = Math.max(0, finalItems.length - 1);
  const itemText = extra ? `${first} 외 ${extra}종` : first;

  const lines = [`제목: ${purpose} 물품 구입`, ''];
  if (related) lines.push(`1. 관련: ${related}`, '');
  lines.push(`${related ? '2' : '1'}. ${purpose}에 필요한 물품을 다음과 같이 구입하고자 합니다.`, '');
  lines.push(`  가. 구입 목적: ${purpose}`);
  lines.push(`  나. 구입 품목: ${itemText}`);
  lines.push(`  다. 소요예산: 금${Math.round(total).toLocaleString('ko-KR')}원`);
  lines.push('  라. 산출내역: 붙임 견적서 참조');
  if (budget) lines.push(`  마. 예산 과목: ${budget}`);
  lines.push(`  ${budget ? '바' : '마'}. 구입 방법: ${method}`, '');
  lines.push(`붙임  ${attachment}.  끝.`);

  $('#draftPreview').textContent = lines.join('\n');
  $('#draftResultCard').classList.remove('hidden');
  $('#draftResultCard').scrollIntoView({ behavior: 'smooth' });
}

function updateMeetingTotal() {
  const count = Math.max(1, num($('#attendeeCount').value));
  const price = num($('#perPerson').value);
  $('#meetingTotal').textContent = won(count * price);
}

function buildMeetingDraft() {
  const name = clean($('#meetingName').value) || '업무 협의회';
  const date = clean($('#meetingDate').value) || '(일시 입력)';
  const place = clean($('#meetingPlace').value) || '(장소 입력)';
  const agenda = clean($('#agenda').value) || '(협의사항 입력)';
  const attendee = clean($('#attendeeName').value) || '(참석자 입력)';
  const count = Math.max(1, num($('#attendeeCount').value));
  const perPerson = num($('#perPerson').value);
  const related = clean($('#meetingRelated').value);
  const total = count * perPerson;

  const lines = [`제목: ${name} 실시`, ''];
  if (related) lines.push(`1. 관련: ${related}`, '');
  lines.push(`${related ? '2' : '1'}. ${name}를 아래와 같이 실시하고자 합니다.`, '');
  lines.push(`  가. 일시: ${date}`);
  lines.push(`  나. 장소: ${place}`);
  lines.push(`  다. 협의사항: ${agenda}`);
  lines.push(`  라. 참석자: ${attendee}, 총 ${count}명`);
  lines.push(`  마. 소요예산: 금${Math.round(total).toLocaleString('ko-KR')}원`);
  lines.push(`  바. 산출내역: ${Math.round(perPerson).toLocaleString('ko-KR')}원 × ${count}명`, '');
  lines.push('붙임  지출품의서 1부.  끝.');

  $('#meetingPreview').textContent = lines.join('\n');
  $('#meetingResultCard').classList.remove('hidden');
  $('#meetingResultCard').scrollIntoView({ behavior: 'smooth' });
}

async function copyText(text, button) {
  await navigator.clipboard.writeText(text);
  const oldText = button.textContent;
  button.textContent = '복사 완료';
  setTimeout(() => {
    button.textContent = oldText;
  }, 1400);
}

$('.tabs').addEventListener('click', (event) => {
  const button = event.target.closest('.tab');
  if (!button) return;

  document.querySelectorAll('.tab').forEach((tab) => {
    tab.classList.toggle('active', tab === button);
  });
  document.querySelectorAll('.panel').forEach((panel) => {
    panel.classList.toggle('active', panel.id === button.dataset.tab);
  });
});

$('#uploadBtn').onclick = () => $('#fileInput').click();
$('#blankBtn').onclick = () => downloadBlank().catch((error) => alert(error.message));
$('#fileInput').onchange = () => handleFile($('#fileInput').files[0]);

const dropzone = $('#dropzone');
['dragenter', 'dragover'].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.add('over');
  });
});
['dragleave', 'drop'].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.remove('over');
  });
});
dropzone.addEventListener('drop', (event) => handleFile(event.dataTransfer.files[0]));

$('#vatMode').onchange = (event) => applyVatMode(event.target.value);

$('#rows').addEventListener('input', (event) => {
  const row = event.target.closest('tr');
  if (!row) return;

  const item = items.find((entry) => entry.id === row.dataset.id);
  if (!item) return;

  const key = event.target.dataset.key;
  item[key] = ['quantity', 'unitPrice'].includes(key)
    ? num(event.target.value)
    : event.target.value;

  if (key === 'unitPrice') item.sourceUnitPrice = num(event.target.value);
  item.amount = Math.round(num(item.quantity) * num(item.unitPrice));
  row.children[5].textContent = won(item.amount);
  updateSummary();
  markDirty();
});

$('#rows').addEventListener('click', (event) => {
  const id = event.target.dataset.delete;
  if (!id) return;
  items = items.filter((item) => item.id !== id);
  renderRowsOnly();
  markDirty();
});

$('#addBtn').onclick = () => {
  items.push({
    id: uuid(),
    name: '',
    spec: '',
    quantity: 1,
    sourceUnitPrice: 0,
    unitPrice: 0,
    amount: 0
  });
  renderRowsOnly();
  markDirty();
};

$('#confirmBtn').onclick = () => {
  const invalid = items.find(
    (item) => !clean(item.name) || num(item.quantity) <= 0 || num(item.unitPrice) < 0
  );

  if (invalid) {
    alert('모든 품목의 내용·수량·예상단가를 확인해 주세요.');
    return;
  }

  confirmed = true;
  $('#downloadBtn').disabled = false;
  $('#confirmState').textContent = '확정 완료';
  $('#confirmState').className = 'good';
  setStatus('최종 확정되었습니다. 현재 화면의 값 그대로 원본 서식에 내려받습니다.', 'success');
};

$('#downloadBtn').onclick = () => exportTemplate().catch((error) => alert(error.message));

$('#resetBtn').onclick = () => {
  items = [];
  confirmed = false;
  documentTotals = { supply: 0, tax: 0, total: 0 };
  $('#fileInput').value = '';
  $('#resultCard').classList.add('hidden');
  $('#draftCard').classList.add('hidden');
  $('#draftResultCard').classList.add('hidden');
  setStatus('새 견적서를 업로드해 주세요.', 'info');
};

$('#makeDraftBtn').onclick = buildPurchaseDraft;
$('#copyDraftBtn').onclick = (event) => copyText(
  $('#draftPreview').textContent,
  event.currentTarget
);

['input', 'change'].forEach((eventName) => {
  $('#attendeeCount').addEventListener(eventName, updateMeetingTotal);
  $('#perPerson').addEventListener(eventName, updateMeetingTotal);
});

$('#makeMeetingBtn').onclick = buildMeetingDraft;
$('#copyMeetingBtn').onclick = (event) => copyText(
  $('#meetingPreview').textContent,
  event.currentTarget
);
