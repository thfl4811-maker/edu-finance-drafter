import './style.css';
import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

const TEMPLATE_URL = '/edufine-template.xls';
const accepted = '.xlsx,.xls,.csv,.pdf,.png,.jpg,.jpeg,.webp';

let items = [];
let sourceName = '';
let quoteTotal = 0;
let supplier = '';
let warnings = [];

const won = value => `${Math.round(Number(value || 0)).toLocaleString('ko-KR')}원`;
const number = value => {
  const n = Number(String(value ?? '').replace(/[^0-9.-]/g, ''));
  return Number.isFinite(n) ? n : 0;
};
const text = value => String(value ?? '').replace(/\s+/g, ' ').trim();
const escapeHtml = value => String(value ?? '').replace(/[&<>"']/g, c => ({
  '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
}[c]));

const app = document.querySelector('#app');
app.innerHTML = `
<header class="hero">
  <div class="hero__inner">
    <div class="eyebrow">AI 학교행정 업무도우미</div>
    <h1>품의서 작성기</h1>
    <p>견적서 형식이 달라도 AI가 실제 품목을 이해해 에듀파인 업로드 파일과 품의서 초안을 만듭니다.</p>
  </div>
</header>
<main class="container">
  <nav class="tabs" aria-label="품의 유형">
    <button class="tab active" data-target="purchase">물품 구입 품의</button>
    <button class="tab" data-target="meeting">업무추진비·협의회</button>
  </nav>

  <section id="purchase" class="panel active">
    <article class="card upload-card">
      <div class="step">1</div>
      <div>
        <h2>견적서를 업로드해 주세요</h2>
        <p class="muted">엑셀, PDF, 사진을 지원합니다. AI가 품목명·규격·수량·부가세 포함 단가를 문맥으로 판단합니다.</p>
      </div>
      <div class="button-row">
        <button id="chooseFile" class="button primary">견적서 업로드</button>
        <button id="blankTemplate" class="button secondary">에듀파인 빈 서식 내려받기</button>
        <input id="fileInput" type="file" accept="${accepted}" hidden>
      </div>
      <div id="dropzone" class="dropzone">
        <strong>파일을 이곳에 끌어다 놓으세요</strong>
        <span>XLSX · XLS · CSV · PDF · PNG · JPG · WEBP</span>
      </div>
      <div id="status" class="status" aria-live="polite"></div>
    </article>

    <article id="resultCard" class="card hidden">
      <div class="card-head">
        <div>
          <div class="step">2</div>
          <h2>AI가 찾은 실제 품목</h2>
          <p id="sourceInfo" class="muted"></p>
        </div>
        <button id="addItem" class="button ghost">행 추가</button>
      </div>
      <div id="warningBox" class="warning hidden"></div>
      <div class="table-scroll">
        <table>
          <thead>
            <tr><th>순번</th><th>내용</th><th>규격</th><th>수량</th><th>예상단가<br><small>부가세 포함</small></th><th>예상금액</th><th></th></tr>
          </thead>
          <tbody id="itemRows"></tbody>
        </table>
      </div>
      <div class="summary">
        <div><span>실제 품목</span><strong id="itemCount">0개</strong></div>
        <div><span>변환 합계</span><strong id="calculatedTotal">0원</strong></div>
        <div><span>견적서 최종 합계</span><strong id="originalTotal">확인 없음</strong></div>
        <div><span>검산</span><strong id="matchStatus">확인 중</strong></div>
      </div>
      <div class="button-row top-space">
        <button id="downloadEdufine" class="button primary">에듀파인 품목내역.xls 내려받기</button>
        <button id="startOver" class="button secondary">다른 견적서 분석</button>
      </div>
      <p class="footnote">에듀파인 원본 양식의 4개 열인 내용·규격·수량·예상단가에 화면의 품목만 입력합니다. 합계·부가세 행은 입력하지 않습니다.</p>
    </article>

    <article id="draftCard" class="card hidden">
      <div class="step">3</div>
      <h2>물품 구입 품의서 초안</h2>
      <div class="form-grid">
        <label class="full">구입 목적<input id="purpose" placeholder="예: 급식실 다지기 칼날 교체"></label>
        <label>예산 과목<input id="budget" placeholder="예: 급식운영비"></label>
        <label>구입 방법<select id="method"><option>카드 결제</option><option>학교장터(S2B)</option><option>계좌이체</option><option>기타</option></select></label>
        <label>관련 문서<input id="related" placeholder="없으면 비워 두세요"></label>
        <label>붙임<input id="attachment" value="견적서 1부"></label>
      </div>
      <button id="makeDraft" class="button primary top-space">품의서 초안 만들기</button>
    </article>

    <article id="draftResult" class="card hidden">
      <div class="card-head"><h2>품의서 초안</h2><button id="copyDraft" class="button secondary">본문 복사</button></div>
      <pre id="draftText" class="document"></pre>
    </article>
  </section>

  <section id="meeting" class="panel">
    <article class="card">
      <h2>업무추진비·협의회 품의</h2>
      <div class="form-grid">
        <label class="full">협의회명<input id="meetingName" placeholder="예: 학교급식 운영 협의회"></label>
        <label>일시<input id="meetingDate" placeholder="예: 2026. 7. 28.(화) 16:00"></label>
        <label>장소<input id="meetingPlace" placeholder="예: 교무실"></label>
        <label class="full">협의사항<textarea id="agenda" placeholder="예: 자율선택급식 운영 방안 협의 외 2건"></textarea></label>
        <label>참석자 표기<input id="attendee" placeholder="예: 김소리 외 4명"></label>
        <label>참석 인원<input id="attendeeCount" type="number" min="1" value="1"></label>
        <label>1인당 단가<input id="perPerson" type="number" min="0" step="100" placeholder="15000"></label>
        <label>관련 문서<input id="meetingRelated" placeholder="없으면 비워 두세요"></label>
      </div>
      <button id="makeMeeting" class="button primary top-space">협의회 품의서 만들기</button>
    </article>
    <article id="meetingResult" class="card hidden">
      <div class="card-head"><h2>협의회 품의서 초안</h2><button id="copyMeeting" class="button secondary">본문 복사</button></div>
      <pre id="meetingText" class="document"></pre>
    </article>
  </section>
</main>`;

const $ = selector => document.querySelector(selector);
const setStatus = (message, kind = 'info') => {
  const box = $('#status');
  box.textContent = message;
  box.className = `status show ${kind}`;
};

function saveBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function downloadBlankTemplate() {
  const response = await fetch(encodeURI(TEMPLATE_URL), { cache: 'no-store' });
  if (!response.ok) throw new Error(`빈 서식을 불러오지 못했습니다. (${response.status})`);
  saveBlob(await response.blob(), '품목내역(통합).xls');
}

async function spreadsheetToText(file) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', raw: false });
  return workbook.SheetNames.map(name => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '', raw: false });
    return `\n[시트: ${name}]\n${rows.map(row => row.map(text).join('\t')).join('\n')}`;
  }).join('\n').slice(0, 120000);
}

function canvasToDataUrl(canvas, quality = 0.78) {
  return canvas.toDataURL('image/jpeg', quality);
}

async function pdfToContent(file) {
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
  const extracted = [];
  const images = [];
  const pageLimit = Math.min(pdf.numPages, 6);

  for (let pageNumber = 1; pageNumber <= pageLimit; pageNumber++) {
    setStatus(`PDF ${pageNumber}/${pageLimit}쪽을 읽고 있습니다…`, 'loading');
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    extracted.push(`[${pageNumber}쪽]\n${content.items.map(item => item.str).join(' ')}`);

    const viewport = page.getViewport({ scale: 1.35 });
    const canvas = document.createElement('canvas');
    canvas.width = Math.round(viewport.width);
    canvas.height = Math.round(viewport.height);
    await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
    images.push(canvasToDataUrl(canvas));
  }
  return { text: extracted.join('\n'), images };
}

async function imageToDataUrl(file) {
  const bitmap = await createImageBitmap(file);
  const maxSide = 1800;
  const scale = Math.min(1, maxSide / Math.max(bitmap.width, bitmap.height));
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(bitmap.width * scale);
  canvas.height = Math.round(bitmap.height * scale);
  canvas.getContext('2d').drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  return canvasToDataUrl(canvas, 0.82);
}

async function prepareFile(file) {
  const extension = file.name.split('.').pop().toLowerCase();
  if (['xlsx', 'xls', 'csv'].includes(extension)) {
    return { text: await spreadsheetToText(file), images: [] };
  }
  if (extension === 'pdf') return pdfToContent(file);
  if (['png', 'jpg', 'jpeg', 'webp'].includes(extension)) {
    return { text: '', images: [await imageToDataUrl(file)] };
  }
  throw new Error('지원하지 않는 파일 형식입니다.');
}

async function analyzeFile(file) {
  sourceName = file.name;
  setStatus('견적서를 준비하고 있습니다…', 'loading');
  const payload = await prepareFile(file);
  setStatus('AI가 품목과 부가세를 판단하고 있습니다…', 'loading');

  const response = await fetch('/api/analyze', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ ...payload, fileName: file.name })
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || 'AI 분석에 실패했습니다.');

  supplier = text(data.supplier);
  quoteTotal = number(data.quoteTotal);
  warnings = Array.isArray(data.warnings) ? data.warnings : [];
  items = (data.items || []).map((item, index) => ({
    id: crypto.randomUUID?.() || `${Date.now()}-${index}`,
    name: text(item.name),
    spec: text(item.spec),
    quantity: number(item.quantity) || 1,
    unitPrice: Math.round(number(item.unitPrice)),
    amount: Math.round((number(item.quantity) || 1) * number(item.unitPrice))
  })).filter(item => item.name);

  if (!items.length) throw new Error('AI가 실제 구매 품목을 찾지 못했습니다. 견적서가 선명한지 확인해 주세요.');
  renderItems();
  $('#resultCard').classList.remove('hidden');
  $('#draftCard').classList.remove('hidden');
  $('#sourceInfo').textContent = `${sourceName}${supplier ? ` · ${supplier}` : ''}`;
  setStatus(`${items.length}개 실제 품목을 찾았습니다. 내려받기 전에 내용을 확인해 주세요.`, 'success');
  $('#resultCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function renderItems() {
  $('#itemRows').innerHTML = items.map((item, index) => `
    <tr data-id="${item.id}">
      <td class="center">${index + 1}</td>
      <td><input data-key="name" value="${escapeHtml(item.name)}"></td>
      <td><input data-key="spec" value="${escapeHtml(item.spec)}"></td>
      <td><input data-key="quantity" type="number" min="0" step="0.01" value="${item.quantity}"></td>
      <td><input data-key="unitPrice" type="number" min="0" step="1" value="${item.unitPrice}"></td>
      <td class="money">${won(item.amount)}</td>
      <td><button class="delete" data-remove="${item.id}" aria-label="삭제">×</button></td>
    </tr>`).join('');

  if (warnings.length) {
    $('#warningBox').textContent = warnings.join(' · ');
    $('#warningBox').classList.remove('hidden');
  } else {
    $('#warningBox').classList.add('hidden');
  }
  updateSummary();
}

function updateSummary() {
  items.forEach(item => { item.amount = Math.round(number(item.quantity) * number(item.unitPrice)); });
  const total = items.reduce((sum, item) => sum + item.amount, 0);
  $('#itemCount').textContent = `${items.length}개`;
  $('#calculatedTotal').textContent = won(total);
  $('#originalTotal').textContent = quoteTotal ? won(quoteTotal) : '확인 없음';
  const difference = quoteTotal ? Math.abs(total - quoteTotal) : 0;
  $('#matchStatus').textContent = !quoteTotal ? '직접 확인' : difference <= 1 ? '일치' : `${won(difference)} 차이`;
  $('#matchStatus').className = quoteTotal && difference > 1 ? 'bad' : 'good';
}

async function exportEdufine() {
  const validItems = items.filter(item => text(item.name) && number(item.quantity) > 0);
  if (!validItems.length) throw new Error('내려받을 품목이 없습니다.');

  const response = await fetch(encodeURI(TEMPLATE_URL), { cache: 'no-store' });
  if (!response.ok) throw new Error(`에듀파인 원본 서식을 불러오지 못했습니다. (${response.status})`);
  const workbook = XLSX.read(await response.arrayBuffer(), { type: 'array', cellStyles: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error('에듀파인 원본 시트를 찾지 못했습니다.');

  // 실제 원본 양식은 A1:D1 = 내용, 규격, 수량, 예상단가이며 품목은 2행부터 입력한다.
  const START_ROW = 2;
  const originalRange = XLSX.utils.decode_range(sheet['!ref'] || 'A1:D1');
  const clearTo = Math.max(originalRange.e.r + 1, START_ROW + 500);

  for (let row = START_ROW; row <= clearTo; row++) {
    for (let col = 0; col < 4; col++) {
      delete sheet[XLSX.utils.encode_cell({ r: row - 1, c: col })];
    }
  }

  validItems.forEach((item, index) => {
    const row = START_ROW - 1 + index;
    const values = [
      text(item.name),
      text(item.spec),
      number(item.quantity),
      Math.round(number(item.unitPrice))
    ];
    values.forEach((value, col) => {
      const headerStyle = sheet[XLSX.utils.encode_cell({ r: 0, c: col })]?.s;
      sheet[XLSX.utils.encode_cell({ r: row, c: col })] = {
        v: value,
        t: typeof value === 'number' ? 'n' : 's',
        ...(headerStyle ? { s: headerStyle } : {})
      };
    });
  });

  sheet['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: Math.max(0, START_ROW - 2 + validItems.length), c: 3 }
  });

  const output = XLSX.write(workbook, { bookType: 'biff8', type: 'array', cellStyles: true });
  saveBlob(new Blob([output], { type: 'application/vnd.ms-excel' }), '품목내역(통합).xls');
}

function makePurchaseDraft() {
  const purpose = text($('#purpose').value) || '업무 운영';
  const budget = text($('#budget').value);
  const method = text($('#method').value);
  const related = text($('#related').value);
  const attachment = text($('#attachment').value) || '견적서 1부';
  const total = items.reduce((sum, item) => sum + item.amount, 0);
  const itemLabel = items.length > 1 ? `${items[0].name} 외 ${items.length - 1}종` : items[0]?.name || '물품';
  const lines = [`제목: ${purpose} 물품 구입`, ''];
  if (related) lines.push(`1. 관련: ${related}`, '');
  lines.push(`${related ? '2' : '1'}. ${purpose}에 필요한 물품을 다음과 같이 구입하고자 합니다.`, '');
  lines.push(`  가. 구입 목적: ${purpose}`);
  lines.push(`  나. 구입 품목: ${itemLabel}`);
  lines.push(`  다. 소요예산: 금${Math.round(total).toLocaleString('ko-KR')}원`);
  lines.push('  라. 산출내역: 붙임 견적서 참조');
  if (budget) lines.push(`  마. 예산 과목: ${budget}`);
  lines.push(`  ${budget ? '바' : '마'}. 구입 방법: ${method}`, '');
  lines.push(`붙임  ${attachment}.  끝.`);
  $('#draftText').textContent = lines.join('\n');
  $('#draftResult').classList.remove('hidden');
  $('#draftResult').scrollIntoView({ behavior: 'smooth' });
}

function makeMeetingDraft() {
  const name = text($('#meetingName').value) || '업무 협의회';
  const date = text($('#meetingDate').value) || '(일시 입력)';
  const place = text($('#meetingPlace').value) || '(장소 입력)';
  const agenda = text($('#agenda').value) || '(협의사항 입력)';
  const attendee = text($('#attendee').value) || '(참석자 입력)';
  const count = Math.max(1, number($('#attendeeCount').value));
  const unitPrice = number($('#perPerson').value);
  const related = text($('#meetingRelated').value);
  const total = count * unitPrice;
  const lines = [`제목: ${name} 실시`, ''];
  if (related) lines.push(`1. 관련: ${related}`, '');
  lines.push(`${related ? '2' : '1'}. ${name}를 아래와 같이 실시하고자 합니다.`, '');
  lines.push(`  가. 일시: ${date}`);
  lines.push(`  나. 장소: ${place}`);
  lines.push(`  다. 협의사항: ${agenda}`);
  lines.push(`  라. 참석자: ${attendee}, 총 ${count}명`);
  lines.push(`  마. 소요예산: 금${Math.round(total).toLocaleString('ko-KR')}원`);
  lines.push(`  바. 산출내역: ${Math.round(unitPrice).toLocaleString('ko-KR')}원 × ${count}명`, '');
  lines.push('붙임  지출품의서 1부.  끝.');
  $('#meetingText').textContent = lines.join('\n');
  $('#meetingResult').classList.remove('hidden');
}

async function copyContent(selector, button) {
  await navigator.clipboard.writeText($(selector).textContent);
  const before = button.textContent;
  button.textContent = '복사 완료';
  setTimeout(() => { button.textContent = before; }, 1300);
}

$('.tabs').addEventListener('click', event => {
  const tab = event.target.closest('.tab');
  if (!tab) return;
  document.querySelectorAll('.tab').forEach(item => item.classList.toggle('active', item === tab));
  document.querySelectorAll('.panel').forEach(panel => panel.classList.toggle('active', panel.id === tab.dataset.target));
});

$('#chooseFile').onclick = () => $('#fileInput').click();
$('#fileInput').onchange = () => analyzeFile($('#fileInput').files[0]).catch(error => setStatus(error.message, 'error'));
$('#blankTemplate').onclick = () => downloadBlankTemplate().catch(error => alert(error.message));

const dropzone = $('#dropzone');
['dragenter', 'dragover'].forEach(type => dropzone.addEventListener(type, event => {
  event.preventDefault(); dropzone.classList.add('over');
}));
['dragleave', 'drop'].forEach(type => dropzone.addEventListener(type, event => {
  event.preventDefault(); dropzone.classList.remove('over');
}));
dropzone.addEventListener('drop', event => analyzeFile(event.dataTransfer.files[0]).catch(error => setStatus(error.message, 'error')));

$('#itemRows').addEventListener('input', event => {
  const row = event.target.closest('tr');
  const item = items.find(entry => entry.id === row?.dataset.id);
  if (!item) return;
  const key = event.target.dataset.key;
  item[key] = ['quantity', 'unitPrice'].includes(key) ? number(event.target.value) : event.target.value;
  item.amount = Math.round(item.quantity * item.unitPrice);
  row.children[5].textContent = won(item.amount);
  updateSummary();
});
$('#itemRows').addEventListener('click', event => {
  const id = event.target.dataset.remove;
  if (!id) return;
  items = items.filter(item => item.id !== id);
  renderItems();
});
$('#addItem').onclick = () => {
  items.push({ id: crypto.randomUUID(), name: '', spec: '', quantity: 1, unitPrice: 0, amount: 0 });
  renderItems();
};
$('#downloadEdufine').onclick = async () => {
  try { await exportEdufine(); }
  catch (error) { alert(error.message || '파일 생성에 실패했습니다.'); }
};
$('#startOver').onclick = () => location.reload();
$('#makeDraft').onclick = makePurchaseDraft;
$('#copyDraft').onclick = event => copyContent('#draftText', event.currentTarget);
$('#makeMeeting').onclick = makeMeetingDraft;
$('#copyMeeting').onclick = event => copyContent('#meetingText', event.currentTarget);
