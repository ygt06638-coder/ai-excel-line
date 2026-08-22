const FN_URL = window.APP_CONFIG.SUPABASE_FUNCTIONS_URL;

const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const fileChip = document.getElementById('fileChip');
const fileName = document.getElementById('fileName');
const rowCountEl = document.getElementById('rowCount');

const stationConfig = document.getElementById('station-config');
const sourceColumnsGrid = document.getElementById('sourceColumnsGrid');
const targetsList = document.getElementById('targetsList');
const addTargetBtn = document.getElementById('addTargetBtn');
const headersListEl = document.getElementById('headersList');
const startBtn = document.getElementById('startBtn');
const etaNote = document.getElementById('etaNote');

const stationBelt = document.getElementById('station-belt');
const beltFill = document.getElementById('beltFill');
const progressCount = document.getElementById('progressCount');
const progressPct = document.getElementById('progressPct');
const beltStatus = document.getElementById('beltStatus');
const quotaPause = document.getElementById('quotaPause');
const quotaPauseText = document.getElementById('quotaPauseText');
const resumeBtn = document.getElementById('resumeBtn');

const stationDone = document.getElementById('station-done');
const downloadBtn = document.getElementById('downloadBtn');

const errorBox = document.getElementById('errorBox');
const logoutBtn = document.getElementById('logoutBtn');

let currentJobId = null;
let currentRowCount = 0;
let currentSheetName = 'Sheet1';
let currentHeaders = [];
const MAX_TARGETS = 6;

// ---------- حفظ رقم الشغلانة في الجهاز عشان لو النت قطع أو الصفحة اتقفلت نقدر نكمل ----------
const JOB_STORAGE_KEY = 'itqan_current_job';

function saveJobState(stage) {
  try {
    localStorage.setItem(
      JOB_STORAGE_KEY,
      JSON.stringify({ jobId: currentJobId, stage, rowCount: currentRowCount }),
    );
  } catch (e) {
    // مفيش localStorage؟ متهمش، بس المعالجة مش هتكمل تلقائي لو حصل قطع
  }
}

function clearJobState() {
  try {
    localStorage.removeItem(JOB_STORAGE_KEY);
  } catch (e) {}
}

function loadJobState() {
  try {
    const raw = localStorage.getItem(JOB_STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function getToken() {
  return localStorage.getItem('app_token') || '';
}

function authHeaders(extra = {}) {
  const token = getToken();
  return token ? { ...extra, 'x-app-token': token } : extra;
}

// api() بترجع خطأ فيه isNetworkError=true لو المشكلة اتصال بالنت (مش رد من السيرفر)
async function api(path, options = {}) {
  let res;
  try {
    res = await fetch(`${FN_URL}/${path}`, {
      ...options,
      headers: authHeaders({ 'Content-Type': 'application/json', ...(options.headers || {}) }),
    });
  } catch (networkErr) {
    const e = new Error('تعذر الاتصال بالسيرفر، تأكد من اتصال النت');
    e.isNetworkError = true;
    throw e;
  }
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || 'حصل خطأ غير متوقع');
  return data;
}

function toArabicDigits(n) {
  return String(n).replace(/[0-9]/g, d => '٠١٢٣٤٥٦٧٨٩'[d]);
}

function showError(msg) {
  errorBox.hidden = false;
  errorBox.textContent = msg;
}

function clearError() {
  errorBox.hidden = true;
  errorBox.textContent = '';
}

// ---------- التحقق من تسجيل الدخول ----------
(async function checkAuthStatus() {
  try {
    const data = await api('auth-status', { method: 'GET' });
    if (data.passwordRequired) {
      logoutBtn.hidden = false;
      if (!data.loggedIn) {
        window.location.href = 'login.html';
        return;
      }
    }
    // بعد ما نتأكد إن الدخول تمام، نشوف لو فيه شغلانة واقفة نكملها
    tryResumeJob();
  } catch (err) {
    // تجاهل، مش حرج لعرض الصفحة
    tryResumeJob();
  }
})();

// ---------- استكمال شغلانة كانت شغالة قبل ما الصفحة تتقفل أو النت يقطع ----------
function tryResumeJob() {
  const saved = loadJobState();
  if (!saved || !saved.jobId) return;

  currentJobId = saved.jobId;
  currentRowCount = saved.rowCount || 0;

  if (saved.stage === 'done') {
    stationDone.hidden = false;
    stationDone.scrollIntoView({ behavior: 'smooth', block: 'start' });
    return;
  }

  if (saved.stage === 'processing') {
    stationBelt.hidden = false;
    beltStatus.textContent = 'رجعنا نكمل من حيث ما وقفنا…';
    stationBelt.scrollIntoView({ behavior: 'smooth', block: 'start' });
    pollStep();
  }
}

logoutBtn.addEventListener('click', async () => {
  try {
    await api('logout', { method: 'POST' });
  } finally {
    localStorage.removeItem('app_token');
    window.location.href = 'login.html';
  }
});

// ---------- Drag & drop ----------
['dragenter', 'dragover'].forEach(evt =>
  dropzone.addEventListener(evt, e => { e.preventDefault(); dropzone.classList.add('drag-over'); })
);
['dragleave', 'drop'].forEach(evt =>
  dropzone.addEventListener(evt, e => { e.preventDefault(); dropzone.classList.remove('drag-over'); })
);
dropzone.addEventListener('drop', e => {
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});
fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) handleFile(fileInput.files[0]);
});

// ---------- قراءة الإكسيل في المتصفح (SheetJS) ثم إرساله كـ JSON ----------
async function handleFile(file) {
  clearError();
  clearJobState(); // ملف جديد = شغلانة جديدة، امسح أي رقم شغلانة قديم محفوظ

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

    if (rows.length === 0) {
      showError('الملف فارغ');
      return;
    }

    const headers = Object.keys(rows[0]);
    currentSheetName = sheetName;

    const data = await api('upload', {
      method: 'POST',
      body: JSON.stringify({ headers, rows, sheetName }),
    });

    currentJobId = data.jobId;
    currentRowCount = data.rowCount;

    fileName.textContent = file.name;
    rowCountEl.textContent = `${toArabicDigits(data.rowCount)} صف`;
    fileChip.hidden = false;

    populateColumns(data.headers);
    stationConfig.hidden = false;
    stationConfig.scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (err) {
    showError(err.message);
  }
}

function populateColumns(headers) {
  currentHeaders = headers;

  sourceColumnsGrid.innerHTML = '';
  headers.forEach((h, idx) => {
    const label = document.createElement('label');
    label.className = 'checkbox-item';

    const input = document.createElement('input');
    input.type = 'checkbox';
    input.value = h;
    input.checked = idx === 0; // اختار أول عمود افتراضياً عشان المستخدم ميبتديش من فاضي

    const span = document.createElement('span');
    span.textContent = h;

    input.addEventListener('change', () => {
      label.classList.toggle('is-checked', input.checked);
    });
    if (input.checked) label.classList.add('is-checked');

    label.appendChild(input);
    label.appendChild(span);
    sourceColumnsGrid.appendChild(label);
  });

  headersListEl.innerHTML = '';
  headers.forEach(h => headersListEl.appendChild(new Option(h, h)));

  targetsList.innerHTML = '';
  addTargetRow();
  updateAddTargetBtnState();
}

function addTargetRow() {
  if (targetsList.children.length >= MAX_TARGETS) return;

  const row = document.createElement('div');
  row.className = 'target-row';

  const head = document.createElement('div');
  head.className = 'target-row-head';

  const colInput = document.createElement('input');
  colInput.type = 'text';
  colInput.className = 'target-column-input';
  colInput.setAttribute('list', 'headersList');
  colInput.placeholder = 'اسم العمود (موجود أو جديد)';

  const removeBtn = document.createElement('button');
  removeBtn.type = 'button';
  removeBtn.className = 'remove-target-btn';
  removeBtn.textContent = '✕';
  removeBtn.title = 'احذف عمود الهدف ده';
  removeBtn.addEventListener('click', () => {
    if (targetsList.children.length <= 1) return; // سيب عمود هدف واحد على الأقل
    row.remove();
    updateAddTargetBtnState();
  });

  head.appendChild(colInput);
  head.appendChild(removeBtn);

  const instrInput = document.createElement('textarea');
  instrInput.className = 'target-instruction-input';
  instrInput.rows = 5;
  instrInput.maxLength = 20000;
  instrInput.placeholder = 'قول للذكاء الاصطناعي عايزه يكتب ايه في العمود ده بالظبط — تقدر تكتب برومبت مفصّل قد ما تحب';

  row.appendChild(head);
  row.appendChild(instrInput);
  targetsList.appendChild(row);
  updateAddTargetBtnState();
}

function updateAddTargetBtnState() {
  addTargetBtn.disabled = targetsList.children.length >= MAX_TARGETS;
}

addTargetBtn.addEventListener('click', () => addTargetRow());

function collectSourceColumns() {
  return Array.from(sourceColumnsGrid.querySelectorAll('input[type="checkbox"]:checked')).map(el => el.value);
}

function collectTargets() {
  return Array.from(targetsList.children).map(row => ({
    column: row.querySelector('.target-column-input').value.trim(),
    instruction: row.querySelector('.target-instruction-input').value.trim(),
  }));
}

// ---------- بدء المعالجة ----------
startBtn.addEventListener('click', async () => {
  clearError();

  const sourceColumns = collectSourceColumns();
  if (sourceColumns.length === 0) {
    showError('اختار عمود مصدر واحد على الأقل');
    return;
  }

  const targets = collectTargets();
  if (targets.some(t => !t.column || !t.instruction)) {
    showError('كل عمود هدف لازم يكون له اسم وتعليمات');
    return;
  }
  const targetCols = targets.map(t => t.column);
  if (new Set(targetCols).size !== targetCols.length) {
    showError('في عمود هدف مكرر أكتر من مرة، خليه اسم مختلف');
    return;
  }

  startBtn.disabled = true;
  const estSeconds = currentRowCount * targets.length * 4.2;
  etaNote.textContent = `الوقت المتوقع تقريباً: ${toArabicDigits(Math.ceil(estSeconds / 60))} دقيقة (بسبب حدود الـ API المجاني)`;

  try {
    await api('start-process', {
      method: 'POST',
      body: JSON.stringify({
        jobId: currentJobId,
        sourceColumns,
        targets,
      }),
    });

    stationBelt.hidden = false;
    quotaPause.hidden = true;
    stationBelt.scrollIntoView({ behavior: 'smooth', block: 'start' });
    saveJobState('processing');
    pollStep();
  } catch (err) {
    showError(err.message);
    startBtn.disabled = false;
  }
});

// ---------- المعالجة صف بصف (كل نداء process-step بيعالج خطوة واحدة) ----------
let reconnectAttempts = 0;
const MAX_AUTO_RETRIES = 5;

async function pollStep() {
  try {
    const data = await api('process-step', {
      method: 'POST',
      body: JSON.stringify({ jobId: currentJobId }),
    });

    reconnectAttempts = 0; // الاتصال رجع تمام، صفّر عداد المحاولات

    const pct = data.total ? Math.round((data.processed / data.total) * 100) : 0;
    beltFill.style.width = pct + '%';
    progressCount.textContent = `${toArabicDigits(data.processed)} / ${toArabicDigits(data.total)}`;
    progressPct.textContent = `٪${toArabicDigits(pct)}`;

    if (data.paused) {
      beltStatus.textContent = 'الأداة واقفة مؤقتاً بسبب حد استخدام Gemini';
      quotaPauseText.textContent = data.message || 'وصلت لحد الطلبات المسموح بيه دلوقتي.';
      resumeBtn.textContent = 'استكمل المعالجة';
      quotaPause.hidden = false;
      return; // وقّف الـ polling التلقائي، المستخدم هيكمل بنفسه من الزرار
    }

    quotaPause.hidden = true;

    if (data.status === 'error') {
      showError('حصل خطأ أثناء المعالجة: ' + data.error);
      startBtn.disabled = false;
      clearJobState();
      return;
    }

    if (data.status === 'done') {
      beltStatus.textContent = 'تمام، خلصنا كل الصفوف';
      quotaPause.hidden = true;
      stationDone.hidden = false;
      saveJobState('done');
      stationDone.scrollIntoView({ behavior: 'smooth', block: 'start' });
      return;
    }

    if (data.etaSeconds) {
      const mins = Math.ceil(data.etaSeconds / 60);
      beltStatus.textContent = `متبقي تقريباً ${toArabicDigits(mins)} دقيقة…`;
    }

    setTimeout(pollStep, data.pollAfterMs || 4300);
  } catch (err) {
    // النت مقطوع؟ جرّب تاني لوحدك كذا مرة قبل ما تسيب المستخدم يحاول يدوي
    if (err.isNetworkError && reconnectAttempts < MAX_AUTO_RETRIES) {
      reconnectAttempts++;
      const delay = Math.min(30000, 3000 * reconnectAttempts); // 3, 6, 9, 12, 15 ثانية
      beltStatus.textContent = `الاتصال بالنت انقطع، بنحاول تاني (محاولة ${toArabicDigits(reconnectAttempts)} من ${toArabicDigits(MAX_AUTO_RETRIES)})…`;
      setTimeout(pollStep, delay);
      return;
    }

    if (err.isNetworkError) {
      beltStatus.textContent = 'مفيش اتصال بالنت دلوقتي';
      quotaPauseText.textContent = 'اتأكد إن جهازك متصل بالنت، وبعدين دوس على الزرار عشان تكمل.';
      resumeBtn.textContent = 'حاول تاني';
      quotaPause.hidden = false;
      return;
    }

    showError(err.message);
    startBtn.disabled = false;
  }
}

resumeBtn.addEventListener('click', () => {
  quotaPause.hidden = true;
  reconnectAttempts = 0;
  pollStep();
});

// ---------- تحميل الملف النهائي ----------
downloadBtn.addEventListener('click', async () => {
  clearError();
  downloadBtn.disabled = true;
  try {
    const res = await fetch(`${FN_URL}/download?jobId=${encodeURIComponent(currentJobId)}`, {
      headers: authHeaders(),
    });
    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      throw new Error(data.error || 'فشل تحميل الملف');
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'نتيجة-المعالجة.xlsx';
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  } catch (err) {
    showError(err.message);
  } finally {
    downloadBtn.disabled = false;
  }
});
