// =====================
// CSV to Excel V4.4.4
// - 新增：自動偵測並解碼 CSV 編碼（UTF-8 / Big5 / GB18030）避免亂碼
// - 金額欄位（支出金額/存入金額/餘額）去除 + 與小數，截去小數，Excel 以 #,##0 顯示
// - 可統計上述三欄合計
// - 保留 4.4.2 的修正（drop 判斷、路徑去重、log 上限、文字欄 z='@'）
// =====================

// ===== 設定區 =====
const SAMPLE_ROWS_FOR_WIDTH = 50;
const MAX_SHEETNAME_LEN = 31;
const LONG_NUMBER_DIGITS = 13;
const longNumRe = new RegExp("^\\d{" + LONG_NUMBER_DIGITS + ",}$");

// 指定文字欄（避免科學記號）
const FORCE_TEXT_FIELDS = ["交易日期", "交易時間", "帳號", "住家電話", "行動電話"];
const PAD_PHONE_FIELDS   = ["住家電話", "行動電話"];
const PAD_PERIOD_FIELD   = "交易期間";

// 金額欄位（本版新增重點）
const AMOUNT_FIELDS = ["支出金額", "存入金額", "餘額"];

// ===== 狀態 / 元件 =====
const fileMap   = new Map(); // key: 相對路徑或檔名 value: File
const duplicates= new Set();

const logBox    = document.getElementById('log');
const fileList  = document.getElementById('fileList');
const bar       = document.getElementById('bar');
const toast     = document.getElementById('toast');
const picker    = document.getElementById('picker');
const btnPick   = document.getElementById('btnPick');
const btnStart  = document.getElementById('btnStart');
const btnClear  = document.getElementById('btnClear');
const dropzone  = document.getElementById('dropzone');
const mergeMode = document.getElementById('mergeMode');
const mergeFilename = document.getElementById('mergeFilename');

// 即時統計標籤
const chipFiles   = document.getElementById('chip-files');
const chipExpense = document.getElementById('chip-expense');
const chipIncome  = document.getElementById('chip-income');
const chipBalance = document.getElementById('chip-balance');

// 累計統計值（跨檔）
const totals = { expense: 0, income: 0, balance: 0 };

// ===== 綁定 =====
btnPick.addEventListener('click', () => picker.click());
picker.addEventListener('change', (e) => handleFiles(e.target.files));
btnStart.addEventListener('click', startConversion);
btnClear.addEventListener('click', () => {
  fileMap.clear(); duplicates.clear();
  resetTotals();
  renderFileList(); renderChips();
  log('🧹 已清除清單與統計');
});

['dragenter','dragover'].forEach(type => dropzone.addEventListener(type, e => {
  e.preventDefault();
  dropzone.classList.add('active');
  e.dataTransfer.dropEffect = 'copy';
}));
['dragleave','drop'].forEach(type => dropzone.addEventListener(type, e => {
  e.preventDefault();
  if (e.type === 'drop') onDrop(e);
  dropzone.classList.remove('active');
}));

// ===== 小工具 =====
function showToast(msg) {
  toast.textContent = msg;
  toast.classList.add('show');
  setTimeout(() => toast.classList.remove('show'), 2200);
}
function log(msg) {
  const line = document.createElement('div');
  line.textContent = msg;
  logBox.appendChild(line);
  if (logBox.childNodes.length > 500) logBox.removeChild(logBox.firstChild);
  logBox.scrollTop = logBox.scrollHeight;
}
function setProgress(p) { bar.style.width = Math.max(0, Math.min(100, p)) + '%'; }
function escapeHtml(s) {
  s = String(s ?? "");
  const map = { "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;" };
  return s.replace(/[&<>"']/g, ch => map[ch]);
}
function renderFileList() {
  if (fileMap.size === 0) {
    fileList.innerHTML = '<div class="muted">目前無檔案</div>';
    return;
  }
  const rows = Array.from(fileMap.values()).map(f => `
    <div class="file-row">
      <div class="file-name">${escapeHtml(f.webkitRelativePath || f.name)}</div>
      <div class="badge">${(f.size/1024).toFixed(1)} KB</div>
    </div>`).join('');
  fileList.innerHTML = rows;
}
function renderChips() {
  chipFiles.textContent   = `檔案 ${fileMap.size}`;
  chipExpense.textContent = `支出合計 ${formatThousands(totals.expense)}`;
  chipIncome.textContent  = `存入合計 ${formatThousands(totals.income)}`;
  chipBalance.textContent = `餘額合計 ${formatThousands(totals.balance)}`;
}
function resetTotals() {
  totals.expense = 0;
  totals.income  = 0;
  totals.balance = 0;
}

// ===== 檔案/資料夾 載入 =====
function onDrop(e) {
  const items = e.dataTransfer && e.dataTransfer.items;
  if (items && items[0] && typeof items[0].webkitGetAsEntry === 'function') {
    const entries = [];
    for (let i=0;i<items.length;i++) {
      const ent = items[i].webkitGetAsEntry();
      if (ent) entries.push(ent);
    }
    Promise.all(entries.map(ent => traverseEntry(ent))).then(() => {
      if (duplicates.size) showToast(`已跳過重複檔案：${duplicates.size} 筆`);
      renderFileList(); renderChips();
      log(`📁 拖曳匯入完成，共 ${fileMap.size} 檔，重複 ${duplicates.size} 檔`);
    });
  } else if (e.dataTransfer && e.dataTransfer.files) {
    handleFiles(e.dataTransfer.files);
  }
}
function traverseEntry(entry) {
  return new Promise((resolve) => {
    if (entry.isFile) {
      entry.file(file => { addFile(file); resolve(); });
    } else if (entry.isDirectory) {
      const reader = entry.createReader();
      const readBatch = () => reader.readEntries(async (batch) => {
        if (!batch.length) return resolve();
        for (const ent of batch) await traverseEntry(ent);
        readBatch();
      });
      readBatch();
    } else resolve();
  });
}
function handleFiles(list) {
  let add = 0, dup = 0;
  for (let i=0;i<list.length;i++) {
    const file = list[i];
    if (!file.name.toLowerCase().endsWith('.csv')) continue;
    const key = file.webkitRelativePath || file.name;
    if (fileMap.has(key)) { duplicates.add(key); dup++; continue; }
    fileMap.set(key, file); add++;
  }
  renderFileList();
  if (dup) showToast(`已跳過重複 ${dup} 檔`);
  log(`📥 新增 ${add} 檔，重複 ${dup} 檔，總計 ${fileMap.size}`);
  renderChips();
}
function addFile(file) {
  if (!file.name.toLowerCase().endsWith('.csv')) return;
  const key = file.webkitRelativePath || file.name;
  if (fileMap.has(key)) { duplicates.add(key); return; }
  fileMap.set(key, file);
}

// ===== 編碼偵測與解碼（關鍵修正） =====
function hasNonASCII(u8) {
  for (let i = 0; i < u8.length; i++) { if (u8[i] > 0x7F) return true; }
  return false;
}
function scoreTextForChinese(t) {
  // 計算 CJK 區段比例與 � 次數（越多 � 分數越差）
  let cjk = 0, total = 0, repl = 0;
  for (let i=0;i<t.length;i++) {
    const ch = t.charCodeAt(i);
    total++;
    if (ch === 0xFFFD) repl++;
    // 中日韓統一表意文字 + 常用全形標點
    if ((ch >= 0x4E00 && ch <= 0x9FFF) || "，、。；：「」『』（）《》【】！？」＂％＄＃＠＋－＝＼｜".includes(t[i])) {
      cjk++;
    }
  }
  const cjkRatio = total ? (cjk/total) : 0;
  return { cjkRatio, repl };
}
function stripBOM(s) {
  if (!s) return s;
  return s.replace(/^\uFEFF/, '');
}
async function decodeFile(file) {
  // 以 ArrayBuffer 取得原始位元組，避免瀏覽器以 UTF-8 直接解讀造成亂碼
  const buf = await file.arrayBuffer();
  const u8 = new Uint8Array(buf);
  // 優先嘗試 UTF-8，其次 Big5，再來 GB18030（台灣 CSV 常見）
  const candidates = ['utf-8', 'big5', 'gb18030'];
  // 如果完全沒有非 ASCII，直接當 UTF-8
  if (!hasNonASCII(u8)) {
    return stripBOM(new TextDecoder('utf-8').decode(u8));
  }
  let best = { enc: 'utf-8', text: '', score: -Infinity };
  for (const enc of candidates) {
    try {
      const td = new TextDecoder(enc, { fatal: false });
      const text = stripBOM(td.decode(u8));
      const { cjkRatio, repl } = scoreTextForChinese(text);
      // 打分：中文比例越高越好，� 越少越好
      const score = cjkRatio * 1000 - repl * 50;
      if (score > best.score) best = { enc, text, score };
    } catch (_) { /* 某些環境可能不支援該編碼，忽略 */ }
  }
  return best.text || stripBOM(new TextDecoder().decode(u8));
}

// ===== 資料偵測 / 格式 =====
function isNumeric(v) {
  v = String(v).trim();
  return /^-?\d+(?:\.\d+)?$/.test(v);
}
// 清洗金額：移除逗號與 +，保留負號，截去小數（不四捨五入）
function sanitizeAmountToInt(v) {
  if (v == null) return null;
  let s = String(v).trim();
  if (!s) return null;
  s = s.replace(/,/g, '');     // 去千分位
  s = s.replace(/^\+/, '');    // 去掉正號
  const n = Number.parseFloat(s);
  if (Number.isNaN(n)) return null;
  // 去小數：截斷（負數用 Math.ceil 避免 -1.9 -> -1.0 的四捨五入）
  return n < 0 ? Math.ceil(n) : Math.trunc(n);
}
function detectTextColumns(data, headers) {
  const set = new Set(FORCE_TEXT_FIELDS);
  for (let h of headers) {
    if (set.has(h)) continue;
    for (let i=0; i<Math.min(data.length, SAMPLE_ROWS_FOR_WIDTH); i++) {
      const val = ((data[i] && data[i][h]) ?? '').toString().trim();
      if (longNumRe.test(val)) { set.add(h); break; }
    }
  }
  // 確保金額欄絕不列入文字欄（要能統計）
  for (const a of AMOUNT_FIELDS) set.delete(a);
  return Array.from(set);
}
function detectNumericColumns(data, headers, textCols) {
  const textSet = new Set(textCols);
  const num = [];
  for (let h of headers) {
    if (textSet.has(h)) continue;
    if (AMOUNT_FIELDS.includes(h)) { num.push(h); continue; } // 金額欄強制視為數值
    let numericCount = 0, nonEmpty = 0;
    for (let i=0; i<Math.min(data.length, 2000); i++) {
      const raw = ((data[i] && data[i][h]) ?? '').toString().trim();
      if (!raw) continue; nonEmpty++;
      if (isNumeric(raw)) numericCount++;
    }
    if (nonEmpty && numericCount / nonEmpty > 0.8) num.push(h);
  }
  return num;
}
function applyCustomFormat(data, headers) {
  for (let r=0; r<data.length; r++) {
    const row = data[r];
    for (let h of headers) {
      let v = ((row && row[h]) ?? '').toString().trim();
      if (PAD_PHONE_FIELDS.includes(h) && /^\d+$/.test(v)) row[h] = v.padStart(10, '0');
      if (h === PAD_PERIOD_FIELD && /^\d+$/.test(v)) row[h] = v.padStart(6, '0');
    }
  }
}
// 金額欄位：轉為整數（去+與小數），並回填；同時回傳「此筆列的金額」
function normalizeAmountsRow(row) {
  const out = { expense: null, income: null, balance: null };
  if ('支出金額' in row) {
    const n = sanitizeAmountToInt(row['支出金額']);
    row['支出金額'] = (n ?? '');
    out.expense = (n ?? null);
  }
  if ('存入金額' in row) {
    const n = sanitizeAmountToInt(row['存入金額']);
    row['存入金額'] = (n ?? '');
    out.income = (n ?? null);
  }
  if ('餘額' in row) {
    const n = sanitizeAmountToInt(row['餘額']);
    row['餘額'] = (n ?? '');
    out.balance = (n ?? null);
  }
  return out;
}
function convertNumeric(data, numericCols) {
  const set = new Set(numericCols);
  for (let r=0; r<data.length; r++) {
    const row = data[r];
    set.forEach(h => {
      if (AMOUNT_FIELDS.includes(h)) return; // 金額欄已先處理
      const t = ((row && row[h]) ?? '').toString().trim();
      row[h] = isNumeric(t) ? parseFloat(t) : '';
    });
  }
}

// 欄寬估算
function getDisplayWidth(str) {
  const s = String(str ?? '');
  let w = 0; for (let i=0;i<s.length;i++) { w += s.charCodeAt(i) > 255 ? 2 : 1; }
  return w;
}
function autoColumnWidths(aoa, sampleRows) {
  const cols = (aoa[0] && aoa[0].length) ? aoa[0].length : 0;
  const widths = new Array(cols).fill(8);
  const limit = Math.min(1 + sampleRows, aoa.length);
  for (let c=0; c<cols; c++) {
    let maxw = 8;
    for (let r=0; r<limit; r++) {
      const w = getDisplayWidth((aoa[r] && aoa[r][c]) ?? '');
      if (w > maxw) maxw = w;
    }
    widths[c] = { wch: Math.max(8, Math.min(50, Math.round(maxw * 1.1))) };
  }
  return widths;
}
function forceTextCells(ws, headers, textCols, rows) {
  const set = new Set(textCols);
  for (let c=0; c<headers.length; c++) {
    if (!set.has(headers[c])) continue;
    for (let r=1; r<rows; r++) {
      const ref = XLSX.utils.encode_cell({ c, r });
      const cell = ws[ref];
      if (!cell) continue;
      cell.t = 's';
      cell.z = '@';
    }
  }
}
// 對金額欄在工作表中套用 Excel 格式：#,##0（顯示千分位、無小數）
function formatAmountCells(ws, headers, rows) {
  for (let c=0; c<headers.length; c++) {
    const h = headers[c];
    if (!AMOUNT_FIELDS.includes(h)) continue;
    for (let r=1; r<rows; r++) {
      const ref = XLSX.utils.encode_cell({ c, r });
      const cell = ws[ref];
      if (!cell) continue;
      cell.t = 'n';
      cell.z = '#,##0';
      if (cell.v === '' || cell.v == null) {
        delete cell.t; delete cell.z;
      }
    }
  }
}
function uniqueSheetName(wb, base) {
  let name = String(base || 'Sheet').replace(/[\\/?*[\]:]/g, '_').slice(0, MAX_SHEETNAME_LEN) || 'Sheet';
  if (!wb.SheetNames.includes(name)) return name;
  let i = 2;
  while (true) {
    const cand = (name.slice(0, MAX_SHEETNAME_LEN - String(i).length - 1) + '_' + i);
    if (!wb.SheetNames.includes(cand)) return cand;
    i++;
  }
}
function formatThousands(n) {
  if (typeof n !== 'number' || !Number.isFinite(n)) return '0';
  return n.toLocaleString('en-US', { maximumFractionDigits: 0 });
}

// ===== 主流程 =====
async function startConversion() {
  if (fileMap.size === 0) { alert('請先選擇 CSV 檔案'); return; }
  const merge = mergeMode.checked;
  let outName = (mergeFilename.value || '').trim() || '合併檔案.xlsx';
  if (!/\.xlsx$/i.test(outName)) outName += '.xlsx';

  // 每次開始轉換，先歸零統計
  resetTotals(); renderChips();

  log(`🚀 開始轉換，共 ${fileMap.size} 個檔案`);
  setProgress(1);

  const wb = XLSX.utils.book_new();
  const files = Array.from(fileMap.values());

  for (let i=0; i<files.length; i++) {
    const f = files[i];
    try {
      log(`處理：${f.name}`);

      // ⤵️ 關鍵修正：用自動偵測的方式將 CSV bytes 解成正確文字
      let text = await decodeFile(f);

      // 清除隱藏控制字元（有些系統產生的 CSV 會混入 \u0000 等）
      text = text.replace(/\u0000/g, '');
      if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);

      const csv = Papa.parse(text, { header: true, skipEmptyLines: 'greedy' });
      if (!csv || !csv.meta) throw new Error('CSV 解析失敗或格式不正確');

      let data = Array.isArray(csv.data) ? csv.data : [];
      const headers = Array.isArray(csv.meta.fields) ? csv.meta.fields : [];
      if (!headers.length) { log(`⚠️ 無標題或空檔，已跳過：${f.name}`); continue; }

      // 去全空列
      data = data.filter(obj => Object.values(obj).some(v => (v ?? '').toString().trim() !== ''));

      // 先做金額欄位整形並累計
      for (const row of data) {
        const { expense, income, balance } = normalizeAmountsRow(row);
        if (typeof expense === 'number') totals.expense += expense;
        if (typeof income  === 'number') totals.income  += income;
        if (typeof balance === 'number') totals.balance += balance;
      }
      renderChips();

      const textCols = detectTextColumns(data, headers);
      const numCols  = detectNumericColumns(data, headers, textCols);
      applyCustomFormat(data, headers);
      convertNumeric(data, numCols);

      // AOA
      const aoa = [headers];
      for (let r=0; r<data.length; r++) aoa.push(headers.map(h => data[r][h] ?? ''));

      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols'] = autoColumnWidths(aoa, SAMPLE_ROWS_FOR_WIDTH);
      forceTextCells(ws, headers, textCols, aoa.length);
      formatAmountCells(ws, headers, aoa.length); // 套用金額格式
      ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s:{c:0,r:0}, e:{c:headers.length-1, r:Math.max(0, aoa.length-1)} }) };

      if (merge) {
        const base = f.name.replace(/\.csv$/i, '').slice(0, MAX_SHEETNAME_LEN);
        const name = uniqueSheetName(wb, base);
        XLSX.utils.book_append_sheet(wb, ws, name);
      } else {
        const wbSingle = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wbSingle, ws, 'Sheet1');
        const buf = XLSX.write(wbSingle, { bookType: 'xlsx', type: 'array' });
        saveAs(new Blob([buf], { type: 'application/octet-stream' }), f.name.replace(/\.csv$/i, '.xlsx'));
      }
    } catch (err) {
      log(`❌ 轉換失敗：${f.name}，原因：${err.message || err}`);
    }
    setProgress(Math.round(((i + 1) / files.length) * 100));
  }

  if (merge) {
    const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    saveAs(new Blob([buf], { type: 'application/octet-stream' }), outName);
  }

  log('✅ 全部轉換完成');
  showToast('轉換完成！');
  setProgress(0);
  // 保留清單與統計，方便查看
}
