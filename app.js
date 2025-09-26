// =====================
// CSV to Excel V4.4.3
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

// 金額欄位（需求重點）
const MONEY_FIELDS = ["支出金額", "存入金額", "餘額"];
const EXCEL_MONEY_FORMAT = "#,##0"; // Excel 顯示為 23,000（仍為數字）

// ===== 狀態 / 元件 =====
const fileMap   = new Map(); // key: 檔案路徑（webkitRelativePath 優先） value: File
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
const mergeMode   = document.getElementById('mergeMode');
const mergeFilename = document.getElementById('mergeFilename');

const sumOutEl = document.getElementById('sumOut');
const sumInEl  = document.getElementById('sumIn');
const sumBalEl = document.getElementById('sumBal');

// 總計（跨所有檔案加總）
const totals = { "支出金額": 0, "存入金額": 0, "餘額": 0 };

// ===== 綁定 =====
btnPick.addEventListener('click', () => picker.click());
picker.addEventListener('change', (e) => handleFiles(e.target.files));
btnStart.addEventListener('click', startConversion);
btnClear.addEventListener('click', () => { fileMap.clear(); duplicates.clear(); renderFileList(); resetTotals(); log('🧹 已清除清單與統計'); });

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
function setProgress(p) {
  bar.style.width = Math.max(0, Math.min(100, p)) + '%';
}
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
function resetTotals() {
  totals["支出金額"] = 0;
  totals["存入金額"] = 0;
  totals["餘額"]    = 0;
  renderTotals();
}
function renderTotals() {
  // 顯示時用本地千分位格式
  sumOutEl.textContent = totals["支出金額"].toLocaleString();
  sumInEl.textContent  = totals["存入金額"].toLocaleString();
  sumBalEl.textContent = totals["餘額"].toLocaleString();
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
      renderFileList();
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
}
function addFile(file) {
  if (!file.name.toLowerCase().endsWith('.csv')) return;
  const key = file.webkitRelativePath || file.name;
  if (fileMap.has(key)) { duplicates.add(key); return; }
  fileMap.set(key, file);
}

// ===== 資料偵測 / 格式 =====
function isNumeric(v) {
  v = String(v).trim();
  return /^-?\d+(?:\.\d+)?$/.test(v);
}
// 移除 + 號、逗號、空白，轉為整數數值；無法解析回傳 null
function parseMoneyToInt(raw) {
  if (raw == null) return null;
  const s = String(raw).replace(/[+,]/g, '').trim();
  if (s === '' || isNaN(Number(s))) return null;
  // 四捨五入去小數
  return Math.round(parseFloat(s));
}
function normalizeMoneyFields(data, headers) {
  const present = MONEY_FIELDS.filter(h => headers.includes(h));
  if (present.length === 0) return present;

  for (let i=0; i<data.length; i++) {
    const row = data[i] || {};
    for (const h of present) {
      const n = parseMoneyToInt(row[h]);
      if (n !== null) {
        row[h] = n; // 直接用數字（Excel 可計算）
      } else {
        // 空字串或非數字：保持空字串，避免 NaN
        row[h] = (row[h] == null || String(row[h]).trim()==='') ? '' : row[h];
      }
    }
    data[i] = row;
  }
  return present;
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
  return Array.from(set);
}
function detectNumericColumns(data, headers, textCols) {
  const textSet = new Set(textCols);
  const num = [];
  for (let h of headers) {
    if (textSet.has(h)) continue;
    if (MONEY_FIELDS.includes(h)) { // 金額欄位一定視為數字
      num.push(h);
      continue;
    }
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
function convertNumeric(data, numericCols) {
  const set = new Set(numericCols);
  for (let r=0; r<data.length; r++) {
    const row = data[r];
    set.forEach(h => {
      if (MONEY_FIELDS.includes(h)) return; // 金額已於 normalizeMoneyFields 處理
      const t = ((row && row[h]) ?? '').toString().trim();
      row[h] = isNumeric(t) ? parseFloat(t) : (t==='' ? '' : t);
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
// 金額欄位套用 Excel 格式 #,##0（確保顯示千分位、且仍為數字）
function applyMoneyFormats(ws, headers, aoaRows) {
  for (let c=0; c<headers.length; c++) {
    const h = headers[c];
    if (!MONEY_FIELDS.includes(h)) continue;
    for (let r=1; r<aoaRows; r++) {
      const ref = XLSX.utils.encode_cell({ c, r });
      const cell = ws[ref];
      if (!cell) continue;
      // 若是數字，就直接套用格式；若是字串但可轉數字，也轉成數字
      if (cell.t === 'n') {
        cell.z = EXCEL_MONEY_FORMAT;
      } else if (cell.t === 's' || cell.t === 'str') {
        const n = parseMoneyToInt(cell.v);
        if (n !== null) {
          cell.v = n;
          cell.t = 'n';
          cell.z = EXCEL_MONEY_FORMAT;
        }
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

// ===== 主流程 =====
async function startConversion() {
  if (fileMap.size === 0) { alert('請先選擇 CSV 檔案'); return; }
  const merge = mergeMode.checked;
  let outName = (mergeFilename.value || '').trim() || '合併檔案.xlsx';
  if (!/\.xlsx$/i.test(outName)) outName += '.xlsx';

  // 重置統計
  resetTotals();

  log(`🚀 開始轉換，共 ${fileMap.size} 個檔案`);
  setProgress(1);

  const wb = XLSX.utils.book_new();
  const files = Array.from(fileMap.values());

  for (let i=0; i<files.length; i++) {
    const f = files[i];
    try {
      log(`處理：${f.name}`);
      let text = await f.text();
      // 去除 BOM
      if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);

      // 同步解析
      const csv = Papa.parse(text, { header: true, skipEmptyLines: 'greedy' });
      if (!csv || !csv.meta) throw new Error('CSV 解析失敗或格式不正確');

      let data = Array.isArray(csv.data) ? csv.data : [];
      const headers = Array.isArray(csv.meta.fields) ? csv.meta.fields : [];
      if (!headers.length) { log(`⚠️ 無標題或空檔，已跳過：${f.name}`); continue; }

      // 先將金額欄位轉成數字（去 +、去小數 → 整數），並累計統計
      const moneyPresent = normalizeMoneyFields(data, headers);
      // 累計 totals
      for (const row of data) {
        for (const h of moneyPresent) {
          const n = typeof row[h] === 'number' ? row[h] : parseMoneyToInt(row[h]);
          if (typeof n === 'number' && !Number.isNaN(n)) totals[h] += n;
        }
      }
      renderTotals();

      // 其他欄位處理
      const textCols = detectTextColumns(data, headers);
      const numCols  = detectNumericColumns(data, headers, textCols);
      applyCustomFormat(data, headers);
      convertNumeric(data, numCols);

      // AOA 組裝
      const aoa = [headers];
      for (let r=0; r<data.length; r++) aoa.push(headers.map(h => data[r][h] ?? ''));

      // 產生 Sheet
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols'] = autoColumnWidths(aoa, SAMPLE_ROWS_FOR_WIDTH);
      forceTextCells(ws, headers, textCols, aoa.length);
      applyMoneyFormats(ws, headers, aoa.length); // 套用金額格式（#,##0）
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

  if (merge && fileMap.size > 0) {
    const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    saveAs(new Blob([buf], { type: 'application/octet-stream' }), outName);
  }

  log('✅ 全部轉換完成');
  showToast('轉換完成！');
  setProgress(0);
}
