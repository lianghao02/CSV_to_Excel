/* =========================================================================
 * CSV to Excel V4.5.1  (SINGLE-FILE LITE)
 * 說明：維持「單檔運行」，不使用模組。加入章節旗標、統一縮排、收斂事件綁定到 init()，
 * 並整合主題切換（GitHub / Police 按鈕）。
 * 目錄：
 * A. 設定常數 & 欄位別名
 * B. 全域狀態 & DOM 取得
 * C. 小工具（toast/log/寬度/格式化/清單/統計）
 * D. 檔案載入（拖曳/資料夾遞迴/清單渲染）
 * E. 編碼偵測 & 解碼（UTF-8/Big5/GB18030）
 * F. 規格化核心（標題清理、別名、近似比對）
 * G. 欄位偵測 & 金額/字串格式
 * H. 解析 CSV → 列資料（含統計）
 * I. 輸出 Excel（工作表欄寬/格式/匯出）
 * J. 轉換流程（原模式 / 依標題合併）
 * K. HeaderMap 對照表
 * L. 啟動與事件綁定（init）
 * ========================================================================= */

// ===== A. 設定常數 & 欄位別名 =====
const SAMPLE_ROWS_FOR_WIDTH = 50;
const MAX_SHEETNAME_LEN = 31;
const LONG_NUMBER_DIGITS = 13;
const longNumRe = new RegExp("^\\d{" + LONG_NUMBER_DIGITS + ",}$");

// 指定文字欄（避免科學記號）
const FORCE_TEXT_FIELDS = ["交易日期", "交易時間", "帳號", "住家電話", "行動電話"];
const PAD_PHONE_FIELDS   = ["住家電話", "行動電話"];
const PAD_PERIOD_FIELD   = "交易期間";

// 金額欄位
const AMOUNT_FIELDS = ["支出金額", "存入金額", "餘額"];

// 標題別名表（先通用清理 → 再按此表歸一化 → 再做近似字補捉）
const HEADER_ALIASES = {
  "身分證統一編號": ["身份證統一編號","身分證統編號","身份證統編號","身分證號","身份證號","身分證","身份證"],
  // 分離「戶名」與「帳號」
  "帳號": ["帳戶","帳戶號碼","帳戶號","帳號 "],
  "戶名": ["戶名(開戶人)","開戶人名稱","開戶人","客戶名稱","帳戶名稱","戶名 "],
  "交易序號": ["交易編號","交易流水號","交易號碼","交易 序號","交易序號 "],
  "交易日期": ["交易日","交易 日","交易日期 ","交易日期　","交易日期(西元)"],
  "交易時間": ["時間","時 間","交易 時間"],
  "交易行": ["交易銀行","金融機構","金融機構名稱","交易行別","交易行(或所屬分行代號)","所屬分行代號"],
  "交易摘要": ["交易說明","摘要","說明","交易內文"],
  "幣別": ["貨幣別","幣 別"],
  "支出金額": ["支出","支出金 額","支 出金額","提款"],
  "存入金額": ["存入","存入金 額","存 入金額","存款"],
  "餘額": ["結餘","結存","餘額金額","餘額 "],
  "ATM或端末機代碼": ["ATM或端未機代碼","ATM或端木機代碼","ATM或端末機 代碼","端末機代碼","端未機代碼","ATM代碼","ATM/端末機代碼","ATM 或端末機代碼"],
  "櫃員代號": ["櫃員","櫃 員代號"],
  "轉出入行庫代碼及帳號": ["轉出入行庫代碼&帳號","轉出入行庫代碼與帳號","轉出入行庫代碼","往來行庫代碼及帳號"],
  "備註": ["備 註","附註","備考","備  註"],
  "被害人": ["受害人","被害 人","被 害 人"],
  "承租人": ["承 租 人","承租 人"],
  "住家電話": ["電話(住家)","住家 電話","家用電話"],
  "行動電話": ["手機","手機號碼","行動 電話"],
  "戶籍地址": ["戶籍 地址"],
  "通訊地址": ["通訊 地址"],
  "資料提供日期": ["資料提供日","資料提供 日","資料提供日期 "],
  "資料提供日帳戶結餘": ["資料提供日結餘","資料提供日 帳戶結餘","資料提供日期帳戶結餘","資料提供日帳戶餘額"],
  "開戶行總分支機構代碼": ["開戶行總、分支機構代碼","開戶行總分支機構 代碼","開戶行總分支機構代碼 "],
  "交易期間": ["期間","交易 期間"]
};

// ===== B. 全域狀態 & DOM 取得 =====
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
const btnGroupByHeader = document.getElementById('btnGroupByHeader');
const chipFiles   = document.getElementById('chip-files');
const chipExpense = document.getElementById('chip-expense');
const chipIncome  = document.getElementById('chip-income');
const chipBalance = document.getElementById('chip-balance');
const btnCopyLog = document.getElementById('btnCopyLog'); // 【優化：複製日誌按鈕】

// 主題：按鈕
const btnThemeGithub = document.getElementById('btnThemeGithub');
const btnThemePolice = document.getElementById('btnThemePolice');

const totals = { expense: 0, income: 0, balance: 0 };

// ===== C. 小工具（toast/log/寬度/格式化/清單/統計） =====
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
function getDisplayWidth(str){
  const s=String(str??''); let w=0; for(let i=0;i<s.length;i++){ w += s.charCodeAt(i)>255 ? 2 : 1; } return w;
}
function autoColumnWidths(aoa,sampleRows=SAMPLE_ROWS_FOR_WIDTH){
  const cols=(aoa[0]&&aoa[0].length)?aoa[0].length:0; const widths=new Array(cols).fill(8);
  const limit=Math.min(1+sampleRows, aoa.length);
  for (let c=0;c<cols;c++){
    let maxw=8;
    for (let r=0;r<limit;r++){
      const w=getDisplayWidth((aoa[r]&&aoa[r][c])??''); if (w>maxw) maxw=w;
    }
    widths[c]={wch:Math.max(8, Math.min(50, Math.round(maxw*1.1)))};
  }
  return widths;
}
function formatThousands(n){ if (typeof n!=="number"||!Number.isFinite(n)) return '0'; return n.toLocaleString('en-US',{maximumFractionDigits:0}); }
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
  
  // 【優化：餘額顏色提示】
  const balanceVal = totals.balance;
  chipBalance.style.color = (balanceVal > 0 ? 'var(--ok)' : (balanceVal < 0 ? 'var(--danger)' : 'var(--chip-text)'));
  chipBalance.style.borderColor = (balanceVal > 0 ? 'var(--ok)' : (balanceVal < 0 ? 'var(--danger)' : 'var(--border)'));
}
function resetTotals() {
  totals.expense = 0;
  totals.income  = 0;
  totals.balance = 0;
}
function isNumeric(v){ v=String(v).trim(); return /^-?\d+(?:\.\d+)?$/.test(v); }

// ===== D. 檔案載入（拖曳/資料夾遞迴/清單渲染） =====
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

// ===== E. 編碼偵測 & 解碼（UTF-8/Big5/GB18030） =====
function hasNonASCII(u8) { for (let i=0;i<u8.length;i++){ if (u8[i]>0x7F) return true; } return false; }
function scoreTextForChinese(t) {
  let cjk=0,total=0,repl=0;
  for (let i=0;i<t.length;i++){
    const ch=t.charCodeAt(i); total++; if (ch===0xFFFD) repl++;
    if ((ch>=0x4E00&&ch<=0x9FFF) || "，、。；：「」『』（）《》【】！？」＂％＄＃＠＋－＝＼｜".includes(t[i])) cjk++;
  }
  const cjkRatio = total ? (cjk/total) : 0; return { cjkRatio, repl };
}
function stripBOM(s){ return s ? s.replace(/^\uFEFF/,'') : s; }
async function decodeFile(file){
  const buf = await file.arrayBuffer(); const u8=new Uint8Array(buf);
  const candidates=['utf-8','big5','gb18030'];
  if (!hasNonASCII(u8)) return stripBOM(new TextDecoder('utf-8').decode(u8));
  let best={enc:'utf-8',text:'',score:-Infinity};
  for (const enc of candidates){
    try{ const td=new TextDecoder(enc,{fatal:false}); const text=stripBOM(td.decode(u8));
      const {cjkRatio,repl}=scoreTextForChinese(text); const score=cjkRatio*1000 - repl*50;
      if (score>best.score) best={enc,text,score};
    }catch(_){/* ignore */}
  }
  return best.text || stripBOM(new TextDecoder().decode(u8));
}

// ===== F. 規格化核心（標題清理、別名、近似比對） =====
function toHalfwidth(str){
  return String(str)
    .replace(/[\uFF01-\uFF5E]/g, ch => String.fromCharCode(ch.charCodeAt(0)-0xFEE0))
    .replace(/\u3000/g,' ');
}
function basicCleanHeader(raw){
  if (raw==null) return '';
  let s = String(raw);
  s = s.replace(/（.*?）/g,'').replace(/\(.*?\)/g,'');
  s = toHalfwidth(s);
  s = s.replace(/[、，,．·・\.\-_/\\]/g,'');
  s = s.replace(/\s+/g,'').replace(/[\u0000-\u001F\u007F]/g,'');
  return s;
}
const CANON_SET = new Set(Object.keys(HEADER_ALIASES));
const ALIAS_LUT = (()=>{
  const m=new Map();
  for (const k of Object.keys(HEADER_ALIASES)){
    m.set(basicCleanHeader(k), k);
    HEADER_ALIASES[k].forEach(a=> m.set(basicCleanHeader(a), k));
  }
  return m;
})();
function levenshtein(a,b){
  const m=a.length,n=b.length; if (m===0) return n; if (n===0) return m;
  const dp=Array(n+1); for(let j=0;j<=n;j++) dp[j]=j;
  for(let i=1;i<=m;i++){
    let prev=dp[0]; dp[0]=i;
    for(let j=1;j<=n;j++){
      const temp=dp[j];
      if (a[i-1]===b[j-1]) dp[j]=prev;
      else dp[j]=Math.min(prev+1, dp[j]+1, dp[j-1]+1);
      prev=temp;
    }
  }
  return dp[n];
}
function chooseClosestCanonical(cleanKey){
  if (cleanKey.includes("帳號") && cleanKey.includes("戶名")) return "帳號";
  if (ALIAS_LUT.has(cleanKey)) return ALIAS_LUT.get(cleanKey);
  let bestKey=null, bestDist=3;
  for (const canonical of CANON_SET.values()){
    if ((cleanKey.includes("承租人") && canonical==="被害人") ||
        (cleanKey.includes("被害人") && canonical==="承租人")) continue;
    const d = levenshtein(cleanKey, basicCleanHeader(canonical));
    if (d<bestDist){ bestDist=d; bestKey=canonical; if (bestDist===0) break; }
  }
  return (bestDist<=2) ? bestKey : null;
}
function normalizeHeaderName(raw){
  const cleaned = basicCleanHeader(raw);
  const mapped  = chooseClosestCanonical(cleaned);
  return mapped ? mapped : cleaned; // 未命中則回 cleaned 版
}
function canonicalizeHeaders(headers){
  const seen = new Set();
  const canon = [];
  for (const h of headers){
    const n = normalizeHeaderName(h);
    if (!seen.has(n)){ seen.add(n); canon.push(n); }
  }
  return canon;
}
function buildHeaderSignature(headers){
  const canon = canonicalizeHeaders(headers);
  return canon.join('\t');
}

// ===== G. 欄位偵測 & 金額/字串格式 =====
function detectTextColumns(data, headers){
  const set=new Set(FORCE_TEXT_FIELDS);
  for (let h of headers){
    if (set.has(h)) continue;
    for (let i=0;i<Math.min(data.length, SAMPLE_ROWS_FOR_WIDTH);i++){
      const val=((data[i]&&data[i][h])??'').toString().trim();
      if (longNumRe.test(val)){ set.add(h); break; }
    }
  }
  for (const a of AMOUNT_FIELDS) set.delete(a);
  return Array.from(set);
}
function detectNumericColumns(data, headers, textCols){
  const textSet=new Set(textCols); const num=[];
  for (let h of headers){
    if (textSet.has(h)) continue;
    if (AMOUNT_FIELDS.includes(h)){ num.push(h); continue; }
    let numericCount=0, nonEmpty=0;
    for (let i=0;i<Math.min(data.length,2000);i++){
      const raw=((data[i]&&data[i][h])??'').toString().trim();
      if (!raw) continue; nonEmpty++; if (isNumeric(raw)) numericCount++;
    }
    if (nonEmpty && numericCount/nonEmpty>0.8) num.push(h);
  }
  return num;
}
function applyCustomFormat(data, headers){
  for (let r=0;r<data.length;r++){
    const row=data[r];
    for (let h of headers){
      let v=((row&&row[h])??'').toString().trim();
      if (PAD_PHONE_FIELDS.includes(h) && /^\d+$/.test(v)) row[h]=v.padStart(10,'0');
      if (h===PAD_PERIOD_FIELD && /^\d+$/.test(v)) row[h]=v.padStart(6,'0');
    }
  }
}
function sanitizeAmountToInt(v){
  if (v==null) return null; let s=String(v).trim(); if (!s) return null;
  s=s.replace(/,/g,'').replace(/^\+/, ''); const n=Number.parseFloat(s);
  if (Number.isNaN(n)) return null; return n<0 ? Math.ceil(n) : Math.trunc(n);
}
function normalizeAmountsRow(row){
  const out={expense:null,income:null,balance:null};
  if ('支出金額' in row){ const n=sanitizeAmountToInt(row['支出金額']); row['支出金額']=(n??''); out.expense=(n??null); }
  if ('存入金額' in row){ const n=sanitizeAmountToInt(row['存入金額']); row['存入金額']=(n??''); out.income=(n??null); }
  if ('餘額' in row){ const n=sanitizeAmountToInt(row['餘額']); row['餘額']=(n??''); out.balance=(n??null); }
  return out;
}

// ===== H. 解析 CSV → 列資料（含統計） =====
async function parseCsvFile(file){
  let text = await decodeFile(file);
  text = text.replace(/\u0000/g,''); if (text.charCodeAt(0)===0xFEFF) text=text.slice(1);

  const csv = Papa.parse(text, { header:true, skipEmptyLines:'greedy' });
  if (!csv || !csv.meta) throw new Error('CSV 解析失敗或格式不正確');

  let data = Array.isArray(csv.data) ? csv.data : [];
  const headersRaw = Array.isArray(csv.meta.fields) ? csv.meta.fields : [];
  if (!headersRaw.length) return null;

  // 去全空列
  data = data.filter(obj => Object.values(obj).some(v => (v??'').toString().trim()!==''));

  // 規格化標題（含去重）
  const headersCanon = canonicalizeHeaders(headersRaw);

  // 建立「原始→規格化」對照，只記錄第一個命中的欄
  const headerMapPairs = [];
  const src2dst = {};
  const seenCanon = new Set();

  for (let i=0;i<headersRaw.length;i++){
    const raw = headersRaw[i];
    let canon = normalizeHeaderName(raw);

    // 額外保護：同時包含帳號/戶名 → 強制帳號
    const cleaned = basicCleanHeader(raw);
    if (cleaned.includes("帳號") && cleaned.includes("戶名")) canon = "帳號";

    if (!seenCanon.has(canon)){
      src2dst[raw]=canon; seenCanon.add(canon); headerMapPairs.push([raw, canon]);
    }
  }

  // 重建資料列：以「規格化後的欄名」為鍵
  const rows = data.map(row => {
    const o={}; headersCanon.forEach(h => { o[h]=''; });
    for (const rawKey of Object.keys(row)){
      const cleaned = basicCleanHeader(rawKey);
      let dstKey = src2dst[rawKey] ?? normalizeHeaderName(rawKey);
      if (cleaned.includes("帳號") && cleaned.includes("戶名")) dstKey = "帳號"; // 再次保護
      if (dstKey in o && o[dstKey]==='') o[dstKey]=row[rawKey];
    }
    // 金額整形 + 統計
    const {expense,income,balance} = normalizeAmountsRow(o);
    if (typeof expense === 'number') totals.expense += expense;
    if (typeof income  === 'number') totals.income  += income;
    if (typeof balance === 'number') totals.balance += balance;
    return o;
  });

  applyCustomFormat(rows, headersCanon);
  const textCols = detectTextColumns(rows, headersCanon);
  const numCols  = detectNumericColumns(rows, headersCanon, textCols);

  return { headerDisplay: headersCanon, rows, textCols, numCols, headerMapPairs };
}

// ===== I. 輸出 Excel（工作表欄寬/格式/匯出） =====
function forceTextCells(ws, headers, textCols, rows){
  const set=new Set(textCols);
  for (let c=0;c<headers.length;c++){
    if (!set.has(headers[c])) continue;
    for (let r=1;r<rows;r++){
      const ref=XLSX.utils.encode_cell({c,r}); const cell=ws[ref]; if (!cell) continue; cell.t='s'; cell.z='@';
    }
  }
}
function formatAmountCells(ws, headers, rows){
  for (let c=0;c<headers.length;c++){
    const h=headers[c]; if (!AMOUNT_FIELDS.includes(h)) continue;
    for (let r=1;r<rows;r++){
      const ref=XLSX.utils.encode_cell({c,r}); const cell=ws[ref]; if (!cell) continue;
      cell.t='n'; cell.z='#,##0'; if (cell.v===''||cell.v==null){ delete cell.t; delete cell.z; }
    }
  }
}
function uniqueSheetNameNumeric(index){ return String(index).padStart(3,'0'); }

// ===== J. 轉換流程（原模式 / 依標題合併） =====
async function startConversion(){
  if (fileMap.size===0){ alert('請先選擇 CSV 檔案'); return; }
  const merge = mergeMode.checked;
  let outName = (mergeFilename.value||'').trim() || '合併檔案.xlsx';
  if (!/\.xlsx$/i.test(outName)) outName += '.xlsx';

  resetTotals(); renderChips();
  log(`🚀 開始轉換（原模式），共 ${fileMap.size} 個檔案`);
  setProgress(1);

  const wb = XLSX.utils.book_new();
  const files = Array.from(fileMap.values());

  // HeaderMap 蒐集
  const headerAudit = [];

  for (let i=0;i<files.length;i++){
    const f = files[i];
    try{
      log(`處理：${f.name}`);
      const parsed = await parseCsvFile(f);
      if (!parsed){ log(`⚠️ 無標題或空檔，已跳過：${f.name}`); continue; }
      const { headerDisplay, rows, textCols, numCols, headerMapPairs } = parsed;
      renderChips();
      headerAudit.push({ src:f.name, map:headerMapPairs, sig:buildHeaderSignature(headerDisplay) });

      const aoa=[headerDisplay]; for (let r=0;r<rows.length;r++) aoa.push(headerDisplay.map(h => rows[r][h] ?? ''));
      const ws=XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols']=autoColumnWidths(aoa, SAMPLE_ROWS_FOR_WIDTH);
      forceTextCells(ws, headerDisplay, textCols, aoa.length);
      formatAmountCells(ws, headerDisplay, aoa.length);
      ws['!autofilter']={ ref: XLSX.utils.encode_range({ s:{c:0,r:0}, e:{c:headerDisplay.length-1, r:Math.max(0,aoa.length-1)} }) };

      if (merge){
        let name = (f.name.replace(/\.csv$/i,'')||'Sheet').replace(/[\\\/?*\[\]:]/g,'_').slice(0,MAX_SHEETNAME_LEN);
        let final=name,k=2; while (wb.SheetNames.includes(final)){ const suffix='_'+(k++); final=name.slice(0,MAX_SHEETNAME_LEN-suffix.length)+suffix; }
        XLSX.utils.book_append_sheet(wb, ws, final);
      }else{
        const wbSingle=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wbSingle, ws, 'Sheet1');
        const buf=XLSX.write(wbSingle,{bookType:'xlsx',type:'array'});
        saveAs(new Blob([buf],{type:'application/octet-stream'}), f.name.replace(/\.csv$/i,'.xlsx'));
      }
    }catch(err){
      log(`❌ 轉換失敗：${f.name}，原因：${err.message||err}`);
    }
    setProgress(Math.round(((i+1)/files.length)*100));
  }

  if (merge){
    appendHeaderMapSheet(wb, headerAudit);
    const buf=XLSX.write(wb,{bookType:'xlsx',type:'array'});
    saveAs(new Blob([buf],{type:'application/octet-stream'}), outName);
  }

  log('✅ 全部轉換完成');
  showToast('轉換完成！');
  setProgress(0);
}

async function groupByHeaderConversion(){
  if (fileMap.size===0){ alert('請先選擇 CSV 檔案'); return; }

  resetTotals(); renderChips();
  log(`🧩 開始依「規格化標題」合併分頁，共 ${fileMap.size} 個檔案`);
  setProgress(1);

  const wb = XLSX.utils.book_new();
  const files = Array.from(fileMap.values());

  // groupMap: key = signature, val = { header, parts:[{rows,textCols,numCols,src}], sig }
  const groupMap = new Map();
  const headerAudit = [];

  for (let i=0;i<files.length;i++){
    const f=files[i];
    try{
      log(`解析：${f.name}`);
      const parsed = await parseCsvFile(f);
      if (!parsed){ log(`⚠️ 無標題或空檔，已跳過：${f.name}`); continue; }
      const { headerDisplay, rows, textCols, numCols, headerMapPairs } = parsed;
      const sig = buildHeaderSignature(headerDisplay);
      headerAudit.push({ src:f.name, map:headerMapPairs, sig });
      if (!groupMap.has(sig)) groupMap.set(sig, { header:headerDisplay, parts:[], sig });
      groupMap.get(sig).parts.push({ rows, textCols, numCols, src:f.name });
      renderChips();
    }catch(err){
      log(`❌ 解析失敗：${f.name}，原因：${err.message||err}`);
    }
    setProgress(Math.round(((i+1)/files.length)*60));
  }

  // 產出 000_HeaderMap 對照
  appendHeaderMapSheet(wb, headerAudit, groupMap);

  // 依組別輸出分頁：001、002、003…
  let sheetIndex=1;
  for (const [, group] of groupMap){
    const sheetName = uniqueSheetNameNumeric(sheetIndex++);
    const headers = group.header;

    const aoa=[headers];
    group.parts.forEach((part,idx)=>{
      for (const row of part.rows) aoa.push(headers.map(h => row[h] ?? ''));
      if (idx !== group.parts.length-1) aoa.push(new Array(headers.length).fill('')); // 分隔空白列
    });

    const ws=XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols']=autoColumnWidths(aoa, SAMPLE_ROWS_FOR_WIDTH);

    const textUnion=new Set(), numUnion=new Set();
    group.parts.forEach(p => { p.textCols.forEach(h=>textUnion.add(h)); p.numCols.forEach(h=>numUnion.add(h)); });
    forceTextCells(ws, headers, Array.from(textUnion), aoa.length);
    formatAmountCells(ws, headers, aoa.length);
    ws['!autofilter']={ ref: XLSX.utils.encode_range({ s:{c:0,r:0}, e:{c:headers.length-1, r:Math.max(0,aoa.length-1)} }) };

    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    log(`🧾 產生工作表：${sheetName}（${group.parts.length} 段）`);
  }

  // ★ 檔名優先使用使用者輸入（兩模式一致）
  let outName = (mergeFilename && mergeFilename.value || '').trim();
  if (!outName) outName = `依標題合併_${new Date().toISOString().slice(0,10)}`;
  if (!/\.xlsx$/i.test(outName)) outName += '.xlsx';

  const buf=XLSX.write(wb,{bookType:'xlsx',type:'array'});
  saveAs(new Blob([buf],{type:'application/octet-stream'}), outName);

  log(`✅ 合併完成，共 ${groupMap.size} 個分頁`);
  showToast('依標題合併完成！');
  setProgress(0);
}

// ===== K. HeaderMap 對照表 =====
function appendHeaderMapSheet(wb, headerAudit, groupMap){
  const aoa=[["來源檔名","原始標題","規格化後標題","簽名/分頁"]];
  headerAudit.forEach(item=>{
    const sigOrSheet = groupMap ? (sheetNameBySig(groupMap, item.sig) || item.sig) : item.sig;
    if (!item.map || item.map.length===0){
      aoa.push([item.src,"(無標題)","(無)",""]);
    }else{
      item.map.forEach(([raw,canon],idx)=>{
        aoa.push([ idx===0?item.src:"", raw, canon, idx===0?sigOrSheet:"" ]);
      });
    }
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols']=autoColumnWidths(aoa, 100);
  XLSX.utils.book_append_sheet(wb, ws, '000_HeaderMap');
}
function sheetNameBySig(groupMap, sig){
  let i=1; for (const [k] of groupMap.entries()){ if (k===sig) return String(i).padStart(3,'0'); i++; }
  return null;
}

// ===== L. 啟動與事件綁定（init） =====
function init(){
  // 主題切換（按鈕）
  const themeKey = 'csv2excel_theme';
  function applyTheme(name){
    document.body.setAttribute('data-theme', name);
    try{ localStorage.setItem(themeKey, name); }catch(_){ }
    // 視覺高亮目前主題
    [btnThemeGithub, btnThemePolice].forEach(btn => btn && btn.classList.remove('active'));
    if (name==='github' && btnThemeGithub) btnThemeGithub.classList.add('active');
    if (name==='police' && btnThemePolice) btnThemePolice.classList.add('active');
  }
  const saved = localStorage.getItem(themeKey) || 'linkedin';
  applyTheme(saved);
  btnThemeGithub && btnThemeGithub.addEventListener('click', () => applyTheme('github'));
  btnThemePolice && btnThemePolice.addEventListener('click', () => applyTheme('police'));

  // 檔案挑選
  btnPick.addEventListener('click', () => picker.click());
  picker.addEventListener('change', (e) => handleFiles(e.target.files));

  // 【優化：拖曳區點擊觸發選擇器】
  dropzone.addEventListener('click', () => picker.click());

  // 合併開關容器
  const mergeSwitch = document.getElementById('mergeSwitchContainer');

  // 主要動作
  btnStart.addEventListener('click', () => {
    // 【優化：原模式啟動時，高亮合併開關】
    if (mergeSwitch) mergeSwitch.style.opacity = '1';
    startConversion();
  });
  btnGroupByHeader.addEventListener('click', () => {
    // 【優化：依標題合併時，淡出合併開關】
    if (mergeSwitch) mergeSwitch.style.opacity = '0.35';
    groupByHeaderConversion();
  });

  // 清除
  btnClear.addEventListener('click', () => {
    fileMap.clear(); duplicates.clear();
    resetTotals();
    renderFileList(); renderChips();
    log('🧹 已清除清單與統計');
  });

  // 【優化：複製日誌】
  btnCopyLog && btnCopyLog.addEventListener('click', () => {
    const logContent = Array.from(logBox.childNodes).map(n => n.textContent).join('\n');
    navigator.clipboard.writeText(logContent).then(() => {
      showToast('日誌已複製到剪貼簿');
    }).catch(err => {
      showToast('複製日誌失敗: ' + err.message);
    });
  });

  // DnD
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

  // 初始狀態
  setProgress(0);
  renderChips();
  log('✅ 準備就緒');
}

// 外部函式庫（由 HTML 載入）：XLSX / Papa / saveAs
// 確保已載入後再 init（此處假設在頁尾載入 app.js）
init();