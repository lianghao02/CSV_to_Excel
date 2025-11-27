const app = (function () {
    // Private variables
    let dropZone, fileInput, fileListEl, convertBtn;
    let filesToProcess = [];

    function init() {
        dropZone = document.getElementById('drop-zone');
        fileInput = document.getElementById('file-input');
        fileListEl = document.getElementById('file-list');
        convertBtn = document.getElementById('convert-btn');

        setupEventListeners();
    }

    function setupEventListeners() {
        // Drag & Drop Events
        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        });

        dropZone.addEventListener('dragleave', () => {
            dropZone.classList.remove('dragover');
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
            handleFiles(e.dataTransfer.files);
        });

        dropZone.addEventListener('click', () => {
            fileInput.click();
        });

        fileInput.addEventListener('change', (e) => {
            handleFiles(e.target.files);
            fileInput.value = ''; // Reset input
        });

        convertBtn.addEventListener('click', handleConvert);
    }

    function handleFiles(files) {
        for (let file of files) {
            if (file.name.toLowerCase().endsWith('.csv')) {
                // Avoid duplicates
                if (!filesToProcess.some(f => f.name === file.name && f.size === file.size)) {
                    filesToProcess.push(file);
                }
            }
        }
        updateFileList();
        updateConvertBtn();
    }

    function updateFileList() {
        fileListEl.innerHTML = '';
        filesToProcess.forEach((file, index) => {
            const item = document.createElement('div');
            item.className = 'file-item';
            item.innerHTML = `
                <span>${file.name}</span>
                <span class="remove-btn" data-index="${index}">×</span>
            `;
            // Use event delegation or direct binding safely
            item.querySelector('.remove-btn').addEventListener('click', (e) => {
                removeFile(index);
            });
            fileListEl.appendChild(item);
        });
    }

    function removeFile(index) {
        filesToProcess.splice(index, 1);
        updateFileList();
        updateConvertBtn();
    }

    function updateConvertBtn() {
        convertBtn.disabled = filesToProcess.length === 0;
        convertBtn.textContent = filesToProcess.length === 0 ? '開始轉換' : `轉換 ${filesToProcess.length} 個檔案`;
    }

    async function handleConvert() {
        if (filesToProcess.length === 0) return;

        const mode = document.querySelector('input[name="export-mode"]:checked').value;
        convertBtn.disabled = true;
        convertBtn.textContent = '處理中...';

        try {
            if (mode === 'individual') {
                await processIndividual(filesToProcess);
            } else {
                await processMerged(filesToProcess);
            }
            alert('轉換完成！');
        } catch (error) {
            console.error(error);
            alert('轉換過程中發生錯誤：' + error.message);
        } finally {
            convertBtn.disabled = false;
            updateConvertBtn();
        }
    }

    async function readCSV(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = (e) => reject(e);
            reader.readAsText(file, 'utf-8');
        });
    }

    function parseCSV(text) {
        const rows = [];
        let currentRow = [];
        let currentCell = '';
        let inQuotes = false;

        for (let i = 0; i < text.length; i++) {
            const char = text[i];
            const nextChar = text[i + 1];

            if (char === '"') {
                if (inQuotes && nextChar === '"') {
                    currentCell += '"';
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (char === ',' && !inQuotes) {
                currentRow.push(currentCell);
                currentCell = '';
            } else if ((char === '\r' || char === '\n') && !inQuotes) {
                if (char === '\r' && nextChar === '\n') i++;
                if (currentCell || currentRow.length > 0) {
                    currentRow.push(currentCell);
                    rows.push(currentRow);
                }
                currentRow = [];
                currentCell = '';
            } else {
                currentCell += char;
            }
        }
        if (currentCell || currentRow.length > 0) {
            currentRow.push(currentCell);
            rows.push(currentRow);
        }
        return rows;
    }

    function cleanData(rows) {
        return rows.map(row => {
            return row.map(cell => {
                let val = cell.trim();
                if (/^[+]?\d+\.00$/.test(val)) {
                    return val.replace('+', '').replace('.00', '');
                }
                if (val.includes('+')) {
                    val = val.replace(/\+/g, '');
                }
                if (val.endsWith('.00')) {
                    val = val.replace(/\.00$/, '');
                }
                return val;
            });
        });
    }

    async function processIndividual(files) {
        for (let file of files) {
            const text = await readCSV(file);
            const rows = parseCSV(text);
            const cleanedRows = cleanData(rows);

            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(cleanedRows);

            const range = XLSX.utils.decode_range(ws['!ref']);
            for (let R = range.s.r; R <= range.e.r; ++R) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cell_address = XLSX.utils.encode_cell({ r: R, c: C });
                    if (!ws[cell_address]) continue;
                    ws[cell_address].t = 's';
                    ws[cell_address].z = '@';
                }
            }

            XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
            XLSX.writeFile(wb, file.name.replace('.csv', '.xlsx'));
        }
    }

    async function processMerged(files) {
        const wb = XLSX.utils.book_new();

        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            const text = await readCSV(file);
            const rows = parseCSV(text);
            const cleanedRows = cleanData(rows);

            const ws = XLSX.utils.aoa_to_sheet(cleanedRows);

            const range = XLSX.utils.decode_range(ws['!ref']);
            for (let R = range.s.r; R <= range.e.r; ++R) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cell_address = XLSX.utils.encode_cell({ r: R, c: C });
                    if (!ws[cell_address]) continue;
                    ws[cell_address].t = 's';
                    ws[cell_address].z = '@';
                }
            }

            const sheetName = (i + 1).toString().padStart(2, '0');
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
        }

        const now = new Date();
        const month = (now.getMonth() + 1).toString().padStart(2, '0');
        const date = now.getDate().toString().padStart(2, '0');
        const hours = now.getHours().toString().padStart(2, '0');
        const minutes = now.getMinutes().toString().padStart(2, '0');
        const filename = `${month}${date}_${hours}${minutes}合併金流.xlsx`;

        XLSX.writeFile(wb, filename);
    }

    return { init };
})();

document.addEventListener('DOMContentLoaded', app.init);
