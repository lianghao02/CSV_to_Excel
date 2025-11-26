document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const fileListEl = document.getElementById('file-list');
    const convertBtn = document.getElementById('convert-btn');
    let filesToProcess = [];

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
                <span class="remove-btn" onclick="removeFile(${index})">×</span>
            `;
            fileListEl.appendChild(item);
        });
    }

    window.removeFile = (index) => {
        filesToProcess.splice(index, 1);
        updateFileList();
        updateConvertBtn();
    };

    function updateConvertBtn() {
        convertBtn.disabled = filesToProcess.length === 0;
        convertBtn.textContent = filesToProcess.length === 0 ? '開始轉換' : `轉換 ${filesToProcess.length} 個檔案`;
    }

    convertBtn.addEventListener('click', async () => {
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
    });

    async function readCSV(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = (e) => reject(e);
            // Use 'big5' encoding for Traditional Chinese CSVs commonly found in Taiwan finance
            // If that fails or looks wrong, we might need 'utf-8'. 
            // However, most banking CSVs in Taiwan are Big5. 
            // Let's try to detect or default to Big5. 
            // Actually, the user's sample file might be UTF-8 or Big5. 
            // Let's assume UTF-8 first as it's standard, but if it looks garbled we might need to switch.
            // Given the user's previous request was about "Traditional Chinese", Big5 is a strong candidate for legacy systems,
            // but modern exports might be UTF-8. 
            // Let's use 'Big5' as a safe bet for Taiwan banking CSVs if they are legacy, 
            // but 'utf-8' is safer for modern web. 
            // Wait, the user provided a sample file content in the prompt history which looked like UTF-8 (readable characters).
            // I will use 'utf-8' for now.
            reader.readAsText(file, 'utf-8');
        });
    }

    function parseCSV(text) {
        // Simple CSV parser that handles quotes
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

                // Rule: Remove '+' and '.00' from amounts
                // Regex looks for optional sign, digits, optional .00
                // Actually user said: "原檔會有+以及小數點.00(請把+號跟.00刪除，留下整數數字)"
                // Example: "+17030.00" -> "17030"
                // Example: "0.00" -> "0"
                // Example: "+0.00" -> "0"

                if (/^[+]?\d+\.00$/.test(val)) {
                    return val.replace('+', '').replace('.00', '');
                }

                // Also handle negative numbers if they exist? User didn't specify, but usually they are just -100.00
                // User specifically mentioned + and .00.
                // Let's be strict to user request: remove + and .00

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

            // Force all cells to be text to prevent auto-formatting (stripping 0s)
            // We can do this by setting cell type to 's' (string) for all cells
            // However, aoa_to_sheet might auto-detect.
            // Better approach: Update the sheet object after creation
            const range = XLSX.utils.decode_range(ws['!ref']);
            for (let R = range.s.r; R <= range.e.r; ++R) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cell_address = XLSX.utils.encode_cell({ r: R, c: C });
                    if (!ws[cell_address]) continue;

                    // If it looks like a number that we want to keep as string (e.g. account number), set type to string
                    // User said: "帳號不要補0(原檔怎麼顯示就怎麼顯示)" -> Keep as string
                    // "交易日期跟交易時間...不要隨意補0" -> Keep as string
                    // "支出金額...留下整數數字" -> This can be number or string. String is safer to preserve exact appearance.

                    // Actually, for Excel to treat it as text, we set .t = 's' and .z = '@'
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

            // Force text format
            const range = XLSX.utils.decode_range(ws['!ref']);
            for (let R = range.s.r; R <= range.e.r; ++R) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cell_address = XLSX.utils.encode_cell({ r: R, c: C });
                    if (!ws[cell_address]) continue;
                    ws[cell_address].t = 's';
                    ws[cell_address].z = '@';
                }
            }

            // Sheet name: 01, 02, 03...
            const sheetName = (i + 1).toString().padStart(2, '0');
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
        }

        // Generate filename: MMDD_HHMM合併金流.xlsx
        const now = new Date();
        const month = (now.getMonth() + 1).toString().padStart(2, '0');
        const date = now.getDate().toString().padStart(2, '0');
        const hours = now.getHours().toString().padStart(2, '0');
        const minutes = now.getMinutes().toString().padStart(2, '0');
        const filename = `${month}${date}_${hours}${minutes}合併金流.xlsx`;

        XLSX.writeFile(wb, filename);
    }
});
