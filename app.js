(function () {
    const STORAGE_KEY = 'dataextract_300_records_v4';
    const MASTER_KEY = 'dataextract_300_master_v2';
    
    const state = { 
        files: [], 
        masterData: loadMasterData(),
        records: loadLocalRecords() 
    };

    const els = {
        dropZone: document.getElementById('dropZone'),
        fileInput: document.getElementById('fileInput'),
        masterFileInput: document.getElementById('masterFileInput'),
        masterStatus: document.getElementById('masterStatus'),
        clearBtn: document.getElementById('clearBtn'),
        exportBtn: document.getElementById('exportBtn'),
        addBtn: document.getElementById('addBtn'),
        totalRecords: document.getElementById('totalRecords'),
        duplicateCount: document.getElementById('duplicateCount'),
        missingCount: document.getElementById('missingCount'),
        progressContainer: document.getElementById('progressContainer'),
        progressBar: document.getElementById('progressBar'),
        progressText: document.getElementById('progressText'),
        tableBody: document.getElementById('tableBody')
    };

    function loadLocalRecords() {
        try {
            const saved = localStorage.getItem(STORAGE_KEY);
            return saved ? JSON.parse(saved) : [];
        } catch (e) { return []; }
    }

    function saveLocalRecords() {
        try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state.records)); } catch (e) {}
    }

    function loadMasterData() {
        try {
            const saved = localStorage.getItem(MASTER_KEY);
            return saved ? JSON.parse(saved) : [];
        } catch (e) { return []; }
    }

    function saveMasterData(data) {
        try { localStorage.setItem(MASTER_KEY, JSON.stringify(data)); state.masterData = data; } catch (e) {}
    }

    function init() {
        if (!els.dropZone) return;

        // Update master status badge on load
        updateMasterStatusUI();

        if (els.masterFileInput) {
            els.masterFileInput.addEventListener('change', handleMasterUpload);
        }

        els.dropZone.addEventListener('click', () => {
            if (state.masterData.length === 0) {
                alert('Please upload your Master List Excel file first before uploading slips.');
                if (els.masterFileInput) els.masterFileInput.focus();
                return;
            }
            els.fileInput.click();
        });

        els.fileInput.addEventListener('change', (e) => handleFiles(e.target.files));
        
        els.dropZone.addEventListener('dragover', (e) => { e.preventDefault(); els.dropZone.classList.add('bg-light'); });
        els.dropZone.addEventListener('dragleave', () => els.dropZone.classList.remove('bg-light'));
        els.dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            els.dropZone.classList.remove('bg-light');
            if (state.masterData.length === 0) {
                alert('Please upload your Master List Excel file first before dropping slips.');
                return;
            }
            handleFiles(e.dataTransfer.files);
        });

        if (els.clearBtn) els.clearBtn.addEventListener('click', clearAll);
        if (els.addBtn) els.addBtn.addEventListener('click', addRecord);
        if (els.exportBtn) els.exportBtn.addEventListener('click', exportToExcel);
        
        updateStats();
    }

    function updateMasterStatusUI() {
        if (!els.masterStatus) return;
        if (state.masterData.length > 0) {
            els.masterStatus.className = 'mt-2 small fw-semibold text-success';
            els.masterStatus.innerHTML = `<i class="fa-solid fa-circle-check me-1"></i> Master list loaded successfully (${state.masterData.length} records active).`;
        } else {
            els.masterStatus.className = 'mt-2 small fw-semibold text-danger';
            els.masterStatus.innerHTML = `<i class="fa-solid fa-circle-exclamation me-1"></i> Master list not loaded yet. Please upload your master Excel file first.`;
        }
    }

    function handleMasterUpload(e) {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function(evt) {
            try {
                const data = new Uint8Array(evt.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const json = XLSX.utils.sheet_to_json(firstSheet);
                
                const normalized = json.map(row => {
                    const name = String(row['NAME'] || row['Name'] || '').trim().toUpperCase();
                    const phone = String(row['PHONE'] || row['Phone'] || '').trim();
                    const lg = String(row['LG'] || row['Lg'] || row['LGA'] || '').trim();
                    const ward = String(row['WARD'] || row['Ward'] || '').trim();
                    const position = String(row['POSITION'] || row['Position'] || '').trim();
                    const appNo = String(row['APPLICATION NO'] || row['Application No'] || row['APPLICATION_NO'] || '').trim();
                    return { name, phone, lg, ward, position, applicationNo: appNo };
                });

                saveMasterData(normalized);
                updateMasterStatusUI();
                alert(`Master list loaded with ${normalized.length} records! You can now upload PDF/image slips.`);
            } catch (err) {
                alert('Error parsing master Excel: ' + err.message);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function handleFiles(fileList) {
        if (!fileList || fileList.length === 0) return;
        for (let i = 0; i < fileList.length; i++) {
            state.files.push(fileList[i]);
        }
        updateStats();
        processDocuments();
    }

    function updateStats() {
        els.totalRecords.textContent = state.records.length;
        
        const seen = new Set();
        let dups = 0;
        state.records.forEach(r => {
            const key = r.applicationNo || r.phone;
            if (key) {
                if (seen.has(key)) dups++;
                else seen.add(key);
            }
        });
        els.duplicateCount.textContent = dups;

        let missing = 0;
        state.records.forEach(r => {
            if (!r.name || !r.phone || !r.lg || !r.position || !r.applicationNo) missing++;
        });
        els.missingCount.textContent = missing;
        
        saveLocalRecords();
        renderTable();
    }

    async function processDocuments() {
        if (state.files.length === 0) return;
        if (els.progressContainer) els.progressContainer.style.display = 'block';

        for (let i = 0; i < state.files.length; i++) {
            const file = state.files[i];
            const pct = Math.round(((i + 1) / state.files.length) * 100);
            if (els.progressBar) els.progressBar.style.width = pct + '%';
            if (els.progressText) els.progressText.textContent = `Processing file ${i + 1} of ${state.files.length}: ${file.name}`;

            try {
                let text = '';
                if (file.type === 'application/pdf' || file.name.endsWith('.pdf')) {
                    text = await extractTextFromPDF(file);
                } else {
                    text = await extractTextFromImage(file);
                }
                
                let record = parseFlexibleText(text);

                // Auto-match with master list
                if (state.masterData.length > 0) {
                    const matched = state.masterData.find(m => 
                        (record.applicationNo && m.applicationNo && m.applicationNo.toLowerCase() === record.applicationNo.toLowerCase()) ||
                        (record.phone && m.phone && m.phone.includes(record.phone.slice(-9))) ||
                        (record.name && m.name && m.name.includes(record.name))
                    );
                    if (matched) {
                        if (!record.applicationNo) record.applicationNo = matched.applicationNo;
                        if (!record.name) record.name = matched.name;
                        if (!record.phone) record.phone = matched.phone;
                        if (!record.lg) record.lg = matched.lg;
                        if (!record.ward) record.ward = matched.ward;
                        if (!record.position) record.position = matched.position;
                    }
                }

                state.records.push(record);
            } catch (err) {
                console.error(`Error processing ${file.name}:`, err);
                state.records.push({
                    name: `Error: ${file.name}`,
                    phone: '', lg: '', ward: '', position: '', institution: '', applicationNo: ''
                });
            }
        }

        if (els.progressText) els.progressText.textContent = 'Processing completed successfully!';
        setTimeout(() => {
            if (els.progressContainer) els.progressContainer.style.display = 'none';
        }, 1500);
        
        state.files = [];
        if (els.fileInput) els.fileInput.value = '';
        updateStats();
    }

    async function extractTextFromPDF(file) {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        let fullText = '';
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const tokenized = await page.getTextContent();
            fullText += tokenized.items.map(item => item.str).join(' ') + '\n';
        }
        return fullText;
    }

    async function extractTextFromImage(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = async function (e) {
                const img = new Image();
                img.onload = async function () {
                    const canvas = document.createElement('canvas');
                    let width = img.width;
                    let height = img.height;
                    const maxDim = 1400;
                    if (width > maxDim || height > maxDim) {
                        if (width > height) {
                            height = Math.round((height * maxDim) / width);
                            width = maxDim;
                        } else {
                            width = Math.round((width * maxDim) / height);
                            height = maxDim;
                        }
                    }
                    canvas.width = width;
                    canvas.height = height;
                    const ctx = canvas.getContext('2d');
                    ctx.drawImage(img, 0, 0, width, height);

                    try {
                        const result = await Tesseract.recognize(canvas, 'eng');
                        resolve(result.data.text);
                    } catch (err) {
                        reject(err);
                    }
                };
                img.onerror = () => reject(new Error('Failed to load image element'));
                img.src = e.target.result;
            };
            reader.onerror = () => reject(new Error('Failed to read file reader'));
            reader.readAsDataURL(file);
        });
    }

    function parseFlexibleText(text) {
        const cleanText = text.replace(/\s+/g, ' ');

        let appNo = "";
        const appMatch = cleanText.match(/(ADA_[0-9]+)/i);
        if (appMatch) appNo = appMatch[1];

        let phone = "";
        const phoneMatch = cleanText.match(/\b(0[789][0-9]{9})\b/);
        if (phoneMatch) phone = phoneMatch[1];

        let position = "";
        const upper = cleanText.toUpperCase();
        if (upper.includes('ASSISTANT') || upper.includes('APO')) {
            position = 'APO';
        } else if (upper.includes('PRESIDING') || upper.includes('PO')) {
            position = 'PO';
        }

        let lg = "";
        const lgas = ['DEMSA', 'SONG', 'YOLA NORTH', 'YOLA SOUTH', 'GIREI', 'FUFORE', 'MAFA', 'GOMBI', 'HONG', 'GANYE', 'TOUNGO', 'JADA', 'MAUYU', 'MICHETA', 'MADAGALI', 'MUBI NORTH', 'MUBI SOUTH', 'NUMAN', 'LAMURDE', 'DIEMA', 'GARKIDA'];
        for (let l of lgas) {
            if (upper.includes(l)) { lg = l; break; }
        }

        return {
            name: '',
            phone: phone,
            lg: lg,
            ward: '',
            position: position,
            institution: '',
            applicationNo: appNo
        };
    }

    function renderTable() {
        if (!els.tableBody) return;
        els.tableBody.innerHTML = '';
        if (state.records.length === 0) {
            els.tableBody.innerHTML = `<tr><td colspan="9" class="text-center text-muted py-5">No records saved. Upload your Master List and Slips above to begin.</td></tr>`;
            return;
        }

        state.records.forEach((r, idx) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td class="fw-bold text-center">${idx + 1}</td>
                <td><input type="text" value="${escapeHtml(r.name)}" data-index="${idx}" data-field="name" class="form-control form-control-sm table-input" placeholder="Name"></td>
                <td><input type="text" value="${escapeHtml(r.phone)}" data-index="${idx}" data-field="phone" class="form-control form-control-sm table-input" placeholder="Phone"></td>
                <td><input type="text" value="${escapeHtml(r.lg)}" data-index="${idx}" data-field="lg" class="form-control form-control-sm table-input" placeholder="LG"></td>
                <td><input type="text" value="${escapeHtml(r.ward)}" data-index="${idx}" data-field="ward" class="form-control form-control-sm table-input" placeholder="Ward"></td>
                <td><input type="text" value="${escapeHtml(r.position)}" data-index="${idx}" data-field="position" class="form-control form-control-sm table-input" placeholder="PO/APO"></td>
                <td><input type="text" value="${escapeHtml(r.institution)}" data-index="${idx}" data-field="institution" class="form-control form-control-sm table-input" placeholder="Institution"></td>
                <td><input type="text" value="${escapeHtml(r.applicationNo)}" data-index="${idx}" data-field="applicationNo" class="form-control form-control-sm table-input ${!r.applicationNo ? 'border-danger' : ''}" placeholder="Application No"></td>
                <td class="text-center"><button class="btn btn-sm btn-outline-danger delete-row px-2 py-0" data-index="${idx}" title="Delete Row">&times;</button></td>
            `;
            els.tableBody.appendChild(tr);
        });

        document.querySelectorAll('.table-input').forEach(input => {
            input.addEventListener('input', (e) => {
                state.records[e.target.dataset.index][e.target.dataset.field] = e.target.value;
                saveLocalRecords();
            });
        });

        document.querySelectorAll('.delete-row').forEach(btn => {
            btn.addEventListener('click', (e) => {
                state.records.splice(e.target.dataset.index, 1);
                updateStats();
            });
        });
    }

    function addRecord() {
        state.records.push({ name: '', phone: '', lg: '', ward: '', position: '', institution: '', applicationNo: '' });
        updateStats();
    }

    function clearAll() {
        if (confirm('Are you sure you want to clear all saved records?')) {
            state.files = [];
            state.records = [];
            localStorage.removeItem(STORAGE_KEY);
            if (els.fileInput) els.fileInput.value = '';
            updateStats();
        }
    }

    function exportToExcel() {
        if (!state.records.length) { alert('No records to export.'); return; }
        const header = ['SN', 'Name', 'Phone', 'Lg', 'Ward', 'Position', 'Name of institution', 'Application No'];
        const rows = [header];
        state.records.forEach((r, i) => {
            rows.push([i + 1, r.name || '', r.phone || '', r.lg || '', r.ward || '', r.position || '', r.institution || '', r.applicationNo || '']);
        });

        const ws = XLSX.utils.aoa_to_sheet(rows);
        if (ws['!ref']) {
            const range = XLSX.utils.decode_range(ws['!ref']);
            for (let R = 1; R <= range.e.r; R++) {
                const addr = XLSX.utils.encode_cell({ r: R, c: 2 });
                if (ws[addr]) { ws[addr].t = 's'; ws[addr].v = String(ws[addr].v); ws[addr].z = '@'; }
            }
        }
        ws['!cols'] = [{ wch: 5 }, { wch: 24 }, { wch: 16 }, { wch: 16 }, { wch: 14 }, { wch: 12 }, { wch: 30 }, { wch: 20 }];
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Records');
        XLSX.writeFile(wb, 'DataExtract_' + new Date().toISOString().slice(0, 10) + '.xlsx');
    }

    function escapeHtml(str) {
        if (!str) return '';
        return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    document.addEventListener('DOMContentLoaded', init);
})();
