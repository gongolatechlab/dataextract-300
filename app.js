(function () {
    const MASTER_KEY = 'dataextract_300_enriched_master_v1';
    
    const state = { 
        files: [], 
        records: loadLocalRecords(),
        unmatchedFiles: []
    };

    const els = {
        dropZone: document.getElementById('dropZone'),
        fileInput: document.getElementById('fileInput'),
        masterFileInput: document.getElementById('masterFileInput'),
        masterStatus: document.getElementById('masterStatus'),
        clearBtn: document.getElementById('clearBtn'),
        exportBtn: document.getElementById('exportBtn'),
        filterSelect: document.getElementById('filterSelect'),
        totalRecords: document.getElementById('totalRecords'),
        matchedCount: document.getElementById('matchedCount'),
        missingCount: document.getElementById('missingCount'),
        unmatchedAlert: document.getElementById('unmatchedAlert'),
        unmatchedList: document.getElementById('unmatchedList'),
        progressContainer: document.getElementById('progressContainer'),
        progressBar: document.getElementById('progressBar'),
        progressText: document.getElementById('progressText'),
        tableBody: document.getElementById('tableBody')
    };

    function loadLocalRecords() {
        try {
            const saved = localStorage.getItem(MASTER_KEY);
            return saved ? JSON.parse(saved) : [];
        } catch (e) { return []; }
    }

    function saveLocalRecords() {
        try { localStorage.setItem(MASTER_KEY, JSON.stringify(state.records)); } catch (e) {}
    }

    function init() {
        if (!els.dropZone) return;

        updateMasterStatusUI();

        if (els.masterFileInput) els.masterFileInput.addEventListener('change', handleMasterUpload);
        if (els.filterSelect) els.filterSelect.addEventListener('change', renderTable);

        els.dropZone.addEventListener('click', () => {
            if (state.records.length === 0) {
                alert('Please upload your Master List Excel file in Step 1 first.');
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
            if (state.records.length === 0) {
                alert('Please upload your Master List Excel file in Step 1 first.');
                return;
            }
            handleFiles(e.dataTransfer.files);
        });

        if (els.clearBtn) els.clearBtn.addEventListener('click', clearAll);
        if (els.exportBtn) els.exportBtn.addEventListener('click', exportToExcel);
        
        updateStats();
    }

    function updateMasterStatusUI() {
        if (!els.masterStatus) return;
        if (state.records.length > 0) {
            els.masterStatus.className = 'mt-2 small fw-semibold text-success';
            els.masterStatus.innerHTML = `<i class="fa-solid fa-circle-check me-1"></i> Master list active (${state.records.length} records loaded). Ready to match slips.`;
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
                
                // Preserve exact master columns and normalize phone for reliable matching
                state.records = json.map((row, idx) => {
                    const sn = row['S/N'] || row['SN'] || (idx + 1);
                    const name = String(row['NAME'] || row['Name'] || '').trim().toUpperCase();
                    let rawPhone = String(row['PHONE'] || row['Phone'] || '').trim();
                    // Ensure leading zero if stripped by Excel (e.g., 9022729924 -> 09022729924)
                    if (rawPhone.length === 10 && !rawPhone.startsWith('0')) {
                        rawPhone = '0' + rawPhone;
                    }
                    const lg = String(row['LG'] || row['Lg'] || row['LGA'] || '').trim().toUpperCase();
                    const ward = String(row['WARD'] || row['Ward'] || '').trim().toUpperCase();
                    const position = String(row['POSITION'] || row['Position'] || '').trim().toUpperCase();
                    const qualification = String(row['QUALIFICATION'] || row['Qualification'] || '').trim().toUpperCase();
                    const applicationNo = String(row['APPLICATION NO'] || row['Application No'] || row['APPLICATION_NO'] || '').trim();
                    
                    return { sn, name, phone: rawPhone, lg, ward, position, qualification, applicationNo };
                });

                saveLocalRecords();
                updateMasterStatusUI();
                updateStats();
                alert(`Loaded ${state.records.length} records from Master List! Now upload your PDFs or images in Step 2.`);
            } catch (err) {
                alert('Error reading Master Excel file: ' + err.message);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function handleFiles(fileList) {
        if (!fileList || fileList.length === 0) return;
        for (let i = 0; i < fileList.length; i++) {
            state.files.push(fileList[i]);
        }
        processDocuments();
    }

    function updateStats() {
        const total = state.records.length;
        const matched = state.records.filter(r => r.applicationNo && r.applicationNo.trim() !== '').length;
        const missing = total - matched;

        if (els.totalRecords) els.totalRecords.textContent = total;
        if (els.matchedCount) els.matchedCount.textContent = matched;
        if (els.missingCount) els.missingCount.textContent = missing;
        
        saveLocalRecords();
        renderTable();
    }

    async function processDocuments() {
        if (state.files.length === 0) return;
        if (els.progressContainer) els.progressContainer.style.display = 'block';
        state.unmatchedFiles = [];

        for (let i = 0; i < state.files.length; i++) {
            const file = state.files[i];
            const pct = Math.round(((i + 1) / state.files.length) * 100);
            if (els.progressBar) els.progressBar.style.width = pct + '%';
            if (els.progressText) els.progressText.textContent = `Scanning ${i + 1} of ${state.files.length}: ${file.name}`;

            try {
                let pdfFile = file;
                if (!file.type.includes('pdf') && !file.name.endsWith('.pdf')) {
                    if (els.progressText) els.progressText.textContent = `Converting image to PDF: ${file.name}`;
                    pdfFile = await convertImageToPDF(file);
                }

                const text = await extractTextFromPDF(pdfFile);
                const extracted = parseFlexibleText(text);

                // Search master list for a match by Phone (last 9 digits) OR Name overlap
                let foundMatch = false;
                if (extracted.applicationNo) {
                    const cleanUpperText = text.toUpperCase().replace(/\s+/g, ' ');
                    
                    for (let r of state.records) {
                        const phoneTail = r.phone ? r.phone.slice(-9) : '';
                        const phoneMatch = (extracted.phone && phoneTail && extracted.phone.includes(phoneTail)) ||
                                           (phoneTail && cleanUpperText.includes(phoneTail));
                        
                        // Check if full name or at least 2 name tokens appear in the slip text
                        const nameTokens = r.name.split(/\s+/).filter(t => t.length > 2);
                        const nameMatch = r.name && (
                            cleanUpperText.includes(r.name) || 
                            (nameTokens.length >= 2 && nameTokens.every(tok => cleanUpperText.includes(tok)))
                        );

                        if (phoneMatch || nameMatch) {
                            r.applicationNo = extracted.applicationNo;
                            foundMatch = true;
                            break;
                        }
                    }
                }

                if (!foundMatch) {
                    state.unmatchedFiles.push(`${file.name}${extracted.applicationNo ? ' (' + extracted.applicationNo + ')' : ''}`);
                }
            } catch (err) {
                console.error(`Error processing ${file.name}:`, err);
                state.unmatchedFiles.push(`${file.name} (Scan Error)`);
            }
        }

        if (state.unmatchedFiles.length > 0 && els.unmatchedAlert && els.unmatchedList) {
            els.unmatchedList.textContent = state.unmatchedFiles.join(', ');
            els.unmatchedAlert.style.display = 'block';
        }

        if (els.progressText) els.progressText.textContent = 'Matching complete!';
        setTimeout(() => {
            if (els.progressContainer) els.progressContainer.style.display = 'none';
        }, 1500);
        
        state.files = [];
        if (els.fileInput) els.fileInput.value = '';
        updateStats();
    }

    async function convertImageToPDF(imageFile) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = function (e) {
                const img = new Image();
                img.onload = function () {
                    try {
                        const { jsPDF } = window.jspdf;
                        const doc = new jsPDF({
                            orientation: img.width > img.height ? 'landscape' : 'portrait',
                            unit: 'px',
                            format: [img.width, img.height]
                        });
                        doc.addImage(img, 'JPEG', 0, 0, img.width, img.height);
                        const pdfOutput = doc.output('blob');
                        resolve(new File([pdfOutput], imageFile.name.replace(/\.[^/.]+$/, "") + '.pdf', { type: 'application/pdf' }));
                    } catch (err) { reject(err); }
                };
                img.onerror = reject;
                img.src = e.target.result;
            };
            reader.onerror = reject;
            reader.readAsDataURL(imageFile);
        });
    }

    async function extractTextFromPDF(file) {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        let fullText = '';
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const tokenized = await page.getTextContent();
            if (tokenized.items.length === 0) {
                const viewport = page.getViewport({ scale: 2.0 });
                const canvas = document.createElement('canvas');
                canvas.width = viewport.width;
                canvas.height = viewport.height;
                const ctx = canvas.getContext('2d');
                await page.render({ canvasContext: ctx, viewport: viewport }).promise;
                const result = await Tesseract.recognize(canvas, 'eng');
                fullText += result.data.text + '\n';
            } else {
                fullText += tokenized.items.map(item => item.str).join(' ') + '\n';
            }
        }
        return fullText;
    }

    function parseFlexibleText(text) {
        const cleanText = text.replace(/\s+/g, ' ');
        let appNo = "";
        const appMatch = cleanText.match(/(ADA_[0-9]+)/i);
        if (appMatch) appNo = appMatch[1].toUpperCase();

        let phone = "";
        const phoneMatch = cleanText.match(/\b(0?[789][0-9]{9})\b/);
        if (phoneMatch) phone = phoneMatch[1];

        return { applicationNo: appNo, phone: phone };
    }

    function renderTable() {
        if (!els.tableBody) return;
        els.tableBody.innerHTML = '';
        
        const filter = els.filterSelect ? els.filterSelect.value : 'all';
        const filteredRecords = state.records.map((r, originalIdx) => ({ ...r, originalIdx })).filter(r => {
            const hasId = r.applicationNo && r.applicationNo.trim() !== '';
            if (filter === 'matched') return hasId;
            if (filter === 'missing') return !hasId;
            return true;
        });

        if (filteredRecords.length === 0) {
            els.tableBody.innerHTML = `<tr><td colspan="8" class="text-center text-muted py-5">No records match the current view.</td></tr>`;
            return;
        }

        filteredRecords.forEach((r) => {
            const hasId = r.applicationNo && r.applicationNo.trim() !== '';
            const tr = document.createElement('tr');
            if (hasId) tr.className = 'row-matched';
            
            tr.innerHTML = `
                <td class="fw-bold text-center">${r.sn}</td>
                <td><input type="text" value="${escapeHtml(r.name)}" data-index="${r.originalIdx}" data-field="name" class="table-input"></td>
                <td><input type="text" value="${escapeHtml(r.phone)}" data-index="${r.originalIdx}" data-field="phone" class="table-input"></td>
                <td><input type="text" value="${escapeHtml(r.lg)}" data-index="${r.originalIdx}" data-field="lg" class="table-input"></td>
                <td><input type="text" value="${escapeHtml(r.ward)}" data-index="${r.originalIdx}" data-field="ward" class="table-input"></td>
                <td><input type="text" value="${escapeHtml(r.position)}" data-index="${r.originalIdx}" data-field="position" class="table-input"></td>
                <td><input type="text" value="${escapeHtml(r.qualification)}" data-index="${r.originalIdx}" data-field="qualification" class="table-input"></td>
                <td><input type="text" value="${escapeHtml(r.applicationNo)}" data-index="${r.originalIdx}" data-field="applicationNo" class="table-input fw-bold ${hasId ? 'text-success' : 'border-danger'}" placeholder="Missing ID"></td>
            `;
            els.tableBody.appendChild(tr);
        });

        document.querySelectorAll('.table-input').forEach(input => {
            input.addEventListener('input', (e) => {
                const idx = parseInt(e.target.dataset.index);
                const field = e.target.dataset.field;
                state.records[idx][field] = e.target.value;
                updateStats();
            });
        });
    }

    function clearAll() {
        if (confirm('Are you sure you want to reset everything and clear saved data?')) {
            state.files = [];
            state.records = [];
            state.unmatchedFiles = [];
            localStorage.removeItem(MASTER_KEY);
            if (els.fileInput) els.fileInput.value = '';
            if (els.masterFileInput) els.masterFileInput.value = '';
            if (els.unmatchedAlert) els.unmatchedAlert.style.display = 'none';
            updateMasterStatusUI();
            updateStats();
        }
    }

    // Export ONLY complete rows that have an Application ID, maintaining exact Master List columns
    function exportToExcel() {
        const completeRecords = state.records.filter(r => r.applicationNo && r.applicationNo.trim() !== '');
        
        if (completeRecords.length === 0) {
            alert('No matched records with an Application ID found yet! Upload PDFs/images in Step 2 first so they match your Master List.');
            return;
        }

        const header = ['S/N', 'NAME', 'PHONE', 'LG', 'WARD', 'POSITION', 'QUALIFICATION', 'APPLICATION NO'];
        const rows = [header];

        completeRecords.forEach((r, i) => {
            rows.push([
                i + 1,
                r.name || '',
                r.phone || '',
                r.lg || '',
                r.ward || '',
                r.position || '',
                r.qualification || '',
                r.applicationNo || ''
            ]);
        });

        const ws = XLSX.utils.aoa_to_sheet(rows);
        if (ws['!ref']) {
            const range = XLSX.utils.decode_range(ws['!ref']);
            for (let R = 1; R <= range.e.r; R++) {
                const addr = XLSX.utils.encode_cell({ r: R, c: 2 });
                if (ws[addr]) { ws[addr].t = 's'; ws[addr].v = String(ws[addr].v); ws[addr].z = '@'; }
            }
        }
        ws['!cols'] = [{ wch: 6 }, { wch: 25 }, { wch: 15 }, { wch: 14 }, { wch: 18 }, { wch: 12 }, { wch: 15 }, { wch: 20 }];
        
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Enriched Records');
        XLSX.writeFile(wb, 'Tangwami_INEC_Enriched_' + new Date().toISOString().slice(0, 10) + '.xlsx');
    }

    function escapeHtml(str) {
        if (!str) return '';
        return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    document.addEventListener('DOMContentLoaded', init);
})();
