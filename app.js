(function () {
    const state = { files: [], records: [] };
    const els = {
        dropZone: document.getElementById('dropZone'),
        fileInput: document.getElementById('fileInput'),
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

    function init() {
        if (!els.dropZone) return;
        els.dropZone.addEventListener('click', () => els.fileInput.click());
        els.fileInput.addEventListener('change', (e) => handleFiles(e.target.files));
        
        els.dropZone.addEventListener('dragover', (e) => { e.preventDefault(); els.dropZone.classList.add('bg-light'); });
        els.dropZone.addEventListener('dragleave', () => els.dropZone.classList.remove('bg-light'));
        els.dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            els.dropZone.classList.remove('bg-light');
            handleFiles(e.dataTransfer.files);
        });

        if (els.clearBtn) els.clearBtn.addEventListener('click', clearAll);
        if (els.addBtn) els.addBtn.addEventListener('click', addRecord);
        if (els.exportBtn) els.exportBtn.addEventListener('click', exportToExcel);
        updateStats();
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
                const parsed = parseExtractedText(text);
                state.records.push(parsed);
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
        
        // Reset file input so same files can be re-selected if needed
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
                    const maxDim = 1600;
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
                        const result = await Tesseract.recognize(canvas, 'eng', {
                            logger: m => {
                                if (m.status === 'recognizing text' && els.progressText) {
                                    els.progressText.textContent = `OCR Scanning: ${Math.round(m.progress * 100)}% - ${file.name}`;
                                }
                            }
                        });
                        resolve(result.data.text);
                    } catch (err) {
                        reject(err);
                    }
                };
                img.onerror = reject;
                img.src = e.target.result;
            };
            reader.onerror = reject;
            reader.readAsDataURL(file);
        });
    }

    function parseExtractedText(text) {
        const cleanText = text.replace(/\s+/g, ' ');
        const appNoMatch = cleanText.match(/APPLICATION\s*NO[:\s]*([A-Z0-9_]+)/i);
        const nameMatch = cleanText.match(/FULL\s*NAME[:\s]*([A-Z\s]+?)(?=\s*(?:PREFERRED|EMAIL|STATE|APPLIED|LGA|NAME|ACKNOWLEDGEMENT|$))/i);
        const phoneMatch = cleanText.match(/(?:PREFERRED\s*CONTACT\s*NUMBER|PHONE)[:\s]*([0-9]+)/i);
        const lgMatch = cleanText.match(/LGA(?:\s*OF\s*RESIDENCE)?[:\s]*([A-Z\s]+?)(?=\s*(?:APPLIED|APPLICANT|NAME|EMAIL|STATE|ACKNOWLEDGEMENT|$))/i);
        const posMatch = cleanText.match(/APPLIED\s*POSITION[:\s]*([A-Z\s\(\)]+?)(?=\s*(?:APPLICANT|NAME|EMAIL|STATE|LGA|ACKNOWLEDGEMENT|$))/i);
        const instMatch = cleanText.match(/NAME\s*OF\s*INSTITUTION[:\s]*([A-Z0-9,\s]+?)(?=\s*(?:ACKNOWLEDGEMENT|INEC|VERIFICATION|Print|$))/i);

        let position = '';
        if (posMatch) {
            const rawPos = posMatch[1].toUpperCase();
            position = rawPos.includes('ASSISTANT') || rawPos.includes('APO') ? 'APO' : (rawPos.includes('PRESIDING') || rawPos.includes('PO') ? 'PO' : rawPos.trim());
        }

        return {
            name: nameMatch ? nameMatch[1].trim() : '',
            phone: phoneMatch ? phoneMatch[1].trim() : '',
            lg: lgMatch ? lgMatch[1].trim() : '',
            ward: '',
            position: position,
            institution: instMatch ? instMatch[1].replace(/\s+/g, ' ').trim() : '',
            applicationNo: appNoMatch ? appNoMatch[1].trim() : ''
        };
    }

    function renderTable() {
        if (!els.tableBody) return;
        els.tableBody.innerHTML = '';
        if (state.records.length === 0) {
            els.tableBody.innerHTML = `<tr><td colspan="9" class="text-center text-muted py-5">No records loaded yet. Upload your PDF or image slips above to begin extraction.</td></tr>`;
            return;
        }

        state.records.forEach((r, idx) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td class="fw-bold text-center">${idx + 1}</td>
                <td><input type="text" value="${escapeHtml(r.name)}" data-index="${idx}" data-field="name" class="form-control form-control-sm table-input"></td>
                <td><input type="text" value="${escapeHtml(r.phone)}" data-index="${idx}" data-field="phone" class="form-control form-control-sm table-input"></td>
                <td><input type="text" value="${escapeHtml(r.lg)}" data-index="${idx}" data-field="lg" class="form-control form-control-sm table-input"></td>
                <td><input type="text" value="${escapeHtml(r.ward)}" data-index="${idx}" data-field="ward" class="form-control form-control-sm table-input"></td>
                <td><input type="text" value="${escapeHtml(r.position)}" data-index="${idx}" data-field="position" class="form-control form-control-sm table-input"></td>
                <td><input type="text" value="${escapeHtml(r.institution)}" data-index="${idx}" data-field="institution" class="form-control form-control-sm table-input"></td>
                <td><input type="text" value="${escapeHtml(r.applicationNo)}" data-index="${idx}" data-field="applicationNo" class="form-control form-control-sm table-input"></td>
                <td class="text-center"><button class="btn btn-sm btn-outline-danger delete-row px-2 py-0" data-index="${idx}" title="Delete Row">&times;</button></td>
            `;
            els.tableBody.appendChild(tr);
        });

        document.querySelectorAll('.table-input').forEach(input => {
            input.addEventListener('input', (e) => {
                state.records[e.target.dataset.index][e.target.dataset.field] = e.target.value;
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
        if (confirm('Are you sure you want to clear all records?')) {
            state.files = [];
            state.records = [];
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
