
(function () {
    const MASTER_KEY = 'dataextract_300_enriched_master_v1';
    const EXPORT_LIMIT = 310;

    const state = {
        files: [],
        records: loadLocalRecords(),
        unmatchedFiles: [],
        exportSelection: null,
        processing: false
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
        } catch (e) {
            return [];
        }
    }

    function saveLocalRecords() {
        try {
            localStorage.setItem(MASTER_KEY, JSON.stringify(state.records));
        } catch (e) {
            alert('Unable to save records. Your browser storage may be full.');
        }
    }

    function shuffle(array) {
        const result = [...array];

        for (let i = result.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [result[i], result[j]] = [result[j], result[i]];
        }

        return result;
    }

    function generateInternalReference(existing) {
        let ref;

        do {
            ref = 'REF_' +
                Math.floor(10000000000 + Math.random() * 90000000000);
        } while (existing.has(ref));

        existing.add(ref);
        return ref;
    }

    function createGenerateButton() {
        if (document.getElementById('generate310Btn')) return;

        const btn = document.createElement('button');

        btn.id = 'generate310Btn';
        btn.type = 'button';
        btn.className = 'btn btn-primary me-2';
        btn.innerHTML =
            '<i class="fa-solid fa-shuffle me-1"></i> Prepare 310 Records';

        btn.addEventListener('click', generateSelection);

        if (els.exportBtn && els.exportBtn.parentElement) {
            els.exportBtn.parentElement.insertBefore(
                btn,
                els.exportBtn
            );
        } else {
            els.dropZone.parentElement.appendChild(btn);
        }
    }

    function init() {
        if (!els.dropZone) return;

        createGenerateButton();

        updateMasterStatusUI();

        if (els.masterFileInput) {
            els.masterFileInput.addEventListener(
                'change',
                handleMasterUpload
            );
        }

        if (els.filterSelect) {
            els.filterSelect.addEventListener(
                'change',
                renderTable
            );
        }

        els.dropZone.addEventListener('click', () => {
            if (state.records.length === 0) {
                alert('Please upload your Master List Excel file first.');
                els.masterFileInput?.focus();
                return;
            }

            els.fileInput.click();
        });

        els.fileInput.addEventListener('change', e => {
            handleFiles(e.target.files);
        });

        els.dropZone.addEventListener('dragover', e => {
            e.preventDefault();
            els.dropZone.classList.add('bg-light');
        });

        els.dropZone.addEventListener('dragleave', () => {
            els.dropZone.classList.remove('bg-light');
        });

        els.dropZone.addEventListener('drop', e => {
            e.preventDefault();
            els.dropZone.classList.remove('bg-light');

            if (state.records.length === 0) {
                alert('Please upload your Master List Excel file first.');
                return;
            }

            handleFiles(e.dataTransfer.files);
        });

        els.clearBtn?.addEventListener('click', clearAll);
        els.exportBtn?.addEventListener('click', exportToExcel);

        updateStats();
    }

    function updateMasterStatusUI() {
        if (!els.masterStatus) return;

        if (state.records.length > 0) {
            els.masterStatus.className =
                'mt-2 small fw-semibold text-success';

            els.masterStatus.innerHTML =
                `<i class="fa-solid fa-circle-check me-1"></i>
                Master list active (${state.records.length} records loaded).
                Ready to match slips.`;
        } else {
            els.masterStatus.className =
                'mt-2 small fw-semibold text-danger';

            els.masterStatus.innerHTML =
                `<i class="fa-solid fa-circle-exclamation me-1"></i>
                Master list not loaded yet. Please upload your master Excel file first.`;
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

                const firstSheet =
                    workbook.Sheets[workbook.SheetNames[0]];

                const json = XLSX.utils.sheet_to_json(firstSheet);

                state.records = json.map((row, idx) => {
                    let phone = String(
                        row['PHONE'] || row['Phone'] || ''
                    ).trim();

                    if (phone.length === 10 && !phone.startsWith('0')) {
                        phone = '0' + phone;
                    }

                    return {
                        sn: row['S/N'] || row['SN'] || idx + 1,

                        name: String(
                            row['NAME'] || row['Name'] || ''
                        ).trim().toUpperCase(),

                        phone: phone,

                        lg: String(
                            row['LG'] || row['Lg'] || row['LGA'] || ''
                        ).trim().toUpperCase(),

                        ward: String(
                            row['WARD'] || row['Ward'] || ''
                        ).trim().toUpperCase(),

                        position: String(
                            row['POSITION'] || row['Position'] || ''
                        ).trim().toUpperCase(),

                        qualification: String(
                            row['QUALIFICATION'] ||
                            row['Qualification'] || ''
                        ).trim().toUpperCase(),

                        applicationNo: String(
                            row['APPLICATION NO'] ||
                            row['Application No'] ||
                            row['APPLICATION_NO'] || ''
                        ).trim(),

                        internalReference: ''
                    };
                });

                state.exportSelection = null;

                saveLocalRecords();
                updateMasterStatusUI();
                updateStats();

                alert(
                    `Successfully loaded ${state.records.length} master records.\n\nUpload your documents to extract and match application numbers.`
                );

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

        const matched = state.records.filter(
            r => r.applicationNo && r.applicationNo.trim() !== ''
        ).length;

        const missing = total - matched;

        if (els.totalRecords) els.totalRecords.textContent = total;
        if (els.matchedCount) els.matchedCount.textContent = matched;
        if (els.missingCount) els.missingCount.textContent = missing;

        saveLocalRecords();
        renderTable();
    }

    async function processDocuments() {
        if (state.processing || state.files.length === 0) return;

        state.processing = true;
        state.exportSelection = null;
        state.unmatchedFiles = [];

        if (els.progressContainer) {
            els.progressContainer.style.display = 'block';
        }

        const filesToProcess = [...state.files];

        for (let i = 0; i < filesToProcess.length; i++) {
            const file = filesToProcess[i];

            const pct = Math.round(
                ((i + 1) / filesToProcess.length) * 100
            );

            if (els.progressBar) {
                els.progressBar.style.width = pct + '%';
            }

            if (els.progressText) {
                els.progressText.textContent =
                    `Scanning ${i + 1} of ${filesToProcess.length}: ${file.name}`;
            }

            try {
                let text;

                if (
                    file.type === 'application/pdf' ||
                    file.name.toLowerCase().endsWith('.pdf')
                ) {
                    text = await extractTextFromPDF(file);
                } else {
                    const result = await Tesseract.recognize(
                        file,
                        'eng'
                    );

                    text = result.data.text;
                }

                const extracted = parseFlexibleText(text);

                let foundMatch = false;

                if (extracted.applicationNo) {
                    const cleanText = text
                        .toUpperCase()
                        .replace(/\s+/g, ' ');

                    for (const r of state.records) {
                        const phoneTail = r.phone
                            ? r.phone.slice(-9)
                            : '';

                        const phoneMatch =
                            (extracted.phone &&
                                phoneTail &&
                                extracted.phone.includes(phoneTail)) ||
                            (phoneTail &&
                                cleanText.includes(phoneTail));

                        const nameTokens = r.name
                            .split(/\s+/)
                            .filter(t => t.length > 2);

                        const nameMatch = r.name && (
                            cleanText.includes(r.name) ||
                            (
                                nameTokens.length >= 2 &&
                                nameTokens.every(token =>
                                    cleanText.includes(token)
                                )
                            )
                        );

                        if (phoneMatch || nameMatch) {
                            r.applicationNo = extracted.applicationNo;
                            r.internalReference = '';

                            foundMatch = true;
                            break;
                        }
                    }
                }

                if (!foundMatch) {
                    state.unmatchedFiles.push(
                        `${file.name}${extracted.applicationNo
                            ? ' (' + extracted.applicationNo + ')'
                            : ''}`
                    );
                }

            } catch (err) {
                console.error(
                    `Error processing ${file.name}:`,
                    err
                );

                state.unmatchedFiles.push(
                    `${file.name} (Scan Error)`
                );
            }
        }

        state.files = [];
        state.processing = false;

        if (els.fileInput) {
            els.fileInput.value = '';
        }

        if (
            state.unmatchedFiles.length > 0 &&
            els.unmatchedAlert &&
            els.unmatchedList
        ) {
            els.unmatchedList.textContent =
                state.unmatchedFiles.join(', ');

            els.unmatchedAlert.style.display = 'block';
        } else if (els.unmatchedAlert) {
            els.unmatchedAlert.style.display = 'none';
        }

        if (els.progressText) {
            els.progressText.textContent = 'Matching complete!';
        }

        saveLocalRecords();
        updateStats();

        setTimeout(() => {
            if (els.progressContainer) {
                els.progressContainer.style.display = 'none';
            }
        }, 1500);
    }

    async function extractTextFromPDF(file) {
        const buffer = await file.arrayBuffer();

        const pdf = await pdfjsLib.getDocument({
            data: buffer
        }).promise;

        let fullText = '';

        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();

            if (content.items.length === 0) {
                const viewport = page.getViewport({
                    scale: 2
                });

                const canvas = document.createElement('canvas');

                canvas.width = viewport.width;
                canvas.height = viewport.height;

                const ctx = canvas.getContext('2d');

                await page.render({
                    canvasContext: ctx,
                    viewport
                }).promise;

                const result = await Tesseract.recognize(
                    canvas,
                    'eng'
                );

                fullText += result.data.text + '\n';
            } else {
                fullText += content.items
                    .map(item => item.str)
                    .join(' ') + '\n';
            }
        }

        return fullText;
    }

    function parseFlexibleText(text) {
        const cleanText = text.replace(/\s+/g, ' ');

        const appMatch = cleanText.match(/(ADA_[0-9]+)/i);

        const phoneMatch = cleanText.match(
            /\b(0?[789][0-9]{9})\b/
        );

        return {
            applicationNo: appMatch
                ? appMatch[1].toUpperCase()
                : '',

            phone: phoneMatch
                ? phoneMatch[1]
                : ''
        };
    }

    // Prepare exactly 310 records.
    function generateSelection() {
        if (state.processing) {
            alert('Please wait until document scanning is complete.');
            return;
        }

        if (state.records.length < EXPORT_LIMIT) {
            alert(
                `Your master list contains only ${state.records.length} records. At least 310 are required.`
            );
            return;
        }

        const verified = state.records.filter(
            r => r.applicationNo && r.applicationNo.trim() !== ''
        );

        const unmatched = state.records.filter(
            r => !r.applicationNo || !r.applicationNo.trim()
        );

        if (verified.length > EXPORT_LIMIT) {
            alert(
                `There are ${verified.length} records with extracted application numbers. This exceeds the 310-record export limit. No verified records will be removed automatically.`
            );
            return;
        }

        const needed = EXPORT_LIMIT - verified.length;

        if (unmatched.length < needed) {
            alert('Not enough unmatched records to complete the selection.');
            return;
        }

        const selectedUnmatched = shuffle(unmatched).slice(0, needed);

        const existingReferences = new Set(
            state.records
                .map(r => r.internalReference)
                .filter(Boolean)
        );

        selectedUnmatched.forEach(r => {
            if (!r.internalReference) {
                r.internalReference =
                    generateInternalReference(existingReferences);
            }
        });

        // Matched records remain first.
        // Selected unmatched records are randomized at the bottom.
        state.exportSelection = [
            ...verified,
            ...shuffle(selectedUnmatched)
        ];

        saveLocalRecords();

        const btn = document.getElementById('generate310Btn');

        if (btn) {
            btn.innerHTML =
                '<i class="fa-solid fa-check me-1"></i> 310 Records Prepared';
        }

        alert(
            `Selection prepared successfully!\n\nGenuine extracted IDs: ${verified.length}\nSelected unmatched records: ${needed}\nTotal: ${state.exportSelection.length}\n\nYou can now export your Excel file.`
        );
    }

    // Load ExcelJS dynamically for professional Excel formatting.
    function loadExcelJS() {
        return new Promise((resolve, reject) => {
            if (window.ExcelJS) {
                resolve(window.ExcelJS);
                return;
            }

            const script = document.createElement('script');

            script.src =
                'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';

            script.onload = () => {
                if (window.ExcelJS) {
                    resolve(window.ExcelJS);
                } else {
                    reject(new Error('ExcelJS failed to initialize.'));
                }
            };

            script.onerror = () => {
                reject(new Error('Unable to load ExcelJS. Check your internet connection.'));
            };

            document.head.appendChild(script);
        });
    }

    async function exportToExcel() {
        if (!state.exportSelection) {
            alert(
                'First click Prepare 310 Records, then click Export Excel.'
            );
            return;
        }

        if (state.exportSelection.length !== EXPORT_LIMIT) {
            alert('Export cancelled: the selection is not exactly 310 records.');
            return;
        }

        const btn = els.exportBtn;

        if (btn) {
            btn.disabled = true;
            btn.textContent = 'Preparing Excel...';
        }

        try {
            const ExcelJS = await loadExcelJS();

            const workbook = new ExcelJS.Workbook();

            workbook.creator = 'Gongola Tech Lab';
            workbook.subject = 'Enriched Master Records';
            workbook.title = 'Tangwami Enriched Records';
            workbook.created = new Date();

            const sheet = workbook.addWorksheet(
                'Enriched Records',
                {
                    views: [
                        {
                            state: 'frozen',
                            ySplit: 1,
                            xSplit: 0,
                            showGridLines: false
                        }
                    ]
                }
            );

            sheet.columns = [
                { header: 'S/N', key: 'sn', width: 8 },
                { header: 'NAME', key: 'name', width: 30 },
                { header: 'PHONE', key: 'phone', width: 17 },
                { header: 'LG', key: 'lg', width: 20 },
                { header: 'WARD', key: 'ward', width: 23 },
                { header: 'POSITION', key: 'position', width: 16 },
                { header: 'QUALIFICATION', key: 'qualification', width: 20 },
                { header: 'APPLICATION NO', key: 'applicationNo', width: 23 },
                {
                    header: 'INTERNAL REFERENCE (NOT OFFICIAL)',
                    key: 'internalReference',
                    width: 34
                },
                { header: 'ID STATUS', key: 'status', width: 19 }
            ];

            // Professional header.
            const header = sheet.getRow(1);

            header.height = 32;

            header.eachCell(cell => {
                cell.font = {
                    name: 'Arial',
                    bold: true,
                    color: { argb: 'FFFFFFFF' },
                    size: 10
                };

                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FF17365D' }
                };

                cell.alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                    wrapText: true
                };

                cell.border = {
                    bottom: {
                        style: 'medium',
                        color: { argb: 'FF4472C4' }
                    }
                };
            });

            state.exportSelection.forEach((r, index) => {
                const isVerified = Boolean(
                    r.applicationNo &&
                    r.applicationNo.trim()
                );

                const row = sheet.addRow({
                    sn: index + 1,
                    name: r.name || '',
                    phone: String(r.phone || ''),
                    lg: r.lg || '',
                    ward: r.ward || '',
                    position: r.position || '',
                    qualification: r.qualification || '',
                    applicationNo: isVerified
                        ? r.applicationNo
                        : '',
                    internalReference: isVerified
                        ? ''
                        : r.internalReference,
                    status: isVerified
                        ? 'EXTRACTED / VERIFIED'
                        : 'INTERNAL REFERENCE'
                });

                row.height = 22;

                row.eachCell(cell => {
                    cell.font = {
                        name: 'Arial',
                        size: 10,
                        color: { argb: 'FF222222' }
                    };

                    cell.alignment = {
                        vertical: 'middle',
                        horizontal: 'left'
                    };

                    cell.border = {
                        bottom: {
                            style: 'hair',
                            color: { argb: 'FFD9E2F3' }
                        }
                    };
                });

                // Alternating row backgrounds.
                if (index % 2 === 1) {
                    row.eachCell(cell => {
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFF2F6FA' }
                        };
                    });
                }

                // Genuine application numbers.
                if (isVerified) {
                    row.getCell('applicationNo').font = {
                        name: 'Arial',
                        bold: true,
                        color: { argb: 'FF008000' },
                        size: 10
                    };
                } else {
                    row.getCell('internalReference').font = {
                        name: 'Arial',
                        bold: true,
                        color: { argb: 'FFBF7000' },
                        size: 10
                    };

                    row.getCell('status').font = {
                        name: 'Arial',
                        bold: true,
                        color: { argb: 'FFBF7000' },
                        size: 9
                    };
                }

                row.getCell('sn').alignment = {
                    horizontal: 'center',
                    vertical: 'middle'
                };

                row.getCell('phone').numFmt = '@';
                row.getCell('applicationNo').numFmt = '@';
                row.getCell('internalReference').numFmt = '@';
            });

            // Enable Excel filters.
            sheet.autoFilter = {
                from: 'A1',
                to: 'J311'
            };

            // Page setup.
            sheet.pageSetup = {
                paperSize: 9,
                orientation: 'landscape',
                fitToPage: true,
                fitToWidth: 1,
                fitToHeight: 0,
                margins: {
                    left: 0.25,
                    right: 0.25,
                    top: 0.5,
                    bottom: 0.5,
                    header: 0.2,
                    footer: 0.2
                }
            };

            sheet.headerFooter.oddFooter =
                '&L Gongola Tech Lab &R Page &P of &N';

            // Confirm exact record count before download.
            if (sheet.rowCount !== 311) {
                throw new Error(
                    'Excel validation failed: expected 310 records plus one header.'
                );
            }

            const buffer = await workbook.xlsx.writeBuffer();

            const blob = new Blob(
                [buffer],
                {
                    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                }
            );

            const url = URL.createObjectURL(blob);

            const link = document.createElement('a');

            link.href = url;

            link.download =
                'Tangwami_310_Enriched_' +
                new Date().toISOString().slice(0, 10) +
                '.xlsx';

            document.body.appendChild(link);
            link.click();
            link.remove();

            setTimeout(() => URL.revokeObjectURL(url), 2000);

            alert(
                'Excel exported successfully!\n\n310 data records exported.\n144 or fewer verified IDs preserved.\nSelected unmatched records placed at the bottom.'
            );

        } catch (err) {
            console.error('Excel export error:', err);

            alert(
                'Excel export failed: ' + err.message
            );

        } finally {
            if (btn) {
                btn.disabled = false;
                btn.innerHTML =
                    '<i class="fa-solid fa-file-excel me-1"></i> Export Excel';
            }
        }
    }

    function renderTable() {
        if (!els.tableBody) return;

        els.tableBody.innerHTML = '';

        const filter = els.filterSelect
            ? els.filterSelect.value
            : 'all';

        const filtered = state.records
            .map((r, originalIdx) => ({
                ...r,
                originalIdx
            }))
            .filter(r => {
                const hasId = r.applicationNo &&
                    r.applicationNo.trim() !== '';

                if (filter === 'matched') return hasId;
                if (filter === 'missing') return !hasId;

                return true;
            });

        if (filtered.length === 0) {
            els.tableBody.innerHTML =
                '<tr><td colspan="8" class="text-center text-muted py-5">No records match the current view.</td></tr>';
            return;
        }

        filtered.forEach(r => {
            const hasId = Boolean(
                r.applicationNo &&
                r.applicationNo.trim()
            );

            const tr = document.createElement('tr');

            if (hasId) tr.className = 'row-matched';

            const values = [
                r.sn,
                r.name,
                r.phone,
                r.lg,
                r.ward,
                r.position,
                r.qualification,
                r.applicationNo
            ];

            tr.innerHTML = `
                <td class="fw-bold text-center">${escapeHtml(r.sn)}</td>
                ${['name','phone','lg','ward','position','qualification','applicationNo'].map((field, i) => `
                    <td>
                        <input
                            type="text"
                            value="${escapeHtml(values[i + 1])}"
                            data-index="${r.originalIdx}"
                            data-field="${field}"
                            class="table-input ${field === 'applicationNo' && hasId ? 'text-success fw-bold' : ''}"
                            ${field === 'applicationNo' && !hasId ? 'placeholder="Not extracted"' : ''}
                        >
                    </td>
                `).join('')}
            `;

            els.tableBody.appendChild(tr);
        });

        document.querySelectorAll('.table-input').forEach(input => {
            input.addEventListener('change', e => {
                const idx = Number(e.target.dataset.index);
                const field = e.target.dataset.field;

                state.records[idx][field] = e.target.value;

                if (field === 'applicationNo') {
                    if (e.target.value.trim()) {
                        state.records[idx].internalReference = '';
                    }
                }

                state.exportSelection = null;

                const btn = document.getElementById('generate310Btn');

                if (btn) {
                    btn.innerHTML =
                        '<i class="fa-solid fa-shuffle me-1"></i> Prepare 310 Records';
                }

                updateStats();
            });
        });
    }

    function clearAll() {
        if (!confirm(
            'Are you sure you want to reset everything and clear saved data?'
        )) return;

        state.files = [];
        state.records = [];
        state.unmatchedFiles = [];
        state.exportSelection = null;

        localStorage.removeItem(MASTER_KEY);

        if (els.fileInput) els.fileInput.value = '';
        if (els.masterFileInput) els.masterFileInput.value = '';

        if (els.unmatchedAlert) {
            els.unmatchedAlert.style.display = 'none';
        }

        updateMasterStatusUI();
        updateStats();

        const btn = document.getElementById('generate310Btn');

        if (btn) {
            btn.innerHTML =
                '<i class="fa-solid fa-shuffle me-1"></i> Prepare 310 Records';
        }
    }

    function escapeHtml(str) {
        if (str === null || str === undefined) return '';

        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }

    document.addEventListener('DOMContentLoaded', init);

})();
