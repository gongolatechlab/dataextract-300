
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

    // LOAD SAVED RECORDS
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
            localStorage.setItem(
                MASTER_KEY,
                JSON.stringify(state.records)
            );
        } catch (e) {
            alert('Unable to save records. Browser storage may be full.');
        }
    }

    // RANDOMIZE RECORDS
    function shuffle(array) {
        const result = [...array];

        for (let i = result.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));

            [result[i], result[j]] = [
                result[j],
                result[i]
            ];
        }

        return result;
    }

    // INITIALIZE APPLICATION
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
                alert(
                    'Please upload your Master List Excel file in Step 1 first.'
                );

                if (els.masterFileInput) {
                    els.masterFileInput.focus();
                }

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
                alert(
                    'Please upload your Master List Excel file first.'
                );
                return;
            }

            handleFiles(e.dataTransfer.files);
        });

        if (els.clearBtn) {
            els.clearBtn.addEventListener('click', clearAll);
        }

        if (els.exportBtn) {
            els.exportBtn.addEventListener('click', exportToExcel);
        }

        updateStats();
    }

    // CREATE RANDOM SELECTION BUTTON
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

    function resetSelectionButton() {
        state.exportSelection = null;

        const btn = document.getElementById('generate310Btn');

        if (btn) {
            btn.innerHTML =
                '<i class="fa-solid fa-shuffle me-1"></i> Prepare 310 Records';
        }
    }

    // MASTER STATUS
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

    // UPLOAD MASTER EXCEL
    function handleMasterUpload(e) {
        const file = e.target.files[0];

        if (!file) return;

        const reader = new FileReader();

        reader.onload = function(evt) {
            try {
                const data = new Uint8Array(evt.target.result);

                const workbook = XLSX.read(data, {
                    type: 'array'
                });

                const firstSheet =
                    workbook.Sheets[workbook.SheetNames[0]];

                const json = XLSX.utils.sheet_to_json(firstSheet);

                state.records = json.map((row, idx) => {
                    const sn =
                        row['S/N'] ||
                        row['SN'] ||
                        (idx + 1);

                    const name = String(
                        row['NAME'] ||
                        row['Name'] ||
                        ''
                    ).trim().toUpperCase();

                    let phone = String(
                        row['PHONE'] ||
                        row['Phone'] ||
                        ''
                    ).trim();

                    // Restore leading zero when Excel removed it.
                    if (
                        phone.length === 10 &&
                        !phone.startsWith('0')
                    ) {
                        phone = '0' + phone;
                    }

                    const lg = String(
                        row['LG'] ||
                        row['Lg'] ||
                        row['LGA'] ||
                        ''
                    ).trim().toUpperCase();

                    const ward = String(
                        row['WARD'] ||
                        row['Ward'] ||
                        ''
                    ).trim().toUpperCase();

                    const position = String(
                        row['POSITION'] ||
                        row['Position'] ||
                        ''
                    ).trim().toUpperCase();

                    const qualification = String(
                        row['QUALIFICATION'] ||
                        row['Qualification'] ||
                        ''
                    ).trim().toUpperCase();

                    // Master list has no application IDs.
                    return {
                        sn,
                        name,
                        phone,
                        lg,
                        ward,
                        position,
                        qualification,
                        applicationNo: ''
                    };
                });

                resetSelectionButton();

                saveLocalRecords();
                updateMasterStatusUI();
                updateStats();

                alert(
                    `Successfully loaded ${state.records.length} records from Master List!\n\nNow upload your PDFs or images in Step 2.`
                );

            } catch (err) {
                alert(
                    'Error reading Master Excel file: ' +
                    err.message
                );
            }
        };

        reader.readAsArrayBuffer(file);
    }

    // UPLOAD DOCUMENTS
    function handleFiles(fileList) {
        if (!fileList || fileList.length === 0) return;

        for (let i = 0; i < fileList.length; i++) {
            state.files.push(fileList[i]);
        }

        processDocuments();
    }

    // STATISTICS
    function updateStats() {
        const total = state.records.length;

        const matched = state.records.filter(
            r => r.applicationNo &&
                r.applicationNo.trim() !== ''
        ).length;

        const missing = total - matched;

        if (els.totalRecords) {
            els.totalRecords.textContent = total;
        }

        if (els.matchedCount) {
            els.matchedCount.textContent = matched;
        }

        if (els.missingCount) {
            els.missingCount.textContent = missing;
        }

        saveLocalRecords();
        renderTable();
    }

    // PROCESS PDF AND IMAGE DOCUMENTS
    async function processDocuments() {
        if (state.processing || state.files.length === 0) return;

        state.processing = true;

        resetSelectionButton();

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
                let text = '';

                // PDF DOCUMENT
                if (
                    file.type === 'application/pdf' ||
                    file.name.toLowerCase().endsWith('.pdf')
                ) {
                    if (els.progressText) {
                        els.progressText.textContent =
                            `Extracting PDF: ${file.name}`;
                    }

                    text = await extractTextFromPDF(file);

                } else {
                    // IMAGE DOCUMENT: DIRECT OCR
                    if (els.progressText) {
                        els.progressText.textContent =
                            `Scanning image directly: ${file.name}`;
                    }

                    const result = await Tesseract.recognize(
                        file,
                        'eng'
                    );

                    text = result.data.text;
                }

                const extracted = parseFlexibleText(text);

                let foundMatch = false;

                if (extracted.applicationNo) {
                    const cleanUpperText = text
                        .toUpperCase()
                        .replace(/\s+/g, ' ');

                    for (const r of state.records) {
                        const phoneTail = r.phone
                            ? r.phone.slice(-9)
                            : '';

                        const phoneMatch =
                            (
                                extracted.phone &&
                                phoneTail &&
                                extracted.phone.includes(phoneTail)
                            ) ||
                            (
                                phoneTail &&
                                cleanUpperText.includes(phoneTail)
                            );

                        const nameTokens = r.name
                            .split(/\s+/)
                            .filter(t => t.length > 2);

                        const nameMatch = r.name && (
                            cleanUpperText.includes(r.name) ||
                            (
                                nameTokens.length >= 2 &&
                                nameTokens.every(token =>
                                    cleanUpperText.includes(token)
                                )
                            )
                        );

                        if (phoneMatch || nameMatch) {
                            r.applicationNo =
                                extracted.applicationNo;

                            foundMatch = true;
                            break;
                        }
                    }
                }

                if (!foundMatch) {
                    state.unmatchedFiles.push(
                        `${file.name}${
                            extracted.applicationNo
                                ? ' (' + extracted.applicationNo + ')'
                                : ''
                        }`
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
            els.progressText.textContent =
                'Matching complete!';
        }

        saveLocalRecords();
        updateStats();

        setTimeout(() => {
            if (els.progressContainer) {
                els.progressContainer.style.display = 'none';
            }
        }, 1500);
    }

    // EXTRACT TEXT FROM PDF
    async function extractTextFromPDF(file) {
        const arrayBuffer = await file.arrayBuffer();

        const pdf = await pdfjsLib.getDocument({
            data: arrayBuffer
        }).promise;

        let fullText = '';

        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);

            const tokenized = await page.getTextContent();

            if (tokenized.items.length === 0) {
                const viewport = page.getViewport({
                    scale: 2.0
                });

                const canvas = document.createElement('canvas');

                canvas.width = viewport.width;
                canvas.height = viewport.height;

                const ctx = canvas.getContext('2d');

                await page.render({
                    canvasContext: ctx,
                    viewport: viewport
                }).promise;

                const result = await Tesseract.recognize(
                    canvas,
                    'eng'
                );

                fullText += result.data.text + '\n';

            } else {
                fullText += tokenized.items
                    .map(item => item.str)
                    .join(' ') + '\n';
            }
        }

        return fullText;
    }

    // EXTRACT APPLICATION NUMBER AND PHONE
    function parseFlexibleText(text) {
        const cleanText = text.replace(/\s+/g, ' ');

        let appNo = '';

        const appMatch = cleanText.match(
            /(ADA_[0-9]+)/i
        );

        if (appMatch) {
            appNo = appMatch[1].toUpperCase();
        }

        let phone = '';

        const phoneMatch = cleanText.match(
            /\b(0?[789][0-9]{9})\b/
        );

        if (phoneMatch) {
            phone = phoneMatch[1];
        }

        return {
            applicationNo: appNo,
            phone: phone
        };
    }

    // PREPARE EXACTLY 310 RECORDS
    function generateSelection() {
        if (state.processing) {
            alert(
                'Please wait until document scanning is complete.'
            );
            return;
        }

        if (state.records.length < EXPORT_LIMIT) {
            alert(
                `Your Master List contains only ${state.records.length} records. At least 310 are required.`
            );
            return;
        }

        const matched = state.records.filter(
            r => r.applicationNo &&
                r.applicationNo.trim() !== ''
        );

        const unmatched = state.records.filter(
            r => !r.applicationNo ||
                !r.applicationNo.trim()
        );

        // Never discard genuine extracted IDs.
        if (matched.length > EXPORT_LIMIT) {
            alert(
                `There are ${matched.length} genuine application IDs. This exceeds the 310-record limit. No genuine IDs will be discarded.`
            );
            return;
        }

        const required = EXPORT_LIMIT - matched.length;

        if (unmatched.length < required) {
            alert(
                'There are not enough unmatched records to complete the selection.'
            );
            return;
        }

        // Randomly select only the required unmatched records.
        const selectedUnmatched = shuffle(unmatched).slice(
            0,
            required
        );

        // Genuine matches first.
        // Randomized unmatched records at the bottom.
        state.exportSelection = [
            ...matched,
            ...selectedUnmatched
        ];

        if (state.exportSelection.length !== EXPORT_LIMIT) {
            state.exportSelection = null;

            alert('Selection validation failed.');
            return;
        }

        const btn = document.getElementById(
            'generate310Btn'
        );

        if (btn) {
            btn.innerHTML =
                '<i class="fa-solid fa-check me-1"></i> 310 Records Prepared';
        }

        alert(
            `310 RECORDS PREPARED SUCCESSFULLY!\n\n` +
            `Genuine extracted IDs: ${matched.length}\n` +
            `Unmatched records: ${required}\n` +
            `Total records: 310\n\n` +
            `Matched records are first.\n` +
            `Unmatched records are randomized at the bottom.\n\n` +
            `Missing application numbers will remain blank.`
        );
    }

    // LOAD EXCELJS FOR PROFESSIONAL FORMATTING
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
                    reject(
                        new Error('ExcelJS failed to initialize.')
                    );
                }
            };

            script.onerror = () => {
                reject(
                    new Error(
                        'Unable to load ExcelJS. Check your internet connection.'
                    )
                );
            };

            document.head.appendChild(script);
        });
    }

    // EXPORT EXACTLY 310 RECORDS AND EIGHT COLUMNS
    async function exportToExcel() {
        if (!state.exportSelection) {
            alert(
                'First click Prepare 310 Records, then click Export Excel.'
            );
            return;
        }

        if (
            state.exportSelection.length !== EXPORT_LIMIT
        ) {
            alert(
                'Export cancelled. Exactly 310 records are required.'
            );
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
                    views: [{
                        state: 'frozen',
                        ySplit: 1,
                        showGridLines: false
                    }]
                }
            );

            // EXACTLY THE ORIGINAL EIGHT COLUMNS
            sheet.columns = [
                {
                    header: 'S/N',
                    key: 'sn',
                    width: 8
                },
                {
                    header: 'NAME',
                    key: 'name',
                    width: 30
                },
                {
                    header: 'PHONE',
                    key: 'phone',
                    width: 18
                },
                {
                    header: 'LG',
                    key: 'lg',
                    width: 20
                },
                {
                    header: 'WARD',
                    key: 'ward',
                    width: 23
                },
                {
                    header: 'POSITION',
                    key: 'position',
                    width: 16
                },
                {
                    header: 'QUALIFICATION',
                    key: 'qualification',
                    width: 20
                },
                {
                    header: 'APPLICATION NO',
                    key: 'applicationNo',
                    width: 25
                }
            ];

            // HEADER FORMATTING
            const header = sheet.getRow(1);

            header.height = 32;

            header.eachCell(cell => {
                cell.font = {
                    name: 'Arial',
                    bold: true,
                    color: {
                        argb: 'FFFFFFFF'
                    },
                    size: 10
                };

                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: {
                        argb: 'FF17365D'
                    }
                };

                cell.alignment = {
                    horizontal: 'center',
                    vertical: 'middle',
                    wrapText: true
                };

                cell.border = {
                    bottom: {
                        style: 'medium',
                        color: {
                            argb: 'FF4472C4'
                        }
                    }
                };
            });

            // ADD 310 RECORDS
            state.exportSelection.forEach((r, index) => {
                const row = sheet.addRow({
                    sn: index + 1,
                    name: r.name || '',
                    phone: String(r.phone || ''),
                    lg: r.lg || '',
                    ward: r.ward || '',
                    position: r.position || '',
                    qualification: r.qualification || '',

                    // Preserve genuine IDs exactly.
                    // Do not fabricate missing IDs.
                    applicationNo: r.applicationNo || ''
                });

                row.height = 23;

                row.eachCell(cell => {
                    cell.font = {
                        name: 'Arial',
                        size: 10,
                        color: {
                            argb: 'FF222222'
                        }
                    };

                    cell.alignment = {
                        vertical: 'middle'
                    };

                    cell.border = {
                        bottom: {
                            style: 'hair',
                            color: {
                                argb: 'FFD9E2F3'
                            }
                        }
                    };

                    // Alternating row colours.
                    if (index % 2 === 1) {
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: {
                                argb: 'FFF2F6FA'
                            }
                        };
                    }
                });

                // Phone and application number as text.
                row.getCell(3).numFmt = '@';
                row.getCell(8).numFmt = '@';

                // Genuine application IDs in green.
                if (
                    r.applicationNo &&
                    r.applicationNo.trim()
                ) {
                    row.getCell(8).font = {
                        name: 'Arial',
                        bold: true,
                        size: 10,
                        color: {
                            argb: 'FF008000'
                        }
                    };
                }

                row.getCell(1).alignment = {
                    horizontal: 'center',
                    vertical: 'middle'
                };
            });

            // EXCEL FILTERS
            sheet.autoFilter = {
                from: 'A1',
                to: 'H311'
            };

            // PRINT SETTINGS
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

            // FINAL VALIDATION
            if (
                sheet.rowCount !== 311 ||
                sheet.columnCount !== 8
            ) {
                throw new Error(
                    'Excel validation failed. Expected 310 records and 8 columns.'
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

            setTimeout(() => {
                URL.revokeObjectURL(url);
            }, 2000);

            alert(
                'EXCEL EXPORTED SUCCESSFULLY!\n\n' +
                'Total records: 310\n' +
                'Columns: 8\n' +
                'Genuine IDs preserved.\n' +
                'Missing IDs left blank.\n' +
                'Matched records first.\n' +
                'Unmatched records randomized at the bottom.'
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

    // RENDER EDITABLE TABLE
    function renderTable() {
        if (!els.tableBody) return;

        els.tableBody.innerHTML = '';

        const filter = els.filterSelect
            ? els.filterSelect.value
            : 'all';

        const filteredRecords = state.records
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

        if (filteredRecords.length === 0) {
            els.tableBody.innerHTML =
                `<tr>
                    <td colspan="8"
                    class="text-center text-muted py-5">
                    No records match the current view.
                    </td>
                </tr>`;

            return;
        }

        filteredRecords.forEach(r => {
            const hasId = Boolean(
                r.applicationNo &&
                r.applicationNo.trim()
            );

            const tr = document.createElement('tr');

            if (hasId) {
                tr.className = 'row-matched';
            }

            const fields = [
                'name',
                'phone',
                'lg',
                'ward',
                'position',
                'qualification',
                'applicationNo'
            ];

            tr.innerHTML = `
                <td class="fw-bold text-center">
                    ${escapeHtml(r.sn)}
                </td>

                ${fields.map(field => `
                    <td>
                        <input
                            type="text"
                            value="${escapeHtml(r[field])}"
                            data-index="${r.originalIdx}"
                            data-field="${field}"
                            class="table-input ${
                                field === 'applicationNo' && hasId
                                    ? 'text-success fw-bold'
                                    : ''
                            }"
                            ${
                                field === 'applicationNo' && !hasId
                                    ? 'placeholder="Missing ID"'
                                    : ''
                            }
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

                resetSelectionButton();

                updateStats();
            });
        });
    }

    // RESET EVERYTHING
    function clearAll() {
        if (!confirm(
            'Are you sure you want to reset everything and clear saved data?'
        )) {
            return;
        }

        state.files = [];
        state.records = [];
        state.unmatchedFiles = [];
        state.exportSelection = null;

        localStorage.removeItem(MASTER_KEY);

        if (els.fileInput) {
            els.fileInput.value = '';
        }

        if (els.masterFileInput) {
            els.masterFileInput.value = '';
        }

        if (els.unmatchedAlert) {
            els.unmatchedAlert.style.display = 'none';
        }

        updateMasterStatusUI();
        updateStats();
        resetSelectionButton();
    }

    // HTML ESCAPING
    function escapeHtml(str) {
        if (str === null || str === undefined) {
            return '';
        }

        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }

    document.addEventListener(
        'DOMContentLoaded',
        init
    );

})();
