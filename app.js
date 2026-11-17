/* DataExtract 300 - browser-only document extractor */
(function () {
  'use strict';

  if (window.pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc =
      'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
  }

  var state = { files: [], records: [], nextId: 1, busy: false };

  var $ = function (s) { return document.querySelector(s); };
  var els = {
    fileInput: $('#fileInput'),
    processBtn: $('#processBtn'),
    exportBtn: $('#exportBtn'),
    addBtn: $('#addBtn'),
    progressFill: $('#progressFill'),
    progressText: $('#progressText'),
    statTotal: $('#statTotal'),
    statDup: $('#statDup'),
    statMissing: $('#statMissing'),
    resultsBody: $('#resultsBody'),
    emptyState: $('#emptyState'),
  };

  var tessWorker = null;
  function getTessWorker() {
    if (!tessWorker) {
      tessWorker = Tesseract.createWorker('eng', 1, {
        workerPath: 'https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/worker.min.js',
        corePath: 'https://cdn.jsdelivr.net/npm/tesseract.js-core@5/',
        langPath: 'https://tessdata.projectnaptha.com/4.0.0'
      });
    }
    return tessWorker;
  }

  function processFile(file, onProgress) {
    var name = (file.name || '').toLowerCase();
    if (name.endsWith('.pdf')) return processPDF(file, onProgress);
    if (/\.(jpe?g|png)$/.test(name)) return processImage(file, onProgress);
    return Promise.resolve('');
  }

  function processPDF(file, onProgress) {
    return file.arrayBuffer().then(function (buf) {
      return pdfjsLib.getDocument({ data: buf }).promise;
    }).then(function (pdf) {
      var chain = Promise.resolve('');
      var pages = [];
      for (var i = 1; i <= pdf.numPages; i++) pages.push(i);
      pages.forEach(function (i) {
        chain = chain.then(function (acc) {
          onProgress && onProgress('Reading ' + file.name + ' - page ' + i + '/' + pdf.numPages);
          return pdf.getPage(i).then(function (page) {
            return page.getTextContent().then(function (content) {
              var pageText = content.items.map(function (it) { return it.str || ''; }).join(' ');
              var next = acc + '\n' + pageText;
              if (pageText.replace(/\s/g, '').length < 30) {
                onProgress && onProgress('OCR ' + file.name + ' - page ' + i + '/' + pdf.numPages);
                return ocrPdfPage(page).then(function (ocr) { return next + '\n' + ocr; });
              }
              return next;
            });
          });
        });
      });
      return chain;
    });
  }

  function ocrPdfPage(page) {
    var viewport = page.getViewport({ scale: 2.0 });
    var canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    var ctx = canvas.getContext('2d');
    return page.render({ canvasContext: ctx, viewport: viewport }).promise
      .then(function () { return canvas.toDataURL('image/png'); })
      .then(function (dataUrl) { return getTessWorker(); })
      .then(function (w) { return w.recognize(canvas.toDataURL('image/png')); })
      .then(function (res) { return (res && res.data && res.data.text) || ''; });
  }

  function processImage(file, onProgress) {
    onProgress && onProgress('OCR ' + file.name);
    return getTessWorker().then(function (w) { return w.recognize(file); })
      .then(function (res) { return (res && res.data && res.data.text) || ''; });
  }

  var NAME_LABEL = '(?:(?:full\\s+)?name|applicant(?:\\'s)?\\s+name|name\\s+of\\s+(?:the\\s+)?(?:applicant|candidate|student))';
  var NAME_LINE_RE = new RegExp('(?:^|\\n)\\s*(?:\\d+[.)]\\s*|[\\u2022\\-\\*]\\s*)?' + NAME_LABEL + '\\s*[:\\-]', 'gi');
  var APP_LABEL_RE = /\\b(?:application|app)\\s*(?:no\\.?|number|#)?\\s*[:\\-]/gi;

  function splitIntoRecords(text) {
    text = String(text || '').replace(/\\r/g, '').replace(/\\u00a0/g, ' ');
    var nameHits = Array.from(text.matchAll(NAME_LINE_RE));
    if (nameHits.length >= 1) {
      var blocks = [];
      for (var i = 0; i < nameHits.length; i++) {
        var start = i === 0 ? 0 : nameHits[i].index;
        var end = i + 1 < nameHits.length ? nameHits[i + 1].index : text.length;
        blocks.push(text.slice(start, end));
      }
      return blocks;
    }
    var appHits = Array.from(text.matchAll(APP_LABEL_RE));
    if (appHits.length >= 1) {
      var blocks2 = [];
      for (var j = 0; j < appHits.length; j++) {
        var s2 = j === 0 ? 0 : appHits[j].index;
        var e2 = j + 1 < appHits.length ? appHits[j + 1].index : text.length;
        blocks2.push(text.slice(s2, e2));
      }
      return blocks2;
    }
    return [text];
  }

  function pick(block, patterns) {
    for (var i = 0; i < patterns.length; i++) {
      var m = block.match(patterns[i]);
      if (m && m[1] !== undefined) {
        return m[1].trim().replace(/^[\\s:.,;\\-]+|[\\s:.,;\\-]+$/g, '');
      }
    }
    return '';
  }

  function extractRecord(block) {
    var name = pick(block, [
      new RegExp('(?:^|\\n)\\s*(?:\\d+[.)]\\s*|[\\u2022\\-\\*]\\s*)?' + NAME_LABEL + '\\s*[:\\-]\\s*([^\\n,;]+)', 'i')
    ]);
    var phone = pick(block, [
      /\\b(?:phone|tel(?:ephone)?|mobile|gsm)\\s*(?:no\\.?|number)?\\s*[:\\-]\\s*([+\\d][\\d\\s\\-]{7,20})/i,
      /\\b(0[789]\\d{9})\\b/,
      /\\b(\\+?234[\\s\\-]?\\d{3}[\\s\\-]?\\d{3}[\\s\\-]?\\d{4})\\b/
    ]).replace(/[^\\d+]/g, '');
    var lg = pick(block, [
      /\\b(?:L\\.?\\s?G\\.?A?\\.?|Local\\s+Government(?:\\s+Area)?)\\s*[:\\-]\\s*([^\\n,;]+)/i
    ]);
    var ward = pick(block, [/\\bward\\s*[:\\-]\\s*([^\\n,;]+)/i]);
    var position = pick(block, [
      /\\b(?:position|post|designation)\\s*[:\\-]\\s*([^\\n,;]+)/i
    ]);
    var institution = pick(block, [
      /\\b(?:name\\s+of\\s+(?:the\\s+)?institution|institution(?:\\s+name)?|school|college|university)\\s*[:\\-]\\s*([^\\n,;]+)/i
    ]);
    var applicationNo = pick(block, [
      /\\b(?:application|app)\\s*(?:no\\.?|number|#)?\\s*[:\\-]\\s*([A-Z0-9][A-Z0-9\\-\\/_.]+)/i
    ]).replace(/[.,;:]+$/, '');
    return { name: name, phone: phone, lg: lg, ward: ward, position: position, institution: institution, applicationNo: applicationNo };
  }

  function parseRecords(text) {
    var blocks = splitIntoRecords(text);
    var out = [];
    for (var i = 0; i < blocks.length; i++) {
      var r = extractRecord(blocks[i]);
      var hasAny = false;
      for (var k in r) { if (String(r[k] || '').trim()) { hasAny = true; break; } }
      if (hasAny) out.push(r);
    }
    return out;
  }

  function normName(s) { return String(s || '').toLowerCase().replace(/[^a-z]/g, ''); }
  function normPhone(s) {
    var d = String(s || '').replace(/\\D/g, '');
    if (d.indexOf('234') === 0) d = '0' + d.slice(3);
    return d;
  }
  function normApp(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }

  function detectDuplicates(records) {
    var byName = new Map(), byPhone = new Map(), byApp = new Map();
    function add(map, key, id) {
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(id);
    }
    for (var i = 0; i < records.length; i++) {
      var r = records[i];
      var n = normName(r.name);
      var words = String(r.name || '').trim().split(/\\s+/).filter(Boolean).length;
      if (n && words >= 2) add(byName, n, r.id);
      var p = normPhone(r.phone);
      if (p && p.length >= 10) add(byPhone, p, r.id);
      var a = normApp(r.applicationNo);
      if (a) add(byApp, a, r.id);
    }
    var dup = new Set();
    var maps = [byName, byPhone, byApp];
    for (var mi = 0; mi < maps.length; mi++) {
      var it = maps[mi].values();
      var v = it.next();
      while (!v.done) {
        if (v.value.length > 1) { v.value.forEach(function (id) { dup.add(id); }); }
        v = it.next();
      }
    }
    return dup;
  }

  var FIELDS = ['name', 'phone', 'lg', 'ward', 'position', 'institution', 'applicationNo'];
  var revalidateTimer = null;

  function renderTable() {
    var tbody = els.resultsBody;
    tbody.innerHTML = '';
    var dupIds = detectDuplicates(state.records);
    var frag = document.createDocumentFragment();
    state.records.forEach(function (r, idx) {
      var tr = document.createElement('tr');
      tr.dataset.id = String(r.id);
      if (dupIds.has(r.id)) tr.classList.add('dup');

      var snTd = document.createElement('td');
      snTd.className = 'sn';
      snTd.textContent = String(idx + 1);
      tr.appendChild(snTd);

      FIELDS.forEach(function (f) {
        var td = document.createElement('td');
        var inp = document.createElement('input');
        inp.type = 'text';
        inp.value = r[f] || '';
        inp.dataset.field = f;
        inp.dataset.id = String(r.id);
        if (!String(r[f] || '').trim()) inp.classList.add('missing');
        inp.addEventListener('input', onCellInput);
        td.appendChild(inp);
        tr.appendChild(td);
      });

      var actTd = document.createElement('td');
      actTd.className = 'actions';
      var del = document.createElement('button');
      del.type = 'button';
      del.className = 'del-btn';
      del.title = 'Delete record';
      del.textContent = '\\u2715';
      del.addEventListener('click', function () { deleteRecord(r.id); });
      actTd.appendChild(del);
      tr.appendChild(actTd);

      frag.appendChild(tr);
    });
    tbody.appendChild(frag);
    els.emptyState.style.display = state.records.length ? 'none' : 'block';
  }

  function onCellInput(e) {
    var inp = e.target;
    var id = Number(inp.dataset.id);
    var field = inp.dataset.field;
    var rec = null;
    for (var i = 0; i < state.records.length; i++) {
      if (state.records[i].id === id) { rec = state.records[i]; break; }
    }
    if (!rec) return;
    rec[field] = inp.value;
    inp.classList.toggle('missing', !inp.value.trim());
    clearTimeout(revalidateTimer);
    revalidateTimer = setTimeout(function () {
      var dupIds = detectDuplicates(state.records);
      var rows = els.resultsBody.querySelectorAll('tr');
      for (var j = 0; j < rows.length; j++) {
        rows[j].classList.toggle('dup', dupIds.has(Number(rows[j].dataset.id)));
      }
      updateStats();
    }, 250);
  }

  function deleteRecord(id) {
    state.records = state.records.filter(function (r) { return r.id !== id; });
    renderTable();
    updateStats();
  }

  function addRecord() {
    state.records.push({
      id: state.nextId++, name: '', phone: '', lg: '', ward: '',
      position: '', institution: '', applicationNo: ''
    });
    renderTable();
    updateStats();
  }

  function updateStats() {
    var total = state.records.length;
    var dup = detectDuplicates(state.records).size;
    var missing = state.records.filter(function (r) {
      return FIELDS.some(function (f) { return !String(r[f] || '').trim(); });
    }).length;
    els.statTotal.textContent = String(total);
    els.statDup.textContent = String(dup);
    els.statMissing.textContent = String(missing);
    els.exportBtn.disabled = total === 0;
  }

  function setProgress(pct, text) {
    els.progressFill.style.width = Math.max(0, Math.min(100, pct)) + '%';
    if (text) els.progressText.textContent = text;
  }

  els.fileInput.addEventListener('change', function () {
    state.files = Array.prototype.slice.call(els.fileInput.files || []);
    els.processBtn.disabled = state.files.length === 0 || state.busy;
    setProgress(0, state.files.length
      ? state.files.length + ' file(s) selected - tap Process'
      : 'No files selected');
  });

  els.processBtn.addEventListener('click', function () {
    if (state.busy || !state.files.length) return;
    state.busy = true;
    els.processBtn.disabled = true;
    els.exportBtn.disabled = true;

    var allText = '';
    var chain = Promise.resolve();
    state.files.forEach(function (f, i) {
      chain = chain.then(function () {
        var basePct = (i / state.files.length) * 90;
        setProgress(basePct, 'Processing ' + f.name + ' (' + (i + 1) + '/' + state.files.length + ')');
        return processFile(f, function (msg) { els.progressText.textContent = msg; })
          .then(function (text) {
            allText += '\\n' + text + '\\n';
            setProgress(((i + 1) / state.files.length) * 90, 'Processed ' + f.name);
          });
      });
    });

    chain.then(function () {
      setProgress(93, 'Parsing records...');
      var parsed = parseRecords(allText);
      parsed.forEach(function (p) {
        state.records.push({
          id: state.nextId++,
          name: p.name || '',
          phone: p.phone || '',
          lg: p.lg || '',
          ward: p.ward || '',
          position: p.position || '',
          institution: p.institution || '',
          applicationNo: p.applicationNo || ''
        });
      });
      renderTable();
      updateStats();
      setProgress(100, 'Done - ' + parsed.length + ' new record(s). Total: ' + state.records.length);
      if (!parsed.length) {
        els.progressText.textContent = 'No records detected. Check document labels.';
      }
    }).catch(function (err) {
      console.error(err);
      var msg = (err && err.message) ? err.message : String(err);
      setProgress(0, 'Error: ' + msg);
      alert('Processing failed: ' + msg);
    }).then(function () {
      state.busy = false;
      els.processBtn.disabled = state.files.length === 0;
      updateStats();
    });
  });

  els.addBtn.addEventListener('click', addRecord);

  els.exportBtn.addEventListener('click', function () {
    if (!state.records.length) return;
    var header = ['SN', 'Name', 'Phone', 'LG', 'Ward', 'Position', 'Name of institution', 'Application No'];
    var rows = [header];
    state.records.forEach(function (r, i) {
      rows.push([
        i + 1,
        r.name || '', r.phone || '', r.lg || '', r.ward || '',
        r.position || '', r.institution || '', r.applicationNo || ''
      ]);
    });
    var ws = XLSX.utils.aoa_to_sheet(rows);
    if (ws['!ref']) {
      var range = XLSX.utils.decode_range(ws['!ref']);
      for (var R = 1; R <= range.e.r; R++) {
        var addr = XLSX.utils.encode_cell({ r: R, c: 2 });
        if (!ws[addr]) continue;
        ws[addr].t = 's';
        ws[addr].v = String(ws[addr].v);
        ws[addr].z = '@';
      }
    }
    ws['!cols'] = [
      { wch: 5 }, { wch: 24 }, { wch: 16 }, { wch: 16 },
      { wch: 14 }, { wch: 16 }, { wch: 30 }, { wch: 20 }
    ];
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Records');
    var stamp = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, 'DataExtract_' + stamp + '.xlsx');
  });

  updateStats();
})();
