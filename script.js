// Excel Searcher - fully client-side, field-specific search
// Reads .xls/.xlsx/.csv entirely in the browser (no backend/server involved).
(() => {
  // ── Auto-load config ──────────────────────────────────────────────
  // Every file in this list that actually exists in the repo root (next
  // to index.html) will be fetched and merged into one combined dataset
  // when the page opens — no manual upload needed. Add/remove entries
  // here as you commit new dated versions.
  const DEFAULT_DATA_FILES = [
    'karaoke-songs.xlsx'
  ];
  // ───────────────────────────────────────────────────────────────────

  const fileInput = document.getElementById('file-input');
  const info = document.getElementById('info');
  const uploadInfo = document.getElementById('upload-info');
  const resultsBody = document.getElementById('results-body');
  const clearBtn = document.getElementById('clear-filters');

  // Fixed column order: 1st column in the file -> code, 2nd -> chinese_title, etc.
  // Mapping is by POSITION, not by header text, since the real file's header
  // row may not literally say "code"/"singer"/etc.
  const fields = ['Number_編號', 'Song_Title_歌名', 'Singer_歌手', 'PinYin_国語拼音', 'Cantonese_粤語拼音', 'Word_Number_字數統計'];
  const inputs = {};
  fields.forEach(f => inputs[f] = document.getElementById('search-' + f));

  let allData = [];

  function setInfo(t) {
    info.textContent = t;
    uploadInfo.textContent = t;
  }

  function renderRows(rows) {
    resultsBody.innerHTML = '';
    if (!rows || rows.length === 0) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = fields.length;
      td.textContent = 'No results';
      td.style.color = '#666';
      tr.appendChild(td);
      resultsBody.appendChild(tr);
      return;
    }
    const frag = document.createDocumentFragment();
    rows.forEach(r => {
      const tr = document.createElement('tr');
      fields.forEach(k => {
        const td = document.createElement('td');
        td.textContent = r[k] ?? '';
        tr.appendChild(td);
      });
      frag.appendChild(tr);
    });
    resultsBody.appendChild(frag);
  }

  function collectParams() {
    const params = {};
    fields.forEach(k => {
      const v = inputs[k].value.trim().toLowerCase();
      if (v) params[k] = v;
    });
    return params;
  }

  function filterRows(params) {
    if (allData.length === 0) return [];
    const paramKeys = Object.keys(params);
    if (paramKeys.length === 0) return allData;

    return allData.filter(row => {
      return paramKeys.every(key => {
        const cellValue = String(row[key] ?? '').toLowerCase();
        return cellValue.includes(params[key]);
      });
    });
  }

  let timer;
  function debounce(fn, ms = 200) {
    return (...args) => {
      clearTimeout(timer);
      timer = setTimeout(() => fn(...args), ms);
    };
  }

  const doSearch = debounce(() => {
    const params = collectParams();
    const results = filterRows(params);
    const paramKeys = Object.keys(params);
    if (paramKeys.length > 0) {
      setInfo(`${results.length} / ${allData.length} matched`);
    } else {
      setInfo(`${allData.length} rows`);
    }
    renderRows(results);
  }, 180);

  // Turn an array-of-arrays (first row = header, rest = data) into
  // objects keyed by our fixed field names, using column POSITION.
  function rowsToObjects(rowsArr) {
    const dataRows = rowsArr.slice(1); // skip header row
    return dataRows
      .map(row => {
        const obj = {};
        fields.forEach((f, i) => { obj[f] = row[i] !== undefined && row[i] !== null ? String(row[i]).trim() : ''; });
        return obj;
      })
      .filter(obj => fields.some(f => obj[f])); // drop fully-empty rows
  }

  function parseCSV(text) {
    // Simple CSV split (handles plain comma-separated files; does not handle
    // quoted commas inside fields).
    const lines = text.split(/\r?\n/).filter(l => l.length > 0);
    const rowsArr = lines.map(line => line.split(',').map(cell => cell.trim()));
    return rowsToObjects(rowsArr);
  }

  function parseExcel(binaryData) {
    if (typeof XLSX === 'undefined') {
      throw new Error('XLSX library is not loaded. Please refresh the page and try again.');
    }
    const workbook = XLSX.read(binaryData, { type: 'binary' });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    // header: 1 -> array-of-arrays, so we control the field mapping ourselves
    // instead of relying on the file's actual header text.
    const rowsArr = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
    return rowsToObjects(rowsArr);
  }

  // Applies parsed rows to the app state and UI. Shared by manual upload
  // and auto-fetch-from-repo paths.
  function applyData(data, sourceLabel) {
    if (!data || data.length === 0) {
      setInfo('The file appears to be empty');
      return;
    }
    allData = data;
    fields.forEach(k => inputs[k].disabled = false);
    setInfo(`Loaded ${data.length} rows${sourceLabel ? ' from ' + sourceLabel : ''}`);
    doSearch();
  }

  function uploadFile(file) {
    setInfo(`Reading ${file.name}...`);
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        let data = [];
        if (file.name.toLowerCase().endsWith('.csv')) {
          data = parseCSV(e.target.result);
        } else {
          data = parseExcel(e.target.result);
        }
        applyData(data, file.name);
      } catch (error) {
        setInfo('Error reading file: ' + error.message);
      }
    };

    reader.onerror = () => setInfo('Error reading file.');

    if (file.name.toLowerCase().endsWith('.csv')) {
      reader.readAsText(file);
    } else {
      reader.readAsBinaryString(file);
    }
  }

  // Fetches a single file (from the repo) and returns its parsed rows,
  // or null if the file doesn't exist / fails to load.
  async function fetchAndParse(name) {
    try {
      const res = await fetch(name, { cache: 'no-store' });
      if (!res.ok) return null;

      if (name.toLowerCase().endsWith('.csv')) {
        const text = await res.text();
        return parseCSV(text);
      } else {
        if (typeof XLSX === 'undefined') {
          setInfo('XLSX library is not loaded. Please refresh the page and try again.');
          return null;
        }
        const buf = await res.arrayBuffer();
        const workbook = XLSX.read(buf, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const rowsArr = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
        return rowsToObjects(rowsArr);
      }
    } catch (err) {
      console.warn(`Could not auto-load "${name}":`, err);
      return null;
    }
  }

  // Fetches every file in DEFAULT_DATA_FILES that actually exists in the
  // repo and merges all of their rows into one combined dataset. Files
  // that aren't present are skipped silently. If none are found, the
  // page stays in "No file uploaded" state and manual upload still works.
  async function tryLoadDefaultFile() {
    setInfo('Loading spreadsheets...');
    const results = await Promise.all(DEFAULT_DATA_FILES.map(fetchAndParse));

    let combined = [];
    let loadedNames = [];
    results.forEach((data, i) => {
      if (data && data.length > 0) {
        combined = combined.concat(data);
        loadedNames.push(DEFAULT_DATA_FILES[i]);
      }
    });

    if (combined.length === 0) {
      setInfo('No file uploaded.');
      return;
    }

    applyData(combined, `${loadedNames.length} file${loadedNames.length > 1 ? 's' : ''}`);
  }

  // Start with inputs disabled until a file is loaded.
  fields.forEach(k => inputs[k].disabled = true);

  fields.forEach(k => inputs[k].addEventListener('input', () => doSearch()));
  clearBtn.addEventListener('click', () => {
    fields.forEach(k => inputs[k].value = '');
    doSearch();
  });

  fileInput.addEventListener('change', (ev) => {
    const f = ev.target.files && ev.target.files[0];
    if (!f) return;
    uploadFile(f);
  });

  // Attempt to auto-load a spreadsheet already committed to the repo.
  tryLoadDefaultFile();
})();
