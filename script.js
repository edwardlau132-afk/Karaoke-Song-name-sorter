// Excel Searcher - fully client-side, field-specific search
// Reads .xls/.xlsx/.csv entirely in the browser (no backend/server involved).
(() => {
  // ── Auto-load config ──────────────────────────────────────────────
  // Every file in this list that actually exists in the repo root (next
  // to index.html) will be fetched and merged into one combined dataset
  // when the page opens — no manual upload needed. Add/remove entries
  // here as you commit new dated versions.
  const DEFAULT_DATA_FILES = [
    'Karaoke Songs Search List.xlsx'
  ];
  // ───────────────────────────────────────────────────────────────────
  const resultsBody = document.getElementById('results-body');

  // Fixed column order: 1st column in the file -> code, 2nd -> chinese_title, etc.
  // Mapping is by POSITION, not by header text, since the real file's header
  // row may not literally say "code"/"singer"/etc.
  const fields = ['Number_編號', 'Song_Title_歌名', 'Singer_歌手', 'PinYin_国語拼音', 'Cantonese_粤語拼音', 'Word_Number_字數統計'];
  const inputs = {};
fields.forEach(f => inputs[f] = document.getElementById('search-' + f));fields.forEach(f => {
  inputs[f] = document.getElementById('search-' + f);
});

console.log(inputs);
  
  let allData = [];

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
  renderRows(results);
}, 180);


fields.forEach(f => {
  inputs[f].addEventListener('input', doSearch);
});

const clearBtn = document.getElementById('clear-filters');

if (clearBtn) {
  clearBtn.addEventListener('click', () => {
    fields.forEach(f => {
      inputs[f].value = '';
    });

    renderRows(allData);
  });
}
  
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

function applyData(data) {
  if (!data || data.length === 0) {
    return;
  }

  allData = data;

  fields.forEach(k => {
    if (inputs[k]) {
      inputs[k].disabled = false;
    }
  });

  renderRows(allData);
}
  
  // Fetches a single file (from the repo) and returns its parsed rows,
  // or null if the file doesn't exist / fails to load.
  async function fetchAndParse(name) {
  try {
    console.log("Trying to load:", name);

    const res = await fetch(name, {
      cache: "no-store"
    });

    console.log("Status:", res.status);
    console.log("URL:", res.url);

    if (!res.ok) {
      throw new Error(`HTTP ${res.status}`);
    }

    if (name.toLowerCase().endsWith(".csv")) {
      const text = await res.text();
      return parseCSV(text);
    }

    const buf = await res.arrayBuffer();

    console.log("Downloaded", buf.byteLength, "bytes");

    const workbook = XLSX.read(buf, {
      type: "array"
    });

    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

    return rowsToObjects(
      XLSX.utils.sheet_to_json(firstSheet, {
        header: 1,
        defval: ""
      })
    );

  } catch (err) {
    console.error(err);
    return null;
  }
}
 
// ── Auto-load files when page opens ──────────────────────────────
async function autoLoadFiles() {

  let combinedData = [];

  for (const file of DEFAULT_DATA_FILES) {
    const data = await fetchAndParse(file);

    if (data) {
      combinedData = combinedData.concat(data);
    }
  }

  if (combinedData.length > 0) {
    applyData(combinedData);
  }
}

// Start automatic loading
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", autoLoadFiles);
} else {
  autoLoadFiles();
}

})();
