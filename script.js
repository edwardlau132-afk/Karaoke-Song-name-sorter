// Excel Searcher - Field-specific search functionality
(() => {
  const fileInput = document.getElementById('file-input');
  const info = document.getElementById('info');
  const uploadInfo = document.getElementById('upload-info');
  const resultsBody = document.getElementById('results-body');
  const searchFieldsContainer = document.getElementById('search-fields');
  const tableHead = document.getElementById('table-head');

  let allData = [];
  let headers = [];
  const inputs = {};

  function setInfo(t) { 
    info.textContent = t; 
    uploadInfo.textContent = t; 
  }

  function createSearchFields() {
    searchFieldsContainer.innerHTML = '';
    inputs.clear();
    
    if (headers.length === 0) {
      searchFieldsContainer.innerHTML = '<p style="color:#999;">Upload a file to create search fields</p>';
      return;
    }

    headers.forEach(field => {
      const label = document.createElement('label');
      const input = document.createElement('input');
      
      input.type = 'search';
      input.id = 'search-' + field;
      input.placeholder = `Search ${field}`;
      input.addEventListener('input', debounce(doSearch, 180));
      
      label.textContent = field;
      label.appendChild(input);
      searchFieldsContainer.appendChild(label);
      
      inputs[field] = input;
    });

    // Add clear button
    const clearBtn = document.createElement('button');
    clearBtn.textContent = 'Clear';
    clearBtn.style.height = '32px';
    clearBtn.addEventListener('click', () => {
      headers.forEach(f => inputs[f].value = '');
      doSearch();
    });
    searchFieldsContainer.appendChild(clearBtn);
  }

  function createTableHeaders() {
    tableHead.innerHTML = '';
    const tr = document.createElement('tr');
    
    headers.forEach(field => {
      const th = document.createElement('th');
      th.textContent = field;
      tr.appendChild(th);
    });
    
    tableHead.appendChild(tr);
  }

  function renderRows(rows) {
    resultsBody.innerHTML = '';
    if (!rows || rows.length === 0) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = headers.length || 1;
      td.textContent = 'No results';
      td.style.color = '#666';
      tr.appendChild(td);
      resultsBody.appendChild(tr);
      return;
    }

    const frag = document.createDocumentFragment();
    rows.forEach(r => {
      const tr = document.createElement('tr');
      headers.forEach(k => {
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
    headers.forEach(k => {
      const v = inputs[k].value.trim().toLowerCase();
      if (v) params[k] = v;
    });
    return params;
  }

  function filterRows(params) {
    if (allData.length === 0) return [];
    
    const paramKeys = Object.keys(params);
    if (paramKeys.length === 0) {
      return allData;
    }

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

  const doSearch = debounce(async () => {
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

  async function uploadFile(file) {
    setInfo(`Uploading ${file.name}...`);
    try {
      const reader = new FileReader();
      
      reader.onload = function(e) {
        try {
          let data = [];
          
          if (file.name.endsWith('.csv')) {
            // Handle CSV
            const text = e.target.result;
            const rows = text.split('\n').map(row => 
              row.split(',').map(cell => cell.trim())
            );
            headers = rows[0] || [];
            data = rows.slice(1).map(row => {
              const obj = {};
              headers.forEach((h, i) => obj[h] = row[i] || '');
              return obj;
            }).filter(row => headers.some((h, i) => row[h]));
          } else {
            // Handle Excel files
            if (typeof XLSX === 'undefined') {
              setInfo('XLSX library is not loaded. Please refresh the page and try again.');
              return;
            }
            const workbook = XLSX.read(e.target.result, { type: 'binary' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet);
            
            if (jsonData.length > 0) {
              headers = Object.keys(jsonData[0]);
              data = jsonData;
            }
          }
          
          if (data.length === 0) {
            setInfo('The file appears to be empty');
            return;
          }
          
          allData = data;
          createTableHeaders();
          createSearchFields();
          setInfo(`Loaded ${data.length} rows`);
          doSearch();
        } catch (error) {
          setInfo('Error reading file: ' + error.message);
        }
      };
      
      if (file.name.endsWith('.csv')) {
        reader.readAsText(file);
      } else {
        reader.readAsBinaryString(file);
      }
    } catch (err) {
      console.error(err);
      setInfo('Upload error');
    }
  }

  fileInput.addEventListener('change', (ev) => {
    const f = ev.target.files && ev.target.files[0];
    if (!f) return;
    uploadFile(f);
  });
})();
