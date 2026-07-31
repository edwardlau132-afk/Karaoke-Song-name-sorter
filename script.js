// Karaoke Searcher - queries the live MySQL database via search.php
// Supports switching between two datasets ("Table 1" / "Table 2"), one
// shown at a time, via the toggle buttons.
(() => {
  const resultsBody = document.getElementById('results-body');

  const fields = ['Number_編號', 'Song_Title_歌名', 'Singer_歌手', 'PinYin_国語拼音', 'Cantonese_粤語拼音', 'Word_Number_字數統計'];
  const inputs = {};

  fields.forEach(f => {
    inputs[f] = document.getElementById('search-' + f);
    if (inputs[f]) inputs[f].disabled = false;
  });

  // Which dataset is currently active: 'old' or 'new'. Must match the
  // keys used in search.php's $tableMap.
  let currentSource = 'old';

  const sourceButtons = document.querySelectorAll('.source-btn');
  sourceButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      const newSource = btn.dataset.source;
      if (newSource === currentSource) return;

      currentSource = newSource;
      sourceButtons.forEach(b => b.classList.toggle('active', b.dataset.source === currentSource));
      runSearch();
    });
  });

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

    if (rows.length >= 300) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = fields.length;
      td.style.color = '#666';
      td.style.fontStyle = 'italic';
      td.textContent = '';
      tr.appendChild(td);
      resultsBody.appendChild(tr);
    }
  }

  function collectParams() {
    const params = new URLSearchParams();
    fields.forEach(k => {
      const v = inputs[k] ? inputs[k].value.trim() : '';
      if (v) params.set(k, v);
    });
    params.set('source', currentSource);
    return params;
  }

  async function runSearch() {
    const params = collectParams();
    resultsBody.innerHTML = '';
    const tr = document.createElement('tr');
    const td = document.createElement('td');
    td.colSpan = fields.length;
    td.style.color = '#666';
    td.textContent = 'Searching...';
    tr.appendChild(td);
    resultsBody.appendChild(tr);

    try {
      const res = await fetch('search.php?' + params.toString(), { cache: 'no-store' });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const rows = await res.json();
      if (rows.error) throw new Error(rows.error);
      renderRows(rows);
    } catch (err) {
      console.error('Search failed:', err);
      resultsBody.innerHTML = '';
      const tr2 = document.createElement('tr');
      const td2 = document.createElement('td');
      td2.colSpan = fields.length;
      td2.style.color = '#b00020';
      td2.textContent = 'Could not reach the search server. Open the browser console (F12) for details.';
      tr2.appendChild(td2);
      resultsBody.appendChild(tr2);
    }
  }

  let timer;
  function debounce(fn, ms = 250) {
    return (...args) => {
      clearTimeout(timer);
      timer = setTimeout(() => fn(...args), ms);
    };
  }

  const doSearch = debounce(runSearch, 250);

  fields.forEach(f => {
    if (inputs[f]) {
      inputs[f].addEventListener('input', doSearch);
    }
  });

  const clearBtn = document.getElementById('clear-filters');
  if (clearBtn) {
    clearBtn.addEventListener('click', () => {
      fields.forEach(f => { if (inputs[f]) inputs[f].value = ''; });
      runSearch();
    });
  }

  // Load the full (first 300) list on page open
  runSearch();
})();
