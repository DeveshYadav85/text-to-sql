/* global state */
const state = { history: [], sortCol: null, sortDir: 1 };

/* ── DOM refs ── */
const questionInput = document.getElementById('questionInput');
const sendBtn       = document.getElementById('sendBtn');
const resultsEl     = document.getElementById('results');
const emptyState    = document.getElementById('emptyState');
const historyList   = document.getElementById('historyList');
const schemaTree    = document.getElementById('schemaTree');
const modelName     = document.getElementById('modelName');

/* ── Boot ── */
window.addEventListener('DOMContentLoaded', () => {
  fetchHealth();
  fetchSchema();
  autoResize(questionInput);
  bindEvents();
});

/* ── Health / model name ── */
async function fetchHealth() {
  try {
    const res = await fetch('/health');
    const data = await res.json();
    const name = data.model.split('/').pop();
    modelName.textContent = name.length > 28 ? name.slice(0, 28) + '…' : name;
  } catch {
    modelName.textContent = 'Not connected';
  }
}

/* ── Schema sidebar ── */
async function fetchSchema() {
  try {
    const res  = await fetch('/schema');
    const data = await res.json();
    renderSchema(data.tables);
  } catch {
    schemaTree.innerHTML = '<span style="font-size:12px;color:var(--text-muted)">Failed to load schema</span>';
  }
}

function renderSchema(tables) {
  schemaTree.innerHTML = '';
  for (const [tbl, cols] of Object.entries(tables)) {
    const wrapper = document.createElement('div');
    wrapper.className = 'schema-table';

    const header = document.createElement('div');
    header.className = 'schema-table-header';
    header.innerHTML = `
      <span class="schema-table-icon">🗂</span>
      <span class="schema-table-name">${tbl}</span>
      <span class="schema-chevron">›</span>`;
    header.addEventListener('click', () => wrapper.classList.toggle('open'));

    const colsEl = document.createElement('div');
    colsEl.className = 'schema-cols';
    colsEl.innerHTML = cols.map(c => `<div class="schema-col">${c}</div>`).join('');

    wrapper.appendChild(header);
    wrapper.appendChild(colsEl);
    schemaTree.appendChild(wrapper);
  }
}

/* ── Event binding ── */
function bindEvents() {
  sendBtn.addEventListener('click', runQuery);

  questionInput.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
      e.preventDefault();
      runQuery();
    }
  });

  questionInput.addEventListener('input', () => autoResize(questionInput));

  document.querySelectorAll('.example-chip').forEach(btn => {
    btn.addEventListener('click', () => {
      questionInput.value = btn.dataset.q;
      autoResize(questionInput);
      runQuery();
    });
  });

  document.getElementById('sidebarToggle').addEventListener('click', () => {
    document.getElementById('sidebar').classList.toggle('collapsed');
  });
}

/* ── Auto-resize textarea ── */
function autoResize(el) {
  el.style.height = 'auto';
  el.style.height = Math.min(el.scrollHeight, 160) + 'px';
}

/* ── Run query ── */
async function runQuery() {
  const question = questionInput.value.trim();
  if (!question) return;

  setLoading(true);
  emptyState && (emptyState.style.display = 'none');
  showSkeleton();

  try {
    const res  = await fetch('/query', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ question })
    });

    if (!res.ok) {
      const err = await res.json();
      showError(question, err.detail || 'Server error');
      return;
    }

    const data = await res.json();
    addToHistory(question);
    renderResult(question, data);

  } catch (err) {
    showError(question, err.message || 'Network error');
  } finally {
    setLoading(false);
  }
}

/* ── Render result ── */
function renderResult(question, data) {
  resultsEl.innerHTML = '';

  const card = document.createElement('div');
  card.className = 'result-card';

  /* Header */
  card.innerHTML = `
    <div class="card-header">
      <span class="card-title">📝 "${truncate(question, 60)}"</span>
      <span class="attempts-badge">⟳ ${data.attempts} attempt${data.attempts !== 1 ? 's' : ''}</span>
    </div>`;

  /* SQL block */
  if (data.sql) {
    const sqlWrap = document.createElement('div');
    sqlWrap.className = 'card-body sql-block';
    const pre  = document.createElement('pre');
    const code = document.createElement('code');
    code.className = 'language-sql';
    code.textContent = data.sql;
    pre.appendChild(code);
    sqlWrap.appendChild(pre);
    card.appendChild(sqlWrap);
    hljs.highlightElement(code);
  }

  /* Error */
  if (data.error) {
    const errEl = document.createElement('div');
    errEl.className = 'error-block';
    errEl.innerHTML = `
      <span class="error-icon">⚠️</span>
      <div class="error-text">Failed after ${data.attempts} attempt(s): ${escHtml(data.error)}</div>`;
    card.appendChild(errEl);
  }

  /* Data table */
  if (data.columns && data.columns.length > 0) {
    const tableSection = buildTable(data.columns, data.rows);
    card.appendChild(tableSection);

    /* Row count */
    const rowCount = document.createElement('div');
    rowCount.className = 'row-count';
    rowCount.textContent = `${data.rows.length} row${data.rows.length !== 1 ? 's' : ''} returned`;
    card.appendChild(rowCount);
  }

  /* Explanation */
  if (data.explanation) {
    const expEl = document.createElement('div');
    expEl.className = 'explanation-block';
    expEl.innerHTML = `<span class="explanation-icon">💡</span><span>${escHtml(data.explanation)}</span>`;
    card.appendChild(expEl);
  }

  resultsEl.appendChild(card);
}

/* ── Build sortable table ── */
function buildTable(columns, rows) {
  let sortedRows = [...rows];
  let sortColIdx = null;
  let sortDir    = 1;

  const wrap = document.createElement('div');
  wrap.className = 'table-wrap';

  const table = document.createElement('table');
  table.className = 'data-table';

  /* Head */
  const thead = document.createElement('thead');
  const trh   = document.createElement('tr');
  columns.forEach((col, i) => {
    const th = document.createElement('th');
    th.innerHTML = `${escHtml(col)} <span class="sort-icon">↕</span>`;
    th.dataset.col = i;
    th.addEventListener('click', () => {
      if (sortColIdx === i) sortDir *= -1;
      else { sortColIdx = i; sortDir = 1; }
      sortedRows.sort((a, b) => {
        const va = a[i], vb = b[i];
        if (va === null) return 1; if (vb === null) return -1;
        return (va > vb ? 1 : va < vb ? -1 : 0) * sortDir;
      });
      thead.querySelectorAll('th').forEach(t => t.classList.remove('sorted'));
      th.classList.add('sorted');
      th.querySelector('.sort-icon').textContent = sortDir === 1 ? '↑' : '↓';
      renderBody();
    });
    trh.appendChild(th);
  });
  thead.appendChild(trh);
  table.appendChild(thead);

  /* Body */
  const tbody = document.createElement('tbody');
  table.appendChild(tbody);

  function renderBody() {
    tbody.innerHTML = '';
    sortedRows.forEach(row => {
      const tr = document.createElement('tr');
      row.forEach(cell => {
        const td = document.createElement('td');
        if (cell === null || cell === undefined)
          td.innerHTML = '<span class="null-val">NULL</span>';
        else
          td.textContent = cell;
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
  }

  renderBody();
  wrap.appendChild(table);
  return wrap;
}

/* ── Skeleton loader ── */
function showSkeleton() {
  resultsEl.innerHTML = `
    <div class="skeleton-card">
      <div class="skeleton-line"></div>
      <div class="skeleton-line short"></div>
      <div class="skeleton-line"></div>
      <div class="skeleton-line tiny"></div>
    </div>
    <div class="skeleton-card">
      <div class="skeleton-line short"></div>
      <div class="skeleton-line"></div>
    </div>`;
}

/* ── Error display ── */
function showError(question, msg) {
  resultsEl.innerHTML = '';
  const card = document.createElement('div');
  card.className = 'result-card';
  card.innerHTML = `
    <div class="card-header">
      <span class="card-title">📝 "${truncate(question, 60)}"</span>
    </div>
    <div class="error-block">
      <span class="error-icon">⚠️</span>
      <div class="error-text">${escHtml(msg)}</div>
    </div>`;
  resultsEl.appendChild(card);
}

/* ── Loading state ── */
function setLoading(on) {
  sendBtn.disabled = on;
  sendBtn.classList.toggle('loading', on);
  sendBtn.innerHTML = on
    ? `<span class="spinner"></span><span class="send-label">Running…</span>`
    : `<svg class="send-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
        <line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/>
       </svg><span class="send-label">Run</span>`;
}

/* ── History ── */
function addToHistory(q) {
  state.history.unshift(q);
  if (state.history.length > 10) state.history.pop();

  historyList.innerHTML = '';
  state.history.forEach(item => {
    const li = document.createElement('li');
    li.className = 'history-item';
    li.textContent = item;
    li.title = item;
    li.addEventListener('click', () => {
      questionInput.value = item;
      autoResize(questionInput);
      runQuery();
    });
    historyList.appendChild(li);
  });
}

/* ── Helpers ── */
function escHtml(s) {
  return String(s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function truncate(s, n) {
  return s.length > n ? s.slice(0, n) + '…' : s;
}
