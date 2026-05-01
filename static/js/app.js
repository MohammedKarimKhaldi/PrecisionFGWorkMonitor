/* PrecisionFG Fundraising Monitor */

// ── API helpers ────────────────────────────────────────────────────────────
const API = {
  async get(url) {
    const r = await fetch(url);
    if (!r.ok) { const t = await r.text(); throw new Error(t); }
    return r.json();
  },
  async post(url, body) {
    const r = await fetch(url, { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(body) });
    if (!r.ok) { const t = await r.text(); throw new Error(t); }
    return r.json();
  },
  async put(url, body) {
    const r = await fetch(url, { method:'PUT', headers:{'Content-Type':'application/json'}, body:JSON.stringify(body) });
    if (!r.ok) { const t = await r.text(); throw new Error(t); }
    return r.json();
  },
};

// ── State ──────────────────────────────────────────────────────────────────
const state = {
  companies:     [],
  emails:        [],
  emailsGrouped: [],
  emailsTotal:   0,
  statuses:      [],
  selectedCompany: null,
  currentEmails: [],
};

// ── Toast ──────────────────────────────────────────────────────────────────
function toast(msg, type = '') {
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.textContent = msg;
  document.getElementById('toast-container').appendChild(el);
  setTimeout(() => el.remove(), 3500);
}

// ── Init ───────────────────────────────────────────────────────────────────
async function init() {
  const [statusData, apiStatus] = await Promise.all([
    fetch('/api/statuses').then(r => r.json()),
    fetch('/api/status').then(r => r.json()),
  ]);

  state.statuses = statusData.statuses || [];
  populateFilterDropdown();

  document.getElementById('user-info').textContent = apiStatus.email || '';
  document.getElementById('settings-email').textContent = apiStatus.email || '—';

  await Promise.all([loadCompanies(), loadEmails()]);
  renderAll();
}

// ── Data loaders ───────────────────────────────────────────────────────────
async function loadCompanies() {
  try {
    const data = await API.get('/api/companies');
    state.companies = data.companies || [];
  } catch (e) {
    toast('Could not load companies: ' + e.message, 'error');
  }
}

async function loadEmails() {
  setConnBadge('unknown', '● fetching emails…');
  try {
    const data = await API.get('/api/emails');
    state.emails        = data.messages  || [];
    state.emailsGrouped = data.grouped   || [];
    state.emailsTotal   = data.total     || state.emails.length;
    setConnBadge('ok', '● connected');
  } catch (e) {
    setConnBadge('error', '● IMAP error');
    toast('Email load failed: ' + e.message, 'error');
  }
}

function setConnBadge(type, text) {
  const b = document.getElementById('conn-badge');
  b.className = `conn-badge conn-${type}`;
  b.textContent = text;
}

// ── Render ─────────────────────────────────────────────────────────────────
function renderAll() {
  renderStats();
  renderSidebar();
  renderCompaniesTable();
}

function renderStats() {
  const counts = {};
  state.statuses.forEach(s => counts[s] = 0);
  state.companies.forEach(c => {
    const s = c['Status'] || '';
    if (s) counts[s] = (counts[s] || 0) + 1;
  });

  const highlights = [
    { label: 'Total',     val: state.companies.length,
      color: '#2E75B6' },
    { label: 'Active',    val: (counts['In Discussion']||0)+(counts['Follow Up']||0)+(counts['Due Diligence']||0),
      color: '#1565C0' },
    { label: 'Proposals', val: counts['Proposal Sent'] || 0,
      color: '#2e7d32' },
    { label: 'Mandates',  val: (counts['Mandate Received']||0)+(counts['Closed – Won']||0),
      color: '#1B5E20' },
    { label: 'Lost',      val: (counts['Not Interested']||0)+(counts['Closed – Lost']||0),
      color: '#c62828' },
    { label: 'Emails',    val: state.emailsTotal,
      color: '#6a1b9a' },
  ];

  document.getElementById('stats-bar').innerHTML = highlights.map(h => `
    <div class="stat-card">
      <div class="stat-val" style="color:${h.color}">${h.val}</div>
      <div class="stat-label">${h.label}</div>
    </div>`).join('');
}

function renderSidebar() {
  const list = document.getElementById('sidebar-list');
  const combined = mergeCompaniesWithEmails();
  if (combined.length === 0) {
    list.innerHTML = '<div class="empty-state" style="padding:20px;font-size:12px">No companies yet.<br>Add one or refresh emails.</div>';
    return;
  }
  list.innerHTML = combined.map(c => `
    <div class="sidebar-item ${state.selectedCompany === c.name ? 'active' : ''}"
         onclick="selectCompany(${JSON.stringify(c.name)})">
      <div class="company-name">${esc(c.name)}</div>
      <div class="company-meta">
        ${c.status ? `<span class="badge badge-${esc(c.status)}">${esc(c.status)}</span>` : ''}
        ${c.emailCount ? ` · ${c.emailCount} emails` : ''}
        ${c.lastDate ? ` · ${c.lastDate.slice(0,10)}` : ''}
      </div>
    </div>`).join('');
}

function mergeCompaniesWithEmails() {
  const result = state.companies.map(c => ({
    name:       c['Company'] || '',
    status:     c['Status'] || '',
    lastDate:   c['Last Email Date'] || '',
    emailCount: null,
    source:     'excel',
  }));

  const companyNames = new Set(state.companies.map(c => (c['Company'] || '').toLowerCase()));

  state.emailsGrouped.forEach(g => {
    const domainBase = g.domain.split('.')[0].toLowerCase();
    const alreadyIn  = [...companyNames].some(n => n.includes(domainBase) || domainBase.includes(n.split(' ')[0]));
    if (!alreadyIn) {
      result.push({
        name:       g.domain,
        status:     '',
        lastDate:   g.last_email_date,
        emailCount: g.message_count,
        source:     'email',
      });
    }
  });

  return result.sort((a, b) => (b.lastDate || '').localeCompare(a.lastDate || ''));
}

function renderCompaniesTable() {
  const tbody    = document.getElementById('companies-tbody');
  const filterSt = document.getElementById('filter-status').value;
  const searchQ  = document.getElementById('search-input').value.toLowerCase();

  const filtered = state.companies.filter(c => {
    if (filterSt && c['Status'] !== filterSt) return false;
    if (searchQ) {
      const haystack = ((c['Company']||'')+(c['Contact Name']||'')+(c['Contact Email']||'')).toLowerCase();
      if (!haystack.includes(searchQ)) return false;
    }
    return true;
  });

  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="7" class="empty-state">No companies match.</td></tr>';
    return;
  }

  tbody.innerHTML = filtered.map(c => `
    <tr>
      <td><strong>${esc(c['Company']||'')}</strong></td>
      <td>${esc(c['Contact Name']||'')}${c['Contact Email'] ? `<br><small style="color:#888">${esc(c['Contact Email'])}</small>` : ''}</td>
      <td><span class="badge badge-${esc(c['Status']||'')}">${esc(c['Status']||'—')}</span></td>
      <td>${esc(c['Last Email Date']||'—')}</td>
      <td>${esc(c['Mandate Type']||'—')}</td>
      <td>${c['AUM (M€)'] ? esc(c['AUM (M€)'])+'M€' : '—'}</td>
      <td style="white-space:nowrap;display:flex;gap:4px">
        <button class="btn btn-outline btn-sm" onclick="openEditModal(${JSON.stringify(c['Company']||'')})">Edit</button>
        <button class="btn btn-primary btn-sm" onclick="openStatusModal(${JSON.stringify(c['Company']||'')},${JSON.stringify(c['Status']||'')})">Status</button>
      </td>
    </tr>`).join('');
}

// ── Company detail ─────────────────────────────────────────────────────────
function selectCompany(name) {
  state.selectedCompany = name;
  renderSidebar();

  const panel = document.getElementById('company-detail');
  panel.style.display = 'block';
  document.getElementById('detail-title').textContent = name;

  const company = state.companies.find(c => c['Company'] === name);
  if (company) {
    document.getElementById('detail-status').innerHTML =
      `<span class="badge badge-${esc(company['Status']||'')}">${esc(company['Status']||'—')}</span>`;
    document.getElementById('detail-contact').textContent =
      company['Contact Name'] ? `${company['Contact Name']} <${company['Contact Email']||''}>` : '—';
    document.getElementById('detail-mandate').textContent = company['Mandate Type'] || '—';
    document.getElementById('detail-aum').textContent     = company['AUM (M€)'] ? company['AUM (M€)'] + ' M€' : '—';
    document.getElementById('detail-notes').textContent   = company['Notes'] || '—';
    document.getElementById('detail-first').textContent   = company['First Contact Date'] || '—';
    document.getElementById('detail-last').textContent    = company['Last Email Date'] || '—';
  }

  document.getElementById('detail-edit-btn').onclick   = () => openEditModal(name);
  document.getElementById('detail-status-btn').onclick = () => openStatusModal(name, company?.['Status'] || '');
  document.getElementById('detail-log-btn').onclick    = () => logEmailsForCompany(name);

  panel.scrollIntoView({ behavior: 'smooth', block: 'start' });
  loadCompanyEmails(name, company);
}

async function loadCompanyEmails(name, company) {
  const panel = document.getElementById('emails-panel');
  panel.innerHTML = '<div style="padding:16px;text-align:center"><span class="spinner"></span></div>';

  const searchTerm = company?.['Contact Email']
    ? company['Contact Email'].split('@')[1] || name
    : name;

  try {
    const data = await API.get(`/api/emails/search?q=${encodeURIComponent(searchTerm)}`);
    state.currentEmails = data.messages || [];
    renderEmailsPanel(state.currentEmails);
  } catch (e) {
    panel.innerHTML = `<div class="empty-state">Could not load emails: ${esc(e.message)}</div>`;
  }
}

function renderEmailsPanel(emails) {
  const panel = document.getElementById('emails-panel');
  if (!emails.length) {
    panel.innerHTML = '<div class="empty-state">No emails found matching this company.</div>';
    return;
  }
  panel.innerHTML = emails.map(e => {
    const ea      = (e.from?.emailAddress) || {};
    const date    = (e.receivedDateTime || '').replace('T',' ').slice(0,16);
    const isOwn   = (ea.address||'').toLowerCase().includes(
      (document.getElementById('user-info').textContent.split('@')[0]||'').toLowerCase()
    );
    return `
      <div class="email-item">
        <div class="email-subject">
          <span class="${isOwn ? 'dir-out' : 'dir-in'}">${isOwn ? 'OUT' : 'IN'}</span>
          ${esc(e.subject || '(no subject)')}
        </div>
        <div class="email-meta">
          ${esc(ea.name||'')} &lt;${esc(ea.address||'')}&gt; · ${esc(date)}
        </div>
        ${e.bodyPreview ? `<div class="email-preview">${esc((e.bodyPreview||'').slice(0,200))}${(e.bodyPreview||'').length>200?'…':''}</div>` : ''}
      </div>`;
  }).join('');
}

async function logEmailsForCompany(name) {
  if (!state.currentEmails.length) {
    toast('No emails loaded for this company', 'error');
    return;
  }
  try {
    const data = await API.post(`/api/companies/${encodeURIComponent(name)}/log-emails`,
      { emails: state.currentEmails });
    toast(`${data.added} new email(s) logged to Excel`, 'success');
  } catch (e) {
    toast('Log failed: ' + e.message, 'error');
  }
}

// ── Modals ─────────────────────────────────────────────────────────────────
function openAddModal() {
  document.getElementById('modal-add').style.display = 'flex';
  document.getElementById('form-add').reset();
  const sel = document.getElementById('add-status');
  sel.innerHTML = state.statuses.map(s => `<option>${esc(s)}</option>`).join('');
}
function closeAddModal() { document.getElementById('modal-add').style.display = 'none'; }

async function submitAddCompany() {
  const f = document.getElementById('form-add');
  const data = {
    Company:         f.elements['company'].value.trim(),
    'Contact Name':  f.elements['contact_name'].value.trim(),
    'Contact Email': f.elements['contact_email'].value.trim(),
    Status:          f.elements['status'].value,
    'Mandate Type':  f.elements['mandate_type'].value.trim(),
    'AUM (M€)':      f.elements['aum'].value.trim(),
    Notes:           f.elements['notes'].value.trim(),
  };
  if (!data.Company) { toast('Company name is required', 'error'); return; }
  try {
    await API.post('/api/companies', data);
    toast('Company saved to Excel', 'success');
    closeAddModal();
    await loadCompanies();
    renderAll();
  } catch (e) { toast('Error: ' + e.message, 'error'); }
}

function openEditModal(name) {
  const c = state.companies.find(c => c['Company'] === name);
  if (!c) return;
  const modal = document.getElementById('modal-edit');
  modal.style.display = 'flex';
  const f = document.getElementById('form-edit');
  f.elements['company'].value       = c['Company'] || '';
  f.elements['contact_name'].value  = c['Contact Name'] || '';
  f.elements['contact_email'].value = c['Contact Email'] || '';
  f.elements['mandate_type'].value  = c['Mandate Type'] || '';
  f.elements['aum'].value           = c['AUM (M€)'] || '';
  f.elements['notes'].value         = c['Notes'] || '';
  f.elements['status'].innerHTML    = state.statuses.map(s =>
    `<option ${s===c['Status']?'selected':''}>${esc(s)}</option>`
  ).join('');
}
function closeEditModal() { document.getElementById('modal-edit').style.display = 'none'; }

async function submitEditCompany() {
  const f = document.getElementById('form-edit');
  const data = {
    Company:         f.elements['company'].value.trim(),
    'Contact Name':  f.elements['contact_name'].value.trim(),
    'Contact Email': f.elements['contact_email'].value.trim(),
    Status:          f.elements['status'].value,
    'Mandate Type':  f.elements['mandate_type'].value.trim(),
    'AUM (M€)':      f.elements['aum'].value.trim(),
    Notes:           f.elements['notes'].value.trim(),
  };
  try {
    await API.post('/api/companies', data);
    toast('Company updated', 'success');
    closeEditModal();
    await loadCompanies();
    renderAll();
    if (state.selectedCompany === data.Company) selectCompany(data.Company);
  } catch (e) { toast('Error: ' + e.message, 'error'); }
}

function openStatusModal(name, currentStatus) {
  document.getElementById('modal-status').style.display = 'flex';
  document.getElementById('status-company-name').textContent = name;
  document.getElementById('status-company-hidden').value     = name;
  document.getElementById('status-select').innerHTML = state.statuses.map(s =>
    `<option ${s===currentStatus?'selected':''}>${esc(s)}</option>`
  ).join('');
  document.getElementById('status-notes').value = '';
}
function closeStatusModal() { document.getElementById('modal-status').style.display = 'none'; }

async function submitStatusChange() {
  const name   = document.getElementById('status-company-hidden').value;
  const status = document.getElementById('status-select').value;
  const notes  = document.getElementById('status-notes').value.trim();
  try {
    await API.put(`/api/companies/${encodeURIComponent(name)}/status`,
      { status, notes: notes || undefined });
    toast(`Status updated → "${status}"`, 'success');
    closeStatusModal();
    await loadCompanies();
    renderAll();
    if (state.selectedCompany === name) selectCompany(name);
  } catch (e) { toast('Error: ' + e.message, 'error'); }
}

// ── Settings modal ─────────────────────────────────────────────────────────
function openSettingsModal() {
  document.getElementById('modal-settings').style.display = 'flex';
  document.getElementById('conn-test-result').textContent   = '';
  document.getElementById('ollama-test-result').textContent = '';
  Promise.all([
    fetch('/api/status').then(r => r.json()),
    fetch('/api/config').then(r => r.json()),
  ]).then(([d, cfg]) => {
    document.getElementById('settings-email').textContent       = d.email       || '—';
    document.getElementById('settings-imap').textContent        = d.imap_server  || '—';
    document.getElementById('settings-excel').textContent       = d.excel_path   || '—';
    document.getElementById('settings-ollama-host').textContent  = cfg.ollama_host  || '—';
    document.getElementById('settings-ollama-model').textContent = cfg.ollama_model || '—';
  }).catch(() => {});
}
function closeSettingsModal() { document.getElementById('modal-settings').style.display = 'none'; }

async function testConnection() {
  const el = document.getElementById('conn-test-result');
  el.textContent = 'Testing Outlook…';
  try {
    const d = await API.get('/api/test-connection');
    if (d.ok) {
      el.style.color = '#2e7d32';
      el.textContent = '✓ ' + d.message;
      setConnBadge('ok', '● connected');
    } else {
      el.style.color = '#c62828';
      el.textContent = '✗ ' + d.message;
      setConnBadge('error', '● disconnected');
    }
  } catch (e) {
    el.style.color = '#c62828';
    el.textContent = '✗ ' + e.message;
  }
}

async function testOllama() {
  const el = document.getElementById('ollama-test-result');
  el.style.color   = '#555';
  el.textContent   = 'Testing Ollama…';
  try {
    const d = await API.get('/api/test-ollama');
    if (d.ok) {
      el.style.color = '#2e7d32';
      el.textContent = '✓ ' + d.message;
    } else {
      el.style.color = '#c62828';
      el.textContent = '✗ ' + d.message;
    }
  } catch (e) {
    el.style.color = '#c62828';
    el.textContent = '✗ ' + e.message;
  }
}

// ── Auto-classify ──────────────────────────────────────────────────────────
async function autoClassify() {
  const btn     = document.getElementById('classify-btn');
  const panel   = document.getElementById('classify-panel');
  const log     = document.getElementById('classify-log');
  const heading = document.getElementById('classify-heading');
  const counter = document.getElementById('classify-counter');
  const summary = document.getElementById('classify-summary');
  const barWrap = document.getElementById('classify-progress-bar-wrap');
  const bar     = document.getElementById('classify-progress-bar');

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Classifying…';

  panel.style.display = 'block';
  log.innerHTML       = '';
  summary.style.display = 'none';
  barWrap.style.display = 'none';
  bar.style.width       = '0%';
  heading.textContent   = '✦ AI Classification — fetching emails…';
  counter.textContent   = '';
  panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

  const addLine = (html) => {
    const div = document.createElement('div');
    div.className = 'classify-status-line';
    div.innerHTML = html;
    log.appendChild(div);
    log.scrollTop = log.scrollHeight;
  };

  let total = 0, done = 0;

  try {
    const resp = await fetch('/api/auto-classify', { method: 'POST',
      headers: { 'Content-Type': 'application/json' }, body: '{}' });

    if (!resp.ok) {
      const t = await resp.text();
      throw new Error(t);
    }

    const reader  = resp.body.getReader();
    const decoder = new TextDecoder();
    let   buf     = '';

    while (true) {
      const { done: streamDone, value } = await reader.read();
      if (streamDone) break;

      buf += decoder.decode(value, { stream: true });
      const parts = buf.split('\n\n');
      buf = parts.pop();          // keep incomplete chunk

      for (const part of parts) {
        const line = part.trim();
        if (!line.startsWith('data:')) continue;
        let evt;
        try { evt = JSON.parse(line.slice(5).trim()); } catch { continue; }

        if (evt.type === 'status') {
          heading.textContent = '✦ AI Classification — ' + evt.message;

        } else if (evt.type === 'start') {
          total = evt.total;
          barWrap.style.display = 'block';
          heading.textContent   = `✦ AI Classification — processing ${total} domain(s)`;
          addLine(`<span style="color:#5c35c8;font-weight:600">Found ${total} email group(s). Classifying with Ollama…</span>`);

        } else if (evt.type === 'progress') {
          done = evt.index;
          bar.style.width  = Math.round((done / total) * 100) + '%';
          counter.textContent = `${done} / ${total}`;
          addLine(`<span class="cl-domain">${esc(evt.domain)}</span><span class="spinner" style="width:10px;height:10px;border-width:2px"></span>`);

        } else if (evt.type === 'result') {
          // Replace last spinner line with result
          const last = log.lastElementChild;
          if (last) last.remove();
          addLine(
            `<span class="cl-company">${esc(evt.company)}</span>` +
            `<span class="cl-domain">${esc(evt.domain)}</span>` +
            `<span class="badge badge-${esc(evt.status)} cl-status">${esc(evt.status)}</span>` +
            (evt.contact ? `<span style="font-size:11px;color:#555">${esc(evt.contact)}</span>` : '')
          );

        } else if (evt.type === 'error') {
          const last = log.lastElementChild;
          if (last) last.remove();
          addLine(
            `<span class="cl-domain">${esc(evt.domain)}</span>` +
            `<span class="cl-err">⚠ ${esc(evt.error)}</span>`
          );

        } else if (evt.type === 'done') {
          bar.style.width       = '100%';
          counter.textContent   = `${evt.classified} classified`;
          heading.textContent   = '✦ AI Classification — complete';
          summary.style.display = 'flex';
          summary.innerHTML =
            `<strong style="color:#2e7d32">✓ ${evt.classified} deal(s) saved to Excel</strong>` +
            (evt.errors ? `<span style="color:#c62828">${evt.errors} error(s)</span>` : '') +
            (evt.message ? `<span style="color:#888;font-size:12px">${esc(evt.message)}</span>` : '');
          // Auto-reload the pipeline table so results appear immediately
          reloadAfterClassify();

        } else if (evt.type === 'fatal') {
          heading.textContent = '✦ AI Classification — failed';
          addLine(`<span class="cl-err" style="white-space:pre-wrap">✗ ${esc(evt.error)}</span>`);
          addLine(`<span style="color:#888;font-size:11px">To fix: install Ollama from https://ollama.ai, then run:<br><code>ollama serve</code> and <code>ollama pull llama3.2</code></span>`);
          toast('Auto-classify failed: ' + evt.error, 'error');
        }
      }
    }
  } catch (e) {
    heading.textContent = '✦ AI Classification — failed';
    addLine(`<span class="cl-err">${esc(e.message)}</span>`);
    toast('Auto-classify error: ' + e.message, 'error');
  } finally {
    btn.disabled = false;
    btn.innerHTML = '✦ Auto-classify All';
  }
}

async function reloadAfterClassify() {
  await loadCompanies();
  renderAll();
  toast('Pipeline reloaded', 'success');
}

// ── Search / filter ────────────────────────────────────────────────────────
function onSearch()       { renderCompaniesTable(); }
function onFilterChange() { renderCompaniesTable(); }

function populateFilterDropdown() {
  const sel = document.getElementById('filter-status');
  const cur = sel.value;
  sel.innerHTML = '<option value="">All Statuses</option>' +
    state.statuses.map(s => `<option ${s===cur?'selected':''}>${esc(s)}</option>`).join('');
}

// ── Refresh / download ─────────────────────────────────────────────────────
async function refreshAll() {
  const btn = document.getElementById('refresh-btn');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Loading…';
  try {
    await Promise.all([loadCompanies(), loadEmails()]);
    renderAll();
    toast('Refreshed', 'success');
  } catch (e) {
    toast('Refresh failed: ' + e.message, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = '↺ Refresh Emails';
  }
}

function downloadExcel() { window.location.href = '/api/excel/download'; }

// ── Helpers ────────────────────────────────────────────────────────────────
function esc(str) {
  return String(str ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

document.addEventListener('DOMContentLoaded', init);
