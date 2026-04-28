/* PrecisionFG Fundraising Email Monitor */

const API = {
  async get(url) {
    const r = await fetch(url);
    if (r.status === 401) {
      const d = await r.json();
      showLoginPrompt(d.login_url);
      throw new Error('not_authenticated');
    }
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  },
  async post(url, body) {
    const r = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (r.status === 401) { const d = await r.json(); showLoginPrompt(d.login_url); throw new Error('not_authenticated'); }
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  },
  async put(url, body) {
    const r = await fetch(url, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (r.status === 401) { const d = await r.json(); showLoginPrompt(d.login_url); throw new Error('not_authenticated'); }
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  },
};

// ── State ─────────────────────────────────────────────────────────────────
let state = {
  companies: [],
  emails: [],
  emailsGrouped: [],
  statuses: [],
  selectedCompany: null,
  loading: false,
};

// ── Toast ─────────────────────────────────────────────────────────────────
function toast(msg, type = '') {
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.textContent = msg;
  document.getElementById('toast-container').appendChild(el);
  setTimeout(() => el.remove(), 3500);
}

// ── Login prompt ──────────────────────────────────────────────────────────
function showLoginPrompt(loginUrl) {
  const box = document.getElementById('login-overlay');
  document.getElementById('login-btn').href = loginUrl || '/auth/login';
  box.style.display = 'flex';
}

function hideLoginOverlay() {
  document.getElementById('login-overlay').style.display = 'none';
}

// ── Bootstrap ─────────────────────────────────────────────────────────────
async function init() {
  const [meData, statusData] = await Promise.all([
    fetch('/auth/me').then(r => r.json()),
    fetch('/api/statuses').then(r => r.json()),
  ]);

  state.statuses = statusData.statuses || [];

  if (!meData.authenticated) {
    showLoginPrompt('/auth/login');
    return;
  }

  hideLoginOverlay();
  document.getElementById('user-info').textContent = `${meData.name || meData.email}`;

  await Promise.all([loadCompanies(), loadEmails()]);
  renderAll();
}

// ── Data loaders ──────────────────────────────────────────────────────────
async function loadCompanies() {
  try {
    const data = await API.get('/api/companies');
    state.companies = data.companies || [];
  } catch (e) {
    if (e.message !== 'not_authenticated') toast('Failed to load companies: ' + e.message, 'error');
  }
}

async function loadEmails() {
  try {
    const data = await API.get('/api/emails?top=200');
    state.emails = data.messages || [];
    state.emailsGrouped = data.grouped || [];
    // Auto-populate companies from email domains if none exist
    autoPopulateFromEmails();
  } catch (e) {
    if (e.message !== 'not_authenticated') toast('Failed to load emails: ' + e.message, 'error');
  }
}

function autoPopulateFromEmails() {
  if (state.companies.length === 0 && state.emailsGrouped.length > 0) {
    // Suggest to user — don't auto-create
  }
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
    const s = c['Status'] || 'Initial Contact';
    counts[s] = (counts[s] || 0) + 1;
  });

  const statsBar = document.getElementById('stats-bar');
  const highlights = [
    { label: 'Total', val: state.companies.length, color: '#2E75B6' },
    { label: 'Active', val: (counts['In Discussion'] || 0) + (counts['Follow Up'] || 0) + (counts['Due Diligence'] || 0), color: '#1565C0' },
    { label: 'Proposals', val: counts['Proposal Sent'] || 0, color: '#2e7d32' },
    { label: 'Mandates', val: counts['Mandate Received'] || 0, color: '#1B5E20' },
    { label: 'Won', val: counts['Closed – Won'] || 0, color: '#1B5E20' },
    { label: 'Emails', val: state.emails.length, color: '#6a1b9a' },
  ];

  statsBar.innerHTML = highlights.map(h => `
    <div class="stat-card">
      <div class="stat-val" style="color:${h.color}">${h.val}</div>
      <div class="stat-label">${h.label}</div>
    </div>
  `).join('');
}

function renderSidebar() {
  const list = document.getElementById('sidebar-list');
  const combined = mergeCompaniesWithEmails();

  if (combined.length === 0) {
    list.innerHTML = '<div class="empty-state" style="padding:20px;font-size:12px;">No companies yet.<br>Add one or sync emails.</div>';
    return;
  }

  list.innerHTML = combined.map(c => `
    <div class="sidebar-item ${state.selectedCompany === c.name ? 'active' : ''}"
         onclick="selectCompany('${escHtml(c.name)}')">
      <div class="company-name">${escHtml(c.name)}</div>
      <div class="company-meta">
        <span class="badge badge-${escHtml(c.status || 'Initial Contact')}">${escHtml(c.status || '—')}</span>
        ${c.emailCount ? `· ${c.emailCount} emails` : ''}
      </div>
    </div>
  `).join('');
}

function mergeCompaniesWithEmails() {
  const byDomain = {};
  state.emailsGrouped.forEach(g => { byDomain[g.domain] = g; });

  const result = state.companies.map(c => ({
    name: c['Company'] || '',
    status: c['Status'] || '',
    lastDate: c['Last Email Date'] || '',
    emailCount: null,
    source: 'excel',
  }));

  // Add email-only domains not in companies list
  state.emailsGrouped.forEach(g => {
    const alreadyIn = state.companies.some(c =>
      (c['Company'] || '').toLowerCase().includes(g.domain.split('.')[0].toLowerCase())
    );
    if (!alreadyIn) {
      result.push({
        name: g.domain,
        status: '',
        lastDate: g.last_email_date,
        emailCount: g.message_count,
        source: 'email',
      });
    }
  });

  return result.sort((a, b) => (b.lastDate || '').localeCompare(a.lastDate || ''));
}

function renderCompaniesTable() {
  const tbody = document.getElementById('companies-tbody');
  const filterStatus = document.getElementById('filter-status').value;
  const searchQ = document.getElementById('search-input').value.toLowerCase();

  let filtered = state.companies.filter(c => {
    if (filterStatus && c['Status'] !== filterStatus) return false;
    if (searchQ && !(c['Company'] || '').toLowerCase().includes(searchQ) &&
        !(c['Contact Name'] || '').toLowerCase().includes(searchQ)) return false;
    return true;
  });

  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="7" class="empty-state">No companies match the filter.</td></tr>';
    return;
  }

  tbody.innerHTML = filtered.map(c => `
    <tr>
      <td><strong>${escHtml(c['Company'] || '')}</strong></td>
      <td>${escHtml(c['Contact Name'] || '')}<br><small style="color:#888">${escHtml(c['Contact Email'] || '')}</small></td>
      <td><span class="badge badge-${escHtml(c['Status'] || '')}">${escHtml(c['Status'] || '—')}</span></td>
      <td>${escHtml(c['Last Email Date'] || '—')}</td>
      <td>${escHtml(c['Mandate Type'] || '—')}</td>
      <td>${escHtml(c['AUM (M€)'] ? c['AUM (M€)'] + 'M€' : '—')}</td>
      <td style="white-space:nowrap">
        <button class="btn btn-outline btn-sm" onclick="openEditModal('${escHtml(c['Company'] || '')}')">Edit</button>
        <button class="btn btn-primary btn-sm" onclick="openStatusModal('${escHtml(c['Company'] || '')}', '${escHtml(c['Status'] || '')}')">Status</button>
      </td>
    </tr>
  `).join('');
}

function renderEmailsPanel(emails) {
  const panel = document.getElementById('emails-panel');
  if (!emails || emails.length === 0) {
    panel.innerHTML = '<div class="empty-state">No emails found for this company.</div>';
    return;
  }
  panel.innerHTML = emails.map(e => {
    const from = e.from?.emailAddress;
    const date = (e.receivedDateTime || '').replace('T', ' ').slice(0, 16);
    const isUnread = !e.isRead;
    return `
      <div class="email-item ${isUnread ? 'unread' : ''}">
        <div class="email-subject">
          <span class="${from?.address?.includes('sent') ? 'dir-out' : 'dir-in'}">${from?.address?.includes(document.getElementById('user-info').textContent.split('@')[0]) ? 'OUT' : 'IN'}</span>
          ${escHtml(e.subject || '(no subject)')}
        </div>
        <div class="email-meta">From: ${escHtml(from?.name || '')} &lt;${escHtml(from?.address || '')}&gt; · ${date}</div>
        <div class="email-preview">${escHtml((e.bodyPreview || '').slice(0, 180))}${e.bodyPreview?.length > 180 ? '…' : ''}</div>
      </div>
    `;
  }).join('');
}

// ── Company selection ─────────────────────────────────────────────────────
function selectCompany(name) {
  state.selectedCompany = name;
  renderSidebar();

  const company = state.companies.find(c => c['Company'] === name);
  const detailPanel = document.getElementById('company-detail');
  const detailTitle = document.getElementById('detail-title');

  if (!company && !state.emailsGrouped.find(g => g.domain === name)) {
    detailPanel.style.display = 'none';
    return;
  }

  detailPanel.style.display = 'block';
  detailTitle.textContent = name;

  if (company) {
    document.getElementById('detail-status').innerHTML = `<span class="badge badge-${escHtml(company['Status'] || '')}">${escHtml(company['Status'] || '—')}</span>`;
    document.getElementById('detail-contact').textContent = company['Contact Name'] ? `${company['Contact Name']} <${company['Contact Email'] || ''}>` : '—';
    document.getElementById('detail-mandate').textContent = company['Mandate Type'] || '—';
    document.getElementById('detail-aum').textContent = company['AUM (M€)'] ? company['AUM (M€)'] + ' M€' : '—';
    document.getElementById('detail-notes').textContent = company['Notes'] || '—';
    document.getElementById('detail-first').textContent = company['First Contact Date'] || '—';
    document.getElementById('detail-last').textContent = company['Last Email Date'] || '—';
    document.getElementById('detail-edit-btn').onclick = () => openEditModal(name);
    document.getElementById('detail-status-btn').onclick = () => openStatusModal(name, company['Status'] || '');
  }

  loadCompanyEmails(name);
}

async function loadCompanyEmails(companyName) {
  const panel = document.getElementById('emails-panel');
  const company = state.companies.find(c => c['Company'] === companyName);
  const searchTerm = company?.['Contact Email']
    ? company['Contact Email'].split('@')[1] || companyName
    : companyName;

  panel.innerHTML = '<div style="padding:20px;text-align:center"><span class="spinner"></span></div>';
  try {
    const data = await API.get(`/api/emails/search?q=${encodeURIComponent(searchTerm)}`);
    renderEmailsPanel(data.messages || []);
  } catch (e) {
    if (e.message !== 'not_authenticated') panel.innerHTML = `<div class="empty-state">Error: ${escHtml(e.message)}</div>`;
  }
}

// ── Modals ─────────────────────────────────────────────────────────────────
function openAddModal() {
  document.getElementById('modal-add').style.display = 'flex';
  document.getElementById('form-add').reset();
  // Populate status select
  const sel = document.getElementById('add-status');
  sel.innerHTML = state.statuses.map(s => `<option value="${escHtml(s)}">${escHtml(s)}</option>`).join('');
}

function closeAddModal() {
  document.getElementById('modal-add').style.display = 'none';
}

async function submitAddCompany() {
  const f = document.getElementById('form-add');
  const data = {
    Company:       f.elements['company'].value.trim(),
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
  } catch (e) {
    toast('Error: ' + e.message, 'error');
  }
}

function openEditModal(name) {
  const company = state.companies.find(c => c['Company'] === name);
  if (!company) return;
  const modal = document.getElementById('modal-edit');
  modal.style.display = 'flex';
  const f = document.getElementById('form-edit');
  f.elements['company'].value = company['Company'] || '';
  f.elements['contact_name'].value = company['Contact Name'] || '';
  f.elements['contact_email'].value = company['Contact Email'] || '';
  f.elements['mandate_type'].value = company['Mandate Type'] || '';
  f.elements['aum'].value = company['AUM (M€)'] || '';
  f.elements['notes'].value = company['Notes'] || '';

  const sel = f.elements['status'];
  sel.innerHTML = state.statuses.map(s => `<option value="${escHtml(s)}" ${s === company['Status'] ? 'selected' : ''}>${escHtml(s)}</option>`).join('');
}

function closeEditModal() {
  document.getElementById('modal-edit').style.display = 'none';
}

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
  } catch (e) {
    toast('Error: ' + e.message, 'error');
  }
}

function openStatusModal(name, currentStatus) {
  const modal = document.getElementById('modal-status');
  modal.style.display = 'flex';
  document.getElementById('status-company-name').textContent = name;
  document.getElementById('status-company-hidden').value = name;

  const sel = document.getElementById('status-select');
  sel.innerHTML = state.statuses.map(s => `<option value="${escHtml(s)}" ${s === currentStatus ? 'selected' : ''}>${escHtml(s)}</option>`).join('');
  document.getElementById('status-notes').value = '';
}

function closeStatusModal() {
  document.getElementById('modal-status').style.display = 'none';
}

async function submitStatusChange() {
  const name = document.getElementById('status-company-hidden').value;
  const status = document.getElementById('status-select').value;
  const notes = document.getElementById('status-notes').value.trim();
  try {
    await API.put(`/api/companies/${encodeURIComponent(name)}/status`, { status, notes: notes || undefined });
    toast(`Status updated to "${status}"`, 'success');
    closeStatusModal();
    await loadCompanies();
    renderAll();
    if (state.selectedCompany === name) selectCompany(name);
  } catch (e) {
    toast('Error: ' + e.message, 'error');
  }
}

// ── Search / filter ────────────────────────────────────────────────────────
function onSearch() { renderCompaniesTable(); }

function onFilterChange() { renderCompaniesTable(); populateFilterDropdown(); }

function populateFilterDropdown() {
  const sel = document.getElementById('filter-status');
  const current = sel.value;
  sel.innerHTML = '<option value="">All Statuses</option>' +
    state.statuses.map(s => `<option value="${escHtml(s)}" ${s === current ? 'selected' : ''}>${escHtml(s)}</option>`).join('');
}

// ── Sync & refresh ─────────────────────────────────────────────────────────
async function refreshAll() {
  const btn = document.getElementById('refresh-btn');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>';
  try {
    await Promise.all([loadCompanies(), loadEmails()]);
    renderAll();
    toast('Refreshed', 'success');
  } catch (e) {
    toast('Refresh failed: ' + e.message, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = 'Refresh';
  }
}

function downloadExcel() {
  window.location.href = '/api/excel/download';
}

// ── Helpers ────────────────────────────────────────────────────────────────
function escHtml(str) {
  return String(str ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── Start ──────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  populateFilterDropdown();
  init();
});
