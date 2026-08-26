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
  cachedAt:      0,
  followUpDays:  5,
  statuses:      [],
  workflowFilter: '',
  selectedCompany: null,
  selectedDeal: null,
  currentEmails: [],
  route: null,
};

const WORKFLOW = {
  needs_reply: { label: 'Needs reply',     stat: 'Needs Reply', color: '#b3261e' },
  no_answer:   { label: 'No answer yet',   stat: 'No Answer',   color: '#b06000' },
  meeting:     { label: 'Meeting signal',  stat: 'Meetings',    color: '#1b6f43' },
  waiting:     { label: 'Waiting',         stat: 'Waiting',     color: '#1565c0' },
  no_email:    { label: 'No email found',  stat: 'No Email',    color: '#6c757d' },
  closed:      { label: 'Closed/paused',   stat: 'Closed',      color: '#555' },
};

const WORKFLOW_RANK = {
  needs_reply: 1,
  no_answer: 2,
  meeting: 3,
  waiting: 4,
  no_email: 8,
  closed: 9,
};

const GENERIC_COMPANY_WORDS = new Set([
  'asset', 'assets', 'management', 'capital', 'partners', 'partner', 'group',
  'holdings', 'holding', 'limited', 'ltd', 'llc', 'inc', 'plc', 'fund', 'funds',
  'ventures', 'venture', 'family', 'office', 'investment', 'investments',
]);

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
  state.route = parseDealRoute();
  applyRouteChrome();

  const [statusData, apiStatus] = await Promise.all([
    fetch('/api/statuses').then(r => r.json()),
    fetch('/api/status').then(r => r.json()),
  ]);

  state.statuses = statusData.statuses || [];
  state.followUpDays = apiStatus.follow_up_days || 5;
  populateFilterDropdown();
  populateWorkflowDropdown();

  document.getElementById('user-info').textContent = apiStatus.email || '';
  document.getElementById('settings-email').textContent = apiStatus.email || '—';

  await Promise.all([loadCompanies(), loadEmails()]);
  renderAll();
  openRouteDeal();
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

async function loadEmails(forceRefresh = false) {
  setConnBadge('unknown', forceRefresh ? '● fetching from Outlook…' : '● loading cache…');
  try {
    const url  = forceRefresh ? '/api/emails?refresh=true' : '/api/emails';
    const data = await API.get(url);
    state.emails        = data.messages  || [];
    state.emailsGrouped = data.grouped   || [];
    state.emailsTotal   = data.total     || state.emails.length;
    state.cachedAt      = data.cached_at || 0;
    if (data.needs_refresh) {
      setConnBadge('unknown', '● emails: refresh needed');
    } else {
      setConnBadge('ok', _cacheAgeBadge(state.cachedAt));
    }
  } catch (e) {
    setConnBadge('error', '● error');
    toast('Email load failed: ' + e.message, 'error');
  }
}

function _cacheAgeBadge(ts) {
  if (!ts) return '● connected';
  const mins = Math.round((Date.now() / 1000 - ts) / 60);
  if (mins < 2)  return '● emails: just refreshed';
  if (mins < 60) return `● emails: ${mins}m ago`;
  const hrs = Math.round(mins / 60);
  return `● emails: ${hrs}h ago`;
}

function setConnBadge(type, text) {
  const b = document.getElementById('conn-badge');
  b.className = `conn-badge conn-${type}`;
  b.textContent = text;
}

// ── Routing ────────────────────────────────────────────────────────────────
function parseDealRoute() {
  const parts = window.location.pathname.split('/').filter(Boolean);
  if (parts[0] !== 'deal' || !parts[1] || !parts[2]) return null;
  return {
    type: parts[1],
    key: decodeURIComponent(parts.slice(2).join('/')),
  };
}

function isDealPage() {
  return Boolean(state.route);
}

function applyRouteChrome() {
  document.body.classList.toggle('deal-page', isDealPage());
  const back = document.getElementById('detail-back-link');
  if (back) back.style.display = isDealPage() ? '' : 'none';
}

function dealHref(row) {
  if (row?.source === 'email' && row.domain) {
    return `/deal/domain/${encodeURIComponent(row.domain)}`;
  }
  return `/deal/company/${encodeURIComponent(row?.name || '')}`;
}

function isSelectedRow(row) {
  if (!state.selectedDeal || !row) return state.selectedCompany === row?.name;
  if (state.selectedDeal.domain && row.domain && state.selectedDeal.domain === row.domain) return true;
  return state.selectedDeal.name === row.name;
}

function findRouteDealRow() {
  if (!state.route) return null;
  const rows = buildWorkflowRows();
  if (state.route.type === 'domain') {
    return rows.find(r => (r.domain || '').toLowerCase() === state.route.key.toLowerCase()) || null;
  }
  const exact = rows.find(r => r.name === state.route.key) ||
    rows.find(r => normalizeKey(r.name) === normalizeKey(state.route.key));
  if (exact) return exact;

  const saved = state.companies.find(c => c['Company'] === state.route.key) ||
    state.companies.find(c => normalizeKey(c['Company']) === normalizeKey(state.route.key));
  const group = saved ? findGroupForCompany(saved) : null;
  if (group?.domain) {
    return rows.find(r => (r.domain || '').toLowerCase() === group.domain.toLowerCase()) || null;
  }
  return null;
}

function openRouteDeal() {
  if (!isDealPage()) return;
  const row = findRouteDealRow();
  if (row) {
    showDealDetail(row, { scroll: false });
    return;
  }

  const panel = document.getElementById('company-detail');
  panel.style.display = 'block';
  document.getElementById('detail-title').textContent = 'Deal not found';
  document.getElementById('detail-status').innerHTML = '';
  document.getElementById('emails-panel').innerHTML =
    '<div class="empty-state">This deal is not in the current Excel file or email cache.</div>';
  document.getElementById('detail-editor').style.display = 'none';
  ['detail-edit-btn', 'detail-status-btn', 'detail-classify-btn', 'detail-log-btn'].forEach(id => {
    const btn = document.getElementById(id);
    if (btn) btn.style.display = 'none';
  });
}

// ── Render ─────────────────────────────────────────────────────────────────
function renderAll() {
  applyRouteChrome();
  renderStats();
  renderFocusQueue();
  renderSidebar();
  renderCompaniesTable();
}

// ── Workflow model ────────────────────────────────────────────────────────
function buildWorkflowRows({ includeEmailOnly = true, sortBy = 'recent' } = {}) {
  const usedDomains = new Set();
  const rows = state.companies.map(company => {
    const group = findGroupForCompany(company);
    const row = buildWorkflowRow(company, group, 'excel');
    if (row.domain) usedDomains.add(row.domain);
    return row;
  });

  if (includeEmailOnly) {
    state.emailsGrouped.forEach(group => {
      if (usedDomains.has(group.domain)) return;
      rows.push(buildWorkflowRow(groupToCompany(group), group, 'email'));
    });
  }

  return dedupeWorkflowRows(rows).sort(sortBy === 'workflow' ? sortByWorkflowPriority : sortByMostRecent);
}

function dedupeWorkflowRows(rows) {
  const byDomain = new Map();
  const withoutDomain = [];

  rows.forEach(row => {
    const key = (row.domain || '').toLowerCase();
    if (!key) {
      withoutDomain.push(row);
      return;
    }
    const existing = byDomain.get(key);
    if (!existing || rowQualityScore(row) > rowQualityScore(existing)) {
      byDomain.set(key, row);
    }
  });

  return [...byDomain.values(), ...withoutDomain];
}

function rowQualityScore(row) {
  const c = row.company || {};
  const filledFields = ['Contact Name', 'Contact Email', 'Status', 'Mandate Type', 'AUM (M€)', 'Notes']
    .filter(key => String(c[key] || '').trim()).length;
  return (
    (row.source === 'excel' ? 1000 : 0) +
    companyNameQuality(row.name, row.domain) +
    filledFields * 10 +
    Math.min(row.messageCount || 0, 40)
  );
}

function companyNameQuality(name, domain) {
  const clean = String(name || '').trim();
  if (!clean) return 0;
  const compact = normalizeKey(clean);
  const domainBase = normalizeKey((domain || '').split('.')[0]);
  let score = compact.length;
  if (/\s/.test(clean)) score += 25;
  if (domainBase && compact === domainBase) score -= 20;
  return score;
}

function sortByMostRecent(a, b) {
  const aTime = dateSortValue(a.lastDate);
  const bTime = dateSortValue(b.lastDate);
  if (aTime !== bTime) return bTime - aTime;
  return String(a.name || '').localeCompare(String(b.name || ''));
}

function sortByWorkflowPriority(a, b) {
  if (a.rank !== b.rank) return a.rank - b.rank;
  if (a.workflowState === 'no_answer' && b.workflowState === 'no_answer') {
    return (b.daysSinceLast ?? -1) - (a.daysSinceLast ?? -1);
  }
  return sortByMostRecent(a, b);
}

function dateSortValue(value) {
  if (!value) return 0;
  const time = new Date(value).getTime();
  return Number.isNaN(time) ? 0 : time;
}

function buildWorkflowRow(company, group, source) {
  const status = company['Status'] || '';
  const lastDate = group?.last_email_date || company['Last Email Date'] || '';
  const days = daysSince(lastDate);
  const direction = group?.latest_direction || '';
  const meeting = getMeetingInfo(group);
  const isClosed = isClosedStatus(status);

  let workflowState = 'no_email';
  if (isClosed) {
    workflowState = 'closed';
  } else if (meeting && meeting.signal !== 'cancelled' && meeting.daysSince <= 45) {
    workflowState = 'meeting';
  } else if (direction === 'IN') {
    workflowState = 'needs_reply';
  } else if (direction === 'OUT') {
    workflowState = days !== null && days >= state.followUpDays ? 'no_answer' : 'waiting';
  } else if (group) {
    workflowState = 'waiting';
  }

  const workflow = WORKFLOW[workflowState] || WORKFLOW.no_email;
  return {
    company,
    source,
    name: company['Company'] || group?.domain || '',
    status,
    group,
    domain: group?.domain || getEmailDomain(company['Contact Email']) || '',
    workflowState,
    workflowLabel: workflow.label,
    workflowColor: workflow.color,
    rank: WORKFLOW_RANK[workflowState] || 10,
    nextAction: nextActionFor(workflowState, days, meeting),
    reason: reasonFor(workflowState, days, meeting, direction),
    lastDate,
    lastDateLabel: formatDateShort(lastDate),
    daysSinceLast: days,
    direction,
    latestSubject: group?.latest_subject || '',
    latestContact: group?.latest_contact || company['Contact Email'] || '',
    messageCount: group?.message_count || 0,
    meeting,
  };
}

function groupToCompany(group) {
  const contact = (group.contacts || [])[0] || '';
  const match = contact.match(/<([^>]+)>/);
  const email = match ? match[1] : (contact.includes('@') ? contact.split(' ').pop() : '');
  return {
    Company: titleFromDomain(group.domain),
    'Contact Name': contact.replace(/<[^>]+>/, '').trim(),
    'Contact Email': email,
    Status: '',
    'Last Email Date': group.last_email_date || '',
  };
}

function findGroupForCompany(company) {
  const contactDomain = getEmailDomain(company['Contact Email']);
  if (contactDomain) {
    const exact = state.emailsGrouped.find(g => g.domain === contactDomain);
    if (exact) return exact;
  }

  const nameTokens = significantTokens(company['Company']);
  if (!nameTokens.length) return null;

  return state.emailsGrouped.find(group => {
    const domainBase = normalizeKey((group.domain || '').split('.')[0]);
    if (!domainBase) return false;
    return nameTokens.some(token => domainBase.includes(token) || token.includes(domainBase));
  }) || null;
}

function getMeetingInfo(group) {
  if (!group?.latest_meeting_subject) return null;
  const days = daysSince(group.latest_meeting_date);
  return {
    signal: group.latest_meeting_signal || 'scheduled',
    subject: group.latest_meeting_subject || '',
    date: group.latest_meeting_date || '',
    folder: group.latest_meeting_folder || '',
    direction: (group.latest_meeting_folder || '').toLowerCase() === 'sent' ? 'OUT' : 'IN',
    daysSince: days ?? 999,
  };
}

function nextActionFor(stateName, days, meeting) {
  if (stateName === 'needs_reply') return 'Reply to them';
  if (stateName === 'no_answer') return 'Send follow-up';
  if (stateName === 'meeting') {
    if (meeting?.signal === 'accepted') return 'Prepare meeting';
    if (meeting?.signal === 'tentative') return 'Confirm meeting';
    return 'Check meeting';
  }
  if (stateName === 'waiting') {
    const left = Math.max(0, state.followUpDays - (days || 0));
    return left ? `Wait ${left}d` : 'Monitor';
  }
  if (stateName === 'no_email') return 'Start outreach';
  return 'No action';
}

function reasonFor(stateName, days, meeting, direction) {
  if (stateName === 'needs_reply') return `They wrote ${ageLabel(days)}`;
  if (stateName === 'no_answer') return `You sent ${ageLabel(days)} with no reply`;
  if (stateName === 'meeting') return `${meetingLabel(meeting?.signal)} ${ageLabel(meeting?.daysSince)}`;
  if (stateName === 'waiting') return direction === 'OUT' ? `You sent ${ageLabel(days)}` : 'Recent thread';
  if (stateName === 'no_email') return 'No matched thread';
  return 'Deal is closed or paused';
}

function isClosedStatus(status) {
  return ['Closed – Won', 'Closed – Lost', 'Not Interested', 'On Hold'].includes(status || '');
}

function meetingLabel(signal) {
  if (signal === 'accepted') return 'Meeting accepted';
  if (signal === 'tentative') return 'Tentative meeting';
  if (signal === 'cancelled') return 'Meeting cancelled';
  return 'Meeting noted';
}

function getEmailDomain(email) {
  const raw = (email || '').toLowerCase().trim();
  const match = raw.match(/[a-z0-9._%+-]+@([a-z0-9.-]+\.[a-z]{2,})/);
  return match ? match[1] : '';
}

function normalizeKey(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function significantTokens(value) {
  const tokens = String(value || '')
    .toLowerCase()
    .split(/[^a-z0-9]+/)
    .filter(token => token.length > 2 && !GENERIC_COMPANY_WORDS.has(token));
  const expanded = new Set(tokens);
  tokens.forEach(token => {
    if (token.endsWith('ical') && token.length > 6) expanded.add(token.slice(0, -4));
    if (token.endsWith('medical') && token.length > 8) expanded.add(`${token.slice(0, -7)}med`);
  });
  return [...expanded];
}

function daysSince(value) {
  if (!value) return null;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return null;
  return Math.max(0, Math.floor((Date.now() - date.getTime()) / 86400000));
}

function ageLabel(days) {
  if (days === null || days === undefined || days === 999) return '';
  if (days <= 0) return 'today';
  if (days === 1) return '1 day ago';
  return `${days} days ago`;
}

function formatDateShort(value) {
  if (!value) return '—';
  return String(value).replace('T', ' ').slice(0, 16);
}

function titleFromDomain(domain) {
  const first = (domain || '').split('.')[0] || '';
  return first
    .split(/[-_]/)
    .filter(Boolean)
    .map(s => s.charAt(0).toUpperCase() + s.slice(1))
    .join(' ') || domain;
}

function workflowChip(row) {
  return `<span class="workflow-pill workflow-${row.workflowState}">${esc(row.workflowLabel)}</span>`;
}

function directionBadge(direction) {
  if (direction === 'OUT') return '<span class="dir-out">OUT</span>';
  if (direction === 'IN') return '<span class="dir-in">IN</span>';
  return '';
}

function renderStats() {
  const rows = buildWorkflowRows();
  const counts = rows.reduce((acc, row) => {
    acc[row.workflowState] = (acc[row.workflowState] || 0) + 1;
    return acc;
  }, {});
  const activeDeals = buildWorkflowRows({ includeEmailOnly: false })
    .filter(row => !isClosedStatus(row.status)).length;

  const highlights = [
    { label: WORKFLOW.needs_reply.stat, val: counts.needs_reply || 0, color: WORKFLOW.needs_reply.color, filter: 'needs_reply' },
    { label: WORKFLOW.no_answer.stat,   val: counts.no_answer || 0,   color: WORKFLOW.no_answer.color,   filter: 'no_answer' },
    { label: WORKFLOW.meeting.stat,     val: counts.meeting || 0,     color: WORKFLOW.meeting.color,     filter: 'meeting' },
    { label: WORKFLOW.waiting.stat,     val: counts.waiting || 0,     color: WORKFLOW.waiting.color,     filter: 'waiting' },
    { label: 'Active Deals',            val: activeDeals,             color: '#2E75B6',                  filter: '' },
    { label: 'Emails',                  val: state.emailsTotal,       color: '#6a1b9a',                  filter: '' },
  ];

  document.getElementById('stats-bar').innerHTML = highlights.map(h => `
    <div class="stat-card ${state.workflowFilter === h.filter && h.filter ? 'active' : ''}"
         onclick="setWorkflowFilter(${JSON.stringify(h.filter)})">
      <div class="stat-val" style="color:${h.color}">${h.val}</div>
      <div class="stat-label">${h.label}</div>
    </div>`).join('');
}

function renderFocusQueue() {
  const el = document.getElementById('focus-queue');
  if (!el) return;

  const rows = buildWorkflowRows({ sortBy: 'workflow' })
    .filter(row => ['needs_reply', 'no_answer', 'meeting'].includes(row.workflowState))
    .slice(0, 8);

  if (!rows.length) {
    el.innerHTML = '<div class="empty-state compact">No urgent workflow items.</div>';
    return;
  }

  el.innerHTML = rows.map(row => `
    <a class="focus-row" href="${esc(dealHref(row))}">
      <div class="focus-main">
        <strong>${esc(row.name)}</strong>
        <span class="focus-reason">${directionBadge(row.direction)} ${esc(row.reason)}</span>
      </div>
      <div class="focus-status">${workflowChip(row)}</div>
      <div class="focus-action">${esc(row.nextAction)}</div>
      <div class="focus-date">${esc(row.lastDateLabel)}</div>
    </a>
  `).join('');
}

function renderSidebar() {
  const list = document.getElementById('sidebar-list');
  const combined = mergeCompaniesWithEmails();
  if (combined.length === 0) {
    list.innerHTML = '<div class="empty-state" style="padding:20px;font-size:12px">No companies yet.<br>Add one or refresh emails.</div>';
    return;
  }
  list.innerHTML = combined.map(row => `
    <a class="sidebar-item ${isSelectedRow(row) ? 'active' : ''}" href="${esc(dealHref(row))}">
      <div class="company-name">${esc(row.name)}</div>
      <div class="company-meta">
        ${workflowChip(row)}
        ${row.messageCount ? ` · ${row.messageCount} emails` : ''}
        ${row.lastDate ? ` · ${row.lastDate.slice(0,10)}` : ''}
      </div>
    </a>`).join('');
}

function mergeCompaniesWithEmails() {
  const rows = buildWorkflowRows({ sortBy: 'recent' });
  if (!state.workflowFilter) return rows;
  return rows.filter(row => row.workflowState === state.workflowFilter);
}

function renderCompaniesTable() {
  const tbody    = document.getElementById('companies-tbody');
  const filterSt = document.getElementById('filter-status').value;
  const filterWf = document.getElementById('filter-workflow')?.value || '';
  const searchQ  = document.getElementById('search-input').value.toLowerCase();

  const filtered = buildWorkflowRows({ includeEmailOnly: false, sortBy: 'recent' }).filter(row => {
    const c = row.company;
    if (filterSt && row.status !== filterSt) return false;
    if (filterWf && row.workflowState !== filterWf) return false;
    if (searchQ) {
      const haystack = [
        c['Company'], c['Contact Name'], c['Contact Email'],
        row.workflowLabel, row.reason, row.latestSubject, row.domain,
      ].join(' ').toLowerCase();
      if (!haystack.includes(searchQ)) return false;
    }
    return true;
  });

  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="8" class="empty-state">No companies match.</td></tr>';
    return;
  }

  tbody.innerHTML = filtered.map(row => {
    const c = row.company;
    const subject = row.latestSubject ? `<div class="muted-line">${esc(row.latestSubject.slice(0, 64))}</div>` : '';
    const href = dealHref(row);
    return `
    <tr class="clickable-row" onclick="window.location.href='${esc(href)}'">
      <td><a class="deal-link" href="${esc(href)}" onclick="event.stopPropagation()"><strong>${esc(c['Company']||'')}</strong></a></td>
      <td>
        ${workflowChip(row)}
        <div class="muted-line">${esc(row.reason)}</div>
      </td>
      <td><span class="badge badge-${esc(c['Status']||'')}">${esc(c['Status']||'—')}</span></td>
      <td>${esc(c['Contact Name']||'')}${c['Contact Email'] ? `<br><small style="color:#888">${esc(c['Contact Email'])}</small>` : ''}</td>
      <td>${directionBadge(row.direction)} ${esc(row.lastDateLabel)}${subject}</td>
      <td><strong>${esc(row.nextAction)}</strong></td>
      <td>${esc(c['Mandate Type']||'—')}${c['AUM (M€)'] ? `<br><small style="color:#888">${esc(c['AUM (M€)'])}M€</small>` : ''}</td>
      <td style="white-space:nowrap;display:flex;gap:4px">
        <button class="btn btn-outline btn-sm" onclick="event.stopPropagation();openEditModal(${JSON.stringify(c['Company']||'')})">Edit</button>
        <button class="btn btn-primary btn-sm" onclick="event.stopPropagation();openStatusModal(${JSON.stringify(c['Company']||'')},${JSON.stringify(c['Status']||'')})">Status</button>
      </td>
    </tr>`;
  }).join('');
}

// ── Company detail ─────────────────────────────────────────────────────────
function selectCompany(name) {
  const company = state.companies.find(c => c['Company'] === name);
  const row = buildWorkflowRows().find(r => r.name === name) ||
    (company ? buildWorkflowRow(company, findGroupForCompany(company), 'excel') : null);
  if (row) showDealDetail(row, { scroll: true });
}

function showDealDetail(row, { scroll = false } = {}) {
  state.selectedDeal = row;
  state.selectedCompany = row.name;
  renderSidebar();

  const panel = document.getElementById('company-detail');
  panel.style.display = 'block';
  document.getElementById('detail-editor').style.display = '';
  document.getElementById('detail-title').textContent = row.name;

  const company = state.companies.find(c => c['Company'] === row.name);
  const detailCompany = company || row?.company || {};

  document.getElementById('detail-status').innerHTML =
    `<span class="badge badge-${esc(detailCompany['Status']||'')}">${esc(detailCompany['Status']||'—')}</span>`;
  document.getElementById('detail-workflow').innerHTML = row ? workflowChip(row) : '—';
  document.getElementById('detail-next-action').textContent = row?.nextAction || '—';
  document.getElementById('detail-workflow-reason').textContent = row?.reason || '—';
  document.getElementById('detail-contact').textContent =
    detailCompany['Contact Name'] ? `${detailCompany['Contact Name']} <${detailCompany['Contact Email']||''}>` : (row?.latestContact || '—');
  document.getElementById('detail-mandate').textContent = detailCompany['Mandate Type'] || '—';
  document.getElementById('detail-aum').textContent     = detailCompany['AUM (M€)'] ? detailCompany['AUM (M€)'] + ' M€' : '—';
  document.getElementById('detail-notes').textContent   = detailCompany['Notes'] || '—';
  document.getElementById('detail-first').textContent   = detailCompany['First Contact Date'] || '—';
  document.getElementById('detail-last').textContent    = row?.lastDateLabel || detailCompany['Last Email Date'] || '—';
  document.getElementById('detail-domain').textContent  = row.domain || '—';

  document.getElementById('detail-edit-btn').onclick = () => {
    if (company) openEditModal(row.name);
    else document.getElementById('form-detail').scrollIntoView({ behavior: 'smooth', block: 'center' });
  };
  document.getElementById('detail-status-btn').onclick = () => openStatusModal(row.name, detailCompany['Status'] || 'Initial Contact');
  document.getElementById('detail-classify-btn').onclick = classifyCurrentDeal;
  document.getElementById('detail-log-btn').onclick = () => logEmailsForCompany(getDetailFormData().Company || row.name);
  document.getElementById('thread-live-btn').onclick = () => loadCompanyEmails(row.name, detailCompany, row, { live: true });
  ['detail-edit-btn', 'detail-status-btn', 'detail-classify-btn', 'detail-log-btn'].forEach(id => {
    document.getElementById(id).style.display = '';
  });
  document.getElementById('detail-edit-btn').textContent = company ? 'Edit' : 'Save Deal';

  populateDetailForm(row, detailCompany, company);

  state.currentEmails = [];
  if (scroll) panel.scrollIntoView({ behavior: 'smooth', block: 'start' });
  loadCompanyEmails(row.name, detailCompany, row);
}

function populateDetailForm(row, company, existingCompany = null) {
  const f = document.getElementById('form-detail');
  if (!f) return;

  const status = company['Status'] || 'Initial Contact';
  f.elements['original_company'].value = existingCompany?.['Company'] || '';
  f.elements['source'].value = row.source || '';
  f.elements['company'].value = company['Company'] || row.name || '';
  f.elements['contact_name'].value = company['Contact Name'] || contactNameFromRow(row);
  f.elements['contact_email'].value = company['Contact Email'] || contactEmailFromRow(row);
  f.elements['mandate_type'].value = company['Mandate Type'] || '';
  f.elements['aum'].value = company['AUM (M€)'] || '';
  f.elements['notes'].value = company['Notes'] || '';
  f.elements['status'].innerHTML = state.statuses.map(s =>
    `<option ${s===status?'selected':''}>${esc(s)}</option>`
  ).join('');

  const source = document.getElementById('detail-editor-source');
  source.textContent = existingCompany ? 'Saved in Excel' : 'Email-only deal';
}

function contactNameFromRow(row) {
  const contact = row?.latestContact || (row?.group?.contacts || [])[0] || '';
  return contact.replace(/<[^>]+>/, '').trim();
}

function contactEmailFromRow(row) {
  const contact = row?.latestContact || (row?.group?.contacts || [])[0] || '';
  const match = contact.match(/<([^>]+)>/) || contact.match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i);
  return match ? (match[1] || match[0]) : '';
}

function getDetailFormData() {
  const f = document.getElementById('form-detail');
  const row = state.selectedDeal || {};
  const contactEmail = f.elements['contact_email'].value.trim();
  return {
    Company:         f.elements['company'].value.trim(),
    'Contact Name':  f.elements['contact_name'].value.trim(),
    'Contact Email': contactEmail,
    Status:          f.elements['status'].value,
    'Mandate Type':  f.elements['mandate_type'].value.trim(),
    'AUM (M€)':      f.elements['aum'].value.trim(),
    'Last Email Date': row.lastDate ? String(row.lastDate).slice(0, 10) : '',
    Notes:           f.elements['notes'].value.trim(),
    _Domain:         row.domain || getEmailDomain(contactEmail),
  };
}

async function submitDetailUpdate() {
  const f = document.getElementById('form-detail');
  const original = f.elements['original_company'].value.trim();
  const data = getDetailFormData();
  if (!data.Company) { toast('Deal name is required', 'error'); return; }

  try {
    if (original) {
      await API.put(`/api/companies/${encodeURIComponent(original)}/details`, data);
    } else {
      await API.post('/api/companies', data);
    }
    toast('Deal corrections saved', 'success');
    await loadCompanies();
    renderAll();

    const href = `/deal/company/${encodeURIComponent(data.Company)}`;
    if (isDealPage() && window.location.pathname !== href) {
      history.replaceState(null, '', href);
      state.route = parseDealRoute();
      applyRouteChrome();
    }
    selectCompany(data.Company);
  } catch (e) {
    toast('Save failed: ' + e.message, 'error');
  }
}

function applyClassificationResult(result) {
  const f = document.getElementById('form-detail');
  if (!f || !result) return;
  if (result.Company) f.elements['company'].value = result.Company;
  if (result['Contact Name']) f.elements['contact_name'].value = result['Contact Name'];
  if (result['Contact Email']) f.elements['contact_email'].value = result['Contact Email'];
  if (result.Status && state.statuses.includes(result.Status)) f.elements['status'].value = result.Status;
  if (result['Mandate Type']) f.elements['mandate_type'].value = result['Mandate Type'];
  if (result['AUM (M€)'] !== undefined) f.elements['aum'].value = result['AUM (M€)'];
  if (result.Notes) f.elements['notes'].value = result.Notes;
}

async function classifyCurrentDeal() {
  const row = state.selectedDeal;
  if (!row) return;

  const btn = document.getElementById('detail-classify-btn');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Classifying...';
  try {
    const emails = state.currentEmails.length
      ? state.currentEmails
      : await loadCompanyEmails(row.name, row.company, row);
    const data = await API.post('/api/deals/classify', {
      domain: row.domain,
      company_name: row.name,
      contacts: row.group?.contacts || [],
      emails,
      last_email_date: row.lastDate,
    });
    applyClassificationResult(data.result);
    toast('Classification loaded for review', 'success');
  } catch (e) {
    toast('Classification failed: ' + e.message, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = 'Classify';
  }
}

async function loadCompanyEmails(name, company, row = null, { live = false } = {}) {
  const panel = document.getElementById('emails-panel');
  panel.innerHTML = '<div style="padding:16px;text-align:center"><span class="spinner"></span></div>';
  setThreadSource(live ? 'Searching Outlook...' : 'Loading cached thread...');

  const searchTerm = row?.domain || (company?.['Contact Email']
    ? company['Contact Email'].split('@')[1] || name
    : name);

  try {
    const url = live
      ? `/api/emails/search?q=${encodeURIComponent(searchTerm)}`
      : `/api/emails/thread?${new URLSearchParams({
          domain: row?.domain || '',
          email: company?.['Contact Email'] || '',
          company: name || '',
          limit: '160',
        }).toString()}`;
    const data = await API.get(url);
    state.currentEmails = data.messages || [];
    renderEmailsPanel(state.currentEmails, {
      source: live ? 'outlook' : data.source,
      total: data.total ?? state.currentEmails.length,
      cachedAt: data.cached_at,
    });
    return state.currentEmails;
  } catch (e) {
    setThreadSource('');
    panel.innerHTML = `<div class="empty-state">Could not load emails: ${esc(e.message)}</div>`;
    return [];
  }
}

function setThreadSource(text) {
  const el = document.getElementById('thread-source');
  if (el) el.textContent = text || '';
}

function renderEmailsPanel(emails, meta = {}) {
  const panel = document.getElementById('emails-panel');
  const orderedEmails = [...emails].sort((a, b) => dateSortValue(b.receivedDateTime) - dateSortValue(a.receivedDateTime));
  state.currentEmails = orderedEmails;
  const source = meta.source === 'outlook' ? 'Live Outlook search' : 'Cached thread';
  const total = meta.total ?? orderedEmails.length;
  const cacheAge = meta.cachedAt ? ` - cache ${_cacheAgeText(meta.cachedAt)}` : '';
  setThreadSource(`${source} - ${orderedEmails.length}${total > orderedEmails.length ? ` of ${total}` : ''} email(s)${cacheAge}`);

  if (!orderedEmails.length) {
    panel.innerHTML = '<div class="empty-state">No emails found matching this company.</div>';
    return;
  }
  panel.innerHTML = orderedEmails.map((e, idx) => {
    const sender  = (e.sender?.emailAddress) || (e.from?.emailAddress) || {};
    const external = ((e.external_participants || [])[0]?.emailAddress) || (e.from?.emailAddress) || {};
    const date    = (e.receivedDateTime || '').replace('T',' ').slice(0,16);
    const isOut   = e.direction ? e.direction === 'OUT' : (e.folder || '').toLowerCase() === 'sent';
    const recipients = [...(e.toRecipients || []), ...(e.ccRecipients || [])]
      .map(r => (r.emailAddress || {}).address)
      .filter(Boolean)
      .slice(0, 5)
      .join(', ');
    return `
      <div class="email-item" tabindex="0" title="Open email"
           ondblclick="openEmailModal(${idx})"
           onkeydown="if(event.key==='Enter')openEmailModal(${idx})">
        <div class="email-subject">
          <span class="${isOut ? 'dir-out' : 'dir-in'}">${isOut ? 'OUT' : 'IN'}</span>
          ${esc(e.subject || '(no subject)')}
        </div>
        <div class="email-meta">
          Client: ${esc(external.name || '')} &lt;${esc(external.address || '')}&gt; ·
          From: ${esc(sender.name || '')} &lt;${esc(sender.address || '')}&gt;${recipients ? ` · To/Cc: ${esc(recipients)}` : ''} · ${esc(date)}
        </div>
        ${e.bodyPreview ? `<div class="email-preview">${esc((e.bodyPreview||'').slice(0,200))}${(e.bodyPreview||'').length>200?'…':''}</div>` : ''}
      </div>`;
  }).join('');
}

function openEmailModal(index) {
  const email = state.currentEmails[index];
  if (!email) return;

  renderEmailModal(email, { loading: Boolean(email.id) });
  document.getElementById('modal-email').style.display = 'flex';
  if (email.id) hydrateEmailModal(index, email.id);
}

function closeEmailModal() {
  document.getElementById('modal-email').style.display = 'none';
}

async function hydrateEmailModal(index, messageId) {
  try {
    const data = await API.post('/api/emails/message', { id: messageId });
    const email = data.message || {};
    if (email.id) state.currentEmails[index] = { ...state.currentEmails[index], ...email };
    renderEmailModal(state.currentEmails[index] || email, {
      loading: false,
      source: data.source,
      warning: data.warning,
    });
    if (data.warning) toast('Opened cached email preview; Outlook full body was unavailable.', 'error');
  } catch (e) {
    renderEmailModal(state.currentEmails[index], { loading: false, warning: e.message });
    toast('Could not load full email: ' + e.message, 'error');
  }
}

function renderEmailModal(email, { loading = false, warning = '', source = '' } = {}) {
  const sender = (email.sender?.emailAddress) || (email.from?.emailAddress) || {};
  const date = (email.receivedDateTime || '').replace('T', ' ').slice(0, 16);
  const isOut = email.direction ? email.direction === 'OUT' : (email.folder || '').toLowerCase() === 'sent';
  const to = formatEmailList(email.toRecipients || []);
  const cc = formatEmailList(email.ccRecipients || []);
  const body = emailBodyText(email);
  const sourceText = source === 'outlook' ? 'Outlook' : source === 'cache' ? 'Cache' : '';

  document.getElementById('email-modal-direction').innerHTML =
    `<span class="${isOut ? 'dir-out' : 'dir-in'}">${isOut ? 'OUT' : 'IN'}</span>`;
  document.getElementById('email-modal-subject').textContent = email.subject || '(no subject)';
  document.getElementById('email-modal-meta').innerHTML = [
    `<div><strong>From:</strong> ${esc(formatEmailPerson(sender))}</div>`,
    to ? `<div><strong>To:</strong> ${esc(to)}</div>` : '',
    cc ? `<div><strong>Cc:</strong> ${esc(cc)}</div>` : '',
    `<div><strong>Date:</strong> ${esc(date || '—')}${sourceText ? ` · ${esc(sourceText)}` : ''}</div>`,
  ].filter(Boolean).join('');
  const visibleBody = esc(body || email.bodyPreview || '');
  document.getElementById('email-modal-body').innerHTML = loading
    ? `${visibleBody}${visibleBody ? '<br><br>' : ''}<div class="email-modal-loading"><span class="spinner"></span> Loading email content...</div>`
    : `${warning ? `<div class="email-modal-loading">${esc(warning)}</div><br>` : ''}${visibleBody || '<div class="empty-state compact">No content available.</div>'}`;
}

function emailBodyText(email) {
  const raw = email?.body?.content || email?.body || email?.bodyPreview || '';
  if (typeof raw !== 'string') return '';
  if (!/<[a-z][\s\S]*>/i.test(raw)) return raw.trim();
  const normalized = raw
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<\/div>/gi, '\n');
  return normalized.replace(/<[^>]+>/g, '').replace(/\n{3,}/g, '\n\n').trim();
}

function formatEmailPerson(person) {
  const name = person?.name || '';
  const address = person?.address || '';
  if (name && address) return `${name} <${address}>`;
  return name || address || '—';
}

function formatEmailList(recipients) {
  return recipients
    .map(r => formatEmailPerson((r || {}).emailAddress || {}))
    .filter(Boolean)
    .join(', ');
}

function _cacheAgeText(ts) {
  const mins = Math.round((Date.now() / 1000 - ts) / 60);
  if (mins < 2) return 'just now';
  if (mins < 60) return `${mins}m old`;
  return `${Math.round(mins / 60)}h old`;
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
  data._Domain = getEmailDomain(data['Contact Email']);
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
  f.elements['original_company'].value = c['Company'] || '';
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
  const original = f.elements['original_company'].value.trim();
  const data = {
    Company:         f.elements['company'].value.trim(),
    'Contact Name':  f.elements['contact_name'].value.trim(),
    'Contact Email': f.elements['contact_email'].value.trim(),
    Status:          f.elements['status'].value,
    'Mandate Type':  f.elements['mandate_type'].value.trim(),
    'AUM (M€)':      f.elements['aum'].value.trim(),
    Notes:           f.elements['notes'].value.trim(),
  };
  data._Domain = getEmailDomain(data['Contact Email']);
  try {
    await API.put(`/api/companies/${encodeURIComponent(original || data.Company)}/details`, data);
    toast('Company updated', 'success');
    closeEditModal();
    await loadCompanies();
    renderAll();
    if (isDealPage()) {
      history.replaceState(null, '', `/deal/company/${encodeURIComponent(data.Company)}`);
      state.route = parseDealRoute();
      applyRouteChrome();
    }
    if (state.selectedCompany === original || state.selectedCompany === data.Company) selectCompany(data.Company);
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
    const existing = state.companies.find(c => c['Company'] === name);
    if (existing) {
      await API.put(`/api/companies/${encodeURIComponent(name)}/status`,
        { status, notes: notes || undefined });
    } else {
      const data = getDetailFormData();
      data.Company = data.Company || name;
      data.Status = status;
      if (notes) data.Notes = notes;
      await API.post('/api/companies', data);
    }
    toast(`Status updated → "${status}"`, 'success');
    closeStatusModal();
    await loadCompanies();
    renderAll();
    const savedName = existing ? name : (getDetailFormData().Company || name);
    if (isDealPage() && !existing) {
      history.replaceState(null, '', `/deal/company/${encodeURIComponent(savedName)}`);
      state.route = parseDealRoute();
      applyRouteChrome();
    }
    if (state.selectedCompany === name || state.selectedCompany === savedName) selectCompany(savedName);
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
    document.getElementById('settings-followup-days').textContent = d.follow_up_days || cfg.follow_up_days || state.followUpDays;
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
function onSearch() { renderCompaniesTable(); }
function onFilterChange() {
  state.workflowFilter = document.getElementById('filter-workflow')?.value || '';
  renderStats();
  renderFocusQueue();
  renderSidebar();
  renderCompaniesTable();
}

function setWorkflowFilter(value) {
  state.workflowFilter = state.workflowFilter === value ? '' : value;
  const sel = document.getElementById('filter-workflow');
  if (sel) sel.value = state.workflowFilter;
  renderStats();
  renderFocusQueue();
  renderSidebar();
  renderCompaniesTable();
}

function populateFilterDropdown() {
  const sel = document.getElementById('filter-status');
  const cur = sel.value;
  sel.innerHTML = '<option value="">All Statuses</option>' +
    state.statuses.map(s => `<option ${s===cur?'selected':''}>${esc(s)}</option>`).join('');
}

function populateWorkflowDropdown() {
  const sel = document.getElementById('filter-workflow');
  if (!sel) return;
  const cur = sel.value;
  sel.innerHTML = '<option value="">All Actions</option>' +
    Object.entries(WORKFLOW).map(([key, cfg]) =>
      `<option value="${key}" ${key===cur?'selected':''}>${esc(cfg.label)}</option>`
    ).join('');
}

// ── Refresh / download ─────────────────────────────────────────────────────
async function refreshAll() {
  const btn = document.getElementById('refresh-btn');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Fetching from Outlook…';
  try {
    await Promise.all([loadCompanies(), loadEmails(true)]);
    renderAll();
    toast('Emails refreshed from Outlook and saved', 'success');
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
