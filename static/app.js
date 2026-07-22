const body = document.querySelector('#forecast-body');
const head = document.querySelector('#head');
const statusEl = document.querySelector('#status');
const retrievedEl = document.querySelector('#retrieved');
const issuedEl = document.querySelector('#issued');
const nextRefreshEl = document.querySelector('#next-refresh');
const adminButton = document.querySelector('#admin-login');
const adminDialog = document.querySelector('#admin-dialog');
const adminForm = document.querySelector('#admin-form');
const adminPassword = document.querySelector('#admin-password');
const adminError = document.querySelector('#admin-error');
const outlookContent = document.querySelector('#outlook-content');
const outlookSource = document.querySelector('#outlook-source');
const riskCount = document.querySelector('#risk-count');
const dateSelect = document.querySelector('#forecast-start-date');
const exportLink = document.querySelector('#export-link');
const REFRESH_MS = 30 * 60 * 1000;
const WINDOW_DAYS = 5;

let nextRefreshAt = Date.now() + REFRESH_MS;
let adminMode = false;
let currentData = null;
let selectedStartDate = '';

function esc(value) {
  return String(value ?? '').replace(/[&<>'"]/g, (char) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;',
  }[char]));
}

function regionSpan(rows, index) {
  if (!rows[index].region) return 0;
  let count = 1;
  for (let row = index + 1; row < rows.length && !rows[row].region; row += 1) count += 1;
  return count;
}

function severityLabel(level) {
  return ({
    green: 'LIGHT',
    yellow: 'MODERATE',
    orange: 'HEAVY',
    red: 'RED OVERRIDE',
    none: 'NO RAIN',
  })[level] || String(level).toUpperCase();
}

function shortDate(label) {
  const parsed = new Date(label);
  if (Number.isNaN(parsed.getTime())) return label;
  const month = String(parsed.getMonth() + 1).padStart(2, '0');
  const day = String(parsed.getDate()).padStart(2, '0');
  return `${month}/${day}`;
}

function severityControl(row, day) {
  const selected = day.severity_override || 'auto';
  const options = [
    ['auto', 'Auto'],
    ['green', 'Green'],
    ['yellow', 'Yellow'],
    ['orange', 'Orange'],
    ['red', 'Red'],
  ].map(([value, label]) => `<option value="${value}"${selected === value ? ' selected' : ''}>${label}</option>`).join('');
  return `<label class="override-control${adminMode ? ' visible' : ''}"><span>Admin severity</span><select data-site="${esc(row.site)}" data-date="${esc(day.date)}">${options}</select></label>`;
}

function weeklyExcerpt(text) {
  if (!text) return 'The model-based outlook is temporarily unavailable.';
  return text.length > 850 ? `${text.slice(0, 850)}…` : text;
}

function availableStartDates(data) {
  const first = data.rows.find((row) => row.days.length)?.days || [];
  return first.slice(0, Math.max(0, first.length - WINDOW_DAYS + 1)).map((day) => day.date);
}

function selectedWindow(data) {
  const starts = availableStartDates(data);
  if (!starts.length) return { start: 0, dates: [] };
  if (!selectedStartDate || !starts.includes(selectedStartDate)) selectedStartDate = starts[0];
  const start = Math.max(0, starts.indexOf(selectedStartDate));
  const first = data.rows.find((row) => row.days.length)?.days || [];
  return { start, dates: first.slice(start, start + WINDOW_DAYS) };
}

function syncDateSelector(data) {
  const starts = availableStartDates(data);
  if (!starts.length) {
    dateSelect.innerHTML = '<option>No dates available</option>';
    dateSelect.disabled = true;
    if (exportLink) exportLink.href = '/api/export';
    return;
  }
  if (!selectedStartDate || !starts.includes(selectedStartDate)) selectedStartDate = starts[0];
  dateSelect.disabled = false;
  dateSelect.innerHTML = starts.map((date) => `<option value="${esc(date)}"${date === selectedStartDate ? ' selected' : ''}>${esc(shortDate(date))}</option>`).join('');
  if (exportLink) {
    const params = new URLSearchParams({ start: selectedStartDate });
    exportLink.href = `/api/export?${params.toString()}`;
  }
}

function render(data) {
  currentData = data;
  syncDateSelector(data);
  const { start, dates } = selectedWindow(data);
  head.innerHTML = '<th>Region</th><th>Site</th>' + dates.map((day) => `<th>${esc(day.date)}</th>`).join('');

  let elevated = 0;
  body.innerHTML = data.rows.map((row, index) => {
    const region = row.region ? `<td class="region" rowspan="${regionSpan(data.rows, index)}">${esc(row.region)}</td>` : '';
    const windowDays = row.days.slice(start, start + WINDOW_DAYS);
    const cells = windowDays.map((day) => {
      if (['yellow', 'orange', 'red'].includes(day.severity) || day.weather_alert) elevated += 1;
      const overlay = day.weather_alert ? `<div class="weather-alert alert-${esc(day.alert_level)}"><span>MODEL HAZARD FLAG</span>⚠ ${esc(day.weather_alert)}<small>${esc(day.overlay_source)}</small></div>` : '';
      const finalBadge = day.severity !== day.base_severity ? `<span class="severity-badge overlay-badge">Display: ${esc(severityLabel(day.severity))}</span>` : '';
      const adminBadge = day.severity_override ? `<span class="severity-badge admin-badge">Admin: ${esc(severityLabel(day.severity_override))}</span>` : '';
      return `<td class="forecast severity-${esc(day.severity)}">${severityControl(row, day)}<div class="badge-row"><span class="severity-badge base-badge">Base: ${esc(severityLabel(day.base_severity || day.automatic_severity))}</span>${finalBadge}${adminBadge}</div><div class="weather-top">${day.icon ? `<img src="${esc(day.icon)}" alt="">` : ''}<span class="condition">${esc(day.condition)}</span></div><div class="metrics"><span class="pill">Rain ${day.rain_chance ?? '—'}%</span></div><div class="forecast-window">${esc(day.narrative)}</div>${overlay}</td>`;
    }).join('');
    return `<tr>${region}<td class="site">${esc(row.site)}<span class="source">Model: Open-Meteo</span></td>${cells || '<td class="error" colspan="5">Forecast unavailable</td>'}</tr>`;
  }).join('');

  issuedEl.textContent = data.issued.replace(/^Issued at:\s*/i, '');
  statusEl.textContent = `Open-Meteo ${data.source_mode}`;
  retrievedEl.textContent = `Retrieved ${new Date(data.retrieved_at).toLocaleString()}`;
  riskCount.textContent = elevated;

  const weekly = data.weekly_outlook || {};
  outlookSource.href = weekly.source_url || 'https://open-meteo.com/';
  outlookContent.innerHTML = `<span class="outlook-meta">${esc(weekly.issued || 'unavailable')}</span>${esc(weeklyExcerpt(weekly.summary))}`;
}

async function load() {
  statusEl.textContent = 'Refreshing…';
  try {
    const response = await fetch('/api/forecast', { cache: 'no-store' });
    if (!response.ok) throw new Error((await response.json()).detail || 'Request failed');
    render(await response.json());
    nextRefreshAt = Date.now() + REFRESH_MS;
  } catch (error) {
    body.innerHTML = `<tr><td class="error" colspan="7">${esc(error.message)}</td></tr>`;
    statusEl.textContent = 'Unable to load';
  }
}

async function checkAdmin() {
  try {
    const response = await fetch('/api/admin/status', { cache: 'no-store' });
    const data = await response.json();
    adminMode = !!data.authenticated;
    adminButton.textContent = adminMode ? 'Admin: signed in' : 'Admin';
    adminButton.classList.toggle('admin-active', adminMode);
  } catch {
    adminMode = false;
  }
}

async function setLevel(site, date, level) {
  const response = await fetch('/api/admin/override', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ site, date, level }),
  });
  if (response.status === 401) {
    adminMode = false;
    await checkAdmin();
    adminDialog.showModal();
    throw new Error('Admin session expired.');
  }
  if (!response.ok) throw new Error((await response.json()).detail || 'Unable to update warning');
  await load();
}

const savedTheme = localStorage.getItem('pagasa-theme') || 'light';
document.documentElement.dataset.theme = savedTheme;
document.querySelector('#theme-toggle').textContent = savedTheme === 'dark' ? '☀' : '☾';

document.querySelector('#theme-toggle').addEventListener('click', (event) => {
  const dark = document.documentElement.dataset.theme !== 'dark';
  document.documentElement.dataset.theme = dark ? 'dark' : 'light';
  localStorage.setItem('pagasa-theme', dark ? 'dark' : 'light');
  event.currentTarget.textContent = dark ? '☀' : '☾';
});

dateSelect.addEventListener('change', (event) => {
  selectedStartDate = event.target.value;
  if (currentData) render(currentData);
});

document.querySelector('#refresh').addEventListener('click', load);
adminButton.addEventListener('click', async () => {
  if (adminMode) {
    await fetch('/api/admin/logout', { method: 'POST' });
    adminMode = false;
    await checkAdmin();
    await load();
  } else {
    adminError.textContent = '';
    adminPassword.value = '';
    adminDialog.showModal();
    adminPassword.focus();
  }
});
document.querySelector('.dialog-close').addEventListener('click', () => adminDialog.close());
adminForm.addEventListener('submit', async (event) => {
  event.preventDefault();
  adminError.textContent = '';
  const response = await fetch('/api/admin/login', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ password: adminPassword.value }),
  });
  if (!response.ok) {
    adminError.textContent = (await response.json()).detail || 'Login failed';
    return;
  }
  adminMode = true;
  adminDialog.close();
  await checkAdmin();
  await load();
});
body.addEventListener('change', async (event) => {
  const select = event.target.closest('.override-control select');
  if (!select || !adminMode) return;
  select.disabled = true;
  try {
    await setLevel(select.dataset.site, select.dataset.date, select.value);
  } catch (error) {
    alert(error.message);
    await load();
  } finally {
    select.disabled = false;
  }
});

setInterval(load, REFRESH_MS);
setInterval(() => {
  const remaining = Math.max(0, nextRefreshAt - Date.now());
  const minutes = Math.floor(remaining / 60000);
  const seconds = Math.floor((remaining % 60000) / 1000);
  nextRefreshEl.textContent = `${minutes}:${String(seconds).padStart(2, '0')}`;
}, 1000);

(async () => {
  await checkAdmin();
  await load();
})();
