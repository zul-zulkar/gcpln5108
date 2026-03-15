// ── Config ────────────────────────────────────────────────────────────────────
function _parseSheetId(url) {
  if (!url) return null;
  const m = url.match(/\/spreadsheets\/d\/([\w-]+)/);
  return m ? m[1] : (url.match(/^[\w-]{20,}$/) ? url : null);
}
const _DEFAULT_SHEET_ID = (typeof CONFIG !== 'undefined')
  ? (_parseSheetId(CONFIG.DEFAULT_SHEET_URL || '') || _parseSheetId(CONFIG.DEFAULT_SHEET_ID || '') || '')
  : '';
let SHEET_ID = _DEFAULT_SHEET_ID;
function _loadSheetConfig() {
  const sid = localStorage.getItem('cfg_sheet_id') || '';
  if (sid) SHEET_ID = sid;
  // clean up legacy GID keys
  ['cfg_gid_utama','cfg_gid_ringkasan','cfg_gid_riwayat'].forEach(k => localStorage.removeItem(k));
}

// All tabs loaded by sheet name — no GID configuration needed
const GVZ_UTAMA     = () => `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=responseHandler:_gsheetUtamaCB&headers=1&sheet=Utama&t=${Date.now()}`;
const GVZ_BASE      = () => `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=responseHandler:_gsheetCB&headers=1&sheet=Riwayat&t=${Date.now()}`;
const GVZ_RINGKASAN = () => `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=responseHandler:_gsheetRingkasanCB&headers=1&sheet=Ringkasan&t=${Date.now()}`;
const GVZ_RIWAYAT   = () => `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=responseHandler:_gsheetRiwayatCB&headers=1&sheet=Riwayat&t=${Date.now()}`;

// ── State ─────────────────────────────────────────────────────────────────────
let ulpMap     = {};   // email.toLowerCase() → ULP name (from Utama tab)
let allData    = [];
let filtered   = [];
let ulpList    = [];   // sorted unique ULPs
let activeUlp  = 'all';
let sortCol    = 'no';
let sortDir    = 'asc';
let chartPasca = null;
let chartPraba = null;
let ringkasanData      = [];
let chartRingkasan     = null;
let riwayatData     = [];   // semua baris dari tab Riwayat
let chartRiwayat    = null;
let riwayatSelectedEmails = new Set(['__all__']); // multi-select state
let riwayatMsQuery    = '';
let riwayatMetric     = 'submit_pasca';
let ringkasanFilter      = { from: '', to: '' };
let riwayatFilter     = { from: '', to: '' };
let ringkasanActiveQuick  = 0;
let riwayatActiveQuick = 0;
let ringkasanTblSortCol   = 'tanggal';
let ringkasanTblSortDir   = 'desc';
let riwayatTblSortCol  = 'tanggal';
let riwayatTblSortDir  = 'desc';
let _msOutsideListenerAttached = false;

// ── JSONP loader ──────────────────────────────────────────────────────────────
function loadData() {
  const btn  = document.getElementById('btnRefresh');
  const icon = document.getElementById('refreshIcon');
  btn.classList.add('loading');
  icon.textContent = '⏳';
  document.getElementById('lastUpdated').textContent = 'Memuat…';
  document.getElementById('errBox').innerHTML = '';

  // Step 1: selalu load Utama (by GID jika dikonfigurasi, by sheet name jika belum)
  const old = document.getElementById('_gsheet_utama_script');
  if (old) old.remove();
  const s = document.createElement('script');
  s.id  = '_gsheet_utama_script';
  s.src = GVZ_UTAMA();
  s.onerror = () => { ulpMap = {}; _loadMainStats(); };
  document.head.appendChild(s);
}

function _loadMainStats() {
  const old = document.getElementById('_gsheet_script');
  if (old) old.remove();
  const script = document.createElement('script');
  script.id  = '_gsheet_script';
  script.src = GVZ_BASE();
  script.onerror = () => {
    document.getElementById('errBox').innerHTML =
      `<div class="err-box">❌ Gagal memuat data. Periksa koneksi internet dan pastikan spreadsheet diset <em>Anyone with the link can view</em>.</div>`;
    document.getElementById('lastUpdated').textContent = 'Gagal';
    const btn = document.getElementById('btnRefresh');
    btn.classList.remove('loading');
    document.getElementById('refreshIcon').textContent = '🔄';
  };
  document.head.appendChild(script);
}

// ── Utama Callback: build email → ULP map ────────────────────────────────────
function _gsheetUtamaCB(data) {
  try {
    ulpMap = {};
    const cols    = data.table.cols || [];
    const allRows = data.table.rows || [];

    // Build colIdx from GViz headers (row 1 of the sheet)
    const colIdx = {};
    cols.forEach((c, i) => { if (c.label) colIdx[c.label.trim().toLowerCase().replace(/_/g, ' ')] = i; });
    const ci = name => colIdx[name.toLowerCase().replace(/_/g, ' ')] ?? -1;

    const EM_KEYS  = ['email','username','user'];
    const ULP_KEYS = ['ulp','nama ulp','kantor cabang','cabang','unit'];
    let iEm  = EM_KEYS.reduce((f, k) => f >= 0 ? f : ci(k), -1);
    let iUlp = ULP_KEYS.reduce((f, k) => f >= 0 ? f : ci(k), -1);

    // Handle 2-row header sheets (e.g. row1: "ULP | PASCABAYAR | PRABAYAR",
    // row2: "No | Nama | Email | ..."). GViz only reads row1 as col labels;
    // row2 becomes the first data row but its cell values are actually column names.
    let dataStart = 0;
    if ((iEm < 0 || iUlp < 0) && allRows.length > 0) {
      const firstCells = allRows[0].c || [];
      const rowIdx = {};
      firstCells.forEach((cell, i) => {
        if (cell && cell.v != null) {
          const key = String(cell.v).trim().toLowerCase().replace(/_/g, ' ');
          if (key) rowIdx[key] = i;
        }
      });
      const ri = name => rowIdx[name.toLowerCase().replace(/_/g, ' ')] ?? -1;
      if (iEm  < 0) iEm  = EM_KEYS.reduce((f, k) => f >= 0 ? f : ri(k), -1);
      if (iUlp < 0) iUlp = ULP_KEYS.reduce((f, k) => f >= 0 ? f : ri(k), -1);
      if (iEm >= 0 || iUlp >= 0) dataStart = 1; // first row was a header row, skip it
    }

    if (iEm >= 0 && iUlp >= 0) {
      allRows.slice(dataStart).forEach(row => {
        const c   = row.c || [];
        const em  = c[iEm]  && c[iEm].v  ? String(c[iEm].v).trim().toLowerCase()  : '';
        const ulp = c[iUlp] && c[iUlp].v ? String(c[iUlp].v).trim() : '';
        if (em && ulp) ulpMap[em] = ulp;
      });
    }
  } catch(e) {
    console.error('Utama CB:', e);
    ulpMap = {};
  }
  _loadMainStats();
}

// ── JSONP Callback ────────────────────────────────────────────────────────────
function _gsheetCB(data) {
  const btn  = document.getElementById('btnRefresh');
  const icon = document.getElementById('refreshIcon');

  try {
    const cols = data.table.cols || [];
    const rows = data.table.rows || [];

    // Build column-name → index map (case-insensitive, underscore = space)
    const colIdx = {};
    cols.forEach((c, i) => {
      if (c.label) colIdx[c.label.trim().toLowerCase().replace(/_/g, ' ')] = i;
    });
    const ci = name => colIdx[name.toLowerCase().replace(/_/g, ' ')] ?? -1;

    // Today in DD/MM/YYYY (same format scraper stores)
    const t = new Date();
    const todayStr = `${String(t.getDate()).padStart(2,'0')}/${String(t.getMonth()+1).padStart(2,'0')}/${t.getFullYear()}`;

    allData = [];
    let seq = 1;
    rows.forEach(row => {
      const c  = row.c || [];
      const cv = idx => idx >= 0 && c[idx] && c[idx].v !== null && c[idx].v !== undefined ? c[idx].v : null;

      // Filter: hanya baris hari ini
      const tglStr = _parseGvizDate(c[ci('tanggal')]);
      if (tglStr !== todayStr) return;

      const nama  = c[ci('nama')]  ? String(c[ci('nama')].v  || '').trim() : '';
      const email = c[ci('email')] ? String(c[ci('email')].v || '').trim() : '';
      if (!nama) return;

      allData.push({
        ulp:          ulpMap[email.toLowerCase()] || '',
        no:           seq++,
        nama_pasca:   nama,
        email_pasca:  email,
        open_pasca:   cv(ci('open pasca')),
        submit_pasca: cv(ci('submit pasca')),
        reject_pasca: cv(ci('reject pasca')),
        nama_praba:   nama,
        email_praba:  email,
        open_praba:   cv(ci('open praba')),
        submit_praba: cv(ci('submit praba')),
        reject_praba: cv(ci('reject praba')),
      });
    });

    activeUlp = 'all';
    render();
    _refreshTimestamp();
  } catch (e) {
    document.getElementById('errBox').innerHTML =
      `<div class="err-box">❌ Gagal memproses data: <strong>${e.message}</strong></div>`;
    document.getElementById('lastUpdated').textContent = 'Error';
  } finally {
    btn.classList.remove('loading');
    icon.textContent = '🔄';
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────
const n = v => (typeof v === 'number' ? v : 0);

function _refreshTimestamp() {
  const el = document.getElementById('lastUpdated');
  if (!el) return;
  if (ringkasanData.length) {
    const last = ringkasanData[ringkasanData.length - 1];
    const ts = last.waktu ? `${last.tanggal} ${last.waktu}` : last.tanggal;
    el.textContent = `Terakhir diperbaharui: ${ts}`;
  } else {
    const now = new Date();
    const ts = now.toLocaleString('id-ID', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    el.textContent = `Dimuat: ${ts}`;
  }
}

function getUlps() {
  const set = new Set(allData.map(d => d.ulp).filter(Boolean));
  return [...set].sort();
}

function getDisplayData() {
  return activeUlp === 'all' ? allData : allData.filter(d => d.ulp === activeUlp);
}

function filterByUlp(idx) {
  activeUlp = idx === 0 ? 'all' : (ulpList[idx - 1] || 'all');
  riwayatSelectedEmails = new Set(['__all__']); // reset pencacah selection when ULP changes
  render();
  buildMsOptions();
  refreshRiwayatViz();
}

// Helper: riwayatData filtered by active ULP
function _ulpRiwayat() {
  return activeUlp === 'all'
    ? riwayatData
    : riwayatData.filter(d => (ulpMap[d.email.toLowerCase()] || '') === activeUlp);
}

function totals(arr, sfx) {
  return arr.reduce((a, d) => ({
    open:   a.open   + n(d[`open_${sfx}`]),
    submit: a.submit + n(d[`submit_${sfx}`]),
    reject: a.reject + n(d[`reject_${sfx}`]),
  }), { open: 0, submit: 0, reject: 0 });
}

function pct(sub, open, rej) {
  const tot = sub + open + rej;
  return tot > 0 ? ((sub / tot) * 100).toFixed(1) : '0.0';
}

function fmt(v) {
  if (v === null || v === undefined) return '<span class="n-zero">–</span>';
  if (typeof v === 'string') {
    if (v === 'ERR')      return '<span class="badge-err">ERR</span>';
    if (v === 'NO_MATCH') return '<span class="badge-nm">N/A</span>';
    return v;
  }
  return v === 0
    ? '<span class="n-zero">0</span>'
    : v.toLocaleString('id-ID');
}

function fmtColored(v, cls) {
  if (v === null || v === undefined) return '<span class="n-zero">–</span>';
  if (typeof v === 'string') {
    if (v === 'ERR')      return '<span class="badge-err">ERR</span>';
    if (v === 'NO_MATCH') return '<span class="badge-nm">N/A</span>';
  }
  if (v === 0) return '<span class="n-zero">0</span>';
  return `<span class="${cls}">${n(v).toLocaleString('id-ID')}</span>`;
}

function toProper(nama) {
  return nama.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
}

function shortName(nama) {
  return toProper(nama).split(' ').slice(0, 2).join(' ');
}

function parseDMY(s) {
  // "DD/MM/YYYY" → "YYYY-MM-DD" for comparison
  if (!s || typeof s !== 'string') return '';
  const p = s.split('/');
  return p.length === 3 ? `${p[2]}-${p[1]}-${p[0]}` : s;
}

function applyDateFilter(data, filter) {
  if (!filter.from && !filter.to) return data;
  return data.filter(d => {
    const ds = parseDMY(d.tanggal);
    if (filter.from && ds < filter.from) return false;
    if (filter.to   && ds > filter.to)   return false;
    return true;
  });
}

// ── Render ────────────────────────────────────────────────────────────────────
function render() {
  ulpList = getUlps();
  const display  = getDisplayData();
  const tPasca   = totals(display, 'pasca');
  const tPraba   = totals(display, 'praba');

  // ULP filter buttons
  const filterBtns = [
    `<button class="ulp-btn${activeUlp === 'all' ? ' active' : ''}" onclick="filterByUlp(0)">Semua ULP</button>`,
    ...ulpList.map((u, i) =>
      `<button class="ulp-btn${activeUlp === u ? ' active' : ''}" onclick="filterByUlp(${i + 1})">${u}</button>`)
  ].join('');

  document.getElementById('content').innerHTML = `
    <!-- ULP Filter Bar -->
    <div class="ulp-filter">${filterBtns}</div>

    <!-- ULP Overview -->
    ${ulpOverviewSection()}

    <!-- Summary Cards -->
    <div class="summary-grid">
      ${summaryCard('pasca', 'PASCABAYAR', tPasca, display.length)}
      ${summaryCard('praba', 'PRABAYAR',   tPraba, display.length)}
    </div>

    <!-- Top Submitters -->
    <div class="top-grid">
      ${topCard('pasca', '🏆 Top Submit – Pascabayar')}
      ${topCard('praba', '🏆 Top Submit – Prabayar')}
    </div>

    <!-- Charts -->
    <div class="chart-wrap">
      <h3>📈 Submit vs Open – Pascabayar</h3>
      <div class="chart-scroll">
        <div class="chart-inner" id="chartWrapPasca"><canvas id="chartPasca"></canvas></div>
      </div>
    </div>
    <div class="chart-wrap">
      <h3>📈 Submit vs Open – Prabayar</h3>
      <div class="chart-scroll">
        <div class="chart-inner" id="chartWrapPraba"><canvas id="chartPraba"></canvas></div>
      </div>
    </div>

    <!-- Table -->
    <div class="tbl-card">
      <div class="tbl-header">
        <h2>Detail Rekap &nbsp;<span id="tblCount"></span></h2>
        <div class="search-box">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
            <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
          </svg>
          <input type="text" id="searchInput" placeholder="Cari nama / email…" oninput="doSearch()" />
        </div>
      </div>
      <div class="tbl-scroll">
        <table>
          <thead>
            <tr>
              <th rowspan="2" class="col-freeze-no" style="vertical-align:middle;text-align:center;cursor:default">No</th>
              <th rowspan="2" onclick="sortBy('nama_pasca')" class="col-nama" style="vertical-align:middle">Nama</th>
              <th colspan="3" class="th-pasca th-span">Pascabayar</th>
              <th colspan="3" class="th-praba th-span">Prabayar</th>
            </tr>
            <tr>
              <th class="th-pasca" onclick="sortBy('open_pasca')">Open</th>
              <th class="th-pasca" onclick="sortBy('submit_pasca')">Submit</th>
              <th class="th-pasca hide-sm" onclick="sortBy('reject_pasca')">Reject</th>
              <th class="th-praba" onclick="sortBy('open_praba')">Open</th>
              <th class="th-praba" onclick="sortBy('submit_praba')">Submit</th>
              <th class="th-praba hide-sm" onclick="sortBy('reject_proba')">Reject</th>
            </tr>
          </thead>
          <tbody id="tblBody"></tbody>
        </table>
      </div>
    </div>
  `;

  renderTable();
  renderChart();
  _animateStatNumbers();
}

function ulpOverviewSection() {
  if (!ulpList.length) return '';

  const cards = ulpList.map((u, idx) => {
    const uData = allData.filter(d => d.ulp === u);
    const tp  = totals(uData, 'pasca');
    const tpr = totals(uData, 'praba');
    const rp  = pct(tp.submit,  tp.open,  tp.reject);
    const rr  = pct(tpr.submit, tpr.open, tpr.reject);
    const isAct = activeUlp === u;
    return `
      <div class="ulp-row${isAct ? ' ulp-active' : ''}" style="--i:${idx}" onclick="filterByUlp(${idx + 1})">
        <div class="ulp-row-top">
          <span class="ulp-row-name">${u}</span>
          <span class="ulp-row-count">${uData.length} petugas</span>
        </div>
        <div class="ulp-prog-row">
          <span class="ulp-badge pasca">Pasca</span>
          <div class="ulp-prog-track"><div class="ulp-prog-fill-p" style="width:${rp}%"></div></div>
          <span class="ulp-prog-pct">${rp}%</span>
          <span class="ulp-prog-val">${tp.submit.toLocaleString('id-ID')} sub</span>
        </div>
        <div class="ulp-prog-row">
          <span class="ulp-badge praba">Praba</span>
          <div class="ulp-prog-track"><div class="ulp-prog-fill-r" style="width:${rr}%"></div></div>
          <span class="ulp-prog-pct">${rr}%</span>
          <span class="ulp-prog-val">${tpr.submit.toLocaleString('id-ID')} sub</span>
        </div>
      </div>`;
  }).join('');

  return `
    <div class="ulp-overview">
      <div class="ulp-overview-header">📍 Progres per ULP</div>
      <div class="ulp-grid">${cards}</div>
    </div>`;
}

function summaryCard(sfx, label, t, count) {
  const rate = pct(t.submit, t.open, t.reject);
  const cls  = sfx === 'pasca' ? 'pasca' : 'praba';
  return `
    <div class="s-card ${cls}">
      <div class="s-card-header">
        <span class="s-badge ${cls}">${label}</span>
        <span class="s-count">${count} petugas${activeUlp !== 'all' ? ' · ' + activeUlp : ''}</span>
      </div>
      <div class="s-stats">
        <div class="s-stat open">
          <div class="s-stat-num">${t.open.toLocaleString('id-ID')}</div>
          <div class="s-stat-lbl">Open</div>
        </div>
        <div class="s-stat submit">
          <div class="s-stat-num">${t.submit.toLocaleString('id-ID')}</div>
          <div class="s-stat-lbl">Submitted</div>
        </div>
        <div class="s-stat reject">
          <div class="s-stat-num">${t.reject.toLocaleString('id-ID')}</div>
          <div class="s-stat-lbl">Rejected</div>
        </div>
      </div>
      <div class="prog-section">
        <div class="prog-meta">
          <span>Tingkat Submit</span>
          <strong>${rate}%</strong>
        </div>
        <div class="prog-track">
          <div class="prog-fill" style="width:${rate}%"></div>
        </div>
      </div>
    </div>`;
}

function topCard(sfx, title) {
  const top5 = [...getDisplayData()]
    .filter(d => typeof d[`submit_${sfx}`] === 'number')
    .sort((a, b) => n(b[`submit_${sfx}`]) - n(a[`submit_${sfx}`]))
    .slice(0, 5);
  const medals = ['🥇','🥈','🥉','4️⃣','5️⃣'];
  const rankCls = ['gold','silver','bronze','',''];
  const rows = top5.map((d, i) => `
    <div class="top-row">
      <span class="top-rank ${rankCls[i]}">${medals[i]}</span>
      <span class="top-nama" title="${toProper(d[`nama_${sfx}`] || d.nama_pasca)}">${toProper(d[`nama_${sfx}`] || d.nama_pasca)}</span>
      <span class="top-val">${n(d[`submit_${sfx}`]).toLocaleString('id-ID')}</span>
    </div>`).join('');
  return `
    <div class="top-card">
      <h3>${title}</h3>
      ${rows || '<p style="color:var(--muted);font-size:.8rem">Belum ada data</p>'}
    </div>`;
}

function renderTable() {
  const q    = (document.getElementById('searchInput')?.value || '').toLowerCase();
  const base = getDisplayData();
  filtered   = base.filter(d => {
    const nama  = (d.nama_pasca || d.nama_praba).toLowerCase();
    const email = (d.email_pasca || d.email_praba).toLowerCase();
    return !q || nama.includes(q) || email.includes(q);
  });

  // Sort
  filtered.sort((a, b) => {
    const va = a[sortCol], vb = b[sortCol];
    if (typeof va === 'number' && typeof vb === 'number')
      return sortDir === 'asc' ? va - vb : vb - va;
    const sa = String(va ?? ''), sb = String(vb ?? '');
    return sortDir === 'asc' ? sa.localeCompare(sb, 'id') : sb.localeCompare(sa, 'id');
  });

  const count = document.getElementById('tblCount');
  if (count) count.textContent = `(${filtered.length} dari ${base.length})`;

  const body = document.getElementById('tblBody');
  if (!body) return;

  if (!filtered.length) {
    body.innerHTML = `<tr><td colspan="8" style="text-align:center;padding:2.5rem;color:var(--muted)">Tidak ditemukan</td></tr>`;
    return;
  }

  body.innerHTML = filtered.map((d, i) => {
    const namaRaw = d.nama_pasca || d.nama_praba;
    const nama    = toProper(namaRaw);
    return `<tr>
      <td class="td-no col-freeze-no">${i + 1}</td>
      <td class="td-nama col-nama" title="${nama}">${nama}</td>
      <td class="td-pasca">${fmtColored(d.open_pasca,   'n-open')}</td>
      <td class="td-pasca">${fmtColored(d.submit_pasca, 'n-submit')}</td>
      <td class="td-pasca hide-sm">${fmtColored(d.reject_pasca, 'n-reject')}</td>
      <td class="td-praba">${fmtColored(d.open_praba,   'n-open')}</td>
      <td class="td-praba">${fmtColored(d.submit_praba, 'n-submit')}</td>
      <td class="td-praba hide-sm">${fmtColored(d.reject_praba, 'n-reject')}</td>
    </tr>`;
  }).join('');

  // Highlight sort column header
  document.querySelectorAll('th').forEach(th => th.classList.remove('sort-asc', 'sort-desc'));
  document.querySelectorAll('th').forEach(th => {
    if (th.getAttribute('onclick') === `sortBy('${sortCol}')`)
      th.classList.add(sortDir === 'asc' ? 'sort-asc' : 'sort-desc');
  });
}

function renderChart() {
  renderOneChart('pasca');
  renderOneChart('praba');
}

function renderOneChart(sfx) {
  const isPasca = sfx === 'pasca';
  const ctx  = document.getElementById(isPasca ? 'chartPasca' : 'chartPraba');
  const wrap = document.getElementById(isPasca ? 'chartWrapPasca' : 'chartWrapPraba');
  if (!ctx || !wrap) return;

  if (isPasca) { if (chartPasca) { chartPasca.destroy(); chartPasca = null; } }
  else         { if (chartPraba) { chartPraba.destroy(); chartPraba = null; } }

  const sorted = [...getDisplayData()]
    .filter(d => typeof d[`submit_${sfx}`] === 'number' || d[`submit_${sfx}`] === 0)
    .sort((a, b) => n(b[`submit_${sfx}`]) - n(a[`submit_${sfx}`]));

  const chartW = Math.max(480, sorted.length * 90);
  wrap.style.width = chartW + 'px';

  const labels    = sorted.map(d => shortName(d[`nama_${sfx}`] || d.nama_pasca || d.nama_praba));
  const fullNames = sorted.map(d => toProper(d[`nama_${sfx}`] || d.nama_pasca || d.nama_praba));
  const rgb    = isPasca ? '59,130,246' : '139,92,246';
  const title  = isPasca ? 'Pascabayar' : 'Prabayar';

  ctx.width  = chartW;
  ctx.height = 320;

  const inst = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {
          label: `Submit ${title}`,
          data: sorted.map(d => n(d[`submit_${sfx}`])),
          backgroundColor: `rgba(${rgb},0.85)`,
          borderRadius: 4,
        },
        {
          label: `Open ${title}`,
          data: sorted.map(d => n(d[`open_${sfx}`])),
          backgroundColor: `rgba(${rgb},0.25)`,
          borderRadius: 4,
        },
      ],
    },
    options: {
      responsive: false,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { font: { size: 11 }, padding: 14 } },
        tooltip: {
          callbacks: {
            title: items => fullNames[items[0].dataIndex],
            label: c => ` ${c.dataset.label}: ${c.parsed.y.toLocaleString('id-ID')}`
          }
        }
      },
      scales: {
        x: { ticks: { font: { size: 10 }, maxRotation: 45 } },
        y: { beginAtZero: true, ticks: { font: { size: 10 } } },
      },
    },
  });

  if (isPasca) chartPasca = inst;
  else         chartPraba = inst;
}

// ── Laporan Ringkasan ────────────────────────────────────────────────────────────
function loadRingkasan() {
  const old = document.getElementById('_gsheet_ringkasan_script');
  if (old) old.remove();
  const s = document.createElement('script');
  s.id  = '_gsheet_ringkasan_script';
  s.src = GVZ_RINGKASAN();
  document.head.appendChild(s);
}

function _parseGvizDate(cell) {
  if (!cell) return '';
  // Prefer formatted value (.f) — e.g. "12/03/2026" as stored in sheet
  if (cell.f) return cell.f;
  const v = cell.v;
  if (v === null || v === undefined) return '';
  // GViz returns date cells as "Date(year,month0,day)"
  if (typeof v === 'string') {
    const m = v.match(/^Date\((\d+),(\d+),(\d+)\)/);
    if (m) return `${String(+m[3]).padStart(2,'0')}/${String(+m[2]+1).padStart(2,'0')}/${m[1]}`;
  }
  return String(v);
}

function _parseGvizTime(cell) {
  if (!cell) return '';
  if (cell.f) return cell.f.slice(0, 5);   // "HH:MM:SS" → "HH:MM"
  const v = cell.v;
  if (v === null || v === undefined) return '';
  // GViz returns time as fraction of day
  if (typeof v === 'number') {
    const totalMin = Math.round(v * 1440);
    const h = Math.floor(totalMin / 60), m2 = totalMin % 60;
    return `${String(h).padStart(2,'0')}:${String(m2).padStart(2,'0')}`;
  }
  return String(v).slice(0, 5);
}

function _gsheetRingkasanCB(data) {
  try {
    ringkasanData = [];
    const cols = data.table.cols || [];
    const colIdx = {};
    cols.forEach((c, i) => { if (c.label) colIdx[c.label.trim().toLowerCase().replace(/_/g, ' ')] = i; });
    const ci = name => colIdx[name.toLowerCase().replace(/_/g, ' ')] ?? -1;
    (data.table.rows || []).forEach(row => {
      const c  = row.c || [];
      const cv = i => (i >= 0 && c[i] && c[i].v !== null && c[i].v !== undefined) ? c[i].v : null;
      const iTgl = ci('tanggal') >= 0 ? ci('tanggal') : 0;
      const iWkt = ci('waktu')   >= 0 ? ci('waktu')   : 1;
      if (!cv(iTgl)) return;
      ringkasanData.push({
        tanggal:      _parseGvizDate(c[iTgl]),
        waktu:        _parseGvizTime(c[iWkt]),
        open_pasca:   cv(ci('open pasca')),
        submit_pasca: cv(ci('submit pasca')),
        reject_pasca: cv(ci('reject pasca')),
        open_praba:   cv(ci('open praba')),
        submit_praba: cv(ci('submit praba')),
        reject_praba: cv(ci('reject praba')),
      });
    });
    _refreshTimestamp();
    renderRingkasan();
  } catch(e) { console.error('Ringkasan CB:', e); }
}

// ── Table sort helpers ─────────────────────────────────────────────────────────
function _sortArr(arr, col, dir) {
  return [...arr].sort((a, b) => {
    let cmp;
    if (col === 'tanggal') {
      cmp = parseDMY(a.tanggal).localeCompare(parseDMY(b.tanggal));
    } else if (col === 'nama') {
      cmp = (a.nama || a.email || '').toLowerCase().localeCompare((b.nama || b.email || '').toLowerCase());
    } else {
      cmp = n(a[col]) - n(b[col]);
    }
    return dir === 'asc' ? cmp : -cmp;
  });
}

function sortRingkasanTbl(col) {
  ringkasanTblSortDir = ringkasanTblSortCol === col && ringkasanTblSortDir === 'desc' ? 'asc' : 'desc';
  ringkasanTblSortCol = col;
  renderRingkasan();
}

function sortRiwayatTbl(col) {
  riwayatTblSortDir = riwayatTblSortCol === col && riwayatTblSortDir === 'desc' ? 'asc' : 'desc';
  riwayatTblSortCol = col;
  refreshRiwayatViz();
}

function _syncRiwayatTblHeaders() {
  document.querySelectorAll('[data-rv-sort]').forEach(span => {
    const col = span.dataset.rvSort;
    span.textContent = riwayatTblSortCol === col
      ? (riwayatTblSortDir === 'asc' ? ' ▲' : ' ▼') : '';
  });
}

function renderRingkasan() {
  const sec = document.getElementById('ringkasanSection');
  if (!ringkasanData.length) {
    sec.innerHTML = '<div class="chart-wrap" style="margin-top:1.25rem"><h3>📅 Tren Ringkasan</h3><p style="color:var(--muted);font-size:.85rem;padding:.5rem 0">Belum ada data ringkasan.</p></div>';
    return;
  }

  // Apply date filter
  const filt        = applyDateFilter(ringkasanData, ringkasanFilter);
  // Chart: newest first (left)
  const display     = [...filt].reverse();
  const labels      = display.map(d => d.tanggal);
  const submitPasca = display.map(d => n(d.submit_pasca));
  const submitPraba = display.map(d => n(d.submit_praba));
  const openPasca   = display.map(d => n(d.open_pasca));
  const openPraba   = display.map(d => n(d.open_praba));

  // Table: sortable
  const arrH    = col => ringkasanTblSortCol === col ? (ringkasanTblSortDir === 'asc' ? ' ▲' : ' ▼') : '';
  const tblRows = _sortArr(filt, ringkasanTblSortCol, ringkasanTblSortDir).map(d => `
    <tr>
      <td>${d.tanggal}</td>
      <td class="td-pasca"><span class="n-open">${n(d.open_pasca).toLocaleString('id-ID')}</span></td>
      <td class="td-pasca"><span class="n-submit">${n(d.submit_pasca).toLocaleString('id-ID')}</span></td>
      <td class="td-pasca"><span class="n-reject">${n(d.reject_pasca).toLocaleString('id-ID')}</span></td>
      <td class="td-praba"><span class="n-open">${n(d.open_praba).toLocaleString('id-ID')}</span></td>
      <td class="td-praba"><span class="n-submit">${n(d.submit_praba).toLocaleString('id-ID')}</span></td>
      <td class="td-praba"><span class="n-reject">${n(d.reject_praba).toLocaleString('id-ID')}</span></td>
    </tr>`).join('');

  const minWH = Math.max(480, display.length * 90);
  const mkQ   = (d, lbl) =>
    `<button class="dfq-btn${ringkasanActiveQuick === d ? ' active' : ''}" onclick="setRingkasanQuickFilter(${d})">${lbl}</button>`;

  sec.innerHTML = `
    <div class="chart-wrap" style="margin-top:1.25rem">
      <h3>📅 Tren Ringkasan</h3>
      <div class="filter-row">
        <div class="date-filter-bar">
          ${mkQ(0,'Semua')}${mkQ(7,'7H')}${mkQ(14,'14H')}${mkQ(30,'30H')}
          <span class="dfq-sep">|</span>
          <input type="date" class="dfq-input" id="ringkasanFrom" value="${ringkasanFilter.from}"
            onchange="setRingkasanDateRange(this.value,document.getElementById('ringkasanTo').value)">
          <span class="dfq-sep">–</span>
          <input type="date" class="dfq-input" id="ringkasanTo" value="${ringkasanFilter.to}"
            onchange="setRingkasanDateRange(document.getElementById('ringkasanFrom').value,this.value)">
        </div>
        <button class="dfq-btn" onclick="resetRingkasanFilter()" title="Reset filter tanggal" style="margin-left:.25rem">↺ Reset</button>
      </div>
      <div class="chart-scroll">
        <div class="chart-inner" style="min-width:${minWH}px;width:100%;height:300px">
          <canvas id="chartRingkasan"></canvas>
        </div>
      </div>
    </div>
    <div class="tbl-card" style="margin-bottom:1.25rem">
      <div class="tbl-header"><h2>Riwayat Ringkasan</h2></div>
      <div class="tbl-scroll" style="max-height:280px">
        <table>
          <thead>
            <tr>
              <th rowspan="2" onclick="sortRingkasanTbl('tanggal')" style="vertical-align:middle;cursor:pointer">Tanggal${arrH('tanggal')}</th>
              <th colspan="3" class="th-pasca th-span">Pascabayar</th>
              <th colspan="3" class="th-praba th-span">Prabayar</th>
            </tr>
            <tr>
              <th class="th-pasca" onclick="sortRingkasanTbl('open_pasca')" style="cursor:pointer">Open${arrH('open_pasca')}</th>
              <th class="th-pasca" onclick="sortRingkasanTbl('submit_pasca')" style="cursor:pointer">Submit${arrH('submit_pasca')}</th>
              <th class="th-pasca" onclick="sortRingkasanTbl('reject_pasca')" style="cursor:pointer">Reject${arrH('reject_pasca')}</th>
              <th class="th-praba" onclick="sortRingkasanTbl('open_praba')" style="cursor:pointer">Open${arrH('open_praba')}</th>
              <th class="th-praba" onclick="sortRingkasanTbl('submit_praba')" style="cursor:pointer">Submit${arrH('submit_praba')}</th>
              <th class="th-praba" onclick="sortRingkasanTbl('reject_praba')" style="cursor:pointer">Reject${arrH('reject_praba')}</th>
            </tr>
          </thead>
          <tbody>${tblRows || '<tr><td colspan="7" style="text-align:center;color:var(--muted);padding:1.5rem">Belum ada data</td></tr>'}</tbody>
        </table>
      </div>
    </div>`;

  if (chartRingkasan) { chartRingkasan.destroy(); chartRingkasan = null; }
  chartRingkasan = new Chart(document.getElementById('chartRingkasan'), {
    type: 'line',
    data: {
      labels,
      datasets: [
        { label: 'Submit Pascabayar', data: submitPasca, borderColor: '#3B82F6', backgroundColor: 'rgba(59,130,246,.1)',  tension: .3, pointRadius: 4, fill: true },
        { label: 'Submit Prabayar',   data: submitPraba, borderColor: '#8B5CF6', backgroundColor: 'rgba(139,92,246,.1)', tension: .3, pointRadius: 4, fill: true },
        { label: 'Open Pascabayar',   data: openPasca,   borderColor: '#93C5FD', borderDash: [4,3], tension: .3, pointRadius: 3, fill: false },
        { label: 'Open Prabayar',     data: openPraba,   borderColor: '#C4B5FD', borderDash: [4,3], tension: .3, pointRadius: 3, fill: false },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { font: { size: 11 }, padding: 14 } },
        tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${c.parsed.y.toLocaleString('id-ID')}` } },
      },
      scales: {
        x: { ticks: { font: { size: 10 }, maxRotation: 30 } },
        y: { beginAtZero: true, ticks: { font: { size: 10 } } },
      },
    },
  });
}

// ── Progres Ringkasan per Pencacah ───────────────────────────────────────────────
function loadRiwayat() {
  const old = document.getElementById('_gsheet_riwayat_script');
  if (old) old.remove();
  const s = document.createElement('script');
  s.id  = '_gsheet_riwayat_script';
  s.src = GVZ_RIWAYAT();
  document.head.appendChild(s);
}

function _gsheetRiwayatCB(data) {
  try {
    riwayatData = [];
    const cols = data.table.cols || [];
    const colIdx = {};
    cols.forEach((c, i) => { if (c.label) colIdx[c.label.trim().toLowerCase().replace(/_/g, ' ')] = i; });
    const ci = name => colIdx[name.toLowerCase().replace(/_/g, ' ')] ?? -1;
    (data.table.rows || []).forEach(row => {
      const c  = row.c || [];
      const cv = i => (i >= 0 && c[i] && c[i].v !== null && c[i].v !== undefined) ? c[i].v : null;
      const iTgl = ci('tanggal') >= 0 ? ci('tanggal') : 0;
      const iEm  = ci('email')   >= 0 ? ci('email')   : 2;
      const iNm  = ci('nama')    >= 0 ? ci('nama')    : 1;
      if (!cv(iTgl) && !cv(iEm)) return;
      riwayatData.push({
        tanggal:      _parseGvizDate(c[iTgl]),
        email:        cv(iEm) ? String(cv(iEm)).trim() : '',
        nama:         cv(iNm) ? String(cv(iNm)).trim() : '',
        open_pasca:   cv(ci('open pasca')),
        submit_pasca: cv(ci('submit pasca')),
        reject_pasca: cv(ci('reject pasca')),
        open_praba:   cv(ci('open praba')),
        submit_praba: cv(ci('submit praba')),
        reject_praba: cv(ci('reject praba')),
      });
    });
    renderRiwayat();
  } catch(e) { console.error('Riwayat CB:', e); }
}

function renderRiwayat() {
  const sec = document.getElementById('riwayatSection');
  if (!riwayatData.length) {
    sec.innerHTML = '<div class="chart-wrap" style="margin-top:1.25rem"><h3>👤 Progres Ringkasan per Pencacah</h3><p style="color:var(--muted);font-size:.85rem;padding:.5rem 0">Belum ada data riwayat per pencacah.</p></div>';
    return;
  }

  const metricLabels = {
    submit_pasca: 'Submit Pascabayar', submit_praba: 'Submit Prabayar',
    open_pasca:   'Open Pascabayar',   open_praba:   'Open Prabayar',
  };
  const metricOptHtml = Object.entries(metricLabels).map(([v, l]) =>
    `<option value="${v}"${v === riwayatMetric ? ' selected' : ''}>${l}</option>`
  ).join('');

  const mkQR = (d, lbl) =>
    `<button class="dfq-btn${riwayatActiveQuick === d ? ' active' : ''}" onclick="setRiwayatQuickFilter(${d})">${lbl}</button>`;

  sec.innerHTML = `
    <div class="chart-wrap" style="margin-top:1.25rem">
      <h3 style="display:flex;align-items:center;gap:.6rem;flex-wrap:wrap">
        👤 Progres Ringkasan per Pencacah
        <select onchange="riwayatMetric=this.value;refreshRiwayatViz()"
          style="font-size:.82rem;padding:.28rem .55rem;border:1px solid var(--border);border-radius:6px;background:#fff;color:var(--text);cursor:pointer">
          ${metricOptHtml}
        </select>
      </h3>
      <div class="filter-row" style="margin-top:.6rem">
        <div class="date-filter-bar">
          ${mkQR(0,'Semua')}${mkQR(7,'7H')}${mkQR(14,'14H')}${mkQR(30,'30H')}
          <span class="dfq-sep">|</span>
          <input type="date" class="dfq-input" id="riwayatFrom" value="${riwayatFilter.from}"
            onchange="setRiwayatDateRange(this.value,document.getElementById('riwayatTo').value)">
          <span class="dfq-sep">–</span>
          <input type="date" class="dfq-input" id="riwayatTo" value="${riwayatFilter.to}"
            onchange="setRiwayatDateRange(document.getElementById('riwayatFrom').value,this.value)">
        </div>
        <button class="dfq-btn" onclick="resetRiwayatFilter()" title="Reset semua filter" style="margin-left:.25rem">↺ Reset</button>
        <div class="ms-wrap" id="msPencacahWrap">
          <button class="ms-trigger" onclick="toggleMsPencacah(event)">
            <span id="msRiwayatLabel">${getMsLabel()}</span>
            <span class="ms-caret">▾</span>
          </button>
          <div class="ms-panel" id="msRiwayatPanel" style="display:none">
            <div class="ms-search-row">
              <input id="msRiwayatSearch" type="text" placeholder="Cari nama petugas…"
                oninput="filterMsPencacah()" autocomplete="off">
            </div>
            <div class="ms-actions">
              <button class="ms-act-btn" onclick="selectAllPencacah()">Pilih Semua</button>
              <button class="ms-act-btn" onclick="clearPencacah()">Hapus Semua</button>
            </div>
            <div class="ms-list" id="msRiwayatList"></div>
          </div>
        </div>
      </div>
      <div class="chart-scroll" style="margin-top:.75rem">
        <div class="chart-inner" id="riwayatChartInner" style="width:100%;height:380px">
          <canvas id="chartRiwayat"></canvas>
        </div>
      </div>
    </div>
    <div class="tbl-card" style="margin-bottom:1.25rem">
      <div class="tbl-header"><h2>Detail Pencacah</h2></div>
      <div class="tbl-scroll" style="max-height:260px">
        <table>
          <thead>
            <tr>
              <th rowspan="2" onclick="sortRiwayatTbl('tanggal')" style="vertical-align:middle;cursor:pointer">Tanggal<span data-rv-sort="tanggal"></span></th>
              <th rowspan="2" onclick="sortRiwayatTbl('nama')" style="vertical-align:middle;cursor:pointer">Pencacah<span data-rv-sort="nama"></span></th>
              <th colspan="3" class="th-pasca th-span">Pascabayar</th>
              <th colspan="3" class="th-praba th-span">Prabayar</th>
            </tr>
            <tr>
              <th class="th-pasca" onclick="sortRiwayatTbl('open_pasca')" style="cursor:pointer">Open<span data-rv-sort="open_pasca"></span></th>
              <th class="th-pasca" onclick="sortRiwayatTbl('submit_pasca')" style="cursor:pointer">Submit<span data-rv-sort="submit_pasca"></span></th>
              <th class="th-pasca" onclick="sortRiwayatTbl('reject_pasca')" style="cursor:pointer">Reject<span data-rv-sort="reject_pasca"></span></th>
              <th class="th-praba" onclick="sortRiwayatTbl('open_praba')" style="cursor:pointer">Open<span data-rv-sort="open_praba"></span></th>
              <th class="th-praba" onclick="sortRiwayatTbl('submit_praba')" style="cursor:pointer">Submit<span data-rv-sort="submit_praba"></span></th>
              <th class="th-praba" onclick="sortRiwayatTbl('reject_praba')" style="cursor:pointer">Reject<span data-rv-sort="reject_praba"></span></th>
            </tr>
          </thead>
          <tbody id="riwayatTblBody"></tbody>
        </table>
      </div>
    </div>`;

  attachMsOutsideListener();
  buildMsOptions();
  refreshRiwayatViz();
}

// ── Reset ─────────────────────────────────────────────────────────────────────
function resetRingkasanFilter() {
  ringkasanFilter      = { from: '', to: '' };
  ringkasanActiveQuick = 0;
  renderRingkasan();
}

function resetRiwayatFilter() {
  riwayatFilter         = { from: '', to: '' };
  riwayatActiveQuick    = 0;
  riwayatSelectedEmails = new Set(['__all__']);
  riwayatMsQuery        = '';
  buildMsOptions();
  refreshRiwayatViz();
}

// ── Date filter helpers ────────────────────────────────────────────────────────
function _localISO(d) {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

function _quickDate(days) {
  const d = new Date();
  d.setDate(d.getDate() - days + 1);
  return _localISO(d);
}

function setRingkasanQuickFilter(days) {
  ringkasanActiveQuick = days;
  ringkasanFilter = days === 0
    ? { from: '', to: '' }
    : { from: _quickDate(days), to: _localISO(new Date()) };
  renderRingkasan();
}

function setRingkasanDateRange(from, to) {
  ringkasanActiveQuick = -1;
  ringkasanFilter = { from, to };
  renderRingkasan();
}

function setRiwayatQuickFilter(days) {
  riwayatActiveQuick = days;
  riwayatFilter = days === 0
    ? { from: '', to: '' }
    : { from: _quickDate(days), to: _localISO(new Date()) };
  refreshRiwayatViz();
  // Sync date inputs
  const fromEl = document.getElementById('riwayatFrom');
  const toEl   = document.getElementById('riwayatTo');
  if (fromEl) fromEl.value = riwayatFilter.from;
  if (toEl)   toEl.value   = riwayatFilter.to;
  // Update quick button active state
  document.querySelectorAll('[onclick^="setRiwayatQuickFilter"]').forEach(btn => {
    const d = parseInt(btn.getAttribute('onclick').match(/\d+/)?.[0] ?? '-1');
    btn.classList.toggle('active', d === riwayatActiveQuick);
  });
}

function setRiwayatDateRange(from, to) {
  riwayatActiveQuick = -1;
  riwayatFilter = { from, to };
  refreshRiwayatViz();
}

// ── Multi-select helpers ───────────────────────────────────────────────────────
function getMsLabel() {
  if (riwayatSelectedEmails.has('__all__')) return 'Semua Petugas';
  const cnt = riwayatSelectedEmails.size;
  if (cnt === 0) return '— Pilih Petugas —';
  const total = new Set(_ulpRiwayat().map(d => d.email).filter(Boolean)).size;
  return cnt >= total ? 'Semua Petugas' : `${cnt} Petugas dipilih`;
}

function updateMsLabel() {
  const lbl = document.getElementById('msRiwayatLabel');
  if (lbl) lbl.textContent = getMsLabel();
}

function toggleMsPencacah(e) {
  if (e) e.stopPropagation();
  const panel = document.getElementById('msRiwayatPanel');
  if (!panel) return;
  const isOpen = panel.style.display !== 'none';
  panel.style.display = isOpen ? 'none' : 'block';
  if (!isOpen) {
    const inp = document.getElementById('msRiwayatSearch');
    if (inp) { inp.value = riwayatMsQuery; inp.focus(); }
    buildMsOptions();
  }
}

function closeMsPencacah() {
  const panel = document.getElementById('msRiwayatPanel');
  if (panel) panel.style.display = 'none';
}

function attachMsOutsideListener() {
  if (_msOutsideListenerAttached) return;
  _msOutsideListenerAttached = true;
  document.addEventListener('click', e => {
    const wrap = document.getElementById('msPencacahWrap');
    if (wrap && !wrap.contains(e.target)) closeMsPencacah();
  });
}

function filterMsPencacah() {
  riwayatMsQuery = document.getElementById('msRiwayatSearch')?.value || '';
  buildMsOptions();
}

function buildMsOptions() {
  const list = document.getElementById('msRiwayatList');
  if (!list) return;
  const q         = riwayatMsQuery.toLowerCase();
  const allEmails = [...new Set(_ulpRiwayat().map(d => d.email).filter(Boolean))].sort();
  const vis = q
    ? allEmails.filter(e => {
        const nama = (riwayatData.find(d => d.email === e)?.nama || '').toLowerCase();
        return nama.includes(q) || e.toLowerCase().includes(q);
      })
    : allEmails;
  const isAll = riwayatSelectedEmails.has('__all__');
  list.innerHTML = vis.map(e => {
    const nama = toProper(riwayatData.find(d => d.email === e)?.nama || e);
    const chk  = (isAll || riwayatSelectedEmails.has(e)) ? ' checked' : '';
    const esc  = e.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
    return `<label class="ms-item"><input type="checkbox"${chk} onchange="togglePencacah('${esc}')"> <span>${nama}</span></label>`;
  }).join('');
  updateMsLabel();
}

function togglePencacah(email) {
  const allEmails = [...new Set(_ulpRiwayat().map(d => d.email).filter(Boolean))];
  const isAll = riwayatSelectedEmails.has('__all__');
  if (isAll) {
    riwayatSelectedEmails = new Set(allEmails);
    riwayatSelectedEmails.delete(email);
    if (riwayatSelectedEmails.size === 0) riwayatSelectedEmails = new Set(['__all__']);
  } else if (riwayatSelectedEmails.has(email)) {
    riwayatSelectedEmails.delete(email);
    if (riwayatSelectedEmails.size === 0) riwayatSelectedEmails = new Set(['__all__']);
  } else {
    riwayatSelectedEmails.add(email);
    if (riwayatSelectedEmails.size >= allEmails.length) riwayatSelectedEmails = new Set(['__all__']);
  }
  buildMsOptions();
  refreshRiwayatViz();
}

function selectAllPencacah() {
  riwayatSelectedEmails = new Set(['__all__']);
  buildMsOptions();
  refreshRiwayatViz();
}

function clearPencacah() {
  riwayatSelectedEmails = new Set();
  buildMsOptions();
  refreshRiwayatViz();
}

// ── Refresh riwayat chart + table (no full DOM rebuild) ───────────────────────
function refreshRiwayatViz() {
  const filteredData = applyDateFilter(_ulpRiwayat(), riwayatFilter);
  const allEmails    = [...new Set(filteredData.map(d => d.email).filter(Boolean))].sort();
  // Latest dates on the LEFT
  const allDates     = [...new Set(filteredData.map(d => d.tanggal))].sort((a, b) => parseDMY(a).localeCompare(parseDMY(b))).reverse();

  const isAll = riwayatSelectedEmails.has('__all__') || riwayatSelectedEmails.size === 0;
  const activeEmails = isAll ? allEmails : allEmails.filter(e => riwayatSelectedEmails.has(e));

  const datasets = activeEmails.map((email, i) => {
    const nama   = filteredData.find(d => d.email === email)?.nama || email;
    const byDate = {};
    filteredData.filter(d => d.email === email).forEach(d => { byDate[d.tanggal] = d[riwayatMetric]; });
    const data = allDates.map(dt => n(byDate[dt]));
    const hue = Math.round(i * 360 / Math.max(1, activeEmails.length));
    return {
      label:           toProper(nama),
      data,
      borderColor:     `hsl(${hue},65%,45%)`,
      backgroundColor: 'transparent',
      borderWidth:     activeEmails.length === 1 ? 2.5 : 1.5,
      pointRadius:     activeEmails.length === 1 ? 4 : 2,
      tension:         .3,
      fill:            false,
    };
  });

  const chartW = Math.max(520, allDates.length * 110);
  const inner  = document.getElementById('riwayatChartInner');
  if (inner) inner.style.minWidth = chartW + 'px';

  if (chartRiwayat) { chartRiwayat.destroy(); chartRiwayat = null; }
  const ctx = document.getElementById('chartRiwayat');
  if (ctx) {
    chartRiwayat = new Chart(ctx, {
      type: 'line',
      data: { labels: allDates, datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${c.parsed.y.toLocaleString('id-ID')}` } },
        },
        scales: {
          x: { ticks: { font: { size: 10 }, maxRotation: 30 } },
          y: { beginAtZero: true, ticks: { font: { size: 10 } } },
        },
      },
    });
  }

  // Update table
  const tblSrc = _sortArr(
    isAll ? filteredData : filteredData.filter(d => riwayatSelectedEmails.has(d.email)),
    riwayatTblSortCol, riwayatTblSortDir);

  const tbody = document.getElementById('riwayatTblBody');
  if (tbody) {
    tbody.innerHTML = tblSrc.length
      ? tblSrc.map(d => `
      <tr>
        <td>${d.tanggal}</td>
        <td style="font-weight:600;max-width:140px;overflow:hidden;text-overflow:ellipsis">${toProper(d.nama || d.email)}</td>
        <td class="td-pasca"><span class="n-open">${n(d.open_pasca).toLocaleString('id-ID')}</span></td>
        <td class="td-pasca"><span class="n-submit">${n(d.submit_pasca).toLocaleString('id-ID')}</span></td>
        <td class="td-pasca"><span class="n-reject">${n(d.reject_pasca).toLocaleString('id-ID')}</span></td>
        <td class="td-praba"><span class="n-open">${n(d.open_praba).toLocaleString('id-ID')}</span></td>
        <td class="td-praba"><span class="n-submit">${n(d.submit_praba).toLocaleString('id-ID')}</span></td>
        <td class="td-praba"><span class="n-reject">${n(d.reject_praba).toLocaleString('id-ID')}</span></td>
      </tr>`).join('')
      : '<tr><td colspan="8" style="text-align:center;color:var(--muted);padding:1.5rem">Belum ada data</td></tr>';
  }
  updateMsLabel();
  _syncRiwayatTblHeaders();
}

// ── Interactions ──────────────────────────────────────────────────────────────
function doSearch() { renderTable(); }

function sortBy(col) {
  sortDir = sortCol === col && sortDir === 'asc' ? 'desc' : 'asc';
  sortCol = col;
  renderTable();
}

// ── Theme ─────────────────────────────────────────────────────────────────────
function _isDark() {
  const t = document.documentElement.getAttribute('data-theme');
  if (t === 'dark')  return true;
  if (t === 'light') return false;
  return window.matchMedia('(prefers-color-scheme: dark)').matches;
}

function _applyThemeIcon() {
  const icon = document.getElementById('themeIcon');
  if (icon) icon.textContent = _isDark() ? '☀️' : '🌙';
}

function _applyChartDefaults() {
  if (typeof Chart === 'undefined') return;
  const dark = _isDark();
  Chart.defaults.color       = dark ? '#94A3B8' : '#64748B';
  Chart.defaults.borderColor = dark ? '#334155' : '#E2E8F0';
}

function toggleTheme() {
  const root = document.documentElement;
  const next = _isDark() ? 'light' : 'dark';
  root.setAttribute('data-theme', next);
  localStorage.setItem('theme', next);
  _applyThemeIcon();
  _applyChartDefaults();
  // Re-render charts with new colors
  if (allData.length)    render();
  if (ringkasanData.length) renderRingkasan();
  if (riwayatData.length) renderRiwayat();
}

// Init theme icon on load
_applyThemeIcon();
_applyChartDefaults();
// Update icon if system preference changes
window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => {
  if (!localStorage.getItem('theme')) {
    _applyThemeIcon();
    _applyChartDefaults();
    if (allData.length)    render();
    if (ringkasanData.length) renderRingkasan();
    if (riwayatData.length) renderRiwayat();
  }
});

// ── Sheet Config Panel ────────────────────────────────────────────────────────
function toggleMobileMenu() {
  const r = document.querySelector('.hdr-right');
  const open = r.classList.toggle('open');
  document.getElementById('mobileMenuBtn').setAttribute('aria-expanded', open);
}
function _closeMobileMenu() {
  const r = document.querySelector('.hdr-right');
  if (r.classList.contains('open')) {
    r.classList.remove('open');
    document.getElementById('mobileMenuBtn').setAttribute('aria-expanded', 'false');
  }
}
function toggleSettings() {
  const p = document.getElementById('settingsPanel');
  const open = !p.classList.contains('open');
  p.classList.toggle('open', open);
  if (open) { _populateSettingsInputs(); _closeMobileMenu(); }
}
function _populateSettingsInputs() {
  const sid = localStorage.getItem('cfg_sheet_id') || SHEET_ID;
  if (sid)
    document.getElementById('cfgUrlSheet').value =
      `https://docs.google.com/spreadsheets/d/${sid}/edit`;
}
function saveSheetConfig() {
  const url = document.getElementById('cfgUrlSheet').value.trim();
  const msg = document.getElementById('cfgMsg');
  const sid = _parseSheetId(url);
  if (!sid) { msg.textContent = '⚠️ URL tidak valid — pastikan URL berasal dari Google Sheets.'; return; }
  localStorage.setItem('cfg_sheet_id', sid);
  SHEET_ID = sid;
  msg.textContent = `✅ Disimpan — Sheet ID: ${sid}`;
  loadData(); loadRingkasan(); loadRiwayat();
}
function resetSheetConfig() {
  localStorage.removeItem('cfg_sheet_id');
  ['cfg_gid_utama','cfg_gid_ringkasan','cfg_gid_riwayat'].forEach(k => localStorage.removeItem(k));
  document.getElementById('cfgUrlSheet').value = '';
  SHEET_ID = _DEFAULT_SHEET_ID;
  document.getElementById('cfgMsg').textContent = '↺ Kembali ke sumber data default.';
  loadData(); loadRingkasan(); loadRiwayat();
}

// ── Count-up animation ────────────────────────────────────────────────────────
function _animateStatNumbers() {
  document.querySelectorAll('.s-stat-num').forEach(el => {
    const raw = parseInt(el.textContent.replace(/[.,\s]/g, ''), 10);
    if (isNaN(raw) || raw === 0) return;
    el.textContent = '0';
    const dur = Math.min(700, 300 + raw * 4);
    const start = performance.now();
    function step(now) {
      const p = Math.min((now - start) / dur, 1);
      const ease = 1 - Math.pow(1 - p, 3);
      el.textContent = Math.round(raw * ease).toLocaleString('id-ID');
      if (p < 1) requestAnimationFrame(step);
    }
    requestAnimationFrame(step);
  });
}

// ── Accent color ──────────────────────────────────────────────────────────────
const ACCENT_PRESETS = {
  blue:   { from:'#1a3f6f', to:'#2E75B6', primary:'#1F4E79', light:'#2E75B6' },
  teal:   { from:'#0a4d4d', to:'#0F9D9D', primary:'#0D6E6E', light:'#14b8a6' },
  green:  { from:'#034d34', to:'#059669', primary:'#065F46', light:'#10b981' },
  purple: { from:'#3b0764', to:'#7C3AED', primary:'#4C1D95', light:'#8b5cf6' },
  rose:   { from:'#7f1d1d', to:'#e11d48', primary:'#881337', light:'#f43f5e' },
  orange: { from:'#451a03', to:'#ea580c', primary:'#7c2d12', light:'#f97316' },
  slate:  { from:'#0f172a', to:'#475569', primary:'#1e293b', light:'#64748b' },
};

function _darkenHex(hex, amt) {
  const n = parseInt(hex.replace('#',''), 16);
  const r = Math.max(0, Math.round(((n >> 16) & 255) * (1 - amt)));
  const g = Math.max(0, Math.round(((n >>  8) & 255) * (1 - amt)));
  const b = Math.max(0, Math.round(( n        & 255) * (1 - amt)));
  return '#' + [r, g, b].map(v => v.toString(16).padStart(2,'0')).join('');
}

function _applyAccentVars(from, to, primary, light) {
  const r = document.documentElement;
  r.style.setProperty('--header-from',   from);
  r.style.setProperty('--header-to',     to);
  r.style.setProperty('--primary',       primary);
  r.style.setProperty('--primary-light', light);
}

function _markActiveSwatch(key) {
  document.querySelectorAll('.accent-swatch').forEach(el =>
    el.classList.toggle('active', el.dataset.accent === key));
}

function loadAccentConfig() {
  const saved = localStorage.getItem('cfg_accent');
  if (!saved) return;
  try {
    const data = JSON.parse(saved);
    if (data.preset && ACCENT_PRESETS[data.preset]) {
      const p = ACCENT_PRESETS[data.preset];
      _applyAccentVars(p.from, p.to, p.primary, p.light);
      _markActiveSwatch(data.preset);
    } else if (data.custom) {
      const hex = data.custom;
      _applyAccentVars(_darkenHex(hex, .4), hex, _darkenHex(hex, .2), hex);
      const el = document.getElementById('accentCustomColor');
      if (el) el.value = hex;
      document.querySelectorAll('.accent-swatch').forEach(el => el.classList.remove('active'));
    }
  } catch(e) {}
}

function selectAccent(preset) {
  const p = ACCENT_PRESETS[preset];
  if (!p) return;
  _applyAccentVars(p.from, p.to, p.primary, p.light);
  localStorage.setItem('cfg_accent', JSON.stringify({ preset }));
  _markActiveSwatch(preset);
}

function resetAccent() {
  localStorage.removeItem('cfg_accent');
  const p = ACCENT_PRESETS.blue;
  _applyAccentVars(p.from, p.to, p.primary, p.light);
  _markActiveSwatch('blue');
  const el = document.getElementById('accentCustomColor');
  if (el) el.value = '#1F4E79';
}

document.addEventListener('DOMContentLoaded', function() {
  const picker = document.getElementById('accentCustomColor');
  if (picker) {
    picker.addEventListener('input', function() {
      const hex = this.value;
      _applyAccentVars(_darkenHex(hex, .4), hex, _darkenHex(hex, .2), hex);
      localStorage.setItem('cfg_accent', JSON.stringify({ custom: hex }));
      document.querySelectorAll('.accent-swatch').forEach(el => el.classList.remove('active'));
    });
  }
});

// ── Scroll-to-top ─────────────────────────────────────────────────────────────
window.addEventListener('scroll', () => {
  document.getElementById('scrollTop').classList.toggle('visible', window.scrollY > 300);
}, { passive: true });

// ── Init ──────────────────────────────────────────────────────────────────────
_loadSheetConfig();
_populateSettingsInputs();   // pre-fill URL input meski panel settings belum dibuka
loadAccentConfig();           // apply saved accent color
setInterval(loadData,    5 * 60 * 1000);
setInterval(loadRingkasan,  5 * 60 * 1000);
setInterval(loadRiwayat, 5 * 60 * 1000);
loadData();
loadRingkasan();
loadRiwayat();
