// ── Left Drawer Sidebar ────────────────────────────────────────────────────────

function openSidebar() {
  document.getElementById('appSidebar').classList.add('open');
  document.getElementById('sidebarOverlay').classList.add('visible');
  document.getElementById('sidebarToggle').setAttribute('aria-expanded', 'true');
  document.body.style.overflow = 'hidden';
}

function closeSidebar() {
  document.getElementById('appSidebar').classList.remove('open');
  document.getElementById('sidebarOverlay').classList.remove('visible');
  document.getElementById('sidebarToggle').setAttribute('aria-expanded', 'false');
  document.body.style.overflow = '';
}

function _sbToggleExpand(bodyId, btnId) {
  const body = document.getElementById(bodyId);
  const btn  = document.getElementById(btnId);
  if (!body) return;
  const isOpen = body.style.display !== 'none';
  body.style.display = isOpen ? 'none' : 'block';
  if (btn) btn.classList.toggle('expanded', !isOpen);
}

function toggleSbAlert() {
  // Close others
  const settBody = document.getElementById('sbSettingsBody');
  const appBody  = document.getElementById('sbAppearanceBody');
  if (settBody && settBody.style.display !== 'none') _sbToggleExpand('sbSettingsBody',  'sbSettingsBtn');
  if (appBody  && appBody.style.display  !== 'none') _sbToggleExpand('sbAppearanceBody','sbAppearanceBtn');

  _sbToggleExpand('sbAlertBody', 'sbAlertBtn');
  const isNowOpen = document.getElementById('sbAlertBody')?.style.display !== 'none';
  if (isNowOpen) {
    alertPanelOpen = true;
    _buildAlertList(_getBelowThreshold());
  } else {
    alertPanelOpen = false;
  }
}

function toggleSbSettings() {
  const alertBody = document.getElementById('sbAlertBody');
  const appBody   = document.getElementById('sbAppearanceBody');
  if (alertBody && alertBody.style.display !== 'none') _sbToggleExpand('sbAlertBody',      'sbAlertBtn');
  if (appBody   && appBody.style.display   !== 'none') _sbToggleExpand('sbAppearanceBody', 'sbAppearanceBtn');

  _sbToggleExpand('sbSettingsBody', 'sbSettingsBtn');
  if (document.getElementById('sbSettingsBody')?.style.display !== 'none') {
    _populateSettingsInputs();
  }
}

function toggleSbAppearance() {
  const alertBody = document.getElementById('sbAlertBody');
  const settBody  = document.getElementById('sbSettingsBody');
  if (alertBody && alertBody.style.display !== 'none') _sbToggleExpand('sbAlertBody',     'sbAlertBtn');
  if (settBody  && settBody.style.display  !== 'none') _sbToggleExpand('sbSettingsBody',  'sbSettingsBtn');

  _sbToggleExpand('sbAppearanceBody', 'sbAppearanceBtn');
  _updateSbThemeLabel();
}

function _updateSbThemeLabel() {
  const lbl = document.getElementById('sbThemeLabel');
  if (lbl) lbl.textContent = _isDark() ? 'Mode Terang' : 'Mode Gelap';
}

// Keyboard: close sidebar on Escape
document.addEventListener('keydown', function(e) {
  if (e.key === 'Escape') closeSidebar();
});

// ── Section Observer for active nav item ────────────────────────────────────
let _navObserver = null;

function navScrollTo(id) {
  const el = document.getElementById(id);
  if (!el) return;
  const hdrH = (document.querySelector('header')?.offsetHeight || 60) + 8;
  const top  = el.getBoundingClientRect().top + window.scrollY - hdrH;
  window.scrollTo({ top, behavior: 'smooth' });
}

function _setupNavObserver() {
  const nav = document.getElementById('sbNavLinks');
  if (!nav) return;

  if (_navObserver) { _navObserver.disconnect(); _navObserver = null; }

  const secIds = ['sec-ringkasan','sec-peringkat','sec-chart','sec-detail','sec-tren','sec-pencacah'];

  // Show/hide nav items based on whether section exists in DOM
  secIds.forEach(id => {
    const item = nav.querySelector(`[data-target="${id}"]`);
    if (item) item.style.display = document.getElementById(id) ? '' : 'none';
  });

  _navObserver = new IntersectionObserver(entries => {
    entries.forEach(entry => {
      const item = nav.querySelector(`[data-target="${entry.target.id}"]`);
      if (item) item.classList.toggle('active', entry.isIntersecting);
    });
  }, { rootMargin: '-8% 0px -72% 0px', threshold: 0 });

  secIds.forEach(id => {
    const el = document.getElementById(id);
    if (el) _navObserver.observe(el);
  });
}
