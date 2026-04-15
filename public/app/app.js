// ══════════════════════════════════════════════════════════
//  API CLIENT — Semua panggilan ke Vercel API Route
//  BASE_URL otomatis deteksi (dev vs prod)
// ══════════════════════════════════════════════════════════
var BASE_URL = (function () {
  var host = window.location.origin;
  if (host.startsWith('file://')) return 'http://localhost:3000';
  return host;
})();

async function callAPI(action, payload) {
  try {
    var res = await fetch(BASE_URL + '/api/' + action, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: action, payload: payload || {} })
    });

    var data = null;
    try { data = await res.json(); } catch (e) { }

    if (!res.ok) {
      if (data && data.message) throw new Error(data.message);
      throw new Error('HTTP ' + res.status);
    }
    return data;
  } catch (err) {
    var badge = document.getElementById('api-status-badge');
    if (badge) { badge.className = 'api-status error'; badge.innerHTML = '<i class="bi bi-circle-fill"></i> Offline'; }
    throw err;
  }
}

// ══════════════════════════════════════════════════════════
//  STATE
// ══════════════════════════════════════════════════════════
var APP = { user: null, currentPage: 'dashboard' };
var currentEditData = null;

// ══════════════════════════════════════════════════════════
//  SPINNER & TOAST
// ══════════════════════════════════════════════════════════
function showSpinner(label) {
  document.getElementById('spinner-label').textContent = label || 'Memproses...';
  document.getElementById('spinner-overlay').classList.add('active');
}
function hideSpinner() {
  document.getElementById('spinner-overlay').classList.remove('active');
}
function showToast(msg, type) {
  var icons = { success: 'check-circle-fill', error: 'x-circle-fill', info: 'info-circle-fill' };
  var el = document.createElement('div');
  el.className = 'toast-msg ' + (type || 'info');
  el.innerHTML = '<i class="bi bi-' + (icons[type] || icons.info) + '"></i><span>' + esc(msg) + '</span>';
  document.getElementById('toast-container').appendChild(el);
  setTimeout(function () { el.remove(); }, 4500);
}

// ══════════════════════════════════════════════════════════
//  THEME (DARK MODE)
// ══════════════════════════════════════════════════════════
function toggleTheme() {
  const isDark = document.body.classList.toggle('dark-mode');
  localStorage.setItem('simadun_theme', isDark ? 'dark' : 'light');
  document.getElementById('theme-icon').className = isDark ? 'bi bi-sun-fill' : 'bi bi-moon-fill';
}

function initTheme() {
  const savedTheme = localStorage.getItem('simadun_theme');
  if (savedTheme === 'dark') {
    document.body.classList.add('dark-mode');
    const ti = document.getElementById('theme-icon');
    if (ti) ti.className = 'bi bi-sun-fill';
  }
}

// ══════════════════════════════════════════════════════════
//  CLOCK
// ══════════════════════════════════════════════════════════
function startClock() {
  function tick() {
    var now = new Date();
    var el = document.getElementById('clock');
    if (el) el.textContent = now.toLocaleTimeString('id-ID');
    var dt = document.getElementById('info-date');
    if (dt) dt.textContent = now.toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
  }
  tick();
  setInterval(tick, 1000);
}

// ══════════════════════════════════════════════════════════
//  NAVIGATION
// ══════════════════════════════════════════════════════════
var PAGE_TITLES = {
  dashboard:  'Dashboard',
  arsip:      'Manajemen Arsip',
  surat:      'Manajemen Surat',
  spt:        'Surat Perintah Tugas',
  panduan:    'Panduan Teknis',
  pengaturan: 'Pengaturan Sistem'
};

var PAGE_SUBS = {
  dashboard:  'Ringkasan data kearsipan sistem',
  arsip:      'Upload & kelola dokumen arsip',
  surat:      'Surat Masuk, Keluar & Undangan',
  spt:        'Arsip & Pembuatan SPT Dinas',
  panduan:    'Dokumentasi lengkap fitur SIMADUN',
  pengaturan: 'Konfigurasi akun & sistem'
};

function navigateTo(page) {
  document.querySelectorAll('.view').forEach(function (v) { v.classList.remove('active'); });
  var target = document.getElementById('page-' + page);
  if (target) target.classList.add('active');
  document.querySelectorAll('.nav-link-item').forEach(function (l) { l.classList.remove('active'); });
  var activeLink = document.querySelector('[data-page="' + page + '"]');
  if (activeLink) activeLink.classList.add('active');
  document.getElementById('topbar-title').textContent = PAGE_TITLES[page] || page;
  document.getElementById('topbar-sub').textContent = PAGE_SUBS[page] || '';
  APP.currentPage = page;
  closeSidebar();
  if (page === 'dashboard') loadDashboard();
  if (page === 'arsip') loadArsip();
  if (page === 'surat') { loadSuratMasuk(); loadSuratKeluar(); loadUndangan(); }
  if (page === 'spt') loadSPT();
  if (page === 'pengaturan') {
    loadUsers();
    loadKopSettings();
    var chgEl = document.getElementById('chg-username');
    if (chgEl) chgEl.value = APP.user ? APP.user.username : '';
  }
}

function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
  document.getElementById('sidebar-overlay').classList.toggle('active');
}
function closeSidebar() {
  document.getElementById('sidebar').classList.remove('open');
  document.getElementById('sidebar-overlay').classList.remove('active');
}

function toggleCollapse(contentId, trigger) {
  var el = document.getElementById(contentId);
  if (!el) return;
  el.classList.toggle('collapsed');
  if (trigger) trigger.classList.toggle('collapsed');
}

function togglePanduanAccordion(header) {
  var body = header.nextElementSibling;
  var isOpen = body.classList.contains('open');
  // close all first
  document.querySelectorAll('.panduan-accordion-body').forEach(function(b) { b.classList.remove('open'); });
  document.querySelectorAll('.panduan-accordion-header').forEach(function(h) { h.classList.remove('open'); });
  if (!isOpen) {
    body.classList.add('open');
    header.classList.add('open');
  }
}

function togglePwd(inputId, btn) {
  var inp = document.getElementById(inputId);
  if (inp.type === 'password') { inp.type = 'text'; btn.innerHTML = '<i class="bi bi-eye-slash"></i>'; }
  else { inp.type = 'password'; btn.innerHTML = '<i class="bi bi-eye"></i>'; }
}

// ══════════════════════════════════════════════════════════
//  LOGIN / LOGOUT
// ══════════════════════════════════════════════════════════
async function doLogin() {
  var u = document.getElementById('login-username').value.trim();
  var p = document.getElementById('login-password').value;
  if (!u || !p) { showToast('Username dan password wajib diisi.', 'error'); return; }
  showSpinner('Memverifikasi...');
  try {
    var res = await callAPI('login', { username: u, password: p });
    hideSpinner();
    if (res.success) {
      APP.user = res.user;
      document.getElementById('view-login').style.display = 'none';
      document.getElementById('app-shell').style.display = 'block';
      document.getElementById('user-nama').textContent = res.user.nama;
      document.getElementById('user-role').textContent = res.user.role;
      document.getElementById('user-avatar').textContent = res.user.nama.charAt(0).toUpperCase();
      document.getElementById('info-user').textContent = res.user.nama;
      document.getElementById('chg-username').value = res.user.username;
      startClock();
      navigateTo('dashboard');
      showToast('Selamat datang, ' + res.user.nama + '!', 'success');
    } else {
      showToast(res.message, 'error');
    }
  } catch (err) {
    hideSpinner();
    showToast('Gagal terhubung ke server: ' + err.message, 'error');
  }
}

function doLogout() {
  if (!confirm('Yakin ingin keluar?')) return;
  APP.user = null;
  document.getElementById('app-shell').style.display = 'none';
  document.getElementById('view-login').style.display = 'flex';
  document.getElementById('login-username').value = '';
  document.getElementById('login-password').value = '';
  showToast('Berhasil keluar.', 'info');
}

// ══════════════════════════════════════════════════════════
//  DASHBOARD
// ══════════════════════════════════════════════════════════
async function loadDashboard() {
  try {
    var res = await callAPI('getDashboard', {});
    
    // FALLBACK CLIENT-SIDE COUNTING
    // Membaca data arsip mentah jika backend api [action].js belum di-deploy.
    var fallbackUd = 0;
    var fallbackInd = 0;
    try {
      var aRes = await callAPI('getArsip', {});
      if (aRes && aRes.success && aRes.data) {
        aRes.data.forEach(function(d) {
          var c = (d['Kategori'] || '').trim().toUpperCase();
          if (c === 'UNDANGAN') fallbackUd++;
          if (c === 'INDISIPLINER ASN') fallbackInd++;
        });
      }
    } catch(e) {}

    if (res.success) {
      animateCount('stat-masuk', res.suratMasuk);
      animateCount('stat-keluar', res.suratKeluar);
      animateCount('stat-undangan', res.undangan);
      animateCount('stat-spt', res.spt);
      animateCount('stat-arsip', res.arsip); // TALLY TOTAL ARSIP FIX
      animateCount('stat-dumas', res.countDumas);
      animateCount('stat-lhp', res.countLhp);
      animateCount('stat-mcsp', res.countMcsp);
      animateCount('stat-pdtt', res.countPdtt);
      animateCount('stat-pkkn', res.countPkkn);
      animateCount('stat-pktpt', res.countPktpt);
      animateCount('stat-sktim', res.countSkTim);
      animateCount('stat-bap', res.countBap);
      animateCount('stat-pka', res.countPka);
      animateCount('stat-iepk', res.countIepk);
      animateCount('stat-nota', res.countNota);
      animateCount('stat-perbup', res.countPerbup);
      animateCount('stat-arsip-undangan', res.countUndanganArsip !== undefined ? res.countUndanganArsip : fallbackUd);
      animateCount('stat-arsip-indisipliner', res.countIndisipliner !== undefined ? res.countIndisipliner : fallbackInd);
    }
  } catch (err) { /* silent */ }
}

function animateCount(id, target) {
  var el = document.getElementById(id);
  if (!el) return;
  var start = 0;
  var step = Math.max(1, Math.ceil(target / 30));
  var timer = setInterval(function () {
    start = Math.min(start + step, target);
    el.textContent = start;
    if (start >= target) clearInterval(timer);
  }, 20);
}

// ══════════════════════════════════════════════════════════
//  FILE HANDLING
// ══════════════════════════════════════════════════════════
function handleFileSelect(inputId, infoId) {
  var file = document.getElementById(inputId).files[0];
  if (!file) return;
  document.getElementById(infoId).textContent = '📎 ' + file.name + ' (' + (file.size / 1024).toFixed(1) + ' KB)';
}

function readFileAsBase64(file) {
  return new Promise(function (resolve, reject) {
    var reader = new FileReader();
    reader.onload = function (e) {
      var base64 = e.target.result.split(',')[1];
      resolve({
        content: base64,
        name: file.name,
        mimeType: file.type || 'application/octet-stream',
        size: (file.size / 1024).toFixed(1) + ' KB'
      });
      document.getElementById('spinner-label').textContent = 'Memproses unggahan (File siap)...';
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ══════════════════════════════════════════════════════════
//  ARSIP
// ══════════════════════════════════════════════════════════
async function submitArsip() {
  var nama = document.getElementById('arsip-nama').value.trim();
  var kategori = document.getElementById('arsip-kategori').value;
  var deskripsi = document.getElementById('arsip-deskripsi').value.trim();
  var tglArsip = document.getElementById('arsip-tgl').value;
  var fileEl = document.getElementById('arsip-file');
  if (!nama) { showToast('Nama file wajib diisi.', 'error'); return; }
  if (!kategori) { showToast('Kategori wajib dipilih.', 'error'); return; }
  if (!tglArsip) { showToast('Tanggal arsip wajib diisi.', 'error'); return; }
  if (!fileEl.files[0]) { showToast('File wajib diunggah.', 'error'); return; }
  showSpinner('Mengunggah arsip ke Google Drive...');
  try {
    var fd = await readFileAsBase64(fileEl.files[0]);
    var res = await callAPI('saveArsip', { data: { namaFile: nama, kategori: kategori, deskripsi: deskripsi, tglArsip: tglArsip }, fileData: fd });
    hideSpinner();
    if (res.success) {
      showToast(res.message, 'success');
      resetFields(['arsip-nama', 'arsip-deskripsi']);
      document.getElementById('arsip-kategori').value = '';
      fileEl.value = '';
      document.getElementById('arsip-file-info').textContent = '';
      togglePanel('arsip-form-panel');
      loadArsip();
    } else showToast(res.message, 'error');
  } catch (err) { hideSpinner(); showToast('Error: ' + err.message, 'error'); }
}

async function loadArsip() {
  try {
    var res = await callAPI('getArsip', {});
    var tbody = document.getElementById('arsip-tbody');
    if (!res.success || !res.data.length) {
      tbody.innerHTML = '<tr class="no-data"><td colspan="9"><i class="bi bi-inbox" style="font-size:2rem;display:block;margin-bottom:8px"></i>Belum ada data arsip</td></tr>'; return;
    }
    tbody.innerHTML = res.data.map(function (d, i) {
      var safeData = encodeURIComponent(JSON.stringify(d));
      var tglArs = d['Tanggal Arsip'] ? fmtDate(d['Tanggal Arsip']) : '-';
      var url = d['URL'] || d['File URL'];
      var actBtn = url ? '<button class="btn-link-custom" onclick="openPreview(\'' + url + '\')"><i class="bi bi-eye"></i> Lihat</button>' : '';
      return '<tr><td>' + (i + 1) + '</td><td><strong>' + esc(d['Nama File']) + '</strong></td><td><span class="badge-cat arsip">' + esc(d['Kategori']) + '</span></td><td>' + esc(d['Folder']) + '</td><td>' + esc(d['Deskripsi']) + '</td><td>' + esc(d['Ukuran']) + '</td><td>' + tglArs + '</td><td>' + fmtDate(d['DibuatPada'] || d['CreatedAt']) + '</td><td class="action-col" style="display:flex;gap:6px">' + actBtn + '<button class="btn-warning-custom" onclick="openEditModal(\'Arsip\', \'' + safeData + '\')"><i class="bi bi-pencil"></i></button><button class="btn-danger-custom" onclick="deleteItem(\'deleteArsip\',\'' + d['ID'] + '\',loadArsip)"><i class="bi bi-trash"></i></button></td></tr>';
    }).join('');
  } catch (err) { showToast('Gagal memuat arsip: ' + err.message, 'error'); }
}

// ══════════════════════════════════════════════════════════
//  SURAT MASUK
// ══════════════════════════════════════════════════════════
async function submitSuratMasuk() {
  var data = { nomorSurat: v('sm-nomor'), tanggal: v('sm-tanggal'), pengirim: v('sm-pengirim'), perihal: v('sm-perihal'), kategori: v('sm-kategori'), catatan: v('sm-catatan') };
  if (!data.nomorSurat || !data.tanggal || !data.pengirim || !data.perihal) { showToast('Lengkapi field yang wajib diisi.', 'error'); return; }
  showSpinner('Menyimpan surat masuk...');
  try {
    var fileEl = document.getElementById('sm-file');
    var fd = fileEl.files[0] ? await readFileAsBase64(fileEl.files[0]) : null;
    var res = await callAPI('saveSuratMasuk', { data: data, fileData: fd });
    hideSpinner();
    if (res.success) {
      showToast(res.message, 'success');
      resetFields(['sm-nomor', 'sm-tanggal', 'sm-pengirim', 'sm-perihal', 'sm-catatan']);
      fileEl.value = ''; document.getElementById('sm-file-info').textContent = '';
      togglePanel('form-masuk'); loadSuratMasuk();
    } else showToast(res.message, 'error');
  } catch (err) { hideSpinner(); showToast('Error: ' + err.message, 'error'); }
}

async function loadSuratMasuk() {
  try {
    var res = await callAPI('getSuratMasuk', {});
    renderSuratTable('tbody-sm', res, ['Nomor Surat', 'Tanggal', 'Pengirim', 'Perihal', 'Kategori'], 'masuk', 'deleteSuratMasuk', loadSuratMasuk);
  } catch (err) { /* silent */ }
}

// ══════════════════════════════════════════════════════════
//  SURAT KELUAR
// ══════════════════════════════════════════════════════════
async function submitSuratKeluar() {
  var data = { nomorSurat: v('sk-nomor'), tanggal: v('sk-tanggal'), tujuan: v('sk-tujuan'), perihal: v('sk-perihal'), kategori: v('sk-kategori'), catatan: v('sk-catatan') };
  if (!data.nomorSurat || !data.tanggal || !data.tujuan || !data.perihal) { showToast('Lengkapi field yang wajib diisi.', 'error'); return; }
  showSpinner('Menyimpan surat keluar...');
  try {
    var fileEl = document.getElementById('sk-file');
    var fd = fileEl.files[0] ? await readFileAsBase64(fileEl.files[0]) : null;
    var res = await callAPI('saveSuratKeluar', { data: data, fileData: fd });
    hideSpinner();
    if (res.success) {
      showToast(res.message, 'success');
      resetFields(['sk-nomor', 'sk-tanggal', 'sk-tujuan', 'sk-perihal', 'sk-catatan']);
      fileEl.value = ''; document.getElementById('sk-file-info').textContent = '';
      togglePanel('form-keluar'); loadSuratKeluar();
    } else showToast(res.message, 'error');
  } catch (err) { hideSpinner(); showToast('Error: ' + err.message, 'error'); }
}

async function loadSuratKeluar() {
  try {
    var res = await callAPI('getSuratKeluar', {});
    renderSuratTable('tbody-sk', res, ['Nomor Surat', 'Tanggal', 'Tujuan', 'Perihal', 'Kategori'], 'keluar', 'deleteSuratKeluar', loadSuratKeluar);
  } catch (err) { /* silent */ }
}

// ══════════════════════════════════════════════════════════
//  UNDANGAN
// ══════════════════════════════════════════════════════════
async function submitUndangan() {
  var data = { nomorSurat: v('ud-nomor'), tanggal: v('ud-tanggal'), penyelenggara: v('ud-penyelenggara'), perihal: v('ud-perihal'), lokasi: v('ud-lokasi'), catatan: v('ud-catatan') };
  if (!data.nomorSurat || !data.tanggal || !data.penyelenggara || !data.perihal) { showToast('Lengkapi field yang wajib diisi.', 'error'); return; }
  showSpinner('Menyimpan undangan...');
  try {
    var fileEl = document.getElementById('ud-file');
    var fd = fileEl.files[0] ? await readFileAsBase64(fileEl.files[0]) : null;
    var res = await callAPI('saveUndangan', { data: data, fileData: fd });
    hideSpinner();
    if (res.success) {
      showToast(res.message, 'success');
      resetFields(['ud-nomor', 'ud-tanggal', 'ud-penyelenggara', 'ud-perihal', 'ud-lokasi', 'ud-catatan']);
      fileEl.value = ''; document.getElementById('ud-file-info').textContent = '';
      togglePanel('form-undangan'); loadUndangan();
    } else showToast(res.message, 'error');
  } catch (err) { hideSpinner(); showToast('Error: ' + err.message, 'error'); }
}

async function loadUndangan() {
  try {
    var res = await callAPI('getUndangan', {});
    renderSuratTable('tbody-ud', res, ['Nomor Surat', 'Tanggal', 'Penyelenggara', 'Perihal', 'Lokasi'], 'undang', 'deleteUndangan', loadUndangan);
  } catch (err) { /* silent */ }
}

// ══════════════════════════════════════════════════════════
//  RENDER SURAT TABLE
// ══════════════════════════════════════════════════════════
function renderSuratTable(tbodyId, res, cols, badgeClass, deleteAction, reloadFn) {
  var tbody = document.getElementById(tbodyId);
  var colCount = cols.length + 3;
  if (!res.success || !res.data.length) {
    tbody.innerHTML = '<tr class="no-data"><td colspan="' + colCount + '"><i class="bi bi-inbox" style="font-size:2rem;display:block;margin-bottom:8px"></i>Belum ada data</td></tr>'; return;
  }
  tbody.innerHTML = res.data.map(function (d, i) {
    var cells = cols.map(function (col) {
      var val = d[col] || '-';
      if (col === 'Kategori') return '<td><span class="badge-cat ' + badgeClass + '">' + esc(val) + '</span></td>';
      if (col === 'Tanggal') return '<td>' + fmtDate(val) + '</td>';
      return '<td>' + esc(val) + '</td>';
    }).join('');

    // Mapping Sheet Name for Edit Modal depending on deleteAction
    var shName = 'Surat Masuk';
    if(deleteAction === 'deleteSuratKeluar') shName = 'Surat Keluar';
    if(deleteAction === 'deleteUndangan') shName = 'Undangan';

    var safeData = encodeURIComponent(JSON.stringify(d));
    var url = d['URL'] || d['File URL'];
    var lampiran = url ? '<button class="btn-link-custom action-col" style="padding:5px 8px;font-size:0.75rem" onclick="openPreview(\'' + url + '\')"><i class="bi bi-eye"></i> View</button>' : '<span style="color:var(--text-muted);font-size:.78rem">-</span>';
    return '<tr><td>' + (i + 1) + '</td>' + cells + '<td>' + lampiran + '</td><td class="action-col" style="display:flex;gap:6px"><button class="btn-warning-custom" onclick="openEditModal(\''+shName+'\', \'' + safeData + '\')"><i class="bi bi-pencil"></i></button><button class="btn-danger-custom" onclick="deleteItem(\'' + deleteAction + '\',\'' + d['ID'] + '\',' + reloadFn.name + ')"><i class="bi bi-trash"></i></button></td></tr>';
  }).join('');
}

// ══════════════════════════════════════════════════════════
//  SPT - TAB HANDLING
// ══════════════════════════════════════════════════════════
function setSptTab(tabId) {
  document.querySelectorAll('#page-spt .custom-tab').forEach(function (t) { t.classList.remove('active'); });
  document.querySelectorAll('#page-spt .tab-content').forEach(function (c) { c.style.display = 'none'; });
  
  var activeBtn = document.getElementById('tab-' + tabId);
  var contentBlock = document.getElementById('content-' + tabId);
  if (activeBtn) activeBtn.classList.add('active');
  if (contentBlock) contentBlock.style.display = 'block';
  
  if (tabId === 'spt-arsip') loadSPT();
}

async function submitArsipSPT() {
  var data = { 
    nomorSpt: v('arsip-spt-nomor'), 
    tujuan: v('arsip-spt-tujuan'), 
    keterangan: v('arsip-spt-ket') 
  };
  var fileEl = document.getElementById('arsip-spt-file');

  if (!data.nomorSpt || !fileEl.files[0]) { showToast('Nomor SPT dan file dokumen (PDF) wajib diisi.', 'error'); return; }
  
  showSpinner('Mengkompres & Mengunggah Arsip SPT...');
  try {
    var fd = await readFileAsBase64(fileEl.files[0]);
    // Populate fake fallback fields so spreadsheet format remains consistent
    data.nama = '-'; data.nip = '-'; data.jabatan = '-'; 
    data.keperluan = 'Arsip Scan PDF'; data.tglBerangkat = '-'; data.tglKembali = '-';

    var res = await callAPI('saveSPT', { data: data, fileData: fd });
    hideSpinner();
    if (res.success) {
      showToast(res.message, 'success');
      resetFields(['arsip-spt-nomor', 'arsip-spt-tujuan', 'arsip-spt-ket']);
      fileEl.value = '';
      var infoEl = document.getElementById('arsip-spt-file-info');
      if (infoEl) infoEl.textContent = '';
      togglePanel('form-spt-arsip'); 
      loadSPT();
    } else {
      showToast(res.message, 'error');
    }
  } catch (err) { hideSpinner(); showToast('Error: ' + err.message, 'error'); }
}

async function submitCreateSPT() {
  var data = { 
    nomorSpt: v('spt-nomor'), 
    nama: v('spt-nama'), 
    nip: v('spt-nip'), 
    jabatan: v('spt-jabatan'), 
    tujuan: v('spt-tujuan'), 
    keperluan: v('spt-keperluan'), 
    tglBerangkat: v('spt-tgl-berangkat'), 
    tglKembali: v('spt-tgl-kembali'), 
    keterangan: '-' 
  };
  if (!data.nomorSpt || !data.nama || !data.tujuan || !data.keperluan || !data.tglBerangkat || !data.tglKembali) { 
    showToast('Lengkapi field SPT yang wajib diisi.', 'error'); return; 
  }
  
  showSpinner('Meregistrasi SPT Baru ke Database...');
  try {
    var res = await callAPI('saveSPT', { data: data, fileData: null });
    hideSpinner();
    if (res.success) {
      showToast('Berhasil disimpan. Mencetak dokumen...', 'success');
      data['Nomor SPT'] = data.nomorSpt; data['Nama'] = data.nama; data['NIP'] = data.nip; data['Jabatan'] = data.jabatan; data['Tujuan'] = data.tujuan; data['Keperluan'] = data.keperluan; data['Tgl Berangkat'] = data.tglBerangkat; data['Tgl Kembali'] = data.tglKembali;
      
      var safeData = encodeURIComponent(JSON.stringify(data));
      printDocSPT(safeData);

      resetFields(['spt-nomor', 'spt-nama', 'spt-nip', 'spt-jabatan', 'spt-tujuan', 'spt-keperluan', 'spt-tgl-berangkat', 'spt-tgl-kembali']);
      setSptTab('spt-arsip');
    } else {
      showToast(res.message, 'error');
    }
  } catch (err) { hideSpinner(); showToast('Error: ' + err.message, 'error'); }
}

async function loadSPT() {
  try {
    var res = await callAPI('getSPT', {});
    var tbody = document.getElementById('tbody-spt');
    if (!res.success || !res.data.length) { tbody.innerHTML = '<tr class="no-data"><td colspan="9"><i class="bi bi-inbox" style="font-size:2rem;display:block;margin-bottom:8px"></i>Belum ada data Arsip SPT</td></tr>'; return; }
    tbody.innerHTML = res.data.map(function (d, i) {
      var url = d['URL'] || d['File URL'];
      var lampiran = url ? '<button class="btn-link-custom action-col" onclick="openPreview(\'' + url + '\')"><i class="bi bi-file-earmark-pdf"></i> Lihat</button>' : '<span style="color:var(--text-muted);font-size:.78rem">-</span>';
      var safeData = encodeURIComponent(JSON.stringify(d));
      
      var target = (d['Nama'] && d['Nama'] !== '-') ? d['Nama'] + '<br><span style="font-family:var(--mono);font-size:.76rem;color:var(--text-muted)">' + (d['NIP'] || '') + '</span>' : '-';
      var berangkat = fmtDate(d['Tgl Berangkat'] || d['Tanggal Berangkat']);
      var kembali = fmtDate(d['Tgl Kembali'] || d['Tanggal Kembali']);
      var jadwal = (berangkat !== '-') ? berangkat + '<br>s/d ' + kembali : '-';
      var tglTerbit = d['Tanggal'] || d['CreatedAt'] || d['created'] || '-';
      var tglTerbitFmt = tglTerbit !== '-' ? fmtDate(tglTerbit) : '<span style="color:var(--text-muted);font-size:.78rem">-</span>';

      return '<tr><td>' + (i + 1) + '</td><td><strong>' + esc(d['Nomor SPT']) + '</strong></td><td>' + target + '</td><td style="font-size:.82rem">' + esc(d['Jabatan'] || '-') + '</td><td>' + esc(d['Tujuan'] || '-') + '</td><td style="font-size:.82rem">' + tglTerbitFmt + '</td><td style="font-size:.82rem">' + jadwal + '</td><td>' + lampiran + '</td><td class="action-col" style="display:flex;gap:6px"><button class="btn-primary-custom" title="Cetak SPT" style="padding:5px 10px;" onclick="printDocSPT(\'' + safeData + '\')"><i class="bi bi-printer"></i></button><button class="btn-warning-custom" onclick="openEditModal(\'SPT\', \'' + safeData + '\')"><i class="bi bi-pencil"></i></button><button class="btn-danger-custom" onclick="deleteItem(\'deleteSPT\',\'' + d['ID'] + '\',loadSPT)"><i class="bi bi-trash"></i></button></td></tr>';
    }).join('');
  } catch (err) { /* silent */ }
}

function printDocSPT(encodedData) {
  var d = JSON.parse(decodeURIComponent(encodedData));
  
  // DEFAULT KOP JIKA BELUM ADA DI PENGATURAN
  var k1 = localStorage.getItem('simadun_kop1') || 'PEMERINTAH KABUPATEN MADIUN';
  var k2 = localStorage.getItem('simadun_kop2') || 'INSPEKTORAT';
  var k3 = localStorage.getItem('simadun_kop3') || 'Jalan M.T. Haryono, Caruban, Jawa Timur 63153, Telepon (0351) 453412,\nLaman www.inspektorat.madiunkab.go.id, Pos-el madiunkab.inspektorat@gmail.com';
  var tKota = localStorage.getItem('simadun_ttd_kota') || 'Caruban';
  var tJab = localStorage.getItem('simadun_ttd_jabatan') || 'Inspektur\\nKabupaten Madiun';
  var tNama = localStorage.getItem('simadun_ttd_nama') || 'Joko Lelono, A.P., M.H., CGCAE';
  var tNip = localStorage.getItem('simadun_ttd_nip') || '197306081993111001';

  var logoKiri = localStorage.getItem('simadun_logo_kiri_data');
  var logoKanan = localStorage.getItem('simadun_logo_kanan_data');
  var logoPos = localStorage.getItem('simadun_logo_pos') || 'left';
  var logoSize = localStorage.getItem('simadun_logo_size') || '90';
  var logoKiriSize = localStorage.getItem('simadun_logo_kiri_size') || logoSize;
  var logoKananSize = localStorage.getItem('simadun_logo_kanan_size') || logoSize;
  var fontSize = localStorage.getItem('simadun_font_size') || '11';
  var penutup = localStorage.getItem('simadun_penutup') || 'Demikian surat tugas ini dibuat untuk dilaksanakan dengan penuh tanggung jawab dan dipergunakan sebagaimana mestinya.';
  
  var defaultLogo = BASE_URL + '/assets/icon-512.png';
  var leftImgSrc = logoKiri || (logoPos === 'left' ? defaultLogo : '');
  var rightImgSrc = logoKanan || (logoPos === 'right' ? defaultLogo : '');
  var leftImgHtml = leftImgSrc ? '<img src="' + leftImgSrc + '" style="width:' + (logoKiri ? logoKiriSize : logoSize) + 'px;margin-right:15px;object-fit:contain;" />' : '';
  var rightImgHtml = rightImgSrc ? '<img src="' + rightImgSrc + '" style="width:' + (logoKanan ? logoKananSize : logoSize) + 'px;margin-left:15px;object-fit:contain;" />' : '';

  // Render K3 with line breaks
  var k3Html = k3.replace(/\\n/g, '<br>');
  var tJabHtml = tJab.replace(/\\n/g, '<br>');

  var w = window.open('', '_blank');
  w.document.write(`
    <html><head><title>Cetak SPT - ${d['Nomor SPT'] || ''}</title>
    <style>
      @media print { 
        body { padding: 0 !important; margin: 15mm 20mm !important; } 
        @page { size: A4 portrait; margin: 0; }
      }
      body { font-family: Arial, Helvetica, sans-serif; padding: 40px; margin: 0; line-height: 1.4; color: #000; font-size: ${fontSize}pt; }
      .kop { display: flex; align-items: center; justify-content: center; margin-bottom: 4px; }
      .kop-text { text-align: center; }
      .kop-text h2 { margin: 0; font-size: calc(${fontSize}pt + 4pt); font-weight: normal; }
      .kop-text h1 { margin: 2px 0; font-size: calc(${fontSize}pt + 8pt); font-weight: bold; letter-spacing: 4px; }
      .kop-text p { margin: 0; font-size: calc(${fontSize}pt - 1pt); }
      .kop-border { border-top: 3px solid #000; border-bottom: 1px solid #000; height: 1px; margin-bottom: 25px; }
      
      .title { text-align: center; margin-bottom: 25px; }
      .title h3 { margin: 0; font-size: calc(${fontSize}pt + 2pt); font-weight: normal; letter-spacing: 5px; }
      .title p { margin: 8px 0 0; font-size: ${fontSize}pt; }
      
      .content-table { width: 100%; border-collapse: collapse; border: none; margin-bottom: 15px; }
      .content-table td { padding: 4px 0; vertical-align: top; }
      .col-label { width: 80px; }
      .col-colon { width: 20px; text-align: center; }
      .col-val { text-align: justify; }

      .m-center { text-align: center; margin: 15px 0; }
      .m-p { margin-top: 15px; margin-bottom: 15px; text-align: justify; }
      
      .kepada-grid { display: grid; grid-template-columns: 20px 1fr; gap: 4px 8px; }
      .kepada-item { display: contents; }
      
      .sig-container { display: flex; justify-content: flex-end; margin-top: 30px; }
      .sig-box { width: 350px; }
      .sig-table { width: 100%; border: none; border-collapse: collapse; }
      .sig-table td { padding: 2px 0; vertical-align: top; }
    </style></head><body>
      <div class="kop">
        ${leftImgHtml}
        <div class="kop-text">
          <h2>${k1}</h2>
          <h1>${k2}</h1>
          <p>${k3Html}</p>
        </div>
        ${rightImgHtml}
      </div>
      <div class="kop-border"></div>
      
            <div class="title">
        <h3>S U R A T  T U G A S</h3>
        <p>Nomor : ${d['Nomor SPT'] || '-'}</p>
      </div>

      <table class="content-table">
        <tr>
          <td class="col-label">Dasar</td>
          <td class="col-colon">:</td>
          <td class="col-val">${(d['Keterangan'] || '-').replace(/\n/g, '<br>')}</td>
        </tr>
      </table>

      <div class="m-center" style="letter-spacing: 2px; font-weight: bold; font-size: calc(${fontSize}pt + 1pt);">M E N U G A S K A N</div>

      <table class="content-table">
        <tr>
          <td class="col-label">Kepada</td>
          <td class="col-colon">:</td>
          <td class="col-val">
            <div class="kepada-grid">
              ${(function() {
                var namaArr = (d['Nama'] || '-').split('\n');
                var nipArr = (d['NIP'] || '-').split('\n');
                var jabArr = (d['Jabatan'] || '-').split('\n');
                var result = '';
                for(var i=0; i<namaArr.length; i++) {
                  if(!namaArr[i].trim()) continue;
                  var n = namaArr[i].trim();
                  var id = nipArr[i] ? nipArr[i].trim() : '-';
                  var j = jabArr[i] ? jabArr[i].trim() : '-';
                  result += '<div class="kepada-item"><span>' + (i+1) + '.</span><span><strong>' + n + '</strong><br>NIP ' + id + '<br>' + j + '</span></div><div style="width:100%; height:8px"></div>';
                }
                return result || '<span>-</span>';
              })()}
            </div>
          </td>
        </tr>
        <tr>
          <td class="col-label">Untuk</td>
          <td class="col-colon">:</td>
          <td class="col-val">${(d['Keperluan'] || '-').replace(/\n/g, '<br>')}</td>
        </tr>
        <tr>
          <td class="col-label">Tanggal</td>
          <td class="col-colon">:</td>
          <td class="col-val">${fmtDate(d['Tgl Berangkat'] || d['Tanggal Berangkat'])} sd ${fmtDate(d['Tgl Kembali'] || d['Tanggal Kembali'])}</td>
        </tr>
      </table>

      <p class="m-p">APIP dalam melaksanakan tugas tidak menerima/meminta gratifikasi dan suap.</p>
      <p class="m-p">Demikian Surat Tugas ini dibuat untuk dipergunakan sebagaimana mestinya.</p>
      
      <div class="sig-container">
        <div class="sig-box">
          <table class="sig-table">
            <tr><td style="width:90px">Ditetapkan di</td><td style="width:15px">:</td><td>${tKota}</td></tr>
            <tr><td>Pada tanggal</td><td>:</td><td>${new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}</td></tr>
          </table>
          <p style="margin: 10px 0 0 0;">${tJabHtml}</p>
          <br/><br/><br/><br/>
          <p style="text-decoration: underline; font-weight: bold; margin: 0;">${tNama}</p>
          <p style="margin: 0;">NIP ${tNip}</p>
        </div>
      </div>
    </body></html>
  `);
  w.document.close();
  setTimeout(() => { w.print(); }, 800);
}

function printPage() {
  var currentPage = APP.currentPage;
  var title = 'Laporan Dokumen';
  var tableIdStr = '';
  
  if (currentPage === 'arsip') {
    title = 'LAPORAN REKAPITULASI ARSIP DOKUMEN';
    tableIdStr = 'arsip-table';
  } else if (currentPage === 'surat') {
    if (document.getElementById('content-masuk') && document.getElementById('content-masuk').style.display !== 'none') {
      title = 'LAPORAN REKAPITULASI SURAT MASUK'; tableIdStr = 'table-sm';
    } else if (document.getElementById('content-keluar') && document.getElementById('content-keluar').style.display !== 'none') {
      title = 'LAPORAN REKAPITULASI SURAT KELUAR'; tableIdStr = 'table-sk';
    } else {
      title = 'LAPORAN REKAPITULASI UNDANGAN'; tableIdStr = 'table-ud';
    }
  } else if (currentPage === 'spt') {
    title = 'LAPORAN REKAPITULASI SURAT PERINTAH TUGAS'; tableIdStr = 'table-spt';
  } else {
    showToast('Tidak ada data yang dapat dicetak pada halaman ini.', 'error'); return;
  }
  
  var tableEl = document.getElementById(tableIdStr);
  if (!tableEl) { showToast('Tabel tidak ditemukan.', 'error'); return; }
  
  var clone = tableEl.cloneNode(true);
  // Remove Aksi column (last column)
  var thr = clone.querySelector('thead tr');
  if (thr && thr.lastElementChild) thr.removeChild(thr.lastElementChild);
  clone.querySelectorAll('tbody tr').forEach(function(tr) {
    if (!tr.classList.contains('no-data') && tr.lastElementChild) tr.removeChild(tr.lastElementChild);
  });
  
  var k1 = localStorage.getItem('simadun_kop1') || 'PEMERINTAH KABUPATEN MADIUN';
  var k2 = localStorage.getItem('simadun_kop2') || 'INSPEKTORAT';
  var k3 = localStorage.getItem('simadun_kop3') || 'Jalan M.T. Haryono, Caruban, Jawa Timur 63153, Telepon (0351) 453412,\\nLaman www.inspektorat.madiunkab.go.id, Pos-el madiunkab.inspektorat@gmail.com';
  var kTelp = localStorage.getItem('simadun_kop_telp') || '';
  var logoKiri = localStorage.getItem('simadun_logo_kiri_data') || '';
  var logoKanan = localStorage.getItem('simadun_logo_kanan_data') || '';
  var logoKiriSize = localStorage.getItem('simadun_logo_kiri_size') || '70';
  var logoKananSize = localStorage.getItem('simadun_logo_kanan_size') || '70';
  var defaultLogo = BASE_URL + '/assets/icon-512.png';
  
  var leftImg = logoKiri ? ('<img src="' + logoKiri + '" style="width:' + logoKiriSize + 'px;margin-right:16px;">') 
                        : ('<img src="' + defaultLogo + '" style="width:70px;margin-right:16px;">');
  var rightImg = logoKanan ? ('<img src="' + logoKanan + '" style="width:' + logoKananSize + 'px;margin-left:16px;">') : '';

  var now = new Date().toLocaleDateString('id-ID', { day:'numeric', month:'long', year:'numeric' });
  var w = window.open('', '_blank');
  w.document.write(`
    <html><head><title>Cetak Laporan - SIMADUN</title>
    <style>
      @page { size: A4 landscape; margin: 15mm 15mm 15mm 15mm; }
      body { font-family: 'Times New Roman', Times, serif; padding: 0; color: #000; font-size: 10pt; }
      .kop { display: flex; align-items: center; justify-content: center; border-bottom: 4px solid #000; padding-bottom: 10px; margin-bottom: 2px; }
      .kop-border { border-top: 1px solid #000; margin-bottom: 16px; }
      .kop-text { text-align: center; line-height: 1.2; flex: 1; }
      .kop-text h2 { margin: 0; font-size: 12pt; font-weight: normal; }
      .kop-text h1 { margin: 3px 0; font-size: 16pt; font-weight: bold; }
      .kop-text p { margin: 0; font-size: 9pt; }
      .title { text-align: center; margin-bottom: 14px; }
      .title h3 { margin: 0; font-size: 12pt; text-transform: uppercase; letter-spacing: 1px; }
      .title p { font-size: 9pt; margin: 4px 0 0; }
      table { width: 100%; border-collapse: collapse; font-size: 9pt; }
      th { border: 1px solid #000; padding: 6px 8px; text-align: center; background: #e8e8e8; font-weight: bold; white-space: nowrap; }
      td { border: 1px solid #000; padding: 5px 8px; vertical-align: top; word-break: break-word; }
      tr:nth-child(even) td { background: #fafafa; }
      @media print { 
        body { margin: 0; padding: 0; }
        table { page-break-inside: auto; }
        tr { page-break-inside: avoid; page-break-after: auto; }
      }
    </style></head><body>
      <div class="kop">
        ${leftImg}
        <div class="kop-text">
          <h2>${k1}</h2>
          <h1>${k2}</h1>
          <p>${k3}${kTelp ? ' &mdash; Telp: ' + kTelp : ''}</p>
        </div>
        ${rightImg}
      </div>
      <div class="kop-border"></div>
      <div class="title">
        <h3>${title}</h3>
        <p>Dicetak pada: ${now}</p>
      </div>
      ${clone.outerHTML}
    </body></html>
  `);
  w.document.close();
  setTimeout(() => { w.print(); }, 800);
}

// ══════════════════════════════════════════════════════════
//  UNIVERSAL EDIT
// ══════════════════════════════════════════════════════════
function closeEditModal() {
  document.getElementById('edit-modal').style.display = 'none';
  currentEditData = null;
  document.getElementById('edit-file').value = '';
  document.getElementById('edit-file-info').textContent = '';
}

function openEditModal(sheetType, dataEnc) {
  var d = JSON.parse(decodeURIComponent(dataEnc));
  currentEditData = { sheet: sheetType, id: d['ID'], oriData: d };
  document.getElementById('edit-modal-title').textContent = 'Edit Data ' + sheetType;
  
  var c = document.getElementById('edit-form-container');
  var html = '';
  
  if (sheetType === 'Arsip') {
    html = `<div class="grid-2">
        <div class="form-group"><label>Nama File</label><input type="text" id="ed-arsip-nama" class="form-control-custom" value="${esc(d['Nama File'])}"></div>
        <div class="form-group"><label>Kategori</label><input type="text" id="ed-arsip-kat" class="form-control-custom" value="${esc(d['Kategori'])}"></div>
      </div>
      <div class="form-group"><label>Deskripsi</label><input type="text" id="ed-arsip-desk" class="form-control-custom" value="${esc(d['Deskripsi'])}"></div>
      <div class="form-group"><label>Tanggal Arsip Asli</label><input type="date" id="ed-arsip-tgl" class="form-control-custom" value="${d['Tanggal Arsip'] ? d['Tanggal Arsip'].substring(0,10) : ''}"></div>`;
  } else if (sheetType === 'Surat Masuk') {
    html = `<div class="grid-2">
        <div class="form-group"><label>Nomor Surat</label><input type="text" id="ed-sm-nomor" class="form-control-custom" value="${esc(d['Nomor Surat'])}"></div>
        <div class="form-group"><label>Tanggal</label><input type="date" id="ed-sm-tgl" class="form-control-custom" value="${d['Tanggal']}"></div>
      </div>
      <div class="grid-2">
        <div class="form-group"><label>Pengirim</label><input type="text" id="ed-sm-pengirim" class="form-control-custom" value="${esc(d['Pengirim'])}"></div>
        <div class="form-group"><label>Kategori</label><input type="text" id="ed-sm-kat" class="form-control-custom" value="${esc(d['Kategori'])}"></div>
      </div>
      <div class="form-group"><label>Perihal</label><input type="text" id="ed-sm-perihal" class="form-control-custom" value="${esc(d['Perihal'])}"></div>
      <div class="form-group"><label>Catatan</label><input type="text" id="ed-sm-catatan" class="form-control-custom" value="${esc(d['Catatan'])}"></div>`;
  } else if (sheetType === 'Surat Keluar') {
    html = `<div class="grid-2">
        <div class="form-group"><label>Nomor Surat</label><input type="text" id="ed-sk-nomor" class="form-control-custom" value="${esc(d['Nomor Surat'])}"></div>
        <div class="form-group"><label>Tanggal</label><input type="date" id="ed-sk-tgl" class="form-control-custom" value="${d['Tanggal']}"></div>
      </div>
      <div class="grid-2">
        <div class="form-group"><label>Tujuan</label><input type="text" id="ed-sk-tujuan" class="form-control-custom" value="${esc(d['Tujuan'])}"></div>
        <div class="form-group"><label>Kategori</label><input type="text" id="ed-sk-kat" class="form-control-custom" value="${esc(d['Kategori'])}"></div>
      </div>
      <div class="form-group"><label>Perihal</label><input type="text" id="ed-sk-perihal" class="form-control-custom" value="${esc(d['Perihal'])}"></div>
      <div class="form-group"><label>Catatan</label><input type="text" id="ed-sk-catatan" class="form-control-custom" value="${esc(d['Catatan'])}"></div>`;
  } else if (sheetType === 'Undangan') {
    html = `<div class="grid-2">
        <div class="form-group"><label>Nomor Surat</label><input type="text" id="ed-ud-nomor" class="form-control-custom" value="${esc(d['Nomor Surat'])}"></div>
        <div class="form-group"><label>Tanggal</label><input type="date" id="ed-ud-tgl" class="form-control-custom" value="${d['Tanggal']}"></div>
      </div>
      <div class="grid-2">
        <div class="form-group"><label>Penyelenggara</label><input type="text" id="ed-ud-peny" class="form-control-custom" value="${esc(d['Penyelenggara'])}"></div>
        <div class="form-group"><label>Lokasi</label><input type="text" id="ed-ud-lokasi" class="form-control-custom" value="${esc(d['Lokasi'])}"></div>
      </div>
      <div class="form-group"><label>Perihal</label><input type="text" id="ed-ud-perihal" class="form-control-custom" value="${esc(d['Perihal'])}"></div>
      <div class="form-group"><label>Catatan</label><input type="text" id="ed-ud-catatan" class="form-control-custom" value="${esc(d['Catatan'])}"></div>`;
  } else if (sheetType === 'SPT') {
    html = `<div class="grid-2">
        <div class="form-group"><label>Nomor SPT</label><input type="text" id="ed-spt-nomor" class="form-control-custom" value="${esc(d['Nomor SPT'])}"></div>
        <div class="form-group"><label>Tujuan</label><input type="text" id="ed-spt-tujuan" class="form-control-custom" value="${esc(d['Tujuan'])}"></div>
      </div>
      <div class="grid-2">
        <div class="form-group"><label>Nama Target</label><input type="text" id="ed-spt-nama" class="form-control-custom" value="${esc(d['Nama'])}"></div>
        <div class="form-group"><label>NIP Target</label><input type="text" id="ed-spt-nip" class="form-control-custom" value="${esc(d['NIP'])}"></div>
      </div>
      <div class="form-group"><label>Jabatan Target</label><input type="text" id="ed-spt-jab" class="form-control-custom" value="${esc(d['Jabatan'])}"></div>
      <div class="form-group"><label>Keperluan</label><input type="text" id="ed-spt-kep" class="form-control-custom" value="${esc(d['Keperluan'])}"></div>
      <div class="form-group"><label>Keterangan</label><input type="text" id="ed-spt-ket" class="form-control-custom" value="${esc(d['Keterangan'])}"></div>`;
  }
  
  c.innerHTML = html;
  document.getElementById('edit-modal').style.display = 'flex';
}

async function submitEditData() {
  if (!currentEditData) return;
  var payload = { id: currentEditData.id, sheetName: currentEditData.sheet, data: {} };
  var t = currentEditData.sheet;
  
  if (t === 'Arsip') { payload.data = { namaFile: v('ed-arsip-nama'), kategori: v('ed-arsip-kat'), deskripsi: v('ed-arsip-desk'), tglArsip: v('ed-arsip-tgl') }; }
  else if (t === 'Surat Masuk') { payload.data = { nomorSurat: v('ed-sm-nomor'), tanggal: v('ed-sm-tgl'), pengirim: v('ed-sm-pengirim'), perihal: v('ed-sm-perihal'), catatan: v('ed-sm-catatan'), kategori: v('ed-sm-kat') }; }
  else if (t === 'Surat Keluar') { payload.data = { nomorSurat: v('ed-sk-nomor'), tanggal: v('ed-sk-tgl'), tujuan: v('ed-sk-tujuan'), perihal: v('ed-sk-perihal'), catatan: v('ed-sk-catatan'), kategori: v('ed-sk-kat') }; }
  else if (t === 'Undangan') { payload.data = { nomorSurat: v('ed-ud-nomor'), tanggal: v('ed-ud-tgl'), penyelenggara: v('ed-ud-peny'), perihal: v('ed-ud-perihal'), catatan: v('ed-ud-catatan'), lokasi: v('ed-ud-lokasi') }; }
  else if (t === 'SPT') { payload.data = { nomorSpt: v('ed-spt-nomor'), nama: v('ed-spt-nama'), nip: v('ed-spt-nip'), jabatan: v('ed-spt-jab'), tujuan: v('ed-spt-tujuan'), keperluan: v('ed-spt-kep'), keterangan: v('ed-spt-ket') }; }
  
  var fileEl = document.getElementById('edit-file');
  var fd = null;
  if (fileEl && fileEl.files[0]) fd = await readFileAsBase64(fileEl.files[0]);
  
  showSpinner('Menyimpan perubahan data...');
  try {
    var realSheetName = t;
    if (t==='Surat Masuk' || t==='Surat Keluar' || t==='Undangan' || t==='SPT' || t==='Arsip') realSheetName = t;
    var res = await callAPI('updateRow', { sheetName: realSheetName, id: payload.id, data: payload.data, fileData: fd });
    hideSpinner();
    if (res.success) {
      showToast('Data berhasil diperbarui!', 'success');
      closeEditModal();
      if (t === 'Arsip') loadArsip();
      else if (t === 'Surat Masuk') loadSuratMasuk();
      else if (t === 'Surat Keluar') loadSuratKeluar();
      else if (t === 'Undangan') loadUndangan();
      else if (t === 'SPT') loadSPT();
    } else {
      showToast(res.message, 'error');
    }
  } catch(e) {
    hideSpinner(); showToast('Error: ' + e.message, 'error');
  }
}

// ══════════════════════════════════════════════════════════
//  PENGATURAN (Setup DB, Users & Cetak)
// ══════════════════════════════════════════════════════════
function setPgTab(tabId) {
  document.querySelectorAll('#page-pengaturan .custom-tab').forEach(function (t) { t.classList.remove('active'); });
  document.querySelectorAll('#page-pengaturan .tab-content').forEach(function (c) { c.style.display = 'none'; });
  
  var activeBtn = document.getElementById('tab-' + tabId);
  var contentBlock = document.getElementById('content-' + tabId);
  if (activeBtn) activeBtn.classList.add('active');
  if (contentBlock) contentBlock.style.display = 'block';
  
  if (tabId === 'pg-cetak') loadKopSettings();
}

function loadKopSettings() {
  var fields = {
    'set-kop1': ['simadun_kop1', 'PEMERINTAH KABUPATEN MADIUN'],
    'set-kop2': ['simadun_kop2', 'INSPEKTORAT'],
    'set-kop3': ['simadun_kop3', 'Jalan M.T. Haryono, Caruban, Jawa Timur 63153, Telepon (0351) 453412,\\nLaman www.inspektorat.madiunkab.go.id, Pos-el madiunkab.inspektorat@gmail.com'],
    'set-kop-telp': ['simadun_kop_telp', ''],
    'set-ttd-kota': ['simadun_ttd_kota', 'Madiun'],
    'set-ttd-jabatan': ['simadun_ttd_jabatan', 'Inspektur Kabupaten Madiun,'],
    'set-ttd-nama': ['simadun_ttd_nama', '________________________'],
    'set-ttd-nip': ['simadun_ttd_nip', '........................................'],
    'set-logo-kiri-size': ['simadun_logo_kiri_size', '90'],
    'set-logo-kanan-size': ['simadun_logo_kanan_size', '90'],
    'set-logo-size': ['simadun_logo_size', '90'],
    'set-font-size': ['simadun_font_size', '12'],
    'set-penutup': ['simadun_penutup', 'Demikian Surat Perintah Tugas ini dibuat untuk dilaksanakan dengan penuh tanggung jawab.']
  };
  Object.keys(fields).forEach(function(id) {
    var el = document.getElementById(id);
    if (el) el.value = localStorage.getItem(fields[id][0]) || fields[id][1];
  });
  // Select fields
  var logoPos = document.getElementById('set-logo-pos');
  if (logoPos) logoPos.value = localStorage.getItem('simadun_logo_pos') || 'left';
  // Restore logo previews
  var kiriData = localStorage.getItem('simadun_logo_kiri_data');
  if (kiriData) {
    var kiriPrev = document.getElementById('set-logo-kiri-preview');
    var kiriImg = document.getElementById('logo-kiri-img');
    if (kiriImg) kiriImg.src = kiriData;
    if (kiriPrev) kiriPrev.style.display = 'block';
  }
  var kananData = localStorage.getItem('simadun_logo_kanan_data');
  if (kananData) {
    var kananPrev = document.getElementById('set-logo-kanan-preview');
    var kananImg = document.getElementById('logo-kanan-img');
    if (kananImg) kananImg.src = kananData;
    if (kananPrev) kananPrev.style.display = 'block';
  }
  updateSptPreview();
}

function handleLogoUpload(side, inputEl) {
  var file = inputEl.files[0];
  if (!file) return;
  if (file.size > 2 * 1024 * 1024) { showToast('Ukuran logo maksimal 2MB.', 'error'); return; }
  var reader = new FileReader();
  reader.onload = function(e) {
    var dataUrl = e.target.result;
    localStorage.setItem('simadun_logo_' + side + '_data', dataUrl);
    var sizeEl = document.getElementById('set-logo-' + side + '-size');
    if (sizeEl) localStorage.setItem('simadun_logo_' + side + '_size', sizeEl.value || '90');
    var prevContainer = document.getElementById('set-logo-' + side + '-preview');
    var prevImg = document.getElementById('logo-' + side + '-img');
    if (prevImg) prevImg.src = dataUrl;
    if (prevContainer) prevContainer.style.display = 'block';
    showToast('Logo ' + side + ' berhasil dimuat!', 'success');
    updateSptPreview();
  };
  reader.readAsDataURL(file);
}

function clearLogo(side) {
  localStorage.removeItem('simadun_logo_' + side + '_data');
  localStorage.removeItem('simadun_logo_' + side + '_size');
  var fileInput = document.getElementById('set-logo-' + side + '-file');
  if (fileInput) fileInput.value = '';
  var prevContainer = document.getElementById('set-logo-' + side + '-preview');
  var prevImg = document.getElementById('logo-' + side + '-img');
  if (prevImg) prevImg.src = '';
  if (prevContainer) prevContainer.style.display = 'none';
  showToast('Logo ' + side + ' berhasil dihapus.', 'info');
  updateSptPreview();
}


function saveKopSettings() {
  localStorage.setItem('simadun_kop1', v('set-kop1'));
  localStorage.setItem('simadun_kop2', v('set-kop2'));
  localStorage.setItem('simadun_kop3', v('set-kop3'));
  localStorage.setItem('simadun_kop_telp', v('set-kop-telp'));
  localStorage.setItem('simadun_ttd_kota', v('set-ttd-kota'));
  localStorage.setItem('simadun_ttd_jabatan', v('set-ttd-jabatan'));
  localStorage.setItem('simadun_ttd_nama', v('set-ttd-nama'));
  localStorage.setItem('simadun_ttd_nip', v('set-ttd-nip'));
  localStorage.setItem('simadun_logo_pos', v('set-logo-pos'));
  localStorage.setItem('simadun_logo_size', v('set-logo-size'));
  localStorage.setItem('simadun_font_size', v('set-font-size'));
  localStorage.setItem('simadun_penutup', v('set-penutup'));
  showToast('Pengaturan format Kop Surat & TTD berhasil disimpan.', 'success');
  updateSptPreview();
}

function resetKopSettings() {
  if (!confirm('Reset semua pengaturan ke default?')) return;
  ['simadun_kop1','simadun_kop2','simadun_kop3','simadun_kop_telp','simadun_ttd_kota','simadun_ttd_jabatan','simadun_ttd_nama','simadun_ttd_nip','simadun_logo_pos','simadun_logo_size','simadun_font_size','simadun_penutup'].forEach(function(k) { localStorage.removeItem(k); });
  loadKopSettings();
  updateSptPreview();
  showToast('Pengaturan dikembalikan ke default.', 'info');
}

function updateSptPreview() {
  var k1 = v('set-kop1') || 'PEMERINTAH KABUPATEN MADIUN';
  var k2 = v('set-kop2') || 'INSPEKTORAT';
  var k3 = v('set-kop3') || 'Jalan M.T. Haryono, Caruban, Jawa Timur 63153, Telepon (0351) 453412,\\nLaman www.inspektorat.madiunkab.go.id, Pos-el madiunkab.inspektorat@gmail.com';
  var kTelp = v('set-kop-telp') || '';
  var tKota = v('set-ttd-kota') || 'Caruban';
  var tJab = v('set-ttd-jabatan') || 'Inspektur\\nKabupaten Madiun';
  var tNama = v('set-ttd-nama') || 'Joko Lelono, A.P., M.H., CGCAE';
  var tNip = v('set-ttd-nip') || '197306081993111001';
  var logoPos = (document.getElementById('set-logo-pos') ? document.getElementById('set-logo-pos').value : 'left');
  var logoSize = v('set-logo-size') || '90';
  var fontSize = v('set-font-size') || '11';
  var penutup = v('set-penutup') || 'Demikian surat tugas ini dibuat untuk dilaksanakan dengan penuh tanggung jawab dan dipergunakan sebagaimana mestinya.';

  var logoKiri = localStorage.getItem('simadun_logo_kiri_data');
  var logoKanan = localStorage.getItem('simadun_logo_kanan_data');
  var logoKiriSize = v('set-logo-kiri-size') || localStorage.getItem('simadun_logo_kiri_size') || logoSize;
  var logoKananSize = v('set-logo-kanan-size') || localStorage.getItem('simadun_logo_kanan_size') || logoSize;
  var defaultLogo = BASE_URL + '/assets/icon-512.png';

  var leftImgSrc = logoKiri || (logoPos === 'left' ? defaultLogo : '');
  var rightImgSrc = logoKanan || (logoPos === 'right' ? defaultLogo : '');

  var leftImgHtml = leftImgSrc ? '<img src="' + leftImgSrc + '" style="width:' + (logoKiri ? logoKiriSize : logoSize) + 'px;margin-right:15px;object-fit:contain;" />' : '';
  var rightImgHtml = rightImgSrc ? '<img src="' + rightImgSrc + '" style="width:' + (logoKanan ? logoKananSize : logoSize) + 'px;margin-left:15px;object-fit:contain;" />' : '';
  
  var k3Html = k3.replace(/\\n/g, '<br>') + (kTelp ? '<br>Telp: ' + kTelp : '');
  var tJabHtml = tJab.replace(/\\n/g, '<br>');

  var kopHtml = leftImgHtml + '<div class="kop-text"><h2>' + k1 + '</h2><h1>' + k2 + '</h1><p>' + k3Html + '</p></div>' + rightImgHtml;

  var html = '<html><head><style>' +
    'body { font-family: Arial, Helvetica, sans-serif; padding:20px; line-height: 1.4; color: #000; font-size: ' + fontSize + 'pt; }' +
    '.kop { display: flex; align-items: center; justify-content: center; margin-bottom: 4px; }' +
    '.kop-text { text-align: center; }' +
    '.kop-text h2 { margin: 0; font-size: calc(' + fontSize + 'pt + 4pt); font-weight: normal; }' +
    '.kop-text h1 { margin: 2px 0; font-size: calc(' + fontSize + 'pt + 8pt); font-weight: bold; letter-spacing: 4px; }' +
    '.kop-text p { margin: 0; font-size: calc(' + fontSize + 'pt - 1pt); }' +
    '.kop-border { border-top: 3px solid #000; border-bottom: 1px solid #000; height: 1px; margin-bottom: 25px; }' +
    '.title { text-align: center; margin-bottom: 25px; }' +
    '.title h3 { margin: 0; font-size: calc(' + fontSize + 'pt + 2pt); font-weight: normal; letter-spacing: 5px; }' +
    '.title p { margin: 8px 0 0; font-size: ' + fontSize + 'pt; }' +
    '.content-table { width: 100%; border-collapse: collapse; border: none; margin-bottom: 15px; }' +
    '.content-table td { padding: 4px 0; vertical-align: top; }' +
    '.col-label { width: 80px; }' +
    '.col-colon { width: 20px; text-align: center; }' +
    '.col-val { text-align: justify; }' +
    '.m-center { text-align: center; margin: 15px 0; }' +
    '.m-p { margin-top: 15px; margin-bottom: 15px; text-align: justify; }' +
    '.kepada-grid { display: grid; grid-template-columns: 20px 1fr; gap: 4px 8px; }' +
    '.kepada-item { display: contents; }' +
    '.sig-container { display: flex; justify-content: flex-end; margin-top: 30px; }' +
    '.sig-box { width: 350px; }' +
    '.sig-table { width: 100%; border: none; border-collapse: collapse; }' +
    '.sig-table td { padding: 2px 0; vertical-align: top; }' +
    '</style></head><body>' +
    '<div class="kop">' + kopHtml + '</div>' +
    '<div class="kop-border"></div>' +
    '<div class="title"><h3>SURAT TUGAS</h3><p>NOMOR : [NOMOR SURAT]</p></div>' +
    '<table class="content-table">' +
    '<tr><td class="col-label">Dasar</td><td class="col-colon">:</td><td class="col-val">Surat Kepala Kejaksaan Negeri Kabupaten Madiun ...</td></tr>' +
    '</table>' +
    '<div class="m-center">MEMERINTAHKAN:</div>' +
    '<table class="content-table">' +
    '<tr><td class="col-label">Kepada</td><td class="col-colon">:</td><td class="col-val">' +
    '<div class="kepada-grid"><div class="kepada-item"><span>1.</span><span>NAMA PETUGAS<br>NIP 19800101 200501 1 001<br>Auditor Ahli Madya</span></div></div>' +
    '</td></tr>' +
    '<tr><td class="col-label">Untuk</td><td class="col-colon">:</td><td class="col-val">Melakukan pemeriksaan...</td></tr>' +
    '<tr><td class="col-label">Tanggal</td><td class="col-colon">:</td><td class="col-val">1 Januari 2026 sd 3 Januari 2026</td></tr>' +
    '</table>' +
    '<p class="m-p">APIP dalam melaksanakan tugas tidak menerima/meminta gratifikasi dan suap.</p>' +
    '<p class="m-p">' + penutup + '</p>' +
    '<div class="sig-container"><div class="sig-box">' +
    '<table class="sig-table">' +
    '<tr><td style="width:90px">Ditetapkan di</td><td style="width:15px">:</td><td>' + tKota + '</td></tr>' +
    '<tr><td>Pada tanggal</td><td>:</td><td>' + new Date().toLocaleDateString('id-ID', {day:'numeric',month:'long',year:'numeric'}) + '</td></tr>' +
    '</table>' +
    '<p style="margin: 10px 0 0 0;">' + tJabHtml + '</p>' +
    '<br><br><br><br>' +
    '<p style="text-decoration:underline;font-weight:bold;margin:0">' + tNama + '</p>' +
    '<p style="margin:0">NIP ' + tNip + '</p>' +
    '</div></div>' +
    '</body></html>';

  var frame = document.getElementById('spt-preview-frame');
  if (frame) {
    frame.srcdoc = html;
  }
}

function printSptPreview() {
  var frame = document.getElementById('spt-preview-frame');
  if (frame && frame.contentWindow) {
    frame.contentWindow.print();
  }
}

async function setupDatabase() {
  if (!confirm('Tindakan ini akan menginisialisasi ulang Sheet pada Google Spreadsheet Anda. Lanjutkan?')) return;
  showSpinner('Inisialisasi Database...');
  try {
    var res = await callAPI('setupDb', {});
    hideSpinner();
    if (res.success) {
      showToast('Database berhasil diinisialisasi!', 'success');
    } else {
      showToast(res.message || 'Gagal inisialisasi database.', 'error');
    }
  } catch (err) { hideSpinner(); showToast('Error: ' + err.message, 'error'); }
}
async function submitAddUser() {
  var data = { nama: v('new-nama'), username: v('new-username'), password: document.getElementById('new-password').value, role: v('new-role') };
  if (!data.nama || !data.username || !data.password) { showToast('Semua field wajib diisi.', 'error'); return; }
  showSpinner('Menambahkan pengguna...');
  try {
    var res = await callAPI('addUser', { data: data });
    hideSpinner();
    if (res.success) { showToast(res.message, 'success'); resetFields(['new-nama', 'new-username', 'new-password']); loadUsers(); }
    else showToast(res.message, 'error');
  } catch (err) { hideSpinner(); showToast('Error: ' + err.message, 'error'); }
}

async function submitChangePassword() {
  var u = v('chg-username');
  var ow = document.getElementById('chg-old').value;
  var nw = document.getElementById('chg-new').value;
  var cf = document.getElementById('chg-confirm').value;
  if (!ow || !nw || !cf) { showToast('Semua field wajib diisi.', 'error'); return; }
  if (nw !== cf) { showToast('Password baru tidak cocok.', 'error'); return; }
  if (nw.length < 6) { showToast('Password baru minimal 6 karakter.', 'error'); return; }
  showSpinner('Mengubah password...');
  try {
    var res = await callAPI('changePassword', { username: u, oldPassword: ow, newPassword: nw });
    hideSpinner();
    if (res.success) { showToast(res.message, 'success'); resetFields(['chg-old', 'chg-new', 'chg-confirm']); }
    else showToast(res.message, 'error');
  } catch (err) { hideSpinner(); showToast('Error: ' + err.message, 'error'); }
}

async function loadUsers() {
  try {
    var res = await callAPI('getUsers', {});
    var tbody = document.getElementById('tbody-users');
    if (!res.success || !res.data.length) { tbody.innerHTML = '<tr class="no-data"><td colspan="6">Tidak ada pengguna.</td></tr>'; return; }
    tbody.innerHTML = res.data.map(function (d, i) {
      var isCurrent = APP.user && d.username === APP.user.username;
      return '<tr><td>' + (i + 1) + '</td><td><div style="display:flex;align-items:center;gap:10px"><div class="user-table-avatar">' + d.nama.charAt(0).toUpperCase() + '</div>' + esc(d.nama) + '</div></td><td><span style="font-family:var(--mono);font-size:.82rem">' + esc(d.username) + '</span></td><td><span class="badge-cat masuk">' + esc(d.role) + '</span></td><td>' + fmtDate(d.created || d.CreatedAt) + '</td><td class="action-col">' + (isCurrent ? '<span style="font-size:.78rem;color:var(--text-muted)">Akun aktif</span>' : '<button class="btn-danger-custom" onclick="deleteItem(\'deleteUser\',\'' + d.id + '\',loadUsers)"><i class="bi bi-person-x"></i> Hapus</button>') + '</td></tr>';
    }).join('');
  } catch (err) { /* silent */ }
}

async function doSetupDb() {
  if (!confirm('Tindakan ini akan menginisialisasi ulang sistem dan Sheet pada spreadsheet Anda. Lanjutkan?')) return;
  showSpinner('Inisialisasi Database...');
  try {
    var res = await callAPI('setupDb', {});
    hideSpinner();
    if (res.success) {
      showToast(res.message, 'success');
    } else {
      showToast(res.message, 'error');
    }
  } catch (err) {
    hideSpinner(); showToast('Network Error: ' + err.message, 'error');
  }
}

// ══════════════════════════════════════════════════════════
//  DELETE GENERIC
// ══════════════════════════════════════════════════════════
async function deleteItem(action, id, reloadFn) {
  if (!confirm('Yakin ingin menghapus data ini?')) return;
  showSpinner('Menghapus data...');
  try {
    var res = await callAPI(action, { id: id });
    hideSpinner();
    if (res.success) { showToast(res.message, 'success'); if (reloadFn) reloadFn(); }
    else showToast(res.message, 'error');
  } catch (err) { hideSpinner(); showToast('Error: ' + err.message, 'error'); }
}

// ══════════════════════════════════════════════════════════
//  TABS
// ══════════════════════════════════════════════════════════
function setTab(tabId) {
  document.querySelectorAll('.custom-tab').forEach(function (t) { t.classList.remove('active'); });
  document.querySelectorAll('.tab-content').forEach(function (c) { c.style.display = 'none'; });
  var tabMap = { 'tab-masuk': 'content-masuk', 'tab-keluar': 'content-keluar', 'tab-undangan': 'content-undangan' };
  document.getElementById(tabId).classList.add('active');
  var content = tabMap[tabId];
  if (content) document.getElementById(content).style.display = 'block';
}

// ══════════════════════════════════════════════════════════
//  TOGGLE PANEL & COLLAPSE
// ══════════════════════════════════════════════════════════
function togglePanel(id) {
  var el = document.getElementById(id);
  if (!el) return;
  el.style.display = el.style.display === 'none' ? 'block' : 'none';
}

function toggleCollapse(contentId, triggerEl) {
  var content = document.getElementById(contentId);
  if (!content) return;
  if (content.classList.contains('collapsed')) {
    content.classList.remove('collapsed');
    triggerEl.classList.remove('collapsed');
  } else {
    content.classList.add('collapsed');
    triggerEl.classList.add('collapsed');
  }
}

// ══════════════════════════════════════════════════════════
//  CETAK REPORT
// ══════════════════════════════════════════════════════════
function printPage() {
  window.print();
}

// ══════════════════════════════════════════════════════════
//  PREVIEW MODAL (IFRAME)
// ══════════════════════════════════════════════════════════
function openPreview(url) {
  var overlay = document.getElementById('preview-overlay');
  var frame = document.getElementById('preview-frame');
  var previewUrl = url;
  if (url.includes('/view')) {
    previewUrl = url.replace(/\/view.*$/, '/preview');
  } else if (!url.includes('/preview')) {
    // If it's a raw google drive link, try to append preview
    previewUrl = url + '/preview';
  }
  frame.src = previewUrl;
  overlay.classList.add('active');
}
function closePreview() {
  document.getElementById('preview-overlay').classList.remove('active');
  document.getElementById('preview-frame').src = '';
}

// ══════════════════════════════════════════════════════════
//  TABLE FILTER
// ══════════════════════════════════════════════════════════
function filterTable(tableId, query) {
  var q = query.toLowerCase();
  document.querySelectorAll('#' + tableId + ' tbody tr:not(.no-data)').forEach(function (row) {
    row.style.display = row.textContent.toLowerCase().includes(q) ? '' : 'none';
  });
}

// ══════════════════════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════════════════════
function esc(str) {
  if (!str) return '-';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function fmtDate(val) {
  if (!val || val === '-') return '-';
  try { var d = new Date(val); if (isNaN(d)) return String(val); return d.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' }); } catch (e) { return String(val); }
}
function v(id) {
  var el = document.getElementById(id);
  return el ? el.value.trim() : '';
}
function resetFields(ids) {
  ids.forEach(function (id) { var el = document.getElementById(id); if (el) el.value = ''; });
}

// ══════════════════════════════════════════════════════════
//  DRAG & DROP EVENTS & INITS
// ══════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', function () {
  initTheme();
  
  ['login-username', 'login-password'].forEach(function (id) {
    var el = document.getElementById(id);
    if (el) el.addEventListener('keydown', function (e) { if (e.key === 'Enter') doLogin(); });
  });

  document.querySelectorAll('.file-drop').forEach(function (zone) {
    zone.addEventListener('dragover', function (e) { e.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', function () { zone.classList.remove('dragover'); });
    zone.addEventListener('drop', function (e) {
      e.preventDefault();
      zone.classList.remove('dragover');
      var fileInput = zone.querySelector('input[type="file"]');
      var infoEl = zone.querySelector('.drop-name');
      if (fileInput && e.dataTransfer.files[0]) {
        var dt = new DataTransfer();
        dt.items.add(e.dataTransfer.files[0]);
        fileInput.files = dt.files;
        var f = e.dataTransfer.files[0];
        if (infoEl) infoEl.textContent = '📎 ' + f.name + ' (' + (f.size / 1024).toFixed(1) + ' KB)';
      }
    });
  });
});
