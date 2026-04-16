// pages/api/[action].js
import { google } from 'googleapis';
import crypto from 'crypto';

/**
 * SIMADUN - Sistem Informasi Manajemen Arsip Dokumen Inspektorat Madiun
 * Vercel Backend implementation replacing Google Apps Script
 */

// Global config mapping from Code.gs
const CONFIG = {
  SHEETS: {
    USERS: 'Users',
    SURAT_MASUK: 'Surat Masuk',
    SURAT_KELUAR: 'Surat Keluar',
    UNDANGAN: 'Undangan',
    SPT: 'SPT',
    ARSIP: 'Arsip'
  }
};

const ARSIP_SALT = 'ARSIP_SALT_2024';

export default async function handler(req, res) {
  // CORS setup
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  if (req.method === 'OPTIONS') return res.status(200).end();

  const { action } = req.query;
  const payload = req.body.payload || {};

  try {
    const auth = await getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    const drive = google.drive({ version: 'v3', auth });
    const spreadsheetId = process.env.GOOGLE_SHEET_ID;
    const rootFolderId = process.env.GOOGLE_DRIVE_FOLDER_ID;

    if (!spreadsheetId) throw new Error('GOOGLE_SHEET_ID is missing');

    let result;
    switch (action) {
      case 'login':
        result = await loginUser(sheets, spreadsheetId, payload.username, payload.password);
        break;
      case 'getDashboard':
        result = await getDashboardData(sheets, spreadsheetId);
        break;
      case 'saveArsip':
        result = await saveArsip(sheets, drive, spreadsheetId, rootFolderId, payload.data, payload.fileData);
        break;
      case 'getArsip':
        result = await getSheetData(sheets, spreadsheetId, CONFIG.SHEETS.ARSIP);
        break;
      case 'deleteArsip':
        result = await deleteRowById(sheets, spreadsheetId, CONFIG.SHEETS.ARSIP, payload.id);
        break;
      case 'saveSuratMasuk':
        result = await saveSurat(sheets, drive, spreadsheetId, rootFolderId, CONFIG.SHEETS.SURAT_MASUK, 'Surat Masuk', payload.data, payload.fileData);
        break;
      case 'getSuratMasuk':
        result = await getSheetData(sheets, spreadsheetId, CONFIG.SHEETS.SURAT_MASUK);
        break;
      case 'deleteSuratMasuk':
        result = await deleteRowById(sheets, spreadsheetId, CONFIG.SHEETS.SURAT_MASUK, payload.id);
        break;
      case 'saveSuratKeluar':
        result = await saveSurat(sheets, drive, spreadsheetId, rootFolderId, CONFIG.SHEETS.SURAT_KELUAR, 'Surat Keluar', payload.data, payload.fileData);
        break;
      case 'getSuratKeluar':
        result = await getSheetData(sheets, spreadsheetId, CONFIG.SHEETS.SURAT_KELUAR);
        break;
      case 'deleteSuratKeluar':
        result = await deleteRowById(sheets, spreadsheetId, CONFIG.SHEETS.SURAT_KELUAR, payload.id);
        break;
      case 'saveUndangan':
        result = await saveSurat(sheets, drive, spreadsheetId, rootFolderId, CONFIG.SHEETS.UNDANGAN, 'Undangan', payload.data, payload.fileData);
        break;
      case 'getUndangan':
        result = await getSheetData(sheets, spreadsheetId, CONFIG.SHEETS.UNDANGAN);
        break;
      case 'deleteUndangan':
        result = await deleteRowById(sheets, spreadsheetId, CONFIG.SHEETS.UNDANGAN, payload.id);
        break;
      case 'saveSPT':
        result = await saveSPT(sheets, drive, spreadsheetId, rootFolderId, payload.data, payload.fileData);
        break;
      case 'getSPT':
        result = await getSheetData(sheets, spreadsheetId, CONFIG.SHEETS.SPT);
        break;
      case 'deleteSPT':
        result = await deleteRowById(sheets, spreadsheetId, CONFIG.SHEETS.SPT, payload.id);
        break;
      case 'getUsers':
        result = await getUsers(sheets, spreadsheetId);
        break;
      case 'addUser':
        result = await addUser(sheets, spreadsheetId, payload.data);
        break;
      case 'deleteUser':
        result = await deleteRowById(sheets, spreadsheetId, CONFIG.SHEETS.USERS, payload.id);
        break;
      case 'changePassword':
        result = await changePassword(sheets, spreadsheetId, payload.username, payload.oldPassword, payload.newPassword);
        break;
      case 'setupDb':
        result = await setupDb(sheets, spreadsheetId);
        break;
      case 'updateRow':
        result = await updateRowAndFile(sheets, drive, spreadsheetId, rootFolderId, payload.sheetName, payload.id, payload.data, payload.fileData);
        break;
      default:
        return res.status(404).json({ success: false, message: `Unknown action: ${action}` });
    }

    return res.status(200).json(result);
  } catch (err) {
    console.error('API Error:', err);
    return res.status(500).json({ success: false, message: 'Server error: ' + err.message });
  }
}

// --------------------------------------------------------------------------
// HELPERS
// --------------------------------------------------------------------------

async function getAuthClient() {
  const keyJson = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
  if (!keyJson) throw new Error('GOOGLE_SERVICE_ACCOUNT_KEY is missing');

  const credentials = JSON.parse(keyJson);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive.file',
      'https://www.googleapis.com/auth/drive'
    ]
  });
  return auth.getClient();
}

function hashPassword(password) {
  return crypto.createHash('sha256').update(password + ARSIP_SALT).digest('hex');
}

function generateId() {
  return 'ID' + Date.now() + Math.floor(Math.random() * 9999);
}

// --------------------------------------------------------------------------
// BUSINESS LOGIC
// --------------------------------------------------------------------------

async function loginUser(sheets, spreadsheetId, username, password) {
  const data = await getSheetData(sheets, spreadsheetId, CONFIG.SHEETS.USERS);
  if (!data.success) return data;

  const hashed = hashPassword(password);
  const user = data.data.find(u => u.Username === username && u.Password === hashed);

  if (user) {
    return {
      success: true,
      message: 'Login berhasil',
      user: { id: user.ID, username: user.Username, nama: user.Nama, role: user.Role }
    };
  }
  return { success: false, message: 'Username atau password salah.' };
}

async function getDashboardData(sheets, spreadsheetId) {
  const getCount = async (sheetName) => {
    const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${sheetName}!A:A` });
    return res.data.values ? res.data.values.length - 1 : 0;
  };

  // Arsip specific counters for SIMADUN
  const arsipRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${CONFIG.SHEETS.ARSIP}!A:K` });
  const arsipRows = arsipRes.data.values || [];
  const arsipHeaders = arsipRows[0] || [];
  const catIdx = arsipHeaders.indexOf('Kategori');

  const arsipData = arsipRows.slice(1);
  const counts = { 
    DUMAS: 0, LHP: 0, MCSP_KPK: 0, PDTT: 0, PKKN: 0, PKTPT: 0,
    SK_TIM: 0, BAP: 0, PKA: 0, IEPK: 0, NOTA_DINAS: 0, PERBUP: 0,
    UNDANGAN: 0, INDISIPLINER_ASN: 0
  };

  arsipData.forEach(row => {
    const rawCat = row[catIdx] || '';
    const cat = String(rawCat).trim().toUpperCase();
    if (cat === 'DUMAS') counts.DUMAS++;
    else if (cat === 'LHP') counts.LHP++;
    else if (cat === 'MCSP KPK') counts.MCSP_KPK++;
    else if (cat === 'PDTT') counts.PDTT++;
    else if (cat === 'PKKN' || cat === 'PEDOMAN AUDIT PENGHITUNGAN KERUGIAN KEUANGAN NEGARA/DAERAH' || cat.includes('KERUGIAN KEUAN')) counts.PKKN++;
    else if (cat === 'PKTPT') counts.PKTPT++;
    else if (cat === 'SK TIM') counts.SK_TIM++;
    else if (cat === 'BERITA ACARA PEMERIKSAAN') counts.BAP++;
    else if (cat === 'PROGRAM KERJA AUDIT') counts.PKA++;
    else if (cat === 'IEPK') counts.IEPK++;
    else if (cat === 'NOTA DINAS') counts.NOTA_DINAS++;
    else if (cat === 'PENGAJUAN PERBUP') counts.PERBUP++;
    else if (cat === 'UNDANGAN') counts.UNDANGAN++;
    else if (cat === 'INDISIPLINER ASN') counts.INDISIPLINER_ASN++;
  });

  return {
    success: true,
    suratMasuk: await getCount(CONFIG.SHEETS.SURAT_MASUK),
    suratKeluar: await getCount(CONFIG.SHEETS.SURAT_KELUAR),
    undangan: await getCount(CONFIG.SHEETS.UNDANGAN),
    arsip: arsipData.length,
    spt: await getCount(CONFIG.SHEETS.SPT),
    // Detailed Arsip Categories
    countDumas: counts.DUMAS,
    countLhp: counts.LHP,
    countMcsp: counts.MCSP_KPK,
    countPdtt: counts.PDTT,
    countPkkn: counts.PKKN,
    countPktpt: counts.PKTPT,
    countSkTim: counts.SK_TIM,
    countBap: counts.BAP,
    countPka: counts.PKA,
    countIepk: counts.IEPK,
    countNota: counts.NOTA_DINAS,
    countPerbup: counts.PERBUP,
    countUndanganArsip: counts.UNDANGAN,
    countIndisipliner: counts.INDISIPLINER_ASN
  };
}

async function uploadToDrive(drive, rootFolderId, folderName, fileData) {
  if (!fileData || !fileData.content) return { url: '', name: '', id: '' };

  // APPS SCRIPT BRIDGE (ABSOLUTE QUOTA BYPASS)
  if (process.env.APPS_SCRIPT_URL) {
    try {
      const resp = await fetch(process.env.APPS_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({ folderId: rootFolderId, folderName: folderName, fileData: fileData })
      });
      const data = await resp.json();
      if (data.error) throw new Error(data.error);
      // Ensure embed view link
      if (data.url && data.url.includes('/view')) data.url = data.url.replace(/\/view.*$/, '/preview');
      return data;
    } catch (e) {
      throw new Error('Apps Script Upload Failed: ' + e.message);
    }
  }

  // Find or create subfolder (Legacy/Standard Service Account mode)
  const listRes = await drive.files.list({
    q: `name = '${folderName}' and '${rootFolderId}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false`,
    fields: 'files(id, name)',
    supportsAllDrives: true,
    includeItemsFromAllDrives: true
  });

  let folderId;
  if (listRes.data.files && listRes.data.files.length > 0) {
    folderId = listRes.data.files[0].id;
  } else {
    const createRes = await drive.files.create({
      resource: { name: folderName, mimeType: 'application/vnd.google-apps.folder', parents: [rootFolderId] },
      fields: 'id',
      supportsAllDrives: true
    });
    folderId = createRes.data.id;
  }

  // Upload file
  const buffer = Buffer.from(fileData.content, 'base64');
  const fileRes = await drive.files.create({
    resource: { name: fileData.name, parents: [folderId] },
    media: { mimeType: fileData.mimeType, body: require('stream').Readable.from(buffer) },
    fields: 'id, name, webViewLink',
    supportsAllDrives: true
  });

  // Set permission
  await drive.permissions.create({
    fileId: fileRes.data.id,
    resource: { role: 'reader', type: 'anyone' },
    supportsAllDrives: true
  });

  return {
    url: fileRes.data.webViewLink.replace('?usp=drivesdk', '').replace(/\/view.*$/, '/preview'),
    name: fileRes.data.name,
    id: fileRes.data.id
  };
}

async function getSheetData(sheets, spreadsheetId, sheetName) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${sheetName}!A:Z` });
  const rows = res.data.values || [];
  if (rows.length === 0) return { success: true, data: [] };

  const headers = rows[0];
  const data = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i] || '');
    return obj;
  });
  return { success: true, data };
}

async function deleteRowById(sheets, spreadsheetId, sheetName, id) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${sheetName}!A:A` });
  const rows = res.data.values || [];
  const index = rows.findIndex(row => String(row[0]) === String(id));

  if (index === -1) return { success: false, message: 'Data tidak ditemukan' };

  // Get sheetId
  const ss = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = ss.data.sheets.find(s => s.properties.title === sheetName);
  const sheetId = sheet.properties.sheetId;

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    resource: {
      requests: [{
        deleteDimension: {
          range: { sheetId, dimension: 'ROWS', startIndex: index, endIndex: index + 1 }
        }
      }]
    }
  });
  return { success: true, message: 'Data berhasil dihapus' };
}

async function saveArsip(sheets, drive, spreadsheetId, rootFolderId, data, fileData) {
  const folderName = data.kategori || 'Dokumentasi';
  const file = await uploadToDrive(drive, rootFolderId, folderName, fileData);
  const id = generateId();

  await sheets.spreadsheets.values.append({
    spreadsheetId, range: `${CONFIG.SHEETS.ARSIP}!A1`,
    valueInputOption: 'USER_ENTERED',
    resource: {
      values: [[id, data.namaFile, data.kategori, folderName, file.url, file.name, file.id, data.deskripsi || '-', fileData.size || '-', new Date().toISOString(), data.tglArsip || new Date().toISOString()]]
    }
  });
  return { success: true, message: 'Arsip berhasil diunggah', id };
}

async function saveSurat(sheets, drive, spreadsheetId, rootFolderId, sheetName, folderName, data, fileData) {
  const file = await uploadToDrive(drive, rootFolderId, folderName, fileData);
  const id = generateId();

  let row;
  if (sheetName === CONFIG.SHEETS.SURAT_MASUK) {
    row = [id, data.nomorSurat, data.tanggal, data.pengirim, data.perihal, data.kategori || 'Umum', file.url, file.name, file.id, data.catatan || '-', new Date().toISOString()];
  } else if (sheetName === CONFIG.SHEETS.SURAT_KELUAR) {
    row = [id, data.nomorSurat, data.tanggal, data.tujuan, data.perihal, data.kategori || 'Umum', file.url, file.name, file.id, data.catatan || '-', new Date().toISOString()];
  } else { // Undangan
    row = [id, data.nomorSurat, data.tanggal, data.penyelenggara, data.perihal, data.lokasi || '-', file.url, file.name, file.id, data.catatan || '-', new Date().toISOString()];
  }

  await sheets.spreadsheets.values.append({
    spreadsheetId, range: `${sheetName}!A1`,
    valueInputOption: 'USER_ENTERED',
    resource: { values: [row] }
  });
  return { success: true, message: 'Data berhasil disimpan', id };
}

async function saveSPT(sheets, drive, spreadsheetId, rootFolderId, data, fileData) {
  const file = await uploadToDrive(drive, rootFolderId, 'SPT', fileData);
  const id = generateId();
  await sheets.spreadsheets.values.append({
    spreadsheetId, range: `${CONFIG.SHEETS.SPT}!A1`,
    valueInputOption: 'USER_ENTERED',
    resource: {
      values: [[id, data.nomorSpt, data.nama, data.nip, data.jabatan, data.tujuan, data.keperluan, data.tglBerangkat, data.tglKembali, data.keterangan || '-', file.url, file.name, file.id, new Date().toISOString()]]
    }
  });
  return { success: true, message: 'SPT berhasil disimpan', id };
}

async function updateRowAndFile(sheets, drive, spreadsheetId, rootFolderId, sheetName, id, data, fileData) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${sheetName}!A:M` }); // Ambil M max kolom
  const rows = res.data.values || [];
  const index = rows.findIndex(row => String(row[0]) === String(id));
  
  if (index === -1) return { success: false, message: 'Data tidak ditemukan' };

  let oldRow = rows[index];
  
  // Deteksi kolom File ID dan URL beda sheet beda letak
  // ARSIP: index 6 (File ID), SURAT MASUK/KELUAR/UNDANGAN: index 8 (File ID), SPT: index 12 (File ID)
  // Periksa secara dinamis jika mungkin lewat Header, jika tidak pakai statis:
  let fileIdIdx = -1, fileUrlIdx = -1, fileNameIdx = -1;
  const headers = rows[0] || [];
  fileIdIdx = headers.indexOf('File ID');
  fileUrlIdx = headers.indexOf('URL');
  fileNameIdx = headers.indexOf('Nama File');

  let fileObj = null;
  // Jika user mengupload file baru
  if (fileData && fileData.content) {
    // 1. Hapus file lama di GDrive (HANYA jika bukan melalui Apps Script Bridge)
    if (fileIdIdx !== -1) {
      let oldFileId = oldRow[fileIdIdx];
      if (oldFileId && oldFileId !== '-' && !process.env.APPS_SCRIPT_URL) {
        try { await drive.files.delete({ fileId: oldFileId }); } catch(e) { console.error('Gagal hapus file lama', e); }
      }
    }
    
    // 2. Upload file baru
    const folderName = data.kategori || 'Update_Data';
    fileObj = await uploadToDrive(drive, rootFolderId, folderName, fileData);
  }

  // Rakit ulang data baris
  let newRow = [...oldRow];
  
  // Inject file info jika diganti
  if (fileObj && fileObj.url) {
    if (fileUrlIdx !== -1) newRow[fileUrlIdx] = fileObj.url;
    if (fileIdIdx !== -1) newRow[fileIdIdx] = fileObj.id;
    if (fileNameIdx !== -1) newRow[fileNameIdx] = fileObj.name;
    
    if (sheetName === CONFIG.SHEETS.ARSIP) {
      newRow[3] = data.kategori || oldRow[3]; // Update Folder Name
      newRow[5] = fileObj.name; // Drive Name khusus arsip
      newRow[8] = fileData.size || oldRow[8]; // Size
    }
  }

  // Update Data Teks berdasarkan Sheet Type
  if (sheetName === CONFIG.SHEETS.ARSIP) {
    newRow[1] = data.namaFile || oldRow[1];
    newRow[2] = data.kategori || oldRow[2];
    newRow[7] = data.deskripsi || oldRow[7];
    if (headers.indexOf('Tanggal Arsip') !== -1) newRow[headers.indexOf('Tanggal Arsip')] = data.tglArsip || oldRow[headers.indexOf('Tanggal Arsip')];
  } 
  else if (sheetName === CONFIG.SHEETS.SURAT_MASUK) {
    newRow[1] = data.nomorSurat || oldRow[1];
    newRow[2] = data.tanggal || oldRow[2];
    newRow[3] = data.pengirim || oldRow[3];
    newRow[4] = data.perihal || oldRow[4];
    newRow[5] = data.kategori || oldRow[5];
    newRow[9] = data.catatan || oldRow[9];
  }
  else if (sheetName === CONFIG.SHEETS.SURAT_KELUAR) {
    newRow[1] = data.nomorSurat || oldRow[1];
    newRow[2] = data.tanggal || oldRow[2];
    newRow[3] = data.tujuan || oldRow[3];
    newRow[4] = data.perihal || oldRow[4];
    newRow[5] = data.kategori || oldRow[5];
    newRow[9] = data.catatan || oldRow[9];
  }
  else if (sheetName === CONFIG.SHEETS.UNDANGAN) {
    newRow[1] = data.nomorSurat || oldRow[1];
    newRow[2] = data.tanggal || oldRow[2];
    newRow[3] = data.penyelenggara || oldRow[3];
    newRow[4] = data.perihal || oldRow[4];
    newRow[5] = data.lokasi || oldRow[5];
    newRow[9] = data.catatan || oldRow[9];
  }
  else if (sheetName === CONFIG.SHEETS.SPT) {
    newRow[1] = data.nomorSpt || oldRow[1];
    newRow[2] = data.nama || oldRow[2];
    newRow[3] = data.nip || oldRow[3];
    newRow[4] = data.jabatan || oldRow[4];
    newRow[5] = data.tujuan || oldRow[5];
    newRow[6] = data.keperluan || oldRow[6];
    newRow[7] = data.tglBerangkat || oldRow[7];
    newRow[8] = data.tglKembali || oldRow[8];
    newRow[9] = data.keterangan || oldRow[9];
  }

  // Lakukan Update ke Sheet
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetName}!A${index + 1}`,
    valueInputOption: 'USER_ENTERED',
    resource: { values: [newRow] }
  });

  return { success: true, message: 'Data berhasil diperbarui', id };
}

async function getUsers(sheets, spreadsheetId) {
  const res = await getSheetData(sheets, spreadsheetId, CONFIG.SHEETS.USERS);
  if (!res.success) return res;
  return {
    success: true,
    data: res.data.map(u => ({ id: u.ID, username: u.Username, nama: u.Nama, role: u.Role, created: u.CreatedAt }))
  };
}

async function addUser(sheets, spreadsheetId, data) {
  const existing = await getSheetData(sheets, spreadsheetId, CONFIG.SHEETS.USERS);
  if (existing.data.some(u => u.Username === data.username)) {
    return { success: false, message: 'Username sudah digunakan.' };
  }

  const id = generateId();
  await sheets.spreadsheets.values.append({
    spreadsheetId, range: `${CONFIG.SHEETS.USERS}!A1`,
    valueInputOption: 'USER_ENTERED',
    resource: {
      values: [[id, data.username, hashPassword(data.password), data.nama, data.role || 'ADMIN', new Date().toISOString()]]
    }
  });
  return { success: true, message: 'Pengguna berhasil ditambahkan.' };
}

async function changePassword(sheets, spreadsheetId, username, oldPassword, newPassword) {
  const res = await getSheetData(sheets, spreadsheetId, CONFIG.SHEETS.USERS);
  const rows = res.data;
  const index = rows.findIndex(u => u.Username === username && u.Password === hashPassword(oldPassword));

  if (index === -1) return { success: false, message: 'Password lama salah.' };

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${CONFIG.SHEETS.USERS}!C${index + 2}`,
    valueInputOption: 'USER_ENTERED',
    resource: { values: [[hashPassword(newPassword)]] }
  });
  return { success: true, message: 'Password berhasil diubah.' };
}

async function setupDb(sheets, spreadsheetId) {
  try {
    // ── Skema Kanonik ──────────────────────────────────────
    const SCHEMA = [
      {
        name: CONFIG.SHEETS.USERS,
        headers: ['ID', 'Username', 'Password', 'Nama', 'Role', 'CreatedAt'],
        defaultAdmin: ['ID-ADMIN123', 'admin', hashPassword('admin123'), 'Administrator Utama', 'ADMIN', new Date().toISOString()]
      },
      {
        name: CONFIG.SHEETS.SURAT_MASUK,
        headers: ['ID', 'Nomor Surat', 'Tanggal', 'Pengirim', 'Perihal', 'Kategori', 'URL', 'Nama File', 'File ID', 'Catatan', 'DibuatPada']
      },
      {
        name: CONFIG.SHEETS.SURAT_KELUAR,
        headers: ['ID', 'Nomor Surat', 'Tanggal', 'Tujuan', 'Perihal', 'Kategori', 'URL', 'Nama File', 'File ID', 'Catatan', 'DibuatPada']
      },
      {
        name: CONFIG.SHEETS.UNDANGAN,
        headers: ['ID', 'Nomor Surat', 'Tanggal', 'Penyelenggara', 'Perihal', 'Lokasi', 'URL', 'Nama File', 'File ID', 'Catatan', 'DibuatPada']
      },
      {
        name: CONFIG.SHEETS.SPT,
        headers: ['ID', 'Nomor SPT', 'Nama', 'NIP', 'Jabatan', 'Tujuan', 'Keperluan', 'Tgl Berangkat', 'Tgl Kembali', 'Keterangan', 'URL', 'Nama File', 'File ID', 'DibuatPada']
      },
      {
        name: CONFIG.SHEETS.ARSIP,
        headers: ['ID', 'Nama File', 'Kategori', 'Folder', 'URL', 'Nama Drive File', 'File ID', 'Deskripsi', 'Ukuran', 'DibuatPada', 'Tanggal Arsip']
      }
    ];

    // ── Blue header format ──────────────────────────────────
    // Header style: white bold text on blue (#1a6fbf) background, freeze row 1
    function buildHeaderFormatRequests(sheetId, numCols) {
      return [
        // Background biru + teks putih bold di header row
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: numCols },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.1, green: 0.435, blue: 0.749 },  // #1a6fbf
                textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 }, fontSize: 10 },
                horizontalAlignment: 'CENTER',
                verticalAlignment: 'MIDDLE',
                wrapStrategy: 'CLIP'
              }
            },
            fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)'
          }
        },
        // Freeze baris pertama (header)
        {
          updateSheetProperties: {
            properties: { sheetId, gridProperties: { frozenRowCount: 1 } },
            fields: 'gridProperties.frozenRowCount'
          }
        }
      ];
    }

    // ── Ambil daftar sheet yang ada ────────────────────────
    const ssInfo = await sheets.spreadsheets.get({ spreadsheetId });
    const existingSheetsMeta = ssInfo.data.sheets;

    const results = { created: [], migrated: [], unchanged: [] };
    const batchFormatRequests = [];

    for (const schema of SCHEMA) {
      const existingMeta = existingSheetsMeta.find(s => s.properties.title === schema.name);

      if (!existingMeta) {
        // ════════════════════════════════════════════════
        // CASE 1: Sheet BELUM ADA → Buat baru
        // ════════════════════════════════════════════════
        const addRes = await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          resource: { requests: [{ addSheet: { properties: { title: schema.name } } }] }
        });
        const newSheetId = addRes.data.replies[0].addSheet.properties.sheetId;

        // Tulis header
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `${schema.name}!A1`,
          valueInputOption: 'USER_ENTERED',
          resource: { values: [schema.headers] }
        });

        // Tambah default admin jika Users
        if (schema.name === CONFIG.SHEETS.USERS && schema.defaultAdmin) {
          await sheets.spreadsheets.values.append({
            spreadsheetId,
            range: `${CONFIG.SHEETS.USERS}!A2`,
            valueInputOption: 'USER_ENTERED',
            resource: { values: [schema.defaultAdmin] }
          });
        }

        // Format: biru + freeze
        batchFormatRequests.push(...buildHeaderFormatRequests(newSheetId, schema.headers.length));
        results.created.push(schema.name);

      } else {
        // ════════════════════════════════════════════════
        // CASE 2: Sheet SUDAH ADA → Cek & migrasi header
        // ════════════════════════════════════════════════
        const sheetId = existingMeta.properties.sheetId;

        // Baca data existing
        const readRes = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `${schema.name}!A:Z`
        });
        const existingRows = readRes.data.values || [];
        const existingHeaders = existingRows[0] || [];
        const dataRows = existingRows.slice(1);

        // Cek apakah header sudah identik
        const headersMatch = existingHeaders.length === schema.headers.length &&
          schema.headers.every((h, i) => h === existingHeaders[i]);

        if (!headersMatch) {
          // ── Ada perbedaan → migrasi data ──────────────
          // Bangun mapping: untuk setiap kolom baru, cari posisi di header lama
          // Kolom yang tidak ada di header lama → diisi '' (kosong)
          const oldIndexOf = (col) => existingHeaders.indexOf(col);

          // Mapping: setiap baris data lama → baris baru mengikuti skema
          const migratedRows = dataRows.map(row => {
            return schema.headers.map(newCol => {
              const oldIdx = oldIndexOf(newCol);
              return oldIdx !== -1 ? (row[oldIdx] || '') : '';
            });
          });

          // Persiapkan semua data: [header, ...migratedRows]
          const allNewData = [schema.headers, ...migratedRows];

          // Hapus semua konten sheet (tanpa hapus sheet-nya)
          await sheets.spreadsheets.values.clear({
            spreadsheetId,
            range: `${schema.name}!A:Z`
          });

          // Tulis ulang dengan header baru + data yang sudah dipetakan
          if (allNewData.length > 0) {
            await sheets.spreadsheets.values.update({
              spreadsheetId,
              range: `${schema.name}!A1`,
              valueInputOption: 'USER_ENTERED',
              resource: { values: allNewData }
            });
          }

          results.migrated.push(schema.name);
        } else {
          results.unchanged.push(schema.name);
        }

        // Selalu terapkan format biru + freeze (idempotent — aman diterapkan berulang)
        batchFormatRequests.push(...buildHeaderFormatRequests(sheetId, schema.headers.length));
      }
    }

    // ── Terapkan semua format sekaligus ────────────────────
    if (batchFormatRequests.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        resource: { requests: batchFormatRequests }
      });
    }

    // Susun pesan ringkasan
    const parts = [];
    if (results.created.length) parts.push(`✅ Dibuat baru: ${results.created.join(', ')}`);
    if (results.migrated.length) parts.push(`🔄 Dimigrasi: ${results.migrated.join(', ')}`);
    if (results.unchanged.length) parts.push(`✔️ Sudah sesuai: ${results.unchanged.join(', ')}`);
    parts.push('Header biru & freeze baris pertama diterapkan.');

    return { success: true, message: parts.join(' | ') };
  } catch (e) {
    console.error('setupDb error:', e);
    return { success: false, message: 'Gagal inisialisasi: ' + e.message };
  }
}


export const config = {
  api: {
    bodyParser: {
      sizeLimit: '15mb'
    }
  }
};
