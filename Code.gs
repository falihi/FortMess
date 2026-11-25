/**
 * FortMess - Google Apps Script Backend (Code.gs)
 * Zero-cost, Google Sheets as Database, Secure Admin PIN + Audit Log.
 */

const SHEET_NAMES = {
  RESIDENTS: 'Master_Penghuni',
  DUES: 'Matriks_Iuran',
  LEDGER: 'Buku_Kas',
  PIKET: 'Jadwal_Piket',
  AUDIT: 'Log_Audit',
};

const DEFAULT_DUES_AMOUNT = 80000;

function getDuesAmount_() {
  const saved = PropertiesService.getScriptProperties().getProperty('DUES_AMOUNT');
  return saved ? Number(saved) : DEFAULT_DUES_AMOUNT;
}

function getPaymentInfo_() {
  return PropertiesService.getScriptProperties().getProperty('PAYMENT_INFO') || 'Belum diatur oleh Bendahara';
}

function getCategories_() {
  const saved = PropertiesService.getScriptProperties().getProperty('CATEGORIES');
  return saved ? JSON.parse(saved) : ['Token Listrik', 'Air Galon', 'Gas', 'Kebersihan', 'Lainnya'];
}

function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('SPREADSHEET_ID');
  if (savedId) {
    try {
      const m = savedId.match(/\/d\/([a-zA-Z0-9_-]+)/);
      const cleanId = m ? m[1] : savedId.trim();
      return SpreadsheetApp.openById(cleanId);
    } catch (e) {}
  }

  const files = DriveApp.getFilesByName('Database Kas FortMess');
  if (files.hasNext()) {
    const file = files.next();
    props.setProperty('SPREADSHEET_ID', file.getId());
    return SpreadsheetApp.openById(file.getId());
  }

  const created = SpreadsheetApp.create('Database Kas FortMess');
  props.setProperty('SPREADSHEET_ID', created.getId());
  return created;
}

function ensureSheets_(ss) {
  // 1. Master_Penghuni
  let resSheet = ss.getSheetByName(SHEET_NAMES.RESIDENTS);
  if (!resSheet) {
    resSheet = ss.insertSheet(SHEET_NAMES.RESIDENTS);
    resSheet.appendRow(['ID', 'Nama', 'Status', 'Tanggal_Masuk', 'Catatan']);
    resSheet.getRange('A1:E1').setFontWeight('bold').setBackground('#e2e8f0');
    resSheet.getRange('A:E').setNumberFormat('@');
    const initialResidents = [
      ['R01', 'Achdi', 'Aktif', '2026-07-01', ''],
      ['R02', 'Falih', 'Aktif', '2026-07-01', ''],
      ['R03', 'Farros', 'Aktif', '2026-08-01', ''],
      ['R04', 'Firman', 'Aktif', '2026-07-01', ''],
      ['R05', 'Rafli', 'Aktif', '2026-07-01', ''],
      ['R06', 'Thoriq', 'Aktif', '2026-07-01', ''],
      ['R07', 'Topan', 'Aktif', '2026-07-01', ''],
      ['R08', 'Rio', 'Aktif', '2026-07-01', ''],
      ['R09', 'Anton', 'Non-Aktif', '2026-07-01', 'Keluar'],
      ['R10', 'Dimas', 'Non-Aktif', '2026-07-01', 'Keluar'],
      ['R11', 'Novarian', 'Non-Aktif', '2026-07-01', 'Keluar'],
      ['R12', 'David', 'Non-Aktif', '2026-07-01', 'Keluar'],
    ];
    resSheet.getRange(2, 1, initialResidents.length, 5).setValues(initialResidents);
  }

  // 2. Matriks_Iuran
  let duesSheet = ss.getSheetByName(SHEET_NAMES.DUES);
  if (!duesSheet) {
    duesSheet = ss.insertSheet(SHEET_NAMES.DUES);
    duesSheet.appendRow(['ID_Penghuni', 'Nama', 'Bulan', 'Nominal', 'Status', 'Tanggal_Bayar', 'Diverifikasi_Oleh']);
    duesSheet.getRange('A1:G1').setFontWeight('bold').setBackground('#e2e8f0');
    duesSheet.getRange('A:C').setNumberFormat('@');
    duesSheet.getRange('F:G').setNumberFormat('@');
  }

  // 3. Buku_Kas
  let ledgerSheet = ss.getSheetByName(SHEET_NAMES.LEDGER);
  if (!ledgerSheet) {
    ledgerSheet = ss.insertSheet(SHEET_NAMES.LEDGER);
    ledgerSheet.appendRow(['ID_Trx', 'Tanggal', 'Tipe', 'Kategori', 'Keterangan', 'Qty', 'Nominal', 'Dicatat_Oleh', 'Bukti_URL']);
    ledgerSheet.getRange('A1:I1').setFontWeight('bold').setBackground('#e2e8f0');
    ledgerSheet.getRange('A:B').setNumberFormat('@');
    const initialLedger = [
      ['TRX-000', '2026-11-01', 'Pemasukan', 'Saldo Awal', 'Saldo Awal Kas November', 1, 591600, 'System', ''],
      ['TRX-001', '2026-11-16', 'Pengeluaran', 'Kebersihan', 'Bersih-bersih (Topan)', '', 518300, 'System', ''],
      ['TRX-002', '2026-11-16', 'Pengeluaran', 'Air Galon', 'Air Galon', 3, 27000, 'System', ''],
      ['TRX-003', '2026-11-16', 'Pengeluaran', 'Gas', 'GAS', 2, 49000, 'System', ''],
      ['TRX-004', '2026-11-17', 'Pengeluaran', 'Kebersihan', 'Sunlight Kilau Nipis', 1, 8500, 'System', ''],
      ['TRX-005', '2026-11-17', 'Pengeluaran', 'Kebersihan', 'Spoon Cuci Piring', 1, 9700, 'System', ''],
      ['TRX-006', '2026-11-19', 'Pengeluaran', 'Kebersihan', 'Stella Ruangan', 1, 39900, 'System', ''],
      ['TRX-007', '2026-11-22', 'Pengeluaran', 'Air Galon', 'Air Galon', 3, 27000, 'System', ''],
    ];
    ledgerSheet.getRange(2, 1, initialLedger.length, 9).setValues(initialLedger);
  }

  // 4. Jadwal_Piket (Mingguan)
  let piketSheet = ss.getSheetByName(SHEET_NAMES.PIKET);
  if (!piketSheet) {
    piketSheet = ss.insertSheet(SHEET_NAMES.PIKET);
    piketSheet.appendRow(['Minggu_Ke', 'Hari_Jadwal', 'Petugas', 'Status', 'Catatan']);
    piketSheet.getRange('A1:E1').setFontWeight('bold').setBackground('#e2e8f0');
    piketSheet.appendRow(['Minggu Ini', 'Minggu', 'Achdi, Falih', 'Terjadwal', 'Bisa diganti ke Jumat Malam jika berhalangan']);
  }

  // 5. Log_Audit
  let auditSheet = ss.getSheetByName(SHEET_NAMES.AUDIT);
  if (!auditSheet) {
    auditSheet = ss.insertSheet(SHEET_NAMES.AUDIT);
    auditSheet.appendRow(['Timestamp', 'Aktor', 'Aksi', 'Detail']);
    auditSheet.getRange('A1:D1').setFontWeight('bold').setBackground('#fee2e2');
  }

  // Bersihkan tab kosong bawaan (Sheet1 / Lembar1) agar user tidak bingung melihat sheet kosong
  ['Sheet1', 'Sheet 1', 'Lembar1', 'Lembar 1'].forEach(defaultName => {
    const s = ss.getSheetByName(defaultName);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) {
      try { ss.deleteSheet(s); } catch (e) {}
    }
  });
}

/**
 * Jalankan fungsi ini SEKALI saja dari Apps Script Editor untuk memastikan
 * izin Google Drive & Sheets aktif serta 5 tab database dibuat.
 */
function setupDatabase() {
  DriveApp.getRootFolder(); // Memastikan izin Google Drive aktif untuk upload foto
  const ss = getSpreadsheet_();
  ensureSheets_(ss);
  Logger.log('Database siap di Spreadsheet: ' + ss.getName() + ' -> ' + ss.getUrl());
  logAudit('System', 'SETUP_DB', 'Inisialisasi tab database FortMess di ' + ss.getName());
}

function hashPin_(pin) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pin);
  return rawHash.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function verifyAdmin_(pin) {
  const props = PropertiesService.getScriptProperties();
  const savedPin = props.getProperty('ADMIN_PIN') || props.getProperty('initPin');
  if (savedPin) return String(pin || '') === String(savedPin);
  const savedHash = props.getProperty('ADMIN_PIN_HASH');
  if (savedHash) return hashPin_(String(pin || '')) === savedHash;
  return false;
}

function getOrCreateFolder_(parent, name) {
  const iter = parent.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : parent.createFolder(name);
}

function uploadProofImage_(proof, dateStr) {
  if (!proof || !proof.base64) return '';
  if (!String(proof.mimeType || '').startsWith('image/')) {
    throw new Error('File bukti harus berupa gambar.');
  }
  const bytes = Utilities.base64Decode(proof.base64);
  if (bytes.length > 2 * 1024 * 1024) {
    throw new Error('Ukuran gambar maksimal 2 MB.');
  }
  const d = (dateStr || new Date().toISOString().slice(0, 10));
  const year = d.slice(0, 4);
  const month = d.slice(0, 7);

  const root = getOrCreateFolder_(DriveApp.getRootFolder(), 'FortMess_Bukti');
  const yearFolder = getOrCreateFolder_(root, year);
  const monthFolder = getOrCreateFolder_(yearFolder, month);

  const blob = Utilities.newBlob(bytes, proof.mimeType, proof.fileName || `bukti-${Date.now()}.jpg`);
  const file = monthFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://drive.google.com/file/d/${file.getId()}/view`;
}

function trashDriveFileByUrl_(url) {
  if (!url) return;
  const match = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/) || String(url).match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (!match) return;
  try {
    DriveApp.getFileById(match[1]).setTrashed(true);
  } catch (e) {
    // Abaikan jika file sudah dihapus manual di Drive
  }
}

function logAudit(actor, action, detail) {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_NAMES.AUDIT);
    if (sheet) {
      sheet.appendRow([new Date().toISOString(), actor, action, String(detail)]);
    }
  } catch (e) {
    // Abaikan jika spreadsheet belum di-link
  }
}

/**
 * HTTP GET: Serve UI HTML atau JSON Data (?action=data)
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'data') {
    return ContentService
      .createTextOutput(JSON.stringify(getDashboardData()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const initialDataJson = JSON.stringify(getDashboardData());
  const rawHtml = HtmlService.createHtmlOutputFromFile('index').getContent();
  const injectedHtml = rawHtml.replace(
    'window.__INITIAL_DATA__ = null;',
    () => `window.__INITIAL_DATA__ = ${initialDataJson};\nwindow.__EXEC_URL__ = ${JSON.stringify(ScriptApp.getService().getUrl() || '')};`
  );
  return HtmlService.createHtmlOutput(injectedHtml)
    .setTitle('FortMess — Kas & Iuran Mes Putra Fortress')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * HTTP POST: Mutasi data (wajib PIN Admin valid)
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    return ContentService
      .createTextOutput(JSON.stringify(handleAction(payload)))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Dipanggil langsung oleh google.script.run dari index.html atau via doPost
 */
function handleAction(payload) {
  const { action, pin, data } = payload || {};
  const cache = CacheService.getScriptCache();
  const failKey = 'PIN_FAIL_COUNT';
  const fails = Number(cache.get(failKey) || 0);
  if (!verifyAdmin_(pin)) {
    Utilities.sleep(1200); // Tahan brute-force tanpa mengunci Bendahara asli yang tahu PIN benar
    cache.put(failKey, String(fails + 1), 300);
    logAudit('Anonymous', 'AUTH_FAILED', `Percobaan akses ${action} (gagal ke-${fails + 1})`);
    return { ok: false, error: 'PIN Bendahara salah atau tidak valid.' };
  }
  cache.remove(failKey);

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const actor = Session.getActiveUser().getEmail() || 'Admin_PIN';

    if (action === 'verifyAuth') {
      return { ok: true, message: 'Akses Bendahara aktif.' };
    }

    // Server-side anti-double-input (idempotency guard 4 detik untuk payload identik)
    const dedupHash = hashPin_(action + ':' + JSON.stringify(data || {}));
    const dedupKey = 'DEDUP_' + dedupHash.slice(0, 24);
    if (cache.get(dedupKey)) {
      return { ok: false, error: 'Permintaan ganda ditolak (sedang diproses).' };
    }
    cache.put(dedupKey, '1', 4);

    if (action === 'linkSpreadsheet') {
      const rawUrl = String(data.spreadsheetUrl || '').trim();
      if (!rawUrl) return { ok: false, error: 'Link / ID Google Sheet wajib diisi.' };
      const m = rawUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
      const cleanId = m ? m[1] : rawUrl;
      const targetSs = SpreadsheetApp.openById(cleanId);
      ensureSheets_(targetSs);
      PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', cleanId);
      logAudit(actor, 'LINK_SHEET', `Spreadsheet dihubungkan ke ${targetSs.getName()}`);
      return { ok: true, data: getDashboardData() };
    }

    const ss = getSpreadsheet_();
    if (!ss.getSheetByName(SHEET_NAMES.LEDGER)) {
      ensureSheets_(ss);
    }

    if (action === 'updateDuesAmount') {
      const newAmount = Number(data.amount);
      if (!newAmount || newAmount < 1000) {
        return { ok: false, error: 'Nominal iuran tidak valid.' };
      }
      PropertiesService.getScriptProperties().setProperty('DUES_AMOUNT', String(newAmount));
      logAudit(actor, 'UPDATE_DUES_AMOUNT', `Tarif iuran diubah menjadi ${newAmount}`);
      return { ok: true, data: getDashboardData() };
    }

    if (action === 'updatePaymentInfo') {
      const info = String(data.paymentInfo || '').trim();
      if (!info) return { ok: false, error: 'Info rekening tidak boleh kosong.' };
      PropertiesService.getScriptProperties().setProperty('PAYMENT_INFO', info);
      logAudit(actor, 'UPDATE_PAYMENT_INFO', `Info rekening diubah: ${info}`);
      return { ok: true, data: getDashboardData() };
    }

    if (action === 'saveCategories') {
      const cats = Array.isArray(data.categories) ? data.categories.map(c => String(c).trim()).filter(Boolean) : [];
      if (cats.length === 0) return { ok: false, error: 'Minimal harus ada 1 kategori.' };
      PropertiesService.getScriptProperties().setProperty('CATEGORIES', JSON.stringify(cats));
      logAudit(actor, 'SAVE_CATEGORIES', `Kategori diperbarui: ${cats.join(', ')}`);
      return { ok: true, data: getDashboardData() };
    }

    if (action === 'addResident') {
      const name = String(data.name || '').trim();
      const note = String(data.note || '').trim();
      if (!name) return { ok: false, error: 'Nama penghuni wajib diisi.' };
      const resSheet = ss.getSheetByName(SHEET_NAMES.RESIDENTS);
      const rows = resSheet.getDataRange().getValues().slice(1);
      let maxNum = 0;
      rows.forEach(r => {
        const m = String(r[0] || '').match(/\d+/);
        if (m) maxNum = Math.max(maxNum, Number(m[0]));
      });
      const id = 'R' + String(maxNum + 1).padStart(2, '0');
      resSheet.appendRow([id, name, 'Aktif', new Date().toISOString().slice(0, 10), note]);
      logAudit(actor, 'ADD_RESIDENT', `Tambah penghuni: ${name} (${id})`);
      return { ok: true, data: getDashboardData() };
    }

    if (action === 'updateResident') {
      const name = String(data.name || '').trim();
      const note = String(data.note || '').trim();
      if (!name) return { ok: false, error: 'Nama penghuni wajib diisi.' };
      const resSheet = ss.getSheetByName(SHEET_NAMES.RESIDENTS);
      const rows = resSheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.id)) {
          const oldName = rows[i][1];
          resSheet.getRange(i + 1, 2).setValue(name);
          resSheet.getRange(i + 1, 5).setValue(note);
          // Sinkronkan nama di Matriks_Iuran
          const duesSheet = ss.getSheetByName(SHEET_NAMES.DUES);
          if (duesSheet && duesSheet.getLastRow() >= 2) {
            const dRows = duesSheet.getDataRange().getValues();
            for (let j = 1; j < dRows.length; j++) {
              if (String(dRows[j][0]) === String(data.id)) {
                duesSheet.getRange(j + 1, 2).setValue(name);
              }
            }
          }
          logAudit(actor, 'UPDATE_RESIDENT', `Edit penghuni ${data.id}: ${oldName} -> ${name}`);
          return { ok: true, data: getDashboardData() };
        }
      }
      return { ok: false, error: 'Penghuni tidak ditemukan.' };
    }

    if (action === 'deleteResident') {
      const duesSheet = ss.getSheetByName(SHEET_NAMES.DUES);
      if (duesSheet && duesSheet.getLastRow() >= 2) {
        const dRows = duesSheet.getDataRange().getValues();
        const hasHistory = dRows.slice(1).some(r => String(r[0]) === String(data.id));
        if (hasHistory) {
          return {
            ok: false,
            error: 'Penghuni ini sudah memiliki riwayat pembayaran iuran. Gunakan tombol "Non-Aktifkan" agar histori kas tetap utuh.',
          };
        }
      }
      const resSheet = ss.getSheetByName(SHEET_NAMES.RESIDENTS);
      const rows = resSheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.id)) {
          const deletedName = rows[i][1];
          resSheet.deleteRow(i + 1);
          logAudit(actor, 'DELETE_RESIDENT', `Hapus penghuni ${data.id}: ${deletedName}`);
          return { ok: true, data: getDashboardData() };
        }
      }
      return { ok: false, error: 'Penghuni tidak ditemukan.' };
    }

    if (action === 'toggleResident') {
      const resSheet = ss.getSheetByName(SHEET_NAMES.RESIDENTS);
      const rows = resSheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.id)) {
          const newStatus = rows[i][2] === 'Aktif' ? 'Non-Aktif' : 'Aktif';
          resSheet.getRange(i + 1, 3).setValue(newStatus);
          logAudit(actor, 'TOGGLE_RESIDENT', `${rows[i][1]} -> ${newStatus}`);
          return { ok: true, data: getDashboardData() };
        }
      }
      return { ok: false, error: 'Penghuni tidak ditemukan.' };
    }

    if (action === 'addTransaction') {
      const sheet = ss.getSheetByName(SHEET_NAMES.LEDGER);
      const id = 'TRX-' + Date.now().toString().slice(-6);
      const qty = data.qty !== '' && data.qty != null ? Math.max(Number(data.qty) || 1, 1) : '';
      const nominal = Number(data.nominal) || 0;
      if (nominal <= 0) return { ok: false, error: 'Nominal harus lebih dari 0.' };

      const trxDate = data.date || new Date().toISOString().slice(0, 10);
      const proofUrl = data.proof ? uploadProofImage_(data.proof, trxDate) : '';

      sheet.appendRow([
        id,
        trxDate,
        data.type || 'Pengeluaran',
        data.category || 'Lainnya',
        data.description || '-',
        qty,
        nominal,
        actor,
        proofUrl,
      ]);
      logAudit(actor, 'ADD_TRX', `${data.type}: ${data.description} (${nominal})`);
      return { ok: true, data: getDashboardData() };
    }

    if (action === 'updateTransaction') {
      const sheet = ss.getSheetByName(SHEET_NAMES.LEDGER);
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.id)) {
          const qty = data.qty !== '' && data.qty != null ? Math.max(Number(data.qty) || 1, 1) : '';
          const nominal = Number(data.nominal) || 0;
          if (nominal <= 0) return { ok: false, error: 'Nominal harus lebih dari 0.' };

          const trxDate = data.date || fmtDate_(rows[i][1]) || new Date().toISOString().slice(0, 10);
          let proofUrl = String(rows[i][8] || '');

          if (data.proof && data.proof.base64) {
            trashDriveFileByUrl_(proofUrl);
            proofUrl = uploadProofImage_(data.proof, trxDate);
          } else if (data.removeProof) {
            trashDriveFileByUrl_(proofUrl);
            proofUrl = '';
          }

          sheet.getRange(i + 1, 2, 1, 8).setValues([[
            trxDate,
            data.type || 'Pengeluaran',
            data.category || 'Lainnya',
            data.description || '-',
            qty,
            nominal,
            actor,
            proofUrl,
          ]]);
          logAudit(actor, 'UPDATE_TRX', `Edit ${data.id}: ${data.description} (${nominal})`);
          return { ok: true, data: getDashboardData() };
        }
      }
      return { ok: false, error: 'Transaksi tidak ditemukan.' };
    }

    if (action === 'deleteTransaction') {
      const sheet = ss.getSheetByName(SHEET_NAMES.LEDGER);
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.id)) {
          const trxId = String(rows[i][0]);
          const desc = String(rows[i][4] || '');
          trashDriveFileByUrl_(rows[i][8]);
          sheet.deleteRow(i + 1);

          // Sinkronkan hapus dari Matriks_Iuran berdasarkan Foreign Key ID (IUR-Rxx-YYYY-MM) atau deskripsi
          const idMatch = trxId.match(/^IUR-(R\d+)-(\d{4}-\d{2})$/i);
          const descMatch = desc.match(/^Iuran\s+(.+)\s+\((\d{4}-\d{2})\)$/i);
          if (idMatch || descMatch) {
            const duesSheet = ss.getSheetByName(SHEET_NAMES.DUES);
            const dRows = duesSheet.getDataRange().getValues();
            for (let j = dRows.length - 1; j >= 1; j--) {
              const rowResId = String(dRows[j][0]).trim();
              const rowName = String(dRows[j][1]).trim().toLowerCase();
              const rowMonth = fmtMonth_(dRows[j][2]);
              if (
                (idMatch && rowResId === idMatch[1] && rowMonth === idMatch[2]) ||
                (descMatch && rowName === descMatch[1].trim().toLowerCase() && rowMonth === descMatch[2])
              ) {
                duesSheet.deleteRow(j + 1);
                break;
              }
            }
          }

          logAudit(actor, 'DELETE_TRX', `Hapus transaksi ${data.id}: ${desc} (${rows[i][6]})`);
          return { ok: true, data: getDashboardData() };
        }
      }
      return { ok: false, error: 'Transaksi tidak ditemukan.' };
    }

    if (action === 'markDuesPaid') {
      const duesSheet = ss.getSheetByName(SHEET_NAMES.DUES);
      const ledgerSheet = ss.getSheetByName(SHEET_NAMES.LEDGER);
      const rows = duesSheet.getDataRange().getValues();
      const { residentId, residentName, monthYear, amount } = data;
      const payAmount = Number(amount) || getDuesAmount_();

      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(residentId) && fmtMonth_(rows[i][2]) === monthYear && rows[i][4] === 'Lunas') {
          return { ok: false, error: `Iuran ${residentName} bulan ${monthYear} sudah tercatat Lunas.` };
        }
      }

      const today = new Date().toISOString().slice(0, 10);
      duesSheet.appendRow([residentId, residentName, String(monthYear), payAmount, 'Lunas', today, actor]);

      const trxId = `IUR-${residentId}-${monthYear}`;
      ledgerSheet.appendRow([
        trxId,
        today,
        'Pemasukan',
        'Iuran Bulanan',
        `Iuran ${residentName} (${monthYear})`,
        1,
        payAmount,
        actor,
        '',
      ]);

      logAudit(actor, 'MARK_DUES_PAID', `${residentName} - ${monthYear} (${payAmount})`);
      return { ok: true, data: getDashboardData() };
    }

    if (action === 'unmarkDuesPaid') {
      const duesSheet = ss.getSheetByName(SHEET_NAMES.DUES);
      const ledgerSheet = ss.getSheetByName(SHEET_NAMES.DUES ? SHEET_NAMES.LEDGER : 'Buku_Kas');
      const { residentId, residentName, monthYear } = data;
      const dRows = duesSheet.getDataRange().getValues();
      let found = false;
      for (let i = dRows.length - 1; i >= 1; i--) {
        if (String(dRows[i][0]) === String(residentId) && fmtMonth_(dRows[i][2]) === monthYear) {
          duesSheet.deleteRow(i + 1);
          found = true;
          break;
        }
      }
      if (!found) return { ok: false, error: 'Data lunas tidak ditemukan.' };

      const targetId = `IUR-${residentId}-${monthYear}`.toLowerCase();
      const targetDesc = `Iuran ${residentName} (${monthYear})`.toLowerCase();
      const lRows = ledgerSheet.getDataRange().getValues();
      for (let j = lRows.length - 1; j >= 1; j--) {
        const rowId = String(lRows[j][0] || '').trim().toLowerCase();
        const rowDesc = String(lRows[j][4] || '').trim().toLowerCase();
        if (rowId === targetId || rowDesc === targetDesc) {
          ledgerSheet.deleteRow(j + 1);
          break;
        }
      }

      logAudit(actor, 'UNMARK_DUES_PAID', `Batal lunas: ${residentName} - ${monthYear}`);
      return { ok: true, data: getDashboardData() };
    }

    if (action === 'saveMonthlyPiket') {
      const piketSheet = ss.getSheetByName(SHEET_NAMES.PIKET);
      const schedules = Array.isArray(data.schedules) ? data.schedules : [];
      if (schedules.length === 0) return { ok: false, error: 'Jadwal piket kosong.' };

      if (piketSheet.getLastRow() >= 2) {
        piketSheet.deleteRows(2, piketSheet.getLastRow() - 1);
      }
      const rowsToInsert = schedules.map(s => [
        String(s.week || ''),
        String(s.day || 'Minggu'),
        String(s.assignees || '-'),
        String(s.status || 'Terjadwal'),
        String(s.note || ''),
      ]);
      piketSheet.getRange(2, 1, rowsToInsert.length, 5).setValues(rowsToInsert);
      logAudit(actor, 'SAVE_MONTHLY_PIKET', `Jadwal piket ${schedules.length} minggu disimpan`);
      return { ok: true, data: getDashboardData() };
    }

    if (action === 'updateWeekPiket') {
      const piketSheet = ss.getSheetByName(SHEET_NAMES.PIKET);
      const rows = piketSheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.week)) {
          const newDay = data.day != null ? String(data.day) : String(rows[i][1]);
          const newAssignees = data.assignees != null ? String(data.assignees) : String(rows[i][2]);
          const newStatus = data.status != null ? String(data.status) : String(rows[i][3]);
          const newNote = data.note != null ? String(data.note) : String(rows[i][4] || '');
          piketSheet.getRange(i + 1, 2, 1, 4).setValues([[newDay, newAssignees, newStatus, newNote]]);
          logAudit(actor, 'UPDATE_WEEK_PIKET', `${data.week}: ${newDay} | ${newAssignees} | ${newStatus}`);
          return { ok: true, data: getDashboardData() };
        }
      }
      return { ok: false, error: 'Minggu piket tidak ditemukan di sheet. Klik "Acak Jadwal 1 Bulan" terlebih dahulu.' };
    }

    if (action === 'generateWeeklyPiket') {
      const resRows = ss.getSheetByName(SHEET_NAMES.RESIDENTS).getDataRange().getValues().slice(1);
      const activeNames = resRows.filter(r => r[2] === 'Aktif').map(r => r[1]);
      if (activeNames.length === 0) return { ok: false, error: 'Tidak ada penghuni aktif.' };

      const count = Math.min(Math.max(Number(data.count) || 2, 1), activeNames.length);
      const dayName = String(data.dayName || 'Minggu').trim();
      const piketSheet = ss.getSheetByName(SHEET_NAMES.PIKET);
      const prevStr = piketSheet.getLastRow() >= 2 ? String(piketSheet.getRange(2, 3).getValue()) : '';
      const prevList = prevStr.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);

      // Fair-Shuffle: prioritaskan yang belum piket minggu lalu
      const fresh = activeNames.filter(n => !prevList.includes(n.toLowerCase()));
      const repeat = activeNames.filter(n => prevList.includes(n.toLowerCase()));
      const shuffle = (arr) => {
        for (let i = arr.length - 1; i > 0; i--) {
          const j = Math.floor(Math.random() * (i + 1));
          [arr[i], arr[j]] = [arr[j], arr[i]];
        }
        return arr;
      };
      const pool = [...shuffle(fresh), ...shuffle(repeat)];
      const picked = pool.slice(0, count).join(', ');

      if (piketSheet.getLastRow() >= 2) {
        piketSheet.getRange(2, 1, 1, 5).setValues([['Minggu Ini', dayName, picked, 'Terjadwal', 'Diacak adil (Fair-Shuffle)']]);
      } else {
        piketSheet.appendRow(['Minggu Ini', dayName, picked, 'Terjadwal', 'Diacak adil (Fair-Shuffle)']);
      }
      logAudit(actor, 'GENERATE_PIKET', `Acak ${count} orang untuk ${dayName}: ${picked}`);
      return { ok: true, data: getDashboardData() };
    }

    if (action === 'markPiketDone') {
      const piketSheet = ss.getSheetByName(SHEET_NAMES.PIKET);
      if (piketSheet.getLastRow() < 2) return { ok: false, error: 'Belum ada jadwal piket.' };
      piketSheet.getRange(2, 4).setValue('Selesai');
      logAudit(actor, 'PIKET_DONE', 'Piket minggu ini ditandai Selesai');
      return { ok: true, data: getDashboardData() };
    }

    if (action === 'updatePiketDay') {
      const piketSheet = ss.getSheetByName(SHEET_NAMES.PIKET);
      const dayName = String(data.dayName || '').trim();
      if (!dayName) return { ok: false, error: 'Nama hari tidak boleh kosong.' };
      if (piketSheet.getLastRow() >= 2) {
        piketSheet.getRange(2, 2).setValue(dayName);
        logAudit(actor, 'UPDATE_PIKET_DAY', `Hari piket diubah ke ${dayName}`);
        return { ok: true, data: getDashboardData() };
      }
      return { ok: false, error: 'Belum ada jadwal piket.' };
    }

    if (action === 'swapPiketMember') {
      const piketSheet = ss.getSheetByName(SHEET_NAMES.PIKET);
      if (piketSheet.getLastRow() < 2) return { ok: false, error: 'Belum ada jadwal piket.' };
      const current = String(piketSheet.getRange(2, 3).getValue());
      const { oldName, newName } = data;
      if (!oldName || !newName) return { ok: false, error: 'Nama asal dan pengganti wajib diisi.' };
      const list = current.split(',').map(s => s.trim()).filter(Boolean);
      const idx = list.findIndex(s => s.toLowerCase() === oldName.toLowerCase());
      if (idx === -1) return { ok: false, error: `${oldName} tidak ada di jadwal minggu ini.` };
      const replaced = list[idx];
      list[idx] = newName;
      piketSheet.getRange(2, 3).setValue(list.join(', '));
      piketSheet.getRange(2, 5).setValue(`Tukar: ${replaced} -> ${newName}`);
      logAudit(actor, 'SWAP_PIKET', `${replaced} diganti ${newName}`);
      return { ok: true, data: getDashboardData() };
    }

    return { ok: false, error: 'Aksi tidak dikenal.' };
  } finally {
    lock.releaseLock();
  }
}

function fmtDate_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM-dd');
  }
  return v == null ? '' : String(v);
}

function fmtMonth_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM');
  }
  return v == null ? '' : String(v).slice(0, 7);
}

/**
 * Mengambil seluruh data dashboard (Read-Only)
 */
function getDashboardData() {
  try {
    const ss = getSpreadsheet_();
    if (!ss.getSheetByName(SHEET_NAMES.LEDGER) || !ss.getSheetByName(SHEET_NAMES.RESIDENTS)) {
      ensureSheets_(ss);
    }
    const resSheet = ss.getSheetByName(SHEET_NAMES.RESIDENTS);
    const duesSheet = ss.getSheetByName(SHEET_NAMES.DUES);
    const ledgerSheet = ss.getSheetByName(SHEET_NAMES.LEDGER);

    const resRows = resSheet ? resSheet.getDataRange().getValues() : [];
    const duesRows = duesSheet ? duesSheet.getDataRange().getValues() : [];
    const ledgerRows = ledgerSheet ? ledgerSheet.getDataRange().getValues() : [];

    const residents = resRows.slice(1).filter(r => r[0]).map(r => ({
      id: String(r[0]),
      name: String(r[1] || ''),
      status: String(r[2] || 'Aktif'),
      joinedAt: fmtDate_(r[3]),
      note: String(r[4] || ''),
    }));

    const dues = duesRows.slice(1).filter(r => r[0]).map(r => ({
      residentId: String(r[0]),
      name: String(r[1] || ''),
      month: fmtMonth_(r[2]),
      amount: Number(r[3]) || 0,
      status: String(r[4] || ''),
      paidAt: fmtDate_(r[5]),
    }));

    let totalIncome = 0;
    let totalExpense = 0;
    const transactions = ledgerRows.slice(1).filter(r => r[0]).map(r => {
      const type = String(r[2] || 'Pengeluaran');
      const nominal = Number(r[6]) || 0;
      if (type === 'Pemasukan') totalIncome += nominal;
      else totalExpense += nominal;
      return {
        id: String(r[0]),
        date: fmtDate_(r[1]),
        type,
        category: String(r[3] || 'Lainnya'),
        description: String(r[4] || '-'),
        qty: r[5] !== '' && r[5] != null ? Number(r[5]) : null,
        nominal,
        proofUrl: String(r[8] || ''),
      };
    }).reverse();

    const piketSheet = ss.getSheetByName(SHEET_NAMES.PIKET);
    const piket = piketSheet ? piketSheet.getDataRange().getValues().slice(1).filter(r => r[0]).map(r => ({
      week: String(r[0] || ''),
      day: String(r[1] || ''),
      assignees: String(r[2] || ''),
      status: String(r[3] || ''),
      note: String(r[4] || ''),
    })) : [];

    const result = {
      summary: {
        balance: totalIncome - totalExpense,
        totalIncome,
        totalExpense,
        activeResidents: residents.filter(r => r.status === 'Aktif').length,
        duesAmount: getDuesAmount_(),
        paymentInfo: getPaymentInfo_(),
        categories: getCategories_(),
        spreadsheetName: ss.getName(),
        spreadsheetUrl: ss.getUrl(),
      },
      residents,
      dues,
      transactions,
      piket,
    };

    return JSON.parse(JSON.stringify(result));
  } catch (err) {
    return {
      error: err.message,
      summary: {
        balance: 0,
        totalIncome: 0,
        totalExpense: 0,
        activeResidents: 0,
        duesAmount: getDuesAmount_(),
        paymentInfo: getPaymentInfo_(),
        categories: getCategories_(),
      },
      residents: [],
      dues: [],
      transactions: [],
      piket: [],
    };
  }
}
