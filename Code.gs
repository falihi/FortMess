/**
 * FortMess - Google Apps Script Backend (Code.gs)
 * Zero-cost, Google Sheets as Database, Secure Admin PIN + Audit Log.
 */

const SHEET_NAMES = {
  RESIDENTS: 'Master_Penghuni',
  DUES: 'Matriks_Iuran',
  LEDGER: 'Buku_Kas',
  AUDIT: 'Log_Audit',
};

const DEFAULT_DUES_AMOUNT = 80000;

/**
 * Jalankan fungsi ini SEKALI saja dari Apps Script Editor untuk membuat
 * 4 tab database beserta header dan data awal.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Master_Penghuni
  let resSheet = ss.getSheetByName(SHEET_NAMES.RESIDENTS);
  if (!resSheet) {
    resSheet = ss.insertSheet(SHEET_NAMES.RESIDENTS);
    resSheet.appendRow(['ID', 'Nama', 'Status', 'Tanggal_Masuk', 'Catatan']);
    resSheet.getRange('A1:E1').setFontWeight('bold').setBackground('#e2e8f0');
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
  }

  // 3. Buku_Kas
  let ledgerSheet = ss.getSheetByName(SHEET_NAMES.LEDGER);
  if (!ledgerSheet) {
    ledgerSheet = ss.insertSheet(SHEET_NAMES.LEDGER);
    ledgerSheet.appendRow(['ID_Trx', 'Tanggal', 'Tipe', 'Kategori', 'Keterangan', 'Qty', 'Nominal', 'Dicatat_Oleh', 'Bukti_URL']);
    ledgerSheet.getRange('A1:I1').setFontWeight('bold').setBackground('#e2e8f0');
    // Saldo awal cutoff (sesuaikan jika perlu)
    ledgerSheet.appendRow(['TRX-000', '2026-11-01', 'Pemasukan', 'Saldo Awal', 'Saldo Awal Cutoff Migrasi', 1, 591600, 'System', '']);
  }

  // 4. Log_Audit
  let auditSheet = ss.getSheetByName(SHEET_NAMES.AUDIT);
  if (!auditSheet) {
    auditSheet = ss.insertSheet(SHEET_NAMES.AUDIT);
    auditSheet.appendRow(['Timestamp', 'Aktor', 'Aksi', 'Detail']);
    auditSheet.getRange('A1:D1').setFontWeight('bold').setBackground('#fee2e2');
  }

  logAudit('System', 'SETUP_DB', 'Inisialisasi 4 tab database FortMess');
}

/**
 * Jalankan fungsi ini SEKALI saja untuk mengatur PIN Admin (disimpan sebagai SHA-256 hash).
 * Contoh: setAdminPin('123456') -> lalu hapus angka PIN-nya sebelum save.
 */
function setAdminPin(plainPin) {
  if (!plainPin || String(plainPin).length < 4) {
    throw new Error('PIN minimal 4 karakter');
  }
  const hash = hashPin_(String(plainPin));
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN_HASH', hash);
  logAudit('System', 'SET_ADMIN_PIN', 'Admin PIN hash diperbarui');
}

function hashPin_(pin) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pin);
  return rawHash.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function verifyAdmin_(pin) {
  const savedHash = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN_HASH');
  if (!savedHash) return false;
  return hashPin_(String(pin || '')) === savedHash;
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
  return file.getUrl();
}

function logAudit(actor, action, detail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.AUDIT);
  if (sheet) {
    sheet.appendRow([new Date().toISOString(), actor, action, String(detail)]);
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
  return HtmlService.createHtmlOutputFromFile('index')
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
  if (!verifyAdmin_(pin)) {
    logAudit('Anonymous', 'AUTH_FAILED', `Percobaan akses ${action}`);
    return { ok: false, error: 'PIN Admin salah atau tidak valid.' };
  }

  const actor = Session.getActiveUser().getEmail() || 'Admin_PIN';
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === 'verifyAuth') {
    return { ok: true, message: 'Akses Bendahara aktif.' };
  }

  if (action === 'addTransaction') {
    const sheet = ss.getSheetByName(SHEET_NAMES.LEDGER);
    const id = 'TRX-' + Date.now().toString().slice(-6);
    const qty = Number(data.qty) || 1;
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

  if (action === 'markDuesPaid') {
    const duesSheet = ss.getSheetByName(SHEET_NAMES.DUES);
    const ledgerSheet = ss.getSheetByName(SHEET_NAMES.LEDGER);
    const rows = duesSheet.getDataRange().getValues();
    const { residentId, residentName, monthYear, amount } = data;
    const payAmount = Number(amount) || DEFAULT_DUES_AMOUNT;

    // Cek idempotency (jangan sampai double-credit)
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === residentId && rows[i][2] === monthYear && rows[i][4] === 'Lunas') {
        return { ok: false, error: `Iuran ${residentName} bulan ${monthYear} sudah tercatat Lunas.` };
      }
    }

    const today = new Date().toISOString().slice(0, 10);
    duesSheet.appendRow([residentId, residentName, monthYear, payAmount, 'Lunas', today, actor]);

    const trxId = 'IUR-' + Date.now().toString().slice(-6);
    ledgerSheet.appendRow([
      trxId,
      today,
      'Pemasukan',
      'Iuran Bulanan',
      `Iuran ${residentName} (${monthYear})`,
      1,
      payAmount,
      actor,
    ]);

    logAudit(actor, 'MARK_DUES_PAID', `${residentName} - ${monthYear} (${payAmount})`);
    return { ok: true, data: getDashboardData() };
  }

  return { ok: false, error: 'Aksi tidak dikenal.' };
}

/**
 * Mengambil seluruh data dashboard (Read-Only)
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resRows = ss.getSheetByName(SHEET_NAMES.RESIDENTS).getDataRange().getValues();
  const duesRows = ss.getSheetByName(SHEET_NAMES.DUES).getDataRange().getValues();
  const ledgerRows = ss.getSheetByName(SHEET_NAMES.LEDGER).getDataRange().getValues();

  const residents = resRows.slice(1).map(r => ({
    id: r[0],
    name: r[1],
    status: r[2],
    joinedAt: r[3],
  }));

  const dues = duesRows.slice(1).map(r => ({
    residentId: r[0],
    name: r[1],
    month: r[2],
    amount: Number(r[3]) || 0,
    status: r[4],
    paidAt: r[5],
  }));

  let totalIncome = 0;
  let totalExpense = 0;
  const transactions = ledgerRows.slice(1).map(r => {
    const type = r[2];
    const nominal = Number(r[6]) || 0;
    if (type === 'Pemasukan') totalIncome += nominal;
    else totalExpense += nominal;
    return {
      id: r[0],
      date: r[1] instanceof Date ? r[1].toISOString().slice(0, 10) : String(r[1]),
      type,
      category: r[3],
      description: r[4],
      qty: Number(r[5]) || 1,
      nominal,
      proofUrl: r[8] || '',
    };
  }).reverse();

  return {
    summary: {
      balance: totalIncome - totalExpense,
      totalIncome,
      totalExpense,
      activeResidents: residents.filter(r => r.status === 'Aktif').length,
      duesAmount: DEFAULT_DUES_AMOUNT,
    },
    residents,
    dues,
    transactions,
  };
}
