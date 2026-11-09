const SPREADSHEET_ID = '1fE1o0UUBWv7Y2vmGMnoHUNpDECWX8cnem8Nf9VeIHTA';
const FOLDER_ID = '1WfP5M3aT015folKfB5PYh6xfmPnytL3A';
const SLIDE_TEMPLATE_ID = '1419udWjSUMZ65OnBjarFRYoVshLPVMedoCU9VAZIBWI';

const ROLE_COLS = {
  KS: { nama: 5, nip: 6, draft: 7, pdf: 8, time: 9, label: 'Kepala Sekolah' },
  Bendahara: { nama: 10, nip: 11, draft: 12, pdf: 13, time: 14, label: 'Bendahara' },
  P2SP: { nama: 15, nip: null, draft: 16, pdf: 17, time: 18, label: 'Ketua P2SP' }
};

const DEFAULT_INDIKATOR = [
  'Indikator Peran Dinas Pendidikan Kab/Kota',
  'Indikator Peran Kepala Sekolah',
  'Indikator Peran Bendahara',
  'Indikator Peran P2SP',
  'Indikator Peran Perencana',
  'Indikator Peran Pengawas (Pengawas Konstruksi)',
  'Indikator Progres Pembangunan dengan Kurva S',
  'Indikator Hambatan yang Muncul dan Penanganannya'
];

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    
    // Sekolah
    if (action === 'loginSekolah') return handleLoginSekolah(params.username, params.npsn);
    if (action === 'saveDraftResponden') return handleSaveDraftResponden(params.npsn, params.role, params.nama, params.nip, params.draftJson);
    if (action === 'generateCover') return handleGenerateCover(params.namaSekolah);
    if (action === 'finalizeResponden') return handleFinalizeResponden(params.npsn, params.role, params.nama, params.nip, params.pdfBase64, params.namaSekolah);
    
    // Admin (Marshal)
    if (action === 'loginAdmin') return handleLoginAdmin(params.username, params.password);
    if (action === 'adminList') return handleAdminList();
    if (action === 'adminUnlockResponden') return handleAdminUnlockResponden(params.npsn, params.role);
    if (action === 'getRoleSettings') return handleGetRoleSettings();
    if (action === 'saveRoleSettings') return handleSaveRoleSettings(params.settings);
    
    return responseJSON({ status: 'error', message: 'Action tidak dikenali' });
  } catch (err) {
    return responseJSON({ status: 'error', message: err.message });
  }
}

function doGet(e) { 
  return ContentService.createTextOutput("API e-Monev Revit is running."); 
}

function responseJSON(data) { 
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); 
}

// --- FUNGSI INISIALISASI / SETUP DATABASE SPREADSHEET ---
function setupSpreadsheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. Inisialisasi Sheet Data_Sekolah
  let sheetSekolah = ss.getSheetByName('Data_Sekolah');
  if (!sheetSekolah) {
    sheetSekolah = ss.insertSheet('Data_Sekolah');
  }
  
  const headers = [
    'Username', 'NPSN', 'Nama_Sekolah', 'Status Sekolah',
    'KS_Nama', 'KS_NIP', 'KS_Draft', 'KS_PDF', 'KS_Timestamp',
    'Bendahara_Nama', 'Bendahara_NIP', 'Bendahara_Draft', 'Bendahara_PDF', 'Bendahara_Timestamp',
    'P2SP_Nama', 'P2SP_Draft', 'P2SP_PDF', 'P2SP_Timestamp'
  ];
  sheetSekolah.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheetSekolah.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f1f5f9');
  
  // 2. Inisialisasi Sheet Pengaturan_Role
  let sheetRole = ss.getSheetByName('Pengaturan_Role');
  if (!sheetRole) {
    sheetRole = ss.insertSheet('Pengaturan_Role');
  }
  
  const roleHeaders = ['Indikator', 'KS', 'Bendahara', 'P2SP'];
  sheetRole.getRange(1, 1, 1, roleHeaders.length).setValues([roleHeaders]);
  sheetRole.getRange(1, 1, 1, roleHeaders.length).setFontWeight('bold').setBackground('#e2e8f0');
  
  // Isi default indikator jika belum ada data
  if (sheetRole.getLastRow() <= 1) {
    const defaultRows = DEFAULT_INDIKATOR.map(ind => [ind, true, true, true]);
    sheetRole.getRange(2, 1, defaultRows.length, 4).setValues(defaultRows);
  }
  
  // 3. Inisialisasi Sheet Data_Admin jika belum ada
  let sheetAdmin = ss.getSheetByName('Data_Admin');
  if (!sheetAdmin) {
    sheetAdmin = ss.insertSheet('Data_Admin');
    sheetAdmin.getRange(1, 1, 1, 2).setValues([['Username', 'Password']]).setFontWeight('bold');
    sheetAdmin.getRange(2, 1, 1, 2).setValues([['1', '1']]);
  }
  
  Logger.log('Setup Spreadsheet Berhasil Diselesaikan!');
  return 'Setup Spreadsheet Berhasil Diselesaikan!';
}

// --- LOGIKA COVER SLIDE ---
function handleGenerateCover(namaSekolah) {
  try {
    const templateFile = DriveApp.getFileById(SLIDE_TEMPLATE_ID);
    const tempFile = templateFile.makeCopy(`Temp_Cover_${namaSekolah}`);
    const tempId = tempFile.getId();
    
    const presentation = SlidesApp.openById(tempId);
    const slides = presentation.getSlides();
    slides.forEach(slide => {
      slide.replaceAllText('<<Peran_Responden>>', '');
      slide.replaceAllText('<<Nama_Sekolah>>', namaSekolah);
    });
    presentation.saveAndClose();
    
    const pdfBlob = tempFile.getAs('application/pdf');
    const pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
    
    tempFile.setTrashed(true);
    return responseJSON({ status: 'success', coverBase64: pdfBase64 });
  } catch (error) {
    return responseJSON({ status: 'error', message: 'Gagal membuat cover: ' + error.message });
  }
}

// --- LOGIKA ROLE SETTINGS ---
function getActiveRoleSettings() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheetRole = ss.getSheetByName('Pengaturan_Role');
  if (!sheetRole || sheetRole.getLastRow() <= 1) {
    setupSpreadsheet();
    sheetRole = ss.getSheetByName('Pengaturan_Role');
  }
  
  const values = sheetRole.getDataRange().getValues();
  const settings = { KS: [], Bendahara: [], P2SP: [] };
  
  for (let i = 1; i < values.length; i++) {
    const indikator = String(values[i][0]).trim();
    if (!indikator) continue;
    if (values[i][1] === true || String(values[i][1]).toLowerCase() === 'true') settings.KS.push(indikator);
    if (values[i][2] === true || String(values[i][2]).toLowerCase() === 'true') settings.Bendahara.push(indikator);
    if (values[i][3] === true || String(values[i][3]).toLowerCase() === 'true') settings.P2SP.push(indikator);
  }
  return settings;
}

function handleGetRoleSettings() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheetRole = ss.getSheetByName('Pengaturan_Role');
  if (!sheetRole || sheetRole.getLastRow() <= 1) {
    setupSpreadsheet();
    sheetRole = ss.getSheetByName('Pengaturan_Role');
  }
  
  const values = sheetRole.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    rows.push({
      indikator: values[i][0],
      KS: values[i][1] === true || String(values[i][1]).toLowerCase() === 'true',
      Bendahara: values[i][2] === true || String(values[i][2]).toLowerCase() === 'true',
      P2SP: values[i][3] === true || String(values[i][3]).toLowerCase() === 'true'
    });
  }
  return responseJSON({ status: 'success', data: rows });
}

function handleSaveRoleSettings(settingsArray) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheetRole = ss.getSheetByName('Pengaturan_Role');
  if (!sheetRole) {
    setupSpreadsheet();
    sheetRole = ss.getSheetByName('Pengaturan_Role');
  }
  
  // Clear data lama
  if (sheetRole.getLastRow() > 1) {
    sheetRole.getRange(2, 1, sheetRole.getLastRow() - 1, 4).clearContent();
  }
  
  const newRows = settingsArray.map(item => [
    item.indikator,
    Boolean(item.KS),
    Boolean(item.Bendahara),
    Boolean(item.P2SP)
  ]);
  
  if (newRows.length > 0) {
    sheetRole.getRange(2, 1, newRows.length, 4).setValues(newRows);
  }
  return responseJSON({ status: 'success' });
}

// --- LOGIKA SEKOLAH (MULTI-RESPONDEN) ---
function handleLoginSekolah(username, npsn) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data_Sekolah');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === username.trim() && String(data[i][1]).trim() === npsn.trim()) {
      const row = data[i];
      const roleSettings = getActiveRoleSettings();
      
      return responseJSON({
        status: 'success',
        data: {
          username: row[0],
          npsn: String(row[1]),
          nama_sekolah: row[2],
          status_sekolah: row[3],
          roles: {
            KS: {
              nama: row[4] || '',
              nip: row[5] || '',
              draft: row[6] || '{}',
              pdf_link: row[7] || '',
              timestamp: row[8] || '',
              status: row[7] ? 'Selesai' : (row[6] ? 'Draf' : 'Belum')
            },
            Bendahara: {
              nama: row[9] || '',
              nip: row[10] || '',
              draft: row[11] || '{}',
              pdf_link: row[12] || '',
              timestamp: row[13] || '',
              status: row[12] ? 'Selesai' : (row[11] ? 'Draf' : 'Belum')
            },
            P2SP: {
              nama: row[14] || '',
              nip: '',
              draft: row[15] || '{}',
              pdf_link: row[16] || '',
              timestamp: row[17] || '',
              status: row[16] ? 'Selesai' : (row[15] ? 'Draf' : 'Belum')
            }
          },
          settings: roleSettings
        }
      });
    }
  }
  return responseJSON({ status: 'error', message: 'Username atau Password salah.' });
}

function handleSaveDraftResponden(npsn, role, nama, nip, draftJson) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data_Sekolah');
  const data = sheet.getDataRange().getValues();
  const cfg = ROLE_COLS[role];
  if (!cfg) return responseJSON({ status: 'error', message: 'Role tidak valid' });
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(npsn).trim()) {
      const rowIdx = i + 1;
      sheet.getRange(rowIdx, cfg.nama).setValue(nama || '');
      if (cfg.nip) sheet.getRange(rowIdx, cfg.nip).setValue(nip || '');
      sheet.getRange(rowIdx, cfg.draft).setValue(draftJson);
      sheet.getRange(rowIdx, cfg.time).setValue(new Date());
      return responseJSON({ status: 'success' });
    }
  }
  return responseJSON({ status: 'error', message: 'NPSN sekolah tidak ditemukan' });
}

function handleFinalizeResponden(npsn, role, nama, nip, pdfBase64, namaSekolah) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data_Sekolah');
  const data = sheet.getDataRange().getValues();
  const cfg = ROLE_COLS[role];
  if (!cfg) return responseJSON({ status: 'error', message: 'Role tidak valid' });
  
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(npsn).trim()) { rowIdx = i + 1; break; }
  }
  if (rowIdx === -1) return responseJSON({ status: 'error', message: 'NPSN tidak ditemukan' });
  
  const sanitizedNama = namaSekolah.replace(/[^a-zA-Z0-9]/g, '_');
  const roleNameSanitized = cfg.label.replace(/ /g, '_');
  const fileName = `Monev_Revit_${roleNameSanitized}_${sanitizedNama}.pdf`;
  
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const blob = Utilities.newBlob(Utilities.base64Decode(pdfBase64), 'application/pdf', fileName);
  const file = folder.createFile(blob);
  
  sheet.getRange(rowIdx, cfg.nama).setValue(nama || '');
  if (cfg.nip) sheet.getRange(rowIdx, cfg.nip).setValue(nip || '');
  sheet.getRange(rowIdx, cfg.pdf).setValue(file.getUrl());
  sheet.getRange(rowIdx, cfg.time).setValue(new Date());
  
  return responseJSON({ status: 'success', pdf_link: file.getUrl() });
}

// --- LOGIKA ADMIN MARSHAL ---
function handleLoginAdmin(username, password) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data_Admin');
  if (!sheet) return responseJSON({ status: 'error', message: 'Sheet admin belum di-setup' });
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === username.trim() && String(data[i][1]).trim() === password.trim()) { 
      return responseJSON({ status: 'success' }); 
    }
  }
  return responseJSON({ status: 'error', message: 'Kredensial Admin salah.' });
}

function handleAdminList() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data_Sekolah');
  const data = sheet.getDataRange().getValues();
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    result.push({
      npsn: row[1],
      nama_sekolah: row[2],
      ks: {
        nama: row[4] || '',
        nip: row[5] || '',
        status: row[7] ? 'Selesai' : (row[6] ? 'Draf' : 'Belum Mengisi'),
        pdf_link: row[7] || '',
        draft_json: row[6] || '{}'
      },
      bendahara: {
        nama: row[9] || '',
        nip: row[10] || '',
        status: row[12] ? 'Selesai' : (row[11] ? 'Draf' : 'Belum Mengisi'),
        pdf_link: row[12] || '',
        draft_json: row[11] || '{}'
      },
      p2sp: {
        nama: row[14] || '',
        status: row[16] ? 'Selesai' : (row[15] ? 'Draf' : 'Belum Mengisi'),
        pdf_link: row[16] || '',
        draft_json: row[15] || '{}'
      }
    });
  }
  return responseJSON({ status: 'success', data: result });
}

function handleAdminUnlockResponden(npsn, role) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data_Sekolah');
  const data = sheet.getDataRange().getValues();
  const cfg = ROLE_COLS[role];
  if (!cfg) return responseJSON({ status: 'error', message: 'Role tidak valid' });
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(npsn).trim()) {
      sheet.getRange(i + 1, cfg.pdf).setValue('');
      return responseJSON({ status: 'success' });
    }
  }
  return responseJSON({ status: 'error', message: 'Sekolah tidak ditemukan' });
}
