const SPREADSHEET_ID = '1fE1o0UUBWv7Y2vmGMnoHUNpDECWX8cnem8Nf9VeIHTA';
const FOLDER_ID = '1WfP5M3aT015folKfB5PYh6xfmPnytL3A';
const SLIDE_TEMPLATE_ID = '1419udWjSUMZ65OnBjarFRYoVshLPVMedoCU9VAZIBWI'; // ID dari link Slide Bapak

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    
    if (action === 'loginSekolah') return handleLoginSekolah(params.username, params.npsn);
    if (action === 'saveDraft') return handleSaveDraft(params.npsn, params.draftJson);
    if (action === 'generateCover') return handleGenerateCover(params.namaSekolah); // Fungsi baru
    if (action === 'finalize') return handleFinalize(params.npsn, params.pdfBase64, params.namaSekolah);
    
    if (action === 'loginAdmin') return handleLoginAdmin(params.username, params.password);
    if (action === 'adminList') return handleAdminList();
    if (action === 'adminUnlock') return handleAdminUnlock(params.npsn);
    
    return responseJSON({ status: 'error', message: 'Action tidak dikenali' });
  } catch (err) {
    return responseJSON({ status: 'error', message: err.message });
  }
}

function doGet(e) { return ContentService.createTextOutput("API e-Monev is running."); }
function responseJSON(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }

// --- LOGIKA COVER SLIDE (BARU) ---
function handleGenerateCover(namaSekolah) {
  try {
    // 1. Duplikasi Template Slide
    const templateFile = DriveApp.getFileById(SLIDE_TEMPLATE_ID);
    const tempFile = templateFile.makeCopy(`Temp_Cover_${namaSekolah}`);
    const tempId = tempFile.getId();
    
    // 2. Buka Slide Duplikat dan Replace Teks
    const presentation = SlidesApp.openById(tempId);
    const slides = presentation.getSlides();
    slides.forEach(slide => {
      slide.replaceAllText('<<Peran_Responden>>', '');
      slide.replaceAllText('<<Nama_Sekolah>>', namaSekolah);
    });
    presentation.saveAndClose(); // Pastikan perubahan teks tersimpan
    
    // 3. Konversi ke PDF Base64 (PERBAIKAN DI SINI)
    // Menggunakan string 'application/pdf' agar dikenali dengan tepat
    const pdfBlob = tempFile.getAs('application/pdf');
    const pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
    
    // 4. Hapus File Duplikat (bersih-bersih)
    tempFile.setTrashed(true);
    
    return responseJSON({ status: 'success', coverBase64: pdfBase64 });
  } catch (error) {
    return responseJSON({ status: 'error', message: 'Gagal membuat cover: ' + error.message });
  }
}

// --- LOGIKA SEKOLAH LAMA ---
function handleLoginSekolah(username, npsn) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data_Sekolah');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === username.trim() && String(data[i][1]).trim() === npsn.trim()) {
      return responseJSON({
        status: 'success',
        data: { username: data[i][0], npsn: data[i][1], nama_sekolah: data[i][2], status_sekolah: data[i][3], draft_json: data[i][4] || '{}', pdf_link: data[i][8] }
      });
    }
  }
  return responseJSON({ status: 'error', message: 'Username atau Password salah.' });
}

function handleSaveDraft(npsn, draftJson) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data_Sekolah');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(npsn)) {
      sheet.getRange(i + 1, 5).setValue(draftJson);
      sheet.getRange(i + 1, 10).setValue(new Date());
      return responseJSON({ status: 'success' });
    }
  }
  return responseJSON({ status: 'error', message: 'Gagal save draft' });
}

function handleFinalize(npsn, pdfBase64, namaSekolah) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data_Sekolah');
  const data = sheet.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(npsn)) { rowIdx = i + 1; break; }
  }
  if (rowIdx === -1) return responseJSON({ status: 'error', message: 'NPSN tidak ditemukan' });
  
  const sanitizedNama = namaSekolah.replace(/[^a-zA-Z0-9]/g, '_');
  const fileName = `Monev_Revit_${sanitizedNama}.pdf`;
  
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const blob = Utilities.newBlob(Utilities.base64Decode(pdfBase64), 'application/pdf', fileName); // Perbaikan decoding split
  const file = folder.createFile(blob);
  
  sheet.getRange(rowIdx, 9).setValue(file.getUrl());
  sheet.getRange(rowIdx, 10).setValue(new Date());
  
  return responseJSON({ status: 'success', pdf_link: file.getUrl() });
}

// --- LOGIKA ADMIN LAMA ---
function handleLoginAdmin(username, password) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data_Admin');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === username.trim() && String(data[i][1]).trim() === password.trim()) { return responseJSON({ status: 'success' }); }
  }
  return responseJSON({ status: 'error', message: 'Kredensial Admin salah.' });
}

function handleAdminList() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data_Sekolah');
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({ npsn: data[i][1], nama_sekolah: data[i][2], status: data[i][8] ? 'Selesai' : (data[i][4] ? 'Draf' : 'Belum Mulai'), pdf_link: data[i][8], draft_json: data[i][4] || '{}' });
  }
  return responseJSON({ status: 'success', data: result });
}

function handleAdminUnlock(npsn) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data_Sekolah');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(npsn)) {
      sheet.getRange(i + 1, 9).setValue('');
      return responseJSON({ status: 'success' });
    }
  }
  return responseJSON({ status: 'error', message: 'Sekolah tidak ditemukan' });
}
