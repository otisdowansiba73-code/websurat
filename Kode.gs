const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const FOLDER_DRIVE_ID = "https://drive.google.com/drive/folders/1agCrD3_4SEyLcX8KQtdCJodwMzn_7eGk"; // Masukkan ID Folder Drive di sini

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sistem Surat Masuk & Surat Keluar')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Helper untuk menyertakan file HTML lain
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ----------------------------------------------------
// VERIFIKASI LOGIN
// ----------------------------------------------------
function processLogin(username, password) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == username && data[i][2] == password) {
      return {
        success: true,
        user: {
          id_user: data[i][0],
          username: data[i][1],
          nama_lengkap: data[i][3],
          role: data[i][4]
        }
      };
    }
  }
  return { success: false, message: 'Username atau password salah!' };
}

// ----------------------------------------------------
// UPLOAD BERKAS KE GOOGLE DRIVE
// ----------------------------------------------------
function saveFileToDrive(fileObj) {
  try {
    const folder = DriveApp.getFolderById(FOLDER_DRIVE_ID);
    const contentType = fileObj.mimeType;
    const bytes = Utilities.base64Decode(fileObj.data);
    const blob = Utilities.newBlob(bytes, contentType, fileObj.fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return {
      fileName: file.getName(),
      fileUrl: file.getUrl()
    };
  } catch (err) {
    throw new Error('Gagal mengunggah file: ' + err.toString());
  }
}

// ----------------------------------------------------
// SURAT MASUK (CRUD)
// ----------------------------------------------------
function getSuratMasuk() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('SuratMasuk');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // Hapus header
  return data.map(row => {
    return {
      id_surat_masuk: row[0],
      no_agenda: row[1],
      no_surat: row[2],
      asal_surat: row[3],
      perihal: row[4],
      tanggal_surat: row[5] ? Utilities.formatDate(new Date(row[5]), "GMT+7", "yyyy-MM-dd") : "",
      tanggal_terima: row[6] ? Utilities.formatDate(new Date(row[6]), "GMT+7", "yyyy-MM-dd") : "",
      file_surat: row[7],
      url_file: row[8]
    };
  });
}

function saveSuratMasuk(form) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('SuratMasuk');
  let fileInfo = { fileName: form.existing_file || "", fileUrl: form.existing_url || "" };

  if (form.file_surat && form.file_surat.data) {
    fileInfo = saveFileToDrive(form.file_surat);
  }

  if (form.id_surat_masuk) {
    // Edit Data
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == form.id_surat_masuk) {
        sheet.getRange(i + 1, 2, 1, 7).setValues([[
          form.no_agenda,
          form.no_surat,
          form.asal_surat,
          form.perihal,
          form.tanggal_surat,
          form.tanggal_terima,
          fileInfo.fileName,
          fileInfo.fileUrl
        ]]);
        return { success: true, message: 'Surat masuk berhasil diperbarui!' };
      }
    }
  } else {
    // Tambah Baru
    const newId = 'SM-' + new Date().getTime();
    sheet.appendRow([
      newId,
      form.no_agenda,
      form.no_surat,
      form.asal_surat,
      form.perihal,
      form.tanggal_surat,
      form.tanggal_terima,
      fileInfo.fileName,
      fileInfo.fileUrl
    ]);
    return { success: true, message: 'Surat masuk berhasil disimpan!' };
  }
}

function deleteSuratMasuk(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('SuratMasuk');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Surat masuk berhasil dihapus!' };
    }
  }
  return { success: false, message: 'Data tidak ditemukan!' };
}

// ----------------------------------------------------
// SURAT KELUAR (CRUD)
// ----------------------------------------------------
function getSuratKeluar() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('SuratKeluar');
  const data = sheet.getDataRange().getValues();
  data.shift();
  return data.map(row => {
    return {
      id_surat_keluar: row[0],
      no_surat: row[1],
      tujuan_surat: row[2],
      perihal: row[3],
      tanggal_surat: row[4] ? Utilities.formatDate(new Date(row[4]), "GMT+7", "yyyy-MM-dd") : "",
      file_surat: row[5],
      url_file: row[6]
    };
  });
}

function saveSuratKeluar(form) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('SuratKeluar');
  let fileInfo = { fileName: form.existing_file || "", fileUrl: form.existing_url || "" };

  if (form.file_surat && form.file_surat.data) {
    fileInfo = saveFileToDrive(form.file_surat);
  }

  if (form.id_surat_keluar) {
    // Edit Data
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == form.id_surat_keluar) {
        sheet.getRange(i + 1, 2, 1, 6).setValues([[
          form.no_surat,
          form.tujuan_surat,
          form.perihal,
          form.tanggal_surat,
          fileInfo.fileName,
          fileInfo.fileUrl
        ]]);
        return { success: true, message: 'Surat keluar berhasil diperbarui!' };
      }
    }
  } else {
    // Tambah Data
    const newId = 'SK-' + new Date().getTime();
    sheet.appendRow([
      newId,
      form.no_surat,
      form.tujuan_surat,
      form.perihal,
      form.tanggal_surat,
      fileInfo.fileName,
      fileInfo.fileUrl
    ]);
    return { success: true, message: 'Surat keluar berhasil disimpan!' };
  }
}

function deleteSuratKeluar(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('SuratKeluar');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Surat keluar berhasil dihapus!' };
    }
  }
  return { success: false, message: 'Data tidak ditemukan!' };
}

// ----------------------------------------------------
// DISPOSISI (CRUD)
// ----------------------------------------------------
function getDisposisi(idSuratMasuk) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Disposisi');
  const data = sheet.getDataRange().getValues();
  data.shift();
  return data
    .filter(row => row[1] == idSuratMasuk)
    .map(row => ({
      id_disposisi: row[0],
      id_surat_masuk: row[1],
      tujuan_disposisi: row[2],
      isi_disposisi: row[3],
      sifat_surat: row[4]
    }));
}

function saveDisposisi(form) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Disposisi');
  const newId = 'DSP-' + new Date().getTime();
  sheet.appendRow([
    newId,
    form.id_surat_masuk,
    form.tujuan_disposisi,
    form.isi_disposisi,
    form.sifat_surat
  ]);
  return { success: true, message: 'Disposisi berhasil disimpan!' };
}

function deleteDisposisi(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Disposisi');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Disposisi berhasil dihapus!' };
    }
  }
  return { success: false, message: 'Data tidak ditemukan!' };
}
