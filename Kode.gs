var SPREADSHEET_ID = "1Rv1x6eeY6DXiZKX-hOSTkH9LcsD07ZymeOV51CEt9gI"; // GANTI DENGAN ID SPREADSHEET DATABASE
var SHEET_NAME_REMINDER = "Reminder"; 
var SHEET_NAME_ANGGOTA = "Anggota"; // harus sama!!!!!!
var DRIVE_FOLDER_NAME = "LampiranReminder"; // nama folder digdrive untuk simpan lampiran (bisa otomatis)

function onOpen() {
  // Logger.log("Spreadsheet dibuka.");
}

/**
 * Fungsi showApp() untuk menampilkan sidebar (jika masih dipanggil via doGet).
 */
function showApp() {
  var html = HtmlService.createHtmlOutputFromFile('WebApp') 
      .setTitle('Aplikasi Pengingat Agenda Gue') 
      .setWidth(1000); 
  SpreadsheetApp.getUi().showSidebar(html); 
}

/**
 * Mencari baris terakhir yang berisi data pada kolom pertama
 */
function getLastDataRow(sheet, columnIndexToConsiderAsMandatory, headerRows) {
  columnIndexToConsiderAsMandatory = columnIndexToConsiderAsMandatory || 1;
  headerRows = headerRows === undefined ? 1 : headerRows; 

  var lastRowReportedBySheet = sheet.getLastRow();
  if (lastRowReportedBySheet < headerRows) return headerRows; 
  if (lastRowReportedBySheet === headerRows) { 
      if (columnIndexToConsiderAsMandatory <= sheet.getLastColumn()) {
        var headerCell = sheet.getRange(headerRows, columnIndexToConsiderAsMandatory).getValue();
        if (headerCell && headerCell.toString().trim() !== "") {
            return headerRows;
        }
      }
      return 0; 
  }

  var dataStartRow = headerRows + 1;
  var numRowsToCheck = lastRowReportedBySheet - headerRows;
  if (numRowsToCheck <= 0) return headerRows;

  var mandatoryColumnValues = sheet.getRange(dataStartRow, columnIndexToConsiderAsMandatory, numRowsToCheck, 1).getValues();
  
  for (var i = mandatoryColumnValues.length - 1; i >= 0; i--) {
    if (mandatoryColumnValues[i][0] !== "" && mandatoryColumnValues[i][0] != null) {
      return dataStartRow + i; 
    }
  }
  return headerRows; 
}

/**
 * get objek spreadsheet
 */
function getSpreadsheetObject() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) throw new Error("Spreadsheet tidak ditemukan dengan ID: " + SPREADSHEET_ID);
    return ss;
  } catch (e) {
    Logger.log("Kritis: Gagal membuka Spreadsheet dengan ID: " + SPREADSHEET_ID + ". Error: " + e.message);
    throw new Error("Gagal mengakses spreadsheet target. Pastikan ID benar dan memiliki akses.");
  }
}

/**
 * Mengambil daftar nama anggota.
 */
function getNamaAnggota() {
  try {
    var ss = getSpreadsheetObject();
    var sheetAnggota = ss.getSheetByName(SHEET_NAME_ANGGOTA);
    if (!sheetAnggota) {
      Logger.log("Sheet '" + SHEET_NAME_ANGGOTA + "' tidak ditemukan.");
      return {error: "Sheet Anggota tidak ditemukan."};
    }
    var headerRowsAnggota = 3; // pastikan kolom anggota di kolom berapa
    var lastDataRowAnggota = getLastDataRow(sheetAnggota, 3, headerRowsAnggota); // Kolom C (Nama)
    
    if (lastDataRowAnggota <= headerRowsAnggota) { 
      return [];
    }
    var dataStartRowAnggota = headerRowsAnggota + 1;
    var numDataRowsAnggota = lastDataRowAnggota - headerRowsAnggota;
    var range = sheetAnggota.getRange(dataStartRowAnggota, 3, numDataRowsAnggota, 1); 
    var values = range.getValues();
    var namaList = values.map(function(row) { return row[0]; }).filter(String);
    return namaList;
  } catch (e) {
    Logger.log("Error getNamaAnggota: " + e.toString());
    return {error: "Gagal mengambil nama anggota: " + e.message};
  }
}

/**
 * Mengambil semua data agenda dari sheet "Reminder", atau yang di state di SHEET_NAME_REMINDER
 */
function getAgendaData() {
  try {
    var ss = getSpreadsheetObject(); 
    var sheet = ss.getSheetByName(SHEET_NAME_REMINDER); 

    if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAME_REMINDER);
        // Header: Jenis Agenda (A), Nama Agenda (B), Tanggal Load (C), DeadLine (D), PIC (E), Keterangan (F), RowID (G)
        // Kolom H: status (Formula), Kolom I: sisa hari (Formula)
        // Kolom J: Link File Lampiran, Kolom K: Link Tambahan
        sheet.getRange("A1:K1").setValues([["Jenis Agenda", "Nama Agenda", "Tanggal Load", "DeadLine", "PIC", "Keterangan", "RowID", "status", "sisa hari", "Link File Lampiran", "Link Tambahan"]]);
        Logger.log("Sheet '" + SHEET_NAME_REMINDER + "' dibuat dengan header (11 kolom).");
        return []; 
    }

    var headerRowsReminder = 1;
    var actualLastDataRow = getLastDataRow(sheet, 2, headerRowsReminder); // Kolom B ("Nama Agenda")

    if (actualLastDataRow <= headerRowsReminder) { 
        return [];
    }
    
    var dataStartRow = headerRowsReminder + 1;
    var numDataRows = actualLastDataRow - headerRowsReminder;

    // Baca semua kolom dari A sampai K (11 kolom)
    var dataRange = sheet.getRange(dataStartRow, 1, numDataRows, 11); 
    var dataValues = dataRange.getValues(); 
    var allAgenda = []; 

    for (var i = 0; i < dataValues.length; i++) {
      var rowData = dataValues[i];
      var rowNumberInSheet = dataStartRow + i; 
      var namaAgendaValid = rowData[1] && rowData[1].toString().trim() !== ""; // Cek ceknama di kolom B (index 1)

      if (namaAgendaValid) {
        var rowId = rowData[6] ? rowData[6].toString().trim() : null; // RowID di kolom G (index 6)
        if (!rowId) {
          rowId = Utilities.getUuid();
          sheet.getRange(rowNumberInSheet, 7).setValue(rowId); // Tulis RowID ke kolom G
        }
        allAgenda.push({ 
          jenisAgenda: rowData[0],  // A
          namaAgenda: rowData[1],   // B
          tanggalLoad: rowData[2] instanceof Date ? rowData[2].toISOString() : (rowData[2] ? new Date(rowData[2]).toISOString() : null), // C
          deadline: rowData[3] instanceof Date ? rowData[3].toISOString() : (rowData[3] ? new Date(rowData[3]).toISOString() : null), // D
          pic: rowData[4],          // E
          keterangan: rowData[5],   // F
          rowId: rowId,             // G
          // Kolom H (status) dan I (sisa hari) adalah formula, tidak perlu diambil untuk data agenda
          linkFile: rowData[9],     // J (index 9)
          linkTambahan: rowData[10] // K (index 10)
        });
      }
    }
    return allAgenda; 
  } catch (e) {
    Logger.log("Error getAgendaData: " + e.toString());
    return {error: "Gagal mengambil data agenda: " + e.message}; 
  }
}

/**
 * Fungsi untuk mendapatkan atau membuat folder di Google Drive.
 */
function getOrCreateFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

/**
 * Menangani upload file dan penambahan/pembaruan agenda.
 */
function prosesAgendaDenganFile(agendaData, fileInfo) {
  var linkFileDrive = agendaData.linkFile || null; // Pertahankan link file yang ada jika tidak ada upload baru
  try {
    if (fileInfo && fileInfo.base64Data) { // Ada file baru yang diupload
      var decodedBytes = Utilities.base64Decode(fileInfo.base64Data, Utilities.Charset.UTF_8);
      var blob = Utilities.newBlob(decodedBytes, fileInfo.mimeType, fileInfo.fileName);
      
      var folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
      var file = folder.createFile(blob);
      linkFileDrive = file.getUrl(); // URL file baru
      Logger.log("File berhasil diupload ke Drive: " + linkFileDrive);
    }

    agendaData.linkFile = linkFileDrive; 

    if (agendaData.rowId) {
      return updateAgenda(agendaData); 
    } else {
      return addAgenda(agendaData);
    }

  } catch (e) {
    Logger.log("Error di prosesAgendaDenganFile: " + e.toString() + "\nStack: " + e.stack);
    return { success: false, message: "Gagal memproses file atau agenda: " + e.message };
  }
}


/**
 * Menambahkan agenda baru ke sheet "Reminder".
 */
function addAgenda(agenda) {
  try {
    var ss = getSpreadsheetObject(); 
    var sheet = ss.getSheetByName(SHEET_NAME_REMINDER); 

     if (!sheet) { 
        sheet = ss.insertSheet(SHEET_NAME_REMINDER);
        sheet.getRange("A1:K1").setValues([["Jenis Agenda", "Nama Agenda", "Tanggal Load", "DeadLine", "PIC", "Keterangan", "RowID", "status", "sisa hari", "Link File Lampiran", "Link Tambahan"]]);
    }

    if (!agenda.namaAgenda || !agenda.deadline) {
      return {success: false, message: "Nama Agenda dan Deadline wajib diisi."};
    }

    var tanggalLoadOtomatis = new Date(); 
    var deadline = new Date(agenda.deadline);
    var rowId = agenda.rowId || Utilities.getUuid();

    var nextRow;
    var headerRowsReminder = 1;
    var lastDataRow = getLastDataRow(sheet, 2, headerRowsReminder); // Kolom B (Nama Agenda)

    if (lastDataRow <= headerRowsReminder) {
        nextRow = headerRowsReminder + 1; 
    } else {
        nextRow = lastDataRow + 1;
    }
    
    // Menulis ke kolom A-G dan J-K
    // Kolom A-F
    sheet.getRange(nextRow, 1, 1, 6).setValues([[
      agenda.jenisAgenda,
      agenda.namaAgenda,
      tanggalLoadOtomatis, 
      deadline,
      agenda.pic,
      agenda.keterangan
    ]]);
    // Kolom G (RowID)
    sheet.getRange(nextRow, 7).setValue(rowId);
    // Kolom J (Link File Lampiran) - kolom ke-10
    sheet.getRange(nextRow, 10).setValue(agenda.linkFile || "");
    // Kolom K (Link Tambahan) - kolom ke-11
    sheet.getRange(nextRow, 11).setValue(agenda.linkTambahan || "");
    
    Logger.log("Agenda ditambahkan: " + agenda.namaAgenda + ", RowID: " + rowId);
    return {success: true, message: "Agenda berhasil ditambahkan!", newRowId: rowId, tanggalLoad: tanggalLoadOtomatis.toISOString(), linkFile: agenda.linkFile};
  } catch (e) {
    Logger.log("Error addAgenda: " + e.toString());
    return {success: false, message: "Gagal menambahkan agenda: " + e.message};
  }
}

/**
 * fungsi update data yang sudah ada
 */
function updateAgenda(agenda) {
  try {
    var ss = getSpreadsheetObject(); 
    var sheet = ss.getSheetByName(SHEET_NAME_REMINDER); 

    if (!sheet) {
        return {success: false, message: "Sheet '" + SHEET_NAME_REMINDER + "' tidak ditemukan."};
    }
    if (!agenda.rowId) {
        return {success: false, message: "RowID tidak ditemukan untuk update."};
    }

    var headerRowsReminder = 1;
    var actualLastDataRow = getLastDataRow(sheet, 2, headerRowsReminder); 
                                                      
    if (actualLastDataRow <= headerRowsReminder) {
        return {success: false, message: "Tidak ada data untuk diupdate."};
    }

    var dataStartRowReminder = headerRowsReminder + 1;
    var numDataRowsReminder = actualLastDataRow - headerRowsReminder;
    if (numDataRowsReminder <= 0) {
        return {success: false, message: "Tidak ada data untuk diupdate."};
    }

    // Baca semua 11 kolom untuk mencari RowID di kolom G (index 6)
    var dataRange = sheet.getRange(dataStartRowReminder, 1, numDataRowsReminder, 11); 
    var dataValues = dataRange.getValues();
    var rowIndexToUpdate = -1;

    for (var i = 0; i < dataValues.length; i++) {
        if (dataValues[i][6] && dataValues[i][6].toString() === agenda.rowId.toString()) { // RowID di kolom G (index 6)
            rowIndexToUpdate = dataStartRowReminder + i; 
            break;
        }
    }

    if (rowIndexToUpdate === -1) {
        return {success: false, message: "Agenda dengan RowID tersebut tidak ditemukan."};
    }

    var existingTanggalLoad = sheet.getRange(rowIndexToUpdate, 3).getValue(); // Kolom C
    var deadline = new Date(agenda.deadline);

    // Update kolom A-F
    sheet.getRange(rowIndexToUpdate, 1, 1, 6).setValues([[ 
      agenda.jenisAgenda,
      agenda.namaAgenda,
      existingTanggalLoad, 
      deadline,
      agenda.pic,
      agenda.keterangan
    ]]);
    // Update kolom J (Link File) - kolom ke-10
    sheet.getRange(rowIndexToUpdate, 10).setValue(agenda.linkFile || ""); // Jika agenda.linkFile null/undefined, kosongkan atau pertahankan yang lama
    // Update kolom K (Link Tambahan) - kolom ke-11
    sheet.getRange(rowIndexToUpdate, 11).setValue(agenda.linkTambahan || "");

    // RowID (kolom G) tidak perlu diupdate karena sudah digunakan untuk mencari baris.
    
    Logger.log("Agenda diupdate: " + agenda.namaAgenda + ", RowID: " + agenda.rowId);
    return {success: true, message: "Agenda berhasil diperbarui!"};
  } catch (e) {
    Logger.log("Error updateAgenda: " + e.toString());
    return {success: false, message: "Gagal memperbarui agenda: " + e.message};
  }
}

/**
 * fungsi hapus
 */
function deleteAgenda(rowId) {
  try {
    var ss = getSpreadsheetObject(); 
    var sheet = ss.getSheetByName(SHEET_NAME_REMINDER); 

    if (!sheet || !rowId) {
        return {success: false, message: "Data tidak valid untuk penghapusan."};
    }
    
    var headerRowsReminder = 1;
    var actualLastDataRow = getLastDataRow(sheet, 2, headerRowsReminder); 
    if (actualLastDataRow <= headerRowsReminder) {
        return {success: false, message: "Tidak ada data untuk dihapus."};
    }

    var dataStartRowReminder = headerRowsReminder + 1;
    var numDataRowsReminder = actualLastDataRow - headerRowsReminder;
     if (numDataRowsReminder <= 0) {
        return {success: false, message: "Tidak ada data untuk dihapus."};
    }

    // Baca semua 11 kolom untuk mencari RowID di kolom G (index 6)
    var dataRange = sheet.getRange(dataStartRowReminder, 1, numDataRowsReminder, 11); 
    var dataValues = dataRange.getValues();
    var rowIndexToDelete = -1;

    for (var i = 0; i < dataValues.length; i++) {
        if (dataValues[i][6] && dataValues[i][6].toString() === rowId.toString()) { // RowID di kolom G (index 6)
            rowIndexToDelete = dataStartRowReminder + i; 
            break;
        }
    }

    if (rowIndexToDelete === -1) {
      return {success: false, message: "Agenda dengan RowID tersebut tidak ditemukan."};
    }

    sheet.deleteRow(rowIndexToDelete); 
    Logger.log("Agenda dihapus, RowID: " + rowId);
    return {success: true, message: "Agenda berhasil dihapus!"};
  } catch (e) {
    Logger.log("Error deleteAgenda: " + e.toString());
    return {success: false, message: "Gagal menghapus agenda: " + e.message};
  }
}

/**
 * fungsi membuat pengingat ulang tahun (harus dijalankan manual)
 */
function buatSemuaPengingatUlangTahunAnggota() {
  try {
    var ss = getSpreadsheetObject();
    var sheetAnggota = ss.getSheetByName(SHEET_NAME_ANGGOTA);
    var sheetReminder = ss.getSheetByName(SHEET_NAME_REMINDER);

    if (!sheetAnggota || !sheetReminder) {
      Logger.log("Sheet Anggota atau Reminder tidak ditemukan.");
      return;
    }

    var headerRowsAnggota = 3; 
    var lastDataRowAnggota = getLastDataRow(sheetAnggota, 3, headerRowsAnggota); 
    
    if (lastDataRowAnggota <= headerRowsAnggota) { 
      Logger.log("Tidak ada data anggota di sheet '" + SHEET_NAME_ANGGOTA + "'.");
      return;
    }

    var dataStartRowAnggota = headerRowsAnggota + 1;
    var numDataRowsAnggota = lastDataRowAnggota - headerRowsAnggota;

    var rangeAnggota = sheetAnggota.getRange(dataStartRowAnggota, 3, numDataRowsAnggota, 2); 
    var dataAnggota = rangeAnggota.getValues();
    var pengingatDibuatCount = 0;
    var pengingatSudahAdaCount = 0;
    var currentYear = new Date().getFullYear(); 

    for (var i = 0; i < dataAnggota.length; i++) {
      var namaAnggota = dataAnggota[i][0]; 
      var ultahTahunIniCell = dataAnggota[i][1];

      if (namaAnggota && namaAnggota.toString().trim() !== "" && ultahTahunIniCell) {
        var deadlineUltah;
        if (ultahTahunIniCell instanceof Date) {
          deadlineUltah = new Date(ultahTahunIniCell);
          deadlineUltah.setFullYear(currentYear); 
          deadlineUltah.setHours(9,0,0,0); 
        } else {
          var parts = ultahTahunIniCell.toString().match(/^(\d{1,2})[\/\.-](\d{1,2})[\/\.-](\d{4})$/);
          if (parts) {
            deadlineUltah = new Date(currentYear, parseInt(parts[2], 10) - 1, parseInt(parts[1], 10), 9, 0, 0);
          } else {
             try {
                var parsedDate = new Date(ultahTahunIniCell);
                if (!isNaN(parsedDate.getTime())) {
                    deadlineUltah = new Date(currentYear, parsedDate.getMonth(), parsedDate.getDate(), 9,0,0,0);
                } else { continue; }
             } catch(parseError){ continue; }
          }
        }
        
        if (isNaN(deadlineUltah.getTime())) {
            continue;
        }

        var keteranganUltah = "Ulang tahun " + namaAnggota.toString().trim();
        var yearForId = deadlineUltah.getFullYear();
        var rowIdUltah = "ULTAH_" + namaAnggota.toString().trim().replace(/\s+/g, '').toUpperCase() + "_" + yearForId;

        var pengingatSudahAda = cekPengingatUltahSudahAdaByIdAtauKeterangan(sheetReminder, rowIdUltah, keteranganUltah, yearForId);

        if (!pengingatSudahAda) {
          var agendaUltah = {
            jenisAgenda: "Pengingat", namaAgenda: keteranganUltah,
            deadline: deadlineUltah.toISOString(), pic: "Sistem", 
            keterangan: "Selamat ulang tahun untuk " + namaAnggota + "!", rowId: rowIdUltah,
            linkFile: null, linkTambahan: null 
          };
          
          var result = addAgenda(agendaUltah); 
          if (result.success) {
            pengingatDibuatCount++;
          }
        } else {
          pengingatSudahAdaCount++;
        }
      }
    }
    Logger.log(pengingatDibuatCount + " pengingat ultah baru dibuat. " + pengingatSudahAdaCount + " sudah ada. Tahun: " + currentYear);

  } catch (e) {
    Logger.log("Error buatSemuaPengingatUlangTahunAnggota: " + e.toString());
  }
}

/**
 * Helper untuk cek duplikasi pengingat ultah.
 */
function cekPengingatUltahSudahAdaByIdAtauKeterangan(sheetReminder, rowIdTarget, keteranganTarget, tahunTarget) {
  var headerRowsReminder = 1;
  var lastDataRowReminder = getLastDataRow(sheetReminder, 2, headerRowsReminder); 
  if (lastDataRowReminder <= headerRowsReminder) return false; 

  var dataStartRowReminder = headerRowsReminder + 1;
  var numDataRowsReminder = lastDataRowReminder - headerRowsReminder;
  if (numDataRowsReminder <= 0) return false;

  // Baca sampai kolom G (RowID) saja sudah cukup untuk pengecekan ini
  var rangeReminder = sheetReminder.getRange(dataStartRowReminder, 1, numDataRowsReminder, 7); 
  var dataReminder = rangeReminder.getValues();

  for (var i = 0; i < dataReminder.length; i++) {
    var rowIdSheet = dataReminder[i][6] ? dataReminder[i][6].toString().trim() : null; // RowID di kolom G (index 6)
    var namaAgendaSheet = dataReminder[i][1] ? dataReminder[i][1].toString().trim() : null; 
    var deadlineCell = dataReminder[i][3]; 

    if (rowIdSheet === rowIdTarget) {
        return true; 
    }

    if (namaAgendaSheet === keteranganTarget && deadlineCell) {
      var deadlineDate;
      if (deadlineCell instanceof Date) {
        deadlineDate = deadlineCell;
      } else {
        try { deadlineDate = new Date(deadlineCell); } catch (e) { continue; } 
      }
      if (!isNaN(deadlineDate.getTime()) && deadlineDate.getFullYear() === tahunTarget) {
        return true; 
      }
    }
  }
  return false; 
}

/**
 * Mengambil data agenda berdasarkan rentang tanggal deadline.
 */
function getAgendaByDeadlineDateRange(startDateString, endDateString) {
  var allAgenda = getAgendaData(); 
  if (allAgenda.error) return allAgenda;

  var startDate = new Date(startDateString);
  startDate.setHours(0, 0, 0, 0); 

  var endDate = new Date(endDateString);
  endDate.setHours(23, 59, 59, 999); 

  return allAgenda.filter(function(agenda) {
    if (!agenda.deadline) return false;
    var deadlineDate = new Date(agenda.deadline);
    return deadlineDate >= startDate && deadlineDate <= endDate;
  });
}

/**
 * Fungsi doGet untuk Web App.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('WebApp') 
      .setTitle('Aplikasi Pengingat Agenda Web') 
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); 
}
