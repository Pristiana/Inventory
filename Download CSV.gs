function doGet(e) {
  try {
    // 1. Ambil parameter tab dari tombol AppSheet
    const sheetName = e.parameter.sheet;
    if (!sheetName) {
      return ContentService.createTextOutput("❌ Error: Parameter 'sheet' tidak ditemukan!").setMimeType(ContentService.MimeType.TEXT);
    }

    // 2. ID Subfolder tujuan di Google Drive (Ganti dengan ID Folder Tahap 1)
    const TARGET_FOLDER_ID = "1N4g-9Ja-hq7-Hqfcm9u_ZAAHzyRLTtIj";
    const targetFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);

    // 3. Ambil data tab spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return ContentService.createTextOutput("❌ Error: Tab '" + sheetName + "' tidak ditemukan di spreadsheet!").setMimeType(ContentService.MimeType.TEXT);
    }

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) {
      return ContentService.createTextOutput("⚠️ Peringatan: Tab '" + sheetName + "' kosong (hanya header).").setMimeType(ContentService.MimeType.TEXT);
    }

    // 4. Konversi baris data menjadi format CSV
    let csvString = values.map(row => {
      return row.map(cell => {
        let text = cell.toString().replace(/"/g, '""');
        return `"${text}"`;
      }).join(",");
    }).join("\r\n");

    const fileName = `${sheetName}.csv`;

    // 5. Timpa file lama jika sudah ada file dengan nama yang sama di folder
    const oldFiles = targetFolder.getFilesByName(fileName);
    while (oldFiles.hasNext()) {
      oldFiles.next().setTrashed(true);
    }

    // 6. Buat file CSV baru di folder Drive
    targetFolder.createFile(fileName, csvString, MimeType.CSV);

    return ContentService.createTextOutput("✅ SUKSES: File '" + fileName + "' berhasil disimpan ke Google Drive!")
      .setMimeType(ContentService.MimeType.TEXT);

  } catch (err) {
    return ContentService.createTextOutput("❌ ERROR: " + err.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}
