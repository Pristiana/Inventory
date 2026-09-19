function doGet(e) {
  try {
    // 1. Ambil parameter sheets dari URL
    const sheetsParam = e.parameter.sheets || e.parameter.sheet;
    if (!sheetsParam) {
      return ContentService.createTextOutput("❌ Error: Parameter 'sheets' wajib diisi!").setMimeType(ContentService.MimeType.TEXT);
    }

    // 2. Folder Target di Google Drive
    const FOLDER_ID = "1N4g-9Ja-hq7-Hqfcm9u_ZAAHzyRLTtIj";
    const targetFolder = DriveApp.getFolderById(FOLDER_ID);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 3. Hitung Otomatis Kode Cabang & Kalender Senin-Minggu (YYMMW#)
    const cabang = "3"; // Kode Cabang Emba 3
    const now = new Date();
    
    // Format YY & MM (WIB)
    const yy = Utilities.formatDate(now, "Asia/Jakarta", "yy");
    const mm = Utilities.formatDate(now, "Asia/Jakarta", "MM");
    
    // Hitung Minggu ke-berapa dalam bulan berjalan (Senin -> Minggu)
    const weekNum = getWeekSOEmba(now);
    
    // Format Prefix Baku: "Inventory 3_2609W3"
    const filePrefix = `Inventory ${cabang}_${yy}${mm}W${weekNum}`;

    // 4. Proses Ekspor Tiap Sheet
    const targetList = sheetsParam.split(",").map(s => s.trim());
    let results = [];

    targetList.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        if (data.length > 1) {
          let csvContent = data.map(row => {
            return row.map(cell => {
              let val = cell.toString().replace(/"/g, '""');
              return `"${val}"`;
            }).join(",");
          }).join("\r\n");

          // Nama File Rapi: "Inventory 3_2609W3 - 1A_Penerimaan_Buah.csv"
          const fileName = `${filePrefix} - ${sheetName}.csv`;

          // Timpa file lama jika nama sama
          const existing = targetFolder.getFilesByName(fileName);
          while (existing.hasNext()) {
            existing.next().setTrashed(true);
          }

          targetFolder.createFile(fileName, csvContent, MimeType.CSV);
          results.push(fileName);
        }
      }
    });

    return ContentService.createTextOutput("✅ SUKSES: " + results.length + " file CSV siap impor POS:\n- " + results.join("\n- "))
      .setMimeType(ContentService.MimeType.TEXT);

  } catch (err) {
    return ContentService.createTextOutput("❌ ERROR: " + err.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

// Penentu Minggu Khusus Siklus SO Toko Emba (W1 = Senin Pertama Penuh)
function getWeekSOEmba(date) {
  const d = new Date(date);
  const day = d.getDate();
  
  // Cari tanggal Senin pertama di bulan ini
  const firstDay = new Date(d.getFullYear(), d.getMonth(), 1);
  const dayOfWeek = (firstDay.getDay() === 0) ? 7 : firstDay.getDay(); // 1=Senin..7=Minggu
  const firstMonday = (dayOfWeek === 1) ? 1 : (9 - dayOfWeek);
  
  // Tanggal 1 s.d sebelum Senin pertama (Masa gajian tgl 4 & transisi)
  if (day < firstMonday) {
    return "0"; 
  }
  
  // Hitung siklus 7 hari penuh SO dari Senin pertama
  return Math.floor((day - firstMonday) / 7) + 1;
}
