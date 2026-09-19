// =====================================================================
// 👑 GLOBAL VARIABLE & SETTINGS
// =====================================================================
var COOKIE_SAKTI = "";
var HEADERS_TOPENG = {
  "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
  "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
};

// =====================================================================
// 🔐 1️⃣ FUNGSI LOGIN iREAP
// =====================================================================
function loginCentratireap() {
  var loginUrl = "https://pro.ireappos.com/login"; 
  var payload = { 
    "email": "pristiana.faisal@gmail.com", 
    "password": "1610IrEaP@" 
  };

  function getCookie(response) {
    var c = response.getAllHeaders()["Set-Cookie"];
    if (!c) return "";
    return Array.isArray(c) ? c.map(x => x.split(';')[0]).join('; ') : c.split(';')[0];
  }

  var resAwal = UrlFetchApp.fetch(loginUrl, { "method": "get", "headers": HEADERS_TOPENG, "followRedirects": false });
  HEADERS_TOPENG["Cookie"] = getCookie(resAwal);

  var resLogin = UrlFetchApp.fetch(loginUrl, { "method": "post", "payload": payload, "headers": HEADERS_TOPENG, "followRedirects": false });
  if (resLogin.getResponseCode() === 200) throw new Error("Email/Password iREAP salah!");

  COOKIE_SAKTI = getCookie(resLogin) || HEADERS_TOPENG["Cookie"];
  HEADERS_TOPENG["Cookie"] = COOKIE_SAKTI;
}

// =====================================================================
// 📦 2️⃣ FETCH MASTER DATA BARANG DARI iREAP ➡️ TAB 'Data Barang'
// =====================================================================
function fetchMasterDataBarang() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var res = UrlFetchApp.fetch("https://pro.ireappos.com/article", {
    "method": "get",
    "headers": HEADERS_TOPENG,
    "followRedirects": false
  });

  var html = res.getContentText();
  var tbody = html.substring(html.search(/<tbody[^>]*>/i), html.search(/<\/tbody>/i));
  if (!tbody) throw new Error("Gagal membaca tabel barang dari iREAP!");

  var finalData = [];
  tbody.split('<tr').slice(1).forEach(row => {
    var rowData = row.split('<td').slice(1).map(col => 
      col.split('</td>')[0].replace(/^[^>]+>/, '').replace(/<[^>]+>/g, '').replace(/>/g, '').replace(/,/g, '').replace(/\s+/g, ' ').trim()
    );
    if (rowData.length > 0) {
      rowData.pop(); 
      finalData.push(rowData);
    }
  });

  if (finalData.length === 0) throw new Error("Data barang kosong!");
  finalData.sort((a, b) => (a[2] || "").toString().toLowerCase().localeCompare((b[2] || "").toString().toLowerCase()));

  var headers = [
    "System ID", "Item Code", "Description", "Category", 
    "Normal Price (IDR)", "Promo Price (IDR)", "Wholesale Price (IDR)", "Wholesale Promo Price (IDR)", 
    "Unit of Measure", "Min Stock", "Tax (%)", "Non Stock", 
    "Unsellable", "Open Selling Price", "Status", "Last Update", "Cost"
  ];

  var sheet = ss.getSheetByName("Data Barang") || ss.insertSheet("Data Barang");
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sheet.getRange(2, 1, finalData.length, headers.length).setValues(finalData);

  SpreadsheetApp.flush(); 
  return finalData.length;
}

// =====================================================================
// 🪡 3️⃣ AUTO COPAS + RESET QC + AUTO-CENTANG JALUR VIP 
// =====================================================================
function copasDataBarangKeSO() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSource = ss.getSheetByName("Data Barang");
  var sheetTarget = ss.getSheetByName("SO");

  var lastRowSrc = sheetSource.getLastRow();
  if (lastRowSrc < 2) throw new Error("Tab 'Data Barang' masih kosong melompong!");

  var numRows = lastRowSrc - 1;

  var dataCode = sheetSource.getRange(2, 2, numRows, 1).getValues();
  var dataCat  = sheetSource.getRange(2, 4, numRows, 1).getValues();
  var dataDesc = sheetSource.getRange(2, 3, numRows, 1).getValues();
  var dataPrce = sheetSource.getRange(2, 5, numRows, 1).getValues();
  var dataCost = sheetSource.getRange(2, 17, numRows, 1).getValues();

  var lastRowSO = Math.max(2, sheetTarget.getLastRow());
  var lastColSO = Math.max(33, sheetTarget.getLastColumn()); 

  // Bersihkan sisa baris di bawah kalau SO kepanjangan
  if (lastRowSO > numRows + 1) {
    sheetTarget.getRange(numRows + 2, 1, lastRowSO - (numRows + 1), lastColSO).clearContent();
  }

  var targetRange = sheetTarget.getRange(2, 1, numRows, lastColSO);
  var values = targetRange.getValues();
  var formulas = targetRange.getFormulas(); 
  var headers = sheetTarget.getRange(1, 1, 1, lastColSO).getValues()[0];

  for (var r = 0; r < numRows; r++) {
    // 💥 EKSTRAKSI NAMA BARANG BIAR GAMPANG DI-SCAN
    var itemNameClean = (dataDesc[r][0] || "").toString().trim().toLowerCase();

    for (var c = 0; c < lastColSO; c++) {
      
      // PERISAI RUMUS (Abaikan kalau ada rumus)
      if (formulas[r][c] !== "") {
        values[r][c] = formulas[r][c]; 
        continue; 
      }

      // INJEKSI DATA MASTER
      if (c === 0) { values[r][c] = dataCode[r][0]; continue; } 
      if (c === 1) { values[r][c] = dataCat[r][0]; continue; }  
      if (c === 2) { values[r][c] = dataDesc[r][0]; continue; } 
      if (c === 20) { values[r][c] = dataPrce[r][0]; continue; }
      if (c === 30) { values[r][c] = dataCost[r][0]; continue; }

      // LOGIKA RESET DEFAULT
      var headName = (headers[c] || "").toString().toLowerCase().trim();

      if (headName === "# rq" || headName === "# q") {
        values[r][c] = 0; 
      } 
      else if (headName === "update manual") {
        // 💥 JALUR VIP: OTOMATIS CENTANG TRUE BUAT BARANG GAIB! (BEBAS DIKLIK MANUAL DI APPSHEET)
        if (itemNameClean === "belum ada barcode" || 
            itemNameClean === "topup (gross only)" || 
            itemNameClean === "ongkos kirim") {
          values[r][c] = true;
        } else {
          values[r][c] = false;
        }
      }
      else if (headName === "aman" || 
               headName === "dekat kadaluarsa / tidak layak display" || 
               headName === "kadaluarsa / rusak" || 
               headName === "salah input") {
        values[r][c] = false; 
      } 
      else {
        values[r][c] = ""; 
      }
    }
  }

  targetRange.setValues(values);
  SpreadsheetApp.flush(); 
}

// =====================================================================
// 🚀 4️⃣ TOMBOL IMPORT UTAMA
// =====================================================================
function IMPORT_MINGGUAN_SO() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  try {
    ss.toast("⏳ 1/3 Login ke server iREAP...", "IMPORT DATA", 5);
    loginCentratireap();

    ss.toast("📦 2/3 Mengambil Data Barang...", "IMPORT DATA", 15);
    var total = fetchMasterDataBarang();

    ss.toast("🪡 3/3 Injeksi Data, Proteksi Rumus & Reset QC...", "IMPORT DATA", 10);
    copasDataBarangKeSO(); 

    ui.alert("✅ SUKSES BRUTAL!\n\nMaster Data: " + total + " item sukses diimpor.\nBarang gaib udah otomatis kecentang 'Update Manual', dan si kuli tetep bisa ngeklik bebas di AppSheet!");
  } catch (err) {
    ui.alert("❌ GAGAL IMPORT!\n\nPenyebab: " + err.message);
  }
}
