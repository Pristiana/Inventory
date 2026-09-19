/**
 * Fungsi gabungan untuk menjalankan penarikan Master Data
 * dan sinkronisasi ke Sheet "SO" dalam sekali klik.
 */
function jalankanProsesLengkap() {
  try {
    // 1. Jalankan Penarikan Master Data Barang
    tarikMasterDataBarang();
    
    // 2. Jalankan Sinkronisasi ke Sheet SO
    sinkronisasiKeSO();
    
  } catch (error) {
    Logger.log(`[Error jalankanProsesLengkap]: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Proses terhenti karena error: ${error.message}`);
  }
}
