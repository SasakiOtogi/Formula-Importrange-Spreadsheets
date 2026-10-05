/**
 * Mengambil nama sheet tepat di sebelah kiri sheet saat ini.
 * Sertakan parameter dummy (misal: RANDBETWEEN dari sheet) agar ter-refresh otomatis.
 * 
 * @param {any} trigger Pemicu agar fungsi me-refresh nilai (opsional).
 * @customfunction
 */
function PREVIOUS_SHEET_NAME(trigger) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const currentSheet = ss.getActiveSheet();
  const currentIndex = currentSheet.getIndex() - 1;
  
  if (currentIndex > 0) {
    return sheets[currentIndex - 1].getName();
  }
  return "N/A";
}

/**
 * Fungsi Automation untuk merefresh seluruh sel IMPORTRANGE di SELURUH sheet.
 * Jalankan fungsi ini menggunakan Time-driven Trigger di Apps Script.
 */
function refreshAllImportRange() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets(); // Mengambil seluruh sheet/tab
  
  sheets.forEach(sheet => {
    const range = sheet.getDataRange();
    const formulas = range.getFormulas();
    
    for (let r = 0; r < formulas.length; r++) {
      for (let c = 0; c < formulas[r].length; c++) {
        let formula = formulas[r][c];
        
        // Memeriksa kata IMPORTRANGE tanpa memperdulikan huruf besar/kecil
        if (formula.toUpperCase().includes("IMPORTRANGE")) {
          const cell = range.getCell(r + 1, c + 1);
          cell.setFormula('');
          SpreadsheetApp.flush();
          cell.setFormula(formula);
        }
      }
    }
  });
}

/**
 * Fungsi untuk men-copy sheet aktif dan otomatis menaikkan nomor baris IMPORTRANGE (+1)
 */
function duplicateSheetAndIncrementFormula() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  
  // 1. Duplikasi sheet yang sedang aktif
  const newSheet = activeSheet.copyTo(ss);
  
  // 2. Ambil semua rumus di sheet baru
  const range = newSheet.getDataRange();
  const formulas = range.getFormulas();
  
  // Regex untuk mencocokkan "HARIAN!" diikuti oleh huruf N, P, R, J, atau L, lalu angka
  // Contoh yang cocok: HARIAN!N460, HARIAN!P12, HARIAN!R5, HARIAN!J100, HARIAN!L20
  const targetPattern = /HARIAN!([NPRJL])(\d+)/gi;

  // 3. Cek sel per sel dan ubah nomor baris (+1)
  for (let r = 0; r < formulas.length; r++) {
    for (let c = 0; c < formulas[r].length; c++) {
      let formula = formulas[r][c];
      
      if (formula && targetPattern.test(formula)) {
        // Reset lastIndex regex setelah pengujian test()
        targetPattern.lastIndex = 0;

        let updatedFormula = formula.replace(targetPattern, function(match, colLetter, rowNumber) {
          let nextRow = parseInt(rowNumber, 10) + 1;
          return "HARIAN!" + colLetter + nextRow;
        });
        
        // Pasang kembali rumus yang sudah diperbarui angkanya
        newSheet.getRange(r + 1, c + 1).setFormula(updatedFormula);
      }
    }
  }
  
  // 4. Buka/aktifkan sheet baru yang sudah diperbarui
  ss.setActiveSheet(newSheet);
}

/**
 * Menambahkan Menu Khusus saat Google Sheets dibuka
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Duplicate Sheet') // Nama Menu
    .addItem('Duplicate', 'duplicateSheetAndIncrementFormula') // Nama Tombol & Fungsi
    .addToUi();
}