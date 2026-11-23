// ============================================
// 1. MENU OTOMATIS DI SPREADSHEET
// ============================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔄 Sinkron')
    .addItem('Sinkron Same Data', 'syncSameData')
    .addToUi();
}

// ============================================
// 2. DETEKSI INPUT TANGGAL & JAM (AUTO TIMESTAMP)
// ============================================
// Backlog: Input di Kolom A -> Isi Kolom H
// Datasuite: Input di Kolom A ATAU L -> Isi Kolom M
const configTimestamp = {
  'Backlog':   { triggerCols: [1], targetCol: 8 },
  'Datasuite': { triggerCols: [1, 12], targetCol: 13 }
};

function onEdit(e) {
  if (!e || !e.range) return;
  
  const sh = e.range.getSheet();
  const conf = configTimestamp[sh.getName()];
  if (!conf) return;
  
  const range = e.range;
  const startRow = range.getRow();
  const startCol = range.getColumn();
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();
  const endCol = startCol + numCols - 1;
  
  // Cek apakah range edit menyentuh kolom pemicu (trigger)
  let isTriggered = false;
  let triggerColIndex = -1;
  
  for (let i = 0; i < conf.triggerCols.length; i++) {
    const c = conf.triggerCols[i];
    // Jika area edit mencakup kolom trigger
    if (startCol <= c && endCol >= c) {
      isTriggered = true;
      triggerColIndex = c; 
      break;
    }
  }
  
  if (!isTriggered) return;
  
  const values = range.getValues();
  const now = new Date();
  const output = [];
  
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    
    // Lewati header (baris 1)
    if (row === 1) {
      output.push(['']); 
      continue;
    }
    
    // Dapatkan nilai dari kolom trigger yang spesifik
    const colOffset = triggerColIndex - startCol;
    const valTrigger = values[i][colOffset];
    
    if (valTrigger !== '' && valTrigger !== null && valTrigger !== undefined) {
      output.push([now]);
    } else {
      output.push(['']);
    }
  }
  
  // Tulis timestamp ke kolom target
  const targetRange = sh.getRange(startRow, conf.targetCol, numRows, 1);
  targetRange.setValues(output);
  targetRange.setNumberFormat('dd/MM/yyyy HH:mm:ss');
}

// ============================================
// 3. FUNGSI UTAMA SINKRONISASI & AGING
// ============================================
function syncSameData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backlogSheet   = ss.getSheetByName('Backlog');
  const datasuiteSheet = ss.getSheetByName('Datasuite');
  const sameSheet      = ss.getSheetByName('Same Data');
  
  // Validasi sheet
  if (!backlogSheet || !datasuiteSheet || !sameSheet) {
    SpreadsheetApp.getUi().alert('❌ Error: Pastikan sheet "Backlog", "Datasuite", dan "Same Data" ada.');
    return;
  }
  
  // ---- 1. Ambil data Backlog (A:G) ----
  const backlogLastRow = backlogSheet.getLastRow();
  const backlogData = backlogLastRow > 1
    ? backlogSheet.getRange(2, 1, backlogLastRow - 1, 7).getValues() // A sampai G
    : [];
  
  // ---- 2. Ambil data Datasuite (A:F) ----
  const datasuiteLastRow = datasuiteSheet.getLastRow();
  const datasuiteRaw = datasuiteLastRow > 1
    ? datasuiteSheet.getRange(2, 1, datasuiteLastRow - 1, 6).getValues() // A sampai F
    : [];
  
  // Buat Map dari Datasuite: Key = Kolom A, Value = Kolom F
  const datasuiteMap = new Map();
  datasuiteRaw.forEach(row => {
    const key = normalizeKey(row[0]);
    if (key !== '') datasuiteMap.set(key, row[5]); // row[5] adalah kolom F
  });
  
  // ---- 3. Cocokkan & Hitung Aging ----
  const result = [];
  const baseUrl = 'https://spx.shopee.co.id/#/orderDetail/'; 
  
  backlogData.forEach(row => {
    const id = row[0]; // Kolom A dari Backlog
    const key = normalizeKey(id);
    
    // Jika ada data yang sama di Datasuite
    if (key !== '' && datasuiteMap.has(key)) {
      
      // row[4] adalah Kolom E (Aging Time)
      const agingText = row[4]; 
      const agingStatus = hitungAging(agingText);
      
      // Susun baris: 
      // A = Link FMS (Base URL + ID)
      // B-H = Data Backlog A-G
      // I = Datasuite F
      // J = Aging Status
      result.push([
        baseUrl + id,          // Kolom A: Link FMS
        row[0],                // Kolom B: ID (dari Backlog A)
        row[1],                // Kolom C: dari Backlog B
        row[2],                // Kolom D: dari Backlog C
        row[3],                // Kolom E: dari Backlog D
        row[4],                // Kolom F: dari Backlog E (Teks Aging)
        row[5],                // Kolom G: dari Backlog F
        row[6],                // Kolom H: dari Backlog G
        datasuiteMap.get(key), // Kolom I: dari Datasuite F
        agingStatus            // Kolom J: Aging Info
      ]);
    }
  });
  
  // ---- 4. Tulis ke sheet Same Data ----
  const lastRowSame = sameSheet.getLastRow();
  const lastColSame = sameSheet.getLastColumn();
  
  // Bersihkan data lama (baris 2 ke bawah)
  if (lastRowSame > 1) {
    sameSheet.getRange(2, 1, lastRowSame - 1, Math.max(lastColSame, 10)).clearContent();
  }
  
  if (result.length === 0) {
    SpreadsheetApp.getUi().alert('⚠️ Tidak ada data yang cocok antara Backlog dan Datasuite.');
    return;
  }
  
  // Tulis data baru (10 kolom: A sampai J)
  sameSheet.getRange(2, 1, result.length, 10).setValues(result);
  
  // Pastikan Header terisi otomatis (hanya kalau masih kosong)
  if (!sameSheet.getRange(1, 1).getValue()) sameSheet.getRange(1, 1).setValue('Link FMS');
  if (!sameSheet.getRange(1, 2).getValue()) sameSheet.getRange(1, 2).setValue('ID');
  if (!sameSheet.getRange(1, 9).getValue()) sameSheet.getRange(1, 9).setValue('Datasuite F');
  if (!sameSheet.getRange(1, 10).getValue()) sameSheet.getRange(1, 10).setValue('Aging Info');
  
  SpreadsheetApp.getUi().alert(
    `✅ Selesai! ${result.length} baris data yang sama berhasil disinkronkan ke sheet "Same Data".`
  );
}

// ============================================
// 4. FUNGSI BANTUAN
// ============================================

// Fungsi untuk mendeteksi angka dari teks dan menentukan status aging
function hitungAging(text) {
  if (text === null || text === undefined || text === '') return '';
  
  // Ambil angka desimal dari teks (contoh: "60.73 hours" -> diambil "60.73")
  const match = text.toString().match(/[\d.]+/);
  if (!match) return '';
  
  const val = parseFloat(match[0]);
  if (isNaN(val)) return '';
  
  // Logika Aturan Threshold
  if (val > 72) return '>72hour';
  if (val > 48) return '>48hour';
  
  // Jika <= 48 jam, biarkan kosong
  return '';
}

// Fungsi untuk menyeragamkan teks kunci (menghapus spasi & ubah ke lowercase)
function normalizeKey(v) {
  if (v === null || v === undefined) return '';
  return v.toString().trim().toLowerCase();
}