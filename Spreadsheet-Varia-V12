/**
 * ============================================================================
 * MASTER SCRIPT V.12 (PT Varia Indo Pangan)
 * Karya : MAHFUD FEBRY S
 * ============================================================================
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(' [MAHFUD - ADMIN MENU]')
      .addItem('📱 Generate Laporan WA', 'generateWALaporan')
      .addItem('📱 Laporan Kilat Boss (Teks)', 'generateLaporanBoss')
      .addItem('📱 Rekap Produksi', 'generateRekapProduksi') 
      .addItem('💡 Cek Insight Hari Ini', 'showDailyInsights')
      .addSeparator()
      .addItem('📊 Buat Rekap Bulanan', 'generateMonthlyRecap')
      .addItem('➕ Sisipkan Produk (Semua Sheet)', 'forceInsertRow') 
      .addItem('❌ Hapus Produk (Semua Sheet)', 'syncDeleteRow')
      .addSeparator() 
      .addItem('📅 Lompat ke Hari Ini', 'jumpToToday')
      .addItem('📄 Simpan PDF', 'exportCurrentSheetToPDF')
      .addItem('🙈 Sembunyikan Tgl Lama', 'hidePastSheets')
      .addToUi();
}

// ============================================================================
// 1. FITUR LAPORAN WHATSAPP (SUPER STRICT FILTER)
// ============================================================================
function generateWALaporan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  if (sheet.getLastRow() < 6) {
    ui.alert('⚠️ Sheet kosong. Pastikan di sheet Laporan Harian.');
    return;
  }

  // --- A. HEADER (DARI B2 & C2) ---
  var rawDay = sheet.getRange("B2").getDisplayValue(); 
  var dd = String(rawDay).trim().padStart(2, '0');
  var rawMonthYear = sheet.getRange("C2").getDisplayValue();
  var parts = rawMonthYear.trim().split(" ");
  var mmm = parts.length >= 2 ? parts[0].substring(0, 3).toUpperCase() : rawMonthYear;
  var yyyy = parts.length >= 2 ? parts[parts.length - 1] : "";
  
  var headerDate = dd + " : " + mmm + (yyyy ? " : " + yyyy : "");
  var text = "Assalamualaikum wr wb.\n";
  text += "*##### LAPORAN #####*\n";
  text += "*### " + headerDate + " ###*\n\n";

  // --- B. AMBIL DATA ---
  var lastRow = sheet.getLastRow();
  // Ambil data B7 s/d O
  var data = sheet.getRange(7, 2, lastRow - 6, 14).getValues();

  // --- C. PRE-PROCESSING (SAPU JAGAT FILTER) ---
  var cleanData = [];
  var lastRememberedName = "";
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    
    // Ambil data dari 3 kolom pertama untuk pengecekan
    var rawNo = String(row[0]).toUpperCase();   // Kolom B
    var rawNama = String(row[1]).toUpperCase(); // Kolom C
    var rawSpec = String(row[2]).toUpperCase(); // Kolom D
    
    // FILTER 1: CEK KATA KUNCI TERLARANG DI 3 KOLOM SEKALIGUS
    // Jika ada kata "TOTAL", "SUB", "JUMLAH", "REKAP" -> SKIP BARIS INI
    if (rawNama.includes("TOTAL") || rawNama.includes("SUB") || rawNama.includes("JUMLAH") ||
        rawNo.includes("TOTAL") || rawNo.includes("SUB") ||
        rawSpec.includes("TOTAL") || rawSpec.includes("SUB")) {
      continue; 
    }

    // FILTER 2: LOGIKA MERGED CELL (Memory)
    var originalName = row[1]; // Nama asli (bukan uppercase)
    
    if (originalName && originalName !== "") {
      lastRememberedName = String(originalName);
    } 
    
    var currentName = "";
    // Jika nama kosong, tapi ada Ukuran(Col E) ATAU Sisa(Col O), pinjam nama atasnya
    if (originalName === "" && lastRememberedName !== "" && (row[3] !== "" || row[13] !== "")) {
      currentName = lastRememberedName; 
    } else if (originalName !== "") {
      currentName = String(originalName);
    }
    
    if (currentName !== "") {
      row[1] = currentName; 
      cleanData.push(row);
    }
  }

  // --- D. PRODUK MASUK ---
  var sectionMasuk = "\n*----- PRODUK MASUK / HASIL ----*\n";
  var hasMasuk = false;
  for (var i = 0; i < cleanData.length; i++) {
    var nama = cleanData[i][1];
    var spec = cleanData[i][2]; // Kol D
    var uk = cleanData[i][3];   // Kol E
    var masuk = cleanData[i][5]; 
    
    if (masuk > 0) {
      var label = nama;
      if (uk) label += " " + uk;
      if (spec) label += " (" + spec + ")";
      sectionMasuk += "• " + label + "\n" + masuk + "\n"; 
      hasMasuk = true;
    }
  }
  if (!hasMasuk) sectionMasuk += "- Nihil -\n";

  // --- E. PRODUK KELUAR ---
  var sectionKeluar = "\n*-----PRODUK KELUAR-----*\n";
  var hasKeluar = false;
  for (var i = 0; i < cleanData.length; i++) {
    var nama = cleanData[i][1];
    var spec = cleanData[i][2];
    var uk = cleanData[i][3];
    var keluar = cleanData[i][12]; 
    
    if (keluar > 0) {
      var label = nama;
      if (uk) label += " " + uk;
      if (spec) label += " (" + spec + ")";
      sectionKeluar += "• " + label + "\n" + keluar + "\n";
      hasKeluar = true;
    }
  }
  if (!hasKeluar) sectionKeluar += "- Nihil -\n";

  // --- F. UPDATE STOCK ---
  var sectionStock = "\n*----UPDATE STOCK HARI INI----*\n";
  var lastProductName = "";
  
  for (var i = 0; i < cleanData.length; i++) {
    var nama = cleanData[i][1];
    var spec = cleanData[i][2];
    var uk = cleanData[i][3];
    var sisa = cleanData[i][13];

    var sisaFmt = (typeof sisa === 'number' && sisa < 10) ? "0" + sisa : sisa;

    if (nama !== lastProductName) {
      sectionStock += "\n*" + nama.toUpperCase() + "*\n"; 
      lastProductName = nama;
    }

    var detailLine = "";
    if (uk && uk !== "") {
      detailLine += uk + " \t: " + sisaFmt;
    } else {
      detailLine += sisaFmt; 
    }
    
    if (spec && spec !== "") {
      detailLine += " (" + spec + ")";
    }

    sectionStock += detailLine + "\n";
  }

  sectionStock += "\n***TERIMAKASIH ATAS PERHATIANNYA***";
  showTextPopup(text + sectionMasuk + sectionKeluar + sectionStock);
}

function showTextPopup(text) {
  var htmlOutput = HtmlService
    .createHtmlOutput('<textarea style="width:100%; height:400px; font-family:monospace; padding:10px;">' + text + '</textarea>')
    .setWidth(450)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Salin Laporan WhatsApp');
}

// ============================================================================
// 2. INSIGHT HARIAN
// ============================================================================
function showDailyInsights() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var lastRow = sheet.getLastRow();
  if (lastRow < 7) { ui.alert('Data kosong.'); return; }
  
  var values = sheet.getRange(7, 2, lastRow - 6, 14).getValues(); 
  var salesData = [];
  var lowStockData = [];
  var totalItemsOut = 0;
  var lastRemembered = ""; 

  for (var i = 0; i < values.length; i++) {
    var rawNo = String(values[i][0]).toUpperCase();
    var rawNama = String(values[i][1]).toUpperCase();
    var rawSpec = String(values[i][2]).toUpperCase();
    
    // FILTER SAMA SEPERTI DI ATAS
    if (rawNama.includes("TOTAL") || rawNama.includes("SUB") || 
        rawNo.includes("TOTAL") || rawNo.includes("SUB") ||
        rawSpec.includes("TOTAL") || rawSpec.includes("SUB")) {
      continue;
    }

    var name = "";
    var originalName = values[i][1];
    if (originalName && originalName !== "") {
      name = String(originalName);
      lastRemembered = name;
    } else if (lastRemembered !== "") {
      name = lastRemembered;
    } else { continue; }

    var sold = values[i][12]; 
    var stock = values[i][13]; 
    
    if (typeof sold === 'number') {
      totalItemsOut += sold;
      if (sold > 0) salesData.push({name: name, qty: sold});
    }
    if (typeof stock === 'number' && stock < 5) {
      lowStockData.push({name: name, qty: stock});
    }
  }

  salesData.sort(function(a, b) { return b.qty - a.qty; }); 
  var top3 = salesData.slice(0, 3);
  
  var message = "📅 PERFORM: " + sheet.getName() + "\n------------------\n";
  message += "📦 Total Keluar: " + totalItemsOut + " pcs\n\n";
  message += "🏆 TOP 3 TERLARIS:\n";
  if (top3.length > 0) {
    top3.forEach(function(item, index) { message += (index + 1) + ". " + item.name + " (" + item.qty + ")\n"; });
  } else { message += "- Nihil -\n"; }
  message += "\n⚠️ STOCK MENIPIS (<5):\n";
  if (lowStockData.length > 0) {
    lowStockData.forEach(function(item) { message += "• " + item.name + " (Sisa: " + item.qty + ")\n"; });
  } else { message += "✅ Aman.\n"; }
  ui.alert(message);
}

// ============================================================================
// 3. REKAP BULANAN
// ============================================================================
function generateMonthlyRecap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var recapSheetName = "REKAP_BULANAN";
  var recapSheet = ss.getSheetByName(recapSheetName);
  
  if (recapSheet) { ss.deleteSheet(recapSheet); }
  recapSheet = ss.insertSheet(recapSheetName, 0);
  
  var refSheet = ss.getSheets()[1]; 
  if (!isNaN(ss.getActiveSheet().getName())) { refSheet = ss.getActiveSheet(); }
  var lastRow = refSheet.getLastRow();
  var products = refSheet.getRange(7, 2, lastRow - 6, 4).getValues(); 
  
  var reportData = [];
  reportData.push(["NAMA PRODUK", "SPEC", "UKURAN", "TOTAL MASUK", "TOTAL JUAL", "STATUS"]);
  var lastProdName = "";

  for (var p = 0; p < products.length; p++) {
    var rawNo = String(products[p][0]).toUpperCase();
    var rawName = String(products[p][1]).toUpperCase();
    
    if (rawName.includes("TOTAL") || rawName.includes("SUB") || 
        rawNo.includes("TOTAL") || rawNo.includes("SUB")) {
      continue;
    }

    var prodName = "";
    var originalName = products[p][1];
    if (originalName && originalName !== "") {
      prodName = String(originalName);
      lastProdName = prodName;
    } else if (lastProdName !== "") {
      prodName = lastProdName;
    } else { continue; }
    
    var prodSpec = products[p][2];
    var prodSize = products[p][3];
    var totalProduksi = 0;
    var totalKeluar = 0;
    var allSheets = ss.getSheets();
    for (var s = 0; s < allSheets.length; s++) {
      var sheet = allSheets[s];
      if (!isNaN(sheet.getName())) {
        try {
          var prodVal = sheet.getRange(p + 7, 7).getValue(); 
          var outVal = sheet.getRange(p + 7, 14).getValue(); 
          if (typeof prodVal === 'number') totalProduksi += prodVal;
          if (typeof outVal === 'number') totalKeluar += outVal;
        } catch(e) {}
      }
    }
    
    var status = "Normal";
    if (totalKeluar > 50) status = "🔥 Best Seller"; 
    if (totalKeluar === 0) status = "💤 Slow Moving";
    
    reportData.push([prodName, prodSpec, prodSize, totalProduksi, totalKeluar, status]);
  }
  
  recapSheet.getRange(1, 1, reportData.length, 6).setValues(reportData);
  recapSheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#4a86e8").setFontColor("white");
  recapSheet.autoResizeColumns(1, 6);
  ui.alert('📊 Rekap Selesai!');
}

// ============================================================================
// 4. SISIPKAN PRODUK
// ============================================================================
function forceInsertRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getActiveSheet();
  var allSheets = ss.getSheets();
  var ui = SpreadsheetApp.getUi();
  var numCols = 15; 
  var currentSheetIndex = currentSheet.getIndex() - 1;
  var activeRange = currentSheet.getActiveRange();
  var rowIdx = activeRange.getRow();      
  var numRows = activeRange.getNumRows(); 
  
  if (rowIdx < 7) { ui.alert('⚠️ Pilih baris produk (Baris 7 ke bawah).'); return; }
  var sourceRange = currentSheet.getRange(rowIdx, 1, numRows, numCols);
  var sourceValues = sourceRange.getValues();
  var sourceFormulas = sourceRange.getFormulas();
  var sourceBackgrounds = sourceRange.getBackgrounds();
  var sourceFontWeights = sourceRange.getFontWeights(); 
  var finalData = [];
  for (var r = 0; r < numRows; r++) {
    var rowData = [];
    for (var c = 0; c < numCols; c++) {
      if (sourceFormulas[r][c] !== "") { rowData.push(sourceFormulas[r][c]); } 
      else { rowData.push(sourceValues[r][c]); }
    }
    finalData.push(rowData);
  }
  var response = ui.alert('KONFIRMASI TAMBAH','Sisipkan baris ke semua sheet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.NO) { return; }
  for (var i = currentSheetIndex + 1; i < allSheets.length; i++) {
    var targetSheet = allSheets[i];
    try { targetSheet.insertRows(rowIdx, numRows); } catch (e) {}
    var targetRange = targetSheet.getRange(rowIdx, 1, numRows, numCols);
    targetRange.setValues(finalData);
    targetRange.setBackgrounds(sourceBackgrounds);
    targetRange.setFontWeights(sourceFontWeights);
    targetSheet.getRange(rowIdx, 6, numRows, 8).setValue(""); 
  }
  ui.alert('✅ Selesai disisipkan.');
}

// ============================================================================
// 5. HAPUS PRODUK
// ============================================================================
function syncDeleteRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getActiveSheet();
  var allSheets = ss.getSheets();
  var ui = SpreadsheetApp.getUi();
  var currentSheetIndex = currentSheet.getIndex() - 1;
  var activeRange = currentSheet.getActiveRange();
  var rowIdx = activeRange.getRow();      
  var numRows = activeRange.getNumRows(); 
  if (rowIdx < 7) { ui.alert('⚠️ Jangan hapus header.'); return; }
  
  var productName = currentSheet.getRange(rowIdx, 3).getValue(); 
  var response = ui.alert('KONFIRMASI HAPUS','Hapus "' + productName + '" dari SEMUA sheet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.NO) { return; }
  
  for (var i = currentSheetIndex + 1; i < allSheets.length; i++) {
    allSheets[i].deleteRows(rowIdx, numRows);
  }
  currentSheet.deleteRows(rowIdx, numRows);
  ui.alert('🗑️ Produk dihapus.');
}

// ============================================================================
// 6. NAVIGASI
// ============================================================================
function jumpToToday() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = new Date().getDate().toString(); 
  var targetSheet = ss.getSheetByName(sheetName);
  if (targetSheet) { ss.setActiveSheet(targetSheet); } 
  else { SpreadsheetApp.getUi().alert('❌ Sheet tanggal ' + sheetName + ' tidak ada.'); }
}

function exportCurrentSheetToPDF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?exportFormat=pdf&format=pdf&size=A4&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenumbers=true&gridlines=true&gid=" + sheet.getSheetId();
  var token = ScriptApp.getOAuthToken();
  var blob = UrlFetchApp.fetch(url, {headers: {'Authorization': 'Bearer ' + token}}).getBlob().setName("Laporan " + sheet.getName() + ".pdf");
  DriveApp.createFile(blob);
  SpreadsheetApp.getUi().alert('📄 PDF tersimpan di Drive.');
}

function hidePastSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var today = new Date().getDate(); 
  var allSheets = ss.getSheets();
  for (var i = 0; i < allSheets.length; i++) {
    var name = allSheets[i].getName();
    if (!isNaN(name) && parseInt(name) < today) { allSheets[i].hideSheet(); }
  }
  SpreadsheetApp.getUi().alert('🙈 Sheet lama disembunyikan.');
}

function generateLaporanBoss() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DASHBOARD"); // Pastikan nama sheet sesuai
  
  // --- BAGIAN 1: MENGAMBIL DATA UTAMA ---
  // Sesuaikan koordinat cell (Baris, Kolom) dengan posisi TOTAL di sheet Anda
  // Asumsi berdasarkan Gambar 2: Baris Total ada di baris 34 (atau baris paling bawah)
  
  var lastRow = sheet.getLastRow(); // Mendeteksi baris terakhir otomatis
  
  // Ambil Total Terjual (Misal Kolom O, baris terakhir)
  var totalTerjual = sheet.getRange("O" + lastRow).getValue();
  
  // Ambil Total Omset (Misal Kolom P, baris terakhir)
  var totalOmset = sheet.getRange("P" + lastRow).getValue();
  
  // Ambil Total Profit (Misal Kolom S, baris terakhir)
  var totalProfit = sheet.getRange("S" + lastRow).getValue();

  // --- BAGIAN 2: MENCARI PRODUK TERLARIS (Opsional tapi Boss suka) ---
  var rangeNama = sheet.getRange("C7:C30").getValues(); // Kolom Nama Produk
  var rangeJual = sheet.getRange("O7:O30").getValues(); // Kolom Terjual
  
  var maxJual = 0;
  var produkLaris = "";
  
  for (var i = 0; i < rangeJual.length; i++) {
    if (rangeJual[i][0] > maxJual) {
      maxJual = rangeJual[i][0];
      produkLaris = rangeNama[i][0];
    }
  }

  // --- BAGIAN 3: FORMATTING RUPIAH ---
  var IDR = new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR' });
  
  // --- BAGIAN 4: MENYUSUN PESAN WA ---
  var bulan = "Januari 2026"; // Bisa dibuat dinamis mengambil dari cell header
  
  var pesan = "*LAPORAN KILAT PRODUKSI & SALES* 🚀\n";
  pesan += "🗓 Periode: " + bulan + "\n";
  pesan += "----------------------------------\n";
  pesan += "📦 Total Item Terjual: *" + totalTerjual + " Pcs*\n";
  pesan += "🏆 Produk Terlaris: *" + produkLaris + "* (" + maxJual + " Pcs)\n";
  pesan += "----------------------------------\n";
  pesan += "💰 *TOTAL OMSET: " + IDR.format(totalOmset) + "*\n";
  pesan += "💵 *TOTAL PROFIT: " + IDR.format(totalProfit) + "*\n";
  pesan += "----------------------------------\n";
  pesan += "_Laporan ini digenerate otomatis dari sistem._";

  // --- BAGIAN 5: OUTPUT (PILIH SALAH SATU) ---
  
  // OPSI A: Tampilkan di layar (Pop-up) untuk di-copy paste manual
  SpreadsheetApp.getUi().alert(pesan);
  
  // OPSI B: Jika Anda punya fungsi kirim WA (API), panggil di sini
  // kirimWAKepadaBoss(pesan); 
  
  Logger.log(pesan); // Untuk cek di log
  return pesan;
}

function generateRekapProduksi() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DASHBOARD");
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ Error: Sheet 'DASHBOARD' tidak ditemukan!");
    return;
  }

  // --- [SETTING] MASUKKAN NOMOR WA DEFAULT (UTAMA) ---
  // Ini adalah nomor yang akan muncul otomatis pertama kali
  var nomorWAPos = "628xxxxxxxxxx"; 
  
  // --- BAGIAN 1: AMBIL DATA PERIODE (DARI CELL C2) ---
  var periode = sheet.getRange("C2").getDisplayValue();
  
  if (periode === "") {
    var now = new Date();
    var bulanIndo = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    periode = bulanIndo[now.getMonth()] + " " + now.getFullYear();
  }

  // --- BAGIAN 2: DATA UTAMA ---
  var lastRow = sheet.getLastRow(); 
  var totalTerjual = sheet.getRange("O" + lastRow).getValue();
  var totalOmset = sheet.getRange("P" + lastRow).getValue();
  var totalProfit = sheet.getRange("S" + lastRow).getValue();

  // --- BAGIAN 3: LOGIKA PENGGABUNGAN NAMA PRODUK ---
  var limitBaris = lastRow - 1;
  var dataRange = sheet.getRange("C7:O" + limitBaris).getValues(); 
  
  var productSummary = {}; 
  var lastProductName = ""; 

  for (var i = 0; i < dataRange.length; i++) {
    var rowName = dataRange[i][0]; 
    var rowQty = dataRange[i][12]; 
    
    if (rowName !== "" && rowName != null) {
      lastProductName = rowName; 
    }
    
    var finalName = lastProductName;
    var qty = (typeof rowQty === 'number') ? rowQty : 0;

    if (productSummary[finalName]) {
      productSummary[finalName] += qty;
    } else {
      productSummary[finalName] = qty;
    }
  }

  // --- BAGIAN 4: MENYUSUN LIST PRODUK ---
  var breakdownText = "";
  for (var key in productSummary) {
    if (productSummary[key] > 0) {
      breakdownText += "• " + key + ": " + productSummary[key] + " Pcs\n";
    }
  }

  // --- BAGIAN 5: FORMATTING ---
  var IDR = new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', minimumFractionDigits: 0 });

  // --- BAGIAN 6: MENYUSUN PESAN FINAL ---
  var pesan = "*REKAP PRODUKSI* 🚀\n";
  pesan += "🗓 Periode: " + periode + "\n";
  pesan += "----------------------------------\n";
  pesan += "*RINCIAN ITEM TERJUAL:*\n";
  pesan += breakdownText; 
  pesan += "----------------------------------\n";
  pesan += "📦 TOTAL ITEM: *" + totalTerjual + " Pcs*\n";
  pesan += "💰 OMSET: *" + IDR.format(totalOmset) + "*\n";
  pesan += "💵 PROFIT: *" + IDR.format(totalProfit) + "*\n";
  pesan += "----------------------------------\n";
  pesan += "_Auto-generated by System_";

  // --- BAGIAN 7: POP-UP DENGAN FORM INPUT NOMOR ---
  
  // Kita menggunakan HTML Template dengan Script di dalamnya (Client-side Script)
  var htmlOutput = HtmlService.createHtmlOutput(
    '<html>' +
    '<head>' +
    '<style>' +
      'body { font-family: sans-serif; padding: 15px; text-align: center; color: #333; }' +
      'textarea { width: 100%; height: 200px; padding: 8px; border: 1px solid #ccc; border-radius: 4px; font-family: monospace; font-size: 12px; box-sizing: border-box; }' +
      '.input-group { margin-top: 15px; text-align: left; background: #f1f1f1; padding: 10px; border-radius: 8px; }' +
      'label { font-weight: bold; font-size: 13px; display: block; margin-bottom: 5px; }' +
      'input[type="text"] { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; box-sizing: border-box; }' +
      '.btn { cursor: pointer; display: block; width: 100%; background-color: #25D366; color: white; padding: 12px; text-decoration: none; border-radius: 50px; font-weight: bold; margin-top: 15px; font-size: 16px; border: none; box-shadow: 0 4px 6px rgba(0,0,0,0.1); transition: background 0.3s; }' +
      '.btn:hover { background-color: #128C7E; box-shadow: 0 6px 8px rgba(0,0,0,0.2); }' +
      '.note { font-size: 11px; color: #666; margin-top: 5px; }' +
    '</style>' +
    '</head>' +
    '<body>' +
      // Area Text Pesan
      '<div style="text-align:left; font-size:13px; margin-bottom:5px;">👇 Cek isi pesan:</div>' +
      '<textarea id="pesanBox">' + pesan + '</textarea>' +
      
      // Area Input Nomor WA
      '<div class="input-group">' +
        '<label>📱 Kirim ke Nomor WA:</label>' +
        '<input type="text" id="nomorWA" value="' + nomorWAPos + '" placeholder="Contoh: 62812345678">' +
        '<div class="note">*Gunakan kode negara (62) tanpa tanda plus (+)</div>' +
      '</div>' +

      // Tombol Kirim (Sekarang menjalankan Fungsi JS, bukan Link biasa)
      '<button class="btn" onclick="kirimSekarang()">🚀 Kirim Rekap ke WA</button>' +

      // Script Javascript untuk menangani Klik Tombol
      '<script>' +
        'function kirimSekarang() {' +
          'var pesan = document.getElementById("pesanBox").value;' +
          'var nomor = document.getElementById("nomorWA").value;' +
          // Validasi sederhana
          'if(nomor === "") { alert("Nomor WA tidak boleh kosong!"); return; }' +
          // Encode dan Buka Link
          'var url = "https://wa.me/" + nomor + "?text=" + encodeURIComponent(pesan);' +
          'window.open(url, "_blank");' +
          // Opsional: Tutup pop-up setelah kirim (hapus baris bawah jika ingin pop-up tetap terbuka)
          'google.script.host.close();' +
        '}' +
      '</script>' +
    '</body>' +
    '</html>'
  )
  .setWidth(400)
  .setHeight(480); // Tinggi ditambah sedikit agar muat form input

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Rekap Produksi Siap Kirim');
}
