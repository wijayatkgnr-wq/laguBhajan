// Konfigurasi Spreadsheet ID (Admin)
const ADMIN_SS_ID = "1_8N8CGpLsHJBvXF0b0qySWAXFAHF7OcxOf9YsJeTniQ";

/**
 * Fungsi utama untuk melayani request web app
 */
function doGet(e) {
  const isAdmin = e.parameter.p === 'admin';
  
  if (isAdmin) {
    return createFormattedHtml('Admin');
  }

  // --- LOGIKA PENGAMBILAN DATA LAGU ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const excludedSheets = ["Sheet1", "DataSeva", "DataSetting", "DataSevaBaru"];
  let allSongs = [];

  sheets.forEach(sh => {
    const name = sh.getName();
    
    if (!excludedSheets.includes(name)) {
      const lastRow = sh.getLastRow();
      if (lastRow > 1) {
        const values = sh.getRange(2, 1, lastRow - 1, 2).getValues();
        
        // Hanya ambil yang ID dan Judulnya tidak kosong
        const songs = values
          .filter(row => row[0] && row[1])
          .map(row => ({
            id: row[0],
            judul: row[1],
            kategori: name
          }));
        
        allSongs = allSongs.concat(songs);
      }
    }
  });

  // Urutkan berdasarkan judul secara alfabetis
  allSongs.sort((a, b) => a.judul.toString().localeCompare(b.judul.toString()));

  const tmp = HtmlService.createTemplateFromFile('Main');
  //tmp.dataBhajan = JSON.stringify(allSongs);
  
  return tmp.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Helper: Membuat template HTML dengan setting standar
 */
function createFormattedHtml(filename) {
  return HtmlService.createTemplateFromFile(filename)
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * MENGAMBIL DATA 1 SLOT UNTUK ADMIN
 */
function getSatuSlot(no) {
  const sh = SpreadsheetApp.openById(ADMIN_SS_ID).getSheetByName("DataSevaBaru");
  const lastRow = sh.getLastRow();
  
  if (lastRow < 2) return null;

  const data = sh.getRange(2, 1, lastRow - 1, 5).getValues();
  const target = data.find(row => row[0].toString() === no.toString());

  if (target) {
    const nama = target[2] ? target[2].toString().trim() : "";
    if (nama === "" || nama === "ID Tidak Terdaftar") return null;
    
    return {
      nama: nama,
      judul: target[3].toString()
    };
  }
  
  return null;
}

/**
 * UPDATE ID MEMBER
 */
function updateIdSeva(slot, idBaru) {
  const sh = SpreadsheetApp.openById(ADMIN_SS_ID).getSheetByName("DataSevaBaru");
  const lastRow = sh.getLastRow();
  
  if (lastRow < 2) return "❌ Data kosong.";

  const data = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  
  // Mencari index baris berdasarkan slot
  const index = data.findIndex(row => row[0].toString() === slot.toString());

  if (index !== -1) {
    const baris = index + 2;
    const rumus = `=IFERROR(VLOOKUP(B${baris}; $H:$I; 2; FALSE); "ID Tidak Terdaftar")`;
    
    sh.getRange(baris, 2).setValue(idBaru.toUpperCase());
    sh.getRange(baris, 3).setFormula(rumus);
    
    return `✅ Berhasil Ganti ID Slot ${slot}`;
  }

  return "❌ Slot tidak ditemukan.";
}



// --- FUNGSI SIMPAN SEVA (DAFTAR) ---
function simpanSeva(obj) {
  const sh = SpreadsheetApp.openById(ADMIN_SS_ID).getSheetByName("DataSevaBaru");
  
  // Pastikan baris target valid (minimal baris ke-2)
  const nomorSlot = parseInt(obj.nomor);
  const barisTarget = Math.max(2, nomorSlot + 1); 
  const waktuSekarang = new Date();

  // Update ID Member dan Nomor
  sh.getRange(barisTarget, 1, 1, 2).setValues([[obj.nomor, obj.idMember.toUpperCase()]]);
  
  // Set Rumus VLOOKUP
  const rumusNama = `=IFERROR(VLOOKUP(B${barisTarget}; $H:$I; 2; FALSE); "ID Tidak Terdaftar")`;
  sh.getRange(barisTarget, 3).setFormula(rumusNama);
  
  // Update Judul, ID Lagu, dan Timestamp
  sh.getRange(barisTarget, 4, 1, 3).setValues([[obj.judul, obj.idLagu, waktuSekarang]]);
  
  return "✅ Berhasil mendaftar!";
}






/**
 * HAPUS DATA DI SLOT SEVA
 */
function hapusSlotSeva(slot) {
  try {
    const sh = SpreadsheetApp.openById(ADMIN_SS_ID).getSheetByName("DataSevaBaru");
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return "❌ Data sudah kosong.";

    const dataSlot = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    const index = dataSlot.findIndex(r => r[0].toString() === slot.toString());

    if (index !== -1) {
      // Hapus kolom A sampai F (1 sampai 6) pada baris tersebut
      sh.getRange(index + 2, 1, 1, 6).clearContent();
      return `✅ Slot ${slot} berhasil dihapus!`;
    }
    return "⚠ Slot tidak ditemukan.";
  } catch (e) {
    return "❌ Error: " + e.message;
  }
}

/**
 * SIMPAN SETTING UTAMA
 */
function simpanSetting(obj) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataSetting");
  if (!sh) return "❌ Sheet DataSetting tidak ditemukan.";
  
  sh.getRange("A2:C2").setValues([[obj.hari, obj.waktu, obj.autoSlot]]);
  return "✅ Setting Disimpan!";
}

/**
 * CARI DETAIL LAGU DI SEMUA SHEET
 */
function getDetailLagu(idLagu) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const idDicari = idLagu.toString().trim().toLowerCase();
    const excluded = ["Sheet1", "DataSeva", "DataSetting", "DataSevaBaru"];

    for (let sh of sheets) {
      const shName = sh.getName();
      if (excluded.includes(shName)) continue;

      const data = sh.getDataRange().getValues();
      // Cari di kolom A (index 0)
      const foundRow = data.find(row => row[0] && row[0].toString().trim().toLowerCase() === idDicari);

      if (foundRow) {
        return {
          judul: foundRow[1] || "",
          lirik: foundRow[2] || "",
          filo:  foundRow[3] || "",
          glos:  foundRow[4] || "",
          mp3:   foundRow[5] || ""
        };
      }
    }
    return null;
  } catch (e) { return null; }
}

/**
 * MANAGEMENT MEMBER (H & I)
 */
function getSemuaMember() {
  const sh = SpreadsheetApp.openById(ADMIN_SS_ID).getSheetByName("DataSevaBaru");
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  return sh.getRange(2, 8, lastRow - 1, 2).getValues()
    .filter(r => r[0] !== "")
    .map(r => ({ id: r[0], nama: r[1] }));
}

function simpanMember(id, nama, isEdit) {
  const sh = SpreadsheetApp.openById(ADMIN_SS_ID).getSheetByName("DataSevaBaru");
  const lastRow = sh.getLastRow();
  const idUpper = id.toUpperCase().trim();
  
  const dataID = sh.getRange(2, 8, Math.max(lastRow - 1, 1), 1).getValues();
  const index = dataID.findIndex(r => r[0].toString().toUpperCase() === idUpper);

  if (index !== -1) {
    sh.getRange(index + 2, 9).setValue(nama);
    return `✅ Member ${idUpper} diperbarui!`;
  } else {
    sh.getRange(Math.max(lastRow + 1, 2), 8, 1, 2).setValues([[idUpper, nama]]);
    sh.getRange(2, 8, sh.getLastRow() - 1, 2).sort({column: 8, ascending: true});
    return "✅ Member Baru ditambah!";
  }
}

function hapusMember(id) {
  const sh = SpreadsheetApp.openById(ADMIN_SS_ID).getSheetByName("DataSevaBaru");
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return "❌ Data kosong.";

  const dataID = sh.getRange(2, 8, lastRow - 1, 1).getValues();
  const index = dataID.findIndex(r => r[0].toString().toUpperCase() === id.toUpperCase());

  if (index !== -1) {
    // Gunakan clearContent lalu sort agar tidak merusak baris kolom lain
    sh.getRange(index + 2, 8, 1, 2).clearContent();
    sh.getRange(2, 8, lastRow - 1, 2).sort({column: 8, ascending: true});
    return `✅ Member ${id} dihapus!`;
  }
  return "❌ Member tidak ditemukan.";
}








/**
 * FUNGSI UNTUK MENERIMA DATA DARI ADMIN.HTML
 */
function kirimKeSheet(obj) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const kategori = obj.kategori.trim();
    
    // 1. CEK TAB: Jika belum ada, buat tab baru
    let sh = ss.getSheetByName(kategori);
    if (!sh) {
      sh = ss.insertSheet(kategori);
      // Buat header otomatis di baris pertama
      sh.getRange(1, 1, 1, 6).setValues([["ID", "Judul", "Lirik & Arti", "Filosofi", "Glosarium", "ID MP3"]]);
      sh.setFrozenRows(1); // Bekukan baris pertama agar rapi
    }

    const lastRow = sh.getLastRow();
    let nextNumber = 1;

    // 2. TENTUKAN NOMOR ID: Lanjutkan dari yang terakhir atau mulai dari 1
    if (lastRow >= 2) {
      const lastIdValue = sh.getRange(lastRow, 1).getValue().toString();
      const lastNumberPart = lastIdValue.substring(lastIdValue.lastIndexOf("-") + 1);
      const lastNum = parseInt(lastNumberPart);
      
      if (!isNaN(lastNum)) {
        nextNumber = lastNum + 1;
      }
    }

    // 3. FORMAT ID (Contoh: Ganesha-001)
    const formattedNumber = nextNumber.toString().padStart(3, '0');
    const customId = `${kategori}-${formattedNumber}`;

    // 4. SIMPAN DATA
    sh.appendRow([
      customId,
      obj.judul,
      obj.lirikArti,
      obj.filosofi,
      obj.glosarium,
      obj.idMp3
    ]);

    return "✅ Berhasil!\nTab: " + kategori + "\nID: " + customId;
  } catch (e) {
    return "❌ Gagal: " + e.message;
  }
}



function getDataLengkapSeva() {
  const ss = SpreadsheetApp.openById(ADMIN_SS_ID);
  const shSet = ss.getSheetByName("DataSetting");
  const shSeva = ss.getSheetByName("DataSevaBaru");
  
  let config = { 
    hari: "4", 
    waktu: "19:00", 
    autoSlot: false, 
    pwAdmin: "", 
    backupFolderId: "", // B5
    logHapusId: "",      // B9
    members: {}, 
    quotes: [], 
    warna: [], 
    daftarLatihan: [], 
    slotTerisi: [] 
  };
  
  if (shSet) {
    // Pengambilan Data Berdasarkan Cell Spesifik (Kolom B)
    config.hari = shSet.getRange("B1").getDisplayValue();
    config.waktu = shSet.getRange("B2").getDisplayValue();
    config.autoSlot = shSet.getRange("B3").getValue().toString().toUpperCase() === "TRUE";
    config.pwAdmin = shSet.getRange("B4").getDisplayValue();
    config.backupFolderId = shSet.getRange("B5").getDisplayValue();
    config.logHapusId = shSet.getRange("B9").getDisplayValue();
    
    const lastRowSet = shSet.getLastRow();
    if (lastRowSet >= 2) {
       // Ambil Data Warna (Kolom F2 ke bawah) dan Quotes (Kolom G2 ke bawah)
       // Kita ambil dari baris 2 sampai baris terakhir, kolom F sampai G (2 kolom)
       const extraData = shSet.getRange(2, 6, lastRowSet - 1, 2).getValues();
       
       extraData.forEach(r => {
         // r[0] adalah Kolom F (Warna)
         if (r[0]) config.warna.push(r[0].toString().trim());
         
         // r[1] adalah Kolom G (Quotes)
         if (r[1]) {
           const matches = r[1].toString().match(/"([^"]+)"/g);
           if (matches && matches.length >= 2) {
             config.quotes.push({ 
               text: matches[0].replace(/"/g, ""), 
               source: matches[1].replace(/"/g, "") 
             });
           }
         }
       });
    }
  }

  if (shSeva) {
    const lastRowSeva = shSeva.getLastRow();
    if (lastRowSeva >= 2) {
      const rawData = shSeva.getRange(2, 1, lastRowSeva - 1, 6).getValues();
      config.daftarLatihan = rawData.map(row => row.map(cell => (cell == null ? "" : cell.toString())));
      
      // Ambil slot terisi (Kolom A / Index 0) jika kolom Nama (Kolom C / Index 2) tidak kosong
      config.slotTerisi = rawData
        .filter(r => {
          let n = (r[2] || "").toString().trim();
          return n !== "" && n !== "0" && n !== "ID Tidak Terdaftar" && n !== "#N/A";
        })
        .map(r => r[0].toString());
      
      // Ambil Database Member (Kolom H & I / Index 7 & 8)
      const rawMembers = shSeva.getRange(2, 8, lastRowSeva - 1, 2).getValues();
      rawMembers.forEach(m => { 
        if (m[0]) config.members[m[0].toString().toLowerCase().trim()] = m[1] || m[0]; 
      });
    }
  }
  return config;
}




function getSemuaLagu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // KUNCI: Urutkan sheet berdasarkan posisi tab-nya (getIndex)
  const sheets = ss.getSheets().sort((a, b) => a.getIndex() - b.getIndex());
  
  const excludedSheets = ["Sheet1", "DataSeva", "DataSetting", "DataSevaBaru"];
  let allSongs = [];

  sheets.forEach(sh => {
    const name = sh.getName();
    if (!excludedSheets.includes(name)) {
      const lastRow = sh.getLastRow();
      if (lastRow > 1) {
        const values = sh.getRange(2, 1, lastRow - 1, 2).getValues();
        
        // Ambil lagu dari tab ini saja
        let songsInSheet = values
          .filter(row => row[0] && row[1])
          .map(row => ({ id: row[0], judul: row[1], kategori: name }));
        
        // URUTKAN JUDUL: Hanya untuk lagu di dalam tab ini saja
        songsInSheet.sort((a, b) => a.judul.toString().localeCompare(b.judul.toString()));
        
        // Gabungkan ke daftar utama tanpa merusak urutan tab
        allSongs = allSongs.concat(songsInSheet);
      }
    }
  });

  // JANGAN ADA allSongs.sort di sini agar urutan tab tidak berantakan lagi
  return allSongs;
}






/** * AUTO BACKUP & RESET - VERSI EFISIENSI TINGGI */
function autoBackupDanResetSeva() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shSetting = ss.getSheetByName("DataSetting");
  const shSeva = ss.getSheetByName("DataSevaBaru");

  if (!shSetting || !shSeva) return;

  try {
    // 1. Ambil semua konfigurasi sekaligus dalam satu array (Efisien)
    const configRange = shSetting.getRange("B1:B9");
    const configValues = configRange.getValues();

    const hariBhajan = parseInt(configValues[0][0]); // B1
    const idFolderBackup = configValues[4][0];      // B5
    const logTerakhir = configValues[8][0].toString(); // B9
    
    const now = new Date();
    const tglHariIni = Utilities.formatDate(now, "GMT+08:00", "yyyy-MM-dd");
    const hariIni = now.getDay();
    const hariEksekusi = (hariBhajan + 1) % 7;

    // 2. Cek Jadwal
    if (hariIni === hariEksekusi) {
      if (logTerakhir.includes(tglHariIni)) return;

      // --- LOGIKA PENAMAAN FILE ---
      const daftarHari = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
      const tglBhajan = new Date(now);
      tglBhajan.setDate(now.getDate() - 1); 
      const formatTgl = Utilities.formatDate(tglBhajan, "GMT+08:00", "yyyy-MM-d");
      const namaFile = `${formatTgl} ${daftarHari[tglBhajan.getDay()]}`;

      // --- PROSES BACKUP & PINDAH FOLDER (Single Action) ---
      const folder = DriveApp.getFolderById(idFolderBackup);
      const newSS = SpreadsheetApp.create(namaFile);
      const fileId = newSS.getId();
      DriveApp.getFileById(fileId).moveTo(folder);

      // --- COPY DATA & HAPUS SHEET BAWAAN (Efisien) ---
      shSeva.copyTo(newSS).setName("DataSeva");
      newSS.deleteSheet(newSS.getSheets()[0]); 

      // --- PROSES RESET DATA (Bulk Action) ---
      const lastRow = shSeva.getLastRow();
      if (lastRow > 1) {
        shSeva.getRange(2, 1, lastRow - 1, 6).clearContent();
      }

      // --- AUTO-RESTORE JADWAL & UPDATE LOG (Batch Update) ---
      // Kita update semua sel sekaligus agar script berjalan secepat kilat
      shSetting.getRange("B1:B3").setValues([[4], ["19:00"], [false]]); 
      shSetting.getRange("B9").setValue(`Sukses: ${tglHariIni}. Jadwal Kembali ke Kamis.`);

    } else {
      const sisa = (hariEksekusi + 7 - hariIni) % 7;
      shSetting.getRange("B9").setValue(`Menunggu H+1 (${sisa} hari lagi)`);
    }

  } catch (e) {
    shSetting.getRange("B9").setValue("ERR: " + e.message);
  }
}




/**
 * BACKUP & RESET MANUAL - BISA DIJALANKAN KAPAN SAJA
 */
function hapusDanBackupManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shSetting = ss.getSheetByName("DataSetting");
  const shSeva = ss.getSheetByName("DataSevaBaru");

  if (!shSetting || !shSeva) throw new Error("Sheet tidak ditemukan!");

  try {
    const configValues = shSetting.getRange("B1:B9").getValues();
    const idFolderBackup = configValues[4][0]; // Ambil ID Folder dari B5
    
    const now = new Date();
    const tglHariIni = Utilities.formatDate(now, "GMT+08:00", "yyyy-MM-dd");
    
    const daftarHari = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
    const namaFile = `${tglHariIni} ${daftarHari[now.getDay()]} Manual`;

    const folder = DriveApp.getFolderById(idFolderBackup);
    const newSS = SpreadsheetApp.create(namaFile);
    const fileId = newSS.getId();
    DriveApp.getFileById(fileId).moveTo(folder);

    shSeva.copyTo(newSS).setName("DataSeva");
    newSS.deleteSheet(newSS.getSheets()[0]);

    const lastRow = shSeva.getLastRow();
    if (lastRow > 1) {
      shSeva.getRange(2, 1, lastRow - 1, 6).clearContent();
    }

    // Update Setting & Log
    shSetting.getRange("B1:B3").setValues([[4], ["19:00"], [false]]);
    shSetting.getRange("B9").setValue(`Sukses Manual: ${tglHariIni}`);
    
    return "Sukses"; // Kirim respon balik ke HTML

  } catch (e) {
    shSetting.getRange("B9").setValue("ERR MANUAL: " + e.message);
    throw new Error(e.message); // Kirim pesan error ke withFailureHandler
  }
}






function getSheetData(fileId) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(fileId);
  if (cached) return JSON.parse(cached);
  try {
    var ss = SpreadsheetApp.openById(fileId);
    var values = ss.getSheets()[0].getDataRange().getValues();
    var data = values.slice(1).map(r => ({
      slot: r[0] || "-", nama: r[2] || "-", judul: r[3] || "-", idLagu: r[4] || "" 
    })).filter(r => r.nama !== "-" && r.nama !== "");
    cache.put(fileId, JSON.stringify(data), 600);
    return data;
  } catch (e) { return []; }
}


function getInitialData() {
  var folderId = "17L_QtPw_xZB0fKzud6vvsgZr9D1W6B0P";
  var folder = DriveApp.getFolderById(folderId);
  
  // Ambil file dengan iterator
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var fileList = [];
  var limit = 50; // Ambil 50 file terbaru saja agar super cepat
  var count = 0;

  // Loop ini sekarang lebih efisien
  while (files.hasNext() && count < limit) {
    var f = files.next();
    fileList.push({
      id: f.getId(), 
      name: f.getName()
    });
    count++;
  }

  // Urutkan berdasarkan nama (Z ke A)
  fileList.sort((a, b) => b.name.localeCompare(a.name));

  // Pengecekan keamanan jika folder kosong
  if (fileList.length === 0) {
    return {
      files: [], 
      firstFileData: []
    };
  }

  // Ambil data dari file yang paling atas setelah di-sort
  return { 
    files: fileList, 
    firstFileData: getSheetData(fileList[0].id) 
  };
}


function getSheetData(fileId) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(fileId);
  if (cached) return JSON.parse(cached);
  try {
    var ss = SpreadsheetApp.openById(fileId);
    var values = ss.getSheets()[0].getDataRange().getValues();
    var data = values.slice(1).map(r => ({
      slot: r[0] || "-", nama: r[2] || "-", judul: r[3] || "-", idLagu: r[4] || "" 
    })).filter(r => r.nama !== "-" && r.nama !== "");
    cache.put(fileId, JSON.stringify(data), 600);
    return data;
  } catch (e) { return []; }
}

