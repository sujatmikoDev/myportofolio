/**
 * === FILE 1: Code.gs (Google Apps Script Backend) ===
 * Backend untuk Aplikasi Portofolio SPA (SujatmikoDev)
 */

const FOLDER_NAME = "SujatmikoDev_Images";
const SPREADSHEET_NAME = "SujatmikoDev_Database";

// 1. SETUP DATABASE
function setupDatabase() {
  let props = PropertiesService.getScriptProperties();
  
  let folderId = props.getProperty('FOLDER_ID');
  if (!folderId) {
    let folders = DriveApp.getFoldersByName(FOLDER_NAME);
    let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    props.setProperty('FOLDER_ID', folder.getId());
  }

  let dbId = props.getProperty('DB_ID');
  let ss;
  if (!dbId) {
    let files = DriveApp.getFilesByName(SPREADSHEET_NAME);
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create(SPREADSHEET_NAME);
    }
    props.setProperty('DB_ID', ss.getId());
  } else {
    ss = SpreadsheetApp.openById(dbId);
  }

  const sheets = {
    'Config': ['Key', 'Value'],
    'Karya': ['ID', 'Judul', 'Kategori', 'Link', 'Deskripsi', 'Gambar1', 'Gambar2', 'Gambar3'],
    'Layanan': ['ID', 'Nama', 'Icon', 'Deskripsi'],
    'Skills': ['ID', 'Nama', 'Persentase', 'Icon']
  };

  for (let sheetName in sheets) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(sheets[sheetName]);
      sheet.getRange(1, 1, 1, sheets[sheetName].length).setFontWeight("bold");
    }
  }

  let configSheet = ss.getSheetByName('Config');
  let data = configSheet.getDataRange().getValues();
  if (!data.some(row => row[0] === 'ADMIN_PASSWORD')) {
    configSheet.appendRow(['ADMIN_PASSWORD', '240be518fabd2724ddb6f04eeb1da5967448d7e831c08c8fa822809f74c720a9']);
  }
  
  // Inject default Layanan jika kosong
  let laySheet = ss.getSheetByName('Layanan');
  if (laySheet.getLastRow() <= 1) {
    laySheet.appendRow(['L1', 'Web Development', 'fas fa-laptop-code', 'Pembuatan website company profile, landing page, hingga aplikasi web custom.']);
    laySheet.appendRow(['L2', 'IT Support', 'fas fa-network-wired', 'Instalasi jaringan, troubleshooting hardware/software, dan maintenance sistem.']);
    laySheet.appendRow(['L3', 'Konten Digital', 'fas fa-paint-brush', 'Desain grafis, pembuatan konten media sosial, dan materi promosi digital.']);
  }

  return "Setup Berhasil! Silahkan cek Google Drive Anda.";
}

// 2. ROUTING & CONTROLLER
function doPost(e) {
  try {
    let payload = JSON.parse(e.postData.contents);
    let action = payload.action;
    let data = payload.data;
    let result;

    switch(action) {
      case 'login': result = doLogin(data.password); break;
      case 'getData': result = getAllData(); break;
      case 'saveKarya': result = checkAuth(payload.token) ? saveRow('Karya', data) : {status: 'error'}; break;
      case 'deleteKarya': result = checkAuth(payload.token) ? deleteRow('Karya', data.id) : {status: 'error'}; break;
      case 'saveData': result = checkAuth(payload.token) ? saveRow(data.sheet, data.payload) : {status: 'error'}; break;
      case 'deleteData': result = checkAuth(payload.token) ? deleteRow(data.sheet, data.id) : {status: 'error'}; break;
      case 'uploadImage': result = checkAuth(payload.token) ? uploadFileToDrive(data.filename, data.base64, data.mimeType) : {status: 'error'}; break;
      case 'changePassword': result = checkAuth(payload.token) ? changePassword(data.oldPwd, data.newPwd) : {status: 'error'}; break;
      case 'updateConfig': result = checkAuth(payload.token) ? updateConfig(data.key, data.value) : {status: 'error'}; break;
      default: result = { status: 'error', message: 'Unknown action' };
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) { return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON); }

// 3. HELPER FUNCTIONS
function getDb() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('DB_ID')); }

// ---- KEAMANAN LOGIN ----
// Browser mengirim SHA-256 dari password. Server menghash sekali lagi (pakai garam rahasia di Script Properties)
// dan hanya hasil itu yang disimpan di Sheet. Jadi, kalaupun isi Sheet bocor, hash di dalamnya TIDAK bisa
// dipakai untuk login. Format lama (hash polos) tetap diterima dan otomatis di-upgrade saat login berhasil.
const LOGIN_MAX_FAIL = 5;     // maksimal salah password berturut-turut
const LOGIN_LOCK_SEC = 600;   // lama kunci (detik) setelah mencapai batas

function toHex_(bytes) { return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join(''); }

function getSalt_() {
  const props = PropertiesService.getScriptProperties();
  let salt = props.getProperty('PWD_SALT');
  if (!salt) { salt = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('PWD_SALT', salt); }
  return salt;
}

function serverHash_(clientHash) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, clientHash + ':' + getSalt_(), Utilities.Charset.UTF_8);
  return 'h2:' + toHex_(raw);
}

function passwordMatches_(clientHash, stored) {
  stored = String(stored);
  if (stored.indexOf('h2:') === 0) return stored === serverHash_(clientHash);
  return stored === clientHash;   // format lama
}

function doLogin(hashedPassword) {
  const cache = CacheService.getScriptCache();
  const fails = parseInt(cache.get('login_fail') || '0', 10);
  if (fails >= LOGIN_MAX_FAIL) return { status: 'error', message: 'Terlalu banyak percobaan. Coba lagi beberapa menit lagi.' };

  const sheet = getDb().getSheetByName('Config');
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.findIndex(row => row[0] === 'ADMIN_PASSWORD');
  if (rowIdx < 0 || !hashedPassword) return { status: 'error', message: 'Password salah!' };

  const stored = data[rowIdx][1];
  if (passwordMatches_(String(hashedPassword), stored)) {
    if (String(stored).indexOf('h2:') !== 0) sheet.getRange(rowIdx + 1, 2).setValue(serverHash_(String(hashedPassword)));
    cache.remove('login_fail');
    const token = Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty('AUTH_TOKEN', token);
    return { status: 'success', token: token };
  }
  cache.put('login_fail', String(fails + 1), LOGIN_LOCK_SEC);
  return { status: 'error', message: 'Password salah!' };
}

function checkAuth(token) {
  let storedToken = PropertiesService.getScriptProperties().getProperty('AUTH_TOKEN');
  return token && storedToken && token === storedToken;
}

function changePassword(oldHash, newHash) {
  let sheet = getDb().getSheetByName('Config');
  let data = sheet.getDataRange().getValues();
  let rowIdx = data.findIndex(row => row[0] === 'ADMIN_PASSWORD');
  if (rowIdx >= 0 && oldHash && newHash && passwordMatches_(String(oldHash), data[rowIdx][1])) {
    sheet.getRange(rowIdx + 1, 2).setValue(serverHash_(String(newHash)));
    return { status: 'success', message: 'Password berhasil diubah' };
  }
  return { status: 'error', message: 'Password lama tidak cocok' };
}

function updateConfig(key, value) {
  let sheet = getDb().getSheetByName('Config');
  let data = sheet.getDataRange().getValues();
  let rowIdx = data.findIndex(row => row[0] === key);
  if (rowIdx >= 0) sheet.getRange(rowIdx + 1, 2).setValue(value);
  else sheet.appendRow([key, value]);
  return { status: 'success' };
}

function getAllData() {
  let ss = getDb();
  let result = {};
  ['Karya', 'Config', 'Layanan'].forEach(sheetName => {
    let data = ss.getSheetByName(sheetName).getDataRange().getValues();
    if(sheetName === 'Config') {
       // PENTING: jangan kirim ADMIN_PASSWORD / token ke publik
       result[sheetName] = data
         .filter(row => row[0] && row[0] !== 'ADMIN_PASSWORD' && row[0] !== 'AUTH_TOKEN' && row[0] !== 'Key')
         .map(row => ({ key: row[0], value: row[1] }));
    } else {
       let headers = data.shift();
       result[sheetName] = data.map(row => {
         let obj = {};
         headers.forEach((h, i) => obj[h.toLowerCase()] = row[i]);
         return obj;
       });
    }
  });
  return { status: 'success', data: result };
}

function ensureColumns(sheet, keys) {
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase());
  keys.forEach(k => {
    if (k && headers.indexOf(k.toLowerCase()) === -1) {
      let col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(k.charAt(0).toUpperCase() + k.slice(1)).setFontWeight("bold");
      headers.push(k.toLowerCase());
    }
  });
}

function saveRow(sheetName, data) {
  let sheet = getDb().getSheetByName(sheetName);
  ensureColumns(sheet, Object.keys(data).filter(k => k !== 'id'));
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let newRow = headers.map(h => data[h.toLowerCase()] || "");
  
  if (data.id) {
    let allData = sheet.getDataRange().getValues();
    let idIndex = headers.findIndex(h => h.toLowerCase() === 'id');
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][idIndex] === data.id) {
        sheet.getRange(i + 1, 1, 1, newRow.length).setValues([newRow]);
        return { status: 'success' };
      }
    }
  }
  
  newRow[0] = "ID_" + new Date().getTime();
  sheet.appendRow(newRow);
  return { status: 'success' };
}

function deleteRow(sheetName, id) {
  let sheet = getDb().getSheetByName(sheetName);
  let data = sheet.getDataRange().getValues();
  let idIndex = data[0].findIndex(h => h.toLowerCase() === 'id');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === id) {
      sheet.deleteRow(i + 1);
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'Data tidak ditemukan' };
}

function uploadFileToDrive(filename, base64Data, mimeType) {
  try {
    if (!base64Data || base64Data.indexOf(',') === -1) return { status: 'error', message: 'Data gambar tidak valid' };
    let props = PropertiesService.getScriptProperties();
    if (!props.getProperty('FOLDER_ID')) setupDatabase();   // otomatis siapkan folder bila belum ada
    let folder = DriveApp.getFolderById(props.getProperty('FOLDER_ID'));
    let blob = Utilities.newBlob(Utilities.base64Decode(base64Data.split(',')[1]), mimeType || 'image/jpeg', filename);
    let file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* folder sudah publik */ }
    return { status: 'success', url: 'https://drive.google.com/uc?export=view&id=' + file.getId() };
  } catch (err) {
    return { status: 'error', message: 'Upload gagal: ' + err.toString() };
  }
}
