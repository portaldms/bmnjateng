/**
 * ============================================================================
 * PORTAL DATA HUB & LINK MANAGER (v2.0) - BACKEND JSON API ENGINE
 * File: Code.gs
 * ============================================================================
 */

/**
 * Entry Point HTTP GET (Public Read / Health Check)
 */
function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : null;
  var result;

  try {
    if (action === 'getPortalData' || !action) {
      result = getPortalData();
    } else {
      result = { status: 'error', message: 'Endpoint GET tidak valid.' };
    }
  } catch (err) {
    result = { status: 'error', message: err.toString() };
  }

  return createJsonResponse(result);
}

/**
 * Entry Point HTTP POST (Action Router & Mutation Engine)
 */
function doPost(e) {
  var result;
  try {
    var payload = {};
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    }

    var action = payload.action || (e.parameter ? e.parameter.action : '');

    switch (action) {
      case 'getPortalData':
        result = getPortalData();
        break;
      case 'login':
        result = verifyAdminLogin(payload.username, payload.password);
        break;
      case 'saveLink':
        result = saveLinkData(payload.linkData);
        break;
      case 'deleteLink':
        result = deleteLinkData(payload.linkId);
        break;
      case 'saveCategory':
        result = saveCategoryData(payload.catData);
        break;
      default:
        result = { status: 'error', message: 'Action backend tidak dikenali: ' + action };
    }
  } catch (err) {
    result = { status: 'error', message: err.toString() };
  }

  return createJsonResponse(result);
}

/**
 * Helper Utility: Membentuk output JSON Standar dengan CORS Header
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Helper Utility: Mengubah isi Sheet menjadi Array of Objects berdasarkan header baris ke-1
 */
function sheetToObjects(sheet) {
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];

  const headers = rows[0];
  return rows.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, idx) => {
      obj[h] = row[idx];
    });
    return obj;
  });
}

/**
 * Mengambil seluruh data Kategori dan Hyperlink dari Spreadsheet
 */
function getPortalData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const catSheet = ss.getSheetByName('Categories');
    const linkSheet = ss.getSheetByName('Links');

    if (!catSheet || !linkSheet) {
      return { 
        status: 'error', 
        message: 'Sheet "Categories" atau "Links" belum dibuat di Google Spreadsheet!' 
      };
    }

    const categories = sheetToObjects(catSheet);
    const links = sheetToObjects(linkSheet);

    return {
      status: 'success',
      categories: categories,
      links: links
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

/**
 * Memverifikasi Login Admin
 */
function verifyAdminLogin(username, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const adminSheet = ss.getSheetByName('Admins');

    // Fallback kredensial jika Sheet Admins belum dibuat
    if (!adminSheet) {
      if (username === 'admin' && password === 'admin123') {
        return { 
          status: 'success', 
          token: 'SESSION_' + Date.now(), 
          role: 'Super Admin' 
        };
      }
      return { status: 'error', message: 'Username atau Password salah (Default: admin / admin123)' };
    }

    const admins = sheetToObjects(adminSheet);
    const userMatch = admins.find(a => 
      String(a.Username).trim() === String(username).trim() && 
      String(a.Password).trim() === String(password).trim()
    );

    if (userMatch) {
      return {
        status: 'success',
        token: 'TOKEN_' + Date.now() + '_' + Math.random().toString(36).substring(2, 9),
        role: userMatch.Role || 'Super Admin'
      };
    }

    return { status: 'error', message: 'Kombinasi Username dan Password tidak cocok.' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

/**
 * Menyimpan (Tambah baru atau Update) Data Hyperlink
 */
function saveLinkData(linkData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Links');
    if (!sheet) return { status: 'error', message: 'Sheet "Links" tidak ditemukan.' };

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    // Cari baris jika melakukan Edit (berdasarkan Link_ID)
    if (linkData.Link_ID) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(linkData.Link_ID)) {
          rowIndex = i + 1;
          break;
        }
      }
    }

    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, 8).setValues([[
        linkData.Link_ID,
        linkData.Category_ID,
        linkData.Sub_Menu || 'Umum',
        linkData.Title,
        linkData.URL,
        linkData.Description || '',
        linkData.Badge || '',
        linkData.Status || 'Active'
      ]]);
    } else {
      const newId = 'LNK-' + Date.now();
      sheet.appendRow([
        newId,
        linkData.Category_ID,
        linkData.Sub_Menu || 'Umum',
        linkData.Title,
        linkData.URL,
        linkData.Description || '',
        linkData.Badge || '',
        'Active'
      ]);
    }
    return { status: 'success' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

/**
 * Menghapus data Hyperlink berdasarkan Link_ID
 */
function deleteLinkData(linkId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Links');
    if (!sheet) return { status: 'error', message: 'Sheet "Links" tidak ditemukan.' };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(linkId)) {
        sheet.deleteRow(i + 1);
        return { status: 'success' };
      }
    }
    return { status: 'error', message: 'Link ID tidak ditemukan.' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

/**
 * Menambahkan Kategori Baru pada Sheet Categories
 */
function saveCategoryData(catData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Categories');
    if (!sheet) return { status: 'error', message: 'Sheet "Categories" tidak ditemukan.' };

    const newId = 'CAT-' + Date.now();
    sheet.appendRow([
      newId,
      catData.Category_Name,
      catData.Description || '',
      catData.Icon || 'fa-folder'
    ]);
    return { status: 'success' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
