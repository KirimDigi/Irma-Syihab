// Google Apps Script untuk Irma & Syihab
// Spreadsheet: https://docs.google.com/spreadsheets/d/1GpGzZWrxae7C449obDyXns-DR2baMOylHfkAZdD9E54/edit?gid=0#gid=0
// Sheet: Sheet1
// Kolom: timestamp | nama tamu | ucapan | konfirmasi kehadiran | jumlah tamu
// Deploy: Publish > Deploy as web app > Anyone, even anonymous > Copy URL

const SPREADSHEET_ID = '1GpGzZWrxae7C449obDyXns-DR2baMOylHfkAZdD9E54';
const SHEET_NAME = 'Sheet1';
const HEADER = ['timestamp','nama tamu','ucapan','konfirmasi kehadiran','jumlah tamu'];

function ensureSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER);
    sheet.getRange(1,1,1,HEADER.length).setFontWeight('bold').setBackground('#415692').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  } else {
    // pastikan header benar
    const firstRow = sheet.getRange(1,1,1,HEADER.length).getValues()[0];
    const needHeader = firstRow.join('') === '';
    if (needHeader) {
      sheet.getRange(1,1,1,HEADER.length).setValues([HEADER]);
    }
  }
  return sheet;
}

function doGet(e) {
  try {
    const sheet = ensureSheet();
    const rows = sheet.getDataRange().getValues();
    const data = [];
    // skip header
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      // skip empty rows
      if (!r[0] && !r[1] && !r[2]) continue;
      data.push({
        timestamp: r[0] ? Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '',
        nama: r[1] || '',
        ucapan: r[2] || '',
        kehadiran: r[3] || '',
        jumlah: r[4] || ''
      });
    }
    // terbaru di atas
    data.reverse();
    return ContentService.createTextOutput(JSON.stringify({status:'success', data: data}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error', message: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const sheet = ensureSheet();
    let payload = {};
    if (e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (err) {
        // fallback ke form-encoded
        payload = e.parameter || {};
      }
    } else {
      payload = e.parameter || {};
    }

    const timestamp = payload.timestamp || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const nama = (payload.nama || payload['nama tamu'] || payload.name || '').toString().trim();
    const ucapan = (payload.ucapan || payload.comment || payload.message || '').toString().trim();
    const kehadiran = (payload.kehadiran || payload['konfirmasi kehadiran'] || payload.attendance || payload.hadir || 'Hadir').toString().trim();
    const jumlah = (payload.jumlah || payload['jumlah tamu'] || payload.guest_count || payload.jumlahTamu || '1').toString().trim();

    if (!nama || !ucapan) {
      return ContentService.createTextOutput(JSON.stringify({status:'error', message:'Nama dan ucapan wajib diisi'}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    sheet.appendRow([timestamp, nama, ucapan, kehadiran, jumlah]);

    return ContentService.createTextOutput(JSON.stringify({status:'success', message:'Ucapan berhasil disimpan'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error', message: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Untuk test di editor
function testAppend() {
  const sheet = ensureSheet();
  sheet.appendRow([new Date(), 'Test Tamu', 'Selamat ya Irma & Syihab!', 'Hadir', '2']);
}
