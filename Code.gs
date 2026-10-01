const SHEET_NAME = 'Nilai';
const CONFIG_SHEET_NAME = 'Config';

function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Apps Script harus dibuat dari Google Sheet tujuan melalui Extensions > Apps Script.');
  }
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function getConfigSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Apps Script harus dibuat dari Google Sheet tujuan melalui Extensions > Apps Script.');
  }
  const sheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME) || spreadsheet.insertSheet(CONFIG_SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(['Kunci', 'Nilai']);
  return sheet;
}

function pindahkanAbsensiLama() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const nilaiSheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!nilaiSheet || nilaiSheet.getLastRow() < 2) return;

  const absensiSheet = spreadsheet.getSheetByName('Absensi') || spreadsheet.insertSheet('Absensi');
  if (absensiSheet.getLastRow() === 0) absensiSheet.appendRow(['Waktu', 'Nama Lengkap', 'Kelas']);

  const rows = nilaiSheet.getDataRange().getValues();
  const headers = rows[0];
  const waktuIndex = headers.indexOf('Waktu');
  const namaIndex = headers.indexOf('Nama Siswa');
  const kelasIndex = headers.indexOf('Kelas');
  const mapelIndex = headers.indexOf('Mata Pelajaran');
  const skorIndex = headers.indexOf('Skor');
  const totalIndex = headers.indexOf('Total Soal');
  if ([waktuIndex, namaIndex, kelasIndex, mapelIndex, skorIndex, totalIndex].some(index => index < 0)) return;

  const legacyRows = [];
  for (let rowIndex = rows.length - 1; rowIndex >= 1; rowIndex -= 1) {
    const row = rows[rowIndex];
    const isLegacyAttendance = row[mapelIndex] === '-' && Number(row[skorIndex]) === 0 && Number(row[totalIndex]) === 0;
    if (!isLegacyAttendance) continue;
    absensiSheet.appendRow([row[waktuIndex], row[namaIndex], row[kelasIndex]]);
    legacyRows.push(rowIndex + 1);
  }

  legacyRows.forEach(rowNumber => nilaiSheet.deleteRow(rowNumber));
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'config') {
    const sheet = getConfigSheet();
    const values = sheet.getDataRange().getValues();
    const data = {};
    values.slice(1).forEach(row => {
      if (!row[0]) return;
      try {
        data[row[0]] = JSON.parse(row[1]);
      } catch (error) {
        data[row[0]] = row[1];
      }
    });
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, data }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService
    .createTextOutput('Web App aktif. Gunakan POST untuk menyimpan nilai.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Data POST tidak ditemukan.');
  }
  const data = JSON.parse(e.postData.contents);

  if (data.action === 'saveConfig') {
    const sheet = getConfigSheet();
    const rows = sheet.getDataRange().getValues();
    const rowIndex = rows.findIndex(row => row[0] === data.key);
    const value = JSON.stringify(data.value);
    if (rowIndex >= 0) sheet.getRange(rowIndex + 1, 2).setValue(value);
    else sheet.appendRow([data.key, value]);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (data.action === 'saveAttendance') {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('Apps Script harus dibuat dari Google Sheet tujuan melalui Extensions > Apps Script.');
    }
    const sheet = spreadsheet.getSheetByName('Absensi') || spreadsheet.insertSheet('Absensi');
    if (sheet.getLastRow() === 0) sheet.appendRow(['Waktu', 'Nama Lengkap', 'Kelas']);

    const nama = String(data.nama || '').trim();
    const kelas = String(data.kelas || '').trim();
    if (!nama || !['VII', 'VIII', 'IX'].includes(kelas)) {
      throw new Error('Nama dan kelas absensi tidak valid.');
    }

    sheet.appendRow([new Date(data.waktu || Date.now()), nama, kelas]);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (data.action !== 'saveQuiz') {
    throw new Error('Aksi penyimpanan tidak dikenali.');
  }

  const sheet = getSheet();

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Waktu', 'Nama Siswa', 'Kelas', 'Mata Pelajaran', 'Skor', 'Total Soal']);
  } else {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers.includes('Kelas')) {
      sheet.insertColumnAfter(2);
      sheet.getRange(1, 3).setValue('Kelas');
    }
  }

  sheet.appendRow([
    new Date(data.waktu || Date.now()),
    data.nama || 'Tanpa nama',
    data.kelas || '-',
    data.mapel || '-',
    Number(data.skor) || 0,
    Number(data.totalSoal) || 0
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
