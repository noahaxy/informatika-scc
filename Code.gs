// ====== ISI DUA BARIS INI ======
const SPREADSHEET_ID = 'GANTI_DENGAN_ID_SPREADSHEET';
const FOLDER_ID      = 'GANTI_DENGAN_ID_FOLDER_DRIVE_UNTUK_PDF';
// ================================

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const now = new Date();
    const out = { ok: true };

    if (d.tipe === 'login') {
      tulis(ss, 'Login', ['Waktu', 'Nama', 'Kelas'],
        [now, d.nama, d.kelas]);

    } else if (d.tipe === 'pretest' || d.tipe === 'posttest') {
      const jwb = Object.keys(d.jawaban).map(k => k + ':' + d.jawaban[k]).join(', ');
      tulis(ss, d.tipe === 'pretest' ? 'Pretest' : 'Posttest',
        ['Waktu', 'Nama', 'Kelas', 'Skor', 'Benar', 'Jumlah Soal', 'Jawaban'],
        [now, d.nama, d.kelas, d.skor, d.benar, d.total, jwb]);

    } else if (d.tipe === 'lkpd') {
      const blob = Utilities.newBlob(Utilities.base64Decode(d.pdfBase64), 'application/pdf', d.namaFile);
      const file = DriveApp.getFolderById(FOLDER_ID).createFile(blob);
      out.url = file.getUrl();
      const ringkas = d.data.formula.map(f => f[0] + ': ' + f[1] + ' => ' + f[2]).join(' | ');
      tulis(ss, 'LKPD', ['Waktu', 'Nama', 'Kelas', 'Link PDF', 'Isian Formula'],
        [now, d.nama, d.kelas, out.url, ringkas]);

    } else {
      throw new Error('Tipe data tidak dikenal: ' + d.tipe);
    }
    return json(out);
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return json({ ok: true, pesan: 'Server LMS aktif' }); }

function tulis(ss, namaSheet, header, baris) {
  let sh = ss.getSheetByName(namaSheet) || ss.insertSheet(namaSheet);
  if (sh.getLastRow() === 0) { sh.appendRow(header); sh.setFrozenRows(1); }
  sh.appendRow(baris.map(bersih));
}

// Cegah teks siswa yang diawali = + - @ dibaca sebagai rumus
function bersih(v) {
  return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v;
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Jalankan SEKALI dari editor untuk memberi izin akses Sheets & Drive
function tesIzin() {
  SpreadsheetApp.openById(SPREADSHEET_ID).getName();
  DriveApp.getFolderById(FOLDER_ID).getName();
  Logger.log('Izin OK');
}
