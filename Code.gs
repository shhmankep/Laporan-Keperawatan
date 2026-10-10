/**
 * Backend Google Drive untuk PWA Laporan Harian Kepala Ruangan.
 * Pasang di Google Sheet: Extensions > Apps Script. Deploy sebagai Web app
 * (Execute as: Me, Who has access: Anyone). Setelah mengubah kode: Deploy > Manage deployments > pensil > New version.
 */
const TOKEN = '';   // dikosongkan = sinkron otomatis tanpa kode. Isi kode bila ingin dikunci.
const FOLDER_NAME = 'Laporan Harian Kepala Ruangan';
const SHEET_NAME = 'Rekap Laporan';
const HEAD = 2;     // 2 baris judul: kelompok (I, II, III, ...) dan nama kolom
const SH = [['pagi', 'Pagi'], ['sore', 'Sore'], ['malam', 'Malam']];

function g_(r, path) {
  let o = r;
  for (const k of path.split('.')) { if (o == null) return ''; o = o[k]; }
  return o == null ? '' : o;
}

// Satu kolom untuk setiap isian di formulir Word
function cols_() {
  const c = [], add = (grp, label, f) => c.push({ grp: grp, label: label, f: f });
  add('Identitas', 'Kunci', r => r.tanggal + '|' + r.ruang);
  add('Identitas', 'Hari / Tanggal', r => r.tanggal);
  add('Identitas', 'Ruang Rawat Inap', r => r.ruang);
  add('Identitas', 'Nama Kepala Ruangan', r => r.kepala);
  add('Identitas', 'Metode Asuhan Keperawatan', r => r.metode);
  SH.forEach(([k, l]) => [
    ['masuk', 'Pasien Masuk Baru (orang)'], ['pulang', 'Pasien Pulang BPL/APS/Meninggal (orang)'],
    ['rujuk', 'Pasien Rujuk (orang)'], ['total', 'Total Pasien Akhir Shift (orang)'],
    ['tt', 'Kapasitas TT'], ['bor', 'BOR (%)']
  ].forEach(([f, t]) => add('I.1 Sensus Mutasi Pasien - ' + l, t, r => g_(r, k + '.' + f))));
  [['minimal', 'Minimal Care'], ['partial', 'Partial Care'], ['total', 'Total Care']]
    .forEach(([f, t]) => add('I.2 Ketergantungan Pasien (orang)', t, r => g_(r, 'kg.' + f)));
  SH.forEach(([k, l]) => [
    ['ketua', 'PPJA / Ketua Tim (orang)'], ['pelaksana', 'Perawat Pelaksana (orang)'],
    ['status', 'Status Ketenagaan'], ['ket', 'Keterangan']
  ].forEach(([f, t]) => add('I.3 Kecukupan Tenaga - Dinas ' + l, t, r => g_(r, 'tn.' + k + '.' + f))));
  [['planning', 'Planning (Perencanaan)'], ['organizing', 'Organizing (Pengorganisasian)'],
   ['directing', 'Directing (Pengarahan)'], ['controlling', 'Controlling (Pengawasan dan Audit)'],
   ['evaluating', 'Evaluating (Evaluasi)']]
    .forEach(([f, t]) => add('II. Fungsi Manajemen - Kegiatan', t, r => g_(r, 'mj.' + f)));
  [['s1', 'SKP 1 Identifikasi Pasien'], ['s2', 'SKP 2 Komunikasi Efektif'], ['s3', 'SKP 3 High-Alert Medications'],
   ['s5', 'SKP 5 Pencegahan Infeksi/HAIs'], ['s6', 'SKP 6 Risiko Jatuh']]
    .forEach(([f, t]) => add('III. Sasaran Keselamatan Pasien (SKP)', t, r => g_(r, 'skp.' + f)));
  add('IV. Masalah Khusus & RTL', 'Masalah Klinis/Manajerial', r => r.masalah);
  add('IV. Masalah Khusus & RTL', 'Rencana Tindak Lanjut (RTL)', r => r.rtl);
  add('V. Fasilitas Alkes dan Non Alkes', 'Alkes - Kondisi', r => g_(r, 'alkes.kondisi'));
  add('V. Fasilitas Alkes dan Non Alkes', 'Alkes - Masalah', r => g_(r, 'alkes.masalah'));
  add('V. Fasilitas Alkes dan Non Alkes', 'Non Alkes - Kondisi', r => g_(r, 'nonalkes.kondisi'));
  add('V. Fasilitas Alkes dan Non Alkes', 'Non Alkes - Masalah', r => g_(r, 'nonalkes.masalah'));
  add('Sistem', 'Diperbarui', () => new Date());
  add('Sistem', 'JSON', r => JSON.stringify(r));
  return c;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_NAME);
  if (s) return s;
  s = ss.insertSheet(SHEET_NAME);
  const c = cols_(), n = c.length, rows = s.getMaxRows();
  s.getRange(1, 1, rows, 2).setNumberFormat('@');            // kunci & tanggal tetap teks
  let i = 0, gi = 0;
  while (i < n) {
    let j = i;
    while (j + 1 < n && c[j + 1].grp === c[i].grp) j++;
    const col = gi % 2 ? ['#16b364', '#d9f4e5'] : ['#1565d8', '#dbe8fb'];
    const top = s.getRange(1, i + 1, 1, j - i + 1);
    if (j > i) top.merge();
    top.setValue(c[i].grp).setBackground(col[0]).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    s.getRange(2, i + 1, 1, j - i + 1).setBackground(col[1]);
    i = j + 1; gi++;
  }
  s.getRange(2, 1, 1, n).setValues([c.map(x => x.label)])
    .setFontWeight('bold').setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.setRowHeight(2, 54);
  s.getRange(HEAD + 1, 1, rows - HEAD, n).setWrap(true).setVerticalAlignment('top');
  c.forEach((x, k) => {
    const wide = /^(II\.|IV\.|V\.)/.test(x.grp) || x.label === 'Keterangan';
    s.setColumnWidth(k + 1, wide ? 300 : x.label === 'Nama Kepala Ruangan' ? 170 : x.label === 'Ruang Rawat Inap' ? 140 : 115);
  });
  s.setFrozenRows(HEAD);
  s.setFrozenColumns(4);
  s.hideColumns(1);        // Kunci (internal)
  s.hideColumns(n);        // JSON (dipakai dashboard)
  return s;
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function doGet(e) {
  try { return get_(e); } catch (err) { return out_({ ok: false, error: String(err) }); }
}

function get_(e) {
  const p = e.parameter || {};
  if (TOKEN && p.token !== TOKEN) return out_({ ok: false, error: 'Kode rahasia salah' });
  const s = sheet_(), n = s.getLastRow() - HEAD, nc = cols_().length;
  const rows = n > 0 ? s.getRange(HEAD + 1, 1, n, nc).getValues() : [];
  const from = p.from || '';
  const reports = rows
    .filter(r => String(r[1]) >= from)
    .map(r => { try { return JSON.parse(r[nc - 1]); } catch (x) { return null; } })
    .filter(Boolean);
  return out_({ ok: true, reports: reports });
}

function doPost(e) {
  try { return post_(e); } catch (err) { return out_({ ok: false, error: String(err) }); }
}

function post_(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, error: 'Data rusak' }); }
  if (TOKEN && b.token !== TOKEN) return out_({ ok: false, error: 'Kode rahasia salah' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (b.action === 'delete') { del_(String(b.key || '')); return out_({ ok: true }); }
    const r = b.report;
    const teks = r.teks; delete r.teks;
    const key = r.tanggal + '|' + r.ruang;
    const s = sheet_(), n = s.getLastRow() - HEAD;
    const keys = n > 0 ? s.getRange(HEAD + 1, 1, n, 1).getValues().flat() : [];
    const row = cols_().map(x => x.f(r));
    const i = keys.indexOf(key);
    if (i >= 0) s.getRange(i + HEAD + 1, 1, 1, row.length).setValues([row]);
    else s.getRange(Math.max(s.getLastRow(), HEAD) + 1, 1, 1, row.length).setValues([row]);
    try { docFor_(r, teks); } catch (err) { /* Sheet tetap tersimpan walau Doc gagal */ }
    return out_({ ok: true });
  } finally { lock.releaseLock(); }
}

function docFor_(r, teks) {
  const f = folder_(), name = r.tanggal + ' - ' + r.ruang, it = f.getFilesByName(name);
  let d;
  if (it.hasNext()) d = DocumentApp.openById(it.next().getId());
  else { d = DocumentApp.create(name); DriveApp.getFileById(d.getId()).moveTo(f); }
  const body = d.getBody().clear();
  String(teks || '').split('\n').forEach(l => {
    if (l.indexOf('# ') === 0) body.appendParagraph(l.slice(2)).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    else body.appendParagraph(l);
  });
  d.saveAndClose();
}

// Hapus laporan: baris di Sheet dihapus, Google Doc dipindah ke tempat sampah Drive
function del_(key) {
  const s = sheet_(), n = s.getLastRow() - HEAD;
  if (n > 0) {
    const i = s.getRange(HEAD + 1, 1, n, 1).getValues().flat().indexOf(key);
    if (i >= 0) s.deleteRow(i + HEAD + 1);
  }
  try {
    const p = key.split('|'), it = folder_().getFilesByName(p[0] + ' - ' + p[1]);
    while (it.hasNext()) it.next().setTrashed(true);
  } catch (err) { /* abaikan */ }
}
