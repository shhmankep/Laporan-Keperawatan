/**
 * Backend Google Drive untuk PWA Laporan Harian Kepala Ruangan.
 * Pasang di Google Sheet: Extensions > Apps Script. Deploy sebagai Web app
 * (Execute as: Me, Who has access: Anyone). Setelah mengubah kode: Deploy > Manage deployments > pensil > New version.
 *
 * Setiap jenis unit (Rawat Inap, IGD, HD, OK, Neonatus, VK, Poliklinik) punya tab "Rekap ..." sendiri.
 * Susunan kolom dikirim oleh aplikasi, jadi tab dibuat otomatis sesuai format laporan.
 */
const TOKEN = '';   // dikosongkan = sinkron otomatis tanpa kode. Isi kode bila ingin dikunci.
const FOLDER_NAME = 'Laporan Harian Kepala Ruangan';
const HEAD = 2;     // 2 baris judul: kelompok (I, II, III, ...) dan nama kolom

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function rekap_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(s => s.getName().indexOf('Rekap') === 0);
}

// cols: [[kelompok, nama kolom, nilai], ...] dari aplikasi
function sheetFor_(name, cols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const labels = ['Kunci'].concat(cols.map(c => c[1]), ['JSON']);
  let s = ss.getSheetByName(name);
  if (s) {
    const hdr = s.getRange(2, 1, 1, Math.max(s.getLastColumn(), 1)).getValues()[0];
    if (hdr.length === labels.length && hdr.every((h, i) => h === labels[i])) return s;
    // Susunan kolom berubah: lembar lama disimpan sebagai cadangan, lembar baru dibuat
    s.setName(name + ' (lama ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd-HHmm') + ')');
  }
  return build_(ss, name, cols);
}

function build_(ss, name, cols) {
  const c = [[cols[0][0], 'Kunci']].concat(cols.map(x => [x[0], x[1]]), [['Sistem', 'JSON']]), n = c.length;
  const s = ss.insertSheet(name);
  if (s.getMaxColumns() < n) s.insertColumnsAfter(s.getMaxColumns(), n - s.getMaxColumns());
  const rows = s.getMaxRows();
  s.getRange(1, 1, rows, 2).setNumberFormat('@');            // kunci & tanggal tetap teks
  let i = 0, gi = 0;
  while (i < n) {
    let j = i;
    while (j + 1 < n && c[j + 1][0] === c[i][0]) j++;
    const col = gi % 2 ? ['#16b364', '#d9f4e5'] : ['#1565d8', '#dbe8fb'];
    const top = s.getRange(1, i + 1, 1, j - i + 1);
    if (j > i) top.merge();
    top.setValue(c[i][0]).setBackground(col[0]).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    s.getRange(2, i + 1, 1, j - i + 1).setBackground(col[1]);
    i = j + 1; gi++;
  }
  s.getRange(2, 1, 1, n).setValues([c.map(x => x[1])])
    .setFontWeight('bold').setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.setRowHeight(2, 54);
  s.getRange(HEAD + 1, 1, rows - HEAD, n).setWrap(true).setVerticalAlignment('top');
  s.setColumnWidths(1, n, 115);
  c.forEach((x, k) => {
    if (/^(II\.|IV\.|V\.)/.test(x[0]) || /^Keterangan/.test(x[1])) s.setColumnWidth(k + 1, 300);
    else if (x[1] === 'Nama Kepala Ruangan') s.setColumnWidth(k + 1, 170);
    else if (x[1] === 'Ruang Rawat Inap' || x[1] === 'Unit / Ruang') s.setColumnWidth(k + 1, 140);
  });
  s.setFrozenRows(HEAD);
  s.setFrozenColumns(4);
  s.hideColumns(1);        // Kunci (internal)
  s.hideColumns(n);        // JSON (dipakai dashboard)
  return s;
}

function doGet(e) {
  try { return get_(e); } catch (err) { return out_({ ok: false, error: String(err) }); }
}

function get_(e) {
  const p = e.parameter || {};
  if (TOKEN && p.token !== TOKEN) return out_({ ok: false, error: 'Kode rahasia salah' });
  const from = p.from || '', found = {};
  rekap_().forEach(s => {
    const n = s.getLastRow() - HEAD, lc = s.getLastColumn();
    if (n < 1 || s.getRange(2, lc).getValue() !== 'JSON') return;
    s.getRange(HEAD + 1, lc, n, 1).getValues().forEach(row => {
      try {
        const r = JSON.parse(row[0]);
        if (String(r.tanggal) < from) return;
        const k = r.tanggal + '|' + r.ruang;
        if (!found[k] || (r.ts || 0) >= (found[k].ts || 0)) found[k] = r;
      } catch (x) { /* abaikan baris rusak */ }
    });
  });
  return out_({ ok: true, reports: Object.keys(found).map(k => found[k]) });
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
    const r = b.report, cols = r.cols, teks = r.teks;
    delete r.cols; delete r.teks;
    const key = r.tanggal + '|' + r.ruang;
    const s = sheetFor_(r.sheet || 'Rekap Laporan', cols), n = s.getLastRow() - HEAD;
    const keys = n > 0 ? s.getRange(HEAD + 1, 1, n, 1).getValues().flat() : [];
    const row = [key].concat(cols.map(c => (c[2] === undefined || c[2] === null) ? '' : c[2]), [JSON.stringify(r)]);
    const i = keys.indexOf(key);
    if (i >= 0) s.getRange(i + HEAD + 1, 1, 1, row.length).setValues([row]);
    else s.getRange(Math.max(s.getLastRow(), HEAD) + 1, 1, 1, row.length).setValues([row]);
    try { docFor_(r, teks); } catch (err) { /* Sheet tetap tersimpan walau Doc gagal */ }
    return out_({ ok: true });
  } finally { lock.releaseLock(); }
}

// Hapus laporan: baris di Sheet dihapus, Google Doc dipindah ke tempat sampah Drive
function del_(key) {
  rekap_().forEach(s => {
    const n = s.getLastRow() - HEAD;
    if (n < 1) return;
    const keys = s.getRange(HEAD + 1, 1, n, 1).getValues().flat();
    for (let i = keys.length - 1; i >= 0; i--) if (keys[i] === key) s.deleteRow(i + HEAD + 1);
  });
  try {
    const p = key.split('|'), it = folder_().getFilesByName(p[0] + ' - ' + p[1]);
    while (it.hasNext()) it.next().setTrashed(true);
  } catch (err) { /* abaikan */ }
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
