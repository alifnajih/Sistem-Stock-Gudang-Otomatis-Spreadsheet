/**
 * SISTEM MANAJEMEN PERSEDIAAN & STOCK OPNAME GUDANG FURNITURE
 * PT. EZRIA FURNISHING BALI - Google Apps Script v2.0 (untuk Google Sheets)
 * Pasang: Extensions > Apps Script > tempel file ini sebagai Code.gs, dan Foto.html sebagai file HTML bernama "Foto".
 * Konstanta baris/kolom HARUS sama dengan workbook.
 */

// ---------- Nama sheet ----------
const SH = {
  DASH: 'DASHBOARD', MASTER: 'MASTER_BARANG', RUANG: 'MASTER_RUANGAN', SALDO: 'SALDO_AWAL', TRX: 'TRANSAKSI_STOK',
  KES: 'STOCK_KESELURUHAN', FORM: 'FORM_STOCK_OPNAME', RIW: 'RIWAYAT_OPNAME', CETAK: 'CETAK_STOCK_OPNAME', SET: 'PENGATURAN'
};
const PREFIX_STOK = 'STOCK_RUANG_';
const FOLDER_FOTO = 'Foto_Produk';
const FOLDER_CETAK = 'Cetak_Stock_Opname';

// ---------- Baris ----------
const R1 = 6, M_LAST = 205, RM_LAST = 14, S_LAST = 1005, T_LAST = 3005, H_LAST = 6005;
const F_R1 = 14, F_LAST = 213;                 // FORM_STOCK_OPNAME
const C_HDR = 15, C_R1 = 16, C_LAST = 215;      // CETAK_STOCK_OPNAME
const C_PRINT_FIRST = 6, C_PRINT_LAST = 225;
const H_MASTER = 90, H_STOK = 60, H_FORM = 55, H_LIST = 45, H_DEFAULT = 21;

// KHUSUS TAMPILAN CETAK STOCK OPNAME
const C_PRINT_PHOTO_WIDTH = 110;
const C_PRINT_ROW_HEIGHT = 90;
const TRX_COLS = [2, 4, 5, 8, 9, 10];          // kolom TRANSAKSI_STOK yang memicu pemrosesan (B, D, E, H, I, J)
const KAPASITAS_PERINGATAN = 0.9;               // peringatan bila kapasitas baris terpakai >= 90%
const KUNCI_FOTO_PUBLIK = 'FOTO_PUBLIK';
const KOL_SALDO1 = 13;                          // kolom M di MASTER_BARANG: stok awal ruangan pertama
const JML_RUANG = RM_LAST - R1 + 1;             // 9 ruangan

// ======================================================================
// UTILITAS
// ======================================================================
function sh_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error('Sheet tidak ditemukan: ' + name);
  return s;
}
function txt_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object' && !(v instanceof Date)) return '';   // CellImage dsb.
  return String(v).trim();
}
function num_(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
function ui_() { return SpreadsheetApp.getUi(); }
function info_(msg, title) { ui_().alert(title || 'Informasi', msg, ui_().ButtonSet.OK); }
function ask_(msg, title) { return ui_().alert(title || 'Konfirmasi', msg, ui_().ButtonSet.YES_NO) === ui_().Button.YES; }
function prompt_(msg, title, def) {
  const r = ui_().prompt(title || 'Input', msg, ui_().ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui_().Button.OK) return null;
  const t = r.getResponseText().trim();
  return t === '' && def ? def : t;
}
function folder_(nama) {
  const it = DriveApp.getFoldersByName(nama);
  return it.hasNext() ? it.next() : DriveApp.createFolder(nama);
}
function jenis_() {
  const v = sh_(SH.SET).getRange('G5:G9').getValues().map(function (r) { return txt_(r[0]); });
  return { masuk: v[0], keluar: v[1], transfer: v[2], adjp: v[3], adjm: v[4] };
}
function lastNonEmpty_(sheet, col, r1, r2) {
  const v = sheet.getRange(r1, col, r2 - r1 + 1, 1).getValues();
  let n = 0;
  for (let i = 0; i < v.length; i++) if (txt_(v[i][0]) !== '') n = i + 1;
  return n === 0 ? r1 - 1 : r1 + n - 1;
}
function barisSku_(sku) {
  const v = sh_(SH.MASTER).getRange(R1, 2, M_LAST - R1 + 1, 1).getValues();
  for (let i = 0; i < v.length; i++) if (txt_(v[i][0]) === sku) return R1 + i;
  return 0;
}
function roomSheets_() {
  return ss_().getSheets().filter(function (s) { return s.getName().indexOf(PREFIX_STOK) === 0; });
}

// ======================================================================
// MENU
// ======================================================================
function onOpen() {
  const u = SpreadsheetApp.getUi();
  u.createMenu('EZRIA GUDANG')
    .addItem('Pasang / Perbarui Sistem (jalankan sekali)', 'PasangSistem')
    .addSeparator()
    .addSubMenu(u.createMenu('Foto Produk')
      .addItem('Tambah / Ganti Foto (SKU di baris aktif)', 'TambahFotoProduk')
      .addItem('Hapus Foto', 'HapusFotoProduk')
      .addItem('Lihat Foto', 'TampilkanFotoProduk')
      .addItem('Sinkronisasi / Periksa Foto', 'SinkronisasiFotoProduk')
      .addItem('Cadangan: salin foto langsung ke sheet lain', 'SalinFotoLangsung'))
    .addItem('Tambah SKU Baru', 'TambahSKU')
    .addItem('Hapus SKU (baris aktif)', 'HapusSKU')
    .addSubMenu(u.createMenu('Transaksi')
      .addItem('Beri Nomor Transaksi', 'BuatNomorTransaksi')
      .addItem('Validasi Transaksi', 'ValidasiTransaksi')
      .addItem('Koreksi Transaksi', 'KoreksiTransaksi'))
    .addSubMenu(u.createMenu('Stock Opname')
      .addItem('Muat Data Ruangan', 'MuatDataRuangan')
      .addItem('Simpan Stock Opname', 'SimpanStockOpname')
      .addItem('Reset Form', 'ResetFormOpname')
      .addItem('Cetak hasil opname ini (PDF)', 'CetakDariForm'))
    .addSubMenu(u.createMenu('Cetak')
      .addItem('Cetak ruangan terpilih (PDF)', 'CetakRuanganTerpilih')
      .addItem('Cetak SEMUA ruangan (PDF)', 'CetakSemuaRuangan'))
    .addSubMenu(u.createMenu('Lainnya')
      .addItem('Sinkronkan Saldo Awal dari Master', 'SinkronSaldoAwal')
      .addItem('Aplikasi HP: tampilkan tautan', 'TautanAplikasi')
      .addItem('Aplikasi HP: atur PIN', 'AturPinWeb')
      .addItem('Aplikasi HP: tampilkan/sembunyikan log', 'TampilkanLog')
      .addItem('Periksa Kesehatan Sistem', 'PeriksaKesehatan')
      .addItem('Pasang ulang pemicu otomatis', 'PasangPemicu')
      .addItem('Lindungi sheet rumus (pemilik)', 'LindungiSistem')
      .addItem('Cadangan foto publik: aktif/nonaktif', 'AturFotoPublik')
      .addSeparator()
      .addItem('Isi foto contoh (SKU001 & SKU002)', 'IsiFotoDemo')
      .addItem('Tentang', 'Tentang'))
    .addToUi();
}

function Tentang() {
  info_('Sistem Persediaan & Stock Opname - PT. EZRIA FURNISHING BALI\nVersi Google Sheets 2.0\n\nFormula menghitung stok otomatis; menu EZRIA GUDANG mengatur foto, nomor transaksi, validasi, opname, dan cetak PDF.');
}

// ======================================================================
// INSTALASI (jalankan sekali)
// ======================================================================
function PasangSistem() {
  wajibMenu_();
  const ss = ss_();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  try { const lg = ss.getSheetByName('LOG_APLIKASI'); if (lg) lg.hideSheet(); } catch (e) { /* abaikan */ }
  const dash = sh_(SH.DASH);
  const nav = [['B5', SH.MASTER, 'Master Barang'], ['D5', SH.RUANG, 'Master Ruangan'], ['F5', SH.SALDO, 'Saldo Awal'], ['H5', SH.TRX, 'Transaksi Stok'],
               ['B6', SH.KES, 'Stok Keseluruhan'], ['D6', SH.FORM, 'Form Stock Opname'], ['F6', SH.RIW, 'Riwayat Opname'], ['H6', SH.SET, 'Pengaturan']];
  nav.forEach(function (n) {
    const g = sh_(n[1]).getSheetId();
    dash.getRange(n[0]).setFormula('=HYPERLINK("#gid=' + g + '","' + n[2] + '")');
  });
  folder_(FOLDER_FOTO); folder_(FOLDER_CETAK);
  AturTinggiBaris();
  let catatan = '';
  try { pasangTrigger_(); catatan += '- Pemicu otomatis (nomor transaksi, cek stok, muat form) dipasang\n'; }
  catch (e) { catatan += '- GAGAL memasang pemicu otomatis: ' + e.message + '\n  (jalankan menu Lainnya > Pasang ulang pemicu otomatis)\n'; }
  try { const n = LindungiSistem_(); catatan += '- ' + n + ' perlindungan sheet/kolom rumus dipasang\n'; }
  catch (e) { catatan += '- Perlindungan sheet dilewati: ' + e.message + '\n'; }
  try {
    PasangKolomSaldo_();
    const imp = imporSaldoKeMaster_();
    const h = dalamKunci_(sinkronSaldo_);
    catatan += '- Kolom stok awal per ruangan dipasang di MASTER_BARANG (M..V)' + (imp ? ', ' + imp + ' saldo lama dipindahkan' : '') + '; SALDO_AWAL terisi otomatis (' + h.out.length + ' baris)\n';
  } catch (e) { catatan += '- GAGAL memasang kolom saldo di master: ' + e.message + '\n'; }
  ss.toast('Sistem terpasang. Menu EZRIA GUDANG siap dipakai.', 'EZRIA', 6);
  info_('Pemasangan selesai:\n- Tautan navigasi dashboard dipasang\n- Folder Drive "' + FOLDER_FOTO + '" dan "' + FOLDER_CETAK + '" dibuat\n- Tinggi baris disesuaikan\n' + catatan +
    '\nLangkah berikut: Lainnya > Periksa Kesehatan Sistem, lalu gunakan menu Foto Produk atau Isi foto contoh.', 'Pasang Sistem');
}

// Pemicu "installable" (bukan simple trigger): berjalan dengan izin pemilik, batas 6 menit, dan boleh memakai LockService.
function pasangTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'saatEdit') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('saatEdit').forSpreadsheet(ss_()).onEdit().create();
}
function adaTrigger_() {
  return ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'saatEdit'; });
}
function PasangPemicu() {
  wajibMenu_();
  try { pasangTrigger_(); info_('Pemicu otomatis dipasang ulang.', 'Pemicu'); }
  catch (e) { info_('Gagal: ' + e.message, 'Pemicu'); }
}

// Sheet khusus-rumus: hanya pemilik yang boleh mengubah. Kolom rumus di sheet input: peringatan saja.
function LindungiSistem_() {
  const ss = ss_();
  const me = Session.getEffectiveUser().getEmail();
  ss.getProtections(SpreadsheetApp.ProtectionType.SHEET).concat(ss.getProtections(SpreadsheetApp.ProtectionType.RANGE)).forEach(function (p) {
    if (String(p.getDescription()).indexOf('EZRIA:') === 0) p.remove();
  });
  let n = 0;
  [SH.DASH, SH.KES, 'HELPER', 'PREVIEW_FOTO'].concat(roomSheets_().map(function (s) { return s.getName(); })).forEach(function (nm) {
    const s = ss.getSheetByName(nm);
    if (!s) return;
    const p = s.protect().setDescription('EZRIA: sheet rumus - hanya pemilik');
    p.removeEditors(p.getEditors().filter(function (u) { return u.getEmail() !== me; }));
    if (p.canDomainEdit()) p.setDomainEdit(false);
    n++;
  });
  [[SH.MASTER, [1, 11, KOL_SALDO1 + JML_RUANG], R1, M_LAST], [SH.SALDO, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], R1, S_LAST], [SH.TRX, [1, 6, 7, 11, 14], R1, T_LAST]].forEach(function (k) {
    const s = sh_(k[0]);
    k[1].forEach(function (c) {
      s.getRange(k[2], c, k[3] - k[2] + 1, 1).protect().setDescription('EZRIA: kolom rumus').setWarningOnly(true);
      n++;
    });
  });
  return n;
}
function LindungiSistem() {
  wajibMenu_();
  if (!ask_('Pasang perlindungan: sheet rumus hanya bisa diubah pemilik; kolom rumus di sheet input memberi peringatan bila diedit.\nJalankan sebagai PEMILIK file. Lanjutkan?', 'Lindungi Sheet')) return;
  try { info_(LindungiSistem_() + ' perlindungan dipasang.', 'Lindungi Sheet'); }
  catch (e) { info_('Gagal: ' + e.message, 'Lindungi Sheet'); }
}

function AturFotoPublik() {
  wajibMenu_();
  const p = PropertiesService.getScriptProperties();
  const aktif = p.getProperty(KUNCI_FOTO_PUBLIK) === 'YA';
  if (aktif) {
    if (ask_('Cadangan foto publik sedang AKTIF (thumbnail di Drive memakai tautan "siapa saja yang memiliki link").\nNonaktifkan?', 'Cadangan Foto')) { p.deleteProperty(KUNCI_FOTO_PUBLIK); info_('Dinonaktifkan.', 'Cadangan Foto'); }
  } else if (ask_('Cadangan foto publik NONAKTIF (aman).\nBila diaktifkan, jika Google menolak gambar langsung ke sel, thumbnail 320 px disimpan di Drive dengan tautan yang dapat dilihat siapa saja yang memiliki link.\nAktifkan?', 'Cadangan Foto')) {
    p.setProperty(KUNCI_FOTO_PUBLIK, 'YA'); info_('Diaktifkan.', 'Cadangan Foto');
  }
}

function kesehatanLaporan_() {
  SpreadsheetApp.flush();
  const L = [];
  [[SH.MASTER, 2, R1, M_LAST, 'SKU'], [SH.SALDO, 2, R1, S_LAST, 'Baris saldo awal'], [SH.TRX, 4, R1, T_LAST, 'Transaksi'], [SH.RIW, 1, R1, H_LAST, 'Baris riwayat opname']].forEach(function (c) {
    const kap = c[3] - c[2] + 1, pakai = Math.max(0, lastNonEmpty_(sh_(c[0]), c[1], c[2], c[3]) - c[2] + 1);
    L.push(c[4] + ': ' + pakai + ' / ' + kap + ' (' + Math.round(pakai * 100 / kap) + '%)' + (pakai / kap >= KAPASITAS_PERINGATAN ? '  <-- HAMPIR PENUH' : ''));
  });
  let trig = false;
  try { trig = adaTrigger_(); } catch (e) { trig = false; }
  L.push('Pemicu otomatis: ' + (trig ? 'TERPASANG' : 'BELUM TERPASANG (menu Lainnya > Pasang ulang pemicu otomatis)'));
  const dup = function (arr) { const s = {}, d = []; arr.forEach(function (r) { const k = txt_(r[0]); if (!k) return; if (s[k] && d.indexOf(k) < 0) d.push(k); s[k] = true; }); return d; };
  const lt = lastNonEmpty_(sh_(SH.TRX), 3, R1, T_LAST), lm = lastNonEmpty_(sh_(SH.MASTER), 2, R1, M_LAST);
  const dn = lt >= R1 ? dup(sh_(SH.TRX).getRange(R1, 3, lt - R1 + 1, 1).getValues()) : [];
  const ds = lm >= R1 ? dup(sh_(SH.MASTER).getRange(R1, 2, lm - R1 + 1, 1).getValues()) : [];
  L.push('Nomor transaksi kembar: ' + (dn.length ? dn.slice(0, 10).join(', ') : 'tidak ada'));
  L.push('SKU kembar: ' + (ds.length ? ds.slice(0, 10).join(', ') : 'tidak ada'));
  const J = jenis_(), ev = bangunPeristiwa_(dataStok_(), J, null), neg = [];
  Object.keys(ev).forEach(function (k) { const m = minKronologis_(ev[k]); if (m < 0) neg.push(ketKunci_(k, m)); });
  L.push('Stok pernah negatif (kronologis): ' + (neg.length ? neg.length + ' kasus -> ' + neg.slice(0, 8).join('; ') : 'tidak ada'));
  try {
    const hm = saldoDariMaster_(), sa = dataStok_().saldo.filter(function (r) { return txt_(r[0]) !== '' && txt_(r[3]) !== '' && txt_(r[5]) !== ''; });
    const s1 = hm.out.reduce(function (a, x) { return a + x[2]; }, 0), s2 = sa.reduce(function (a, x) { return a + num_(x[5]); }, 0);
    L.push('Saldo Awal sinkron dengan MASTER_BARANG: ' + (hm.out.length === sa.length && s1 === s2 ? 'YA' : 'TIDAK (jalankan Lainnya > Sinkronkan Saldo Awal)') + (hm.bad.length ? ' | ' + hm.bad.length + ' nilai perlu diperiksa' : ''));
  } catch (e) { L.push('Saldo Awal: tidak dapat diperiksa (' + e.message + ')'); }
  return L;
}
function PeriksaKesehatan() {
  wajibMenu_();
  info_(kesehatanLaporan_().join('\n'), 'Kesehatan Sistem');
}

function AturTinggiBaris() {
  wajibMenu_();
  function atur(nama, col, r1, r2, h) {
    try {
      const s = sh_(nama);
      const last = lastNonEmpty_(s, col, r1, r2);
      const n = last - r1 + 1;
      if (n > 0) s.setRowHeightsForced(r1, n, h);
      const sisa = (r2 - r1 + 1) - Math.max(n, 0);
      if (sisa > 0) s.setRowHeightsForced(r1 + Math.max(n, 0), sisa, H_DEFAULT);
    } catch (e) { /* sheet terlindungi / tidak dapat diubah oleh pengguna ini */ }
  }
  atur(SH.MASTER, 2, R1, M_LAST, H_MASTER);
  atur(SH.KES, 2, R1, M_LAST, H_STOK);
  roomSheets_().forEach(function (s) { atur(s.getName(), 2, R1, M_LAST, H_STOK); });
  atur(SH.SALDO, 2, R1, S_LAST, H_LIST);
  atur(SH.TRX, 5, R1, T_LAST, H_LIST);
  atur(SH.FORM, 3, F_R1, F_LAST, H_FORM);
  atur(SH.RIW, 4, R1, H_LAST, H_LIST);
atur(SH.CETAK, 3, C_R1, C_LAST, C_PRINT_ROW_HEIGHT);

const cetak = sh_(SH.CETAK);
cetak.setColumnWidth(2, C_PRINT_PHOTO_WIDTH);
}

// ======================================================================
// FOTO PRODUK
// ======================================================================
function skuAktif_(judul) {
  judul = judul || 'Foto Produk';
  const s = SpreadsheetApp.getActiveSheet();
  if (s.getName() !== SH.MASTER) { info_('Buka sheet MASTER_BARANG lalu klik salah satu sel pada baris SKU yang dimaksud.', judul); return null; }
  const r = s.getActiveCell().getRow();
  if (r < R1 || r > M_LAST) { info_('Pilih sel pada baris data SKU (baris ' + R1 + ' sampai ' + M_LAST + ').', judul); return null; }
  const sku = txt_(s.getRange(r, 2).getValue());
  if (!sku) { info_('Baris yang dipilih belum memiliki Kode Barang (SKU).', judul); return null; }
  return { row: r, sku: sku };
}

function TambahFotoProduk() {
  const a = skuAktif_();
  if (!a) return;
  const t = HtmlService.createTemplateFromFile('Foto');
  t.sku = a.sku; t.row = a.row;
  ui_().showModalDialog(t.evaluate().setWidth(430).setHeight(470), 'Foto Produk - ' + a.sku);
}

// dipanggil dari Foto.html
// Dipanggil dari dialog foto di Sheet (pengguna sudah masuk). Browser anonim ditolak; HP memakai api('simpanFoto').
function simpanFoto(p) {
  if (!Session.getActiveUser().getEmail()) throw new Error('Hanya dapat dipakai dari Google Sheets.');
  const r = simpanFotoInti_(p);
  CacheService.getScriptCache().remove('foto_' + String(p.sku));
  return r;
}
function simpanFotoInti_(p) {
  const m = sh_(SH.MASTER);
  p.sku = String(p.sku).replace(/^["'\s]+|["'\s]+$/g, '');      // toleran bila nilai tiba dengan tanda kutip
  p.row = Number(String(p.row).replace(/[^0-9]/g, ''));
  if (!p.row || txt_(m.getRange(p.row, 2).getValue()) !== p.sku) throw new Error('Baris SKU berubah. Tutup dialog dan ulangi.');
  const folder = folder_(FOLDER_FOTO);
  const safe = p.sku.replace(/[\\\/:*?"<>|]/g, '_');
  const ext = p.mime === 'image/png' ? 'png' : 'jpg';
  // hapus arsip lama dan thumbnail cadangan lama untuk SKU ini
  const old = folder.getFiles();
  while (old.hasNext()) {
    const f = old.next(), nm = f.getName().replace(/\.[^.]+$/, '');
    if (nm === safe || nm === 'thumb_' + safe) f.setTrashed(true);
  }
  let fileId = '';
  if (p.orig) {
    const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(p.orig), p.mime, safe + '.' + ext));
    fileId = file.getId();
  }
  const b64t = String(p.thumb).split(',')[1];
  const tf = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64t), p.mime, 'thumb_' + safe + '.' + ext));   // thumbnail pribadi (untuk aplikasi HP)
  try {
    const img = SpreadsheetApp.newCellImage().setSourceUrl(p.thumb).setAltTextTitle(p.sku).setAltTextDescription('Foto ' + p.sku).build();
    m.getRange(p.row, 3).setValue(img);
  } catch (e) {
    // cadangan (hanya bila diaktifkan pemilik): thumbnail di Drive dengan tautan publik
    if (PropertiesService.getScriptProperties().getProperty(KUNCI_FOTO_PUBLIK) !== 'YA') {
      throw new Error('Google menolak gambar langsung ke sel (' + e.message + '). Cadangan foto publik NONAKTIF demi keamanan; aktifkan lewat EZRIA GUDANG > Lainnya > Cadangan foto publik, lalu ulangi.');
    }
    tf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const img2 = SpreadsheetApp.newCellImage().setSourceUrl('https://drive.google.com/uc?export=view&id=' + tf.getId()).setAltTextTitle(p.sku).build();
    m.getRange(p.row, 3).setValue(img2);
  }
  m.getRange(p.row, 12).setValue(fileId ? 'https://drive.google.com/open?id=' + fileId : '(file asli tidak diarsipkan)');
  m.setRowHeightsForced(p.row, 1, H_MASTER);
  return 'Foto ' + p.sku + ' tersimpan.';
}

function HapusFotoProduk() {
  const a = skuAktif_();
  if (!a) return;
  if (!ask_('Hapus foto untuk SKU ' + a.sku + '?\nData produk TIDAK dihapus (hanya foto).', 'Hapus Foto')) return;
  const m = sh_(SH.MASTER);
  m.getRange(a.row, 3).clearContent();
  m.getRange(a.row, 12).clearContent();
  info_('Foto SKU ' + a.sku + ' dihapus. (File arsip asli di Drive tidak ikut dihapus.)', 'Hapus Foto');
}

function TampilkanFotoProduk() {
  let sku = '', row = 0;
  const s = SpreadsheetApp.getActiveSheet();
  if (s.getName() === SH.MASTER && s.getActiveCell().getRow() >= R1 && txt_(s.getRange(s.getActiveCell().getRow(), 2).getValue())) {
    row = s.getActiveCell().getRow(); sku = txt_(s.getRange(row, 2).getValue());
  } else {
    sku = prompt_('Masukkan Kode Barang (SKU):', 'Lihat Foto');
    if (!sku) return;
    row = barisSku_(sku);
    if (!row) { info_('SKU tidak ditemukan.', 'Lihat Foto'); return; }
  }
  const m = sh_(SH.MASTER);
  const link = txt_(m.getRange(row, 12).getValue());
  let html = '';
  const mt = link.match(/id=([\w-]+)/);
  if (mt) {
    try {
      const blob = DriveApp.getFileById(mt[1]).getBlob();
      html = '<img style="max-width:100%;max-height:520px" src="data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) + '">';
    } catch (e) { html = ''; }
  }
  if (!html) {
    const v = m.getRange(row, 3).getValue();
    if (v && typeof v === 'object') {
      try { html = '<img style="max-width:100%;max-height:520px" src="' + v.getUrl() + '">'; } catch (e) { html = ''; }
    }
  }
  if (!html) { info_('SKU ' + sku + ' belum memiliki foto.', 'Lihat Foto'); return; }
  const out = HtmlService.createHtmlOutput('<div style="font-family:Arial;text-align:center"><h3>' + sku + ' - ' + txt_(m.getRange(row, 4).getValue()) + '</h3>' + html + '</div>').setWidth(620).setHeight(620);
  ui_().showModalDialog(out, 'Foto Produk');
}

function SinkronisasiFotoProduk() {
  const m = sh_(SH.MASTER);
  AturTinggiBaris();
  const v = m.getRange(R1, 2, M_LAST - R1 + 1, 3).getValues();
  let ada = 0; const tanpa = [];
  for (let i = 0; i < v.length; i++) {
    const sku = txt_(v[i][0]);
    if (!sku) continue;
    if (v[i][2] !== '' && v[i][2] !== null) ada++; else tanpa.push(sku);
  }
  info_('Pemeriksaan foto selesai.\nSKU dengan foto: ' + ada + '\nSKU tanpa foto: ' + tanpa.length + (tanpa.length ? '\n' + tanpa.slice(0, 30).join(', ') : '') +
    '\n\nFoto di sheet lain mengikuti MASTER_BARANG lewat rumus (aman terhadap pengurutan/filter). Jika foto tidak tampil di sheet lain, jalankan menu "Cadangan: salin foto langsung".', 'Sinkronisasi Foto');
}

// Cadangan: menuliskan gambar langsung ke sel foto sheet turunan (jika rumus tidak menampilkan gambar)
function SalinFotoLangsung() {
  wajibMenu_();
  const m = sh_(SH.MASTER);
  const sv = m.getRange(R1, 2, M_LAST - R1 + 1, 2).getValues();
  const map = {};
  sv.forEach(function (r) { const k = txt_(r[0]); if (k && r[1] && typeof r[1] === 'object') map[k] = r[1]; });
  const targets = [];
  roomSheets_().forEach(function (s) { targets.push([s, 2, 3, R1, M_LAST]); });
  targets.push([sh_(SH.KES), 2, 3, R1, M_LAST], [sh_(SH.SALDO), 2, 3, R1, S_LAST], [sh_(SH.TRX), 5, 6, R1, T_LAST],
               [sh_(SH.FORM), 3, 2, F_R1, F_LAST], [sh_(SH.RIW), 4, 5, R1, H_LAST], [sh_(SH.CETAK), 3, 2, C_R1, C_LAST]);
  let n = 0, gagal = 0;
  targets.forEach(function (t) {
    const s = t[0];
    try {
      const skus = s.getRange(t[3], t[1], t[4] - t[3] + 1, 1).getValues();
      for (let i = 0; i < skus.length; i++) {
        const k = txt_(skus[i][0]);
        if (k && map[k]) { s.getRange(t[3] + i, t[2]).setValue(map[k]); n++; }
      }
    } catch (e) { gagal++; }
  });
  info_(n + ' foto disalin langsung. Catatan: hasil ini bernilai tetap; setelah data diurutkan atau berubah jalankan lagi menu ini.' +
    (gagal ? '\n' + gagal + ' sheet dilewati (terlindungi). Jalankan sebagai pemilik file.' : ''), 'Salin Foto');
}

function IsiFotoDemo() {
  wajibMenu_();
  const demo = {
    'SKU001': 'iVBORw0KGgoAAAANSUhEUgAAAPAAAADwCAIAAACxN37FAAAEF0lEQVR42u3cXUscVxzA4f9sxhSa6s5GW402tr0KCkssDbRQSK9aCr3wi+Tj5BsIRXojUtpK0tBqSV8CscHGKqQtItE1xsS3xAuxO70QFhPpKiiLa57nSuaMg875cTiO7CRb68sBp0XBLUDQIGgQNAgaQYOgQdAgaDh60EPDI2M3x8e+H//2xg9//TO3e/DLr0ZrJ9S+Hhoe+fGnX2vHJ27fGRoeqT+0tr4x+s2NPM93j3/93a3VtQ3zwRGldcbOFAqff/pJROzs7Nwa/zlN03d73/6/M9c3NvM8T5IkIjafPTtTKNQfyoptWbE4N//wvXcuPlyotLaeK2Vt5oNGbDnSNL3yfvnP2Qd1zmk/X1p5shoRT1fXSll2mKGBct/U/Zk8z6emZwfK/SaDxu2hS1lxc/N5nRN6ujsXKksRsbC41NPdeZihYrGtlGW/3Jl849zrWdHyTAODzvO8UEjqnNB9obNSWY6IyqPH3V1vHXLocrnvwd9zly3PNDjox0+eZlkxIiJJdv+Sq+Z5sifx186ejSSeb21FREtLy97vrTNUbGttSdNiW6uZoHFBb29v3/39j3L/pYjoaC8tVh5FxOLiUkf7+Re3Fl2T96ZfWoMPHIIGPeX4t1oduzkeSVSr1XL/pa7ONyPioysDt3+7OzU9GxEff/jB3vMv9lyYvHd/8IvP9l+qzhAco8QnVngV99AgaBA0CBoEjaBB0CBoEDSCdgsQNJxQ6en7la4NXjWvh3d9dMIKDYIGQcOrvofeq6+3wxzvNzO/YoUGQYOgwR76mHeWR9x8n8xLWaFB0CBoEDQIGkGDoEHQIGgQNIIGQYOgQdAgaAQNggZBg6ARNAgaBA2CBkEjaBA0CBoEDYJG0CBoEDQIGgSNoEHQIGgQNAgaQYOgQdAgaBA0ggZBg6BB0CBoBA2CBkGDoBG0W4CgQdAgaBA0ggZBg6BB0CBoBA2CBkGDoEHQCBoEDYIGQYOgETQ0m9QtqOnr7Tjdl7JCg6BB0GAPfbCZ+RVzbIUGQcMJkGytLzfdD31t8KqZa4zroxNWaBA0CBpe0vSP7fxn+Ng19bNOKzS2HCBoEDQIGkGDoEHQIGgQNIIGQYOgQdAgaAQNgoYToOk/seJVMlihETQIGhqsKd+ctN+B71JqujcAuVdWaBA0ggZBg6BB0AgaBA2CBkGDoBE0CBoEDYIGQSNoEDQIGgQNgkbQIGgQNAgaBI2gQdAgaBA0CBpBg6BB0CBoEDSCBkGDoEHQCBoEDYIGQYOgETQIGgQNggZBI2gQNAgaBA2CRtAgaBA0CBoEjaBB0CBoEDQIGkGDoEHQIGioSbbWl90FrNAgaBA0CBpBg6BB0CBoEDSCBkGDoEHQIGgEDYIGQYOgQdAIGgQNggZBw4v+A8Zh8TqzsZNyAAAAAElFTkSuQmCC',
    'SKU002': 'iVBORw0KGgoAAAANSUhEUgAAAPAAAADwCAIAAACxN37FAAAD8UlEQVR42u3c70tddRzA8c+5Oy5oec+5c6XTturRmHCZ0aAgWI+KoAf7R/pz+g+EkJ6IRCVbY9NYPwYzsVnCKmRMr3NtenXzgdg9PRDEip0ZbkfaXq9Hcr5fr/I9b758lXNvstFeDnhW1CwBggZBg6BB0AgaBA2CBkHD/oMeHhkdvzgx/vXElxcu//r7/PbFTz8b25mw8/XwyOiVb77fuT559drwyGj50Gp7beyLC0VRbF///KtLK6tr7gf7lJaMHarVPnz/vYjY2tq6NPFtmqavn3z1UTPba+tFUSRJEhHrDx4cqtXKh/KsnmfZ/K3bb7x24vZCq7v7SCOvux9UceRI0/Tsm82f526WzOk52vjj3kpE3F9ZbeT5XoaGmqdnbvxSFMXM7NxQc9DNoLozdCPP1tcflkwY6O9daC1FxMLi0kB/716GsqzeyPPvrk29dOTFPLM9U2HQRVHUaknJhP7jva3WckS07tzt73tlj0Nnmqdv/jZ/xvZMxUHfvXc/z7OIiCTZ/kuuUxTJrsRfOHw4kni4sRERXV1du7+3ZCird3elaVbvdieoLujNzc3rP/7UHDwVEcd6GoutOxGxuLh0rOfo348WfVPTs//Ygx87BBX9l+PPTmf84kQk0el0moOn+npfjoh3zg5d/eH6zOxcRLz79lu7558YOD41feP8Rx/8+6VKhuAJSrxjhefxDA2CBkGDoEHQCBoEDYIGQSNoS4CgQdAgaBA0ggZBg6BB0CBoBA2CBkGDoEHQCBoEDYIGQYOgETQIGgQNggZBI2gQNAgaBA2CRtAgaBA0CBoEjaBB0CBoEDSCBkGDoEHQIGgEDYIGQYOgQdAIGgQNgoYnLn16L/3x+XPWl0f5ZGzSDg2CRtDgDF3m3NBJa83k9C07NAgaQYOgQdAgaBA0ggZBg6BB0CBoBA2ChoNW0eOjFTw3CHZoBA2ChmokG+3lA/nBj/2Qg5K3ufuAhMrs5y48pQ8qsEPjyAGCBkGDoEHQCBoEDYIGQYOgeaal/8df+kAeEsAODYIGQYOgETQIGgQNgkbQlgBBg6BB0CBoBA2CBkGDoEHQCBoEDYIGQYOgETQIGgQNggZBI2gQNAgaBA2CRtAgaBA0CBoEjaBB0CBoEDQIGkGDoEHQIGgEDYIGQYOgQdAIGgQNggZBg6ARNAgaBA2CBkEjaBA0CBoEDYJG0CBoEDQIGgSNoEHQIGgQNAgaQYOgQdAgaBA0ggZBg6Dhv0k22stWATs0CBoEDYJG0CBoEDQIGgSNoEHQIGgQNAgaQYOgQdAgaBA0ggZBg6BB0CBoBA2CBkGDoEHQCBoEDYIGQYOgETQIGgQNguY59hesx+CEKisitwAAAABJRU5ErkJggg=='
  };
  const m = sh_(SH.MASTER);
  let n = 0;
  Object.keys(demo).forEach(function (k) {
    const r = barisSku_(k);
    if (!r) return;
    const img = SpreadsheetApp.newCellImage().setSourceUrl('data:image/png;base64,' + demo[k]).setAltTextTitle(k).build();
    m.getRange(r, 3).setValue(img);
    m.setRowHeightsForced(r, 1, H_MASTER);
    n++;
  });
  info_(n + ' foto contoh dipasang (SKU001, SKU002).', 'Foto Contoh');
}

// ======================================================================
// SKU BARU
// ======================================================================
function TambahSKU() {
  const m = sh_(SH.MASTER);
  let sku = prompt_('Kode Barang / SKU baru (contoh: SKU003):', 'Tambah SKU');
  if (!sku) return;
  sku = sku.toUpperCase();
  const nama = prompt_('Nama Barang:', 'Tambah SKU'); if (!nama) return;
  const kat = prompt_('Kategori (mis. Kursi, Meja, Lemari):', 'Tambah SKU') || '';
  const sat = prompt_('Satuan:', 'Tambah SKU', 'Unit') || 'Unit';
  const hasil = dalamKunci_(function () {
    const v = m.getRange(R1, 2, M_LAST - R1 + 1, 1).getValues();
    let r = 0, pakai = 0;
    for (let i = 0; i < v.length; i++) {
      const k = txt_(v[i][0]);
      if (k === sku) return { dup: true };
      if (k !== '') pakai++; else if (!r) r = R1 + i;
    }
    if (!r) return { penuh: true };
    m.getRange(r, 2).setNumberFormat('@').setValue(sku);
    m.getRange(r, 4).setValue(nama); m.getRange(r, 5).setValue(kat); m.getRange(r, 8).setValue(sat); m.getRange(r, 10).setValue('Aktif');
    m.setRowHeightsForced(r, 1, H_MASTER);
    return { r: r, pakai: pakai + 1 };
  });
  if (hasil.dup) { info_('SKU ' + sku + ' sudah ada. SKU harus unik.', 'Tambah SKU'); return; }
  if (hasil.penuh) { info_('Kapasitas SKU penuh (' + (M_LAST - R1 + 1) + ').', 'Tambah SKU'); return; }
  m.setActiveRange(m.getRange(hasil.r, 2));
  kapasitasToast_('SKU', hasil.pakai, M_LAST - R1 + 1);
  if (ask_('SKU ' + sku + ' ditambahkan di baris ' + hasil.r + '.\nTambahkan foto sekarang?', 'Tambah SKU')) TambahFotoProduk();
}

// ======================================================================
// HAPUS SKU (barang) - baris master dikosongkan, bukan dihapus, agar rumus & kapasitas tetap utuh
// ======================================================================
// HARUS dipanggil dari dalam kode yang sudah memegang kunci ATAU lewat hapusSkuInti_ (yang mengunci sendiri).
// Ditolak bila SKU sudah dipakai di TRANSAKSI_STOK (riwayat tidak boleh kehilangan barangnya).
function hapusSkuInti_(sku) {
  sku = txt_(sku);
  return dalamKunci_(function () {
    const row = barisSku_(sku);
    if (!row) throw new Error('SKU tidak ditemukan.');
    const jml = dataStok_().trx.filter(function (r) { return txt_(r[3]) === sku; }).length;
    if (jml > 0) {
      return { ditolak: true, jml: jml, pesan: 'SKU ' + sku + ' sudah dipakai di ' + jml + ' transaksi, jadi tidak dapat dihapus (riwayat transaksi akan rusak).\nGunakan status Nonaktif untuk menyembunyikannya.' };
    }
    const m = sh_(SH.MASTER);
    const nama = txt_(m.getRange(row, 4).getValue());
    // kosongkan kolom isian saja: B..J, L (arsip foto), M..U (stok awal per ruangan). Kolom A, K, V berisi rumus: jangan disentuh.
    m.getRange(row, 2, 1, 9).clearContent();
    m.getRange(row, 12, 1, JML_RUANG + 1).clearContent();
    // buang file foto di Drive (ke tempat sampah, masih bisa dipulihkan)
    try {
      const folder = folder_(FOLDER_FOTO), safe = sku.replace(/[\\\/:*?"<>|]/g, '_'), it = folder.getFiles();
      while (it.hasNext()) {
        const f = it.next(), nm = f.getName().replace(/\.[^.]+$/, '');
        if (nm === safe || nm === 'thumb_' + safe) f.setTrashed(true);
      }
    } catch (e) { /* foto di Drive bersifat tambahan */ }
    CacheService.getScriptCache().remove('foto_' + sku);
    try { m.setRowHeightsForced(row, 1, H_DEFAULT); } catch (e) { /* abaikan */ }
    const h = sinkronSaldo_();
    return { ok: true, nama: nama, bad: h.bad };
  });
}

function HapusSKU() {
  const a = skuAktif_('Hapus SKU');
  if (!a) return;
  const nama = txt_(sh_(SH.MASTER).getRange(a.row, 4).getValue());
  if (!ask_('Hapus SKU ' + a.sku + (nama ? ' - ' + nama : '') + '?\nData barang, foto, dan stok awal di semua ruangan ikut terhapus. Tindakan ini tidak dapat dibatalkan.', 'Hapus SKU')) return;
  try {
    const h = hapusSkuInti_(a.sku);
    if (h.ditolak) { info_(h.pesan, 'Hapus SKU'); return; }
    catat_('Hapus SKU', Session.getActiveUser().getEmail() || 'Sheets', a.sku + ' ' + h.nama);
    info_('SKU ' + a.sku + ' dihapus.' + (h.bad.length ? '\n\nPerlu diperiksa:\n' + h.bad.slice(0, 5).join('\n') : ''), 'Hapus SKU');
  } catch (e) { info_('Gagal: ' + e.message, 'Hapus SKU'); }
}

// ======================================================================
// NOMOR TRANSAKSI, STOK, VALIDASI
// ======================================================================
// ======================================================================
// KUNCI, NOMOR, PERISTIWA STOK (v1.1)
// ======================================================================
// Semua penulisan nomor/baris dilindungi kunci dokumen agar tidak bentrok antar pengguna.
function dalamKunci_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('Sistem sedang dipakai pengguna lain. Coba lagi dalam beberapa detik.');
  try { return fn(); } finally { lock.releaseLock(); }
}
function tglDari_(v) { return v instanceof Date ? v : new Date(); }
function waktu_(r) { return r[0] instanceof Date ? r[0].getTime() : Infinity; }
function dorong_(ev, k, o) { (ev[k] || (ev[k] = [])).push(o); }

function kapasitasToast_(nama, pakai, kap) {
  if (kap > 0 && pakai / kap >= KAPASITAS_PERINGATAN) {
    ss_().toast(nama + ' terpakai ' + pakai + ' dari ' + kap + ' (' + Math.round(pakai * 100 / kap) + '%). Segera perluas workbook/arsipkan data.', 'Kapasitas hampir penuh', 10);
  }
}

// Membuat n nomor unik sekaligus. HARUS dipanggil di dalam dalamKunci_().
// commit=false: hanya pratinjau (penghitung tidak dimajukan).
function nomorBatch_(kunci, tglList, commit) {
  const set = sh_(SH.SET);
  const cfg = kunci === 'TRX' ? { pf: 'C14', no: 'C15', dg: 'C16' } : { pf: 'C17', no: 'C18', dg: 'C19' };
  const pf = txt_(set.getRange(cfg.pf).getValue());
  let n = Math.max(1, num_(set.getRange(cfg.no).getValue()));
  const dg = Math.max(1, num_(set.getRange(cfg.dg).getValue()) || 4);
  const ada = {};
  const rng = kunci === 'TRX' ? sh_(SH.TRX).getRange(R1, 3, T_LAST - R1 + 1, 1) : sh_(SH.RIW).getRange(R1, 1, H_LAST - R1 + 1, 1);
  rng.getValues().forEach(function (r) { const t = txt_(r[0]); if (t) ada[t] = true; });
  const out = [];
  tglList.forEach(function (tgl) {
    let s;
    for (;;) {
      s = pf + '-' + tgl.getFullYear() + '-' + String(n).padStart(dg, '0');
      n++;
      if (!ada[s]) break;
    }
    ada[s] = true;
    out.push(s);
  });
  if (commit !== false) set.getRange(cfg.no).setValue(n);
  return out;
}

function BuatNomorTransaksi() {
  wajibMenu_();
  const t = sh_(SH.TRX);
  const n = dalamKunci_(function () {
    const last = lastNonEmpty_(t, 4, R1, T_LAST);
    if (last < R1) return -1;
    const v = t.getRange(R1, 2, last - R1 + 1, 3).getValues();   // B (tgl), C (no), D (jenis)
    const idx = [];
    v.forEach(function (r, i) { if (txt_(r[2]) !== '' && txt_(r[1]) === '') idx.push(i); });
    if (!idx.length) return 0;
    const nos = nomorBatch_('TRX', idx.map(function (i) { return tglDari_(v[i][0]); }));
    const crg = t.getRange(R1, 3, last - R1 + 1, 1), cv = crg.getValues();
    idx.forEach(function (i, k) { cv[i][0] = nos[k]; });
    crg.setNumberFormat('@'); crg.setValues(cv);
    return idx.length;
  });
  info_(n < 0 ? 'Belum ada transaksi.' : n + ' nomor transaksi dibuat.', 'Nomor Transaksi');
}

// Membaca hanya baris yang terpakai (jauh lebih cepat dari membaca seluruh kapasitas).
function dataStok_() {
  const sa = sh_(SH.SALDO), t = sh_(SH.TRX);
  const ls = lastNonEmpty_(sa, 2, R1, S_LAST), lt = lastNonEmpty_(t, 4, R1, T_LAST);
  return {
    saldo: ls >= R1 ? sa.getRange(R1, 2, ls - R1 + 1, 6).getValues() : [],     // B..G : B sku, E ruang, G qty
    trx: lt >= R1 ? t.getRange(R1, 2, lt - R1 + 1, 14).getValues() : []        // B..O
  };
}

function jenisKeluarRuang_(j, J) { return j === J.keluar || j === J.adjm || j === J.transfer; }

function negatifGlobal_() { return txt_(sh_(SH.SET).getRange('C20').getValue()).toUpperCase() === 'YA'; }
function negatifOk_(approval, global) { return global || txt_(approval) !== ''; }

// efek satu baris transaksi: [[kunci 'SKU|RUANG', delta], ...]
function efekTrx_(r, J) {
  const jn = txt_(r[2]), sku = txt_(r[3]), asal = txt_(r[6]), tuj = txt_(r[7]), q = Number(r[8]);
  if (!sku || r[8] === '' || r[8] === null || isNaN(q)) return [];
  const out = [];
  if (jn === J.masuk) { if (tuj) out.push([sku + '|' + tuj, q]); }
  else if (jn === J.keluar) { if (asal) out.push([sku + '|' + asal, -q]); }
  else if (jn === J.transfer) { if (asal) out.push([sku + '|' + asal, -q]); if (tuj) out.push([sku + '|' + tuj, q]); }
  else if (jn === J.adjp) { if (tuj) out.push([sku + '|' + tuj, q]); }
  else if (jn === J.adjm) { if (asal) out.push([sku + '|' + asal, -q]); }
  return out;
}

// indeks peristiwa per kunci SKU|RUANG: saldo awal (waktu paling awal) + semua transaksi
function bangunPeristiwa_(data, J, kecuali) {
  const ev = {};
  data.saldo.forEach(function (r) {
    const sku = txt_(r[0]), ru = txt_(r[3]);
    if (!sku || !ru || txt_(r[5]) === '' || isNaN(Number(r[5]))) return;
    dorong_(ev, sku + '|' + ru, { t: -Infinity, i: -1, d: Number(r[5]) });
  });
  data.trx.forEach(function (r, i) {
    if (kecuali && kecuali[i]) return;
    const w = waktu_(r);
    efekTrx_(r, J).forEach(function (e) { dorong_(ev, e[0], { t: w, i: i, d: e[1] }); });
  });
  return ev;
}

// saldo berjalan terendah menurut urutan tanggal, lalu urutan baris
function minKronologis_(list) {
  if (!list || !list.length) return 0;
  const a = list.slice().sort(function (x, y) { return x.t === y.t ? x.i - y.i : (x.t < y.t ? -1 : 1); });
  let s = 0, mn = 0;
  a.forEach(function (o) { s += o.d; if (s < mn) mn = s; });
  return mn;
}

function ketKunci_(kunci, nilai) {
  const p = kunci.split('|');
  return p[0] + ' di ruangan ' + p[1] + ' menjadi ' + nilai;
}

// ======================================================================
// PEMICU OTOMATIS (installable) - tangani edit satu sel MAUPUN tempel banyak baris
// ======================================================================
function saatEdit(e) {
  try {
    const r = e.range, s = r.getSheet(), nm = s.getName();
    if (nm === SH.TRX) {
      const c1 = r.getColumn(), c2 = c1 + r.getNumColumns() - 1;
      const kena = TRX_COLS.some(function (c) { return c >= c1 && c <= c2; });
      const r1 = Math.max(r.getRow(), R1), r2 = Math.min(r.getLastRow(), T_LAST);
      if (kena && r2 >= r1) prosesRentangTrx_(s, r1, r2);
    } else if (nm === SH.FORM && r.getA1Notation() === 'D7') {
      muatForm_();
    } else if (nm === SH.MASTER) {
      const c1 = r.getColumn(), c2 = c1 + r.getNumColumns() - 1;
      const kenaSku = c1 <= 2 && c2 >= 2, kenaQty = c2 >= KOL_SALDO1 && c1 <= KOL_SALDO1 + JML_RUANG - 1;
      const r1 = Math.max(r.getRow(), R1), r2 = Math.min(r.getLastRow(), M_LAST);
      if ((kenaSku || kenaQty) && r2 >= r1) {
        const h = dalamKunci_(sinkronSaldo_);
        const pesan = [];
        if (h.bad.length) pesan.push('Perlu diperiksa: ' + h.bad.slice(0, 3).join('; '));
        if (kenaQty && lastNonEmpty_(sh_(SH.TRX), 4, R1, T_LAST) >= R1) pesan.push('Sudah ada transaksi: mengubah stok awal ikut mengubah stok akhir.');
        if (pesan.length) ss_().toast(pesan.join(' '), 'Saldo Awal', 8);
      }
    }
  } catch (err) {
    ss_().toast(String(err && err.message ? err.message : err), 'EZRIA - galat pemicu', 10);
  }
}

function prosesRentangTrx_(s, r1, r2) {
  const J = jenis_();
  const hasil = dalamKunci_(function () {
    const data = dataStok_();
    const nR = data.trx.length;
    const target = [], kecuali = {};
    for (let row = r1; row <= r2; row++) {
      const i = row - R1;
      if (i < 0 || i >= nR) continue;
      const v = data.trx[i];
      if (!txt_(v[2]) || !txt_(v[3]) || v[8] === '' || isNaN(Number(v[8]))) continue;
      target.push(i); kecuali[i] = true;
    }
    if (!target.length) return null;
    const global = negatifGlobal_();
    const ev = bangunPeristiwa_(data, J, kecuali);
    const diterima = [], ditolak = [];
    target.forEach(function (i) {
      const v = data.trx[i], jn = txt_(v[2]), w = waktu_(v), efek = efekTrx_(v, J);
      efek.forEach(function (e) { dorong_(ev, e[0], { t: w, i: i, d: e[1] }); });
      let bad = '';
      if (jenisKeluarRuang_(jn, J) && !negatifOk_(v[13], global)) {
        for (let k = 0; k < efek.length && !bad; k++) {
          if (efek[k][1] < 0) { const m = minKronologis_(ev[efek[k][0]]); if (m < 0) bad = ketKunci_(efek[k][0], m); }
        }
      }
      if (bad) {
        efek.forEach(function (e) { ev[e[0]] = ev[e[0]].filter(function (x) { return x.i !== i; }); });
        ditolak.push({ i: i, info: bad, no: R1 + i });
      } else diterima.push(i);
    });
    ditolak.forEach(function (d) { s.getRange(R1 + d.i, 10).clearContent(); });
    const perlu = diterima.filter(function (i) { return txt_(data.trx[i][1]) === ''; });
    if (perlu.length) {
      const nos = nomorBatch_('TRX', perlu.map(function (i) { return tglDari_(data.trx[i][0]); }));
      const a = Math.min.apply(null, perlu), b = Math.max.apply(null, perlu);
      const rg = s.getRange(R1 + a, 3, b - a + 1, 1), cur = rg.getValues();
      perlu.forEach(function (i, k) { cur[i - a][0] = nos[k]; });
      rg.setNumberFormat('@'); rg.setValues(cur);
    }
    return { ditolak: ditolak, ok: diterima.length, pakai: nR };
  });
  if (!hasil) return;
  if (hasil.ditolak.length) {
    ss_().toast('Ditolak ' + hasil.ditolak.length + ' baris (qty dikosongkan): ' +
      hasil.ditolak.slice(0, 5).map(function (d) { return 'baris ' + d.no + ' - ' + d.info; }).join('; ') +
      '. Isi kolom Persetujuan bila disetujui khusus.', 'Stok Negatif', 12);
  }
  kapasitasToast_('Transaksi', hasil.pakai, T_LAST - R1 + 1);
}

// ---- ValidasiTransaksi : simulasi stok berurutan menurut TANGGAL (lalu urutan baris) ----
function validasiInti_() {
  SpreadsheetApp.flush();
  const t = sh_(SH.TRX), J = jenis_(), global = negatifGlobal_();
  const data = dataStok_();
  const stok = {};
  data.saldo.forEach(function (r) {
    const k = txt_(r[0]), ru = txt_(r[3]);
    if (k && ru && !isNaN(Number(r[5])) && txt_(r[5]) !== '') stok[k + '|' + ru] = (stok[k + '|' + ru] || 0) + Number(r[5]);
  });
  const n = T_LAST - R1 + 1, nr = data.trx.length;
  const hasil = [], warna = [];
  for (let i = 0; i < n; i++) { hasil.push(['']); warna.push([null]); }
  const urut = [];
  for (let i = 0; i < nr; i++) if (txt_(data.trx[i][2]) !== '' || txt_(data.trx[i][3]) !== '') urut.push(i);
  urut.sort(function (a, b) {
    const x = waktu_(data.trx[a]), y = waktu_(data.trx[b]);
    return x === y ? a - b : (x < y ? -1 : 1);
  });
  let ok = 0, bad = 0, neg = 0;
  urut.forEach(function (i) {
    const r = data.trx[i], jn = txt_(r[2]), sku = txt_(r[3]);
    const st = txt_(r[12]);                // N (Validasi, rumus)
    if (st !== 'OK') { bad++; hasil[i][0] = 'TIDAK VALID: ' + st; warna[i][0] = '#F8D7DA'; return; }
    const asal = txt_(r[6]), tuj = txt_(r[7]), q = num_(r[8]);
    const kS = sku + '|' + asal, kD = sku + '|' + tuj;
    if (jenisKeluarRuang_(jn, J)) stok[kS] = (stok[kS] || 0) - q;
    if (jn === J.masuk || jn === J.adjp || jn === J.transfer) stok[kD] = (stok[kD] || 0) + q;
    if (jenisKeluarRuang_(jn, J) && stok[kS] < 0 && !negatifOk_(r[13], global)) {
      neg++; hasil[i][0] = 'STOK NEGATIF di ruangan ' + asal + ' (' + stok[kS] + ')'; warna[i][0] = '#FFF3CD';
    } else { ok++; hasil[i][0] = 'OK'; }
  });
  const rg = t.getRange(R1, 16, n, 1);
  rg.setValues(hasil); rg.setBackgrounds(warna);
  return { ok: ok, bad: bad, neg: neg };
}

function ValidasiTransaksi() {
  wajibMenu_();
  const h = dalamKunci_(validasiInti_);
  info_('Hasil validasi transaksi (diurutkan menurut tanggal):\nValid: ' + h.ok + '\nTidak valid (data): ' + h.bad + '\nMenyebabkan stok negatif: ' + h.neg +
    (h.bad + h.neg ? '\n\nLihat kolom "Hasil Cek Makro" (P). Perbaiki lewat transaksi koreksi.' : ''), 'Validasi Transaksi');
}

function KoreksiTransaksi() {
  const t = sh_(SH.TRX), J = jenis_();
  const no = prompt_('Nomor transaksi yang dikoreksi (mis. TRX-2026-000003):', 'Koreksi Transaksi');
  if (!no) return;
  const data = dataStok_();
  let idx = -1;
  for (let i = 0; i < data.trx.length; i++) if (txt_(data.trx[i][1]) === no) { idx = i; break; }
  if (idx < 0) { info_('Nomor transaksi tidak ditemukan.', 'Koreksi'); return; }
  const r = data.trx[idx];
  const jn = txt_(r[2]), asal = txt_(r[6]), tuj = txt_(r[7]);
  let jb, a2 = '', t2 = '';
  if (jn === J.masuk) { jb = J.keluar; a2 = tuj; }
  else if (jn === J.keluar) { jb = J.masuk; t2 = asal; }
  else if (jn === J.transfer) { jb = J.transfer; a2 = tuj; t2 = asal; }
  else if (jn === J.adjp) { jb = J.adjm; a2 = tuj; }
  else if (jn === J.adjm) { jb = J.adjp; t2 = asal; }
  else { info_('Jenis transaksi tidak dikenali.', 'Koreksi'); return; }
  const sudah = data.trx.some(function (x) { return txt_(x[10]).indexOf('KOREKSI atas ' + no) === 0; });
  if (sudah && !ask_('Transaksi ' + no + ' SUDAH pernah dikoreksi.\nKoreksi lagi akan menggandakan pembalikan. Tetap lanjutkan?', 'Koreksi')) return;
  const alasan = prompt_('Alasan koreksi:', 'Koreksi'); if (!alasan) return;
  const petugas = prompt_('Nama petugas:', 'Koreksi'); if (!petugas) return;

  // periksa dampak stok dari transaksi pembalik (tanggal hari ini)
  const rv = new Array(14).fill('');
  rv[0] = new Date(); rv[2] = jb; rv[3] = txt_(r[3]); rv[6] = a2; rv[7] = t2; rv[8] = Number(r[8]);
  const ev = bangunPeristiwa_(data, J, null);
  const efek = efekTrx_(rv, J);
  efek.forEach(function (e) { dorong_(ev, e[0], { t: rv[0].getTime(), i: 1e9, d: e[1] }); });
  let persetujuan = '', neg = '';
  efek.forEach(function (e) { if (e[1] < 0 && !neg) { const m = minKronologis_(ev[e[0]]); if (m < 0) neg = ketKunci_(e[0], m); } });
  if (neg) {
    if (!ask_('PERINGATAN: koreksi ini membuat stok ' + neg + '.\nLanjutkan dengan persetujuan khusus (dicatat pada kolom Persetujuan)?', 'Koreksi - stok negatif')) return;
    persetujuan = 'Koreksi disetujui: ' + petugas;
  }
  if (!ask_('Buat transaksi pembalik untuk ' + no + ' (' + jn + ' ' + txt_(r[3]) + ' qty ' + r[8] + ')?\nTransaksi lama tetap tersimpan.', 'Koreksi')) return;

  const rn = dalamKunci_(function () {
    const baris = lastNonEmpty_(t, 4, R1, T_LAST) + 1;
    if (baris > T_LAST) return -1;
    const tgl = new Date();
    const nomor = nomorBatch_('TRX', [tgl])[0];
    t.getRange(baris, 2).setValue(tgl);
    t.getRange(baris, 3).setNumberFormat('@').setValue(nomor);
    t.getRange(baris, 4).setValue(jb);
    t.getRange(baris, 5).setNumberFormat('@').setValue(txt_(r[3]));
    if (a2) t.getRange(baris, 8).setNumberFormat('@').setValue(a2);
    if (t2) t.getRange(baris, 9).setNumberFormat('@').setValue(t2);
    t.getRange(baris, 10).setValue(r[8]);
    t.getRange(baris, 12).setValue('KOREKSI atas ' + no + ' - ' + alasan);
    t.getRange(baris, 13).setValue(petugas);
    if (persetujuan) t.getRange(baris, 15).setValue(persetujuan);
    t.setRowHeightsForced(baris, 1, H_LIST);
    return baris;
  });
  if (rn < 0) { info_('Kapasitas transaksi penuh.', 'Koreksi'); return; }
  info_('Transaksi koreksi dibuat di baris ' + rn + '.', 'Koreksi');
}

// ======================================================================
// STOCK OPNAME
// ======================================================================
function muatForm_() {
  const f = sh_(SH.FORM);
  f.getRange(F_R1, 7, F_LAST - F_R1 + 1, 1).clearContent();
  f.getRange(F_R1, 9, F_LAST - F_R1 + 1, 1).clearContent();
  SpreadsheetApp.flush();
  const kode = txt_(f.getRange('D8').getValue());
  const skus = f.getRange(F_R1, 3, F_LAST - F_R1 + 1, 1).getValues().map(function (r) { return [txt_(r[0])]; });
  f.getRange(F_R1, 12, F_LAST - F_R1 + 1, 1).setValues(skus);
  f.getRange('L5').setNumberFormat('@').setValue(kode);
  f.getRange('L6').setValue('');
  f.getRange('I10').setValue('BELUM DISIMPAN');
  const n = skus.filter(function (r) { return r[0] !== ''; }).length;
  const last = lastNonEmpty_(f, 3, F_R1, F_LAST);
  if (last >= F_R1) f.setRowHeightsForced(F_R1, last - F_R1 + 1, H_FORM);
  if (last < F_LAST) f.setRowHeightsForced(Math.max(last + 1, F_R1), F_LAST - Math.max(last + 1, F_R1) + 1, H_DEFAULT);
  return n;
}

function MuatDataRuangan() {
  wajibMenu_();
  const f = sh_(SH.FORM);
  if (!txt_(f.getRange('D7').getValue())) { info_('Pilih Nama Ruangan terlebih dahulu (sel D7).', 'Stock Opname'); return; }
  const ada = f.getRange(F_R1, 7, F_LAST - F_R1 + 1, 1).getValues().some(function (r) { return txt_(r[0]) !== ''; });
  if (ada && !ask_('Isian Stok Fisik/Keterangan yang ada akan DIKOSONGKAN (riwayat tidak terpengaruh). Lanjutkan?', 'Muat Data Ruangan')) return;
  const n = muatForm_();
  info_(n + ' barang dimuat dari ' + txt_(f.getRange('D7').getValue()) + '. Silakan isi Stok Fisik.', 'Muat Data Ruangan');
}

function ResetFormOpname() {
  if (!ask_('Kosongkan Stok Fisik, Keterangan baris, dan Keterangan form?\nRiwayat yang sudah tersimpan TIDAK dihapus.', 'Reset Form')) return;
  const f = sh_(SH.FORM);
  f.getRange(F_R1, 7, F_LAST - F_R1 + 1, 1).clearContent();
  f.getRange(F_R1, 9, F_LAST - F_R1 + 1, 1).clearContent();
  f.getRange('D10').clearContent();
  f.getRange('L6').setValue('');
  f.getRange('I10').setValue('BELUM DISIMPAN');
  info_('Form direset.', 'Reset Form');
}

function SimpanStockOpname(dariCetak) {
  SpreadsheetApp.flush();
  const f = sh_(SH.FORM), h = sh_(SH.RIW);
  const noSimpan = txt_(f.getRange('L6').getValue());
  const lastH0 = lastNonEmpty_(h, 1, R1, H_LAST);
  const riw = lastH0 >= R1 ? h.getRange(R1, 1, lastH0 - R1 + 1, 3).getValues() : [];   // A no, B tgl, C ruangan
  // penjaga dobel-simpan: berdasarkan RIWAYAT (bukan hanya teks status di form)
  if ((noSimpan !== '' && riw.some(function (x) { return txt_(x[0]) === noSimpan; })) || txt_(f.getRange('I10').getValue()).indexOf('TERSIMPA') === 0) {
    info_('Form ini sudah disimpan (' + noSimpan + '). Gunakan "Muat Data Ruangan" untuk pemeriksaan baru.', 'Simpan Stock Opname'); return false;
  }
  const tgl = f.getRange('D6').getValue();
  if (!(tgl instanceof Date)) { info_('Tanggal Pemeriksaan (D6) belum diisi.', 'Simpan Stock Opname'); return false; }
  const ruang = txt_(f.getRange('D7').getValue()), kode = txt_(f.getRange('D8').getValue()), petugas = txt_(f.getRange('D9').getValue());
  if (!ruang || !kode) { info_('Nama Ruangan belum dipilih.', 'Simpan Stock Opname'); return false; }
  if (!petugas) { info_('Petugas Pemeriksa (D9) belum diisi.', 'Simpan Stock Opname'); return false; }
  if (txt_(f.getRange('L5').getValue()) !== kode) { info_('Klik "Muat Data Ruangan" terlebih dahulu untuk ruangan ini.', 'Simpan Stock Opname'); return false; }
  const rows = f.getRange(F_R1, 1, F_LAST - F_R1 + 1, 12).getValues();   // A..L
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i], sku = txt_(r[2]);
    if (sku !== txt_(r[11])) { info_('Daftar barang berubah sejak dimuat (baris ' + (F_R1 + i) + '). Klik "Muat Data Ruangan" ulang.', 'Simpan Stock Opname'); return false; }
    if (!sku) continue;
    if (txt_(r[6]) === '' || isNaN(Number(r[6]))) { info_('Stok Fisik belum terisi/tidak valid pada baris ' + (F_R1 + i) + ' (' + sku + ').', 'Simpan Stock Opname'); return false; }
    if (Number(r[6]) < 0) { info_('Stok Fisik tidak boleh negatif (baris ' + (F_R1 + i) + ').', 'Simpan Stock Opname'); return false; }
    out.push({ sku: sku, nama: txt_(r[3]), sistem: num_(r[5]), fisik: Number(r[6]), ket: txt_(r[8]) });
  }
  if (!out.length) { info_('Tidak ada barang untuk disimpan.', 'Simpan Stock Opname'); return false; }
  // peringatan opname ganda untuk ruangan + tanggal yang sama
  const tz = ss_().getSpreadsheetTimeZone();
  const hari = Utilities.formatDate(tgl, tz, 'yyyyMMdd');
  const sama = {};
  riw.forEach(function (x) { if (txt_(x[2]) === ruang && x[1] instanceof Date && Utilities.formatDate(x[1], tz, 'yyyyMMdd') === hari) sama[txt_(x[0])] = true; });
  const daftarSama = Object.keys(sama);
  if (daftarSama.length && !ask_('Sudah ada opname ' + ruang + ' pada tanggal yang sama (' + daftarSama.join(', ') + ').\nTetap simpan sebagai opname baru?', 'Opname Ganda')) return false;
  const perkiraan = dalamKunci_(function () { return nomorBatch_('OP', [tgl], false)[0]; });
  if (!ask_('Simpan Stock Opname ' + ruang + ' (' + out.length + ' item), perkiraan nomor ' + perkiraan + '?\nStok sistem TIDAK berubah. Selisih ditindaklanjuti dengan transaksi Penyesuaian.', 'Simpan Stock Opname')) return false;
  const n = out.length;
  const hasil = dalamKunci_(function () {
    const lastH = lastNonEmpty_(h, 1, R1, H_LAST);
    const baris = Math.max(lastH, R1 - 1) + 1;
    if (baris + n - 1 > H_LAST) return null;
    const no = nomorBatch_('OP', [tgl])[0];
    h.getRange(baris, 1, n, 4).setNumberFormat('@');
    h.getRange(baris, 1, n, 4).setValues(out.map(function (o) { return [no, tgl, ruang, o.sku]; }));
    h.getRange(baris, 2, n, 1).setNumberFormat('DD/MM/YYYY');
    h.getRange(baris, 6, n, 6).setValues(out.map(function (o) { return [o.nama, o.sistem, o.fisik, o.fisik - o.sistem, o.ket, petugas]; }));
    h.setRowHeightsForced(baris, n, H_LIST);
    return { no: no, pakai: baris + n - R1 };
  });
  if (!hasil) { info_('Kapasitas riwayat tidak cukup.', 'Simpan Stock Opname'); return false; }
  const no = hasil.no;
  f.getRange('L6').setValue(no);
  f.getRange('I10').setValue('TERSIMPAN: ' + no);
  const c = sh_(SH.CETAK);
  c.getRange('D2').setValue(ruang); c.getRange('D3').setValue(no);
  SpreadsheetApp.flush();
  kapasitasToast_('Riwayat opname', hasil.pakai, H_LAST - R1 + 1);
  if (!dariCetak && ask_('Stock Opname ' + no + ' tersimpan di RIWAYAT_OPNAME (' + n + ' item).\nBuat PDF untuk dicetak sekarang?', 'Simpan Stock Opname')) CetakRuanganTerpilih();
  else if (dariCetak) ss_().toast('Tersimpan: ' + no, 'EZRIA', 5);
  return true;
}

// ======================================================================
// CETAK (PDF)  - pilih ruangan di CETAK_STOCK_OPNAME!D2
// ======================================================================
function exportPdf_(sheet, a1, nama, folder) {
  const id = ss_().getId();
  const url = 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=pdf&gid=' + sheet.getSheetId() + '&range=' + a1 +
    '&portrait=true&size=A4&fitw=true&gridlines=false&sheetnames=false&printtitle=false&pagenum=UNDEFINED&attachment=false' +
    '&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4&horizontal_alignment=CENTER';
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Gagal membuat PDF (kode ' + res.getResponseCode() + ').');
  return folder.createFile(res.getBlob().setName(nama + '.pdf'));
}

// membuat PDF untuk ruangan yang sedang terpilih; mengembalikan {nama,url,n} atau null
function buatPdfCetak_(senyap) {
const c = sh_(SH.CETAK);

// Atur ukuran foto khusus untuk cetak
c.setColumnWidth(2, C_PRINT_PHOTO_WIDTH);
c.setRowHeightsForced(C_R1, C_LAST - C_R1 + 1, C_PRINT_ROW_HEIGHT);

SpreadsheetApp.flush();
  if (!txt_(c.getRange('D2').getValue())) { if (!senyap) info_('Pilih ruangan terlebih dahulu pada kotak kuning (D2).', 'Cetak'); return null; }
  const kode = txt_(c.getRange('K2').getValue());
  if (!kode) { if (!senyap) info_('Ruangan tidak dikenali.', 'Cetak'); return null; }
  const n = c.getRange(C_R1, 3, C_LAST - C_R1 + 1, 1).getValues().filter(function (r) { return txt_(r[0]) !== ''; }).length;
  if (n === 0) { if (!senyap) info_('Ruangan ini belum memiliki barang (belum ada saldo awal atau transaksi).', 'Cetak'); return null; }
  c.showRows(C_R1, C_LAST - C_R1 + 1);
  if (n < C_LAST - C_R1 + 1) c.hideRows(C_R1 + n, C_LAST - C_R1 + 1 - n);
  SpreadsheetApp.flush();
  try {
    const tz = Session.getScriptTimeZone();
    const nama = 'StockOpname_' + kode + '_' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
    const file = exportPdf_(c, 'A' + C_PRINT_FIRST + ':I' + C_PRINT_LAST, nama, folder_(FOLDER_CETAK));
    return { nama: txt_(c.getRange('D2').getValue()), url: file.getUrl(), n: n, id: file.getId() };
  } finally {
    c.showRows(C_R1, C_LAST - C_R1 + 1);
  }
}

function tampilkanLink_(daftar) {
  let html = '<div style="font-family:Arial;font-size:13px"><p>PDF siap. Klik tautan, lalu cetak dengan <b>Ctrl+P</b> (atau ikon printer):</p><ul>';
  daftar.forEach(function (d) { html += '<li><a href="' + d.url + '" target="_blank">' + d.nama + ' (' + d.n + ' item)</a></li>'; });
  html += '</ul><p style="color:#666">Semua PDF tersimpan di folder Drive "' + FOLDER_CETAK + '".</p></div>';
  ui_().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(460).setHeight(120 + daftar.length * 28), 'Cetak Stock Opname');
}

function CetakRuanganTerpilih() {
  wajibMenu_();
  const r = buatPdfCetak_(false);
  if (r) tampilkanLink_([r]);
}

function CetakSemuaRuangan() {
  if (!ask_('Buat PDF Stock Opname untuk SEMUA ruangan aktif (satu dokumen per ruangan)?\nRuangan tanpa barang dilewati.', 'Cetak Semua Ruangan')) return;
  const c = sh_(SH.CETAK), rr = sh_(SH.RUANG);
  const namaAsal = c.getRange('D2').getValue(), noAsal = c.getRange('D3').getValue();
  const rooms = rr.getRange(R1, 2, RM_LAST - R1 + 1, 4).getValues();   // B..E
  const hasil = [];
  try {
    c.getRange('D3').clearContent();
    rooms.forEach(function (r) {
      if (txt_(r[0]) && txt_(r[3]) !== 'Nonaktif') {
        c.getRange('D2').setValue(r[1]);
        const x = buatPdfCetak_(true);
        if (x) hasil.push(x);
      }
    });
  } finally {
    c.getRange('D2').setValue(namaAsal);
    if (noAsal !== '') c.getRange('D3').setValue(noAsal);
  }
  if (hasil.length) tampilkanLink_(hasil); else info_('Tidak ada ruangan yang memiliki barang.', 'Cetak Semua Ruangan');
}

function CetakDariForm() {
  wajibMenu_();
  const f = sh_(SH.FORM);
  if (txt_(f.getRange('I10').getValue()).indexOf('TERSIMPA') !== 0) {
    if (!ask_('Stock Opname ini belum disimpan. Simpan sekarang agar bisa dicetak?', 'Cetak')) return;
    if (!SimpanStockOpname(true)) return;
  }
  const c = sh_(SH.CETAK);
  c.getRange('D2').setValue(txt_(f.getRange('D7').getValue()));
  c.getRange('D3').setValue(txt_(f.getRange('L6').getValue()));
  CetakRuanganTerpilih();
}

// ======================================================================
// SALDO AWAL DIISI DARI MASTER_BARANG (v1.2)
// Kolom M..U di MASTER_BARANG = stok awal per ruangan (urutan sama dengan MASTER_RUANGAN).
// Kosong = barang tidak ada di ruangan itu; angka 0 = terdaftar di ruangan itu dengan stok nol.
// SALDO_AWAL diisi otomatis dari sini (jangan diisi manual).
// ======================================================================
function kodeRuangan_() {
  return sh_(SH.RUANG).getRange(R1, 2, JML_RUANG, 1).getValues().map(function (r) { return txt_(r[0]); });
}

function PasangKolomSaldo_() {
  const m = sh_(SH.MASTER), n = JML_RUANG, c1 = KOL_SALDO1, rows = M_LAST - R1 + 1;
  const P = SpreadsheetApp.CopyPasteType.PASTE_FORMAT;
  const hdr = [];
  for (let k = 0; k < n; k++) hdr.push('="Saldo "&MASTER_RUANGAN!C' + (R1 + k));
  m.getRange(5, c1, 1, n).setFormulas([hdr]);
  m.getRange(5, c1 + n).setValue('Total Saldo');
  m.getRange(5, 11).copyTo(m.getRange(5, c1, 1, n + 1), P, false);                 // gaya judul (K5)
  m.getRange(R1, 8).copyTo(m.getRange(R1, c1, rows, n), P, false);                  // gaya sel isian (H)
  m.getRange(R1, 11).copyTo(m.getRange(R1, c1 + n, rows, 1), P, false);             // gaya kolom rumus (K)
  m.getRange(R1, c1, rows, n).setHorizontalAlignment('center');
  const f = [];
  for (let i = 0; i < rows; i++) {
    const r = R1 + i;
    f.push(['=IF($B' + r + '="","",SUM(' + m.getRange(r, c1).getA1Notation() + ':' + m.getRange(r, c1 + n - 1).getA1Notation() + '))']);
  }
  m.getRange(R1, c1 + n, rows, 1).setFormulas(f);
  m.getRange(R1, c1, rows, n).setDataValidation(SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false)
    .setHelpText('Stok awal di ruangan ini (angka ≥ 0). Kosong = barang tidak ada di ruangan ini; 0 = terdaftar dengan stok nol.').build());
  m.setColumnWidths(c1, n + 1, 80);
  m.getRange('A3').setValue('ISI SEMUA DATA BARANG DI SINI: SKU, nama, foto, dan stok awal per ruangan (kolom M dst.). Sheet lain mengikuti otomatis.');
  m.getRange('A4').setValue('SKU baru: isi pada baris kosong berikutnya (kolom kuning). Stok awal: kosong = tidak ada di ruangan itu, 0 = terdaftar dengan stok nol. Perubahan stok berikutnya lewat TRANSAKSI_STOK.');
  const sa = sh_(SH.SALDO);
  sa.getRange('A3').setValue('Sheet ini DIISI OTOMATIS dari kolom Saldo di MASTER_BARANG. Jangan diedit di sini (akan ditimpa).');
}

// Pindahkan saldo manual lama (SALDO_AWAL) ke kolom master, hanya bila kolom master masih kosong.
function imporSaldoKeMaster_() {
  const m = sh_(SH.MASTER), n = JML_RUANG, c1 = KOL_SALDO1, rows = M_LAST - R1 + 1;
  if (m.getRange(R1, c1, rows, n).getValues().some(function (r) { return r.some(function (v) { return txt_(v) !== ''; }); })) return 0;
  const sa = dataStok_().saldo;
  if (!sa.length) return 0;
  const kode = kodeRuangan_(), idx = {};
  m.getRange(R1, 2, rows, 1).getValues().forEach(function (r, i) { const s = txt_(r[0]); if (s && idx[s] === undefined) idx[s] = i; });
  const grid = [];
  for (let i = 0; i < rows; i++) { const a = []; for (let k = 0; k < n; k++) a.push(''); grid.push(a); }
  let c = 0;
  sa.forEach(function (r) {
    const sku = txt_(r[0]), ru = txt_(r[3]), q = r[5], i = idx[sku], k = kode.indexOf(ru);
    if (!sku || !ru || txt_(q) === '' || isNaN(Number(q)) || i === undefined || k < 0) return;
    grid[i][k] = grid[i][k] === '' ? Number(q) : grid[i][k] + Number(q);
    c++;
  });
  if (c) m.getRange(R1, c1, rows, n).setValues(grid);
  return c;
}

// daftar baris SALDO_AWAL yang seharusnya, dihitung dari MASTER_BARANG
function saldoDariMaster_() {
  const m = sh_(SH.MASTER), n = JML_RUANG, rows = M_LAST - R1 + 1;
  const sku = m.getRange(R1, 2, rows, 1).getValues();
  const q = m.getRange(R1, KOL_SALDO1, rows, n).getValues();
  const kode = kodeRuangan_(), out = [], bad = [], lihat = {};
  for (let i = 0; i < rows; i++) {
    const s = txt_(sku[i][0]);
    if (!s) continue;
    if (lihat[s]) { bad.push('SKU ganda: ' + s + ' (baris ' + (R1 + i) + ' diabaikan)'); continue; }
    lihat[s] = true;
    for (let k = 0; k < n; k++) {
      const v = q[i][k];
      if (txt_(v) === '' || !kode[k]) continue;
      if (isNaN(Number(v)) || Number(v) < 0) { bad.push(s + ' ruang ' + kode[k] + ': "' + v + '" tidak valid'); continue; }
      out.push([s, kode[k], Number(v), 'Dari MASTER_BARANG']);
    }
  }
  return { out: out, bad: bad };
}

// HARUS dipanggil di dalam dalamKunci_()
function sinkronSaldo_() {
  const sa = sh_(SH.SALDO), total = S_LAST - R1 + 1;
  const h = saldoDariMaster_();
  if (h.out.length > total) throw new Error('Jumlah baris saldo (' + h.out.length + ') melebihi kapasitas SALDO_AWAL (' + total + ').');
  [[2, 0], [5, 1], [7, 2], [9, 3]].forEach(function (c) {
    const arr = [];
    for (let i = 0; i < total; i++) arr.push([i < h.out.length ? h.out[i][c[1]] : '']);
    const rg = sa.getRange(R1, c[0], total, 1);
    if (c[0] === 2 || c[0] === 5) rg.setNumberFormat('@');
    rg.setValues(arr);
  });
  return h;
}

function SinkronSaldoAwal() {
  wajibMenu_();
  try {
    const h = dalamKunci_(sinkronSaldo_);
    info_('Saldo Awal disinkronkan dari MASTER_BARANG: ' + h.out.length + ' baris.' + (h.bad.length ? '\n\nPerlu diperiksa:\n' + h.bad.slice(0, 10).join('\n') : ''), 'Sinkron Saldo Awal');
  } catch (e) { info_('Gagal: ' + e.message, 'Sinkron Saldo Awal'); }
}

// ======================================================================
// APLIKASI WEB UNTUK HP (v2.0)
// Deploy: Deploy > New deployment > Web app | Execute as: Me | Who has access: Anyone
// Satu pintu masuk: api(nama, args, pin). Fungsi berakhiran "_" tidak dapat dipanggil dari browser.
// ======================================================================
function ss_() {
  const a = SpreadsheetApp.getActive();
  if (a) return a;
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  throw new Error('Spreadsheet belum terpasang. Jalankan EZRIA GUDANG > Pasang / Perbarui Sistem sebagai pemilik.');
}
function wajibMenu_() {
  try { SpreadsheetApp.getUi(); } catch (e) { throw new Error('Fungsi ini hanya dapat dijalankan dari menu Google Sheets.'); }
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('App').setTitle('EZRIA Gudang')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function infoAplikasi() {
  let nama = 'Gudang';
  try { nama = txt_(sh_(SH.SET).getRange('C5').getValue()) || nama; } catch (e) { /* abaikan */ }
  return { perluPin: !!PropertiesService.getScriptProperties().getProperty('PIN_WEB'), nama: nama };
}

function api(nama, args, pin) {
  const sp = PropertiesService.getScriptProperties(), cache = CacheService.getScriptCache();
  const butuh = sp.getProperty('PIN_WEB');
  if (butuh) {
    const gagal = Number(cache.get('gagal_pin') || 0);
    if (gagal >= 8) throw new Error('Terlalu banyak PIN salah. Coba lagi dalam 15 menit.');
    if (String(pin || '') !== butuh) { cache.put('gagal_pin', String(gagal + 1), 900); throw new Error('PIN salah.'); }
  }
  const fn = API_[nama];
  if (!fn) throw new Error('Perintah tidak dikenal: ' + nama);
  return fn(args || {});
}

function AturPinWeb() {
  wajibMenu_();
  const sp = PropertiesService.getScriptProperties();
  const sekarang = sp.getProperty('PIN_WEB');
  const p = prompt_('PIN aplikasi HP (4-20 karakter). Kosongkan untuk MENONAKTIFKAN PIN.\nStatus sekarang: ' + (sekarang ? 'PIN aktif' : 'tanpa PIN'), 'PIN Aplikasi HP');
  if (p === null) return;
  if (p === '') { sp.deleteProperty('PIN_WEB'); info_('PIN dinonaktifkan. Siapa pun yang memiliki tautan dapat masuk.', 'PIN Aplikasi HP'); return; }
  if (p.length < 4 || p.length > 20) { info_('PIN harus 4 sampai 20 karakter.', 'PIN Aplikasi HP'); return; }
  sp.setProperty('PIN_WEB', p);
  info_('PIN disimpan. Bagikan PIN hanya kepada petugas.', 'PIN Aplikasi HP');
}
function TautanAplikasi() {
  wajibMenu_();
  ss_(); PropertiesService.getScriptProperties().setProperty('SS_ID', ss_().getId());
  let url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) { url = ''; }
  info_(url ? 'Tautan aplikasi HP:\n' + url + '\n\nBuka di browser HP, lalu "Tambahkan ke layar utama".' : 'Aplikasi web belum di-deploy.\nExtensions > Apps Script > Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).', 'Aplikasi HP');
}

// ---------- utilitas ----------
function tz_() { return ss_().getSpreadsheetTimeZone(); }
function tglStr_(d) { return d instanceof Date ? Utilities.formatDate(d, tz_(), 'yyyy-MM-dd') : ''; }
function parseTgl_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(s || ''))) throw new Error('Tanggal tidak valid.');
  return Utilities.parseDate(String(s), tz_(), 'yyyy-MM-dd');
}
function catat_(aksi, petugas, rincian) {
  try {
    const ss = ss_();
    let s = ss.getSheetByName('LOG_APLIKASI');
    if (!s) { s = ss.insertSheet('LOG_APLIKASI'); s.appendRow(['Waktu', 'Aksi', 'Petugas', 'Rincian']); try { s.hideSheet(); } catch (e2) { /* abaikan */ } }
    s.appendRow([new Date(), aksi, petugas, rincian]);
  } catch (e) { /* log bersifat tambahan */ }
}
function masterData_() {
  const m = sh_(SH.MASTER), last = lastNonEmpty_(m, 2, R1, M_LAST);
  if (last < R1) return [];
  const out = [];
  m.getRange(R1, 2, last - R1 + 1, 11).getValues().forEach(function (r, i) {      // B..L
    const sku = txt_(r[0]);
    if (!sku) return;
    out.push({ row: R1 + i, sku: sku, nama: txt_(r[2]), kat: txt_(r[3]), jenis: txt_(r[4]), ukuran: txt_(r[5]), sat: txt_(r[6]), ket: txt_(r[7]), status: txt_(r[8]), foto: txt_(r[9]) === 'ADA' });
  });
  return out;
}
function ruanganData_() {
  return sh_(SH.RUANG).getRange(R1, 2, JML_RUANG, 4).getValues().filter(function (r) { return txt_(r[0]) !== ''; })
    .map(function (r) { return { kode: txt_(r[0]), nama: txt_(r[1]), aktif: txt_(r[3]) !== 'Nonaktif' }; });
}
function stokPeta_(data, J) {
  const ev = bangunPeristiwa_(data, J, null), peta = {};
  Object.keys(ev).forEach(function (k) { let s = 0; ev[k].forEach(function (o) { s += o.d; }); peta[k] = s; });
  return peta;
}
function daftarPengaturan_(kolom) {
  return sh_(SH.SET).getRange(5, kolom, 50, 1).getValues().map(function (r) { return txt_(r[0]); }).filter(function (x) { return x !== ''; });
}
function tulisTrx_(t, baris, o) {
  t.getRange(baris, 2).setValue(o.tgl);
  t.getRange(baris, 3, 1, 3).setNumberFormat('@').setValues([[o.nomor, o.jenis, o.sku]]);
  t.getRange(baris, 8, 1, 2).setNumberFormat('@').setValues([[o.asal || '', o.tuj || '']]);
  t.getRange(baris, 10).setValue(o.qty);
  t.getRange(baris, 12, 1, 2).setValues([[o.ket || '', o.petugas]]);
  if (o.setuju) t.getRange(baris, 15).setValue(o.setuju);
  t.setRowHeightsForced(baris, 1, H_LIST);
}

// ---------- baca ----------
function apiBootstrap_() {
  const J = jenis_(), data = dataStok_(), peta = stokPeta_(data, J), ru = ruanganData_(), md = masterData_();
  const perRuang = {}, ringkas = { jmlSku: md.length, totalStok: 0, trx: 0, tanpaFoto: 0, negatif: 0 };
  ru.forEach(function (r) { perRuang[r.kode] = 0; });
  const barang = md.map(function (b) {
    const per = {}; let tot = 0;
    ru.forEach(function (r) { const v = peta[b.sku + '|' + r.kode]; if (v !== undefined) { per[r.kode] = v; tot += v; perRuang[r.kode] += v; if (v < 0) ringkas.negatif++; } });
    b.per = per; b.total = tot; ringkas.totalStok += tot; if (!b.foto) ringkas.tanpaFoto++;
    return b;
  });
  data.trx.forEach(function (r) { if (txt_(r[2]) !== '') ringkas.trx++; });
  const ket = { masuk: 'Barang masuk', keluar: 'Barang keluar', transfer: 'Pindah antar ruangan', adjp: 'Penyesuaian tambah', adjm: 'Penyesuaian kurang' };
  return {
    cfg: {
      nama: txt_(sh_(SH.SET).getRange('C5').getValue()),
      jenis: [['masuk', false, true], ['keluar', true, false], ['transfer', true, true], ['adjp', false, true], ['adjm', true, false]].map(function (x) { return { key: x[0], label: J[x[0]], info: ket[x[0]], asal: x[1], tuj: x[2] }; }),
      kategori: daftarPengaturan_(5), satuan: daftarPengaturan_(6)
    },
    ruang: ru, barang: barang, ringkas: ringkas, perRuang: perRuang
  };
}

function apiSkuDetail_(a) {
  const sku = txt_(a.sku), row = barisSku_(sku);
  if (!row) throw new Error('SKU tidak ditemukan.');
  const b = masterData_().filter(function (x) { return x.sku === sku; })[0];
  const kode = kodeRuangan_(), q = sh_(SH.MASTER).getRange(row, KOL_SALDO1, 1, JML_RUANG).getValues()[0], saldo = {};
  kode.forEach(function (k, i) { if (k) saldo[k] = txt_(q[i]) === '' ? '' : q[i]; });
  b.saldoAwal = saldo;
  return b;
}

function apiTransaksiTerakhir_(a) {
  const n = Math.min(Number(a.n) || 30, 100), sku = txt_(a.sku), data = dataStok_(), nama = {};
  masterData_().forEach(function (b) { nama[b.sku] = b.nama; });
  const out = [];
  for (let i = data.trx.length - 1; i >= 0 && out.length < n; i--) {
    const r = data.trx[i];
    if (txt_(r[2]) === '' && txt_(r[3]) === '') continue;
    if (sku && txt_(r[3]) !== sku) continue;
    out.push({ no: txt_(r[1]), tgl: tglStr_(r[0]), jenis: txt_(r[2]), sku: txt_(r[3]), nama: nama[txt_(r[3])] || '', asal: txt_(r[6]), tuj: txt_(r[7]), qty: r[8], ket: txt_(r[10]), petugas: txt_(r[11]), valid: txt_(r[12]) });
  }
  return out;
}

function apiFoto_(a) {
  const sku = txt_(a.sku), row = barisSku_(sku);
  if (!row) return null;
  const m = sh_(SH.MASTER);
  if (txt_(m.getRange(row, 11).getValue()) !== 'ADA') return null;
  const cache = CacheService.getScriptCache(), ck = 'foto_' + sku;
  const c = cache.get(ck);
  if (c) return c;
  const folder = folder_(FOLDER_FOTO), safe = sku.replace(/[\\\/:*?"<>|]/g, '_');
  let blob = null;
  ['jpg', 'png'].forEach(function (ext) {
    if (blob) return;
    const it = folder.getFilesByName('thumb_' + safe + '.' + ext);
    if (it.hasNext()) blob = it.next().getBlob();
  });
  if (!blob) {
    const mt = txt_(m.getRange(row, 12).getValue()).match(/id=([\w-]+)/);
    if (mt) { try { const f = DriveApp.getFileById(mt[1]); blob = f.getThumbnail() || (f.getSize() < 600000 ? f.getBlob() : null); } catch (e) { blob = null; } }
  }
  if (!blob) return null;
  const url = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  if (url.length < 95000) cache.put(ck, url, 21600);
  return url;
}

// ---------- transaksi ----------
function apiCatat_(a) {
  const J = jenis_(), jn = txt_(a.jenis), sku = txt_(a.sku), qty = Number(a.qty), petugas = txt_(a.petugas), ket = txt_(a.ket);
  const asal = txt_(a.asal), tuj = txt_(a.tuj);
  if (!petugas) throw new Error('Nama petugas wajib diisi.');
  if ([J.masuk, J.keluar, J.transfer, J.adjp, J.adjm].indexOf(jn) < 0) throw new Error('Jenis transaksi tidak valid.');
  if (!(qty > 0) || !isFinite(qty)) throw new Error('Qty harus angka lebih dari 0.');
  const tgl = parseTgl_(a.tgl), kode = ruanganData_().map(function (r) { return r.kode; });
  const perluAsal = jn === J.keluar || jn === J.transfer || jn === J.adjm, perluTuj = jn === J.masuk || jn === J.transfer || jn === J.adjp;
  if (perluAsal && kode.indexOf(asal) < 0) throw new Error('Ruangan asal belum dipilih.');
  if (perluTuj && kode.indexOf(tuj) < 0) throw new Error('Ruangan tujuan belum dipilih.');
  if (jn === J.transfer && asal === tuj) throw new Error('Ruangan asal dan tujuan tidak boleh sama.');
  const hasil = dalamKunci_(function () {
    if (!masterData_().some(function (b) { return b.sku === sku; })) throw new Error('SKU tidak ditemukan.');
    const t = sh_(SH.TRX), data = dataStok_(), baris = lastNonEmpty_(t, 4, R1, T_LAST) + 1;
    if (baris > T_LAST) throw new Error('Kapasitas transaksi penuh.');
    const rv = new Array(14).fill('');
    rv[0] = tgl; rv[2] = jn; rv[3] = sku; rv[6] = perluAsal ? asal : ''; rv[7] = perluTuj ? tuj : ''; rv[8] = qty;
    const ev = bangunPeristiwa_(data, J, null), efek = efekTrx_(rv, J);
    efek.forEach(function (e) { dorong_(ev, e[0], { t: tgl.getTime(), i: baris - R1, d: e[1] }); });
    let neg = '';
    if (jenisKeluarRuang_(jn, J)) efek.forEach(function (e) { if (e[1] < 0 && !neg) { const m = minKronologis_(ev[e[0]]); if (m < 0) neg = ketKunci_(e[0], m); } });
    let setuju = '';
    if (neg && !negatifGlobal_()) {
      if (!a.setujuNegatif) return { perluPersetujuan: true, pesan: 'Stok akan menjadi negatif: ' + neg };
      setuju = 'Disetujui via aplikasi HP oleh ' + petugas;
    }
    const nomor = nomorBatch_('TRX', [tgl])[0];
    tulisTrx_(t, baris, { tgl: tgl, nomor: nomor, jenis: jn, sku: sku, asal: rv[6], tuj: rv[7], qty: qty, ket: ket, petugas: petugas, setuju: setuju });
    return { ok: true, nomor: nomor };
  });
  if (hasil.ok) catat_('Transaksi', petugas, hasil.nomor + ' ' + jn + ' ' + sku + ' x' + qty);
  return hasil;
}

function koreksiRencana_(data, J, no) {
  let idx = -1;
  for (let i = 0; i < data.trx.length; i++) if (txt_(data.trx[i][1]) === no) { idx = i; break; }
  if (idx < 0) throw new Error('Nomor transaksi tidak ditemukan.');
  const r = data.trx[idx], jn = txt_(r[2]), asal = txt_(r[6]), tuj = txt_(r[7]);
  let jb, a2 = '', t2 = '';
  if (jn === J.masuk) { jb = J.keluar; a2 = tuj; }
  else if (jn === J.keluar) { jb = J.masuk; t2 = asal; }
  else if (jn === J.transfer) { jb = J.transfer; a2 = tuj; t2 = asal; }
  else if (jn === J.adjp) { jb = J.adjm; a2 = tuj; }
  else if (jn === J.adjm) { jb = J.adjp; t2 = asal; }
  else throw new Error('Jenis transaksi tidak dikenali.');
  const sudah = data.trx.some(function (x) { return txt_(x[10]).indexOf('KOREKSI atas ' + no) === 0; });
  const rv = new Array(14).fill('');
  rv[0] = new Date(); rv[2] = jb; rv[3] = txt_(r[3]); rv[6] = a2; rv[7] = t2; rv[8] = Number(r[8]);
  const ev = bangunPeristiwa_(data, J, null), efek = efekTrx_(rv, J);
  efek.forEach(function (e) { dorong_(ev, e[0], { t: rv[0].getTime(), i: 1e9, d: e[1] }); });
  let neg = '';
  efek.forEach(function (e) { if (e[1] < 0 && !neg) { const m = minKronologis_(ev[e[0]]); if (m < 0) neg = ketKunci_(e[0], m); } });
  return { r: r, jn: jn, jb: jb, a2: a2, t2: t2, sudah: sudah, neg: neg, no: no };
}
function apiKoreksiCek_(a) {
  const p = koreksiRencana_(dataStok_(), jenis_(), txt_(a.no));
  return { no: p.no, jenis: p.jn, sku: txt_(p.r[3]), qty: p.r[8], balik: p.jb, sudah: p.sudah, neg: p.neg };
}
function apiKoreksi_(a) {
  const J = jenis_(), alasan = txt_(a.alasan), petugas = txt_(a.petugas);
  if (!alasan) throw new Error('Alasan koreksi wajib diisi.');
  if (!petugas) throw new Error('Nama petugas wajib diisi.');
  const hasil = dalamKunci_(function () {
    const t = sh_(SH.TRX), p = koreksiRencana_(dataStok_(), J, txt_(a.no));
    if (p.sudah && !a.abaikanSudah) return { perlu: 'sudah', pesan: 'Transaksi ini SUDAH pernah dikoreksi. Koreksi lagi akan menggandakan pembalikan.' };
    let setuju = '';
    if (p.neg && !negatifGlobal_()) {
      if (!a.setujuNegatif) return { perlu: 'negatif', pesan: 'Koreksi ini membuat stok ' + p.neg + '.' };
      setuju = 'Koreksi disetujui: ' + petugas;
    }
    const baris = lastNonEmpty_(t, 4, R1, T_LAST) + 1;
    if (baris > T_LAST) throw new Error('Kapasitas transaksi penuh.');
    const tgl = new Date(), nomor = nomorBatch_('TRX', [tgl])[0];
    tulisTrx_(t, baris, { tgl: tgl, nomor: nomor, jenis: p.jb, sku: txt_(p.r[3]), asal: p.a2, tuj: p.t2, qty: p.r[8], ket: 'KOREKSI atas ' + p.no + ' - ' + alasan, petugas: petugas, setuju: setuju });
    return { ok: true, nomor: nomor };
  });
  if (hasil.ok) catat_('Koreksi', petugas, hasil.nomor + ' atas ' + a.no + ': ' + alasan);
  return hasil;
}

// ---------- stock opname ----------
function ruangDari_(kode) {
  const r = ruanganData_().filter(function (x) { return x.kode === txt_(kode); })[0];
  if (!r) throw new Error('Ruangan tidak dikenali.');
  return r;
}
function itemRuang_(kode) {
  const J = jenis_(), peta = stokPeta_(dataStok_(), J);
  return masterData_().filter(function (b) { return peta[b.sku + '|' + kode] !== undefined; })
    .map(function (b) { return { sku: b.sku, nama: b.nama, sat: b.sat, foto: b.foto, sistem: peta[b.sku + '|' + kode] }; });
}
function apiOpnameMuat_(a) {
  const ru = ruangDari_(a.kode);
  return { kode: ru.kode, nama: ru.nama, items: itemRuang_(ru.kode) };
}
function apiOpnameSimpan_(a) {
  const ru = ruangDari_(a.kode), tgl = parseTgl_(a.tgl), petugas = txt_(a.petugas);
  if (!petugas) throw new Error('Nama petugas wajib diisi.');
  if (!a.items || !a.items.length) throw new Error('Tidak ada barang untuk disimpan.');
  const hasil = dalamKunci_(function () {
    const h = sh_(SH.RIW), sekarang = itemRuang_(ru.kode), map = {};
    sekarang.forEach(function (x) { map[x.sku] = x; });
    const kirim = {};
    a.items.forEach(function (x) { kirim[txt_(x.sku)] = x; });
    if (sekarang.length !== Object.keys(kirim).length || sekarang.some(function (x) { return !kirim[x.sku]; })) return { berubah: true, pesan: 'Daftar barang di ruangan berubah sejak dimuat. Muat ulang data ruangan.' };
    const out = sekarang.map(function (x) {
      const f = kirim[x.sku].fisik;
      if (f === '' || f === null || isNaN(Number(f))) throw new Error('Stok fisik ' + x.sku + ' belum terisi/tidak valid.');
      if (Number(f) < 0) throw new Error('Stok fisik ' + x.sku + ' tidak boleh negatif.');
      return { sku: x.sku, nama: x.nama, sistem: x.sistem, fisik: Number(f), ket: txt_(kirim[x.sku].ket) };
    });
    const lastH = lastNonEmpty_(h, 1, R1, H_LAST), riw = lastH >= R1 ? h.getRange(R1, 1, lastH - R1 + 1, 3).getValues() : [];
    const hari = tglStr_(tgl), sama = {};
    riw.forEach(function (x) { if (txt_(x[2]) === ru.nama && tglStr_(x[1]) === hari) sama[txt_(x[0])] = true; });
    if (Object.keys(sama).length && !a.abaikanGanda) return { ganda: Object.keys(sama), pesan: 'Sudah ada opname ' + ru.nama + ' pada tanggal yang sama (' + Object.keys(sama).join(', ') + ').' };
    const n = out.length, baris = Math.max(lastH, R1 - 1) + 1;
    if (baris + n - 1 > H_LAST) throw new Error('Kapasitas riwayat opname tidak cukup.');
    const no = nomorBatch_('OP', [tgl])[0];
    h.getRange(baris, 1, n, 4).setNumberFormat('@');
    h.getRange(baris, 1, n, 4).setValues(out.map(function (o) { return [no, tgl, ru.nama, o.sku]; }));
    h.getRange(baris, 2, n, 1).setNumberFormat('DD/MM/YYYY');
    h.getRange(baris, 6, n, 6).setValues(out.map(function (o) { return [o.nama, o.sistem, o.fisik, o.fisik - o.sistem, o.ket, petugas]; }));
    h.setRowHeightsForced(baris, n, H_LIST);
    return { ok: true, nomor: no, item: n, selisih: out.filter(function (o) { return o.fisik !== o.sistem; }).length };
  });
  if (hasil.ok) catat_('Opname', petugas, hasil.nomor + ' ' + ru.nama + ' (' + hasil.item + ' item, ' + hasil.selisih + ' selisih)');
  return hasil;
}
function apiRiwayatOpname_() {
  const h = sh_(SH.RIW), last = lastNonEmpty_(h, 1, R1, H_LAST);
  if (last < R1) return [];
  const v = h.getRange(R1, 1, last - R1 + 1, 11).getValues(), por = {}, urut = [];
  v.forEach(function (r) {
    const no = txt_(r[0]);
    if (!no) return;
    if (!por[no]) { por[no] = { no: no, tgl: tglStr_(r[1]), ruang: txt_(r[2]), item: 0, selisih: 0, petugas: txt_(r[10]) }; urut.push(no); }
    por[no].item++; if (num_(r[8]) !== 0) por[no].selisih++;
  });
  return urut.reverse().slice(0, 40).map(function (n) { return por[n]; });
}
function apiOpnameDetail_(a) {
  const h = sh_(SH.RIW), last = lastNonEmpty_(h, 1, R1, H_LAST), no = txt_(a.no), out = [];
  if (last >= R1) h.getRange(R1, 1, last - R1 + 1, 11).getValues().forEach(function (r) {
    if (txt_(r[0]) === no) out.push({ sku: txt_(r[3]), nama: txt_(r[5]), sistem: r[6], fisik: r[7], selisih: r[8], ket: txt_(r[9]) });
  });
  return out;
}

// ---------- PDF (dikirim sebagai data, karena file ada di Drive pemilik) ----------
function apiCetak_(a) {
  const ru = ruangDari_(a.kode), nomor = txt_(a.nomor);
  const hasil = dalamKunci_(function () {
    const c = sh_(SH.CETAK);
    c.getRange('D2').setValue(ru.nama);
    if (nomor) c.getRange('D3').setValue(nomor); else c.getRange('D3').clearContent();
    SpreadsheetApp.flush();
    return buatPdfCetak_(true);
  });
  if (!hasil) throw new Error('Ruangan ' + ru.nama + ' belum memiliki barang.');
  const blob = DriveApp.getFileById(hasil.id).getBlob();
  return { id: hasil.id, nama: ru.nama, file: blob.getName(), n: hasil.n, b64: Utilities.base64Encode(blob.getBytes()) };
}
function apiCetakTautan_(a) {
  const f = DriveApp.getFileById(txt_(a.id)), p = f.getParents();
  let ok = false;
  while (p.hasNext()) if (p.next().getName() === FOLDER_CETAK) ok = true;
  if (!ok) throw new Error('File tidak berada di folder cetak.');
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/file/d/' + f.getId() + '/view';
}

// ---------- SKU & foto ----------
function apiTambahSku_(a) {
  const sku = txt_(a.sku).toUpperCase(), nama = txt_(a.nama), petugas = txt_(a.petugas);
  if (!/^[A-Z0-9._-]{1,30}$/.test(sku)) throw new Error('SKU hanya boleh huruf, angka, titik, minus, garis bawah (maks 30 karakter).');
  if (!nama) throw new Error('Nama barang wajib diisi.');
  const hasil = dalamKunci_(function () {
    const m = sh_(SH.MASTER), v = m.getRange(R1, 2, M_LAST - R1 + 1, 1).getValues();
    let r = 0, pakai = 0;
    for (let i = 0; i < v.length; i++) {
      const k = txt_(v[i][0]);
      if (k === sku) throw new Error('SKU ' + sku + ' sudah ada.');
      if (k !== '') pakai++; else if (!r) r = R1 + i;
    }
    if (!r) throw new Error('Kapasitas SKU penuh (' + (M_LAST - R1 + 1) + ').');
    m.getRange(r, 2).setNumberFormat('@').setValue(sku);
    m.getRange(r, 4, 1, 2).setValues([[nama, txt_(a.kat)]]);
    m.getRange(r, 6, 1, 3).setValues([[txt_(a.jenis), txt_(a.ukuran), txt_(a.sat) || 'Unit']]);
    m.getRange(r, 9, 1, 2).setValues([[txt_(a.ket), 'Aktif']]);
    isiSaldoMaster_(m, r, a.saldo);
    m.setRowHeightsForced(r, 1, H_MASTER);
    const h = sinkronSaldo_();
    return { row: r, pakai: pakai + 1, bad: h.bad };
  });
  catat_('Tambah SKU', petugas, sku + ' ' + nama);
  return hasil;
}
function isiSaldoMaster_(m, row, saldo) {
  if (!saldo) return;
  const kode = kodeRuangan_();
  Object.keys(saldo).forEach(function (k) {
    const i = kode.indexOf(k);
    if (i < 0) return;
    const v = saldo[k];
    if (v === '' || v === null) m.getRange(row, KOL_SALDO1 + i).clearContent();
    else if (!isNaN(Number(v)) && Number(v) >= 0) m.getRange(row, KOL_SALDO1 + i).setValue(Number(v));
  });
}
function apiEditSku_(a) {
  const sku = txt_(a.sku), petugas = txt_(a.petugas);
  const hasil = dalamKunci_(function () {
    const row = barisSku_(sku);
    if (!row) throw new Error('SKU tidak ditemukan.');
    const m = sh_(SH.MASTER);
    if (a.nama !== undefined) { if (!txt_(a.nama)) throw new Error('Nama barang wajib diisi.'); m.getRange(row, 4, 1, 2).setValues([[txt_(a.nama), txt_(a.kat)]]); }
    if (a.jenis !== undefined) m.getRange(row, 6, 1, 3).setValues([[txt_(a.jenis), txt_(a.ukuran), txt_(a.sat) || 'Unit']]);
    if (a.ket !== undefined) m.getRange(row, 9, 1, 2).setValues([[txt_(a.ket), txt_(a.status) || 'Aktif']]);
    isiSaldoMaster_(m, row, a.saldo);
    const h = sinkronSaldo_();
    return { ok: true, bad: h.bad, adaTransaksi: lastNonEmpty_(sh_(SH.TRX), 4, R1, T_LAST) >= R1 };
  });
  catat_('Ubah SKU', petugas, sku);
  return hasil;
}
function apiHapusSku_(a) {
  const sku = txt_(a.sku), petugas = txt_(a.petugas);
  if (!petugas) throw new Error('Nama petugas wajib diisi.');
  const h = hapusSkuInti_(sku);
  if (h.ok) catat_('Hapus SKU', petugas, sku + ' ' + h.nama);
  return h;
}
function apiSimpanFoto_(a) {
  const sku = txt_(a.sku), row = barisSku_(sku);
  if (!row) throw new Error('SKU tidak ditemukan.');
  const msg = simpanFotoInti_({ sku: sku, row: row, mime: 'image/jpeg', orig: a.orig, thumb: a.thumb });
  CacheService.getScriptCache().remove('foto_' + sku);
  catat_('Foto', txt_(a.petugas), sku);
  return msg;
}
function apiHapusFoto_(a) {
  const sku = txt_(a.sku), row = barisSku_(sku);
  if (!row) throw new Error('SKU tidak ditemukan.');
  const m = sh_(SH.MASTER);
  m.getRange(row, 3).clearContent(); m.getRange(row, 12).clearContent();
  CacheService.getScriptCache().remove('foto_' + sku);
  catat_('Hapus foto', txt_(a.petugas), sku);
  return true;
}

// ---------- pemeriksaan ----------
function apiValidasi_() { return dalamKunci_(validasiInti_); }
function apiKesehatan_() { return kesehatanLaporan_(); }

const API_ = {
  bootstrap: apiBootstrap_, skuDetail: apiSkuDetail_, trxTerakhir: apiTransaksiTerakhir_, foto: apiFoto_,
  catat: apiCatat_, koreksiCek: apiKoreksiCek_, koreksi: apiKoreksi_,
  opnameMuat: apiOpnameMuat_, opnameSimpan: apiOpnameSimpan_, riwayatOpname: apiRiwayatOpname_, opnameDetail: apiOpnameDetail_,
  cetak: apiCetak_, cetakTautan: apiCetakTautan_,
  tambahSku: apiTambahSku_, editSku: apiEditSku_, hapusSku: apiHapusSku_, simpanFoto: apiSimpanFoto_, hapusFoto: apiHapusFoto_,
  validasi: apiValidasi_, kesehatan: apiKesehatan_
};

function TampilkanLog() {
  wajibMenu_();
  const s = ss_().getSheetByName('LOG_APLIKASI');
  if (!s) { info_('Belum ada log. Log dibuat otomatis saat aplikasi HP dipakai untuk mengubah data.', 'Log Aplikasi'); return; }
  if (s.isSheetHidden()) { s.showSheet(); ss_().setActiveSheet(s); } else s.hideSheet();
}