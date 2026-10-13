/**
 * Bimbel Kotagede - Web App (Google Apps Script)
 * Terikat ke Spreadsheet Fase 1 (Bimbel_Kotagede_Fase1_v0.1.xlsx -> Google Sheets).
 * Deploy: Execute as "User accessing the web app". Peran dibaca dari tab user_account.
 * CATATAN: peran hanya membatasi tampilan aplikasi. Akses data sebenarnya = siapa yang
 * diberi akses (share) ke Spreadsheet.
 */
const TZ = 'Asia/Jakarta';
const SS_ = () => SpreadsheetApp.getActiveSpreadsheet();
const ROLES = {
  dash:     ['PEM', 'AL', 'ADM'],
  pantau:   ['PEM', 'AL', 'ADM', 'TUT'],
  presensi: ['PEM', 'AL', 'TUT'],
  siswa:    ['PEM', 'AL', 'ADM'],
  bayar:    ['PEM', 'ADM']
};
const STATUS_HADIR = ['hadir', 'izin', 'sakit', 'alfa'];
const METODE = ['QRIS', 'Transfer', 'Tunai'];

/* ---------- Menu & setup ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Bimbel').addItem('Setup awal (sekali saja)', 'setup').addToUi();
}

function setup() {
  const ss = SS_();
  let sh = ss.getSheetByName('user_account');
  if (!sh) {
    sh = ss.insertSheet('user_account');
    sh.appendRow(['user_id', 'email', 'role', 'tutor_id', 'status']);
    sh.getRange('A1:E1').setFontWeight('bold').setBackground('#1F4E78').setFontColor('#ffffff');
    sh.getRange('C2:C100').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['PEM', 'AL', 'ADM', 'TUT'], true).build());
    sh.getRange('E2:E100').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['aktif', 'nonaktif'], true).build());
    sh.setColumnWidth(2, 240);
  }
  const me = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  if (me && !table_('user_account').some(x => String(x.email).toLowerCase() === me)) {
    sh.appendRow(['U001', me, 'PEM', '', 'aktif']);
  }
  if (!ss.getSheetByName('audit_log')) {
    const a = ss.insertSheet('audit_log');
    a.appendRow(['waktu', 'email', 'aksi', 'objek']);
    a.getRange('A1:D1').setFontWeight('bold').setBackground('#1F4E78').setFontColor('#ffffff');
    a.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
  // kolom teks agar jam/periode tidak berubah jadi tanggal
  const cl = ss.getSheetByName('class');   if (cl) cl.getRange('G2:G201').setNumberFormat('@');
  const iv = ss.getSheetByName('invoice'); if (iv) iv.getRange('C2:C201').setNumberFormat('@');
  SpreadsheetApp.getUi().alert('Setup selesai. Email Anda terdaftar sebagai PEM di tab user_account.');
}

/* ---------- Web app ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Bimbel Kotagede')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- Helper umum ---------- */
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function pad_(n) { return ('000' + n).slice(-3); }
function conv_(v) {
  if (v instanceof Date) {
    return v.getFullYear() < 1950 ? Utilities.formatDate(v, TZ, 'HH:mm') : Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  }
  return v;
}
function parseDate_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  if (!m) throw new Error('Tanggal harus format yyyy-mm-dd.');
  return new Date(+m[1], +m[2] - 1, +m[3]);
}
function table_(name) {
  const sh = SS_().getSheetByName(name);
  if (!sh) throw new Error('Tab tidak ditemukan: ' + name + '. Sudah jalankan menu Bimbel > Setup awal?');
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return [];
  const v = sh.getRange(1, 1, lr, lc).getValues();
  const h = v[0], out = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] === '' || v[i][0] === null) continue;
    const o = {};
    h.forEach((k, j) => { if (k !== '') o[k] = conv_(v[i][j]); });
    out.push(o);
  }
  return out;
}
function idx_(rows, key) { const m = {}; rows.forEach(r => { m[r[key]] = r; }); return m; }
function nextNum_(name, re) {
  const sh = SS_().getSheetByName(name);
  const n = Math.max(sh.getLastRow() - 1, 1);
  const v = sh.getRange(2, 1, n, 1).getValues();
  let m = 0;
  v.forEach(r => { const x = re.exec(String(r[0])); if (x) m = Math.max(m, +x[1]); });
  return m + 1;
}
/** Tulis ke baris kosong pertama (kolom A kosong), hanya kolom yang ada di obj.
 *  Kolom rumus (mis. invoice D,F,G,H) dipertahankan. JANGAN pakai appendRow untuk tab berumus. */
function append_(name, obj) {
  const sh = SS_().getSheetByName(name);
  if (!sh) throw new Error('Tab tidak ditemukan: ' + name);
  const lc = sh.getLastColumn();
  const h = sh.getRange(1, 1, 1, lc).getValues()[0];
  const maxR = sh.getMaxRows();
  const colA = sh.getRange(1, 1, maxR, 1).getValues();
  let r = -1;
  for (let i = 1; i < maxR; i++) { if (colA[i][0] === '') { r = i + 1; break; } }
  if (r < 0) { sh.insertRowAfter(maxR); r = maxR + 1; }
  const rng = sh.getRange(r, 1, 1, lc);
  const f = rng.getFormulas()[0];
  const row = h.map((k, i) => {
    if (Object.prototype.hasOwnProperty.call(obj, k)) {
      const v = obj[k];
      if (typeof v === 'string' && /^(0\d|\d{4}-\d{2}$)/.test(v)) sh.getRange(r, i + 1).setNumberFormat('@');
      return v;
    }
    return f[i] || '';
  });
  rng.setValues([row]);
  return r;
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { SpreadsheetApp.flush(); lock.releaseLock(); }
}
function audit_(u, aksi, objek) {
  const sh = SS_().getSheetByName('audit_log');
  if (sh) sh.appendRow([new Date(), u.email, aksi, objek]);
}

/* ---------- Pengguna & peran ---------- */
function user_() {
  const email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) throw new Error('Email tidak terbaca. Deploy harus "Execute as: User accessing the web app" dan buka dengan akun Google.');
  const u = table_('user_account').find(x => String(x.email).toLowerCase() === email && x.status === 'aktif');
  if (!u) throw new Error('Akses ditolak untuk ' + email + '. Minta pemilik menambahkan email ini di tab user_account.');
  return { email: email, role: u.role, tutor_id: u.tutor_id || '' };
}
function need_(u, roles) { if (roles.indexOf(u.role) < 0) throw new Error('Peran ' + u.role + ' tidak punya akses.'); }
function classAllowed_(u, cid) {
  const c = table_('class').find(x => x.class_id === cid);
  if (!c) throw new Error('Kelas tidak ditemukan.');
  if (u.role === 'TUT' && u.tutor_id && c.tutor_id !== u.tutor_id) throw new Error('Bukan kelas Anda.');
  return c;
}
function classLabels_() {
  const sj = idx_(table_('subject'), 'subject_id'), tt = idx_(table_('tutor'), 'tutor_id'), m = {};
  table_('class').forEach(c => {
    m[c.class_id] = ((sj[c.subject_id] || {}).nama || '?') + ' T' + c.tier + ' · ' + c.hari + ' ' + c.jam + ' · ' + ((tt[c.tutor_id] || {}).nama || '?');
  });
  return m;
}

/* ---------- API (dipanggil dari Index.html) ---------- */
function boot() {
  const u = user_();
  const tabs = Object.keys(ROLES).filter(k => ROLES[k].indexOf(u.role) >= 0);
  return { user: u, tabs: tabs, meta: meta_(u), today: today_() };
}
function getMeta() { return meta_(user_()); }
function meta_(u) {
  const cl = classLabels_();
  return {
    classes: table_('class')
      .filter(c => !(u.role === 'TUT' && u.tutor_id && c.tutor_id !== u.tutor_id))
      .map(c => ({ class_id: c.class_id, label: cl[c.class_id] })),
    programs: table_('program').map(p => ({ program_id: p.program_id, label: p.nama + ' (Rp ' + Number(p.harga).toLocaleString('id-ID') + ')' })),
    parents: table_('parent').map(p => ({ parent_id: p.parent_id, nama: p.nama })),
    students: table_('student').map(s => ({ student_id: s.student_id, nama: s.nama })),
    errors: table_('Kode_Kesalahan').map(e => ({ code: e.error_code, label: e.error_code + ' ' + e.deskripsi }))
  };
}

function getDashboard() {
  const u = user_(); need_(u, ROLES.dash);
  SpreadsheetApp.flush();
  const rows = SS_().getSheetByName('Dashboard').getRange('A4:C19').getDisplayValues();
  const kelas = table_('Rekap_Kelas').map(k => ({
    class_id: k.class_id, mapel: k.mapel, tier: k.tier, tutor: k.tutor, kapasitas: k.kapasitas,
    terdaftar: k.terdaftar, okupansi: k.okupansi === '' ? '' : Math.round(Number(k.okupansi) * 100) + '%', catatan: k.catatan
  }));
  return { rows: rows, kelas: kelas };
}

function getPantau() {
  const u = user_(); need_(u, ROLES.pantau);
  SpreadsheetApp.flush();
  let allowed = null;
  if (u.role === 'TUT' && u.tutor_id) {
    const cls = {};
    table_('class').forEach(c => { if (c.tutor_id === u.tutor_id) cls[c.class_id] = 1; });
    allowed = {};
    table_('enrollment').forEach(e => { if (cls[e.class_id] && e.status === 'aktif') allowed[e.student_id] = 1; });
  }
  return table_('Rekap_Siswa').filter(r => r.PANTAU === 'PANTAU' && (!allowed || allowed[r.student_id])).map(r => {
    const why = [];
    if (Number(r.absen) >= 2) why.push('absen ' + r.absen + '×');
    if (r.PR_selesai !== '' && Number(r.PR_selesai) < 0.6) why.push('PR ' + Math.round(Number(r.PR_selesai) * 100) + '%');
    if (Number(r['jml_exit_≤1']) >= 2) why.push('exit ticket ≤1 ' + r['jml_exit_≤1'] + '×');
    return { student_id: r.student_id, nama: r.nama, alasan: why.join(', ') };
  });
}

function getRoster(class_id) {
  const u = user_(); need_(u, ROLES.presensi); classAllowed_(u, class_id);
  const st = idx_(table_('student'), 'student_id');
  return table_('enrollment').filter(e => e.class_id === class_id && e.status === 'aktif')
    .map(e => ({ student_id: e.student_id, nama: (st[e.student_id] || {}).nama || e.student_id }));
}

function saveSession(p) {
  const u = user_(); need_(u, ROLES.presensi); classAllowed_(u, p.class_id);
  const date = parseDate_(p.tanggal);
  if (!p.rows || !p.rows.length) throw new Error('Tidak ada siswa di kelas ini.');
  return withLock_(() => {
    if (table_('session').some(s => s.class_id === p.class_id && s.tanggal === p.tanggal)) {
      throw new Error('Sesi kelas ini pada tanggal tersebut sudah tercatat.');
    }
    const roster = {};
    table_('enrollment').forEach(e => { if (e.class_id === p.class_id && e.status === 'aktif') roster[e.student_id] = 1; });
    const codes = {};
    table_('Kode_Kesalahan').forEach(e => { codes[e.error_code] = 1; });
    p.rows.forEach(r => {
      if (!roster[r.student_id]) throw new Error('Siswa ' + r.student_id + ' tidak terdaftar di kelas ini.');
      if (STATUS_HADIR.indexOf(r.status) < 0) throw new Error('Status presensi tidak valid.');
      if (r.status === 'hadir') {
        const x = Number(r.exit), k = Number(r.libat);
        if (!(x >= 0 && x <= 3)) throw new Error('Exit ticket harus 0-3.');
        if (!(k >= 1 && k <= 3)) throw new Error('Keterlibatan harus 1-3.');
        if (r.pr !== 'Y' && r.pr !== 'N') throw new Error('PR harus Y atau N.');
        if (r.error && !codes[r.error]) throw new Error('Kode kesalahan tidak dikenal: ' + r.error);
      }
    });
    const sid = 'SE' + pad_(nextNum_('session', /^SE(\d+)$/));
    append_('session', { session_id: sid, class_id: p.class_id, tanggal: date, ringkasan: String(p.ringkasan || '').slice(0, 200) });
    let ln = nextNum_('session_log', /^LG(\d+)$/), hadir = 0;
    p.rows.forEach(r => {
      append_('attendance', { session_id: sid, student_id: r.student_id, status: r.status, catatan: '' });
      if (r.status === 'hadir') {
        hadir++;
        append_('session_log', {
          log_id: 'LG' + pad_(ln++), session_id: sid, student_id: r.student_id,
          skor_exit_ticket: Number(r.exit), error_code: r.error || '', pr_status: r.pr,
          keterlibatan: Number(r.libat), catatan: String(r.catatan || '').slice(0, 200)
        });
      }
    });
    audit_(u, 'simpan_sesi', sid);
    return { session_id: sid, hadir: hadir, total: p.rows.length };
  });
}

function getStudents() {
  const u = user_(); need_(u, ROLES.siswa);
  const pr = idx_(table_('parent'), 'parent_id');
  return table_('student').map(s => ({
    student_id: s.student_id, nama: s.nama, sekolah: s.sekolah, kelas_sekolah: s.kelas_sekolah, status: s.status,
    ortu: (pr[s.parent_id] || {}).nama || '', wa: (pr[s.parent_id] || {}).whatsapp || ''
  }));
}

function addStudent(p) {
  const u = user_(); need_(u, ['PEM', 'ADM']);
  const s = p.student || {};
  if (!String(s.nama || '').trim()) throw new Error('Nama siswa wajib diisi.');
  return withLock_(() => {
    let pid = p.parent_id, note = '';
    if (!pid) {
      const pr = p.parent || {};
      if (!String(pr.nama || '').trim()) throw new Error('Nama orang tua wajib diisi.');
      const wa = String(pr.whatsapp || '').trim();
      const ex = wa ? table_('parent').find(x => String(x.whatsapp) === wa) : null;
      if (ex) { pid = ex.parent_id; note = ' (ortu sudah ada: ' + ex.nama + ')'; }
      else {
        pid = 'P' + pad_(nextNum_('parent', /^P(\d+)$/));
        append_('parent', { parent_id: pid, nama: pr.nama.trim(), whatsapp: wa, email: pr.email || '', relasi: pr.relasi || '', alamat: pr.alamat || '' });
      }
    }
    const sid = 'S' + pad_(nextNum_('student', /^S(\d+)$/));
    append_('student', {
      student_id: sid, parent_id: pid, nama: s.nama.trim(),
      tgl_lahir: s.tgl_lahir ? parseDate_(s.tgl_lahir) : '',
      sekolah: s.sekolah || '', kelas_sekolah: s.kelas_sekolah ? Number(s.kelas_sekolah) : '',
      status: s.status || 'aktif', tgl_daftar: parseDate_(today_()), sumber: s.sumber || ''
    });
    let en = nextNum_('enrollment', /^E(\d+)$/), n = 0;
    (p.enroll || []).forEach(e => {
      if (!e.class_id || !e.program_id) return;
      append_('enrollment', { enrollment_id: 'E' + pad_(en++), student_id: sid, class_id: e.class_id, program_id: e.program_id, mulai: parseDate_(today_()), selesai: '', status: 'aktif' });
      n++;
    });
    audit_(u, 'tambah_siswa', sid);
    return { student_id: sid, parent_id: pid, kelas: n, note: note };
  });
}

function addEnrollment(student_id, class_id, program_id) {
  const u = user_(); need_(u, ['PEM', 'ADM']);
  return withLock_(() => {
    if (!table_('student').some(s => s.student_id === student_id)) throw new Error('Siswa tidak ditemukan.');
    if (!table_('class').some(c => c.class_id === class_id)) throw new Error('Kelas tidak ditemukan.');
    if (!table_('program').some(p => p.program_id === program_id)) throw new Error('Program tidak ditemukan.');
    if (table_('enrollment').some(e => e.student_id === student_id && e.class_id === class_id && e.status === 'aktif')) {
      throw new Error('Siswa sudah aktif di kelas ini.');
    }
    const eid = 'E' + pad_(nextNum_('enrollment', /^E(\d+)$/));
    append_('enrollment', { enrollment_id: eid, student_id: student_id, class_id: class_id, program_id: program_id, mulai: parseDate_(today_()), selesai: '', status: 'aktif' });
    audit_(u, 'tambah_enrollment', eid);
    return { enrollment_id: eid };
  });
}

function generateInvoices(periode, jatuh_tempo) {
  const u = user_(); need_(u, ROLES.bayar);
  if (!/^\d{4}-\d{2}$/.test(String(periode))) throw new Error('Periode harus yyyy-mm.');
  const jt = parseDate_(jatuh_tempo);
  return withLock_(() => {
    const inv = table_('invoice');
    let n = nextNum_('invoice', /^INV(\d+)$/), made = 0;
    table_('enrollment').filter(e => e.status === 'aktif').forEach(e => {
      if (inv.some(i => i.enrollment_id === e.enrollment_id && String(i.periode) === periode)) return;
      append_('invoice', { invoice_id: 'INV' + pad_(n++), enrollment_id: e.enrollment_id, periode: periode, jatuh_tempo: jt });
      made++;
    });
    audit_(u, 'buat_tagihan', periode + ' x' + made);
    return { dibuat: made };
  });
}

function getInvoices(onlyOpen) {
  const u = user_(); need_(u, ROLES.bayar);
  SpreadsheetApp.flush();
  const en = idx_(table_('enrollment'), 'enrollment_id'), st = idx_(table_('student'), 'student_id'), cl = classLabels_();
  return table_('invoice').filter(i => !onlyOpen || i.status !== 'Lunas').map(i => {
    const e = en[i.enrollment_id] || {};
    return {
      invoice_id: i.invoice_id, periode: String(i.periode), jumlah: Number(i.jumlah) || 0, dibayar: Number(i.dibayar) || 0,
      sisa: (Number(i.jumlah) || 0) - (Number(i.dibayar) || 0), status: i.status, hari_telat: Number(i.hari_telat) || 0,
      jatuh_tempo: i.jatuh_tempo, siswa: (st[e.student_id] || {}).nama || '?', kelas: cl[e.class_id] || ''
    };
  }).sort((a, b) => b.hari_telat - a.hari_telat);
}

function addPayment(p) {
  const u = user_(); need_(u, ROLES.bayar);
  const jml = Number(p.jumlah);
  if (!(jml > 0)) throw new Error('Jumlah harus lebih dari 0.');
  if (METODE.indexOf(p.metode) < 0) throw new Error('Metode tidak valid.');
  const tgl = parseDate_(p.tgl);
  return withLock_(() => {
    SpreadsheetApp.flush();
    const i = table_('invoice').find(x => x.invoice_id === p.invoice_id);
    if (!i) throw new Error('Invoice tidak ditemukan.');
    const sisa = (Number(i.jumlah) || 0) - (Number(i.dibayar) || 0);
    if (jml > sisa) throw new Error('Melebihi sisa tagihan (Rp ' + sisa.toLocaleString('id-ID') + ').');
    const pid = 'PAY' + pad_(nextNum_('payment', /^PAY(\d+)$/));
    append_('payment', { payment_id: pid, invoice_id: p.invoice_id, jumlah: jml, tgl: tgl, metode: p.metode });
    audit_(u, 'catat_bayar', pid + ' ' + p.invoice_id + ' ' + jml);
    return { payment_id: pid, sisa_baru: sisa - jml };
  });
}
