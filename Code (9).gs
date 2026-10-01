/**
 * Code.gs — Backend ระบบคลังภาพ โรงเรียนกุศลวิทยา (Google Apps Script)
 * ข้อมูลเก็บใน Google Sheets, รูปเก็บใน Google Drive
 *
 * วิธีติดตั้ง
 *  1) script.google.com > โปรเจกต์ใหม่ > วางไฟล์นี้แทน Code.gs
 *  2) เลือกฟังก์ชัน setup > Run (อนุญาตสิทธิ์) — สร้างชีตและโฟลเดอร์รูปให้อัตโนมัติ
 *  3) แก้ค่าใน createAdmin() แล้ว Run ครั้งเดียว จากนั้นลบรหัสผ่านออกจากโค้ด
 *  4) Deploy > New deployment > Web app > Execute as: Me / Who has access: Anyone
 *  5) คัดลอก URL (ลงท้าย /exec) ไปใส่ที่ตัวแปร API_URL ในไฟล์ HTML
 */
const CFG = {
  REQUIRE_APPROVAL: true,            // สมัครใหม่ต้องให้ครูเปลี่ยน status เป็น active ในชีต Users
  SESSION_SEC: 6 * 3600,             // อายุการเข้าสู่ระบบ (สูงสุด 6 ชม.)
  MAX_FAILS: 5,                      // กรอกรหัสผิดได้กี่ครั้งก่อนล็อก
  LOCK_SEC: 600,                     // ล็อกกี่วินาที
  HASH_ROUNDS: 1000,
  MAX_B64: 14e6,                     // ~10 MB ต่อรูป
  SELF_ROLES: ['นักเรียน', 'ครูและบุคลากร', 'ศิษย์เก่า'],
  ADMIN_ROLE: 'ครูผู้ดูแลระบบ',
  CATEGORIES: ['กิจกรรมวิชาการ', 'กีฬาสี', 'คุณธรรมและศาสนา', 'ภาพบรรยากาศโรงเรียน'],
  PRIVACY: ['สาธารณะ', 'เฉพาะภายใน', 'ส่วนตัว'],
  MIME: ['image/jpeg', 'image/png', 'image/webp']
};
const SHEETS = {
  Users: ['id', 'name', 'role', 'salt', 'hash', 'status', 'created'],
  Albums: ['id', 'title', 'category', 'description', 'ownerId', 'author', 'role', 'privacy', 'date', 'created'],
  Photos: ['albumId', 'fileId', 'uploaderId', 'created']
};
const TH_MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

/* ---------- ติดตั้งครั้งแรก ---------- */
function setup() {
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('SS_ID')) { Logger.log('ตั้งค่าไว้แล้ว'); return; }
  const ss = SpreadsheetApp.create('Kuson Gallery DB');
  Object.keys(SHEETS).forEach((n, i) => {
    const sh = i === 0 ? ss.getSheets()[0].setName(n) : ss.insertSheet(n);
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setNumberFormat('@'); // เก็บเป็นข้อความ กันเลขศูนย์หาย/สูตรแทรก
    sh.appendRow(SHEETS[n]);
    sh.setFrozenRows(1);
  });
  p.setProperty('SS_ID', ss.getId());
  p.setProperty('FOLDER_ID', DriveApp.createFolder('Kuson Gallery Photos').getId());
  Logger.log('สร้างฐานข้อมูลแล้ว: ' + ss.getUrl());
}

/** ใช้แทน setup() เมื่อมีสเปรดชีตอยู่แล้ว (เช่นนำเข้าจาก kuson_gallery_db.xlsx) */
function connectExisting() {
  const SS_ID = 'วาง ID สเปรดชีตที่นี่';
  const FOLDER_ID = ''; // เว้นว่าง = สร้างโฟลเดอร์รูปใหม่
  const ss = SpreadsheetApp.openById(SS_ID);
  Object.keys(SHEETS).forEach(n => {
    const sh = ss.getSheetByName(n);
    if (!sh) throw new Error('ไม่พบชีต ' + n);
    const head = sh.getRange(1, 1, 1, SHEETS[n].length).getValues()[0].join('|');
    if (head !== SHEETS[n].join('|')) throw new Error('หัวคอลัมน์ชีต ' + n + ' ต้องเป็น: ' + SHEETS[n].join(', '));
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setNumberFormat('@');
  });
  const p = PropertiesService.getScriptProperties();
  p.setProperty('SS_ID', SS_ID);
  p.setProperty('FOLDER_ID', FOLDER_ID || p.getProperty('FOLDER_ID') || DriveApp.createFolder('Kuson Gallery Photos').getId());
  Logger.log('เชื่อมต่อฐานข้อมูลแล้ว');
}

function createAdmin() {
  const id = 'teacher@kuson.ac.th', name = 'ครูสมชาย ใจดี', password = 'CHANGE_ME_BEFORE_RUN';
  if (password === 'CHANGE_ME_BEFORE_RUN' || password.length < 8) throw new Error('กรุณาตั้งรหัสผ่านอย่างน้อย 8 ตัวก่อนรัน');
  addUser(id, name, CFG.ADMIN_ROLE, password, 'active');
  Logger.log('สร้างผู้ดูแลระบบแล้ว — อย่าลืมลบรหัสผ่านออกจากโค้ด');
}

/* ---------- Web app entry ---------- */
function handle(d) {
  try {
    if (!Object.prototype.hasOwnProperty.call(H, d.action)) throw new Error('คำสั่งไม่ถูกต้อง');
    return Object.assign({ ok: true }, H[d.action](d));
  } catch (err) {
    return { ok: false, error: String(err.message || err) };
  }
}
/** เรียกจากหน้าเว็บผ่าน google.script.run (โหมดเสิร์ฟหน้าเว็บจาก Apps Script) */
function api(json) { return JSON.stringify(handle(JSON.parse(json))); }
/** โหมดเรียกผ่าน fetch จากเว็บภายนอก (ต้องตั้งสิทธิ์เป็น Anyone) */
function doPost(e) {
  let out;
  try { out = handle(JSON.parse(e.postData.contents)); } catch (err) { out = { ok: false, error: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
/** เปิด URL /exec จะได้หน้าเว็บ (ไฟล์ index.html ในโปรเจกต์) */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบคลังภาพ โรงเรียนกุศลวิทยา')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- Handlers ---------- */
const H = {
  register(d) {
    const id = String(d.id || '').trim();
    if (!/^[A-Za-z0-9._@-]{3,64}$/.test(id)) throw new Error('รหัสประจำตัว/อีเมลใช้ได้เฉพาะ a-z 0-9 . _ @ - (3-64 ตัว)');
    const name = clean(d.name, 100);
    if (!name) throw new Error('กรุณากรอกชื่อ-นามสกุล');
    if (CFG.SELF_ROLES.indexOf(d.role) < 0) throw new Error('บทบาทไม่ถูกต้อง');
    if (String(d.password || '').length < 6) throw new Error('รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร');
    return withLock(() => {
      if (findUser(id)) throw new Error('รหัสประจำตัว/อีเมลนี้ถูกใช้งานแล้ว');
      addUser(id, name, d.role, String(d.password), CFG.REQUIRE_APPROVAL ? 'pending' : 'active');
      return { message: CFG.REQUIRE_APPROVAL ? 'สมัครสำเร็จ รอครูผู้ดูแลระบบอนุมัติบัญชี' : 'สมัครสมาชิกสำเร็จ กรุณาเข้าสู่ระบบ' };
    });
  },

  login(d) {
    const id = String(d.id || '').trim();
    const cache = CacheService.getScriptCache(), key = 'f_' + id.toLowerCase().slice(0, 64);
    const fails = Number(cache.get(key) || 0);
    if (fails >= CFG.MAX_FAILS) throw new Error('กรอกรหัสผิดหลายครั้ง กรุณารอ 10 นาทีแล้วลองใหม่');
    const u = findUser(id);
    if (!u || hashPw(String(d.password || ''), u.salt) !== u.hash) {
      cache.put(key, String(fails + 1), CFG.LOCK_SEC);
      throw new Error('รหัสประจำตัวหรือรหัสผ่านไม่ถูกต้อง');
    }
    if (u.status !== 'active') throw new Error('บัญชียังไม่ได้รับการอนุมัติ กรุณาติดต่อครูผู้ดูแลระบบ');
    cache.remove(key);
    const user = { id: String(u.id), name: u.name, role: u.role };
    const token = Utilities.getUuid() + Utilities.getUuid();
    cache.put('s_' + token, JSON.stringify(user), CFG.SESSION_SEC);
    return { token, user };
  },

  logout(d) { CacheService.getScriptCache().remove('s_' + d.token); return {}; },

  list(d) {
    const me = session(d.token), admin = me.role === CFG.ADMIN_ROLE;
    const photos = {};
    readAll('Photos').forEach(p => (photos[p.albumId] = photos[p.albumId] || []).push(String(p.fileId)));
    const albums = readAll('Albums')
      .filter(a => a.privacy !== 'ส่วนตัว' || String(a.ownerId) === me.id || admin) // บังคับสิทธิ์ที่เซิร์ฟเวอร์
      .reverse()
      .map(a => ({
        id: a.id, title: a.title, category: a.category, description: a.description,
        ownerId: String(a.ownerId), author: a.author, role: a.role, privacy: a.privacy,
        date: a.date, photos: photos[a.id] || []
      }));
    return { albums };
  },

  createAlbum(d) {
    const me = session(d.token);
    const title = clean(d.title, 150);
    if (!title) throw new Error('กรุณากรอกชื่ออัลบั้ม');
    if (CFG.CATEGORIES.indexOf(d.category) < 0 || CFG.PRIVACY.indexOf(d.privacy) < 0) throw new Error('ข้อมูลไม่ถูกต้อง');
    const id = 'alb-' + Utilities.getUuid().slice(0, 8);
    withLock(() => db().getSheetByName('Albums').appendRow([
      id, title, d.category, clean(d.description, 500), me.id, me.name, me.role, d.privacy, thaiDate(), new Date().toISOString()
    ]));
    return { albumId: id };
  },

  addPhoto(d) {
    const me = session(d.token);
    if (CFG.MIME.indexOf(d.mime) < 0) throw new Error('รองรับเฉพาะ JPG, PNG, WEBP');
    const b64 = String(d.data || '');
    if (!b64 || b64.length > CFG.MAX_B64) throw new Error('ไฟล์ว่างหรือใหญ่เกินไป');
    const a = readAll('Albums').find(x => x.id === d.albumId);
    if (!a) throw new Error('ไม่พบอัลบั้ม');
    if (String(a.ownerId) !== me.id && me.role !== CFG.ADMIN_ROLE) throw new Error('คุณไม่มีสิทธิ์เพิ่มรูปในอัลบั้มนี้');
    const name = clean(d.name, 80).replace(/[^\w.\-ก-๙ ]/g, '_') || 'photo.jpg';
    const blob = Utilities.newBlob(Utilities.base64Decode(b64), d.mime, name);
    const file = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID')).createFile(blob);
    // หมายเหตุ: รูปเปิดได้ด้วยลิงก์ (รู้ ID เท่านั้น) เพื่อให้แสดงในหน้าเว็บได้
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    withLock(() => db().getSheetByName('Photos').appendRow([a.id, file.getId(), me.id, new Date().toISOString()]));
    return { fileId: file.getId() };
  }
};

/* ---------- Helpers ---------- */
function db() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SS_ID')); }
function readAll(name) {
  const v = db().getSheetByName(name).getDataRange().getValues();
  const h = v.shift();
  return v.map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])));
}
function findUser(id) { return readAll('Users').find(u => String(u.id).toLowerCase() === String(id).toLowerCase()); }
function addUser(id, name, role, password, status) {
  const salt = Utilities.getUuid();
  withLock(() => db().getSheetByName('Users').appendRow([id, name, role, salt, hashPw(password, salt), status, new Date().toISOString()]));
}
function hashPw(pw, salt) {
  let d = Utilities.newBlob(salt + pw).getBytes();
  for (let i = 0; i < CFG.HASH_ROUNDS; i++) d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, d);
  return d.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function session(token) {
  const raw = token && CacheService.getScriptCache().get('s_' + token);
  if (!raw) throw new Error('SESSION');
  return JSON.parse(raw);
}
function clean(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max); }
function thaiDate() {
  const n = new Date(), p = k => Number(Utilities.formatDate(n, 'Asia/Bangkok', k));
  return p('d') + ' ' + TH_MONTHS[p('M') - 1] + ' ' + (p('yyyy') + 543);
}
function withLock(fn) {
  const l = LockService.getScriptLock();
  l.waitLock(20000);
  try { return fn(); } finally { l.releaseLock(); }
}
