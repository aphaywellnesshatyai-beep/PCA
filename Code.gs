// ==========================================================
// อพัย เวลเนส — Code.gs (v5: ข้อมูลคลินิก/แพทย์/โลโก้ ตั้งค่าได้จากหน้าแอดมิน)
// ใช้ชีตที่ผูกกับสคริปต์นี้โดยตรง (ไม่ต้องใส่ลิงก์ชีตในโค้ด)
// ฟังก์ชันของใบรับรองแพทย์/ใบเสร็จ/ห้องยา อยู่ในไฟล์ Clinic.gs
// ==========================================================
function getSS() {
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  throw new Error('ไม่พบสเปรดชีต — ต้องเปิด Apps Script จากเมนู "ส่วนขยาย > Apps Script" ของชีต แล้วรัน setupSystem 1 ครั้ง');
}
function include(name) { return HtmlService.createTemplateFromFile(name).evaluate().getContent(); }
function getAppUrl() { return ScriptApp.getService().getUrl(); }

/* ---------- ตัวช่วยกลาง ---------- */
function pzNorm_(v) {
  var s = String(v == null ? "" : v).replace(/'/g, "").trim();
  return (s !== "" && !isNaN(s)) ? s.padStart(4, '0') : s;
}
function pzDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, "GMT+7", "yyyy-MM-dd");
  return String(v == null ? "" : v).replace(/'/g, "").trim();
}
function pzThumb_(fileId) { return "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1000"; }
function pzIdFromUrl_(url) { var m = String(url || "").match(/[-\w]{25,}/); return m ? m[0] : ""; }
function isActive_(v) { return !(v === false || String(v).trim().toUpperCase() === 'FALSE'); }
function S_(v) { return String(v == null ? "" : v).trim(); }

/** วันเกิด → {y,m,d} รองรับ Date / yyyy-mm-dd / dd/mm/yyyy / ปี พ.ศ. */
function pzDobParts_(v) {
  var p = null, m;
  if (v instanceof Date) p = { y: +Utilities.formatDate(v, "GMT+7", "yyyy"), m: +Utilities.formatDate(v, "GMT+7", "M"), d: +Utilities.formatDate(v, "GMT+7", "d") };
  else {
    var s = String(v || "").replace(/'/g, "").trim();
    if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) p = { y: +m[1], m: +m[2], d: +m[3] };
    else if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) p = { y: +m[3], m: +m[2], d: +m[1] };
  }
  if (!p) return null;
  if (p.y > 2400) p.y -= 543;
  return p;
}
function pzDobIso_(v) {
  var p = pzDobParts_(v);
  return p ? p.y + "-" + String(p.m).padStart(2, '0') + "-" + String(p.d).padStart(2, '0') : "";
}

// ==========================================================
// ⚙️ การตั้งค่า (ชีต Setting / Links)
// ==========================================================
var SET_KEYS = {
  postureFolder: "Tracking / Posture",
  signatureFolder: "Signature / ลายเซ็น",
  receiptFolder: "โฟลเดอร์เก็บใบเสร็จ",
  certFolder: "โฟลเดอร์เก็บใบรับรองแพทย์"
};

function settingSheet_() {
  var ss = getSS(), sh = ss.getSheetByName("Setting");
  if (!sh) { sh = ss.insertSheet("Setting"); sh.appendRow(["ชื่อการตั้งค่า", "ค่า", "คำอธิบาย"]); }
  return sh;
}
function linksSheet_() {
  var ss = getSS(), sh = ss.getSheetByName("Links");
  if (!sh) { sh = ss.insertSheet("Links"); sh.appendRow(["ชื่อลิงก์", "URL", "หมายเหตุ"]); }
  return sh;
}
function readSettings_() {
  var out = {};
  try {
    var sh = getSS().getSheetByName("Setting");
    if (!sh || sh.getLastRow() < 1) return out;
    sh.getRange(1, 1, sh.getLastRow(), 2).getValues().forEach(function (r) {
      var k = S_(r[0]); if (k) out[k] = S_(r[1]);
    });
  } catch (e) {}
  return out;
}
function writeSetting_(key, val, desc) {
  var sh = settingSheet_(), d = sh.getDataRange().getValues();
  for (var i = 0; i < d.length; i++) {
    if (S_(d[i][0]) === key) { sh.getRange(i + 1, 2).setValue(val); return; }
  }
  sh.appendRow([key, val, desc || ""]);
}
/** เข้ากันได้กับโค้ดเดิม */
function getAppConfigs() {
  var s = readSettings_();
  return {
    trackingFolderId: s[SET_KEYS.postureFolder] || "", signatureFolderId: s[SET_KEYS.signatureFolder] || "",
    receiptFolderId: s[SET_KEYS.receiptFolder] || "", certFolderId: s[SET_KEYS.certFolder] || ""
  };
}
function pzFolder_(cfgVal, fallbackName) {
  var id = pzIdFromUrl_(cfgVal), f = null;
  if (id) { try { f = DriveApp.getFolderById(id); } catch (e) { f = null; } }
  if (!f) {
    var it = DriveApp.getFoldersByName(fallbackName);
    f = it.hasNext() ? it.next() : DriveApp.createFolder(fallbackName);
  }
  return f;
}
/** ลิงก์ที่เกี่ยวข้อง (แสดงที่ปุ่ม "ลิงก์" บนหัวเว็บ) */
function getLinks() {
  try {
    var sh = getSS().getSheetByName("Links");
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
      .filter(function (r) { return /^https?:\/\//i.test(S_(r[1])); })
      .map(function (r) { return { name: S_(r[0]) || S_(r[1]), url: S_(r[1]), note: S_(r[2]) }; });
  } catch (e) { return []; }
}

// ==========================================================
// 🏥 ข้อมูลคลินิก / ชื่อเว็บ / โลโก้ (แอดมินตั้งค่าได้ — เก็บในชีต Setting)
// ==========================================================
var BRAND_DEFAULTS = {
  name: "อพัย เวลเนส", subtitle: "ระบบเวชระเบียนคลินิกการแพทย์แผนไทย", crest: "อ",
  title: "อพัย เวลเนส คลินิกการแพทย์แผนไทย",
  clinicFull: "อพัย เวลเนส คลินิกการแพทย์แผนไทย", clinicType: "คลินิกการแพทย์แผนไทย",
  address: "1326 ถ.กาญจนวณิชย์ ต.หาดใหญ่ อ.หาดใหญ่ จ.สงขลา 90110",
  phone: "", email: "", taxId: "", licenseNo: "90108000769",
  receiver: "", receiptNote: "",
  logo: "https://drive.google.com/file/d/1Q-zlnmDdVkqR9vKPVz5u0QxbeU3thLUd/view",
  loginTitle: "เข้าสู่ระบบ", loginSub: "กรุณาเข้าสู่ระบบเพื่อใช้งานระบบเวชระเบียน",
  colors: { primary: "#1f4d3f", gold: "#b08d3c", bg: "#f4f2ed" },
  nav: { Register: "ลงทะเบียน", Search: "ค้นหา", Index: "บันทึกรักษา", Pharmacy: "จ่ายยา", tracking: "สรีระ",
         element: "ธาตุเจ้าเรือน", History: "ประวัติ", Admin: "ตั้งค่า" }
};
var BRAND_KEYS = {
  name: "ชื่อเว็บ (หัวเว็บและหน้าเข้าสู่ระบบ)", subtitle: "คำอธิบายใต้ชื่อ", crest: "อักษรย่อในโลโก้",
  title: "ชื่อแท็บเบราว์เซอร์", clinicFull: "ชื่อเต็มคลินิก (เอกสารพิมพ์)", phone: "เบอร์โทรคลินิก (เอกสารพิมพ์)",
  loginTitle: "หัวข้อหน้าเข้าสู่ระบบ", loginSub: "ข้อความหน้าเข้าสู่ระบบ",
  clinicType: "ประเภทคลินิก (บรรทัดที่ 2 หัวเอกสาร)", address: "ที่ตั้งคลินิก", email: "อีเมลคลินิก",
  taxId: "เลขประจำตัวผู้เสียภาษี", licenseNo: "เลขที่ใบอนุญาตสถานพยาบาล",
  receiver: "ชื่อผู้รับเงิน (ใบเสร็จ)", receiptNote: "หมายเหตุท้ายใบเสร็จ", logo: "ลิงก์โลโก้คลินิก"
};
var BRAND_LIM = { name: 60, subtitle: 100, crest: 3, title: 100, clinicFull: 120, phone: 40, loginTitle: 60, loginSub: 160,
  clinicType: 80, address: 250, email: 80, taxId: 30, licenseNo: 40, receiver: 120, receiptNote: 200, logo: 400 };

var COLOR_KEYS = { primary: "สีหลัก (หัวเว็บ/เมนู/ปุ่ม)", gold: "สีเน้น (เส้นขอบ/ปุ่มรอง)", bg: "สีพื้นหลังหน้าเว็บ" };
function isHex_(v) { return /^#[0-9a-fA-F]{6}$/.test(S_(v)); }

/** ลิงก์โลโก้ → ลิงก์รูปที่ใช้แสดงในหน้าเว็บได้ (ไฟล์ Drive ต้องแชร์ "ทุกคนที่มีลิงก์") */
function logoSrc_(v) {
  v = S_(v);
  if (!v) return '';
  if (/^data:image\//.test(v)) return v;
  if (!/^https?:\/\//i.test(v)) return '';
  if (/drive\.google\.com|docs\.google\.com/i.test(v)) {
    var id = pzIdFromUrl_(v);
    if (id) return 'https://drive.google.com/thumbnail?id=' + id + '&sz=w400';
  }
  return v;
}
/** โลโก้แบบฝังในเอกสาร (ใช้กับใบรับรองแพทย์/ใบเสร็จ เพื่อให้ติดไปใน PDF ด้วย) */
function getLogoDataUri_(b) {
  var src = S_(b && b.logo);
  if (!src) return '';
  var id = pzIdFromUrl_(src);
  if (!(/drive\.google\.com|docs\.google\.com/i.test(src) && id)) return b.logoUrl || '';
  var cache = CacheService.getScriptCache(), key = 'aw_logo_' + id;
  var c = cache.get(key);
  if (c) return c;
  try {
    var blob = DriveApp.getFileById(id).getBlob(), bytes = blob.getBytes();
    if (bytes.length > 700 * 1024) return b.logoUrl || '';
    var uri = 'data:' + (blob.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(bytes);
    try { cache.put(key, uri, 3600); } catch (e) {}
    return uri;
  } catch (e) { return b.logoUrl || ''; }
}

/** เปิดให้เรียกได้ตั้งแต่หน้าล็อคอิน (ไม่มีข้อมูลลับ) */
function getBranding() {
  var cache = CacheService.getScriptCache(), c = cache.get('aw_brand');
  if (c) { try { return JSON.parse(c); } catch (e) {} }
  var s = readSettings_(), b = JSON.parse(JSON.stringify(BRAND_DEFAULTS));
  Object.keys(BRAND_KEYS).forEach(function (k) { if (S_(s[BRAND_KEYS[k]])) b[k] = S_(s[BRAND_KEYS[k]]); });
  Object.keys(b.nav).forEach(function (p) { var v = S_(s['เมนู: ' + p]); if (v) b.nav[p] = v; });
  Object.keys(COLOR_KEYS).forEach(function (k) { var v = S_(s[COLOR_KEYS[k]]); if (isHex_(v)) b.colors[k] = v.toLowerCase(); });
  b.noLogo = (b.logo === 'none');
  if (b.noLogo) b.logo = '';
  b.logoUrl = logoSrc_(b.logo);
  b.receiverName = b.receiver || b.clinicFull;
  try { b.doctors = getDoctors(); } catch (e) { b.doctors = []; }
  try { cache.put('aw_brand', JSON.stringify(b), 300); } catch (e) {}
  return b;
}
/** เขียนค่าเป็นข้อความล้วน (เบอร์โทรไม่ถูกตัดเลข 0 / ข้อความขึ้นต้น = ไม่กลายเป็นสูตร) */
function writeSettingsBatch_(map, desc) {
  var sh = settingSheet_(), d = sh.getDataRange().getValues(), idx = {};
  for (var i = 0; i < d.length; i++) idx[S_(d[i][0])] = i + 1;
  Object.keys(map).forEach(function (k) {
    var v = String(map[k]), r = idx[k];
    if (!r) { r = sh.getLastRow() + 1; sh.getRange(r, 1).setValue(k); sh.getRange(r, 3).setValue(desc || ""); idx[k] = r; }
    sh.getRange(r, 2).setNumberFormat('@').setValue(v);
  });
}
function adminSaveBranding(token, b) {
  requireAdmin_(token);
  b = b || {};
  var map = {};
  Object.keys(BRAND_KEYS).forEach(function (k) { map[BRAND_KEYS[k]] = S_(b[k]).substring(0, BRAND_LIM[k] || 200); });
  if (b.noLogo) map[BRAND_KEYS.logo] = 'none';
  Object.keys(BRAND_DEFAULTS.nav).forEach(function (p) { map['เมนู: ' + p] = S_((b.nav || {})[p]).substring(0, 30); });
  Object.keys(COLOR_KEYS).forEach(function (k) { var v = S_((b.colors || {})[k]); map[COLOR_KEYS[k]] = isHex_(v) ? v.toLowerCase() : ''; });
  writeSettingsBatch_(map, 'ตั้งค่าจากหน้า Admin (เว้นว่าง = ใช้ค่าเริ่มต้น)');
  CacheService.getScriptCache().remove('aw_brand');
  return '✅ บันทึกข้อมูลคลินิกและหน้าตาเว็บเรียบร้อย';
}
/** อัปโหลดโลโก้จากหน้าแอดมิน → เก็บลง Drive + เปิดแชร์แบบดูอย่างเดียว + บันทึกลิงก์ */
function adminUploadLogo(token, dataUrl) {
  requireAdmin_(token);
  var parts = String(dataUrl || '').split(';base64,');
  if (parts.length < 2 || !/^data:image\//.test(parts[0])) return { ok: false, message: 'ไฟล์รูปไม่ถูกต้อง' };
  var ct = parts[0].split(':')[1] || 'image/png', bytes = Utilities.base64Decode(parts[1]);
  if (bytes.length > 2 * 1024 * 1024) return { ok: false, message: 'รูปใหญ่เกิน 2 MB' };
  var folder = pzFolder_('', 'AphaiWellness_Logo');
  var file = folder.createFile(Utilities.newBlob(bytes, ct, 'Logo_' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd_HHmmss') + (ct.indexOf('png') > -1 ? '.png' : '.jpg')));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  var url = file.getUrl(), m = {};
  m[BRAND_KEYS.logo] = url;
  writeSettingsBatch_(m, 'โลโก้คลินิก (อัปโหลดจากหน้า Admin)');
  CacheService.getScriptCache().remove('aw_brand');
  return { ok: true, url: url, src: logoSrc_(url) };
}

// ==========================================================
// 👨‍⚕️ แพทย์ผู้ตรวจ (ชีต "Doctors") — ใช้ในใบรับรองแพทย์
// ==========================================================
var DOC_HEADERS = ['ใช้งาน (TRUE/FALSE)', 'ชื่อ-นามสกุล (รวมคำนำหน้า)', 'เลขที่ใบอนุญาตประกอบวิชาชีพ', 'วิชาชีพ (พิมพ์ใต้ลายเซ็น)'];
var DOC_DEFAULTS = [
  [true, 'นายพีรธัช ศรีพันธุ์', 'พท.ว.25339', 'ผู้ประกอบวิชาชีพเวชกรรมไทย'],
  [true, 'นางสาวนัชชา บุญมี', 'พท.ว.27505', 'ผู้ประกอบวิชาชีพเวชกรรมไทย']
];
function doctorSheet_() {
  var ss = getSS(), sh = ss.getSheetByName('Doctors');
  if (!sh) {
    sh = ss.insertSheet('Doctors');
    sh.appendRow(DOC_HEADERS);
    sh.getRange(2, 1, DOC_DEFAULTS.length, DOC_HEADERS.length).setValues(DOC_DEFAULTS);
    sh.setFrozenRows(1);
  }
  return sh;
}
function readDoctors_(all) {
  var d = doctorSheet_().getDataRange().getValues(), out = [];
  for (var i = 1; i < d.length; i++) {
    var r = d[i];
    if (!S_(r[1])) continue;
    var act = isActive_(r[0]);
    if (!all && !act) continue;
    out.push({ active: act, name: S_(r[1]), license: S_(r[2]), title: S_(r[3]) || 'ผู้ประกอบวิชาชีพเวชกรรมไทย' });
  }
  return out;
}
function getDoctors() { try { return readDoctors_(false); } catch (e) { return []; } }
function adminSaveDoctors(token, list) {
  requireAdmin_(token);
  var rows = [];
  (list || []).forEach(function (x) {
    var n = S_(x && x.name).substring(0, 120);
    if (n) rows.push([x.active !== false, n, S_(x.license).substring(0, 40), S_(x.title).substring(0, 80) || 'ผู้ประกอบวิชาชีพเวชกรรมไทย']);
  });
  var sh = doctorSheet_();
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, DOC_HEADERS.length).clearContent();
  if (rows.length) {
    sh.getRange(2, 3, rows.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, rows.length, DOC_HEADERS.length).setValues(rows);
  }
  CacheService.getScriptCache().remove('aw_brand');
  return '✅ บันทึกรายชื่อแพทย์ ' + rows.length + ' ท่าน';
}

function doGet(e) {
  var prm = (e && e.parameter) || {};
  var p = String(prm.p || "Login").toLowerCase();
  
  // 🌟 บังคับเปิดหน้า PublicElement ทันทีโดยไม่ต้องผ่านระบบล็อกอิน
  if (p === 'publicelement') {
    var brand = getBranding();
    var t = HtmlService.createTemplateFromFile("PublicElement");
    t.brand = brand;
    t.brandJson = JSON.stringify(brand).replace(/</g, '\\u003c');
    t.logo = brand.logoUrl || "";
    return t.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle(brand.title);
  }

  var map = {
    login: "Login", register: "Register", search: "Search", index: "Index", history: "History",
    tracking: "posture_analyzer", element: "Element", admin: "Admin",
    pharmacy: "Pharmacy", certtemplate: "CertTemplate", receipttemplate: "ReceiptTemplate", labeltemplate: "LabelTemplate"
  };
  var fileName = map[p] || "Login";
  try {
    var brand = getBranding();
    var isDoc = (p === "certtemplate" || p === "receipttemplate" || p === "labeltemplate");
    var t = HtmlService.createTemplateFromFile(fileName);
    t.mode = prm.mode || "add";
    t.cn = prm.cn || "";
    t.row = prm.row || "";
    t.embed = prm.embed || "";
    t.session = prm.session || "";
    t.focus = prm.focus || "";

    // ----- ข้อมูลคลินิกสำหรับทุกหน้า -----
    t.brand = brand;
    t.brandJson = JSON.stringify(brand).replace(/</g, '\\u003c');
    t.logo = isDoc ? getLogoDataUri_(brand) : (brand.logoUrl || "");

    // ----- ตัวแปรสำหรับหน้าเอกสาร (ใบรับรองแพทย์ / ใบเสร็จ / ฉลากยา) -----
    var fullName = prm.name || "";
    if ((p === "certtemplate" || p === "receipttemplate") && prm.cn) {
      var pd = searchPatientForEdit(prm.cn);                       
      if (pd) {
        var pre = pd.prefix || "", nm = pd.name || "";
        fullName = (nm && !String(nm).startsWith(pre)) ? pre + nm : nm;
      }
    }
    t.name = fullName;
    t.patientName = prm.patient || fullName;
    t.idCard = prm.idCard || "";
    t.address = prm.address || "";
    t.dx = prm.dx || "";
    t.treatment = prm.treatment || "";
    t.advice = prm.advice || "";
    t.date = prm.date || "";
    t.total = prm.total || "";
    t.medName = prm.medName || prm.name || "";
    t.medUsage = prm.usage || "";
    t.description = prm.description || "";
    t.warning = prm.warning || "";
    t.amount = prm.amount || "";
    t.recNo = (p === "receipttemplate") ? getNextReceiptNumber(prm.cn || "") : "";

    return t.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle(brand.title);
  } catch (err) {
    return HtmlService.createHtmlOutput("<div style='text-align:center;padding:20px;font-family:Sarabun,sans-serif;'><h1>❌ ระบบหาไฟล์ไม่พบ</h1><p>ไฟล์: <b style='color:red;'>" + fileName + "</b></p><p>" + err.message + "</p></div>");
  }
}

// ==========================================================
// 🔐 ผู้ใช้งาน / เข้าสู่ระบบ (User_DB: ชื่อผู้ใช้, รหัสผ่าน, ชื่อที่แสดง, สิทธิ์, ใช้งาน)
// ==========================================================
var USER_HEADERS = ['ชื่อผู้ใช้งาน', 'รหัสผ่าน', 'ชื่อที่แสดง', 'สิทธิ์ (admin/staff)', 'ใช้งาน (TRUE/FALSE)'];

function usersSheet_() {
  var ss = getSS(), sh = ss.getSheetByName('User_DB');
  if (!sh) { sh = ss.insertSheet('User_DB'); sh.appendRow(USER_HEADERS); }
  else if (sh.getLastRow() === 0) sh.appendRow(USER_HEADERS);
  else if (sh.getLastColumn() < USER_HEADERS.length) sh.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]);
  return sh;
}
function hashPass_(u, p) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(u).toLowerCase() + '|' + String(p), Utilities.Charset.UTF_8);
  return 'sha256:' + b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}
function passOk_(u, stored, p) {
  stored = String(stored == null ? '' : stored);
  if (stored.indexOf('sha256:') === 0) return stored === hashPass_(u, p);
  return stored !== '' && stored === String(p);          // รองรับรหัสเดิมที่ยังไม่เข้ารหัส
}
function recordLoginHistory(username) {
  try {
    var ss = getSS(), sheet = ss.getSheetByName('Login_Logs');
    if (!sheet) { sheet = ss.insertSheet('Login_Logs'); sheet.appendRow(['วัน-เวลา', 'ชื่อผู้ใช้งาน', 'สถานะ']); }
    sheet.appendRow([new Date(), username, "เข้าสู่ระบบ"]);
    return true;
  } catch (e) { return false; }
}
function checkLoginFromSheet(userInput, passInput) {
  var data = usersSheet_().getDataRange().getValues(), u = S_(userInput);
  for (var i = 1; i < data.length; i++) {
    var su = S_(data[i][0]);
    if (su && su === u && isActive_(data[i][4]) && passOk_(su, data[i][1], passInput)) {
      var role = S_(data[i][3]).toLowerCase() === 'admin' ? 'admin' : 'staff';
      var token = Utilities.getUuid();
      CacheService.getScriptCache().put('aw_tok_' + token, JSON.stringify({ u: su, role: role }), 21600);
      recordLoginHistory(su);
      return { success: true, username: su, displayName: S_(data[i][2]) || su, role: role, token: token };
    }
  }
  return { success: false };
}
function requireAdmin_(token) {
  var v = token ? CacheService.getScriptCache().get('aw_tok_' + token) : null;
  var a = v ? JSON.parse(v) : null;
  if (!a || a.role !== 'admin') throw new Error('เซสชันผู้ดูแลระบบหมดอายุหรือไม่มีสิทธิ์ — กรุณาออกจากระบบแล้วเข้าสู่ระบบใหม่');
  return a;
}

// ---------- ฟังก์ชันหน้า Admin ----------
function adminGetData(token) {
  requireAdmin_(token);
  var d = usersSheet_().getDataRange().getValues(), users = [];
  for (var i = 1; i < d.length; i++) {
    if (!S_(d[i][0])) continue;
    users.push({ username: S_(d[i][0]), displayName: S_(d[i][2]), role: S_(d[i][3]).toLowerCase() === 'admin' ? 'admin' : 'staff', active: isActive_(d[i][4]) });
  }
  var s = readSettings_(), docs = [];
  try { docs = readDoctors_(true); } catch (e) {}
  return {
    users: users, postureFolder: s[SET_KEYS.postureFolder] || "", signatureFolder: s[SET_KEYS.signatureFolder] || "",
    receiptFolder: s[SET_KEYS.receiptFolder] || "", certFolder: s[SET_KEYS.certFolder] || "",
    links: getLinks(), music: getMusicList(), doctors: docs, sheetUrl: getSS().getUrl(), appUrl: getAppUrl()
  };
}
function adminSaveUser(token, u) {
  var me = requireAdmin_(token), lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var name = S_(u && u.username);
    if (!/^\S{3,40}$/.test(name)) return '❌ ชื่อผู้ใช้งานต้องมี 3-40 ตัวอักษร และห้ามมีช่องว่าง';
    var sh = usersSheet_(), d = sh.getDataRange().getValues(), row = -1;
    for (var i = 1; i < d.length; i++) if (S_(d[i][0]) === name) { row = i + 1; break; }
    if (u.isNew && row > -1) return '❌ มีชื่อผู้ใช้งานนี้อยู่แล้ว';
    if (!u.isNew && row === -1) return '❌ ไม่พบผู้ใช้งานนี้';
    var role = u.role === 'admin' ? 'admin' : 'staff', active = u.active !== false;
    if (name === me.u && (role !== 'admin' || !active)) return '❌ ไม่สามารถลดสิทธิ์หรือปิดการใช้งานบัญชีของตัวเองได้';
    var pass;
    if (u.password) {
      if (String(u.password).length < 4) return '❌ รหัสผ่านต้องมีอย่างน้อย 4 ตัวอักษร';
      pass = hashPass_(name, u.password);
    } else if (row > -1) pass = d[row - 1][1];
    else return '❌ กรุณาตั้งรหัสผ่านสำหรับผู้ใช้ใหม่';
    var vals = [name, pass, S_(u.displayName) || name, role, active];
    if (row > -1) sh.getRange(row, 1, 1, 5).setValues([vals]); else sh.appendRow(vals);
    return '✅ บันทึกผู้ใช้งาน "' + name + '" เรียบร้อย';
  } finally { try { lock.releaseLock(); } catch (e) {} }
}
function adminDeleteUser(token, username) {
  var me = requireAdmin_(token), name = S_(username);
  if (name === me.u) return '❌ ไม่สามารถลบบัญชีของตัวเองได้';
  var sh = usersSheet_(), d = sh.getDataRange().getValues();
  for (var i = d.length - 1; i >= 1; i--) if (S_(d[i][0]) === name) { sh.deleteRow(i + 1); return '✅ ลบผู้ใช้งาน "' + name + '" แล้ว'; }
  return '❌ ไม่พบผู้ใช้งาน';
}
function adminSaveSettings(token, s) {
  requireAdmin_(token);
  s = s || {};
  writeSetting_(SET_KEYS.postureFolder, S_(s.postureFolder), 'โฟลเดอร์ Google Drive เก็บภาพ Posture Tracking');
  writeSetting_(SET_KEYS.signatureFolder, S_(s.signatureFolder), 'โฟลเดอร์ Google Drive เก็บลายเซ็น/รูปยืนยันการรักษา');
  writeSetting_(SET_KEYS.receiptFolder, S_(s.receiptFolder), 'โฟลเดอร์ Google Drive เก็บ PDF ใบเสร็จ');
  writeSetting_(SET_KEYS.certFolder, S_(s.certFolder), 'โฟลเดอร์ Google Drive เก็บ PDF ใบรับรองแพทย์');
  var all = s.links || [], rows = [];
  all.forEach(function (l) {
    var url = S_(l && l.url);
    if (/^https?:\/\//i.test(url)) rows.push([S_(l.name) || url, url, S_(l.note)]);
  });
  var sh = linksSheet_();
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
  var skipped = all.filter(function (l) { return l && (S_(l.url) || S_(l.name)); }).length - rows.length;
  return '✅ บันทึกการตั้งค่าเรียบร้อย' + (skipped > 0 ? ' (ข้าม ' + skipped + ' ลิงก์ที่ไม่ได้ขึ้นต้นด้วย http/https)' : '');
}
function adminTestFolder(token, url) {
  requireAdmin_(token);
  var id = pzIdFromUrl_(url);
  if (!id) return { ok: false, message: 'ไม่พบรหัสโฟลเดอร์ในลิงก์' };
  try { var f = DriveApp.getFolderById(id); return { ok: true, name: f.getName() }; }
  catch (e) { return { ok: false, message: 'เปิดโฟลเดอร์ไม่ได้: ' + e.message }; }
}

/** ▶️ รันครั้งเดียวจาก Editor: สร้างชีตที่จำเป็น + ตั้งแอดมินคนแรก */
function setupSystem() {
  var sh = usersSheet_();
  settingSheet_(); linksSheet_(); postureDxSheet_(); pthSheet_(); visitSheet_(); treatmentSheet_(); musicSheet_(); doctorSheet_();
  var d = sh.getDataRange().getValues(), hasAdmin = false;
  for (var i = 1; i < d.length; i++) if (S_(d[i][3]).toLowerCase() === 'admin') hasAdmin = true;
  if (!hasAdmin) {
    if (d.length > 1) { sh.getRange(2, 4).setValue('admin'); Logger.log('ตั้ง "' + d[1][0] + '" เป็นแอดมินแล้ว'); }
    else { sh.appendRow(['admin', hashPass_('admin', 'admin1234'), 'ผู้ดูแลระบบ', 'admin', true]); Logger.log('สร้างผู้ใช้ admin / admin1234 (กรุณาเปลี่ยนรหัสทันที)'); }
  }
  PropertiesService.getScriptProperties().setProperty('SS_ID', getSS().getId());
  Logger.log('✅ ตั้งค่าระบบเรียบร้อย');
}

// ==========================================================
// 👤 ผู้ป่วย (Patient_DB)
// ==========================================================
function pzNextCN_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return "0001";
  var vals = sheet.getRange(2, 1, lastRow - 1, 1).getValues(), max = 0;
  for (var i = 0; i < vals.length; i++) {
    var n = parseInt(String(vals[i][0]).replace(/'/g, "").trim(), 10);
    if (!isNaN(n) && n > max) max = n;
  }
  return String(max + 1).padStart(4, '0');
}
function getNextCN() {
  try { var s = getSS().getSheetByName('Patient_DB'); return s ? pzNextCN_(s) : "0001"; } catch (e) { return "0001"; }
}
function patientRow_(r) {
  return {
    cn: pzNorm_(r[0]), am: r[1], prefix: r[2], name: r[3], nickname: r[4],
    dob: pzDate_(r[5]), phone: r[6], gender: r[7], weight: r[8], height: r[9], job: r[10],
    idCard: r[11], address: r[12], relName: r[13], relPhone: r[14], disease: r[15], medicalPlan: r[16]
  };
}
function searchPatientForEdit(cn) {
  try {
    var data = getSS().getSheetByName('Patient_DB').getDataRange().getValues(), target = pzNorm_(cn);
    for (var i = 1; i < data.length; i++) if (pzNorm_(data[i][0]) === target) return patientRow_(data[i]);
    return null;
  } catch (e) { return null; }
}
function registerNewPatient(p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var sheet = getSS().getSheetByName('Patient_DB');
    if (!sheet) return "❌ ไม่พบ Sheet ชื่อ Patient_DB";
    var data = sheet.getDataRange().getValues();
    var cn = p.isNew ? pzNextCN_(sheet) : String(p.cn || "").replace(/'/g, "").trim().padStart(4, '0');
    var rowData = [
      "'" + cn, p.am || "", p.prefix || "", p.name || "", p.nickname || "",
      "'" + (p.dob || ""), "'" + (p.phone || ""), p.gender || "",
      "'" + (p.weight || ""), "'" + (p.height || ""), p.job || "",
      "'" + (p.idCard || ""), "'" + (p.address || ""), p.relName || "",
      "'" + (p.relPhone || ""), p.disease || "", p.medicalPlan || ""
    ];
    var rowIndex = -1;
    if (!p.isNew) for (var i = 1; i < data.length; i++) if (pzNorm_(data[i][0]) === cn) { rowIndex = i + 1; break; }
    if (rowIndex > -1) { sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]); return "✅ อัปเดตข้อมูลสำเร็จ (CN: " + cn + ")"; }
    sheet.appendRow(rowData);
    return "✅ ลงทะเบียนสำเร็จ (CN: " + cn + ")";
  } catch (e) { return "❌ ผิดพลาดที่ Server: " + e.message; }
  finally { try { lock.releaseLock(); } catch (e2) {} }
}
function searchAllPatients(query) {
  var sheet = getSS().getSheetByName('Patient_DB');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues(), out = [], q = String(query || "").toLowerCase().trim();
  for (var i = 1; i < data.length; i++) {
    var cn = pzNorm_(data[i][0]), nm = String(data[i][3] || "").toLowerCase(), ph = String(data[i][6] || "");
    if (cn.includes(q) || nm.includes(q) || ph.includes(q)) out.push(patientRow_(data[i]));
  }
  return out;
}
/** ดึงผู้ป่วยตาม CN สำหรับหน้าบันทึกการรักษา + วิเคราะห์ธาตุเจ้าเรือนให้ทันที */
function searchPatient(query) {
  try {
    var data = getSS().getSheetByName('Patient_DB').getDataRange().getValues(), term = pzNorm_(query);
    for (var i = 1; i < data.length; i++) {
      if (pzNorm_(data[i][0]) !== term) continue;
      var r = data[i], dob = pzDobIso_(r[5]), el = null;
      if (dob) { try { el = getElementInfo(dob); } catch (e) { el = null; } }
      return {
        cn: pzNorm_(r[0]), name: (r[2] || "") + (r[3] || "") || "-", phone: r[6] || "-",
        weight: r[8] || "", height: r[9] || "", dob: dob, element: el,
        medicalInfo: "⚠️ แพ้ยา/โรคประจำตัว: " + (r[15] || "-") + " | สิทธิ: " + (r[16] || "-")
      };
    }
    return null;
  } catch (e) { return null; }
}

// ==========================================================
// 🩺 บันทึกการรักษา (Visit_Logs) 28 คอลัมน์
// ==========================================================
var VISIT_HEADERS = ['วันที่', 'เวลา', 'CN', 'ชื่อ', 'น้ำหนัก', 'ส่วนสูง', 'BP', 'PR', 'RR', 'T', 'CC', 'PI', 'PH', 'PE', 'Dx', 'หัตถการ',
  'คำแนะนำ', 'Plan', 'นัด', 'ผู้ตรวจ', 'ภาพยืนยัน', '(สำรอง)', 'ปวดก่อน', 'ปวดหลัง', 'ธาตุเจ้าเรือน', 'Posture ครั้งที่', 'ประเภทการยืนยัน', 'ไฟล์ยืนยันใน Drive'];
var VISIT_COLS = 29;

function visitSheet_() {
  var ss = getSS(), sh = ss.getSheetByName('Visit_Logs');
  if (!sh) { sh = ss.insertSheet('Visit_Logs'); sh.appendRow(VISIT_HEADERS); }
  if (sh.getMaxColumns() < VISIT_COLS) sh.insertColumnsAfter(sh.getMaxColumns(), VISIT_COLS - sh.getMaxColumns());
  if (sh.getLastRow() >= 1) {
    var h = sh.getRange(1, 26, 1, 3).getValues()[0];
    if (!h[0] && !h[1] && !h[2]) sh.getRange(1, 26, 1, 3).setValues([VISIT_HEADERS.slice(25)]);
  }
  return sh;
}
function saveConfirmToDrive_(cn, dataUrl, type) {
  var folder = pzFolder_(readSettings_()[SET_KEYS.signatureFolder], "AphaiWellness_Signatures");
  var parts = String(dataUrl).split(";base64,");
  var ct = (parts[0].split(":")[1]) || "image/jpeg";
  var fname = (type === 'ลายเซ็น' ? 'Sign_' : 'Confirm_') + pzNorm_(cn) + '_' + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd_HHmmss") + '.jpg';
  return folder.createFile(Utilities.newBlob(Utilities.base64Decode(parts[1]), ct, fname)).getUrl();   // ไม่เปิดแชร์สาธารณะ
}
function saveVisit(d) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var sh = visitSheet_();
    var ph = "", type = "", drive = "", pain = S_(d.painPoints);
    if (d.photoChanged && /^data:image\/.*$/.test(String(d.photo || ""))) {
      ph = String(d.photo);
      if (ph.length > 49000) return "❌ ไฟล์ลายเซ็น/รูปยืนยันใหญ่เกินไป กรุณาลองใหม่";
      type = d.confirmType || "";
      try { drive = saveConfirmToDrive_(d.cn, ph, type); } catch (e) { drive = ""; }
    }
    var build = function (date, time) {
      return [
        date,                     // 1. วันที่ (A)
        time,                     // 2. เวลา (B)
        "" + pzNorm_(d.cn),       // 3. CN (C)
        d.name,                   // 4. ชื่อ (D)
        d.weight,                 // 5. น้ำหนัก (E)
        d.height,                 // 6. ส่วนสูง (F)
        d.bp,                     // 7. BP (G)
        d.pr,                     // 8. PR (H)
        d.rr,                     // 9. RR (I)
        d.temp,                   // 10. T (J)
        d.cc,                     // 11. CC (K)
        d.pi,                     // 12. PI (L)
        d.ph,                     // 13. PH (M)
        d.pe,                     // 14. PE (N)
        d.dx,                     // 15. Dx (O)
        d.treatment,              // 16. หัตถการ (P)
        d.advice,                 // 17. คำแนะนำ (Q)
        d.plan,                   // 18. Plan (R)
        d.appointment,            // 19. นัด (S)
        d.staffName,              // 20. ผู้ตรวจ (T)
        ph,                       // 21. ภาพยืนยัน (U)
        "",                       // 22. (สำรอง) (V)
        d.prePain || 0,           // 23. ปวดก่อน (W)
        d.postPain || 0,          // 24. ปวดหลัง (X)
        d.element || "",          // 25. ธาตุเจ้าเรือน (Y)
        d.postureSession || "",   // 26. Posture ครั้งที่ (Z)
        type,                     // 27. ประเภทการยืนยัน (AA)
        drive,                    // 28. ไฟล์ยืนยันใน Drive (AB)
        pain                      // 29. painPoints (AC)
      ];
    };
    if (d.rowIndex) {
      var r = parseInt(d.rowIndex, 10);
      var old = sh.getRange(r, 1, 1, VISIT_COLS).getValues()[0];
      if (pzNorm_(old[2]) !== pzNorm_(d.cn)) return "❌ แถวที่จะแก้ไขไม่ตรงกับผู้ป่วย กรุณาเปิดจากหน้าประวัติใหม่";
      if (!d.photoChanged) { ph = old[20]; type = old[26]; drive = old[27]; }
      if (!pain) { pain = S_(old[28]); }
      sh.getRange(r, 1, 1, VISIT_COLS).setValues([build(old[0], old[1])]);
      return "✅ บันทึกการแก้ไขสำเร็จ";
    }
    var now = new Date();
    sh.appendRow(build("" + Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy")), "" + Utilities.formatDate(now, "GMT+7", "HH:mm:ss"));
    return "✅ บันทึกสำเร็จ";
  } catch (e) { return "❌ Error: " + e.message; }
  finally { try { lock.releaseLock(); } catch (e2) {} }
}

function getVisitHistory(cnQuery) {
  try {
    var ss = getSS(), vs = ss.getSheetByName('Visit_Logs'), ps = ss.getSheetByName('Patient_DB');
    if (!vs) return [];
    var vd = vs.getDataRange().getValues(), pMap = {};
    var clean = function (v) { return String(v == null ? "" : v).replace(/'/g, "").trim(); };
    if (ps) {
      var pd = ps.getDataRange().getValues();
      for (var j = 1; j < pd.length; j++) {
        pMap[pzNorm_(pd[j][0])] = {
          disease: pd[j][15], idCard: clean(pd[j][11]), address: clean(pd[j][12]), phone: clean(pd[j][6])
        };
      }
    }
    var key = (cnQuery && cnQuery !== "") ? pzNorm_(cnQuery) : null;
    var S = function (v, d) { return String(v == null || v === "" ? (d !== undefined ? d : "-") : v).trim(); };
    var out = [];
    for (var i = 1; i < vd.length; i++) {
      var v = vd[i];
      if (!v[2]) continue;
      var cn = pzNorm_(v[2]);
      if (key && cn !== key) continue;
      var pi = pMap[cn] || { disease: "-", idCard: "", address: "", phone: "" };
      out.push({
        rowIndex: i + 1,
        date: (v[0] instanceof Date) ? Utilities.formatDate(v[0], "GMT+7", "dd/MM/yyyy") : String(v[0] || "").replace(/'/g, "") + "",
        time: (v[1] instanceof Date) ? Utilities.formatDate(v[1], "GMT+7", "HH:mm:ss") : String(v[1] || "").replace(/'/g, "") + "",
        cn: cn,
        name: S(v[3], "-"),
        weight: S(v[4], "-"),
        height: S(v[5], "-"),
        bp: S(v[6], "-"),
        pr: S(v[7], "-"),
        rr: S(v[8], "-"),
        temp: S(v[9], "-"),
        cc: S(v[10], "-"),
        pi: S(v[11], "-"),
        ph: S(v[12], "-"),
        pe: S(v[13], "-"),
        dx: S(v[14], "-"),
        treatment: S(v[15], "-"),
        advice: S(v[16], "-"),
        plan: S(v[17], "-"),
        appointment: (v[18] instanceof Date) ? Utilities.formatDate(v[18], "GMT+7", "yyyy-MM-dd") : S(v[18], "-"),
        staff: S(v[19], "-"),         // ผู้ตรวจ (index 19 / คอลัมน์ T)
        photo: S(v[20], ""),         // ภาพยืนยัน (index 20 / คอลัมน์ U)
        prePain: S(v[22], "0"),      // ปวดก่อน (index 22 / คอลัมน์ W)
        postPain: S(v[23], "0"),     // ปวดหลัง (index 23 / คอลัมน์ X)
        element: S(v[24], ""),       // ธาตุเจ้าเรือน (index 24 / คอลัมน์ Y)
        postureSession: S(v[25], ""),// Posture ครั้งที่ (index 25 / คอลัมน์ Z)
        confirmType: S(v[26], ""),   // ประเภทการยืนยัน (index 26 / คอลัมน์ AA)
        confirmDrive: S(v[27], ""),  // ไฟล์ยืนยันใน Drive (index 27 / คอลัมน์ AB)
        painPoints: S(v[28], ""),    // painPoints (index 28 / คอลัมน์ AC)
        disease: pi.disease, idCard: pi.idCard, address: pi.address, phone: pi.phone
      });
    }
    return out.reverse();
  } catch (e) { return "❌ เกิดข้อผิดพลาดที่ Server: " + e.toString(); }
}

function getVisitByRow(row) {
  var all = getVisitHistory(""), r = parseInt(row, 10);
  if (!Array.isArray(all)) return null;
  for (var i = 0; i < all.length; i++) if (all[i].rowIndex === r) return all[i];
  return null;
}
function getVisitStats() {
  try {
    var ss = getSS(), vs = ss.getSheetByName('Visit_Logs'), ps = ss.getSheetByName('Patient_DB');
    var totalPatients = ps ? Math.max(0, ps.getLastRow() - 1) : 0, today = 0, cycle = 0;
    var now = new Date(), todayStr = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy");
    var start = now.getDate() >= 26 ? new Date(now.getFullYear(), now.getMonth(), 26) : new Date(now.getFullYear(), now.getMonth() - 1, 26);
    var end = now.getDate() >= 26 ? new Date(now.getFullYear(), now.getMonth() + 1, 25, 23, 59, 59) : new Date(now.getFullYear(), now.getMonth(), 25, 23, 59, 59);
    if (vs) {
      var vd = vs.getDataRange().getValues();
      for (var i = 1; i < vd.length; i++) {
        var d = vd[i][0], dt = (d instanceof Date) ? d : null;
        if (!dt && d) { var p = d.toString().replace(/'/g, "").split('/'); if (p.length >= 3) dt = new Date(p[2], p[1] - 1, p[0]); }
        if (dt && !isNaN(dt.getTime())) {
          if (Utilities.formatDate(dt, "GMT+7", "dd/MM/yyyy") === todayStr) today++;
          if (dt >= start && dt <= end) cycle++;
        }
      }
    }
    return { totalAll: totalPatients, totalToday: today, totalMonthCycle: cycle,
      cycleRange: Utilities.formatDate(start, "GMT+7", "dd/MM") + " - " + Utilities.formatDate(end, "GMT+7", "dd/MM") };
  } catch (e) { return { totalAll: 0, totalToday: 0, totalMonthCycle: 0, cycleRange: "Error" }; }
}
function deleteVisit(rowNum, cn) {
  try {
    var sh = getSS().getSheetByName('Visit_Logs'), r = parseInt(rowNum, 10);
    if (cn && pzNorm_(sh.getRange(r, 3).getValue()) !== pzNorm_(cn)) return "❌ แถวไม่ตรงกับผู้ป่วย กรุณารีเฟรชหน้าแล้วลองใหม่";
    sh.deleteRow(r);
    return "✅ ลบข้อมูลเรียบร้อยแล้ว";
  } catch (e) { return "❌ Error: " + e.message; }
}
function getDiseaseList() {
  try {
    var sheet = getSS().getSheetByName("ICD");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues(), list = [];
    for (var i = 1; i < data.length; i++) {
      var code = data[i][0] ? String(data[i][0]).trim() : "", name = data[i][1] ? String(data[i][1]).trim() : "";
      var modern = data[i][2] ? String(data[i][2]).trim() : "";          // คอลัมน์ C = ชื่อโรคแผนปัจจุบัน (ถ้ามี)
      if (code || name) list.push({ code: code, name: name, modern: modern, display: code + " - " + name });
    }
    return list;
  } catch (err) { return []; }
}

// ==========================================================
// 🍃 ธาตุเจ้าเรือน / สมุนไพร (ชีต "ชีต1")
// ==========================================================
function getElementInfo(birthDateStr) {
  var p = pzDobParts_(birthDateStr);
  if (!p) return null;
  var sheet = getSS().getSheetByName("ชีต1");
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  // ชีตนับ 365 วัน (ไม่มี 29 ก.พ.) → ใช้ปีสมมติที่ไม่อธิกสุรทิน และให้ 29 ก.พ. = 28 ก.พ.
  var d = (p.m === 2 && p.d > 28) ? 28 : p.d;
  var dayOfYear = Math.round((Date.UTC(2001, p.m - 1, d) - Date.UTC(2001, 0, 0)) / 86400000);
  for (var i = 1; i < data.length; i++) {
    if (data[i][3] == dayOfYear) {                               // คอลัมน์ D = ลำดับวัน
      var bg = ["", "", ""];
      try { bg = sheet.getRange(i + 1, 6, 1, 3).getBackgrounds()[0]; } catch (e) {}   // สีพื้นหลังช่อง F, G, H ของแถวนี้
      return {
        dayOrder: dayOfYear, element: data[i][4], appearance: data[i][5], caution: data[i][6], mind: data[i][7],
        commonIllness: data[i][9], food: data[i][10], lifestyle: data[i][11], oil: data[i][12],
        colorR: S_(data[i][5]), colorS: S_(data[i][6]), colorT: S_(data[i][7]),
        bgR: bg[0], bgS: bg[1], bgT: bg[2]                        // F = วงนอก, G = วงกลาง, H = วงใน
      };
    }
  }
  return null;
}
function searchHerb(keyword) {
  var sheet = getSS().getSheetByName("ชีต1");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues(), out = [], kw = String(keyword || "").toLowerCase().trim();
  for (var i = 1; i < data.length; i++) {
    var symptom = String(data[i][45] || "");
    if (kw !== "" && symptom.toLowerCase().includes(kw)) {
      out.push({ disease: data[i][44], symptom: symptom, modernGroup: data[i][46],
        sideEffect: data[i][47], herbSolution: data[i][48], interaction: data[i][49], testLink: data[i][50] });
    }
  }
  return out;
}

// ==========================================================
// 🧠 กฎวินิจฉัยจาก Posture (ชีต "Posture_Dx_Rules")
// ==========================================================
var DX_HEADERS = ['ใช้งาน (TRUE/FALSE)', 'มุมมอง (ด้านหน้า/ด้านหลัง/ด้านข้าง — ว่าง = ทุกมุม)', 'ส่วนที่วิเคราะห์ (ทั้งตัว/ส่วนบน/ส่วนล่าง — ว่าง = ทุกส่วน)',
  'รายการที่ตรวจ (ศีรษะ, คอ, ไหล่, ศอก, ข้อมือ, แนวสันหลัง, สะโพก, เข่า, ข้อเท้า)', 'คำที่ต้องพบในผล', 'องศาขั้นต่ำ',
  'อาการ/ข้อสรุป', 'โรคทางแผนไทย', 'ชื่อแผนปัจจุบัน', 'รหัสโรค', 'หัตถการแนะนำ (คั่นด้วย ,)', 'คำแนะนำผู้ป่วย'];

function postureDxSheet_() {
  var ss = getSS(), sh = ss.getSheetByName('Posture_Dx_Rules');
  if (!sh) {
    sh = ss.insertSheet('Posture_Dx_Rules');
    sh.appendRow(DX_HEADERS);
    sh.appendRow([true, 'ด้านหลัง', 'ส่วนบน', 'ไหล่', 'ไหล่ขวาต่ำกว่า', 4, 'บ่าขวาตึง', 'ลมปลายปัตคาตสัญญาณ 4 หลัง', 'Office syndrome', '',
      'นวดรักษา, ประคบสมุนไพร', 'ยืดเหยียดกล้ามเนื้อบ่า-คอทุก 1 ชั่วโมงระหว่างทำงาน']);
    sh.setFrozenRows(1);
  }
  return sh;
}
function getPostureDxRules() {
  try {
    var d = postureDxSheet_().getDataRange().getValues(), out = [];
    for (var i = 1; i < d.length; i++) {
      var r = d[i];
      if (!isActive_(r[0])) continue;
      if (!S_(r[3]) && !S_(r[4]) && !S_(r[6]) && !S_(r[7])) continue;
      out.push({ view: S_(r[1]), region: S_(r[2]), item: S_(r[3]), keyword: S_(r[4]), minDeg: parseFloat(r[5]) || 0,
        symptom: S_(r[6]), thaiDx: S_(r[7]), modernDx: S_(r[8]), code: S_(r[9]), treatment: S_(r[10]), advice: S_(r[11]) });
    }
    return out;
  } catch (e) { return []; }
}

// ==========================================================
// 📐 Posture Tracking
// ==========================================================
function searchPatientForPosture(searchKey) {
  try {
    var sheet = getSS().getSheetByName('Patient_DB');
    if (!sheet) return { success: false, message: "ไม่พบชีต 'Patient_DB'" };
    var data = sheet.getDataRange().getValues();
    var key = String(searchKey || "").trim().toLowerCase();
    if (!key) return { success: false, message: "กรุณาระบุคำค้นหา" };
    var keyPad = pzNorm_(key), matches = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i], cn = pzNorm_(r[0]), name = String(r[3] || "").trim();
      if (!cn && !name) continue;
      var phone = String(r[6] || "").replace(/'/g, "").trim(), idCard = String(r[11] || "").replace(/'/g, "").trim();
      var hay = [cn, r[1], r[3], r[4], phone, idCard].join("|").toLowerCase();
      if (cn === keyPad || hay.indexOf(key) > -1) {
        matches.push({
          cn: cn, prefix: String(r[2] || "").trim(), name: name, nickname: String(r[4] || "").trim(),
          gender: String(r[7] || "-").trim(), weight: r[8] || "-", height: r[9] || "-", phone: phone || "-",
          dob: pzDate_(r[5]), idCard: idCard || "-", address: String(r[12] || "-").trim(),
          relName: String(r[13] || "-").trim(), relPhone: String(r[14] || "-").trim(),
          disease: String(r[15] || "-").trim(), medicalPlan: String(r[16] || "-").trim(),
          cc: "แพ้: " + (r[15] || "-") + " | สิทธิ: " + (r[16] || "-")
        });
        if (matches.length >= 30) break;
      }
    }
    if (matches.length === 1) return { success: true, data: matches[0] };
    if (matches.length > 1) return { success: false, list: matches, message: "พบหลายรายการ" };
    return { success: false, message: "ไม่พบข้อมูลคนไข้ที่ตรงกับ: " + searchKey };
  } catch (e) { return { success: false, message: "Server Error: " + e.message }; }
}

var PTH_HEADERS = ["CN", "ชื่อ-นามสกุล", "วันที่", "ศีรษะ/คอ", "แนวสันหลัง", "ไหล่", "ศอก", "ข้อมือ", "สะโพก", "เข่า", "ข้อเท้า",
  "URL รูปภาพ", "มุมมอง", "ส่วนที่วิเคราะห์", "รายละเอียดเพิ่มเติม", "ครั้งที่", "หมายเหตุครั้งนี้", "รหัสระเบียน", "ภาพย่อ (Base64)"];
var PTH_ID_COL = 17, PTH_THUMB_COL = 18;

function pthSheet_() {
  var ss = getSS(), sheet = ss.getSheetByName("Posture_Tracking_History");
  if (!sheet) {
    sheet = ss.insertSheet("Posture_Tracking_History");
    sheet.appendRow(PTH_HEADERS); sheet.setColumnWidth(12, 220); sheet.setColumnWidth(19, 120);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(PTH_HEADERS);
  } else {
    var cols = sheet.getLastColumn();
    if (cols < PTH_HEADERS.length) sheet.getRange(1, cols + 1, 1, PTH_HEADERS.length - cols).setValues([PTH_HEADERS.slice(cols)]);
  }
  return sheet;
}
function saveBodyTracking(cn, name, base64Image, angles) {
  try {
    var folder = pzFolder_(readSettings_()[SET_KEYS.postureFolder], "AphaiWellness_Posture_Images");
    var parts = String(base64Image).split(";base64,");
    if (parts.length < 2) return { success: false, message: "รูปภาพไม่ถูกต้อง" };
    var contentType = parts[0].split(":")[1] || "image/jpeg";
    var blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), contentType, "Posture_" + pzNorm_(cn) + "_" + new Date().getTime() + ".jpg");
    var file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}

    var sheet = pthSheet_();
    var t = function (k) { return (angles && angles[k] && angles[k].text) ? angles[k].text : "-"; };
    var thumb = (angles && angles.thumb) ? String(angles.thumb) : "";
    if (thumb.length > 48000) thumb = "";
    sheet.appendRow([
      "'" + pzNorm_(cn), name, Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"),
      t('head'), t('spine'), t('shoulder'), t('elbow'), t('wrist'), t('hip'), t('knee'), t('ankle'),
      pzThumb_(file.getId()),
      (angles && angles.view) ? angles.view : "ด้านหน้า", (angles && angles.region) ? angles.region : "ทั้งตัว",
      (angles && angles.extra) ? angles.extra : "",
      (angles && angles.session) ? (parseInt(angles.session, 10) || "") : "",
      (angles && angles.note) ? String(angles.note).trim() : "",
      Utilities.getUuid(), thumb
    ]);
    return { success: true, url: pzThumb_(file.getId()) };
  } catch (err) { return { success: false, message: err.message }; }
}
function getTrackingHistory(cn) {
  try {
    var sheet = getSS().getSheetByName("Posture_Tracking_History");
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0] || "").trim() !== "" && (data[r].length <= PTH_ID_COL || !String(data[r][PTH_ID_COL] || "").trim())) {
        var nid = Utilities.getUuid();
        while (data[r].length <= PTH_ID_COL) data[r].push("");
        data[r][PTH_ID_COL] = nid;
        sheet.getRange(r + 1, PTH_ID_COL + 1).setValue(nid);
      }
    }
    var target = pzNorm_(cn), out = [];
    for (var i = 1; i < data.length; i++) {
      if (pzNorm_(data[i][0]) !== target) continue;
      var raw = String(data[i][11] || ""), fid = pzIdFromUrl_(raw), thumb = String(data[i][PTH_THUMB_COL] || "");
      out.push({
        id: String(data[i][PTH_ID_COL] || ""), cn: pzNorm_(data[i][0]), name: data[i][1],
        date: (data[i][2] instanceof Date) ? Utilities.formatDate(data[i][2], "GMT+7", "yyyy-MM-dd HH:mm") : String(data[i][2] || "-"),
        head: data[i][3], spine: data[i][4], shoulder: data[i][5], elbow: data[i][6],
        wrist: data[i][7], hip: data[i][8], knee: data[i][9], ankle: data[i][10],
        imageUrl: thumb ? ("data:image/jpeg;base64," + thumb) : (fid ? pzThumb_(fid) : raw),
        openUrl: fid ? ("https://drive.google.com/file/d/" + fid + "/view") : raw,
        view: String(data[i][12] || ""), region: String(data[i][13] || ""), extra: String(data[i][14] || ""),
        session: data[i][15] || "", note: String(data[i][16] || "")
      });
    }
    return out;
  } catch (err) { return []; }
}

// ==========================================================
// 🗑️ ลบระเบียน Posture Tracking ตาม ID
// ==========================================================
function deleteTrackingRecords(cn, ids) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var sheet = getSS().getSheetByName("Posture_Tracking_History");
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: "ไม่พบชีตประวัติสรีระ" };
    var idSet = {};
    (ids || []).forEach(function (id) { if (id) idSet[String(id)] = true; });
    if (!Object.keys(idSet).length) return { success: false, message: "ไม่พบรหัสระเบียนที่ต้องการลบ" };
    var data = sheet.getDataRange().getValues(), target = pzNorm_(cn), rows = [];
    for (var i = 1; i < data.length; i++) {
      if (pzNorm_(data[i][0]) === target && idSet[String(data[i][PTH_ID_COL] || "")]) {
        rows.push(i + 1);
        try { var fid = pzIdFromUrl_(String(data[i][11] || "")); if (fid) DriveApp.getFileById(fid).setTrashed(true); } catch (e) {}
      }
    }
    if (!rows.length) return { success: false, message: "ไม่พบระเบียนที่ตรงกับรหัสที่ส่งมา" };
    rows.sort(function (a, b) { return b - a; });
    rows.forEach(function (r) { sheet.deleteRow(r); });
    return { success: true, deleted: rows.length };
  } catch (err) { return { success: false, message: err.message }; }
  finally { try { lock.releaseLock(); } catch (e2) {} }
}

// ==========================================================
// 📍 ฐานข้อมูลโรคตามจุดปวด (ชีต "Pain_Map")
// ==========================================================
var PAIN_HEADERS = ['ใช้งาน (TRUE/FALSE)', 'ส่วนของร่างกาย (ศีรษะ, คอ, บ่า, ไหล่, ต้นแขน, ศอก, ข้อมือ, หลังส่วนบน, อก, เอว, ท้อง, สะโพก, ต้นขา, เข่า, น่อง, ข้อเท้า)',
  'ข้าง (ซ้าย/ขวา — ว่าง = ทั้งสองข้าง)', 'อาการ/ข้อสรุป', 'โรคทางแผนไทย', 'ชื่อแผนปัจจุบัน', 'รหัสโรค', 'หัตถการแนะนำ (คั่นด้วย ,)', 'คำแนะนำผู้ป่วย'];
var PAIN_DEFAULTS = [
  [true, 'ศีรษะ', '', 'ปวดศีรษะ', 'ลมปะกัง', 'Tension-type headache', '', 'นวดรักษา, ประคบสมุนไพร', 'พักผ่อนให้เพียงพอ ลดความเครียด'],
  [true, 'คอ', '', 'ปวดตึงคอ', 'ลมปลายปัตคาต', 'Neck pain', '', 'นวดรักษา, ประคบสมุนไพร', 'หลีกเลี่ยงการก้มใช้มือถือนาน ๆ'],
  [true, 'บ่า', '', 'ปวดตึงบ่า', 'ลมปลายปัตคาตสัญญาณ 4 หลัง', 'Office syndrome / Myofascial pain', '', 'นวดรักษา, ประคบสมุนไพร, ครอบแก้ว', 'ยืดเหยียดบ่า-คอทุก 1 ชั่วโมงระหว่างทำงาน'],
  [true, 'ไหล่', '', 'ปวดไหล่', 'ลมปลายปัตคาต', 'Shoulder pain', '', 'นวดรักษา, ประคบสมุนไพร', ''],
  [true, 'หลังส่วนบน', '', 'ปวดสะบัก/หลังส่วนบน', 'ลมปลายปัตคาต', 'Upper back pain', '', 'นวดรักษา, ประคบสมุนไพร', ''],
  [true, 'เอว', '', 'ปวดหลังส่วนล่าง', 'ลมปลายปัตคาตสัญญาณ 1 หลัง', 'Low back pain', '', 'นวดรักษา, ประคบสมุนไพร', 'หลีกเลี่ยงการยกของหนัก'],
  [true, 'สะโพก', '', 'ปวดสะโพก', 'ลมปลายปัตคาต', 'Hip pain', '', 'นวดรักษา', ''],
  [true, 'เข่า', '', 'ปวดเข่า', 'จับโปงแห้งเข่า', 'Knee osteoarthritis', '', 'นวดรักษา, พอกยา', 'หลีกเลี่ยงการนั่งยอง/พับเพียบนาน ๆ'],
  [true, 'ศอก', '', 'ปวดข้อศอก', '', 'Elbow pain', '', 'นวดรักษา', ''],
  [true, 'ข้อมือ', '', 'ปวดข้อมือ/มือชา', '', 'Wrist pain', '', 'นวดรักษา', ''],
  [true, 'ข้อเท้า', '', 'ปวดข้อเท้า', '', 'Ankle pain', '', 'ประคบสมุนไพร', '']
];
function painMapSheet_() {
  var ss = getSS(), sh = ss.getSheetByName('Pain_Map');
  if (!sh) {
    sh = ss.insertSheet('Pain_Map');
    sh.appendRow(PAIN_HEADERS);
    sh.getRange(2, 1, PAIN_DEFAULTS.length, PAIN_HEADERS.length).setValues(PAIN_DEFAULTS);
    sh.setFrozenRows(1);
  }
  return sh;
}
function getPainMap() {
  try {
    var d = painMapSheet_().getDataRange().getValues(), out = [];
    for (var i = 1; i < d.length; i++) {
      var r = d[i];
      if (!isActive_(r[0]) || !S_(r[1])) continue;
      out.push({ part: S_(r[1]), side: S_(r[2]), symptom: S_(r[3]), thaiDx: S_(r[4]), modernDx: S_(r[5]),
        code: S_(r[6]), treatment: S_(r[7]), advice: S_(r[8]) });
    }
    return out;
  } catch (e) { return []; }
}

// ==========================================================
// 🧘 วิธีการรักษา / วิธีนวดตามลำดับ (ชีต "Treatment_Methods")
// ==========================================================
var TM_HEADERS = ['ใช้งาน (TRUE/FALSE)', 'รหัสโรค (เช่น U57.20)', 'ชื่อโรค (ใส่เมื่อไม่มีรหัส)', 'ลำดับ (1, 2, 3...)',
  'ชื่อวิธีนวด/ท่า', 'รายละเอียดขั้นตอน', 'เวลา (นาที)', 'ข้อควรระวัง / อุปกรณ์', 'ลิงก์ภาพ/วิดีโอประกอบ (ไม่บังคับ)'];

function treatmentSheet_() {
  var ss = getSS(), sh = ss.getSheetByName('Treatment_Methods');
  if (!sh) {
    sh = ss.insertSheet('Treatment_Methods');
    sh.appendRow(TM_HEADERS);
    sh.appendRow([false, 'U57.20', 'ลมปลายปัตคาตบ่า', 1, 'ตัวอย่าง: ชื่อท่านวดที่ 1',
      'ใส่รายละเอียดขั้นตอนที่นี่ — เปลี่ยนคอลัมน์ A เป็น TRUE เพื่อให้แถวนี้แสดงในหน้าบันทึกการรักษา', 10, '', '']);
    sh.setFrozenRows(1);
  }
  return sh;
}
/** items = [{code, name}] ของโรคที่เลือก → คืนวิธีนวดของแต่ละโรคเรียงตามลำดับ */
function getTreatmentMethods(items) {
  try {
    var d = treatmentSheet_().getDataRange().getValues(), rows = [];
    for (var i = 1; i < d.length; i++) {
      var r = d[i];
      if (!isActive_(r[0])) continue;
      var title = S_(r[4]), detail = S_(r[5]);
      if (!title && !detail) continue;
      rows.push({ code: S_(r[1]).toUpperCase(), name: S_(r[2]).toLowerCase(), order: parseFloat(r[3]) || 9999,
        title: title, detail: detail, minutes: S_(r[6]), note: S_(r[7]), media: S_(r[8]) });
    }
    return (items || []).map(function (it) {
      var code = S_(it && it.code).toUpperCase(), name = S_(it && it.name).toLowerCase();
      var ms = rows.filter(function (m) {
        return (m.code && code && m.code === code) || (!m.code && m.name && name && m.name === name);
      }).sort(function (a, b) { return a.order - b.order; });
      return { code: S_(it && it.code), name: S_(it && it.name), methods: ms };
    });
  } catch (e) { return []; }
}

// ==========================================================
// 🎵 เพลงผ่อนคลาย (ชีต "Music")
// ==========================================================
function musicSheet_() {
  var ss = getSS(), sh = ss.getSheetByName('Music');
  if (!sh) { sh = ss.insertSheet('Music'); sh.appendRow(['ชื่อเพลง', 'URL (Google Drive)', 'หมายเหตุ']); sh.setFrozenRows(1); }
  return sh;
}
function getMusicList() {
  try {
    var sh = musicSheet_();
    if (sh.getLastRow() < 2) return [];
    return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
      .filter(function (r) { return /^https?:\/\//i.test(S_(r[1])); })
      .map(function (r) { return { name: S_(r[0]) || S_(r[1]), url: S_(r[1]), note: S_(r[2]) }; });
  } catch (e) { return []; }
}
function adminSaveMusic(token, list) {
  requireAdmin_(token);
  var rows = [], all = list || [];
  all.forEach(function (m) {
    var url = S_(m && m.url);
    if (/^https?:\/\//i.test(url)) rows.push([S_(m.name) || url, url, S_(m.note)]);
  });
  var sh = musicSheet_();
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
  var skipped = all.filter(function (m) { return m && (S_(m.url) || S_(m.name)); }).length - rows.length;
  return '✅ บันทึกรายการเพลง ' + rows.length + ' รายการ' + (skipped > 0 ? ' (ข้าม ' + skipped + ' รายการที่ลิงก์ไม่ขึ้นต้นด้วย http/https)' : '');
}
/** สำรองกรณีเล่นจากลิงก์ Drive ตรง ๆ ไม่ได้ — จำกัดเฉพาะไฟล์ที่อยู่ในรายการเพลง ขนาดไม่เกิน 18 MB */
function getMusicData(fileId) {
  try {
    var id = pzIdFromUrl_(fileId), ok = false;
    getMusicList().forEach(function (m) { if (pzIdFromUrl_(m.url) === id) ok = true; });
    if (!id || !ok) return { ok: false, message: 'ไฟล์นี้ไม่อยู่ในรายการเพลงที่แอดมินตั้งไว้' };
    var f = DriveApp.getFileById(id);
    if (f.getSize() > 18 * 1024 * 1024) return { ok: false, message: 'ไฟล์ใหญ่เกิน 18 MB — กรุณาใช้ไฟล์ MP3 ที่เล็กลง' };
    var blob = f.getBlob();
    return { ok: true, mime: blob.getContentType() || 'audio/mpeg', b64: Utilities.base64Encode(blob.getBytes()) };
  } catch (e) { return { ok: false, message: e.message }; }
}

// ==========================================================
// 📋 ดึงการรักษาครั้งก่อน (ใช้ในหน้าบันทึกการรักษา)
// ==========================================================
function getLastVisit(cn, beforeRow) {
  try {
    var sh = getSS().getSheetByName('Visit_Logs');
    if (!sh || sh.getLastRow() < 2) return null;
    var target = pzNorm_(cn), limit = parseInt(beforeRow, 10) || 0;
    var d = sh.getRange(2, 1, sh.getLastRow() - 1, 24).getValues();
    var T = function (v) { var s = S_(v).replace(/^'/, ''); return s === '-' ? '' : s; };
    for (var i = d.length - 1; i >= 0; i--) {
      var row = i + 2, v = d[i];
      if (limit && row >= limit) continue;
      if (pzNorm_(v[2]) !== target) continue;
      var dt = (v[0] instanceof Date) ? Utilities.formatDate(v[0], "GMT+7", "dd/MM/yyyy") : T(v[0]);
      return {
        row: row, date: dt, cc: T(v[10]), pi: T(v[11]), ph: T(v[12]), pe: T(v[13]),
        dx: T(v[14]), treatment: T(v[15]), advice: T(v[16]), plan: T(v[17]),
        staff: T(v[19]), prePain: T(v[22]), postPain: T(v[23])
      };
    }
    return null;
  } catch (e) { return null; }
}
// ==========================================================
// 📍 บันทึกภาพมาร์กจุดเจ็บ/อาการชา (Pain Map)
// ==========================================================
function savePainMapImage(cn, base64Image) {
  try {
    var folder = pzFolder_("", "AphaiWellness_PainMaps");
    var parts = String(base64Image).split(";base64,");
    if (parts.length < 2) return { success: false, message: "ข้อมูลรูปภาพไม่ถูกต้อง" };
    var blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), "image/png", "PainMap_" + pzNorm_(cn) + "_" + new Date().getTime() + ".png");
    var file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    return { success: true, url: pzThumb_(file.getId()) };
  } catch (err) { return { success: false, message: err.message }; }
}
function getBase64ImageFromUrl(url) {
  try {
    var fileId = pzIdFromUrl_(url);
    var blob = fileId ? DriveApp.getFileById(fileId).getBlob() : UrlFetchApp.fetch(url).getBlob();
    return "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    return "";
  }
}
