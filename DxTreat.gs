// ==========================================================
// 🩺 DxTreat.gs — ฐานข้อมูลโรค + วิธีการนวด + เพลงผ่อนคลาย
// ใช้ฟังก์ชันกลางจาก Code.gs: getSS, S_, readSettings_, writeSettingsBatch_, requireAdmin_, pzIdFromUrl_
// ==========================================================
var DXT_KEY_SHEET = 'ฐานข้อมูลโรค: ชื่อชีต';
var DXT_KEY_MUSIC = 'เพลงผ่อนคลาย (ลิงก์ Google Drive)';
var DXT_DEFAULT_SHEET = 'Dx_Treatment';
var DXT_HEADERS = ['ใช้งาน (TRUE/FALSE)', 'คำค้น / ส่วนของร่างกาย', 'ชื่อโรคแผนไทย', 'ชื่อโรคแผนปัจจุบัน', 'รหัสโรค',
  'วิธีการนวด ขั้นที่ 1', 'วิธีการนวด ขั้นที่ 2', 'วิธีการนวด ขั้นที่ 3', 'วิธีการนวด ขั้นที่ 4', 'วิธีการนวด ขั้นที่ 5',
  'หัตถการเสริม (คั่นด้วย ,)', 'คำแนะนำผู้ป่วย', 'ลิงก์รูป/วิดีโอประกอบ'];

function dxtSheetName_() { return S_(readSettings_()[DXT_KEY_SHEET]) || DXT_DEFAULT_SHEET; }

function dxtSheet_(create) {
  var ss = getSS(), name = dxtSheetName_(), sh = ss.getSheetByName(name);
  if (!sh && create) {
    sh = ss.insertSheet(name);
    sh.appendRow(DXT_HEADERS);
    sh.appendRow([true, 'บ่า, คอ, สะบัก', 'ลมปลายปัตคาตสัญญาณ 4 หลัง', 'Office syndrome / Myofascial pain syndrome', '',
      '(ตัวอย่าง — แก้เป็นวิธีของคลินิก) จัดท่าผู้ป่วย นวดพื้นฐานบ่าและคอ', '(ตัวอย่าง) กดจุดสัญญาณ 4 หลัง', '(ตัวอย่าง) นวดคลายกล้ามเนื้อบ่า-สะบัก แล้วประคบสมุนไพร', '', '',
      'ประคบสมุนไพร', 'ยืดเหยียดบ่า-คอทุก 1 ชั่วโมงระหว่างทำงาน', '']);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** หาตำแหน่งคอลัมน์จากหัวตาราง (รองรับชีตที่ตั้งชื่อหัวต่างกัน) */
function dxtCols_(head) {
  var c = { active: -1, kw: -1, thai: -1, modern: -1, code: -1, steps: [], extra: -1, advice: -1, media: -1 };
  head.forEach(function (h, i) {
    var t = S_(h).toLowerCase();
    if (!t) return;
    if (/ใช้งาน|active/.test(t)) { if (c.active < 0) c.active = i; }
    else if (/ขั้น|step|^\d+\s*[.)]?$/.test(t) || (/วิธี/.test(t) && /นวด|รักษา/.test(t))) c.steps.push(i);
    else if (/ปัจจุบัน|modern|english|อังกฤษ|diagnos/.test(t)) { if (c.modern < 0) c.modern = i; }
    else if (/รหัส|icd|code/.test(t)) { if (c.code < 0) c.code = i; }
    else if (/คำค้น|ส่วน|ตำแหน่ง|keyword|อาการ/.test(t)) { if (c.kw < 0) c.kw = i; }
    else if (/หัตถการ/.test(t)) { if (c.extra < 0) c.extra = i; }
    else if (/คำแนะนำ|advice/.test(t)) { if (c.advice < 0) c.advice = i; }
    else if (/ลิงก์|link|url|รูป|วิดีโอ|video/.test(t)) { if (c.media < 0) c.media = i; }
    else if (/โรค|แผนไทย|thai|ชื่อ/.test(t)) { if (c.thai < 0) c.thai = i; }
  });
  return c;
}
function dxtSplit_(t) {
  var parts = String(t).split(/\r?\n+/);
  if (parts.length < 2) parts = String(t).split(/\s+(?=(?:ขั้นที่\s*)?\d+[.)]\s)/);
  return parts.map(function (s) { return S_(s).replace(/^(?:ขั้นที่\s*)?\d+\s*[.):-]\s*/, ''); }).filter(Boolean);
}
function dxtRead_(sh) {
  var out = { cols: null, list: [], byPos: false };
  if (!sh || sh.getLastRow() < 1) return out;
  var data = sh.getDataRange().getDisplayValues(), c = dxtCols_(data[0]);
  if (c.thai < 0 && c.modern < 0) {             // หาหัวคอลัมน์ไม่เจอ → ใช้ตำแหน่งมาตรฐาน
    c = { active: 0, kw: 1, thai: 2, modern: 3, code: 4, steps: [5, 6, 7, 8, 9], extra: 10, advice: 11, media: 12 };
    out.byPos = true;
  }
  out.cols = c;
  var g = function (r, i) { return i > -1 ? S_(r[i]) : ''; };
  for (var i = 1; i < data.length; i++) {
    var r = data[i], act = g(r, c.active);
    if (/^(false|no|ไม่|ไม่ใช้|ปิด)$/i.test(act)) continue;
    var thai = g(r, c.thai), modern = g(r, c.modern);
    if (!thai && !modern) continue;
    var steps = [];
    c.steps.forEach(function (k) { var v = g(r, k); if (v) steps.push(v); });
    if (steps.length === 1) steps = dxtSplit_(steps[0]);
    out.list.push({ row: i + 1, thai: thai, modern: modern, code: g(r, c.code), kw: g(r, c.kw), steps: steps,
      extra: g(r, c.extra), advice: g(r, c.advice), media: g(r, c.media) });
  }
  return out;
}
function dxtInfo_(sh) {
  var r = dxtRead_(sh), c = r.cols || { steps: [] }, miss = [];
  if (c.thai < 0) miss.push('ชื่อโรคแผนไทย');
  if (c.modern < 0) miss.push('ชื่อโรคแผนปัจจุบัน');
  if (!c.steps.length) miss.push('ขั้นตอนการนวด');
  return 'พบ ' + r.list.length + ' โรค · คอลัมน์ขั้นตอนการนวด ' + c.steps.length + ' คอลัมน์'
    + (miss.length ? ' · ⚠️ ไม่พบคอลัมน์: ' + miss.join(', ') : '') + (r.byPos ? ' (ใช้ตำแหน่งคอลัมน์มาตรฐาน)' : '');
}

/** หน้าบันทึกการรักษาเรียกใช้: รายชื่อโรคทั้งหมด */
function getDxTreatList() {
  var sh = dxtSheet_(true);
  return { sheet: sh.getName(), list: dxtRead_(sh).list };
}

// ---------- เพลงผ่อนคลาย ----------
function dxtMusicRaw_() {
  return S_(readSettings_()[DXT_KEY_MUSIC]).split(/\r?\n/).map(S_).filter(Boolean).map(function (ln, i) {
    var name = '', url = ln, k = ln.indexOf('|');
    if (k > -1) { name = S_(ln.substring(0, k)); url = S_(ln.substring(k + 1)); }
    var id = /drive\.google|docs\.google/.test(url) ? pzIdFromUrl_(url) : '';
    return { i: i, name: name, id: id, url: url };
  }).filter(function (t) { return t.id || /^https?:\/\//i.test(t.url); });
}
function getMusicList() {
  return dxtMusicRaw_().map(function (t, n) {
    if (!t.name && t.id) { try { t.name = DriveApp.getFileById(t.id).getName().replace(/\.[^.]+$/, ''); } catch (e) {} }
    if (!t.name) t.name = 'เพลงที่ ' + (n + 1);
    return t;
  });
}
/** สำรอง: ดึงไฟล์เสียงจาก Drive ผ่านเซิร์ฟเวอร์ (เมื่อเล่นจากลิงก์ตรงไม่ได้) */
function getMusicData(id) {
  var ok = dxtMusicRaw_().some(function (t) { return t.id === id; });
  if (!ok) return { ok: false, message: 'ไม่พบเพลงนี้ในการตั้งค่า' };
  try {
    var f = DriveApp.getFileById(id);
    if (f.getSize() > 20 * 1024 * 1024) return { ok: false, message: 'ไฟล์ใหญ่เกิน 20 MB — ใช้ไฟล์ mp3 ที่เล็กลง หรือแชร์ไฟล์แบบ "ทุกคนที่มีลิงก์"' };
    var b = f.getBlob(), ct = b.getContentType() || 'audio/mpeg';
    if (ct === 'application/octet-stream') ct = 'audio/mpeg';
    return { ok: true, dataUrl: 'data:' + ct + ';base64,' + Utilities.base64Encode(b.getBytes()) };
  } catch (e) { return { ok: false, message: e.message }; }
}

// ---------- หน้า Admin ----------
function adminGetDxtSettings(token) {
  requireAdmin_(token);
  var ss = getSS(), name = dxtSheetName_(), sh = ss.getSheetByName(name);
  return {
    sheet: name, exists: !!sh, info: sh ? dxtInfo_(sh) : '', sheetUrl: ss.getUrl(),
    sheets: ss.getSheets().map(function (s) { return { name: s.getName(), gid: s.getSheetId() }; }),
    music: S_(readSettings_()[DXT_KEY_MUSIC])
  };
}
function dxtCheckMusic_(text) {
  return String(text || '').split(/\r?\n/).map(S_).filter(Boolean).map(function (ln) {
    var url = ln.indexOf('|') > -1 ? S_(ln.substring(ln.indexOf('|') + 1)) : ln;
    var id = /drive\.google|docs\.google/.test(url) ? pzIdFromUrl_(url) : '';
    if (!id) return { line: ln, ok: /^https?:\/\//i.test(url), msg: /^https?:\/\//i.test(url) ? 'ลิงก์ภายนอก (ไม่ใช่ Drive)' : 'ลิงก์ไม่ถูกต้อง' };
    try {
      var f = DriveApp.getFileById(id), mt = f.getMimeType(), mb = (f.getSize() / 1048576).toFixed(1);
      var warn = /^audio\//.test(mt) ? '' : ' ⚠️ ไม่ใช่ไฟล์เสียง (' + mt + ')';
      return { line: ln, ok: !warn, msg: f.getName() + ' · ' + mb + ' MB' + warn };
    } catch (e) { return { line: ln, ok: false, msg: 'เปิดไฟล์ไม่ได้: ' + e.message }; }
  });
}
function adminCheckMusic(token, text) { requireAdmin_(token); return dxtCheckMusic_(text); }
function adminSaveDxtSettings(token, o) {
  requireAdmin_(token);
  o = o || {};
  var lines = String(o.music || '').split(/\r?\n/).map(S_).filter(Boolean), map = {};
  map[DXT_KEY_SHEET] = S_(o.sheet) || DXT_DEFAULT_SHEET;
  map[DXT_KEY_MUSIC] = lines.join('\n');
  writeSettingsBatch_(map, 'ตั้งค่าจากหน้า Admin (ฐานข้อมูลโรค / เพลงผ่อนคลาย)');
  var sh = dxtSheet_(true), msg = '✅ บันทึกเรียบร้อย\n• ชีต "' + sh.getName() + '": ' + dxtInfo_(sh);
  if (lines.length) {
    var chk = dxtCheckMusic_(lines.join('\n')), bad = chk.filter(function (x) { return !x.ok; }).length;
    msg += '\n• เพลง ' + lines.length + ' รายการ' + (bad ? ' (⚠️ มีปัญหา ' + bad + ' รายการ — กด "ตรวจสอบลิงก์เพลง" ดูรายละเอียด)' : ' ใช้งานได้ทั้งหมด');
  }
  return msg;
}