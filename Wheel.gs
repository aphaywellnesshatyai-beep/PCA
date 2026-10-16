// ==========================================================
// 🎡 Wheel.gs — ธาตุเจ้าเรือนตามแผ่นหมุน (แผ่นใหญ่ + แผ่นกลาง + แผ่นเล็ก)
// ไฟล์นี้เป็นไฟล์ใหม่ วางเพิ่มในโปรเจกต์ได้เลย (ไม่ต้องแก้ Code.gs)
// ใช้ฟังก์ชันกลางจาก Code.gs: pzDobParts_, S_, readSettings_, writeSettingsBatch_, requireAdmin_, settingSheet_, getSS
//
// หลักการ (ตรวจกับรูปตัวอย่างจริง 2 ราย: 27 ส.ค. 2541 และ 19 ม.ค. 2545 แล้วตรงกับที่ช่องแสดงสี):
//  1) ปีเกิด พ.ศ. → ตารางด้านหลังแผ่นใหญ่ → วัน/เดือน D
//  2) หมุนแผ่นกลางให้ลูกศรชี้ D บนสเกลวันที่ของแผ่นใหญ่ (365 ช่อง = 360°, 1 ม.ค. อยู่บนสุด, ตามเข็มนาฬิกา)
//  3) หมุนแผ่นเล็กให้ลูกศรซ้ายชี้วันเกิด (ลูกศรบน = +90° = ฝั่งวันปฏิสนธิ)
//  4) ช่องสามเหลี่ยมที่อยู่ตรงลูกศรบน เปิดให้เห็นสีของแผ่นกลาง 4 วง → อ่านค่า
// มุมทั้งหมดวัดจากรูปถ่ายแผ่นจริง (คลาดเคลื่อนประมาณ ±1°) และการหมุนแผ่นจริงด้วยมือคลาดได้ราว 2–5 วัน จึงแจ้ง "ใกล้ขอบสี" เมื่อแกนช่องห่างเส้นแบ่งไม่ถึง 3°
// ==========================================================

// ตารางด้านหลังแผ่นใหญ่ พ.ศ. 2475–2590 เรียงตามปี แต่ละช่อง = [วัน, เดือน]
// ⚠️ ตามที่อ่านจากรูป พ.ศ. 2483 และ 2484 เป็น 13 มี.ค. ทั้งคู่ — โปรดเทียบกับแผ่นจริงอีกครั้ง
// แก้เฉพาะปีที่ผิดได้โดยสร้างชีต "Wheel_Years" (คอลัมน์ A=พ.ศ., B=วัน, C=เดือน) ระบบจะใช้ค่าในชีตก่อน
var WHEEL_YEARS_START = 2475;
var WHEEL_YEARS = [
  [11,3],[28,2],[20,3],[8,3],[25,2],[16,3],[5,3],[24,3],[13,3],[13,3],[2,3],[21,3],[9,3],
  [26,2],[18,3],[7,3],[25,3],[14,3],[3,3],[23,3],[11,3],[28,2],[19,3],[8,3],[26,2],[16,3],[5,3],[24,3],[12,3],[1,3],[21,3],[10,3],[27,2],[17,3],
  [6,3],[26,3],[14,3],[3,3],[22,3],[11,3],[29,2],[19,3],[8,3],[27,3],[15,3],[4,3],[24,3],[13,3],[1,3],[21,3],[10,3],[27,2],[17,3],[6,3],[25,3],[14,3],[3,3],[22,3],
  [11,3],[28,2],[19,3],[8,3],[27,3],[16,3],[4,3],[24,3],[13,3],[2,3],[20,3],[9,3],[27,2],[18,3],[6,3],[25,3],[14,3],[3,3],[22,3],[11,3],[28,2],[19,3],[7,3],[27,3],
  [16,3],[5,3],[23,3],[12,3],[2,3],[21,3],[9,3],[26,2],[18,3],[7,3],[25,3],[14,3],[3,3],[22,3],[10,3],[28,2],[19,3],[8,3],[26,3],[15,3],[5,3],
  [24,3],[12,3],[1,3],[21,3],[10,3],[27,2],[17,3],[6,3],[26,3],[14,3],[3,3],[22,3],[11,3]
];

// สีที่ใช้วาดบนหน้าเว็บ (ชื่อตามตำนานสีบนแผ่นเล็ก)
var WHEEL_PALETTE = {
  pink:   { name: 'พัทธปิตตะ',  hex: '#e79bd0' },
  orange: { name: 'อพัทธปิตตะ', hex: '#ee9462' },
  red:    { name: 'กำเดา',      hex: '#a8404f' },
  yellow: { name: 'หทัยวาตะ',   hex: '#f3ea3a' },
  lgreen: { name: 'สัตถกวาตะ',  hex: '#8cc63f' },
  green:  { name: 'สุมนาวาตะ',  hex: '#2e8b4a' },
  lblue:  { name: 'คอเสมหะ',    hex: '#44a5e8' },
  indigo: { name: 'อุระเสมหะ',  hex: '#3b4a9e' },
  purple: { name: 'คูถเสมหะ',   hex: '#6a3a86' },
  black:  { name: 'สีดำ',       hex: '#26262b' }
};

// เรขาคณิตแผ่นกลาง (หน่วย: แผ่นกลางรัศมี 100) มุมวัดตามเข็มจากลูกศรบนแผ่นกลาง
var WHEEL_GEOM = {
  discR: 100, smallR: 95.4, bigR: 132,
  ringR: { ring1: [70.3, 80], ring2: [58.6, 70.3], ring3: [47, 58.6], center: [0, 47] },
  ring1: [[3.8, 92.0, 'red'], [92.0, 181.5, 'yellow'], [181.5, 270.0, 'lblue'], [270.0, 363.8, 'black']],
  ring2: [[0, 42.9, 'pink'], [42.9, 80.0, 'orange'], [80.0, 118.3, 'red'], [118.3, 157.8, 'yellow'], [157.8, 197.6, 'lgreen'],
          [197.6, 235.8, 'green'], [235.8, 274.0, 'lblue'], [274.0, 312.2, 'indigo'], [312.2, 360, 'purple']],
  ring3: [[0, 23.0, 'pink', 1], [23.0, 42.8, 'orange', 2], [42.8, 62.0, 'red', 3], [62.0, 79.8, 'pink', 4], [79.8, 98.2, 'orange', 5], [98.2, 118.2, 'red', 6],
          [118.2, 138.2, 'yellow', 7], [138.2, 157.8, 'lgreen', 8], [157.8, 180.0, 'green', 9], [180.0, 197.5, 'yellow', 10], [197.5, 216.2, 'lgreen', 11], [216.2, 235.8, 'green', 12],
          [235.8, 254.8, 'lblue', 13], [254.8, 274.5, 'indigo', 14], [274.5, 292.0, 'purple', 15], [292.0, 312.5, 'lblue', 16], [312.5, 333.0, 'indigo', 17], [333.0, 360, 'purple', 18]],
  center: [[0, 119.0, 'red'], [119.0, 235.2, 'yellow'], [235.2, 360, 'lblue']],
  // ช่องสามเหลี่ยมของแผ่นเล็ก: ปลายแหลมอยู่ห่างจุดศูนย์กลาง apexR, ขอบสองข้างเอียง halfDeg จากแกน (แกน = ลูกศรบน), ตัดที่รัศมี outerR
  win: { apexR: 22.9, halfDeg: 14.9, outerR: 82 },
  rows: { r1: 75.2, r2: 64.5, r3: 52.8, s: 35 }
};

var WHEEL_FIELDS = ['name', 'appearance', 'caution', 'mind', 'commonIllness', 'food', 'lifestyle', 'oil'];
var WHEEL_FIELD_HEADERS = ['ชื่อที่แสดง', 'ลักษณะทั่วไป', 'โรคที่ควรระวัง', 'ลักษณะจิตใจเมื่อกำเริบ', 'โรคที่พบบ่อย', 'การกินอาหาร', 'การปรับวิถีชีวิต', 'น้ำมันหอมระเหยที่แนะนำ'];
var WHEEL_PRIMARY_KEY = 'ธาตุเจ้าเรือน: แถวหลัก (r1/r2/r3/s)';

// รายการค่าที่ให้แอดมินตั้ง: [แถว, สี, ชื่อเริ่มต้น]
// r1 = วงนอก 4 สี (ไฟ/ลม/น้ำ/ดิน ตามตำนานข้างช่อง), r2 = วงที่ 2 (9 สี), r3 = วงที่ 3 ช่องเลข 1-18 (9 สี), s = วงกลาง ฤดูสมุฏฐาน (3 สี)
var WHEEL_ENTRIES = [
  ['r1', 'red', 'ไฟ'], ['r1', 'yellow', 'ลม'], ['r1', 'lblue', 'น้ำ'], ['r1', 'black', 'ดิน'],
  ['r2', 'pink', 'พัทธปิตตะ'], ['r2', 'orange', 'อพัทธปิตตะ'], ['r2', 'red', 'กำเดา'], ['r2', 'yellow', 'หทัยวาตะ'], ['r2', 'lgreen', 'สัตถกวาตะ'],
  ['r2', 'green', 'สุมนาวาตะ'], ['r2', 'lblue', 'คอเสมหะ'], ['r2', 'indigo', 'อุระเสมหะ'], ['r2', 'purple', 'คูถเสมหะ'],
  ['r3', 'pink', 'พัทธปิตตะ'], ['r3', 'orange', 'อพัทธปิตตะ'], ['r3', 'red', 'กำเดา'], ['r3', 'yellow', 'หทัยวาตะ'], ['r3', 'lgreen', 'สัตถกวาตะ'],
  ['r3', 'green', 'สุมนาวาตะ'], ['r3', 'lblue', 'คอเสมหะ'], ['r3', 'indigo', 'อุระเสมหะ'], ['r3', 'purple', 'คูถเสมหะ'],
  ['s', 'red', 'ร้อน'], ['s', 'yellow', 'ฝน'], ['s', 'lblue', 'หนาว']
];

// ---------- ตัวช่วย ----------
function wheelCum_() { return [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334]; }
/** วัน/เดือน → มุม (องศาตามเข็มจาก 1 ม.ค. บนสุด) บนสเกล 365 ช่อง; 29 ก.พ. อยู่กึ่งกลางระหว่าง 28 ก.พ. กับ 1 มี.ค. */
function wheelDateDeg_(m, d) {
  var pos = wheelCum_()[m - 1] + d;
  if (m === 2 && d >= 29) pos = 59.5;
  return (pos - 1) / 365 * 360;
}
function wheelYearTable_() {
  var t = {};
  WHEEL_YEARS.forEach(function (e, i) { t[WHEEL_YEARS_START + i] = [e[0], e[1]]; });
  try {                                                   // ค่าแก้ไขจากชีต Wheel_Years (ถ้ามี)
    var sh = getSS().getSheetByName('Wheel_Years');
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) {
        var be = parseInt(r[0], 10), d = parseInt(r[1], 10), m = parseInt(r[2], 10);
        if (be > 2300 && d >= 1 && d <= 31 && m >= 1 && m <= 12) t[be] = [d, m];
      });
    }
  } catch (e) {}
  return t;
}
function wheelNorm_(a) { return ((a % 360) + 360) % 360; }
function wheelFind_(ring, ang) {
  ang = wheelNorm_(ang);
  for (var i = 0; i < ring.length; i++) {
    if ((ang >= ring[i][0] && ang < ring[i][1]) || (ang + 360 >= ring[i][0] && ang + 360 < ring[i][1])) return i;
  }
  return 0;
}
function wheelOverlap_(ring, a0, a1) {
  var out = [], total = a1 - a0;
  for (var i = 0; i < ring.length; i++) {
    var ov = 0;
    for (var k = -360; k <= 360; k += 360) ov += Math.max(0, Math.min(a1, ring[i][1] + k) - Math.max(a0, ring[i][0] + k));
    if (ov > 0.05) out.push({ idx: i, color: ring[i][2], no: ring[i][3] || null, share: Math.round(ov / total * 100) });
  }
  out.sort(function (x, y) { return y.share - x.share; });
  return out;
}
/** ครึ่งความกว้างเชิงมุม (องศา) ของช่องที่รัศมี r — ขอบช่องเป็นเส้นตรงจากปลายแหลม */
function wheelWinHalf_(r) {
  var w = WHEEL_GEOM.win, h = w.halfDeg * Math.PI / 180, a = w.apexR;
  if (r <= a) return 0;
  var t = -a * Math.cos(h) + Math.sqrt(a * a * Math.cos(h) * Math.cos(h) - a * a + r * r);
  return Math.atan2(t * Math.sin(h), a + t * Math.cos(h)) * 180 / Math.PI;
}
function wheelDist_(ring, ang) {                           // ระยะ (องศา) จากแกนถึงเส้นแบ่งสีที่ใกล้สุดของวงนั้น
  var best = 999;
  ring.forEach(function (c) {
    [c[0], c[1]].forEach(function (b) {
      var d = Math.abs(((wheelNorm_(ang) - wheelNorm_(b) + 540) % 360) - 180);
      if (d < best) best = d;
    });
  });
  return best;
}

// ---------- ค่าที่แอดมินตั้ง ----------
function wheelDefaults_() {
  var o = {};
  WHEEL_ENTRIES.forEach(function (e) {
    var k = e[0] + '_' + e[1]; o[k] = {};
    WHEEL_FIELDS.forEach(function (f) { o[k][f] = (f === 'name') ? e[2] : ''; });
  });
  return o;
}
function wheelConfig_() {
  var cfg = wheelDefaults_();
  try {
    var sh = getSS().getSheetByName('Wheel_Config');
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 3 + WHEEL_FIELDS.length).getValues().forEach(function (r) {
        var k = S_(r[0]);
        if (!cfg[k]) return;
        WHEEL_FIELDS.forEach(function (f, i) { var v = S_(r[3 + i]); if (v || f !== 'name') cfg[k][f] = v; });
      });
    }
  } catch (e) {}
  return cfg;
}
function wheelPrimary_() {
  var p = S_(readSettings_()[WHEEL_PRIMARY_KEY]);
  return (p === 'r1' || p === 'r2' || p === 'r3' || p === 's') ? p : 'r1';
}

/** ข้อมูลทั้งหมดที่หน้าเว็บต้องใช้วาดแผ่นและแสดงผล */
function getWheelBundle() {
  var yrs = wheelYearTable_();
  return { geom: WHEEL_GEOM, palette: WHEEL_PALETTE, years: yrs, config: wheelConfig_(), primary: wheelPrimary_(),
           entries: WHEEL_ENTRIES, fields: WHEEL_FIELDS, fieldHeaders: WHEEL_FIELD_HEADERS,
           yearMin: WHEEL_YEARS_START, yearMax: WHEEL_YEARS_START + WHEEL_YEARS.length - 1 };
}

/** คำนวณตามขั้นตอนของแผ่นหมุน — รับวันเกิดแบบ yyyy-mm-dd (ค.ศ.) */
function getWheelReading(dobIso) {
  var p = pzDobParts_(dobIso);
  if (!p) return { ok: false, message: 'วันเกิดไม่ถูกต้อง' };
  var be = p.y + 543, tbl = wheelYearTable_(), e = tbl[be];
  if (!e) return { ok: false, message: 'ไม่มีข้อมูล พ.ศ. ' + be + ' ในตารางของแผ่น (รองรับ พ.ศ. ' + WHEEL_YEARS_START + '–' + (WHEEL_YEARS_START + WHEEL_YEARS.length - 1) + ')' };
  var G = WHEEL_GEOM, cfg = wheelConfig_();
  var thetaB = wheelDateDeg_(p.m, p.d), thetaD = wheelDateDeg_(e[1], e[0]);
  var midRot = thetaD;                                     // ลูกศรแผ่นกลางชี้วันเดือนตามตาราง
  var smallRot = wheelNorm_(thetaB - 270);                 // ลูกศรซ้ายแผ่นเล็ก (อยู่ที่ 270° ของแผ่นเล็ก) ชี้วันเกิด
  var conception = wheelNorm_(thetaB + 90);                // ลูกศรบน (ฝั่งวันปฏิสนธิ)
  var axis = wheelNorm_(smallRot - midRot);                // แกนช่องบนแผ่นกลาง (องศาตามเข็มจากลูกศรแผ่นกลาง)
  var defs = [['r1', 1, G.ring1], ['r2', 2, G.ring2], ['r3', 3, G.ring3], ['s', 4, G.center]];
  var rows = defs.map(function (d) {
    var ring = d[2], r = G.rows[d[0]], half = wheelWinHalf_(r);
    var i = wheelFind_(ring, axis), cell = ring[i], color = cell[2];
    var cov = wheelOverlap_(ring, axis - half, axis + half).map(function (c) {
      return { color: c.color, name: WHEEL_PALETTE[c.color].name, no: c.no, share: c.share };
    });
    var k = d[0] + '_' + color, c = cfg[k] || {};
    return { key: d[0], row: d[1], color: color, colorName: WHEEL_PALETTE[color].name, no: cell[3] || null,
             name: c.name || WHEEL_PALETTE[color].name, cfgKey: k, covered: cov,
             near: wheelDist_(ring, axis) < 3 };
  });
  return { ok: true,
    birth: { y: p.y, be: be, m: p.m, d: p.d },
    table: { d: e[0], m: e[1] },
    angles: { birth: thetaB, table: thetaD, midRot: midRot, smallRot: smallRot, conception: conception, axis: axis },
    rows: rows, primary: wheelPrimary_() };
}

/** รูปแบบเดียวกับ getElementInfo เดิม เพื่อให้หน้าบันทึกการรักษาสลับมาเรียกตัวนี้ได้ */
function getWheelElementInfo(dobIso) {
  var r = getWheelReading(dobIso);
  if (!r.ok) return null;
  var cfg = wheelConfig_(), prim = null;
  r.rows.forEach(function (x) { if (x.key === r.primary) prim = x; });
  var c = cfg[prim.cfgKey] || {};
  var oils = [];
  r.rows.forEach(function (x) { var o = S_((cfg[x.cfgKey] || {}).oil); if (o && oils.indexOf(o) === -1) oils.push(o); });
  return { dayOrder: r.table.d + '/' + r.table.m, element: prim.name,
    appearance: c.appearance || '', caution: c.caution || '', mind: c.mind || '', commonIllness: c.commonIllness || '',
    food: c.food || '', lifestyle: c.lifestyle || '', oil: oils.join(' / '),
    summary: r.rows.map(function (x) { return x.row + ':' + x.name + (x.no ? ' (ช่อง ' + x.no + ')' : ''); }).join(' | ') };
}

/** บันทึกความหมายของแต่ละสี (เฉพาะแอดมิน) cfg = { primary:'r1', entries:{ 'r1_red':{name,appearance,...}, ... } } */
function adminSaveWheelConfig(token, cfgIn) {
  requireAdmin_(token);
  cfgIn = cfgIn || {};
  var ent = cfgIn.entries || {}, lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss = getSS(), sh = ss.getSheetByName('Wheel_Config');
    if (!sh) sh = ss.insertSheet('Wheel_Config');
    sh.clear();
    var head = ['รหัส (ห้ามแก้)', 'แถว', 'สี'].concat(WHEEL_FIELD_HEADERS);
    var rows = WHEEL_ENTRIES.map(function (e) {
      var k = e[0] + '_' + e[1], v = ent[k] || {};
      return [k, e[0], WHEEL_PALETTE[e[1]].name].concat(WHEEL_FIELDS.map(function (f) { return S_(v[f]).substring(0, 2000); }));
    });
    var all = [head].concat(rows);
    sh.getRange(1, 1, all.length, head.length).setNumberFormat('@').setValues(all);
    sh.setFrozenRows(1);
    var prim = (['r1', 'r2', 'r3', 's'].indexOf(cfgIn.primary) > -1) ? cfgIn.primary : 'r1';
    writeSettingsBatch_((function () { var o = {}; o[WHEEL_PRIMARY_KEY] = prim; return o; })(), 'แถวของแผ่นธาตุที่ใช้เป็นธาตุเจ้าเรือนหลัก');
    return '✅ บันทึกการตั้งค่าแผ่นธาตุเรียบร้อย';
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

/** โลโก้คลินิกจากหน้าแอดมินในรูปแบบ data URI (ใช้ในใบพิมพ์/การคัดลอกเป็นรูปภาพ เพื่อไม่ติดข้อจำกัดข้ามโดเมน) */
function getBrandLogoData() {
  try { return getLogoDataUri_(getBranding()) || ''; } catch (e) { return ''; }
}