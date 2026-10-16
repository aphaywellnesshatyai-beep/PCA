// ==========================================================
// Netlify.gs — วางเป็นไฟล์ใหม่ใน Apps Script (ไฟล์ละส่วนกับ Code.gs)
// ทำให้หน้าเว็บบน Netlify เรียกฟังก์ชันของ GAS ได้
// ==========================================================

// ฟังก์ชันที่เรียกได้โดยไม่ต้องล็อกอิน (ใช้ตอนเปิดหน้า login)
var API_OPEN_ = ['checkLoginFromSheet', 'getBranding'];

// ฟังก์ชันที่เรียกได้หลังล็อกอินแล้วเท่านั้น
// ถ้าหน้าไหนขึ้น error "ไม่อนุญาตให้เรียก: ชื่อฟังก์ชัน" ให้เพิ่มชื่อนั้นต่อท้ายลิสต์นี้
var API_ = [
  'getNextCN', 'searchPatientForEdit', 'registerNewPatient', 'searchAllPatients', 'searchPatient',
  'saveVisit', 'getVisitHistory', 'getVisitByRow', 'getVisitStats', 'deleteVisit', 'getDiseaseList',
  'getElementInfo', 'searchHerb', 'getPostureDxRules', 'searchPatientForPosture', 'saveBodyTracking',
  'getTrackingHistory', 'deleteTrackingRecords', 'getPainMap', 'getTreatmentMethods', 'getMusicList',
  'getMusicData', 'getLastVisit', 'savePainMapImage', 'getBase64ImageFromUrl', 'getLinks', 'getDoctors',
  'getAppConfigs', 'getNextReceiptNumber',
  'adminGetData', 'adminSaveUser', 'adminDeleteUser', 'adminSaveSettings', 'adminTestFolder',
  'adminSaveBranding', 'adminUploadLogo', 'adminSaveDoctors', 'adminSaveMusic',
  // ฟังก์ชันจากไฟล์ Wheel.gs / Clinic.gs ฯลฯ
  'getDxTreatList', 'getWheelBundle', 'getWheelReading'
];

function tokenOk_(t) {
  return !!t && !!CacheService.getScriptCache().get('aw_tok_' + t);
}

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    var fn = String(req.fn || '');
    var open = API_OPEN_.indexOf(fn) > -1;
    if (!open && API_.indexOf(fn) < 0) throw new Error('ไม่อนุญาตให้เรียก: ' + fn);
    if (!open && !tokenOk_(req.token)) throw new Error('กรุณาเข้าสู่ระบบใหม่ (เซสชันหมดอายุ)');
    var f = globalThis[fn];
    if (typeof f !== 'function') throw new Error('ไม่พบฟังก์ชัน ' + fn + ' ใน GAS');
    out = { result: f.apply(null, req.args || []) };
  } catch (err) {
    out = { error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}