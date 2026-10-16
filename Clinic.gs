// ==========================================================
// 🧾 Clinic.gs — ใบรับรองแพทย์ / ใบเสร็จ / ฉลากยา / ห้องยา (ย้ายมาจากระบบเก่า)
// ไฟล์ใหม่ วางเพิ่มในโปรเจกต์ได้เลย
// ใช้ฟังก์ชันกลางจาก Code.gs: getSS, S_, pzNorm_, pzFolder_, readSettings_, SET_KEYS
// ==========================================================

// ---------- เลขที่ใบเสร็จ (ลำดับของผู้ป่วยในวันนั้น) ----------
function getNextReceiptNumber(targetCN) {
  try {
    var sheet = getSS().getSheetByName('Visit_Logs');
    if (!sheet) return "0001";
    var data = sheet.getDataRange().getValues();
    var todayStr = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy");
    var dailyVisits = [];
    for (var i = 1; i < data.length; i++) {
      var rowDate = data[i][0] ? data[i][0].toString().trim().replace(/'/g, "") : "";
      if (data[i][0] instanceof Date) rowDate = Utilities.formatDate(data[i][0], "GMT+7", "dd/MM/yyyy");
      if (rowDate.split(' ')[0] === todayStr) dailyVisits.push(pzNorm_(data[i][2]));
    }
    if (!targetCN) return (dailyVisits.length + 1).toString().padStart(4, '0');
    var idx = dailyVisits.indexOf(pzNorm_(targetCN));
    return idx !== -1 ? (idx + 1).toString().padStart(4, '0') : (dailyVisits.length + 1).toString().padStart(4, '0');
  } catch (e) { return "0001"; }
}

// ---------- บันทึก PDF ลง Google Drive ----------
// เว้นช่องโฟลเดอร์ในหน้า "ตั้งค่า" ว่างได้ — ระบบจะสร้างโฟลเดอร์ให้อัตโนมัติ
function savePdfToFolder_(settingKey, fallbackName, label, base64Data, filename) {
  try {
    var folder = pzFolder_(readSettings_()[settingKey], fallbackName);
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', filename);
    folder.createFile(blob);
    return "✅ บันทึก" + label + "ลง Google Drive (โฟลเดอร์ \"" + folder.getName() + "\") เรียบร้อยแล้ว!";
  } catch (error) {
    return "❌ บันทึกไม่สำเร็จ: ตรวจสอบลิงก์โฟลเดอร์ หรือสิทธิ์การเข้าถึง (" + error.message + ")";
  }
}
function saveMedicalCertificatePDF(base64Data, filename) {
  return savePdfToFolder_(SET_KEYS.certFolder, "AphaiWellness_MedCerts", "ใบรับรองแพทย์", base64Data, filename);
}
function saveCertPDF(base64Data, filename) { return saveMedicalCertificatePDF(base64Data, filename); }
function saveReceiptPDF(base64Data, filename) {
  return savePdfToFolder_(SET_KEYS.receiptFolder, "AphaiWellness_Receipts", "ใบเสร็จ", base64Data, filename);
}

// ---------- 💊 ห้องยา (Medicine_DB & Dispensing_Log) ----------
function getMedicineList() {
  try {
    var sheet = getSS().getSheetByName("Medicine_DB");
    if (!sheet) return [];
    var data = sheet.getDataRange().getDisplayValues();
    return data.slice(1)
      .filter(function (row) { return row[0] !== ""; })
      .map(function (row) {
        return { name: row[0], usage: row[1] || "", price: parseFloat(row[2]) || 0, description: row[3] || "", warning: row[45] || "" };
      });
  } catch (e) { return []; }
}

function searchPatientForPharmacy(query) {
  var sheet = getSS().getSheetByName("Patient_DB");
  if (!sheet) return null;
  var data = sheet.getDataRange().getDisplayValues();
  var s = String(query || "").trim().toLowerCase();
  var results = data.slice(1).filter(function (row) {
    var cn = row[0] ? row[0].toString().toLowerCase() : "";
    var name = row[3] ? row[3].toString().toLowerCase() : "";
    return cn.indexOf(s) > -1 || name.indexOf(s) > -1;
  }).map(function (row) {
    return { cn: row[0] || "-", name: (row[2] || "") + (row[3] || ""), address: row[12] || "-", allergy: row[15] || "ไม่มี" };
  });
  return results.length > 0 ? results : null;
}

function saveDispensingData(payload) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName("Dispensing_Log");
    if (!sheet) {
      sheet = ss.insertSheet("Dispensing_Log");
      sheet.appendRow(["CN", "วันที่", "รายการยา", "ราคารวม"]);
    }
    sheet.appendRow(["'" + pzNorm_(payload.cn), "'" + String(payload.date || "").trim(),
      String(payload.items || payload.medList || "").trim(), String(payload.total || "0").trim()]);
    return "Success";
  } catch (e) { return "Error: " + e.message; }
}

function getPatientDispensingHistory(cn) {
  try {
    var sheet = getSS().getSheetByName("Dispensing_Log");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues(), target = pzNorm_(cn), out = [];
    for (var i = 1; i < data.length; i++) {
      if (pzNorm_(data[i][0]) !== target) continue;
      var d = data[i][1];
      out.push({
        date: (d instanceof Date) ? Utilities.formatDate(d, "GMT+7", "dd/MM/yyyy HH:mm") : String(d || "-").replace(/'/g, ""),
        items: data[i][2] || "-", total: data[i][3] || "0.00"
      });
    }
    return out;
  } catch (err) { return []; }
}

function getPrescriptionHistory(cn, date) {
  try {
    var sheet = getSS().getSheetByName("Dispensing_Log");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues(), target = pzNorm_(cn), out = [];
    for (var i = 1; i < data.length; i++) {
      if (pzNorm_(data[i][0]) !== target) continue;
      var rawDrug = String(data[i][2] || "-"), qty = "1", drugName = rawDrug;
      var m = rawDrug.match(/(.*)\s*\((.*?)\)$/);
      if (m) { drugName = m[1].trim(); qty = m[2].trim(); }
      var d = data[i][1];
      out.push({
        date: (d instanceof Date) ? Utilities.formatDate(d, "GMT+7", "dd/MM/yyyy HH:mm") : String(d || "-").replace(/'/g, ""),
        drugName: drugName, qty: qty, price: data[i][3] || "0.00"
      });
    }
    return out;
  } catch (err) { return []; }
}