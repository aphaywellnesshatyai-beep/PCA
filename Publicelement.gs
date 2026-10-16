// ==========================================================
// 🌐 PublicElement.gs — ฟังก์ชันของหน้าธาตุเจ้าเรือนสำหรับบุคคลทั่วไป (เปิดผ่าน QR Code ไม่ต้องล็อกอิน)
// ไฟล์ใหม่ — วางเพิ่มในโปรเจกต์ได้เลย (ไม่ต้องแก้ Code.gs)
// ส่งเฉพาะข้อมูลที่ไม่เป็นความลับ: ชื่อ/ที่อยู่/เบอร์โทร/โลโก้/สี ของคลินิก (ไม่ส่งรายชื่อแพทย์หรือข้อมูลผู้ป่วย)
// ==========================================================
function getPublicBrand() {
  var b = getBranding();
  return {
    name: b.name || '', clinicFull: b.clinicFull || '', clinicType: b.clinicType || '',
    address: b.address || '', phone: b.phone || '', email: b.email || '',
    crest: b.crest || 'อ', noLogo: !!b.noLogo, logoUrl: b.logoUrl || '',
    colors: b.colors || {}, title: b.title || ''
  };
}

/** โลโก้แบบฝังในรูป (data URI) — ใช้ตอนพิมพ์/บันทึกรูป เพื่อไม่ติดข้อจำกัดข้ามโดเมน */
function getBrandLogoData() {
  try {
    var b = getBranding();
    if (b.noLogo || !b.logo) return '';
    return getLogoDataUri_(b) || b.logoUrl || '';
  } catch (e) { return ''; }
}