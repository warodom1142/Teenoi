// ===== แจ้งเตือนไลน์ (วางใน Apps Script ของ Code.gs) =====
// 1) Project Settings > Script Properties เพิ่ม
//    LINE_TOKEN     = Channel access token (long-lived) จาก LINE Developers
//    LINE_GROUP_ID  = Group ID ของกลุ่มที่บอทอยู่ (ขึ้นต้นด้วย C)
// 2) วางฟังก์ชันนี้ไว้ใน Code.gs
// 3) ใน doPost เดิม (หลังตรวจ token แล้ว) เพิ่มบรรทัดนี้ โดยปรับชื่อตัวแปรให้ตรงกับของคุณ
//    if (body.action === 'notify') return notifyLineHandler_(body.text);

function notifyLine_(text) {
  const props = PropertiesService.getScriptProperties();
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + props.getProperty('LINE_TOKEN') },
    payload: JSON.stringify({
      to: props.getProperty('LINE_GROUP_ID'),
      messages: [{ type: 'text', text: String(text).slice(0, 5000) }]
    }),
    muteHttpExceptions: true
  });
  return { ok: res.getResponseCode() === 200, message: res.getContentText() };
}

function notifyLineHandler_(text) {
  const r = notifyLine_(text);
  return ContentService.createTextOutput(JSON.stringify({ ok: r.ok, message: r.message }))
    .setMimeType(ContentService.MimeType.JSON);
}
