// วางในชีทเดิม: ส่วนขยาย > Apps Script
const TOKEN = 'bomcal';           // ต้องตรงกับ STOCK_TOKEN ในหน้าเว็บ
const STOCK_SHEET = 'Stock';      // คอลัมน์ A=รหัส, B=ชื่อ, C=คงเหลือ
const LOG_SHEET = 'Log';
const PLAN_SHEET = 'Plan';        // แผนงานรายเดือน: A=เดือน, B=หมวด, C=รหัสงาน, D=จำนวน
const PLAN_HEADERS = ['เดือน', 'หมวด', 'รหัสงาน', 'จำนวน', 'หมายเหตุ', 'รหัสแถว', 'ปั๊ม', 'ยกเลิก', 'เลขกำกับ', 'ปิดงาน'];
const UNIT_SHEET = 'Units';       // ยูนิตต่อ 1 เครื่อง ของแต่ละรหัสงาน: A=รหัสงาน, B=ยูนิตต่อเครื่อง
const UNIT_HEADERS = ['รหัสงาน', 'ยูนิตต่อเครื่อง'];

// คอลัมน์ใน Log (A-U เดิม, V-X เพิ่มใหม่ต่อท้ายทางขวา ไม่กระทบข้อมูลเก่า)
const LOG_HEADERS = ['รหัสรายการ', 'เวลา', 'รหัสงาน', 'ชื่องาน', 'ปั๊ม', 'จำนวนผลิต',
  'รหัสเหล็ก', 'ชื่อเหล็ก', 'BOM/ตัว', 'Partlist', 'เบิก', 'คืน', 'ใช้จริง',
  'ส่วนต่าง', 'สต็อกก่อน', 'สต็อกหลัง',
  'ผู้ตัด', 'เศษทิ้ง', 'สถานะ', 'ยกเลิกโดย', 'เวลายกเลิก',
  'เดือนงาน', 'groupId', 'โน้ต', 'รหัสแถวแผน'];
const CANCELLED = 'ยกเลิกแล้ว';
const MAX_NUM = 10000000;

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function fail_(error, message) { return out_({ ok: false, error: error, message: message }); }

// แปลงเป็นตัวเลข: ว่าง = 0, ติดลบ/ไม่ใช่ตัวเลข/ใหญ่ผิดปกติ = ปฏิเสธ
function num_(v, label) {
  if (v === undefined || v === null || v === '') return 0;
  const n = Number(v);
  if (!isFinite(n) || n < 0 || n > MAX_NUM) throw new Error('ค่าไม่ถูกต้องที่ช่อง ' + label + ': ' + v);
  return n;
}

function round_(n) { return Math.round(n * 1000) / 1000; }

// เปิดชีท Log (สร้างถ้ายังไม่มี) และเติมหัวคอลัมน์ใหม่ให้ครบ
function ensureLog_(ss) {
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) sh = ss.insertSheet(LOG_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(LOG_HEADERS);
  } else {
    // เติมหัวคอลัมน์ตั้งแต่ตำแหน่ง 17 (ผู้ทำรายการ) ไปจนสุด รองรับทั้งชีทเก่าที่มีแค่ A-U และชีทใหม่ A-X
    const n = LOG_HEADERS.length - 16;
    const have = sh.getRange(1, 17, 1, n).getValues()[0];
    for (let i = 0; i < n; i++) {
      if (have[i] === '' || have[i] === null) sh.getRange(1, 17 + i).setValue(LOG_HEADERS[16 + i]);
    }
  }
  return sh;
}

// หน้าเว็บเรียกอ่านสต็อกปัจจุบัน / ประวัติการตัด
function doGet(e) {
  if (e.parameter.token !== TOKEN) return out_({ ok: false, error: 'unauthorized' });
  const ss = SpreadsheetApp.getActive();
  if (e.parameter.action === 'history') return out_(getHistory_(ss));
  if (e.parameter.action === 'plan') { const pl = getPlan_(ss); pl.collapsed = getPlanCollapsed_(); return out_(pl); }
  if (e.parameter.action === 'units') return out_(getUnits_(ss));
  if (e.parameter.action === 'summary') return out_(getSummary_(ss));
  const rows = ss.getSheetByName(STOCK_SHEET).getDataRange().getValues();
  const stock = {};
  rows.slice(1).forEach(r => { if (r[0]) stock[String(r[0]).trim()] = Number(r[2]) || 0; });
  return out_({ ok: true, stock: stock });
}

function fmt_(d) {
  return d ? Utilities.formatDate(new Date(d), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm') : '';
}

// รวมแถว Log ที่มี "groupId" เดียวกันเป็น 1 งาน (รองรับตัดสต็อกหลายวันของงานเดียวกัน)
// แถวเก่าที่ยังไม่มี groupId จะ fallback ไปใช้ "รหัสรายการ" ของตัวเองแทน (ไม่กระทบประวัติเดิม)
// เรียงงานตามเวลาล่าสุดที่มีการตัดสต็อกของกลุ่มนั้น (ใหม่สุดอยู่บนสุด)
function getHistory_(ss) {
  const sh = ss.getSheetByName(LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return { ok: true, jobs: [] };
  const last = sh.getLastRow();
  const from = Math.max(2, last - 2999);
  const cols = Math.min(LOG_HEADERS.length, Math.max(16, sh.getLastColumn()));
  const rows = sh.getRange(from, 1, last - from + 1, cols).getValues();
  const byId = {};
  const order = [];
  rows.forEach(r => {
    const rid = r[0];
    if (!rid) return;
    const gid = r[22] || rid;
    const ts = r[1] ? new Date(r[1]).getTime() : 0;
    if (!byId[gid]) {
      byId[gid] = {
        id: gid,
        time: fmt_(r[1]), _ts: ts,
        jobCode: String(r[2]), jobName: r[3], pump: r[4], qty: r[5],
        by: r[16] || '',
        month: r[21] || '',
        note: r[23] || '',
        planRowId: String(r[24] || ''),
        cancelled: r[18] === CANCELLED,
        cancelledAt: fmt_(r[20]),
        cancelledBy: r[19] || '',
        items: []
      };
      order.push(gid);
    } else {
      // อัปเดตข้อมูลให้เป็นของแถวล่าสุดในกลุ่ม
      if (ts >= byId[gid]._ts) { byId[gid].time = fmt_(r[1]); byId[gid]._ts = ts; }
      if (r[16]) byId[gid].by = r[16];
      if (r[23]) byId[gid].note = r[23];
      if (r[21]) byId[gid].month = r[21];
      if (r[18] === CANCELLED) { byId[gid].cancelled = true; byId[gid].cancelledAt = fmt_(r[20]); byId[gid].cancelledBy = r[19] || ''; }
    }
    byId[gid].items.push({
      code: String(r[6]), name: r[7], bom: r[8], partlist: r[9],
      withdraw: r[10], ret: r[11], used: r[12], diff: r[13], scrap: Number(r[17]) || 0
    });
  });
  const jobs = order.map(gid => byId[gid]).sort((a, b) => b._ts - a._ts);
  jobs.forEach(j => delete j._ts);
  return { ok: true, jobs: jobs.slice(0, 150) };
}

// หน้าเว็บส่งรายการตัดสต็อก หรือสั่งยกเลิกรายการ
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
      const d = JSON.parse(e.postData.contents);
      if (d.token !== TOKEN) return out_({ ok: false, error: 'unauthorized' });
      if (d.action === 'notify') return notifyLineHandler_(d.text);   // เพิ่มบรรทัดนี้
      const ss = SpreadsheetApp.getActive();
    if (d.action === 'planCollapsed') { PropertiesService.getScriptProperties().setProperty('planCollapsed', JSON.stringify(Array.isArray(d.collapsed) ? d.collapsed : [])); return out_({ ok: true }); }
    if (d.action === 'cancel') return cancel_(ss, d);
    if (d.action === 'edit') return edit_(ss, d);
    if (d.action === 'planSave') return planSave_(ss, d);
    if (d.action === 'noteSave') return noteSave_(ss, d);
    if (d.action === 'doneSave') return doneSave_(ss, d);
    if (d.action === 'unitSave') return unitSave_(ss, d);
    return commit_(ss, d);
  } catch (err) {
    const m = String(err && err.message ? err.message : err);
    return fail_(m.indexOf('ค่าไม่ถูกต้อง') === 0 ? 'validation' : 'exception', m);
  } finally {
    lock.releaseLock();
  }
}

function readStock_(stockSh) {
  const data = stockSh.getDataRange().getValues();
  const rowOf = {};
  data.forEach((r, i) => { if (i > 0 && r[0]) rowOf[String(r[0]).trim()] = i; });
  return { data: data, rowOf: rowOf };
}

function commit_(ss, d) {
  const stockSh = ss.getSheetByName(STOCK_SHEET);
  const logSh = ensureLog_(ss);

  // รองรับทั้ง d.by และ d.user (หน้าเว็บบางเวอร์ชันส่งชื่อฟิลด์ไม่ตรงกัน กันพังไว้ก่อน)
  const by = String(d.by || d.user || '').trim();
  if (!by) return fail_('validation', 'ต้องระบุชื่อผู้ทำรายการ');
  if (by.length > 60) return fail_('validation', 'ชื่อผู้ทำรายการยาวเกินไป');
  if (!d.id) return fail_('validation', 'ไม่มีรหัสรายการ');
  if (!Array.isArray(d.items) || d.items.length === 0) return fail_('validation', 'ไม่มีรายการเหล็ก');
  const qty = num_(d.qty, 'จำนวนผลิต');

  // กันกดส่งซ้ำ (ตรวจ 300 แถวล่าสุด) — ใช้ "รหัสรายการ" (id) เป็น idempotency key ของการกดยืนยันครั้งนี้เท่านั้น
  const last = logSh.getLastRow();
  if (last > 1) {
    const from = Math.max(2, last - 299);
    const ids = logSh.getRange(from, 1, last - from + 1, 1).getValues().flat();
    if (ids.indexOf(d.id) !== -1) return fail_('duplicate', 'รายการนี้บันทึกไปแล้ว');
  }

  const st = readStock_(stockSh);

  // ตรวจทั้งหมดก่อน ถ้ามีปัญหาจะไม่ตัดอะไรเลย
  // partial = ตัดเฉพาะเหล็กที่สต็อกพอ (ใช้กับงานผลิตเท่านั้น) รายการที่สต็อกไม่พอจะถูกข้ามและส่งกลับให้หน้าเว็บแจ้ง
  const partial = d.partial === true && String(d.jobCode || '') !== 'ADJUST';
  const problems = [];
  const skipped = [];
  const seen = {};
  const items = d.items.map(it => {
    const x = {
      code: String(it.code || '').trim(), name: it.name, bom: num_(it.bom, 'BOM'),
      withdraw: num_(it.withdraw, 'เบิก (' + it.name + ')'),
      ret: num_(it.ret, 'คืน (' + it.name + ')'),
      scrap: num_(it.scrap, 'เศษทิ้ง (' + it.name + ')')
    };
    x.used = x.withdraw - x.ret;
    if (seen[x.code]) problems.push('รหัส ' + x.code + ' ซ้ำในรายการ');
    seen[x.code] = true;
    const i = st.rowOf[x.code];
    if (i === undefined) problems.push('ไม่พบรหัส ' + x.code + ' (' + x.name + ') ในชีท Stock');
    else if (round_(Number(st.data[i][2]) - x.used) < 0) {
      if (partial) { x.skip = true; skipped.push({ code: x.code, name: x.name, have: Number(st.data[i][2]) || 0, need: x.used }); }
      else problems.push(x.name + ' สต็อกไม่พอ (เหลือ ' + st.data[i][2] + ' ต้องตัด ' + x.used + ')');
    }
    return x;
  });
  if (problems.length) return fail_('validation', problems.join('\n'));
  const cutItems = items.filter(x => !x.skip);
  if (!cutItems.length) return fail_('validation', 'สต็อกไม่พอทุกรายการ ยังตัดไม่ได้\n' + skipped.map(x => x.name + ' (เหลือ ' + x.have + ' ต้องตัด ' + x.need + ')').join('\n'));

  // groupId = ตัวจับกลุ่ม "งานเดียวกัน" (คนละวันก็รวมประวัติเป็นการ์ดเดียว) — ถ้าไม่ส่งมา fallback เป็น id ของแถวนี้เอง
  const groupId = String(d.groupId || d.id);
  const month = String(d.month || '');
  const note = String(d.note || '').trim();

  // ผู้ใช้เลือกรวมกับงานเดิม: ตรวจฝั่งเซิร์ฟเวอร์ว่ากลุ่มนั้นยังไม่ถูกยกเลิก และไม่มีเหล็กชนิดเดียวกันซ้ำ
  if (d.merge === true) {
    const lastR = logSh.getLastRow();
    const fromR = Math.max(2, lastR - 2999);
    const vals = lastR < 2 ? [] : logSh.getRange(fromR, 1, lastR - fromR + 1, 23).getValues();
    const found = vals.filter(v => String(v[22] || v[0]) === groupId);
    if (!found.length) return fail_('validation', 'ไม่พบงานที่เลือกรวม (อาจถูกลบหรือเก่าเกินไป)');
    if (found.some(v => v[18] === CANCELLED)) return fail_('validation', 'งานที่เลือกรวมถูกยกเลิกไปแล้ว');
    const have = {};
    found.forEach(v => { have[String(v[6]).trim()] = true; });
    const dup = cutItems.filter(x => have[x.code]).map(x => x.name);
    if (dup.length) return fail_('validation', 'เหล็กนี้ถูกตัดในงานที่เลือกรวมแล้ว: ' + dup.join(', '));
  }

  const now = new Date();
  const logRows = [];
  cutItems.forEach(x => {
    const i = st.rowOf[x.code];
    const before = Number(st.data[i][2]);
    const after = round_(before - x.used);
    stockSh.getRange(i + 1, 3).setValue(after);
    logRows.push([d.id, now, d.jobCode, d.jobName, d.pump, qty,
      x.code, x.name, x.bom, x.bom * qty, x.withdraw, x.ret, x.used,
      x.used - x.bom * qty, before, after,
      by, x.scrap, '', '', '',
      month, groupId, note, String(d.planRowId || '').slice(0, 60)]);
  });
  logSh.getRange(logSh.getLastRow() + 1, 22, logRows.length, 1).setNumberFormat('@');   // เดือนงาน (คอลัมน์ V) เป็นข้อความ
  logSh.getRange(logSh.getLastRow() + 1, 1, logRows.length, LOG_HEADERS.length).setValues(logRows);
  SpreadsheetApp.flush();

  return out_({ ok: true, count: logRows.length, groupId: groupId, skipped: skipped.map(x => ({ code: x.code, name: x.name, have: x.have, need: x.need })) });
}

// แก้ไขรายการที่ตัดไปแล้ว: คืนสต็อกของรอบเดิมทั้งหมด แล้วตัดตามค่าใหม่ในครั้งเดียว (ทำในล็อกเดียวกัน)
// รอบเดิมถูกทำเครื่องหมายยกเลิก และสร้างกลุ่มใหม่ 1 กลุ่มแทน
// d.gids = รหัสกลุ่มของรอบเดิม, d.items = [{code,name,bom,withdraw,ret,scrap}], d.keepDate = ใช้เวลาเดิมหรือไม่
function edit_(ss, d) {
  const stockSh = ss.getSheetByName(STOCK_SHEET);
  const logSh = ensureLog_(ss);
  const by = String(d.by || d.user || '').trim();
  if (!by) return fail_('validation', 'ต้องระบุชื่อผู้แก้ไข');
  if (by.length > 60) return fail_('validation', 'ชื่อผู้แก้ไขยาวเกินไป');
  const gids = (Array.isArray(d.gids) ? d.gids : []).map(String);
  if (!gids.length) return fail_('validation', 'ไม่มีรหัสรายการ');
  if (!Array.isArray(d.items) || !d.items.length) return fail_('validation', 'ต้องมีเหล็กอย่างน้อย 1 รายการ');

  const last = logSh.getLastRow();
  if (last < 2) return fail_('not_found', 'ไม่พบรายการนี้ใน Log');
  const cols = Math.min(LOG_HEADERS.length, Math.max(16, logSh.getLastColumn()));
  const rows = logSh.getRange(2, 1, last - 1, cols).getValues();
  const hit = [];
  rows.forEach((r, k) => { if (gids.indexOf(String(r[22] || r[0])) !== -1) hit.push({ sheetRow: k + 2, r: r }); });
  if (!hit.length) return fail_('not_found', 'ไม่พบรายการนี้ใน Log');
  if (hit.some(h => h.r[18] === CANCELLED)) return fail_('already', 'รายการนี้ถูกยกเลิกไปแล้ว');

  hit.sort((a, b) => new Date(a.r[1]).getTime() - new Date(b.r[1]).getTime());
  const head = hit[hit.length - 1].r;             // หัวงานจากรอบล่าสุด
  const qty = Number(head[5]) || 0;
  const time = d.keepDate === true ? head[1] : new Date();
  const now = new Date();

  // จำนวนที่จะคืนของรอบเดิม รวมต่อรหัส
  const restore = {};
  hit.forEach(h => { const c = String(h.r[6]).trim(); restore[c] = (restore[c] || 0) + (Number(h.r[12]) || 0); });

  const st = readStock_(stockSh);
  const problems = [], seen = {};
  const items = d.items.map(it => {
    const x = {
      code: String(it.code || '').trim(), name: it.name, bom: num_(it.bom, 'BOM'),
      withdraw: num_(it.withdraw, 'เบิก (' + it.name + ')'),
      ret: num_(it.ret, 'คืน (' + it.name + ')'),
      scrap: num_(it.scrap, 'เศษทิ้ง (' + it.name + ')')
    };
    x.used = x.withdraw - x.ret;
    if (seen[x.code]) problems.push('รหัส ' + x.code + ' ซ้ำในรายการ');
    seen[x.code] = true;
    const i = st.rowOf[x.code];
    if (i === undefined) problems.push('ไม่พบรหัส ' + x.code + ' (' + x.name + ') ในชีท Stock');
    else if (round_(Number(st.data[i][2]) + (restore[x.code] || 0) - x.used) < 0) {
      problems.push(x.name + ' สต็อกไม่พอ (มี ' + (Number(st.data[i][2]) + (restore[x.code] || 0)) + ' ต้องตัด ' + x.used + ')');
    }
    return x;
  });
  if (problems.length) return fail_('validation', problems.join('\n'));

  // ทำจริง: คำนวณสต็อกทีละรหัสจาก (สต็อกเดิม + คืนรอบเดิม) แล้วหักค่าใหม่
  const cur = {};
  const touched = Object.keys(restore).concat(items.map(x => x.code));
  touched.forEach(c => { const i = st.rowOf[c]; if (i !== undefined) cur[c] = Number(st.data[i][2]) + (restore[c] || 0); });

  const rowId = Utilities.getUuid();
  const groupId = Utilities.getUuid();
  const logRows = [];
  const pumpVal = String(head[2]) === 'ADJUST' ? (String(d.note || '').trim().slice(0, 200) || '-') : head[4];
  items.forEach(x => {
    const before = cur[x.code];
    const after = round_(before - x.used);
    cur[x.code] = after;
    logRows.push([rowId, time, head[2], head[3], pumpVal, qty,
      x.code, x.name, x.bom, x.bom * qty, x.withdraw, x.ret, x.used,
      x.used - x.bom * qty, before, after,
      by, x.scrap, '', '', '',
      head[21], groupId, head[23], head[24] || '']);
  });
  // ตั้งสต็อกทุกรหัสที่เกี่ยวข้อง
  touched.forEach(c => { const i = st.rowOf[c]; if (i !== undefined) stockSh.getRange(i + 1, 3).setValue(cur[c]); });
  // ทำเครื่องหมายรอบเดิมว่ายกเลิก
  hit.forEach(h => logSh.getRange(h.sheetRow, 19, 1, 3).setValues([[CANCELLED, by, now]]));
  logSh.getRange(logSh.getLastRow() + 1, 22, logRows.length, 1).setNumberFormat('@');   // เดือนงาน (คอลัมน์ V) เป็นข้อความ
  logSh.getRange(logSh.getLastRow() + 1, 1, logRows.length, LOG_HEADERS.length).setValues(logRows);
  SpreadsheetApp.flush();

  return out_({ ok: true, count: logRows.length, groupId: groupId });
}

// ยกเลิกรายการ: คืนสต็อกตาม "ใช้จริง" ของทุกแถวใน groupId นี้ (รวมของวันอื่นๆ ที่ตัดเพิ่มทีหลังด้วย) แล้วทำเครื่องหมายใน Log (ไม่ลบแถว)
function cancel_(ss, d) {
  const stockSh = ss.getSheetByName(STOCK_SHEET);
  const logSh = ensureLog_(ss);

  // รองรับทั้ง d.by และ d.user
  const by = String(d.by || d.user || '').trim();
  if (!by) return fail_('validation', 'ต้องระบุชื่อผู้ยกเลิก');
  if (by.length > 60) return fail_('validation', 'ชื่อผู้ยกเลิกยาวเกินไป');
  if (!d.id) return fail_('validation', 'ไม่มีรหัสรายการ');

  const last = logSh.getLastRow();
  if (last < 2) return fail_('not_found', 'ไม่พบรายการนี้ใน Log');
  const cols = Math.min(LOG_HEADERS.length, Math.max(16, logSh.getLastColumn()));
  const rows = logSh.getRange(2, 1, last - 1, cols).getValues();

  const hit = [];   // เลขแถวจริงในชีท
  // d.id ที่ส่งมาคือ "id" ของงานที่แสดงในประวัติ ซึ่งอาจเป็น groupId (ถ้ามีการตัดหลายวัน) หรือ id เดิม (ถ้ามีแถวเดียว)
  rows.forEach((r, k) => {
    const gid = r[22] || r[0];
    if (gid === d.id || r[0] === d.id) hit.push({ sheetRow: k + 2, r: r });
  });
  if (hit.length === 0) return fail_('not_found', 'ไม่พบรายการนี้ใน Log');
  if (hit.some(h => h.r[18] === CANCELLED)) return fail_('already', 'รายการนี้ถูกยกเลิกไปแล้ว');

  const st = readStock_(stockSh);
  const delta = {};   // รหัสเหล็ก -> จำนวนที่จะบวกคืน
  hit.forEach(h => {
    const code = String(h.r[6]).trim();
    delta[code] = (delta[code] || 0) + (Number(h.r[12]) || 0);
  });

  const problems = [];
  Object.keys(delta).forEach(code => {
    const i = st.rowOf[code];
    if (i === undefined) { problems.push('ไม่พบรหัส ' + code + ' ในชีท Stock'); return; }
    if (round_(Number(st.data[i][2]) + delta[code]) < 0) {
      problems.push('คืนสต็อก ' + code + ' ไม่ได้: จะติดลบ (เหลือ ' + st.data[i][2] + ', ปรับ ' + delta[code] + ')');
    }
  });
  if (problems.length) return fail_('validation', problems.join('\n'));

  const now = new Date();
  Object.keys(delta).forEach(code => {
    const i = st.rowOf[code];
    stockSh.getRange(i + 1, 3).setValue(round_(Number(st.data[i][2]) + delta[code]));
  });
  hit.forEach(h => {
    logSh.getRange(h.sheetRow, 19, 1, 3).setValues([[CANCELLED, by, now]]);
  });
  SpreadsheetApp.flush();

  return out_({ ok: true, count: hit.length });
}

// ===== สรุปการใช้เหล็กตั้งแต่เริ่มต้น (ทุกรายการใน Log เรียงตามเวลา) =====
function getSummary_(ss) {
  const names = {};
  ss.getSheetByName(STOCK_SHEET).getDataRange().getValues().slice(1).forEach(r => {
    if (r[0]) names[String(r[0]).trim()] = { name: r[1], now: Number(r[2]) || 0 };
  });
  const sh = ss.getSheetByName(LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return { ok: true, events: [], names: names };
  const last = sh.getLastRow();
  const cols = Math.min(LOG_HEADERS.length, Math.max(16, sh.getLastColumn()));
  const rows = sh.getRange(2, 1, last - 1, cols).getValues();
  const map = {}, out = [];
  rows.forEach(r => {
    if (!r[0] || !r[1]) return;
    const ts = new Date(r[1]).getTime();
    const gid = String(r[22] || r[0]);
    const cancelled = r[18] === CANCELLED;
    const key = gid + '|' + ts;
    let e = map[key];
    if (!e) {
      e = map[key] = { type: String(r[2]) === 'ADJUST' ? 'adjust' : 'job', ts: ts, time: fmt_(r[1]),
        jobCode: String(r[2]), jobName: r[3], pump: r[4], qty: Number(r[5]) || 0, month: String(r[21] || ''),
        by: r[16] || '', cancelled: cancelled, gid: gid, items: [] };
      out.push(e);
    }
    if (cancelled) e.cancelled = true;
    e.items.push({ code: String(r[6]).trim(), name: r[7], withdraw: Number(r[10]) || 0, ret: Number(r[11]) || 0,
      used: Number(r[12]) || 0, scrap: Number(r[17]) || 0, before: Number(r[14]) || 0, after: Number(r[15]) || 0 });
    if (cancelled && r[20]) {
      const cts = new Date(r[20]).getTime();
      const ck = gid + '|cancel|' + cts;
      let c = map[ck];
      if (!c) {
        c = map[ck] = { type: 'cancel', ts: cts, time: fmt_(r[20]), jobCode: String(r[2]), jobName: r[3], pump: r[4],
          qty: Number(r[5]) || 0, month: String(r[21] || ''), by: r[19] || '', cancelled: false, items: [] };
        out.push(c);
      }
      const back = Number(r[12]) || 0;
      c.items.push({ code: String(r[6]).trim(), name: r[7], withdraw: 0, ret: back, used: -back, scrap: 0, before: null, after: null });
    }
  });
  out.sort((a, b) => a.ts - b.ts);
  return { ok: true, events: out.slice(-3000), names: names };
}

// ===== แผนงานรายเดือน =====
function ensurePlan_(ss) {
  let sh = ss.getSheetByName(PLAN_SHEET);
  if (!sh) sh = ss.insertSheet(PLAN_SHEET);
  if (sh.getLastRow() === 0) sh.appendRow(PLAN_HEADERS);
  else if (!sh.getRange(1, 5).getValue()) sh.getRange(1, 5).setValue('หมายเหตุ');
  if (!sh.getRange(1, 6).getValue()) sh.getRange(1, 6).setValue('รหัสแถว');
  if (!sh.getRange(1, 7).getValue()) sh.getRange(1, 7).setValue('ปั๊ม');
  if (!sh.getRange(1, 8).getValue()) sh.getRange(1, 8).setValue('ยกเลิก');
  if (!sh.getRange(1, 9).getValue()) sh.getRange(1, 9).setValue('เลขกำกับ');
  if (!sh.getRange(1, 10).getValue()) sh.getRange(1, 10).setValue('ปิดงาน');
  return sh;
}

function getPlanCollapsed_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('planCollapsed') || '[]'); } catch (e) { return []; }
}
function getPlan_(ss) {
  const sh = ensurePlan_(ss);
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, plan: [] };
  const v = sh.getRange(2, 1, last - 1, 10).getValues();
  // ย้อนหลัง: รหัสซ้ำในเดือนเดียวกันที่ยังไม่มีเลขกำกับ ให้เรียงตามลำดับในชีตแล้วบันทึกลงชีตครั้งเดียว
  const cnt = {};
  v.forEach(r => { const c = String(r[2]).trim(); if (c) { const k = String(r[0]) + '|' + c; cnt[k] = (cnt[k] || 0) + 1; } });
  const nextSeq = {}, fill = [];
  let changed = false;
  v.forEach(r => {
    const c = String(r[2]).trim(); if (!c) { fill.push([r[8]]); return; }
    const k = String(r[0]) + '|' + c;
    if (cnt[k] < 2) { fill.push([Number(r[8]) || 0]); return; }
    nextSeq[k] = nextSeq[k] || 0;
    let sq = Number(r[8]) || 0;
    if (!sq) { sq = nextSeq[k] + 1; changed = true; }
    if (sq > nextSeq[k]) nextSeq[k] = sq;
    fill.push([sq]);
  });
  if (changed) sh.getRange(2, 9, v.length, 1).setValues(fill);
  return {
    ok: true,
    plan: v.filter(r => String(r[2]).trim() !== '').map(r => ({
      id: String(r[5] || ''), month: String(r[0]), section: String(r[1] || ''), code: String(r[2]).trim(), qty: Number(r[3]) || 0, note: String(r[4] || ''), pump: String(r[6] || ''), cancelled: String(r[7] || '').trim() === 'ยกเลิก',
      seq: (function (i) { return Number(fill[i][0]) || 0; })(v.indexOf(r)),
      done: String(r[9] || '').trim() === 'เสร็จ'
    }))
  };
}

// แก้หมายเหตุของแถวเดียว (ไม่ต้องใช้รหัส PIN) จับแถวจาก รหัสแถว + เดือน + รหัสงาน ให้ตรงตัว
function noteSave_(ss, d) {
  const id = String(d.id || '').trim(), month = String(d.month || '').trim(), code = String(d.code || '').trim();
  if (!id || !month || !code) return fail_('validation', 'ข้อมูลแถวไม่ครบ');
  const note = String(d.note || '').trim().slice(0, 200);
  const sh = ensurePlan_(ss);
  const last = sh.getLastRow();
  if (last < 2) return fail_('not_found', 'ไม่พบแถวแผน');
  const v = sh.getRange(2, 1, last - 1, 9).getValues();
  const hit = [];
  v.forEach((r, k) => { if (String(r[5]) === id && String(r[0]) === month && String(r[2]).trim() === code) hit.push(k); });
  if (hit.length !== 1) return fail_('not_found', 'ไม่พบแถวแผนนี้ (อาจถูกแก้ไขหรือลบไปแล้ว)');
  sh.getRange(hit[0] + 2, 5).setValue(note);
  SpreadsheetApp.flush();
  return out_({ ok: true });
}

// ปิดงาน/เปิดงานใหม่ของงานลูกค้า (ไม่ใช้รหัส) จับแถวด้วย รหัสแถว + เดือน + รหัสงาน
function doneSave_(ss, d) {
  const id = String(d.id || '').trim(), month = String(d.month || '').trim(), code = String(d.code || '').trim();
  if (!id || !month || !code) return fail_('validation', 'ข้อมูลแถวไม่ครบ');
  const sh = ensurePlan_(ss);
  const last = sh.getLastRow();
  if (last < 2) return fail_('not_found', 'ไม่พบแถวแผน');
  const v = sh.getRange(2, 1, last - 1, 10).getValues();
  const hit = [];
  v.forEach((r, k) => { if (String(r[5]) === id && String(r[0]) === month && String(r[2]).trim() === code) hit.push(k); });
  if (hit.length !== 1) return fail_('not_found', 'ไม่พบแถวแผนนี้ (อาจถูกแก้ไขหรือลบไปแล้ว)');
  sh.getRange(hit[0] + 2, 10).setValue(d.done ? 'เสร็จ' : '');
  SpreadsheetApp.flush();
  return out_({ ok: true });
}

// บันทึกแผนของ "เดือนเดียว" แทนที่ของเดิมทั้งเดือน (เดือนอื่นไม่กระทบ)
function planSave_(ss, d) {
  const month = String(d.month || '').trim();
  if (!/^(\d{4}-(0[1-9]|1[0-2])|1[0-2]|[1-9])$/.test(month)) return fail_('validation', 'เดือนไม่ถูกต้อง');   // "1".."12" = ปีเดิม, "2027-01" = ระบุปี
  if (!Array.isArray(d.rows) || d.rows.length > 300) return fail_('validation', 'ข้อมูลแผนไม่ถูกต้อง');
  const rows = d.rows.map(r => ({
    section: String(r.section || '').trim().slice(0, 60),
    code: String(r.code || '').trim(),
    qty: num_(r.qty, 'จำนวน'),
    note: String(r.note || '').trim().slice(0, 200),
    id: String(r.id || '').trim().slice(0, 60),
    pump: String(r.pump || '').trim().slice(0, 60),
    cancelled: !!r.cancelled,
    seq: Math.max(0, Math.floor(Number(r.seq) || 0)),
    done: !!r.done
  }));
  if (rows.some(r => !r.code || r.code.length > 20)) return fail_('validation', 'มีบรรทัดที่ไม่มีรหัสงาน');

  // เลขกำกับต่อรหัสในเดือนเดียวกัน: เลขเดิมคงไว้ (ซ้ำหรือว่างจะได้เลขใหม่ต่อจากสูงสุด)
  const groups = {};
  rows.forEach(r => { const k = r.code; (groups[k] = groups[k] || []).push(r); });
  Object.keys(groups).forEach(k => {
    const list = groups[k];
    const taken = {};
    list.forEach(r => { if (r.seq > 0 && !taken[r.seq]) taken[r.seq] = true; else r.seq = 0; });
    let top = 0; Object.keys(taken).forEach(x => { if (Number(x) > top) top = Number(x); });
    list.forEach(r => { if (!r.seq) { top += 1; r.seq = top; } });
  });

  const sh = ensurePlan_(ss);
  const last = sh.getLastRow();
  const keep = last < 2 ? [] : sh.getRange(2, 1, last - 1, 10).getValues().filter(r => String(r[0]) !== month);
  const all = keep.concat(rows.map(r => [month, r.section, r.code, r.qty, r.note, r.id, r.pump, r.cancelled ? 'ยกเลิก' : '', r.seq, r.done ? 'เสร็จ' : '']));
  if (last >= 2) sh.getRange(2, 1, last - 1, 10).clearContent();
  if (all.length) {
    sh.getRange(2, 1, all.length, 1).setNumberFormat('@');   // เดือน (รองรับ "2027-01") เก็บเป็นข้อความ ไม่ให้ชีทแปลงเป็นวันที่
    sh.getRange(2, 3, all.length, 1).setNumberFormat('@');   // เก็บรหัสเป็นข้อความ ไม่ให้เลขศูนย์นำหน้าหาย
    sh.getRange(2, 1, all.length, 10).setValues(all);
  }
  SpreadsheetApp.flush();
  return out_({ ok: true, count: rows.length, seqs: rows.map(r => r.seq) });
}

// ===== ยูนิตต่อ 1 เครื่อง (ฐานข้อมูลรายรหัสงาน) =====
function ensureUnits_(ss) {
  let sh = ss.getSheetByName(UNIT_SHEET);
  if (!sh) sh = ss.insertSheet(UNIT_SHEET);
  if (sh.getLastRow() === 0) sh.appendRow(UNIT_HEADERS);
  return sh;
}

function getUnits_(ss) {
  const sh = ensureUnits_(ss);
  const last = sh.getLastRow();
  const units = {};
  if (last >= 2) sh.getRange(2, 1, last - 1, 2).getValues().forEach(r => {
    const c = String(r[0]).trim();
    if (c) units[c] = Number(r[1]) || 0;
  });
  return { ok: true, units: units };
}

// เพิ่ม/อัปเดตยูนิตต่อเครื่องของรหัสที่ส่งมา (รหัสอื่นไม่กระทบ)
function unitSave_(ss, d) {
  if (!Array.isArray(d.rows) || d.rows.length > 500) return fail_('validation', 'ข้อมูลยูนิตไม่ถูกต้อง');
  const upd = {};
  d.rows.forEach(r => {
    const code = String(r.code || '').trim();
    if (!code || code.length > 20) throw new Error('ค่าไม่ถูกต้องที่ช่อง รหัสงาน: ' + code);
    upd[code] = num_(r.units, 'ยูนิตต่อเครื่อง (' + code + ')');
  });
  const sh = ensureUnits_(ss);
  const last = sh.getLastRow();
  const cur = last < 2 ? [] : sh.getRange(2, 1, last - 1, 2).getValues().map(r => [String(r[0]).trim(), r[1]]);
  const seen = {};
  cur.forEach(r => { if (upd[r[0]] !== undefined) { r[1] = upd[r[0]]; } seen[r[0]] = true; });
  Object.keys(upd).forEach(c => { if (!seen[c]) cur.push([c, upd[c]]); });
  if (last >= 2) sh.getRange(2, 1, last - 1, 2).clearContent();
  if (cur.length) {
    sh.getRange(2, 1, cur.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, cur.length, 2).setValues(cur);
  }
  SpreadsheetApp.flush();
  return out_({ ok: true, count: Object.keys(upd).length });
}

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