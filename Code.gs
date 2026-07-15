/**
 * ================================================================
 * Khaophun Note — Google Sheets Sync Backend (Apps Script)
 * ================================================================
 * วิธีติดตั้ง:
 * 1. สร้าง Google Sheet ใหม่ (ตั้งชื่ออะไรก็ได้ เช่น "Khaophun Note Data")
 * 2. เมนู Extensions → Apps Script → ลบโค้ดเดิม แล้ววางไฟล์นี้ทั้งไฟล์
 * 3. กด Deploy → New deployment → ประเภท "Web app"
 *      - Execute as:  Me
 *      - Who has access:  Anyone
 * 4. กด Deploy แล้ว copy "Web app URL" (ลงท้ายด้วย /exec)
 * 5. เปิดแอป Khaophun Note → หน้า "จัดการข้อมูล" → วาง URL → กด "เปิดซิงก์"
 *
 * ชีตที่สร้างอัตโนมัติ:
 *  - Snapshot    : สถานะล่าสุดทั้งหมด (JSON) — แอปใช้ดึงข้อมูลข้ามเครื่อง
 *  - Tasks       : ตารางงานล่าสุด อ่านง่ายสำหรับคน
 *  - Notes       : ตารางบันทึกล่าสุด
 *  - ActivityLog : ประวัติย้อนหลังแบบ append-only (ใครทำอะไร เมื่อไหร่)
 * ================================================================
 */

var TZ = 'Australia/Melbourne'; // เปลี่ยน timezone ได้ตามต้องการ เช่น 'Asia/Bangkok'

// ---------- Entry points ----------

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'load';
  if (action === 'load') {
    var snap = readSnapshot_();
    return json_({ ok: true, data: snap });
  }
  return json_({ ok: false, error: 'unknown action' });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'save' && body.data) {
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var prev = readSnapshot_();
        writeSnapshot_(body.data);
        rewriteTasks_(body.data.tasks || [], body.data.workspaces || []);
        rewriteNotes_(body.data.notes || [], body.data.workspaces || []);
        logDiff_(prev, body.data);
      } finally {
        lock.releaseLock();
      }
      return json_({ ok: true, savedAt: now_() });
    }
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// ---------- Snapshot (JSON state) ----------

function readSnapshot_() {
  var sh = sheet_('Snapshot');
  var raw = sh.getRange('A2').getValue();
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function writeSnapshot_(data) {
  var sh = sheet_('Snapshot');
  sh.getRange('A1').setValue('JSON state (อย่าแก้มือ)');
  sh.getRange('B1').setValue('อัปเดตล่าสุด');
  sh.getRange('A2').setValue(JSON.stringify(data));
  sh.getRange('B2').setValue(now_());
}

// ---------- Human-readable mirrors ----------

function rewriteTasks_(tasks, wss) {
  var sh = sheet_('Tasks');
  sh.clearContents();
  var head = ['ID', 'งาน', 'รายละเอียด', 'พื้นที่งาน', 'สถานะ', 'ความสำคัญ', 'กำหนดส่ง', 'สร้างเมื่อ'];
  var rows = tasks.map(function (t) {
    return [t.id, t.title || '', t.description || '', wsName_(t.workspace, wss),
            t.status || '', t.priority || '', t.dueDate || '', t.createdAt || ''];
  });
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
}

function rewriteNotes_(notes, wss) {
  var sh = sheet_('Notes');
  sh.clearContents();
  var head = ['ID', 'หัวข้อ', 'เนื้อหา', 'พื้นที่งาน', 'สร้างเมื่อ'];
  var rows = notes.map(function (n) {
    return [n.id, n.title || '', n.content || '', wsName_(n.workspace, wss), n.createdAt || ''];
  });
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
}

// ---------- ActivityLog: server-side diff = ประวัติย้อนหลัง ----------

function logDiff_(prev, next) {
  var sh = sheet_('ActivityLog');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 5)
      .setValues([['เวลา', 'ประเภท', 'การกระทำ', 'รายการ', 'รายละเอียด']])
      .setFontWeight('bold');
  }
  var rows = [];
  var ts = now_();
  var wss = (next && next.workspaces) || [];

  diffById_(prev && prev.tasks, next && next.tasks, function (kind, oldT, newT) {
    if (kind === 'added') {
      rows.push([ts, 'งาน', 'เพิ่ม', newT.title, wsName_(newT.workspace, wss) + ' · ' + (newT.status || '')]);
    } else if (kind === 'removed') {
      rows.push([ts, 'งาน', 'ลบ', oldT.title, wsName_(oldT.workspace, wss)]);
    } else {
      if (oldT.status !== newT.status) {
        rows.push([ts, 'งาน', 'ย้ายสถานะ', newT.title, oldT.status + ' → ' + newT.status]);
      }
      if (oldT.title !== newT.title || (oldT.description || '') !== (newT.description || '') ||
          (oldT.dueDate || '') !== (newT.dueDate || '') || oldT.priority !== newT.priority) {
        rows.push([ts, 'งาน', 'แก้ไข', newT.title, '']);
      }
    }
  });

  diffById_(prev && prev.notes, next && next.notes, function (kind, oldN, newN) {
    if (kind === 'added')   rows.push([ts, 'บันทึก', 'เพิ่ม', newN.title, wsName_(newN.workspace, wss)]);
    if (kind === 'removed') rows.push([ts, 'บันทึก', 'ลบ', oldN.title, '']);
    if (kind === 'changed' && (oldN.title !== newN.title || (oldN.content || '') !== (newN.content || ''))) {
      rows.push([ts, 'บันทึก', 'แก้ไข', newN.title, '']);
    }
  });

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
  }
}

function diffById_(oldArr, newArr, cb) {
  oldArr = oldArr || []; newArr = newArr || [];
  var oldMap = {}, newMap = {};
  oldArr.forEach(function (x) { oldMap[x.id] = x; });
  newArr.forEach(function (x) { newMap[x.id] = x; });
  newArr.forEach(function (x) {
    if (!oldMap[x.id]) cb('added', null, x);
    else cb('changed', oldMap[x.id], x);
  });
  oldArr.forEach(function (x) { if (!newMap[x.id]) cb('removed', x, null); });
}

// ---------- Helpers ----------

function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function wsName_(id, wss) {
  for (var i = 0; i < wss.length; i++) if (wss[i].id === id) return wss[i].name;
  return id || '';
}

function now_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
