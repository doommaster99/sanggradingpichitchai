/**
 * สั่นกระดิ่งพิชิตชัย — ฝั่งเซิร์ฟเวอร์ (Google Apps Script)
 *
 * หน้าที่
 *  1) เสิร์ฟหน้าเกม (Index.html) เป็นเว็บแอป
 *  2) getQuestionSets(): อ่านชุดคำถามจากแท็บในสเปรดชีต (1 แท็บ = 1 ชุดคำถาม)
 *  3) saveResult(): บันทึกผลคะแนนลงแท็บ "ผลคะแนน"
 *
 * รูปแบบแท็บคำถาม (แถวแรกเป็นหัวตาราง):
 *  หัวข้อ | ระดับ | คำถาม | คำตอบที่ถูก | ตัวเลือกผิด 1 | ตัวเลือกผิด 2 | ตัวเลือกผิด 3 | คำอธิบาย
 */
const CONFIG = {
  SPREADSHEET_ID: '',        // เว้นว่าง = ใช้สเปรดชีตที่สคริปต์นี้ผูกอยู่ (เปิดจาก ส่วนขยาย > Apps Script)
  RESULT_SHEET: 'ผลคะแนน',   // แท็บเก็บผลคะแนน (สร้างให้เองถ้ายังไม่มี)
  SKIP_PREFIX: '_',          // แท็บที่ชื่อขึ้นต้นด้วยตัวอักษรนี้จะไม่ถูกนับเป็นชุดคำถาม
  MAX_ROWS_PER_SET: 1000,    // จำนวนแถวสูงสุดที่อ่านต่อชุด
  MAX_COLS: 12,              // จำนวนคอลัมน์สูงสุดที่อ่านต่อแท็บ
  MAX_RESULT_ROWS: 12        // จำนวนผู้เล่น/ทีมสูงสุดที่บันทึกต่อเกม
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('สั่นกระดิ่งพิชิตชัย')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ss_() {
  return CONFIG.SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

/** ส่งข้อมูลดิบ (2 มิติ) ของแต่ละแท็บคำถามให้หน้าเกมตรวจและแปลงเอง */
function getQuestionSets() {
  const out = [];
  ss_().getSheets().forEach(function (sh) {
    const name = sh.getName();
    if (name.indexOf(CONFIG.SKIP_PREFIX) === 0 || name === CONFIG.RESULT_SHEET) return;
    const lastRow = Math.min(sh.getLastRow(), CONFIG.MAX_ROWS_PER_SET + 1);
    const lastCol = Math.min(sh.getLastColumn(), CONFIG.MAX_COLS);
    if (lastRow < 2 || lastCol < 3) return;
    const rows = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    if (!looksLikeQuestionSheet_(rows[0])) return;
    out.push({ name: name, rows: rows });
  });
  return out;
}

function looksLikeQuestionSheet_(header) {
  return header.some(function (h) {
    const n = String(h).replace(/\s+/g, '').toLowerCase();
    return n === 'คำถาม' || n === 'question' || n === 'โจทย์';
  });
}

/** บันทึกผลหนึ่งเกม (หนึ่งแถวต่อผู้เล่นหรือทีม) */
function saveResult(payload) {
  if (!payload || !Array.isArray(payload.players) || !payload.players.length) {
    throw new Error('ข้อมูลผลคะแนนไม่ถูกต้อง');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = ss_();
    let sh = ss.getSheetByName(CONFIG.RESULT_SHEET);
    if (!sh) {
      sh = ss.insertSheet(CONFIG.RESULT_SHEET);
      sh.appendRow(['เวลา', 'ชุดคำถาม', 'โหมด', 'ผู้เล่น/ทีม', 'อันดับ', 'คะแนน',
                    'ตอบแล้ว', 'ตอบถูก', 'ตอบผิด', 'หมดเวลา', 'จำนวนข้อ', 'ใช้เวลา (วินาที)']);
      sh.setFrozenRows(1);
      sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    }
    const now = new Date();
    const mode = payload.mode === 'bell' ? 'แข่งกดกระดิ่ง' : 'เล่นคนเดียว';
    const rows = payload.players.slice(0, CONFIG.MAX_RESULT_ROWS).map(function (p) {
      return [now, txt_(payload.setName), mode, txt_(p.name), num_(p.rank), num_(p.score),
              num_(p.answered), num_(p.correct), num_(p.wrong), num_(p.timeout),
              num_(payload.questions), num_(payload.seconds)];
    });
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { saved: rows.length };
  } finally {
    lock.releaseLock();
  }
}

/** ตัดความยาว และกันข้อความที่ขึ้นต้นด้วย = + - @ ไม่ให้ถูกตีเป็นสูตร */
function txt_(v) {
  const s = String(v == null ? '' : v).slice(0, 100);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function num_(v) {
  const n = Math.round(Number(v));
  return isFinite(n) ? n : 0;
}

/** รันครั้งเดียวจากตัวแก้ไข Apps Script เพื่อสร้างแท็บตัวอย่างชุดคำถาม */
function setupSample() {
  const ss = ss_();
  if (ss.getSheetByName('ตัวอย่างชุดคำถาม')) return;
  const sh = ss.insertSheet('ตัวอย่างชุดคำถาม');
  sh.getRange(1, 1, 3, 8).setValues([
    ['หัวข้อ', 'ระดับ', 'คำถาม', 'คำตอบที่ถูก', 'ตัวเลือกผิด 1', 'ตัวเลือกผิด 2', 'ตัวเลือกผิด 3', 'คำอธิบาย'],
    ['ระบบสุริยะ', 'กลาง', 'ดาวเคราะห์ดวงใดใหญ่ที่สุดในระบบสุริยะ?', 'ดาวพฤหัสบดี', 'ดาวเสาร์', 'ดาวยูเรนัส', 'ดาวเนปจูน', 'ดาวพฤหัสบดีมีมวลมากกว่าดาวเคราะห์ดวงอื่นรวมกัน'],
    ['คณิตศาสตร์', 'ง่าย', '2 + 3 เท่ากับเท่าไร?', '5', '4', '6', '', '2 บวก 3 ได้ 5']
  ]);
  sh.setFrozenRows(1);
}
