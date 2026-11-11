/**
 * 國立臺南特殊教育學校 派車申請系統（Google Apps Script 網頁應用程式）
 *
 * 流程：老師送出 → 試算表記為「待審核」→ 總務處（管理者）核准 → 排入「南特派車行事曆」並寄出含 Word 的通知
 *
 * 第一次使用：在 Apps Script 編輯器選 setup 執行一次。
 */

const CFG = {
  DOMAIN: 'tnmr.tn.edu.tw',
  SCHOOL: '國立臺南特殊教育學校',
  TZ: 'Asia/Taipei',
  CAL_NAME: '南特派車行事曆',
  SHEET_BOOKINGS: '申請紀錄',
  SHEET_VEHICLES: '車輛',
  SHEET_SETTINGS: '設定',
};

// 欄位順序很重要：程式用位置存取（第 1 欄 = index 0）
const BOOKING_HEADERS = [
  '申請編號', '送出時間', '申請人Email', '申請人', '處室', '聯絡電話',   // 0-5
  '出發時間', '返回時間', '集合地點', '目的地', '事由',                  // 6-10
  '一般學生', '輪椅生', '隨行教師', '總人數', '車輛',                    // 11-15
  '狀態', '行事曆事件ID', '申請單文件ID', 'Word檔ID', '備註',            // 16-20
  '審核人', '審核時間', '車輛代號',                                      // 21-23
];

const ST = { WAIT: '待審核', OK: '已核准', BACK: '已退回', CANCEL: '已取消' };

/* ======================== 網頁進入點 ======================== */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('南特派車申請')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/* ======================== 初始化（管理者執行一次） ======================== */

function setup() {
  const p = PropertiesService.getScriptProperties();

  let folder;
  if (p.getProperty('FOLDER_ID')) {
    folder = DriveApp.getFolderById(p.getProperty('FOLDER_ID'));
  } else {
    folder = DriveApp.createFolder('派車申請系統_資料');
    p.setProperty('FOLDER_ID', folder.getId());
  }

  if (!p.getProperty('DOCS_FOLDER_ID')) {
    p.setProperty('DOCS_FOLDER_ID', folder.createFolder('申請單').getId());
  }

  if (!p.getProperty('CAL_ID')) {
    const cal = CalendarApp.createCalendar(CFG.CAL_NAME, {
      summary: CFG.SCHOOL + ' 派車行事曆（藍色為已核准、黃色為待審，由派車申請系統自動維護）',
      timeZone: CFG.TZ,
      color: CalendarApp.Color.BLUE,
    });
    p.setProperty('CAL_ID', cal.getId());
  }

  if (!p.getProperty('SHEET_ID')) {
    const ss = SpreadsheetApp.create('派車申請紀錄');
    DriveApp.getFileById(ss.getId()).moveTo(folder);

    const bookings = ss.getSheets()[0].setName(CFG.SHEET_BOOKINGS);
    bookings.getRange(1, 1, 1, BOOKING_HEADERS.length).setValues([BOOKING_HEADERS])
      .setFontWeight('bold').setBackground('#e8eef7');
    bookings.setFrozenRows(1);

    const vehicles = ss.insertSheet(CFG.SHEET_VEHICLES);
    vehicles.getRange(1, 1, 6, 7).setValues([
      ['車輛代號', '車輛名稱', '車牌', '座位數', '輪椅位', '啟用', '說明'],
      ['L', '大型巴士', '待填', 29, 3, true, ''],
      ['MA', '中型巴士A', '待填', 12, 1, true, ''],
      ['MB', '中型巴士B', '待填', 12, 2, true, ''],
      ['MC', '中型巴士C', '待填', 9, 2, true, ''],
      ['S', '小校車', '待填', 6, 0, true, ''],
    ]);
    vehicles.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e8eef7');
    vehicles.getRange(2, 6, 5, 1).insertCheckboxes();
    vehicles.setFrozenRows(1);

    const settings = ss.insertSheet(CFG.SHEET_SETTINGS);
    settings.getRange(1, 1, 3, 3).setValues([
      ['項目', '值', '說明'],
      ['管理者Email', Session.getEffectiveUser().getEmail(), '總務處人員：負責核准／退回申請，可看全部申請。多位用逗號分隔'],
      ['通知副本Email', '', '每筆申請核准時副本寄送（例如學務處交通車承辦），多位用逗號分隔'],
    ]);
    settings.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8eef7');
    settings.setColumnWidth(2, 320);
    settings.setColumnWidth(3, 420);

    p.setProperty('SHEET_ID', ss.getId());
  }

  if (!p.getProperty('TEMPLATE_ID')) {
    p.setProperty('TEMPLATE_ID', createDefaultTemplate_(folder));
  }

  Logger.log('完成。資料夾：' + folder.getUrl());
  Logger.log('試算表：' + SpreadsheetApp.openById(p.getProperty('SHEET_ID')).getUrl());
  Logger.log('範本：' + DocumentApp.openById(p.getProperty('TEMPLATE_ID')).getUrl());
}

/** 改用正式範本：把 Word 範本存成 Google 文件後，把文件 ID 填進來執行一次。 */
function setTemplate(docId) {
  DocumentApp.openById(docId);
  PropertiesService.getScriptProperties().setProperty('TEMPLATE_ID', docId);
  Logger.log('已改用範本：' + docId);
}

function createDefaultTemplate_(folder) {
  const doc = DocumentApp.create('派車申請單_範本（暫用）');
  const body = doc.getBody();
  body.setMarginTop(50).setMarginBottom(50).setMarginLeft(56).setMarginRight(56);

  body.appendParagraph(CFG.SCHOOL).setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setFontSize(16).setBold(true);
  body.appendParagraph('派車申請單').setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setFontSize(18).setBold(true);
  body.appendParagraph('申請編號：{{申請編號}}　　申請日期：{{申請日期}}')
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setFontSize(11).setBold(false);

  const rows = [
    ['申請人', '{{申請人}}（{{處室}}）'],
    ['聯絡電話', '{{聯絡電話}}'],
    ['用車時間', '{{用車時間}}'],
    ['集合地點', '{{集合地點}}'],
    ['目的地', '{{目的地}}'],
    ['事由', '{{事由}}'],
    ['乘車人數', '一般學生 {{一般學生}} 人、輪椅生 {{輪椅生}} 人、隨行教師 {{隨行教師}} 人，合計 {{總人數}} 人'],
    ['派用車輛', '{{車輛}}'],
    ['備註', '{{備註}}'],
  ];
  const t = body.appendTable(rows);
  for (let r = 0; r < t.getNumRows(); r++) {
    t.getCell(r, 0).setWidth(90).setBackgroundColor('#f2f2f2');
    for (let c = 0; c < 2; c++) {
      t.getCell(r, c).setPaddingTop(6).setPaddingBottom(6).editAsText().setFontSize(12).setBold(c === 0);
    }
  }

  body.appendParagraph('');
  const sign = body.appendTable([
    ['申請人', '處室主任', '總務主任', '校長'],
    ['\n\n', '\n\n', '\n\n', '\n\n'],
  ]);
  for (let c = 0; c < 4; c++) {
    sign.getCell(0, c).setBackgroundColor('#f2f2f2').editAsText().setBold(true).setFontSize(12);
    sign.getCell(0, c).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }

  try { body.editAsText().setFontFamily('DFKai-SB'); } catch (e) { /* 字型不支援時維持預設 */ }

  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(folder);
  return doc.getId();
}

/* ======================== 共用工具 ======================== */

function prop_(k) {
  const v = PropertiesService.getScriptProperties().getProperty(k);
  if (!v) throw new Error('系統尚未初始化，請管理者先執行 setup()。');
  return v;
}
function ss_() { return SpreadsheetApp.openById(prop_('SHEET_ID')); }
function cal_() { return CalendarApp.getCalendarById(prop_('CAL_ID')); }

function requireUser_() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email || !email.endsWith('@' + CFG.DOMAIN)) {
    throw new Error('請使用學校帳號（@' + CFG.DOMAIN + '）登入後再使用。');
  }
  return email;
}

function settings_() {
  const values = ss_().getSheetByName(CFG.SHEET_SETTINGS).getDataRange().getValues();
  const out = {};
  values.slice(1).forEach(r => { out[String(r[0]).trim()] = String(r[1] || '').trim(); });
  return out;
}
function splitEmails_(s) {
  return String(s || '').split(/[,，;；\s]+/).map(x => x.trim().toLowerCase()).filter(Boolean);
}
function isAdmin_(email) {
  return splitEmails_(settings_()['管理者Email']).indexOf(email) >= 0;
}
function requireAdmin_() {
  const email = requireUser_();
  if (!isAdmin_(email)) throw new Error('只有總務處管理者可以執行此操作。');
  return email;
}

function getVehicles_() {
  const values = ss_().getSheetByName(CFG.SHEET_VEHICLES).getDataRange().getValues();
  return values.slice(1)
    .filter(r => String(r[0]).trim() && r[5] !== false && String(r[5]).toUpperCase() !== 'FALSE')
    .map(r => ({
      code: String(r[0]).trim(),
      name: String(r[1]).trim(),
      plate: String(r[2] || '').trim(),
      seats: Number(r[3]) || 0,
      wheel: Number(r[4]) || 0,
      note: String(r[6] || '').trim(),
    }));
}

function allRows_() {
  const sh = ss_().getSheetByName(CFG.SHEET_BOOKINGS);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, BOOKING_HEADERS.length).getValues();
}

function parseDT_(date, time) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date || '') || !/^\d{2}:\d{2}$/.test(time || '')) {
    throw new Error('日期或時間格式不正確');
  }
  return Utilities.parseDate(date + ' ' + time, CFG.TZ, 'yyyy-MM-dd HH:mm');
}
function fmt_(d, p) { return Utilities.formatDate(d, CFG.TZ, p); }
function roc_(d) {
  const y = Number(fmt_(d, 'yyyy')) - 1911;
  const w = '日一二三四五六'[Number(fmt_(d, 'u')) % 7];
  return y + '年' + Number(fmt_(d, 'M')) + '月' + Number(fmt_(d, 'd')) + '日（' + w + '）';
}
function rocRange_(s, e) {
  const sameDay = fmt_(s, 'yyyyMMdd') === fmt_(e, 'yyyyMMdd');
  return roc_(s) + ' ' + fmt_(s, 'HH:mm') + ' 至 ' + (sameDay ? '' : roc_(e) + ' ') + fmt_(e, 'HH:mm');
}

/** 行事曆事件佔用哪些車：系統建立的看標籤；手動建立的看標題是否含【代號】或車名。 */
function eventVehicles_(ev, vehicles) {
  const tag = ev.getTag('vehicles');
  if (tag) return tag.split(',').filter(Boolean);
  const title = ev.getTitle();
  return vehicles
    .filter(v => title.indexOf('【' + v.code + '】') >= 0 || (v.name && title.indexOf(v.name) >= 0))
    .map(v => v.code);
}

function eventInfo_(ev, vehicles) {
  return {
    title: ev.getTitle(),
    start: fmt_(ev.getStartTime(), 'MM/dd HH:mm'),
    end: fmt_(ev.getEndTime(), 'MM/dd HH:mm'),
    vehicles: eventVehicles_(ev, vehicles),
    bookingId: ev.getTag('bookingId') || '',
    status: isApproved_(ev) ? 'approved' : 'pending',
  };
}

function overlapping_(s, e) {
  return cal_().getEvents(s, e).filter(ev => ev.getStartTime() < e && ev.getEndTime() > s);
}

/* ======================== 前端呼叫的 API ======================== */

function getInit() {
  const email = requireUser_();
  return {
    email: email,
    isAdmin: isAdmin_(email),
    school: CFG.SCHOOL,
    vehicles: getVehicles_(),
    today: fmt_(new Date(), 'yyyy-MM-dd'),
    calId: prop_('CAL_ID'),
  };
}

/** 查詢時段內每台車：已核准的派車（紅）與待審中的申請（黃）。 */
function getAvailability(q) {
  requireUser_();
  const s = parseDT_(q.startDate, q.startTime);
  const e = parseDT_(q.endDate, q.endTime);
  if (!(e > s)) return { ok: false, msg: '返回時間需晚於出發時間' };

  const vehicles = getVehicles_();
  const busy = {};
  overlapping_(s, e).filter(isApproved_).forEach(ev => {
    const info = eventInfo_(ev, vehicles);
    info.vehicles.forEach(code => { (busy[code] = busy[code] || []).push(info); });
  });

  const pending = {};
  allRows_().forEach(r => {
    if (String(r[16]) !== ST.WAIT) return;
    const rs = new Date(r[6]), re = new Date(r[7]);
    if (!(rs < e && re > s)) return;
    const info = {
      title: r[9] + '－' + r[3] + '（' + r[4] + '）',
      start: fmt_(rs, 'MM/dd HH:mm'),
      end: fmt_(re, 'MM/dd HH:mm'),
    };
    String(r[23] || '').split(',').filter(Boolean).forEach(code => { (pending[code] = pending[code] || []).push(info); });
  });

  const dayStart = parseDT_(q.startDate, '00:00');
  const dayEnd = new Date(parseDT_(q.endDate, '00:00').getTime() + 24 * 3600 * 1000);
  const day = cal_().getEvents(dayStart, dayEnd)
    .sort((a, b) => a.getStartTime() - b.getStartTime())
    .map(ev => eventInfo_(ev, vehicles));

  return {
    ok: true,
    vehicles: vehicles.map(v => Object.assign({}, v, { conflicts: busy[v.code] || [], pending: pending[v.code] || [] })),
    day: day,
  };
}

/** 老師送出：只寫入試算表並產生申請單，狀態為「待審核」，不排行事曆。 */
function submitBooking(f) {
  const email = requireUser_();

  const need = ['name', 'dept', 'destination', 'purpose'];
  need.forEach(k => { if (!String(f[k] || '').trim()) throw new Error('請填寫所有必填欄位'); });
  const s = parseDT_(f.startDate, f.startTime);
  const e = parseDT_(f.endDate, f.endTime);
  if (!(e > s)) throw new Error('返回時間需晚於出發時間');
  if (s < new Date()) throw new Error('出發時間已經過去，請重新選擇');

  const students = Math.max(0, parseInt(f.students, 10) || 0);
  const wheel = Math.max(0, parseInt(f.wheel, 10) || 0);
  const teachers = Math.max(0, parseInt(f.teachers, 10) || 0);
  if (students + wheel + teachers === 0) throw new Error('乘車人數不可為 0');

  const vehicles = getVehicles_();
  const chosen = vehicles.filter(v => (f.vehicles || []).indexOf(v.code) >= 0);
  if (!chosen.length) throw new Error('請至少選擇一台車');

  const seats = chosen.reduce((a, v) => a + v.seats, 0);
  const wheelSlots = chosen.reduce((a, v) => a + v.wheel, 0);
  if (students + teachers > seats) throw new Error('座位不足：需要 ' + (students + teachers) + ' 個座位，所選車輛只有 ' + seats + ' 個');
  if (wheel > wheelSlots) throw new Error('輪椅位不足：需要 ' + wheel + ' 個，所選車輛只有 ' + wheelSlots + ' 個');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const codes = chosen.map(v => v.code);
    const clash = overlapping_(s, e).filter(ev => isApproved_(ev) && eventVehicles_(ev, vehicles).some(c => codes.indexOf(c) >= 0));
    if (clash.length) {
      throw new Error('所選車輛在這個時段已有核准的派車：' + clash.map(ev => ev.getTitle()).join('、') + '。請重新查詢後再送出。');
    }

    const sheet = ss_().getSheetByName(CFG.SHEET_BOOKINGS);
    const now = new Date();
    const prefix = 'P' + (Number(fmt_(now, 'yyyy')) - 1911) + fmt_(now, 'MMdd') + '-';
    const ids = allRows_().map(r => String(r[0]));
    const seq = ids.filter(x => x.indexOf(prefix) === 0).length + 1;
    const id = prefix + ('00' + seq).slice(-3);

    const carText = chosen.map(v => v.name + (v.plate && v.plate !== '待填' ? '（' + v.plate + '）' : '')).join('、');
    const total = students + wheel + teachers;
    const data = {
      '申請編號': id,
      '申請日期': roc_(now),
      '申請人': f.name.trim(),
      '處室': f.dept.trim(),
      '聯絡電話': String(f.phone || '').trim(),
      '用車時間': rocRange_(s, e),
      '集合地點': String(f.meet || '').trim() || '本校',
      '目的地': f.destination.trim(),
      '事由': f.purpose.trim(),
      '一般學生': String(students),
      '輪椅生': String(wheel),
      '隨行教師': String(teachers),
      '總人數': String(total),
      '車輛': carText,
      '備註': String(f.note || '').trim(),
    };

    const docId = fillTemplate_(id + '_派車申請單_' + data['申請人'], data);
    const docx = exportBlob_(docId, 'docx', id + '_派車申請單.docx');
    const docxFile = DriveApp.getFolderById(prop_('DOCS_FOLDER_ID')).createFile(docx);

    const row = [
      id, now, email, data['申請人'], data['處室'], data['聯絡電話'],
      s, e, data['集合地點'], data['目的地'], data['事由'],
      students, wheel, teachers, total, carText,
      ST.WAIT, '', docId, docxFile.getId(), data['備註'], '', '', codes.join(','),
    ];
    // 行事曆先放一筆黃色「待審」活動，核准後才改成藍色已核准
    const pendEv = cal_().createEvent(titleFor_(row, ST.WAIT), s, e, { location: row[9], description: eventDesc_(row) });
    pendEv.setTag('vehicles', codes.join(','));
    pendEv.setTag('bookingId', id);
    pendEv.setTag('status', 'pending');
    pendEv.setColor(CalendarApp.EventColor.YELLOW);
    row[17] = pendEv.getId();
    sheet.appendRow(row);

    // 通知老師（確認信，尚未附 Word；核准後才附）
    let mailed = true;
    try {
      MailApp.sendEmail({
        to: email,
        name: '南特派車申請系統',
        subject: '【派車申請已送出】' + id + '｜' + rocRange_(s, e) + '｜' + data['目的地'],
        htmlBody: '<p>' + escapeHtml_(data['申請人']) + ' 老師您好：</p>' +
          '<p>您的派車申請已送出，目前為<b>待審核</b>狀態。總務處核准後，系統會寄出含申請單 Word 檔的通知，並排入派車行事曆。</p>' +
          detailTable_({ id: id, s: s, e: e, meet: data['集合地點'], dest: data['目的地'], purpose: data['事由'], cars: carText, students: students, wheel: wheel, teachers: teachers, total: total }) +
          '<p style="color:#666">此信由系統自動寄出。</p>',
      });
    } catch (err) { mailed = false; console.error(err); }

    // 通知總務處管理者
    notifyAdmins_('【派車待審】' + id + '｜' + rocRange_(s, e) + '｜' + data['目的地'],
      '<p>有新的派車申請待審核：</p>' +
      detailTable_({ id: id, s: s, e: e, meet: data['集合地點'], dest: data['目的地'], purpose: data['事由'], cars: carText, students: students, wheel: wheel, teachers: teachers, total: total, who: data['申請人'] + '（' + data['處室'] + '）' }) +
      '<p><a href="' + ScriptApp.getService().getUrl() + '">前往派車申請系統審核</a></p>');

    return { id: id, when: rocRange_(s, e), cars: carText, mailed: mailed };
  } finally {
    lock.releaseLock();
  }
}

/** 總務處核准：檢查衝突 → 建立行事曆 → 寄含 Word 的通知。 */
function approveBooking(id) {
  const adminEmail = requireAdmin_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const found = findRow_(id);
    const r = found.r;
    if (String(r[16]) !== ST.WAIT) throw new Error('這筆目前狀態為「' + r[16] + '」，無法核准。');

    const s = new Date(r[6]), e = new Date(r[7]);
    const codes = String(r[23] || '').split(',').filter(Boolean);
    if (!codes.length) throw new Error('這筆申請沒有車輛代號，無法核准，請確認試算表。');
    const vehicles = getVehicles_();
    const clash = overlapping_(s, e).filter(ev => isApproved_(ev) && eventVehicles_(ev, vehicles).some(c => codes.indexOf(c) >= 0));
    if (clash.length) {
      throw new Error('核准失敗，所選車輛已有其他核准的派車：' + clash.map(ev => ev.getTitle()).join('、') + '。請退回此申請。');
    }

    // 把送出時建立的黃色待審活動改成藍色已核准；找不到就重建
    let ev = findEvent_(r[17]);
    if (ev) {
      ev.setTitle(titleFor_(r, ST.OK));
      ev.setDescription(eventDesc_(r));
    } else {
      ev = cal_().createEvent(titleFor_(r, ST.OK), s, e, { location: String(r[9]), description: eventDesc_(r) });
      ev.setTag('vehicles', codes.join(','));
      ev.setTag('bookingId', String(r[0]));
    }
    ev.setTag('status', 'approved');
    ev.setColor(CalendarApp.EventColor.BLUE);

    found.sheet.getRange(found.row, 17).setValue(ST.OK);
    found.sheet.getRange(found.row, 18).setValue(ev.getId());
    found.sheet.getRange(found.row, 22, 1, 2).setValues([[adminEmail, new Date()]]);

    const wordFile = DriveApp.getFileById(String(r[19])).getBlob().setName(r[0] + '_派車申請單.docx');
    mailTeacher_(r, '【派車已核准】' + r[0] + '｜' + rocRange_(s, e) + '｜' + r[9],
      '您的派車申請已經<b>核准</b>，並已排入派車行事曆。申請單 Word 檔已附在這封信中，請列印後依流程核章。',
      [wordFile]);
    return true;
  } finally {
    lock.releaseLock();
  }
}

/** 總務處退回：寫入原因並寄信給老師。 */
function rejectBooking(id, reason) {
  const adminEmail = requireAdmin_();
  reason = String(reason || '').trim();
  if (!reason) throw new Error('請填寫退回原因。');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const found = findRow_(id);
    const r = found.r;
    if (String(r[16]) !== ST.WAIT) throw new Error('這筆目前狀態為「' + r[16] + '」，無法退回。');
    const pev = findEvent_(r[17]);
    if (pev) pev.deleteEvent();
    found.sheet.getRange(found.row, 17).setValue(ST.BACK);
    found.sheet.getRange(found.row, 21).setValue(String(r[20] || '') + (r[20] ? '；' : '') + '退回原因：' + reason);
    found.sheet.getRange(found.row, 22, 1, 2).setValues([[adminEmail, new Date()]]);
    mailTeacher_(r, '【派車申請已退回】' + r[0] + '｜' + r[9],
      '很抱歉，您的派車申請<b>未能通過審核</b>。<br>退回原因：' + escapeHtml_(reason) + '<br>如需重新申請，請到派車申請網頁重新送出。',
      []);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function listBookings(all) {
  const email = requireUser_();
  const admin = isAdmin_(email);
  const now = new Date();
  const rows = allRows_()
    .filter(r => (all && admin) || String(r[2]).toLowerCase() === email)
    .map(r => {
      const st = String(r[16]);
      const start = new Date(r[6]);
      return {
        id: String(r[0]),
        applicant: r[3] + '（' + r[4] + '）',
        when: rocRange_(start, new Date(r[7])),
        start: start.getTime(),
        destination: String(r[9]),
        purpose: String(r[10]),
        people: '學生 ' + r[11] + '／輪椅 ' + r[12] + '／教師 ' + r[13],
        cars: String(r[15]),
        status: st,
        canCancel: (st === ST.WAIT || st === ST.OK) && start > now,
        canReview: admin && st === ST.WAIT,
      };
    })
    .sort((a, b) => b.start - a.start)
    .slice(0, 200);
  return { isAdmin: admin, rows: rows };
}

function findRow_(id) {
  const sheet = ss_().getSheetByName(CFG.SHEET_BOOKINGS);
  if (sheet.getLastRow() < 2) throw new Error('找不到申請：' + id);
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  if (i < 0) throw new Error('找不到申請：' + id);
  const row = i + 2;
  const r = sheet.getRange(row, 1, 1, BOOKING_HEADERS.length).getValues()[0];
  return { sheet: sheet, row: row, r: r };
}

function checkOwner_(email, r) {
  if (String(r[2]).toLowerCase() !== email && !isAdmin_(email)) throw new Error('只能操作自己的申請');
}

/** 下載申請單。fmt = 'docx'（Word）或 'pdf'（列印用）。 */
function getBookingFile(id, fmt) {
  const email = requireUser_();
  const found = findRow_(id);
  checkOwner_(email, found.r);
  const blob = fmt === 'pdf'
    ? DriveApp.getFileById(String(found.r[18])).getAs('application/pdf').setName(id + '_派車申請單.pdf')
    : DriveApp.getFileById(String(found.r[19])).getBlob().setName(id + '_派車申請單.docx');
  return packBlob_(blob);
}

/** 下載空白範本（Word）。 */
function getBlankTemplate() {
  requireUser_();
  const tmp = fillTemplate_('派車申請單_空白', {}, true);
  try {
    return packBlob_(exportBlob_(tmp, 'docx', '派車申請單_空白範本.docx'));
  } finally {
    DriveApp.getFileById(tmp).setTrashed(true);
  }
}

/** 取消：待審核直接取消；已核准則同時刪除行事曆活動。 */
function cancelBooking(id) {
  const email = requireUser_();
  const found = findRow_(id);
  checkOwner_(email, found.r);
  const st = String(found.r[16]);
  if (st !== ST.WAIT && st !== ST.OK) throw new Error('這筆申請目前狀態為「' + st + '」，不能取消。');
  const cev = findEvent_(found.r[17]);
  if (cev) cev.deleteEvent();
  found.sheet.getRange(found.row, 17).setValue(ST.CANCEL);
  found.sheet.getRange(found.row, 21).setValue(
    String(found.r[20] || '') + (found.r[20] ? '；' : '') + '取消：' + email + ' ' + fmt_(new Date(), 'yyyy/MM/dd HH:mm'));
  return true;
}

/* ======================== 行事曆輔助 ======================== */

/** 手動建立的活動沒有 status 標籤，視為已核准。 */
function isApproved_(ev) { return ev.getTag('status') !== 'pending'; }

function titleFor_(r, st) {
  const codes = String(r[23] || '').split(',').filter(Boolean).join('、');
  return (st === ST.WAIT ? '【待審】' : '') + '【' + codes + '】' + r[9] + '－' + r[3] + '（' + r[4] + '）';
}

function eventDesc_(r) {
  return [
    '申請編號：' + r[0],
    '事由：' + r[10],
    '申請人：' + r[3] + '（' + r[4] + '）' + r[2] + ' ' + r[5],
    '集合地點：' + r[8],
    '人數：一般學生 ' + r[11] + '、輪椅生 ' + r[12] + '、隨行教師 ' + r[13] + '，共 ' + r[14] + ' 人',
    '車輛：' + r[15],
    r[20] ? '備註：' + r[20] : '',
  ].filter(Boolean).join('\n');
}

function findEvent_(eid) {
  if (!eid) return null;
  try { return cal_().getEventById(String(eid)); } catch (e) { return null; }
}

/* ======================== 信件與文件 ======================== */

function detailTable_(d) {
  const rows = [
    ['申請編號', d.id],
    ['用車時間', rocRange_(d.s, d.e)],
    ['集合地點', d.meet],
    ['目的地', d.dest],
    ['事由', d.purpose],
    ['車輛', d.cars],
    ['人數', '一般學生 ' + d.students + '、輪椅生 ' + d.wheel + '、隨行教師 ' + d.teachers + '，共 ' + d.total + ' 人'],
  ];
  if (d.who) rows.splice(1, 0, ['申請人', d.who]);
  return '<table cellpadding="6" style="border-collapse:collapse;border:1px solid #ccc">' +
    rows.map(x => '<tr><td style="background:#f2f2f2;border:1px solid #ccc"><b>' + x[0] + '</b></td><td style="border:1px solid #ccc">' + escapeHtml_(x[1]) + '</td></tr>').join('') +
    '</table>';
}

/** 寄信給申請老師（核准或退回）。r 為試算表那一列。 */
function mailTeacher_(r, subject, leadHtml, attachments) {
  try {
    const cc = splitEmails_(settings_()['通知副本Email']).filter(x => x !== String(r[2]).toLowerCase()).join(',');
    MailApp.sendEmail({
      to: String(r[2]),
      cc: cc,
      name: '南特派車申請系統',
      subject: subject,
      htmlBody: '<p>' + escapeHtml_(r[3]) + ' 老師您好：</p><p>' + leadHtml + '</p>' +
        detailTable_({ id: r[0], s: new Date(r[6]), e: new Date(r[7]), meet: r[8], dest: r[9], purpose: r[10], cars: r[15], students: r[11], wheel: r[12], teachers: r[13], total: r[14] }) +
        '<p style="color:#666">此信由系統自動寄出。</p>',
      attachments: attachments || [],
    });
    return true;
  } catch (err) { console.error(err); return false; }
}

function notifyAdmins_(subject, html) {
  const admins = splitEmails_(settings_()['管理者Email']);
  if (!admins.length) return;
  try {
    MailApp.sendEmail({ to: admins.join(','), name: '南特派車申請系統', subject: subject, htmlBody: html });
  } catch (err) { console.error(err); }
}

/** 複製範本並填入資料，回傳新文件 ID。blank=true 時把所有 {{欄位}} 清成空白。 */
function fillTemplate_(name, data, blank) {
  const folder = DriveApp.getFolderById(prop_('DOCS_FOLDER_ID'));
  const copy = DriveApp.getFileById(prop_('TEMPLATE_ID')).makeCopy(name, folder);
  const doc = DocumentApp.openById(copy.getId());
  const parts = [doc.getBody(), doc.getHeader(), doc.getFooter()].filter(Boolean);
  parts.forEach(part => {
    if (blank) {
      part.replaceText('\\{\\{[^}]*\\}\\}', '');
    } else {
      Object.keys(data).forEach(k => {
        part.replaceText('\\{\\{' + k + '\\}\\}', String(data[k]).replace(/\$/g, '＄'));
      });
      part.replaceText('\\{\\{[^}]*\\}\\}', '');
    }
  });
  doc.saveAndClose();
  return copy.getId();
}

function exportBlob_(docId, format, name) {
  const res = UrlFetchApp.fetch('https://docs.google.com/document/d/' + docId + '/export?format=' + format, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error('匯出 ' + format + ' 失敗（' + res.getResponseCode() + '）');
  return res.getBlob().setName(name);
}

function packBlob_(blob) {
  return { name: blob.getName(), mime: blob.getContentType(), b64: Utilities.base64Encode(blob.getBytes()) };
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}
