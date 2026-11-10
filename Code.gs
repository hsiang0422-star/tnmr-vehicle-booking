/**
 * 國立臺南特殊教育學校 派車申請系統（Google Apps Script 網頁應用程式）
 *
 * 第一次使用：在 Apps Script 編輯器選 setup 執行一次，
 * 會自動建立「南特派車行事曆」、「派車申請紀錄」試算表、資料夾與暫用 Word 範本。
 * 之後依 README 部署為網頁應用程式。
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

const BOOKING_HEADERS = [
  '申請編號', '送出時間', '申請人Email', '申請人', '處室', '聯絡電話',
  '出發時間', '返回時間', '集合地點', '目的地', '事由',
  '一般學生', '輪椅生', '隨行教師', '總人數', '車輛',
  '狀態', '行事曆事件ID', '申請單文件ID', 'Word檔ID', '備註',
];

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
      summary: CFG.SCHOOL + ' 派車預約行事曆（由派車申請系統自動維護）',
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
    vehicles.getRange(1, 1, 4, 7).setValues([
      ['車輛代號', '車輛名稱', '車牌', '座位數(不含駕駛)', '輪椅位', '啟用', '說明'],
      ['A', '一號車', '待填', 20, 2, true, '數字待確認'],
      ['B', '二號車', '待填', 15, 4, true, '數字待確認'],
      ['C', '三號車', '待填', 8, 1, true, '數字待確認'],
    ]);
    vehicles.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e8eef7');
    vehicles.getRange(2, 6, 3, 1).insertCheckboxes();
    vehicles.setFrozenRows(1);

    const settings = ss.insertSheet(CFG.SHEET_SETTINGS);
    settings.getRange(1, 1, 3, 3).setValues([
      ['項目', '值', '說明'],
      ['管理者Email', Session.getEffectiveUser().getEmail(), '可看全部申請、取消任何申請；多位用逗號分隔'],
      ['通知副本Email', '', '每筆申請成功時副本寄送（例如總務處、學務處交通車承辦），多位用逗號分隔'],
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
  DocumentApp.openById(docId); // 確認可開啟
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

/** 判斷某個行事曆事件佔用哪些車輛：系統建立的看標籤；手動建立的看標題是否含【代號】或車名。 */
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
  };
}

/** 查詢時段內每台車是否可用，以及當天所有派車。 */
function getAvailability(q) {
  requireUser_();
  const s = parseDT_(q.startDate, q.startTime);
  const e = parseDT_(q.endDate, q.endTime);
  if (!(e > s)) return { ok: false, msg: '返回時間需晚於出發時間' };

  const vehicles = getVehicles_();
  const busy = {};
  overlapping_(s, e).forEach(ev => {
    const info = eventInfo_(ev, vehicles);
    info.vehicles.forEach(code => { (busy[code] = busy[code] || []).push(info); });
  });

  const dayStart = parseDT_(q.startDate, '00:00');
  const dayEnd = new Date(parseDT_(q.endDate, '00:00').getTime() + 24 * 3600 * 1000);
  const day = cal_().getEvents(dayStart, dayEnd)
    .sort((a, b) => a.getStartTime() - b.getStartTime())
    .map(ev => eventInfo_(ev, vehicles));

  return {
    ok: true,
    vehicles: vehicles.map(v => Object.assign({}, v, { conflicts: busy[v.code] || [] })),
    day: day,
  };
}

function submitBooking(f) {
  const email = requireUser_();

  // ---- 檢查欄位 ----
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
    // ---- 送出當下再檢查一次衝突，避免兩人同時預約 ----
    const codes = chosen.map(v => v.code);
    const clash = overlapping_(s, e).filter(ev => eventVehicles_(ev, vehicles).some(c => codes.indexOf(c) >= 0));
    if (clash.length) {
      throw new Error('所選車輛在這個時段已被預約：' + clash.map(ev => ev.getTitle()).join('、') + '。請重新查詢後再送出。');
    }

    const sheet = ss_().getSheetByName(CFG.SHEET_BOOKINGS);
    const now = new Date();
    const prefix = 'P' + (Number(fmt_(now, 'yyyy')) - 1911) + fmt_(now, 'MMdd') + '-';
    const ids = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
    const seq = ids.filter(x => x.indexOf(prefix) === 0).length + 1;
    const id = prefix + ('00' + seq).slice(-3);

    const carText = chosen.map(v => v.name + (v.plate && v.plate !== '待填' ? '（' + v.plate + '）' : '')).join('、');
    const total = students + wheel + teachers;
    const data = {
      '申請編號': id,
      '申請日期': roc_(now),
      '申請人': f.name.trim(),
      '處室': f.dept.trim(),
      'Email': email,
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

    // ---- 行事曆 ----
    const ev = cal_().createEvent(
      '【' + codes.join('、') + '】' + data['目的地'] + '－' + data['申請人'] + '（' + data['處室'] + '）',
      s, e,
      {
        location: data['目的地'],
        description: [
          '申請編號：' + id,
          '事由：' + data['事由'],
          '申請人：' + data['申請人'] + '（' + data['處室'] + '）' + email + ' ' + data['聯絡電話'],
          '集合地點：' + data['集合地點'],
          '人數：一般學生 ' + students + '、輪椅生 ' + wheel + '、隨行教師 ' + teachers + '，共 ' + total + ' 人',
          '車輛：' + carText,
          data['備註'] ? '備註：' + data['備註'] : '',
        ].filter(Boolean).join('\n'),
        guests: email,
        sendInvites: false,
      }
    );
    ev.setTag('vehicles', codes.join(','));
    ev.setTag('bookingId', id);
    ev.setTag('applicant', email);

    // ---- 申請單（Google 文件 → Word） ----
    const docId = fillTemplate_(id + '_派車申請單_' + data['申請人'], data);
    const docx = exportBlob_(docId, 'docx', id + '_派車申請單.docx');
    const docxFile = DriveApp.getFolderById(prop_('DOCS_FOLDER_ID')).createFile(docx);

    sheet.appendRow([
      id, now, email, data['申請人'], data['處室'], data['聯絡電話'],
      s, e, data['集合地點'], data['目的地'], data['事由'],
      students, wheel, teachers, total, carText,
      '已預約', ev.getId(), docId, docxFile.getId(), data['備註'],
    ]);

    // ---- 寄信 ----
    let mailed = true;
    try {
      const cc = splitEmails_(settings_()['通知副本Email']).filter(x => x !== email).join(',');
      MailApp.sendEmail({
        to: email,
        cc: cc,
        name: '南特派車申請系統',
        subject: '【派車預約成功】' + id + '｜' + rocRange_(s, e) + '｜' + data['目的地'],
        htmlBody:
          '<p>' + data['申請人'] + ' 老師您好：</p>' +
          '<p>您的派車申請已預約成功，申請單 Word 檔已附在這封信中，請列印後依流程核章。</p>' +
          '<table cellpadding="6" style="border-collapse:collapse;border:1px solid #ccc">' +
          ['申請編號', '用車時間', '集合地點', '目的地', '事由', '車輛'].map(k =>
            '<tr><td style="background:#f2f2f2;border:1px solid #ccc"><b>' + k + '</b></td><td style="border:1px solid #ccc">' + escapeHtml_(data[k]) + '</td></tr>').join('') +
          '<tr><td style="background:#f2f2f2;border:1px solid #ccc"><b>人數</b></td><td style="border:1px solid #ccc">一般學生 ' + students + '、輪椅生 ' + wheel + '、隨行教師 ' + teachers + '，共 ' + total + ' 人</td></tr>' +
          '</table><p style="color:#666">此信由系統自動寄出。如需取消，請到派車申請網頁的「我的申請」操作。</p>',
        attachments: [docx],
      });
    } catch (err) {
      mailed = false;
      console.error(err);
    }

    return { id: id, when: rocRange_(s, e), cars: carText, mailed: mailed };
  } finally {
    lock.releaseLock();
  }
}

function listBookings(all) {
  const email = requireUser_();
  const admin = isAdmin_(email);
  const sheet = ss_().getSheetByName(CFG.SHEET_BOOKINGS);
  if (sheet.getLastRow() < 2) return { isAdmin: admin, rows: [] };
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, BOOKING_HEADERS.length).getValues();
  const now = new Date();
  const rows = values
    .filter(r => (all && admin) || String(r[2]).toLowerCase() === email)
    .map(r => ({
      id: String(r[0]),
      applicant: r[3] + '（' + r[4] + '）',
      when: rocRange_(new Date(r[6]), new Date(r[7])),
      start: new Date(r[6]).getTime(),
      destination: String(r[9]),
      purpose: String(r[10]),
      people: '學生 ' + r[11] + '／輪椅 ' + r[12] + '／教師 ' + r[13],
      cars: String(r[15]),
      status: String(r[16]),
      canCancel: String(r[16]) === '已預約' && new Date(r[6]) > now,
    }))
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

function cancelBooking(id) {
  const email = requireUser_();
  const found = findRow_(id);
  checkOwner_(email, found.r);
  if (String(found.r[16]) !== '已預約') throw new Error('這筆申請目前狀態為「' + found.r[16] + '」，不能取消');
  try {
    const ev = cal_().getEventById(String(found.r[17]));
    if (ev) ev.deleteEvent();
  } catch (e) { console.error(e); }
  found.sheet.getRange(found.row, 17).setValue('已取消');
  found.sheet.getRange(found.row, 21).setValue(
    String(found.r[20] || '') + (found.r[20] ? '；' : '') + '取消：' + email + ' ' + fmt_(new Date(), 'yyyy/MM/dd HH:mm'));
  return true;
}

/* ======================== 文件處理 ======================== */

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
      part.replaceText('\\{\\{[^}]*\\}\\}', ''); // 範本中未使用的欄位清空
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
