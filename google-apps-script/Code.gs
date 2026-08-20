const CONFIG = Object.freeze({
  SHEET_NAME: '문의',
  ADMIN_EMAIL: 'dabe0001@naver.com'
});

const HEADERS = [
  '견적번호', '접수일시', '상태', '이름', '연락처', '이메일',
  '희망 컨셉', '기념일·웨딩 날짜', '제작 인원', '완성본 장수',
  '고난도 합성', '수정 횟수', '맞춤 배경', '빠른 제작',
  '예상 견적', '요청사항', '마케팅 동의', '선택한 사진명', '접수 페이지'
];

function setupSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#2E2823')
      .setFontColor('#F7F3EC');
  }

  return sheet;
}

function doPost(event) {
  try {
    const data = event && event.parameter ? event.parameter : {};
    if (data.website) return jsonResponse({ success: true, inquiryId: '' });

    requireFields(data, ['name', 'phone', 'email', 'concept', 'estimateTotal']);

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    let inquiryId;
    try {
      const sheet = setupSheet();
      inquiryId = createInquiryId();
      sheet.appendRow([
        inquiryId,
        new Date(),
        '신규',
        safeCell(data.name),
        safeCell(data.phone),
        safeCell(data.email),
        safeCell(data.concept),
        safeCell(data.weddingDate || '미정'),
        toNumber(data.peopleCount),
        toNumber(data.imageCount),
        toNumber(data.complexCount),
        toNumber(data.revisionCount),
        safeCell(data.customBackground),
        safeCell(data.rushOrder),
        toNumber(data.estimateTotal),
        safeCell(data.message || ''),
        safeCell(data.marketing || '미동의'),
        safeCell(data.selectedFiles || '없음'),
        safeCell(data.sourceUrl || '')
      ]);
    } finally {
      lock.releaseLock();
    }

    sendAdminNotification(inquiryId, data);
    return jsonResponse({ success: true, inquiryId: inquiryId });
  } catch (error) {
    console.error(error);
    return jsonResponse({ success: false, message: '문의 접수 중 오류가 발생했습니다.' });
  }
}

function createInquiryId() {
  const timezone = Session.getScriptTimeZone() || 'Asia/Seoul';
  const timestamp = Utilities.formatDate(new Date(), timezone, 'yyMMdd-HHmmss');
  return 'DG-' + timestamp + '-' + Utilities.getUuid().slice(0, 4).toUpperCase();
}

function requireFields(data, fields) {
  fields.forEach(function (field) {
    if (!String(data[field] || '').trim()) throw new Error('필수값 누락: ' + field);
  });
}

function toNumber(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function safeCell(value) {
  const text = String(value == null ? '' : value).trim();
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function escapeHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function won(value) {
  return Number(value || 0).toLocaleString('ko-KR') + '원';
}

function sendAdminNotification(inquiryId, data) {
  const rows = [
    ['견적번호', inquiryId],
    ['이름', data.name],
    ['연락처', data.phone],
    ['회신 이메일', data.email],
    ['희망 컨셉', data.concept],
    ['제작 인원', data.peopleCount + '명'],
    ['완성본', data.imageCount + '장'],
    ['고난도 합성', data.complexCount + '장'],
    ['수정 횟수', data.revisionCount + '회'],
    ['맞춤 배경', data.customBackground],
    ['빠른 제작', data.rushOrder],
    ['예상 견적', won(data.estimateTotal)],
    ['요청사항', data.message || '없음']
  ];

  const htmlRows = rows.map(function (row) {
    return '<tr><th style="padding:8px 12px;text-align:left;background:#f5f1ea">' +
      escapeHtml(row[0]) + '</th><td style="padding:8px 12px">' +
      escapeHtml(row[1]) + '</td></tr>';
  }).join('');

  MailApp.sendEmail({
    to: CONFIG.ADMIN_EMAIL,
    subject: '[다시, 그날] 새 견적 문의 ' + inquiryId,
    body: rows.map(function (row) { return row[0] + ': ' + row[1]; }).join('\n'),
    htmlBody: '<h2>새 견적 문의가 접수되었습니다.</h2><table style="border-collapse:collapse">' + htmlRows + '</table>',
    name: '다시, 그날 문의 알림'
  });
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
