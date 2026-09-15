/**
 * FLEX-day 비교과 프로그램 학과별 온라인 접수 시스템 - Apps Script 백엔드
 * -----------------------------------------------------------------
 * 이 스크립트를 "이 스프레드시트에 바인딩된" Apps Script 프로젝트로 붙여넣고
 * 웹앱으로 배포하면, index.html이 이 URL로 제출/조회/상태변경 요청을 보냅니다.
 *
 * 배포 방법은 README.md를 참고하세요.
 */

// ⚠️ 반드시 각각 변경하세요. (완전한 보안 수단은 아니며, PIN을 아는 사람만 접근한다는 정도의 가벼운 방어입니다.)
// - 전체 관리자 PIN: 조회 + 검토(검토중/승인/반려) + 삭제 + 접수 마감·재개까지 기존 관리자 권한 전부
// - 센터 관리자 PIN: 목록 조회만 가능 (상태 변경/삭제/접수 마감·재개 불가)
const FULL_ADMIN_PIN = '0070';
const CENTER_ADMIN_PIN = '7563';

// 입력된 PIN이 어느 역할에 해당하는지 판별 ('full' | 'center' | null)
function adminRole_(pin) {
  if (pin === FULL_ADMIN_PIN) return 'full';
  if (pin === CENTER_ADMIN_PIN) return 'center';
  return null;
}

const SHEET_NAME = '신청현황';

const HEADERS = [
  'id', '접수일시', '학과', '교수명',
  '희망일자1', '희망시간1', '희망일자2', '희망시간2', '희망장소', '참여대상', '예상인원',
  '협조요청',
  '상태', '검토자', '검토일시', '검토의견'
];

function doGet(e) {
  try {
    const action = (e.parameter.action || 'list');
    if (action === 'ping') return json({ ok: true });
    if (action === 'status') return json({ ok: true, accepting: isAccepting_() });
    if (action === 'list') {
      const role = adminRole_(e.parameter.pin);
      if (!role) return json({ error: 'unauthorized' });
      return json({ ok: true, items: listApplications(), role: role });
    }
    if (action === 'lookup') {
      // 신청자가 로그인 없이 신청번호만으로 자기 신청 현황을 확인하는 용도.
      // 신청번호를 정확히 알아야만 해당 1건만 조회되므로 목록 전체 노출 위험은 없음.
      const item = findApplicationById_(e.parameter.id || '');
      if (!item) return json({ error: 'id_not_found' });
      return json({ ok: true, item: publicView_(item) });
    }
    return json({ error: 'unknown_action' });
  } catch (err) {
    return json({ error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    const action = body.action;

    if (action === 'submit') {
      if (!isAccepting_()) return json({ error: 'closed' });
      const id = submitApplication(body.data || {});
      return json({ ok: true, id: id });
    }

    if (action === 'updateStatus') {
      const role = adminRole_(body.pin);
      if (!role) return json({ error: 'unauthorized' });
      if (role !== 'full') return json({ error: 'forbidden' }); // 센터 관리자는 상태 변경 불가
      updateStatus(body.id, body.status, body.reviewer || '', body.comment || '');
      return json({ ok: true });
    }

    if (action === 'setStatus') {
      const role = adminRole_(body.pin);
      if (!role) return json({ error: 'unauthorized' });
      if (role !== 'full') return json({ error: 'forbidden' }); // 센터 관리자는 접수 마감/재개 불가
      setAccepting_(!!body.accepting);
      return json({ ok: true, accepting: isAccepting_() });
    }

    if (action === 'delete') {
      const role = adminRole_(body.pin);
      if (!role) return json({ error: 'unauthorized' });
      if (role !== 'full') return json({ error: 'forbidden' }); // 센터 관리자는 삭제 불가
      deleteApplication(body.id);
      return json({ ok: true });
    }

    return json({ error: 'unknown_action' });
  } catch (err) {
    return json({ error: String(err) });
  }
}

// 접수 마감/재개 상태는 스프레드시트가 아니라 스크립트 속성(Script Properties)에 저장합니다.
function isAccepting_() {
  const v = PropertiesService.getScriptProperties().getProperty('ACCEPTING');
  return v === null ? true : v === 'true'; // 한 번도 설정한 적 없으면 기본값: 접수중
}
function setAccepting_(accepting) {
  PropertiesService.getScriptProperties().setProperty('ACCEPTING', accepting ? 'true' : 'false');
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  ensureHeaders_(sheet);
  return sheet;
}

// 헤더 행이 아직 없으면 새로 만들고, 이미 있으면 HEADERS 중 그 시트에 없는 항목만
// 맨 뒤에 이어붙인다. 기존 열은 순서·내용을 그대로 두므로, 나중에 신청서 항목을
// 추가/삭제/이름 변경해도 이미 쌓인 데이터가 깨지지 않는다. (README "항목을 바꾸고
// 싶을 때" 참고)
function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    return HEADERS.slice();
  }
  const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = HEADERS.filter(function (h) { return existing.indexOf(h) === -1; });
  if (missing.length) {
    sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
    return existing.concat(missing);
  }
  return existing;
}

function submitApplication(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet_();
    // 실제 저장은 (혹시 예전 항목이 남아있는) 시트의 현재 헤더 행 순서를 그대로 따른다.
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const now = new Date();
    const id = 'FLEX-' + Utilities.formatDate(now, 'Asia/Seoul', 'yyyyMMdd-HHmmss') +
      '-' + Math.floor(Math.random() * 900 + 100);

    const row = headers.map(function (h) {
      if (h === 'id') return id;
      // 앞에 붙인 작은따옴표(')는 구글 시트가 날짜/숫자로 자동 변환하지 못하게 "텍스트 강제" 표시입니다.
      // (실제 저장/조회되는 값에는 따옴표가 남지 않습니다.)
      if (h === '접수일시') return "'" + Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
      if (h === '희망일자1' || h === '희망일자2') return data[h] ? "'" + data[h] : '';
      if (h === '상태') return '접수';
      if (h === '검토자' || h === '검토일시' || h === '검토의견') return '';
      return data[h] != null ? data[h] : '';
    });

    sheet.appendRow(row);
    return id;
  } finally {
    lock.releaseLock();
  }
}

function listApplications() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .map(function (r) {
      const obj = {};
      headers.forEach(function (h, i) { obj[h] = r[i]; });
      return obj;
    })
    .filter(function (o) { return o.id; })
    .sort(function (a, b) {
      return String(b['접수일시'] || '').localeCompare(String(a['접수일시'] || ''));
    });
}

function findApplicationById_(id) {
  if (!id) return null;
  const items = listApplications();
  for (let i = 0; i < items.length; i++) {
    if (items[i].id === id) return items[i];
  }
  return null;
}

// 신청자 조회 화면에 내보낼 항목만 추림 (검토자 이름 등 내부 정보는 제외)
function publicView_(it) {
  return {
    id: it['id'],
    학과: it['학과'],
    교수명: it['교수명'],
    희망일자1: it['희망일자1'],
    희망시간1: it['희망시간1'],
    희망일자2: it['희망일자2'],
    희망시간2: it['희망시간2'],
    희망장소: it['희망장소'],
    접수일시: it['접수일시'],
    상태: it['상태'],
    검토의견: it['검토의견'] || ''
  };
}

function updateStatus(id, status, reviewer, comment) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet_();
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const idCol = headers.indexOf('id');
    const statusCol = headers.indexOf('상태');
    const reviewerCol = headers.indexOf('검토자');
    const reviewedAtCol = headers.indexOf('검토일시');
    const commentCol = headers.indexOf('검토의견');

    for (let i = 1; i < values.length; i++) {
      if (values[i][idCol] === id) {
        const rowIndex = i + 1;
        sheet.getRange(rowIndex, statusCol + 1).setValue(status);
        sheet.getRange(rowIndex, reviewerCol + 1).setValue(reviewer);
        sheet.getRange(rowIndex, reviewedAtCol + 1).setValue(
          "'" + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm')
        );
        sheet.getRange(rowIndex, commentCol + 1).setValue(comment);
        return;
      }
    }
    throw new Error('id_not_found');
  } finally {
    lock.releaseLock();
  }
}

function deleteApplication(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet_();
    const values = sheet.getDataRange().getValues();
    const idCol = values[0].indexOf('id');
    for (let i = 1; i < values.length; i++) {
      if (values[i][idCol] === id) {
        sheet.deleteRow(i + 1);
        return;
      }
    }
    throw new Error('id_not_found');
  } finally {
    lock.releaseLock();
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
