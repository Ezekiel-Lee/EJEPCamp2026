// ============================================================
// EJEP Camp 2026 — Google Apps Script 백엔드
// ============================================================

const SHEET_NAME = '명단';
const ADMIN_PASSWORD = 'ejep2026'; // ★ 나중에 변경하세요!

// 컬럼 인덱스 (1부터 시작)
const COL = {
  NAME: 1,        // 이름
  PHONE: 2,       // 전화번호
  FAMILY: 3,      // 가족ID
  ROOM: 4,         // 방배정
  TEAM: 5,      // 팀배정
  NOTE: 6,        // Dietary
  CHECKED_IN: 7,  // 체크인여부
  CHECKED_AT: 8  // 체크인시각
};

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const params = e.parameter;
  const action = params.action;
  const method = e.requestMethod || 'GET';

  try {
    let result;

    switch (action) {
      case 'stats':
        result = getStats();
        break;
      case 'checkin':
        result = checkin(JSON.parse(e.postData.contents));
        break;
      case 'checkinFamily':
        result = checkinFamily(JSON.parse(e.postData.contents));
        break;
      case 'getMembers':
        result = getMembers(params.password);
        break;
      case 'toggleCheckin':
        result = toggleCheckin(JSON.parse(e.postData.contents));
        break;
      case 'export':
        result = exportCSV(params.password);
        break;
      default:
        result = { error: '알 수 없는 액션입니다.' };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 헬퍼 ──────────────────────────────────────────────────
function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function getLastRow() {
  const sheet = getSheet();
  return sheet.getLastRow();
}

function getAllMembers() {
  const sheet = getSheet();
  const lastRow = getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  return data.map((row, i) => ({
    id: i + 2,
    name: String(row[0] || '').trim(),
    phone: String(row[1] || '').replace(/\D/g, ''),
    familyGroup: String(row[2] || '').trim(),
    room: String(row[3] || '').trim(),
    team: String(row[4] || '').trim(),
    note: String(row[5] || '').trim(),
    checkedIn: row[6] === true || row[6] === 'TRUE' || row[6] === '✓',
    checkedInAt: row[7] ? String(row[7]) : null
  })).filter(m => m.name !== '');
}

// ── 통계 ──────────────────────────────────────────────────
function getStats() {
  const members = getAllMembers();
  const total = members.length;
  const attended = members.filter(m => m.checkedIn).length;
  return { total, attended, absent: total - attended };
}

// ── 체크인 ────────────────────────────────────────────────
function checkin(body) {
  const { name, phoneLast3 } = body;
  const members = getAllMembers();

  const member = members.find(m =>
    m.name === name.trim() &&
    m.phone.slice(-3) === phoneLast3.trim()
  );

  if (!member) {
    return { error: '명단에서 찾을 수 없습니다. 이름과 전화번호를 확인해주세요.' };
  }

  if (member.checkedIn) {
    return { error: 'ALREADY', member };
  }

  // 체크인 처리
  const sheet = getSheet();
  const now = new Date().toISOString();
  sheet.getRange(member.id, COL.CHECKED_IN).setValue('✓');
  sheet.getRange(member.id, COL.CHECKED_AT).setValue(now);
  member.checkedIn = true;
  member.checkedInAt = now;

  // 가족 조회
  let family = [];
  if (member.familyGroup) {
    family = members.filter(m =>
      m.familyGroup === member.familyGroup && m.id !== member.id
    );
  }

  return { success: true, member, family };
}

// ── 가족 일괄 체크인 ──────────────────────────────────────
function checkinFamily(body) {
  const { ids } = body;
  const sheet = getSheet();
  const now = new Date().toISOString();
  const members = getAllMembers();
  const checkedMembers = [];

  ids.forEach(id => {
    const member = members.find(m => m.id === id);
    if (member && !member.checkedIn) {
      sheet.getRange(id, COL.CHECKED_IN).setValue('✓');
      sheet.getRange(id, COL.CHECKED_AT).setValue(now);
      member.checkedIn = true;
      member.checkedInAt = now;
      checkedMembers.push(member);
    }
  });

  return { success: true, members: checkedMembers };
}

// ── 관리자: 전체 명단 ─────────────────────────────────────
function getMembers(password) {
  if (password !== ADMIN_PASSWORD) return { error: '인증 실패' };
  return { members: getAllMembers() };
}

// ── 관리자: 체크인 토글 ───────────────────────────────────
function toggleCheckin(body) {
  const { id, checkedIn, password } = body;
  if (password !== ADMIN_PASSWORD) return { error: '인증 실패' };

  const sheet = getSheet();
  const now = checkedIn ? new Date().toISOString() : '';
  sheet.getRange(id, COL.CHECKED_IN).setValue(checkedIn ? '✓' : '');
  sheet.getRange(id, COL.CHECKED_AT).setValue(now);

  const members = getAllMembers();
  const member = members.find(m => m.id === id);
  return { success: true, member };
}

// ── 관리자: CSV 내보내기 ──────────────────────────────────
function exportCSV(password) {
  if (password !== ADMIN_PASSWORD) return { error: '인증 실패' };
  const members = getAllMembers();
  const headers = ['이름','전화번호','가족그룹ID','방배정','팀배정','Dietary','체크인여부','체크인시각'];
  const rows = members.map(m => [
    m.name, m.phone, m.familyGroup,
    m.room, m.team, m.note,
    m.checkedIn ? '✓' : '',
    m.checkedInAt ? new Date(m.checkedInAt).toLocaleString('ko-KR') : ''
  ]);
  return { headers, rows };
}

// ── 초기 헤더 설정 (최초 1회 실행) ───────────────────────
function setupHeaders() {
  const sheet = getSheet() || SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);
  const headers = ['이름','전화번호','가족그룹ID','방배정','팀배정','Dietary','체크인여부','체크인시각'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  SpreadsheetApp.flush();
  Logger.log('헤더 설정 완료!');
}
