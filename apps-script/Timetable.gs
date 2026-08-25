/**
 * 시간표 생성기(index4.html) 공유 데이터 백엔드.
 *
 * 프로젝트팀 관리 시스템(Code.gs)과는 **별개의 Apps Script 프로젝트**로 배포합니다.
 * 서로 영향을 주지 않도록 데이터 시트도 따로 만들어집니다.
 *
 * 배포 방법
 *  1. https://script.google.com 에서 새 프로젝트를 만들고 이 파일 내용을 붙여넣습니다.
 *  2. 배포 > 새 배포 > 유형: 웹 앱
 *       - 실행 사용자: 나
 *       - 액세스 권한: 모든 사용자
 *  3. 처음 배포할 때 Google 계정 권한 승인 (교수 계정 1회, 보는 사람은 불필요)
 *  4. 발급된 웹앱 URL(.../exec)을 index4.html 의 API_URL 에 넣습니다.
 *
 * 코드를 고친 뒤에는 반드시 "배포 관리 > 편집 > 버전: 새 버전"으로 다시 배포해야
 * 바뀐 내용이 반영됩니다. (URL은 그대로 유지됩니다)
 *
 * 데이터는 스크립트가 자동 생성한 구글 시트 한 줄에 저장됩니다.
 *   A1 = 시간표 상태 JSON / B1 = 비밀번호 / C1 = 리비전 번호
 */

var DEFAULT_PASSWORD = '1004';
var SHEET_NAME = 'state';
var PROP_SHEET_ID = 'timetableSheetId';
var GRADES = ['1', '2', '3'];

function emptyState() {
  return {
    subjects: [],
    rooms: [],
    totalWeeks: 15,
    boardsByGrade: { '1': [], '2': [], '3': [] }
  };
}

// 비밀번호 힌트(첫 글자 + * + 끝 글자). 어느 PC에서 접속하든 서버 기준으로
// 같은 힌트를 보여주기 위해 서버에서 만들어 내려준다.
function maskPassword(pw) {
  pw = String(pw || '');
  if (pw.length <= 2) return pw;
  return pw.charAt(0) + repeatChar('*', pw.length - 2) + pw.charAt(pw.length - 1);
}

function repeatChar(ch, n) {
  var s = '';
  for (var i = 0; i < n; i++) s += ch;
  return s;
}

function getSheet() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(PROP_SHEET_ID);
  var ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('시간표 생성기 데이터');
    props.setProperty(PROP_SHEET_ID, ss.getId());
  }
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  return sheet;
}

function readStore() {
  var sheet = getSheet();
  var raw = sheet.getRange('A1').getValue();
  var state;
  try {
    state = raw ? JSON.parse(raw) : emptyState();
  } catch (e) {
    state = emptyState();
  }
  state = normalizeState(state);

  var pw = sheet.getRange('B1').getValue();
  var rev = parseInt(sheet.getRange('C1').getValue(), 10);
  return {
    state: state,
    password: pw ? String(pw) : DEFAULT_PASSWORD,
    rev: isNaN(rev) ? 0 : rev
  };
}

/* 저장된 값이 비었거나 예전 형식이어도 화면이 깨지지 않도록 모양을 맞춘다.
 * index4.html 의 loadState() 와 같은 규칙. */
function normalizeState(state) {
  if (!state || typeof state !== 'object') state = emptyState();
  if (!state.subjects) state.subjects = [];
  if (!state.rooms) state.rooms = [];
  if (!state.totalWeeks) state.totalWeeks = 15;
  if (!state.boardsByGrade) state.boardsByGrade = {};
  GRADES.forEach(function (g) {
    if (!state.boardsByGrade[g]) state.boardsByGrade[g] = [];
  });
  return state;
}

function writeStore(state, password, rev) {
  var sheet = getSheet();
  sheet.getRange('A1').setValue(JSON.stringify(state));
  if (password) sheet.getRange('B1').setValue(password);
  sheet.getRange('C1').setValue(rev);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  var store = readStore();
  return jsonOut({
    ok: true,
    state: store.state,
    rev: store.rev,
    hint: maskPassword(store.password)
  });
}

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ ok: false, error: 'bad request' });
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return jsonOut({ ok: false, error: 'busy' });
  }

  try {
    var store = readStore();
    var result = applyAction(store, req.action, req.payload || {}, req.password);

    // 실패해도 최신 상태를 함께 내려준다. 충돌(stale)일 때 클라이언트가
    // 한 번의 왕복으로 바로 서버 내용에 맞출 수 있다.
    if (!result.ok) {
      return jsonOut({
        ok: false,
        error: result.error,
        state: store.state,
        rev: store.rev,
        hint: maskPassword(store.password)
      });
    }

    if (result.changed) store.rev = store.rev + 1;
    if (result.changed || result.newPassword) {
      writeStore(store.state, result.newPassword, store.rev);
    }

    // 비밀번호가 바뀌었으면 바뀐 값 기준으로 힌트를 내려준다
    var effectivePw = result.newPassword || store.password;
    return jsonOut({
      ok: true,
      state: store.state,
      rev: store.rev,
      message: result.message,
      hint: maskPassword(effectivePw)
    });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * state 변형 로직
 *
 * 시간표는 프로젝트팀과 달리 **편집자가 사실상 한 명**(담당 교수)이라
 * 항목별 액션을 두지 않고 화면의 상태 전체를 통째로 저장한다.
 * 대신 리비전(rev)으로 덮어쓰기 사고를 막는다:
 *   클라이언트는 자기가 마지막으로 받은 rev를 baseRev로 함께 보내고,
 *   그 사이에 다른 사람이 저장했으면(rev 불일치) 거부한다.
 * ============================================================ */

function applyAction(store, action, payload, password) {
  switch (action) {
    // 시간표 저장은 담당 교수 전용. 보기만 하는 사람은 doGet으로 읽기만 한다.
    case 'saveState': {
      if (password !== store.password) return { ok: false, error: 'wrong password' };
      if (!payload.state) return { ok: false, error: 'empty state' };
      // baseRev 가 없다는 건 아직 서버 내용을 한 번도 못 받았다는 뜻이다.
      // 이미 저장된 시간표가 있는데 그걸 덮어쓰게 두면 안 된다.
      if (payload.baseRev == null) {
        if (store.rev > 0) return { ok: false, error: 'stale' };
      } else if (Number(payload.baseRev) !== store.rev) {
        return { ok: false, error: 'stale' };
      }
      store.state = normalizeState(payload.state);
      return { ok: true, message: 'saved', changed: true };
    }

    case 'resetAll':
      if (password !== store.password) return { ok: false, error: 'wrong password' };
      store.state = emptyState();
      return { ok: true, message: 'reset', changed: true };

    case 'changePassword':
      if (password !== store.password) return { ok: false, error: 'wrong password' };
      if (!payload.newPassword) return { ok: false, error: 'empty password' };
      return { ok: true, message: 'password changed', newPassword: String(payload.newPassword) };

    case 'verifyPassword':
      if (password !== store.password) return { ok: false, error: 'wrong password' };
      return { ok: true, message: 'verified' };

    default:
      return { ok: false, error: 'unknown action' };
  }
}
