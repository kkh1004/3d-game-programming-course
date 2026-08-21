/**
 * 프로젝트팀 생성기(index3.html) 공유 데이터 백엔드.
 *
 * 배포 방법
 *  1. https://script.google.com 에서 새 프로젝트를 만들고 이 파일 내용을 붙여넣습니다.
 *  2. 배포 > 새 배포 > 유형: 웹 앱
 *       - 실행 사용자: 나
 *       - 액세스 권한: 모든 사용자
 *  3. 처음 배포할 때 Google 계정 권한 승인 (교수 계정 1회, 학생은 불필요)
 *  4. 발급된 웹앱 URL(.../exec)을 index3.html 의 API_URL 에 넣습니다.
 *
 * 데이터는 스크립트가 자동 생성한 구글 시트 1칸에 JSON 문자열로 저장됩니다.
 */

var ROLES = ['기획', '아트', '개발'];
var INDIVIDUAL_GRADES = ['2', '3'];
var DEFAULT_PASSWORD = '1004';
var SHEET_NAME = 'state';
var PROP_SHEET_ID = 'teamProjectSheetId';

function emptyState() {
  return { roster: [], teamsByGrade: {}, assignedMap: {}, evaluations: {}, evalEnabled: false };
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
    ss = SpreadsheetApp.create('프로젝트팀 생성기 데이터');
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
  if (!state.roster) state.roster = [];
  if (!state.teamsByGrade) state.teamsByGrade = {};
  if (!state.assignedMap) state.assignedMap = {};
  if (!state.evaluations) state.evaluations = {};
  if (typeof state.evalEnabled !== 'boolean') state.evalEnabled = false;
  var pw = sheet.getRange('B1').getValue();
  return { state: state, password: pw ? String(pw) : DEFAULT_PASSWORD };
}

function writeStore(state, password) {
  var sheet = getSheet();
  sheet.getRange('A1').setValue(JSON.stringify(state));
  if (password) sheet.getRange('B1').setValue(password);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  var store = readStore();
  return jsonOut({ ok: true, state: store.state, hint: maskPassword(store.password) });
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
    if (!result.ok) return jsonOut(result);
    writeStore(store.state, result.newPassword);
    // 비밀번호가 바뀌었으면 바뀐 값 기준으로 힌트를 내려준다
    var effectivePw = result.newPassword || store.password;
    return jsonOut({
      ok: true,
      state: store.state,
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
 * state 변형 로직 — index3.html 클라이언트와 동일한 규칙
 * ============================================================ */

function ensureGradeState(state, grade) {
  if (!state.teamsByGrade[grade]) {
    state.teamsByGrade[grade] = {
      teamCount: 0,
      teams: [],
      individual: INDIVIDUAL_GRADES.indexOf(grade) !== -1 ? [] : undefined
    };
  }
  return state.teamsByGrade[grade];
}

function removeStudentFromAllSlots(state, grade, studentId) {
  var gd = state.teamsByGrade[grade];
  if (gd) {
    (gd.teams || []).forEach(function (t) {
      ROLES.forEach(function (role) {
        t.roles[role] = (t.roles[role] || []).filter(function (id) { return id !== studentId; });
      });
    });
    if (gd.individual) {
      gd.individual = gd.individual.filter(function (id) { return id !== studentId; });
    }
  }
  delete state.assignedMap[studentId];
}

function applyAction(store, action, payload, password) {
  var state = store.state;

  switch (action) {
    // 명단 업로드 / 학년 배정 초기화 / 전체 리셋 / 비밀번호 변경은 담당 교수 전용.
    // 팀 배정과 기여도 평가는 학생이 직접 하므로 비밀번호를 요구하지 않는다.
    case 'uploadRoster':
      if (password !== store.password) return { ok: false, error: 'wrong password' };
      if (!payload.roster || !payload.roster.length) return { ok: false, error: 'empty roster' };
      state.roster = payload.roster;
      state.teamsByGrade = {};
      state.assignedMap = {};
      state.evaluations = {};
      return { ok: true, message: 'roster replaced' };

    case 'mergeRoster': {
      if (password !== store.password) return { ok: false, error: 'wrong password' };
      if (!payload.roster || !payload.roster.length) return { ok: false, error: 'empty roster' };
      var newIds = {};
      payload.roster.forEach(function (s) { newIds[s.studentId] = true; });
      state.roster.forEach(function (s) {
        if (!newIds[s.studentId]) {
          var info = state.assignedMap[s.studentId];
          if (info) removeStudentFromAllSlots(state, info.grade, s.studentId);
        }
      });
      state.roster = payload.roster;
      return { ok: true, message: 'roster merged' };
    }

    case 'assignStudent': {
      var grade = String(payload.grade);
      removeStudentFromAllSlots(state, grade, payload.studentId);
      var gd = ensureGradeState(state, grade);
      if (payload.target === 'individual') {
        gd.individual.push(payload.studentId);
        state.assignedMap[payload.studentId] = { grade: grade, target: 'individual' };
      } else {
        var team = gd.teams.filter(function (t) { return t.id === payload.target; })[0];
        if (!team) return { ok: false, error: 'team not found' };
        team.roles[payload.role].push(payload.studentId);
        state.assignedMap[payload.studentId] = { grade: grade, target: payload.target, role: payload.role };
      }
      return { ok: true, message: 'assigned' };
    }

    case 'removeStudent':
      removeStudentFromAllSlots(state, String(payload.grade), payload.studentId);
      return { ok: true, message: 'removed' };

    case 'setTeamCount': {
      var g = String(payload.grade);
      var newCount = Math.max(0, Math.min(30, parseInt(payload.count, 10) || 0));
      var gd2 = ensureGradeState(state, g);
      var oldCount = gd2.teamCount;
      if (newCount > oldCount) {
        for (var i = oldCount + 1; i <= newCount; i++) {
          gd2.teams.push({ id: 't' + i, name: i + '팀', roles: { '기획': [], '아트': [], '개발': [] } });
        }
      } else if (newCount < oldCount) {
        gd2.teams.slice(newCount).forEach(function (t) {
          ROLES.forEach(function (role) {
            (t.roles[role] || []).forEach(function (sid) { delete state.assignedMap[sid]; });
          });
        });
        gd2.teams = gd2.teams.slice(0, newCount);
      }
      gd2.teamCount = newCount;
      return { ok: true, message: 'team count set' };
    }

    case 'clearGrade': {
      if (password !== store.password) return { ok: false, error: 'wrong password' };
      var g2 = String(payload.grade);
      var gd3 = state.teamsByGrade[g2];
      if (gd3) {
        (gd3.teams || []).forEach(function (t) {
          ROLES.forEach(function (role) {
            (t.roles[role] || []).forEach(function (sid) { delete state.assignedMap[sid]; });
            t.roles[role] = [];
          });
        });
        if (gd3.individual) {
          gd3.individual.forEach(function (sid) { delete state.assignedMap[sid]; });
          gd3.individual = [];
        }
      }
      return { ok: true, message: 'grade cleared' };
    }

    // 평가 활성화/마감은 담당 교수만. 활성화된 동안에만 학생이 평가를 낼 수 있다.
    case 'setEvalEnabled':
      if (password !== store.password) return { ok: false, error: 'wrong password' };
      state.evalEnabled = !!payload.enabled;
      return { ok: true, message: state.evalEnabled ? 'evaluation opened' : 'evaluation closed' };

    case 'submitEvaluation':
      if (!state.evalEnabled) return { ok: false, error: 'evaluation not open' };
      if (!payload.studentId || !payload.scores) return { ok: false, error: 'bad payload' };
      state.evaluations[payload.studentId] = payload.scores;
      return { ok: true, message: 'evaluation saved' };

    case 'resetAll':
      if (password !== store.password) return { ok: false, error: 'wrong password' };
      store.state = emptyState();
      return { ok: true, message: 'reset' };

    case 'changePassword':
      if (password !== store.password) return { ok: false, error: 'wrong password' };
      if (!payload.newPassword) return { ok: false, error: 'empty password' };
      return { ok: true, message: 'password changed', newPassword: payload.newPassword };

    case 'verifyPassword':
      if (password !== store.password) return { ok: false, error: 'wrong password' };
      return { ok: true, message: 'verified' };

    default:
      return { ok: false, error: 'unknown action' };
  }
}
