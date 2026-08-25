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
 * 코드를 고친 뒤에는 반드시 "배포 관리 > 편집 > 버전: 새 버전"으로 다시
 * 배포해야 바뀐 내용이 반영됩니다. (URL은 그대로 유지됩니다)
 *
 * 데이터는 스크립트가 자동 생성한 구글 시트 1칸에 JSON 문자열로 저장됩니다.
 *   A1 = 상태 JSON / B1 = 교수 비밀번호
 *
 * ⚠️ 공개 범위
 *   doGet 과 doPost 는 publicState() 를 거친 데이터만 내려줍니다.
 *   평가 점수(evaluations)와 학생 비밀번호 해시(studentAuth)는 여기에
 *   포함되지 않습니다. 점수는 본인이 getMyEvaluation, 교수가 getResults
 *   로 각자의 비밀번호를 내고 따로 받아갑니다.
 */

var ROLES = ['기획', '아트', '개발'];
var INDIVIDUAL_GRADES = ['2', '3'];
var DEFAULT_PASSWORD = '1004';
var SHEET_NAME = 'state';
var PROP_SHEET_ID = 'teamProjectSheetId';

/* 반(수업반) 구분은 1학년만 사용한다. 2·3학년은 반 없이 팀 번호만 쓴다. */
var SECTION_GRADES = ['1'];
var SECTIONS = ['a', 'b', 'c'];
var TEAMS_PER_SECTION = 10;   // 1~10팀 = a반, 11~20팀 = b반, 21~30팀 = c반

function emptyState() {
  return {
    roster: [],
    teamsByGrade: {},
    assignedMap: {},
    evaluations: {},
    evalEnabled: false,
    studentAuth: {}   // 학번 → { salt, hash } — 학생이 직접 만든 평가용 비밀번호
  };
}

/* ============================================================
 * 학생 비밀번호 — 평문으로 저장하지 않는다.
 * 학번마다 다른 salt를 붙여 SHA-256으로 해시하므로, 시트를 열어봐도
 * (교수 포함 누구도) 학생의 비밀번호 자체는 알 수 없다.
 * 대신 잊어버렸을 때 되돌려줄 방법도 없으므로 교수가 초기화해 준다.
 * ============================================================ */
function newSalt() {
  return Utilities.getUuid();
}

function hashPassword(password, salt) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + '|' + String(password),
    Utilities.Charset.UTF_8
  );
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i] < 0 ? bytes[i] + 256 : bytes[i];
    var s = b.toString(16);
    hex += (s.length === 1 ? '0' : '') + s;
  }
  return hex;
}

/* 학생 비밀번호 확인. 아직 만들지 않았으면 'no password'를 돌려준다. */
function checkStudentPassword(state, studentId, password) {
  var rec = state.studentAuth[studentId];
  if (!rec || !rec.hash || !rec.salt) return 'no password';
  if (hashPassword(password, rec.salt) !== rec.hash) return 'wrong student password';
  return null;
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
  if (!state.studentAuth) state.studentAuth = {};
  if (typeof state.evalEnabled !== 'boolean') state.evalEnabled = false;
  migrateTeamSections(state);
  var pw = sheet.getRange('B1').getValue();
  return { state: state, password: pw ? String(pw) : DEFAULT_PASSWORD };
}

/* ============================================================
 * 반 정보 이관 (한 번만 효과가 있고, 여러 번 실행해도 안전하다)
 *
 * 반 필드가 생기기 전에는 팀 번호 범위로 반을 구분해서 쓰고 있었다.
 *   1~10팀 = a반 / 11~20팀 = b반 / 21~30팀 = c반
 * 이 규칙대로 기존 팀에 section을 채워 넣는다. 팀 배정 내용(roles)은
 * 건드리지 않고 필드만 추가하므로 배정된 학생은 그대로 남는다.
 * 이관 뒤에는 교수가 팀별로 반을 자유롭게 바꿀 수 있다.
 * ============================================================ */
function sectionForTeamNumber(n) {
  var idx = Math.floor((n - 1) / TEAMS_PER_SECTION);
  if (idx < 0) idx = 0;
  if (idx >= SECTIONS.length) idx = SECTIONS.length - 1;
  return SECTIONS[idx];
}

function teamNumberOf(teamId) {
  var n = parseInt(String(teamId).replace('t', ''), 10);
  return isNaN(n) ? 0 : n;
}

function migrateTeamSections(state) {
  SECTION_GRADES.forEach(function (grade) {
    var gd = state.teamsByGrade[grade];
    if (!gd) return;
    (gd.teams || []).forEach(function (t) {
      if (SECTIONS.indexOf(t.section) !== -1) return;   // 이미 값이 있으면 그대로 둔다
      t.section = sectionForTeamNumber(teamNumberOf(t.id));
    });
    migrateSectionCounts(gd);
  });
}

/* 반마다 쓰는 팀 번호 구간의 시작값. a=0(1~10), b=10(11~20), c=20(21~30).
 * 번호로 반을 알 수 있어야 하므로 구간은 고정이다. */
function sectionBase(section) {
  return SECTIONS.indexOf(section) * TEAMS_PER_SECTION;
}

function teamsInSection(gd, section) {
  return (gd.teams || []).filter(function (t) { return t.section === section; });
}

/* 팀 개수를 학년 단위로 세던 것을 반 단위로 바꾼다.
 * 이미 있는 팀을 반별로 세기만 하므로 팀 배정에는 영향이 없다. */
function migrateSectionCounts(gd) {
  if (!gd.teamCountBySection || typeof gd.teamCountBySection !== 'object') {
    gd.teamCountBySection = {};
  }
  SECTIONS.forEach(function (sec) {
    var n = parseInt(gd.teamCountBySection[sec], 10);
    if (isNaN(n) || n < 0 || n > TEAMS_PER_SECTION) {
      gd.teamCountBySection[sec] = teamsInSection(gd, sec).length;
    }
  });
  gd.teamCount = totalTeamCount(gd);
}

function totalTeamCount(gd) {
  var sum = 0;
  SECTIONS.forEach(function (sec) { sum += gd.teamCountBySection[sec] || 0; });
  return sum;
}

/* 팀 목록을 번호 순으로 정렬 — 반 구간을 섞어 만들어도 화면 순서가 흔들리지 않는다 */
function sortTeams(gd) {
  gd.teams.sort(function (a, b) { return teamNumberOf(a.id) - teamNumberOf(b.id); });
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

/* ============================================================
 * 공개 상태 — 누구나 받아가는 데이터.
 *
 * ⚠️ evaluations(누가 누구에게 몇 점)와 studentAuth(비밀번호 해시)는
 *    절대 여기에 넣지 않는다. 예전에는 state를 통째로 내려주고 있어서
 *    개발자도구나 API 주소만으로 전교생 점수를 볼 수 있었다.
 *    평가 점수는 본인(getMyEvaluation) 또는 교수(getResults)만
 *    비밀번호를 내고 따로 받아간다.
 *
 * 화면에 필요한 "누가 제출했는지 / 누가 비밀번호를 만들었는지"는
 * 점수가 아니라 학번 목록이므로 그대로 내려준다.
 * ============================================================ */
function publicState(state) {
  return {
    roster: state.roster,
    teamsByGrade: state.teamsByGrade,
    assignedMap: state.assignedMap,
    evalEnabled: state.evalEnabled,
    submittedIds: Object.keys(state.evaluations),
    authedIds: Object.keys(state.studentAuth)
  };
}

function doGet() {
  var store = readStore();
  return jsonOut({ ok: true, state: publicState(store.state), hint: maskPassword(store.password) });
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
    // 조회만 하는 동작(readOnly)은 시트에 쓰지 않는다
    if (!result.readOnly) writeStore(store.state, result.newPassword);
    // 비밀번호가 바뀌었으면 바뀐 값 기준으로 힌트를 내려준다
    var effectivePw = result.newPassword || store.password;
    var out = {
      ok: true,
      state: publicState(store.state),
      message: result.message,
      hint: maskPassword(effectivePw)
    };
    // 본인 평가 / 교수용 결과처럼 비밀번호를 확인한 뒤에만 주는 데이터
    if (result.data !== undefined) out.data = result.data;
    return jsonOut(out);
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

/* 학생을 모든 학년의 모든 슬롯에서 빼낸다.
 *
 * ⚠️ 예전에는 넘겨받은 학년의 보드만 훑었다. 2·3학년 학생을 1학년 팀에
 *    넣을 수 있게 되면서, 원래 있던 학년의 팀 명단에 학번이 그대로 남는
 *    유령 배정이 생겼다. 그래서 학년을 가리지 않고 전부 훑는다. */
function removeStudentFromAllSlots(state, studentId) {
  Object.keys(state.teamsByGrade).forEach(function (g) {
    var gd = state.teamsByGrade[g];
    if (!gd) return;
    (gd.teams || []).forEach(function (t) {
      ROLES.forEach(function (role) {
        t.roles[role] = (t.roles[role] || []).filter(function (id) { return id !== studentId; });
      });
    });
    if (gd.individual) {
      gd.individual = gd.individual.filter(function (id) { return id !== studentId; });
    }
  });
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
      state.studentAuth = {};   // 명단을 갈아치우면 학생 비밀번호도 의미가 없다
      return { ok: true, message: 'roster replaced' };

    case 'mergeRoster': {
      if (password !== store.password) return { ok: false, error: 'wrong password' };
      if (!payload.roster || !payload.roster.length) return { ok: false, error: 'empty roster' };
      var newIds = {};
      payload.roster.forEach(function (s) { newIds[s.studentId] = true; });
      state.roster.forEach(function (s) {
        if (!newIds[s.studentId]) removeStudentFromAllSlots(state, s.studentId);
      });
      state.roster = payload.roster;
      return { ok: true, message: 'roster merged' };
    }

    /* 배정 — grade 는 "배정되는 팀이 속한 학년"이고, 학생의 명단상 학년과
     * 다를 수 있다. (2·3학년 학생이 1학년 팀에 들어가는 경우)
     * 명단(roster)의 grade 는 건드리지 않으므로 원 소속은 그대로 남는다. */
    case 'assignStudent': {
      var grade = String(payload.grade);
      removeStudentFromAllSlots(state, payload.studentId);
      var gd = ensureGradeState(state, grade);
      if (payload.target === 'individual') {
        if (!gd.individual) return { ok: false, error: 'no individual slot' };
        gd.individual.push(payload.studentId);
        state.assignedMap[payload.studentId] = { grade: grade, target: 'individual' };
      } else {
        var team = gd.teams.filter(function (t) { return t.id === payload.target; })[0];
        if (!team) return { ok: false, error: 'team not found' };
        if (ROLES.indexOf(payload.role) === -1) return { ok: false, error: 'bad role' };
        team.roles[payload.role].push(payload.studentId);
        state.assignedMap[payload.studentId] = { grade: grade, target: payload.target, role: payload.role };
      }
      return { ok: true, message: 'assigned' };
    }

    case 'removeStudent':
      removeStudentFromAllSlots(state, payload.studentId);
      return { ok: true, message: 'removed' };

    /* 팀 개수 — 반을 쓰지 않는 학년(2·3학년)만 학년 단위로 정한다.
     * 1학년은 반마다 따로 정하므로 setSectionTeamCount 를 쓴다. */
    case 'setTeamCount': {
      var g = String(payload.grade);
      if (SECTION_GRADES.indexOf(g) !== -1) {
        return { ok: false, error: 'use section team count' };
      }
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

    /* 반별 팀 개수 (1학년) — 반마다 자기 번호 구간 안에서 앞에서부터 채운다.
     *   a반 t1~t10 / b반 t11~t20 / c반 t21~t30
     * 줄이면 그 반의 뒤쪽 팀부터 사라지고, 거기 배정됐던 학생은 미배정으로
     * 돌아간다. 다른 반의 팀은 건드리지 않는다. */
    case 'setSectionTeamCount': {
      var sg2 = String(payload.grade);
      if (SECTION_GRADES.indexOf(sg2) === -1) return { ok: false, error: 'grade has no sections' };
      if (SECTIONS.indexOf(payload.section) === -1) return { ok: false, error: 'bad section' };
      var sec2 = payload.section;
      var want = Math.max(0, Math.min(TEAMS_PER_SECTION, parseInt(payload.count, 10) || 0));
      var sgd2 = ensureGradeState(state, sg2);
      migrateSectionCounts(sgd2);
      var base = sectionBase(sec2);
      var have = teamsInSection(sgd2, sec2).length;

      if (want > have) {
        for (var j = have + 1; j <= want; j++) {
          var num = base + j;
          sgd2.teams.push({
            id: 't' + num, name: num + '팀', section: sec2,
            roles: { '기획': [], '아트': [], '개발': [] }
          });
        }
        sortTeams(sgd2);
      } else if (want < have) {
        var cutoff = base + want;
        var doomed = sgd2.teams.filter(function (t) {
          return t.section === sec2 && teamNumberOf(t.id) > cutoff;
        });
        doomed.forEach(function (t) {
          ROLES.forEach(function (role) {
            (t.roles[role] || []).forEach(function (sid) { delete state.assignedMap[sid]; });
          });
        });
        sgd2.teams = sgd2.teams.filter(function (t) {
          return !(t.section === sec2 && teamNumberOf(t.id) > cutoff);
        });
      }

      sgd2.teamCountBySection[sec2] = want;
      sgd2.teamCount = totalTeamCount(sgd2);
      return { ok: true, message: 'section team count set' };
    }

    /* 팀명(title) 지정 — "1팀/2팀"이라는 순번(name)은 고정이고, 그 옆에
     * 붙는 팀 이름만 바꾼다. 팀 배정과 같은 성격이라 비밀번호는 없다. */
    case 'setTeamTitle': {
      var rg = String(payload.grade);
      var rgd = state.teamsByGrade[rg];
      if (!rgd) return { ok: false, error: 'grade not found' };
      var rteam = (rgd.teams || []).filter(function (t) { return t.id === payload.teamId; })[0];
      if (!rteam) return { ok: false, error: 'team not found' };
      rteam.title = String(payload.title == null ? '' : payload.title).trim().slice(0, 30);
      // 순번 이름이 지워졌거나 바뀌어 있으면 기본값으로 되돌린다
      rteam.name = String(payload.teamId).replace('t', '') + '팀';
      return { ok: true, message: 'team title set' };
    }

    /* 팀의 반을 개별로 바꾸는 setTeamSection 은 없앴다.
     * 반은 팀 번호 구간(a=1~10, b=11~20, c=21~30)으로 정해진다. 이 규칙이
     * 있어야 반별 팀 개수를 세고 새 팀을 어느 번호로 만들지 알 수 있다.
     * 임의로 반만 바꾸면 번호와 반이 어긋나 개수 계산이 깨진다.
     * 반을 옮기려면 한쪽 반 개수를 줄이고 다른 쪽을 늘린다. */

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

    /* ========================================================
     * 학생 본인 확인 — 평가는 본인만 낼 수 있고, 본인만 볼 수 있다.
     *
     * 예전에는 학번만 보내면 누구든 남의 이름으로 평가를 제출하거나
     * 덮어쓸 수 있었다. 이제 학생이 처음 평가할 때 스스로 비밀번호를
     * 만들고, 그 뒤로는 매번 그 비밀번호로 본인을 증명한다.
     * ======================================================== */

    // 최초 1회 — 학번+이름이 명단과 일치할 때만 만들 수 있다
    case 'setStudentPassword': {
      var apId = String(payload.studentId || '');
      var apName = String(payload.name || '');
      if (!apId || !apName) return { ok: false, error: 'bad payload' };
      if (!payload.password || String(payload.password).length < 4) {
        return { ok: false, error: 'password too short' };
      }
      var apStudent = state.roster.filter(function (s) {
        return String(s.studentId) === apId && String(s.name) === apName;
      })[0];
      if (!apStudent) return { ok: false, error: 'student not found' };
      if (state.studentAuth[apId]) return { ok: false, error: 'password exists' };
      var apSalt = newSalt();
      state.studentAuth[apId] = { salt: apSalt, hash: hashPassword(payload.password, apSalt) };
      return { ok: true, message: 'student password created' };
    }

    // 본인 비밀번호로 자기가 낸 평가만 받아간다 (남의 점수는 내려주지 않는다)
    case 'getMyEvaluation': {
      var myId = String(payload.studentId || '');
      var myErr = checkStudentPassword(state, myId, payload.password);
      if (myErr) return { ok: false, error: myErr };
      return {
        ok: true,
        readOnly: true,
        message: 'my evaluation',
        data: { scores: state.evaluations[myId] || null }
      };
    }

    case 'submitEvaluation': {
      if (!state.evalEnabled) return { ok: false, error: 'evaluation not open' };
      if (!payload.studentId || !payload.scores) return { ok: false, error: 'bad payload' };
      var subId = String(payload.studentId);
      var subErr = checkStudentPassword(state, subId, payload.password);
      if (subErr) return { ok: false, error: subErr };
      state.evaluations[subId] = payload.scores;
      return { ok: true, message: 'evaluation saved' };
    }

    // 교수용 — 전체 평가 결과. 교수 비밀번호를 확인한 뒤에만 내려준다.
    case 'getResults': {
      if (password !== store.password) return { ok: false, error: 'wrong password' };
      return {
        ok: true,
        readOnly: true,
        message: 'results',
        data: { evaluations: state.evaluations }
      };
    }

    /* 평가 결과 초기화 — 학년을 주면 그 학년 팀에 배정된 학생의 평가만,
     * 주지 않으면 전체를 지운다. 비밀번호(studentAuth)는 기본적으로
     * 남겨 두고, withPasswords 를 주면 같이 지운다. */
    case 'clearEvaluations': {
      if (password !== store.password) return { ok: false, error: 'wrong password' };
      var ceGrade = payload.grade == null ? null : String(payload.grade);
      var ceIds = Object.keys(state.evaluations);
      var ceCleared = 0;
      ceIds.forEach(function (sid) {
        if (ceGrade !== null) {
          var info = state.assignedMap[sid];
          if (!info || String(info.grade) !== ceGrade) return;
        }
        delete state.evaluations[sid];
        if (payload.withPasswords) delete state.studentAuth[sid];
        ceCleared++;
      });
      return { ok: true, message: 'evaluations cleared: ' + ceCleared };
    }

    /* 학생 한 명을 처음 상태로 되돌린다 — 제출한 점수와 비밀번호를 함께 지운다.
     *
     * 점수만 지우는 것은 의미가 적다. 학생은 자기 비밀번호로 다시 들어와
     * 언제든 점수를 고칠 수 있었으니, 교수가 초기화를 누르는 상황은 그 학생을
     * 처음부터 다시 시작하게 하려는 것이다. 비밀번호를 남겨두면 비밀번호를
     * 잊은 학생은 초기화를 해줘도 여전히 들어오지 못한다. */
    case 'clearStudentEvaluation': {
      if (password !== store.password) return { ok: false, error: 'wrong password' };
      var csId = String(payload.studentId || '');
      if (!csId) return { ok: false, error: 'bad payload' };
      var hadSubmission = state.evaluations[csId] !== undefined;
      var hadPassword = state.studentAuth[csId] !== undefined;
      if (!hadSubmission && !hadPassword) return { ok: false, error: 'nothing to clear' };
      delete state.evaluations[csId];
      delete state.studentAuth[csId];
      return { ok: true, message: 'student reset' };
    }

    // 학생이 비밀번호를 잊었을 때 — 교수가 지워 주면 다시 만들 수 있다
    case 'resetStudentPassword': {
      if (password !== store.password) return { ok: false, error: 'wrong password' };
      var rsId = String(payload.studentId || '');
      if (!rsId) return { ok: false, error: 'bad payload' };
      if (!state.studentAuth[rsId]) return { ok: false, error: 'no password' };
      delete state.studentAuth[rsId];
      return { ok: true, message: 'student password reset' };
    }

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
