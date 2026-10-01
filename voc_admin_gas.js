/**
 * VOC 대시보드 — 관리자 권한 중계 서버 (Google Apps Script)
 *
 * 역할: GitHub 토큰과 계정 정보를 서버에만 보관하고, 요청자의 역할을 확인해
 *       허용된 작업만 GitHub에 반영한다. 보조 관리자는 삭제를 요청해도 거부된다.
 *
 * ── 설치 (최초 1회) ─────────────────────────────────────────────
 * 1) script.google.com → 새 프로젝트 → 이 파일 내용 전체 붙여넣기
 * 2) 편집기에서 setup() 한 번 실행 (권한 승인) → 메인 계정이 생성됨
 *    ※ setup() 안의 MAIN_ID / MAIN_PW / GH_TOKEN 을 먼저 본인 값으로 수정
 * 3) 배포 → 새 배포 → 유형 '웹 앱'
 *    - 실행 사용자: 나
 *    - 액세스 권한: 모든 사용자
 * 4) 생성된 웹앱 URL을 VOC 대시보드 07 섹션에 입력
 *
 * ※ 코드 수정 후에는 반드시 '배포 관리 → 편집 → 새 버전'으로 다시 배포해야 반영된다.
 */

var REPO = 'HANI-4455/S.HAN';
var DATA_PATH = 'voc_data.json';

/** 최초 1회 실행 — 본인 값으로 바꾼 뒤 편집기에서 실행 */
function setup() {
  var MAIN_ID = 'shan';                 // 메인 관리자 아이디
  var MAIN_NAME = '이성한';
  var MAIN_PW = '여기에_메인_비밀번호';   // 실행 후 이 줄은 지워도 됨
  var GH_TOKEN = '여기에_github_토큰';    // github_pat_... (Contents: Read and write)

  var P = PropertiesService.getScriptProperties();
  P.setProperty('GH_TOKEN', GH_TOKEN);
  P.setProperty('USERS', JSON.stringify([mkUser(MAIN_ID, MAIN_NAME, 'main', MAIN_PW)]));
  Logger.log('설치 완료 — 메인 계정: ' + MAIN_ID);
}

/* ───────── 계정 유틸 ───────── */
function b64(bytes) { return Utilities.base64Encode(bytes); }

function hashPw(pw, saltB64) {
  var salt = Utilities.base64Decode(saltB64);
  var data = salt.concat(Utilities.newBlob(pw).getBytes());
  // 비공개 저장이므로 SHA-256 + salt 로 충분 (공개 노출 시에는 PBKDF2 필요)
  return b64(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, data));
}

function mkUser(id, name, role, pw) {
  var salt = b64(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,
    Utilities.getUuid() + new Date().getTime()));
  return { id: String(id).trim(), name: String(name || id).trim(),
           role: role === 'main' ? 'main' : 'sub', salt: salt, hash: hashPw(pw, salt) };
}

function getUsers() {
  var raw = PropertiesService.getScriptProperties().getProperty('USERS');
  return raw ? JSON.parse(raw) : [];
}

function putUsers(list) {
  PropertiesService.getScriptProperties().setProperty('USERS', JSON.stringify(list));
}

/** 자격 확인 — 실패 시 null */
function auth(id, pw) {
  var u = getUsers().filter(function (x) { return x.id === String(id || '').trim(); })[0];
  if (!u) return null;
  return hashPw(pw, u.salt) === u.hash ? u : null;
}

/* ───────── GitHub ───────── */
function ghGet(path) {
  var tok = PropertiesService.getScriptProperties().getProperty('GH_TOKEN');
  var res = UrlFetchApp.fetch('https://api.github.com/repos/' + REPO + '/contents/' + path, {
    headers: { Authorization: 'Bearer ' + tok, Accept: 'application/vnd.github+json' },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() === 404) return null;
  if (res.getResponseCode() >= 300) throw new Error('GitHub 조회 실패 ' + res.getResponseCode());
  return JSON.parse(res.getContentText());
}

function ghPut(path, contentStr, message) {
  var tok = PropertiesService.getScriptProperties().getProperty('GH_TOKEN');
  var cur = ghGet(path);
  var body = {
    message: message,
    content: Utilities.base64Encode(contentStr, Utilities.Charset.UTF_8)
  };
  if (cur && cur.sha) body.sha = cur.sha;
  var res = UrlFetchApp.fetch('https://api.github.com/repos/' + REPO + '/contents/' + path, {
    method: 'put', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + tok, Accept: 'application/vnd.github+json' },
    payload: JSON.stringify(body), muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300)
    throw new Error('GitHub 반영 실패 ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 150));
  return true;
}

function readData() {
  var f = ghGet(DATA_PATH);
  if (!f) return { updated: '', months: [] };
  return JSON.parse(Utilities.newBlob(Utilities.base64Decode(f.content)).getDataAsString('UTF-8'));
}

/* ───────── 엔드포인트 ───────── */
function doPost(e) {
  var out = { ok: false };
  try {
    var q = JSON.parse(e.postData.contents || '{}');
    var me = auth(q.id, q.pw);
    if (!me) throw new Error('아이디 또는 비밀번호가 올바르지 않습니다.');
    var isMain = me.role === 'main';
    var need = function () { if (!isMain) throw new Error('메인 관리자만 가능한 작업입니다.'); };

    switch (q.action) {
      case 'login':
        out = { ok: true, id: me.id, name: me.name, role: me.role };
        break;

      case 'publish': {                       // 메인·보조 모두 허용
        if (!q.data || !q.data.months) throw new Error('데이터가 비어 있습니다.');
        var cnt = q.data.months.length;
        ghPut(DATA_PATH, JSON.stringify(q.data, null, 1),
          'VOC 데이터 업데이트 (' + me.id + ', ' + cnt + '개월)');
        out = { ok: true, months: cnt };
        break;
      }

      case 'deleteMonth': {                   // 메인 전용 — 보조는 여기서 막힌다
        need();
        var d = readData();
        var before = d.months.length;
        d.months = d.months.filter(function (m) { return m.key !== q.key; });
        if (d.months.length === before) throw new Error('해당 월이 없습니다: ' + q.key);
        d.updated = new Date().toISOString().slice(0, 10);
        ghPut(DATA_PATH, JSON.stringify(d, null, 1), 'VOC ' + q.key + ' 삭제 (' + me.id + ')');
        out = { ok: true, removed: q.key, months: d.months.length };
        break;
      }

      case 'listUsers':
        need();
        out = { ok: true, users: getUsers().map(function (u) {
          return { id: u.id, name: u.name, role: u.role };
        }) };
        break;

      case 'saveUser': {                      // 추가 또는 수정(비번 재설정 포함)
        need();
        var t = q.target || {};
        if (!t.id) throw new Error('아이디를 입력하세요.');
        var list = getUsers();
        var i = -1;
        for (var k = 0; k < list.length; k++) if (list[k].id === t.id) i = k;
        if (i < 0) {
          if (!t.pw) throw new Error('새 계정은 초기 비밀번호가 필요합니다.');
          list.push(mkUser(t.id, t.name, t.role, t.pw));
        } else {
          if (list[i].role === 'main' && t.role === 'sub' &&
              list.filter(function (u) { return u.role === 'main'; }).length <= 1)
            throw new Error('마지막 메인 관리자는 역할을 바꿀 수 없습니다.');
          list[i].name = String(t.name || list[i].name).trim();
          list[i].role = t.role === 'main' ? 'main' : 'sub';
          if (t.pw) { var nu = mkUser(t.id, list[i].name, list[i].role, t.pw);
                      list[i].salt = nu.salt; list[i].hash = nu.hash; }
        }
        putUsers(list);
        out = { ok: true };
        break;
      }

      case 'deleteUser': {
        need();
        if (q.targetId === me.id) throw new Error('본인 계정은 삭제할 수 없습니다.');
        var ul = getUsers().filter(function (u) { return u.id !== q.targetId; });
        if (!ul.filter(function (u) { return u.role === 'main'; }).length)
          throw new Error('메인 관리자가 최소 1명은 있어야 합니다.');
        putUsers(ul);
        out = { ok: true };
        break;
      }

      case 'changePw': {                      // 본인 비밀번호 변경
        if (!q.newPw || String(q.newPw).length < 8) throw new Error('새 비밀번호는 8자 이상이어야 합니다.');
        var l2 = getUsers();
        for (var j = 0; j < l2.length; j++) if (l2[j].id === me.id) {
          var n2 = mkUser(me.id, l2[j].name, l2[j].role, q.newPw);
          l2[j].salt = n2.salt; l2[j].hash = n2.hash;
        }
        putUsers(l2);
        out = { ok: true };
        break;
      }

      default:
        throw new Error('알 수 없는 요청: ' + q.action);
    }
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, service: 'VOC admin relay' }))
    .setMimeType(ContentService.MimeType.JSON);
}
