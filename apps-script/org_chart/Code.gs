/**
 * PAS사업부 조직도 관리 — 구글시트 연동판 (서버)
 * 보기·과거 조회·인쇄: 웹 앱 주소를 아는 누구나 (배포: 실행 계정 = 나, 액세스 = 모든 사용자)
 * 명단 수정·기준일 확정: 관리자 로그인. 계정은 스크립트 속성에 비밀번호 해시로만 저장한다
 * 시트 탭: 데이터(현재 명단) / 조직도_이력(확정 명단 누적) / 조직도_기준일정보 / 조직도_설정
 * @OnlyCurrentDoc
 */
const TAB = {data: '데이터', hist: '조직도_이력', info: '조직도_기준일정보', set: '조직도_설정'};
const DRAFT = '작성중';
const SET_KEYS = {
  orgName: '조직명', topUnit: '최상위 단위 팀명', coachRoles: '별도 집계 직무', maxRows: '팀당 최대 줄',
  showGroups: '직무별 소제목', owner: '관리 담당자', draftDate: '작성중 기준일'
};
const SET_DEFAULTS = {orgName: 'PAS사업부', topUnit: '사업부', coachRoles: '코칭매니저', maxRows: '15', showGroups: 'TRUE', owner: '', draftDate: ''};
const USERS_KEY = 'ADMIN_USERS', TOKEN_SECONDS = 6 * 3600, MAX_FAILS = 5, LOCK_SECONDS = 600;

/* ---------- 시트 메뉴 ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('조직도 관리')
    .addItem('조직도 화면 열기', 'openApp')
    .addItem('초기 설정 (처음 한 번)', 'setup')
    .addItem('admin 비밀번호 초기화', 'resetAdminPassword')
    .addToUi();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('조직도 관리 · 파리크라상')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function openApp() {
  const ui = SpreadsheetApp.getUi(), url = ScriptApp.getService().getUrl();
  if (!url) return ui.alert('웹 앱이 아직 배포되지 않았습니다. 설치 안내의 "웹 앱 배포" 단계를 먼저 해 주세요.');
  const html = HtmlService.createHtmlOutput(
    '<p style="font-family:sans-serif;font-size:14px">아래 링크를 누르면 새 탭에서 조직도 화면이 열립니다.</p>' +
    '<p><a href="' + url + '" target="_blank" style="font-family:sans-serif;font-size:16px">조직도 화면 열기</a></p>'
  ).setWidth(360).setHeight(130);
  ui.showModalDialog(html, '조직도 관리');
}

/** 필요한 탭을 만들고, 관리자 계정이 하나도 없으면 admin 계정을 임시 비밀번호로 만든다. 여러 번 실행해도 안전 */
function setup() {
  const ss = ss_();
  sheet_(ss, TAB.data);
  prepare_(ss, TAB.hist, ['기준일', '본부/부서', '팀명', '직책', '직무/역할', '이름', '재직 상태', '특이사항/비고']);
  prepare_(ss, TAB.info, ['기준일', '점포 인원', '당월 퇴사 및 부서이동', '휴·복직 현황', '팀 표시 순서', '저장 시각', '저장한 관리자']);
  const set = prepare_(ss, TAB.set, ['항목', '값']);
  set.getRange('B:B').setNumberFormat('@');
  const have = rows_(set, 2).map(r => r[0]);
  Object.keys(SET_KEYS).forEach(k => { if (have.indexOf(SET_KEYS[k]) < 0) set.appendRow([SET_KEYS[k], SET_DEFAULTS[k]]); });
  let temp = '', msg = '초기 설정이 끝났습니다. 조직도_이력, 조직도_기준일정보, 조직도_설정 탭을 확인하세요.';
  if (!Object.keys(users_()).length) {
    temp = tempPassword_();
    setUser_('admin', '관리자', temp);
    msg += '\n\n첫 관리자 계정을 만들었습니다.\n아이디: admin\n임시 비밀번호: ' + temp + '\n\n조직도 화면에서 로그인한 뒤 설정 탭에서 비밀번호를 바꾸세요.';
  }
  notify_(msg);
  return {ok: true, tempPassword: temp};
}

/** 비밀번호를 잊었을 때: 시트 편집자가 메뉴에서 실행하면 admin 계정 비밀번호를 새 임시 비밀번호로 바꾼다 */
function resetAdminPassword() {
  const temp = tempPassword_(), u = users_();
  setUser_('admin', (u.admin && u.admin.name) || '관리자', temp);
  notify_('admin 계정 비밀번호를 초기화했습니다.\n임시 비밀번호: ' + temp + '\n\n로그인한 뒤 설정 탭에서 바꾸세요.');
  return temp;
}

/* ---------- 누구나 부르는 함수 ---------- */
function api_load() { return load_(ss_()); }

function api_version() { return hash_(read_(ss_())); }

function api_login(id, pw) {
  id = String(id || '').trim();
  const cache = CacheService.getScriptCache(), failKey = 'fail:' + id, fails = +(cache.get(failKey) || 0);
  if (fails >= MAX_FAILS) return {ok: false, msg: '로그인에 여러 번 실패해 10분 동안 잠겼습니다.'};
  const u = users_()[id];
  if (!u || hashPassword_(u.salt, String(pw || '')) !== u.hash) {
    cache.put(failKey, String(fails + 1), LOCK_SECONDS);
    return {ok: false, msg: '아이디 또는 비밀번호가 맞지 않습니다.'};
  }
  cache.remove(failKey);
  const token = Utilities.getUuid();
  cache.put('tok:' + token, id, TOKEN_SECONDS);
  return {ok: true, token: token, id: id, name: u.name || id};
}

function api_logout(token) {
  if (token) CacheService.getScriptCache().remove('tok:' + token);
  return true;
}

/* ---------- 관리자만 부르는 함수 (첫 인자는 로그인 토큰) ---------- */
function api_saveDraft(token, draft, baseVersion, force) {
  const me = who_(token);
  return withLock_(ss => {
    if (!force && hash_(read_(ss)) !== baseVersion) return {conflict: true};
    writeDraft_(ss, draft, me);
    return {version: hash_(read_(ss))};
  });
}

function api_confirm(token, date, draft, baseVersion, force) {
  const me = who_(token);
  date = normDate_(date);
  if (!date) throw new Error('기준일 형식이 올바르지 않습니다.');
  return withLock_(ss => {
    if (!force && hash_(read_(ss)) !== baseVersion) return {conflict: true};
    draft.settings.draftDate = date;
    writeDraft_(ss, draft, me);
    const org = draft.settings.orgName;
    const add = draft.people.map(p => [date, org, p.team, p.position, p.role, p.name, p.status, p.note]);
    replaceRows_(sheet_(ss, TAB.hist), 8, r => normDate_(r[0]) !== date, add, true);
    upsertInfo_(ss, date, draft, me);
    return load_(ss);
  });
}

function api_deleteDate(token, date, baseVersion) {
  who_(token);
  date = normDate_(date);
  return withLock_(ss => {
    if (hash_(read_(ss)) !== baseVersion) return {conflict: true};
    replaceRows_(sheet_(ss, TAB.hist), 8, r => normDate_(r[0]) !== date, [], false);
    replaceRows_(sheet_(ss, TAB.info), 7, r => r[0] === DRAFT || normDate_(r[0]) !== date, [], false);
    return load_(ss);
  });
}

function api_listUsers(token) {
  const me = who_(token), u = users_();
  return Object.keys(u).sort().map(k => ({id: k, name: u[k].name, self: k === me}));
}

function api_addUser(token, id, name, pw) {
  who_(token);
  id = String(id || '').trim();
  if (!/^[A-Za-z0-9._-]{3,30}$/.test(id)) throw new Error('아이디는 영문·숫자 3~30자로 입력하세요.');
  if (users_()[id]) throw new Error('이미 있는 아이디입니다.');
  checkPassword_(pw);
  setUser_(id, String(name || '').trim() || id, pw);
  return api_listUsers(token);
}

function api_removeUser(token, id) {
  const me = who_(token);
  if (id === me) throw new Error('지금 로그인한 계정은 지울 수 없습니다.');
  const u = users_();
  delete u[id];
  saveUsers_(u);
  return api_listUsers(token);
}

function api_changePassword(token, oldPw, newPw) {
  const me = who_(token), u = users_()[me];
  if (hashPassword_(u.salt, String(oldPw || '')) !== u.hash) throw new Error('현재 비밀번호가 맞지 않습니다.');
  checkPassword_(newPw);
  setUser_(me, u.name, newPw);
  return true;
}

/* ---------- 관리자 계정 ---------- */
function users_() { return JSON.parse(PropertiesService.getScriptProperties().getProperty(USERS_KEY) || '{}'); }

function saveUsers_(u) { PropertiesService.getScriptProperties().setProperty(USERS_KEY, JSON.stringify(u)); }

function hashPassword_(salt, pw) {
  let h = salt + '|' + pw;
  for (let i = 0; i < 300; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8));
  }
  return h;
}

function setUser_(id, name, pw) {
  const u = users_(), salt = Utilities.getUuid();
  u[id] = {name: name || id, salt: salt, hash: hashPassword_(salt, String(pw))};
  saveUsers_(u);
}

function checkPassword_(pw) {
  if (String(pw || '').length < 8) throw new Error('비밀번호는 8자 이상이어야 합니다.');
}

function who_(token) {
  const id = token ? CacheService.getScriptCache().get('tok:' + token) : null;
  if (!id || !users_()[id]) throw new Error('관리자 로그인이 필요합니다. 다시 로그인해 주세요.');
  return id;
}

function tempPassword_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

function notify_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

/* ---------- 시트 읽고 쓰기 ---------- */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error("'" + name + "' 탭이 없습니다. 시트 메뉴 [조직도 관리 > 초기 설정]을 먼저 실행하세요.");
  return sh;
}

function prepare_(ss, name, head) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('@');
  return sh;
}

function rows_(sh, ncol) {
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, ncol).getDisplayValues().map(r => r.map(v => String(v).trim())) : [];
}

function normDate_(v) {
  const m = String(v || '').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : '';
}

function person_(r) { return {team: r[1], position: r[2], role: r[3], name: r[4], status: r[5], note: r[6]}; }

function parseStore_(s) {
  return String(s || '').split('\n').map(l => l.trim()).filter(Boolean).map(l => {
    const i = l.lastIndexOf('=');
    return i < 0 ? {label: l, n: ''} : {label: l.slice(0, i).trim(), n: l.slice(i + 1).trim()};
  });
}

function read_(ss) {
  const raw = {};
  rows_(sheet_(ss, TAB.set), 2).forEach(r => { if (r[0]) raw[r[0]] = r[1]; });
  const settings = {};
  Object.keys(SET_KEYS).forEach(k => { settings[k] = raw[SET_KEYS[k]] !== undefined ? raw[SET_KEYS[k]] : SET_DEFAULTS[k]; });
  const current = rows_(sheet_(ss, TAB.data), 7).filter(r => r[4]).map(person_);
  const hist = {};
  rows_(sheet_(ss, TAB.hist), 8).forEach(r => {
    const d = normDate_(r[0]);
    if (d && r[5]) (hist[d] = hist[d] || []).push(person_(r.slice(1)));
  });
  const infos = {};
  rows_(sheet_(ss, TAB.info), 5).forEach(r => {
    const k = r[0] === DRAFT ? DRAFT : normDate_(r[0]);
    if (k) infos[k] = {store: parseStore_(r[1]), notes: {moves: r[2], leaves: r[3]}, teamOrder: r[4].split('\n').map(x => x.trim()).filter(Boolean)};
  });
  return {settings: settings, current: current, hist: hist, infos: infos};
}

function hash_(a) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(a), Utilities.Charset.UTF_8);
  return Utilities.base64Encode(bytes);
}

function load_(ss) {
  const a = read_(ss);
  a.version = hash_(read_(ss));
  a.sheetUrl = ss.getUrl();
  return a;
}

/** 수식으로 해석되지 않게 막는다 (= + @ 로 시작하는 글자) */
function safe_(v) {
  const s = v === null || v === undefined ? '' : String(v);
  return /^[=+@]/.test(s) ? "'" + s : s;
}

function writeDraft_(ss, d, me) {
  const sh = sheet_(ss, TAB.data), old = sh.getLastRow() - 1;
  if (old > 0) sh.getRange(2, 1, old, 7).clearContent();
  const rows = d.people.map(p => [d.settings.orgName, p.team, p.position, p.role, p.name, p.status, p.note].map(safe_));
  if (rows.length) sh.getRange(2, 1, rows.length, 7).setValues(rows);
  upsertInfo_(ss, DRAFT, d, me);
  const set = sheet_(ss, TAB.set), cur = rows_(set, 2), labels = cur.map(r => r[0]);
  Object.keys(SET_KEYS).forEach(k => {
    const v = safe_(d.settings[k]), i = labels.indexOf(SET_KEYS[k]);
    if (i < 0) { set.appendRow([SET_KEYS[k], v]); labels.push(SET_KEYS[k]); }
    else if (cur[i][1] !== String(d.settings[k] === undefined ? '' : d.settings[k])) set.getRange(i + 2, 2).setValue(v);
  });
}

function upsertInfo_(ss, key, d, me) {
  const sh = sheet_(ss, TAB.info);
  const row = [key, d.store.map(x => x.label + '=' + x.n).join('\n'), d.notes.moves, d.notes.leaves, d.teamOrder.join('\n'),
    Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm'), me || ''].map(safe_);
  const keys = rows_(sh, 1).map(r => r[0] === DRAFT ? DRAFT : normDate_(r[0]));
  const at = keys.indexOf(key) >= 0 ? keys.indexOf(key) + 2 : sh.getLastRow() + 1;
  sh.getRange(at, 1).setNumberFormat('@');
  sh.getRange(at, 1, 1, 7).setValues([row]);
}

function replaceRows_(sh, ncol, keep, add, sortByDate) {
  const cur = rows_(sh, ncol);
  let next = cur.filter(keep).concat(add.map(r => r.map(safe_)));
  if (sortByDate) next = next.sort((a, b) => normDate_(a[0]).localeCompare(normDate_(b[0])));
  if (cur.length) sh.getRange(2, 1, cur.length, ncol).clearContent();
  if (next.length) {
    sh.getRange(2, 1, next.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, next.length, ncol).setValues(next);
  }
}

function withLock_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try { return fn(ss_()); } finally { lock.releaseLock(); }
}
