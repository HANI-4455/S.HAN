/**
 * FOOD 제품 생산 Check-list — Google Apps Script 백엔드
 * ------------------------------------------------------
 * 사용법
 *  1) 구글 시트를 새로 만든다 (이름 예: FOOD 체크리스트 DB)
 *  2) 확장 프로그램 > Apps Script 를 열고 이 파일 내용을 전부 붙여넣는다
 *  3) 상단 함수 목록에서 setup 을 선택하고 실행 → 시트/기본 계정/기본 제품이 생성된다
 *  4) 배포 > 새 배포 > 유형: 웹 앱
 *       - 실행 사용자: 나
 *       - 액세스 권한: 모든 사용자
 *     배포 후 나오는 웹 앱 URL 을 food_checklist.html 첫 화면에 입력한다
 *  5) 코드를 수정했으면 반드시 [배포 > 배포 관리 > 편집 > 버전: 새 버전] 으로 재배포해야 반영된다
 *
 * 기본 계정 (setup 실행 시 생성 · 최초 로그인 후 비밀번호 변경 필요)
 *    admin / 1234   전체 관리자
 *    area1 / 1234   사업1팀 관리자
 *    area2 / 1234   사업2팀 관리자
 *    area3 / 1234   사업3팀 관리자
 *
 * 이미 쓰던 시트에 이 코드를 덮어쓴 경우에도 setup 을 다시 실행하면 된다.
 * 열이 늘어난 만큼 머리글이 채워지고, 기존 데이터는 그대로 남는다.
 */

var TZ = 'Asia/Seoul';
var SESSION_HOURS = 12;          // 로그인 유지 시간
var RECORD_FETCH_LIMIT = 3000;   // 한 번에 내려주는 최대 이력 건수

/* ===================== 시트 정의 ===================== */

/* Regions 시트는 '팀'을 담는다 (사업1팀 · 사업2팀 · 사업3팀).
   areas 열에 그 팀이 담당하는 광역시도를 JSON 배열로 넣는다.
   시트 이름과 regionId 열 이름은 기존 데이터를 살리려고 그대로 둔다. */
/* ⚠ 열 순서를 절대 중간에 바꾸지 말 것.
   readAll 은 시트의 열 위치로 값을 읽는데, setup 은 머리글만 다시 쓰고
   기존 데이터는 옮기지 않는다. 중간에 열을 끼우면 그 뒤 값이 전부 한 칸씩
   밀려서 읽힌다. 새 항목은 반드시 맨 뒤에 붙인다. */
var SHEETS = {
  Regions: ['regionId', 'regionName', 'sortOrder', 'areas'],
  Users: ['userId', 'pwHash', 'salt', 'role', 'name', 'regionId', 'storeCode', 'phone', 'active', 'mustChangePw', 'createdAt', 'createdBy', 'area'],
  Products: ['productId', 'name', 'methods', 'qualityHours', 'qualityLabel', 'note', 'sortOrder', 'active'],
  Records: ['recordId', 'regionId', 'regionName', 'storeId', 'storeName', 'productName', 'qty',
            'mfgMs', 'mfgText', 'thawMs', 'thawText', 'thawMaxMs', 'expireMs', 'expireText',
            'storage', 'status', 'memo', 'createdAt', 'updatedAt', 'updatedBy', 'area'],
  Sessions: ['token', 'userId', 'expireMs']
};

/* 전국 광역자치단체 17곳 — 점포 소재 지역 */
var AREAS = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
             '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주'];

/* 기본 제품 기준값 — 첨부된 체크리스트 양식 그대로 */
var SEED_PRODUCTS = [
  { name: '조각케익',                        methods: [{ storage: '냉장', min: 2,  max: 3 }],  qualityHours: 48, qualityLabel: '48시간' },
  { name: '홀케익',                          methods: [{ storage: '냉장', min: 4,  max: 11 }], qualityHours: 48, qualityLabel: '48시간' },
  { name: '파니니류',                        methods: [{ storage: '냉장', min: 12, max: 12 }], qualityHours: 0,  qualityLabel: '당일' },
  { name: '브라우니',                        methods: [{ storage: '냉장', min: 2,  max: 2 }],  qualityHours: 48, qualityLabel: '48시간' },
  { name: '마카롱',                          methods: [{ storage: '냉장', min: 2,  max: 2 }],  qualityHours: 72, qualityLabel: '72시간' },
  { name: '후레쉬 카스텔라',                 methods: [{ storage: '냉장', min: 2,  max: 3 }],  qualityHours: 48, qualityLabel: '2일' },
  { name: '타르트(에그/치즈)',               methods: [{ storage: '냉장', min: 2,  max: 2 }],  qualityHours: 24, qualityLabel: '24시간' },
  { name: '베이글 / 6쪽식빵',                methods: [{ storage: '냉장', min: 12, max: 12 }], qualityHours: 72, qualityLabel: '72시간' },
  { name: '햄토스트',                        methods: [{ storage: '냉장', min: 12, max: 12 }], qualityHours: 0,  qualityLabel: '당일' },
  { name: '크라상 / 푀이테쇼콜라 / 뺑오레젱', methods: [{ storage: '상온', min: 2,  max: 2 }],  qualityHours: 0,  qualityLabel: '당일' },
  { name: '마카다미아 데니쉬',               methods: [{ storage: '상온', min: 2,  max: 2 }],  qualityHours: 15, qualityLabel: '15시간' },
  { name: '퀸아망 / 몽블랑',                 methods: [{ storage: '상온', min: 1,  max: 1 }],  qualityHours: 15, qualityLabel: '15시간' },
  { name: '스콘',                            methods: [{ storage: '상온', min: 1,  max: 1 }],  qualityHours: 36, qualityLabel: '36시간' },
  { name: '파운드케익',                      methods: [{ storage: '상온', min: 1,  max: 1 }],  qualityHours: 48, qualityLabel: '2일' },
  { name: '판 포카챠',                       methods: [{ storage: '냉장', min: 12, max: 12 }, { storage: '상온', min: 4, max: 4 }], qualityHours: 0, qualityLabel: '당일' },
  { name: '제조 샌드위치',                   methods: [{ storage: '냉장', min: 0,  max: 0 }],  qualityHours: 0,  qualityLabel: '당일', note: '당일생산 당일폐기' }
];

var SEED_REGIONS = [
  { regionId: 'R1', regionName: '사업1팀', sortOrder: 1, areas: '[]' },
  { regionId: 'R2', regionName: '사업2팀', sortOrder: 2, areas: '[]' },
  { regionId: 'R3', regionName: '사업3팀', sortOrder: 3, areas: '[]' }
];

/* 처음 만들 때 넣었던 임시 팀명 — 아직 그대로면 사업N팀으로 바꿔 준다 */
var OLD_DEFAULT_TEAM_NAMES = ['지역 1', '지역 2', '지역 3'];

/* ===================== 최초 설치 ===================== */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(SHEETS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    var head = SHEETS[name];
    sh.getRange(1, 1, 1, head.length).setValues([head])
      .setFontWeight('bold').setBackground('#2F497D').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  });

  var regions = readAll('Regions');
  if (regions.length === 0) {
    SEED_REGIONS.forEach(function (r) { insertRow('Regions', r); });
  } else {
    // 이미 쓰던 시트라면 — 임시 팀명만 바꾸고, areas 가 비어 있으면 빈 배열로 채운다
    regions.forEach(function (r) {
      var patch = {};
      var i = OLD_DEFAULT_TEAM_NAMES.indexOf(String(r.regionName).trim());
      if (i >= 0) patch.regionName = '사업' + (i + 1) + '팀';
      if (!String(r.areas || '').length) patch.areas = '[]';
      if (Object.keys(patch).length) updateRow('Regions', r._row, patch);
    });
  }

  if (readAll('Products').length === 0) {
    SEED_PRODUCTS.forEach(function (p, i) {
      insertRow('Products', {
        productId: 'P' + pad(i + 1, 3),
        name: p.name,
        methods: JSON.stringify(p.methods),
        qualityHours: p.qualityHours,
        qualityLabel: p.qualityLabel,
        note: p.note || '',
        sortOrder: i + 1,
        active: 'Y'
      });
    });
  }

  if (readAll('Users').length === 0) {
    createUserRow('admin', '1234', 'super',  '전체 관리자',    '',   '', '', 'system');
    createUserRow('area1', '1234', 'region', '사업1팀 관리자', 'R1', '', '', 'system');
    createUserRow('area2', '1234', 'region', '사업2팀 관리자', 'R2', '', '', 'system');
    createUserRow('area3', '1234', 'region', '사업3팀 관리자', 'R3', '', '', 'system');
  }

  var sh = ss.getSheetByName('Sessions');
  if (sh) sh.hideSheet();

  return '설치 완료 — admin/1234 로 로그인하세요.';
}

/**
 * 비밀번호 초기화 — 관리자가 비밀번호를 잊었을 때 쓴다.
 *
 * 아래 두 값을 고친 뒤 함수 목록에서 resetPassword 를 선택해 실행한다.
 * 편집기에서 도는 것이라 웹 앱 재배포는 필요 없다.
 * 실행하면 그 계정은 '초기비번' 상태가 되므로, 로그인 후 바로 바꾸게 한다.
 */
function resetPassword() {
  var USER_ID = 'admin';     // 초기화할 아이디
  var NEW_PW = '1234';       // 새 비밀번호

  var u = findUser(USER_ID);
  if (!u) throw new Error(USER_ID + ' 계정을 찾을 수 없습니다.');

  var salt = newToken().slice(0, 12);
  updateRow('Users', u._row, {
    salt: salt,
    pwHash: hash(String(NEW_PW), salt),
    mustChangePw: 'Y',
    active: 'Y'              // 실수로 중지된 계정도 함께 되살린다
  });
  return USER_ID + ' 비밀번호를 ' + NEW_PW + ' 로 초기화했습니다. 로그인 후 바로 변경하세요.';
}

/* ===================== HTTP 진입점 ===================== */

function doGet(e) {
  return json({ ok: true, service: 'FOOD 체크리스트 API', time: new Date().toISOString() });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json({ ok: false, error: '서버가 혼잡합니다. 잠시 후 다시 시도해 주세요.' });
  }
  try {
    var req = JSON.parse(e.postData.contents || '{}');
    var fn = ACTIONS[req.action];
    if (!fn) return json({ ok: false, error: '알 수 없는 요청: ' + req.action });
    var out = fn(req) || {};
    out.ok = true;
    return json(out);
  } catch (err) {
    return json({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ===================== 액션 ===================== */

var ACTIONS = {

  /* ---- 인증 ---- */

  login: function (req) {
    var u = findUser(req.userId);
    if (!u || u.active !== 'Y') throw new Error('아이디 또는 비밀번호가 올바르지 않습니다.');
    if (hash(String(req.pw), u.salt) !== u.pwHash) throw new Error('아이디 또는 비밀번호가 올바르지 않습니다.');
    var token = newToken();
    insertRow('Sessions', { token: token, userId: u.userId, expireMs: Date.now() + SESSION_HOURS * 3600000 });
    cleanSessions();
    return { token: token, user: publicUser(u) };
  },

  me: function (req) {
    var u = auth(req);
    return { user: publicUser(u) };
  },

  logout: function (req) {
    var rows = readAll('Sessions');
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].token === req.token) { deleteRowAt('Sessions', rows[i]._row); break; }
    }
    return {};
  },

  changePw: function (req) {
    var u = auth(req);
    if (hash(String(req.oldPw), u.salt) !== u.pwHash) throw new Error('현재 비밀번호가 올바르지 않습니다.');
    if (!req.newPw || String(req.newPw).length < 4) throw new Error('새 비밀번호는 4자 이상이어야 합니다.');
    var salt = newToken().slice(0, 12);
    updateRow('Users', u._row, { salt: salt, pwHash: hash(String(req.newPw), salt), mustChangePw: 'N' });
    return {};
  },

  /* ---- 기준 정보 ---- */

  bootstrap: function (req) {
    var u = auth(req);
    var out = {
      user: publicUser(u),
      areas: AREAS,
      regions: readAll('Regions').map(function (r) {
        return {
          regionId: r.regionId, regionName: r.regionName,
          sortOrder: Number(r.sortOrder) || 0,
          areas: parseJson(r.areas, [])
        };
      }).sort(function (a, b) { return a.sortOrder - b.sortOrder; }),
      products: readAll('Products')
        .filter(function (p) { return p.active !== 'N'; })
        .map(function (p) {
          return {
            productId: p.productId, name: p.name,
            methods: parseJson(p.methods, []),
            qualityHours: Number(p.qualityHours) || 0,
            qualityLabel: p.qualityLabel, note: p.note,
            sortOrder: Number(p.sortOrder) || 0
          };
        })
        .sort(function (a, b) { return a.sortOrder - b.sortOrder; })
    };
    if (u.role !== 'store') out.users = visibleUsers(u).map(publicUser);
    return out;
  },

  /* ---- 지역 (전체 관리자 전용) ---- */

  saveRegion: function (req) {
    auth(req, ['super']);
    var r = req.region || {};
    if (!r.regionName) throw new Error('팀 이름을 입력해 주세요.');

    // 담당 지역은 광역시도 목록 안에서만 고를 수 있다
    var areas = (r.areas || []).filter(function (a) { return AREAS.indexOf(a) >= 0; });

    // 한 지역을 두 팀이 함께 담당하면 집계가 겹치므로 막는다
    var rows = readAll('Regions');
    var dup = [];
    rows.forEach(function (x) {
      if (String(x.regionId) === String(r.regionId)) return;
      parseJson(x.areas, []).forEach(function (a) {
        if (areas.indexOf(a) >= 0) dup.push(a + '(' + x.regionName + ')');
      });
    });
    if (dup.length) throw new Error('이미 다른 팀이 담당하는 지역입니다 — ' + dup.join(', '));

    var hit = find(rows, 'regionId', r.regionId);
    if (hit) {
      updateRow('Regions', hit._row, {
        regionName: r.regionName,
        sortOrder: r.sortOrder || hit.sortOrder,
        areas: JSON.stringify(areas)
      });
    } else {
      insertRow('Regions', {
        regionId: r.regionId || 'R' + (rows.length + 1),
        regionName: r.regionName,
        sortOrder: r.sortOrder || rows.length + 1,
        areas: JSON.stringify(areas)
      });
    }
    return { regions: teamList() };
  },

  deleteRegion: function (req) {
    auth(req, ['super']);
    var users = readAll('Users').filter(function (x) { return x.regionId === req.regionId && x.active === 'Y'; });
    if (users.length) throw new Error('이 팀에 소속된 계정이 ' + users.length + '개 있습니다. 먼저 정리해 주세요.');
    var hit = find(readAll('Regions'), 'regionId', req.regionId);
    if (hit) deleteRowAt('Regions', hit._row);
    return { regions: teamList() };
  },

  /* ---- 계정 ---- */

  saveUser: function (req) {
    var me = auth(req, ['super', 'region']);
    var d = req.user || {};
    if (!d.userId) throw new Error('아이디를 입력해 주세요.');
    if (!d.name) throw new Error('이름(점포명)을 입력해 주세요.');

    // 팀 관리자는 자기 팀의 점포 계정만 다룰 수 있다
    if (me.role === 'region') {
      if (d.role !== 'store') throw new Error('팀 관리자는 점포 계정만 만들 수 있습니다.');
      d.regionId = me.regionId;
    }
    if (d.role === 'store') {
      if (!d.regionId) throw new Error('팀을 선택해 주세요.');
      if (!d.area) throw new Error('지역을 선택해 주세요.');
      if (AREAS.indexOf(d.area) < 0) throw new Error('지역 값이 올바르지 않습니다 — ' + d.area);

      // 팀에 담당 지역을 지정해 두었으면 그 안에서만 고를 수 있다
      var team = find(readAll('Regions'), 'regionId', d.regionId);
      var teamAreas = team ? parseJson(team.areas, []) : [];
      if (teamAreas.length && teamAreas.indexOf(d.area) < 0) {
        throw new Error(d.area + '은(는) ' + (team ? team.regionName : '이 팀') + ' 담당 지역이 아닙니다.');
      }
    }

    var hit = findUser(d.userId);
    if (hit) {
      if (me.role === 'region' && (hit.role !== 'store' || hit.regionId !== me.regionId)) {
        throw new Error('수정 권한이 없는 계정입니다.');
      }
      var patch = {
        name: d.name, role: d.role, regionId: d.regionId || '', area: d.area || '',
        storeCode: d.storeCode || '', phone: d.phone || '',
        active: d.active === false ? 'N' : 'Y'
      };
      if (d.pw) {
        var salt = newToken().slice(0, 12);
        patch.salt = salt;
        patch.pwHash = hash(String(d.pw), salt);
        patch.mustChangePw = 'Y';
      }
      updateRow('Users', hit._row, patch);
    } else {
      if (!d.pw) throw new Error('비밀번호를 입력해 주세요.');
      createUserRow(d.userId, d.pw, d.role, d.name, d.regionId || '', d.area || '',
                    d.storeCode || '', me.userId, d.phone || '');
    }
    return { users: visibleUsers(me).map(publicUser) };
  },

  deleteUser: function (req) {
    var me = auth(req, ['super', 'region']);
    var hit = findUser(req.userId);
    if (!hit) throw new Error('계정을 찾을 수 없습니다.');
    if (hit.userId === me.userId) throw new Error('본인 계정은 삭제할 수 없습니다.');
    if (me.role === 'region' && (hit.role !== 'store' || hit.regionId !== me.regionId)) {
      throw new Error('삭제 권한이 없는 계정입니다.');
    }
    deleteRowAt('Users', hit._row);
    return { users: visibleUsers(me).map(publicUser) };
  },

  /* ---- 제품 기준값 (전체 관리자 전용) ---- */

  saveProduct: function (req) {
    auth(req, ['super']);
    var p = req.product || {};
    if (!p.name) throw new Error('제품명을 입력해 주세요.');
    if (!p.methods || !p.methods.length) throw new Error('해동(보관) 조건을 1개 이상 입력해 주세요.');
    var rows = readAll('Products');
    var hit = find(rows, 'productId', p.productId);
    var patch = {
      name: p.name,
      methods: JSON.stringify(p.methods),
      qualityHours: Number(p.qualityHours) || 0,
      qualityLabel: p.qualityLabel || (Number(p.qualityHours) ? p.qualityHours + '시간' : '당일'),
      note: p.note || '',
      sortOrder: Number(p.sortOrder) || rows.length + 1,
      active: p.active === false ? 'N' : 'Y'
    };
    if (hit) updateRow('Products', hit._row, patch);
    else {
      patch.productId = 'P' + pad(rows.length + 1, 3);
      insertRow('Products', patch);
    }
    return {};
  },

  deleteProduct: function (req) {
    auth(req, ['super']);
    var hit = find(readAll('Products'), 'productId', req.productId);
    if (hit) deleteRowAt('Products', hit._row);
    return {};
  },

  /* ---- 사용 이력 ---- */

  saveRecord: function (req) {
    var u = auth(req);
    var d = req.record || {};
    if (!d.productName) throw new Error('제품을 선택해 주세요.');
    if (!d.mfgMs) throw new Error('제조일시를 입력해 주세요.');

    var storeId, storeName, regionId, area;
    if (u.role === 'store') {
      storeId = u.userId; storeName = u.name; regionId = u.regionId; area = u.area;
    } else {
      storeId = d.storeId || u.userId;
      var su = findUser(storeId);
      storeName = su ? su.name : (d.storeName || u.name);
      regionId = su ? su.regionId : (d.regionId || u.regionId);
      area = su ? su.area : '';
    }
    var region = find(readAll('Regions'), 'regionId', regionId);

    var row = {
      regionId: regionId || '',
      regionName: region ? region.regionName : '',
      area: area || '',
      storeId: storeId,
      storeName: storeName,
      productName: d.productName,
      qty: d.qty || '',
      mfgMs: Number(d.mfgMs),
      mfgText: fmt(d.mfgMs),
      thawMs: Number(d.thawMs),
      thawText: fmt(d.thawMs),
      thawMaxMs: Number(d.thawMaxMs || d.thawMs),
      expireMs: Number(d.expireMs),
      expireText: fmt(d.expireMs),
      storage: d.storage || '',
      status: d.status || '진행중',
      memo: d.memo || '',
      updatedAt: fmt(Date.now()),
      updatedBy: u.userId
    };

    var hit = d.recordId ? find(readAll('Records'), 'recordId', d.recordId) : null;
    if (hit) {
      if (u.role === 'store' && hit.storeId !== u.userId) throw new Error('수정 권한이 없습니다.');
      updateRow('Records', hit._row, row);
    } else {
      row.recordId = 'REC' + Date.now() + Math.floor(Math.random() * 1000);
      row.createdAt = fmt(Date.now());
      insertRow('Records', row);
    }
    return {};
  },

  setRecordStatus: function (req) {
    var u = auth(req);
    var hit = find(readAll('Records'), 'recordId', req.recordId);
    if (!hit) throw new Error('이력을 찾을 수 없습니다.');
    if (u.role === 'store' && hit.storeId !== u.userId) throw new Error('수정 권한이 없습니다.');
    updateRow('Records', hit._row, { status: req.status, updatedAt: fmt(Date.now()), updatedBy: u.userId });
    return {};
  },

  deleteRecord: function (req) {
    var u = auth(req);
    var hit = find(readAll('Records'), 'recordId', req.recordId);
    if (!hit) throw new Error('이력을 찾을 수 없습니다.');
    if (u.role === 'store' && hit.storeId !== u.userId) throw new Error('삭제 권한이 없습니다.');
    if (u.role === 'region' && hit.regionId !== u.regionId) throw new Error('삭제 권한이 없습니다.');
    deleteRowAt('Records', hit._row);
    return {};
  },

  listRecords: function (req) {
    var u = auth(req);
    var rows = readAll('Records');
    var fromMs = req.fromMs ? Number(req.fromMs) : 0;
    var toMs = req.toMs ? Number(req.toMs) : Infinity;

    var out = rows.filter(function (r) {
      if (u.role === 'store' && r.storeId !== u.userId) return false;
      if (u.role === 'region' && r.regionId !== u.regionId) return false;
      if (req.storeId && r.storeId !== req.storeId) return false;
      if (req.regionId && r.regionId !== req.regionId) return false;
      var m = Number(r.mfgMs) || 0;
      return m >= fromMs && m <= toMs;
    }).map(strip);

    out.sort(function (a, b) { return Number(b.mfgMs) - Number(a.mfgMs); });
    var total = out.length;
    if (out.length > RECORD_FETCH_LIMIT) out = out.slice(0, RECORD_FETCH_LIMIT);
    return { records: out, total: total, truncated: total > out.length };
  }
};

/* ===================== 권한 ===================== */

function auth(req, roles) {
  if (!req.token) throw new Error('로그인이 필요합니다.');
  var s = find(readAll('Sessions'), 'token', req.token);
  if (!s || Number(s.expireMs) < Date.now()) throw new Error('로그인이 만료되었습니다. 다시 로그인해 주세요.');
  var u = findUser(s.userId);
  if (!u || u.active !== 'Y') throw new Error('사용할 수 없는 계정입니다.');
  if (roles && roles.indexOf(u.role) < 0) throw new Error('권한이 없습니다.');
  return u;
}

function visibleUsers(me) {
  var all = readAll('Users');
  if (me.role === 'super') return all;
  if (me.role === 'region') {
    return all.filter(function (x) { return x.regionId === me.regionId; });
  }
  return [];
}

function publicUser(u) {
  return {
    userId: u.userId, role: u.role, name: u.name, regionId: u.regionId, area: u.area,
    storeCode: u.storeCode, phone: u.phone,
    active: u.active !== 'N', mustChangePw: u.mustChangePw === 'Y',
    createdAt: u.createdAt, createdBy: u.createdBy
  };
}

function createUserRow(userId, pw, role, name, regionId, area, storeCode, createdBy, phone) {
  var salt = newToken().slice(0, 12);
  insertRow('Users', {
    userId: userId, pwHash: hash(String(pw), salt), salt: salt,
    role: role, name: name, regionId: regionId || '', area: area || '',
    storeCode: storeCode || '', phone: phone || '', active: 'Y', mustChangePw: 'Y',
    createdAt: fmt(Date.now()), createdBy: createdBy || ''
  });
}

/* 팀 목록 (담당 지역 포함) */
function teamList() {
  return readAll('Regions').map(function (r) {
    return {
      regionId: r.regionId, regionName: r.regionName,
      sortOrder: Number(r.sortOrder) || 0,
      areas: parseJson(r.areas, [])
    };
  }).sort(function (a, b) { return a.sortOrder - b.sortOrder; });
}

function findUser(userId) {
  if (!userId) return null;
  return find(readAll('Users'), 'userId', String(userId).trim());
}

/* ===================== 시트 유틸 ===================== */

function sheet(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error(name + ' 시트가 없습니다. setup 함수를 먼저 실행해 주세요.');
  return sh;
}

function readAll(name) {
  var sh = sheet(name);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var head = SHEETS[name];
  var vals = sh.getRange(2, 1, last - 1, head.length).getValues();
  return vals.map(function (row, i) {
    var o = { _row: i + 2 };
    head.forEach(function (h, c) { o[h] = row[c]; });
    return o;
  }).filter(function (o) { return String(o[SHEETS[name][0]]).length > 0; });
}

function insertRow(name, obj) {
  var head = SHEETS[name];
  sheet(name).appendRow(head.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; }));
}

function updateRow(name, rowIndex, patch) {
  var sh = sheet(name), head = SHEETS[name];
  var cur = sh.getRange(rowIndex, 1, 1, head.length).getValues()[0];
  head.forEach(function (h, c) { if (patch[h] !== undefined) cur[c] = patch[h]; });
  sh.getRange(rowIndex, 1, 1, head.length).setValues([cur]);
}

function deleteRowAt(name, rowIndex) {
  sheet(name).deleteRow(rowIndex);
}

function find(rows, key, val) {
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][key]) === String(val)) return rows[i];
  }
  return null;
}

function strip(o) {
  var c = {};
  Object.keys(o).forEach(function (k) { if (k !== '_row') c[k] = o[k]; });
  return c;
}

function cleanSessions() {
  var rows = readAll('Sessions');
  var now = Date.now();
  for (var i = rows.length - 1; i >= 0; i--) {
    if (Number(rows[i].expireMs) < now) deleteRowAt('Sessions', rows[i]._row);
  }
}

/* ===================== 기타 유틸 ===================== */

function hash(pw, salt) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '::' + pw, Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function newToken() {
  return Utilities.getUuid().replace(/-/g, '');
}

function fmt(ms) {
  if (!ms) return '';
  return Utilities.formatDate(new Date(Number(ms)), TZ, 'yyyy-MM-dd HH:mm');
}

function pad(n, len) {
  var s = String(n);
  while (s.length < len) s = '0' + s;
  return s;
}

function parseJson(s, dflt) {
  try { return JSON.parse(s); } catch (e) { return dflt; }
}
