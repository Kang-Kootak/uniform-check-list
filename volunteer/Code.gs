/**
 * 자치회 봉사 수첩 — 구글 스프레드시트 백엔드
 *
 * 학생자치회 부원이 선 아침 교통지도 · 점심 급식지도를 날마다 입력하고,
 * 학기마다 봉사활동 시간을 누계합니다.
 *
 * 설치: 새 스프레드시트 → 확장 프로그램 → Apps Script → 이 파일과 Index.html 붙여넣기
 *       → 스프레드시트 새로고침 → [자치회 봉사] 메뉴 → ① 초기 설정
 * 자세한 순서는 저장소의 volunteer/설치_방법.md 참고.
 */

var SH_M = '부원', SH_S = '순번', SH_A = '활동', SH_C = '설정', SH_T = '누계';
var HDR_M = ['학번', '이름', '학년', '반', '번호', '부서', '상태', '개인코드', '메모'];
var HDR_S = ['요일', '구분', '학번', '이름'];
var HDR_A = ['기록ID', '날짜', '구분', '학번', '이름', '인정분', '입력자', '입력일시', '메모'];
var SLOT = { am: '아침 교통지도', pm: '점심 급식지도' };
var DAYS = ['월', '화', '수', '목', '금'];
var ST_STOP = '중단';
var DEF_PM = 25;                    // 점심 급식지도 1회 인정 시간(분). 아침 교통지도는 입력할 때마다 넣습니다.
var ROUNDS = ['내림', '반올림', '그대로'];
var MAX_MIN = 600;                  // 한 번에 인정할 수 있는 최대 분
var BULK_MAX = 500;                 // 부원 일괄 등록에서 한 번에 받는 최대 줄 수
var CFG_TTL = 30;                   // 설정 캐시 (초). PIN을 바꾸면 최대 이만큼 뒤에 적용됩니다.

/* ─────────────────────────── 스프레드시트 메뉴 ─────────────────────────── */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('자치회 봉사')
    .addItem('① 초기 설정 (시트 만들기)', '초기설정')
    .addItem('② 앱 주소 보기', '앱주소보기')
    .addSeparator()
    .addItem('이번 학기 누계 시트 만들기', '학기누계')
    .addItem('백업 사본 만들기', '백업사본')
    .addToUi();
}

function 초기설정() {
  var ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty('ssid', ss.getId());

  var ms = ensureSheet_(ss, SH_M, HDR_M);
  ms.setColumnWidth(1, 80); ms.setColumnWidth(2, 100); ms.setColumnWidth(6, 110); ms.setColumnWidth(8, 100); ms.setColumnWidth(9, 200);
  ms.getRange(2, 8, Math.max(ms.getMaxRows() - 1, 1), 1).setNumberFormat('@');   // 개인코드는 글자로

  var ps = ensureSheet_(ss, SH_S, HDR_S);
  ps.setColumnWidth(2, 130);

  var as = ensureSheet_(ss, SH_A, HDR_A);
  as.setColumnWidth(1, 90); as.setColumnWidth(2, 100); as.setColumnWidth(3, 130); as.setColumnWidth(8, 150); as.setColumnWidth(9, 200);
  as.getRange(2, 2, Math.max(as.getMaxRows() - 1, 1), 1).setNumberFormat('@');   // 날짜는 2026-10-21 글자로

  var cs = ss.getSheetByName(SH_C) || ss.insertSheet(SH_C);
  if (cs.getLastRow() === 0) {
    cs.getRange(1, 1, 1, 3).setValues([['항목', '값', '설명']]);
    cs.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e2ecf6');
    cs.setColumnWidth(1, 110); cs.setColumnWidth(2, 200); cs.setColumnWidth(3, 480);
    cs.setFrozenRows(1);
  }
  addCfg_(cs, '기록PIN', '', '입력을 돕는 학생(회장단)·선생님이 들어올 때 쓰는 번호. 오늘 날짜만 입력할 수 있습니다.');
  addCfg_(cs, '관리PIN', '', '담당 선생님용. 지난 날짜 입력·수정, 부원 관리, 누계와 설정까지 모두 할 수 있습니다. 두 PIN을 모두 비워두면 주소를 아는 누구나 관리자로 들어옵니다.');
  addCfg_(cs, '학교명', '', '앱 위쪽에 표시할 이름 (선택)');
  addCfg_(cs, '급식지도분', DEF_PM, '점심 급식지도 한 번에 인정하는 시간(분). 아침 교통지도는 앱에서 입력할 때마다 시간을 넣습니다.');
  addCfg_(cs, '학기시작', '', '이번 학기 첫날 (예: 2026-08-17). 비워두면 1학기 3/1~8/15, 2학기 8/16~2월 말로 계산합니다.');
  addCfg_(cs, '학기종료', '', '이번 학기 마지막 날 (예: 2027-02-12)');
  addCfg_(cs, '시간계산', '내림', '누계 분을 결재용 시간으로 바꾸는 방법: 내림(5시간 50분 → 5시간) · 반올림 · 그대로(5.8시간)');
  ss.setSpreadsheetTimeZone('Asia/Seoul');
  bumpRev_();
  cache_().remove('cfg');

  SpreadsheetApp.getUi().alert(
    '시트를 준비했습니다.\n\n' +
    '다음 순서로 진행하세요.\n' +
    '1) [설정] 시트에서 기록PIN·관리PIN을 정합니다. (점심 급식지도는 1회 ' + DEF_PM + '분으로 들어가 있습니다)\n' +
    '2) 오른쪽 위 [배포] → [새 배포] → 유형 [웹 앱]\n' +
    '   · 다음 사용자로 실행: 나\n' +
    '   · 액세스 권한이 있는 사용자: 모든 사용자\n' +
    '3) 만들어진 웹 앱 주소를 관리 PIN으로 열고 [관리] 탭에서 부원을 등록합니다.\n' +
    '   ([자치회 봉사] → [② 앱 주소 보기]에서 다시 볼 수 있습니다.)');
}

function ensureSheet_(ss, name, hdr) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0 || String(sh.getRange(1, 1).getValue()).trim() !== hdr[0]) {
    sh.getRange(1, 1, 1, hdr.length).setValues([hdr]);
  }
  sh.getRange(1, 1, 1, hdr.length).setFontWeight('bold').setBackground('#e2ecf6');
  sh.setFrozenRows(1);
  return sh;
}

function appUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

function 앱주소보기() {
  var url = appUrl_();
  var ui = SpreadsheetApp.getUi();
  if (!url) {
    ui.alert('아직 배포되지 않았습니다.\n\n오른쪽 위 [배포] → [새 배포] → 유형 [웹 앱]으로 먼저 배포해 주세요.');
    return;
  }
  ui.alert('앱 주소\n\n' + url + '\n\n선생님·회장단은 이 주소와 PIN으로 들어옵니다.\n부원 개인 조회 링크는 앱의 [관리] 탭에서 만듭니다.');
}

function 학기누계() {
  var cfg = config_();
  var term = term_(cfg, today_());
  var r = writeTally_(term.from, term.to);
  SpreadsheetApp.getActive().setActiveSheet(r.sheet);
  SpreadsheetApp.getUi().alert('[' + SH_T + '] 시트에 ' + term.name + ' (' + term.from + ' ~ ' + term.to + ') 누계를 만들었습니다.\n\n부원 ' + r.rows + '명');
}

function 백업사본() {
  var ss = ss_();
  var name = '자치회 봉사 백업 ' + Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HHmm');
  var copy = ss.copy(name);
  try { SpreadsheetApp.getUi().alert('백업 사본을 만들었습니다.\n\n' + name + '\n\n내 드라이브에서 확인하세요.'); } catch (e) { /* 앱에서 부르면 알림창이 없다 */ }
  return copy.getId();
}

/* ─────────────────────────── 웹 앱 ─────────────────────────── */

function doGet(e) {
  var html = HtmlService.createHtmlOutputFromFile('Index').getContent();
  html = html.replace('/*__SEED__*/', 'SEED=' + JSON.stringify(seed_(e)) + ';');
  return HtmlService.createHtmlOutput(html)
    .setTitle('자치회 봉사 수첩')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/**
 * 주소에 실어 보낸 값을 읽는다.
 *   …/exec?c=K7HDQ2MN        부원 개인 조회 링크
 *   …/exec?pin=2026&by=김선생 PIN·입력자 이름을 다시 묻지 않게
 */
function seed_(e) {
  var p = (e && e.parameter) || {};
  var clean = function (v, n) {
    return String(v == null ? '' : v).replace(/[<>]/g, '').trim().slice(0, n);
  };
  return { pin: clean(p.pin, 20), by: clean(p.by, 40), code: clean(p.c, 20) };
}

/** 앱에서 오는 모든 요청의 단일 입구 */
function api(p) {
  p = p || {};
  try {
    var who = identify_(p);
    if (!who) return { ok: false, err: String(p.code || '').trim() ? 'CODE' : 'PIN' };
    var admin = who.role === 'admin';
    var staff = admin || who.role === 'record';
    switch (p.op) {
      case 'rev':       return ok_({ rev: rev_() });
      case 'boot':      return ok_(who.role === 'member' ? bootMember_(who) : bootStaff_(who.role));
      case 'save':      return staff ? ok_(save_(p, admin)) : deny_();
      case 'del':       return staff ? ok_(delAct_(p, admin)) : deny_();
      case 'sched_set': return admin ? ok_(schedSet_(p.rows || [])) : deny_();
      case 'm_bulk':    return admin ? ok_(memberBulk_(p.list || [])) : deny_();
      case 'm_state':   return admin ? ok_(memberState_(p)) : deny_();
      case 'm_del':     return admin ? ok_(memberDel_(normNo_(p.no))) : deny_();
      case 'code_gen':  return admin ? ok_(codeGen_(p)) : deny_();
      case 'cfg_set':   return admin ? ok_(cfgSet_(p)) : deny_();
      case 'tally':     return admin ? ok_(tally_(p)) : deny_();
      case 'backup':    return admin ? ok_({ id: 백업사본() }) : deny_();
      default:          return { ok: false, err: 'OP', message: '알 수 없는 요청입니다.' };
    }
  } catch (e) {
    return { ok: false, err: 'ERR', message: String((e && e.message) || e) };
  }
}

function ok_(data) { return { ok: true, data: data }; }
function deny_() { return { ok: false, err: 'ROLE', message: '관리 PIN이 필요한 기능입니다.' }; }

/**
 * 누가 들어왔는지
 *  · 개인코드(?c=…) 가 부원 시트에 있으면 → 그 부원 본인 (자기 기록만 봅니다)
 *  · 아니면 PIN 으로 → 관리 / 기록
 */
function identify_(p) {
  var code = String(p.code || '').trim().toUpperCase();
  if (code) {
    var mem = readMembers_(ss_());
    for (var i = 0; i < mem.length; i++) {
      if (mem[i].code && mem[i].code === code) return { role: 'member', no: mem[i].no, name: mem[i].name };
    }
  }
  var role = auth_(p.pin);
  return role ? { role: role } : null;
}

/**
 * PIN 규칙
 *  · 둘 다 비어 있으면  → 누구나 관리자 (초기 상태, 앱이 경고를 띄웁니다)
 *  · 관리PIN 일치       → 관리자
 *  · 기록PIN 일치       → 관리PIN이 있으면 기록 권한, 없으면 관리자
 *  · 그 외              → 거부
 */
function auth_(pin) {
  var cfg = config_();
  var rec = String(cfg['기록PIN'] || '').trim();
  var adm = String(cfg['관리PIN'] || '').trim();
  pin = String(pin == null ? '' : pin).trim();
  if (!rec && !adm) return 'admin';
  if (adm && pin === adm) return 'admin';
  if (rec && pin === rec) return adm ? 'record' : 'admin';
  return null;
}

/* ─────────────────────────── 읽기 ─────────────────────────── */

function bootStaff_(role) {
  var ss = ss_(), cfg = config_(), today = today_();
  var admin = role === 'admin';
  var members = readMembers_(ss).map(function (m) {
    return { no: m.no, name: m.name, grade: m.grade, klass: m.klass, num: m.num, dept: m.dept,
             state: m.state, code: admin ? m.code : '', memo: admin ? m.memo : '' };
  });
  var acts = readActs_(ss);
  // 기록 권한은 오늘만 다루므로 오늘 기록만 내려준다
  if (!admin) acts = acts.filter(function (a) { return a.d === today; });
  return {
    role: role,
    members: members,
    sched: readSched_(ss),
    acts: acts.map(packAct_),
    cfg: cfgOut_(cfg, today),
    today: today,
    pinSet: !!(String(cfg['기록PIN'] || '').trim() || String(cfg['관리PIN'] || '').trim()),
    sheetUrl: admin ? ss.getUrl() : '',
    appUrl: admin ? appUrl_() : '',
    pins: admin ? { rec: String(cfg['기록PIN'] || '').trim(), adm: String(cfg['관리PIN'] || '').trim() } : null,
    rev: rev_()
  };
}

/** 부원 본인 — 자기 기록과 자기 순번만 */
function bootMember_(who) {
  var ss = ss_(), cfg = config_(), today = today_();
  var me = null;
  readMembers_(ss).forEach(function (m) { if (m.no === who.no) me = m; });
  return {
    role: 'member',
    me: { no: who.no, name: who.name, dept: me ? me.dept : '', state: me ? me.state : '' },
    sched: readSched_(ss).filter(function (s) { return s.no === who.no; }),
    acts: readActs_(ss).filter(function (a) { return a.no === who.no; }).map(packAct_),
    cfg: cfgOut_(cfg, today),
    today: today,
    rev: rev_()
  };
}

function packAct_(a) {
  return { id: a.id, d: a.d, s: a.s, no: a.no, name: a.name, m: a.m, by: a.by, at: a.at, memo: a.memo };
}

function cfgOut_(cfg, today) {
  var t = term_(cfg, today);
  return {
    school: String(cfg['학교명'] || '').trim(),
    pm: pmMin_(cfg),
    round: roundMode_(cfg),
    from: t.from, to: t.to, termName: t.name, termSet: t.set
  };
}

function readMembers_(ss) {
  var sh = ss.getSheetByName(SH_M);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, HDR_M.length).getValues();
  var out = [], seen = {};
  for (var i = 0; i < vals.length; i++) {
    var no = normNo_(vals[i][0]);
    if (!no || seen[no]) continue;
    seen[no] = true;
    var d = derive_(no);
    out.push({
      row: i + 2, no: no,
      name: String(vals[i][1] || '').trim(),
      grade: num_(vals[i][2]) || d.grade, klass: num_(vals[i][3]) || d.klass, num: num_(vals[i][4]) || d.num,
      dept: String(vals[i][5] || '').trim(),
      state: String(vals[i][6] || '').trim() === ST_STOP ? ST_STOP : '활동',
      code: String(vals[i][7] || '').trim().toUpperCase(),
      memo: String(vals[i][8] || '').trim()
    });
  }
  out.sort(function (a, b) { return cmpNo_(a.no, b.no); });
  return out;
}

function memberMap_(ss) {
  var map = {};
  readMembers_(ss).forEach(function (m) { map[m.no] = m; });
  return map;
}

function readSched_(ss) {
  var sh = ss.getSheetByName(SH_S);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, HDR_S.length).getValues();
  var out = [], seen = {};
  for (var i = 0; i < vals.length; i++) {
    var day = String(vals[i][0] || '').trim().charAt(0);
    var slot = slotKey_(vals[i][1]);
    var no = normNo_(vals[i][2]);
    if (DAYS.indexOf(day) < 0 || !slot || !no) continue;
    var k = day + slot + no;
    if (seen[k]) continue;
    seen[k] = true;
    out.push({ day: day, slot: slot, no: no });
  }
  return out;
}

function readActs_(ss) {
  var sh = ss.getSheetByName(SH_A);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, HDR_A.length).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i];
    var d = dstr_(v[1]), s = slotKey_(v[2]), no = normNo_(v[3]);
    if (!d || !s || !no) continue;
    out.push({
      row: i + 2, id: String(v[0] || ''), d: d, s: s, no: no,
      name: String(v[4] || '').trim(),
      m: Math.max(0, num_(v[5]) || 0),
      by: String(v[6] || ''), at: iso_(v[7]), memo: String(v[8] || '')
    });
  }
  return out;
}

/** '아침 교통지도' · '교통' · 'am' 무엇이든 am/pm 으로 */
function slotKey_(v) {
  var t = String(v || '').trim();
  if (t === 'am' || /아침|교통|등교/.test(t)) return 'am';
  if (t === 'pm' || /점심|급식/.test(t)) return 'pm';
  return '';
}

/* ─────────────────────────── 지도 확인 ─────────────────────────── */

function by_(p) {
  var v = String(p.by || '').trim().slice(0, 40);
  if (!v) throw new Error('입력자 이름을 먼저 넣어 주세요.');
  return v;
}

/** 기록 권한은 오늘만, 관리자는 오늘까지의 모든 날짜 */
function checkDay_(date, slot, admin) {
  if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('날짜가 올바르지 않습니다.');
  if (!SLOT[slot]) throw new Error('구분(아침/점심)이 올바르지 않습니다.');
  var today = today_();
  if (date > today) throw new Error('아직 오지 않은 날짜는 기록할 수 없습니다.');
  if (!admin && date !== today) throw new Error('지난 날짜는 담당 선생님(관리 PIN)만 고칠 수 있습니다.');
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('다른 사람이 기록하는 중입니다. 잠시 후 다시 시도해 주세요.');
  try { return fn(); } finally { lock.releaseLock(); }
}

/**
 * 그날 선 학생 한 명과 인정 시간(분)을 입력한다.
 * 같은 날 같은 구분에 같은 학생은 한 줄만 남습니다. 다시 입력하면 시간·메모를 고칩니다.
 * 두 사람이 동시에 입력해도 겹치지 않습니다.
 */
function save_(p, admin) {
  var date = dstr_(p.date), slot = String(p.slot || ''), no = normNo_(p.no);
  checkDay_(date, slot, admin);
  var by = by_(p);
  var mins = Math.round(Number(p.min));
  if (String(p.min == null ? '' : p.min).trim() === '' || !isFinite(mins) || mins < 1 || mins > MAX_MIN) {
    throw new Error('봉사 시간을 1~' + MAX_MIN + '분으로 넣어 주세요.');
  }
  var memo = String(p.memo == null ? '' : p.memo).trim().slice(0, 100);
  return withLock_(function () {
    var ss = ss_();
    var m = memberMap_(ss)[no];
    var sh = sheet_(ss, SH_A);
    var cur = null;
    readActs_(ss).forEach(function (a) { if (a.d === date && a.s === slot && a.no === no) cur = a; });
    if (!m && !cur) throw new Error('부원 명단에 없는 학번입니다.');
    var name = m ? m.name : cur.name;
    var now = new Date();
    var id = cur ? cur.id : newId_();
    var row = [id, date, SLOT[slot], no, name, mins, by, now, memo];
    if (cur) sh.getRange(cur.row, 1, 1, HDR_A.length).setValues([row]);
    else appendText_(sh, row);
    bumpRev_();
    return {
      rec: { id: id, d: date, s: slot, no: no, name: name, m: mins, by: by, at: now.toISOString(), memo: memo },
      updated: !!cur, before: cur ? cur.m : null, rev: rev_()
    };
  });
}

/** 입력한 한 건을 지운다. 기록 권한은 오늘 것만 지울 수 있습니다. */
function delAct_(p, admin) {
  var id = String(p.id || '');
  return withLock_(function () {
    var ss = ss_(), cur = null;
    readActs_(ss).forEach(function (a) { if (a.id === id) cur = a; });
    if (!cur) throw new Error('기록을 찾지 못했습니다. 새로고침 후 다시 시도해 주세요.');
    checkDay_(cur.d, cur.s, admin);
    sheet_(ss, SH_A).deleteRow(cur.row);
    bumpRev_();
    return { id: id, rev: rev_() };
  });
}

function appendText_(sh, row) {
  var r = sh.getLastRow() + 1;
  sh.getRange(r, 2).setNumberFormat('@');
  sh.getRange(r, 1, 1, row.length).setValues([row]);
}

function newId_() {
  return Utilities.formatDate(new Date(), tz_(), 'MMdd') + '-' + Math.random().toString(36).slice(2, 7);
}

/* ─────────────────────────── 순번표 ─────────────────────────── */

/** 순번표 전체를 앱에서 보낸 것으로 바꾼다. 이미 남긴 활동 기록은 건드리지 않습니다. */
function schedSet_(rows) {
  return withLock_(function () {
    var ss = ss_(), mem = memberMap_(ss);
    var sh = ensureSheet_(ss, SH_S, HDR_S);
    var out = [], seen = {};
    rows.forEach(function (r) {
      var day = String(r.day || '').charAt(0), slot = String(r.slot || ''), no = normNo_(r.no);
      if (DAYS.indexOf(day) < 0 || !SLOT[slot] || !mem[no]) return;
      var k = day + slot + no;
      if (seen[k]) return;
      seen[k] = true;
      out.push([day, SLOT[slot], no, mem[no].name]);
    });
    out.sort(function (a, b) {
      return DAYS.indexOf(a[0]) - DAYS.indexOf(b[0]) || (a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0) || cmpNo_(a[2], b[2]);
    });
    var last = sh.getLastRow();
    if (last >= 2) sh.getRange(2, 1, last - 1, HDR_S.length).clearContent();
    if (out.length) sh.getRange(2, 1, out.length, HDR_S.length).setValues(out);
    bumpRev_();
    return { total: out.length, rev: rev_() };
  });
}

/* ─────────────────────────── 부원 ─────────────────────────── */

function memberBulk_(list) {
  if (list.length > BULK_MAX) throw new Error('한 번에 ' + BULK_MAX + '명까지 등록할 수 있습니다.');
  return withLock_(function () {
    var ss = ss_(), sh = ensureSheet_(ss, SH_M, HDR_M);
    var have = {};
    readMembers_(ss).forEach(function (m) { have[m.no] = m.row; });
    var added = 0, updated = 0, rows = [];
    list.forEach(function (x) {
      var no = normNo_(x.no), name = String(x.name || '').trim().slice(0, 30);
      var dept = String(x.dept || '').trim().slice(0, 30);
      if (!no || !name) return;
      var d = derive_(no);
      if (have[no]) {
        sh.getRange(have[no], 2).setValue(name);
        if (dept) sh.getRange(have[no], 6).setValue(dept);
        updated++;
      } else {
        have[no] = -1;
        rows.push([no, name, d.grade, d.klass, d.num, dept, '활동', '', '']);
        added++;
      }
    });
    if (rows.length) {
      var start = sh.getLastRow() + 1;
      sh.getRange(start, 8, rows.length, 1).setNumberFormat('@');
      sh.getRange(start, 1, rows.length, HDR_M.length).setValues(rows);
    }
    sortMembers_(sh);
    bumpRev_();
    return { added: added, updated: updated, total: readMembers_(ss).length };
  });
}

function memberState_(p) {
  var no = normNo_(p.no), state = String(p.state || '') === ST_STOP ? ST_STOP : '활동';
  return withLock_(function () {
    var ss = ss_(), m = memberMap_(ss)[no];
    if (!m) throw new Error('부원 명단에 없는 학번입니다.');
    sheet_(ss, SH_M).getRange(m.row, 7).setValue(state);
    if (state === ST_STOP) dropFromSched_(ss, no);   // 활동을 멈추면 순번표에서도 뺀다
    bumpRev_();
    return { no: no, state: state };
  });
}

/** 부원을 지워도 이미 남긴 활동 기록은 그대로 둡니다 (누계에 계속 잡힙니다) */
function memberDel_(no) {
  return withLock_(function () {
    var ss = ss_(), m = memberMap_(ss)[no];
    if (!m) throw new Error('부원 명단에 없는 학번입니다.');
    sheet_(ss, SH_M).deleteRow(m.row);
    dropFromSched_(ss, no);
    bumpRev_();
    return { no: no };
  });
}

function dropFromSched_(ss, no) {
  var sh = ss.getSheetByName(SH_S);
  if (!sh) return;
  var last = sh.getLastRow();
  if (last < 2) return;
  var vals = sh.getRange(2, 3, last - 1, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (normNo_(vals[i][0]) === no) sh.deleteRow(i + 2);
  }
}

function sortMembers_(sh) {
  var last = sh.getLastRow();
  if (last < 3) return;
  sh.getRange(2, 1, last - 1, HDR_M.length).sort({ column: 1, ascending: true });
}

/**
 * 부원 개인 조회 링크의 코드
 *  · p.no 가 있으면 그 사람 코드를 새로 만든다 (예전 링크는 더 이상 열리지 않습니다)
 *  · 없으면 코드가 없는 모든 부원에게 만들어 준다
 */
function codeGen_(p) {
  var one = normNo_(p.no);
  return withLock_(function () {
    var ss = ss_(), sh = sheet_(ss, SH_M), mem = readMembers_(ss);
    var used = {};
    mem.forEach(function (m) { if (m.code) used[m.code] = true; });
    var made = 0;
    mem.forEach(function (m) {
      if (one ? m.no !== one : m.code) return;
      var c = newCode_();
      while (used[c]) c = newCode_();
      used[c] = true;
      sh.getRange(m.row, 8).setNumberFormat('@').setValue(c);
      made++;
    });
    if (one && !made) throw new Error('부원 명단에 없는 학번입니다.');
    bumpRev_();
    return { made: made };
  });
}

/** 헷갈리는 글자(0 O 1 l I)는 빼고 만든다 */
function newCode_() {
  var abc = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  var out = '';
  for (var i = 0; i < 8; i++) out += abc.charAt(Math.floor(Math.random() * abc.length));
  return out;
}

/* ─────────────────────────── 설정 ─────────────────────────── */

function cfgSet_(p) {
  var pm = Math.round(Number(p.pm));
  if (!(pm > 0 && pm <= MAX_MIN)) throw new Error('점심 급식지도 시간은 1~' + MAX_MIN + '분으로 넣어 주세요.');
  var from = dstr_(p.from), to = dstr_(p.to);
  if ((from && !to) || (!from && to)) throw new Error('학기 시작일과 종료일을 함께 넣어 주세요. (둘 다 비우면 자동으로 계산합니다)');
  if (from && from > to) throw new Error('학기 종료일이 시작일보다 빠릅니다.');
  var round = ROUNDS.indexOf(String(p.round)) >= 0 ? String(p.round) : '내림';
  return withLock_(function () {
    var cs = sheet_(ss_(), SH_C);
    setCfg_(cs, '급식지도분', pm);
    setCfg_(cs, '학기시작', from);
    setCfg_(cs, '학기종료', to);
    setCfg_(cs, '시간계산', round);
    cache_().remove('cfg');
    bumpRev_();
    return { cfg: cfgOut_(config_(), today_()) };
  });
}

/** 점심 급식지도 1회 인정 시간(분) */
function pmMin_(cfg) {
  var n = Math.round(Number(cfg['급식지도분']));
  return n > 0 && n <= MAX_MIN ? n : DEF_PM;
}

function roundMode_(cfg) {
  var v = String(cfg['시간계산'] || '').trim();
  return ROUNDS.indexOf(v) >= 0 ? v : '내림';
}

/** 누계 분 → 결재에 올릴 시간 */
function credit_(mins, mode) {
  var h = mins / 60;
  if (mode === '반올림') return Math.round(h);
  if (mode === '그대로') return Math.round(h * 10) / 10;
  return Math.floor(h);
}

/**
 * 이번 학기 기간. 설정 시트에 적어두면 그 날짜를,
 * 비어 있으면 1학기 3/1~8/15 · 2학기 8/16~이듬해 2월 말로 계산합니다.
 */
function term_(cfg, today) {
  var from = dstr_(cfg['학기시작']), to = dstr_(cfg['학기종료']);
  if (/^\d{4}-\d{2}-\d{2}$/.test(from) && /^\d{4}-\d{2}-\d{2}$/.test(to) && from <= to) {
    return { from: from, to: to, name: termName_(from), set: true };
  }
  var y = +today.slice(0, 4), md = today.slice(5);
  if (md >= '03-01' && md <= '08-15') return { from: y + '-03-01', to: y + '-08-15', name: y + '학년도 1학기', set: false };
  var sy = md < '03-01' ? y - 1 : y;
  var leap = ((sy + 1) % 4 === 0 && (sy + 1) % 100 !== 0) || (sy + 1) % 400 === 0;
  return { from: sy + '-08-16', to: (sy + 1) + '-02-' + (leap ? '29' : '28'), name: sy + '학년도 2학기', set: false };
}

function termName_(from) {
  var y = +from.slice(0, 4), m = +from.slice(5, 7);
  if (m >= 2 && m <= 7) return y + '학년도 1학기';
  return (m === 1 ? y - 1 : y) + '학년도 2학기';
}

/* ─────────────────────────── 누계 ─────────────────────────── */

function tally_(p) {
  var from = dstr_(p.from), to = dstr_(p.to);
  if (!from || !to || from > to) throw new Error('기간을 확인해 주세요.');
  var r = writeTally_(from, to, String(p.title || '').slice(0, 60));
  return { rows: r.rows, url: ss_().getUrl() + '#gid=' + r.sheet.getSheetId() };
}

/** 기간 안의 부원별 누계를 [누계] 시트에 새로 쓴다 — 결재 서류에 그대로 붙여 씁니다 */
function writeTally_(from, to, title) {
  var ss = ss_(), cfg = config_(), mode = roundMode_(cfg);
  var by = {};
  readMembers_(ss).forEach(function (m) {
    by[m.no] = { no: m.no, name: m.name, dept: m.dept, am: 0, pm: 0, min: 0, keep: m.state !== ST_STOP };
  });
  readActs_(ss).forEach(function (a) {
    if (a.d < from || a.d > to) return;
    var t = by[a.no] || (by[a.no] = { no: a.no, name: a.name, dept: '', am: 0, pm: 0, min: 0, keep: false });
    t[a.s]++;
    t.min += a.m;
  });
  var list = Object.keys(by).map(function (k) { return by[k]; })
    .filter(function (t) { return t.keep || t.am || t.pm; })
    .sort(function (a, b) { return cmpNo_(a.no, b.no); });

  var sh = ss.getSheetByName(SH_T) || ss.insertSheet(SH_T);
  sh.clear();
  var head = ['학번', '이름', '부서', '교통지도(회)', '급식지도(회)', '인정 분', '시간 환산', '인정 시간'];
  var school = String(cfg['학교명'] || '').trim();
  sh.getRange(1, 1).setValue((school ? school + ' ' : '') + (title || termName_(from)) + ' 학생자치회 지도 봉사활동 누계');
  sh.getRange(2, 1).setValue('기간 ' + from + ' ~ ' + to + '  ·  점심 급식지도 1회 ' + pmMin_(cfg) + '분, 아침 교통지도는 날마다 입력한 시간  ·  시간 계산: ' + mode +
    '  ·  만든 때 ' + Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'));
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  sh.getRange(2, 1).setFontColor('#6a7589');
  sh.getRange(4, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e2ecf6');
  if (list.length) {
    var rows = list.map(function (t) {
      return [t.no, t.name, t.dept, t.am, t.pm, t.min, hm_(t.min), credit_(t.min, mode)];
    });
    sh.getRange(5, 1, rows.length, head.length).setValues(rows);
    var sum = rows.reduce(function (s, r) { s[0] += r[3]; s[1] += r[4]; s[2] += r[5]; return s; }, [0, 0, 0]);
    sh.getRange(5 + rows.length, 1, 1, head.length)
      .setValues([['합계', list.length + '명', '', sum[0], sum[1], sum[2], hm_(sum[2]), '']])
      .setFontWeight('bold').setBorder(true, null, null, null, null, null);
  }
  sh.setFrozenRows(4);
  sh.setColumnWidth(2, 100); sh.setColumnWidth(3, 110); sh.setColumnWidth(7, 110);
  return { sheet: sh, rows: list.length };
}

function hm_(mins) {
  var h = Math.floor(mins / 60), m = mins % 60;
  return (h ? h + '시간' : '') + (h && m ? ' ' : '') + (m || !h ? m + '분' : '');
}

/* ─────────────────────────── 도우미 ─────────────────────────── */

function addCfg_(cs, key, value, desc) {
  if (cfgRow_(cs, key)) return;
  cs.appendRow([key, value, desc || '']);
  cs.getRange(cs.getLastRow(), 3).setWrap(true);
}

function setCfg_(cs, key, value) {
  var row = cfgRow_(cs, key);
  if (row) cs.getRange(row, 2).setValue(value);
  else cs.appendRow([key, value, '']);
}

function cfgRow_(cs, key) {
  var last = cs.getLastRow();
  if (last < 2) return 0;
  var vals = cs.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) if (String(vals[i][0] || '').trim() === key) return i + 2;
  return 0;
}

function ss_() {
  var id = PropertiesService.getScriptProperties().getProperty('ssid');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) {}
  }
  var active = SpreadsheetApp.getActive();
  if (!active) throw new Error('스프레드시트를 찾지 못했습니다. [자치회 봉사] → [① 초기 설정]을 먼저 실행해 주세요.');
  return active;
}

function sheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('“' + name + '” 시트가 없습니다. [자치회 봉사] → [① 초기 설정]을 실행해 주세요.');
  return sh;
}

function cache_() { return CacheService.getScriptCache(); }

function config_() {
  var hit = cache_().get('cfg');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var out = {};
  try {
    var cs = sheet_(ss_(), SH_C);
    var last = cs.getLastRow();
    if (last >= 2) {
      var vals = cs.getRange(2, 1, last - 1, 2).getValues();
      for (var i = 0; i < vals.length; i++) {
        var k = String(vals[i][0] || '').trim();
        // 날짜 칸은 셀이 날짜로 바뀌어 있어도 글자로 맞춰 둔다 (캐시에 넣기 전에)
        if (k) out[k] = (vals[i][1] instanceof Date) ? dstr_(vals[i][1]) : vals[i][1];
      }
    }
  } catch (e) { /* 초기 설정 전 */ }
  cache_().put('cfg', JSON.stringify(out), CFG_TTL);
  return out;
}

function rev_() {
  return Number(PropertiesService.getScriptProperties().getProperty('rev') || 0);
}
function bumpRev_() {
  PropertiesService.getScriptProperties().setProperty('rev', String(Date.now()));
}

function cmpNo_(a, b) { return a.length - b.length || (a < b ? -1 : a > b ? 1 : 0); }

function normNo_(v) {
  if (v == null) return '';
  return String(v).replace(/\D/g, '');
}

function num_(v) {
  var n = Number(v);
  return isFinite(n) && v !== '' && v !== null ? n : null;
}

/** 5자리 10312 → 1학년 3반 12번, 4자리 1312 → 1학년 3반 12번 */
function derive_(no) {
  var o = { grade: '', klass: '', num: '' };
  if (no.length === 5) { o.grade = +no[0]; o.klass = +no.slice(1, 3); o.num = +no.slice(3, 5); }
  else if (no.length === 4) { o.grade = +no[0]; o.klass = +no[1]; o.num = +no.slice(2); }
  return o;
}

function iso_(v) {
  if (!v) return null;
  var d = (v instanceof Date) ? v : new Date(v);
  return isNaN(d.getTime()) ? null : d.toISOString();
}

/** 무엇이 들어와도 yyyy-MM-dd 문자열로 (알아볼 수 없으면 빈 글자) */
function dstr_(v) {
  if (!v && v !== 0) return '';
  if (v instanceof Date) {
    return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  }
  var m = /^(\d{4})[-.\/]\s*(\d{1,2})[-.\/]\s*(\d{1,2})/.exec(String(v).trim());
  return m ? m[1] + '-' + p2_(m[2]) + '-' + p2_(m[3]) : '';
}
function p2_(n) { n = String(n); return n.length < 2 ? '0' + n : n; }

function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }

function tz_() {
  try { return ss_().getSpreadsheetTimeZone() || 'Asia/Seoul'; } catch (e) { return 'Asia/Seoul'; }
}
