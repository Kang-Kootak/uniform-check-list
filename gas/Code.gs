/**
 * 교복 지도 수첩 — 구글 스프레드시트 백엔드
 *
 * 설치: 스프레드시트 → 확장 프로그램 → Apps Script → 이 파일과 Index.html 붙여넣기
 *       → 스프레드시트 새로고침 → [교복 지도] 메뉴 → ① 초기 설정
 * 자세한 순서는 저장소의 gas/설치_방법.md 참고.
 */

var SH_R = '명단', SH_L = '기록', SH_C = '설정', SH_E = '면제';
var HDR_R = ['학번', '이름', '학년', '반', '번호', '누적횟수', '최근적발'];
var HDR_L = ['기록ID', '일시', '학번', '이름', '학년', '반', '사유', '메모', '기록자'];
var HDR_E = ['면제ID', '학번', '이름', '사유', '시작일', '종료일', '등록자', '등록일시'];
var EXEMPT_LIMIT = 500;
var REASON = '교복 미착용';   // 적발 사유는 이 한 가지로 통일
var DEF_WARN = 10;            // 이 횟수부터 회부 경고를 띄운다
var DEF_REFER = 15;           // 이 횟수에 도달하면 학생선도위원회 회부 대상
var LOG_LIMIT = 1000;      // 앱이 한 번에 받아가는 최근 기록 수
var CFG_TTL = 30;          // 설정 캐시 (초). PIN을 바꾸면 최대 이만큼 뒤에 적용됩니다.

/* ─────────────────────────── 스프레드시트 메뉴 ─────────────────────────── */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('교복 지도')
    .addItem('① 초기 설정 (시트 만들기)', '초기설정')
    .addItem('② 앱 주소 보기', '앱주소보기')
    .addSeparator()
    .addItem('오늘 자료 백업 사본 만들기', '백업사본')
    .addItem('적발 기록만 초기화', '기록초기화')
    .addToUi();
}

function 초기설정() {
  var ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty('ssid', ss.getId());

  var rs = ss.getSheetByName(SH_R) || ss.insertSheet(SH_R);
  if (rs.getLastRow() === 0 || String(rs.getRange(1, 1).getValue()).trim() !== HDR_R[0]) {
    rs.getRange(1, 1, 1, HDR_R.length).setValues([HDR_R]);
  }
  rs.getRange(1, 1, 1, HDR_R.length).setFontWeight('bold').setBackground('#e2ecf6');
  rs.setFrozenRows(1);
  rs.setColumnWidth(1, 90); rs.setColumnWidth(2, 110); rs.setColumnWidth(7, 150);

  var ls = ss.getSheetByName(SH_L) || ss.insertSheet(SH_L);
  if (ls.getLastRow() === 0 || String(ls.getRange(1, 1).getValue()).trim() !== HDR_L[0]) {
    ls.getRange(1, 1, 1, HDR_L.length).setValues([HDR_L]);
  }
  ls.getRange(1, 1, 1, HDR_L.length).setFontWeight('bold').setBackground('#e2ecf6');
  ls.setFrozenRows(1);
  ls.setColumnWidth(1, 90); ls.setColumnWidth(2, 150); ls.setColumnWidth(7, 160); ls.setColumnWidth(8, 200);

  var es = ensureExemptSheet_(ss);
  es.getRange(1, 1, 1, HDR_E.length).setFontWeight('bold').setBackground('#e2ecf6');
  es.setColumnWidth(1, 90); es.setColumnWidth(4, 220); es.setColumnWidth(8, 150);

  var cs = ss.getSheetByName(SH_C) || ss.insertSheet(SH_C);
  if (cs.getLastRow() === 0) {
    cs.getRange(1, 1, 4, 3).setValues([
      ['항목', '값', '설명'],
      ['기록PIN', '', '선도부원·선생님이 앱에 들어올 때 입력하는 숫자입니다. 비워두면 링크를 아는 누구나 관리자로 들어옵니다.'],
      ['관리PIN', '', '명단 편집·삭제 권한. 기록PIN과 다르게 정하세요.'],
      ['학교명', '', '앱 화면에 표시할 이름 (선택)']
    ]);
    cs.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e2ecf6');
    cs.setColumnWidth(1, 100); cs.setColumnWidth(2, 260); cs.setColumnWidth(3, 460);
    cs.getRange(2, 3, 3, 1).setWrap(true);
    cs.setFrozenRows(1);
  }
  setCfg_(cs, '경고기준', DEF_WARN, '누적 적발이 이 횟수에 이르면 앱에 회부 경고가 뜹니다. 앱의 [명단 → 선도위원회 기준]에서도 고칠 수 있습니다.');
  setCfg_(cs, '회부기준', DEF_REFER, '누적 적발이 이 횟수에 도달하면 학생선도위원회 회부 대상으로 표시합니다.');
  ss.setSpreadsheetTimeZone('Asia/Seoul');
  bumpRev_();
  cache_().remove('cfg');

  SpreadsheetApp.getUi().alert(
    '시트를 준비했습니다.\n\n' +
    '다음 순서로 진행하세요.\n' +
    '1) [설정] 시트에서 기록PIN·관리PIN을 정합니다. (예: 기록 2024 / 관리 7788)\n' +
    '2) 오른쪽 위 [배포] → [새 배포] → 유형 [웹 앱]\n' +
    '   · 다음 사용자로 실행: 나\n' +
    '   · 액세스 권한이 있는 사용자: 모든 사용자\n' +
    '3) 만들어진 웹 앱 URL을 선도부원에게 공유하세요.\n' +
    '   ([교복 지도] → [② 앱 주소 보기]에서 다시 볼 수 있습니다.)');
}

function 앱주소보기() {
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  var ui = SpreadsheetApp.getUi();
  if (!url) {
    ui.alert('아직 배포되지 않았습니다.\n\n오른쪽 위 [배포] → [새 배포] → 유형 [웹 앱]으로 먼저 배포해 주세요.');
    return;
  }
  ui.alert('앱 주소\n\n' + url + '\n\n이 주소를 선도부원에게 공유하세요.\n휴대폰에서 열고 홈 화면에 추가하면 앱처럼 쓸 수 있습니다.');
}

function 백업사본() {
  var ss = SpreadsheetApp.getActive();
  var name = '교복지도 백업 ' + Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HHmm');
  var copy = ss.copy(name);
  SpreadsheetApp.getUi().alert('백업 사본을 만들었습니다.\n\n' + name + '\n\n내 드라이브에서 확인하세요.');
  return copy.getId();
}

function 기록초기화() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert('적발 기록 초기화',
    '학생 명단은 그대로 두고 적발 기록을 모두 지웁니다.\n먼저 [오늘 자료 백업 사본 만들기]를 실행하시길 권합니다.\n\n계속할까요?',
    ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;
  resetCounts_();
  ui.alert('적발 기록을 초기화했습니다.');
}

/* ─────────────────────────── 웹 앱 ─────────────────────────── */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('교복 지도 수첩')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** 앱에서 오는 모든 요청의 단일 입구 */
function api(p) {
  p = p || {};
  try {
    var role = auth_(p.pin);
    if (!role) return { ok: false, err: 'PIN' };
    var admin = (role === 'admin');
    switch (p.op) {
      case 'rev':     return ok_({ rev: rev_() });
      case 'boot':    return ok_(boot_(role));
      case 'history': return ok_({ no: normNo_(p.no), items: history_(normNo_(p.no)) });
      case 'add':     return ok_(addRecord_(p));
      case 'undo':    return admin ? ok_(undoRecord_(p)) : deny_();
      case 'bulk':    return admin ? ok_(bulkRoster_(p.list || [])) : deny_();
      case 'put':     return admin ? ok_(putStudent_(p)) : deny_();
      case 'del':     return admin ? ok_(delStudent_(normNo_(p.no))) : deny_();
      case 'ex_add':  return admin ? ok_(exemptAdd_(p)) : deny_();
      case 'ex_del':  return admin ? ok_(exemptDel_(String(p.id || ''))) : deny_();
      case 'th_set':  return admin ? ok_(saveThresholds_(p)) : deny_();
      case 'reset':   return admin ? ok_(resetCounts_()) : deny_();
      case 'wipe':    return admin ? ok_(wipeRoster_()) : deny_();
      case 'backup':  return admin ? ok_({ id: 백업사본() }) : deny_();
      default:        return { ok: false, err: 'OP', message: '알 수 없는 요청입니다.' };
    }
  } catch (e) {
    if (e && e.dup) {
      return { ok: false, err: 'DUP', at: e.at, count: e.count, message: e.message };
    }
    return { ok: false, err: 'ERR', message: String((e && e.message) || e) };
  }
}

function ok_(data) { return { ok: true, data: data }; }
function deny_() { return { ok: false, err: 'ROLE', message: '관리 PIN이 필요한 기능입니다.' }; }

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

function boot_(role) {
  var ss = ss_();
  var rs = sheet_(ss, SH_R);
  var students = [];
  var last = rs.getLastRow();
  if (last >= 2) {
    var vals = rs.getRange(2, 1, last - 1, 7).getValues();
    for (var i = 0; i < vals.length; i++) {
      var no = normNo_(vals[i][0]);
      if (!no) continue;
      students.push({
        no: no,
        name: String(vals[i][1] || '').trim(),
        grade: num_(vals[i][2]), klass: num_(vals[i][3]), num: num_(vals[i][4]),
        count: num_(vals[i][5]) || 0,
        lastAt: iso_(vals[i][6])
      });
    }
  }
  var cfg = config_();
  return {
    role: role,
    students: students,
    logs: recentLogs_(LOG_LIMIT),
    exempts: readExempts_(),
    thresholds: thresholds_(cfg),
    school: String(cfg['학교명'] || '').trim(),
    pinSet: !!(String(cfg['기록PIN'] || '').trim() || String(cfg['관리PIN'] || '').trim()),
    sheetUrl: role === 'admin' ? ss.getUrl() : '',
    logLimit: LOG_LIMIT,
    rev: rev_()
  };
}

function recentLogs_(n) {
  var ls = sheet_(ss_(), SH_L);
  var last = ls.getLastRow();
  if (last < 2) return [];
  var from = Math.max(2, last - n + 1);
  var vals = ls.getRange(from, 1, last - from + 1, 9).getValues();
  var out = [];
  for (var i = vals.length - 1; i >= 0; i--) {
    var r = vals[i];
    var no = normNo_(r[2]);
    if (!no) continue;
    out.push({
      id: String(r[0]), at: iso_(r[1]), no: no, name: String(r[3] || ''),
      grade: num_(r[4]), klass: num_(r[5]),
      reason: String(r[6] || ''), memo: String(r[7] || ''), by: String(r[8] || '')
    });
  }
  return out;   // 최신순
}

function history_(no) {
  if (!no) return [];
  var ls = sheet_(ss_(), SH_L);
  var last = ls.getLastRow();
  if (last < 2) return [];
  var vals = ls.getRange(2, 1, last - 1, 9).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    if (normNo_(vals[i][2]) !== no) continue;
    out.push({
      id: String(vals[i][0]), at: iso_(vals[i][1]),
      reason: String(vals[i][6] || ''), memo: String(vals[i][7] || ''), by: String(vals[i][8] || '')
    });
  }
  out.sort(function (a, b) { return a.at < b.at ? 1 : -1; });
  return out;
}

/** 선도위원회 기준 — 설정 시트 값이 없거나 이상하면 기본값 */
function thresholds_(cfg) {
  cfg = cfg || config_();
  var w = parseInt(cfg['경고기준'], 10);
  var r = parseInt(cfg['회부기준'], 10);
  if (!(w > 0)) w = DEF_WARN;
  if (!(r > 0)) r = DEF_REFER;
  if (w > r) w = r;
  return { warn: w, refer: r };
}

function saveThresholds_(p) {
  var w = parseInt(p.warn, 10), r = parseInt(p.refer, 10);
  if (!(w > 0) || !(r > 0)) throw new Error('기준 횟수를 1 이상으로 입력해 주세요.');
  if (w > 99 || r > 99) throw new Error('기준 횟수는 99 이하로 입력해 주세요.');
  if (w > r) throw new Error('경고 시작 횟수가 회부 기준보다 클 수 없습니다.');
  var cs = sheet_(ss_(), SH_C);
  setCfg_(cs, '경고기준', w, '누적 적발이 이 횟수에 이르면 앱에 회부 경고가 뜹니다.');
  setCfg_(cs, '회부기준', r, '누적 적발이 이 횟수에 도달하면 학생선도위원회 회부 대상으로 표시합니다.');
  cache_().remove('cfg');
  bumpRev_();
  return { warn: w, refer: r };
}

/** 설정 시트의 한 줄을 고치거나, 없으면 만든다 */
function setCfg_(cs, key, value, desc) {
  var row = cfgRow_(cs, key);
  if (row) {
    cs.getRange(row, 2).setValue(value);
  } else {
    cs.appendRow([key, value, desc || '']);
    cs.getRange(cs.getLastRow(), 3).setWrap(true);
  }
}

/** 면제 시트는 나중에 추가된 기능이라, 없으면 그때 만든다 */
function ensureExemptSheet_(ss) {
  var sh = ss.getSheetByName(SH_E);
  if (!sh) sh = ss.insertSheet(SH_E);
  if (sh.getLastRow() === 0 || String(sh.getRange(1, 1).getValue()).trim() !== HDR_E[0]) {
    sh.getRange(1, 1, 1, HDR_E.length).setValues([HDR_E]);
    sh.setFrozenRows(1);
  }
  sh.getRange(2, 5, Math.max(sh.getMaxRows() - 1, 1), 2).setNumberFormat('@');  // 시작일·종료일은 글자로
  return sh;
}

function readExempts_() {
  var es;
  try { es = ensureExemptSheet_(ss_()); } catch (e) { return []; }
  var last = es.getLastRow();
  if (last < 2) return [];
  var from = Math.max(2, last - EXEMPT_LIMIT + 1);
  var vals = es.getRange(from, 1, last - from + 1, HDR_E.length).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var no = normNo_(vals[i][1]);
    if (!no) continue;
    out.push({
      id: String(vals[i][0]), no: no, name: String(vals[i][2] || ''),
      reason: String(vals[i][3] || ''), from: dstr_(vals[i][4]), to: dstr_(vals[i][5]),
      by: String(vals[i][6] || '')
    });
  }
  out.sort(function (a, b) { return a.from < b.from ? 1 : (a.from > b.from ? -1 : 0); });
  return out;
}

function exemptAdd_(p) {
  var no = normNo_(p.no);
  var reason = String(p.reason || '').trim().slice(0, 100);
  var from = dstr_(p.from), to = dstr_(p.to);
  if (!no) throw new Error('학생을 선택해 주세요.');
  if (!reason) throw new Error('면제 사유를 입력해 주세요.');
  if (!from || !to) throw new Error('시작일과 종료일을 모두 선택해 주세요.');
  if (from > to) throw new Error('종료일이 시작일보다 빠릅니다.');
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('잠시 후 다시 시도해 주세요.');
  try {
    var ss = ss_(), rs = sheet_(ss, SH_R), es = ensureExemptSheet_(ss);
    var row = rosterIndex_(rs)[no];
    if (!row) throw new Error('명단에 없는 학번입니다.');
    var name = String(rs.getRange(row, 2).getValue() || '');
    var id = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    es.appendRow([id, no, name, reason, from, to, String(p.by || '').slice(0, 40), new Date()]);
    bumpRev_();
    return { id: id, no: no, name: name, reason: reason, from: from, to: to };
  } finally {
    lock.releaseLock();
  }
}

function exemptDel_(id) {
  if (!id) throw new Error('삭제할 면제를 찾지 못했습니다.');
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('잠시 후 다시 시도해 주세요.');
  try {
    var es = ensureExemptSheet_(ss_());
    var last = es.getLastRow();
    if (last >= 2) {
      var vals = es.getRange(2, 1, last - 1, 1).getValues();
      for (var i = vals.length - 1; i >= 0; i--) {
        if (String(vals[i][0]) === id) { es.deleteRow(i + 2); break; }
      }
    }
    bumpRev_();
    return { id: id };
  } finally {
    lock.releaseLock();
  }
}

/* ─────────────────────────── 쓰기 ─────────────────────────── */

function addRecord_(p) {
  var no = normNo_(p.no);
  if (!no) throw new Error('학번이 올바르지 않습니다.');
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('다른 기록을 처리 중입니다. 잠시 후 다시 눌러주세요.');
  try {
    var ss = ss_(), rs = sheet_(ss, SH_R), ls = sheet_(ss, SH_L);
    var row = rosterIndex_(rs)[no];
    if (!row) throw new Error('명단에 없는 학번입니다. 명단을 먼저 등록해 주세요.');
    var info = rs.getRange(row, 1, 1, 7).getValues()[0];
    var now = new Date();
    // 하루 한 번 원칙 — 오늘 이미 기록된 학생은 막는다
    if (sameDay_(info[6], now)) {
      throw { dup: true, at: iso_(info[6]), count: num_(info[5]) || 0,
              message: '오늘 이미 체크된 학생입니다.' };
    }
    var id = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    ls.appendRow([id, now, no, info[1], info[2], info[3],
      REASON, '', String(p.by || '').slice(0, 40)]);
    var count = (num_(info[5]) || 0) + 1;
    rs.getRange(row, 6, 1, 2).setValues([[count, now]]);
    bumpRev_();
    return { id: id, no: no, count: count, at: iso_(now), rev: rev_() };
  } finally {
    lock.releaseLock();
  }
}

function undoRecord_(p) {
  var no = normNo_(p.no), id = String(p.id || '');
  if (!no || !id) throw new Error('취소할 기록을 찾지 못했습니다.');
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('다른 기록을 처리 중입니다. 잠시 후 다시 시도해 주세요.');
  try {
    var ss = ss_(), ls = sheet_(ss, SH_L);
    var row = findLogRow_(ls, id);
    if (row) ls.deleteRow(row);
    var re = recalcOne_(ss, no);
    bumpRev_();
    return { no: no, count: re.count, lastAt: re.lastAt, removed: !!row, rev: rev_() };
  } finally {
    lock.releaseLock();
  }
}

function putStudent_(p) {
  var no = normNo_(p.no), name = String(p.name || '').trim();
  var oldNo = normNo_(p.oldNo || p.no);
  if (!no || !name) throw new Error('학번과 이름을 모두 입력해 주세요.');
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('잠시 후 다시 시도해 주세요.');
  try {
    var ss = ss_(), rs = sheet_(ss, SH_R), ls = sheet_(ss, SH_L);
    var idx = rosterIndex_(rs);
    var d = derive_(no);
    if (oldNo && oldNo !== no && idx[oldNo]) {
      // 학번 변경: 기존 행을 옮기고 기록의 학번·이름도 갱신
      var old = rs.getRange(idx[oldNo], 1, 1, 7).getValues()[0];
      rs.deleteRow(idx[oldNo]);
      idx = rosterIndex_(rs);
      var keep = [no, name, d.grade, d.klass, d.num, num_(old[5]) || 0, old[6] || ''];
      if (idx[no]) rs.getRange(idx[no], 1, 1, 7).setValues([keep]);
      else rs.appendRow(keep);
      renameInLog_(ls, oldNo, no, name, d);
    } else if (idx[no]) {
      rs.getRange(idx[no], 2, 1, 4).setValues([[name, d.grade, d.klass, d.num]]);
      renameInLog_(ls, no, no, name, d);
    } else {
      rs.appendRow([no, name, d.grade, d.klass, d.num, 0, '']);
    }
    sortRoster_(rs);
    bumpRev_();
    return { no: no };
  } finally {
    lock.releaseLock();
  }
}

function delStudent_(no) {
  if (!no) throw new Error('학번이 올바르지 않습니다.');
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('잠시 후 다시 시도해 주세요.');
  try {
    var ss = ss_(), rs = sheet_(ss, SH_R), ls = sheet_(ss, SH_L);
    var row = rosterIndex_(rs)[no];
    if (row) rs.deleteRow(row);
    // 해당 학생의 기록도 함께 삭제 (아래에서 위로 지워야 행 번호가 밀리지 않음)
    var last = ls.getLastRow();
    if (last >= 2) {
      var vals = ls.getRange(2, 3, last - 1, 1).getValues();
      for (var i = vals.length - 1; i >= 0; i--) {
        if (normNo_(vals[i][0]) === no) ls.deleteRow(i + 2);
      }
    }
    bumpRev_();
    return { no: no };
  } finally {
    lock.releaseLock();
  }
}

function bulkRoster_(list) {
  if (!list.length) throw new Error('등록할 학생이 없습니다.');
  if (list.length > 5000) throw new Error('한 번에 5,000명까지 등록할 수 있습니다.');
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) throw new Error('잠시 후 다시 시도해 주세요.');
  try {
    var rs = sheet_(ss_(), SH_R);
    var last = rs.getLastRow();
    var cur = {}, order = [];
    if (last >= 2) {
      var vals = rs.getRange(2, 1, last - 1, 7).getValues();
      for (var i = 0; i < vals.length; i++) {
        var n = normNo_(vals[i][0]);
        if (!n || cur[n]) continue;
        cur[n] = vals[i]; order.push(n);
      }
    }
    var added = 0, updated = 0;
    for (var j = 0; j < list.length; j++) {
      var s = list[j], no = normNo_(s.no), name = String(s.name || '').trim();
      if (!no || !name) continue;
      var d = derive_(no);
      if (cur[no]) { cur[no] = [no, name, d.grade, d.klass, d.num, num_(cur[no][5]) || 0, cur[no][6] || '']; updated++; }
      else { cur[no] = [no, name, d.grade, d.klass, d.num, 0, '']; order.push(no); added++; }
    }
    var keys = order.slice().sort(cmpNo_);
    var rows = [];
    for (var k = 0; k < keys.length; k++) rows.push(cur[keys[k]]);
    if (last >= 2) rs.getRange(2, 1, last - 1, 7).clearContent();
    if (rows.length) rs.getRange(2, 1, rows.length, 7).setValues(rows);
    bumpRev_();
    return { total: rows.length, added: added, updated: updated };
  } finally {
    lock.releaseLock();
  }
}

function resetCounts_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) throw new Error('잠시 후 다시 시도해 주세요.');
  try {
    var ss = ss_(), rs = sheet_(ss, SH_R), ls = sheet_(ss, SH_L);
    var lastL = ls.getLastRow();
    if (lastL >= 2) ls.getRange(2, 1, lastL - 1, HDR_L.length).clearContent();
    var lastR = rs.getLastRow();
    if (lastR >= 2) {
      var zeros = [];
      for (var i = 0; i < lastR - 1; i++) zeros.push([0, '']);
      rs.getRange(2, 6, lastR - 1, 2).setValues(zeros);
    }
    bumpRev_();
    return { cleared: Math.max(0, lastL - 1) };
  } finally {
    lock.releaseLock();
  }
}

function wipeRoster_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) throw new Error('잠시 후 다시 시도해 주세요.');
  try {
    var ss = ss_(), rs = sheet_(ss, SH_R), ls = sheet_(ss, SH_L);
    if (rs.getLastRow() >= 2) rs.getRange(2, 1, rs.getLastRow() - 1, HDR_R.length).clearContent();
    if (ls.getLastRow() >= 2) ls.getRange(2, 1, ls.getLastRow() - 1, HDR_L.length).clearContent();
    bumpRev_();
    return {};
  } finally {
    lock.releaseLock();
  }
}

/* ─────────────────────────── 도우미 ─────────────────────────── */

function ss_() {
  var id = PropertiesService.getScriptProperties().getProperty('ssid');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) {}
  }
  var active = SpreadsheetApp.getActive();
  if (!active) throw new Error('스프레드시트를 찾지 못했습니다. [교복 지도] → [① 초기 설정]을 먼저 실행해 주세요.');
  return active;
}

function sheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('“' + name + '” 시트가 없습니다. [교복 지도] → [① 초기 설정]을 실행해 주세요.');
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
        if (k) out[k] = vals[i][1];
      }
    }
  } catch (e) { /* 초기 설정 전 */ }
  cache_().put('cfg', JSON.stringify(out), CFG_TTL);
  return out;
}

function cfgRow_(cs, key) {
  var last = cs.getLastRow();
  if (last < 2) return 0;
  var vals = cs.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) if (String(vals[i][0] || '').trim() === key) return i + 2;
  return 0;
}

function rev_() {
  return Number(PropertiesService.getScriptProperties().getProperty('rev') || 0);
}
function bumpRev_() {
  PropertiesService.getScriptProperties().setProperty('rev', String(Date.now()));
}

function rosterIndex_(rs) {
  var map = {};
  var last = rs.getLastRow();
  if (last < 2) return map;
  var vals = rs.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    var no = normNo_(vals[i][0]);
    if (no && !map[no]) map[no] = i + 2;
  }
  return map;
}

function findLogRow_(ls, id) {
  var last = ls.getLastRow();
  if (last < 2) return 0;
  var vals = ls.getRange(2, 1, last - 1, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) if (String(vals[i][0]) === id) return i + 2;
  return 0;
}

function recalcOne_(ss, no) {
  var ls = sheet_(ss, SH_L), rs = sheet_(ss, SH_R);
  var count = 0, lastAt = null;
  var last = ls.getLastRow();
  if (last >= 2) {
    var vals = ls.getRange(2, 2, last - 1, 2).getValues();   // 일시, 학번
    for (var i = 0; i < vals.length; i++) {
      if (normNo_(vals[i][1]) !== no) continue;
      count++;
      var d = vals[i][0];
      if (d instanceof Date && (!lastAt || d > lastAt)) lastAt = d;
    }
  }
  var row = rosterIndex_(rs)[no];
  if (row) rs.getRange(row, 6, 1, 2).setValues([[count, lastAt || '']]);
  return { count: count, lastAt: iso_(lastAt) };
}

function renameInLog_(ls, oldNo, newNo, name, d) {
  var last = ls.getLastRow();
  if (last < 2) return;
  var vals = ls.getRange(2, 3, last - 1, 4).getValues();   // 학번, 이름, 학년, 반
  var changed = false;
  for (var i = 0; i < vals.length; i++) {
    if (normNo_(vals[i][0]) !== oldNo) continue;
    vals[i] = [newNo, name, d.grade, d.klass];
    changed = true;
  }
  if (changed) ls.getRange(2, 3, last - 1, 4).setValues(vals);
}

function sortRoster_(rs) {
  var last = rs.getLastRow();
  if (last < 3) return;
  var vals = rs.getRange(2, 1, last - 1, 7).getValues().filter(function (r) { return normNo_(r[0]); });
  vals.sort(function (a, b) { return cmpNo_(normNo_(a[0]), normNo_(b[0])); });
  rs.getRange(2, 1, last - 1, 7).clearContent();
  if (vals.length) rs.getRange(2, 1, vals.length, 7).setValues(vals);
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

/** 무엇이 들어와도 yyyy-MM-dd 문자열로 */
function dstr_(v) {
  if (!v && v !== 0) return '';
  if (v instanceof Date) {
    return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  }
  var t = String(v).trim();
  var m = /^(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})/.exec(t);
  if (m) return m[1] + '-' + p2_(m[2]) + '-' + p2_(m[3]);
  return t.slice(0, 10);
}
function p2_(n) { n = String(n); return n.length < 2 ? '0' + n : n; }

/** 스프레드시트 시간대 기준으로 같은 날인지 */
function sameDay_(a, b) {
  if (!a || !b) return false;
  var da = (a instanceof Date) ? a : new Date(a);
  var db = (b instanceof Date) ? b : new Date(b);
  if (isNaN(da.getTime()) || isNaN(db.getTime())) return false;
  var tz = tz_();
  return Utilities.formatDate(da, tz, 'yyyy-MM-dd') === Utilities.formatDate(db, tz, 'yyyy-MM-dd');
}

function tz_() {
  try { return ss_().getSpreadsheetTimeZone() || 'Asia/Seoul'; } catch (e) { return 'Asia/Seoul'; }
}
