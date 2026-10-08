/**
 * 한백건축배 연말파티 접수 서버 (Google Apps Script)
 *
 * 설치: 접수 데이터를 담을 구글시트를 새로 만든 뒤
 *   확장 프로그램 > Apps Script 에 이 코드를 통째로 붙여넣기
 *   → 아래 ADMIN_ID, ADMIN_PASSWORD, SALT 를 바꾸고 저장
 *   → 배포 > 새 배포 > 유형 "웹 앱" > 실행: 나 / 액세스: 모든 사용자 > 배포
 *   → 나오는 "웹 앱 URL"을 index.html 맨 위 API_URL 에 넣기
 *
 * 보안 구조
 *  - 관리자 비밀번호는 이 서버 코드에만 있고, 웹사이트 소스에는 없음
 *  - 일반 방문자에게는 "접수 인원 숫자"만 내려감 (이름 등은 절대 안 내려감)
 *  - 접수자는 '이름 + 수정 비밀번호'로 본인 것만 조회/수정 (비밀번호는 해시로만 저장)
 *  - 비밀번호 5회 연속 오류 시 10분 잠금
 */

const ADMIN_ID = '여기에-관리자-아이디';              // ← 반드시 변경
const ADMIN_PASSWORD = '여기에-관리자-비밀번호를-적으세요'; // ← 반드시 변경
const SALT = '아무-긴-문자열로-바꾸세요-예-snow-2026';     // ← 반드시 변경 (변경 후엔 기존 접수자 비번이 무효화되니 오픈 전에만)
const SHEET_NAME = '접수';
const MAX_ENTRIES = 500; // 안전장치

const HEADERS = ['ID', '이름', '좋아하는색상', '색상코드', '각오', '바라는점', '비번해시', '접수일시', '수정일시', '상태'];
const COL = { id: 0, name: 1, colorName: 2, colorHex: 3, resolve: 4, wish: 5, hash: 6, created: 7, updated: 8, status: 9 };
const STATUSES = ['접수', '확정', '취소'];

/* ───────── 진입점 ───────── */

function doGet() {
  return json_({ ok: true, msg: 'hanbaek-party api' });
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const a = req.action;
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      switch (a) {
        case 'stats':          return json_(stats_());
        case 'submit':         return json_(submit_(req));
        case 'lookup':         return json_(lookup_(req));
        case 'update':         return json_(update_(req));
        case 'cancel':         return json_(cancel_(req));
        case 'adminLogin':     return json_(adminLogin_(req));
        case 'adminList':      return json_(adminList_(req));
        case 'adminSetStatus': return json_(adminSetStatus_(req));
        case 'adminDelete':    return json_(adminDelete_(req));
        case 'adminSetOpen':   return json_(adminSetOpen_(req));
        default:               return json_({ ok: false, error: '알 수 없는 요청이에요.' });
      }
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, error: '서버 오류: ' + err });
  }
}

/* ───────── 일반 방문자 ───────── */

function stats_() {
  const rows = rows_();
  const count = rows.filter(r => r[COL.status] !== '취소').length;
  return { ok: true, count: count, open: isOpen_() };
}

function submit_(req) {
  if (!isOpen_()) return { ok: false, error: '접수가 마감되었어요.' };
  const v = validate_(req, true);
  if (v.error) return { ok: false, error: v.error };

  const rows = rows_();
  if (rows.length >= MAX_ENTRIES) return { ok: false, error: '접수 한도를 넘었어요. 관리자에게 문의해주세요.' };

  const id = Utilities.getUuid().slice(0, 8);
  const hash = hash_(id, req.pin);
  // 같은 이름+같은 비밀번호 중복 방지
  const dup = rows.some(r => sameName_(r[COL.name], v.name) && r[COL.status] !== '취소' && sameHashAcrossId_(r, req.pin));
  if (dup) return { ok: false, error: '같은 이름·비밀번호로 이미 접수되어 있어요. 아래 "내 접수 수정"을 이용해주세요.' };

  const now = now_();
  sheet_().appendRow([id, v.name, v.colorName, v.colorHex, v.resolve, v.wish, hash, now, now, '접수']);
  return { ok: true, count: stats_().count };
}

function lookup_(req) {
  const found = authUser_(req.name, req.pin);
  if (found.error) return { ok: false, error: found.error };
  const r = found.row;
  return { ok: true, record: pub_(r) };
}

function update_(req) {
  if (!isOpen_()) return { ok: false, error: '접수가 마감되어 수정할 수 없어요.' };
  const found = authUserById_(req.id, req.pin);
  if (found.error) return { ok: false, error: found.error };
  const v = validate_(req, false);
  if (v.error) return { ok: false, error: v.error };
  const sh = sheet_();
  const i = found.index; // 시트 행 번호
  const status = found.row[COL.status] === '취소' ? '접수' : found.row[COL.status];
  sh.getRange(i, COL.name + 1, 1, 5).setValues([[v.name, v.colorName, v.colorHex, v.resolve, v.wish]]);
  sh.getRange(i, COL.updated + 1).setValue(now_());
  sh.getRange(i, COL.status + 1).setValue(status);
  return { ok: true, count: stats_().count };
}

function cancel_(req) {
  const found = authUserById_(req.id, req.pin);
  if (found.error) return { ok: false, error: found.error };
  sheet_().getRange(found.index, COL.status + 1).setValue('취소');
  sheet_().getRange(found.index, COL.updated + 1).setValue(now_());
  return { ok: true, count: stats_().count };
}

/* ───────── 관리자 ───────── */

function adminLogin_(req) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('admin_fail') || 0);
  if (fails >= 5) return { ok: false, error: '로그인 시도가 너무 많아요. 10분 뒤에 다시 시도해주세요.' };
  if (!req.password || String(req.id || '').trim() !== ADMIN_ID || String(req.password) !== ADMIN_PASSWORD) {
    cache.put('admin_fail', String(fails + 1), 600);
    return { ok: false, error: '아이디 또는 비밀번호가 맞지 않아요.' };
  }
  cache.remove('admin_fail');
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('tok_' + token, '1', 6 * 60 * 60); // 6시간
  return { ok: true, token: token };
}

function requireAdmin_(token) {
  if (!token || !CacheService.getScriptCache().get('tok_' + token)) return false;
  return true;
}

function adminList_(req) {
  if (!requireAdmin_(req.token)) return { ok: false, auth: false, error: '로그인이 만료되었어요.' };
  const rows = rows_().map(r => ({
    id: r[COL.id], name: r[COL.name], colorName: r[COL.colorName], colorHex: r[COL.colorHex],
    resolve: r[COL.resolve], wish: r[COL.wish], created: r[COL.created], updated: r[COL.updated], status: r[COL.status]
  }));
  return { ok: true, rows: rows, open: isOpen_() };
}

function adminSetStatus_(req) {
  if (!requireAdmin_(req.token)) return { ok: false, auth: false, error: '로그인이 만료되었어요.' };
  if (STATUSES.indexOf(req.status) < 0) return { ok: false, error: '잘못된 상태값이에요.' };
  const i = indexById_(req.id);
  if (!i) return { ok: false, error: '대상을 찾지 못했어요.' };
  sheet_().getRange(i, COL.status + 1).setValue(req.status);
  return { ok: true };
}

function adminDelete_(req) {
  if (!requireAdmin_(req.token)) return { ok: false, auth: false, error: '로그인이 만료되었어요.' };
  const i = indexById_(req.id);
  if (!i) return { ok: false, error: '대상을 찾지 못했어요.' };
  sheet_().deleteRow(i);
  return { ok: true };
}

function adminSetOpen_(req) {
  if (!requireAdmin_(req.token)) return { ok: false, auth: false, error: '로그인이 만료되었어요.' };
  PropertiesService.getScriptProperties().setProperty('OPEN', req.open ? '1' : '0');
  return { ok: true, open: isOpen_() };
}

/* ───────── 내부 유틸 ───────── */

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function now_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'); }
function isOpen_() { return PropertiesService.getScriptProperties().getProperty('OPEN') !== '0'; }

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.getRange(1, 1, sh.getMaxRows(), HEADERS.length).setNumberFormat('@'); // 전부 텍스트로 (날짜 자동변환·수식 방지)
    sh.setFrozenRows(1);
    sh.hideColumns(COL.hash + 1); // 비번해시 열 숨김
  }
  return sh;
}

function rows_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
}

function indexById_(id) {
  const rows = rows_();
  for (let k = 0; k < rows.length; k++) if (String(rows[k][COL.id]) === String(id)) return k + 2;
  return 0;
}

function hash_(id, pin) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, id + ':' + String(pin) + ':' + SALT);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function sameHashAcrossId_(row, pin) { return row[COL.hash] === hash_(row[COL.id], pin); }
function norm_(s) { return String(s || '').replace(/\s+/g, '').toLowerCase(); }
function sameName_(a, b) { return norm_(a) === norm_(b); }

function clean_(s, max) {
  let t = String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim();
  if (t.length > max) t = t.slice(0, max);
  if (/^[=+\-@]/.test(t)) t = "'" + t; // 수식 주입 방지
  return t;
}

function validate_(req, needPin) {
  const name = clean_(req.name, 20);
  const colorName = clean_(req.colorName, 20);
  const colorHex = /^#[0-9a-fA-F]{6}$/.test(req.colorHex || '') ? req.colorHex : '';
  const resolve = clean_(req.resolve, 200);
  const wish = clean_(req.wish, 300);
  if (!name) return { error: '이름을 입력해주세요.' };
  if (!colorName) return { error: '좋아하는 색상을 골라주세요.' };
  if (!resolve) return { error: '각오를 한 줄 적어주세요.' };
  if (needPin && !/^.{4,20}$/.test(String(req.pin || ''))) return { error: '수정 비밀번호는 4자 이상으로 정해주세요.' };
  return { name: name, colorName: colorName, colorHex: colorHex, resolve: resolve, wish: wish };
}

function pub_(r) {
  return { id: r[COL.id], name: r[COL.name], colorName: r[COL.colorName], colorHex: r[COL.colorHex],
           resolve: r[COL.resolve], wish: r[COL.wish], status: r[COL.status] };
}

function tooMany_(key) { return Number(CacheService.getScriptCache().get(key) || 0) >= 5; }
function addFail_(key) {
  const c = CacheService.getScriptCache();
  c.put(key, String(Number(c.get(key) || 0) + 1), 600);
}

// 이름+비번으로 본인 찾기
function authUser_(name, pin) {
  const key = 'f_' + norm_(name);
  if (tooMany_(key)) return { error: '시도가 너무 많아요. 10분 뒤에 다시 해주세요.' };
  const rows = rows_();
  for (let k = 0; k < rows.length; k++) {
    const r = rows[k];
    if (sameName_(r[COL.name], name) && sameHashAcrossId_(r, pin)) {
      CacheService.getScriptCache().remove(key);
      return { row: r, index: k + 2 };
    }
  }
  addFail_(key);
  return { error: '이름 또는 비밀번호가 맞지 않아요.' };
}

// 수정/취소 시 id+비번 재검증
function authUserById_(id, pin) {
  const key = 'f_id_' + id;
  if (tooMany_(key)) return { error: '시도가 너무 많아요. 10분 뒤에 다시 해주세요.' };
  const i = indexById_(id);
  if (!i) return { error: '접수 내역을 찾지 못했어요.' };
  const r = sheet_().getRange(i, 1, 1, HEADERS.length).getValues()[0];
  if (!sameHashAcrossId_(r, pin)) { addFail_(key); return { error: '비밀번호가 맞지 않아요.' }; }
  return { row: r, index: i };
}
