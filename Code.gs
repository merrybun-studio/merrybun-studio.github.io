/**
 * 메리번 예식 체크리스트 → 구글 시트 저장
 * 시트 메뉴 [확장 프로그램 > Apps Script]에 이 코드를 통째로 붙여넣고 저장하세요.
 *
 * - doGet은 없어요. 모든 요청은 POST로만 받아요.
 * - 고객은 처음 제출할 때 만들어진 "수정 키"(비밀 링크)로만 자기 답변 한 건을 불러오고 고칠 수 있어요.
 * - 시트의 "확인" 칸을 체크하면 그 건은 잠겨서 더 이상 수정할 수 없어요.
 */
const SHEET_NAME = '체크리스트';
const TIME_FMT = 'yyyy. m. d am/pm h:mm:ss';
const FIXED = ['제출 시각', '확인', '개인정보 동의', '최종 수정 시각'];
const KEY_COL = '수정 키';
const MAX_LEN = 3000;      // 답변 하나의 최대 글자 수
const MAX_KEYS = 60;       // 문항 수 상한
const MIN_ELAPSED = 3000;  // 페이지를 연 뒤 이 시간(ms)보다 빨리 온 제출은 버려요
// iframe 보조 경로로 결과를 보내도 되는 페이지 주소
const ALLOWED_ORIGINS = ['https://kwy-del.github.io', 'https://merrybun-studio.github.io', 'http://localhost:8000'];

function doPost(e) {
  let data;
  try {
    // fetch는 본문(JSON)으로, 보조 경로(숨은 iframe 폼)는 payload 칸으로 보내요
    const raw = (e.parameter && e.parameter.payload) || e.postData.contents;
    data = JSON.parse(raw);
  } catch (err) {
    return reply({ ok: false, error: 'bad_request' });
  }
  let res;
  try {
    res = data.action === 'load' ? load(data) : save(data);
  } catch (err) {
    console.error(err);
    res = { ok: false, error: 'server' };
  }
  return data.via === 'frame' ? replyFrame(res, data.rid, data.origin) : reply(res);
}

function save(data) {
  // 스팸: 허니팟 칸이 채워져 있거나 너무 빨리 제출되면 저장하지 않아요
  if (data.hp) return { ok: true };
  if (!(Number(data.elapsed) >= MIN_ELAPSED)) return { ok: false, error: 'too_fast' };

  if (data.consent !== true) return { ok: false, error: 'no_consent' };
  const key = cleanKey(data.key);
  if (!key) return { ok: false, error: 'bad_request' };
  const answers = cleanAnswers(data.answers);
  if (!answers) return { ok: false, error: 'bad_request' };
  const keys = Object.keys(answers);
  if (!answers[keys[0]] || !answers[keys[1]]) return { ok: false, error: 'missing_required' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: 'busy' };
  try {
    const sheet = getSheet();

    // 첫 제출 때 헤더를 만들고, 나중에 문항이 늘면 오른쪽에 열을 추가해요 (수정 키는 맨 끝)
    let header = readHeader(sheet);
    const want = FIXED.concat(keys).concat([KEY_COL]);
    const missing = want.filter(function (h) { return header.indexOf(h) === -1; });
    if (missing.length) {
      sheet.getRange(1, header.length + 1, 1, missing.length).setValues([missing]);
      header = header.concat(missing);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
    }

    const now = new Date();
    const r = findRow(sheet, header, key);
    const old = r ? sheet.getRange(r, 1, 1, header.length).getValues()[0] : null;
    if (old && isLocked(old[header.indexOf('확인')])) return { ok: false, error: 'locked' };

    const row = header.map(function (h, i) {
      if (h === '제출 시각') return old ? old[i] : now;
      if (h === '확인') return old ? old[i] : false;
      if (h === '개인정보 동의') return '동의함';
      if (h === '최종 수정 시각') return old ? now : '';
      if (h === KEY_COL) return key;
      if (h in answers) return safe(answers[h]);
      return old ? old[i] : '';  // 지금 페이지에 없는 예전 문항 칸은 그대로 둬요
    });

    const target = r || sheet.getLastRow() + 1;
    sheet.getRange(target, 1, 1, header.length).setValues([row]);
    sheet.getRange(target, header.indexOf('확인') + 1).insertCheckboxes();
    // 날짜만 보이지 않게 시각까지 나오는 서식으로 맞춰요
    ['제출 시각', '최종 수정 시각'].forEach(function (h) {
      sheet.getRange(target, header.indexOf(h) + 1).setNumberFormat(TIME_FMT);
    });
    return { ok: true, updated: !!old, submittedAt: iso(row[header.indexOf('제출 시각')]), updatedAt: old ? iso(now) : '' };
  } finally {
    lock.releaseLock();
  }
}

// 수정 키로 그 한 건만 돌려줘요
function load(data) {
  const key = cleanKey(data.key);
  if (!key) return { ok: false, error: 'bad_request' };
  const sheet = getSheet();
  const header = readHeader(sheet);
  const r = findRow(sheet, header, key);
  if (!r) return { ok: false, error: 'not_found' };
  const values = sheet.getRange(r, 1, 1, header.length).getValues()[0];
  const shown = sheet.getRange(r, 1, 1, header.length).getDisplayValues()[0];
  const answers = {};
  header.forEach(function (h, i) { if (/^\d{2}\. /.test(h)) answers[h] = shown[i]; });
  return {
    ok: true,
    locked: isLocked(values[header.indexOf('확인')]),
    submittedAt: iso(values[header.indexOf('제출 시각')]),
    updatedAt: iso(values[header.indexOf('최종 수정 시각')]),
    answers: answers
  };
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function readHeader(sheet) {
  return sheet.getLastRow() === 0 ? [] : sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function findRow(sheet, header, key) {
  const c = header.indexOf(KEY_COL);
  if (c === -1 || sheet.getLastRow() < 2) return 0;
  const hit = sheet.getRange(2, c + 1, sheet.getLastRow() - 1, 1).createTextFinder(key).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

// 체크박스가 체크됐거나, 확인 칸에 글자가 있으면 잠긴 걸로 봐요
function isLocked(v) {
  return v === true || (typeof v === 'string' && v.trim() !== '' && v.trim().toUpperCase() !== 'FALSE');
}

function iso(v) {
  return v instanceof Date ? v.toISOString() : '';
}

// "01. 문항" 형식의 키만 받고, 값은 문자열로 자르기
function cleanAnswers(a) {
  if (!a || typeof a !== 'object') return null;
  const out = {};
  const keys = Object.keys(a).filter(function (k) { return /^\d{2}\. .{1,100}$/.test(k); }).slice(0, MAX_KEYS);
  if (keys.length < 2) return null;
  keys.forEach(function (k) { out[k] = String(a[k] == null ? '' : a[k]).slice(0, MAX_LEN); });
  return out;
}

function cleanKey(k) {
  return /^[a-f0-9]{32}$/.test(String(k || '')) ? String(k) : '';
}

// = + - @ 로 시작하는 값이 수식으로 실행되지 않게 글자로 저장해요
function safe(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 보조 경로: 숨은 iframe 안에서 결과를 체크리스트 페이지로만 전달해요
function replyFrame(obj, rid, origin) {
  if (ALLOWED_ORIGINS.indexOf(origin) === -1) return HtmlService.createHtmlOutput('');
  const msg = JSON.stringify({ mb: String(rid || '').slice(0, 64), res: obj }).replace(/</g, '\\u003c');
  return HtmlService.createHtmlOutput('<script>window.top.postMessage(' + msg + ', ' + JSON.stringify(origin) + ');</script>')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Apps Script 편집기에서 이 함수를 선택해 ▶실행하면, 웹 없이도 저장 → 수정 → 불러오기를 시험할 수 있어요.
// 시험이 끝나면 시트에 생긴 테스트 줄은 지워주세요.
function testPost() {
  const key = Utilities.getUuid().replace(/-/g, '');
  const send = function (o) { return JSON.parse(doPost({ postData: { contents: JSON.stringify(o) } }).getContent()); };
  const base = { consent: true, hp: '', elapsed: 9999, key: key };
  console.log('저장', send(Object.assign({}, base, { answers: { '01. 계약자 이름 / 연락처': '테스트 / 010-0000-0000', '02. 예식 일시 · 웨딩홀': '=테스트 수식 방지 확인' } })));
  console.log('수정', send(Object.assign({}, base, { answers: { '01. 계약자 이름 / 연락처': '테스트(수정) / 010-0000-0000', '02. 예식 일시 · 웨딩홀': '=테스트 수식 방지 확인' } })));
  console.log('불러오기', send({ action: 'load', key: key }));
}
