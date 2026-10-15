/**
 * 메리번 예식 체크리스트 → 구글 시트 저장
 * 시트 메뉴 [확장 프로그램 > Apps Script]에 이 코드를 통째로 붙여넣고 저장하세요.
 * 조회용 엔드포인트(doGet)는 일부러 만들지 않았어요. 웹앱으로는 저장만 할 수 있어요.
 */
const SHEET_NAME = '체크리스트';
const FIXED = ['제출 시각', '확인', '개인정보 동의'];
const MAX_LEN = 3000;      // 답변 하나의 최대 글자 수
const MAX_KEYS = 60;       // 문항 수 상한
const MIN_ELAPSED = 3000;  // 페이지를 연 뒤 이 시간(ms)보다 빨리 온 제출은 버려요

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return reply({ ok: false, error: 'bad_request' });
  }

  // 스팸: 허니팟 칸이 채워져 있거나 너무 빨리 제출되면 저장하지 않아요
  if (data.hp) return reply({ ok: true });
  if (!(Number(data.elapsed) >= MIN_ELAPSED)) return reply({ ok: false, error: 'too_fast' });

  if (data.consent !== true) return reply({ ok: false, error: 'no_consent' });
  const answers = cleanAnswers(data.answers);
  if (!answers) return reply({ ok: false, error: 'bad_request' });
  const keys = Object.keys(answers);
  if (!answers[keys[0]] || !answers[keys[1]]) return reply({ ok: false, error: 'missing_required' });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return reply({ ok: false, error: 'busy' });
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    // 첫 제출 때 헤더를 만들고, 나중에 문항이 늘면 오른쪽에 열을 추가해요
    let header = sheet.getLastRow() === 0 ? [] : sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const missing = FIXED.concat(keys).filter(function (h) { return header.indexOf(h) === -1; });
    if (missing.length) {
      sheet.getRange(1, header.length + 1, 1, missing.length).setValues([missing]);
      header = header.concat(missing);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
    }

    const now = new Date();
    const row = header.map(function (h) {
      if (h === '제출 시각') return now;
      if (h === '확인') return '';
      if (h === '개인정보 동의') return '동의함';
      return safe(answers[h] || '');
    });
    sheet.appendRow(row);
    return reply({ ok: true });
  } catch (err) {
    console.error(err);
    return reply({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
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

// = + - @ 로 시작하는 값이 수식으로 실행되지 않게 글자로 저장해요
function safe(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Apps Script 편집기에서 이 함수를 선택해 ▶실행하면, 웹 없이도 시트 저장을 시험할 수 있어요.
function testPost() {
  const res = doPost({ postData: { contents: JSON.stringify({
    consent: true, hp: '', elapsed: 9999,
    answers: { '01. 계약자 이름 / 연락처': '테스트 / 010-0000-0000', '02. 예식 일시 · 웨딩홀': '=테스트 수식 방지 확인' }
  }) } });
  console.log(res.getContent());
}
