/**
 * 인사조직 진단 설문 — 응답 수신용 Apps Script
 *
 * 설치: 스프레드시트 「인사조직 진단 설문 응답」 → 확장 프로그램 → Apps Script
 *       (메뉴가 안 보이면 https://script.google.com → 새 프로젝트 에서 해도 됩니다)
 *       이 코드를 전부 붙여 넣고 저장 → 배포 → 새 배포 → 유형: 웹 앱
 *       (실행 사용자: 나 / 액세스 권한: 모든 사용자) → 배포 → 웹 앱 URL 복사
 *
 * 하는 일: 설문 페이지에서 제출된 응답을 「응답」 시트에 한 줄씩 추가하고,
 *          제출될 때마다 스프레드시트 소유자에게 이메일로 알려 줍니다.
 */

// ===== 설정 =====
const SPREADSHEET_ID = '1dfiPqZFOdF8_3eWhypKB1IXlGwBimgwUvpCp5BTI0Js'; // 응답 스프레드시트
const SHEET_NAME = '응답';      // 응답이 쌓일 시트 이름
const SEND_EMAIL = true;        // 제출 알림 메일 보내기 (false로 바꾸면 끔)
const NOTIFY_TO  = '';          // 받을 주소. 비워 두면 스프레드시트 소유자(배포한 계정)에게 보냄
// ================

const NQ = 40;
const POS = {1:'임원', 2:'실장·팀장급', 3:'차·과장', 4:'대리·사원'};
const JOB = {1:'기획', 2:'총무인사', 3:'영업', 4:'생산', 5:'기타'};
const GEN = {M:'남', F:'여'};
const AREAS = [
  ['경영방침 및 계획', 1, 5], ['경영진 및 상사의 리더십', 6, 10],
  ['조직운영 및 조직분위기', 11, 15], ['커뮤니케이션 및 업무협조', 16, 20],
  ['급여 및 복지수준', 21, 25], ['인사 및 교육체계', 26, 30],
  ['근무의욕 및 직무만족도', 31, 36], ['기업문화', 37, 40],
];
const HEADER = ['응답ID','제출일시','직위','직종','직종(기타)','성별','근속(개월)']
  .concat(Array.from({length: NQ}, (_, i) => 'Q' + (i + 1)), ['기타의견']);

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // --- 입력 검증 ---
    const id = String(d.id || '');
    if (!/^[a-z0-9]{6,24}$/.test(id)) return out_({ok: false, error: 'bad id'});
    if (!Array.isArray(d.a) || d.a.length !== NQ) return out_({ok: false, error: 'bad answers'});
    const a = d.a.map(Number);
    if (a.some(v => !(v >= 1 && v <= 5 && Math.floor(v) === v))) return out_({ok: false, error: 'bad answers'});

    const sh = getSheet_();
    // 같은 응답이 두 번 들어오지 않도록 응답ID 확인
    const last = sh.getLastRow();
    if (last > 1) {
      const ids = sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]));
      if (ids.indexOf(id) !== -1) return out_({ok: true, duplicate: true});
    }

    const ten = (typeof d.ten === 'number' && d.ten >= 0 && d.ten < 1000) ? Math.round(d.ten) : '';
    const row = [
      id,
      new Date(),
      POS[d.p] || '',
      JOB[d.j] || '',
      safeText_(d.je, 30),
      GEN[d.g] || '',
      ten,
    ].concat(a, [safeText_(d.c, 2000)]);
    sh.appendRow(row);

    if (SEND_EMAIL) notify_(row, a, sh);
    return out_({ok: true});
  } catch (err) {
    return out_({ok: false, error: String(err)});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// 브라우저에서 웹 앱 주소를 열었을 때 작동 확인용
function doGet() {
  return out_({ok: true, message: '인사조직 진단 설문 수신기가 작동 중입니다.'});
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(HEADER);
  return sh;
}

// 스프레드시트가 수식으로 해석하지 않도록 =, +, -, @ 로 시작하는 글은 앞에 ' 를 붙임
function safeText_(v, max) {
  let s = String(v == null ? '' : v).slice(0, max).trim();
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function notify_(row, a, sh) {
  try {
    const to = NOTIFY_TO || Session.getEffectiveUser().getEmail();
    if (!to) return;
    const ss = sh.getParent();
    const count = sh.getLastRow() - 1;
    const avg = arr => arr.reduce((s, v) => s + v, 0) / arr.length;
    const total = avg(a);
    const lines = AREAS.map(([name, s, e]) => '  · ' + name + ': ' + avg(a.slice(s - 1, e)).toFixed(2));
    const when = Utilities.formatDate(row[1], 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
    const body = [
      '인사조직 진단 설문에 새 응답이 도착했습니다. (누적 ' + count + '명)',
      '',
      '제출일시: ' + when,
      '직위: ' + (row[2] || '-') + ' / 직종: ' + (row[3] || '-') + (row[4] ? '(' + row[4] + ')' : '') +
        ' / 성별: ' + (row[5] || '-') + ' / 근속: ' + (row[6] === '' ? '-' : row[6] + '개월'),
      '',
      '이 응답의 종합 평균: ' + total.toFixed(2) + ' / 5점',
      lines.join('\n'),
      '',
      '기타 의견: ' + (row[row.length - 1] || '(없음)'),
      '',
      '전체 응답과 자동 분석 보기: ' + ss.getUrl(),
    ].join('\n');
    MailApp.sendEmail(to, '[인사조직 진단] 새 응답 도착 (누적 ' + count + '명)', body);
  } catch (err) {
    console.error('메일 발송 실패: ' + err);
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// (선택) 편집기에서 한 번 실행하면 권한 승인 창이 뜹니다. 승인 후 테스트 메일이 옵니다.
function testSetup() {
  const sh = getSheet_();
  MailApp.sendEmail(NOTIFY_TO || Session.getEffectiveUser().getEmail(),
    '[인사조직 진단] 알림 설정 확인', '설정이 완료되었습니다.\n응답 시트: ' + sh.getParent().getUrl());
}
