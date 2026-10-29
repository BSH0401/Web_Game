/**
 * 프로게이머 키우기 건의함 → 구글 시트 자동 정리
 *
 * Firebase(Firestore)의 feedback 컬렉션에서 새 제보를 읽어 '건의함' 시트에 한 줄씩 추가한다.
 * setup()을 한 번 실행하면 5분마다 자동으로 동기화된다.
 * 읽기는 시트 주인(Firebase 프로젝트 관리자)의 구글 계정 권한으로 한다.
 */
const PROJECT_ID = 'web-game-progamer';
const SHEET_NAME = '건의함';
const HEADERS = ['받은 시각', '유형', '내용', '연락처', '게임 상태', '처리 상태', '메모', '브라우저', '문서 ID'];
const STATUS = ['🆕 새 제보', '👀 확인함', '🔧 작업 중', '✅ 해결', '⏸ 보류'];
const TYPE_LABEL = { '건의': '💡 건의', '버그': '🐞 버그', '기타': '💬 기타' };
const SEND_EMAIL = false; // true로 바꾸면 새 제보가 올 때마다 내 메일로도 알려준다

/** 시트를 열면 위쪽 메뉴에 '🎮 건의함'이 생긴다 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎮 건의함')
    .addItem('지금 새 제보 가져오기', 'syncFeedback')
    .addItem('처음 설정 (한 번만)', 'setup')
    .addToUi();
}

/** 처음 한 번만 실행: 시트 꾸미기 + 5분마다 자동 실행 등록 + 기존 제보 가져오기 */
function setup() {
  const sh = getSheet_();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncFeedback')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncFeedback').timeBased().everyMinutes(5).create();
  const added = syncFeedback();
  SpreadsheetApp.getActive().toast(`설정 완료! 지금 ${added}건을 가져왔고, 이제 5분마다 새 제보를 자동으로 가져와요.`, '🎮 건의함', 10);
  return sh;
}

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME, 0);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    [140, 80, 420, 170, 250, 110, 220, 180, 120].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.getRange('C:C').setWrap(true);
    sh.hideColumns(9); // 문서 ID: 중복 방지용
    // '✅ 해결'로 바꾼 줄은 회색으로
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$F2="✅ 해결"')
      .setFontColor('#9aa0a6')
      .setRanges([sh.getRange('A2:H')])
      .build();
    sh.setConditionalFormatRules([rule]);
  }
  return sh;
}

/** Firestore에서 feedback 문서를 모두 읽는다 (관리자 권한) */
function fetchAll_() {
  const token = ScriptApp.getOAuthToken();
  const base = `https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents/feedback`;
  let docs = [], pageToken = '';
  do {
    const url = base + '?pageSize=300' + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
    const res = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token, 'X-Goog-User-Project': PROJECT_ID },
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) {
      throw new Error(`Firestore 읽기 실패 (${res.getResponseCode()}): ${res.getContentText().slice(0, 400)}`);
    }
    const j = JSON.parse(res.getContentText());
    docs = docs.concat(j.documents || []);
    pageToken = j.nextPageToken || '';
  } while (pageToken);
  return docs;
}

/** 시트에 '=' 등으로 시작하는 글이 수식으로 실행되지 않게 막는다 */
function safe_(s) {
  s = String(s == null ? '' : s);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/** 새 제보만 시트 아래에 추가하고, 추가한 개수를 돌려준다 */
function syncFeedback() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return 0;
  try {
    const sh = getSheet_();
    const last = sh.getLastRow();
    const seen = new Set(last > 1 ? sh.getRange(2, 9, last - 1, 1).getValues().map(r => String(r[0])) : []);
    const val = (f, k) => !f[k] ? '' : (f[k].stringValue !== undefined ? f[k].stringValue : (f[k].integerValue || ''));

    const rows = fetchAll_()
      .map(d => ({ id: d.name.split('/').pop(), t: new Date(d.createTime), f: d.fields || {} }))
      .filter(d => !seen.has(d.id))
      .sort((a, b) => a.t - b.t)
      .map(d => [
        d.t,
        TYPE_LABEL[val(d.f, 'type')] || safe_(val(d.f, 'type')),
        safe_(val(d.f, 'message')),
        safe_(val(d.f, 'contact')),
        safe_(val(d.f, 'state')),
        STATUS[0],
        '',
        safe_(val(d.f, 'ua')),
        d.id,
      ]);
    if (!rows.length) return 0;

    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, rows.length, HEADERS.length).setValues(rows);
    sh.getRange(start, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange(start, 6, rows.length, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).build());
    rows.forEach((r, i) => { if (String(r[1]).includes('버그')) sh.getRange(start + i, 2).setBackground('#fde2e2'); });

    if (SEND_EMAIL) {
      MailApp.sendEmail(
        Session.getEffectiveUser().getEmail(),
        `[프로게이머 키우기] 새 제보 ${rows.length}건`,
        rows.map(r => `${r[1]}  ${r[2]}`).join('\n\n') + '\n\n' + SpreadsheetApp.getActive().getUrl()
      );
    }
    return rows.length;
  } finally {
    lock.releaseLock();
  }
}
