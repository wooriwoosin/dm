/**
 * 아빠 대여금 관리 앱 - Google Apps Script (JSON API 백엔드)
 * 프론트엔드는 별도로 호스팅되는 정적 HTML(예: GitHub Pages)에서 fetch()로 이 웹앱을 호출합니다.
 * 스프레드시트: 1umwRDxeqYwLQ8nug9Xc_JVJeEKbNyaCwFgSsT51i8Lw / 시트: 아빠상환
 * build v2026.08.11-01
 */

const BUILD = 'v2026.08.11-01'; // 응답 JSON에 그대로 찍혀서, 지금 이 코드가 실제로 실행 중인지 확실히 확인 가능

const SHEET_ID = '1umwRDxeqYwLQ8nug9Xc_JVJeEKbNyaCwFgSsT51i8Lw';
const SHEET_NAME = '아빠상환';
const START_ROW = 3;          // 데이터 시작 행 (2행은 헤더)
const PRINCIPAL = 95000000;   // 아빠 대여금 원금 (자금출처 표 합계 기준)

const APP_ID = '민정';
const APP_PW = 'alswjd2';

// ---------- 웹 요청 라우팅 ----------
// 전부 GET 하나로 처리한다: ?action=getData|checkLogin|addEntry|updateEntry&payload=<JSON 문자열>
// (POST + JSON은 배포 환경에 따라 CORS/사전요청에서 막히는 경우가 있어 GET으로 통일)

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'getData';
    const payload = (e && e.parameter && e.parameter.payload) ? JSON.parse(e.parameter.payload) : {};
    let result;
    if (action === 'getData') {
      result = getData();
    } else if (action === 'checkLogin') {
      result = checkLogin(payload.id, payload.pw);
    } else if (action === 'addEntry') {
      result = addEntry(payload);
    } else if (action === 'updateEntry') {
      result = updateEntry(payload);
    } else {
      throw new Error('알 수 없는 요청입니다: ' + action);
    }
    return jsonOutput_({ ok: true, data: result });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err.message || err) });
  }
}

function jsonOutput_(obj) {
  obj.build = BUILD; // 모든 응답에 현재 실행 중인 코드 버전을 찍는다
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- 비즈니스 로직 ----------

function checkLogin(id, pw) {
  return id === APP_ID && pw === APP_PW;
}

function getSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
}

// 실제 값이 채워진 마지막 행 찾기.
// (예전엔 B열/No.가 비어있지 않은 마지막 행을 직접 스캔했는데, 새로 삽입된 행이
//  일시적으로 No.가 비어있는 경우 범위 계산에서 통째로 빠지는 문제가 반복됐다.
//  대신 시트가 자체적으로 관리하는 "실제 데이터가 있는 마지막 행"을 그대로 쓴다 —
//  어떤 열에든 값이 있으면 잡아내므로 훨씬 안전하다.)
function getLastRow_(sheet) {
  const lr = sheet.getLastRow();
  return lr < START_ROW ? START_ROW - 1 : lr;
}

// 셀 값이 실제 Date 객체든, "2022-03-28" / "2022. 3. 28" / "2022.3.28" 같은
// 텍스트로 된 날짜든 최대한 실제 날짜로 인식해서 반환한다.
//
// 중요: Apps Script에서 시트 셀로부터 읽어온 Date 값은 "val instanceof Date"가
// false로 나올 때가 있다 (실제로는 정상 날짜인데도). Object.prototype.toString으로
// 확인하는 방식은 이런 경우에도 안전하게 "진짜 Date"임을 잡아낸다.
function parseDateCell_(val) {
  if (val && Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    return val;
  }
  if (typeof val === 'string') {
    const s = val.trim();
    if (!s) return null;
    const m = s.match(/^(\d{4})[.\-\/\s]+(\d{1,2})[.\-\/\s]+(\d{1,2})/);
    if (m) {
      const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
      const dt = new Date(y, mo - 1, d);
      if (!isNaN(dt.getTime())) return dt;
    }
  }
  return null;
}

// 정렬용 "월 인덱스" 키.
// "26.7월분" 같은 라벨이면 그 연·월(라벨 기준)을 쓰고, 라벨에 월 정보가 없으면
// (주형이상환·추가대여금 등) 입금일자의 연·월을 쓴다. 둘 다 없으면 맨 뒤로 보낸다.
//
// 왜 입금일자가 아니라 "월분 라벨"을 우선하나:
//  예) 23.6월분은 실제 입금일이 2023.8.1 인데, 23.7월분(미입금)은 날짜가 없다.
//  입금일 기준으로 정렬하면 23.7월분이 23.6월분보다 앞서버린다. 장부의 진짜 순서는
//  "몇 월분"이라는 슬롯이므로 라벨의 연·월을 우선한다.
function monthKeyOf_(label, parsedDate) {
  if (label) {
    const m = String(label).match(/(\d{2})\s*\.\s*(\d{1,2})\s*월/);
    if (m) {
      return (2000 + Number(m[1])) * 12 + (Number(m[2]) - 1);
    }
  }
  if (parsedDate) {
    return parsedDate.getFullYear() * 12 + parsedDate.getMonth();
  }
  return Number.MAX_SAFE_INTEGER;
}

function getData() {
  const sheet = getSheet_();
  const lastRow = getLastRow_(sheet);
  const numRows = lastRow - START_ROW + 1;
  const tz = Session.getScriptTimeZone();

  const raw = [];
  if (numRows > 0) {
    const values = sheet.getRange(START_ROW, 2, numRows, 7).getValues(); // B~H
    values.forEach((row, idx) => {
      const [no, label, date, deposit, insurance, interest, repay] = row;
      const labelStr = (label === null || label === undefined) ? '' : String(label).trim();
      const parsedDate = parseDateCell_(date);
      const dateStr = parsedDate ? Utilities.formatDate(parsedDate, tz, 'yyyy-MM-dd') : '';
      const depositNum = (typeof deposit === 'number') ? deposit : 0;
      const insuranceNum = (typeof insurance === 'number') ? insurance : 0;
      const interestNum = (typeof interest === 'number') ? interest : 0;
      const repayNum = (typeof repay === 'number') ? repay : 0;

      // 구분·입금일자·금액이 모두 없는 완전 빈 버퍼 행은 목록에서 제외
      const isEmpty = !labelStr && !dateStr &&
        depositNum === 0 && insuranceNum === 0 && interestNum === 0 && repayNum === 0;
      if (isEmpty) return;

      raw.push({
        rowIndex: START_ROW + idx,
        origOrder: idx,
        sortKey: monthKeyOf_(labelStr, parsedDate),
        no: no,
        label: labelStr,
        date: dateStr,
        deposit: depositNum,
        insurance: insuranceNum,
        interest: interestNum,
        repay: repayNum,
        hasData: !!dateStr   // 실제 입금일자가 있는 행만 "실제 기록"으로 처리 (예정된 미래 행은 제외)
      });
    });
  }

  // 시트의 물리적 행 순서가 뒤죽박죽이어도(예: 새 항목이 맨 위에 잘못 삽입돼도)
  // 항상 올바른 시간순(오래된→최신)으로 계산/반환되도록 월 기준으로 정렬한다.
  raw.sort((a, b) => (a.sortKey - b.sortKey) || (a.origOrder - b.origOrder));

  let cumulative = 0;
  const entries = raw.map(e => {
    cumulative += e.repay;
    return {
      rowIndex: e.rowIndex,
      no: e.no,
      label: e.label,
      date: e.date,
      deposit: e.deposit,
      insurance: e.insurance,
      interest: e.interest,
      repay: e.repay,
      cumulative: cumulative,
      hasData: e.hasData
    };
  });

  const totalRepay = cumulative;
  const remaining = PRINCIPAL - totalRepay;

  return {
    entries: entries,
    principal: PRINCIPAL,
    totalRepay: totalRepay,
    remaining: remaining,
    progress: Math.round((totalRepay / PRINCIPAL) * 1000) / 10
  };
}

// 기존 행 수정
function updateEntry(entry) {
  const sheet = getSheet_();
  const row = entry.rowIndex;
  const repay = (Number(entry.deposit) || 0) + (Number(entry.insurance) || 0) + (Number(entry.interest) || 0);

  sheet.getRange(row, 3).setValue(entry.label || '');
  if (entry.date) {
    sheet.getRange(row, 4).setValue(new Date(entry.date));
  } else {
    sheet.getRange(row, 4).clearContent();
  }
  sheet.getRange(row, 5).setValue(Number(entry.deposit) || 0);
  sheet.getRange(row, 6).setValue(Number(entry.insurance) || 0);
  sheet.getRange(row, 7).setValue(Number(entry.interest) || 0);
  sheet.getRange(row, 8).setValue(repay);

  return getData();
}

// 새 항목 추가: 실제 데이터가 있는 마지막 행 바로 다음에 삽입한다.
// (예전엔 "마지막 '날짜 있는' 행" 다음에 넣었는데, 날짜 인식이 한 번이라도 어긋나면
//  삽입 위치가 최상단으로 떨어져서 새 항목이 시트 맨 위에 잘못 꽂히는 문제가 있었다.
//  구분·입금일자·금액 중 하나라도 있으면 "데이터가 있는 행"으로 보고 그 다음에 넣는다.
//  어차피 getData가 월 기준으로 다시 정렬하므로 물리적 위치는 표시에 영향을 주지 않는다.)
function addEntry(entry) {
  const sheet = getSheet_();
  const lastRow = getLastRow_(sheet);

  let lastDataRow = START_ROW - 1;
  if (lastRow >= START_ROW) {
    const block = sheet.getRange(START_ROW, 2, lastRow - START_ROW + 1, 7).getValues(); // B~H
    for (let i = 0; i < block.length; i++) {
      const [no, label, date, deposit, insurance, interest, repay] = block[i];
      const has = (label && String(label).trim()) || parseDateCell_(date) ||
        (typeof deposit === 'number' && deposit !== 0) ||
        (typeof insurance === 'number' && insurance !== 0) ||
        (typeof interest === 'number' && interest !== 0) ||
        (typeof repay === 'number' && repay !== 0);
      if (has) lastDataRow = START_ROW + i;
    }
  }
  const insertRow = lastDataRow + 1;

  sheet.insertRowBefore(insertRow);

  // 새로 삽입된 행은 No.(B열)가 비어있는 상태로 시작한다. 이 칸이 비어있으면
  // getLastRow_()가 이 행을 "마지막 행" 범위 계산에서 놓칠 수 있으므로,
  // renumberRows_가 다시 정리하기 전까지 쓸 임시 번호를 즉시 채워둔다.
  sheet.getRange(insertRow, 2).setValue(insertRow - START_ROW + 1);

  const repay = (Number(entry.deposit) || 0) + (Number(entry.insurance) || 0) + (Number(entry.interest) || 0);
  sheet.getRange(insertRow, 3).setValue(entry.label || '');
  if (entry.date) sheet.getRange(insertRow, 4).setValue(new Date(entry.date));
  sheet.getRange(insertRow, 5).setValue(Number(entry.deposit) || 0);
  sheet.getRange(insertRow, 6).setValue(Number(entry.insurance) || 0);
  sheet.getRange(insertRow, 7).setValue(Number(entry.interest) || 0);
  sheet.getRange(insertRow, 8).setValue(repay);

  SpreadsheetApp.flush();
  renumberRows_(sheet);
  SpreadsheetApp.flush();
  return getData();
}

function renumberRows_(sheet) {
  const lastRow = getLastRow_(sheet);
  const n = lastRow - START_ROW + 1;
  if (n <= 0) return; // 안전장치: 계산된 범위가 없으면 그냥 건너뛴다 (전체 삭제 방지)
  const nos = [];
  for (let i = 1; i <= n; i++) nos.push([i]);
  sheet.getRange(START_ROW, 2, n, 1).setValues(nos);
}

// ---------- 진단용: D열(입금일자) 셀에 실제로 뭐가 들어있는지 확인 ----------
// Apps Script 편집기 상단 함수 선택 드롭다운에서 "debugDateCell"을 고르고
// ▶ 실행 버튼을 누른 뒤, "실행" 탭(또는 보기 > 로그, Ctrl+Enter)에서 결과를 확인한다.
// 배포/웹앱 URL과 무관하게 편집기 안에서 바로 실행되므로 가장 확실한 진단이다.
function debugDateCell() {
  const sheet = getSheet_();
  const testRows = [3, 8, 25, 63]; // 26.7월분(신규), 22.2월분, 23.4월분, 26.6월분 근처
  testRows.forEach(function(r) {
    const val = sheet.getRange(r, 4).getValue();
    Logger.log(
      'row ' + r +
      ' | typeof=' + typeof val +
      ' | instanceof Date=' + (val instanceof Date) +
      ' | JSON=' + JSON.stringify(val) +
      ' | String()=' + String(val)
    );
  });
}
