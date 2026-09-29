/**
 * 맛집투어관리 앱 — Apps Script 백엔드
 * v1.0.0
 *
 * 배포 방법: 배포 > 배포 관리 > 새 버전으로 재배포 (반드시 "저장"이 아니라 "새 버전 배포")
 *
 * 시트 구성
 *  - 방문기록 : 맛집 방문 기록 (메인 데이터)
 *  - 그룹     : 그룹 및 구성원 관리
 *
 * 컬럼은 헤더 이름 기반으로 읽고 쓰기 때문에(getColMap_), 나중에 컬럼을 추가해도
 * 기존 코드의 range 크기를 직접 손댈 필요가 없습니다. (여행이력 앱에서 겪었던
 * "범위 크기가 잘못됨" 문제를 구조적으로 방지)
 */

var APP_VERSION = '1.0.0';

var VISIT_SHEET_NAME = '방문기록';
var GROUP_SHEET_NAME = '그룹';

var VISIT_HEADERS = [
  'id', '방문일', '방문시간', '가게명', '위치', '위도', '경도',
  '카테고리', '메뉴', '평점', '지출금액', '재방문의사',
  '영업시간', '주차가능여부', '웨이팅시간', '동행자',
  '사진', '등록자', 'groupId', '메모', '생성일시', '수정일시'
];

var GROUP_HEADERS = ['groupId', '그룹명', '구성원', '생성일시'];

var PHOTO_FOLDER_NAME = '맛집투어관리앱_사진';

var LOCK_DAYS = 7; // 그룹 모드에서 등록 후 이 일수가 지나면 등록자만 수정/삭제 가능

// ---------------------------------------------------------------------------
// 진입점
// ---------------------------------------------------------------------------

function doGet(e) {
  try {
    var action = (e.parameter.action || 'ping');
    var result;
    switch (action) {
      case 'ping':
        result = { ok: true, version: APP_VERSION, serverTime: new Date().toISOString() };
        break;
      case 'getVisits':
        result = { ok: true, visits: getVisits_() };
        break;
      case 'getGroups':
        result = { ok: true, groups: getGroups_() };
        break;
      default:
        result = { ok: false, error: 'Unknown GET action: ' + action };
    }
    return jsonOut_(result);
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function doPost(e) {
  try {
    var body = {};
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    var action = body.action || (e.parameter && e.parameter.action) || '';
    var result;
    switch (action) {
      case 'addVisit':
        result = { ok: true, visit: addVisit_(body.visit) };
        break;
      case 'updateVisit':
        result = { ok: true, visit: updateVisit_(body.visit, body.requester) };
        break;
      case 'deleteVisit':
        result = { ok: true, deletedId: deleteVisit_(body.id, body.requester) };
        break;
      case 'uploadPhoto':
        result = { ok: true, photo: uploadPhoto_(body.filename, body.mimeType, body.data) };
        break;
      case 'deletePhoto':
        result = { ok: true, deleted: deletePhoto_(body.fileId) };
        break;
      case 'addGroup':
        result = { ok: true, group: addGroup_(body.name, body.members) };
        break;
      case 'updateGroupMembers':
        result = { ok: true, group: updateGroupMembers_(body.groupId, body.members) };
        break;
      case 'ensureHeaders':
        ensureHeaders_();
        result = { ok: true };
        break;
      default:
        result = { ok: false, error: 'Unknown POST action: ' + action };
    }
    return jsonOut_(result);
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// 시트 / 헤더 유틸
// ---------------------------------------------------------------------------

function getSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet_(name, headers) {
  var ss = getSS_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  ensureHeadersOnSheet_(sheet, headers);
  return sheet;
}

// 헤더 행에 없는 컬럼이 있으면 뒤에 추가한다 (기존 데이터/컬럼 순서는 건드리지 않음)
function ensureHeadersOnSheet_(sheet, headers) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var existingSet = {};
  for (var i = 0; i < existing.length; i++) existingSet[existing[i]] = true;
  var toAdd = [];
  for (var j = 0; j < headers.length; j++) {
    if (!existingSet[headers[j]]) toAdd.push(headers[j]);
  }
  if (toAdd.length > 0) {
    sheet.getRange(1, lastCol + 1, 1, toAdd.length).setValues([toAdd]);
  }
}

function ensureHeaders_() {
  getOrCreateSheet_(VISIT_SHEET_NAME, VISIT_HEADERS);
  getOrCreateSheet_(GROUP_SHEET_NAME, GROUP_HEADERS);
  fixDateTextCells_();
}

// 기존 행 중 방문일/방문시간이 구글시트에 의해 실제 Date로 저장돼버린 것들을
// 찾아서 문자열로 되돌리고, 셀 서식도 텍스트로 고정한다. "예전 기록 정렬 점검하기"
// 버튼을 누르면 이 함수까지 같이 실행된다.
function fixDateTextCells_() {
  var sheet = getOrCreateSheet_(VISIT_SHEET_NAME, VISIT_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var colMap = getColMap_(sheet);
  var tz = Session.getScriptTimeZone() || 'Asia/Seoul';

  ['방문일', '방문시간'].forEach(function (key) {
    var col = colMap[key];
    if (!col) return;
    var range = sheet.getRange(2, col, lastRow - 1, 1);
    var values = range.getValues();
    var changed = false;
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] instanceof Date) {
        values[i][0] = Utilities.formatDate(values[i][0], tz, key === '방문일' ? 'yyyy-MM-dd' : 'HH:mm');
        changed = true;
      }
    }
    range.setNumberFormat('@'); // 앞으로도 자동 날짜 변환되지 않도록 텍스트로 고정
    if (changed) range.setValues(values);
  });
}

// 헤더이름 -> 1-based 컬럼번호 맵
function getColMap_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var i = 0; i < headerRow.length; i++) {
    if (headerRow[i]) map[headerRow[i]] = i + 1;
  }
  return map;
}

function rowToObject_(rowValues, colMap) {
  var obj = {};
  for (var key in colMap) {
    obj[key] = rowValues[colMap[key] - 1];
  }
  return obj;
}

// obj의 각 필드를 colMap 순서에 맞는 배열로 변환 (없는 값은 빈 문자열)
function objectToRow_(obj, colMap, width) {
  var row = new Array(width).fill('');
  for (var key in colMap) {
    if (Object.prototype.hasOwnProperty.call(obj, key)) {
      row[colMap[key] - 1] = obj[key];
    }
  }
  return row;
}

function genId_() {
  return Utilities.getUuid();
}

function nowIso_() {
  return new Date().toISOString();
}

// ---------------------------------------------------------------------------
// 방문기록 (Visits)
// ---------------------------------------------------------------------------

function getVisits_() {
  var sheet = getOrCreateSheet_(VISIT_SHEET_NAME, VISIT_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = sheet.getLastColumn();
  var colMap = getColMap_(sheet);
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var obj = rowToObject_(values[i], colMap);
    if (!obj.id) continue; // 빈 행 skip
    out.push(normalizeVisitOut_(obj));
  }
  return out;
}

function normalizeVisitOut_(obj) {
  // 사진은 JSON 문자열로 저장되어 있으므로 배열로 풀어준다
  try {
    obj.사진 = obj.사진 ? JSON.parse(obj.사진) : [];
  } catch (e) {
    obj.사진 = [];
  }
  coerceDateTextFields_(obj);
  return obj;
}

// 구글시트가 "2026-09-20", "15:32" 같은 문자열을 날짜/시간으로 자동 인식해서
// 실제 Date 객체로 바꿔버리는 경우가 있다 (특히 새로 만든 시트의 기본 셀 서식이
// "자동"이기 때문). 그렇게 되면 JSON으로 내보낼 때 "2026-09-20T15:00:00.000Z" 같은
// 이상한 값으로 직렬화된다. 여기서 다시 원래 형식의 문자열로 되돌린다.
function coerceDateTextFields_(obj) {
  var tz = Session.getScriptTimeZone() || 'Asia/Seoul';
  if (obj.방문일 instanceof Date) {
    obj.방문일 = Utilities.formatDate(obj.방문일, tz, 'yyyy-MM-dd');
  }
  if (obj.방문시간 instanceof Date) {
    obj.방문시간 = Utilities.formatDate(obj.방문시간, tz, 'HH:mm');
  }
  return obj;
}

// 날짜로 오인되기 쉬운 컬럼(방문일/방문시간)은 셀 서식을 "일반 텍스트"로 고정해서
// 앞으로 저장할 때 구글시트가 다시 Date로 자동 변환하지 못하게 막는다.
function forceTextFormatForDateCols_(sheet, colMap, rowNum) {
  ['방문일', '방문시간'].forEach(function (key) {
    var col = colMap[key];
    if (col) sheet.getRange(rowNum, col).setNumberFormat('@');
  });
}

function findVisitRow_(sheet, colMap, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var idCol = colMap['id'];
  var ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return i + 2; // 실제 시트 row 번호
  }
  return -1;
}

function addVisit_(visit) {
  if (!visit) throw new Error('visit 데이터가 없습니다.');
  var sheet = getOrCreateSheet_(VISIT_SHEET_NAME, VISIT_HEADERS);
  var colMap = getColMap_(sheet);
  var lastCol = Math.max(sheet.getLastColumn(), VISIT_HEADERS.length);

  var record = Object.assign({}, visit);
  record.id = genId_();
  record.생성일시 = nowIso_();
  record.수정일시 = record.생성일시;
  if (Array.isArray(record.사진)) record.사진 = JSON.stringify(record.사진);
  if (Array.isArray(record.카테고리)) record.카테고리 = record.카테고리.join(',');
  if (Array.isArray(record.동행자)) record.동행자 = record.동행자.join(',');

  var targetRow = sheet.getLastRow() + 1;
  forceTextFormatForDateCols_(sheet, colMap, targetRow); // setValues보다 먼저 서식을 텍스트로 고정
  var row = objectToRow_(record, colMap, lastCol);
  sheet.getRange(targetRow, 1, 1, lastCol).setValues([row]);
  return normalizeVisitOut_(record);
}

function updateVisit_(visit, requester) {
  if (!visit || !visit.id) throw new Error('수정할 방문기록의 id가 없습니다.');
  var sheet = getOrCreateSheet_(VISIT_SHEET_NAME, VISIT_HEADERS);
  var colMap = getColMap_(sheet);
  var rowNum = findVisitRow_(sheet, colMap, visit.id);
  if (rowNum === -1) throw new Error('해당 방문기록을 찾을 수 없습니다: ' + visit.id);

  var lastCol = sheet.getLastColumn();
  var currentValues = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
  var current = rowToObject_(currentValues, colMap);

  assertEditable_(current, requester);

  var merged = Object.assign({}, current, visit);
  merged.수정일시 = nowIso_();
  merged.생성일시 = current.생성일시; // 생성일시는 보존
  merged.등록자 = current.등록자; // 등록자는 변경 불가
  merged.groupId = current.groupId; // groupId는 변경 불가
  if (Array.isArray(merged.사진)) merged.사진 = JSON.stringify(merged.사진);
  if (Array.isArray(merged.카테고리)) merged.카테고리 = merged.카테고리.join(',');
  if (Array.isArray(merged.동행자)) merged.동행자 = merged.동행자.join(',');
  coerceDateTextFields_(merged); // 기존 셀이 이미 Date로 오염돼 있었다면 여기서 문자열로 되돌림

  forceTextFormatForDateCols_(sheet, colMap, rowNum); // setValues보다 먼저 서식을 텍스트로 고정
  var row = objectToRow_(merged, colMap, lastCol);
  sheet.getRange(rowNum, 1, 1, lastCol).setValues([row]);
  return normalizeVisitOut_(merged);
}

function deleteVisit_(id, requester) {
  if (!id) throw new Error('삭제할 방문기록의 id가 없습니다.');
  var sheet = getOrCreateSheet_(VISIT_SHEET_NAME, VISIT_HEADERS);
  var colMap = getColMap_(sheet);
  var rowNum = findVisitRow_(sheet, colMap, id);
  if (rowNum === -1) throw new Error('해당 방문기록을 찾을 수 없습니다: ' + id);

  var lastCol = sheet.getLastColumn();
  var currentValues = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
  var current = rowToObject_(currentValues, colMap);

  assertEditable_(current, requester);

  sheet.deleteRow(rowNum);
  return id;
}

// 그룹 모드(레코드에 groupId가 있음) + 등록 7일 경과 + 요청자가 등록자 본인이 아니면 거부
function assertEditable_(current, requester) {
  var hasGroup = current.groupId && String(current.groupId).length > 0;
  if (!hasGroup) return; // 개인 모드 기록은 기간 제한 없음

  var isOwner = requester && current.등록자 && String(requester) === String(current.등록자);
  if (isOwner) return; // 본인은 언제나 수정/삭제 가능

  var created = current.생성일시 ? new Date(current.생성일시) : null;
  if (!created) return;
  var days = (Date.now() - created.getTime()) / (1000 * 60 * 60 * 24);
  if (days > LOCK_DAYS) {
    throw new Error('등록 후 ' + LOCK_DAYS + '일이 지난 그룹 기록은 등록자 본인만 수정/삭제할 수 있습니다.');
  }
}

// ---------------------------------------------------------------------------
// 그룹 (Groups)
// ---------------------------------------------------------------------------

function getGroups_() {
  var sheet = getOrCreateSheet_(GROUP_SHEET_NAME, GROUP_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = sheet.getLastColumn();
  var colMap = getColMap_(sheet);
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var obj = rowToObject_(values[i], colMap);
    if (!obj.groupId) continue;
    obj.구성원 = obj.구성원 ? String(obj.구성원).split(',').map(function (s) { return s.trim(); }).filter(Boolean) : [];
    out.push(obj);
  }
  return out;
}

function addGroup_(name, members) {
  if (!name) throw new Error('그룹명이 없습니다.');
  var sheet = getOrCreateSheet_(GROUP_SHEET_NAME, GROUP_HEADERS);
  var colMap = getColMap_(sheet);
  var lastCol = Math.max(sheet.getLastColumn(), GROUP_HEADERS.length);

  var record = {
    groupId: genId_(),
    그룹명: name,
    구성원: Array.isArray(members) ? members.join(',') : (members || ''),
    생성일시: nowIso_()
  };
  var row = objectToRow_(record, colMap, lastCol);
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, lastCol).setValues([row]);
  record.구성원 = record.구성원 ? record.구성원.split(',') : [];
  return record;
}

function updateGroupMembers_(groupId, members) {
  if (!groupId) throw new Error('groupId가 없습니다.');
  var sheet = getOrCreateSheet_(GROUP_SHEET_NAME, GROUP_HEADERS);
  var colMap = getColMap_(sheet);
  var lastRow = sheet.getLastRow();
  var idCol = colMap['groupId'];
  var membersCol = colMap['구성원'];
  for (var r = 2; r <= lastRow; r++) {
    var gid = sheet.getRange(r, idCol).getValue();
    if (gid === groupId) {
      var joined = Array.isArray(members) ? members.join(',') : (members || '');
      sheet.getRange(r, membersCol).setValue(joined);
      return { groupId: groupId, 구성원: Array.isArray(members) ? members : joined.split(',') };
    }
  }
  throw new Error('해당 그룹을 찾을 수 없습니다: ' + groupId);
}

// ---------------------------------------------------------------------------
// 사진 (Google Drive)
// ---------------------------------------------------------------------------

function getPhotoFolder_() {
  var folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

function uploadPhoto_(filename, mimeType, base64Data) {
  if (!base64Data) throw new Error('업로드할 이미지 데이터가 없습니다.');
  var folder = getPhotoFolder_();
  var bytes = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg', filename || (genId_() + '.jpg'));
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var fileId = file.getId();
  return {
    fileId: fileId,
    url: 'https://drive.google.com/uc?export=view&id=' + fileId,
    thumbUrl: 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w480'
  };
}

function deletePhoto_(fileId) {
  if (!fileId) throw new Error('fileId가 없습니다.');
  DriveApp.getFileById(fileId).setTrashed(true);
  return fileId;
}
