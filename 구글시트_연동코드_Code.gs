/**
 * 스마트 주보 - 구글 시트 통합 데이터 연동 엔진 스크립트
 * 처리 대상: 접속 로그, 마음으로 드리는 고백 제출, 목장 나눔, 심방 및 양육 자율 신청서
 */

function doPost(e) {
  if (e && e.parameter && e.parameter.type === '특별기록') return saveSpecialRecord_(e);
  // CORS 처리 헤더 구성
  var output = ContentService.createTextOutput();
  
  try {
    var param = e.parameter;
    var type = param.type; // 데이터 전송 구분을 위한 식별 키
    
    // 현재 활성화된 구글 스프레드시트 가져오기
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "";
    var headers = [];
    var rowData = [];
    
    // 1. 요청 타입별 적치할 시트 이름 및 헤더 필드 매핑 구조 분기
    if (type === "접속자") {
      sheetName = "1_접속자로그";
      headers = ["기록일시", "주보설정날짜", "활동내용"];
      rowData = [param.date, param.bulletinDate, param.content];
      
    } else if (type === "마드고") {
      sheetName = "2_마음의고백";
      headers = ["제출일시", "작성자", "성경본문", "설교제목", "받은은혜와 깨달음", "기도제목"];
      rowData = [param.date, param.author, param.biblePassage, param.sermonTitle, param.grace, param.prayer];
      
    } else if (type === "목장나눔") {
      sheetName = "3_목장나눔기록";
      headers = ["작성일자", "소속목장", "작성성도", "나눔내용", "공유범위"];
      rowData = [param.date, param.mokjang, param.author, param.text, param.공유여부];
      
    } else if (type === "심방양육신청") {
      sheetName = "4_심방양육신청서";
      headers = ["신청일자", "소속목장", "신청성도", "희망일정 및 기입내용", "공유범위"];
      rowData = [param.date, param.mokjang, param.author, param.text, param.공유여부];
    } else {
      return output.setContent(JSON.stringify({status: "SUCCESS", message: "Unknown type bypassed"})).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 2. 대상 시트가 존재하지 않을 경우 자동 생성 및 헤더 배치 인프라 가동
    var targetSheet = ss.getSheetByName(sheetName);
    if (!targetSheet) {
      targetSheet = ss.insertSheet(sheetName);
      targetSheet.appendRow(headers);
      // 첫 행 디자인 가독성 확보 (볼드 처리 및 배경색 하이라이트)
      targetSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#F1F5F9");
    }
    
    // 3. 최하단 행에 데이터 자동 추가 적치
    targetSheet.appendRow(rowData);
    
    // 최종 가공 완료 응답 반환
    return output.setContent(JSON.stringify({status: "SUCCESS", sheet: sheetName}))
                 .setMimeType(ContentService.MimeType.JSON);
                 
  } catch(error) {
    // 예외 발생 시 에러 추적 결과 반환
    return output.setContent(JSON.stringify({status: "ERROR", message: error.toString()}))
                 .setMimeType(ContentService.MimeType.JSON);
  }
}

// GET 요청에 대한 기본 보안 확인 테스트 핑
function doGet(e) {
  return ContentService.createTextOutput("스마트 주보 클라우드 서버 엔진 작동 중.");
}
function saveSpecialRecord_(e) {
  var p = e.parameter;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var name = String(p.recordTitle || '특별기록').replace(/[\[\]\*\?\/\\:]/g, ' ').trim().slice(0, 90) || '특별기록';
    var sheet = ss.getSheetByName(name) || ss.insertSheet(name);

    var fields = [];
    try { fields = JSON.parse(p.fieldsJson || '[]'); } catch (err) { fields = []; }

    var fixed = ['제출일시', '이름', '목장/셀', '제출구분', '작성 항목 수'];
    var header = sheet.getLastRow() > 0
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String)
      : [];
    if (header.length === 0) header = fixed.slice();
    // 기존 열의 순서는 절대 바꾸지 않고, 새로 생긴 항목만 오른쪽 끝에 덧붙입니다.
    fields.forEach(function (f) { if (f && f.label && header.indexOf(f.label) === -1) header.push(f.label); });
    if (header.indexOf('전체 내용') === -1) header.push('전체 내용');
    if (header.indexOf('기록번호') === -1) header.push('기록번호');   // 수정·삭제할 줄을 찾는 번호

    sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#FEFCBF');
    sheet.setFrozenRows(1);

    // 같은 성도(같은 기록번호)가 이미 낸 줄 찾기
    var idCol = header.indexOf('기록번호') + 1;
    var foundRow = 0;
    var entryId = String(p.entryId || '');
    if (entryId && sheet.getLastRow() > 1) {
      var ids = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getValues();
      for (var i = ids.length - 1; i >= 0; i--) { if (String(ids[i][0]) === entryId) { foundRow = i + 2; break; } }
    }

    // 삭제 요청
    if (p.action === 'delete') {
      if (foundRow) sheet.deleteRow(foundRow);
      return ContentService.createTextOutput('deleted');
    }

    var row = header.map(function () { return ''; });
    var put = function (col, val) { var k = header.indexOf(col); if (k !== -1) row[k] = val; };
    put('제출일시', p.date || new Date());
    put('이름', p.author || '');
    put('목장/셀', p.mokjang || '');
    put('제출구분', foundRow ? '수정 제출' : (p.submitKind || '최초 제출'));
    put('작성 항목 수', p.filled || '');
    fields.forEach(function (f) { if (f && f.label) put(f.label, f.text || ''); });
    put('전체 내용', p.content || '');
    put('기록번호', entryId);

    if (foundRow) {
      sheet.getRange(foundRow, 1, 1, header.length).setValues([row]);   // 원래 줄을 새 내용으로 교체
    } else {
      sheet.appendRow(row);
      foundRow = sheet.getLastRow();
    }
    sheet.getRange(foundRow, 1, 1, header.length).setWrap(true).setVerticalAlignment('top');
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}
