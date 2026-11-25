/**
 * 스마트 주보 - 구글 시트 통합 데이터 연동 엔진 스크립트
 * 처리 대상: 접속 로그, 마음으로 드리는 고백 제출, 목장 나눔, 심방 및 양육 자율 신청서
 * v1.1: '주보' 탭에 쓴 이번 주 내용과 유튜브 최신 예배 영상을 주보 앱에 보내 줍니다 (아래 doGet).
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

// GET 요청: ?action=bulletin 이면 '주보' 탭 내용 + 최신 예배 영상을 보내고, 그 밖에는 작동 확인 문구만 보냅니다.
// '주보' 탭 말고 다른 탭(고백·목장 나눔 등 성도가 낸 글)은 절대 내보내지 않습니다.
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'bulletin') {
    var data;
    try { data = getBulletin_(); } catch (err) { data = { ok: false, reason: String(err) }; }
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
  }
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


/* ===================================================================
 * v1.1  '주보' 탭으로 이번 주 주보 쓰기 + 유튜브 예배 영상 자동 연결
 * -------------------------------------------------------------------
 * 1) 시트 위쪽 메뉴 [스마트 주보] → [① 주보 입력 탭 만들기]
 * 2) '주보' 탭의 '내용' 칸에 이번 주 내용을 씁니다. 쓰는 즉시 주보 앱에 반영됩니다.
 * 3) '유튜브 채널' 칸에 채널 주소를 넣고 [② 유튜브 연결 확인하기]로 확인합니다.
 * =================================================================== */

var BULLETIN_SHEET = '주보';
var BULLETIN_CACHE = 'bulletin-v1';

// [항목 이름, 앱 안의 이름, 처음 값, 쓰는 방법]
// 앱 안의 이름이 #으로 시작하면 설정값이라 주보 앱으로 보내지 않습니다.
// 탭을 처음 만들 때 내용 칸은 비워 둡니다. 내용 칸이 모두 비어 있는 동안에는 주보 파일 내용이 그대로 보입니다.
var BULLETIN_ROWS = [
  ['주보 날짜', 'bulletinDate', '', '예: 2026-11-29 · 이 날짜에 맞춰 첫 화면의 절기 색이 바뀝니다'],
  ['설교 제목', 'sermonTitle', '', '예: 어둠 속에 비친 큰 빛'],
  ['설교자', 'preacher', '', '예: 홍길동 목사'],
  ['성경 본문 제목', 'bibleHeading', '', '예: 요한복음 3장 16절 · 첫 화면 카드에 보이는 본문 이름'],
  ['성경 본문', 'biblePassage', '', '절마다 번호(16. 17.)를 붙여 쓰세요 · 칸 안에서 줄 바꾸기: Ctrl+Enter (맥은 ⌘+Enter)'],
  ['설교 노트', 'summary', '', '1. 2. 3. 번호를 붙여 쓰면 가지런히 정렬됩니다'],
  ['나눔 질문', 'questions', '', '1. 2. 번호를 붙여 쓰세요'],
  ['예배 순서', 'worshipOrder', '', '순서마다 * 로 시작하고, 설명 줄은 : 로 시작하세요 · 예: * 설교 (줄 바꿈) : 요나 1장 1-3절'],
  ['교회 소식', 'churchNews', '', '1. 2. 3. 번호로 제목을 쓰고, 설명 줄은 : 로 시작하세요'],
  ['예배 안내', 'serviceInfo', '', '예배 시간 안내 · 예: 주일 예배 : 오전 11시'],
  ['주일 예배 영상 링크', 'youtubeLink', '', '비워 두면 아래 [유튜브 채널]에서 가장 최근 예배 영상을 자동으로 찾아 연결합니다'],
  ['금요 기도회 링크', 'fridayLink', '', '비워 두면 원래 주보 파일에 있는 링크를 씁니다'],
  ['금요 기도회 안내 문구', 'fridayNotice', '', '버튼에 보일 짧은 글 · 비워 두면 원래 값'],
  ['말씀 도시락 링크', 'dosirakLink', '', '비워 두면 원래 주보 파일에 있는 링크를 씁니다'],
  ['말씀 쇼츠 링크', 'shortsLink', '', '비워 두면 원래 주보 파일에 있는 링크를 씁니다'],
  ['유튜브 채널', '#channel', '', '[설정] 교회 유튜브 채널 주소 · 예: @우리교회 또는 https://www.youtube.com/@우리교회'],
  ['예배 영상 제목 단어', '#keyword', '주일예배', '[설정] 영상 제목에 이 말이 들어 있으면 예배 영상으로 봅니다 · 여러 개는 쉼표로: 주일예배, 주일 낮 예배']
];

// 교회별 처음 값 (선택): 다른 .gs 파일에 var BULLETIN_PRESET = { '#channel': '@채널', '#keyword': '주일예배' }; 처럼 두면 탭을 만들 때 그 값으로 채웁니다.
function presetValue_(key, fallback) {
  var p = (typeof BULLETIN_PRESET === 'object' && BULLETIN_PRESET) || {};
  return Object.prototype.hasOwnProperty.call(p, key) ? p[key] : fallback;
}

function normLabel_(s) { return String(s == null ? '' : s).replace(/\s+/g, ''); }

function cellText_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v == null ? '' : v).replace(/\r\n?/g, '\n');
}

// '주보' 탭 → { 앱 안의 이름: 내용 }  (탭이 없으면 null)
function readBulletinSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BULLETIN_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var byLabel = {};
  BULLETIN_ROWS.forEach(function (r) { byLabel[normLabel_(r[0])] = r[1]; });
  var tz = ss.getSpreadsheetTimeZone();
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  var out = {};
  values.forEach(function (row) {
    var key = byLabel[normLabel_(row[0])];
    if (key) out[key] = cellText_(row[1], tz).trim();
  });
  return out;
}

function getBulletin_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(BULLETIN_CACHE);
  if (hit) return JSON.parse(hit);

  var rows = readBulletinSheet_();
  if (!rows) return { ok: false, reason: "'" + BULLETIN_SHEET + "' 탭이 없습니다" };
  var content = {};
  Object.keys(rows).forEach(function (k) { if (k.charAt(0) !== '#') content[k] = rows[k]; });
  // 내용 칸이 모두 비어 있으면(탭을 막 만든 상태) 주보 파일 내용을 그대로 두도록 내용은 보내지 않고 영상만 보냅니다.
  if (!Object.keys(content).some(function (k) { return content[k] !== ''; })) content = {};
  var found = findLatestVideo_(rows['#channel'], rows['#keyword'], false);
  var data = { ok: true, content: content, video: found.video, at: new Date().toISOString() };
  cache.put(BULLETIN_CACHE, JSON.stringify(data), 60);
  return data;
}

// 유튜브 채널의 최신 영상 목록(RSS, API 키 필요 없음)에서 제목 단어가 들어간 가장 최근 영상을 찾습니다.
function findLatestVideo_(channelText, keywordText, skipCache) {
  channelText = String(channelText || '').trim();
  if (!channelText) return { video: null, note: '유튜브 채널 칸이 비어 있습니다' };
  var cache = CacheService.getScriptCache();
  var ck = 'yt-v1-' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, channelText + '|' + keywordText, Utilities.Charset.UTF_8));
  if (!skipCache) {
    var hit = cache.get(ck);
    if (hit) return JSON.parse(hit);
  }
  var result;
  try {
    var id = resolveChannelId_(channelText);
    if (!id) {
      result = { video: null, note: '채널을 찾지 못했습니다. 채널 주소(@이름) 또는 UC로 시작하는 채널 ID를 넣어 주세요' };
    } else {
      var res = UrlFetchApp.fetch('https://www.youtube.com/feeds/videos.xml?channel_id=' + id, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) {
        result = { video: null, note: '유튜브 영상 목록을 읽지 못했습니다 (응답 ' + res.getResponseCode() + ')' };
      } else {
        var video = pickVideo_(res.getContentText(), keywordText);
        result = video ? { video: video, note: '' } : { video: null, note: '제목에 [' + (keywordText || '') + '] 이(가) 들어간 최근 영상이 없습니다' };
      }
    }
  } catch (err) {
    result = { video: null, note: '유튜브 확인 중 오류: ' + err };
  }
  cache.put(ck, JSON.stringify(result), result.video ? 300 : 120);   // 최대 5분마다 새 영상 확인
  return result;
}

function resolveChannelId_(text) {
  var direct = /(UC[\w-]{22})/.exec(text);
  if (direct) return direct[1];
  var m = /@([^\/?#\s]+)/.exec(text);
  var page;   // 채널 첫 화면 주소
  if (m) page = 'https://www.youtube.com/@' + encodeURIComponent(decodeURIComponent(m[1]));
  else if (/^https?:\/\/(www\.|m\.)?youtube\.com\//i.test(text)) page = text.replace(/[?#].*$/, '');   // 예전 /c/이름, /user/이름 주소
  else page = 'https://www.youtube.com/@' + encodeURIComponent(text);
  var props = PropertiesService.getScriptProperties();
  var key = 'channelId:' + page.toLowerCase();
  var saved = props.getProperty(key);
  if (saved) return saved;
  var res = UrlFetchApp.fetch(page, {
    muteHttpExceptions: true, headers: { 'Accept-Language': 'ko-KR,ko;q=0.9' }
  });
  if (res.getResponseCode() !== 200) return '';
  var html = res.getContentText();
  var found = /feeds\/videos\.xml\?channel_id=(UC[\w-]{22})/.exec(html) ||
              /"externalId":"(UC[\w-]{22})"/.exec(html) ||
              /<link rel="canonical" href="https:\/\/www\.youtube\.com\/channel\/(UC[\w-]{22})"/.exec(html);
  if (!found) return '';
  props.setProperty(key, found[1]);
  return found[1];
}

function pickVideo_(xml, keywordText) {
  var atom = XmlService.getNamespace('http://www.w3.org/2005/Atom');
  var yt = XmlService.getNamespace('yt', 'http://www.youtube.com/xml/schemas/2015');
  var words = String(keywordText || '').split(/[,，]/).map(normLabel_).filter(String).map(function (w) { return w.toLowerCase(); });
  var entries = XmlService.parse(xml).getRootElement().getChildren('entry', atom);
  var list = [];
  entries.forEach(function (en) {
    var title = en.getChildText('title', atom) || '';
    var id = en.getChildText('videoId', yt) || '';
    if (!id) return;
    var t = normLabel_(title).toLowerCase();
    if (words.length && !words.some(function (w) { return t.indexOf(w) !== -1; })) return;
    list.push({ id: id, title: title, url: 'https://www.youtube.com/watch?v=' + id, published: en.getChildText('published', atom) || '' });
  });
  list.sort(function (a, b) { return a.published < b.published ? 1 : a.published > b.published ? -1 : 0; });
  return list[0] || null;
}

// '주보' 탭을 고치면 바로 앱에 보이도록 저장해 둔 사본을 지웁니다.
function onEdit(e) {
  try {
    if (e && e.range && e.range.getSheet().getName() === BULLETIN_SHEET) CacheService.getScriptCache().remove(BULLETIN_CACHE);
  } catch (err) {}
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('스마트 주보')
    .addItem('① 주보 입력 탭 만들기', 'setupBulletinSheet')
    .addItem('② 유튜브 연결 확인하기', 'testYoutubeLink')
    .addToUi();
}

function setupBulletinSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  if (ss.getSheetByName(BULLETIN_SHEET)) {
    ss.setActiveSheet(ss.getSheetByName(BULLETIN_SHEET));
    ui.alert("'" + BULLETIN_SHEET + "' 탭이 이미 있습니다. 그 탭의 '내용' 칸을 고쳐 쓰시면 됩니다.");
    return;
  }
  var sheet = ss.insertSheet(BULLETIN_SHEET, 0);
  var data = [['항목', '내용', '쓰는 방법']].concat(BULLETIN_ROWS.map(function (r) { return [r[0], presetValue_(r[1], r[2]), r[3]]; }));
  sheet.getRange(2, 2, BULLETIN_ROWS.length, 1).setNumberFormat('@');   // 날짜가 숫자로 바뀌지 않게 글자로 저장
  sheet.getRange(1, 1, data.length, 3).setValues(data).setWrap(true).setVerticalAlignment('top');
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#1E1F4B').setFontColor('#FFFFFF');
  sheet.getRange(2, 1, BULLETIN_ROWS.length, 1).setFontWeight('bold').setBackground('#F1F5F9');
  sheet.getRange(2, 2, BULLETIN_ROWS.length, 1).setBackground('#FFFDF5');
  sheet.getRange(2, 3, BULLETIN_ROWS.length, 1).setFontColor('#64748B').setFontSize(9);
  var firstSetting = BULLETIN_ROWS.length - 1;   // 설정 두 줄은 색을 달리
  sheet.getRange(firstSetting + 1, 1, 2, 3).setBackground('#FEF3C7');
  sheet.setColumnWidth(1, 150); sheet.setColumnWidth(2, 420); sheet.setColumnWidth(3, 320);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, data.length, 1).protect().setDescription('항목 이름을 바꾸면 주보 앱이 그 줄을 못 읽습니다').setWarningOnly(true);
  CacheService.getScriptCache().remove(BULLETIN_CACHE);
  ui.alert("'" + BULLETIN_SHEET + "' 탭을 만들었습니다.\n\n가운데 '내용' 칸에 이번 주 내용을 쓰면 주보 앱에 바로 보입니다.\n(내용 칸이 모두 비어 있는 동안에는 지금 주보 내용이 그대로 보입니다.)\n맨 아래 [유튜브 채널] 칸에 교회 채널 주소를 넣고, 메뉴의 [② 유튜브 연결 확인하기]를 눌러 보세요.");
}

function testYoutubeLink() {
  var ui = SpreadsheetApp.getUi();
  var rows = readBulletinSheet_();
  if (!rows) { ui.alert("먼저 메뉴의 [① 주보 입력 탭 만들기]를 눌러 주세요."); return; }
  var found = findLatestVideo_(rows['#channel'], rows['#keyword'], true);
  CacheService.getScriptCache().remove(BULLETIN_CACHE);
  if (found.video) {
    ui.alert('연결되었습니다 ✅\n\n주보의 [주일 예배] 버튼을 누르면 이 영상이 열립니다.\n\n' + found.video.title + '\n' + found.video.url);
  } else {
    ui.alert('아직 연결되지 않았습니다\n\n' + found.note);
  }
}
