// =====================================================
// 하연통합과학학원 학습도우미 - Apps Script 최종판
// ★ 아래 정보를 입력하세요 ★
// =====================================================
var TEACHER_EMAIL = '여기에_선생님_이메일_입력@gmail.com';

// ★ OneSignal 설정 (onesignal.com 가입 후 입력) ★
var ONESIGNAL_APP_ID  = '45b6f78e-327a-4150-b39a-9a1e0b972c03';
var ONESIGNAL_REST_KEY = 'foav2j6she6u4ch3gj5h25nfh';

function sendPushToTeacher(title, message) {
  if(ONESIGNAL_APP_ID.indexOf('여기에') >= 0) return;
  try {
    UrlFetchApp.fetch('https://onesignal.com/api/v1/notifications', {
      method: 'post',
      headers: {'Content-Type':'application/json','Authorization':'Basic '+ONESIGNAL_REST_KEY},
      payload: JSON.stringify({
        app_id: ONESIGNAL_APP_ID,
        filters: [{"field":"tag","key":"role","relation":"=","value":"teacher"}],
        headings: {"ko": title}, contents: {"ko": message}
      }),
      muteHttpExceptions: true
    });
  } catch(e) {}
}

function sendPushToStudent(studentName, title, message) {
  if(ONESIGNAL_APP_ID.indexOf('여기에') >= 0) return;
  try {
    UrlFetchApp.fetch('https://onesignal.com/api/v1/notifications', {
      method: 'post',
      headers: {'Content-Type':'application/json','Authorization':'Basic '+ONESIGNAL_REST_KEY},
      payload: JSON.stringify({
        app_id: ONESIGNAL_APP_ID,
        filters: [{"field":"tag","key":"name","relation":"=","value":studentName}],
        headings: {"ko": title}, contents: {"ko": message}
      }),
      muteHttpExceptions: true
    });
  } catch(e) {}
}

var SHEET_Q  = '질문게시판';
var SHEET_HW = '숙제오답';
var SHEET_STU = '학생명부';
var SHEET_VID = '풀이영상';
var SHEET_NTC = '공지사항';
var HEADERS_Q = ['id','name','grade','week','unit','question','ts','answered','reply','replyTs','hasImg','imgData','replyImgData','confirmedBy','comments','replyVideoUrl'];
var HEADERS_HW = ['id','name','grade','week','unit','total','wrong','ts'];
var HEADERS_STU = ['id','name','grade','phone','parentPhone','subject','memo','ts','status'];
var HEADERS_VID = ['id','grade','title','url','desc','ts'];
var HEADERS_NTC = ['id','grade','title','content','important','ts'];

function getSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,headers.length).setBackground('#1a73e8').setFontColor('#fff').setFontWeight('bold');
  }
  return sheet;
}

function cors(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var action = e.parameter.action;
  var result;
  if (action === 'getQuestions') result = getQuestions();
  else if (action === 'getHW') result = getHW();
  else if (action === 'getStudents') result = getStudents();
  else if (action === 'getVideos') result = getVideos();
  else if (action === 'getNotices') result = getNotices();
  else result = {error: '알 수 없는 요청'};
  return cors(ContentService.createTextOutput(JSON.stringify(result)));
}

function doPost(e) {
  var data;
  try { data = JSON.parse(e.postData.contents); } catch(err) { return cors(ContentService.createTextOutput(JSON.stringify({error:'잘못된 요청'}))); }
  var action = data.action;
  var result;
  if      (action === 'addQuestion')    result = addQuestion(data);
  else if (action === 'addReply')       result = addReply(data);
  else if (action === 'editReply')      result = editReply(data);
  else if (action === 'deleteQuestion') result = deleteQuestion(data);
  else if (action === 'addHW')          result = addHW(data);
  else if (action === 'confirmReply')   result = confirmReply(data);
  else if (action === 'addComment')     result = addComment(data);
  else if (action === 'deleteComment')  result = deleteComment(data);
  else if (action === 'addStudent')     result = addStudent(data);
  else if (action === 'updateStudent')  result = updateStudent(data);
  else if (action === 'deleteStudent')  result = deleteStudentRow(data);
  else if (action === 'addVideo')       result = addVideo(data);
  else if (action === 'deleteVideo')    result = deleteVideo(data);
  else if (action === 'addNotice')      result = addNotice(data);
  else if (action === 'deleteNotice')   result = deleteNotice(data);
  else result = {error: '알 수 없는 요청'};
  return cors(ContentService.createTextOutput(JSON.stringify(result)));
}

// ── 질문 전체 불러오기 ──
function getQuestions() {
  var sheet = getSheet(SHEET_Q, HEADERS_Q);
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var comments = [];
    var confirmedBy = [];
    try { comments = r[14] ? JSON.parse(String(r[14])) : []; } catch(e) { comments = []; }
    try { confirmedBy = r[13] ? String(r[13]).split(',').filter(function(s){return s.trim()!=='';}) : []; } catch(e) { confirmedBy = []; }
    result.push({
      id: String(r[0]),
      name: String(r[1]||''),
      grade: String(r[2]||''),
      week: String(r[3]||''),
      unit: String(r[4]||''),
      question: String(r[5]||''),
      ts: r[6]||'',
      answered: r[7]===true||String(r[7]).toUpperCase()==='TRUE',
      reply: String(r[8]||''),
      replyTs: r[9]||'',
      hasImg: r[10]===true||String(r[10]).toUpperCase()==='TRUE',
      imgData: String(r[11]||''),
      replyImgData: String(r[12]||''),
      confirmedBy: confirmedBy,
      comments: comments,
      replyVideoUrl: String(r[15]||'')
    });
  }
  result.reverse();
  return result;
}

// ── 구글 드라이브에 이미지 저장 ──
function saveImgToDrive(base64Data, fileName) {
  if (!base64Data || base64Data.length < 10) return '';
  try {
    // base64에서 실제 데이터 추출
    var matches = base64Data.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/);
    if (!matches || matches.length !== 3) return '';
    var mimeType = matches[1];
    var data = Utilities.base64Decode(matches[2]);
    var blob = Utilities.newBlob(data, mimeType, fileName);
    // 드라이브에 저장
    var folder = getOrCreateDriveFolder('하연과학_학생사진');
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    // 바로 볼 수 있는 URL 반환
    return 'https://drive.google.com/uc?export=view&id=' + file.getId();
  } catch(e) {
    return '';
  }
}

function getOrCreateDriveFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

// ── 질문 추가 ──
function addQuestion(data) {
  var sheet = getSheet(SHEET_Q, HEADERS_Q);
  var id = String(Date.now());
  // 이미지를 구글 드라이브에 저장
  var imgUrl = '';
  if (data.imgData && data.imgData.length > 10) {
    imgUrl = saveImgToDrive(data.imgData, 'q_' + id + '.jpg');
  }
  sheet.appendRow([
    id, data.name||'', data.grade||'', data.week||'', data.unit||'',
    data.question||'', data.ts||Date.now(), false, '', '', !!(imgUrl), imgUrl,
    '', '', '[]', ''
  ]);
  // 이메일 + 푸시 알람
  try {
    if (TEACHER_EMAIL && TEACHER_EMAIL.indexOf('@') > 0 && TEACHER_EMAIL.indexOf('여기에') < 0) {
      MailApp.sendEmail({
        to: TEACHER_EMAIL,
        subject: '[하연과학] ' + data.name + '(' + data.grade + ') 학생이 질문을 올렸어요',
        body: '학생: ' + data.name + ' (' + data.grade + ')\n주차: ' + data.week + '\n단원: ' + data.unit + '\n\n질문:\n' + data.question
      });
    }
    sendPushToTeacher(
      '⚗️ 새 질문! — ' + data.name + ' (' + data.grade + ')',
      data.unit + ' | ' + String(data.question).substring(0, 50)
    );
  } catch(e) {}
  return {success: true, id: id};
}

// ── 답변 등록 ──
function addReply(data) {
  var sheet = getSheet(SHEET_Q, HEADERS_Q);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      var studentName = String(rows[i][1]||'');
      var unit = String(rows[i][4]||'');
      // 답변 사진을 구글 드라이브에 저장
      var replyImgUrl = '';
      if (data.replyImgData && data.replyImgData.length > 10) {
        replyImgUrl = saveImgToDrive(data.replyImgData, 'r_' + data.id + '.jpg');
      }
      sheet.getRange(i+1, 8).setValue(true);
      sheet.getRange(i+1, 9).setValue(data.reply||'');
      sheet.getRange(i+1, 10).setValue(data.replyTs||Date.now());
      sheet.getRange(i+1, 13).setValue(replyImgUrl);
      sheet.getRange(i+1, 14).setValue('');
      sheet.getRange(i+1, 16).setValue(data.replyVideoUrl||'');
      // 학생에게 푸시 알림
      sendPushToStudent(
        studentName,
        '⚗️ 선생님 답변 도착!',
        unit + ' 질문에 선생님 답변이 달렸어요!'
      );
      return {success: true};
    }
  }
  return {error: '질문을 찾을 수 없어요'};
}

// ── 답변 수정 ──
function editReply(data) {
  var sheet = getSheet(SHEET_Q, HEADERS_Q);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.getRange(i+1, 9).setValue(data.reply||'');
      sheet.getRange(i+1, 10).setValue(data.replyTs||Date.now());
      sheet.getRange(i+1, 13).setValue(data.replyImgData||'');
      sheet.getRange(i+1, 16).setValue(data.replyVideoUrl||'');
      return {success: true};
    }
  }
  return {error: '질문을 찾을 수 없어요'};
}

// ── 질문 삭제 ──
function deleteQuestion(data) {
  var sheet = getSheet(SHEET_Q, HEADERS_Q);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.deleteRow(i+1);
      return {success: true};
    }
  }
  return {error: '질문을 찾을 수 없어요'};
}

// ── 답변 확인 (학생) ──
function confirmReply(data) {
  var sheet = getSheet(SHEET_Q, HEADERS_Q);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      var confirmed = [];
      try { confirmed = rows[i][13] ? String(rows[i][13]).split(',').filter(function(s){return s.trim()!=='';}) : []; } catch(e) { confirmed = []; }
      if (confirmed.indexOf(data.name) < 0) {
        confirmed.push(data.name);
        sheet.getRange(i+1, 14).setValue(confirmed.join(','));
      }
      return {success: true};
    }
  }
  return {error: '질문을 찾을 수 없어요'};
}

// ── 댓글 추가 ──
function addComment(data) {
  var sheet = getSheet(SHEET_Q, HEADERS_Q);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      var comments = [];
      try { comments = rows[i][14] ? JSON.parse(String(rows[i][14])) : []; } catch(e) { comments = []; }
      comments.push({id: String(Date.now()), name: data.name||'', text: data.text||'', ts: data.ts||Date.now()});
      sheet.getRange(i+1, 15).setValue(JSON.stringify(comments));
      return {success: true};
    }
  }
  return {error: '질문을 찾을 수 없어요'};
}

// ── 댓글 삭제 ──
function deleteComment(data) {
  var sheet = getSheet(SHEET_Q, HEADERS_Q);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.qid)) {
      var comments = [];
      try { comments = rows[i][14] ? JSON.parse(String(rows[i][14])) : []; } catch(e) { comments = []; }
      comments = comments.filter(function(c){ return String(c.id) !== String(data.cid); });
      sheet.getRange(i+1, 15).setValue(JSON.stringify(comments));
      return {success: true};
    }
  }
  return {error: '질문을 찾을 수 없어요'};
}

// ── 숙제 오답 불러오기 ──
function getHW() {
  var sheet = getSheet(SHEET_HW, HEADERS_HW);
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var wrong = [];
    try { wrong = r[6] ? String(r[6]).split(',').map(Number).filter(function(n){return !isNaN(n)&&n>0;}) : []; } catch(e) { wrong = []; }
    result.push({
      id: String(r[0]), name: String(r[1]||''), grade: String(r[2]||''),
      week: String(r[3]||''), unit: String(r[4]||''),
      total: Number(r[5])||0, wrong: wrong, ts: r[7]||''
    });
  }
  return result;
}

// ── 숙제 오답 추가 ──
// ── 숙제 오답 추가 ──
function addHW(data) {
  var sheet = getSheet(SHEET_HW, HEADERS_HW);
  var id = String(Date.now());
  sheet.appendRow([id, data.name||'', data.grade||'', data.week||'', data.unit||'', data.total||0, (data.wrong||[]).join(','), data.ts||Date.now()]);
  // 이메일 알람
  try {
    if (TEACHER_EMAIL && TEACHER_EMAIL.indexOf('@') > 0 && TEACHER_EMAIL.indexOf('여기에') < 0) {
      var wrongList = data.wrong&&data.wrong.length>0 ? data.wrong.join(', ')+'번' : '없음';
      MailApp.sendEmail({
        to: TEACHER_EMAIL,
        subject: '[하연과학] ' + data.name + '(' + data.grade + ') 오답 제출 — ' + data.week,
        body: '학생: ' + data.name + ' (' + data.grade + ')\n주차: ' + data.week + '\n단원: ' + data.unit + '\n틀린 문제: ' + wrongList + '\n오답률: ' + Math.round((data.wrong||[]).length/(data.total||1)*100) + '%'
      });
    }
  } catch(e) {}
  return {success: true, id: id};
}

// ── 학생 명부 ──
function getStudents() {
  var sheet = getSheet(SHEET_STU, HEADERS_STU);
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    result.push({
      id: String(r[0]),
      name: String(r[1]||''),
      grade: String(r[2]||''),
      phone: String(r[3]||''),
      parentPhone: String(r[4]||''),
      subject: String(r[5]||''),
      memo: String(r[6]||''),
      ts: r[7]||'',
      status: String(r[8]||'pending')
    });
  }
  return result;
}

function addStudent(data) {
  var sheet = getSheet(SHEET_STU, HEADERS_STU);
  var id = String(Date.now());
  sheet.appendRow([id, data.name||'', data.grade||'', data.phone||'', data.parentPhone||'', data.subject||'', data.memo||'', data.ts||Date.now(), data.status||'pending']);
  // 선생님에게 신청 알람
  try {
    if (TEACHER_EMAIL && TEACHER_EMAIL.indexOf('@') > 0 && TEACHER_EMAIL.indexOf('여기에') < 0) {
      MailApp.sendEmail({
        to: TEACHER_EMAIL,
        subject: '[하연과학] 새 수강 신청 — ' + data.name + '(' + data.grade + ')',
        body: '이름: ' + data.name + '\n학년: ' + data.grade + '\n연락처: ' + (data.phone||'') + '\n학부모: ' + (data.parentPhone||'') + '\n과목: ' + (data.subject||'') + '\n메모: ' + (data.memo||'')
      });
    }
  } catch(e) {}
  return {success: true, id: id};
}

function updateStudent(data) {
  var sheet = getSheet(SHEET_STU, HEADERS_STU);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      if (data.status !== undefined) sheet.getRange(i+1, 9).setValue(data.status);
      if (data.name !== undefined) sheet.getRange(i+1, 2).setValue(data.name);
      if (data.grade !== undefined) sheet.getRange(i+1, 3).setValue(data.grade);
      if (data.phone !== undefined) sheet.getRange(i+1, 4).setValue(data.phone);
      if (data.parentPhone !== undefined) sheet.getRange(i+1, 5).setValue(data.parentPhone);
      if (data.subject !== undefined) sheet.getRange(i+1, 6).setValue(data.subject);
      if (data.memo !== undefined) sheet.getRange(i+1, 7).setValue(data.memo);
      return {success: true};
    }
  }
  return {error: '학생을 찾을 수 없어요'};
}

function deleteStudentRow(data) {
  var sheet = getSheet(SHEET_STU, HEADERS_STU);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.deleteRow(i+1);
      return {success: true};
    }
  }
  return {error: '학생을 찾을 수 없어요'};
}

// ── 풀이 영상 ──
function getVideos() {
  var sheet = getSheet(SHEET_VID, HEADERS_VID);
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    result.push({
      id: String(r[0]),
      grade: String(r[1]||'전체'),
      title: String(r[2]||''),
      url: String(r[3]||''),
      desc: String(r[4]||''),
      ts: r[5]||''
    });
  }
  result.reverse();
  return result;
}

function addVideo(data) {
  var sheet = getSheet(SHEET_VID, HEADERS_VID);
  var id = String(Date.now());
  sheet.appendRow([id, data.grade||'전체', data.title||'', data.url||'', data.desc||'', data.ts||Date.now()]);
  return {success: true, id: id};
}

function deleteVideo(data) {
  var sheet = getSheet(SHEET_VID, HEADERS_VID);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.deleteRow(i+1);
      return {success: true};
    }
  }
  return {error: '영상을 찾을 수 없어요'};
}

// ── 공지사항 ──
function getNotices() {
  var sheet = getSheet(SHEET_NTC, HEADERS_NTC);
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    result.push({
      id: String(r[0]),
      grade: String(r[1]||'전체'),
      title: String(r[2]||''),
      content: String(r[3]||''),
      important: r[4]===true||String(r[4]).toUpperCase()==='TRUE'?'TRUE':'FALSE',
      ts: r[5]||''
    });
  }
  result.reverse();
  return result;
}

function addNotice(data) {
  var sheet = getSheet(SHEET_NTC, HEADERS_NTC);
  var id = String(Date.now());
  sheet.appendRow([id, data.grade||'전체', data.title||'', data.content||'', data.important||false, data.ts||Date.now()]);
  return {success: true, id: id};
}

function deleteNotice(data) {
  var sheet = getSheet(SHEET_NTC, HEADERS_NTC);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.deleteRow(i+1);
      return {success: true};
    }
  }
  return {error: '공지사항을 찾을 수 없어요'};
}
