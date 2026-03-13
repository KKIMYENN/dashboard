// ============================================================
// 양영학원 확인학습 실시간 대시보드 — Google Apps Script v2
// ============================================================
// 사용법:
// 1. 마스터 스프레드시트에서 [확장 프로그램] > [Apps Script] 열기
// 2. 이 코드를 전체 붙여넣기 → 저장
// 3. 스프레드시트로 돌아가서 [📊 대시보드] > [시트 초기 설정]
// 4. [📊 대시보드] > [트리거 등록]
// 5. [배포] > [새 배포] > 웹 앱 (누구나 접근, 본인으로 실행)
// 6. 배포 URL을 dashboard.html에 입력
// ============================================================

// ---- 상수 ----
var SHEET_STUDENTS = '학생명단';
var SHEET_TASKS    = '할일배정';
var SHEET_SUMMARY  = 'Summary';
var SHEET_PERIODS  = '기간관리';
var SHEET_WEEKLY   = '주간과제';
var CACHE_KEY      = 'dashboard_v8';
var CACHE_TTL      = 300; // 5분 (초)

// ============================================================
// 1. doGet — 대시보드 데이터 JSON 제공
// ============================================================
function doGet(e) {
  try {
    var param = (e && e.parameter) || {};

    // ?action=setup → 초기 시트 구조 확인
    if (param.action === 'setup') {
      setupSheets();
      return jsonResponse({ success: true, message: '시트 구조가 생성되었습니다.' });
    }

    // ?period=기간명 → 특정 기간 조회 (없으면 활성 기간)
    var requestedPeriod = param.period || null;

    var data = getDashboardData(requestedPeriod);
    return jsonResponse(data);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ============================================================
// 2. doPost — 선생님 액션 처리
// ============================================================
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 최대 10초 대기
  } catch (lockErr) {
    return jsonResponse({ error: '서버가 바쁩니다. 잠시 후 다시 시도하세요.' });
  }

  try {
    if (!e.postData || !e.postData.contents) {
      return jsonResponse({ error: '요청 본문이 비어있습니다.' });
    }
    var params = JSON.parse(e.postData.contents);
    var action = params.action;

    if (!action) {
      return jsonResponse({ error: 'action이 필요합니다.' });
    }

    var result;
    switch (action) {
      case '수동완료':
        result = handleManualDone(params);
        break;
      case '수동완료취소':
        result = handleManualUndone(params);
        break;
      case '일괄완료':
        result = handleBulkDone(params);
        break;
      case '점수수정':
        result = handleScoreEdit(params);
        break;
      case '상태변경':
        result = handleStatusChange(params);
        break;
      case '새기간':
        result = handleNewPeriod(params);
        break;
      case '기간전환':
        result = handleSwitchPeriod(params);
        break;
      case '주간과제저장':
        result = handleWeeklyTaskSave(params);
        break;
      case '주간과제삭제':
        result = handleWeeklyTaskDelete(params);
        break;
      case '주간과제조회':
        result = handleWeeklyTaskList();
        break;
      default:
        result = { error: '유효하지 않은 액션: ' + action };
    }

    // 캐시 무효화
    invalidateCache();

    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message });
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 2-1. 수동완료 — Summary에 수동 완료 행 추가
// ============================================================
function handleManualDone(params) {
  var studentName = (params.studentName || '').trim();
  var taskTitle = (params.taskTitle || '').trim();
  if (!studentName || !taskTitle) {
    return { error: 'studentName과 taskTitle이 필요합니다.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSummary(ss);
  var periodName = getActivePeriodName(ss);

  sheet.appendRow([
    new Date(),         // A: 타임스탬프
    studentName,        // B: 이름
    taskTitle,          // C: 퀴즈제목
    100,                // D: 점수
    100,                // E: 만점
    100,                // F: 정답률
    periodName,         // G: 기간명
    '수동'              // H: 수정여부
  ]);

  return { success: true, action: '수동완료', name: studentName, task: taskTitle };
}

// ============================================================
// 2-2. 수동완료취소 — Summary에서 수동 완료 행 삭제
// ============================================================
function handleManualUndone(params) {
  var studentName = (params.studentName || '').trim();
  var taskTitle = (params.taskTitle || '').trim();
  var timestampStr = (params.timestamp || '').trim();
  if (!studentName || !taskTitle) {
    return { error: 'studentName과 taskTitle이 필요합니다.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SUMMARY);
  if (!sheet) return { error: 'Summary 시트를 찾을 수 없습니다.' };

  var data = sheet.getDataRange().getValues();
  // 역순으로 탐색하여 매칭되는 수동 행 삭제
  for (var i = data.length - 1; i >= 1; i--) {
    var tsStr = data[i][0] instanceof Date ? data[i][0].toISOString() : String(data[i][0]);
    var rowName = String(data[i][1]).trim();
    var rowTitle = String(data[i][2]).trim();
    var rowFlag = String(data[i][7]).trim();
    
    if (rowName === studentName && rowTitle === taskTitle && rowFlag === '수동') {
      if (timestampStr && tsStr !== timestampStr) continue;
      sheet.deleteRow(i + 1);
      return { success: true, action: '수동완료취소', name: studentName, task: taskTitle };
    }
  }

  return { success: true, found: false, message: '삭제할 수동완료 기록을 찾지 못했습니다.' };
}

// ============================================================
// 2-3. 일괄완료 — 특정 퀴즈를 배정받은 모든 학생에 수동완료
// NOTE: API 전용 기능. 대시보드 UI에는 버튼 없음 (필요 시 추가 가능)
// ============================================================
function handleBulkDone(params) {
  var taskTitle = (params.taskTitle || '').trim();
  if (!taskTitle) {
    return { error: 'taskTitle이 필요합니다.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = getOrCreateSummary(ss);
  var periodName = getActivePeriodName(ss);

  // 할일배정에서 이 퀴즈가 배정된 대상 파악
  var students = getStudentsForTask(ss, taskTitle);

  // 이미 제출된 학생은 제외
  var existingSubmissions = getExistingSubmissions(ss, taskTitle, periodName);
  var count = 0;

  for (var i = 0; i < students.length; i++) {
    var name = students[i];
    if (!existingSubmissions[name]) {
      summarySheet.appendRow([
        new Date(), name, taskTitle, 100, 100, 100, periodName, '수동'
      ]);
      count++;
    }
  }

  return { success: true, action: '일괄완료', task: taskTitle, count: count };
}

// ============================================================
// 2-4. 점수수정 — Summary에서 최신 행의 점수 수정
// ============================================================
function handleScoreEdit(params) {
  var studentName = (params.studentName || '').trim();
  var taskTitle = (params.taskTitle || '').trim();
  var newScore = Number(params.newScore);
  var newTotal = Number(params.newTotal);
  var timestampStr = (params.timestamp || '').trim();
  if (!studentName || !taskTitle || isNaN(newScore) || isNaN(newTotal) || newTotal <= 0) {
    return { error: 'studentName, taskTitle, newScore, newTotal이 모두 필요합니다.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SUMMARY);
  if (!sheet) return { error: 'Summary 시트를 찾을 수 없습니다.' };

  var data = sheet.getDataRange().getValues();
  // 역순으로 최신 매칭 행 찾기
  for (var i = data.length - 1; i >= 1; i--) {
    var tsStr = data[i][0] instanceof Date ? data[i][0].toISOString() : String(data[i][0]);
    var rowName = String(data[i][1]).trim();
    var rowTitle = String(data[i][2]).trim();
    if (rowName === studentName && rowTitle === taskTitle) {
      if (timestampStr && tsStr !== timestampStr) continue;
      var newPct = Math.round(newScore / newTotal * 100);
      sheet.getRange(i + 1, 4).setValue(newScore);  // D열: 점수
      sheet.getRange(i + 1, 5).setValue(newTotal);  // E열: 만점
      sheet.getRange(i + 1, 6).setValue(newPct);    // F열: 정답률
      sheet.getRange(i + 1, 8).setValue('수정됨');   // H열: 수정여부
      return { success: true, action: '점수수정', name: studentName, task: taskTitle, newScore: newScore, newTotal: newTotal, newPct: newPct };
    }
  }

  return { error: '해당 제출 기록을 찾을 수 없습니다: ' + studentName + ' / ' + taskTitle };
}

// ============================================================
// 2-4b. 상태변경 — 학생명단 D열 상태 변경 (결석/귀가/추가문제/출석)
// ============================================================
function handleStatusChange(params) {
  var studentName = (params.studentName || '').trim();
  var newStatus = (params.newStatus || '').trim();
  if (!studentName || !newStatus) {
    return { error: 'studentName과 newStatus가 필요합니다.' };
  }

  var validStatuses = ['출석', '결석', '귀가', '지각', '추가문제'];
  if (validStatuses.indexOf(newStatus) === -1) {
    return { error: '유효하지 않은 상태: ' + newStatus + ' (출석/결석/귀가/지각/추가문제)' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet) return { error: '학생명단 시트를 찾을 수 없습니다.' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][0] != null ? data[i][0] : '').trim();
    if (name === studentName) {
      sheet.getRange(i + 1, 4).setValue(newStatus); // D열: 상태
      return { success: true, action: '상태변경', name: studentName, status: newStatus };
    }
  }

  return { error: '학생을 찾을 수 없습니다: ' + studentName };
}

// ============================================================
// 2-5. 새기간 — 새 기간 생성 (기존 활성 기간은 종료로 변경)
// ============================================================
function handleNewPeriod(params) {
  var periodName = (params.periodName || '').trim();
  var startDate = (params.startDate || '').trim();
  var endDate = (params.endDate || '').trim();
  if (!periodName) {
    return { error: '기간명이 필요합니다.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreatePeriods(ss);
  var data = sheet.getDataRange().getValues();

  // 기존 활성 기간을 종료로 변경
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim() === '활성') {
      sheet.getRange(i + 1, 4).setValue('종료');
    }
  }

  // 새 기간 추가
  sheet.appendRow([periodName, startDate || '', endDate || '', '활성']);

  return { success: true, action: '새기간', periodName: periodName };
}

// ============================================================
// 2-6. 기간전환 — 활성 기간 변경
// ============================================================
function handleSwitchPeriod(params) {
  var periodName = (params.periodName || '').trim();
  if (!periodName) {
    return { error: '기간명이 필요합니다.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PERIODS);
  if (!sheet) return { error: '기간관리 시트를 찾을 수 없습니다.' };

  var data = sheet.getDataRange().getValues();
  var found = false;

  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][0]).trim();
    if (name === periodName) {
      sheet.getRange(i + 1, 4).setValue('활성');
      found = true;
    } else {
      sheet.getRange(i + 1, 4).setValue('종료');
    }
  }

  if (!found) {
    return { error: '해당 기간을 찾을 수 없습니다: ' + periodName };
  }

  return { success: true, action: '기간전환', periodName: periodName };
}

// ============================================================
// 3. getDashboardData — 핵심 데이터 조합
// ============================================================
function getDashboardData(requestedPeriod) {
  // --- 캐시 확인 (활성 기간만 캐시) ---
  var cache = CacheService.getScriptCache();
  var cacheKey = CACHE_KEY;
  if (requestedPeriod) {
    cacheKey = CACHE_KEY + '_' + requestedPeriod;
  }
  var chunkCountStr = cache.get(cacheKey + '_chunk_count');
  if (chunkCountStr) {
    var chunkCount = parseInt(chunkCountStr, 10);
    var cachedStr = '';
    var isValid = true;
    for (var ci = 0; ci < chunkCount; ci++) {
      var chunk = cache.get(cacheKey + '_chunk_' + ci);
      if (chunk) {
        cachedStr += chunk;
      } else {
        isValid = false;
        break;
      }
    }
    if (isValid && cachedStr) {
      try { return JSON.parse(cachedStr); } catch (_) { /* 파싱 오류 시 무시 */ }
    }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 교과서 매핑 ---
  var TEXTBOOK_MAP = {
    "가오고_1": "공통영어1 미래엔김",
    "관저고_1": "공통영어1 미래엔김",
    "괴정고_1": "공통영어1 천재강",
    "금산고_1": "공통영어1 천재조",
    "노은고_1": "공통영어1 능률오",
    "대성고_1": "공통영어1 능률민",
    "대성고_2": "미래엔김 영어1",
    "대신고_1": "공통영어1 YBM박",
    "대전고_1": "공통영어1 미래엔김",
    "대전여고_1": "공통영어1 미래엔김",
    "대전외고_1": "원서",
    "대전외고_2": "원서",
    "도안고_1": "공통영어1 미래엔김",
    "도안고_2": "미래엔김 영어1",
    "동대전고_1": "공통영어1 YBM김",
    "동산고_1": "공통영어1 비상홍",
    "둔산여고_1": "공통영어1 YBM박",
    "둔원고_1": "공통영어1 YBM김",
    "보문고_1": "공통영어1 YBM박",
    "복수고_1": "공통영어1 YBM박",
    "서대전고_1": "공통영어1 YBM박",
    "서대전고_2": "YBM 박 영어1",
    "세종국제고_": "원서",
    "세종국제고_1": "원서",
    "세종국제고_2": "원서",
    "세종국제고_3": "원서",
    "세종보람고_1": "공통영어1 미래엔김",
    "세종보람고_2": "천재조 영어1",
    "우송고_1": "공통영어1 능률민",
    "유성고_1": "공통영어1 천재조",
    "유성고_2": "미래엔김 영어1",
    "유성여고_2": "비상홍 영어1",
    "이문고_1": "공통영어1 능률민",
    "전민고_1": "공통영어1 미래엔김",
    "전민고_2": "지학사신 영어1",
    "중앙고_2": "능률오 영어1",
    "지족고_1": "공통영어1 비상홍",
    "충남고_1": "공통영어1 YBM김",
    "충남고_2": "능률오 영어1",
    "충남여고_1": "공통영어1 YBM박",
    "충남여고_2": "YBM 박 영어1",
    "한밭고_1": "공통영어1 비상홍",
    "호수돈여고_1": "공통영어1 비상홍"
  };

  // --- 기간 정보 ---
  var periodsInfo = getPeriodsInfo(ss);
  var activePeriod = requestedPeriod || periodsInfo.activePeriod || '';

  // --- 학생명단 ---
  var studentsSheet = ss.getSheetByName(SHEET_STUDENTS);
  var studentsRaw = studentsSheet ? studentsSheet.getDataRange().getValues() : [];
  // [이름, 학교, 학년, 상태]

  // --- 할일배정 ---
  var tasksSheet = ss.getSheetByName(SHEET_TASKS);
  var tasksRaw = tasksSheet ? tasksSheet.getDataRange().getValues() : [];
  // [학교, 학년, 퀴즈제목, 폼링크, 폼ID, 대상학생]

  // --- Summary ---
  var summarySheet = ss.getSheetByName(SHEET_SUMMARY);
  var summaryRaw = summarySheet ? summarySheet.getDataRange().getValues() : [];
  // [타임스탬프, 이름, 퀴즈제목, 점수, 만점, 정답률, 기간명, 수정여부]

  // --- 학교+학년 → 할일 목록 맵 + 개별 학생 할일 ---
  var taskMap = {};       // key: "학교_학년" → [{title, link}]
  var studentTaskMap = {}; // key: "학생이름" → [{title, link}]

  for (var i = 1; i < tasksRaw.length; i++) {
    var row = tasksRaw[i];
    var tSchool = String(row[0] != null ? row[0] : '').trim();
    
    var rawTGrade = String(row[1] != null ? row[1] : '').trim();
    var tGrade = rawTGrade.replace(/[^0-9]/g, '');
    // If the spreadsheet actually contained no numbers, fallback to original to avoid empty string
    if (!tGrade) tGrade = rawTGrade;
    
    var tTitle = String(row[2] != null ? row[2] : '').trim();
    var tLink = String(row[3] != null ? row[3] : '').trim();
    var tTargets = String(row[5] != null ? row[5] : '').trim();

    // B2 fix: 빈 퀴즈제목 건너뛰기
    if (!tTitle) continue;

    var taskObj = { title: tTitle, link: tLink || '' };

    if (tTargets) {
      // F열(대상학생) 지정 → 특정 학생에게만 배정
      var names = tTargets.split(',');
      for (var n = 0; n < names.length; n++) {
        var targetName = names[n].trim();
        if (!targetName) continue;
        if (!studentTaskMap[targetName]) studentTaskMap[targetName] = [];
        studentTaskMap[targetName].push(taskObj);
      }
    } else if (tSchool) {
      // 학교+학년 지정 → 해당 학생 전원
      var key = tSchool + '_' + tGrade;
      if (!taskMap[key]) taskMap[key] = [];
      taskMap[key].push(taskObj);
    }
  }

  // --- 주간과제 전체 불러오기 ---
  var allWeeklyTasks = getWeeklyTasksData(ss);

  // --- 기간별 히스토리 맵 (활성 기간 외 종료된 기간들) ---
  var historyMap = {}; // key: 학생이름 → { 기간명: { sum: 0, count: 0 } }

  // --- 유효한 학생 명단 Set 생성 ---
  var validStudentNamesSet = {};
  for (var pi = 1; pi < studentsRaw.length; pi++) {
    var pRow = studentsRaw[pi];
    var pName = String(pRow[0] != null ? pRow[0] : '').trim();
    if (pName) validStudentNamesSet[pName] = true;
  }

  // --- 학생별 제출 기록 맵 (기간 필터 + 전체 이력) ---
  var submissionMap = {}; // key: 학생이름 → { 퀴즈제목: { history: [...], latest: {...} } }
  // --- 미매칭 제출 기록 맵 ---
  var unmatchedSubmissionMap = {}; // key: 학생이름 → { 퀴즈제목: { history: [...], latest: {...} } }

  for (var si = 1; si < summaryRaw.length; si++) {
    var sRow = summaryRaw[si];
    var ts = sRow[0];
    var sName = String(sRow[1] != null ? sRow[1] : '').trim();  // B1 fix
    var sFormTitle = String(sRow[2] != null ? sRow[2] : '').trim();
    var sScore = Number(sRow[3]) || 0;
    var sTotal = Number(sRow[4]) || 0;
    // B9 fix: total=0일 때 pct를 null로
    var sPct = sTotal > 0 ? (Number(sRow[5]) || Math.round(sScore / sTotal * 100)) : null;
    var sPeriod = String(sRow[6] != null ? sRow[6] : '').trim();
    var sEdited = String(sRow[7] != null ? sRow[7] : '').trim();

    if (!sName || !sFormTitle) continue;

    // 기간 필터: 활성 기간이 설정된 경우, 해당 기간의 데이터만 포함
    if (activePeriod && sPeriod && sPeriod !== activePeriod) {
      if (validStudentNamesSet[sName] && sPct !== null) {
        if (!historyMap[sName]) historyMap[sName] = {};
        if (!historyMap[sName][sPeriod]) historyMap[sName][sPeriod] = { sum: 0, count: 0 };
        historyMap[sName][sPeriod].sum += sPct;
        historyMap[sName][sPeriod].count++;
      }
      continue;
    }
    // 기간이 아직 없는 기존 데이터는 포함 (sPeriod가 비어있으면 통과)

    var entry = {
      timestamp: ts instanceof Date ? ts.toISOString() : String(ts),
      score: sScore,
      total: sTotal,
      pct: sPct,
      edited: sEdited === '수정됨',
      manual: sEdited === '수동',
      pending: sEdited === '점수대기'
    };
    var entryTime = ts instanceof Date ? ts.getTime() : new Date(ts).getTime();

    if (validStudentNamesSet[sName]) {
      if (!submissionMap[sName]) submissionMap[sName] = {};
      if (!submissionMap[sName][sFormTitle]) {
        submissionMap[sName][sFormTitle] = { history: [], latest: null };
      }

      submissionMap[sName][sFormTitle].history.push(entry);

      // 최신 제출 추적
      var current = submissionMap[sName][sFormTitle].latest;
      if (!current || entryTime > (current._time || 0)) {
        entry._time = entryTime;
        submissionMap[sName][sFormTitle].latest = entry;
      }
    } else {
      // 미매칭 제출 기록 수집
      if (!unmatchedSubmissionMap[sName]) unmatchedSubmissionMap[sName] = {};
      if (!unmatchedSubmissionMap[sName][sFormTitle]) {
        unmatchedSubmissionMap[sName][sFormTitle] = { history: [], latest: null };
      }

      unmatchedSubmissionMap[sName][sFormTitle].history.push(entry);
      var currentUnmatched = unmatchedSubmissionMap[sName][sFormTitle].latest;
      if (!currentUnmatched || entryTime > (currentUnmatched._time || 0)) {
        entry._time = entryTime;
        unmatchedSubmissionMap[sName][sFormTitle].latest = entry;
      }
    }
  }

  // --- 학생 데이터 조합 ---
  var students = [];
  var unmatchedSubmissions = [];
  var schoolsSet = {};
  var gradesSet = {};

  for (var pi = 1; pi < studentsRaw.length; pi++) {
    var pRow = studentsRaw[pi];
    var pName = String(pRow[0] != null ? pRow[0] : '').trim();  // B1 fix
    var pSchool = String(pRow[1] != null ? pRow[1] : '').trim();
    
    var rawGrade = String(pRow[2] != null ? pRow[2] : '').trim();
    var pGrade = rawGrade.replace(/[^0-9]/g, '');
    if (!pGrade) pGrade = rawGrade;
    
    var pStatus = String(pRow[3] != null ? pRow[3] : '').trim() || '출석';
    if (!pName) continue;

    // 학교/학년 목록 수집
    if (pSchool) schoolsSet[pSchool] = true;
    if (pGrade) gradesSet[pGrade] = true;

    // 이 학생에게 배정된 할일 조합 (학교+학년 + 개별 배정)
    var groupKey = pSchool + '_' + pGrade;
    var assignedTasks = [];
    var seenTitles = {};

    // 학교+학년 기반 할일
    var groupTasks = taskMap[groupKey] || [];
    for (var gi = 0; gi < groupTasks.length; gi++) {
      if (!seenTitles[groupTasks[gi].title]) {
        assignedTasks.push(groupTasks[gi]);
        seenTitles[groupTasks[gi].title] = true;
      }
    }

    // 개별 학생 지정 할일
    var personalTasks = studentTaskMap[pName] || [];
    for (var pti = 0; pti < personalTasks.length; pti++) {
      if (!seenTitles[personalTasks[pti].title]) {
        assignedTasks.push(personalTasks[pti]);
        seenTitles[personalTasks[pti].title] = true;
      }
    }

    // --- 주간과제 자동 연동 ---
    for (var wTag in allWeeklyTasks) {
      if (allWeeklyTasks.hasOwnProperty(wTag)) {
        var weekList = allWeeklyTasks[wTag];
        for (var wl = 0; wl < weekList.length; wl++) {
          var wItem = weekList[wl];
          if (wItem.schools.indexOf('전체') !== -1 || wItem.schools.indexOf(pSchool) !== -1) {
            if (!seenTitles[wItem.label]) {
              assignedTasks.push({ title: wItem.label, link: wItem.link || '' });
              seenTitles[wItem.label] = true;
            }
          }
        }
      }
    }

    var subs = submissionMap[pName] || {};
    var tasks = [];

    for (var ai = 0; ai < assignedTasks.length; ai++) {
      var at = assignedTasks[ai];
      var subData = subs[at.title];
      if (subData && subData.latest) {
        var lt = subData.latest;
        tasks.push({
          title: at.title,
          link: at.link,
          done: true,
          score: lt.score,
          total: lt.total,
          pct: lt.pct,
          edited: lt.edited || false,
          manual: lt.manual || false,
          pending: lt.pending || false,
          attempts: subData.history.length,
          history: subData.history.map(function(h) {
            return {
              timestamp: h.timestamp,
              score: h.score,
              total: h.total,
              pct: h.pct,
              edited: h.edited,
              manual: h.manual,
              pending: h.pending || false
            };
          })
        });
      } else {
        tasks.push({
          title: at.title,
          link: at.link,
          done: false,
          attempts: 0
        });
      }
    }

    var doneCount = 0;
    var doneScoreSum = 0;
    var doneScoreCount = 0;
    for (var di = 0; di < tasks.length; di++) {
      if (tasks[di].done) {
        doneCount++;
        if (tasks[di].pct !== null && tasks[di].pct !== undefined) {
          doneScoreSum += tasks[di].pct;
          doneScoreCount++;
        }
      }
    }
    var avgPct = doneScoreCount > 0 ? Math.round(doneScoreSum / doneScoreCount) : 0;

    var textbook = TEXTBOOK_MAP[pSchool + '_' + pGrade] || TEXTBOOK_MAP[pSchool + '_'] || '';

    // 기간별 히스토리 배열 생성 (종료된 기간들, 기간명 사전순)
    var studentHistory = [];
    if (historyMap[pName]) {
      var hPeriods = Object.keys(historyMap[pName]).sort();
      for (var hp = 0; hp < hPeriods.length; hp++) {
        var hd = historyMap[pName][hPeriods[hp]];
        studentHistory.push({
          period: hPeriods[hp],
          avgPct: hd.count > 0 ? Math.round(hd.sum / hd.count) : 0,
          count: hd.count
        });
      }
    }

    students.push({
      name: pName,
      school: pSchool,
      grade: pGrade,
      status: pStatus,
      tasks: tasks,
      doneCount: doneCount,
      totalCount: tasks.length,
      avgPct: avgPct,
      textbook: textbook,
      history: studentHistory
    });
  }

  // --- 미매칭 제출 감지 ---
  for (var subName in unmatchedSubmissionMap) {
    if (unmatchedSubmissionMap.hasOwnProperty(subName)) {
      var subTitles = Object.keys(unmatchedSubmissionMap[subName]);
      var firstSub = unmatchedSubmissionMap[subName][subTitles[0]];
      unmatchedSubmissions.push({
        name: subName,
        titles: subTitles,
        timestamp: firstSub && firstSub.latest ? firstSub.latest.timestamp : ''
      });
    }
  }

  // --- 통계 ---
  var stats = {
    total: students.length,
    allDone: 0,
    inProgress: 0,
    notStarted: 0,
    noTasks: 0,
    avgScore: 0,
    above90: 0,
    below70: 0,
    absent: 0,
    dismissed: 0,
    late: 0
  };

  var scoredStudents = [];
  for (var sti = 0; sti < students.length; sti++) {
    var st = students[sti];

    // 상태 카운트
    if (st.status === '결석') stats.absent++;
    else if (st.status === '귀가') stats.dismissed++;
    else if (st.status === '지각') stats.late++;

    if (st.totalCount === 0) {
      stats.noTasks++;
    } else if (st.doneCount === st.totalCount) {
      stats.allDone++;
    } else if (st.doneCount > 0) {
      stats.inProgress++;
    } else {
      stats.notStarted++;
    }
    if (st.doneCount > 0) {
      scoredStudents.push(st.avgPct);
      if (st.avgPct >= 90) stats.above90++;
      if (st.avgPct < 70) stats.below70++;
    }
  }

  if (scoredStudents.length > 0) {
    var totalScore = 0;
    for (var sci = 0; sci < scoredStudents.length; sci++) {
      totalScore += scoredStudents[sci];
    }
    stats.avgScore = Math.round(totalScore / scoredStudents.length);
  }

  // --- 결과 조합 ---
  var result = {
    students: students,
    stats: stats,
    unmatchedSubmissions: unmatchedSubmissions,
    lastUpdate: new Date().toISOString(),
    periods: periodsInfo.periods,
    activePeriod: activePeriod,
    schools: Object.keys(schoolsSet).sort(),
    grades: Object.keys(gradesSet).sort(),
    weeklyTasks: allWeeklyTasks
  };

  // --- 캐시 저장 (Chunking) ---
  try {
    var jsonStr = JSON.stringify(result);
    var chunkSize = 90000; // 안전 마진 (100KB 한도)
    var chunks = [];
    for (var i = 0; i < jsonStr.length; i += chunkSize) {
      chunks.push(jsonStr.substring(i, i + chunkSize));
    }
    cache.put(cacheKey + '_chunk_count', chunks.length.toString(), CACHE_TTL);
    for (var j = 0; j < chunks.length; j++) {
      cache.put(cacheKey + '_chunk_' + j, chunks[j], CACHE_TTL);
    }
  } catch (_) { /* 캐시 저장 실패 무시 */ }

  return result;
}

// ============================================================
// 4. onFormSubmitHandler — 폼 제출 시 Summary에 기록
// ============================================================
function onFormSubmitHandler(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (lockErr) {
    Logger.log('onFormSubmit Lock 실패: ' + lockErr.message);
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var summarySheet = getOrCreateSummary(ss);

    // 어떤 시트에 응답이 기록되었는지 확인
    var responseSheet = e.range.getSheet();
    var sheetName = responseSheet.getName();

    // B11 fix: 구글 퀴즈 채점 완료 대기 (점수 컬럼이 아직 비어있을 수 있음)
    Utilities.sleep(3000);

    // 헤더 행 읽기 — 딜레이 후 다시 읽기
    var lastCol = responseSheet.getLastColumn();
    var headers = responseSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var responseRow = e.range.getRow();
    var values = responseSheet.getRange(responseRow, 1, 1, lastCol).getValues()[0];

    // 이름 찾기
    var studentName = '';
    var score = 0;
    var total = 0;

    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] != null ? headers[c] : '').trim().toLowerCase();  // B10 fix: null 체크 + trim

      // 이름 컬럼
      if (h === '이름' || h === 'name' || h === '성명') {
        // B1 fix: undefined/null 체크
        if (values[c] != null && String(values[c]).trim().length > 0) {
          studentName = String(values[c]).trim();
        }
      }

      // 점수 컬럼 (구글폼 퀴즈 모드)
      if (h === '점수' || h === 'score' || h === '총점') {
        var scoreStr = String(values[c] != null ? values[c] : '');
        // "85 / 100" 형식 처리
        if (scoreStr.indexOf('/') !== -1) {
          var parts = scoreStr.split('/');
          score = Number(parts[0].trim()) || 0;
          total = Number(parts[1].trim()) || 0;
        } else {
          score = Number(values[c]) || 0;
        }
      }
    }

    // 이름을 못 찾았으면 타임스탬프/점수 아닌 첫 번째 컬럼 시도
    if (!studentName) {
      for (var c2 = 0; c2 < headers.length; c2++) {
        var h2 = String(headers[c2] != null ? headers[c2] : '').trim().toLowerCase();
        if (h2 !== 'timestamp' && h2 !== '타임스탬프' && h2 !== '점수' && h2 !== 'score'
            && h2 !== '이메일' && h2 !== 'email address' && h2 !== '이메일 주소') {
          if (values[c2] != null) {
            var candidate = String(values[c2]).trim();
            if (candidate.length > 0 && candidate.length < 20) {
              studentName = candidate;
              break;
            }
          }
        }
      }
    }

    // 이름을 여전히 못 찾으면 기록하지 않음
    if (!studentName) {
      Logger.log('이름을 찾을 수 없음: ' + sheetName + ' row ' + responseRow);
      return;
    }

    // B8 fix: 5초 내 중복 제출 방지
    var summaryData = summarySheet.getDataRange().getValues();
    var now = new Date();
    for (var di = summaryData.length - 1; di >= Math.max(1, summaryData.length - 10); di--) {
      var dRow = summaryData[di];
      var dName = String(dRow[1] != null ? dRow[1] : '').trim();
      var dTitle = String(dRow[2] != null ? dRow[2] : '').trim();
      var dTime = dRow[0] instanceof Date ? dRow[0] : new Date(dRow[0]);
      if (dName === studentName && (now.getTime() - dTime.getTime()) < 5000) {
        // 같은 시트이름 기반 매칭도 확인
        if (sheetName.indexOf(dTitle) !== -1 || dTitle.indexOf(sheetName) !== -1) {
          Logger.log('5초 내 중복 제출 무시: ' + studentName + ' / ' + sheetName);
          return;
        }
      }
    }

    // B12 fix: 폼 제목 매칭 — 3단계 전략
    // 1단계: 연결된 Google Form의 실제 제목 가져오기
    var formTitle = sheetName;
    var actualFormTitle = '';
    try {
      var formUrl = responseSheet.getFormUrl();
      if (formUrl) {
        var linkedForm = FormApp.openByUrl(formUrl);
        actualFormTitle = linkedForm.getTitle().trim();
        Logger.log('연결된 폼 제목: ' + actualFormTitle);
      }
    } catch (formErr) {
      Logger.log('FormApp 접근 실패: ' + formErr.message);
    }

    // 2단계: 할일배정에서 매칭 시도 (폼 제목 우선, 시트명 폴백)
    var tasksSheet2 = ss.getSheetByName(SHEET_TASKS);
    if (tasksSheet2) {
      var tasksData = tasksSheet2.getDataRange().getValues();
      var matched = false;

      // 2a: 실제 폼 제목으로 매칭
      if (actualFormTitle) {
        var formTitleNorm = actualFormTitle.toLowerCase().replace(/\s+/g, '');
        for (var ti = 1; ti < tasksData.length; ti++) {
          var taskTitle = String(tasksData[ti][2] != null ? tasksData[ti][2] : '').trim();
          if (!taskTitle) continue;
          var taskNorm = taskTitle.toLowerCase().replace(/\s+/g, '');
          if (formTitleNorm === taskNorm || formTitleNorm.indexOf(taskNorm) !== -1 || taskNorm.indexOf(formTitleNorm) !== -1) {
            formTitle = taskTitle;
            matched = true;
            break;
          }
        }
        // 정확히 일치하는 할일배정이 없어도 폼 제목 자체를 사용
        if (!matched) {
          formTitle = actualFormTitle;
        }
      }

      // 2b: 폼 제목을 못 가져왔으면 시트명으로 폴백
      if (!actualFormTitle) {
        var sheetNameNorm = sheetName.trim().toLowerCase().replace(/\s+/g, '');
        for (var ti = 1; ti < tasksData.length; ti++) {
          var taskTitle = String(tasksData[ti][2] != null ? tasksData[ti][2] : '').trim();
          if (!taskTitle) continue;
          var taskNorm = taskTitle.toLowerCase().replace(/\s+/g, '');
          if (sheetNameNorm === taskNorm || sheetNameNorm.indexOf(taskNorm) !== -1 || taskNorm.indexOf(sheetNameNorm) !== -1) {
            formTitle = taskTitle;
            break;
          }
        }
      }
    } else if (actualFormTitle) {
      formTitle = actualFormTitle;
    }

    // B13 fix: 점수를 못 읽었으면 최대 10초(2초 x 5번)까지 대기하며 재시도
    var retryCount = 0;
    var maxRetries = 5;
    while (total === 0 && retryCount < maxRetries) {
      Logger.log('점수 컬럼 비어있음 — 2초 후 재시도 (' + (retryCount+1) + '/' + maxRetries + ')');
      Utilities.sleep(2000);
      values = responseSheet.getRange(responseRow, 1, 1, lastCol).getValues()[0];
      for (var c3 = 0; c3 < headers.length; c3++) {
        var h3 = String(headers[c3] != null ? headers[c3] : '').trim().toLowerCase();
        if (h3 === '점수' || h3 === 'score' || h3 === '총점') {
          var scoreStr3 = String(values[c3] != null ? values[c3] : '');
          if (scoreStr3.indexOf('/') !== -1) {
            var parts3 = scoreStr3.split('/');
            score = Number(parts3[0].trim()) || 0;
            total = Number(parts3[1].trim()) || 0;
          } else if (Number(values[c3]) > 0) {
            score = Number(values[c3]);
          }
        }
      }
      retryCount++;
    }

    // B14 fix: 여전히 total=0이면 FormApp에서 퀴즈 총점 가져오기
    if (total === 0) {
      try {
        var formUrl2 = responseSheet.getFormUrl();
        if (formUrl2) {
          var form2 = FormApp.openByUrl(formUrl2);
          if (form2.isQuiz()) {
            var responses = form2.getResponses();
            if (responses.length > 0) {
              var lastResp = responses[responses.length - 1];
              var gradableItems = lastResp.getGradableItemResponses();
              var calcScore = 0, calcTotal = 0;
              for (var gi = 0; gi < gradableItems.length; gi++) {
                calcScore += gradableItems[gi].getScore();
                try {
                  var origItem = gradableItems[gi].getItem();
                  var iType = origItem.getType();
                  if (iType == FormApp.ItemType.MULTIPLE_CHOICE) calcTotal += origItem.asMultipleChoiceItem().getPoints();
                  else if (iType == FormApp.ItemType.CHECKBOX) calcTotal += origItem.asCheckboxItem().getPoints();
                  else if (iType == FormApp.ItemType.LIST) calcTotal += origItem.asListItem().getPoints();
                  else if (iType == FormApp.ItemType.TEXT) calcTotal += origItem.asTextItem().getPoints();
                } catch(ptErr) {}
              }
              if (calcTotal > 0) {
                score = calcScore;
                total = calcTotal;
                Logger.log('FormApp에서 점수 복원: ' + score + '/' + total);
              }
            }
          }
        }
      } catch (formScoreErr) {
        Logger.log('FormApp 점수 조회 실패: ' + formScoreErr.message);
      }
    }

    // B9 fix: 정답률 — total=0이면 빈 값
    var pct = total > 0 ? Math.round(score / total * 100) : '';

    // 현재 활성 기간명
    var periodName = getActivePeriodName(ss);

    // Summary에 기록
    summarySheet.appendRow([
      new Date(),       // A: 타임스탬프
      studentName,      // B: 이름
      formTitle,        // C: 퀴즈제목
      score,            // D: 점수
      total,            // E: 만점
      pct,              // F: 정답률
      periodName,       // G: 기간명
      total === 0 ? '점수대기' : ''  // H: 수정여부 (점수 미확인 시 보정 대기)
    ]);

    // 캐시 무효화
    invalidateCache();

    Logger.log('폼 제출 기록: ' + studentName + ' / ' + formTitle + ' / ' + score + '/' + total);
  } catch (err) {
    Logger.log('onFormSubmit 오류: ' + err.message);
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 5. registerAllTriggers — 스프레드시트 제출 트리거 등록
// ============================================================
function registerAllTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 기존 트리거 모두 삭제 (onFormSubmit + fixMissingScores)
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'onFormSubmitHandler' || fn === 'fixMissingScores') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 스프레드시트 레벨 onFormSubmit 트리거 1개 등록
  ScriptApp.newTrigger('onFormSubmitHandler')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  // 점수 보정 트리거 — 5분마다
  ScriptApp.newTrigger('fixMissingScores')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('트리거 등록 완료: onFormSubmitHandler + fixMissingScores');
  try {
    SpreadsheetApp.getUi().alert(
      '트리거가 등록되었습니다.\n' +
      '이 스프레드시트에 연결된 모든 구글폼의 제출이\n' +
      '자동으로 Summary에 기록됩니다.\n' +
      '(점수 미확인 건은 2분 이내 자동 보정)'
    );
  } catch (_) { /* 스크립트 에디터에서 실행 시 UI 없음 — 무시 */ }
}

// ============================================================
// 5-1. fixMissingScores — 점수대기 행 자동 보정 (2분마다 실행)
// ============================================================
function fixMissingScores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SUMMARY);
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();

  // 할일배정에서 퀴즈제목 → 폼URL 맵 구성
  var formMap = {};
  var tasksSheet = ss.getSheetByName(SHEET_TASKS);
  if (tasksSheet) {
    var tasksData = tasksSheet.getDataRange().getValues();
    for (var i = 1; i < tasksData.length; i++) {
      var title = String(tasksData[i][2] || '').trim();
      var link  = String(tasksData[i][3] || '').trim();
      if (title && link && !formMap[title]) formMap[title] = link;
    }
  }

  var fixed = 0;
  for (var r = data.length - 1; r >= 1; r--) {
    if (String(data[r][7] || '').trim() !== '점수대기') continue;

    var studentName = String(data[r][1] || '').trim();
    var quizTitle   = String(data[r][2] || '').trim();
    var formUrl     = formMap[quizTitle];
    if (!formUrl) continue;

    try {
      var form = FormApp.openByUrl(formUrl);
      if (!form.isQuiz()) continue;

      var responses = form.getResponses();
      for (var ri = responses.length - 1; ri >= 0; ri--) {
        var resp = responses[ri];
        // 응답에서 이름 확인
        var items = resp.getItemResponses();
        var respName = '';
        for (var ii = 0; ii < items.length; ii++) {
          var itemTitle = items[ii].getItem().getTitle().toLowerCase().trim();
          if (itemTitle === '이름' || itemTitle === 'name' || itemTitle === '성명') {
            respName = String(items[ii].getResponse()).trim();
            break;
          }
        }
        if (respName !== studentName) continue;

        // 점수 계산
        var gradable = resp.getGradableItemResponses();
        var calcScore = 0, calcTotal = 0;
        for (var gi = 0; gi < gradable.length; gi++) {
          calcScore += gradable[gi].getScore();
          try {
            var origItem = gradable[gi].getItem();
            var iType = origItem.getType();
            if      (iType === FormApp.ItemType.MULTIPLE_CHOICE) calcTotal += origItem.asMultipleChoiceItem().getPoints();
            else if (iType === FormApp.ItemType.CHECKBOX)        calcTotal += origItem.asCheckboxItem().getPoints();
            else if (iType === FormApp.ItemType.LIST)            calcTotal += origItem.asListItem().getPoints();
            else if (iType === FormApp.ItemType.TEXT)            calcTotal += origItem.asTextItem().getPoints();
          } catch (_) {}
        }

        if (calcTotal > 0) {
          var newPct = Math.round(calcScore / calcTotal * 100);
          sheet.getRange(r + 1, 4).setValue(calcScore);
          sheet.getRange(r + 1, 5).setValue(calcTotal);
          sheet.getRange(r + 1, 6).setValue(newPct);
          sheet.getRange(r + 1, 8).setValue('');
          fixed++;
          Logger.log('점수 보정: ' + studentName + ' / ' + quizTitle + ' = ' + calcScore + '/' + calcTotal);
          break;
        }
      }
    } catch (e) {
      Logger.log('fixMissingScores 오류: ' + quizTitle + ' — ' + e.message);
    }
  }

  if (fixed > 0) invalidateCache();
  Logger.log('fixMissingScores 완료: ' + fixed + '개 보정');
}

// ============================================================
// 6. setupSheets — 초기 시트 구조 생성
// ============================================================
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 학생명단 (4열: 이름, 학교, 학년, 상태)
  if (!ss.getSheetByName(SHEET_STUDENTS)) {
    var s1 = ss.insertSheet(SHEET_STUDENTS);
    s1.appendRow(['이름', '학교', '학년', '상태']);
    s1.getRange('A1:D1').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
    s1.setColumnWidth(1, 100);
    s1.setColumnWidth(2, 120);
    s1.setColumnWidth(3, 60);
    s1.setColumnWidth(4, 80);
    // 상태 드롭다운
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['출석', '결석', '귀가', '추가문제'])
      .build();
    s1.getRange('D2:D200').setDataValidation(statusRule);
  }

  // 할일배정 (6열: 학교, 학년, 퀴즈제목, 폼링크, 폼ID, 대상학생)
  if (!ss.getSheetByName(SHEET_TASKS)) {
    var s2 = ss.insertSheet(SHEET_TASKS);
    s2.appendRow(['학교', '학년', '퀴즈제목', '폼링크', '폼ID', '대상학생']);
    s2.getRange('A1:F1').setFontWeight('bold').setBackground('#34a853').setFontColor('white');
    s2.setColumnWidth(1, 120);
    s2.setColumnWidth(2, 60);
    s2.setColumnWidth(3, 180);
    s2.setColumnWidth(4, 300);
    s2.setColumnWidth(5, 200);
    s2.setColumnWidth(6, 150);
  }

  // Summary (8열)
  getOrCreateSummary(ss);

  // 기간관리 (4열)
  if (!ss.getSheetByName(SHEET_PERIODS)) {
    var s4 = ss.insertSheet(SHEET_PERIODS);
    s4.appendRow(['기간명', '시작일', '종료일', '상태']);
    s4.getRange('A1:D1').setFontWeight('bold').setBackground('#9c27b0').setFontColor('white');
    s4.setColumnWidth(1, 150);
    s4.setColumnWidth(2, 120);
    s4.setColumnWidth(3, 120);
    s4.setColumnWidth(4, 80);
    // 상태 드롭다운
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['활성', '종료'])
      .build();
    s4.getRange('D2:D100').setDataValidation(rule);
  }

  // 주간과제 (6열)
  getOrCreateWeeklySheet(ss);
}

// ============================================================
// 7. 메뉴 추가 (편의 기능)
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 대시보드')
    .addItem('시트 초기 설정', 'setupSheets')
    .addItem('트리거 등록', 'registerAllTriggers')
    .addSeparator()
    .addItem('캐시 초기화', 'clearCache')
    .addItem('Summary 데이터 확인', 'checkSummary')
    .addToUi();
}

function clearCache() {
  invalidateCache();
  SpreadsheetApp.getUi().alert('캐시가 초기화되었습니다.');
}

function checkSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SUMMARY);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Summary 시트가 없습니다. 시트 초기 설정을 먼저 실행하세요.');
    return;
  }
  var count = Math.max(0, sheet.getLastRow() - 1);
  var periodName = getActivePeriodName(ss);
  var msg = 'Summary에 총 ' + count + '개의 제출 기록이 있습니다.';
  if (periodName) {
    msg += '\n현재 활성 기간: ' + periodName;
  }
  SpreadsheetApp.getUi().alert(msg);
}

// ============================================================
// 헬퍼 함수들
// ============================================================

/** Summary 시트가 없으면 생성 */
function getOrCreateSummary(ss) {
  var sheet = ss.getSheetByName(SHEET_SUMMARY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SUMMARY);
    sheet.appendRow(['타임스탬프', '이름', '퀴즈제목', '점수', '만점', '정답률', '기간명', '수정여부']);
    sheet.getRange('A1:H1').setFontWeight('bold').setBackground('#ea4335').setFontColor('white');
  }
  return sheet;
}

/** 기간관리 시트가 없으면 생성 */
function getOrCreatePeriods(ss) {
  var sheet = ss.getSheetByName(SHEET_PERIODS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PERIODS);
    sheet.appendRow(['기간명', '시작일', '종료일', '상태']);
    sheet.getRange('A1:D1').setFontWeight('bold').setBackground('#9c27b0').setFontColor('white');
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['활성', '종료'])
      .build();
    sheet.getRange('D2:D100').setDataValidation(rule);
  }
  return sheet;
}

/** 현재 활성 기간명 반환 (없으면 빈 문자열) */
function getActivePeriodName(ss) {
  var sheet = ss.getSheetByName(SHEET_PERIODS);
  if (!sheet) return '';
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim() === '활성') {
      return String(data[i][0]).trim();
    }
  }
  return '';
}

/** 기간 정보 목록 반환 */
function getPeriodsInfo(ss) {
  var sheet = ss.getSheetByName(SHEET_PERIODS);
  var periods = [];
  var activePeriod = '';
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][0]).trim();
      var active = String(data[i][3]).trim() === '활성';
      if (!name) continue;
      periods.push({
        name: name,
        startDate: String(data[i][1] || ''),
        endDate: String(data[i][2] || ''),
        active: active
      });
      if (active) activePeriod = name;
    }
  }
  return { periods: periods, activePeriod: activePeriod };
}

/** 특정 퀴즈가 배정된 학생 이름 목록 반환 */
function getStudentsForTask(ss, taskTitle) {
  var tasksSheet = ss.getSheetByName(SHEET_TASKS);
  var studentsSheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!tasksSheet || !studentsSheet) return [];

  var tasksData = tasksSheet.getDataRange().getValues();
  var studentsData = studentsSheet.getDataRange().getValues();

  var targetStudents = [];
  var seen = {};

  for (var i = 1; i < tasksData.length; i++) {
    var tTitle = String(tasksData[i][2] != null ? tasksData[i][2] : '').trim();
    if (tTitle !== taskTitle) continue;

    var tSchool = String(tasksData[i][0] != null ? tasksData[i][0] : '').trim();
    var tGrade = String(tasksData[i][1] != null ? tasksData[i][1] : '').trim();
    var tTargets = String(tasksData[i][5] != null ? tasksData[i][5] : '').trim();

    if (tTargets) {
      // 개별 학생 지정
      var names = tTargets.split(',');
      for (var n = 0; n < names.length; n++) {
        var nm = names[n].trim();
        if (nm && !seen[nm]) { targetStudents.push(nm); seen[nm] = true; }
      }
    } else if (tSchool) {
      // 학교+학년 기반
      for (var j = 1; j < studentsData.length; j++) {
        var sName = String(studentsData[j][0] != null ? studentsData[j][0] : '').trim();
        var sSchool = String(studentsData[j][1] != null ? studentsData[j][1] : '').trim();
        var sGrade = String(studentsData[j][2] != null ? studentsData[j][2] : '').trim();
        if (sSchool === tSchool && sGrade === tGrade && sName && !seen[sName]) {
          targetStudents.push(sName);
          seen[sName] = true;
        }
      }
    }
  }

  return targetStudents;
}

/** 특정 퀴즈+기간의 기존 제출 학생 맵 반환 */
function getExistingSubmissions(ss, taskTitle, periodName) {
  var sheet = ss.getSheetByName(SHEET_SUMMARY);
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][1] != null ? data[i][1] : '').trim();
    var title = String(data[i][2] != null ? data[i][2] : '').trim();
    var period = String(data[i][6] != null ? data[i][6] : '').trim();
    if (title === taskTitle && (!periodName || !period || period === periodName)) {
      map[name] = true;
    }
  }
  return map;
}

/** 캐시 무효화 */
function invalidateCache() {
  var cache = CacheService.getScriptCache();
  
  function removeChunks(baseKey) {
    var countStr = cache.get(baseKey + '_chunk_count');
    cache.remove(baseKey + '_chunk_count');
    if (countStr) {
      var count = parseInt(countStr, 10);
      for (var i = 0; i < count; i++) {
        cache.remove(baseKey + '_chunk_' + i);
      }
    }
  }
  
  removeChunks(CACHE_KEY);
  
  // 모든 기간 캐시 제거
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_PERIODS);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var name = String(data[i][0] || '').trim();
        if (name) removeChunks(CACHE_KEY + '_' + name);
      }
    }
  } catch (_) { /* 무시 */ }
}

// ============================================================
// 주간과제 관리 — CRUD 핸들러 + 헬퍼
// ============================================================

/** 주간과제 시트 가져오기 (없으면 생성) */
function getOrCreateWeeklySheet(ss) {
  var sheet = ss.getSheetByName(SHEET_WEEKLY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_WEEKLY);
    sheet.appendRow(['주차', '과제명', '대상학교', '문제형태', '링크URL', '문제유형']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** 주간과제 시트 → { weekNum: [{label, schools[], format, link, items[]}] } */
function getWeeklyTasksData(ss) {
  var sheet = ss.getSheetByName(SHEET_WEEKLY);
  if (!sheet) return {};

  var data = sheet.getDataRange().getValues();
  var result = {};

  for (var i = 1; i < data.length; i++) {
    var weekNum = parseInt(data[i][0], 10);
    if (isNaN(weekNum)) continue;

    var label = String(data[i][1] || '').trim();
    if (!label) continue;

    var schoolsStr = String(data[i][2] || '').trim();
    var schools = schoolsStr === '전체' ? ['전체'] : schoolsStr.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    var format = String(data[i][3] || '').trim() || '프린트';
    var link = String(data[i][4] || '').trim();
    var itemsStr = String(data[i][5] || '').trim();
    var items = itemsStr ? itemsStr.split(',').map(function(s) { return s.trim(); }).filter(Boolean) : [];

    if (!result[weekNum]) result[weekNum] = [];
    result[weekNum].push({
      label: label,
      schools: schools,
      format: format,
      link: link,
      items: items
    });
  }

  return result;
}

/** 주간과제 저장 핸들러 */
function handleWeeklyTaskSave(params) {
  var weekNum = parseInt(params.weekNum, 10);
  var label = (params.label || '').trim();
  if (isNaN(weekNum) || !label) {
    return { error: 'weekNum과 label이 필요합니다.' };
  }

  var schools = params.schools || '전체';  // 문자열 (쉼표 구분) or '전체'
  var format = params.format || '프린트';
  var link = params.link || '';
  var items = params.items || '';  // 문자열 (쉼표 구분)

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateWeeklySheet(ss);

  sheet.appendRow([weekNum, label, schools, format, link, items]);

  return { success: true, action: '주간과제저장', weekNum: weekNum, label: label };
}

/** 주간과제 삭제 핸들러 */
function handleWeeklyTaskDelete(params) {
  var weekNum = parseInt(params.weekNum, 10);
  var label = (params.label || '').trim();
  if (isNaN(weekNum) || !label) {
    return { error: 'weekNum과 label이 필요합니다.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_WEEKLY);
  if (!sheet) {
    return { error: '주간과제 시트가 없습니다.' };
  }

  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    var rowWeek = parseInt(data[i][0], 10);
    var rowLabel = String(data[i][1] || '').trim();
    if (rowWeek === weekNum && rowLabel === label) {
      sheet.deleteRow(i + 1); // 시트 행 번호는 1-based
      return { success: true, action: '주간과제삭제', weekNum: weekNum, label: label };
    }
  }

  return { error: '해당 과제를 찾을 수 없습니다: ' + weekNum + '주차 ' + label };
}

/** 주간과제 조회 핸들러 */
function handleWeeklyTaskList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tasks = getWeeklyTasksData(ss);
  return { success: true, action: '주간과제조회', weeklyTasks: tasks };
}

/** JSON 응답 생성 유틸리티 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
