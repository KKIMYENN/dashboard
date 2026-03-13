# 대시보드 개선 3종 구현 계획

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 기능 2(요약 복사), 기능 4(기간별 성적 히스토리), 기능 5(과제별 완료율 통계) 구현

**Architecture:** Google Apps Script(Code.gs) + 단일 HTML 프론트엔드. 기능 4만 백엔드 변경 필요(Summary 스캔 루프에 historyMap 수집 추가). 나머지 2개는 이미 로드된 데이터에서 프론트만 변경.

**Tech Stack:** Google Apps Script (ES5), vanilla JS, dashboard.html 인라인 CSS/JS

---

## Task 1: Code.gs — 기간별 히스토리 집계 (기능 4 백엔드)

**Files:**
- Modify: `/Users/mac4/Desktop/확인학습_대시보드/Code.gs`

### Step 1: historyMap 변수 선언 추가

`validStudentNamesSet` 블록(line 477~483) 바로 위, 빈 줄 사이에 추가:

```javascript
// --- 기간별 히스토리 맵 (활성 기간 외 종료된 기간들) ---
var historyMap = {}; // key: 학생이름 → { 기간명: { sum: 0, count: 0 } }
```

### Step 2: Summary 스캔 루프의 continue 교체 (line 505)

현재 코드:
```javascript
    if (activePeriod && sPeriod && sPeriod !== activePeriod) continue;
```

교체 후:
```javascript
    if (activePeriod && sPeriod && sPeriod !== activePeriod) {
      if (validStudentNamesSet[sName] && sPct !== null) {
        if (!historyMap[sName]) historyMap[sName] = {};
        if (!historyMap[sName][sPeriod]) historyMap[sName][sPeriod] = { sum: 0, count: 0 };
        historyMap[sName][sPeriod].sum += sPct;
        historyMap[sName][sPeriod].count++;
      }
      continue;
    }
```

### Step 3: student 객체에 history 필드 추가 (line 664~674)

현재 코드:
```javascript
    students.push({
      name: pName,
      school: pSchool,
      grade: pGrade,
      status: pStatus,
      tasks: tasks,
      doneCount: doneCount,
      totalCount: tasks.length,
      avgPct: avgPct,
      textbook: textbook
    });
```

교체 후:
```javascript
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
```

### Step 4: 변경사항 저장 후 확인

파일 저장. 다음 task에서 배포.

---

## Task 2: dashboard.html — 기능 4 프론트엔드 (히스토리 섹션)

**Files:**
- Modify: `/Users/mac4/Desktop/확인학습_대시보드/dashboard.html`

### Step 1: CSS 추가

`</style>` 직전에 추가:

```css
/* ══════ 기간별 히스토리 (기능 4) ══════ */
.period-history { margin-top: 10px; padding-top: 8px; border-top: 1px solid var(--border); }
.period-history-title { font-size: 11px; color: var(--text-3); margin-bottom: 6px; font-weight: 600; letter-spacing: 0.04em; }
.period-chips { display: flex; flex-wrap: wrap; gap: 6px; }
.period-chip { display: inline-flex; flex-direction: column; align-items: center; font-size: 11px; padding: 4px 8px; border-radius: 6px; background: var(--bg); border: 1px solid var(--border); min-width: 52px; }
.period-chip .pc-name { color: var(--text-3); font-size: 10px; line-height: 1.2; }
.period-chip .pc-pct { font-weight: 700; font-size: 13px; line-height: 1.4; }
```

### Step 2: renderDetail(s) 함수에 히스토리 섹션 추가

`renderDetail(s)` 함수 끝 (line 1800) `return tasks + actions;` 를 아래로 교체:

```javascript
  // 기간별 히스토리 섹션 (2개 이상 기간 데이터 있을 때만 표시)
  let historyHtml = '';
  if (s.history && s.history.length >= 1) {
    const chips = s.history.map(h => {
      const cls = h.avgPct >= 90 ? 'sc-green' : h.avgPct >= 70 ? 'sc-blue' : h.avgPct >= 50 ? 'sc-amber' : 'sc-red';
      return `<span class="period-chip"><span class="pc-name">${esc(h.period)}</span><span class="pc-pct ${cls}">${h.avgPct}%</span></span>`;
    }).join('');
    historyHtml = `<div class="period-history"><div class="period-history-title">과거 기록</div><div class="period-chips">${chips}</div></div>`;
  }

  return tasks + historyHtml + actions;
```

---

## Task 3: dashboard.html — 기능 5 (과제별 완료율 통계 패널)

**Files:**
- Modify: `/Users/mac4/Desktop/확인학습_대시보드/dashboard.html`

### Step 1: CSS 추가

기능 4 CSS 아래에 추가:

```css
/* ══════ 과제별 통계 패널 (기능 5) ══════ */
.task-stats-panel { background: var(--card); border: 1px solid var(--border); border-radius: 10px; margin: 8px 16px; overflow: hidden; }
.task-stats-toggle { padding: 10px 16px; cursor: pointer; font-size: 13px; font-weight: 600; user-select: none; display: flex; align-items: center; gap: 6px; }
.task-stats-toggle:hover { background: var(--bg); }
.task-stats-body { display: none; overflow-x: auto; }
.task-stats-body.open { display: block; }
.task-stats-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.task-stats-table th { background: var(--bg); padding: 8px 12px; text-align: left; font-weight: 600; color: var(--text-2); border-bottom: 1px solid var(--border); white-space: nowrap; }
.task-stats-table td { padding: 8px 12px; border-bottom: 1px solid var(--border); }
.task-stats-table tr:last-child td { border-bottom: none; }
.ts-rate-bar { display: inline-block; height: 6px; background: var(--border); border-radius: 3px; width: 60px; vertical-align: middle; margin-right: 6px; overflow: hidden; }
.ts-rate-fill { height: 100%; border-radius: 3px; }
```

### Step 2: HTML에 패널 컨테이너 추가

`<div class="dash-stats" id="statsBar"></div>` (line 831) 바로 아래에 추가:

```html
<!-- Task Stats Panel (기능 5) -->
<div class="task-stats-panel" id="taskStatsPanel" style="display:none;">
  <div class="task-stats-toggle" onclick="toggleTaskStats()">📊 과제별 통계 <span id="taskStatsArrow">▾</span></div>
  <div class="task-stats-body" id="taskStatsBody"></div>
</div>
```

### Step 3: JavaScript 함수 추가

`renderStats()` 함수 바로 아래에 추가:

```javascript
function renderTaskStats() {
  const panel = document.getElementById('taskStatsPanel');
  const body = document.getElementById('taskStatsBody');
  if (!dashboardData || !dashboardData.students) { panel.style.display = 'none'; return; }

  // 결석 학생 제외한 활성 학생 목록
  const activeStudents = dashboardData.students.filter(s => s.status !== '결석');
  if (activeStudents.length === 0) { panel.style.display = 'none'; return; }

  // 과제별 집계
  const taskMap = {}; // title → { done, scoreSum, scoreCount, total }
  activeStudents.forEach(s => {
    (s.tasks || []).forEach(t => {
      if (!taskMap[t.title]) taskMap[t.title] = { done: 0, scoreSum: 0, scoreCount: 0, total: 0 };
      taskMap[t.title].total++;
      if (t.done) {
        taskMap[t.title].done++;
        if (t.pct !== null && t.pct !== undefined) {
          taskMap[t.title].scoreSum += t.pct;
          taskMap[t.title].scoreCount++;
        }
      }
    });
  });

  const rows = Object.entries(taskMap).map(([title, d]) => {
    const rate = d.total > 0 ? Math.round(d.done / d.total * 100) : 0;
    const avgScore = d.scoreCount > 0 ? Math.round(d.scoreSum / d.scoreCount) : null;
    const fillColor = rate >= 90 ? 'var(--green)' : rate >= 70 ? 'var(--blue)' : rate >= 50 ? 'var(--amber)' : 'var(--red)';
    return { title, done: d.done, notDone: d.total - d.done, total: d.total, rate, avgScore, fillColor };
  }).sort((a, b) => a.rate - b.rate); // 완료율 낮은 순

  if (rows.length === 0) { panel.style.display = 'none'; return; }

  panel.style.display = 'block';
  body.innerHTML = `<table class="task-stats-table">
    <thead><tr>
      <th>과제명</th><th>제출</th><th>미제출</th>
      <th>완료율</th><th>평균점수</th>
    </tr></thead>
    <tbody>${rows.map(r => `
      <tr>
        <td>${esc(r.title)}</td>
        <td>${r.done}</td>
        <td>${r.notDone}</td>
        <td>
          <span class="ts-rate-bar"><span class="ts-rate-fill" style="width:${r.rate}%;background:${r.fillColor};"></span></span>
          ${r.rate}%
        </td>
        <td>${r.avgScore !== null ? r.avgScore + '%' : '-'}</td>
      </tr>`).join('')}
    </tbody>
  </table>`;
}

function toggleTaskStats() {
  const body = document.getElementById('taskStatsBody');
  const arrow = document.getElementById('taskStatsArrow');
  body.classList.toggle('open');
  arrow.textContent = body.classList.contains('open') ? '▴' : '▾';
}
```

### Step 4: renderStats() 호출 뒤에 renderTaskStats() 호출 추가

`renderStats()` 함수가 호출되는 모든 곳에 `renderTaskStats()` 함께 호출.

`renderStats()` 함수 본문 끝, `document.getElementById('statsBar').innerHTML = ...` 블록 마지막에 추가:

```javascript
  renderTaskStats();
```

---

## Task 4: dashboard.html — 기능 2 (요약 복사 버튼)

**Files:**
- Modify: `/Users/mac4/Desktop/확인학습_대시보드/dashboard.html`

### Step 1: CSS 추가

기능 5 CSS 아래에 추가:

```css
/* ══════ 요약 복사 버튼 (기능 2) ══════ */
.summary-copy-btn { font-size: 12px; padding: 5px 10px; border-radius: 6px; border: 1px solid var(--border); background: var(--card); color: var(--text-1); cursor: pointer; white-space: nowrap; }
.summary-copy-btn:hover { background: var(--bg); }
.toast { position: fixed; bottom: 24px; left: 50%; transform: translateX(-50%) translateY(20px); background: #1E293B; color: #fff; padding: 10px 20px; border-radius: 8px; font-size: 13px; opacity: 0; pointer-events: none; transition: opacity 0.2s, transform 0.2s; z-index: 9999; }
.toast.show { opacity: 1; transform: translateX(-50%) translateY(0); }
```

### Step 2: 버튼 HTML 추가

header (line 808) `<div class="dash-meta">` 안에 설정 링크 직전에 추가:

```html
<button class="summary-copy-btn" id="summaryCopyBtn" onclick="copySummaryText()" style="display:none;">📋 요약 복사</button>
```

즉, 현재:
```html
      <div class="dash-meta">
        <span id="dateDisplay"></span>
        <span><span class="live-dot" id="pollDot"></span> <span id="pollStatus">연결 중...</span></span>
        <span class="settings-link" onclick="showSettings()">설정</span>
      </div>
```

변경 후:
```html
      <div class="dash-meta">
        <span id="dateDisplay"></span>
        <span><span class="live-dot" id="pollDot"></span> <span id="pollStatus">연결 중...</span></span>
        <button class="summary-copy-btn" id="summaryCopyBtn" onclick="copySummaryText()" style="display:none;">📋 요약 복사</button>
        <span class="settings-link" onclick="showSettings()">설정</span>
      </div>
```

### Step 3: toast HTML 추가

`</body>` 직전에 추가:

```html
<div class="toast" id="toastMsg"></div>
```

### Step 4: JavaScript 함수 추가

```javascript
function showToast(msg) {
  const el = document.getElementById('toastMsg');
  el.textContent = msg;
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), 2000);
}

function copySummaryText() {
  if (!dashboardData) return;
  const s = dashboardData.stats || {};
  const period = (dashboardData.activePeriod) || '이번 기간';
  const students = dashboardData.students || [];

  // 70% 미만 (출석 + 완료)
  const low = students.filter(st => st.status !== '결석' && st.doneCount > 0 && st.avgPct < 70)
    .map(st => st.name).join(', ');
  // 미완료 (출석 + doneCount < totalCount)
  const notDone = students.filter(st => st.status !== '결석' && st.doneCount < st.totalCount)
    .map(st => st.name).join(', ');

  const absent = s.absent || 0;
  const late = s.late || 0;
  const present = s.total - absent;

  let text = `[${period}] 확인학습 요약\n`;
  text += `출석 ${present}명 / 결석 ${absent}명` + (late ? ` / 지각 ${late}명` : '') + '\n';
  text += `완료 ${s.allDone || 0}명 | 평균 ${s.avgScore || 0}% | 완료율 ${present > 0 ? Math.round((s.allDone || 0) / present * 100) : 0}%\n`;
  text += '───\n';
  if (low) text += `70% 미만: ${low}\n`;
  if (notDone) text += `미완료: ${notDone}\n`;

  navigator.clipboard.writeText(text).then(() => showToast('복사됨 ✓'))
    .catch(() => { prompt('아래 텍스트를 복사하세요:', text); });
}
```

### Step 5: 데이터 로드 완료 시 버튼 표시

`renderStats()` 또는 `fetchData()` 성공 콜백 부근에서 버튼 표시:

```javascript
  document.getElementById('summaryCopyBtn').style.display = '';
```

`renderStats()` 함수 맨 첫 줄에 추가:
```javascript
  document.getElementById('summaryCopyBtn').style.display = '';
```

---

## Task 5: Code.gs 재배포

**Files:**
- Deploy: Code.gs

### Step 1: Google Apps Script 편집기에서 배포

1. script.google.com 에서 프로젝트 열기
2. 배포 → 새 배포 → 유형: 웹 앱
3. 설명: "기능4 히스토리 집계 추가"
4. 실행: 나(본인), 액세스: 모든 사용자
5. 배포 → URL 복사
6. dashboard.html의 `SCRIPT_URL` 상수 업데이트 (이미 설정된 경우 스킵)

---

## 검증 체크리스트

1. **기능 2**: 데이터 로드 후 헤더에 "📋 요약 복사" 버튼 표시 → 클릭 시 "복사됨 ✓" 토스트 → 클립보드 내용 확인
2. **기능 4**: 학생 카드 펼침 → 이전 기간 데이터가 있으면 "과거 기록" 칩 표시 → 이번 주 첫 수업이면 히스토리 없어서 섹션 숨김 (정상)
3. **기능 5**: 데이터 로드 후 stats 바 아래 "📊 과제별 통계" 패널 표시 → 클릭하면 테이블 펼침 → 완료율 낮은 순 정렬 확인 → 결석 학생 제외 확인
