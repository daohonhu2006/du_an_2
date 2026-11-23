(() => {
  'use strict';

  const STORAGE_KEYS = {
    CATALOG: 'huit_dkhp_catalog_v2',
    REGISTERED: 'huit_dkhp_registered_v2',
    THEME: 'huit_dkhp_theme_v1',
    TUITION_RATE_LT: 'huit_dkhp_tuition_rate_lt_v1',
    TUITION_RATE_TH: 'huit_dkhp_tuition_rate_v1',
    TUITION_SELECTIONS: 'huit_dkhp_tuition_selections_v1',
    CUSTOM_TUITION: 'huit_dkhp_custom_tuition_v1',
    CALENDAR_EVENTS: 'huit_dkhp_calendar_events_v1',
    USERS: 'huit_dkhp_users_v1',
    CURRENT_USER: 'huit_dkhp_current_user_v1',
  };

  // State (khởi tạo trống, người dùng tự tải lên file Excel TKB của mình)
  let currentUser = (() => {
    try {
      return localStorage.getItem(STORAGE_KEYS.CURRENT_USER) || null;
    } catch {
      return null;
    }
  })();
  let catalog = loadState(STORAGE_KEYS.CATALOG, []);
  let registered = loadState(STORAGE_KEYS.REGISTERED, []);
  let customTuitionItems = loadState(STORAGE_KEYS.CUSTOM_TUITION, []);
  let calendarEvents = loadObjectState(STORAGE_KEYS.CALENDAR_EVENTS, {});
  let tuitionSelections = loadObjectState(STORAGE_KEYS.TUITION_SELECTIONS, {});

  // DOM Elements
  const themeSelect = document.getElementById('theme-select');
  const catalogTbody = document.getElementById('catalog-tbody');
  const catalogTableWrap = document.getElementById('catalog-table-wrap');
  const excelDropzone = document.getElementById('excel-dropzone');
  const catalogCount = document.getElementById('catalog-count');
  const registeredListEl = document.getElementById('registered-list');
  const timetableTbody = document.getElementById('timetable-tbody');
  const statCourses = document.getElementById('stat-courses');
  const statCredits = document.getElementById('stat-credits');
  const searchInput = document.getElementById('search-input');
  const searchSuggestionsEl = document.getElementById('search-suggestions');
  const filterSubject = document.getElementById('filter-subject');
  const filterClass = document.getElementById('filter-class');
  const filterLecturer = document.getElementById('filter-lecturer');
  const filterDay = document.getElementById('filter-day');
  const filterSession = document.getElementById('filter-session');
  const filterStatus = document.getElementById('filter-status');
  const btnResetFilters = document.getElementById('btn-reset-filters');
  const importStatusEl = document.getElementById('import-status');
  const tabRegCount = document.getElementById('tab-reg-count');
  const viewTabs = Array.from(document.querySelectorAll('.view-tab'));
  const panelRegister = document.getElementById('panel-register');
  const panelTimetable = document.getElementById('panel-timetable');
  const panelCalendar = document.getElementById('panel-calendar');
  const panelTuition = document.getElementById('panel-tuition');
  const panelExcel = document.getElementById('panel-excel');

  const tuitionRateLtInput = document.getElementById('tuition-rate-lt');
  const tuitionRateInput = document.getElementById('tuition-rate-input');
  const tuitionCheckAll = document.getElementById('tuition-check-all');
  const tuitionTotalLtCreditsEl = document.getElementById('tuition-total-lt-credits');
  const tuitionTotalThCreditsEl = document.getElementById('tuition-total-th-credits');
  const tuitionTotalAmountEl = document.getElementById('tuition-total-amount');
  const tuitionTbody = document.getElementById('tuition-tbody');

  function loadObjectState(key, fallback) {
    try {
      const raw = localStorage.getItem(key);
      if (!raw) return structuredClone(fallback);
      const parsed = JSON.parse(raw);
      return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : structuredClone(fallback);
    } catch {
      return structuredClone(fallback);
    }
  }

  const conflictDialog = document.getElementById('conflict-dialog');
  const conflictSummaryEl = document.getElementById('conflict-summary');
  const conflictAlternativesEl = document.getElementById('conflict-alternatives');
  const excelGuideDialog = document.getElementById('excel-guide-dialog');

  // --- Utilities ---
  function loadState(key, fallback) {
    try {
      const raw = localStorage.getItem(key);
      if (!raw) return structuredClone(fallback);
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : structuredClone(fallback);
    } catch {
      return structuredClone(fallback);
    }
  }

  function getUserDataKey(username) {
    return `huit_dkhp_user_data_${username}`;
  }

  function saveState() {
    try {
      localStorage.setItem(STORAGE_KEYS.CATALOG, JSON.stringify(catalog));
      localStorage.setItem(STORAGE_KEYS.REGISTERED, JSON.stringify(registered));
      localStorage.setItem(STORAGE_KEYS.CUSTOM_TUITION, JSON.stringify(customTuitionItems));
      localStorage.setItem(STORAGE_KEYS.CALENDAR_EVENTS, JSON.stringify(calendarEvents));
      if (currentUser) {
        const userPayload = {
          catalog,
          registered,
          customTuitionItems,
          calendarEvents,
          tuitionSelections,
          rateLt: tuitionRateLtInput?.value || '750000',
          rateTh: tuitionRateInput?.value || '850000',
          theme: themeSelect?.value || 'system',
        };
        localStorage.setItem(getUserDataKey(currentUser), JSON.stringify(userPayload));
      }
    } catch {
      // Ignore storage quota errors
    }
  }

  function stripDiacritics(str) {
    return String(str || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/đ/g, 'd')
      .replace(/Đ/g, 'D')
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .trim();
  }

  function levenshtein(a, b) {
    if (a === b) return 0;
    const m = a.length;
    const n = b.length;
    if (m === 0) return n;
    if (n === 0) return m;
    const prev = Array.from({ length: n + 1 }, (_, i) => i);
    const curr = new Array(n + 1);
    for (let i = 1; i <= m; i++) {
      curr[0] = i;
      for (let j = 1; j <= n; j++) {
        const cost = a[i - 1] === b[j - 1] ? 0 : 1;
        curr[j] = Math.min(curr[j - 1] + 1, prev[j] + 1, prev[j - 1] + cost);
      }
      for (let j = 0; j <= n; j++) prev[j] = curr[j];
    }
    return prev[n];
  }

  /**
   * Chấm điểm khớp chuỗi không dấu + gần đúng ("na ná") cho tìm kiếm & gợi ý.
   */
  function fuzzyScore(rawQuery, rawTarget) {
    const q = stripDiacritics(rawQuery);
    const t = stripDiacritics(rawTarget);
    if (!q || !t) return 0;
    if (t === q) return 100;
    if (t.startsWith(q)) return 90;
    if (t.includes(q)) return 80;

    const qTokens = q.split(' ').filter(Boolean);
    const tTokens = t.split(/[\s.\-_/]+/).filter(Boolean);
    if (!qTokens.length || !tTokens.length) return 0;

    let total = 0;
    let matchedCount = 0;

    for (const qt of qTokens) {
      let bestTokenScore = 0;
      for (const tt of tTokens) {
        if (tt === qt) {
          bestTokenScore = Math.max(bestTokenScore, 30);
        } else if (tt.startsWith(qt)) {
          bestTokenScore = Math.max(bestTokenScore, 24);
        } else if (tt.includes(qt)) {
          bestTokenScore = Math.max(bestTokenScore, 18);
        } else if (qt.length >= 2) {
          const maxDist = qt.length <= 4 ? 1 : 2;
          const dist = levenshtein(qt, tt.slice(0, Math.min(tt.length, qt.length + maxDist)));
          const fullDist = levenshtein(qt, tt);
          const bestD = Math.min(dist, fullDist);
          if (bestD <= maxDist) {
            bestTokenScore = Math.max(bestTokenScore, 14 - bestD * 4);
          }
        }
      }
      if (bestTokenScore > 0) {
        matchedCount++;
        total += bestTokenScore;
      }
    }

    if (matchedCount === 0) return 0;
    // Yêu cầu khớp ít nhất 1 token và ưu tiên khớp đủ số từ đã nhập
    const coverage = matchedCount / qTokens.length;
    if (qTokens.length > 1 && coverage < 0.5) return 0;
    return Math.round(total * coverage);
  }

  let activeSuggestionIndex = -1;
  let currentSuggestions = [];

  function buildSuggestionPool() {
    const map = new Map();
    const add = (type, label) => {
      const clean = String(label || '').trim();
      if (!clean) return;
      const key = `${type}::${clean}`;
      if (!map.has(key)) map.set(key, { type, label: clean });
    };

    for (const s of catalog) {
      add('Môn học', s.tenHP);
      if (s.lopHoc) add('Lớp', s.lopHoc);
      if (s.ghiChu) add('Hình thức', s.ghiChu);
      add('Mã LHP', s.maLHP);
      for (const item of getSectionSessions(s).all) {
        if (item.sess.gv) add('Giảng viên', item.sess.gv);
        if (item.sess.phong) add('Phòng', item.sess.phong);
      }
    }
    return Array.from(map.values());
  }

  function hideSuggestions() {
    searchSuggestionsEl.hidden = true;
    searchInput.setAttribute('aria-expanded', 'false');
    activeSuggestionIndex = -1;
  }

  function renderSuggestions() {
    const q = searchInput.value.trim();
    if (!q) {
      hideSuggestions();
      return;
    }

    const pool = buildSuggestionPool();
    currentSuggestions = pool
      .map((item) => ({ ...item, score: fuzzyScore(q, item.label) }))
      .filter((item) => item.score > 0)
      .sort((a, b) => b.score - a.score || a.label.localeCompare(b.label, 'vi'))
      .slice(0, 8);

    if (currentSuggestions.length === 0) {
      hideSuggestions();
      return;
    }

    activeSuggestionIndex = -1;
    searchSuggestionsEl.innerHTML = currentSuggestions
      .map(
        (item, idx) => `
        <li class="suggestion-item" role="option" aria-selected="false" data-suggestion-idx="${idx}">
          <span>${escapeHtml(item.label)}</span>
          <span class="suggestion-type">${escapeHtml(item.type)}</span>
        </li>
      `
      )
      .join('');
    searchSuggestionsEl.hidden = false;
    searchInput.setAttribute('aria-expanded', 'true');
  }

  function selectSuggestion(idx) {
    const item = currentSuggestions[idx];
    if (!item) return;

    // Đặt lại các bộ lọc đang kẹt để hiển thị đúng duy nhất lựa chọn vừa chọn
    filterSubject.value = '';
    filterClass.value = '';
    filterLecturer.value = '';
    filterDay.value = '';
    filterSession.value = '';
    filterStatus.value = 'all';

    if (item.type === 'Môn học') {
      filterSubject.value = item.label;
      searchInput.value = item.label;
    } else if (item.type === 'Lớp') {
      filterClass.value = item.label;
      searchInput.value = item.label;
    } else if (item.type === 'Giảng viên') {
      filterLecturer.value = item.label;
      searchInput.value = item.label;
    } else {
      searchInput.value = item.label;
    }

    hideSuggestions();
    renderCatalog();
  }

  function escapeHtml(str) {
    return String(str ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function formatThu(thu) {
    if (Number(thu) === 8) return 'Chủ nhật';
    return `Thứ ${thu}`;
  }

  // Khung giờ 16 tiết chính thức của HUIT (Lý thuyết có nghỉ giải lao 30p sau tiết 3 & 9; Thực hành không nghỉ giải lao)
  const HUIT_LT_SLOTS = {
    1: { start: '06h30', end: '07h20', sMin: 390, eMin: 440 },
    2: { start: '07h20', end: '08h10', sMin: 440, eMin: 490 },
    3: { start: '08h10', end: '09h00', sMin: 490, eMin: 540 },
    4: { start: '09h30', end: '10h20', sMin: 570, eMin: 620 },
    5: { start: '10h20', end: '11h10', sMin: 620, eMin: 670 },
    6: { start: '11h10', end: '12h00', sMin: 670, eMin: 720 },
    7: { start: '12h30', end: '13h20', sMin: 750, eMin: 800 },
    8: { start: '13h20', end: '14h10', sMin: 800, eMin: 850 },
    9: { start: '14h10', end: '15h00', sMin: 850, eMin: 900 },
    10: { start: '15h30', end: '16h20', sMin: 930, eMin: 980 },
    11: { start: '16h20', end: '17h10', sMin: 980, eMin: 1030 },
    12: { start: '17h10', end: '18h00', sMin: 1030, eMin: 1080 },
    13: { start: '18h10', end: '19h00', sMin: 1090, eMin: 1140 },
    14: { start: '19h00', end: '19h50', sMin: 1140, eMin: 1190 },
    15: { start: '19h50', end: '20h40', sMin: 1190, eMin: 1240 },
    16: { start: '20h40', end: '21h30', sMin: 1240, eMin: 1290 },
  };

  const HUIT_TH_SLOTS = {
    1: { start: '07h00', end: '07h50', sMin: 420, eMin: 470 },
    2: { start: '07h50', end: '08h40', sMin: 470, eMin: 520 },
    3: { start: '08h40', end: '09h30', sMin: 520, eMin: 570 },
    4: { start: '09h30', end: '10h20', sMin: 570, eMin: 620 },
    5: { start: '10h20', end: '11h10', sMin: 620, eMin: 670 },
    6: { start: '11h10', end: '12h00', sMin: 670, eMin: 720 },
    7: { start: '12h30', end: '13h20', sMin: 750, eMin: 800 },
    8: { start: '13h20', end: '14h10', sMin: 800, eMin: 850 },
    9: { start: '14h10', end: '15h00', sMin: 850, eMin: 900 },
    10: { start: '15h00', end: '15h50', sMin: 900, eMin: 950 },
    11: { start: '15h50', end: '16h40', sMin: 950, eMin: 1000 },
    12: { start: '16h40', end: '17h30', sMin: 1000, eMin: 1050 },
    13: { start: '18h00', end: '18h50', sMin: 1080, eMin: 1130 },
    14: { start: '18h50', end: '19h40', sMin: 1130, eMin: 1180 },
    15: { start: '19h40', end: '20h30', sMin: 1180, eMin: 1230 },
    16: { start: '20h30', end: '21h20', sMin: 1230, eMin: 1280 },
  };

  function getSessionClockRange(sess, typeCode = 'LT') {
    if (!sess) return '';
    const table = typeCode === 'TH' ? HUIT_TH_SLOTS : HUIT_LT_SLOTS;
    const s = table[sess.start];
    const e = table[sess.end];
    if (!s || !e) return '';
    return `${s.start}–${e.end}`;
  }

  function getSessionMinutes(sess, typeCode = 'LT') {
    const table = typeCode === 'TH' ? HUIT_TH_SLOTS : HUIT_LT_SLOTS;
    const s = table[sess.start];
    const e = table[sess.end];
    if (!s || !e) return null;
    return { sMin: s.sMin, eMin: e.eMin };
  }

  function formatScheduleText(sess, typeCode = 'LT') {
    if (!sess) return '—';
    const clock = getSessionClockRange(sess, typeCode);
    return `${formatThu(sess.thu)}, Tiết ${sess.start}–${sess.end}${clock ? ` (${clock})` : ''}`;
  }

  function parseWeekSet(tuanStr) {
    const weeks = new Set();
    const raw = String(tuanStr || '1-15').trim();
    if (!raw) {
      for (let i = 1; i <= 15; i++) weeks.add(i);
      return weeks;
    }
    for (const part of raw.split(',')) {
      const trimmed = part.trim();
      if (!trimmed) continue;
      const range = trimmed.match(/^(\d+)\s*[-–]\s*(\d+)$/);
      if (range) {
        const a = Number(range[1]);
        const b = Number(range[2]);
        for (let w = Math.min(a, b); w <= Math.max(a, b); w++) weeks.add(w);
      } else if (/^\d+$/.test(trimmed)) {
        weeks.add(Number(trimmed));
      }
    }
    if (weeks.size === 0) {
      for (let i = 1; i <= 15; i++) weeks.add(i);
    }
    return weeks;
  }

  function weeksOverlap(tuanA, tuanB) {
    const setA = parseWeekSet(tuanA);
    const setB = parseWeekSet(tuanB);
    for (const w of setA) {
      if (setB.has(w)) return true;
    }
    return false;
  }

  function parseExcelDate(val) {
    if (val === undefined || val === null) return null;
    const raw = String(val).trim();
    if (!raw) return null;

    // dd/mm/yyyy hoặc d/m/yyyy hoặc dd-mm-yyyy
    const dmy = raw.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (dmy) {
      const day = Number(dmy[1]);
      const month = Number(dmy[2]);
      const year = dmy[3].length === 2 ? 2000 + Number(dmy[3]) : Number(dmy[3]);
      if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        const iso = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const display = `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${year}`;
        return { iso, display };
      }
    }

    // yyyy-mm-dd
    const ymd = raw.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (ymd) {
      const year = Number(ymd[1]);
      const month = Number(ymd[2]);
      const day = Number(ymd[3]);
      const iso = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      const display = `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${year}`;
      return { iso, display };
    }

    // Excel serial date
    if (/^\d{4,5}$/.test(raw)) {
      const serial = Number(raw);
      if (serial > 20000 && serial < 70000) {
        const utc = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
        const year = utc.getUTCFullYear();
        const month = utc.getUTCMonth() + 1;
        const day = utc.getUTCDate();
        return {
          iso: `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`,
          display: `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${year}`,
        };
      }
    }
    return null;
  }

  function dateRangesOverlap(a, b) {
    if (a.batDau && b.ketThuc && a.batDau > b.ketThuc) return false;
    if (b.batDau && a.ketThuc && b.batDau > a.ketThuc) return false;
    return true;
  }

  function sessionsOverlap(a, b, typeA = 'LT', typeB = 'LT') {
    if (!a || !b) return false;
    if (Number(a.thu) !== Number(b.thu)) return false;
    if (!dateRangesOverlap(a, b)) return false;
    if (!weeksOverlap(a.tuan, b.tuan)) return false;

    const periodOverlap = a.start <= b.end && b.start <= a.end;
    if (periodOverlap) return true;

    // Kiểm tra trùng theo giờ thực tế HUIT (vì giờ LT và giờ TH lệch nhau ở một số tiết)
    const minA = getSessionMinutes(a, typeA);
    const minB = getSessionMinutes(b, typeB);
    if (minA && minB) {
      return minA.sMin < minB.eMin && minB.sMin < minA.eMin;
    }
    return false;
  }

  function getSectionSessions(sec) {
    if (!sec) return { ltList: [], thList: [], all: [] };
    const ltList = Array.isArray(sec.ltList) ? sec.ltList : sec.lt ? [sec.lt] : [];
    const thList = Array.isArray(sec.thList) ? sec.thList : sec.th ? [sec.th] : [];
    const all = [
      ...ltList.map((sess) => ({ type: 'Lý thuyết', code: 'LT', sess })),
      ...thList.map((sess) => ({ type: 'Thực hành', code: 'TH', sess })),
    ];
    return { ltList, thList, all };
  }

  /**
   * Kiểm tra xung đột lịch (cả Lý thuyết & Thực hành) và trùng môn học.
   * @param {object} candidate - Lớp học phần muốn đăng ký
   * @param {string|null} ignoreMaHP - Bỏ qua môn này khi kiểm tra chuyển lớp cùng môn
   */
  function evaluateCandidate(candidate, ignoreMaHP = null) {
    const sameSubjectSection =
      !ignoreMaHP
        ? registered.find((r) => r.maHP === candidate.maHP) || null
        : null;

    const timeConflicts = [];
    const candidateSessions = getSectionSessions(candidate).all;

    for (const existing of registered) {
      if (ignoreMaHP && existing.maHP === ignoreMaHP) continue;
      if (existing.maLHP === candidate.maLHP) continue;

      const existingSessions = getSectionSessions(existing).all;

      for (const cItem of candidateSessions) {
        for (const eItem of existingSessions) {
          if (sessionsOverlap(cItem.sess, eItem.sess, cItem.code, eItem.code)) {
            timeConflicts.push({
              candidateType: cItem.type,
              candidateCode: cItem.code,
              candidateSess: cItem.sess,
              existingSection: existing,
              existingType: eItem.type,
              existingCode: eItem.code,
              existingSess: eItem.sess,
            });
          }
        }
      }
    }

    return {
      sameSubjectSection,
      timeConflicts,
      hasConflict: timeConflicts.length > 0 || Boolean(sameSubjectSection),
    };
  }

  // --- Registration Actions ---
  function attemptRegister(maLHP) {
    const candidate = catalog.find((s) => s.maLHP === maLHP);
    if (!candidate) return;

    if (registered.some((r) => r.maLHP === maLHP)) return;

    const evalResult = evaluateCandidate(candidate);
    if (evalResult.hasConflict) {
      showConflictModal(candidate, evalResult);
      return;
    }

    registered.push(structuredClone(candidate));
    saveState();
    renderAll();
  }

  function switchToSection(maLHP) {
    const candidate = catalog.find((s) => s.maLHP === maLHP);
    if (!candidate) return;

    const evalResult = evaluateCandidate(candidate, candidate.maHP);
    if (evalResult.timeConflicts.length > 0) {
      showConflictModal(candidate, evalResult);
      return;
    }

    // Thay thế lớp cũ của cùng môn học (nếu có) bằng lớp mới không trùng
    registered = registered.filter((r) => r.maHP !== candidate.maHP);
    registered.push(structuredClone(candidate));
    saveState();
    if (conflictDialog.open) conflictDialog.close();
    renderAll();
  }

  function unregisterSection(maLHP) {
    registered = registered.filter((r) => r.maLHP !== maLHP);
    saveState();
    renderAll();
  }

  // --- Conflict Warning Modal ---
  function showConflictModal(candidate, evalResult) {
    const parts = [];

    if (evalResult.sameSubjectSection && evalResult.timeConflicts.length === 0) {
      parts.push(
        `<p>Môn <strong>${escapeHtml(candidate.tenHP)}</strong> (${escapeHtml(candidate.maHP)}) đã được đăng ký ở lớp <code>${escapeHtml(evalResult.sameSubjectSection.maLHP)}</code>.</p>
         <p>Bạn có thể chọn lại lớp học phần khác bên dưới để thay thế.</p>`
      );
    } else {
      parts.push(
        `<p>Không thể đăng ký lớp <strong>${escapeHtml(candidate.tenHP)}</strong> (<code>${escapeHtml(candidate.maLHP)}</code>) vì bị trùng lịch với học phần đã đăng ký:</p>`
      );
      const listItems = evalResult.timeConflicts.map((c) => {
        const cRange = formatScheduleText(c.candidateSess, c.candidateCode);
        const eRange = formatScheduleText(c.existingSess, c.existingCode);
        return `<li>
          Lịch <strong>${escapeHtml(c.candidateType)}</strong> (${escapeHtml(cRange)})
          bị trùng với lịch <strong>${escapeHtml(c.existingType)}</strong> (${escapeHtml(eRange)})
          của môn <strong>${escapeHtml(c.existingSection.tenHP)}</strong> (<code>${escapeHtml(c.existingSection.maLHP)}</code>).
        </li>`;
      });
      if (evalResult.sameSubjectSection) {
        listItems.push(
          `<li>Ngoài ra, môn này đã đăng ký ở lớp <code>${escapeHtml(evalResult.sameSubjectSection.maLHP)}</code>.</li>`
        );
      }
      parts.push(`<ul>${listItems.join('')}</ul>`);
    }

    conflictSummaryEl.innerHTML = parts.join('');

    // Tìm các lớp khác của cùng môn học để gợi ý chọn lại
    const otherSections = catalog.filter(
      (s) => s.maHP === candidate.maHP && s.maLHP !== candidate.maLHP
    );

    if (otherSections.length === 0) {
      conflictAlternativesEl.innerHTML = `
        <h3>Chọn lại lịch môn bị trùng này</h3>
        <p class="muted">Hiện không có lớp học phần nào khác cho môn <strong>${escapeHtml(candidate.tenHP)}</strong>. Bạn có thể hủy môn đang trùng ở danh sách đã đăng ký hoặc nhập thêm lớp từ Excel.</p>
      `;
    } else {
      const altCards = otherSections.map((alt) => {
        const altEval = evaluateCandidate(alt, candidate.maHP);
        const isOk = altEval.timeConflicts.length === 0;
        const altSess = getSectionSessions(alt);
        const ltText = altSess.ltList.length
          ? altSess.ltList
              .map(
                (lt) =>
                  `<span class="badge badge-lt">LT</span> ${escapeHtml(formatScheduleText(lt, 'LT'))} (${escapeHtml(lt.phong || 'Chưa rõ')})`
              )
              .join('<br>')
          : '<span class="muted">Không có LT</span>';
        const thText = altSess.thList.length
          ? altSess.thList
              .map(
                (th) =>
                  `<span class="badge badge-th">TH</span> ${escapeHtml(formatScheduleText(th, 'TH'))} (${escapeHtml(th.phong || 'Chưa rõ')})`
              )
              .join('<br>')
          : '<span class="muted">Không có TH</span>';

        const conflictNote = isOk
          ? '<span class="badge badge-ok">Phù hợp — Không trùng lịch</span>'
          : `<span class="badge badge-conflict">Cũng bị trùng (${escapeHtml(
              altEval.timeConflicts[0].existingSection.tenHP
            )})</span>`;

        const actionBtn = isOk
          ? `<button type="button" class="btn btn-sm btn-success" data-switch-lhp="${escapeHtml(alt.maLHP)}">Chọn lại lớp này</button>`
          : `<button type="button" class="btn btn-sm btn-ghost" disabled>Bị trùng lịch</button>`;

        return `
          <div class="alt-card ${isOk ? 'available' : 'blocked'}">
            <div>
              <div><strong>${escapeHtml(alt.maLHP)}</strong> ${conflictNote}</div>
              <div class="sched-meta" style="margin-top:0.25rem">${ltText} &nbsp;|&nbsp; ${thText}</div>
            </div>
            <div>${actionBtn}</div>
          </div>
        `;
      });

      conflictAlternativesEl.innerHTML = `
        <h3>Gợi ý chọn lại lịch cho môn ${escapeHtml(candidate.tenHP)}</h3>
        <div class="alt-list">${altCards.join('')}</div>
      `;
    }

    conflictDialog.showModal();
  }

  // --- Rendering ---
  function formatWeekSummary(tuanStr) {
    const nums = Array.from(parseWeekSet(tuanStr)).sort((a, b) => a - b);
    if (nums.length === 0) return String(tuanStr || '1-15');
    const ranges = [];
    let start = nums[0];
    let prev = nums[0];
    for (let i = 1; i < nums.length; i++) {
      if (nums[i] === prev + 1) {
        prev = nums[i];
      } else {
        ranges.push(start === prev ? `${start}` : `${start}–${prev}`);
        start = nums[i];
        prev = nums[i];
      }
    }
    ranges.push(start === prev ? `${start}` : `${start}–${prev}`);
    return ranges.join(', ');
  }

  function renderSessionCell(sessOrList, badgeClass, badgeLabel) {
    const list = Array.isArray(sessOrList) ? sessOrList : sessOrList ? [sessOrList] : [];
    if (list.length === 0) return '<span class="muted">—</span>';
    return list
      .map((sess) => {
        const dateNote = sess.batDauText
          ? sess.ketThucText
            ? `${sess.batDauText} – ${sess.ketThucText}`
            : `Bắt đầu: ${sess.batDauText}`
          : '';
        return `
      <div class="sched-block" style="margin-bottom:0.35rem;">
        <div><span class="badge ${badgeClass}">${badgeLabel}</span> <strong>${escapeHtml(formatScheduleText(sess, badgeLabel))}</strong></div>
        <div class="sched-meta">Phòng: ${escapeHtml(sess.phong || '—')} • Tuần: ${escapeHtml(formatWeekSummary(sess.tuan || '1-15'))}</div>
        ${dateNote ? `<div class="sched-meta">Thời gian: ${escapeHtml(dateNote)}</div>` : ''}
        <div class="sched-meta">GV: ${escapeHtml(sess.gv || '—')}</div>
      </div>
    `;
      })
      .join('');
  }

  function populateFilterOptions() {
    const prevSubject = filterSubject.value;
    const prevClass = filterClass.value;
    const prevLecturer = filterLecturer.value;

    const subjects = Array.from(new Set(catalog.map((s) => s.tenHP).filter(Boolean))).sort((a, b) =>
      a.localeCompare(b, 'vi')
    );
    const classes = Array.from(new Set(catalog.map((s) => s.lopHoc).filter(Boolean))).sort((a, b) =>
      a.localeCompare(b, 'vi')
    );
    const lecturerSet = new Set();
    for (const s of catalog) {
      for (const item of getSectionSessions(s).all) {
        if (item.sess.gv) lecturerSet.add(item.sess.gv);
      }
    }
    const lecturers = Array.from(lecturerSet).sort((a, b) => a.localeCompare(b, 'vi'));

    filterSubject.innerHTML =
      '<option value="">Tất cả môn học</option>' +
      subjects.map((sub) => `<option value="${escapeHtml(sub)}">${escapeHtml(sub)}</option>`).join('');
    filterClass.innerHTML =
      '<option value="">Tất cả lớp học</option>' +
      classes.map((cls) => `<option value="${escapeHtml(cls)}">${escapeHtml(cls)}</option>`).join('');
    filterLecturer.innerHTML =
      '<option value="">Tất cả giảng viên</option>' +
      lecturers.map((gv) => `<option value="${escapeHtml(gv)}">${escapeHtml(gv)}</option>`).join('');

    if (subjects.includes(prevSubject)) filterSubject.value = prevSubject;
    if (classes.includes(prevClass)) filterClass.value = prevClass;
    if (lecturers.includes(prevLecturer)) filterLecturer.value = prevLecturer;
  }

  function sessionMatchesPeriodRange(sess, sessionType) {
    if (!sess) return false;
    if (sessionType === 'morning') return sess.start <= 6;
    if (sessionType === 'afternoon') return sess.end >= 7 && sess.start <= 12;
    if (sessionType === 'evening') return sess.end >= 13;
    return true;
  }

  function renderCatalog() {
    if (catalog.length === 0) {
      catalogCount.textContent = 'Chưa có dữ liệu học phần';
      catalogTbody.innerHTML = `
        <tr>
          <td colspan="7" style="text-align:center; padding: 2rem; color: var(--text-muted);">
            Chưa có danh sách học phần. Vui lòng sang tab
            <button type="button" class="btn btn-sm btn-outline" id="btn-goto-excel" style="margin-left:0.35rem;">File Excel &amp; Quy định</button>
            để tải file lên.
          </td>
        </tr>
      `;
      const gotoBtn = document.getElementById('btn-goto-excel');
      if (gotoBtn) gotoBtn.addEventListener('click', () => switchView('excel'));
      return;
    }

    const rawQuery = searchInput.value.trim();
    const normQuery = stripDiacritics(rawQuery);
    const subjectVal = filterSubject.value;
    const classVal = filterClass.value;
    const lecturerVal = filterLecturer.value;
    const dayVal = filterDay.value ? Number(filterDay.value) : null;
    const sessionVal = filterSession.value;
    const filterMode = filterStatus.value;

    const matchesOtherFilters = (sec) => {
      const secSessions = getSectionSessions(sec).all;
      if (subjectVal && sec.tenHP !== subjectVal) return false;
      if (classVal && (sec.lopHoc || '') !== classVal) return false;
      if (lecturerVal && !secSessions.some((i) => i.sess.gv === lecturerVal)) return false;

      if (dayVal && !secSessions.some((i) => Number(i.sess.thu) === dayVal)) return false;

      if (sessionVal && !secSessions.some((i) => sessionMatchesPeriodRange(i.sess, sessionVal))) return false;

      const isRegistered = registered.some((r) => r.maLHP === sec.maLHP);
      const evalRes = evaluateCandidate(sec);

      if (filterMode === 'valid') return isRegistered || !evalRes.hasConflict;
      if (filterMode === 'conflict') return !isRegistered && evalRes.hasConflict;
      if (filterMode === 'registered') return isRegistered;
      if (filterMode === 'elearning') {
        const allText = [
          sec.ghiChu || '',
          sec.tinhChat || '',
          sec.tenHP || '',
          sec.lopHoc || '',
          ...secSessions.flatMap((i) => [i.sess.phong || '', i.sess.gv || '']),
        ].join(' ');
        return /e[_\s-]?learn|elen|huitx/i.test(allText);
      }
      return true;
    };

    const baseList = catalog.filter(matchesOtherFilters);
    let usedFuzzyFallback = false;

    let filtered = baseList;
    if (normQuery) {
      const exactMatches = baseList.filter((sec) => {
        const secSessions = getSectionSessions(sec).all;
        const haystack = stripDiacritics(
          [
            sec.maHP,
            sec.tenHP,
            sec.maLHP,
            sec.lopHoc || '',
            sec.ghiChu || '',
            ...secSessions.flatMap((i) => [i.sess.gv || '', i.sess.phong || '']),
          ].join(' ')
        );
        return haystack.includes(normQuery);
      });

      if (exactMatches.length > 0) {
        filtered = exactMatches;
      } else {
        // Fallback sang tìm kiếm gần đúng ("na ná") nếu nhập thiếu hoặc sai chính tả nhẹ
        usedFuzzyFallback = true;
        filtered = baseList
          .map((sec) => {
            const secSessions = getSectionSessions(sec).all;
            const fields = [
              sec.tenHP,
              sec.lopHoc || '',
              sec.maLHP,
              sec.maHP,
              ...secSessions.flatMap((i) => [i.sess.gv || '', i.sess.phong || '']),
            ];
            const bestScore = Math.max(...fields.map((f) => fuzzyScore(rawQuery, f)));
            return { sec, bestScore };
          })
          .filter((x) => x.bestScore > 0)
          .sort((a, b) => b.bestScore - a.bestScore)
          .map((x) => x.sec);
      }
    }

    catalogCount.textContent = usedFuzzyFallback
      ? `Không khớp chính xác "${rawQuery}" — đang hiển thị ${filtered.length} gợi ý gần đúng`
      : `Hiển thị ${filtered.length} / ${catalog.length} lớp học phần`;

    if (filtered.length === 0) {
      catalogTbody.innerHTML = `
        <tr>
          <td colspan="7" style="text-align:center; padding: 2rem; color: var(--text-muted);">
            Không tìm thấy lớp học phần phù hợp.
          </td>
        </tr>
      `;
      return;
    }

    catalogTbody.innerHTML = filtered
      .map((sec) => {
        const isRegistered = registered.some((r) => r.maLHP === sec.maLHP);
        const evalRes = evaluateCandidate(sec);

        let rowClass = '';
        let statusHtml = '<span class="badge badge-ok">Hợp lệ</span>';
        let actionHtml = `<button type="button" class="btn btn-sm btn-primary" data-register="${escapeHtml(sec.maLHP)}">Đăng ký</button>`;

        if (isRegistered) {
          rowClass = 'row-registered';
          statusHtml = '<span class="badge badge-registered">Đã đăng ký</span>';
          actionHtml = `<button type="button" class="btn btn-sm btn-danger-outline" data-drop="${escapeHtml(sec.maLHP)}">Hủy</button>`;
        } else if (evalRes.timeConflicts.length > 0) {
          rowClass = 'row-conflict';
          const firstC = evalRes.timeConflicts[0];
          statusHtml = `
            <span class="badge badge-conflict">Trùng lịch</span>
            <div class="conflict-hint">Trùng ${escapeHtml(firstC.candidateCode)} với ${escapeHtml(firstC.existingSection.tenHP)}</div>
          `;
          actionHtml = `<button type="button" class="btn btn-sm btn-warning-outline" data-register="${escapeHtml(sec.maLHP)}">Xem &amp; Chọn lại</button>`;
        } else if (evalRes.sameSubjectSection) {
          statusHtml = `
            <span class="badge badge-registered">Đã có lớp khác</span>
            <div class="conflict-hint">Đang chọn: ${escapeHtml(evalRes.sameSubjectSection.maLHP)}</div>
          `;
          actionHtml = `<button type="button" class="btn btn-sm btn-outline" data-switch-lhp="${escapeHtml(sec.maLHP)}">Đổi sang lớp này</button>`;
        }

        const secSess = getSectionSessions(sec);
        const showMaHP = sec.maHP && sec.maHP !== normalizeHeader(sec.tenHP);
        const subMeta = [
          sec.lopHoc ? `Lớp: ${escapeHtml(sec.lopHoc)}` : '',
          showMaHP ? `Mã HP: ${escapeHtml(sec.maHP)}` : '',
        ]
          .filter(Boolean)
          .join(' • ');
        return `
          <tr class="${rowClass}">
            <td>
              <div class="course-name">${escapeHtml(sec.tenHP)}</div>
              ${subMeta ? `<div class="course-code">${subMeta}</div>` : ''}
              ${sec.ghiChu ? `<div><span class="badge badge-elearning">${escapeHtml(sec.ghiChu)}</span></div>` : ''}
            </td>
            <td><code>${escapeHtml(sec.maLHP)}</code></td>
            <td><strong>${escapeHtml(sec.soTC)}</strong></td>
            <td>${renderSessionCell(secSess.ltList, 'badge-lt', 'LT')}</td>
            <td>${renderSessionCell(secSess.thList, 'badge-th', 'TH')}</td>
            <td>${statusHtml}</td>
            <td>${actionHtml}</td>
          </tr>
        `;
      })
      .join('');
  }

  function renderRegistered() {
    statCourses.textContent = String(registered.length);
    if (tabRegCount) tabRegCount.textContent = String(registered.length);
    const totalCredits = registered.reduce((sum, item) => sum + (Number(item.soTC) || 0), 0);
    statCredits.textContent = String(totalCredits);

    if (registered.length === 0) {
      registeredListEl.innerHTML = `
        <div class="registered-empty">
          Chưa có học phần nào được đăng ký.<br>
          Hãy chọn lớp ở bảng bên trái để xếp Thời khóa biểu.
        </div>
      `;
      return;
    }

    registeredListEl.innerHTML = registered
      .map((sec) => {
        const secSess = getSectionSessions(sec);
        const schedLines = [
          ...secSess.ltList.map(
            (lt) =>
              `<div><span class="badge badge-lt">LT</span> ${escapeHtml(formatScheduleText(lt, 'LT'))} • ${escapeHtml(lt.phong || '—')}${lt.gv ? ` • GV: ${escapeHtml(lt.gv)}` : ''}</div>`
          ),
          ...secSess.thList.map(
            (th) =>
              `<div><span class="badge badge-th">TH</span> ${escapeHtml(formatScheduleText(th, 'TH'))} • ${escapeHtml(th.phong || '—')}${th.gv ? ` • GV: ${escapeHtml(th.gv)}` : ''}</div>`
          ),
        ].join('');

        return `
          <article class="reg-item">
            <div class="reg-item-top">
              <div>
                <div class="reg-item-title">${escapeHtml(sec.tenHP)}</div>
                <div class="course-code">${escapeHtml(sec.maLHP)} • ${escapeHtml(sec.soTC)} TC</div>
              </div>
              <button type="button" class="btn btn-sm btn-danger-outline" data-drop="${escapeHtml(sec.maLHP)}">Hủy</button>
            </div>
            <div class="reg-item-schedules">
              ${schedLines}
            </div>
          </article>
        `;
      })
      .join('');
  }

  function renderTimetable() {
    // Xây dựng ma trận chiếm chỗ: occupied[thu][period] = { section, typeCode, span, isStart }
    const occupied = {};
    for (let thu = 2; thu <= 8; thu++) occupied[thu] = {};

    for (const sec of registered) {
      const secSess = getSectionSessions(sec);
      const items = [
        ...secSess.ltList.map((sess) => ({ typeCode: 'lt', label: 'LT', sess })),
        ...secSess.thList.map((sess) => ({ typeCode: 'th', label: 'TH', sess })),
      ];

      for (const item of items) {
        const { thu, start, end } = item.sess;
        if (!occupied[thu]) continue;
        const s = Math.max(1, Math.min(16, Number(start)));
        const e = Math.max(s, Math.min(16, Number(end)));
        occupied[thu][s] = {
          sec,
          typeCode: item.typeCode,
          label: item.label,
          sess: item.sess,
          span: e - s + 1,
          isStart: true,
        };
        for (let p = s + 1; p <= e; p++) {
          occupied[thu][p] = { isStart: false };
        }
      }
    }

    const rows = [];
    for (let p = 1; p <= 16; p++) {
      if (p === 1) rows.push('<tr class="session-break"><td colspan="8">Buổi Sáng (Tiết 1 – 6 • LT nghỉ giải lao 30p sau tiết 3)</td></tr>');
      if (p === 7) rows.push('<tr class="session-break"><td colspan="8">Buổi Chiều (Tiết 7 – 12 • LT nghỉ giải lao 30p sau tiết 9)</td></tr>');
      if (p === 13) rows.push('<tr class="session-break"><td colspan="8">Buổi Tối (Tiết 13 – 16)</td></tr>');

      const ltSlot = HUIT_LT_SLOTS[p];
      const thSlot = HUIT_TH_SLOTS[p];
      const cells = [
        `<td class="period-cell"><div>Tiết ${p}</div><div class="sched-meta" style="font-weight:400;font-size:0.68rem">LT: ${ltSlot.start}<br>TH: ${thSlot.start}</div></td>`,
      ];
      for (let thu = 2; thu <= 8; thu++) {
        const slot = occupied[thu][p];
        if (!slot) {
          cells.push('<td></td>');
        } else if (!slot.isStart) {
          // Covered by rowspan above
          continue;
        } else {
          cells.push(`
            <td rowspan="${slot.span}">
              <div class="tt-slot ${slot.typeCode}">
                <div class="tt-slot-title">[${slot.label}] ${escapeHtml(slot.sec.tenHP)}</div>
                <div class="tt-slot-meta">${escapeHtml(getSessionClockRange(slot.sess, slot.label))}</div>
                <div class="tt-slot-meta">${escapeHtml(slot.sec.maLHP)} • ${escapeHtml(slot.sess.phong || '—')}</div>
                ${slot.sess.gv ? `<div class="tt-slot-meta">GV: ${escapeHtml(slot.sess.gv)}</div>` : ''}
                <div class="tt-slot-meta">Tuần: ${escapeHtml(formatWeekSummary(slot.sess.tuan || '1-15'))}</div>
              </div>
            </td>
          `);
        }
      }
      rows.push(`<tr>${cells.join('')}</tr>`);
    }

    timetableTbody.innerHTML = rows.join('');
  }

  function getAllTuitionRows() {
    return [
      ...registered.map((sec) => ({ ...sec, isCustom: false })),
      ...customTuitionItems.map((item) => ({
        maLHP: item.id,
        tenHP: item.tenHP,
        soTC: item.soTC,
        defaultType: item.type === 'th' ? 'th' : 'lt',
        isCustom: true,
      })),
    ];
  }

  function getTuitionSelection(sec) {
    const existing = tuitionSelections[sec.maLHP];
    if (existing && typeof existing === 'object') {
      return {
        checked: existing.checked !== false,
        type: existing.type === 'lt' || existing.type === 'th' ? existing.type : 'lt',
      };
    }
    if (sec.isCustom) {
      return {
        checked: true,
        type: sec.defaultType === 'th' ? 'th' : 'lt',
      };
    }
    const secSess = getSectionSessions(sec);
    return {
      checked: true,
      type: secSess.thList.length > 0 ? 'th' : 'lt',
    };
  }

  function saveTuitionSelections() {
    try {
      localStorage.setItem(STORAGE_KEYS.TUITION_SELECTIONS, JSON.stringify(tuitionSelections));
      saveState();
    } catch {
      // Ignore
    }
  }

  function formatVnd(amount) {
    return `${Math.round(amount).toLocaleString('vi-VN')} đ`;
  }

  function renderTuition() {
    if (!tuitionTbody) return;
    const rateLt = Math.max(0, Number(tuitionRateLtInput?.value) || 0);
    const rateTh = Math.max(0, Number(tuitionRateInput?.value) || 0);
    const allRows = getAllTuitionRows();

    if (allRows.length === 0) {
      if (tuitionCheckAll) tuitionCheckAll.checked = false;
      if (tuitionTotalLtCreditsEl) tuitionTotalLtCreditsEl.textContent = '0';
      tuitionTotalThCreditsEl.textContent = '0';
      tuitionTotalAmountEl.textContent = '0 đ';
      tuitionTbody.innerHTML = `
        <tr>
          <td colspan="6" style="text-align:center;padding:2rem;color:var(--text-muted);">
            Chưa có học phần nào. Hãy chọn học phần ở tab <strong>Đăng ký học phần</strong> hoặc thêm môn chưa có lịch ở khung phía trên.
          </td>
        </tr>
      `;
      return;
    }

    let totalLtCredits = 0;
    let totalThCredits = 0;
    let totalAmount = 0;
    let allChecked = true;

    tuitionTbody.innerHTML = allRows
      .map((sec) => {
        const sel = getTuitionSelection(sec);
        if (!sel.checked) allChecked = false;
        const credits = Number(sec.soTC) || 0;
        const unitRate = sel.type === 'th' ? rateTh : rateLt;
        const rowAmount = sel.checked ? unitRate * credits : 0;

        if (sel.checked) {
          if (sel.type === 'th') totalThCredits += credits;
          else totalLtCredits += credits;
          totalAmount += rowAmount;
        }

        const radioName = `tuition-type-${sec.maLHP}`;
        const subInfo = sec.isCustom
          ? `<div class="course-code">Tự thêm • Chưa có lịch <button type="button" class="btn btn-sm btn-danger-outline" style="padding:0.1rem 0.45rem;font-size:0.72rem;margin-left:0.35rem;" data-remove-custom-tuition="${escapeHtml(sec.maLHP)}">Xóa</button></div>`
          : sec.lopHoc
          ? `<div class="course-code">Lớp: ${escapeHtml(sec.lopHoc)}</div>`
          : '';

        return `
          <tr class="${sel.checked ? '' : 'tuition-row-unchecked'}">
            <td class="col-check">
              <input type="checkbox" ${sel.checked ? 'checked' : ''} data-tuition-check="${escapeHtml(sec.maLHP)}" aria-label="Tích chọn ${escapeHtml(sec.tenHP)}">
            </td>
            <td>
              <div class="course-name">${escapeHtml(sec.tenHP)}</div>
              ${subInfo}
            </td>
            <td><code>${sec.isCustom ? 'Tự thêm' : escapeHtml(sec.maLHP)}</code></td>
            <td><strong>${escapeHtml(credits)}</strong></td>
            <td>
              <div class="tuition-choice-group">
                <label class="tuition-radio-label">
                  <input type="radio" name="${escapeHtml(radioName)}" value="lt" ${sel.type === 'lt' ? 'checked' : ''} data-tuition-type="${escapeHtml(sec.maLHP)}">
                  <span>Lý thuyết</span>
                </label>
                <label class="tuition-radio-label">
                  <input type="radio" name="${escapeHtml(radioName)}" value="th" ${sel.type === 'th' ? 'checked' : ''} data-tuition-type="${escapeHtml(sec.maLHP)}">
                  <span>Thực hành</span>
                </label>
              </div>
            </td>
            <td><strong>${formatVnd(rowAmount)}</strong></td>
          </tr>
        `;
      })
      .join('');

    if (tuitionCheckAll) tuitionCheckAll.checked = allChecked;
    if (tuitionTotalLtCreditsEl) tuitionTotalLtCreditsEl.textContent = String(totalLtCredits);
    tuitionTotalThCreditsEl.textContent = String(totalThCredits);
    tuitionTotalAmountEl.textContent = formatVnd(totalAmount);
  }

  function renderAll() {
    populateFilterOptions();
    renderCatalog();
    renderRegistered();
    renderTimetable();
    renderCalendar();
    renderTuition();
  }

  // --- Excel Helpers ---
  function showImportStatus(message, type = 'info') {
    importStatusEl.textContent = message;
    importStatusEl.className = `import-status ${type}`;
    importStatusEl.hidden = false;
  }

  function normalizeHeader(h) {
    return String(h || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/đ/g, 'd')
      .replace(/Đ/g, 'D')
      .toLowerCase()
      .replace(/[^a-z0-9]/g, '');
  }

  function parseThuValue(val) {
    if (val === undefined || val === null || String(val).trim() === '') return null;
    const raw = normalizeHeader(val);
    if (raw === 'cn' || raw === 'chunhat' || raw === '8') return 8;
    const m = raw.match(/([2-8])/);
    return m ? Number(m[1]) : null;
  }

  function parseTietRange(val) {
    if (val === undefined || val === null || String(val).trim() === '') return null;
    const raw = String(val).trim();
    const range = raw.match(/(\d+)\s*[-–>:]+\s*(\d+)/);
    if (range) {
      const a = Number(range[1]);
      const b = Number(range[2]);
      return {
        start: Math.max(1, Math.min(16, Math.min(a, b))),
        end: Math.max(1, Math.min(16, Math.max(a, b))),
      };
    }
    const nums = raw.match(/\d+/g);
    if (nums && nums.length > 0) {
      const parsed = nums.map(Number).filter((n) => n >= 1 && n <= 16);
      if (parsed.length > 0) {
        return { start: Math.min(...parsed), end: Math.max(...parsed) };
      }
    }
    return null;
  }

  function pickField(rowMap, keys) {
    for (const k of keys) {
      if (rowMap[k] !== undefined && String(rowMap[k]).trim() !== '') {
        return String(rowMap[k]).trim();
      }
    }
    return '';
  }

  function expandSheetMergesToMatrix(sheet) {
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
    const merges = sheet['!merges'] || [];
    for (const m of merges) {
      const topVal = matrix[m.s.r]?.[m.s.c] ?? '';
      for (let r = m.s.r; r <= m.e.r; r++) {
        if (!matrix[r]) matrix[r] = [];
        for (let c = m.s.c; c <= m.e.c; c++) {
          if (matrix[r][c] === undefined || matrix[r][c] === '') {
            matrix[r][c] = topVal;
          }
        }
      }
    }
    return matrix;
  }

  function isHuitTheoryRoom(phong) {
    const raw = String(phong || '').trim();
    if (!raw) return null;
    // Lý thuyết HUIT: Zoom, hoặc phòng B + số, A + số, V. + số, G + số (VD: A201, B104, V.202, G301, Zoom01)
    if (/zoom/i.test(raw)) return true;
    if (/^(?:p\.?\s*)?[abvg]\.?\s*\d+/i.test(raw)) return true;
    return false;
  }

  function spansBreakContinuously(sess) {
    if (!sess) return false;
    // Lịch học liên tiếp qua mốc giải lao (tiết 3->4 hoặc tiết 9->10) là Thực hành không nghỉ giải lao
    return (sess.start <= 3 && sess.end >= 4) || (sess.start <= 9 && sess.end >= 10);
  }

  function classifyRowScheduleType(tinhChat, phong, existingSection, sessObj) {
    if (spansBreakContinuously(sessObj)) return 'th';

    const roomCheck = isHuitTheoryRoom(phong);
    if (roomCheck === true) return 'lt';
    if (roomCheck === false) return 'th';

    // Nếu ô Phòng trống, xét theo Tính chất học phần hoặc thứ tự dòng
    const tcNorm = normalizeHeader(tinhChat);
    if (tcNorm.includes('thuchanh') || tcNorm === 'th') return 'th';
    if (tcNorm.includes('lythuyet') || tcNorm === 'lt') return 'lt';
    if (existingSection && existingSection.lt && !existingSection.th) return 'th';
    return 'lt';
  }

  function parseHuitMatrix(matrix) {
    // Tìm dòng tiêu đề chứa Mã LHP và Tên môn học (trong 25 dòng đầu)
    let headerRowIdx = -1;
    for (let r = 0; r < Math.min(25, matrix.length); r++) {
      const rowNorm = (matrix[r] || []).map(normalizeHeader);
      const hasMaLHP = rowNorm.some((h) => h === 'malhp' || h === 'malophocphan');
      const hasTenMon = rowNorm.some((h) => h === 'tenmonhoc' || h === 'tenhp' || h === 'tenhocphan');
      if (hasMaLHP && hasTenMon) {
        headerRowIdx = r;
        break;
      }
    }
    if (headerRowIdx === -1) return null;

    const row1 = matrix[headerRowIdx] || [];
    const row2 = matrix[headerRowIdx + 1] || [];
    // Kiểm tra xem có phải tiêu đề 2 tầng (Lịch học -> Thứ, Từ tiết, Đến tiết...) không
    const row2Norm = row2.map(normalizeHeader);
    const hasSubHeader = row2Norm.some(
      (h) => h === 'thu' || h === 'tutiet' || h === 'dentiet' || h === 'batdau'
    );
    const dataStartIdx = hasSubHeader ? headerRowIdx + 2 : headerRowIdx + 1;

    const colCount = Math.max(row1.length, row2.length);
    const colIdx = {
      maLHP: -1,
      maHP: -1,
      tenHP: -1,
      tinhChat: -1,
      phong: -1,
      soTC: -1,
      lopHoc: -1,
      thu: -1,
      tuTiet: -1,
      denTiet: -1,
      batDau: -1,
      ketThuc: -1,
      giangVien: -1,
      ghiChu: -1,
      weekCols: [],
    };

    for (let c = 0; c < colCount; c++) {
      const h1 = normalizeHeader(row1[c]);
      const h2 = hasSubHeader ? normalizeHeader(row2[c]) : '';
      const key = h2 && h2 !== h1 ? h2 : h1;

      if ((key === 'malhp' || key === 'malophocphan') && colIdx.maLHP === -1) colIdx.maLHP = c;
      else if ((key === 'mahp' || key === 'mahocphan') && colIdx.maHP === -1) colIdx.maHP = c;
      else if ((key === 'tenmonhoc' || key === 'tenhp' || key === 'tenhocphan') && colIdx.tenHP === -1) colIdx.tenHP = c;
      else if ((key === 'tinhchathocphan' || key === 'tinhchat' || key === 'loai') && colIdx.tinhChat === -1) colIdx.tinhChat = c;
      else if ((key === 'phong' || key === 'phonghoc') && colIdx.phong === -1) colIdx.phong = c;
      else if ((key === 'sotc' || key === 'tc' || key === 'tinchi') && colIdx.soTC === -1) colIdx.soTC = c;
      else if ((key === 'lophoc' || key === 'lop') && colIdx.lopHoc === -1) colIdx.lopHoc = c;
      else if (key === 'thu' && colIdx.thu === -1) colIdx.thu = c;
      else if ((key === 'tutiet' || key === 'tietbd') && colIdx.tuTiet === -1) colIdx.tuTiet = c;
      else if ((key === 'dentiet' || key === 'tietkt') && colIdx.denTiet === -1) colIdx.denTiet = c;
      else if (key === 'batdau' && colIdx.batDau === -1) colIdx.batDau = c;
      else if (key === 'ketthuc' && colIdx.ketThuc === -1) colIdx.ketThuc = c;
      else if ((key === 'giangvien' || key === 'gv') && colIdx.giangVien === -1) colIdx.giangVien = c;
      else if (
        (key === 'ghichu' || key === 'hinhthuc' || key === 'hinhthuchoc' || key === 'elearning' || key === 'huitx' || key === 'note') &&
        colIdx.ghiChu === -1
      ) {
        colIdx.ghiChu = c;
      } else if (
        colIdx.batDau !== -1 &&
        colIdx.ketThuc === -1 &&
        (/^t0*\d+$/.test(key) || (h1.includes('thoigianhoc') && /^\d+$/.test(h2)))
      ) {
        const numMatch = key.match(/\d+/);
        const weekLabel = numMatch ? `T${Number(numMatch[0])}` : `T${colIdx.weekCols.length + 1}`;
        colIdx.weekCols.push({ col: c, label: weekLabel });
      }
    }

    // Nếu cột cuối cùng (sau Giảng viên) không có tiêu đề nhưng chứa dữ liệu như E_learning(HUITx)
    if (colIdx.ghiChu === -1 && colIdx.giangVien !== -1) {
      for (let c = colIdx.giangVien + 1; c < colCount + 3; c++) {
        const hasNoteData = matrix
          .slice(dataStartIdx, Math.min(matrix.length, dataStartIdx + 80))
          .some((r) => r && String(r[c] || '').trim() !== '');
        if (hasNoteData) {
          colIdx.ghiChu = c;
          break;
        }
      }
    }

    if (colIdx.maLHP === -1 || colIdx.tenHP === -1 || colIdx.thu === -1 || colIdx.tuTiet === -1) {
      return null;
    }

    const grouped = new Map();
    const errors = [];

    for (let r = dataStartIdx; r < matrix.length; r++) {
      const row = matrix[r] || [];
      const maLHP = String(row[colIdx.maLHP] || '').trim();
      const tenHP = String(row[colIdx.tenHP] || '').trim();
      if (!maLHP && !tenHP) continue;

      const thu = parseThuValue(row[colIdx.thu]);
      const startTiet = Number(String(row[colIdx.tuTiet] || '').trim());
      const endTiet = colIdx.denTiet !== -1 ? Number(String(row[colIdx.denTiet] || '').trim()) || startTiet : startTiet;

      if (!maLHP || !tenHP || !thu || !startTiet || startTiet < 1 || startTiet > 16) {
        errors.push(`Dòng ${r + 1}: thiếu Mã LHP / Tên môn / Thứ / Tiết hợp lệ.`);
        continue;
      }

      const phong = colIdx.phong !== -1 ? String(row[colIdx.phong] || '').trim() : '';
      const tinhChat = colIdx.tinhChat !== -1 ? String(row[colIdx.tinhChat] || '').trim() : '';
      const lopHoc = colIdx.lopHoc !== -1 ? String(row[colIdx.lopHoc] || '').trim() : '';
      const gv = colIdx.giangVien !== -1 ? String(row[colIdx.giangVien] || '').trim() : '';
      let ghiChu = colIdx.ghiChu !== -1 ? String(row[colIdx.ghiChu] || '').trim() : '';
      const elenCell = row.find((cell) => /e[_\s-]?learn|elen|huitx/i.test(String(cell || '')));
      if (elenCell && !/e[_\s-]?learn|elen|huitx/i.test(ghiChu)) {
        ghiChu = ghiChu ? `${ghiChu} • ${String(elenCell).trim()}` : String(elenCell).trim();
      }
      const soTC = colIdx.soTC !== -1 ? Number(row[colIdx.soTC]) || 3 : 3;

      // Xác định tuần học từ các cột T... nếu có đánh dấu, ngược lại dùng 1-15
      let tuan = '1-15';
      if (colIdx.weekCols.length > 0) {
        const activeWeeks = [];
        colIdx.weekCols.forEach((wc, idx) => {
          const val = String(row[wc.col] || '').trim();
          if (val !== '') {
            const numMatch = wc.label.match(/\d+/);
            activeWeeks.push(numMatch ? Number(numMatch[0]) : idx + 1);
          }
        });
        if (activeWeeks.length > 0) {
          tuan = activeWeeks.join(',');
        }
      }

      const maHP =
        colIdx.maHP !== -1 && String(row[colIdx.maHP] || '').trim()
          ? String(row[colIdx.maHP]).trim()
          : normalizeHeader(tenHP);

      let sec = grouped.get(maLHP);
      if (!sec) {
        sec = { maHP, tenHP, maLHP, lopHoc, tinhChat, ghiChu, soTC, ltList: [], thList: [] };
        grouped.set(maLHP, sec);
      } else {
        if (!sec.lopHoc && lopHoc) sec.lopHoc = lopHoc;
        if (!sec.tinhChat && tinhChat) sec.tinhChat = tinhChat;
        if (!sec.ghiChu && ghiChu) sec.ghiChu = ghiChu;
      }

      const batDauParsed = colIdx.batDau !== -1 ? parseExcelDate(row[colIdx.batDau]) : null;
      const ketThucParsed = colIdx.ketThuc !== -1 ? parseExcelDate(row[colIdx.ketThuc]) : null;

      const sessObj = {
        thu,
        start: Math.max(1, Math.min(16, Math.min(startTiet, endTiet))),
        end: Math.max(1, Math.min(16, Math.max(startTiet, endTiet))),
        phong,
        gv,
        tuan,
        batDau: batDauParsed ? batDauParsed.iso : '',
        batDauText: batDauParsed ? batDauParsed.display : '',
        ketThuc: ketThucParsed ? ketThucParsed.iso : '',
        ketThucText: ketThucParsed ? ketThucParsed.display : '',
      };

      const kind = classifyRowScheduleType(tinhChat, phong, sec, sessObj);
      if (kind === 'th') {
        sec.thList.push(sessObj);
      } else {
        sec.ltList.push(sessObj);
      }
    }

    return { sections: Array.from(grouped.values()), errors };
  }

  function importExcelFile(file) {
    if (typeof XLSX === 'undefined') {
      showImportStatus('Không tải được thư viện đọc Excel (SheetJS). Vui lòng kiểm tra kết nối mạng.', 'error');
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

        // 1. Thử đọc theo định dạng Thời khóa biểu dự kiến HUIT (header 2 tầng + merge ô)
        const matrix = expandSheetMergesToMatrix(firstSheet);
        const huitResult = parseHuitMatrix(matrix);
        if (huitResult && huitResult.sections.length > 0) {
          catalog = huitResult.sections;
          const validCodes = new Set(catalog.map((s) => s.maLHP));
          registered = registered.filter((r) => validCodes.has(r.maLHP));
          saveState();
          renderAll();

          const warnText =
            huitResult.errors.length > 0
              ? ` (Bỏ qua ${huitResult.errors.length} dòng không đủ lịch: ${huitResult.errors[0]})`
              : '';
          showImportStatus(
            `Đã nhận ${huitResult.sections.length} lớp học phần theo chuẩn HUIT từ "${file.name}".${warnText}`,
            'success'
          );
          return;
        }

        // 2. Fallback: đọc theo định dạng bảng phẳng 1 dòng / lớp học phần
        const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });
        if (!rows.length) {
          showImportStatus('File Excel trống hoặc không có dữ liệu học phần.', 'error');
          return;
        }

        const parsedSections = [];
        const errors = [];

        rows.forEach((rawRow, idx) => {
          const rowMap = {};
          for (const [k, v] of Object.entries(rawRow)) {
            rowMap[normalizeHeader(k)] = v;
          }

          const maHP = pickField(rowMap, ['mahp', 'mahocphan', 'mamon', 'code']);
          const tenHP = pickField(rowMap, ['tenhp', 'tenhocphan', 'tenmon', 'tenmonhoc', 'subject']);
          const maLHP = pickField(rowMap, ['malhp', 'malophocphan', 'malop', 'nhom']) || `${maHP}-${idx + 1}`;
          const soTC = Number(pickField(rowMap, ['sotc', 'tc', 'tinchi', 'credits'])) || 3;

          if (!maHP || !tenHP) {
            errors.push(`Dòng ${idx + 2}: thiếu MaHP hoặc TenHP.`);
            return;
          }

          const thuLT = parseThuValue(pickField(rowMap, ['thult', 'thu', 'thulythuyet']));
          const tietLT = parseTietRange(pickField(rowMap, ['tietlt', 'tiet', 'tietlythuyet']));
          const lt =
            thuLT && tietLT
              ? {
                  thu: thuLT,
                  start: tietLT.start,
                  end: tietLT.end,
                  phong: pickField(rowMap, ['phonglt', 'phong', 'phonglythuyet']),
                  gv: pickField(rowMap, ['giangvienlt', 'gvlt', 'giangvien', 'gv']),
                  tuan: pickField(rowMap, ['tuanlt', 'tuan', 'tuanlythuyet']) || '1-15',
                }
              : null;

          const thuTH = parseThuValue(pickField(rowMap, ['thuth', 'thuthuchanh']));
          const tietTH = parseTietRange(pickField(rowMap, ['tietth', 'tietthuchanh']));
          const th =
            thuTH && tietTH
              ? {
                  thu: thuTH,
                  start: tietTH.start,
                  end: tietTH.end,
                  phong: pickField(rowMap, ['phongth', 'phongthuchanh']),
                  gv: pickField(rowMap, ['giangvienth', 'gvth', 'giangvienthuchanh']),
                  tuan: pickField(rowMap, ['tuanth', 'tuanthuchanh']) || '1-15',
                }
              : null;

          if (!lt && !th) {
            errors.push(`Dòng ${idx + 2} (${maLHP}): cần ít nhất lịch Lý thuyết hoặc Thực hành hợp lệ.`);
            return;
          }

          parsedSections.push({ maHP, tenHP, maLHP, soTC, lt, th });
        });

        if (parsedSections.length === 0) {
          showImportStatus(`Không nhập được lớp nào. ${errors.slice(0, 3).join(' ')}`, 'error');
          return;
        }

        catalog = parsedSections;
        const validCodes = new Set(catalog.map((s) => s.maLHP));
        registered = registered.filter((r) => validCodes.has(r.maLHP));
        saveState();
        renderAll();

        const warnText = errors.length > 0 ? ` (Bỏ qua ${errors.length} dòng lỗi: ${errors[0]})` : '';
        showImportStatus(`Đã nhận ${parsedSections.length} lớp học phần từ file "${file.name}".${warnText}`, 'success');
      } catch (err) {
        showImportStatus(`Lỗi đọc file Excel: ${err.message}`, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function exportRegisteredExcel() {
    if (typeof XLSX === 'undefined') {
      showImportStatus('Không tải được thư viện xuất Excel (SheetJS).', 'error');
      return;
    }
    if (registered.length === 0) {
      showImportStatus('Bạn chưa đăng ký học phần nào để xuất Thời khóa biểu.', 'error');
      return;
    }

    const listRows = registered.map((s, idx) => {
      const secSess = getSectionSessions(s);
      return {
        STT: idx + 1,
        MaHP: s.maHP,
        TenHP: s.tenHP,
        MaLHP: s.maLHP,
        LopHoc: s.lopHoc || '',
        SoTC: s.soTC,
        LichLyThuyet: secSess.ltList.length
          ? secSess.ltList.map((lt) => `${formatScheduleText(lt, 'LT')} (${lt.phong || '—'})`).join(' | ')
          : 'Không',
        GiangVienLT: Array.from(new Set(secSess.ltList.map((lt) => lt.gv).filter(Boolean))).join(', '),
        LichThucHanh: secSess.thList.length
          ? secSess.thList.map((th) => `${formatScheduleText(th, 'TH')} (${th.phong || '—'})`).join(' | ')
          : 'Không',
        GiangVienTH: Array.from(new Set(secSess.thList.map((th) => th.gv).filter(Boolean))).join(', '),
      };
    });

    // Sheet 2: Ma trận Thời khóa biểu (16 tiết HUIT)
    const ttRows = [];
    for (let p = 1; p <= 16; p++) {
      const row = {
        Tiet: `Tiết ${p}`,
        GioLT: `${HUIT_LT_SLOTS[p].start}–${HUIT_LT_SLOTS[p].end}`,
        GioTH: `${HUIT_TH_SLOTS[p].start}–${HUIT_TH_SLOTS[p].end}`,
      };
      for (let thu = 2; thu <= 8; thu++) {
        const colName = formatThu(thu);
        const matches = [];
        for (const s of registered) {
          const secSess = getSectionSessions(s);
          for (const lt of secSess.ltList) {
            if (Number(lt.thu) === thu && p >= lt.start && p <= lt.end) {
              matches.push(`[LT] ${s.tenHP} (${lt.phong || '—'})`);
            }
          }
          for (const th of secSess.thList) {
            if (Number(th.thu) === thu && p >= th.start && p <= th.end) {
              matches.push(`[TH] ${s.tenHP} (${th.phong || '—'})`);
            }
          }
        }
        row[colName] = matches.join(' | ');
      }
      ttRows.push(row);
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(listRows), 'DanhSachDangKy');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ttRows), 'ThoiKhoaBieu');
    XLSX.writeFile(wb, 'TKB_HUIT_Da_Dang_Ky.xlsx');
    showImportStatus('Đã xuất Thời khóa biểu ra file "TKB_HUIT_Da_Dang_Ky.xlsx".', 'success');
  }

  function exportRegisteredWord() {
    if (registered.length === 0) {
      showImportStatus('Bạn chưa đăng ký học phần nào để xuất Word.', 'error');
      return;
    }

    const totalCredits = registered.reduce((sum, item) => sum + (Number(item.soTC) || 0), 0);
    const listRowsHtml = registered
      .map((s, idx) => {
        const secSess = getSectionSessions(s);
        const ltHtml = secSess.ltList.length
          ? secSess.ltList.map((lt) => `${escapeHtml(formatScheduleText(lt, 'LT'))} (${escapeHtml(lt.phong || '—')})`).join('<br>')
          : '—';
        const thHtml = secSess.thList.length
          ? secSess.thList.map((th) => `${escapeHtml(formatScheduleText(th, 'TH'))} (${escapeHtml(th.phong || '—')})`).join('<br>')
          : '—';
        return `
          <tr>
            <td style="text-align:center;">${idx + 1}</td>
            <td><strong>${escapeHtml(s.tenHP)}</strong>${s.lopHoc ? `<br><small>Lớp: ${escapeHtml(s.lopHoc)}</small>` : ''}</td>
            <td>${escapeHtml(s.maLHP)}</td>
            <td style="text-align:center;">${escapeHtml(s.soTC)}</td>
            <td>${ltHtml}</td>
            <td>${thHtml}</td>
          </tr>
        `;
      })
      .join('');

    const ttTableHtml = document.getElementById('timetable-grid')?.outerHTML || '';
    const userNote = currentUser ? ` • Sinh viên: <strong>${escapeHtml(currentUser)}</strong>` : '';

    const docHtml = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
      <head>
        <meta charset="utf-8">
        <title>Thời Khóa Biểu HUIT</title>
        <style>
          body { font-family: Arial, sans-serif; font-size: 11pt; color: #0f172a; }
          h1 { color: #0f2a5f; font-size: 16pt; margin-bottom: 4pt; }
          h2 { color: #0f2a5f; font-size: 13pt; margin-top: 16pt; margin-bottom: 6pt; }
          table { width: 100%; border-collapse: collapse; margin-top: 6pt; }
          th, td { border: 1px solid #94a3b8; padding: 5pt 6pt; font-size: 9.5pt; vertical-align: top; }
          th { background: #0f2a5f; color: #ffffff; }
          .tt-slot.lt { background: #dbeafe; color: #1e3a8a; padding: 3pt; }
          .tt-slot.th { background: #dcfce7; color: #14532d; padding: 3pt; }
          .session-break td { background: #e2e8f0; font-weight: bold; text-align: center; }
        </style>
      </head>
      <body>
        <h1>THỜI KHÓA BIỂU ĐĂNG KÝ HỌC PHẦN DỰ KIẾN HUIT</h1>
        <p>Đã đăng ký: <strong>${registered.length}</strong> học phần • Tổng tín chỉ: <strong>${totalCredits}</strong> TC${userNote}</p>
        <h2>1. Danh sách học phần đã đăng ký</h2>
        <table>
          <thead>
            <tr>
              <th>STT</th><th>Tên học phần</th><th>Mã LHP</th><th>Số TC</th><th>Lịch Lý thuyết (LT)</th><th>Lịch Thực hành (TH)</th>
            </tr>
          </thead>
          <tbody>${listRowsHtml}</tbody>
        </table>
        <h2>2. Thời khóa biểu trực quan (16 tiết)</h2>
        ${ttTableHtml}
      </body>
      </html>
    `;

    const blob = new Blob(['\ufeff', docHtml], { type: 'application/msword;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'TKB_HUIT_Da_Dang_Ky.doc';
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    showImportStatus('Đã xuất Thời khóa biểu ra file Word "TKB_HUIT_Da_Dang_Ky.doc".', 'success');
  }

  async function exportRegisteredImage() {
    if (registered.length === 0) {
      showImportStatus('Bạn chưa đăng ký học phần nào để xuất ảnh.', 'error');
      return;
    }
    if (typeof html2canvas === 'undefined') {
      showImportStatus('Không tải được thư viện chụp ảnh (html2canvas). Vui lòng kiểm tra kết nối mạng.', 'error');
      return;
    }

    const ttSection = document.querySelector('#panel-timetable .timetable-section');
    const wasHidden = panelTimetable.hidden;
    if (wasHidden) panelTimetable.hidden = false;

    try {
      const canvas = await html2canvas(ttSection, {
        scale: 2,
        backgroundColor: getComputedStyle(document.body).getPropertyValue('--surface').trim() || '#ffffff',
        useCORS: true,
      });
      const url = canvas.toDataURL('image/png');
      const a = document.createElement('a');
      a.href = url;
      a.download = 'TKB_HUIT_Da_Dang_Ky.png';
      document.body.appendChild(a);
      a.click();
      a.remove();
      showImportStatus('Đã xuất Thời khóa biểu ra file ảnh "TKB_HUIT_Da_Dang_Ky.png".', 'success');
    } catch (err) {
      showImportStatus(`Lỗi xuất ảnh Thời khóa biểu: ${err.message}`, 'error');
    } finally {
      if (wasHidden) panelTimetable.hidden = true;
    }
  }

  // --- Event Listeners ---
  catalogTbody.addEventListener('click', (e) => {
    const regBtn = e.target.closest('[data-register]');
    if (regBtn) {
      attemptRegister(regBtn.dataset.register);
      return;
    }
    const dropBtn = e.target.closest('[data-drop]');
    if (dropBtn) {
      unregisterSection(dropBtn.dataset.drop);
      return;
    }
    const switchBtn = e.target.closest('[data-switch-lhp]');
    if (switchBtn) {
      switchToSection(switchBtn.dataset.switchLhp);
    }
  });

  registeredListEl.addEventListener('click', (e) => {
    const dropBtn = e.target.closest('[data-drop]');
    if (dropBtn) unregisterSection(dropBtn.dataset.drop);
  });

  conflictAlternativesEl.addEventListener('click', (e) => {
    const switchBtn = e.target.closest('[data-switch-lhp]');
    if (switchBtn) switchToSection(switchBtn.dataset.switchLhp);
  });

  conflictDialog.addEventListener('close', () => {
    conflictDialog.returnValue = '';
  });

  excelGuideDialog.addEventListener('close', () => {
    excelGuideDialog.returnValue = '';
  });

  if (!('closedBy' in HTMLDialogElement.prototype)) {
    const isOutsideDialog = (e, dlg) => {
      const r = dlg.getBoundingClientRect();
      return e.clientX < r.left || e.clientX > r.right || e.clientY < r.top || e.clientY > r.bottom;
    };
    for (const dlg of document.querySelectorAll('dialog[closedby="any"]')) {
      let pressedOutside = false;
      dlg.addEventListener('pointerdown', (e) => {
        pressedOutside = isOutsideDialog(e, dlg);
      });
      dlg.addEventListener('pointerup', (e) => {
        if (pressedOutside && isOutsideDialog(e, dlg)) dlg.close();
      });
    }
  }

  function applyTheme(mode) {
    if (mode === 'light' || mode === 'dark') {
      document.documentElement.setAttribute('data-theme', mode);
    } else {
      document.documentElement.removeAttribute('data-theme');
    }
    try {
      localStorage.setItem(STORAGE_KEYS.THEME, mode);
    } catch {
      // Ignore
    }
  }

  if (themeSelect) {
    const savedTheme = (() => {
      try {
        return localStorage.getItem(STORAGE_KEYS.THEME) || 'system';
      } catch {
        return 'system';
      }
    })();
    themeSelect.value = savedTheme;
    applyTheme(savedTheme);
    themeSelect.addEventListener('change', () => applyTheme(themeSelect.value));
  }

  if (tuitionRateLtInput) {
    try {
      const savedLtRate = localStorage.getItem(STORAGE_KEYS.TUITION_RATE_LT);
      if (savedLtRate !== null && savedLtRate !== '') tuitionRateLtInput.value = savedLtRate;
    } catch {
      // Ignore
    }
    tuitionRateLtInput.addEventListener('input', () => {
      try {
        localStorage.setItem(STORAGE_KEYS.TUITION_RATE_LT, tuitionRateLtInput.value);
        saveState();
      } catch {
        // Ignore
      }
      renderTuition();
    });
  }

  if (tuitionRateInput) {
    try {
      const savedRate = localStorage.getItem(STORAGE_KEYS.TUITION_RATE_TH);
      if (savedRate !== null && savedRate !== '') tuitionRateInput.value = savedRate;
    } catch {
      // Ignore
    }
    tuitionRateInput.addEventListener('input', () => {
      try {
        localStorage.setItem(STORAGE_KEYS.TUITION_RATE_TH, tuitionRateInput.value);
        saveState();
      } catch {
        // Ignore
      }
      renderTuition();
    });
  }

  if (tuitionCheckAll) {
    tuitionCheckAll.addEventListener('change', () => {
      const isChecked = tuitionCheckAll.checked;
      for (const sec of getAllTuitionRows()) {
        const cur = getTuitionSelection(sec);
        tuitionSelections[sec.maLHP] = { ...cur, checked: isChecked };
      }
      saveTuitionSelections();
      renderTuition();
    });
  }

  const customTuitionForm = document.getElementById('custom-tuition-form');
  const customTuitionName = document.getElementById('custom-tuition-name');
  const customTuitionCredits = document.getElementById('custom-tuition-credits');
  const customTuitionType = document.getElementById('custom-tuition-type');

  if (customTuitionForm) {
    customTuitionForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const tenHP = customTuitionName.value.trim();
      const soTC = Number(customTuitionCredits.value);
      const type = customTuitionType.value === 'th' ? 'th' : 'lt';
      if (!tenHP || !soTC || soTC <= 0) return;

      const id = `CUSTOM-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
      customTuitionItems.push({ id, tenHP, soTC, type });
      tuitionSelections[id] = { checked: true, type };
      customTuitionName.value = '';
      saveTuitionSelections();
      renderTuition();
    });
  }

  if (tuitionTbody) {
    tuitionTbody.addEventListener('click', (e) => {
      const removeBtn = e.target.closest('[data-remove-custom-tuition]');
      if (!removeBtn) return;
      const id = removeBtn.dataset.removeCustomTuition;
      customTuitionItems = customTuitionItems.filter((item) => item.id !== id);
      delete tuitionSelections[id];
      saveTuitionSelections();
      renderTuition();
    });

    tuitionTbody.addEventListener('change', (e) => {
      const checkEl = e.target.closest('[data-tuition-check]');
      if (checkEl) {
        const maLHP = checkEl.dataset.tuitionCheck;
        const sec = getAllTuitionRows().find((r) => r.maLHP === maLHP);
        if (sec) {
          const cur = getTuitionSelection(sec);
          tuitionSelections[maLHP] = { ...cur, checked: checkEl.checked };
          saveTuitionSelections();
          renderTuition();
        }
        return;
      }

      const radioEl = e.target.closest('[data-tuition-type]');
      if (radioEl) {
        const maLHP = radioEl.dataset.tuitionType;
        const sec = getAllTuitionRows().find((r) => r.maLHP === maLHP);
        if (sec) {
          const cur = getTuitionSelection(sec);
          tuitionSelections[maLHP] = { ...cur, type: radioEl.value === 'th' ? 'th' : 'lt' };
          saveTuitionSelections();
          renderTuition();
        }
      }
    });
  }

  // --- Monthly / Weekly / Daily Calendar ---
  const calPrevBtn = document.getElementById('cal-prev-btn');
  const calNextBtn = document.getElementById('cal-next-btn');
  const calTodayBtn = document.getElementById('cal-today-btn');
  const calHeading = document.getElementById('calendar-heading');
  const calGridContainer = document.getElementById('cal-grid-container');
  const calSelectedTitle = document.getElementById('cal-selected-title');
  const calSelectedItems = document.getElementById('cal-selected-items');
  const calEventForm = document.getElementById('cal-event-form');
  const calEventTitle = document.getElementById('cal-event-title');
  const calEventTime = document.getElementById('cal-event-time');
  const calModeBtns = Array.from(document.querySelectorAll('[data-cal-mode]'));

  let calMode = 'month';
  let calSelectedDate = new Date();
  let calViewYear = calSelectedDate.getFullYear();
  let calViewMonth = calSelectedDate.getMonth();

  function toDateKey(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  function jsDayToHuitThu(jsDay) {
    return jsDay === 0 ? 8 : jsDay + 1;
  }

  function getEarliestRegisteredStartDate() {
    let earliest = '';
    for (const sec of registered) {
      for (const item of getSectionSessions(sec).all) {
        if (item.sess.batDau && (!earliest || item.sess.batDau < earliest)) {
          earliest = item.sess.batDau;
        }
      }
    }
    if (!earliest) return null;
    const [y, m, d] = earliest.split('-').map(Number);
    return new Date(y, m - 1, d);
  }

  function isSessionActiveOnDate(sess, d, dateKey) {
    if (sess.batDau && dateKey < sess.batDau) return false;
    if (sess.ketThuc && dateKey > sess.ketThuc) return false;

    // Nếu có ngày bắt đầu và danh sách tuần cụ thể (không phải mặc định '1-15'), đối chiếu đúng tuần học
    if (sess.batDau && sess.tuan && sess.tuan !== '1-15') {
      const weekSet = parseWeekSet(sess.tuan);
      if (weekSet.size > 0) {
        const minWeek = Math.min(...weekSet);
        const [sy, sm, sd] = sess.batDau.split('-').map(Number);
        const startMs = Date.UTC(sy, sm - 1, sd);
        const curMs = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
        const diffDays = Math.round((curMs - startMs) / 86400000);
        if (diffDays < 0) return false;
        const weekOffset = Math.floor(diffDays / 7);
        if (!weekSet.has(minWeek + weekOffset)) return false;
      }
    }
    return true;
  }

  function getRegisteredClassesForDate(d) {
    const thu = jsDayToHuitThu(d.getDay());
    const dateKey = toDateKey(d);
    const list = [];
    for (const sec of registered) {
      const secSess = getSectionSessions(sec);
      for (const item of secSess.all) {
        if (Number(item.sess.thu) !== thu) continue;
        if (!isSessionActiveOnDate(item.sess, d, dateKey)) continue;
        list.push({
          kind: item.code.toLowerCase(),
          label: item.code,
          tenHP: sec.tenHP,
          maLHP: sec.maLHP,
          clock: getSessionClockRange(item.sess, item.code),
          phong: item.sess.phong,
          gv: item.sess.gv,
        });
      }
    }
    return list;
  }

  function getCustomEventsForDate(d) {
    const key = toDateKey(d);
    return Array.isArray(calendarEvents[key]) ? calendarEvents[key] : [];
  }

  function dateHasAnyItem(d) {
    return getCustomEventsForDate(d).length > 0 || getRegisteredClassesForDate(d).length > 0;
  }

  function renderCalendar() {
    if (!calGridContainer || !calHeading) return;
    calHeading.textContent = `tháng ${calViewMonth + 1}, ${calViewYear}`;

    const todayKey = toDateKey(new Date());
    const selectedKey = toDateKey(calSelectedDate);

    if (calMode === 'month') {
      const firstOfMonth = new Date(calViewYear, calViewMonth, 1);
      const startOffset = (firstOfMonth.getDay() + 6) % 7; // Monday = 0 .. Sunday = 6
      const startDate = new Date(calViewYear, calViewMonth, 1 - startOffset);

      const cellsHtml = [];
      for (let i = 0; i < 42; i++) {
        const d = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate() + i);
        const key = toDateKey(d);
        const isCurrentMonth = d.getMonth() === calViewMonth;
        const isWeekend = d.getDay() === 0 || d.getDay() === 6;
        const hasDot = dateHasAnyItem(d);

        const cls = [
          'cal-day-btn',
          isCurrentMonth ? '' : 'other-month',
          isWeekend ? 'weekend' : '',
          key === todayKey ? 'is-today' : '',
          key === selectedKey ? 'is-selected' : '',
        ]
          .filter(Boolean)
          .join(' ');

        cellsHtml.push(`
          <button type="button" class="${cls}" data-cal-date="${key}">
            <span>${d.getDate()}</span>
            ${hasDot ? '<span class="cal-dot" aria-hidden="true"></span>' : '<span style="height:6px;"></span>'}
          </button>
        `);
      }

      calGridContainer.innerHTML = `
        <div class="cal-weekdays">
          <div class="cal-weekday">Th 2</div>
          <div class="cal-weekday">Th 3</div>
          <div class="cal-weekday">Th 4</div>
          <div class="cal-weekday">Th 5</div>
          <div class="cal-weekday">Th 6</div>
          <div class="cal-weekday weekend">Th 7</div>
          <div class="cal-weekday weekend">CN</div>
        </div>
        <div class="cal-days-grid">${cellsHtml.join('')}</div>
      `;
    } else if (calMode === 'week') {
      const offset = (calSelectedDate.getDay() + 6) % 7;
      const monday = new Date(
        calSelectedDate.getFullYear(),
        calSelectedDate.getMonth(),
        calSelectedDate.getDate() - offset
      );
      const weekCells = [];
      for (let i = 0; i < 7; i++) {
        const d = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + i);
        const key = toDateKey(d);
        const isWeekend = d.getDay() === 0 || d.getDay() === 6;
        const hasDot = dateHasAnyItem(d);
        const cls = [
          'cal-day-btn',
          isWeekend ? 'weekend' : '',
          key === todayKey ? 'is-today' : '',
          key === selectedKey ? 'is-selected' : '',
        ]
          .filter(Boolean)
          .join(' ');
        weekCells.push(`
          <button type="button" class="${cls}" data-cal-date="${key}">
            <span>${d.getDate()}/${d.getMonth() + 1}</span>
            ${hasDot ? '<span class="cal-dot" aria-hidden="true"></span>' : '<span style="height:6px;"></span>'}
          </button>
        `);
      }
      calGridContainer.innerHTML = `
        <div class="cal-weekdays">
          <div class="cal-weekday">Th 2</div>
          <div class="cal-weekday">Th 3</div>
          <div class="cal-weekday">Th 4</div>
          <div class="cal-weekday">Th 5</div>
          <div class="cal-weekday">Th 6</div>
          <div class="cal-weekday weekend">Th 7</div>
          <div class="cal-weekday weekend">CN</div>
        </div>
        <div class="cal-days-grid">${weekCells.join('')}</div>
      `;
    } else {
      const thuLabel = formatThu(jsDayToHuitThu(calSelectedDate.getDay()));
      calGridContainer.innerHTML = `
        <div style="text-align:center;padding:1.5rem 0.5rem;">
          <div class="muted">${escapeHtml(thuLabel)}</div>
          <div style="font-size:2.4rem;font-weight:800;color:#dc2626;line-height:1.1;margin:0.3rem 0;">${calSelectedDate.getDate()}</div>
          <div style="font-weight:600;">Tháng ${calSelectedDate.getMonth() + 1}, ${calSelectedDate.getFullYear()}</div>
        </div>
      `;
    }

    // Chi tiết ngày đã chọn
    const thuStr = formatThu(jsDayToHuitThu(calSelectedDate.getDay()));
    calSelectedTitle.textContent = `${thuStr}, ngày ${calSelectedDate.getDate()}/${calSelectedDate.getMonth() + 1}/${calSelectedDate.getFullYear()}`;

    const customItems = getCustomEventsForDate(calSelectedDate);
    const classItems = getRegisteredClassesForDate(calSelectedDate);

    if (customItems.length === 0 && classItems.length === 0) {
      calSelectedItems.innerHTML = '<div class="muted small-note">Chưa có lịch học hoặc nhắc nhở nào trong ngày này.</div>';
      return;
    }

    const customHtml = customItems.map(
      (ev) => `
      <div class="cal-agenda-item custom">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:0.4rem;">
          <div>
            <strong>${escapeHtml(ev.title)}</strong>
            ${ev.time ? `<div class="sched-meta">Giờ: ${escapeHtml(ev.time)}</div>` : ''}
          </div>
          <button type="button" class="btn btn-sm btn-danger-outline" style="padding:0.15rem 0.45rem;font-size:0.74rem;" data-del-cal-ev="${escapeHtml(ev.id)}">Xóa</button>
        </div>
      </div>
    `
    );

    const classHtml = classItems.map(
      (c) => `
      <div class="cal-agenda-item ${c.kind}">
        <div><span class="badge badge-${c.kind}">${c.label}</span> <strong>${escapeHtml(c.tenHP)}</strong></div>
        <div class="sched-meta">${escapeHtml(c.clock || '')} • ${escapeHtml(c.phong || '—')}</div>
        ${c.gv ? `<div class="sched-meta">GV: ${escapeHtml(c.gv)}</div>` : ''}
      </div>
    `
    );

    calSelectedItems.innerHTML = [...customHtml, ...classHtml].join('');
  }

  if (calPrevBtn) {
    calPrevBtn.addEventListener('click', () => {
      if (calViewMonth === 0) {
        calViewMonth = 11;
        calViewYear--;
      } else {
        calViewMonth--;
      }
      renderCalendar();
    });
  }

  if (calNextBtn) {
    calNextBtn.addEventListener('click', () => {
      if (calViewMonth === 11) {
        calViewMonth = 0;
        calViewYear++;
      } else {
        calViewMonth++;
      }
      renderCalendar();
    });
  }

  const calStartBtn = document.getElementById('cal-start-btn');
  let calUserNavigated = false;

  if (calTodayBtn) {
    calTodayBtn.addEventListener('click', () => {
      calUserNavigated = true;
      const now = new Date();
      calSelectedDate = now;
      calViewYear = now.getFullYear();
      calViewMonth = now.getMonth();
      renderCalendar();
    });
  }

  if (calStartBtn) {
    calStartBtn.addEventListener('click', () => {
      const earliest = getEarliestRegisteredStartDate();
      if (!earliest) return;
      calUserNavigated = true;
      calSelectedDate = earliest;
      calViewYear = earliest.getFullYear();
      calViewMonth = earliest.getMonth();
      renderCalendar();
    });
  }

  for (const btn of calModeBtns) {
    btn.addEventListener('click', () => {
      calMode = btn.dataset.calMode;
      for (const b of calModeBtns) {
        b.setAttribute('aria-pressed', String(b === btn));
      }
      renderCalendar();
    });
  }

  if (calGridContainer) {
    calGridContainer.addEventListener('click', (e) => {
      const btn = e.target.closest('[data-cal-date]');
      if (!btn) return;
      const [y, m, d] = btn.dataset.calDate.split('-').map(Number);
      calSelectedDate = new Date(y, m - 1, d);
      calViewYear = calSelectedDate.getFullYear();
      calViewMonth = calSelectedDate.getMonth();
      renderCalendar();
    });
  }

  if (calEventForm) {
    calEventForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const title = calEventTitle.value.trim();
      const time = calEventTime.value.trim();
      if (!title) return;
      const key = toDateKey(calSelectedDate);
      if (!Array.isArray(calendarEvents[key])) calendarEvents[key] = [];
      calendarEvents[key].push({ id: `EV-${Date.now()}-${Math.floor(Math.random() * 1000)}`, title, time });
      calEventTitle.value = '';
      calEventTime.value = '';
      saveState();
      renderCalendar();
    });
  }

  if (calSelectedItems) {
    calSelectedItems.addEventListener('click', (e) => {
      const delBtn = e.target.closest('[data-del-cal-ev]');
      if (!delBtn) return;
      const key = toDateKey(calSelectedDate);
      if (!Array.isArray(calendarEvents[key])) return;
      calendarEvents[key] = calendarEvents[key].filter((ev) => ev.id !== delBtn.dataset.delCalEv);
      if (calendarEvents[key].length === 0) delete calendarEvents[key];
      saveState();
      renderCalendar();
    });
  }

  function switchView(viewMode) {
    for (const tab of viewTabs) {
      const isSelected = tab.dataset.view === viewMode;
      tab.setAttribute('aria-selected', String(isSelected));
      if (isSelected) {
        tab.scrollIntoView({ block: 'nearest', inline: 'nearest', container: 'nearest' });
      }
    }
    panelRegister.hidden = !(viewMode === 'register' || viewMode === 'both');
    panelTimetable.hidden = !(viewMode === 'timetable' || viewMode === 'both');
    if (panelCalendar) {
      const showCal = viewMode === 'calendar' || viewMode === 'both';
      panelCalendar.hidden = !showCal;
      if (showCal && !calUserNavigated) {
        const earliest = getEarliestRegisteredStartDate();
        if (earliest) {
          calSelectedDate = earliest;
          calViewYear = earliest.getFullYear();
          calViewMonth = earliest.getMonth();
          renderCalendar();
        }
      }
    }
    if (panelTuition) panelTuition.hidden = !(viewMode === 'tuition' || viewMode === 'both');
    if (panelExcel) panelExcel.hidden = !(viewMode === 'excel' || viewMode === 'both');
  }

  for (const tab of viewTabs) {
    tab.addEventListener('click', () => switchView(tab.dataset.view));
  }

  searchInput.addEventListener('input', () => {
    renderSuggestions();
    renderCatalog();
  });

  searchInput.addEventListener('focus', () => {
    if (searchInput.value.trim()) renderSuggestions();
  });

  searchInput.addEventListener('keydown', (e) => {
    if (searchSuggestionsEl.hidden || currentSuggestions.length === 0) return;
    const items = Array.from(searchSuggestionsEl.querySelectorAll('.suggestion-item'));
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      activeSuggestionIndex = (activeSuggestionIndex + 1) % currentSuggestions.length;
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      activeSuggestionIndex = (activeSuggestionIndex - 1 + currentSuggestions.length) % currentSuggestions.length;
    } else if (e.key === 'Enter' && activeSuggestionIndex >= 0) {
      e.preventDefault();
      selectSuggestion(activeSuggestionIndex);
      return;
    } else if (e.key === 'Escape') {
      hideSuggestions();
      return;
    } else {
      return;
    }

    items.forEach((el, i) => {
      const active = i === activeSuggestionIndex;
      el.classList.toggle('active', active);
      el.setAttribute('aria-selected', String(active));
    });
  });

  searchSuggestionsEl.addEventListener('mousedown', (e) => {
    const item = e.target.closest('[data-suggestion-idx]');
    if (item) {
      e.preventDefault();
      selectSuggestion(Number(item.dataset.suggestionIdx));
    }
  });

  document.addEventListener('click', (e) => {
    if (!e.target.closest('.search-autocomplete')) {
      hideSuggestions();
    }
  });
  filterSubject.addEventListener('change', () => {
    if (searchInput.value.trim() && filterSubject.value && searchInput.value.trim() !== filterSubject.value) {
      searchInput.value = '';
    }
    renderCatalog();
  });
  filterClass.addEventListener('change', renderCatalog);
  filterLecturer.addEventListener('change', renderCatalog);
  filterDay.addEventListener('change', renderCatalog);
  filterSession.addEventListener('change', renderCatalog);
  filterStatus.addEventListener('change', renderCatalog);

  btnResetFilters.addEventListener('click', () => {
    searchInput.value = '';
    filterSubject.value = '';
    filterClass.value = '';
    filterLecturer.value = '';
    filterDay.value = '';
    filterSession.value = '';
    filterStatus.value = 'all';
    renderCatalog();
  });

  document.getElementById('btn-clear-registered').addEventListener('click', () => {
    registered = [];
    saveState();
    renderAll();
  });

  document.getElementById('btn-clear-catalog').addEventListener('click', () => {
    catalog = [];
    registered = [];
    saveState();
    renderAll();
    showImportStatus('Đã xóa dữ liệu học phần. Bạn có thể tải lên file Excel mới.', 'info');
  });

  if (excelDropzone) {
    excelDropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      excelDropzone.classList.add('drag-over');
    });
    excelDropzone.addEventListener('dragleave', () => {
      excelDropzone.classList.remove('drag-over');
    });
    excelDropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      excelDropzone.classList.remove('drag-over');
      const file = e.dataTransfer?.files?.[0];
      if (file) importExcelFile(file);
    });
  }

  for (const btn of document.querySelectorAll('[data-export]')) {
    btn.addEventListener('click', () => {
      const fmt = btn.dataset.export;
      if (fmt === 'excel') exportRegisteredExcel();
      else if (fmt === 'word') exportRegisteredWord();
      else if (fmt === 'image') exportRegisteredImage();
    });
  }
  const hoursGuideDialog = document.getElementById('hours-guide-dialog');
  const hoursGuideBody = document.getElementById('hours-guide-body');

  const HUIT_EXAM_SLOTS = {
    1: { start: '07h00', end: '07h50' },
    2: { start: '07h50', end: '08h40' },
    3: { start: '08h40', end: '09h30' },
    4: { start: '09h30', end: '10h20' },
    5: { start: '10h20', end: '11h10' },
    6: { start: '11h10', end: '12h00' },
    7: { start: '13h00', end: '13h50' },
    8: { start: '13h50', end: '14h40' },
    9: { start: '14h40', end: '15h30' },
    10: { start: '15h30', end: '16h20' },
    11: { start: '16h20', end: '17h10' },
    12: { start: '17h10', end: '18h00' },
    13: { start: '18h00', end: '18h50' },
    14: { start: '18h50', end: '19h40' },
    15: { start: '19h40', end: '20h30' },
    16: { start: '20h30', end: '21h20' },
  };

  function buildHuitHoursTableHtml(title, slots, hasBreak, footerNote = '') {
    const rows = [];
    for (let i = 0; i < 6; i++) {
      const pSang = i + 1;
      const pChieu = i + 7;
      const sSlot = slots[pSang];
      const cSlot = slots[pChieu];

      if (hasBreak && i === 3) {
        const t16 = slots[16];
        rows.push(`
          <tr>
            <td colspan="4" style="text-align:center;background:#eff6ff;color:#1e3a8a;font-style:italic;font-weight:600;">
              Nghỉ giải lao 30 phút
            </td>
            <td><strong>16</strong></td>
            <td>Từ ${t16.start} đến ${t16.end}</td>
          </tr>
        `);
      }

      const pToi = hasBreak ? (i < 3 ? i + 13 : null) : i < 4 ? i + 13 : null;
      const tSlot = pToi ? slots[pToi] : null;

      rows.push(`
        <tr>
          <td><strong>${pSang}</strong></td>
          <td>Từ ${sSlot.start} đến ${sSlot.end}</td>
          <td><strong>${pChieu}</strong></td>
          <td>Từ ${cSlot.start} đến ${cSlot.end}</td>
          <td>${pToi ? `<strong>${pToi}</strong>` : ''}</td>
          <td>${tSlot ? `Từ ${tSlot.start} đến ${tSlot.end}` : ''}</td>
        </tr>
      `);
    }

    return `
      <div>
        <h3 style="margin:0 0 0.45rem;font-size:0.95rem;color:var(--huit-navy);">${escapeHtml(title)}</h3>
        <div class="table-wrap">
          <table class="guide-table" style="font-size:0.8rem;">
            <thead>
              <tr>
                <th colspan="2" style="text-align:center;">BUỔI SÁNG</th>
                <th colspan="2" style="text-align:center;">BUỔI CHIỀU</th>
                <th colspan="2" style="text-align:center;">BUỔI TỐI</th>
              </tr>
              <tr>
                <th>Tiết</th><th>Giờ</th>
                <th>Tiết</th><th>Giờ</th>
                <th>Tiết</th><th>Giờ</th>
              </tr>
            </thead>
            <tbody>${rows.join('')}</tbody>
          </table>
        </div>
        ${footerNote ? `<p class="muted small-note" style="margin-top:0.35rem;"><em>${escapeHtml(footerNote)}</em></p>` : ''}
      </div>
    `;
  }

  function renderHoursGuideTable() {
    if (!hoursGuideBody) return;
    hoursGuideBody.innerHTML = [
      buildHuitHoursTableHtml(
        '1. Giờ học đối với các học phần lý thuyết, thực hành học tại phòng lý thuyết, phòng Zoom:',
        HUIT_LT_SLOTS,
        true
      ),
      buildHuitHoursTableHtml(
        '2. Giờ học đối với các học phần học tại phòng máy tính, phòng thực hành, phòng thí nghiệm, thực hành GDQP-AN và GDTC:',
        HUIT_TH_SLOTS,
        false,
        'Các lớp học thí nghiệm/thực hành không nghỉ giải lao.'
      ),
      buildHuitHoursTableHtml('3. Giờ thi:', HUIT_EXAM_SLOTS, false),
    ].join('');
  }

  hoursGuideDialog.addEventListener('close', () => {
    hoursGuideDialog.returnValue = '';
  });

  document.getElementById('btn-excel-guide').addEventListener('click', () => excelGuideDialog.showModal());
  document.getElementById('btn-hours-guide').addEventListener('click', () => {
    renderHoursGuideTable();
    hoursGuideDialog.showModal();
  });

  document.getElementById('excel-file-input').addEventListener('change', (e) => {
    const file = e.target.files?.[0];
    if (file) {
      importExcelFile(file);
      e.target.value = '';
    }
  });

  // --- Account / Auth System (giữ dữ liệu theo tài khoản khi thoát trang) ---
  const accountBar = document.getElementById('account-bar');
  const authDialog = document.getElementById('auth-dialog');
  const authDialogTitle = document.getElementById('auth-dialog-title');
  const authForm = document.getElementById('auth-form');
  const authUsernameInput = document.getElementById('auth-username');
  const authFullnameWrap = document.getElementById('auth-fullname-wrap');
  const authFullnameInput = document.getElementById('auth-fullname');
  const authPasswordInput = document.getElementById('auth-password');
  const authErrorEl = document.getElementById('auth-error');
  const authSubmitBtn = document.getElementById('auth-submit-btn');
  const authModeBtns = Array.from(document.querySelectorAll('[data-auth-mode]'));
  let authMode = 'login';

  async function hashPassword(pw) {
    const raw = String(pw || '');
    if (window.crypto?.subtle) {
      const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(raw));
      return Array.from(new Uint8Array(buf))
        .map((b) => b.toString(16).padStart(2, '0'))
        .join('');
    }
    // Fallback nhẹ nếu chạy file:// trên trình duyệt chặn crypto.subtle
    let h = 2166136261;
    for (let i = 0; i < raw.length; i++) {
      h ^= raw.charCodeAt(i);
      h = Math.imul(h, 16777619);
    }
    return `fnv_${h >>> 0}`;
  }

  function getUsersMap() {
    return loadObjectState(STORAGE_KEYS.USERS, {});
  }

  function saveUsersMap(users) {
    try {
      localStorage.setItem(STORAGE_KEYS.USERS, JSON.stringify(users));
    } catch {
      // Ignore
    }
  }

  function setAuthMode(mode) {
    authMode = mode === 'signup' ? 'signup' : 'login';
    for (const btn of authModeBtns) {
      btn.setAttribute('aria-selected', String(btn.dataset.authMode === authMode));
    }
    const isSignup = authMode === 'signup';
    authFullnameWrap.hidden = !isSignup;
    authDialogTitle.textContent = isSignup ? 'Tạo tài khoản sinh viên mới' : 'Đăng nhập để lưu dữ liệu';
    authSubmitBtn.textContent = isSignup ? 'Tạo tài khoản & Đăng nhập' : 'Đăng nhập';
    authPasswordInput.setAttribute('autocomplete', isSignup ? 'new-password' : 'current-password');
    authErrorEl.hidden = true;
  }

  function renderAccountBar() {
    if (!accountBar) return;
    const users = getUsersMap();
    if (currentUser && users[currentUser]) {
      const u = users[currentUser];
      const displayName = u.fullName ? `${u.fullName} (${u.username})` : u.username;
      accountBar.innerHTML = `
        <div class="account-chip">
          <span><strong>${escapeHtml(displayName)}</strong></span>
          <button type="button" id="btn-logout-account" class="account-logout-btn">Đăng xuất</button>
        </div>
      `;
      document.getElementById('btn-logout-account').addEventListener('click', logoutAccount);
    } else {
      accountBar.innerHTML = `<button type="button" id="btn-open-auth" class="btn btn-sm account-btn">Đăng nhập</button>`;
      document.getElementById('btn-open-auth').addEventListener('click', () => {
        setAuthMode('login');
        authForm.reset();
        authDialog.showModal();
      });
    }
  }

  function applyAccountSession(username, migrateCurrentIfEmpty = false) {
    currentUser = username;
    try {
      localStorage.setItem(STORAGE_KEYS.CURRENT_USER, username);
    } catch {
      // Ignore
    }

    const saved = loadObjectState(getUserDataKey(username), null);
    if (saved && (Array.isArray(saved.catalog) || Array.isArray(saved.registered))) {
      const hasSavedContent = (saved.catalog?.length || 0) > 0 || (saved.registered?.length || 0) > 0;
      if (hasSavedContent || !migrateCurrentIfEmpty) {
        catalog = Array.isArray(saved.catalog) ? saved.catalog : [];
        registered = Array.isArray(saved.registered) ? saved.registered : [];
        customTuitionItems = Array.isArray(saved.customTuitionItems) ? saved.customTuitionItems : [];
        calendarEvents =
          saved.calendarEvents && typeof saved.calendarEvents === 'object' ? saved.calendarEvents : {};
        tuitionSelections =
          saved.tuitionSelections && typeof saved.tuitionSelections === 'object' ? saved.tuitionSelections : {};
        if (tuitionRateLtInput && saved.rateLt !== undefined) tuitionRateLtInput.value = saved.rateLt;
        if (tuitionRateInput && saved.rateTh !== undefined) tuitionRateInput.value = saved.rateTh;
        if (themeSelect && saved.theme) {
          themeSelect.value = saved.theme;
          applyTheme(saved.theme);
        }
      }
    }
    saveState();
    renderAccountBar();
    renderAll();
  }

  function logoutAccount() {
    saveState();
    currentUser = null;
    catalog = [];
    registered = [];
    customTuitionItems = [];
    calendarEvents = {};
    tuitionSelections = {};
    try {
      localStorage.removeItem(STORAGE_KEYS.CURRENT_USER);
      localStorage.removeItem(STORAGE_KEYS.CATALOG);
      localStorage.removeItem(STORAGE_KEYS.REGISTERED);
      localStorage.removeItem(STORAGE_KEYS.CUSTOM_TUITION);
      localStorage.removeItem(STORAGE_KEYS.CALENDAR_EVENTS);
      localStorage.removeItem(STORAGE_KEYS.TUITION_SELECTIONS);
    } catch {
      // Ignore
    }
    renderAccountBar();
    renderAll();
    showImportStatus('Đã đăng xuất. Dữ liệu tài khoản của bạn vẫn được lưu an toàn cho lần đăng nhập sau.', 'info');
  }

  for (const btn of authModeBtns) {
    btn.addEventListener('click', () => setAuthMode(btn.dataset.authMode));
  }

  if (authDialog) {
    authDialog.addEventListener('close', () => {
      authDialog.returnValue = '';
    });
  }

  if (authForm) {
    authForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      authErrorEl.hidden = true;

      const username = authUsernameInput.value.trim();
      const fullName = authFullnameInput.value.trim();
      const password = authPasswordInput.value;

      if (!username || password.length < 4) {
        authErrorEl.textContent = 'Vui lòng nhập tên đăng nhập và mật khẩu tối thiểu 4 ký tự.';
        authErrorEl.hidden = false;
        return;
      }

      const users = getUsersMap();
      const pwHash = await hashPassword(password);

      if (authMode === 'signup') {
        if (users[username]) {
          authErrorEl.textContent = 'Tên đăng nhập / MSSV này đã tồn tại. Vui lòng chọn Đăng nhập.';
          authErrorEl.hidden = false;
          return;
        }
        users[username] = { username, fullName, passwordHash: pwHash };
        saveUsersMap(users);
        applyAccountSession(username, true);
        authDialog.close();
        showImportStatus(`Đã tạo tài khoản "${username}" và lưu toàn bộ dữ liệu hiện tại.`, 'success');
      } else {
        const user = users[username];
        if (!user || user.passwordHash !== pwHash) {
          authErrorEl.textContent = 'Sai tên đăng nhập hoặc mật khẩu.';
          authErrorEl.hidden = false;
          return;
        }
        applyAccountSession(username, true);
        authDialog.close();
        showImportStatus(`Chào mừng ${user.fullName || username} đã quay trở lại! Dữ liệu đã được khôi phục.`, 'success');
      }
    });
  }

  // --- Sao lưu & Khôi phục dữ liệu (.json) để chuyển giữa Điện thoại & Máy tính ---
  function exportBackupJson() {
    const users = getUsersMap();
    const userRecord = currentUser && users[currentUser] ? users[currentUser] : null;
    const payload = {
      app: 'HUIT_DKHP',
      version: 1,
      exportedAt: new Date().toISOString(),
      account: userRecord,
      data: {
        catalog,
        registered,
        customTuitionItems,
        calendarEvents,
        tuitionSelections,
        rateLt: tuitionRateLtInput?.value || '750000',
        rateTh: tuitionRateInput?.value || '850000',
        theme: themeSelect?.value || 'system',
      },
    };

    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const fileSuffix = userRecord ? `_${userRecord.username}` : '';
    a.download = `HUIT_SaoLuu${fileSuffix}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    showImportStatus(`Đã xuất file sao lưu "${a.download}". Bạn có thể gửi file này sang điện thoại/máy tính khác để khôi phục.`, 'success');
  }

  function importBackupJson(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const parsed = JSON.parse(String(e.target.result || ''));
        if (!parsed || parsed.app !== 'HUIT_DKHP' || !parsed.data) {
          showImportStatus('File sao lưu không hợp lệ hoặc không đúng định dạng HUIT.', 'error');
          return;
        }

        const d = parsed.data;
        catalog = Array.isArray(d.catalog) ? d.catalog : [];
        registered = Array.isArray(d.registered) ? d.registered : [];
        customTuitionItems = Array.isArray(d.customTuitionItems) ? d.customTuitionItems : [];
        calendarEvents = d.calendarEvents && typeof d.calendarEvents === 'object' ? d.calendarEvents : {};
        tuitionSelections = d.tuitionSelections && typeof d.tuitionSelections === 'object' ? d.tuitionSelections : {};
        if (tuitionRateLtInput && d.rateLt !== undefined) tuitionRateLtInput.value = d.rateLt;
        if (tuitionRateInput && d.rateTh !== undefined) tuitionRateInput.value = d.rateTh;
        if (themeSelect && d.theme) {
          themeSelect.value = d.theme;
          applyTheme(d.theme);
        }

        if (parsed.account && parsed.account.username) {
          const users = getUsersMap();
          users[parsed.account.username] = parsed.account;
          saveUsersMap(users);
          currentUser = parsed.account.username;
          try {
            localStorage.setItem(STORAGE_KEYS.CURRENT_USER, currentUser);
          } catch {
            // Ignore
          }
        }

        saveState();
        renderAccountBar();
        renderAll();
        showImportStatus(
          `Đã khôi phục thành công ${catalog.length} lớp học phần và ${registered.length} môn đã đăng ký từ "${file.name}".`,
          'success'
        );
      } catch (err) {
        showImportStatus(`Lỗi đọc file sao lưu: ${err.message}`, 'error');
      }
    };
    reader.readAsText(file, 'utf-8');
  }

  const btnExportBackup = document.getElementById('btn-export-backup');
  const backupFileInput = document.getElementById('backup-file-input');
  if (btnExportBackup) btnExportBackup.addEventListener('click', exportBackupJson);
  if (backupFileInput) {
    backupFileInput.addEventListener('change', (e) => {
      const file = e.target.files?.[0];
      if (file) {
        importBackupJson(file);
        e.target.value = '';
      }
    });
  }

  // Khôi phục phiên đăng nhập nếu đã đăng nhập trước đó
  if (currentUser && getUsersMap()[currentUser]) {
    applyAccountSession(currentUser, false);
  } else {
    currentUser = null;
    renderAccountBar();
    renderAll();
  }
})();
