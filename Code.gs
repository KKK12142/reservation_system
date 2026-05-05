/**
 * 부여여자고등학교 2026 교육과정 박람회 - 실시간 상담 시스템 백엔드
 * 단일 Apps Script 프로젝트 안에서 학생/교사/전광판 3개 웹앱이 같은 doGet/doPost를 호출.
 *
 * 배포: Apps Script Editor → 배포 → 새 배포 → 웹앱
 *   - 실행 계정: 나 (스프레드시트 소유자)
 *   - 액세스: 모든 사용자 (익명 포함)
 *
 * 최초 1회 setupSheets() 함수를 Apps Script 에디터에서 직접 실행하여 시트/헤더/시드 데이터 생성.
 */

// ===== 상수 =====
const TZ = 'Asia/Seoul';
const SHEETS = {
  STUDENTS: 'students',
  BOOTHS: 'booths',
  RESERVATIONS: 'reservations',
  LOG: 'consultation_log',
  STATE: 'current_state',
  CONFIG: 'config',
  BLOCKS: 'blocks'
};
const ACTIVE_STATUS = ['waiting', 'calling', 'in_progress'];

// ===== 진입점 =====
function doGet(e) {
  const params = (e && e.parameter) || {};
  // 페이지 라우팅: ?page=student|teacher|display
  const page = params.page || (params.action ? null : 'student');
  if (page) {
    const allowed = ['student', 'teacher', 'display'];
    if (allowed.indexOf(page) === -1) {
      return HtmlService.createHtmlOutput('Not found: ' + page);
    }
    const t = HtmlService.createTemplateFromFile(page);
    t.appUrl = ScriptApp.getService().getUrl();
    t.boothId = params.booth_id || '';
    return t.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, user-scalable=no')
      .setTitle('부여여고 박람회 상담 시스템')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return handle(e);
}

function doPost(e) { return handle(e); }

// HtmlService 페이지에서 google.script.run으로 호출하는 통합 API 래퍼
function api(action, params) {
  return dispatch(action, params || {});
}

function handle(e) {
  try {
    const params = parseParams(e);
    const action = params.action || '';
    const data = dispatch(action, params);
    return jsonResponse({ ok: true, data });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err && err.message || err) });
  }
}

function parseParams(e) {
  const p = {};
  if (e && e.parameter) Object.assign(p, e.parameter);
  if (e && e.postData && e.postData.contents) {
    try { Object.assign(p, JSON.parse(e.postData.contents)); } catch (_) { /* not JSON */ }
  }
  return p;
}

function dispatch(action, params) {
  const handlers = {
    // 공용
    'ping':                () => ({ now: nowIso(), block: getCurrentBlock() }),
    'getBooths':           () => readSheet(SHEETS.BOOTHS).filter(b => isTrue(b.active)),
    'getCurrentBlock':     () => ({ block: getCurrentBlock() }),

    // 학생용
    'login':               () => handleLogin(params),
    'getMyQR':             () => handleGetMyQR(params),
    'getMyReservations':   () => handleGetMyReservations(params),
    'createReservation':   () => handleCreateReservation(params),
    'cancelReservation':   () => handleCancelReservation(params),
    'getMyQueueStatus':    () => handleGetMyQueueStatus(params),

    // 교사 태블릿(부스)용
    'boothLogin':          () => handleBoothLogin(params),
    'callNext':            () => handleCallNext(params),
    'scanQR':              () => handleScanQR(params),
    'completeConsultation':() => handleCompleteConsultation(params),
    'skipNoShow':          () => handleSkipNoShow(params),
    'getBoothQueue':       () => handleGetBoothQueue(params),

    // 전광판
    'getDashboard':        () => handleGetDashboard(),

    // 셋업/관리
    'setup':               () => setupSheets()
  };
  const fn = handlers[action];
  if (!fn) throw new Error('Unknown action: ' + action);
  return fn();
}

// ===== 응답/유틸 =====
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function nowIso() {
  return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function shortId(prefix) {
  return prefix + Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();
}

function isTrue(v) {
  return String(v).toUpperCase() === 'TRUE' || v === true;
}

// ===== 시트 헬퍼 =====
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getSheet(name) {
  const s = ss().getSheetByName(name);
  if (!s) throw new Error('Sheet not found: ' + name);
  return s;
}

function readSheet(name) {
  return readSheetWithRowIndex(name).rows;
}

function readSheetWithRowIndex(name) {
  const s = getSheet(name);
  const data = s.getDataRange().getValues();
  if (data.length < 2) return { headers: data[0] || [], rows: [] };
  const headers = data[0];
  const rows = data.slice(1).map((row, i) => {
    const obj = { _rowIndex: i + 2 };
    headers.forEach((h, idx) => obj[h] = row[idx]);
    return obj;
  });
  return { headers, rows };
}

function appendRow(sheetName, rowObj) {
  const s = getSheet(sheetName);
  const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const row = headers.map(h => rowObj[h] === undefined ? '' : rowObj[h]);
  s.appendRow(row);
}

function updateRowByIndex(sheetName, rowIndex, updates) {
  const s = getSheet(sheetName);
  const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  headers.forEach((h, i) => {
    if (updates[h] !== undefined) s.getRange(rowIndex, i + 1).setValue(updates[h]);
  });
}

// ===== 설정/시간 =====
function getConfig() {
  const rows = readSheet(SHEETS.CONFIG);
  const c = {};
  rows.forEach(r => c[r.key] = r.value);
  return c;
}

function getCurrentBlock() {
  const c = getConfig();
  const ov = c.block_override;
  if (ov !== '' && ov !== undefined && ov !== null) return Number(ov);
  return calcBlockByTime(new Date());
}

function calcBlockByTime(date) {
  const hhmm = Utilities.formatDate(date, TZ, 'HH:mm');
  const [h, m] = hhmm.split(':').map(Number);
  const t = h * 60 + m;
  if (t < 9 * 60 + 30)  return 0;   // ~09:30
  if (t < 10 * 60 + 30) return 1;   // 09:30~10:30
  if (t < 11 * 60 + 30) return 2;   // 10:30~11:30
  if (t < 13 * 60 + 30) return 3;   // 11:30~13:30 (점심 포함)
  if (t < 14 * 60 + 30) return 4;   // 13:30~14:30
  if (t < 15 * 60 + 30) return 5;   // 14:30~15:30
  if (t < 16 * 60 + 30) return 6;   // 15:30~16:30
  return 7;                          // 종료
}

// ===== QR 토큰 =====
function generateQrToken(studentId) {
  const c = getConfig();
  const salt = c.qr_salt || 'default-salt';
  const raw = String(studentId) + ':' + salt;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
  return digest.slice(0, 4)
    .map(b => ((b & 0xff) + 0x100).toString(16).slice(1)).join('');
}

// ===== 학생 핸들러 =====
function handleLogin(params) {
  const sid = String(params.student_id || '').trim();
  const name = String(params.name || '').trim();
  if (!sid || !name) throw new Error('학번과 이름을 모두 입력해 주세요');

  const sheet = readSheetWithRowIndex(SHEETS.STUDENTS);
  const row = sheet.rows.find(r => String(r.student_id) === sid && String(r.name).trim() === name);
  if (!row) throw new Error('학번 또는 이름이 일치하지 않습니다');

  // 토큰이 비어 있으면 즉시 채워줌(시트에 누락 시 안전장치)
  if (!row.qr_token) {
    const token = generateQrToken(row.student_id);
    updateRowByIndex(SHEETS.STUDENTS, row._rowIndex, { qr_token: token });
    row.qr_token = token;
  }
  return {
    student_id: String(row.student_id),
    name: row.name,
    grade: Number(row.grade),
    class_no: Number(row.class_no),
    number: Number(row.number),
    qr_token: row.qr_token
  };
}

function handleGetMyQR(params) {
  const sid = String(params.student_id || '');
  const stu = readSheet(SHEETS.STUDENTS).find(s => String(s.student_id) === sid);
  if (!stu) throw new Error('학생을 찾을 수 없습니다');
  return {
    student_id: String(stu.student_id),
    name: stu.name,
    qr_payload: 'STU:' + stu.student_id + ':' + stu.qr_token
  };
}

function handleGetMyReservations(params) {
  const sid = String(params.student_id || '');
  const all = readSheet(SHEETS.RESERVATIONS).filter(r => String(r.student_id) === sid);
  const booths = readSheet(SHEETS.BOOTHS);
  return all.map(r => {
    const b = booths.find(x => x.booth_id === r.booth_id);
    return {
      ...r,
      _rowIndex: undefined,
      booth_label: b ? `${b.space} · ${b.subject_group} · ${b.target_grade === 0 ? '전체' : b.target_grade + '학년'}` : r.booth_id,
      room: b ? b.room : ''
    };
  });
}

function handleCreateReservation(params) {
  const sid = String(params.student_id || '');
  const boothId = String(params.booth_id || '');
  if (!sid || !boothId) throw new Error('student_id, booth_id 필요');

  return withLock(() => {
    const c = getConfig();
    if (!isTrue(c.reservation_open)) throw new Error('현재 예약 접수가 중지되었습니다');

    const stu = readSheet(SHEETS.STUDENTS).find(s => String(s.student_id) === sid);
    if (!stu) throw new Error('학생을 찾을 수 없습니다');

    const booth = readSheet(SHEETS.BOOTHS).find(b => b.booth_id === boothId);
    if (!booth) throw new Error('부스를 찾을 수 없습니다');
    if (!isTrue(booth.active)) throw new Error('해당 부스는 운영하지 않습니다');

    const currentBlock = getCurrentBlock();
    const isPreReg = currentBlock === 0;

    const reservations = readSheet(SHEETS.RESERVATIONS);
    const myActive = reservations.filter(r =>
      String(r.student_id) === sid && ACTIVE_STATUS.indexOf(r.status) !== -1
    );

    // 규칙 1: 1교과군 1건
    if (myActive.some(r => r.subject_group === booth.subject_group)) {
      throw new Error(`이미 ${booth.subject_group} 교과군에 예약이 있습니다`);
    }

    // 규칙 2: 사전예약 기간이면 학생당 최대 N건
    if (isPreReg) {
      const max = Number(c.pre_reg_max_per_student || 2);
      if (myActive.length >= max) {
        throw new Error(`사전예약은 학생당 최대 ${max}건까지 가능합니다`);
      }
    }

    // 블록 결정: 사전예약은 학생이 선택, 당일 추가는 현재 블록
    let blockNum;
    if (isPreReg) {
      blockNum = Number(params.block);
      if (!(blockNum >= 1 && blockNum <= 6)) throw new Error('블록은 1~6 중 선택해야 합니다');
    } else {
      blockNum = currentBlock >= 7 ? 6 : currentBlock;
    }

    const reservation = {
      reservation_id: shortId('R'),
      student_id: sid,
      booth_id: boothId,
      subject_group: booth.subject_group,
      target_grade: booth.target_grade,
      block: blockNum,
      type: isPreReg ? 'reserved' : 'walkin',
      status: 'waiting',
      created_at: nowIso(),
      called_at: ''
    };
    appendRow(SHEETS.RESERVATIONS, reservation);
    return reservation;
  });
}

function handleCancelReservation(params) {
  const rid = String(params.reservation_id || '');
  const sid = String(params.student_id || '');
  return withLock(() => {
    const sheet = readSheetWithRowIndex(SHEETS.RESERVATIONS);
    const row = sheet.rows.find(r => r.reservation_id === rid);
    if (!row) throw new Error('예약을 찾을 수 없습니다');
    if (String(row.student_id) !== sid) throw new Error('본인 예약만 취소할 수 있습니다');
    if (['completed', 'in_progress'].indexOf(row.status) !== -1) {
      throw new Error('이미 진행 중이거나 완료된 예약은 취소할 수 없습니다');
    }
    updateRowByIndex(SHEETS.RESERVATIONS, row._rowIndex, { status: 'cancelled' });
    return { ok: true };
  });
}

function handleGetMyQueueStatus(params) {
  const sid = String(params.student_id || '');
  const reservations = readSheet(SHEETS.RESERVATIONS);
  const myActive = reservations.filter(r =>
    String(r.student_id) === sid && ACTIVE_STATUS.indexOf(r.status) !== -1
  );
  const booths = readSheet(SHEETS.BOOTHS);

  return myActive.map(r => {
    const booth = booths.find(b => b.booth_id === r.booth_id);
    if (r.status === 'in_progress') {
      return { reservation: r, booth, position: 0, ahead_count: 0, status_label: '상담 중' };
    }
    if (r.status === 'calling') {
      return { reservation: r, booth, position: 0, ahead_count: 0, status_label: '본인 호출됨 — 즉시 이동' };
    }
    const queue = computeQueue(r.booth_id);
    const idx = queue.findIndex(q => q.reservation_id === r.reservation_id);
    return {
      reservation: r,
      booth,
      position: idx === -1 ? null : idx + 1,
      ahead_count: idx === -1 ? null : idx,
      status_label: idx === 0 ? '곧 호출됩니다' : `앞에 ${idx}명 대기 중`
    };
  });
}

// 부스의 대기열을 정렬하여 반환 (waiting만)
function computeQueue(boothId) {
  return readSheet(SHEETS.RESERVATIONS)
    .filter(r => r.booth_id === boothId && r.status === 'waiting')
    .sort((a, b) => {
      if (a.block !== b.block) return Number(a.block) - Number(b.block);
      if (a.type !== b.type) return a.type === 'reserved' ? -1 : 1;
      return new Date(a.created_at) - new Date(b.created_at);
    });
}

// ===== 부스(태블릿) 핸들러 =====
function handleBoothLogin(params) {
  const boothId = String(params.booth_id || '');
  const booth = readSheet(SHEETS.BOOTHS).find(b => b.booth_id === boothId);
  if (!booth) throw new Error('부스를 찾을 수 없습니다');
  if (!isTrue(booth.active)) throw new Error('해당 부스는 비활성 상태입니다');
  return booth;
}

function handleCallNext(params) {
  const boothId = String(params.booth_id || '');
  return withLock(() => {
    // 이미 calling 상태인 학생이 있으면 그대로 반환(중복 호출 방지)
    const existingCalling = readSheet(SHEETS.RESERVATIONS).find(r =>
      r.booth_id === boothId && r.status === 'calling'
    );
    if (existingCalling) return enrichWithStudent(existingCalling);

    const queue = computeQueue(boothId);
    if (queue.length === 0) throw new Error('대기 중인 학생이 없습니다');

    const next = queue[0];
    const sheet = readSheetWithRowIndex(SHEETS.RESERVATIONS);
    const row = sheet.rows.find(r => r.reservation_id === next.reservation_id);
    updateRowByIndex(SHEETS.RESERVATIONS, row._rowIndex, {
      status: 'calling',
      called_at: nowIso()
    });
    return enrichWithStudent({ ...next, status: 'calling', called_at: nowIso() });
  });
}

function handleScanQR(params) {
  const boothId = String(params.booth_id || '');
  const payload = String(params.qr_payload || '');
  return withLock(() => {
    const parts = payload.split(':');
    if (parts.length !== 3 || parts[0] !== 'STU') throw new Error('QR 형식이 올바르지 않습니다');
    const [, scannedSid, scannedToken] = parts;

    const stu = readSheet(SHEETS.STUDENTS).find(s => String(s.student_id) === String(scannedSid));
    if (!stu) throw new Error('등록되지 않은 학생입니다');
    if (stu.qr_token !== scannedToken) throw new Error('QR 검증 실패 — 본인 QR이 맞는지 확인해 주세요');

    // 같은 부스의 calling/waiting 예약 중에서 매칭
    const sheet = readSheetWithRowIndex(SHEETS.RESERVATIONS);
    const candidates = sheet.rows.filter(r =>
      String(r.student_id) === String(scannedSid) &&
      r.booth_id === boothId &&
      ['calling', 'waiting'].indexOf(r.status) !== -1
    );
    if (candidates.length === 0) {
      throw new Error(`${stu.name} 학생은 본 부스에 예약이 없습니다 (부스: ${boothId})`);
    }

    // calling 상태가 우선, 없으면 waiting 중 첫 번째
    const target = candidates.find(r => r.status === 'calling') || candidates[0];
    const startedAt = nowIso();

    updateRowByIndex(SHEETS.RESERVATIONS, target._rowIndex, {
      status: 'in_progress',
      called_at: target.called_at || startedAt
    });

    appendRow(SHEETS.LOG, {
      log_id: shortId('L'),
      reservation_id: target.reservation_id,
      student_id: String(scannedSid),
      booth_id: boothId,
      subject_group: target.subject_group,
      started_at: startedAt,
      ended_at: '',
      duration_sec: '',
      type: target.type,
      memo: ''
    });

    upsertCurrentState(boothId, {
      current_reservation_id: target.reservation_id,
      current_student_id: String(scannedSid),
      current_started_at: startedAt,
      updated_at: startedAt
    });

    return {
      student: {
        student_id: String(stu.student_id),
        name: stu.name,
        grade: Number(stu.grade),
        class_no: Number(stu.class_no),
        number: Number(stu.number)
      },
      reservation_id: target.reservation_id,
      type: target.type,
      block: Number(target.block),
      started_at: startedAt
    };
  });
}

function handleCompleteConsultation(params) {
  const boothId = String(params.booth_id || '');
  const memo = String(params.memo || '');
  return withLock(() => {
    const stateSheet = readSheetWithRowIndex(SHEETS.STATE);
    const stateRow = stateSheet.rows.find(r => r.booth_id === boothId);
    if (!stateRow || !stateRow.current_reservation_id) {
      throw new Error('진행 중인 상담이 없습니다');
    }
    const endedAt = nowIso();
    const reservationId = stateRow.current_reservation_id;

    const resSheet = readSheetWithRowIndex(SHEETS.RESERVATIONS);
    const resRow = resSheet.rows.find(r => r.reservation_id === reservationId);
    if (resRow) {
      updateRowByIndex(SHEETS.RESERVATIONS, resRow._rowIndex, { status: 'completed' });
    }

    const logSheet = readSheetWithRowIndex(SHEETS.LOG);
    const logRow = [...logSheet.rows].reverse()
      .find(r => r.reservation_id === reservationId && !r.ended_at);
    if (logRow) {
      const startedAt = new Date(logRow.started_at);
      const durationSec = Math.max(0, Math.round((new Date(endedAt) - startedAt) / 1000));
      const updates = { ended_at: endedAt, duration_sec: durationSec };
      if (memo) updates.memo = memo;
      updateRowByIndex(SHEETS.LOG, logRow._rowIndex, updates);
    }

    updateRowByIndex(SHEETS.STATE, stateRow._rowIndex, {
      current_reservation_id: '',
      current_student_id: '',
      current_started_at: '',
      updated_at: endedAt
    });

    return { ok: true, ended_at: endedAt, reservation_id: reservationId };
  });
}

function handleSkipNoShow(params) {
  const boothId = String(params.booth_id || '');
  return withLock(() => {
    const sheet = readSheetWithRowIndex(SHEETS.RESERVATIONS);
    const callingRow = sheet.rows.find(r => r.booth_id === boothId && r.status === 'calling');
    if (!callingRow) throw new Error('호출 중인 학생이 없습니다');
    // 대기열 끝으로: status를 waiting으로 되돌리고 created_at을 현재로 갱신
    updateRowByIndex(SHEETS.RESERVATIONS, callingRow._rowIndex, {
      status: 'waiting',
      called_at: '',
      created_at: nowIso()
    });
    return { ok: true, reservation_id: callingRow.reservation_id };
  });
}

function handleGetBoothQueue(params) {
  const boothId = String(params.booth_id || '');
  const queue = computeQueue(boothId);
  const enrichedQueue = queue.map(enrichWithStudent);

  const states = readSheet(SHEETS.STATE);
  const state = states.find(s => s.booth_id === boothId);
  let current = null;
  if (state && state.current_reservation_id) {
    const r = readSheet(SHEETS.RESERVATIONS).find(x => x.reservation_id === state.current_reservation_id);
    if (r) current = { ...enrichWithStudent(r), started_at: state.current_started_at };
  }

  const calling = readSheet(SHEETS.RESERVATIONS).find(r =>
    r.booth_id === boothId && r.status === 'calling'
  );

  return {
    booth_id: boothId,
    current,
    calling: calling ? enrichWithStudent(calling) : null,
    queue: enrichedQueue,
    queue_count: enrichedQueue.length
  };
}

function enrichWithStudent(reservation) {
  const stu = readSheet(SHEETS.STUDENTS).find(s => String(s.student_id) === String(reservation.student_id));
  if (!stu) return { ...reservation, _rowIndex: undefined };
  return {
    ...reservation,
    _rowIndex: undefined,
    student_name: stu.name,
    student_grade: Number(stu.grade),
    student_class_no: Number(stu.class_no),
    student_number: Number(stu.number),
    student_label: `${stu.grade}-${stu.class_no}-${stu.number} ${stu.name}`
  };
}

function upsertCurrentState(boothId, updates) {
  const sheet = readSheetWithRowIndex(SHEETS.STATE);
  const row = sheet.rows.find(r => r.booth_id === boothId);
  if (row) {
    updateRowByIndex(SHEETS.STATE, row._rowIndex, updates);
  } else {
    appendRow(SHEETS.STATE, Object.assign({ booth_id: boothId }, updates));
  }
}

// ===== 전광판 =====
function handleGetDashboard() {
  const booths = readSheet(SHEETS.BOOTHS).filter(b => isTrue(b.active));
  const reservations = readSheet(SHEETS.RESERVATIONS);
  const students = readSheet(SHEETS.STUDENTS);
  const states = readSheet(SHEETS.STATE);

  const stuMap = {};
  students.forEach(s => stuMap[String(s.student_id)] = s);

  const result = booths.map(b => {
    const state = states.find(s => s.booth_id === b.booth_id) || {};
    const waiting = reservations.filter(r => r.booth_id === b.booth_id && r.status === 'waiting');
    const calling = reservations.find(r => r.booth_id === b.booth_id && r.status === 'calling');

    const sortedQueue = waiting.sort((a, b) => {
      if (a.block !== b.block) return Number(a.block) - Number(b.block);
      if (a.type !== b.type) return a.type === 'reserved' ? -1 : 1;
      return new Date(a.created_at) - new Date(b.created_at);
    });

    return {
      booth_id: b.booth_id,
      space: b.space,
      subject_group: b.subject_group,
      target_grade: Number(b.target_grade),
      room: b.room,
      current: state.current_student_id ? maskStudent(stuMap[String(state.current_student_id)]) : null,
      calling: calling ? maskStudent(stuMap[String(calling.student_id)]) : null,
      waiting_count: waiting.length,
      next: sortedQueue.slice(0, 3).map(r => ({
        ...maskStudent(stuMap[String(r.student_id)]),
        type: r.type,
        block: Number(r.block)
      }))
    };
  });

  return {
    current_block: getCurrentBlock(),
    updated_at: nowIso(),
    booths: result
  };
}

function maskStudent(s) {
  if (!s) return null;
  return {
    name: maskName(s.name),
    grade_class: `${s.grade}-${s.class_no}`,
    number: Number(s.number)
  };
}

function maskName(name) {
  if (!name || String(name).length <= 1) return String(name || '');
  const n = String(name);
  if (n.length === 2) return n[0] + '*';
  return n[0] + '*' + n.slice(2);
}

// ===== 셋업 (최초 1회 Apps Script 에디터에서 직접 실행) =====
function setupSheets() {
  const schemas = {
    students: ['student_id', 'name', 'grade', 'class_no', 'number', 'qr_token'],
    booths: ['booth_id', 'subject_group', 'target_grade', 'space', 'room', 'teacher_am', 'teacher_pm', 'active'],
    reservations: ['reservation_id', 'student_id', 'booth_id', 'subject_group', 'target_grade', 'block', 'type', 'status', 'created_at', 'called_at'],
    consultation_log: ['log_id', 'reservation_id', 'student_id', 'booth_id', 'subject_group', 'started_at', 'ended_at', 'duration_sec', 'type', 'memo'],
    current_state: ['booth_id', 'current_reservation_id', 'current_student_id', 'current_started_at', 'updated_at'],
    config: ['key', 'value'],
    blocks: ['block', 'start', 'end', 'label']
  };

  const spreadsheet = ss();
  Object.keys(schemas).forEach(name => {
    const headers = schemas[name];
    let s = spreadsheet.getSheetByName(name);
    if (!s) s = spreadsheet.insertSheet(name);
    s.getRange(1, 1, 1, headers.length).setValues([headers]);
    s.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    s.setFrozenRows(1);
  });

  // config 시드
  const configSheet = spreadsheet.getSheetByName('config');
  if (configSheet.getLastRow() < 2) {
    const seed = [
      ['event_date', '2026-05-14'],
      ['current_block', 0],
      ['block_override', ''],
      ['reservation_open', 'TRUE'],
      ['no_show_minutes', 3],
      ['pre_reg_max_per_student', 2],
      ['qr_salt', Utilities.getUuid()]
    ];
    configSheet.getRange(2, 1, seed.length, 2).setValues(seed);
  }

  // blocks 시드
  const blocksSheet = spreadsheet.getSheetByName('blocks');
  if (blocksSheet.getLastRow() < 2) {
    const seed = [
      [1, '09:40', '10:30', '1블록'],
      [2, '10:40', '11:30', '2블록'],
      [3, '11:40', '12:30', '3블록'],
      [4, '13:30', '14:20', '4블록'],
      [5, '14:30', '15:20', '5블록'],
      [6, '15:30', '16:20', '6블록']
    ];
    blocksSheet.getRange(2, 1, seed.length, 4).setValues(seed);
  }

  return { ok: true, sheets: Object.keys(schemas) };
}

/**
 * 학생 명단 일괄 업로드 시 호출 (선택) — students 시트에 학번/이름/학년/반/번호 입력 후
 * 이 함수를 한 번 실행하면 qr_token이 비어있는 행을 모두 자동 채워줌.
 */
function backfillQrTokens() {
  const sheet = readSheetWithRowIndex(SHEETS.STUDENTS);
  let count = 0;
  sheet.rows.forEach(r => {
    if (!r.qr_token && r.student_id) {
      const token = generateQrToken(r.student_id);
      updateRowByIndex(SHEETS.STUDENTS, r._rowIndex, { qr_token: token });
      count++;
    }
  });
  return { backfilled: count };
}
