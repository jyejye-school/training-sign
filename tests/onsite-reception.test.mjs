import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..');
const backend = fs.readFileSync(path.join(root, 'apps-script', 'Code.gs'), 'utf8');
const index = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const app = fs.readFileSync(path.join(root, 'assets', 'app.js'), 'utf8');
const styles = fs.readFileSync(path.join(root, 'assets', 'styles.css'), 'utf8');

const loadHarness = new Function('PropertiesService', 'CacheService', 'Utilities', 'LockService', 'Date', `${backend}
return {
  normalizeTraining_: normalizeTraining_,
  publicTraining_: publicTraining_,
  validateSigningWindow_: validateSigningWindow_,
  validateTrainingReception_: validateTrainingReception_,
  getTrainingReception_: getTrainingReception_,
  startTrainingReception_: startTrainingReception_,
  closeTrainingReception_: closeTrainingReception_,
  readTrainingReceptionState_: readTrainingReceptionState_,
  cleanupExpiredTrainingReceptions_: cleanupExpiredTrainingReceptions_,
  clearCopiedInstanceProperties_: clearCopiedInstanceProperties_,
  trainingReceptionSchedule_: trainingReceptionSchedule_,
  saveTraining_: saveTraining_,
  dispatch_: dispatch_,
  propertyKey: trainingReceptionPropertyKey_,
  configureTraining: function(training) {
    findRow_ = function(definition, key, value) {
      if (definition === SHEETS.TRAININGS && key === 'id' && value === training.id) return training;
      return null;
    };
    audit_ = function() {};
  },
  setAuditHandler: function(handler) {
    audit_ = handler;
  },
  configureTrainingSave: function(existing) {
    readSettings_ = function() {
      return {
        schoolName: '테스트학교', subtitle: '연수 참여 확인',
        privacyPurpose: '연수 참여 확인', privacyItems: '부서, 성명, 서명', privacyRetention: '출력 뒤 삭제'
      };
    };
    readRowsWithNumbers_ = function(definition) {
      return definition === SHEETS.TRAININGS ? [{ data: existing, rowNumber: 2 }] : [];
    };
    sheet_ = function() { return {}; };
    writeObjectRow_ = function() {};
    audit_ = function() {};
  },
  configureAdminSession: function(validToken) {
    requireAdminSession_ = function(token) {
      if (token !== validToken) apiError_('SESSION_EXPIRED', '관리자 로그인이 필요합니다.');
      return token;
    };
    withAdminMutationLock_ = function(callback) { return callback(); };
  }
};`);

function formatDate(date, timeZone, pattern) {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: timeZone || 'Asia/Seoul',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hourCycle: 'h23'
  }).formatToParts(date);
  const value = type => parts.find(part => part.type === type)?.value || '';
  if (pattern === 'yyyy-MM-dd') return `${value('year')}-${value('month')}-${value('day')}`;
  if (pattern === 'HH:mm') return `${value('hour')}:${value('minute')}`;
  if (pattern === 'HH:mm:ss') return `${value('hour')}:${value('minute')}:${value('second')}`;
  return '';
}

function createHarness(initialProperties = {}, fixedNow = Date.now()) {
  const propertyValues = new Map(Object.entries({
    SPREADSHEET_ID: 'spreadsheet-0001',
    INSTANCE_ID: 'instance-0001',
    ONSITE_CODE_SECRET: 'test-only-onsite-secret',
    ...initialProperties
  }));
  const propertyStore = {
    getProperty: key => propertyValues.get(key) ?? null,
    setProperty(key, value) { propertyValues.set(key, String(value)); return this; },
    setProperties(values) { Object.entries(values).forEach(([key, value]) => propertyValues.set(key, String(value))); return this; },
    deleteProperty(key) { propertyValues.delete(key); return this; },
    getProperties: () => Object.fromEntries(propertyValues)
  };
  const cacheValues = new Map();
  const cache = {
    get: key => cacheValues.get(key) ?? null,
    put(key, value) { cacheValues.set(key, String(value)); },
    remove(key) { cacheValues.delete(key); }
  };
  const bytes = value => [...Buffer.from(value)].map(byte => byte > 127 ? byte - 256 : byte);
  const utilities = {
    Charset: { UTF_8: 'UTF-8' },
    DigestAlgorithm: { SHA_256: 'SHA-256' },
    getUuid: () => crypto.randomUUID(),
    computeDigest: (algorithm, value) => bytes(crypto.createHash('sha256').update(String(value)).digest()),
    computeHmacSha256Signature: (value, key) => bytes(crypto.createHmac('sha256', String(key)).update(String(value)).digest()),
    base64EncodeWebSafe: value => Buffer.from(value.map(byte => byte & 255)).toString('base64url'),
    formatDate
  };
  const lock = { waitLock() {}, releaseLock() {} };
  const nowMs = new Date(fixedNow).getTime();
  class ControlledDate extends Date {
    constructor(...args) {
      super(...(args.length ? args : [nowMs]));
    }
    static now() { return nowMs; }
  }
  const api = loadHarness(
    { getScriptProperties: () => propertyStore },
    { getScriptCache: () => cache },
    utilities,
    { getScriptLock: () => lock },
    ControlledDate
  );
  return { api, propertyStore, propertyValues, cacheValues };
}

const person = { id: 'staff-0001', active: true, department: '교무부', name: '홍길동' };

function dailyTraining(overrides = {}) {
  return {
    id: 'training-0001',
    title: '현장 연수',
    daily: true,
    date: '',
    active: true,
    startTime: '',
    endTime: '',
    audienceMode: 'all',
    audienceDepartments: '',
    onsiteCodeRequired: true,
    ...overrides
  };
}

test('기존 연수는 현장 코드가 기본 꺼짐이고 7월 13일 연수도 그대로 서명할 수 있다', () => {
  const { api } = createHarness();
  const normalized = api.normalizeTraining_({
    id: 'training-0713',
    title: '7월 13일 기존 연수',
    daily: false,
    date: '2026-07-13',
    active: true
  });
  assert.equal(Object.hasOwn(normalized, 'onsiteCodeRequired'), false);
  assert.equal(api.publicTraining_({ ...normalized, sortOrder: 1 }).onsiteCodeRequired, false);
  assert.deepEqual(
    api.validateSigningWindow_(normalized, person),
    { verificationMethod: 'share_link', onsiteSessionId: '' }
  );

  const enabled = api.normalizeTraining_({ ...normalized, onsiteCodeRequired: true });
  const disabled = api.normalizeTraining_({ ...normalized, onsiteCodeRequired: false });
  assert.equal(enabled.onsiteCodeRequired, true);
  assert.equal(disabled.onsiteCodeRequired, false);
  assert.throws(
    () => api.normalizeTraining_({ ...normalized, onsiteCodeRequired: 'true' }),
    error => error.apiCode === 'VALIDATION'
  );
});

test('접수 시작은 6자리 코드만 한 번 반환하고 저장값·상태 조회에는 원문 코드가 없다', () => {
  const harness = createHarness();
  const training = dailyTraining();
  harness.api.configureTraining(training);

  const started = harness.api.startTrainingReception_(training.id, 5);
  assert.match(started.code, /^\d{6}$/);
  assert.equal(started.required, true);
  assert.equal(started.status, 'open');
  assert.equal(started.durationMinutes, 5);

  const stored = JSON.parse(harness.propertyStore.getProperty(harness.api.propertyKey(training.id)));
  assert.equal(Object.hasOwn(stored, 'code'), false);
  assert.match(stored.codeHash, /^[A-Za-z0-9_-]{20,}$/);
  assert.notEqual(stored.codeHash, started.code);

  const status = harness.api.getTrainingReception_(training.id);
  assert.equal(status.status, 'open');
  assert.equal(status.sessionId, started.sessionId);
  assert.equal(Object.hasOwn(status, 'code'), false);
  assert.equal(Object.hasOwn(status, 'codeHash'), false);
  assert.equal(Object.hasOwn(status, 'salt'), false);
});

test('관리자가 선택하지 않은 연수는 코드 없이 동작하고 접수 시작은 거부된다', () => {
  const harness = createHarness();
  const training = dailyTraining({ onsiteCodeRequired: false });
  harness.api.configureTraining(training);

  assert.deepEqual(
    harness.api.validateTrainingReception_(training, person, '', new Date(), true),
    { verificationMethod: 'share_link', onsiteSessionId: '' }
  );
  assert.equal(harness.api.getTrainingReception_(training.id).required, false);
  assert.throws(
    () => harness.api.startTrainingReception_(training.id, 5),
    error => error.apiCode === 'RECEPTION_NOT_ENABLED'
  );
  assert.throws(
    () => {
      training.onsiteCodeRequired = true;
      harness.api.startTrainingReception_(training.id, 7);
    },
    error => error.apiCode === 'RECEPTION_DURATION'
  );
});

test('접수 코드는 미입력·오입력·5회 제한·교체·종료 상태를 서버에서 검증한다', () => {
  const harness = createHarness();
  const training = dailyTraining();
  harness.api.configureTraining(training);
  const first = harness.api.startTrainingReception_(training.id, 10);

  assert.throws(
    () => harness.api.validateTrainingReception_(training, person, '', new Date(), true),
    error => error.apiCode === 'ONSITE_CODE_REQUIRED'
  );
  assert.throws(
    () => harness.api.validateTrainingReception_(training, person, '12A456', new Date(), true),
    error => error.apiCode === 'ONSITE_CODE_INVALID'
  );
  assert.deepEqual(
    harness.api.validateTrainingReception_(training, person, first.code, new Date(), true),
    { verificationMethod: 'onsite_code', onsiteSessionId: first.sessionId }
  );

  const second = harness.api.startTrainingReception_(training.id, 5);
  assert.notEqual(second.sessionId, first.sessionId);
  assert.throws(
    () => harness.api.validateTrainingReception_(training, person, first.code, new Date(), true),
    error => error.apiCode === 'ONSITE_CODE_INVALID'
  );
  assert.deepEqual(
    harness.api.validateTrainingReception_(training, person, second.code, new Date(), true),
    { verificationMethod: 'onsite_code', onsiteSessionId: second.sessionId }
  );

  for (let attempt = 1; attempt <= 4; attempt += 1) {
    assert.throws(
      () => harness.api.validateTrainingReception_(training, person, '000000', new Date(), true),
      error => error.apiCode === 'ONSITE_CODE_INVALID'
    );
  }
  assert.throws(
    () => harness.api.validateTrainingReception_(training, person, '000000', new Date(), true),
    error => error.apiCode === 'ONSITE_CODE_LOCKED'
  );
  assert.throws(
    () => harness.api.validateTrainingReception_(training, person, second.code, new Date(), true),
    error => error.apiCode === 'ONSITE_CODE_LOCKED'
  );

  const third = harness.api.startTrainingReception_(training.id, 5);
  assert.deepEqual(
    harness.api.validateTrainingReception_(training, person, third.code, new Date(), true),
    { verificationMethod: 'onsite_code', onsiteSessionId: third.sessionId }
  );
  harness.api.closeTrainingReception_(training.id);
  assert.throws(
    () => harness.api.validateTrainingReception_(training, person, third.code, new Date(), true),
    error => error.apiCode === 'TRAINING_RECEPTION_CLOSED'
  );
});

test('오답 횟수는 Script Properties에 구성원별·세션 전체로 저장되고 새 코드에서 초기화된다', () => {
  const harness = createHarness();
  const training = dailyTraining();
  harness.api.configureTraining(training);
  const first = harness.api.startTrainingReception_(training.id, 15);
  const propertyKey = harness.api.propertyKey(training.id);

  assert.throws(
    () => harness.api.validateTrainingReception_(training, person, '000000', new Date(), true),
    error => error.apiCode === 'ONSITE_CODE_INVALID'
  );
  let stored = JSON.parse(harness.propertyStore.getProperty(propertyKey));
  assert.equal(stored.failureTotal, 1);
  assert.equal(Object.values(stored.failureCounts).reduce((sum, count) => sum + count, 0), 1);
  assert.equal(harness.cacheValues.size, 0, '오답 횟수를 CacheService에만 저장하면 인스턴스 간 제한이 풀릴 수 있습니다.');

  assert.deepEqual(
    harness.api.validateTrainingReception_(training, person, first.code, new Date(), true),
    { verificationMethod: 'onsite_code', onsiteSessionId: first.sessionId }
  );
  stored = JSON.parse(harness.propertyStore.getProperty(propertyKey));
  assert.equal(stored.failureTotal, 1, '세션 전체 오답 횟수는 성공 뒤에도 유지해야 합니다.');
  assert.deepEqual(stored.failureCounts, {}, '성공한 구성원의 개인 오답 횟수만 지워야 합니다.');

  for (let index = 0; index < 49; index += 1) {
    const anotherPerson = { ...person, id: `staff-${String(index).padStart(4, '0')}` };
    const expectedCode = index === 48 ? 'ONSITE_CODE_LOCKED' : 'ONSITE_CODE_INVALID';
    assert.throws(
      () => harness.api.validateTrainingReception_(training, anotherPerson, '000000', new Date(), true),
      error => error.apiCode === expectedCode
    );
  }
  stored = JSON.parse(harness.propertyStore.getProperty(propertyKey));
  assert.equal(stored.failureTotal, 50);
  assert.throws(
    () => harness.api.validateTrainingReception_(training, { ...person, id: 'staff-final1' }, first.code, new Date(), true),
    error => error.apiCode === 'ONSITE_CODE_LOCKED'
  );

  const replacement = harness.api.startTrainingReception_(training.id, 5);
  stored = JSON.parse(harness.propertyStore.getProperty(propertyKey));
  assert.equal(stored.failureTotal, 0);
  assert.deepEqual(stored.failureCounts, {});
  assert.deepEqual(
    harness.api.validateTrainingReception_(training, { ...person, id: 'staff-final1' }, replacement.code, new Date(), true),
    { verificationMethod: 'onsite_code', onsiteSessionId: replacement.sessionId }
  );
});

test('오늘·매일 연수 접수는 운영 시각을 따르고 만료 시각이 연수 종료를 넘지 않는다', () => {
  const fixedNow = '2026-07-21T01:30:30.000Z'; // 서울 10:30:30

  const beforeHarness = createHarness({}, fixedNow);
  const beforeTraining = dailyTraining({ id: 'training-early', startTime: '10:31', endTime: '11:00' });
  beforeHarness.api.configureTraining(beforeTraining);
  assert.throws(
    () => beforeHarness.api.startTrainingReception_(beforeTraining.id, 5),
    error => error.apiCode === 'TOO_EARLY'
  );

  const afterHarness = createHarness({}, fixedNow);
  const afterTraining = dailyTraining({
    id: 'training-ended', daily: false, date: '2026-07-21', startTime: '09:00', endTime: '10:29'
  });
  afterHarness.api.configureTraining(afterTraining);
  assert.throws(
    () => afterHarness.api.startTrainingReception_(afterTraining.id, 5),
    error => error.apiCode === 'TOO_LATE'
  );

  const cappedHarness = createHarness({}, fixedNow);
  const cappedTraining = dailyTraining({
    id: 'training-capped', daily: false, date: '2026-07-21', startTime: '10:00', endTime: '10:31'
  });
  cappedHarness.api.configureTraining(cappedTraining);
  const capped = cappedHarness.api.startTrainingReception_(cappedTraining.id, 15);
  assert.equal(capped.durationMinutes, 15);
  assert.ok(Date.parse(capped.expiresAt) <= Date.parse('2026-07-21T01:32:00.000Z'));
  assert.ok(capped.remainingSeconds <= 90 && capped.remainingSeconds > 0);

  const pastHarness = createHarness({}, fixedNow);
  const pastTraining = dailyTraining({
    id: 'training-past1', daily: false, date: '2026-07-13', startTime: '09:00', endTime: '10:00'
  });
  pastHarness.api.configureTraining(pastTraining);
  const past = pastHarness.api.startTrainingReception_(pastTraining.id, 15);
  assert.equal(past.status, 'open');
  assert.ok(past.remainingSeconds >= 899, '과거 고정 연수는 당시 운영 시각 때문에 재수합이 막히면 안 됩니다.');
});

test('연수 운영 시각을 바꾸면 진행 중인 현장 접수가 자동 종료된다', () => {
  const harness = createHarness({}, '2026-07-21T01:30:30.000Z');
  const training = dailyTraining({ startTime: '09:00', endTime: '12:00' });
  harness.api.configureTraining(training);
  harness.api.startTrainingReception_(training.id, 15);
  const propertyKey = harness.api.propertyKey(training.id);
  assert.ok(harness.propertyStore.getProperty(propertyKey));

  harness.api.configureTrainingSave(training);
  harness.api.saveTraining_({
    id: training.id,
    title: training.title,
    daily: true,
    date: '',
    startTime: '09:00',
    endTime: '11:30',
    active: true,
    audienceMode: 'all',
    audienceDepartments: [],
    onsiteCodeRequired: true
  });
  assert.equal(harness.propertyStore.getProperty(propertyKey), null);
});

test('접수 시작 감사 기록 실패는 이전 상태를 복구하고 종료 감사 실패는 닫힘을 유지한다', () => {
  const harness = createHarness();
  const training = dailyTraining();
  harness.api.configureTraining(training);
  harness.api.startTrainingReception_(training.id, 10);
  const propertyKey = harness.api.propertyKey(training.id);
  const previousRaw = harness.propertyStore.getProperty(propertyKey);

  harness.api.setAuditHandler(action => {
    if (action === 'start_training_reception' || action === 'close_training_reception') {
      throw new Error('감사 기록 저장 실패');
    }
  });
  assert.throws(() => harness.api.startTrainingReception_(training.id, 5), /감사 기록 저장 실패/);
  assert.equal(harness.propertyStore.getProperty(propertyKey), previousRaw);

  const originalConsoleError = console.error;
  console.error = () => {};
  try {
    const closed = harness.api.closeTrainingReception_(training.id);
    assert.equal(closed.status, 'closed');
  } finally {
    console.error = originalConsoleError;
  }
  assert.equal(harness.propertyStore.getProperty(propertyKey), null);
});

test('만료된 접수와 복사된 인스턴스의 접수 상태는 정리된다', () => {
  const harness = createHarness();
  const training = dailyTraining();
  harness.api.configureTraining(training);
  harness.api.startTrainingReception_(training.id, 5);
  const key = harness.api.propertyKey(training.id);
  const expired = JSON.parse(harness.propertyStore.getProperty(key));
  expired.expiresAt = new Date(Date.now() - 1000).toISOString();
  harness.propertyStore.setProperty(key, JSON.stringify(expired));

  assert.equal(harness.api.getTrainingReception_(training.id).status, 'expired');
  assert.throws(
    () => harness.api.validateTrainingReception_(training, person, '123456', new Date(), true),
    error => error.apiCode === 'TRAINING_RECEPTION_CLOSED'
  );
  assert.equal(harness.api.cleanupExpiredTrainingReceptions_(), 1);
  assert.equal(harness.propertyStore.getProperty(key), null);

  harness.propertyStore.setProperty('SHARE_TOKEN', 'old-share-token');
  harness.propertyStore.setProperty(key, JSON.stringify(expired));
  harness.api.clearCopiedInstanceProperties_(harness.propertyStore);
  assert.equal(harness.propertyStore.getProperty('SHARE_TOKEN'), null);
  assert.equal(harness.propertyStore.getProperty('ONSITE_CODE_SECRET'), null);
  assert.equal(harness.propertyStore.getProperty(key), null);
});

test('현장 접수 관리 API는 관리자 세션 뒤에서만 호출된다', () => {
  const harness = createHarness();
  const training = dailyTraining();
  harness.api.configureTraining(training);
  harness.api.configureAdminSession('valid-admin-session');

  assert.throws(
    () => harness.api.dispatch_({ action: 'get_training_reception_status', trainingId: training.id }),
    error => error.apiCode === 'SESSION_EXPIRED'
  );
  const result = harness.api.dispatch_({
    action: 'get_training_reception_status',
    sessionToken: 'valid-admin-session',
    trainingId: training.id
  });
  assert.equal(result.required, true);
  assert.equal(result.status, 'closed');
  assert.equal(Object.hasOwn(result, 'code'), false);
});

test('서명 제출은 이미지 생성 전과 잠금 후에 현장 코드를 다시 검증한다', () => {
  const body = backend.slice(backend.indexOf('function submitSignature_'), backend.indexOf('function validateSigningWindow_'));
  const firstValidation = body.indexOf('validateSigningWindow_(training, person, onsiteCode, true)');
  const createFile = body.indexOf('.createFile(');
  const secondValidation = body.indexOf('validateSigningWindow_(freshTraining, freshPerson, onsiteCode, false)');
  const append = body.indexOf('appendObject_(SHEETS.SIGNATURES');
  assert.ok(firstValidation >= 0 && firstValidation < createFile, '파일 생성 전에 코드를 검증해야 합니다.');
  assert.ok(secondValidation > createFile && secondValidation < append, '잠금 뒤 저장 직전에 코드를 다시 검증해야 합니다.');
  assert.match(body, /catch \(error\)[\s\S]*file\.setTrashed\(true\)/);
  assert.match(body, /verificationMethod:\s*verification\.verificationMethod/);
  assert.match(body, /onsiteSessionId:\s*verification\.onsiteSessionId/);
});

test('공개 화면과 관리자 화면에 선택형 현장 접수 UI와 6자리 모바일 입력이 있다', () => {
  assert.match(index, /id="trainingOnsiteCodeRequired"[^>]*type="checkbox"/);
  assert.match(index, /id="onsiteCodeInput"[^>]*inputmode="numeric"[^>]*pattern="\[0-9\]\{6\}"[^>]*maxlength="6"/);
  assert.match(index, /id="receptionDuration"[\s\S]*value="5"[\s\S]*value="10"[\s\S]*value="15"/);
  assert.match(index, /id="startTrainingReception"/);
  assert.match(index, /id="closeTrainingReception"/);
  assert.match(app, /onsiteCodeRequired\s*\?\s*\{ onsiteCode \}/);
  assert.match(app, /\^\\d\{6\}\$/);
  assert.match(app, /get_training_reception_status/);
  assert.match(app, /start_training_reception/);
  assert.match(app, /close_training_reception/);
  assert.doesNotMatch(app, /localStorage[^\n]*(?:onsite|reception)|(?:onsite|reception)[^\n]*localStorage/i);
  assert.match(styles, /@media\s*\([^)]*max-width[^)]*\)[\s\S]*(?:onsite-code|reception)/);
});

test('프런트는 오래된 접수 응답을 버리고 서버 설정·접수 시간·확인 방식을 다시 반영한다', () => {
  const fetchBody = app.slice(app.indexOf('async function fetchReceptionStatus'), app.indexOf('function startReceptionTimers'));
  const startBody = app.slice(app.indexOf('async function startTrainingReception'), app.indexOf('async function closeTrainingReception'));
  const closeBody = app.slice(app.indexOf('async function closeTrainingReception'), app.indexOf('function collapseTrainingReception'));
  const submitBody = app.slice(app.indexOf('async function submitSignature'), app.indexOf('function renderPrivacy'));
  const normalizeBody = app.slice(app.indexOf('function normalizeReceptionStatus'), app.indexOf('function clearReceptionCode'));
  const renderBody = app.slice(app.indexOf('function renderTrainingReceptionPanel'), app.indexOf('async function fetchReceptionStatus'));

  assert.match(app, /receptionGeneration:\s*0/);
  assert.match(app, /function deactivateTrainingReception\([\s\S]*receptionGeneration \+= 1/);
  for (const body of [fetchBody, startBody, closeBody]) {
    assert.match(body, /const generation = state\.receptionGeneration/);
    assert.match(body, /state\.receptionGeneration !== generation/);
  }

  assert.match(submitBody, /refreshSelectedTrainingForSubmission\(\)/);
  assert.match(submitBody, /error\.code === 'ONSITE_CODE_REQUIRED'[\s\S]*selectedTraining\.onsiteCodeRequired = true/);
  assert.match(submitBody, /publicTraining\.onsiteCodeRequired = true/);
  assert.match(normalizeBody, /data\.durationMinutes[\s\S]*previous\?\.durationMinutes/);
  assert.match(renderBody, /entry\.durationMinutes[\s\S]*receptionDuration'\)\.value = String\(entry\.durationMinutes\)/);
  assert.match(startBody, /durationMinutes:\s*state\.receptionDurationMinutes/);

  assert.match(app, /verificationMethod === 'onsite_code'\s*\?\s*'현장 코드 확인'/);
  assert.match(backend, /verificationMethod:\s*String\(row\.verificationMethod \|\| 'legacy'\)/);
});

test('현장 접수 상태·공개 연수에는 코드 원문이나 Drive 정보가 노출되지 않는다', () => {
  const publicTrainingBody = backend.slice(backend.indexOf('function publicTraining_'), backend.indexOf('function staffSort_'));
  const publicReceptionBody = backend.slice(backend.indexOf('function publicTrainingReception_'), backend.indexOf('function validateTrainingReception_'));
  assert.match(publicTrainingBody, /onsiteCodeRequired:\s*bool_/);
  assert.doesNotMatch(publicTrainingBody, /codeHash|sessionId|expiresAt|fileId|Drive/);
  assert.doesNotMatch(publicReceptionBody, /codeHash|salt|imageFileId|DriveApp|getUrl/);
  assert.match(backend, /ONSITE_CODE_SECRET/);
  assert.match(backend, /computeHmacSha256Signature/);
  assert.match(backend, /ONSITE_CODE_MAX_FAILURES:\s*5/);
  assert.match(backend, /TRAININGS:[^\n]*'audienceDepartments', 'onsiteCodeRequired'/);
  assert.match(backend, /SIGNATURES:[^\n]*'scopeDate', 'verificationMethod', 'onsiteSessionId'/);
});
