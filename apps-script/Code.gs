/**
 * 학교 연수 전자서명 API
 * - Google 시트에 연결된 Apps Script 웹앱으로 배포합니다.
 * - 데이터와 파일은 이 스크립트를 소유한 Google 계정에만 저장합니다.
 */

const APP = Object.freeze({
  VERSION: '1.1.1',
  TIME_ZONE: 'Asia/Seoul',
  DATA_FILE: '학교 연수 전자서명 데이터',
  GUIDE_SHEET: '사용설명서',
  ROOT_FOLDER: '학교 연수 전자서명',
  SIGNATURE_FOLDER: '서명 원본',
  EXPORT_FOLDER: '출력 보관',
  SESSION_SECONDS: 1800,
  MAX_SIGNATURE_BYTES: 400 * 1024,
  MAX_STAFF: 500,
  MAX_EXPORT_ROWS: 200,
  EXPORT_BATCH_SIZE: 15,
  DOWNLOAD_CHUNK_SIZE: 512 * 1024
});

const SHEETS = Object.freeze({
  SETTINGS: { name: '설정', headers: ['key', 'value'] },
  STAFF: { name: '구성원', headers: ['id', 'department', 'name', 'active', 'sortOrder', 'createdAt'] },
  TRAININGS: { name: '연수', headers: ['id', 'title', 'target', 'date', 'daily', 'startTime', 'endTime', 'active', 'sortOrder', 'createdAt', 'updatedAt'] },
  SIGNATURES: { name: '서명', headers: ['id', 'trainingId', 'staffId', 'signDate', 'signTime', 'department', 'name', 'imageFileId', 'createdAt'] },
  EXPORTS: { name: '출력 작업', headers: ['jobId', 'trainingId', 'trainingTitle', 'date', 'sort', 'columns', 'showRate', 'status', 'progress', 'total', 'tempSpreadsheetId', 'pdfFileId', 'xlsxFileId', 'createdAt', 'updatedAt', 'error', 'purgedAt'] },
  AUDIT: { name: '감사 기록', headers: ['timestamp', 'action', 'target', 'count', 'detail'] }
});

const SETTING_KEYS = Object.freeze([
  'schoolName', 'subtitle', 'notice', 'brandColor',
  'privacyPurpose', 'privacyItems', 'privacyRetention'
]);

const INSTANCE_PROPERTIES = Object.freeze([
  'SPREADSHEET_ID', 'INSTANCE_ID', 'ROOT_FOLDER_ID', 'SIGNATURE_FOLDER_ID', 'EXPORT_FOLDER_ID',
  'SHARE_TOKEN', 'SETUP_CODE', 'ADMIN_PEPPER', 'ADMIN_EPOCH', 'ADMIN_SALT', 'ADMIN_HASH', 'FRONTEND_URL'
]);

/** 시트를 열 때 관리용 메뉴를 표시합니다. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🖊️ 전자서명 관리')
    .addItem('초기 설정 실행', 'initializeSystemFromMenu')
    .addItem('초기 설정 코드 보기', 'showSetupCode')
    .addItem('웹앱 주소 확인', 'showWebAppUrl')
    .addSeparator()
    .addItem('데이터 탭 표시·숨기기', 'toggleDataSheets')
    .addItem('관리자 비밀번호 복구', 'resetAdminPasswordFromMenu')
    .addItem('사용설명서 다시 만들기', 'rebuildGuideSheetFromMenu')
    .addToUi();
}

function onInstall() {
  onOpen();
}

function initializeSystemFromMenu() {
  try {
    const result = initializeSystem();
    SpreadsheetApp.getUi().alert(
      '초기 설정 완료',
      '현재 학교용 시트를 안전한 데이터 파일로 초기화했습니다.\n\n초기 설정 코드: ' + result.setupCode +
        '\n\n이 코드는 관리자 첫 비밀번호를 설정한 뒤 자동으로 폐기됩니다.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('초기 설정을 완료하지 못했습니다', String(error && error.message || error), SpreadsheetApp.getUi().ButtonSet.OK);
    throw error;
  }
}

function showSetupCode() {
  const properties = PropertiesService.getScriptProperties();
  const code = properties.getProperty('SETUP_CODE');
  const message = !properties.getProperty('SPREADSHEET_ID')
    ? '먼저 ‘초기 설정 실행’을 선택해 주세요.'
    : code
      ? '초기 설정 코드: ' + code + '\n\n관리자 첫 비밀번호 설정이 끝나면 이 코드는 자동 폐기됩니다.'
      : '관리자 첫 설정이 이미 완료되어 초기 설정 코드가 폐기되었습니다.';
  SpreadsheetApp.getUi().alert('초기 설정 코드', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function showWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  const message = url
    ? '현재 웹앱 주소:\n\n' + url + '\n\n주소가 /exec로 끝나는지 확인하세요.'
    : '아직 웹앱으로 배포되지 않았습니다. Apps Script에서 ‘배포 → 새 배포 → 웹 앱’을 실행해 주세요.';
  SpreadsheetApp.getUi().alert('웹앱 주소', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function toggleDataSheets() {
  const spreadsheet = boundSpreadsheet_();
  ensureGuideSheet_(spreadsheet, false);
  const sheets = dataSheetDefinitions_().map(function(definition) { return spreadsheet.getSheetByName(definition.name); }).filter(Boolean);
  const shouldShow = sheets.some(function(sheet) { return sheet.isSheetHidden(); });
  sheets.forEach(function(sheet) { if (shouldShow) sheet.showSheet(); else sheet.hideSheet(); });
  SpreadsheetApp.getUi().alert('데이터 탭', shouldShow ? '데이터 탭을 표시했습니다.' : '데이터 탭을 숨겼습니다.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function resetAdminPasswordFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('관리자 비밀번호 복구', '기존 관리자 세션을 모두 끝내고 임시 비밀번호를 발급할까요?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;
  const temporaryPassword = resetAdminPasswordFromEditor();
  ui.alert('임시 관리자 비밀번호', temporaryPassword + '\n\n관리자 화면에 로그인한 뒤 즉시 새 비밀번호로 변경해 주세요.', ui.ButtonSet.OK);
}

function rebuildGuideSheetFromMenu() {
  const spreadsheet = boundSpreadsheet_();
  ensureGuideSheet_(spreadsheet, true);
  SpreadsheetApp.getUi().alert('사용설명서', '사용설명서 탭을 다시 만들었습니다. 기존 데이터 탭은 변경하지 않았습니다.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function doGet() {
  return jsonOutput_({ ok: true, data: { service: '학교 연수 전자서명 API', version: APP.VERSION } });
}

function doPost(event) {
  try {
    if (!event || !event.postData || !event.postData.contents) apiError_('BAD_REQUEST', '요청 본문이 없습니다.');
    const request = JSON.parse(event.postData.contents);
    const data = dispatch_(request || {});
    return jsonOutput_({ ok: true, data: data === undefined ? null : data });
  } catch (error) {
    console.error(error && error.stack ? error.stack : error);
    return jsonOutput_({
      ok: false,
      error: {
        code: error.apiCode || 'SERVER_ERROR',
        message: error.apiCode ? error.message : '서버에서 요청을 처리하지 못했습니다.',
        details: error.details || null
      }
    });
  }
}

function dispatch_(request) {
  const action = string_(request.action, 80);
  if (!action) apiError_('BAD_REQUEST', '작업 이름이 없습니다.');
  if (action === 'get_setup_status') return getSetupStatus_();
  if (action === 'complete_setup') return completeSetup_(request);
  if (action === 'admin_login') return adminLogin_(request.password);
  if (action === 'get_public_data') return getPublicData_(request.shareToken);
  if (action === 'submit_signature') return submitSignature_(request);

  const sessionToken = requireAdminSession_(request.sessionToken);
  if (action === 'logout') return logout_(sessionToken);
  if (action === 'get_admin_data') return getAdminData_();
  if (action === 'save_settings') return saveSettings_(request.settings, request.frontendUrl);
  if (action === 'save_training') return saveTraining_(request.training);
  if (action === 'delete_training') return deleteTraining_(request.trainingId);
  if (action === 'move_training') return moveTraining_(request.trainingId, request.direction);
  if (action === 'save_staff_batch') return saveStaffBatch_(request.people);
  if (action === 'update_staff') return updateStaff_(request.person);
  if (action === 'delete_staff') return deleteStaff_(request.staffId);
  if (action === 'rename_department') return renameDepartment_(request.oldDepartment, request.newDepartment);
  if (action === 'list_records') return listRecords_(request.trainingId, request.date);
  if (action === 'delete_record') return deleteRecord_(request.recordId);
  if (action === 'rotate_share_token') return rotateShareToken_(request.frontendUrl);
  if (action === 'change_password') return changePassword_(request.currentPassword, request.newPassword);
  if (action === 'start_export') return startExport_(request);
  if (action === 'continue_export') return continueExport_(request.jobId);
  if (action === 'download_export_chunk') return downloadExportChunk_(request.jobId, request.format, request.offset, request.chunkSize);
  if (action === 'purge_originals') return purgeOriginals_(request.jobId, request.confirmation);
  apiError_('UNKNOWN_ACTION', '지원하지 않는 작업입니다.');
}

/**
 * 학교용 시트의 메뉴 또는 연결형 Apps Script 편집기에서 실행합니다.
 * 웹앱 요청에서는 호출하지 않으며, 반복 실행해도 기존 데이터와 비밀값을 보존합니다.
 */
function initializeSystem() {
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('TEMPLATE_LOCK') === '1') {
    apiError_('TEMPLATE_LOCKED', '이 파일은 비어 있는 배포용 원본입니다. 학교용 사본을 만든 뒤 사본에서 초기화해 주세요.');
  }

  const spreadsheet = boundSpreadsheet_();
  const spreadsheetId = spreadsheet.getId();
  const storedId = properties.getProperty('SPREADSHEET_ID');
  if (storedId && storedId !== spreadsheetId) {
    clearCopiedInstanceProperties_(properties);
  }

  DriveApp.getFileById(spreadsheetId).setName(APP.DATA_FILE);
  spreadsheet.setSpreadsheetTimeZone(APP.TIME_ZONE);
  ensureGuideSheet_(spreadsheet, false);
  dataSheetDefinitions_().forEach(function(definition) { ensureSheet_(spreadsheet, definition); });

  const rootFolder = getOrRepairFolder_(properties, 'ROOT_FOLDER_ID', null, APP.ROOT_FOLDER);
  const signatureFolder = getOrRepairFolder_(properties, 'SIGNATURE_FOLDER_ID', rootFolder, APP.SIGNATURE_FOLDER);
  const exportFolder = getOrRepairFolder_(properties, 'EXPORT_FOLDER_ID', rootFolder, APP.EXPORT_FOLDER);

  const secrets = {
    SPREADSHEET_ID: spreadsheetId,
    INSTANCE_ID: properties.getProperty('INSTANCE_ID') || randomToken_(24),
    ROOT_FOLDER_ID: rootFolder.getId(),
    SIGNATURE_FOLDER_ID: signatureFolder.getId(),
    EXPORT_FOLDER_ID: exportFolder.getId(),
    SHARE_TOKEN: properties.getProperty('SHARE_TOKEN') || randomToken_(24),
    ADMIN_PEPPER: properties.getProperty('ADMIN_PEPPER') || randomToken_(32),
    ADMIN_EPOCH: properties.getProperty('ADMIN_EPOCH') || '1'
  };
  if (!properties.getProperty('ADMIN_HASH')) secrets.SETUP_CODE = properties.getProperty('SETUP_CODE') || randomToken_(24);
  properties.setProperties(secrets, false);

  if (!readRows_(SHEETS.SETTINGS).length) {
    writeSettings_({
      schoolName: '학교 연수 전자서명',
      subtitle: '연수 참여 확인',
      notice: '',
      brandColor: '#315c54',
      privacyPurpose: '',
      privacyItems: '',
      privacyRetention: ''
    });
  }
  ensureCleanupTrigger_();
  hideDataSheets_(spreadsheet);
  audit_('initialize', spreadsheetId, 1, storedId ? '시스템 구성 복구' : '학교용 사본 초기화');
  const setupCode = properties.getProperty('SETUP_CODE') || '(이미 관리자 설정 완료)';
  console.log('초기 설정 코드: ' + setupCode);
  console.log('학교용 데이터 시트: ' + spreadsheet.getUrl());
  return { spreadsheetId: spreadsheetId, setupCode: setupCode };
}

function getSetupStatus_() {
  const properties = PropertiesService.getScriptProperties();
  return {
    initialized: Boolean(properties.getProperty('SPREADSHEET_ID')),
    adminConfigured: Boolean(properties.getProperty('ADMIN_HASH'))
  };
}

function completeSetup_(request) {
  requireInitialized_();
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('ADMIN_HASH')) apiError_('ALREADY_CONFIGURED', '관리자 설정이 이미 완료되었습니다.');
  if (!safeEqual_(string_(request.setupCode, 100), properties.getProperty('SETUP_CODE') || '')) apiError_('BAD_SETUP_CODE', '초기 설정 코드가 올바르지 않습니다.');
  const password = validatePassword_(request.password);
  const salt = randomToken_(18);
  properties.setProperties({
    ADMIN_SALT: salt,
    ADMIN_HASH: passwordHash_(password, salt),
    FRONTEND_URL: normalizeFrontendUrl_(request.frontendUrl)
  });
  properties.deleteProperty('SETUP_CODE');
  audit_('complete_setup', 'admin', 1, '관리자 비밀번호 최초 설정');
  return createAdminLoginResult_();
}

function adminLogin_(password) {
  requireInitialized_();
  const properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty('ADMIN_HASH')) apiError_('SETUP_REQUIRED', '관리자 첫 설정이 필요합니다.');
  const cache = CacheService.getScriptCache();
  const lockedUntil = Number(cache.get('admin-login-locked-until') || 0);
  if (lockedUntil > Date.now()) apiError_('LOGIN_LOCKED', '로그인 시도가 잠시 제한되었습니다. 5분 뒤 다시 시도해 주세요.');
  const valid = verifyPassword_(String(password || ''));
  if (!valid) {
    const failures = Number(cache.get('admin-login-failures') || 0) + 1;
    cache.put('admin-login-failures', String(failures), 300);
    if (failures >= 5) cache.put('admin-login-locked-until', String(Date.now() + 300000), 300);
    apiError_('BAD_PASSWORD', failures >= 5 ? '로그인 시도가 잠시 제한되었습니다. 5분 뒤 다시 시도해 주세요.' : '관리자 비밀번호가 올바르지 않습니다.');
  }
  cache.remove('admin-login-failures');
  cache.remove('admin-login-locked-until');
  cleanupStaleExportJobs();
  audit_('admin_login', 'admin', 1, '관리자 로그인');
  return createAdminLoginResult_();
}

function createAdminLoginResult_() {
  const token = randomToken_(32);
  const epoch = PropertiesService.getScriptProperties().getProperty('ADMIN_EPOCH') || '1';
  CacheService.getScriptCache().put('admin-session:' + token, epoch, APP.SESSION_SECONDS);
  return { sessionToken: token, expiresIn: APP.SESSION_SECONDS, adminData: getAdminData_() };
}

function requireAdminSession_(token) {
  const value = string_(token, 100);
  const epoch = PropertiesService.getScriptProperties().getProperty('ADMIN_EPOCH') || '1';
  if (!value || CacheService.getScriptCache().get('admin-session:' + value) !== epoch) apiError_('SESSION_EXPIRED', '관리자 로그인이 만료되었습니다. 다시 로그인해 주세요.');
  CacheService.getScriptCache().put('admin-session:' + value, epoch, APP.SESSION_SECONDS);
  return value;
}

function logout_(token) {
  CacheService.getScriptCache().remove('admin-session:' + token);
  return { loggedOut: true };
}

function changePassword_(currentPassword, newPassword) {
  if (!verifyPassword_(String(currentPassword || ''))) apiError_('BAD_PASSWORD', '현재 비밀번호가 올바르지 않습니다.');
  const password = validatePassword_(newPassword);
  const salt = randomToken_(18);
  PropertiesService.getScriptProperties().setProperties({ ADMIN_SALT: salt, ADMIN_HASH: passwordHash_(password, salt) });
  audit_('change_password', 'admin', 1, '관리자 비밀번호 변경');
  return { changed: true };
}

/**
 * 관리자 비밀번호를 잊었을 때 관리용 Google 계정으로 편집기를 열어 직접 실행합니다.
 * 모든 기존 관리자 세션을 무효화하고 임시 비밀번호를 실행 로그에 표시합니다.
 */
function resetAdminPasswordFromEditor() {
  requireInitialized_();
  const temporaryPassword = 'R' + randomToken_(18) + '9';
  const salt = randomToken_(18);
  const properties = PropertiesService.getScriptProperties();
  const nextEpoch = number_(properties.getProperty('ADMIN_EPOCH')) + 1;
  properties.setProperties({ ADMIN_SALT: salt, ADMIN_HASH: passwordHash_(temporaryPassword, salt), ADMIN_EPOCH: String(nextEpoch) });
  audit_('reset_admin_password', 'admin', 1, '편집기에서 임시 비밀번호 발급');
  console.log('임시 관리자 비밀번호: ' + temporaryPassword);
  console.log('로그인 후 공유·보안 메뉴에서 즉시 새 비밀번호로 변경하세요.');
  return temporaryPassword;
}

function passwordHash_(password, salt) {
  const pepper = PropertiesService.getScriptProperties().getProperty('ADMIN_PEPPER') || '';
  const bytes = Utilities.computeHmacSha256Signature(salt + ':' + password, pepper, Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

function verifyPassword_(password) {
  const properties = PropertiesService.getScriptProperties();
  const salt = properties.getProperty('ADMIN_SALT') || '';
  const expected = properties.getProperty('ADMIN_HASH') || '';
  return expected && safeEqual_(passwordHash_(password, salt), expected);
}

function validatePassword_(password) {
  const value = String(password || '');
  if (/^\d{4}$/.test(value)) return value;
  if (value.length < 10 || value.length > 100 || !/[A-Za-z가-힣]/.test(value) || !/\d/.test(value)) {
    apiError_('WEAK_PASSWORD', '관리자 비밀번호는 숫자 4자리 또는 문자와 숫자를 포함한 10자 이상 100자 이하로 설정해 주세요.');
  }
  return value;
}

function getPublicData_(shareToken) {
  requireInitialized_();
  requireShareToken_(shareToken);
  const settings = readSettings_();
  const privacyReady = privacyReady_(settings);
  if (!privacyReady) apiError_('PRIVACY_NOT_READY', '관리자가 개인정보 처리 안내를 완료하지 않았습니다.');
  const today = today_();
  const staff = readRows_(SHEETS.STAFF)
    .filter(row => bool_(row.active))
    .sort(staffSort_)
    .map(publicStaff_);
  const trainings = readRows_(SHEETS.TRAININGS)
    .filter(row => bool_(row.active) && (bool_(row.daily) || row.date === today))
    .sort(orderSort_)
    .map(publicTraining_);
  return { settings: settings, staff: staff, trainings: trainings, privacyReady: true, serverDate: today };
}

function submitSignature_(request) {
  requireInitialized_();
  requireShareToken_(request.shareToken);
  const trainingId = id_(request.trainingId, '연수');
  const staffId = id_(request.staffId, '구성원');
  const signatureData = String(request.signatureData || '');
  if (signatureData.length > APP.MAX_SIGNATURE_BYTES * 1.5) apiError_('SIGNATURE_TOO_LARGE', '서명 이미지가 너무 큽니다. 다시 작성해 주세요.');
  const match = signatureData.match(/^data:image\/png;base64,([A-Za-z0-9+/=]+)$/);
  if (!match) apiError_('BAD_SIGNATURE', '서명 이미지 형식이 올바르지 않습니다.');
  const bytes = Utilities.base64Decode(match[1]);
  if (bytes.length < 100 || bytes.length > APP.MAX_SIGNATURE_BYTES) apiError_('BAD_SIGNATURE', '서명 이미지 크기가 올바르지 않습니다.');
  if ((bytes[0] & 255) !== 137 || (bytes[1] & 255) !== 80 || (bytes[2] & 255) !== 78 || (bytes[3] & 255) !== 71) apiError_('BAD_SIGNATURE', 'PNG 서명 이미지만 등록할 수 있습니다.');

  const training = findRow_(SHEETS.TRAININGS, 'id', trainingId);
  const person = findRow_(SHEETS.STAFF, 'id', staffId);
  validateSigningWindow_(training, person);
  const now = new Date();
  const date = formatDate_(now, 'yyyy-MM-dd');
  const time = formatDate_(now, 'HH:mm:ss');
  const folder = getOrCreateTrainingFolder_(trainingId, training.title);
  const fileName = safeFileName_(date + '_' + person.department + '_' + person.name) + '.png';
  const file = folder.createFile(Utilities.newBlob(bytes, 'image/png', fileName));

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const freshTraining = findRow_(SHEETS.TRAININGS, 'id', trainingId);
    const freshPerson = findRow_(SHEETS.STAFF, 'id', staffId);
    validateSigningWindow_(freshTraining, freshPerson);
    const duplicate = readRows_(SHEETS.SIGNATURES).some(row => row.trainingId === trainingId && row.staffId === staffId && row.signDate === date);
    if (duplicate) apiError_('DUPLICATE', '[' + freshTraining.title + '] ' + freshPerson.name + '님은 오늘 이미 서명을 완료했습니다.');
    appendObject_(SHEETS.SIGNATURES, {
      id: Utilities.getUuid(), trainingId: trainingId, staffId: staffId,
      signDate: date, signTime: time, department: freshPerson.department, name: freshPerson.name,
      imageFileId: file.getId(), createdAt: now.toISOString()
    });
    return { registeredAt: now.toISOString(), signDate: date, signTime: time };
  } catch (error) {
    try { file.setTrashed(true); } catch (ignore) { /* Best effort orphan cleanup. */ }
    throw error;
  } finally {
    try { lock.releaseLock(); } catch (ignore) { /* Lock may not have been acquired. */ }
  }
}

function validateSigningWindow_(training, person) {
  if (!training || !bool_(training.active)) apiError_('TRAINING_CLOSED', '현재 서명할 수 없는 연수입니다.');
  if (!person || !bool_(person.active)) apiError_('STAFF_NOT_FOUND', '구성원 명단에서 확인할 수 없습니다.');
  const today = today_();
  if (!bool_(training.daily) && training.date !== today) apiError_('TRAINING_DATE', '오늘 서명할 수 있는 연수가 아닙니다.');
  const nowTime = formatDate_(new Date(), 'HH:mm');
  if (training.startTime && nowTime < training.startTime) apiError_('TOO_EARLY', '아직 서명 가능 시간이 아닙니다. ' + training.startTime + '부터 서명할 수 있습니다.');
  if (training.endTime && nowTime > training.endTime) apiError_('TOO_LATE', '서명 가능 시간이 종료되었습니다.');
}

function getAdminData_() {
  requireInitialized_();
  cleanupStaleExportJobs();
  const properties = PropertiesService.getScriptProperties();
  const shareToken = properties.getProperty('SHARE_TOKEN') || '';
  const frontendUrl = properties.getProperty('FRONTEND_URL') || '';
  const staff = readRows_(SHEETS.STAFF).sort(staffSort_).map(publicStaff_);
  const trainings = readRows_(SHEETS.TRAININGS).sort(orderSort_).map(publicTraining_);
  const signatures = readRows_(SHEETS.SIGNATURES);
  const exports = readRows_(SHEETS.EXPORTS).sort(function(a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)); });
  return {
    settings: readSettings_(), staff: staff, trainings: trainings, exports: exports,
    shareToken: shareToken, shareUrl: buildShareUrl_(frontendUrl, shareToken),
    stats: { staff: staff.length, trainings: trainings.length, signatures: signatures.length }
  };
}

function saveSettings_(input, frontendUrl) {
  const settings = {};
  SETTING_KEYS.forEach(function(key) { settings[key] = string_(input && input[key], key === 'notice' ? 1000 : 500); });
  if (!privacyReady_(settings)) apiError_('PRIVACY_REQUIRED', '학교명과 개인정보 처리 안내를 모두 입력해 주세요.');
  if (!/^#[0-9a-f]{6}$/i.test(settings.brandColor)) settings.brandColor = '#315c54';
  writeSettings_(settings);
  if (frontendUrl) PropertiesService.getScriptProperties().setProperty('FRONTEND_URL', normalizeFrontendUrl_(frontendUrl));
  audit_('save_settings', 'settings', 1, '기관 설정 변경');
  return { settings: settings };
}

function saveTraining_(input) {
  const training = normalizeTraining_(input);
  if (training.active && !privacyReady_(readSettings_())) apiError_('PRIVACY_REQUIRED', '개인정보 처리 안내를 모두 입력해야 연수를 활성화할 수 있습니다.');
  const sheet = sheet_(SHEETS.TRAININGS);
  const rows = readRowsWithNumbers_(SHEETS.TRAININGS);
  const existing = training.id ? rows.find(function(item) { return item.data.id === training.id; }) : null;
  const now = new Date().toISOString();
  if (existing) {
    const updated = Object.assign({}, existing.data, training, { updatedAt: now });
    writeObjectRow_(sheet, SHEETS.TRAININGS.headers, existing.rowNumber, updated);
  } else {
    training.id = Utilities.getUuid();
    training.sortOrder = rows.length ? Math.max.apply(null, rows.map(function(item) { return number_(item.data.sortOrder); })) + 1 : 1;
    training.createdAt = now;
    training.updatedAt = now;
    appendObject_(SHEETS.TRAININGS, training);
  }
  audit_('save_training', training.id, 1, training.title);
  return { training: publicTraining_(training) };
}

function normalizeTraining_(input) {
  const title = string_(input && input.title, 100);
  const target = string_(input && input.target, 200);
  const daily = bool_(input && input.daily);
  const date = string_(input && input.date, 10);
  const startTime = string_(input && input.startTime, 5);
  const endTime = string_(input && input.endTime, 5);
  if (!title) apiError_('VALIDATION', '연수명을 입력해 주세요.');
  if (!daily && !/^\d{4}-\d{2}-\d{2}$/.test(date)) apiError_('VALIDATION', '연수 날짜가 올바르지 않습니다.');
  if (startTime && !/^\d{2}:\d{2}$/.test(startTime)) apiError_('VALIDATION', '시작 시각이 올바르지 않습니다.');
  if (endTime && !/^\d{2}:\d{2}$/.test(endTime)) apiError_('VALIDATION', '종료 시각이 올바르지 않습니다.');
  if (startTime && endTime && startTime >= endTime) apiError_('VALIDATION', '종료 시각은 시작 시각보다 늦어야 합니다.');
  return { id: input && input.id ? id_(input.id, '연수') : '', title: title, target: target, date: daily ? date || today_() : date, daily: daily, startTime: startTime, endTime: endTime, active: bool_(input && input.active) };
}

function deleteTraining_(trainingId) {
  const id = id_(trainingId, '연수');
  deleteRowById_(SHEETS.TRAININGS, id);
  audit_('delete_training', id, 1, '서명 기록은 유지');
  return { deleted: true };
}

function moveTraining_(trainingId, direction) {
  const id = id_(trainingId, '연수');
  if (direction !== 'up' && direction !== 'down') apiError_('VALIDATION', '이동 방향이 올바르지 않습니다.');
  const rows = readRowsWithNumbers_(SHEETS.TRAININGS).sort(function(a, b) { return orderSort_(a.data, b.data); });
  const index = rows.findIndex(function(item) { return item.data.id === id; });
  const target = direction === 'up' ? index - 1 : index + 1;
  if (index < 0) apiError_('NOT_FOUND', '연수를 찾을 수 없습니다.');
  if (target < 0 || target >= rows.length) return { moved: false };
  const firstOrder = number_(rows[index].data.sortOrder) || index + 1;
  const secondOrder = number_(rows[target].data.sortOrder) || target + 1;
  const sheet = sheet_(SHEETS.TRAININGS);
  const sortColumn = SHEETS.TRAININGS.headers.indexOf('sortOrder') + 1;
  sheet.getRange(rows[index].rowNumber, sortColumn).setValue(secondOrder);
  sheet.getRange(rows[target].rowNumber, sortColumn).setValue(firstOrder);
  return { moved: true };
}

function saveStaffBatch_(people) {
  if (!Array.isArray(people) || !people.length) apiError_('VALIDATION', '등록할 구성원이 없습니다.');
  if (people.length > APP.MAX_STAFF) apiError_('VALIDATION', '한 번에 등록할 수 있는 인원은 ' + APP.MAX_STAFF + '명입니다.');
  const existing = readRows_(SHEETS.STAFF);
  if (existing.length >= APP.MAX_STAFF) apiError_('LIMIT', '구성원은 최대 ' + APP.MAX_STAFF + '명까지 등록할 수 있습니다.');
  const seen = {};
  existing.forEach(function(person) { seen[staffKey_(person.department, person.name)] = true; });
  let added = 0;
  let skipped = 0;
  let nextOrder = existing.length ? Math.max.apply(null, existing.map(function(person) { return number_(person.sortOrder); })) + 1 : 1;
  people.forEach(function(person) {
    const department = string_(person && person.department, 50);
    const name = string_(person && person.name, 50);
    const key = staffKey_(department, name);
    if (!department || !name || seen[key] || existing.length + added >= APP.MAX_STAFF) { skipped += 1; return; }
    appendObject_(SHEETS.STAFF, { id: Utilities.getUuid(), department: department, name: name, active: true, sortOrder: nextOrder++, createdAt: new Date().toISOString() });
    seen[key] = true;
    added += 1;
  });
  audit_('save_staff_batch', 'staff', added, '건너뜀 ' + skipped);
  return { added: added, skipped: skipped };
}

function updateStaff_(input) {
  const id = id_(input && input.id, '구성원');
  const department = string_(input && input.department, 50);
  const name = string_(input && input.name, 50);
  if (!department || !name) apiError_('VALIDATION', '부서와 성명을 입력해 주세요.');
  const rows = readRowsWithNumbers_(SHEETS.STAFF);
  const current = rows.find(function(item) { return item.data.id === id; });
  if (!current) apiError_('NOT_FOUND', '구성원을 찾을 수 없습니다.');
  const duplicate = rows.some(function(item) { return item.data.id !== id && staffKey_(item.data.department, item.data.name) === staffKey_(department, name); });
  if (duplicate) apiError_('DUPLICATE_STAFF', '같은 부서와 성명의 구성원이 이미 있습니다.');
  writeObjectRow_(sheet_(SHEETS.STAFF), SHEETS.STAFF.headers, current.rowNumber, Object.assign({}, current.data, { department: department, name: name }));
  audit_('update_staff', id, 1, department + ' ' + name);
  return { updated: true };
}

function deleteStaff_(staffId) {
  const id = id_(staffId, '구성원');
  deleteRowById_(SHEETS.STAFF, id);
  audit_('delete_staff', id, 1, '기존 서명 기록 유지');
  return { deleted: true };
}

function renameDepartment_(oldDepartment, newDepartment) {
  const oldName = string_(oldDepartment, 50);
  const newName = string_(newDepartment, 50);
  if (!oldName || !newName) apiError_('VALIDATION', '기존 부서와 새 부서명을 입력해 주세요.');
  const sheet = sheet_(SHEETS.STAFF);
  const rows = readRowsWithNumbers_(SHEETS.STAFF);
  const departmentColumn = SHEETS.STAFF.headers.indexOf('department') + 1;
  let updated = 0;
  rows.forEach(function(item) {
    if (item.data.department === oldName) {
      sheet.getRange(item.rowNumber, departmentColumn).setValue(newName);
      updated += 1;
    }
  });
  if (!updated) apiError_('NOT_FOUND', '변경할 부서를 찾지 못했습니다.');
  audit_('rename_department', oldName, updated, newName);
  return { updated: updated };
}

function listRecords_(trainingId, date) {
  const id = id_(trainingId, '연수');
  const signDate = validDate_(date);
  const records = readRows_(SHEETS.SIGNATURES)
    .filter(function(row) { return row.trainingId === id && row.signDate === signDate; })
    .sort(function(a, b) { return String(a.createdAt).localeCompare(String(b.createdAt)); })
    .map(function(row) { return { id: row.id, trainingId: row.trainingId, signDate: row.signDate, signTime: row.signTime, department: row.department, name: row.name }; });
  return { records: records };
}

function deleteRecord_(recordId) {
  const id = id_(recordId, '서명 기록');
  const rows = readRowsWithNumbers_(SHEETS.SIGNATURES);
  const record = rows.find(function(item) { return item.data.id === id; });
  if (!record) apiError_('NOT_FOUND', '서명 기록을 찾을 수 없습니다.');
  trashFileIfExists_(record.data.imageFileId);
  sheet_(SHEETS.SIGNATURES).deleteRow(record.rowNumber);
  audit_('delete_record', id, 1, record.data.department + ' ' + record.data.name);
  return { deleted: true };
}

function rotateShareToken_(frontendUrl) {
  const properties = PropertiesService.getScriptProperties();
  const token = randomToken_(24);
  const url = normalizeFrontendUrl_(frontendUrl || properties.getProperty('FRONTEND_URL') || '');
  properties.setProperties({ SHARE_TOKEN: token, FRONTEND_URL: url });
  audit_('rotate_share_token', 'share', 1, '기존 공유 링크 무효화');
  return { shareToken: token, shareUrl: buildShareUrl_(url, token) };
}

function requireShareToken_(token) {
  const expected = PropertiesService.getScriptProperties().getProperty('SHARE_TOKEN') || '';
  if (!expected || !safeEqual_(string_(token, 100), expected)) apiError_('INVALID_LINK', '공유 링크가 올바르지 않거나 교체되었습니다. 담당자에게 새 링크를 받아 주세요.');
}

function startExport_(request) {
  const trainingId = id_(request.trainingId, '연수');
  const date = validDate_(request.date);
  const columns = Math.max(1, Math.min(3, number_(request.columns) || 1));
  const sort = request.sort === 'name' ? 'name' : 'department';
  const showRate = bool_(request.showRate);
  const training = findRow_(SHEETS.TRAININGS, 'id', trainingId);
  if (!training) apiError_('NOT_FOUND', '연수를 찾을 수 없습니다.');
  const roster = buildExportRoster_(trainingId, date, sort);
  if (roster.length > APP.MAX_EXPORT_ROWS) apiError_('EXPORT_LIMIT', '한 번에 출력할 수 있는 인원은 ' + APP.MAX_EXPORT_ROWS + '명입니다.');

  const exportFolder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('EXPORT_FOLDER_ID'));
  const temporary = SpreadsheetApp.create('임시_' + safeFileName_(training.title) + '_' + date + '_' + Date.now());
  DriveApp.getFileById(temporary.getId()).moveTo(exportFolder);
  const output = temporary.getSheets()[0];
  output.setName('서명등록부');
  prepareExportSheet_(output, training, date, roster, columns, showRate);
  const dataSheet = temporary.insertSheet('_DATA');
  dataSheet.getRange(1, 1, 1, 6).setValues([['layoutIndex', 'staffId', 'department', 'name', 'time', 'fileId']]);
  if (roster.length) dataSheet.getRange(2, 1, roster.length, 6).setValues(roster.map(function(row, index) { return [index, row.staffId, row.department, row.name, row.time || '', row.fileId || '']; }));
  dataSheet.hideSheet();
  SpreadsheetApp.flush();

  const totalImages = roster.filter(function(row) { return Boolean(row.fileId); }).length;
  const now = new Date().toISOString();
  const job = {
    jobId: Utilities.getUuid(), trainingId: trainingId, trainingTitle: training.title, date: date,
    sort: sort, columns: columns, showRate: showRate, status: 'queued', progress: 0, total: totalImages,
    tempSpreadsheetId: temporary.getId(), pdfFileId: '', xlsxFileId: '', createdAt: now, updatedAt: now, error: '', purgedAt: ''
  };
  appendObject_(SHEETS.EXPORTS, job);
  audit_('start_export', job.jobId, roster.length, training.title + ' ' + date);
  return publicJob_(job);
}

function buildExportRoster_(trainingId, date, sort) {
  const signatures = readRows_(SHEETS.SIGNATURES).filter(function(row) { return row.trainingId === trainingId && row.signDate === date; });
  const signedByStaff = {};
  signatures.forEach(function(row) { signedByStaff[row.staffId] = row; });
  const roster = readRows_(SHEETS.STAFF).filter(function(row) { return bool_(row.active); }).map(function(person) {
    const signature = signedByStaff[person.id];
    return { staffId: person.id, department: person.department, name: person.name, time: signature ? signature.signTime : '', fileId: signature ? signature.imageFileId : '' };
  });
  roster.sort(function(a, b) {
    if (sort === 'name') return compareKo_(a.name, b.name) || compareKo_(a.department, b.department);
    return compareKo_(a.department, b.department) || compareKo_(a.name, b.name);
  });
  return roster;
}

function prepareExportSheet_(sheet, training, date, roster, columns, showRate) {
  const totalColumns = columns * 4;
  sheet.clear();
  sheet.getRange(1, 1, 1, totalColumns).merge().setValue(training.title + ' 서명등록부').setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
  const signedCount = roster.filter(function(row) { return Boolean(row.fileId); }).length;
  const rate = roster.length ? Math.round(signedCount / roster.length * 1000) / 10 : 0;
  const subline = formatKoreanDate_(date) + (training.target ? ' · ' + training.target : '') + (showRate ? ' · 서명 ' + signedCount + '/' + roster.length + '명 (' + rate + '%)' : '');
  sheet.getRange(2, 1, 1, totalColumns).merge().setValue(subline).setHorizontalAlignment('center').setFontColor('#4c5c57');
  const rowsPerBlock = Math.max(1, Math.ceil(roster.length / columns));
  for (let block = 0; block < columns; block += 1) {
    const base = 1 + block * 4;
    sheet.getRange(3, base, 1, 4).setValues([['부서', '성명', '서명', '시각']]).setFontWeight('bold').setBackground('#dfece8').setHorizontalAlignment('center');
    sheet.setColumnWidth(base, 90);
    sheet.setColumnWidth(base + 1, 85);
    sheet.setColumnWidth(base + 2, 150);
    sheet.setColumnWidth(base + 3, 70);
  }
  roster.forEach(function(row, index) {
    const position = exportPosition_(index, rowsPerBlock);
    const base = 1 + position.block * 4;
    sheet.getRange(position.row, base, 1, 4).setValues([[row.department, row.name, '', row.time || '']]);
    sheet.getRange(position.row, base, 1, 4).setBorder(true, true, true, true, true, true).setVerticalAlignment('middle');
    sheet.getRange(position.row, base + 2).setHorizontalAlignment('center');
    sheet.setRowHeight(position.row, 58);
  });
  sheet.setFrozenRows(3);
  sheet.setHiddenGridlines(true);
}

function continueExport_(jobId) {
  const id = id_(jobId, '출력 작업');
  const entry = findRowWithNumber_(SHEETS.EXPORTS, 'jobId', id);
  if (!entry) apiError_('NOT_FOUND', '출력 작업을 찾을 수 없습니다.');
  let job = entry.data;
  if (job.status === 'complete' || job.status === 'failed' || job.status === 'expired') return publicJob_(job);
  try {
    updateExportJob_(entry.rowNumber, { status: 'processing', error: '' });
    const spreadsheet = SpreadsheetApp.openById(job.tempSpreadsheetId);
    const dataSheet = spreadsheet.getSheetByName('_DATA');
    const output = spreadsheet.getSheetByName('서명등록부');
    if (!dataSheet || !output) throw new Error('임시 출력 데이터를 찾을 수 없습니다.');
    const rowCount = Math.max(0, dataSheet.getLastRow() - 1);
    const data = rowCount ? dataSheet.getRange(2, 1, rowCount, 6).getValues() : [];
    const withImages = data.filter(function(row) { return Boolean(row[5]); });
    const start = number_(job.progress);
    const batch = withImages.slice(start, start + APP.EXPORT_BATCH_SIZE);
    const rowsPerBlock = Math.max(1, Math.ceil(data.length / number_(job.columns)));
    batch.forEach(function(row) {
      const layoutIndex = number_(row[0]);
      const position = exportPosition_(layoutIndex, rowsPerBlock);
      const column = 3 + position.block * 4;
      try {
        const blob = DriveApp.getFileById(String(row[5])).getBlob();
        const image = output.insertImage(blob, column, position.row);
        image.setWidth(140).setHeight(50);
      } catch (imageError) {
        output.getRange(position.row, column).setValue('이미지 없음').setFontColor('#b4473d').setFontSize(8);
      }
    });
    SpreadsheetApp.flush();
    const nextProgress = start + batch.length;
    updateExportJob_(entry.rowNumber, { progress: nextProgress, total: withImages.length, status: 'processing' });
    if (nextProgress >= withImages.length) {
      job = finalizeExport_(entry.rowNumber, Object.assign({}, job, { progress: nextProgress, total: withImages.length }));
    } else {
      job = findRowWithNumber_(SHEETS.EXPORTS, 'jobId', id).data;
    }
    return publicJob_(job);
  } catch (error) {
    updateExportJob_(entry.rowNumber, { status: 'failed', error: String(error.message || error).slice(0, 500) });
    return publicJob_(findRowWithNumber_(SHEETS.EXPORTS, 'jobId', id).data);
  }
}

function finalizeExport_(rowNumber, job) {
  const spreadsheet = SpreadsheetApp.openById(job.tempSpreadsheetId);
  const dataSheet = spreadsheet.getSheetByName('_DATA');
  if (dataSheet) spreadsheet.deleteSheet(dataSheet);
  SpreadsheetApp.flush();
  Utilities.sleep(1000);
  const output = spreadsheet.getSheetByName('서명등록부');
  const base = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + '/export';
  const auth = { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: false };
  const pdfUrl = base + '?format=pdf&gid=' + output.getSheetId() + '&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&pagenumbers=true&gridlines=false&fzr=true';
  const xlsxUrl = base + '?format=xlsx';
  const exportFolder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('EXPORT_FOLDER_ID'));
  const safeName = safeFileName_(job.trainingTitle + '_' + job.date + '_서명등록부');
  const pdfFile = exportFolder.createFile(UrlFetchApp.fetch(pdfUrl, auth).getBlob().setName(safeName + '.pdf'));
  const xlsxFile = exportFolder.createFile(UrlFetchApp.fetch(xlsxUrl, auth).getBlob().setName(safeName + '.xlsx'));
  DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
  updateExportJob_(rowNumber, { status: 'complete', pdfFileId: pdfFile.getId(), xlsxFileId: xlsxFile.getId(), tempSpreadsheetId: '', error: '' });
  audit_('complete_export', job.jobId, number_(job.total), safeName);
  return findRowWithNumber_(SHEETS.EXPORTS, 'jobId', job.jobId).data;
}

function downloadExportChunk_(jobId, format, offset, chunkSize) {
  const job = findRow_(SHEETS.EXPORTS, 'jobId', id_(jobId, '출력 작업'));
  if (!job || job.status !== 'complete') apiError_('EXPORT_NOT_READY', '출력 파일이 아직 준비되지 않았습니다.');
  const normalizedFormat = format === 'pdf' ? 'pdf' : format === 'xlsx' ? 'xlsx' : '';
  if (!normalizedFormat) apiError_('VALIDATION', '파일 형식이 올바르지 않습니다.');
  const fileId = normalizedFormat === 'pdf' ? job.pdfFileId : job.xlsxFileId;
  if (!fileId) apiError_('NOT_FOUND', '출력 파일을 찾을 수 없습니다.');
  const file = DriveApp.getFileById(fileId);
  const bytes = file.getBlob().getBytes();
  const start = Math.max(0, number_(offset));
  const size = Math.max(32768, Math.min(APP.DOWNLOAD_CHUNK_SIZE, number_(chunkSize) || APP.DOWNLOAD_CHUNK_SIZE));
  const end = Math.min(bytes.length, start + size);
  return {
    base64: Utilities.base64Encode(bytes.slice(start, end)), nextOffset: end, totalBytes: bytes.length,
    fileName: file.getName(), mimeType: file.getMimeType()
  };
}

function purgeOriginals_(jobId, confirmation) {
  const entry = findRowWithNumber_(SHEETS.EXPORTS, 'jobId', id_(jobId, '출력 작업'));
  if (!entry || entry.data.status !== 'complete' || !entry.data.pdfFileId || !entry.data.xlsxFileId) apiError_('EXPORT_REQUIRED', 'PDF와 엑셀 생성이 완료된 작업만 원본을 삭제할 수 있습니다.');
  const job = entry.data;
  if (!safeEqual_(String(confirmation || ''), String(job.trainingTitle || ''))) apiError_('CONFIRMATION_MISMATCH', '연수명이 일치하지 않습니다.');
  assertFileExists_(job.pdfFileId);
  assertFileExists_(job.xlsxFileId);
  const sheet = sheet_(SHEETS.SIGNATURES);
  const records = readRowsWithNumbers_(SHEETS.SIGNATURES)
    .filter(function(item) { return item.data.trainingId === job.trainingId && item.data.signDate === job.date; })
    .sort(function(a, b) { return b.rowNumber - a.rowNumber; });
  let deleted = 0;
  let failed = 0;
  records.forEach(function(item) {
    try {
      trashFileIfExists_(item.data.imageFileId);
      sheet.deleteRow(item.rowNumber);
      deleted += 1;
    } catch (error) {
      failed += 1;
    }
  });
  if (!failed) updateExportJob_(entry.rowNumber, { purgedAt: new Date().toISOString() });
  audit_('purge_originals', job.jobId, deleted, '실패 ' + failed);
  return { deleted: deleted, failed: failed };
}

function cleanupStaleExportJobs() {
  const cutoff = Date.now() - 24 * 60 * 60 * 1000;
  const rows = readRowsWithNumbers_(SHEETS.EXPORTS);
  rows.forEach(function(item) {
    const status = item.data.status;
    const created = new Date(item.data.createdAt).getTime();
    if (status !== 'complete' && status !== 'expired' && created && created < cutoff) {
      if (item.data.tempSpreadsheetId) trashFileIfExists_(item.data.tempSpreadsheetId);
      updateExportJob_(item.rowNumber, { status: 'expired', tempSpreadsheetId: '', error: '24시간이 지나 자동 정리됨' });
    }
  });
}

function updateExportJob_(rowNumber, changes) {
  const sheet = sheet_(SHEETS.EXPORTS);
  const current = rowObject_(SHEETS.EXPORTS.headers, sheet.getRange(rowNumber, 1, 1, SHEETS.EXPORTS.headers.length).getValues()[0]);
  writeObjectRow_(sheet, SHEETS.EXPORTS.headers, rowNumber, Object.assign({}, current, changes, { updatedAt: new Date().toISOString() }));
}

function publicJob_(job) {
  return {
    jobId: job.jobId, trainingId: job.trainingId, trainingTitle: job.trainingTitle, date: job.date,
    sort: job.sort, columns: number_(job.columns), showRate: bool_(job.showRate), status: job.status,
    progress: number_(job.progress), total: number_(job.total), createdAt: job.createdAt, updatedAt: job.updatedAt,
    error: job.error || '', purgedAt: job.purgedAt || ''
  };
}

function exportPosition_(index, rowsPerBlock) {
  const block = Math.floor(index / rowsPerBlock);
  return { block: block, row: 4 + (index % rowsPerBlock) };
}

function spreadsheet_() {
  requireInitialized_();
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
}

/** 초기화·시트 메뉴에서만 사용합니다. 웹앱 요청 경로에서는 호출하지 않습니다. */
function boundSpreadsheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) apiError_('BOUND_SHEET_REQUIRED', '이 함수는 학교용 Google 시트에 연결된 Apps Script에서 실행해야 합니다.');
  return spreadsheet;
}

function dataSheetDefinitions_() {
  return [SHEETS.SETTINGS, SHEETS.STAFF, SHEETS.TRAININGS, SHEETS.SIGNATURES, SHEETS.EXPORTS, SHEETS.AUDIT];
}

function clearCopiedInstanceProperties_(properties) {
  INSTANCE_PROPERTIES.forEach(function(key) { properties.deleteProperty(key); });
}

function getOrRepairFolder_(properties, propertyKey, parentFolder, name) {
  const existingId = properties.getProperty(propertyKey);
  if (existingId) {
    try {
      const existing = DriveApp.getFolderById(existingId);
      existing.getName();
      return existing;
    } catch (error) {
      console.warn(propertyKey + ' 폴더를 찾을 수 없어 다시 만듭니다: ' + String(error && error.message || error));
    }
  }
  return parentFolder ? parentFolder.createFolder(name) : DriveApp.createFolder(name);
}

function ensureGuideSheet_(spreadsheet, rebuild) {
  let sheet = spreadsheet.getSheetByName(APP.GUIDE_SHEET);
  if (!sheet) {
    const reusable = spreadsheet.getSheets().find(function(candidate) {
      return dataSheetDefinitions_().every(function(definition) { return candidate.getName() !== definition.name; });
    });
    sheet = reusable || spreadsheet.insertSheet(APP.GUIDE_SHEET, 0);
    sheet.setName(APP.GUIDE_SHEET);
    rebuild = true;
  }
  if (!rebuild && String(sheet.getRange('A1').getValue()).indexOf('학교 연수 전자서명') >= 0) {
    sheet.showSheet();
    spreadsheet.setActiveSheet(sheet);
    return sheet;
  }

  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(3);
  sheet.setColumnWidth(1, 34);
  sheet.setColumnWidth(2, 170);
  sheet.setColumnWidth(3, 170);
  sheet.setColumnWidth(4, 170);
  sheet.setColumnWidth(5, 170);
  sheet.setColumnWidth(6, 34);
  for (let row = 1; row <= 38; row += 1) sheet.setRowHeight(row, row === 1 ? 18 : 28);

  sheet.getRange('A1:F38').setBackground('#F6F4EE').setFontFamily('Arial').setFontColor('#23342F').setVerticalAlignment('middle');
  sheet.getRange('B2:E2').merge().setValue('학교 연수 전자서명 시스템').setFontSize(22).setFontWeight('bold').setFontColor('#244C43');
  sheet.getRange('B3:E3').merge().setValue('Google 시트 사본 하나로 시작하는 비공개 운영 안내').setFontSize(11).setFontColor('#5E6F69');
  sheet.getRange('B5:E5').merge().setValue('처음 설치할 때').setFontSize(14).setFontWeight('bold').setBackground('#DDEBE5');

  const installRows = [
    ['1', '학교용 사본 만들기', '이 배포용 원본은 비워 두고, 파일 → 사본 만들기로 학교용 파일을 만듭니다.'],
    ['2', '초기 설정 실행', '학교용 사본에서 🖊️ 전자서명 관리 → 초기 설정 실행을 선택하고 권한을 승인합니다.'],
    ['3', '웹앱 배포', '확장 프로그램 → Apps Script → 배포 → 새 배포 → 웹 앱에서 ‘나로 실행 / 모든 사용자’를 선택합니다.'],
    ['4', '화면 연결', '웹앱의 /exec 주소를 GitHub Pages의 assets/config.js에 입력합니다.'],
    ['5', '관리자 첫 설정', '초기 설정 코드로 숫자 4자리 또는 문자·숫자 포함 10자 이상의 비밀번호를 만들고 학교 정보·개인정보 안내·명단·연수를 등록합니다.']
  ];
  installRows.forEach(function(row, index) {
    const targetRow = 7 + index * 2;
    sheet.getRange(targetRow, 2).setValue(row[0]).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#315C54').setFontColor('#FFFFFF');
    sheet.getRange(targetRow, 3, 1, 2).merge().setValue(row[1]).setFontWeight('bold').setBackground('#FFFFFF');
    sheet.getRange(targetRow + 1, 3, 1, 3).merge().setValue(row[2]).setWrap(true).setFontSize(10).setFontColor('#56645F');
  });

  sheet.getRange('B18:E18').merge().setValue('운영할 때 지킬 점').setFontSize(14).setFontWeight('bold').setBackground('#F2D9B8');
  const cautions = [
    '공유 링크에는 본인 인증이 없습니다. 학교 내부에서만 공유하고 유출되면 공유 키를 교체하세요.',
    '개인정보 안내의 목적·항목·보관기간을 모두 입력하기 전에는 연수를 활성화하지 마세요.',
    'PDF와 XLSX가 모두 생성된 것을 확인한 뒤에만 해당 날짜의 원본 서명 기록과 PNG를 삭제하세요.',
    '이 시스템은 연수 참여 확인용이며 공식 동의·결재·법적 전자서명 용도로 사용하지 않습니다.'
  ];
  cautions.forEach(function(text, index) {
    const row = 20 + index * 2;
    sheet.getRange(row, 2).setValue('•').setFontWeight('bold').setFontColor('#B05A31').setHorizontalAlignment('center');
    sheet.getRange(row, 3, 1, 3).merge().setValue(text).setWrap(true).setFontSize(10);
  });

  sheet.getRange('B29:E29').merge().setValue('시트 메뉴').setFontSize(14).setFontWeight('bold').setBackground('#DDEBE5');
  sheet.getRange('B31:E34').setValues([
    ['초기 설정 코드 보기', '관리자 첫 비밀번호 설정 전 코드 확인', '', ''],
    ['웹앱 주소 확인', '현재 배포된 /exec 주소 확인', '', ''],
    ['데이터 탭 표시·숨기기', '기본적으로 숨긴 데이터 탭 점검', '', ''],
    ['관리자 비밀번호 복구', '기존 세션 무효화 후 임시 비밀번호 발급', '', '']
  ]);
  sheet.getRange('B31:B34').setFontWeight('bold').setBackground('#FFFFFF');
  sheet.getRange('C31:E34').setFontColor('#56645F').setWrap(true);
  sheet.getRange('B36:E36').merge().setValue('참고: developers.google.com/apps-script/guides/bound · developers.google.com/apps-script/guides/web').setFontSize(9).setFontColor('#6D7773');
  sheet.getRange('B37:E37').merge().setValue('배포용 원본에는 명단·서명·비밀번호·공유 키·웹앱 주소를 저장하지 않습니다.').setFontSize(9).setFontWeight('bold').setFontColor('#8C4A2F');
  sheet.getRange('B2:E37').setBorder(false, false, false, false, false, false);
  sheet.showSheet();
  spreadsheet.setActiveSheet(sheet);
  return sheet;
}

function hideDataSheets_(spreadsheet) {
  ensureGuideSheet_(spreadsheet, false).showSheet();
  dataSheetDefinitions_().forEach(function(definition) {
    const sheet = spreadsheet.getSheetByName(definition.name);
    if (sheet) sheet.hideSheet();
  });
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(APP.GUIDE_SHEET));
}

function sheet_(definition) {
  return ensureSheet_(spreadsheet_(), definition);
}

function ensureSheet_(spreadsheet, definition) {
  let sheet = spreadsheet.getSheetByName(definition.name);
  if (!sheet) sheet = spreadsheet.insertSheet(definition.name);
  const width = definition.headers.length;
  const current = sheet.getRange(1, 1, 1, width).getValues()[0];
  if (current.join('\u0000') !== definition.headers.join('\u0000')) {
    sheet.getRange(1, 1, 1, width).setValues([definition.headers]).setFontWeight('bold').setBackground('#dfece8');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readRows_(definition) {
  return readRowsWithNumbers_(definition).map(function(item) { return item.data; });
}

function readRowsWithNumbers_(definition) {
  const sheet = sheet_(definition);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, definition.headers.length).getValues()
    .map(function(values, index) { return { rowNumber: index + 2, data: rowObject_(definition.headers, values) }; })
    .filter(function(item) { return Object.keys(item.data).some(function(key) { return item.data[key] !== ''; }); });
}

function rowObject_(headers, values) {
  const result = {};
  headers.forEach(function(header, index) { result[header] = values[index]; });
  return result;
}

function objectValues_(headers, object) {
  return headers.map(function(header) { return object[header] === undefined || object[header] === null ? '' : object[header]; });
}

function appendObject_(definition, object) {
  sheet_(definition).appendRow(objectValues_(definition.headers, object));
}

function writeObjectRow_(sheet, headers, rowNumber, object) {
  sheet.getRange(rowNumber, 1, 1, headers.length).setValues([objectValues_(headers, object)]);
}

function findRow_(definition, key, value) {
  const entry = findRowWithNumber_(definition, key, value);
  return entry ? entry.data : null;
}

function findRowWithNumber_(definition, key, value) {
  return readRowsWithNumbers_(definition).find(function(item) { return String(item.data[key]) === String(value); }) || null;
}

function deleteRowById_(definition, id) {
  const entry = findRowWithNumber_(definition, 'id', id);
  if (!entry) apiError_('NOT_FOUND', definition.name + ' 항목을 찾을 수 없습니다.');
  sheet_(definition).deleteRow(entry.rowNumber);
}

function readSettings_() {
  const values = {};
  SETTING_KEYS.forEach(function(key) { values[key] = ''; });
  readRows_(SHEETS.SETTINGS).forEach(function(row) {
    if (SETTING_KEYS.indexOf(String(row.key)) >= 0) values[String(row.key)] = String(row.value || '');
  });
  return values;
}

function writeSettings_(settings) {
  const sheet = sheet_(SHEETS.SETTINGS);
  const rows = readRowsWithNumbers_(SHEETS.SETTINGS);
  SETTING_KEYS.forEach(function(key) {
    const value = settings[key] === undefined ? '' : String(settings[key]);
    const existing = rows.find(function(item) { return item.data.key === key; });
    if (existing) writeObjectRow_(sheet, SHEETS.SETTINGS.headers, existing.rowNumber, { key: key, value: value });
    else appendObject_(SHEETS.SETTINGS, { key: key, value: value });
  });
}

function privacyReady_(settings) {
  return ['schoolName', 'subtitle', 'privacyPurpose', 'privacyItems', 'privacyRetention']
    .every(function(key) { return Boolean(String(settings[key] || '').trim()); });
}

function publicStaff_(row) {
  return { id: String(row.id), department: String(row.department), name: String(row.name), active: bool_(row.active), sortOrder: number_(row.sortOrder) };
}

function publicTraining_(row) {
  return {
    id: String(row.id), title: String(row.title), target: String(row.target || ''), date: String(row.date), daily: bool_(row.daily),
    startTime: String(row.startTime || ''), endTime: String(row.endTime || ''), active: bool_(row.active), sortOrder: number_(row.sortOrder)
  };
}

function staffSort_(a, b) {
  return number_(a.sortOrder) - number_(b.sortOrder) || compareKo_(a.department, b.department) || compareKo_(a.name, b.name);
}

function orderSort_(a, b) {
  return number_(a.sortOrder) - number_(b.sortOrder) || compareKo_(a.title, b.title);
}

function getOrCreateTrainingFolder_(trainingId, title) {
  const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('SIGNATURE_FOLDER_ID'));
  const folderName = safeFileName_(trainingId.slice(0, 8) + '_' + title);
  const iterator = root.getFoldersByName(folderName);
  return iterator.hasNext() ? iterator.next() : root.createFolder(folderName);
}

function trashFileIfExists_(fileId) {
  if (!fileId) return;
  try {
    const file = DriveApp.getFileById(String(fileId));
    if (!file.isTrashed()) file.setTrashed(true);
  } catch (error) {
    const message = String(error && error.message || error);
    if (/not found|does not exist|찾을 수|유효하지/i.test(message)) return;
    throw error;
  }
}

function assertFileExists_(fileId) {
  try {
    DriveApp.getFileById(String(fileId)).getName();
  } catch (error) {
    apiError_('EXPORT_MISSING', '보관된 출력 파일을 찾을 수 없어 원본을 삭제하지 않았습니다.');
  }
}

function audit_(action, target, count, detail) {
  appendObject_(SHEETS.AUDIT, { timestamp: new Date().toISOString(), action: action, target: String(target || ''), count: number_(count), detail: String(detail || '').slice(0, 500) });
}

function ensureCleanupTrigger_() {
  const exists = ScriptApp.getProjectTriggers().some(function(trigger) { return trigger.getHandlerFunction() === 'cleanupStaleExportJobs'; });
  if (!exists) ScriptApp.newTrigger('cleanupStaleExportJobs').timeBased().everyHours(6).create();
}

function requireInitialized_() {
  if (!PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')) apiError_('NOT_INITIALIZED', 'Apps Script 편집기에서 initializeSystem 함수를 먼저 실행해 주세요.');
}

function jsonOutput_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function apiError_(code, message, details) {
  const error = new Error(message);
  error.apiCode = code;
  error.details = details || null;
  throw error;
}

function string_(value, maxLength) {
  return String(value === undefined || value === null ? '' : value).replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, maxLength || 500);
}

function bool_(value) {
  return value === true || value === 1 || String(value).toLowerCase() === 'true';
}

function number_(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function id_(value, label) {
  const result = string_(value, 100);
  if (!/^[A-Za-z0-9_-]{8,100}$/.test(result)) apiError_('VALIDATION', (label || 'ID') + ' 값이 올바르지 않습니다.');
  return result;
}

function validDate_(value) {
  const date = string_(value, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) apiError_('VALIDATION', '날짜가 올바르지 않습니다.');
  return date;
}

function today_() {
  return formatDate_(new Date(), 'yyyy-MM-dd');
}

function formatDate_(date, pattern) {
  return Utilities.formatDate(date, APP.TIME_ZONE, pattern);
}

function formatKoreanDate_(date) {
  const parts = String(date).split('-');
  return parts.length === 3 ? Number(parts[0]) + '년 ' + Number(parts[1]) + '월 ' + Number(parts[2]) + '일' : String(date);
}

function compareKo_(a, b) {
  return String(a || '').localeCompare(String(b || ''), 'ko');
}

function staffKey_(department, name) {
  return String(department || '').trim().toLowerCase() + '\u0000' + String(name || '').trim().toLowerCase();
}

function safeFileName_(value) {
  return String(value || '').replace(/[\\/:*?"<>|\r\n]+/g, '_').replace(/\s+/g, ' ').trim().slice(0, 100) || '연수';
}

function randomToken_(length) {
  const seed = Utilities.getUuid() + Utilities.getUuid() + new Date().getTime() + Math.random();
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed, Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '').slice(0, length || 24);
}

function safeEqual_(left, right) {
  const a = String(left || '');
  const b = String(right || '');
  let mismatch = a.length ^ b.length;
  const length = Math.max(a.length, b.length);
  for (let index = 0; index < length; index += 1) mismatch |= (a.charCodeAt(index % Math.max(1, a.length)) || 0) ^ (b.charCodeAt(index % Math.max(1, b.length)) || 0);
  return mismatch === 0;
}

function normalizeFrontendUrl_(url) {
  const value = string_(url, 500).split('#')[0].replace(/\?.*$/, '');
  if (!/^https:\/\/[A-Za-z0-9.-]+(?::\d+)?(?:\/[^\s]*)?$/.test(value)) apiError_('VALIDATION', 'GitHub Pages 주소가 올바르지 않습니다.');
  return value;
}

function buildShareUrl_(baseUrl, token) {
  return baseUrl && token ? baseUrl + '#k=' + encodeURIComponent(token) : '';
}
