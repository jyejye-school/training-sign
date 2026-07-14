import {
  buildShareUrl,
  formatKoreanDate,
  groupStaffByDepartment,
  isPrivacyReady,
  localDuplicateKey,
  normalizeRosterRows,
  parseShareToken,
  splitNames,
  todaySeoul,
  trainingTimeLabel,
  validateTraining
} from './core.js';

const $ = id => document.getElementById(id);
const config = window.TRAINING_SIGN_CONFIG || {};
const DEMO = new URLSearchParams(location.search).get('demo') === '1';
const API_URL = String(config.API_URL || '');
const shareToken = DEMO ? 'DEMO_TOKEN_1234567890123456' : parseShareToken(location.hash);
const baseUrl = `${location.origin}${location.pathname}`;

const state = {
  publicData: null,
  selectedTraining: null,
  selectedStaff: null,
  adminSession: '',
  adminData: null,
  records: [],
  strokes: [],
  drawing: false,
  currentStroke: null,
  demoAdminData: null
};

const demoData = {
  settings: {
    schoolName: '한빛고등학교',
    subtitle: '교직원 연수 참여 확인',
    notice: '연수 내용을 확인한 뒤 본인의 부서와 성명을 선택해 서명해 주세요.',
    brandColor: '#315c54',
    privacyPurpose: '교직원 연수 참여 여부 확인 및 서명등록부 작성',
    privacyItems: '부서, 성명, 서명 이미지, 서명 날짜와 시각',
    privacyRetention: 'PDF와 엑셀 보관 완료 후 시스템 원본을 삭제합니다.',
    privacyContact: '교무기획부 000-0000-0000'
  },
  trainings: [
    { id: 'demo-training-1', title: '2026 개인정보 보호 연수', target: '전 교직원', date: todaySeoul(), daily: false, startTime: '', endTime: '', active: true, sortOrder: 1 },
    { id: 'demo-training-2', title: '학교 안전교육', target: '전 교직원', date: todaySeoul(), daily: false, startTime: '09:00', endTime: '18:00', active: true, sortOrder: 2 }
  ],
  staff: [
    { id: 'staff-1', department: '교무기획부', name: '김하늘', active: true },
    { id: 'staff-2', department: '교무기획부', name: '박서준', active: true },
    { id: 'staff-3', department: '교육연구부', name: '이도윤', active: true },
    { id: 'staff-4', department: '교육연구부', name: '최지우', active: true },
    { id: 'staff-5', department: '행정실', name: '정민서', active: true }
  ],
  privacyReady: true,
  serverDate: todaySeoul()
};

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function setHidden(element, hidden) {
  element?.classList.toggle('hidden', Boolean(hidden));
}

function showToast(message, timeout = 2600) {
  const toast = $('toast');
  toast.textContent = message;
  setHidden(toast, false);
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => setHidden(toast, true), timeout);
}

function showStatus(message = '', isError = true) {
  const banner = $('statusBanner');
  banner.textContent = message;
  banner.style.background = isError ? '' : 'var(--brand-soft)';
  banner.style.color = isError ? '' : 'var(--brand-dark)';
  setHidden(banner, !message);
}

async function rpc(action, payload = {}, options = {}) {
  if (DEMO) return demoRpc(action, payload);
  if (!API_URL || API_URL.includes('__APPS_SCRIPT_WEB_APP_URL__')) {
    throw new Error('아직 Apps Script 주소가 연결되지 않았습니다. 관리자에게 알려 주세요.');
  }
  const body = { action, ...payload };
  if (options.admin !== false && state.adminSession && !body.sessionToken) body.sessionToken = state.adminSession;
  const response = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body),
    redirect: 'follow',
    cache: 'no-store'
  });
  if (!response.ok) throw new Error(`서버 응답 오류 (${response.status})`);
  const result = await response.json();
  if (!result.ok) {
    const error = new Error(result.error?.message || '요청을 처리하지 못했습니다.');
    error.code = result.error?.code || 'UNKNOWN';
    error.details = result.error?.details;
    throw error;
  }
  return result.data;
}

function demoRpc(action, payload) {
  if (!state.demoAdminData) {
    state.demoAdminData = {
      settings: { ...demoData.settings },
      trainings: demoData.trainings.map(item => ({ ...item })),
      staff: demoData.staff.map(item => ({ ...item })),
      exports: [],
      shareToken: shareToken,
      shareUrl: buildShareUrl(baseUrl, shareToken),
      stats: { staff: demoData.staff.length, trainings: demoData.trainings.length, signatures: 2 }
    };
  }
  if (action === 'get_public_data') return Promise.resolve(demoData);
  if (action === 'submit_signature') return Promise.resolve({ registeredAt: new Date().toISOString(), demo: true });
  if (action === 'get_setup_status') return Promise.resolve({ initialized: true, adminConfigured: true });
  if (action === 'admin_login') {
    if (payload.password !== 'demo-admin') return Promise.reject(Object.assign(new Error('데모 비밀번호는 demo-admin입니다.'), { code: 'BAD_PASSWORD' }));
    return Promise.resolve({ sessionToken: 'demo-session', expiresIn: 1800, adminData: state.demoAdminData });
  }
  if (action === 'get_admin_data') return Promise.resolve(state.demoAdminData);
  if (action === 'list_records') return Promise.resolve({ records: [
    { id: 'record-1', department: '교무기획부', name: '김하늘', signDate: todaySeoul(), signTime: '10:12:03', trainingId: 'demo-training-1' },
    { id: 'record-2', department: '교육연구부', name: '이도윤', signDate: todaySeoul(), signTime: '10:20:14', trainingId: 'demo-training-1' }
  ] });
  if (['logout', 'download_export_chunk'].includes(action)) return Promise.resolve({});
  return Promise.reject(Object.assign(new Error('데모에서는 변경 내용을 저장하지 않습니다.'), { code: 'DEMO_READ_ONLY' }));
}

function applySettings(settings) {
  const color = /^#[0-9a-f]{6}$/i.test(settings.brandColor || '') ? settings.brandColor : '#315c54';
  document.documentElement.style.setProperty('--brand', color);
  $('schoolName').textContent = settings.schoolName || config.APP_NAME || '학교 연수 전자서명';
  $('schoolSubtitle').textContent = settings.subtitle || '연수 참여 확인';
  document.title = `${settings.schoolName || '학교'} 연수 전자서명`;
  $('noticeText').textContent = settings.notice || '';
  setHidden($('noticePanel'), !settings.notice);
}

function showPanel(panelId) {
  ['loadingPanel', 'invalidPanel', 'trainingPanel', 'personPanel', 'signaturePanel', 'successPanel']
    .forEach(id => setHidden($(id), id !== panelId));
  const step = panelId === 'personPanel' ? 2 : panelId === 'signaturePanel' ? 3 : 1;
  document.querySelectorAll('#stepper li').forEach(item => {
    const value = Number(item.dataset.step);
    item.classList.toggle('active', value === step);
    item.classList.toggle('done', value < step);
  });
}

function renderTrainings() {
  const list = $('trainingList');
  const trainings = state.publicData?.trainings || [];
  list.replaceChildren();
  setHidden($('noTraining'), trainings.length > 0);
  trainings.forEach(training => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'choice-card';
    button.dataset.trainingId = training.id;
    button.innerHTML = `<span><strong>${escapeHtml(training.title)}</strong><span class="choice-meta"><span class="pill">${escapeHtml(training.target || '대상 안내 없음')}</span><small>${escapeHtml(trainingTimeLabel(training))}</small></span></span><b aria-hidden="true">›</b>`;
    button.addEventListener('click', () => selectTraining(training.id));
    list.append(button);
  });
}

function selectTraining(trainingId) {
  state.selectedTraining = state.publicData.trainings.find(item => item.id === trainingId) || null;
  state.selectedStaff = null;
  if (!state.selectedTraining) return;
  $('selectedTrainingLabel').textContent = `${state.selectedTraining.title} · ${trainingTimeLabel(state.selectedTraining)}`;
  renderDepartments();
  showPanel('personPanel');
}

function renderDepartments() {
  const groups = groupStaffByDepartment(state.publicData?.staff || []);
  const select = $('departmentSelect');
  select.innerHTML = '<option value="">부서를 선택하세요</option>';
  for (const department of groups.keys()) {
    const option = document.createElement('option');
    option.value = department;
    option.textContent = department;
    select.append(option);
  }
  $('staffSelect').innerHTML = '<option value="">먼저 부서를 선택하세요</option>';
  $('staffSelect').disabled = true;
  $('goToSignature').disabled = true;
}

function renderStaffForDepartment() {
  const department = $('departmentSelect').value;
  const people = (state.publicData?.staff || []).filter(person => person.department === department && person.active !== false);
  const select = $('staffSelect');
  select.innerHTML = '<option value="">성명을 선택하세요</option>';
  people.forEach(person => {
    const option = document.createElement('option');
    option.value = person.id;
    option.textContent = person.name;
    select.append(option);
  });
  select.disabled = !department;
  state.selectedStaff = null;
  $('goToSignature').disabled = true;
}

function goToSignature() {
  const person = state.publicData.staff.find(item => item.id === $('staffSelect').value);
  if (!person) return;
  state.selectedStaff = person;
  $('signerSummary').textContent = `${state.selectedTraining.title} · ${person.department} ${person.name}`;
  $('privacyConfirm').checked = false;
  clearSignature();
  showPanel('signaturePanel');
  requestAnimationFrame(resizeCanvas);
}

function clearSignature() {
  state.strokes = [];
  state.currentStroke = null;
  drawSignature();
}

function undoSignature() {
  state.strokes.pop();
  drawSignature();
}

function resizeCanvas() {
  const canvas = $('signatureCanvas');
  if (!canvas) return;
  const rect = canvas.getBoundingClientRect();
  const ratio = Math.min(window.devicePixelRatio || 1, 2);
  canvas.width = Math.max(1, Math.round(rect.width * ratio));
  canvas.height = Math.max(1, Math.round(rect.height * ratio));
  drawSignature();
}

function drawSignature(targetCanvas = $('signatureCanvas'), width = targetCanvas.width, height = targetCanvas.height) {
  if (!targetCanvas) return;
  const context = targetCanvas.getContext('2d');
  context.clearRect(0, 0, width, height);
  context.lineCap = 'round';
  context.lineJoin = 'round';
  context.strokeStyle = '#13201d';
  context.lineWidth = Math.max(2, width / 250);
  for (const stroke of state.strokes) {
    if (!stroke.length) continue;
    context.beginPath();
    context.moveTo(stroke[0].x * width, stroke[0].y * height);
    for (const point of stroke.slice(1)) context.lineTo(point.x * width, point.y * height);
    if (stroke.length === 1) context.lineTo(stroke[0].x * width + .01, stroke[0].y * height + .01);
    context.stroke();
  }
}

function canvasPoint(event) {
  const rect = $('signatureCanvas').getBoundingClientRect();
  return {
    x: Math.min(1, Math.max(0, (event.clientX - rect.left) / rect.width)),
    y: Math.min(1, Math.max(0, (event.clientY - rect.top) / rect.height))
  };
}

function startDrawing(event) {
  event.preventDefault();
  state.drawing = true;
  state.currentStroke = [canvasPoint(event)];
  state.strokes.push(state.currentStroke);
  $('signatureCanvas').setPointerCapture?.(event.pointerId);
  drawSignature();
}

function continueDrawing(event) {
  if (!state.drawing || !state.currentStroke) return;
  event.preventDefault();
  state.currentStroke.push(canvasPoint(event));
  drawSignature();
}

function stopDrawing() {
  state.drawing = false;
  state.currentStroke = null;
}

function signatureDataUrl() {
  const canvas = document.createElement('canvas');
  canvas.width = 600;
  canvas.height = 220;
  const context = canvas.getContext('2d');
  context.fillStyle = '#ffffff';
  context.fillRect(0, 0, canvas.width, canvas.height);
  drawSignature(canvas, canvas.width, canvas.height);
  return canvas.toDataURL('image/png');
}

async function submitSignature() {
  if (!state.selectedTraining || !state.selectedStaff) return;
  if (state.strokes.length === 0 || !state.strokes.some(stroke => stroke.length > 1)) {
    showToast('서명을 먼저 작성해 주세요.');
    return;
  }
  if (!$('privacyConfirm').checked) {
    showToast('개인정보 처리 안내를 확인해 주세요.');
    return;
  }
  const date = state.publicData.serverDate || todaySeoul();
  const duplicateKey = localDuplicateKey(state.selectedTraining.id, state.selectedStaff.id, date);
  if (localStorage.getItem(duplicateKey) && !confirm('이 기기에서 이미 서명한 기록이 있습니다. 그래도 서버에 확인할까요?')) return;
  const button = $('submitSignature');
  button.disabled = true;
  button.textContent = '등록 중…';
  try {
    const result = await rpc('submit_signature', {
      shareToken,
      trainingId: state.selectedTraining.id,
      staffId: state.selectedStaff.id,
      signatureData: signatureDataUrl()
    }, { admin: false });
    localStorage.setItem(duplicateKey, result.registeredAt || new Date().toISOString());
    $('successMessage').textContent = `${state.selectedStaff.department} ${state.selectedStaff.name}님의 참여 확인이 완료되었습니다.${result.demo ? ' 데모이므로 실제 저장되지는 않았습니다.' : ''}`;
    showPanel('successPanel');
  } catch (error) {
    showToast(error.message, 4200);
  } finally {
    button.disabled = false;
    button.textContent = '서명 제출';
  }
}

function renderPrivacy() {
  const settings = state.publicData?.settings || state.adminData?.settings || {};
  const fields = [
    ['수집 목적', settings.privacyPurpose],
    ['수집 항목', settings.privacyItems],
    ['보관·삭제', settings.privacyRetention],
    ['담당자', settings.privacyContact],
    ['이용 제한', '연수 참여 확인용이며, 본인 인증이 필요한 법적 전자서명에는 사용할 수 없습니다.']
  ];
  $('privacyDetails').innerHTML = fields.map(([title, value]) => `<dt>${escapeHtml(title)}</dt><dd>${escapeHtml(value || '관리자가 아직 입력하지 않았습니다.')}</dd>`).join('');
  $('privacyDialog').showModal();
}

function renderQr(container, url) {
  container.replaceChildren();
  if (!url || typeof window.qrcode !== 'function') {
    container.textContent = 'QR 생성 기능을 불러오지 못했습니다.';
    return;
  }
  const qr = window.qrcode(0, 'M');
  qr.addData(url, 'Byte');
  qr.make();
  container.innerHTML = qr.createSvgTag({ cellSize: 6, margin: 4, scalable: true });
}

function openShareDialog() {
  const url = DEMO ? buildShareUrl(baseUrl, shareToken) : location.href;
  $('shareUrl').value = url;
  renderQr($('qrCode'), url);
  $('shareDialog').showModal();
}

async function copyText(value, message = '복사했습니다.') {
  try {
    await navigator.clipboard.writeText(value);
    showToast(message);
  } catch {
    const input = document.createElement('textarea');
    input.value = value;
    document.body.append(input);
    input.select();
    document.execCommand('copy');
    input.remove();
    showToast(message);
  }
}

async function openAdminLogin() {
  $('adminLoginError').textContent = '';
  setHidden($('adminLoginError'), true);
  $('adminPassword').value = '';
  $('adminPasswordConfirm').value = '';
  $('setupCode').value = '';
  try {
    const status = await rpc('get_setup_status', {}, { admin: false });
    const setupRequired = !status.adminConfigured;
    $('adminLoginTitle').textContent = setupRequired ? '관리자 첫 설정' : '관리자 로그인';
    $('adminLoginSubmit').textContent = setupRequired ? '비밀번호 설정' : '로그인';
    setHidden($('setupCodeLabel'), !setupRequired);
    setHidden($('adminPasswordConfirmLabel'), !setupRequired);
    $('setupCode').required = setupRequired;
    $('adminPasswordConfirm').required = setupRequired;
    $('adminPassword').autocomplete = setupRequired ? 'new-password' : 'current-password';
    $('adminLoginForm').dataset.setupRequired = setupRequired ? '1' : '0';
    $('adminLoginDialog').showModal();
  } catch (error) {
    showToast(error.message, 4200);
  }
}

async function handleAdminLogin(event) {
  event.preventDefault();
  const setupRequired = event.currentTarget.dataset.setupRequired === '1';
  const password = $('adminPassword').value;
  const errorBox = $('adminLoginError');
  try {
    let data;
    if (setupRequired) {
      if (password !== $('adminPasswordConfirm').value) throw new Error('비밀번호 확인이 일치하지 않습니다.');
      data = await rpc('complete_setup', {
        setupCode: $('setupCode').value.trim(),
        password,
        frontendUrl: baseUrl
      }, { admin: false });
    } else {
      data = await rpc('admin_login', { password }, { admin: false });
    }
    state.adminSession = data.sessionToken;
    state.adminData = data.adminData;
    $('adminLoginDialog').close();
    renderAdmin();
    $('adminDialog').showModal();
  } catch (error) {
    errorBox.textContent = error.message;
    setHidden(errorBox, false);
  }
}

function switchAdminTab(tab) {
  document.querySelectorAll('#adminTabs button').forEach(button => button.classList.toggle('active', button.dataset.adminTab === tab));
  document.querySelectorAll('[data-admin-panel]').forEach(panel => panel.classList.toggle('active', panel.dataset.adminPanel === tab));
}

function trainingMeta(training) {
  return [training.target || '대상 미지정', trainingTimeLabel(training), training.active ? '활성' : '비활성'].join(' · ');
}

function renderTrainingAdmin() {
  const container = $('trainingAdminList');
  const trainings = state.adminData?.trainings || [];
  container.innerHTML = trainings.length ? trainings.map((training, index) => `
    <div class="admin-row" data-training-id="${escapeHtml(training.id)}">
      <div class="admin-row-main"><strong>${escapeHtml(training.title)}</strong><small>${escapeHtml(trainingMeta(training))}</small></div>
      <div class="row-actions">
        <button data-action="move-up" ${index === 0 ? 'disabled' : ''}>위</button>
        <button data-action="move-down" ${index === trainings.length - 1 ? 'disabled' : ''}>아래</button>
        <button data-action="edit-training">수정</button>
        <button data-action="delete-training" class="danger">삭제</button>
      </div>
    </div>`).join('') : '<div class="empty-state">등록된 연수가 없습니다.</div>';
}

function openTrainingForm(training = null) {
  $('trainingId').value = training?.id || '';
  $('trainingTitle').value = training?.title || '';
  $('trainingTarget').value = training?.target || '';
  $('trainingDate').value = training?.date && training.date !== '매일' ? training.date : todaySeoul();
  $('trainingDaily').checked = Boolean(training?.daily);
  $('trainingStart').value = training?.startTime || '';
  $('trainingEnd').value = training?.endTime || '';
  $('trainingActive').checked = training ? training.active !== false : true;
  setHidden($('trainingFormError'), true);
  setHidden($('trainingForm'), false);
  $('trainingTitle').focus();
}

async function saveTraining(event) {
  event.preventDefault();
  const training = {
    id: $('trainingId').value || undefined,
    title: $('trainingTitle').value.trim(),
    target: $('trainingTarget').value.trim(),
    date: $('trainingDate').value,
    daily: $('trainingDaily').checked,
    startTime: $('trainingStart').value,
    endTime: $('trainingEnd').value,
    active: $('trainingActive').checked
  };
  const errors = validateTraining(training);
  if (training.active && !isPrivacyReady(state.adminData.settings)) errors.push('기관 설정에서 개인정보 안내를 모두 입력해야 활성화할 수 있습니다.');
  if (errors.length) {
    $('trainingFormError').textContent = errors.join(' ');
    setHidden($('trainingFormError'), false);
    return;
  }
  try {
    await rpc('save_training', { training });
    await refreshAdminData();
    setHidden($('trainingForm'), true);
    showToast('연수를 저장했습니다.');
  } catch (error) { showToast(error.message, 4200); }
}

async function handleTrainingListClick(event) {
  const button = event.target.closest('button[data-action]');
  const row = event.target.closest('[data-training-id]');
  if (!button || !row) return;
  const training = state.adminData.trainings.find(item => item.id === row.dataset.trainingId);
  if (!training) return;
  try {
    if (button.dataset.action === 'edit-training') return openTrainingForm(training);
    if (button.dataset.action === 'delete-training') {
      if (!confirm(`'${training.title}' 연수를 삭제할까요? 기존 서명 기록은 남습니다.`)) return;
      await rpc('delete_training', { trainingId: training.id });
    } else {
      await rpc('move_training', { trainingId: training.id, direction: button.dataset.action === 'move-up' ? 'up' : 'down' });
    }
    await refreshAdminData();
  } catch (error) { showToast(error.message, 4200); }
}

function renderStaffAdmin() {
  const staff = state.adminData?.staff || [];
  const groups = groupStaffByDepartment(staff);
  const container = $('staffAdminList');
  const departmentOptions = [...groups.keys()];
  $('oldDepartment').innerHTML = '<option value="">기존 부서 선택</option>' + departmentOptions.map(value => `<option>${escapeHtml(value)}</option>`).join('');
  container.innerHTML = departmentOptions.length ? departmentOptions.map(department => `
    <div class="subcard"><h4>${escapeHtml(department)} <small>${groups.get(department).length}명</small></h4>
    ${groups.get(department).map(person => `<div class="admin-row" data-staff-id="${escapeHtml(person.id)}"><div class="admin-row-main"><strong>${escapeHtml(person.name)}</strong></div><div class="row-actions"><button data-action="edit-staff">수정</button><button data-action="delete-staff" class="danger">삭제</button></div></div>`).join('')}</div>`).join('') : '<div class="empty-state">등록된 구성원이 없습니다.</div>';
}

async function addStaff(event) {
  event.preventDefault();
  const department = $('staffDepartment').value.trim();
  const names = splitNames($('staffNames').value);
  if (!department || !names.length) return;
  try {
    const result = await rpc('save_staff_batch', { people: names.map(name => ({ department, name })) });
    $('staffNames').value = '';
    await refreshAdminData();
    showToast(`${result.added}명 등록, ${result.skipped}명 건너뜀`);
  } catch (error) { showToast(error.message, 4200); }
}

async function handleStaffListClick(event) {
  const button = event.target.closest('button[data-action]');
  const row = event.target.closest('[data-staff-id]');
  if (!button || !row) return;
  const person = state.adminData.staff.find(item => item.id === row.dataset.staffId);
  if (!person) return;
  try {
    if (button.dataset.action === 'edit-staff') {
      const department = prompt('부서', person.department);
      if (department === null) return;
      const name = prompt('성명', person.name);
      if (name === null) return;
      await rpc('update_staff', { person: { id: person.id, department: department.trim(), name: name.trim() } });
    } else {
      if (!confirm(`${person.department} ${person.name} 구성원을 삭제할까요? 기존 서명 기록은 남습니다.`)) return;
      await rpc('delete_staff', { staffId: person.id });
    }
    await refreshAdminData();
  } catch (error) { showToast(error.message, 4200); }
}

async function renameDepartment(event) {
  event.preventDefault();
  const oldDepartment = $('oldDepartment').value;
  const newDepartment = $('newDepartment').value.trim();
  if (!oldDepartment || !newDepartment) return;
  try {
    const result = await rpc('rename_department', { oldDepartment, newDepartment });
    $('newDepartment').value = '';
    await refreshAdminData();
    showToast(`${result.updated}명의 부서명을 변경했습니다.`);
  } catch (error) { showToast(error.message, 4200); }
}

function downloadRosterTemplate() {
  if (!window.XLSX) return showToast('엑셀 기능을 불러오지 못했습니다.');
  const sheet = XLSX.utils.aoa_to_sheet([['부서', '성명'], ['교무기획부', '홍길동']]);
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, '구성원명단');
  XLSX.writeFile(book, '구성원_등록_양식.xlsx');
}

async function importRosterFile(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  const status = $('rosterImportStatus');
  try {
    if (!window.XLSX) throw new Error('엑셀 기능을 불러오지 못했습니다.');
    status.textContent = '파일을 읽는 중…';
    const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const people = normalizeRosterRows(rows);
    if (!people.length) throw new Error('부서와 성명 열을 찾지 못했습니다. 양식을 확인해 주세요.');
    if (!confirm(`${people.length}명을 가져올까요? 이미 등록된 같은 부서·성명은 건너뜁니다.`)) return;
    const result = await rpc('save_staff_batch', { people });
    await refreshAdminData();
    status.textContent = `${result.added}명 등록, ${result.skipped}명 건너뜀`;
  } catch (error) {
    status.textContent = error.message;
  } finally {
    event.target.value = '';
  }
}

function fillSettingsForm() {
  const settings = state.adminData?.settings || {};
  $('settingsSchoolName').value = settings.schoolName || '';
  $('settingsSubtitle').value = settings.subtitle || '';
  $('settingsNotice').value = settings.notice || '';
  $('settingsBrandColor').value = /^#[0-9a-f]{6}$/i.test(settings.brandColor || '') ? settings.brandColor : '#315c54';
  $('settingsPrivacyPurpose').value = settings.privacyPurpose || '';
  $('settingsPrivacyItems').value = settings.privacyItems || '';
  $('settingsPrivacyRetention').value = settings.privacyRetention || '';
  $('settingsPrivacyContact').value = settings.privacyContact || '';
}

async function saveSettings(event) {
  event.preventDefault();
  const settings = {
    schoolName: $('settingsSchoolName').value.trim(),
    subtitle: $('settingsSubtitle').value.trim(),
    notice: $('settingsNotice').value.trim(),
    brandColor: $('settingsBrandColor').value,
    privacyPurpose: $('settingsPrivacyPurpose').value.trim(),
    privacyItems: $('settingsPrivacyItems').value.trim(),
    privacyRetention: $('settingsPrivacyRetention').value.trim(),
    privacyContact: $('settingsPrivacyContact').value.trim()
  };
  if (!isPrivacyReady(settings)) return showToast('개인정보 안내 항목을 모두 입력해 주세요.');
  try {
    await rpc('save_settings', { settings, frontendUrl: baseUrl });
    await refreshAdminData();
    applySettings(settings);
    showToast('기관 설정을 저장했습니다.');
  } catch (error) { showToast(error.message, 4200); }
}

function populateTrainingSelects() {
  const options = (state.adminData?.trainings || []).map(training => `<option value="${escapeHtml(training.id)}">${escapeHtml(training.title)}</option>`).join('');
  ['recordTraining', 'exportTraining'].forEach(id => { $(id).innerHTML = `<option value="">연수 선택</option>${options}`; });
  if (!$('recordDate').value) $('recordDate').value = todaySeoul();
  if (!$('exportDate').value) $('exportDate').value = todaySeoul();
}

async function loadRecords(event) {
  event?.preventDefault();
  const trainingId = $('recordTraining').value;
  const date = $('recordDate').value;
  if (!trainingId || !date) return;
  try {
    const data = await rpc('list_records', { trainingId, date });
    state.records = data.records || [];
    $('recordSummary').innerHTML = `<p class="selection-summary">서명 ${state.records.length}건</p>`;
    $('recordList').innerHTML = state.records.length ? state.records.map(record => `
      <div class="admin-row" data-record-id="${escapeHtml(record.id)}"><div class="admin-row-main"><strong>${escapeHtml(record.department)} ${escapeHtml(record.name)}</strong><small>${escapeHtml(record.signDate)} ${escapeHtml(record.signTime)}</small></div><div class="row-actions"><button class="danger" data-action="delete-record">기록 삭제</button></div></div>`).join('') : '<div class="empty-state">서명 기록이 없습니다.</div>';
  } catch (error) { showToast(error.message, 4200); }
}

async function handleRecordClick(event) {
  const button = event.target.closest('[data-action="delete-record"]');
  const row = event.target.closest('[data-record-id]');
  if (!button || !row) return;
  const record = state.records.find(item => item.id === row.dataset.recordId);
  if (!record || !confirm(`${record.department} ${record.name}의 서명 기록과 이미지 파일을 삭제할까요?`)) return;
  try {
    await rpc('delete_record', { recordId: record.id });
    await loadRecords();
    showToast('서명 기록을 삭제했습니다.');
  } catch (error) { showToast(error.message, 4200); }
}

function renderShareAdmin() {
  const url = state.adminData?.shareUrl || buildShareUrl(baseUrl, state.adminData?.shareToken || '');
  $('adminShareUrl').value = url;
  renderQr($('adminQrCode'), url);
}

async function rotateShareToken() {
  if (!confirm('공유 키를 교체하면 기존 링크와 QR은 즉시 사용할 수 없게 됩니다. 교체할까요?')) return;
  try {
    const data = await rpc('rotate_share_token', { frontendUrl: baseUrl });
    state.adminData.shareToken = data.shareToken;
    state.adminData.shareUrl = data.shareUrl;
    renderShareAdmin();
    showToast('공유 키를 교체했습니다. 새 링크를 안내해 주세요.', 4200);
  } catch (error) { showToast(error.message, 4200); }
}

async function changePassword(event) {
  event.preventDefault();
  try {
    await rpc('change_password', { currentPassword: $('currentPassword').value, newPassword: $('newPassword').value });
    $('currentPassword').value = '';
    $('newPassword').value = '';
    showToast('관리자 비밀번호를 변경했습니다.');
  } catch (error) { showToast(error.message, 4200); }
}

function exportStatusLabel(job) {
  if (job.status === 'complete') return '완료';
  if (job.status === 'failed') return `실패: ${job.error || '알 수 없는 오류'}`;
  if (job.status === 'expired') return '만료됨';
  return `생성 중 ${job.progress || 0}/${job.total || 0}`;
}

function renderExportJobs() {
  const jobs = state.adminData?.exports || [];
  $('exportJobList').innerHTML = jobs.length ? jobs.map(job => {
    const training = state.adminData.trainings.find(item => item.id === job.trainingId);
    return `<div class="admin-row" data-job-id="${escapeHtml(job.jobId)}">
      <div class="admin-row-main"><strong>${escapeHtml(training?.title || job.trainingTitle || '삭제된 연수')}</strong><small>${escapeHtml(job.date)} · ${escapeHtml(exportStatusLabel(job))}</small></div>
      <div class="row-actions">
        ${job.status === 'complete' ? '<button data-action="download-pdf">PDF</button><button data-action="download-xlsx">엑셀</button><button class="danger" data-action="purge-originals">원본 삭제</button>' : ''}
        ${job.status === 'processing' || job.status === 'queued' ? '<button data-action="resume-export">계속 만들기</button>' : ''}
      </div>
    </div>`;
  }).join('') : '<div class="empty-state">생성한 출력 파일이 없습니다.</div>';
}

async function startExport(event) {
  event.preventDefault();
  const payload = {
    trainingId: $('exportTraining').value,
    date: $('exportDate').value,
    columns: Number($('exportColumns').value),
    sort: $('exportSort').value,
    showRate: $('exportShowRate').checked
  };
  if (!payload.trainingId || !payload.date) return showToast('연수와 날짜를 선택해 주세요.');
  try {
    const job = await rpc('start_export', payload);
    setHidden($('exportProgress'), false);
    await runExportJob(job.jobId);
  } catch (error) { showToast(error.message, 5200); }
}

async function runExportJob(jobId) {
  const box = $('exportProgress');
  setHidden(box, false);
  let job;
  try {
    do {
      job = await rpc('continue_export', { jobId });
      const percent = job.total ? Math.round(job.progress / job.total * 100) : job.status === 'complete' ? 100 : 0;
      box.querySelector('progress').value = percent;
      box.querySelector('p').textContent = job.status === 'complete' ? 'PDF와 엑셀을 만들었습니다.' : `서명 이미지를 배치 처리하는 중입니다. ${job.progress}/${job.total}`;
    } while (job.status === 'processing' || job.status === 'queued');
    await refreshAdminData();
    if (job.status === 'complete') showToast('PDF와 엑셀 생성이 완료되었습니다.');
    else if (job.status === 'failed') throw new Error(job.error || '출력 파일 생성에 실패했습니다.');
  } catch (error) {
    box.querySelector('p').textContent = error.message;
    showToast(error.message, 5200);
  }
}

async function downloadExport(jobId, format) {
  let offset = 0;
  let total = null;
  let fileName = `연수_서명등록부.${format}`;
  let mimeType = format === 'pdf' ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const chunks = [];
  showToast('파일을 내려받는 중입니다…', 5000);
  do {
    const chunk = await rpc('download_export_chunk', { jobId, format, offset, chunkSize: 524288 });
    chunks.push(Uint8Array.from(atob(chunk.base64), character => character.charCodeAt(0)));
    offset = chunk.nextOffset;
    total = chunk.totalBytes;
    fileName = chunk.fileName || fileName;
    mimeType = chunk.mimeType || mimeType;
  } while (offset < total);
  const blob = new Blob(chunks, { type: mimeType });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

async function purgeOriginals(job) {
  const training = state.adminData.trainings.find(item => item.id === job.trainingId);
  const expected = training?.title || job.trainingTitle;
  const confirmation = prompt(`PDF와 엑셀 파일을 내려받아 보관했는지 확인해 주세요.\n원본을 삭제하려면 연수명을 그대로 입력하세요.\n\n${expected}`);
  if (confirmation !== expected) return showToast('연수명이 일치하지 않아 삭제하지 않았습니다.');
  try {
    const result = await rpc('purge_originals', { jobId: job.jobId, confirmation });
    await refreshAdminData();
    showToast(result.failed ? `${result.deleted}건 삭제, ${result.failed}건 실패. 다시 시도할 수 있습니다.` : `${result.deleted}건의 원본을 삭제했습니다.`, 5200);
  } catch (error) { showToast(error.message, 5200); }
}

async function handleExportJobClick(event) {
  const button = event.target.closest('button[data-action]');
  const row = event.target.closest('[data-job-id]');
  if (!button || !row) return;
  const job = state.adminData.exports.find(item => item.jobId === row.dataset.jobId);
  if (!job) return;
  try {
    if (button.dataset.action === 'download-pdf') await downloadExport(job.jobId, 'pdf');
    if (button.dataset.action === 'download-xlsx') await downloadExport(job.jobId, 'xlsx');
    if (button.dataset.action === 'resume-export') await runExportJob(job.jobId);
    if (button.dataset.action === 'purge-originals') await purgeOriginals(job);
  } catch (error) { showToast(error.message, 5200); }
}

function renderAdmin() {
  renderTrainingAdmin();
  renderStaffAdmin();
  fillSettingsForm();
  populateTrainingSelects();
  renderShareAdmin();
  renderExportJobs();
}

async function refreshAdminData() {
  state.adminData = await rpc('get_admin_data');
  renderAdmin();
}

async function logoutAdmin() {
  try { await rpc('logout'); } catch { /* Session may already be gone. */ }
  state.adminSession = '';
  state.adminData = null;
  $('adminDialog').close();
  showToast('로그아웃했습니다.');
}

async function initializePublicApp() {
  if (!shareToken) {
    setHidden($('loadingPanel'), true);
    setHidden($('invalidPanel'), false);
    return;
  }
  try {
    state.publicData = await rpc('get_public_data', { shareToken }, { admin: false });
    if (!state.publicData.privacyReady) throw Object.assign(new Error('관리자가 개인정보 처리 안내를 완료하지 않았습니다.'), { code: 'PRIVACY_NOT_READY' });
    applySettings(state.publicData.settings || {});
    renderTrainings();
    showPanel('trainingPanel');
    if (DEMO) showStatus('데모 화면입니다. 입력 내용은 실제로 저장되지 않습니다.', false);
  } catch (error) {
    $('invalidMessage').textContent = error.message;
    showPanel('invalidPanel');
  }
}

function bindEvents() {
  $('departmentSelect').addEventListener('change', renderStaffForDepartment);
  $('staffSelect').addEventListener('change', () => { $('goToSignature').disabled = !$('staffSelect').value; });
  $('goToSignature').addEventListener('click', goToSignature);
  $('backToTraining').addEventListener('click', () => showPanel('trainingPanel'));
  $('backToPerson').addEventListener('click', () => showPanel('personPanel'));
  $('submitSignature').addEventListener('click', submitSignature);
  $('clearSignature').addEventListener('click', clearSignature);
  $('undoSignature').addEventListener('click', undoSignature);
  $('signAnother').addEventListener('click', () => { state.selectedTraining = null; state.selectedStaff = null; showPanel('trainingPanel'); });
  $('signatureCanvas').addEventListener('pointerdown', startDrawing);
  $('signatureCanvas').addEventListener('pointermove', continueDrawing);
  $('signatureCanvas').addEventListener('pointerup', stopDrawing);
  $('signatureCanvas').addEventListener('pointercancel', stopDrawing);
  new ResizeObserver(resizeCanvas).observe($('signatureCanvas'));
  $('shareButton').addEventListener('click', openShareDialog);
  $('copyShareUrl').addEventListener('click', () => copyText($('shareUrl').value, '공유 링크를 복사했습니다.'));
  ['openPrivacy', 'footerPrivacy'].forEach(id => $(id).addEventListener('click', renderPrivacy));
  $('adminButton').addEventListener('click', openAdminLogin);
  $('adminLoginForm').addEventListener('submit', handleAdminLogin);
  $('adminTabs').addEventListener('click', event => { const button = event.target.closest('[data-admin-tab]'); if (button) switchAdminTab(button.dataset.adminTab); });
  $('closeAdmin').addEventListener('click', () => $('adminDialog').close());
  $('adminLogout').addEventListener('click', logoutAdmin);
  $('newTraining').addEventListener('click', () => openTrainingForm());
  $('cancelTraining').addEventListener('click', () => setHidden($('trainingForm'), true));
  $('trainingForm').addEventListener('submit', saveTraining);
  $('trainingAdminList').addEventListener('click', handleTrainingListClick);
  $('staffAddForm').addEventListener('submit', addStaff);
  $('staffAdminList').addEventListener('click', handleStaffListClick);
  $('renameDepartmentForm').addEventListener('submit', renameDepartment);
  $('downloadRosterTemplate').addEventListener('click', downloadRosterTemplate);
  $('rosterFile').addEventListener('change', importRosterFile);
  $('settingsForm').addEventListener('submit', saveSettings);
  $('recordFilterForm').addEventListener('submit', loadRecords);
  $('recordList').addEventListener('click', handleRecordClick);
  $('rotateShareToken').addEventListener('click', rotateShareToken);
  $('adminCopyShare').addEventListener('click', () => copyText($('adminShareUrl').value, '공유 링크를 복사했습니다.'));
  $('changePasswordForm').addEventListener('submit', changePassword);
  $('exportForm').addEventListener('submit', startExport);
  $('exportJobList').addEventListener('click', handleExportJobClick);
}

bindEvents();
initializePublicApp();
