import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');
const index = read('index.html');
const app = read('assets/app.js');
const backend = read('apps-script/Code.gs');

new Function(backend);

const failures = [];
const expect = (condition, message) => { if (!condition) failures.push(message); };

expect(!/hirame|문의\s*@/i.test(index + app + backend), '원본 제작자 표기가 코드에 남아 있습니다.');
expect(!/supabase/i.test(index + app + backend), 'Supabase 의존성이 남아 있습니다.');
expect(!/setSharing\s*\(/.test(backend), '드라이브 파일 공개 공유 코드가 있습니다.');
expect(/frame-ancestors 'none'/.test(index), '클릭재킹 방지 CSP가 없습니다.');
expect(!/unsafe-inline|unsafe-eval/.test(index), 'CSP에 unsafe 설정이 있습니다.');
expect(/get_public_data/.test(backend) && /submit_signature/.test(backend), '공개 API가 누락되었습니다.');
expect(/requireAdminSession_/.test(backend), '관리자 세션 검증이 누락되었습니다.');
expect(/view:\s*'bootstrap'/.test(app), '관리자 로그인이 빠른 bootstrap 보기를 요청하지 않습니다.');
expect(/get_admin_section/.test(app + backend) && /function getAdminSection_/.test(backend), '관리자 탭별 지연 로딩 API가 누락되었습니다.');
expect(/id="adminRefresh"/.test(index) && /refreshActiveAdminSection/.test(app), '관리자 수동 새로고침 기능이 누락되었습니다.');
expect(/ADMIN_SYNC_MS\s*=\s*30000/.test(app), '관리자 화면의 30초 백그라운드 동기화가 누락되었습니다.');
expect(!/await\s+refreshAdminData\s*\(/.test(app), '단순 저장 뒤 전체 관리자 데이터를 다시 읽고 있습니다.');
expect(/REQUEST_CONTEXT_/.test(backend) && /context\.spreadsheet/.test(backend) && /context\.rows/.test(backend), 'Apps Script 요청 범위 시트·행 캐시가 누락되었습니다.');
const loginBody = backend.slice(backend.indexOf('function adminLogin_'), backend.indexOf('function createAdminLoginResult_'));
expect(!/cleanupStaleExportJobs/.test(loginBody), '관리자 로그인 중 오래된 출력 작업 정리가 실행됩니다.');
expect(/created\.length, SHEETS\.STAFF\.headers\.length\)[\s\S]*setValues/.test(backend), '구성원 일괄 등록이 범위 단위로 저장되지 않습니다.');
expect(/withAdminMutationLock_/.test(backend), '동시 관리자 수정 잠금이 누락되었습니다.');
expect(/\^\\d\{4\}\$/.test(backend), '숫자 4자리 관리자 비밀번호 검증이 누락되었습니다.');
expect(!/privacyContact|settingsPrivacyContact/.test(index + app + backend), '개인정보 담당자·연락처 항목이 남아 있습니다.');
expect(!/trainingTarget/.test(index + app), '연수 대상·내용 입력 항목이 남아 있습니다.');
expect(!/training\.target/.test(app), '연수 카드가 대상·내용 값을 사용하고 있습니다.');
expect(!/올바른 공유 링크가 아닙니다/.test(index + app), '공개 첫 화면에 잘못된 링크 오류 문구가 남아 있습니다.');
expect(!/서명 화면을 열 수 없습니다/.test(index + app), '연수가 없을 때 오류 중심 문구가 남아 있습니다.');
expect(/오늘 참여할 수 있는 연수가 없습니다/.test(index), '연수가 없을 때의 안내 문구가 없습니다.');
expect(/날짜·시간·활성 상태/.test(index), '연수가 없을 때 관리자 확인 안내가 없습니다.');
expect(/rel="icon"[^>]+favicon\.svg/.test(index), '브라우저 파비콘 연결이 없습니다.');
expect(/class="brand-mark-icon"/.test(index) && !/<div class="brand-mark"[^>]*>서명/.test(index), '왼쪽 위 브랜드 표시가 펜 아이콘으로 바뀌지 않았습니다.');
expect(/--mark-blue:\s*#2563eb/.test(read('assets/styles.css')), '왼쪽 위 펜 아이콘에 파란색이 고정되지 않았습니다.');
expect(/id="schoolDate"/.test(index) && /formatKoreanHeaderDate/.test(app), '메인 화면 날짜 표시가 없습니다.');
expect(/data-panel="trainingPanel"[\s\S]*overflow:\s*hidden/.test(read('assets/styles.css')), '첫 단계 한 화면 배치 규칙이 없습니다.');
expect(/LockService/.test(backend), '동시 제출 잠금이 누락되었습니다.');
expect(/setTrashed\(true\)/.test(backend), '원본 파일 삭제 처리가 누락되었습니다.');
expect(/function onOpen\(\)/.test(backend) && /🖊️ 전자서명 관리/.test(backend), '연결형 시트 관리 메뉴가 누락되었습니다.');
expect(/TEMPLATE_LOCK/.test(backend), '배포용 원본 초기화 잠금이 누락되었습니다.');
expect(/SpreadsheetApp\.getActiveSpreadsheet\(\)/.test(backend), '학교용 사본 초기화가 현재 연결 시트를 사용하지 않습니다.');
expect(!/SpreadsheetApp\.create\(APP\.DATA_FILE\)/.test(backend), '초기화가 별도 데이터 스프레드시트를 만들고 있습니다.');
expect(/SpreadsheetApp\.openById\(PropertiesService\.getScriptProperties\(\)\.getProperty\('SPREADSHEET_ID'\)\)/.test(backend), '웹앱 데이터 접근이 저장된 시트 ID를 사용하지 않습니다.');
expect(/hideDataSheets_/.test(backend) && /사용설명서/.test(backend), '사용설명서 표시·데이터 탭 숨김 처리가 누락되었습니다.');
expect(/setProperties\(secrets, false\)/.test(backend), '재초기화가 기존 Script Properties를 보존하지 않습니다.');
expect(/outputType/.test(index + app + backend) && /previewFileId/.test(backend), '선택 출력 형식 또는 비공개 미리보기 작업 정보가 누락되었습니다.');
expect(/pagenum=CENTER/.test(backend), 'PDF 페이지 번호가 가운데에 표시되도록 설정되지 않았습니다.');
expect(/base \+ 1\)\.setWrap\(true\)/.test(backend), '긴 부서명이 출력에서 잘리지 않도록 줄바꿈 처리되지 않았습니다.');
expect(/recordPrintOpened_/.test(backend) && /canPurgeExport_/.test(backend), '인쇄 기록 또는 출력별 원본 삭제 차단이 누락되었습니다.');
expect(/frame-src 'self' blob:/.test(index), '비공개 Blob PDF 미리보기 CSP가 누락되었습니다.');
expect(fs.existsSync(path.join(root, 'vendor/qrcode.js')), 'QR 라이브러리가 없습니다.');
expect(fs.existsSync(path.join(root, 'vendor/xlsx.full.min.js')), '엑셀 라이브러리가 없습니다.');
expect(read('vendor/xlsx.full.min.js').includes('0.20.3'), 'SheetJS가 고정 버전 0.20.3이 아닙니다.');

if (failures.length) {
  console.error(failures.map(message => `- ${message}`).join('\n'));
  process.exit(1);
}

console.log('정적 보안·구성 검사 통과');
