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
expect(/\^\\d\{4\}\$/.test(backend), '숫자 4자리 관리자 비밀번호 검증이 누락되었습니다.');
expect(/LockService/.test(backend), '동시 제출 잠금이 누락되었습니다.');
expect(/setTrashed\(true\)/.test(backend), '원본 파일 삭제 처리가 누락되었습니다.');
expect(/function onOpen\(\)/.test(backend) && /🖊️ 전자서명 관리/.test(backend), '연결형 시트 관리 메뉴가 누락되었습니다.');
expect(/TEMPLATE_LOCK/.test(backend), '배포용 원본 초기화 잠금이 누락되었습니다.');
expect(/SpreadsheetApp\.getActiveSpreadsheet\(\)/.test(backend), '학교용 사본 초기화가 현재 연결 시트를 사용하지 않습니다.');
expect(!/SpreadsheetApp\.create\(APP\.DATA_FILE\)/.test(backend), '초기화가 별도 데이터 스프레드시트를 만들고 있습니다.');
expect(/SpreadsheetApp\.openById\(PropertiesService\.getScriptProperties\(\)\.getProperty\('SPREADSHEET_ID'\)\)/.test(backend), '웹앱 데이터 접근이 저장된 시트 ID를 사용하지 않습니다.');
expect(/hideDataSheets_/.test(backend) && /사용설명서/.test(backend), '사용설명서 표시·데이터 탭 숨김 처리가 누락되었습니다.');
expect(/setProperties\(secrets, false\)/.test(backend), '재초기화가 기존 Script Properties를 보존하지 않습니다.');
expect(fs.existsSync(path.join(root, 'vendor/qrcode.js')), 'QR 라이브러리가 없습니다.');
expect(fs.existsSync(path.join(root, 'vendor/xlsx.full.min.js')), '엑셀 라이브러리가 없습니다.');
expect(read('vendor/xlsx.full.min.js').includes('0.20.3'), 'SheetJS가 고정 버전 0.20.3이 아닙니다.');

if (failures.length) {
  console.error(failures.map(message => `- ${message}`).join('\n'));
  process.exit(1);
}

console.log('정적 보안·구성 검사 통과');
