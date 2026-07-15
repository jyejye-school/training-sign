import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildShareUrl,
  formatKoreanHeaderDate,
  groupStaffByDepartment,
  isPrivacyReady,
  isValidAdminPassword,
  localDuplicateKey,
  normalizeNameEntryText,
  normalizeRosterRows,
  parseShareToken,
  safeFileName,
  sortRecords,
  splitNames,
  validateTraining
} from '../assets/core.js';

test('메인 화면 날짜에 요일을 함께 표시한다', () => {
  assert.equal(formatKoreanHeaderDate('2026-07-15'), '7월 15일 (수)');
  assert.equal(formatKoreanHeaderDate('잘못된 날짜'), '');
});

test('공유 키는 올바른 형식만 읽는다', () => {
  assert.equal(parseShareToken('#k=Abcd_12345678901234567890'), 'Abcd_12345678901234567890');
  assert.equal(parseShareToken('#admin'), '');
  assert.equal(parseShareToken('#k=short'), '');
  assert.equal(parseShareToken('#k=%3Cscript%3E'), '');
});

test('공유 주소는 기존 쿼리와 해시를 제거한다', () => {
  assert.equal(buildShareUrl('https://school.github.io/training-sign/?demo=1#old', 'abc_123'), 'https://school.github.io/training-sign/#k=abc_123');
});

test('엑셀 명단을 정규화하고 중복과 빈 행을 제거한다', () => {
  const result = normalizeRosterRows([
    ['부서', '성명'],
    ['교무부', '홍길동'],
    ['교무부', '홍길동'],
    ['', '빈부서'],
    ['연구부', '김하늘']
  ]);
  assert.deepEqual(result, [
    { department: '교무부', name: '홍길동' },
    { department: '연구부', name: '김하늘' }
  ]);
});

test('명단은 부서별로 묶고 등록 순서를 유지한다', () => {
  const groups = groupStaffByDepartment([
    { id: '2', department: '교무부', name: '최교사', active: true, sortOrder: 2 },
    { id: '1', department: '교무부', name: '김교사', active: true, sortOrder: 1 },
    { id: '3', department: '연구부', name: '박교사', active: false, sortOrder: 3 }
  ]);
  assert.deepEqual([...groups.keys()], ['교무부']);
  assert.deepEqual(groups.get('교무부').map(person => person.name), ['김교사', '최교사']);
});

test('출력 정렬은 등록순을 기본으로 하고 부서순과 이름순을 지원한다', () => {
  const records = [
    { department: '연구부', name: '김교사', sortOrder: 3 },
    { department: '교무부', name: '최교사', sortOrder: 1 },
    { department: '교무부', name: '박교사', sortOrder: 2 }
  ];
  assert.deepEqual(sortRecords(records).map(person => person.name), ['최교사', '박교사', '김교사']);
  assert.deepEqual(sortRecords(records, 'department').map(person => person.name), ['박교사', '최교사', '김교사']);
  assert.deepEqual(sortRecords(records, 'name').map(person => person.name), ['김교사', '박교사', '최교사']);
});

test('개인정보 안내 필수값을 검사한다', () => {
  const valid = { schoolName: '학교', subtitle: '연수', privacyPurpose: '목적', privacyItems: '항목', privacyRetention: '삭제' };
  assert.equal(isPrivacyReady(valid), true);
  assert.equal(isPrivacyReady({ ...valid, privacyRetention: '' }), false);
});

test('관리자 비밀번호는 숫자 4자리 또는 문자·숫자 포함 10자 이상을 허용한다', () => {
  assert.equal(isValidAdminPassword('1234'), true);
  assert.equal(isValidAdminPassword('0000'), true);
  assert.equal(isValidAdminPassword('학교연수2026비밀번호'), true);
  assert.equal(isValidAdminPassword('12345'), false);
  assert.equal(isValidAdminPassword('abcdefghij'), false);
  assert.equal(isValidAdminPassword('abcd'), false);
});

test('연수 날짜와 시각을 검증한다', () => {
  assert.deepEqual(validateTraining({ title: '연수', date: '2026-07-14', startTime: '09:00', endTime: '10:00' }), []);
  assert.equal(validateTraining({ title: '', date: '', startTime: '11:00', endTime: '10:00' }).length, 3);
  assert.deepEqual(validateTraining({ title: '매일', daily: true, startTime: '', endTime: '' }), []);
});

test('이름 나누기와 파일명 안전화', () => {
  assert.equal(normalizeNameEntryText('홍길동 홍수박'), '홍길동\n홍수박');
  assert.equal(normalizeNameEntryText('홍길동\t홍수박,김하늘;박서준\r\n최지우'), '홍길동\n홍수박\n김하늘\n박서준\n최지우');
  assert.deepEqual(splitNames('김하늘, 박서준\n\n김하늘  홍수박'), ['김하늘', '박서준', '홍수박']);
  assert.equal(safeFileName('2026/연수:*?'), '2026_연수_');
  assert.equal(localDuplicateKey('t1', 's1', '2026-07-14'), 'training-sign:t1:s1:2026-07-14');
});
