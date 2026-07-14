import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildShareUrl,
  groupStaffByDepartment,
  isPrivacyReady,
  localDuplicateKey,
  normalizeRosterRows,
  parseShareToken,
  safeFileName,
  splitNames,
  validateTraining
} from '../assets/core.js';

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

test('명단은 부서별로 묶고 이름순으로 정렬한다', () => {
  const groups = groupStaffByDepartment([
    { id: '2', department: '교무부', name: '최교사', active: true },
    { id: '1', department: '교무부', name: '김교사', active: true },
    { id: '3', department: '연구부', name: '박교사', active: false }
  ]);
  assert.deepEqual([...groups.keys()], ['교무부']);
  assert.deepEqual(groups.get('교무부').map(person => person.name), ['김교사', '최교사']);
});

test('개인정보 안내 필수값을 검사한다', () => {
  const valid = { schoolName: '학교', subtitle: '연수', privacyPurpose: '목적', privacyItems: '항목', privacyRetention: '삭제', privacyContact: '담당자' };
  assert.equal(isPrivacyReady(valid), true);
  assert.equal(isPrivacyReady({ ...valid, privacyContact: '' }), false);
});

test('연수 날짜와 시각을 검증한다', () => {
  assert.deepEqual(validateTraining({ title: '연수', date: '2026-07-14', startTime: '09:00', endTime: '10:00' }), []);
  assert.equal(validateTraining({ title: '', date: '', startTime: '11:00', endTime: '10:00' }).length, 3);
  assert.deepEqual(validateTraining({ title: '매일', daily: true, startTime: '', endTime: '' }), []);
});

test('이름 나누기와 파일명 안전화', () => {
  assert.deepEqual(splitNames('김하늘, 박서준\n김하늘'), ['김하늘', '박서준']);
  assert.equal(safeFileName('2026/연수:*?'), '2026_연수_');
  assert.equal(localDuplicateKey('t1', 's1', '2026-07-14'), 'training-sign:t1:s1:2026-07-14');
});
