// 2회차 자동 채점: 교육 칸 판정 · 전체 수료 판정 (DESIGN.md §10.2 1~5, 8~11번)

import { describe, expect, it } from 'vitest';
import { judgeCourseCell, judgeOverall } from '../judge';
import type { CourseRecord } from '../types';

const rec = (raw: string): CourseRecord => ({ empNo: 'A001', name: '김민준', dept: '인사팀', raw });

describe('교육 칸 판정 (DESIGN §4.2)', () => {
  it('1. 원본 값 "수료" → 수료', () => {
    expect(judgeCourseCell([rec('수료')])).toBe('수료');
  });

  it('2. 원본 값 " 수강 " (공백 포함) → 미수료', () => {
    expect(judgeCourseCell([rec(' 수강 ')])).toBe('미수료');
  });

  it('3. 원본 값 "이수" → 확인 필요 (추측하지 않음)', () => {
    expect(judgeCourseCell([rec('이수')])).toBe('확인 필요');
    expect(judgeCourseCell([rec('')])).toBe('확인 필요');
  });

  it('4. 교육 파일에 사번 없음 → 교육 기록 없음', () => {
    expect(judgeCourseCell([])).toBe('교육 기록 없음');
  });

  it('5. 같은 교육 파일에 사번 2줄 → 확인 필요 (D-1)', () => {
    expect(judgeCourseCell([rec('수강'), rec('수료')])).toBe('확인 필요');
    expect(judgeCourseCell([rec('수료'), rec('수료')])).toBe('확인 필요');
  });
});

describe('전체 수료 판정 (DESIGN §4.3)', () => {
  it('8. 교육 3개 모두 수료 → 수료', () => {
    expect(judgeOverall(['수료', '수료', '수료'])).toBe('수료');
  });

  it('9. 수료·미수료·기록 없음 섞임 → 미수료(교육기록 미흡)', () => {
    expect(judgeOverall(['수료', '미수료', '교육 기록 없음'])).toBe('미수료(교육기록 미흡)');
  });

  it('10. 수료·미수료 섞임 → 미수료', () => {
    expect(judgeOverall(['수료', '미수료'])).toBe('미수료');
  });

  it('11. 수료·확인 필요 섞임 → 확인 필요', () => {
    expect(judgeOverall(['수료', '확인 필요'])).toBe('확인 필요');
  });

  it('순서: 미수료가 확인 필요보다 먼저, 기록 없음이 가장 먼저', () => {
    expect(judgeOverall(['확인 필요', '미수료'])).toBe('미수료');
    expect(judgeOverall(['확인 필요', '교육 기록 없음'])).toBe('미수료(교육기록 미흡)');
  });
});
