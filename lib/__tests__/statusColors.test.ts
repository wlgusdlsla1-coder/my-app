// 화면 미리보기와 결과 엑셀이 같이 쓰는 색 (DESIGN.md §4.2, §4.3)

import { describe, expect, it } from 'vitest';
import { COLORS } from '../constants';
import { COURSE_FILL, OVERALL_FILL, toCssColor } from '../statusColors';

describe('판정별 색', () => {
  it('교육 칸: 미수료 빨강, 교육 기록 없음 노랑, 확인 필요 회색, 수료는 색 없음', () => {
    expect(COURSE_FILL).toEqual({ 미수료: COLORS.red, '교육 기록 없음': COLORS.yellow, '확인 필요': COLORS.gray });
    expect(COURSE_FILL['수료']).toBeUndefined();
  });

  it('전체 수료 여부: 미수료 빨강, 미수료(교육기록 미흡) 노랑, 확인 필요 회색, 수료는 색 없음', () => {
    expect(OVERALL_FILL).toEqual({ 미수료: COLORS.red, '미수료(교육기록 미흡)': COLORS.yellow, '확인 필요': COLORS.gray });
    expect(OVERALL_FILL['수료']).toBeUndefined();
  });

  it('엑셀 색(ARGB)을 화면 색으로 바꿈', () => {
    expect(toCssColor(COLORS.red)).toBe('#FFC7CE');
    expect(toCssColor(COLORS.yellow)).toBe('#FFEB9C');
    expect(toCssColor(undefined)).toBeUndefined();
  });
});
