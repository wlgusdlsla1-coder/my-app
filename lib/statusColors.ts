// 설계 참고: DESIGN.md §4.2, §4.3 — 판정 결과별 배경색. 결과 엑셀과 화면 미리보기가 같은 색을 쓰도록 한 곳에 둔다.
// 이 파일은 엑셀 도구(ExcelJS)를 불러오지 않으므로 화면 부품이 가져다 써도 첫 화면이 무거워지지 않는다.

import { COLORS } from './constants';
import type { CourseStatus, OverallStatus } from './types';

/** 교육 칸 판정 → 배경색 (엑셀 ARGB). 수료는 색 없음 */
export const COURSE_FILL: Partial<Record<CourseStatus, string>> = {
  미수료: COLORS.red,
  '교육 기록 없음': COLORS.yellow,
  '확인 필요': COLORS.gray,
};

/** 전체 수료 여부 → 배경색 (엑셀 ARGB). 수료는 색 없음 */
export const OVERALL_FILL: Partial<Record<OverallStatus, string>> = {
  미수료: COLORS.red,
  '미수료(교육기록 미흡)': COLORS.yellow,
  '확인 필요': COLORS.gray,
};

/** 엑셀 색(ARGB, 예 FFFFC7CE)을 화면 색(#FFC7CE)으로 바꾼다 */
export function toCssColor(argb: string | undefined): string | undefined {
  return argb ? `#${argb.slice(2)}` : undefined;
}
