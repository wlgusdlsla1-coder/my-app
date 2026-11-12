// 설계 참고: DESIGN.md §4.2, §4.3 — 교육 칸 판정과 전체 수료 여부 판정 (PRD 5번 규칙 그대로)

import { RAW_COMPLETED, RAW_NOT_COMPLETED } from './constants';
import type { CourseRecord, CourseStatus, OverallStatus } from './types';

/**
 * 한 사람의 한 교육 칸을 판정한다.
 * @param records 그 교육 파일에서 이 사번으로 찾은 줄들
 *
 * - 0줄 → "교육 기록 없음"
 * - 2줄 이상 → "확인 필요" (D-1: 재수강 등, 운영자가 직접 판단)
 * - 원본 값 "수료" → "수료"
 * - 원본 값 "수강" → "미수료"
 * - 그 밖의 값 → "확인 필요" (추측하지 않는다)
 */
export function judgeCourseCell(records: readonly CourseRecord[]): CourseStatus {
  if (records.length === 0) return '교육 기록 없음';
  if (records.length > 1) return '확인 필요';
  const raw = records[0].raw.trim();
  if (raw === RAW_COMPLETED) return '수료';
  if (raw === RAW_NOT_COMPLETED) return '미수료';
  return '확인 필요';
}

/**
 * 교육 칸들을 보고 전체 수료 여부를 정한다. 위에서부터 먼저 맞는 것 (DESIGN.md §4.3)
 * 1. "교육 기록 없음"이 하나라도 있음 → "미수료(교육기록 미흡)"
 * 2. "미수료"가 하나라도 있음 → "미수료"
 * 3. "확인 필요"가 하나라도 있음 → "확인 필요" (D-2)
 * 4. 모두 "수료" → "수료"
 */
export function judgeOverall(statuses: readonly CourseStatus[]): OverallStatus {
  if (statuses.includes('교육 기록 없음')) return '미수료(교육기록 미흡)';
  if (statuses.includes('미수료')) return '미수료';
  if (statuses.includes('확인 필요')) return '확인 필요';
  return '수료';
}
