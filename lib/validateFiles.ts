// 설계 참고: DESIGN.md §8 (E01·E02·E03·E06) — 파일을 읽기 전에 형식·용량·개수·이름 겹침 검사

import { ALLOWED_EXTENSION, MAX_COURSE_FILES, MAX_FILE_BYTES } from './constants';
import type { ErrorCode } from './types';

/** 검사에 필요한 파일 정보 (브라우저의 File 객체도 이 모양을 갖고 있다) */
export type FileInfo = { name: string; size: number };

/**
 * 교육 이름 = 파일 이름에서 ".xlsx"를 뺀 부분 (PLAN 작업 6, 사용자 답변: 파일 이름에서)
 * 예: "성희롱예방교육.xlsx" → "성희롱예방교육"
 */
export function courseNameFromFile(fileName: string): string {
  return fileName.replace(/\.xlsx$/i, '').trim();
}

/** 파일 하나 검사: 문제가 없으면 null, 있으면 오류 코드 */
export function checkFile(file: FileInfo): ErrorCode | null {
  if (!file.name.toLowerCase().endsWith(ALLOWED_EXTENSION)) return 'E01';
  if (file.size > MAX_FILE_BYTES) return 'E02';
  return null;
}

/** 교육 파일 여러 개를 목록에 더할 때의 검사 결과 */
export type CourseAddResult<T extends FileInfo> = {
  /** 목록에 넣을 파일 */
  accepted: T[];
  /** 넣지 못한 파일과 이유 */
  rejected: { file: T; code: ErrorCode }[];
};

/**
 * 이미 올린 교육 파일 목록(existing)에 새 파일들(incoming)을 더할 수 있는지 검사한다.
 * - E01/E02: 형식·용량
 * - E06: 교육 이름(파일 이름)이 이미 있는 것과 겹침
 * - E03: 합쳐서 20개를 넘는 파일
 */
export function checkCourseFiles<T extends FileInfo>(
  existing: readonly FileInfo[],
  incoming: readonly T[],
): CourseAddResult<T> {
  const names = new Set(existing.map((f) => courseNameFromFile(f.name)));
  const accepted: T[] = [];
  const rejected: { file: T; code: ErrorCode }[] = [];

  for (const file of incoming) {
    const basic = checkFile(file);
    if (basic) {
      rejected.push({ file, code: basic });
      continue;
    }
    const courseName = courseNameFromFile(file.name);
    if (names.has(courseName)) {
      rejected.push({ file, code: 'E06' });
      continue;
    }
    if (existing.length + accepted.length >= MAX_COURSE_FILES) {
      rejected.push({ file, code: 'E03' });
      continue;
    }
    names.add(courseName);
    accepted.push(file);
  }
  return { accepted, rejected };
}
