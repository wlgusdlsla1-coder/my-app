// 설계 참고: DESIGN.md §2.2, §8 — 화면에서 쓰는 "파일 하나를 읽어 목록 칸으로 만들기"와 오류 메시지
// 화면 코드(app/)와 분리해 두어서 자동 채점으로 확인할 수 있다.

import { detectColumns, isMappingComplete } from './detectColumns';
import { errorMessage, FIELD_LABELS } from './messages';
import { extractCourse, readSheet } from './readExcel';
import type { ColumnMapping, CourseFile, RawSheet } from './types';
import { courseNameFromFile } from './validateFiles';

/** 화면 목록에 보이는 파일 한 칸 */
export type FileEntry = {
  id: string;
  fileName: string;
  /** 파일 이름에서 확장자를 뺀 부분 = 교육 이름 */
  courseName: string;
  /**
   * reading: 읽는 중 / ready: 열이 모두 정해져 합칠 수 있음
   * needsColumns: 열을 못 찾아 운영자가 골라야 함 / error: 열 수 없는 파일
   */
  status: 'reading' | 'ready' | 'needsColumns' | 'error';
  /** ready일 때만: 사번·이름·부서·수료 여부 네 열만 뽑아낸 기록 */
  course?: CourseFile;
  /** needsColumns일 때만: 열을 고를 때까지 잠깐 들고 있는 표. 고르고 나면 버린다 */
  sheet?: RawSheet;
  /** needsColumns일 때만: 자동으로 찾은 열 (못 찾은 열은 빠져 있음) */
  mapping?: ColumnMapping;
  /** error일 때 화면에 보일 메시지 */
  error?: string;
};

/** 엑셀 파일 내용(바이트)을 읽어 목록 칸을 만든다. 열을 다 찾으면 ready, 못 찾으면 needsColumns, 못 열면 error */
export async function readCourseEntry(id: string, fileName: string, data: ArrayBuffer): Promise<FileEntry> {
  const base = { id, fileName, courseName: courseNameFromFile(fileName) };
  let sheet: RawSheet;
  try {
    sheet = await readSheet(data);
  } catch {
    return { ...base, status: 'error', error: errorMessage('E04', fileName) };
  }
  const { mapping, missing } = detectColumns(sheet.headers);
  if (missing.length === 0) {
    return { ...base, status: 'ready', course: extractCourse(sheet, mapping, base.courseName) };
  }
  return {
    ...base,
    status: 'needsColumns',
    sheet,
    mapping,
    error: errorMessage('E05', fileName, missing.map((f) => FIELD_LABELS[f]).join('·')),
  };
}

/** 운영자가 고른 열을 적용한다. 열이 다 정해졌으면 ready로 바꾸고 표는 버린다 */
export function applyMapping(entry: FileEntry, mapping: ColumnMapping): FileEntry {
  if (!entry.sheet || !isMappingComplete(mapping)) return entry;
  return {
    id: entry.id,
    fileName: entry.fileName,
    courseName: entry.courseName,
    status: 'ready',
    course: extractCourse(entry.sheet, mapping, entry.courseName),
  };
}

/** [정리하기]를 누를 수 있는 상태인지: 1개 이상이고 전부 ready */
export function canProcess(entries: readonly FileEntry[]): boolean {
  return entries.length > 0 && entries.every((e) => e.status === 'ready');
}
