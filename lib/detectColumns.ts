// 설계 참고: DESIGN.md §4.4 — 사번·이름·부서·수료 여부 열을 비슷한 이름으로 자동으로 찾기

import { COLUMN_SYNONYMS, CONTAINS_MATCH_MIN_LENGTH, REQUIRED_FIELDS } from './constants';
import type { ColumnMapping, Field, FileKind } from './types';

/** 제목 칸 비교용으로 다듬기: 모든 공백 제거 + 소문자 */
export function normalizeHeader(text: string): string {
  return text.replace(/\s+/g, '').toLowerCase();
}

/**
 * 제목 줄에서 한 종류의 열을 찾는다. 못 찾으면 undefined.
 * 1단계: 사전의 이름과 똑같은 칸 (사전 앞쪽 이름 우선)
 * 2단계: 사전의 3글자 이상 이름을 "포함"하는 칸 (예: "수료여부(Y/N)")
 */
export function findColumn(headers: readonly string[], field: Field): number | undefined {
  const normalized = headers.map(normalizeHeader);
  const synonyms = COLUMN_SYNONYMS[field];

  for (const word of synonyms) {
    const idx = normalized.indexOf(word);
    if (idx !== -1) return idx;
  }
  for (const word of synonyms) {
    if (word.length < CONTAINS_MATCH_MIN_LENGTH) continue;
    const idx = normalized.findIndex((h) => h.includes(word));
    if (idx !== -1) return idx;
  }
  return undefined;
}

/** 이 파일 종류에서 찾아볼 열 목록 (꼭 필요한 열 + 있으면 쓰는 열) */
const FIELDS_TO_LOOK: Record<FileKind, Field[]> = {
  roster: ['empNo', 'name', 'dept'],
  // 교육 파일의 이름 열은 있으면 이름 불일치 확인(D-3)에 쓴다
  course: ['empNo', 'status', 'name'],
};

export type DetectResult = {
  /** 찾은 열 위치 */
  mapping: ColumnMapping;
  /** 꼭 필요한데 못 찾은 열 → 열 고르기 창에서 운영자가 고른다 (E05) */
  missing: Field[];
};

/** 파일 종류에 맞게 열을 자동으로 찾는다 */
export function detectColumns(headers: readonly string[], kind: FileKind): DetectResult {
  const mapping: ColumnMapping = {};
  for (const field of FIELDS_TO_LOOK[kind]) {
    const idx = findColumn(headers, field);
    if (idx !== undefined) mapping[field] = idx;
  }
  const missing = REQUIRED_FIELDS[kind].filter((f) => mapping[f] === undefined);
  return { mapping, missing };
}

/** 운영자가 고른 열까지 합쳐서 꼭 필요한 열이 모두 정해졌는지 */
export function isMappingComplete(mapping: ColumnMapping, kind: FileKind): boolean {
  return REQUIRED_FIELDS[kind].every((f) => mapping[f] !== undefined);
}
