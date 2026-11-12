// 설계 참고: DESIGN.md §4.4, §7, §8 — 열 찾기 사전, 색, 제한값

import type { Field } from './types';

/** 교육 파일 최대 개수 (PRD 5번). 직원 명단 파일은 받지 않는다 */
export const MAX_COURSE_FILES = 20;

/** 파일 1개당 최대 용량: 10MB */
export const MAX_FILE_BYTES = 10 * 1024 * 1024;

/** 받는 파일 확장자 (.xlsx만) */
export const ALLOWED_EXTENSION = '.xlsx';

/** 제목 줄을 찾을 때 위에서부터 살펴볼 줄 수 */
export const HEADER_SEARCH_ROWS = 10;

/**
 * 열 자동 찾기 사전.
 * 제목 칸의 공백을 지우고 소문자로 바꾼 뒤 비교한다. 앞에 있는 이름일수록 먼저 고른다.
 */
export const COLUMN_SYNONYMS: Record<Field, string[]> = {
  empNo: ['사번', '사원번호', '직원번호', '사원id', '직번', '개인번호', 'id'],
  name: ['이름', '성명', '직원명', '사원명', 'name'],
  dept: ['부서', '부서명', '소속', '소속부서', '조직', '팀'],
  status: ['수료여부', '이수여부', '수료상태', '이수상태', '수료결과', '결과', '상태'],
};

/**
 * 이 길이 이상인 이름은 "포함"만 돼도 같은 열로 본다 (예: "수료여부(Y/N)").
 * "id", "팀", "결과"처럼 짧은 이름은 다른 열과 헷갈리기 쉬워 똑같을 때만 인정한다.
 */
export const CONTAINS_MATCH_MIN_LENGTH = 3;

/** 교육 파일에 꼭 있어야 하는 열. 교육 파일이 직원 명단 역할도 하므로 이름·부서도 필요하다 */
export const REQUIRED_FIELDS: readonly Field[] = ['empNo', 'name', 'dept', 'status'];

/** 원본 수료 여부 값 (PRD 5번) */
export const RAW_COMPLETED = '수료';
export const RAW_NOT_COMPLETED = '수강';

/** 셀 배경색 (엑셀 ARGB 형식) — DESIGN.md §4.2 */
export const COLORS = {
  /** 미수료: 엑셀 기본 "나쁨" 색 */
  red: 'FFFFC7CE',
  /** 교육 기록 없음: 엑셀 기본 "보통" 색 */
  yellow: 'FFFFEB9C',
  /** 확인 필요 */
  gray: 'FFD9D9D9',
  /** 제목 줄 */
  header: 'FFF2F2F2',
} as const;

/** 부서 칸이 비어 있을 때 결과 엑셀에 적는 이름 (부서별 요약에서 한 묶음으로 세기 위해) */
export const NO_DEPT_LABEL = '(부서 없음)';

/** 결과 엑셀 시트 이름 */
export const SHEET_NAMES = { merged: '통합', summary: '요약' } as const;
