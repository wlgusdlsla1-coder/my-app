// 설계 참고: DESIGN.md §8 — 화면에 보이는 열 이름과 오류 메시지
// 엑셀 도구(ExcelJS)를 불러오지 않는 파일이라, 화면 부품이 이 파일을 가져다 써도 첫 화면이 무거워지지 않는다.

import { MAX_COURSE_FILES, MAX_FILE_BYTES } from './constants';
import type { ErrorCode, Field } from './types';

/** 열 이름 (화면 글자) */
export const FIELD_LABELS: Record<Field, string> = {
  empNo: '사번',
  name: '이름',
  dept: '부서',
  status: '수료 여부',
};

/** 오류 메시지 (DESIGN.md §8). fileName은 문제가 된 파일 이름, extra는 오류별 덧붙임 */
export function errorMessage(code: ErrorCode, fileName = '', extra = ''): string {
  switch (code) {
    case 'E01':
      return `${fileName} 파일은 .xlsx 형식이 아니에요. 엑셀에서 .xlsx로 저장한 뒤 올려 주세요.`;
    case 'E02':
      return `${fileName} 파일이 ${MAX_FILE_BYTES / 1024 / 1024}MB를 넘어요.`;
    case 'E03':
      return `교육 파일은 ${MAX_COURSE_FILES}개까지 올릴 수 있어요. ${fileName} 파일은 넣지 못했어요.`;
    case 'E04':
      return `${fileName} 파일을 열 수 없어요. 암호가 걸려 있거나 손상됐을 수 있어요.`;
    case 'E05':
      return `${fileName} 파일에서 ${extra} 열을 찾지 못했어요. 직접 골라 주세요.`;
    case 'E06':
      return `같은 이름의 교육 파일이 이미 있어요: ${fileName}`;
    case 'E08':
      return '합계가 대상자 수와 달라요. 결과를 확인해 주세요.';
  }
}
