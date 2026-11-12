// 설계 참고: DESIGN.md §2.2 ②, §4.4 — 엑셀 첫 번째 시트를 글자 표로 읽고, 필요한 열만 뽑아내기
// ※ 모든 처리는 브라우저 안에서 일어나며, 읽은 내용은 어디로도 보내지 않는다 (PRD 7번)

import ExcelJS from 'exceljs';
import { HEADER_SEARCH_ROWS } from './constants';
import { findColumn } from './detectColumns';
import type { ColumnMapping, CourseFile, RawSheet } from './types';

/** 엑셀 파일을 열 수 없을 때 (암호, 손상 등) — E04 */
export class ExcelReadError extends Error {
  readonly code = 'E04' as const;
  constructor(message: string) {
    super(message);
    this.name = 'ExcelReadError';
  }
}

/**
 * 엑셀 칸 값을 글자로 바꾼다.
 * 숫자, 수식 결과, 서식 있는 글자, 링크, 날짜 등 여러 모양을 모두 글자로 통일한다.
 */
export function cellToText(value: ExcelJS.CellValue): string {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') return value;
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === 'object') {
    if ('richText' in value) return value.richText.map((r) => r.text).join('');
    if ('result' in value) return cellToText(value.result as ExcelJS.CellValue); // 수식은 계산 결과를 쓴다
    if ('text' in value) return cellToText(value.text as ExcelJS.CellValue); // 링크
    if ('error' in value) return '';
  }
  return String(value);
}

/** 사번 다듬기: 글자로 바꾸고 앞뒤 공백만 지운다. 앞자리 0은 그대로 둔다 (PRD 5번) */
export function normalizeEmpNo(value: string): string {
  return value.trim();
}

/**
 * 엑셀 파일(바이트)을 읽어 첫 번째 시트를 글자 표로 돌려준다.
 * 제목 줄은 위에서 10줄 안에서 "사번" 열이 있는 줄로 찾고, 없으면 첫 줄을 제목으로 본다.
 */
export async function readSheet(data: ArrayBuffer): Promise<RawSheet> {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(data);
  } catch {
    throw new ExcelReadError('엑셀 파일을 열 수 없어요. 암호가 걸려 있거나 손상됐을 수 있어요.');
  }
  const ws = wb.worksheets[0];
  if (!ws) throw new ExcelReadError('엑셀 파일에 시트가 없어요.');

  // 시트 전체를 글자 표로 바꾼다 (빈 줄도 자리를 지켜 줄 번호가 어긋나지 않게)
  const table: string[][] = [];
  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const cells: string[] = [];
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      cells[colNumber - 1] = cellToText(cell.value).trim();
    });
    table[rowNumber - 1] = Array.from(cells, (c) => c ?? '');
  });
  for (let i = 0; i < table.length; i++) table[i] ??= [];

  // 제목 줄 찾기
  let headerIdx = 0;
  const limit = Math.min(HEADER_SEARCH_ROWS, table.length);
  for (let i = 0; i < limit; i++) {
    if (findColumn(table[i], 'empNo') !== undefined) {
      headerIdx = i;
      break;
    }
  }

  const headers = table[headerIdx] ?? [];
  const rows = table.slice(headerIdx + 1).filter((r) => r.some((c) => c !== ''));
  return { headers, rows };
}

/**
 * 정해진 열 위치로 교육 파일 한 개의 기록을 뽑아낸다. 사번이 빈 줄은 건너뛴다.
 * 사번·이름·부서·수료 여부 네 열만 남기고 나머지 열은 버린다 (PRD 7번)
 */
export function extractCourse(sheet: RawSheet, mapping: ColumnMapping, courseName: string): CourseFile {
  const { empNo, name, dept, status } = mapping;
  if (empNo === undefined || name === undefined || dept === undefined || status === undefined) {
    throw new Error('교육 파일의 사번·이름·부서·수료 여부 열이 모두 정해져야 해요.');
  }
  const records = sheet.rows
    .map((r) => ({
      empNo: normalizeEmpNo(r[empNo] ?? ''),
      name: (r[name] ?? '').trim(),
      dept: (r[dept] ?? '').trim(),
      raw: (r[status] ?? '').trim(),
    }))
    .filter((rec) => rec.empNo !== '');
  return { courseName, records };
}
