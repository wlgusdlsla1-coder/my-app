// 설계 참고: DESIGN.md §7 — 통합 시트(색칠)와 요약 시트(COUNTIF 수식)가 담긴 엑셀 파일 만들기

import ExcelJS from 'exceljs';
import { COLORS, SHEET_NAMES } from './constants';
import { summarize, type MergeResult } from './merge';
import type { CourseStatus, OverallStatus } from './types';

/** 교육 칸 판정 → 배경색 (DESIGN.md §4.2). 수료는 색 없음 */
const COURSE_FILL: Partial<Record<CourseStatus, string>> = {
  미수료: COLORS.red,
  '교육 기록 없음': COLORS.yellow,
  '확인 필요': COLORS.gray,
};

/** 전체 수료 여부 → 배경색 (DESIGN.md §4.3). 수료는 색 없음 */
const OVERALL_FILL: Partial<Record<OverallStatus, string>> = {
  미수료: COLORS.red,
  '미수료(교육기록 미흡)': COLORS.yellow,
  '확인 필요': COLORS.gray,
};

function fillOf(argb: string | undefined): ExcelJS.Fill | undefined {
  return argb ? { type: 'pattern', pattern: 'solid', fgColor: { argb } } : undefined;
}

/** 통합 시트 고정 열 수: 사번·이름·부서 */
const FIXED_COLUMNS = 3;

/**
 * 합친 결과로 엑셀 통합 문서를 만든다.
 * - 시트 1 "통합": 사번 → 이름 → 부서 → 교육1…N → 전체 수료 여부, 미수료·기록 없음·확인 필요 칸 색칠
 * - 시트 2 "요약": 수료/미수료/확인 필요/합계/대상자 수를 COUNTIF 수식으로 통합 시트와 연결
 */
export function buildWorkbook(result: MergeResult): ExcelJS.Workbook {
  const { courseNames, rows } = result;
  const wb = new ExcelJS.Workbook();

  // ── 시트 1: 통합 ──────────────────────────────────────────
  const merged = wb.addWorksheet(SHEET_NAMES.merged, { views: [{ state: 'frozen', xSplit: FIXED_COLUMNS, ySplit: 1 }] });
  const overallCol = FIXED_COLUMNS + courseNames.length + 1; // 전체 수료 여부 열 번호 (1부터)

  const header = merged.addRow(['사번', '이름', '부서', ...courseNames, '전체 수료 여부']);
  header.font = { bold: true };
  header.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  header.eachCell((cell) => {
    cell.fill = fillOf(COLORS.header)!;
    cell.border = { bottom: { style: 'thin' } };
  });

  for (const row of rows) {
    const excelRow = merged.addRow([
      row.empNo,
      row.name,
      row.dept,
      ...courseNames.map((c) => row.courses[c]),
      row.overall,
    ]);
    // 사번은 글자로 저장해서 앞자리 0이 사라지지 않게 한다
    excelRow.getCell(1).numFmt = '@';

    courseNames.forEach((c, i) => {
      const fill = fillOf(COURSE_FILL[row.courses[c]]);
      const cell = excelRow.getCell(FIXED_COLUMNS + 1 + i);
      cell.alignment = { horizontal: 'center' };
      if (fill) cell.fill = fill;
    });
    const overallCell = excelRow.getCell(overallCol);
    overallCell.alignment = { horizontal: 'center' };
    overallCell.font = { bold: true };
    const overallFill = fillOf(OVERALL_FILL[row.overall]);
    if (overallFill) overallCell.fill = overallFill;
  }

  // 열 너비
  merged.getColumn(1).width = 12;
  merged.getColumn(2).width = 12;
  merged.getColumn(3).width = 18;
  courseNames.forEach((c, i) => {
    merged.getColumn(FIXED_COLUMNS + 1 + i).width = Math.min(Math.max(c.length * 2 + 2, 14), 40);
  });
  merged.getColumn(overallCol).width = 22;
  merged.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: overallCol } };

  // ── 시트 2: 요약 ──────────────────────────────────────────
  const summary = wb.addWorksheet(SHEET_NAMES.summary);
  const lastRow = Math.max(rows.length + 1, 2); // 직원 줄 마지막 번호 (직원이 없어도 범위가 깨지지 않게)
  const overallLetter = merged.getColumn(overallCol).letter;
  const overallRange = `'${SHEET_NAMES.merged}'!$${overallLetter}$2:$${overallLetter}$${lastRow}`;
  const empNoRange = `'${SHEET_NAMES.merged}'!$A$2:$A$${lastRow}`;
  const s = summarize(rows);

  const head = summary.addRow(['구분', '인원']);
  head.font = { bold: true };
  head.eachCell((cell) => {
    cell.fill = fillOf(COLORS.header)!;
    cell.border = { bottom: { style: 'thin' } };
  });

  // 수식에는 계산 결과도 함께 적어, 엑셀이 다시 계산하기 전에도(미리보기 등) 숫자가 보이게 한다
  summary.addRow(['전체 수료', { formula: `COUNTIF(${overallRange},"수료")`, result: s.completed }]);
  summary.addRow(['미수료', { formula: `COUNTIF(${overallRange},"미수료*")`, result: s.notCompleted }]);
  summary.addRow(['└ 그중 교육기록 미흡', { formula: `COUNTIF(${overallRange},"미수료(교육기록 미흡)")`, result: s.missingRecord }]);
  summary.addRow(['확인 필요', { formula: `COUNTIF(${overallRange},"확인 필요")`, result: s.needsCheck }]);
  // 합계 = 전체 수료 + 미수료 + 확인 필요 (교육기록 미흡은 미수료에 이미 들어 있으므로 더하지 않는다)
  summary.addRow(['합계', { formula: 'B2+B3+B5', result: s.total }]);
  summary.addRow(['대상자 수', { formula: `COUNTA(${empNoRange})`, result: s.targets }]);
  summary.addRow(['일치 여부', { formula: 'IF(B6=B7,"일치","불일치")', result: s.total === s.targets ? '일치' : '불일치' }]);

  summary.getRow(6).font = { bold: true };
  summary.getColumn(1).width = 26;
  summary.getColumn(2).width = 12;
  summary.getColumn(2).alignment = { horizontal: 'right' };

  return wb;
}

/** 엑셀 파일 내용(바이트)으로 바꾼다 */
export async function workbookToBytes(wb: ExcelJS.Workbook): Promise<Uint8Array> {
  const buffer = await wb.xlsx.writeBuffer();
  return new Uint8Array(buffer);
}
