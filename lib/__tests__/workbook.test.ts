// 3회차 자동 채점: 결과 엑셀을 만들고 다시 읽어서 열 순서·색칠·수식·앞자리 0을 확인 (SC-3·SC-4·SC-5)

import { readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { buildWorkbook, workbookToBytes } from '../buildWorkbook';
import { COLORS } from '../constants';
import { detectColumns } from '../detectColumns';
import { resultFileName } from '../download';
import { mergeCourses } from '../merge';
import { extractCourse, readSheet } from '../readExcel';
import type { ColumnMapping, CourseFile, CourseRecord } from '../types';
import { courseNameFromFile } from '../validateFiles';

const SAMPLES = path.resolve(__dirname, '..', '..', 'samples');

const r = (empNo: string, raw: string, name = '김민준', dept = '인사팀'): CourseRecord => ({ empNo, name, dept, raw });
const course = (courseName: string, records: CourseRecord[]): CourseFile => ({ courseName, records });

/** 만든 엑셀을 파일로 저장했다가 다시 읽은 것처럼 되돌린다 */
async function roundTrip(courses: CourseFile[]) {
  const merged = mergeCourses(courses);
  const bytes = await workbookToBytes(buildWorkbook(merged));
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
  return { merged, wb };
}

const fillOf = (cell: ExcelJS.Cell) => (cell.fill as ExcelJS.FillPattern | undefined)?.fgColor?.argb;

describe('통합 시트', () => {
  it('시트 2개(통합, 요약), 열 순서는 사번→이름→부서→교육들→전체 수료 여부', async () => {
    const { wb } = await roundTrip([
      course('교육1', [r('A001', '수료')]),
      course('교육2', [r('A001', '수료')]),
    ]);
    expect(wb.worksheets.map((w) => w.name)).toEqual(['통합', '요약']);
    const header = wb.getWorksheet('통합')!.getRow(1).values as string[];
    expect(header.slice(1)).toEqual(['사번', '이름', '부서', '교육1', '교육2', '전체 수료 여부']);
  });

  it('칸 색: 미수료 빨강, 교육 기록 없음 노랑, 확인 필요 회색, 수료는 색 없음', async () => {
    const { wb } = await roundTrip([
      course('교육1', [r('A1', '수료'), r('A2', '수강'), r('A3', '이수'), r('A4', '수료')]),
      course('교육2', [r('A1', '수료'), r('A2', '수료'), r('A3', '수료')]),
    ]);
    const ws = wb.getWorksheet('통합')!;
    // 줄: 2=A1 3=A2 4=A3 5=A4 / 열: 4=교육1 5=교육2 6=전체
    expect(ws.getCell('D2').value).toBe('수료');
    expect(fillOf(ws.getCell('D2'))).toBeUndefined();
    expect(fillOf(ws.getCell('F2'))).toBeUndefined(); // 전체 수료

    expect(ws.getCell('D3').value).toBe('미수료');
    expect(fillOf(ws.getCell('D3'))).toBe(COLORS.red);
    expect(ws.getCell('F3').value).toBe('미수료');
    expect(fillOf(ws.getCell('F3'))).toBe(COLORS.red);

    expect(ws.getCell('D4').value).toBe('확인 필요');
    expect(fillOf(ws.getCell('D4'))).toBe(COLORS.gray);
    expect(fillOf(ws.getCell('F4'))).toBe(COLORS.gray);

    expect(ws.getCell('E5').value).toBe('교육 기록 없음');
    expect(fillOf(ws.getCell('E5'))).toBe(COLORS.yellow);
    expect(ws.getCell('F5').value).toBe('미수료(교육기록 미흡)');
    expect(fillOf(ws.getCell('F5'))).toBe(COLORS.yellow);
  });

  it('사번은 글자로 저장되어 앞자리 0이 유지됨', async () => {
    const { wb } = await roundTrip([course('교육1', [r('00123', '수료'), r('123', '수료')])]);
    const ws = wb.getWorksheet('통합')!;
    expect(ws.getCell('A2').value).toBe('00123');
    expect(ws.getCell('A3').value).toBe('123');
    expect(ws.getCell('A2').numFmt).toBe('@');
  });

  it('틀 고정과 필터가 켜져 있음', async () => {
    const { wb } = await roundTrip([course('교육1', [r('A1', '수료')])]);
    const ws = wb.getWorksheet('통합')!;
    expect(ws.views[0]).toMatchObject({ state: 'frozen', xSplit: 3, ySplit: 1 });
    expect(ws.autoFilter).toBeTruthy();
  });
});

/**
 * 요약 시트 수식을 엑셀 대신 직접 계산해 본다 (COUNTIF·COUNTA·셀 더하기·IF만 쓰므로 이것으로 충분).
 * 통합 시트 값만 보고 계산하므로, 수식이 통합 시트와 제대로 연결됐는지 확인할 수 있다.
 */
function evaluateSummary(wb: ExcelJS.Workbook): Record<string, string | number> {
  const merged = wb.getWorksheet('통합')!;
  const summary = wb.getWorksheet('요약')!;
  const cellsIn = (range: string): (string | number | null)[] => {
    const m = /^'통합'!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)$/.exec(range);
    if (!m) throw new Error(`알 수 없는 범위: ${range}`);
    const out: (string | number | null)[] = [];
    for (let row = Number(m[2]); row <= Number(m[4]); row++) {
      out.push(merged.getCell(`${m[1]}${row}`).value as string | number | null);
    }
    return out;
  };
  const values: Record<string, number | string> = {};
  const evalCell = (addr: string): number | string => {
    if (addr in values) return values[addr];
    const f = (summary.getCell(addr).value as ExcelJS.CellFormulaValue).formula;
    let v: number | string;
    let m: RegExpExecArray | null;
    if ((m = /^COUNTIF\((.+),"(.*)"\)$/.exec(f))) {
      const crit = m[2];
      const test = crit.endsWith('*') ? (x: string) => x.startsWith(crit.slice(0, -1)) : (x: string) => x === crit;
      v = cellsIn(m[1]).filter((x) => typeof x === 'string' && test(x)).length;
    } else if ((m = /^COUNTA\((.+)\)$/.exec(f))) {
      v = cellsIn(m[1]).filter((x) => x !== null && x !== '').length;
    } else if ((m = /^([A-Z]\d+)\+([A-Z]\d+)\+([A-Z]\d+)$/.exec(f))) {
      v = [m[1], m[2], m[3]].reduce((sum, a) => sum + (evalCell(a) as number), 0);
    } else if ((m = /^IF\(([A-Z]\d+)=([A-Z]\d+),"(.*)","(.*)"\)$/.exec(f))) {
      v = evalCell(m[1]) === evalCell(m[2]) ? m[3] : m[4];
    } else {
      throw new Error(`알 수 없는 수식: ${f}`);
    }
    values[addr] = v;
    return v;
  };
  const labels: Record<string, string | number> = {};
  for (let row = 2; row <= 8; row++) labels[summary.getCell(`A${row}`).value as string] = evalCell(`B${row}`);
  return labels;
}

describe('요약 시트', () => {
  it('수식이 통합 시트와 연결되어 있고, 계산 결과가 저장된 값과 같음 (SC-5)', async () => {
    const { merged, wb } = await roundTrip([
      course('교육1', [r('A1', '수료'), r('A2', '수강'), r('A3', '이수'), r('A4', '수료')]),
      course('교육2', [r('A1', '수료'), r('A2', '수료'), r('A3', '수료')]),
    ]);
    const summary = wb.getWorksheet('요약')!;
    expect(summary.getCell('B2').value).toMatchObject({ formula: expect.stringContaining("COUNTIF('통합'!$F$2:$F$5") });

    const calc = evaluateSummary(wb);
    expect(calc).toEqual({
      '전체 수료': 1,
      미수료: 2,
      '└ 그중 교육기록 미흡': 1,
      '확인 필요': 1,
      합계: 4,
      '대상자 수': 4,
      '일치 여부': '일치',
    });
    // 엑셀이 다시 계산하기 전에 보이는 값(result)도 같아야 한다
    for (let row = 2; row <= 8; row++) {
      const cell = summary.getCell(`B${row}`).value as ExcelJS.CellFormulaValue;
      expect(cell.result).toBe(calc[summary.getCell(`A${row}`).value as string]);
    }
    expect(merged.rows).toHaveLength(4);
  });

  it('통합 시트 값을 바꾸면 요약 숫자가 따라 바뀜 (수식으로 연결된 증거)', async () => {
    const { wb } = await roundTrip([course('교육1', [r('A1', '수강'), r('A2', '수료')])]);
    expect(evaluateSummary(wb)['전체 수료']).toBe(1);
    wb.getWorksheet('통합')!.getCell('E2').value = '수료'; // A1의 전체 수료 여부를 수정
    const after = evaluateSummary(wb);
    expect(after['전체 수료']).toBe(2);
    expect(after['미수료']).toBe(0);
  });

  it('통합 시트에서 전체 수료 여부 칸을 비우면 "불일치"로 잡아 줌', async () => {
    const { wb } = await roundTrip([course('교육1', [r('A1', '수강'), r('A2', '수료')])]);
    wb.getWorksheet('통합')!.getCell('E2').value = null;
    expect(evaluateSummary(wb)['일치 여부']).toBe('불일치');
  });
});

describe('파일 이름', () => {
  it('의무교육_수료현황_YYYYMMDD.xlsx', () => {
    expect(resultFileName(new Date(2026, 10, 12))).toBe('의무교육_수료현황_20261112.xlsx');
    expect(resultFileName(new Date(2026, 0, 5))).toBe('의무교육_수료현황_20260105.xlsx');
  });
});

describe('가짜 데이터 전체 흐름 → 엑셀 (samples/basic, samples/full)', () => {
  async function load(dir: string, overrides: Record<string, ColumnMapping> = {}): Promise<CourseFile[]> {
    const files = readdirSync(dir).filter((f) => f.endsWith('.xlsx')).sort();
    const out: CourseFile[] = [];
    for (const f of files) {
      const buf = readFileSync(path.join(dir, f));
      const sheet = await readSheet(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer);
      const name = courseNameFromFile(f);
      out.push(extractCourse(sheet, { ...detectColumns(sheet.headers).mapping, ...overrides[name] }, name));
    }
    return out;
  }

  it('basic: 요약 시트 숫자가 정답표와 같음', async () => {
    const { wb } = await roundTrip(await load(path.join(SAMPLES, 'basic'), { 산업안전보건교육: { empNo: 0, status: 3 } }));
    expect(evaluateSummary(wb)).toEqual({
      '전체 수료': 20, 미수료: 9, '└ 그중 교육기록 미흡': 3, '확인 필요': 2, 합계: 31, '대상자 수': 31, '일치 여부': '일치',
    });
    const ws = wb.getWorksheet('통합')!;
    expect(ws.rowCount).toBe(32); // 제목 1줄 + 대상자 31명
    expect(ws.getCell('A32').value).toBe('Z999');
  });

  // 결과 엑셀을 눈으로 확인하고 싶을 때: EXPORT_DIR=폴더경로 npm test  (설정하지 않으면 건너뜀)
  it.skipIf(!process.env.EXPORT_DIR)('basic 결과를 EXPORT_DIR 폴더에 저장 (눈으로 확인용)', async () => {
    const merged = mergeCourses(await load(path.join(SAMPLES, 'basic'), { 산업안전보건교육: { empNo: 0, status: 3 } }));
    const bytes = await workbookToBytes(buildWorkbook(merged));
    const { writeFileSync } = await import('node:fs');
    writeFileSync(path.join(process.env.EXPORT_DIR!, resultFileName()), bytes);
  });

  it('full: 500명 × 20개를 엑셀로 만들고 숫자가 정답표와 같음, 걸린 시간 기록', async () => {
    const courses = await load(path.join(SAMPLES, 'full'));
    const start = performance.now();
    const { wb } = await roundTrip(courses);
    const ms = performance.now() - start;
    expect(evaluateSummary(wb)).toMatchObject({ '전체 수료': 258, 미수료: 242, '확인 필요': 0, 합계: 500, '일치 여부': '일치' });
    expect(wb.getWorksheet('통합')!.columnCount).toBe(24); // 사번·이름·부서 + 교육 20 + 전체 = 24
    expect(ms).toBeLessThan(30_000);
    console.log(`samples/full 합치기+엑셀 만들기+다시 읽기: ${Math.round(ms)}ms`);
  }, 60_000);
});
