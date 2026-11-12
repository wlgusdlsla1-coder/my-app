// 5회차 최종 확인: 사용자가 직접 만든 "수료내역 취합본" 안의 교육별 시트를 교육 파일처럼 쪼개서 우리 앱 흐름에 넣고,
// 사용자가 손으로 만든 "수료현황_종합" 시트와 같은 결과가 나오는지 비교한다.
// samples/user는 Git에 올리지 않으므로, 파일이 있을 때만 실행한다.

import { existsSync, readFileSync } from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { mergeCourses } from '../merge';
import { readCourseEntry } from '../pipeline';
import type { CourseFile } from '../types';

const FILE = path.resolve(__dirname, '..', '..', 'samples', 'user', '수료내역 취합본 - 샘플.xlsx');
const has = existsSync(FILE);
const NOT_COURSE_SHEETS = ['정리', '수료현황_종합'];

/** 시트 하나를 교육 파일 하나(.xlsx 바이트)로 만든다 — 실제로 교육 플랫폼 파일을 올리는 것과 같은 경로를 타게 한다 */
async function sheetToFileBytes(ws: ExcelJS.Worksheet): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  const copy = wb.addWorksheet('Sheet1');
  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      copy.getCell(rowNumber, colNumber).value = cell.value as ExcelJS.CellValue;
    });
  });
  const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

describe.skipIf(!has)('사용자 샘플 "수료내역 취합본" — 사용자가 손으로 만든 종합 시트와 비교', () => {
  it('교육별 시트 18개를 교육 파일처럼 읽어 합친 결과가 종합 시트와 같음', async () => {
    const src = new ExcelJS.Workbook();
    const buf = readFileSync(FILE);
    await src.xlsx.load(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer);

    // ① 교육별 시트 → 우리 앱의 읽기 흐름(열 자동 찾기 포함)
    const courses: CourseFile[] = [];
    for (const ws of src.worksheets) {
      if (NOT_COURSE_SHEETS.includes(ws.name)) continue;
      const entry = await readCourseEntry(ws.name, `${ws.name}.xlsx`, await sheetToFileBytes(ws));
      expect(entry.status, `${ws.name}: 열을 자동으로 찾아야 해요 (${entry.error ?? ''})`).toBe('ready');
      courses.push(entry.course!);
    }
    expect(courses).toHaveLength(18);
    const merged = mergeCourses(courses);

    // ② 사용자가 만든 종합 시트: 사번·성명·부서 + 교육별 수료/수강
    const total = src.getWorksheet('수료현황_종합')!;
    const header = (total.getRow(1).values as (string | undefined)[]).slice(1);
    const rows = new Map<string, Record<string, string>>();
    total.eachRow((row, n) => {
      if (n === 1) return;
      // 수식 칸(최종 수료여부)은 계산 결과 값을 쓴다
      const v = (row.values as unknown[]).slice(1).map((x) =>
        x && typeof x === 'object' && 'result' in x ? String((x as { result: unknown }).result) : String(x),
      );
      rows.set(v[0].trim(), Object.fromEntries(header.map((h, i) => [h!, v[i]])));
    });

    // ③ 같은 사람들이 나오고, 교육별 판정과 최종 판정이 같은지 비교
    expect(merged.rows.map((r) => r.empNo).sort()).toEqual([...rows.keys()].sort());
    const courseHeaders = header.slice(3).filter((h) => h !== '최종 수료여부') as string[];
    expect(new Set(merged.courseNames)).toEqual(new Set(courseHeaders));

    for (const mr of merged.rows) {
      const theirs = rows.get(mr.empNo)!;
      expect(theirs['성명']).toBe(mr.name);
      for (const c of courseHeaders) {
        const expected = theirs[c] === '수료' ? '수료' : theirs[c] === '수강' ? '미수료' : theirs[c];
        expect(mr.courses[c], `${mr.empNo} / ${c}`).toBe(expected);
      }
      // 사용자의 최종 판정은 수료/미수료 두 가지. 우리 판정은 "미수료(교육기록 미흡)"를 따로 구분할 뿐 수료 여부는 같아야 한다
      expect(mr.overall === '수료' ? '수료' : '미수료', `${mr.empNo} 최종`).toBe(theirs['최종 수료여부']);
    }

    // ④ 결과를 눈으로 볼 수 있게 출력
    console.log(
      merged.rows.map((r) => `${r.empNo} ${r.name} ${r.dept}: ${r.overall}`).join('\n'),
      '\n종합 시트 부서:',
      [...rows.entries()].map(([k, v]) => `${k}=${v['부서']}`).join(', '),
    );
  });
});
