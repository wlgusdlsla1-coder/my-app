// 부서별 요약 자동 채점 (DESIGN.md §7.2, FR-11): 집계 규칙, 엑셀 수식(SUMPRODUCT), 특수문자 부서 이름

import { readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { buildWorkbook, workbookToBytes } from '../buildWorkbook';
import { NO_DEPT_LABEL } from '../constants';
import { detectColumns } from '../detectColumns';
import { deptLabel, mergeCourses, summarize, summarizeByDept } from '../merge';
import { extractCourse, readSheet } from '../readExcel';
import type { ColumnMapping, CourseFile, CourseRecord } from '../types';
import { courseNameFromFile } from '../validateFiles';

const SAMPLES = path.resolve(__dirname, '..', '..', 'samples');

const r = (empNo: string, raw: string, dept: string, name = '김민준'): CourseRecord => ({ empNo, name, dept, raw });
const course = (courseName: string, records: CourseRecord[]): CourseFile => ({ courseName, records });

describe('부서별 집계 규칙', () => {
  it('부서는 처음 나온 순서, 수료/미수료/확인 필요를 부서마다 셈', () => {
    const { rows } = mergeCourses([
      course('교육1', [
        r('A1', '수료', '인사팀'),
        r('A2', '수강', '재무팀'),
        r('A3', '수료', '인사팀'),
        r('A4', '이수', '재무팀'),
        r('A5', '수료', '개발팀'),
      ]),
      course('교육2', [r('A1', '수료', '인사팀'), r('A2', '수료', '재무팀'), r('A3', '수강', '인사팀'), r('A4', '수료', '재무팀')]),
    ]);
    expect(summarizeByDept(rows)).toEqual([
      { dept: '인사팀', total: 2, completed: 1, notCompleted: 1, needsCheck: 0 }, // A1 수료, A3 미수료
      { dept: '재무팀', total: 2, completed: 0, notCompleted: 1, needsCheck: 1 }, // A2 미수료, A4 확인 필요
      { dept: '개발팀', total: 1, completed: 0, notCompleted: 1, needsCheck: 0 }, // A5는 교육2 기록 없음 → 미수료(미흡)
    ]);
  });

  it('부서가 비어 있으면 "(부서 없음)", 앞뒤 공백은 지우고, 대소문자만 다른 부서는 하나로 묶음', () => {
    expect(deptLabel('  ')).toBe(NO_DEPT_LABEL);
    expect(deptLabel(' 인사팀 ')).toBe('인사팀');
    const { rows } = mergeCourses([
      course('교육1', [r('A1', '수료', 'HR'), r('A2', '수료', 'hr'), r('A3', '수료', ''), r('A4', '수료', '인사팀 '), r('A5', '수료', '인사팀')]),
    ]);
    expect(summarizeByDept(rows).map((d) => [d.dept, d.total])).toEqual([
      ['HR', 2],
      [NO_DEPT_LABEL, 1],
      ['인사팀', 2],
    ]);
  });

  it('부서별 합이 전체 요약과 같음', () => {
    const { rows } = mergeCourses([course('교육1', [r('A1', '수료', 'X'), r('A2', '수강', 'Y'), r('A3', '이수', 'Y')])]);
    const depts = summarizeByDept(rows);
    const s = summarize(rows);
    expect(depts.reduce((a, d) => a + d.total, 0)).toBe(s.targets);
    expect(depts.reduce((a, d) => a + d.completed, 0)).toBe(s.completed);
    expect(depts.reduce((a, d) => a + d.notCompleted, 0)).toBe(s.notCompleted);
    expect(depts.reduce((a, d) => a + d.needsCheck, 0)).toBe(s.needsCheck);
  });
});

// ── 엑셀 수식을 엑셀 대신 직접 계산 (SUMPRODUCT·SUM만 쓴다) ─────────────────
// 글자를 그대로 비교하는 수식만 쓰기 때문에 흉내 내기 쉽다. (COUNTIF의 와일드카드 규칙은 흉내가 실제 엑셀과 달라서 쓰지 않는다.)
// 실제 엑셀에서의 결과는 EXPORT_DIR 검사로 파일을 저장해 따로 확인했다.

const RANGE = String.raw`'통합'!\$[A-Z]+\$\d+:\$[A-Z]+\$\d+`;
// --(범위="글자")  또는  --(LEFT(범위,3)="글자")  한 덩어리
const TERM = new RegExp(String.raw`--\((?:LEFT\((${RANGE}),3\)|(${RANGE}))="((?:[^"]|"")*)"\)`, 'g');

function evaluateDeptTable(wb: ExcelJS.Workbook) {
  const merged = wb.getWorksheet('통합')!;
  const summary = wb.getWorksheet('요약')!;

  const column = (range: string): string[] => {
    const m = /^'통합'!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)$/.exec(range);
    if (!m) throw new Error(`알 수 없는 범위: ${range}`);
    const out: string[] = [];
    for (let row = Number(m[2]); row <= Number(m[4]); row++) out.push(String(merged.getCell(`${m[1]}${row}`).value ?? ''));
    return out;
  };

  const evalCell = (cell: ExcelJS.Cell): number => {
    const f = (cell.value as ExcelJS.CellFormulaValue).formula;
    let m: RegExpExecArray | null;
    if ((m = /^SUMPRODUCT\((.+)\)$/.exec(f))) {
      // 모르는 수식이 숨어 있지 않은지: 이해한 덩어리를 빼면 쉼표만 남아야 한다
      expect(m[1].replace(TERM, ''), `이해하지 못한 수식: ${f}`).toMatch(/^,*$/);
      let ok: boolean[] | null = null;
      for (const t of m[1].matchAll(TERM)) {
        const wanted = t[3].replace(/""/g, '"').toLowerCase(); // 엑셀의 "=" 비교는 대소문자를 구분하지 않는다
        const values = column((t[1] ?? t[2])!).map((v) => (t[1] ? v.slice(0, 3) : v).toLowerCase());
        const hit = values.map((v) => v === wanted);
        ok = ok ? ok.map((x, k) => x && hit[k]) : hit;
      }
      return ok!.filter(Boolean).length;
    }
    if ((m = /^SUM\(([A-Z])(\d+):([A-Z])(\d+)\)$/.exec(f))) {
      let sum = 0;
      for (let row = Number(m[2]); row <= Number(m[4]); row++) sum += evalCell(summary.getCell(`${m[1]}${row}`));
      return sum;
    }
    throw new Error(`알 수 없는 수식: ${f}`);
  };

  // 제목 줄("부서")을 찾아 그 아래를 읽는다
  let headRow = 0;
  summary.eachRow((row, n) => {
    if (row.getCell(1).value === '부서') headRow = n;
  });
  expect(headRow).toBeGreaterThan(0);
  const table: { dept: string; total: number; completed: number; notCompleted: number; needsCheck: number; stored: number[] }[] = [];
  for (let n = headRow + 1; n <= summary.rowCount; n++) {
    const label = summary.getCell(`A${n}`).value as string;
    const cells = ['B', 'C', 'D', 'E'].map((c) => summary.getCell(`${c}${n}`));
    const [total, completed, notCompleted, needsCheck] = cells.map(evalCell);
    // 저장된 결과값이 0이면 엑셀 파일에는 <v>0</v>으로 제대로 적혀 있지만 ExcelJS가 다시 읽을 때 빈 값(undefined)으로 돌려주므로 0으로 본다
    const stored = cells.map((c) => ((c.value as ExcelJS.CellFormulaValue).result as number | undefined) ?? 0);
    table.push({ dept: label, total, completed, notCompleted, needsCheck, stored });
  }
  return table;
}

async function roundTrip(courses: CourseFile[]) {
  const merged = mergeCourses(courses);
  const bytes = await workbookToBytes(buildWorkbook(merged));
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
  return { merged, wb };
}

describe('요약 시트의 부서별 표', () => {
  it('수식을 계산하면 집계와 같고, 저장된 값(result)도 같으며, 합계 줄이 있음', async () => {
    const { merged, wb } = await roundTrip([
      course('교육1', [r('A1', '수료', '인사팀'), r('A2', '수강', '재무팀'), r('A3', '수료', '인사팀'), r('A4', '이수', '재무팀')]),
      course('교육2', [r('A1', '수료', '인사팀'), r('A2', '수료', '재무팀'), r('A3', '수강', '인사팀')]),
    ]);
    const table = evaluateDeptTable(wb);
    const expected = summarizeByDept(merged.rows);

    expect(table.slice(0, -1).map(({ dept, total, completed, notCompleted, needsCheck }) => ({ dept, total, completed, notCompleted, needsCheck }))).toEqual(expected);
    for (const row of table) expect(row.stored).toEqual([row.total, row.completed, row.notCompleted, row.needsCheck]);

    const totalRow = table.at(-1)!;
    const s = summarize(merged.rows);
    expect(totalRow).toMatchObject({ dept: '합계', total: s.targets, completed: s.completed, notCompleted: s.notCompleted, needsCheck: s.needsCheck });
  });

  it('부서 이름에 * ? ~ " , = 가 있어도 글자 그대로 세어짐, 부서 없음·대소문자도 처리', async () => {
    const depts = ['개발*팀', '개발A팀', '기획?팀', '~특수', '"따옴표"', 'A,B팀', 'HR', 'hr', '', '=시작'];
    const { merged, wb } = await roundTrip([
      course('교육1', depts.map((d, i) => r(`A${i}`, i % 2 === 0 ? '수료' : '수강', d))),
    ]);
    const table = evaluateDeptTable(wb);
    const expected = summarizeByDept(merged.rows);
    expect(table.slice(0, -1).map((t) => [t.dept, t.total, t.completed, t.notCompleted, t.needsCheck])).toEqual(
      expected.map((e) => [e.dept, e.total, e.completed, e.notCompleted, e.needsCheck]),
    );
    // "개발*팀"은 "개발A팀"까지 끌어들이지 않아야 한다
    expect(table.find((t) => t.dept === '개발*팀')?.total).toBe(1);
    expect(table.find((t) => t.dept === '기획?팀')?.total).toBe(1);
    expect(table.find((t) => t.dept === '~특수')?.total).toBe(1);
    expect(table.find((t) => t.dept === NO_DEPT_LABEL)?.total).toBe(1);
    expect(table.find((t) => t.dept === 'HR')?.total).toBe(2);
    // 합계는 대상자 수와 같다
    expect(table.at(-1)!.total).toBe(depts.length);
  });

  // 실제 엑셀로 열어 확인하고 싶을 때: EXPORT_DIR=폴더경로 npm test  (설정하지 않으면 건너뜀)
  it.skipIf(!process.env.EXPORT_DIR)('특수문자 부서 이름 결과를 EXPORT_DIR 폴더에 저장 (실제 엑셀 확인용)', async () => {
    const depts = ['개발*팀', '개발A팀', '기획?팀', '~특수', '"따옴표"', 'A,B팀', 'HR', 'hr', '', '=시작', '인사팀', '인사팀'];
    const merged = mergeCourses([course('교육1', depts.map((d, i) => r(`A${i}`, i % 3 === 0 ? '수강' : '수료', d)))]);
    const bytes = await workbookToBytes(buildWorkbook(merged));
    const { writeFileSync } = await import('node:fs');
    writeFileSync(path.join(process.env.EXPORT_DIR!, '부서특수문자.xlsx'), bytes);
  });

  it('통합 시트에서 부서를 "(부서 없음)"으로 적어 부서별 표와 같은 이름을 씀', async () => {
    const { wb } = await roundTrip([course('교육1', [r('A1', '수료', '')])]);
    expect(wb.getWorksheet('통합')!.getCell('C2').value).toBe(NO_DEPT_LABEL);
  });

  it('기존 요약 표(B2~B8)는 그대로이고 부서별 표는 그 아래에 있음', async () => {
    const { wb } = await roundTrip([course('교육1', [r('A1', '수료', '인사팀')])]);
    const s = wb.getWorksheet('요약')!;
    expect(s.getCell('A2').value).toBe('전체 수료');
    expect(s.getCell('A8').value).toBe('일치 여부');
    expect(s.getCell('A10').value).toBe('부서별 요약');
    expect(s.getRow(11).values).toEqual([undefined, '부서', '인원', '수료', '미수료', '확인 필요']);
  });
});

describe('가짜 데이터 (samples/basic, samples/full)', () => {
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

  it('basic: 부서 4개, 부서별 합계가 전체 요약과 같음', async () => {
    const { merged, wb } = await roundTrip(await load(path.join(SAMPLES, 'basic'), { 산업안전보건교육: { empNo: 0, status: 3 } }));
    const table = evaluateDeptTable(wb);
    expect(table.map((t) => t.dept)).toEqual(['인사팀', '재무팀', '영업팀', '개발팀', '합계']);
    expect(table.at(-1)).toMatchObject({ total: 31, completed: 20, notCompleted: 9, needsCheck: 2 });
    expect(table.slice(0, -1).map(({ dept, total }) => [dept, total])).toEqual(summarizeByDept(merged.rows).map((d) => [d.dept, d.total]));
  });

  it('full: 부서 8개 × 500명, 수식 계산이 집계와 같음', async () => {
    const { merged, wb } = await roundTrip(await load(path.join(SAMPLES, 'full')));
    const table = evaluateDeptTable(wb);
    expect(table).toHaveLength(9);
    expect(table.slice(0, -1).map((t) => t.total)).toEqual(summarizeByDept(merged.rows).map((d) => d.total));
    expect(table.at(-1)).toMatchObject({ total: 500, completed: 258, notCompleted: 242, needsCheck: 0 });
  }, 60_000);
});
