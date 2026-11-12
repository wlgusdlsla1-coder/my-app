// 공개 전 보안 점검에서 나온 항목의 자동 채점 (DESIGN.md §9)

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import nextConfig from '../../next.config';
import { buildWorkbook, workbookToBytes } from '../buildWorkbook';
import { MAX_FORMULA_TEXT_LENGTH } from '../constants';
import { mergeCourses, summarize } from '../merge';
import type { CourseFile, CourseRecord } from '../types';

const r = (empNo: string, raw: string, dept = '인사팀', name = '김민준'): CourseRecord => ({ empNo, name, dept, raw });
const course = (courseName: string, records: CourseRecord[]): CourseFile => ({ courseName, records });

async function roundTrip(courses: CourseFile[]) {
  const merged = mergeCourses(courses);
  const bytes = await workbookToBytes(buildWorkbook(merged));
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
  return { merged, wb };
}

describe('보안 응답 헤더 (next.config.ts)', () => {
  it('모든 주소에 보안 헤더가 붙고, 핵심 규칙이 들어 있음', async () => {
    const rules = await nextConfig.headers!();
    expect(rules).toHaveLength(1);
    expect(rules[0].source).toBe('/:path*');
    const h = Object.fromEntries(rules[0].headers.map((x) => [x.key, x.value]));

    expect(h['X-Content-Type-Options']).toBe('nosniff');
    expect(h['X-Frame-Options']).toBe('DENY');
    expect(h['Referrer-Policy']).toBe('no-referrer');
    expect(h['Permissions-Policy']).toContain('camera=()');

    const csp = h['Content-Security-Policy'];
    // 다른 주소로 데이터를 보내는 연결을 막는다 (개인정보를 서버로 보내지 않는다는 원칙을 브라우저가 지켜 줌)
    expect(csp).toContain("connect-src 'self'");
    expect(csp).not.toMatch(/connect-src[^;]*(https?:|\*)/);
    expect(csp).toContain("frame-ancestors 'none'");
    expect(csp).toContain("form-action 'none'");
    expect(csp).toContain("object-src 'none'");
    expect(csp).toContain("base-uri 'self'");
    expect(csp).toContain("default-src 'self'");
  });

  it('운영(배포)용 규칙에는 unsafe-eval과 웹소켓 허용이 없음', async () => {
    const [rule] = await nextConfig.headers!();
    const csp = rule.headers.find((x) => x.key === 'Content-Security-Policy')!.value;
    // 이 테스트는 NODE_ENV=test에서 돌아가므로 개발 전용 허용이 빠진 운영용 모양과 같다
    expect(csp).not.toContain('unsafe-eval');
    expect(csp).not.toMatch(/\bwss?:/);
  });
});

describe('특별한 교육 이름 (파일 이름 "__proto__.xlsx" 등)', () => {
  it('"__proto__", "constructor", "toString"이라는 이름도 칸이 사라지지 않고 판정에 들어감', () => {
    const names = ['__proto__', 'constructor', 'toString', '정상교육'];
    const { rows } = mergeCourses([
      course(names[0], [r('A1', '수강')]),
      course(names[1], [r('A1', '수료')]),
      course(names[2], [r('A1', '수료')]),
      course(names[3], [r('A1', '수료')]),
    ]);
    const cells = rows[0].courses;
    for (const n of names) expect(Object.keys(cells)).toContain(n);
    expect(cells['__proto__']).toBe('미수료');
    expect(rows[0].overall).toBe('미수료'); // "__proto__" 교육의 미수료가 전체 판정에 반영됨
    expect(summarize(rows).total).toBe(1);
  });

  it('결과 엑셀의 통합 시트에도 그 이름의 열과 값이 그대로 나옴', async () => {
    const { wb } = await roundTrip([course('__proto__', [r('A1', '수강')]), course('정상교육', [r('A1', '수료')])]);
    const ws = wb.getWorksheet('통합')!;
    expect(ws.getCell('D1').value).toBe('__proto__');
    expect(ws.getCell('D2').value).toBe('미수료');
    expect(ws.getCell('F2').value).toBe('미수료');
  });
});

describe('아주 긴 부서 이름 (엑셀 수식 글자 한도 255자)', () => {
  it('한도를 넘는 부서는 수식 대신 숫자로 적고, 나머지 부서와 합계는 정상', async () => {
    const longDept = '가'.repeat(MAX_FORMULA_TEXT_LENGTH + 20);
    const { wb } = await roundTrip([
      course('교육1', [r('A1', '수료', longDept), r('A2', '수강', longDept), r('A3', '수료', '인사팀')]),
    ]);
    const s = wb.getWorksheet('요약')!;
    // 11행 = 제목 줄, 12행 = 긴 부서(첫 번째로 나온 부서), 13행 = 인사팀, 14행 = 합계
    expect(s.getCell('A11').value).toBe('부서');
    expect(s.getCell('A12').value).toBe(longDept);
    for (const col of ['B', 'C', 'D', 'E']) {
      expect(typeof s.getCell(`${col}12`).value).toBe('number'); // 수식이 아니라 숫자
    }
    expect([s.getCell('B12').value, s.getCell('C12').value, s.getCell('D12').value, s.getCell('E12').value]).toEqual([2, 1, 1, 0]);
    expect((s.getCell('B13').value as ExcelJS.CellFormulaValue).formula).toContain('SUMPRODUCT'); // 짧은 부서는 그대로 수식
    expect((s.getCell('B14').value as ExcelJS.CellFormulaValue).formula).toBe('SUM(B12:B13)');
  });

  it('한도 안의 긴 이름(따옴표가 섞여 두 배가 되는 경우 포함)은 수식을 그대로 씀', async () => {
    const okDept = '나'.repeat(MAX_FORMULA_TEXT_LENGTH);
    const quoted = '"'.repeat(MAX_FORMULA_TEXT_LENGTH / 2 + 1); // 따옴표는 ""로 두 배가 되어 한도를 넘음
    const { wb } = await roundTrip([course('교육1', [r('A1', '수료', okDept), r('A2', '수료', quoted)])]);
    const s = wb.getWorksheet('요약')!;
    expect((s.getCell('B12').value as ExcelJS.CellFormulaValue).formula).toContain('SUMPRODUCT');
    expect(typeof s.getCell('B13').value).toBe('number');
  });
});
