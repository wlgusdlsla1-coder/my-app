// 2회차 자동 채점: 사번 기준 합치기 · 인원 집계 · 가짜 데이터 정답표와 비교 (DESIGN.md §10.2 6·7·12번)

import { readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { detectColumns } from '../detectColumns';
import { mergeCourses, summarize } from '../merge';
import { extractCourse, readSheet } from '../readExcel';
import type { ColumnMapping, CourseFile, CourseRecord } from '../types';
import { courseNameFromFile } from '../validateFiles';

const SAMPLES = path.resolve(__dirname, '..', '..', 'samples');

const r = (empNo: string, raw: string, name = '김민준', dept = '인사팀'): CourseRecord => ({ empNo, name, dept, raw });
const course = (courseName: string, records: CourseRecord[]): CourseFile => ({ courseName, records });

describe('사번 기준 합치기 (D-4)', () => {
  it('교육 파일에 나온 사번 전체가 대상자, 처음 나온 순서대로', () => {
    const { rows, courseNames } = mergeCourses([
      course('교육1', [r('A002', '수료'), r('A001', '수료')]),
      course('교육2', [r('A001', '수료'), r('A003', '수료')]),
    ]);
    expect(courseNames).toEqual(['교육1', '교육2']);
    expect(rows.map((x) => x.empNo)).toEqual(['A002', 'A001', 'A003']);
  });

  it('이름·부서는 처음 나온 값을 씀', () => {
    const { rows } = mergeCourses([
      course('교육1', [r('A001', '수료', '김민준', '인사팀')]),
      course('교육2', [r('A001', '수료', '김민준', '재무팀')]),
    ]);
    expect(rows[0]).toMatchObject({ name: '김민준', dept: '인사팀' });
  });

  it('6. 사번 앞뒤 공백은 읽을 때 지워져 같은 사람으로 합쳐짐', () => {
    // 공백 제거는 extractCourse에서 하므로, 여기서는 지운 뒤의 값으로 확인
    const { rows } = mergeCourses([course('교육1', [r('A001', '수료')]), course('교육2', [r('A001', '수강')])]);
    expect(rows).toHaveLength(1);
    expect(rows[0].courses).toEqual({ 교육1: '수료', 교육2: '미수료' });
  });

  it('7. 사번 "00123"과 "123"은 다른 사람 (앞자리 0 유지)', () => {
    const { rows } = mergeCourses([course('교육1', [r('00123', '수료'), r('123', '수료')])]);
    expect(rows.map((x) => x.empNo)).toEqual(['00123', '123']);
  });

  it('12. 한 교육 파일에만 나온 사번 → 대상자에 들어가고, 나머지 교육은 교육 기록 없음', () => {
    const { rows } = mergeCourses([
      course('교육1', [r('A001', '수료')]),
      course('교육2', [r('A001', '수료'), r('Z999', '수료')]),
    ]);
    const z = rows.find((x) => x.empNo === 'Z999');
    expect(z?.courses).toEqual({ 교육1: '교육 기록 없음', 교육2: '수료' });
    expect(z?.overall).toBe('미수료(교육기록 미흡)');
  });

  it('교육 파일이 없으면 오류', () => {
    expect(() => mergeCourses([])).toThrow();
  });
});

describe('인원 집계', () => {
  it('미수료에는 교육기록 미흡도 포함하고, 합계 = 대상자 수', () => {
    const { rows } = mergeCourses([
      course('교육1', [r('A1', '수료'), r('A2', '수강'), r('A3', '이수'), r('A4', '수료')]),
      course('교육2', [r('A1', '수료'), r('A2', '수료'), r('A3', '수료')]),
    ]);
    expect(summarize(rows)).toEqual({
      completed: 1, // A1
      notCompleted: 2, // A2(미수료), A4(교육2 기록 없음)
      missingRecord: 1, // A4
      needsCheck: 1, // A3
      total: 4,
      targets: 4,
    });
  });
});

// ── 가짜 데이터 전체 흐름: 읽기 → 열 찾기 → 합치기 → 판정 ─────────────

/** 폴더의 교육 파일들을 읽어 CourseFile로. 자동으로 못 찾은 열은 overrides로 직접 지정 (열 고르기 창 역할) */
async function loadFolder(dir: string, overrides: Record<string, ColumnMapping> = {}): Promise<CourseFile[]> {
  const files = readdirSync(dir).filter((f) => f.endsWith('.xlsx')).sort();
  const result: CourseFile[] = [];
  for (const f of files) {
    const buf = readFileSync(path.join(dir, f));
    const sheet = await readSheet(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer);
    const name = courseNameFromFile(f);
    const { mapping } = detectColumns(sheet.headers);
    result.push(extractCourse(sheet, { ...mapping, ...overrides[name] }, name));
  }
  return result;
}

describe('가짜 데이터 정답표와 비교 (samples/basic)', () => {
  it('요약 숫자와 수료가 아닌 직원이 정답표와 같음', async () => {
    // 산업안전보건교육은 열 제목이 "직원ID"·"진행상태"라 운영자가 고른 것처럼 직접 지정
    const courses = await loadFolder(path.join(SAMPLES, 'basic'), {
      산업안전보건교육: { empNo: 0, status: 3 },
    });
    const { rows } = mergeCourses(courses);

    expect(summarize(rows)).toEqual({
      completed: 20, notCompleted: 9, missingRecord: 3, needsCheck: 2, total: 31, targets: 31,
    });

    const notDone = Object.fromEntries(rows.filter((x) => x.overall !== '수료').map((x) => [x.empNo, x.overall]));
    expect(notDone).toEqual({
      A003: '미수료', A008: '미수료', A011: '미수료', A014: '미수료', A022: '미수료', A027: '미수료',
      A015: '확인 필요', A016: '확인 필요',
      A020: '미수료(교육기록 미흡)', A021: '미수료(교육기록 미흡)', Z999: '미수료(교육기록 미흡)',
    });

    // 공백 섞인 사번(A005·A006)이 따로 생기지 않고, 앞자리 0 사번이 유지됨
    expect(rows.filter((x) => x.empNo.includes('A005'))).toHaveLength(1);
    expect(rows.some((x) => x.empNo === '00123')).toBe(true);
    // 이 파일에만 나온 Z999는 맨 끝 줄
    expect(rows.at(-1)?.empNo).toBe('Z999');
  });
});

describe('가짜 데이터 정답표와 비교 (samples/full, 500명 × 20개)', () => {
  it('요약 숫자가 정답표와 같고, 처리 시간이 넉넉히 짧음', async () => {
    const start = performance.now();
    const courses = await loadFolder(path.join(SAMPLES, 'full'));
    const { rows, courseNames } = mergeCourses(courses);
    const ms = performance.now() - start;

    expect(courseNames).toHaveLength(20);
    expect(summarize(rows)).toEqual({
      completed: 258, notCompleted: 242, missingRecord: 51, needsCheck: 0, total: 500, targets: 500,
    });
    // 목표는 5분 이내. 읽기·합치기만으로 30초를 넘으면 문제가 있는 것으로 본다
    expect(ms).toBeLessThan(30_000);
    console.log(`samples/full 읽기+합치기: ${Math.round(ms)}ms`);
  }, 60_000);
});
