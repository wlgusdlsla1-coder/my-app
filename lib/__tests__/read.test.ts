// 1회차 자동 채점: 파일 검사 · 열 자동 찾기 · 엑셀 읽기 (DESIGN.md §10.2 13번 등)

import { existsSync, readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { detectColumns, findColumn, isMappingComplete } from '../detectColumns';
import { extractCourse, readSheet } from '../readExcel';
import { checkCourseFiles, checkFile, courseNameFromFile } from '../validateFiles';

const SAMPLES = path.resolve(__dirname, '..', '..', 'samples');

/** 샘플 엑셀을 브라우저와 같은 ArrayBuffer 모양으로 읽기 */
function load(...parts: string[]): ArrayBuffer {
  const buf = readFileSync(path.join(SAMPLES, ...parts));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
}

describe('파일 검사 (E01·E02·E03·E06)', () => {
  it('.xlsx가 아니면 E01', () => {
    expect(checkFile({ name: '명단.csv', size: 10 })).toBe('E01');
    expect(checkFile({ name: '명단.xls', size: 10 })).toBe('E01');
  });

  it('10MB를 넘으면 E02', () => {
    expect(checkFile({ name: '명단.xlsx', size: 10 * 1024 * 1024 + 1 })).toBe('E02');
    expect(checkFile({ name: '명단.XLSX', size: 10 * 1024 * 1024 })).toBeNull();
  });

  it('교육 파일은 20개까지만 받고 넘는 것은 E03', () => {
    const files = Array.from({ length: 21 }, (_, i) => ({ name: `교육${i + 1}.xlsx`, size: 1 }));
    const r = checkCourseFiles([], files);
    expect(r.accepted).toHaveLength(20);
    expect(r.rejected).toEqual([{ file: files[20], code: 'E03' }]);
  });

  it('이미 있는 교육 이름과 겹치면 E06', () => {
    const r = checkCourseFiles([{ name: '청렴교육.xlsx', size: 1 }], [{ name: '청렴교육.xlsx', size: 1 }]);
    expect(r.rejected[0].code).toBe('E06');
  });

  it('교육 이름은 파일 이름에서 확장자를 뺀 부분', () => {
    expect(courseNameFromFile('성희롱예방교육.xlsx')).toBe('성희롱예방교육');
    expect(courseNameFromFile('청렴교육.XLSX')).toBe('청렴교육');
  });
});

describe('열 자동 찾기', () => {
  it('"직원번호", "소속", "이수여부"를 사번·부서·수료 여부로 인식 (DESIGN §10.2-13)', () => {
    const r = detectColumns(['직원번호', '성명', '소속', '이수여부']);
    expect(r.mapping).toEqual({ empNo: 0, name: 1, dept: 2, status: 3 });
    expect(r.missing).toEqual([]);
  });

  it('제목 칸의 공백·대소문자는 무시', () => {
    expect(findColumn(['수료 여부'], 'status')).toBe(0);
    expect(findColumn(['NAME'], 'name')).toBe(0);
  });

  it('3글자 이상 이름은 포함만 돼도 인식 (예: "수료여부(Y/N)")', () => {
    expect(findColumn(['수료여부(Y/N)'], 'status')).toBe(0);
  });

  it('"직원ID", "진행상태"는 인식하지 않고 열 고르기로 넘김', () => {
    const r = detectColumns(['직원ID', '성명', '소속부서', '진행상태']);
    expect(r.missing).toEqual(['empNo', 'status']);
    expect(isMappingComplete(r.mapping)).toBe(false);
    expect(isMappingComplete({ ...r.mapping, empNo: 0, status: 3 })).toBe(true);
  });

  it('부서와 팀이 함께 있으면 부서를 고름', () => {
    expect(findColumn(['사번', '부서', '팀'], 'dept')).toBe(1);
  });
});

describe('엑셀 읽기 — 가짜 데이터(samples/basic)', () => {
  it('기본형 교육 파일: 네 열 인식, 30명, 앞자리 0 유지', async () => {
    const sheet = await readSheet(load('basic', '성희롱예방교육.xlsx'));
    const { mapping, missing } = detectColumns(sheet.headers);
    expect(missing).toEqual([]);
    const course = extractCourse(sheet, mapping, '성희롱예방교육');
    expect(course.records).toHaveLength(30);
    expect(course.records.map((r) => r.empNo)).toContain('00123');
    expect(course.records[0]).toMatchObject({ empNo: 'A001', dept: '인사팀', raw: '수료' });
  });

  it('위에 안내 줄이 있어도 제목 줄을 찾고, 사번·값의 앞뒤 공백을 지움', async () => {
    const sheet = await readSheet(load('basic', '개인정보보호교육.xlsx'));
    expect(sheet.headers).toEqual(['직원번호', '성명', '소속', '이수여부']);
    const { mapping } = detectColumns(sheet.headers);
    const course = extractCourse(sheet, mapping, '개인정보보호교육');
    expect(course.records).toHaveLength(30);
    const a005 = course.records.find((r) => r.empNo === 'A005');
    expect(a005?.raw).toBe('수료');
    expect(course.records.find((r) => r.empNo === 'A006')).toBeDefined();
  });

  it('잘못된 파일은 E04 오류', async () => {
    const junk = new TextEncoder().encode('엑셀 아님').buffer as ArrayBuffer;
    await expect(readSheet(junk)).rejects.toMatchObject({ code: 'E04' });
  });
});

// 사용자가 준 샘플(samples/user)은 Git에 올리지 않으므로, 파일이 있을 때만 검사한다
const USER_COURSE = '고용노동부와_보건복지부가_함께하는_장애인인식개선교육] - 샘플.xlsx';
const hasUserSample = existsSync(path.join(SAMPLES, 'user', USER_COURSE));

describe.skipIf(!hasUserSample)('엑셀 읽기 — 사용자 샘플(samples/user)', () => {
  it('실제 교육 플랫폼 양식(12열)에서 사번·성명·부서·수료여부를 자동 인식 (팀 열이 아니라 부서 열)', async () => {
    const sheet = await readSheet(load('user', USER_COURSE));
    const { mapping, missing } = detectColumns(sheet.headers);
    expect(missing).toEqual([]);
    expect(sheet.headers[mapping.empNo!]).toBe('사번');
    expect(sheet.headers[mapping.name!]).toBe('성명');
    expect(sheet.headers[mapping.dept!]).toBe('부서');
    expect(sheet.headers[mapping.status!]).toBe('수료여부');
  });

  it('번호만 채워진 빈 줄 2천여 개는 건너뛰고, 사번 앞자리 0을 유지', async () => {
    const sheet = await readSheet(load('user', USER_COURSE));
    const { mapping } = detectColumns(sheet.headers);
    const course = extractCourse(sheet, mapping, '장애인인식개선교육');
    expect(course.records).toEqual([
      { empNo: '123456', name: '홍길동', dept: '서울지역본부', raw: '수료' },
      { empNo: '024567', name: '김치만두', dept: '국유재산관리처', raw: '수료' },
    ]);
  });
});
