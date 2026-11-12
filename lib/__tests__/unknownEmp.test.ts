// 사번 칸이 빈 줄 → "알수없음1", "알수없음2" … (DESIGN.md D-5)

import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { buildWorkbook, workbookToBytes } from '../buildWorkbook';
import { mergeCourses, summarize } from '../merge';
import { readCourseEntry } from '../pipeline';
import { extractCourse } from '../readExcel';
import type { CourseFile, CourseRecord } from '../types';

const r = (empNo: string, raw: string, name = '김민준', dept = '인사팀'): CourseRecord => ({ empNo, name, dept, raw });
const course = (courseName: string, records: CourseRecord[]): CourseFile => ({ courseName, records });

describe('읽을 때: 사번이 빈 줄을 버리지 않는다', () => {
  const sheet = {
    headers: ['번호', '사번', '이름', '부서', '수료여부'],
    rows: [
      ['1', 'A001', '가', '인사팀', '수료'],
      ['2', '', '나', '재무팀', '수강'], // 사번만 빔 → 살린다
      ['3', '', '', '', ''], // 번호만 있는 빈 줄 → 건너뜀
      ['4', '', '', '', '수료'], // 수료 여부만 있음 → 살린다 (값이 있는 줄이므로)
      ['5', '  ', '다', '', ''], // 사번이 공백뿐, 이름만 있음 → 살린다
    ],
  };
  const mapping = { empNo: 1, name: 2, dept: 3, status: 4 };

  it('네 열이 모두 빈 줄만 건너뛰고 나머지는 남긴다', () => {
    const c = extractCourse(sheet, mapping, '교육1');
    expect(c.records.map((x) => [x.empNo, x.name, x.raw])).toEqual([
      ['A001', '가', '수료'],
      ['', '나', '수강'],
      ['', '', '수료'],
      ['', '다', ''],
    ]);
  });
});

describe('합칠 때: 알수없음N 이름표', () => {
  it('빈 줄 하나가 한 사람, 번호는 파일이 달라도 겹치지 않게 전체에서 하나씩 늘어남', () => {
    const { rows } = mergeCourses([
      course('교육1', [r('', '수료', '가나'), r('A1', '수료'), r('', '수강', '다라')]),
      course('교육2', [r('', '수료', '마바'), r('A1', '수료')]),
    ]);
    expect(rows.map((x) => x.empNo)).toEqual(['알수없음1', 'A1', '알수없음2', '알수없음3']);
    // 서로 다른 파일의 빈 줄이 같은 사람으로 합쳐지지 않는다
    expect(rows.find((x) => x.empNo === '알수없음3')?.name).toBe('마바');
    expect(rows.find((x) => x.empNo === '알수없음3')?.courses).toEqual({ 교육1: '교육 기록 없음', 교육2: '수료' });
  });

  it('같은 파일의 빈 줄 두 개는 "같은 사번 두 줄(확인 필요)"이 아니라 서로 다른 두 사람', () => {
    const { rows } = mergeCourses([course('교육1', [r('', '수료', '가'), r('', '수료', '나')])]);
    expect(rows.map((x) => [x.empNo, x.courses['교육1']])).toEqual([
      ['알수없음1', '수료'],
      ['알수없음2', '수료'],
    ]);
  });

  it('이름·부서·수료 값은 그대로 보존되고, 다른 교육에 기록이 없으면 미수료(교육기록 미흡)', () => {
    const { rows } = mergeCourses([course('교육1', [r('', '수료', '가나', '재무팀'), r('A1', '수료')]), course('교육2', [r('A1', '수료')])]);
    const u = rows[0];
    expect(u).toMatchObject({ empNo: '알수없음1', name: '가나', dept: '재무팀', overall: '미수료(교육기록 미흡)' });
  });

  it('같은 입력으로 다시 합쳐도 같은 이름표가 나오고, 입력 자료는 바뀌지 않음', () => {
    const input = [course('교육1', [r('', '수료'), r('', '수강')]), course('교육2', [r('', '수료')])];
    const a = mergeCourses(input).rows.map((x) => x.empNo);
    const b = mergeCourses(input).rows.map((x) => x.empNo);
    expect(a).toEqual(['알수없음1', '알수없음2', '알수없음3']);
    expect(b).toEqual(a);
    expect(input[0].records[0].empNo).toBe(''); // 원본은 그대로
  });

  it('인원 집계의 합계는 대상자 수와 같음', () => {
    const { rows } = mergeCourses([course('교육1', [r('', '수료'), r('', '수강'), r('A1', '수료')])]);
    const s = summarize(rows);
    expect(s.targets).toBe(3);
    expect(s.total).toBe(3);
  });
});

describe('엑셀 파일부터 결과 엑셀까지', () => {
  async function fileBytes(rows: (string | number)[][]): Promise<ArrayBuffer> {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.addRow(['번호', '사번', '이름', '부서', '수료여부']);
    rows.forEach((row) => ws.addRow(row));
    const u = new Uint8Array(await wb.xlsx.writeBuffer());
    return u.buffer.slice(u.byteOffset, u.byteOffset + u.byteLength) as ArrayBuffer;
  }

  it('번호만 있는 빈 줄은 버리고 사번만 빈 줄은 알수없음N으로 결과 엑셀에 나옴', async () => {
    const bytes = await fileBytes([
      [1, '123456', '홍길동', '인사팀', '수료'],
      [2, '', '이름만있음', '재무팀', '수강'], // 사번 빔
      [3, '', '', '', ''], // 번호만 있는 줄
      [4, '', '', '', ''],
    ]);
    const entry = await readCourseEntry('1', '교육A.xlsx', bytes);
    expect(entry.status).toBe('ready');
    expect(entry.course!.records).toHaveLength(2); // 번호만 있는 두 줄은 건너뜀

    const merged = mergeCourses([entry.course!]);
    const wb = buildWorkbook(merged);
    const loaded = new ExcelJS.Workbook();
    const out = await workbookToBytes(wb);
    await loaded.xlsx.load(out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength) as ArrayBuffer);
    const ws = loaded.getWorksheet('통합')!;
    expect(ws.getCell('A2').value).toBe('123456');
    expect(ws.getCell('A3').value).toBe('알수없음1');
    expect(ws.getCell('B3').value).toBe('이름만있음');
    expect(ws.getCell('C3').value).toBe('재무팀');
    expect(ws.getCell('D3').value).toBe('미수료');
  });
});
