// 4회차 자동 채점: 화면이 쓰는 "파일 읽기 → 목록 칸" 로직과 오류 메시지 (DESIGN.md §8)

import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { errorMessage } from '../messages';
import { applyMapping, canProcess, dropSheet, readCourseEntry } from '../pipeline';

const SAMPLES = path.resolve(__dirname, '..', '..', 'samples', 'basic');

function load(name: string): ArrayBuffer {
  const buf = readFileSync(path.join(SAMPLES, name));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
}

describe('readCourseEntry', () => {
  it('열을 다 찾으면 ready, 교육 이름은 파일 이름에서, 네 열만 기록에 남음', async () => {
    const e = await readCourseEntry('1', '성희롱예방교육.xlsx', load('성희롱예방교육.xlsx'));
    expect(e.status).toBe('ready');
    expect(e.courseName).toBe('성희롱예방교육');
    expect(e.course?.records).toHaveLength(30);
    expect(Object.keys(e.course!.records[0]).sort()).toEqual(['dept', 'empNo', 'name', 'raw']);
    expect(e.sheet).toBeUndefined(); // 표는 들고 있지 않는다
  });

  it('열을 못 찾으면 needsColumns, 못 찾은 열 이름을 메시지에 적음', async () => {
    const e = await readCourseEntry('2', '산업안전보건교육.xlsx', load('산업안전보건교육.xlsx'));
    expect(e.status).toBe('needsColumns');
    expect(e.error).toContain('사번·수료 여부');
    expect(e.sheet?.headers).toEqual(['직원ID', '성명', '소속부서', '진행상태']);
    expect(e.mapping).toEqual({ name: 1, dept: 2 });
    expect(canProcess([e])).toBe(false);
  });

  it('엑셀이 아닌 내용은 error (E04)', async () => {
    const junk = new TextEncoder().encode('엑셀 아님').buffer as ArrayBuffer;
    const e = await readCourseEntry('3', '깨진파일.xlsx', junk);
    expect(e.status).toBe('error');
    expect(e.error).toBe(errorMessage('E04', '깨진파일.xlsx'));
    expect(canProcess([e])).toBe(false);
  });
});

describe('applyMapping', () => {
  it('운영자가 고른 열을 적용하면 ready가 되고 표를 버림', async () => {
    const e = await readCourseEntry('2', '산업안전보건교육.xlsx', load('산업안전보건교육.xlsx'));
    const done = applyMapping(e, { ...e.mapping, empNo: 0, status: 3 });
    expect(done.status).toBe('ready');
    expect(done.sheet).toBeUndefined();
    expect(done.mapping).toBeUndefined();
    expect(done.error).toBeUndefined();
    expect(done.course?.records).toHaveLength(30);
    expect(done.course?.records.filter((r) => r.raw === '수강')).toHaveLength(1);
    expect(canProcess([done])).toBe(true);
  });

  it('열이 덜 정해졌으면 그대로 둠', async () => {
    const e = await readCourseEntry('2', '산업안전보건교육.xlsx', load('산업안전보건교육.xlsx'));
    expect(applyMapping(e, { empNo: 0 })).toBe(e);
  });
});

describe('dropSheet (열 고르기 취소 뒤 읽어 둔 표 버리기)', () => {
  it('표만 버리고 칸은 "열 고르기 필요"로 남김. 자동으로 찾은 열과 안내 문구는 유지', async () => {
    const e = await readCourseEntry('2', '산업안전보건교육.xlsx', load('산업안전보건교육.xlsx'));
    expect(e.sheet).toBeDefined();
    const dropped = dropSheet(e);
    expect(dropped.sheet).toBeUndefined();
    expect(dropped.status).toBe('needsColumns');
    expect(dropped.mapping).toEqual(e.mapping);
    expect(dropped.error).toBe(e.error);
    expect(dropped.fileName).toBe(e.fileName);
    expect(canProcess([dropped])).toBe(false);
  });

  it('다시 읽으면 같은 제목과 자동 인식 결과가 돌아옴 (취소 뒤 [열 고르기]를 다시 눌렀을 때)', async () => {
    const first = await readCourseEntry('2', '산업안전보건교육.xlsx', load('산업안전보건교육.xlsx'));
    const again = await readCourseEntry('2', '산업안전보건교육.xlsx', load('산업안전보건교육.xlsx'));
    expect(again.sheet?.headers).toEqual(first.sheet?.headers);
    expect(again.mapping).toEqual(first.mapping);
  });

  it('표가 없거나 준비된 칸은 그대로 돌려줌', async () => {
    const ready = await readCourseEntry('1', '성희롱예방교육.xlsx', load('성희롱예방교육.xlsx'));
    expect(dropSheet(ready)).toBe(ready);
    const needs = await readCourseEntry('2', '산업안전보건교육.xlsx', load('산업안전보건교육.xlsx'));
    const dropped = dropSheet(needs);
    expect(dropSheet(dropped)).toBe(dropped);
  });

  it('표를 버린 칸에는 열 고르기를 마쳐도 적용되지 않음 (표가 있어야 열을 뽑을 수 있다)', async () => {
    const e = await readCourseEntry('2', '산업안전보건교육.xlsx', load('산업안전보건교육.xlsx'));
    const dropped = dropSheet(e);
    expect(applyMapping(dropped, { empNo: 0, name: 1, dept: 2, status: 3 })).toBe(dropped);
  });
});

describe('canProcess', () => {
  it('파일이 없으면 false', () => {
    expect(canProcess([])).toBe(false);
  });
});

describe('오류 메시지 (DESIGN §8)', () => {
  it('파일 이름과 제한값이 메시지에 들어감', () => {
    expect(errorMessage('E01', '명단.csv')).toContain('명단.csv');
    expect(errorMessage('E02', 'a.xlsx')).toContain('10MB');
    expect(errorMessage('E03', 'b.xlsx')).toContain('20개');
    expect(errorMessage('E06', 'c.xlsx')).toContain('같은 이름');
    expect(errorMessage('E08')).toContain('대상자 수');
  });
});
