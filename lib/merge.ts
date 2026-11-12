// 설계 참고: DESIGN.md §2.2 ⑤~⑦, D-4 — 교육 파일들을 사번 기준으로 합쳐 직원 한 명당 한 줄로 만들기

import { judgeCourseCell, judgeOverall } from './judge';
import type { CourseFile, CourseRecord, Employee, MergedRow, OverallStatus } from './types';

export type MergeResult = {
  /** 교육 이름들 (올린 순서) = 통합 시트의 교육 열 순서 */
  courseNames: string[];
  /** 대상자 한 사람당 한 줄 (교육 파일에 처음 나온 순서) */
  rows: MergedRow[];
};

/**
 * 교육 파일들을 사번 기준으로 합친다.
 * - 직원 명단 파일은 없다. 교육 파일에 한 번이라도 나온 사번 전체가 대상자다 (D-4)
 * - 줄 순서와 이름·부서는 처음 나온 것을 쓴다 (이름·부서는 파일마다 같다)
 * - 사번은 읽을 때 이미 앞뒤 공백을 지웠으므로 그대로 비교한다 (앞자리 0 유지)
 */
export function mergeCourses(courses: readonly CourseFile[]): MergeResult {
  if (courses.length === 0) throw new Error('교육 파일이 1개 이상 있어야 해요.');

  // ① 대상자 목록: 올린 순서대로 파일을 훑으며 처음 나온 사번을 차례로 모은다
  const employees = new Map<string, Employee>();
  for (const course of courses) {
    for (const rec of course.records) {
      if (!employees.has(rec.empNo)) {
        employees.set(rec.empNo, { empNo: rec.empNo, name: rec.name, dept: rec.dept });
      }
    }
  }

  // ② 교육마다 사번 → 그 사번의 줄들 (같은 사번이 여러 줄이면 판정에서 "확인 필요")
  const byCourse = courses.map((course) => {
    const map = new Map<string, CourseRecord[]>();
    for (const rec of course.records) {
      const list = map.get(rec.empNo);
      if (list) list.push(rec);
      else map.set(rec.empNo, [rec]);
    }
    return { courseName: course.courseName, map };
  });

  // ③ 직원마다 교육 칸 판정 → 전체 판정
  const rows: MergedRow[] = [];
  for (const emp of employees.values()) {
    const cells: MergedRow['courses'] = {};
    for (const { courseName, map } of byCourse) {
      cells[courseName] = judgeCourseCell(map.get(emp.empNo) ?? []);
    }
    rows.push({ ...emp, courses: cells, overall: judgeOverall(Object.values(cells)) });
  }

  return { courseNames: courses.map((c) => c.courseName), rows };
}

/** 결과 화면·요약 시트용 인원 집계 */
export type Summary = {
  /** 전체 수료 */
  completed: number;
  /** 미수료 ("미수료" + "미수료(교육기록 미흡)") */
  notCompleted: number;
  /** 그중 교육기록 미흡 */
  missingRecord: number;
  /** 확인 필요 */
  needsCheck: number;
  /** 합계 = 수료 + 미수료 + 확인 필요 */
  total: number;
  /** 대상자 수 (고유 사번 수). 합계와 다르면 E08 */
  targets: number;
};

export function summarize(rows: readonly MergedRow[]): Summary {
  const count = (s: OverallStatus) => rows.filter((r) => r.overall === s).length;
  const completed = count('수료');
  const missingRecord = count('미수료(교육기록 미흡)');
  const notCompleted = count('미수료') + missingRecord;
  const needsCheck = count('확인 필요');
  return {
    completed,
    notCompleted,
    missingRecord,
    needsCheck,
    total: completed + notCompleted + needsCheck,
    targets: new Set(rows.map((r) => r.empNo)).size,
  };
}
