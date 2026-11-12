// 설계 참고: DESIGN.md §2.2 ⑤~⑦, D-4 — 교육 파일들을 사번 기준으로 합쳐 직원 한 명당 한 줄로 만들기

import { NO_DEPT_LABEL, UNKNOWN_EMP_PREFIX } from './constants';
import { judgeCourseCell, judgeOverall } from './judge';
import type { CourseFile, CourseRecord, Employee, MergedRow, OverallStatus } from './types';

export type MergeResult = {
  /** 교육 이름들 (올린 순서) = 통합 시트의 교육 열 순서 */
  courseNames: string[];
  /** 대상자 한 사람당 한 줄 (교육 파일에 처음 나온 순서) */
  rows: MergedRow[];
};

/**
 * 사번 칸이 빈 줄에 "알수없음1", "알수없음2" … 이름표를 붙인다 (DESIGN.md D-5).
 * - 사번이 없어서 같은 사람인지 알 수 없으므로 **빈 줄 하나가 곧 한 사람**이다.
 * - 번호는 올린 파일 순서와 줄 순서대로 매기고, 파일이 달라도 겹치지 않게 **전체에서 하나씩** 늘린다.
 *   (파일마다 1부터 매기면 서로 다른 파일의 "알수없음1"이 같은 사람으로 합쳐진다.)
 */
function labelUnknownEmpNos(courses: readonly CourseFile[]): CourseFile[] {
  let count = 0;
  return courses.map((course) => ({
    ...course,
    records: course.records.map((rec) => (rec.empNo === '' ? { ...rec, empNo: `${UNKNOWN_EMP_PREFIX}${++count}` } : rec)),
  }));
}

/**
 * 교육 파일들을 사번 기준으로 합친다.
 * - 직원 명단 파일은 없다. 교육 파일에 한 번이라도 나온 사번 전체가 대상자다 (D-4)
 * - 줄 순서와 이름·부서는 처음 나온 것을 쓴다 (이름·부서는 파일마다 같다)
 * - 사번은 읽을 때 이미 앞뒤 공백을 지웠으므로 그대로 비교한다 (앞자리 0 유지)
 */
export function mergeCourses(input: readonly CourseFile[]): MergeResult {
  if (input.length === 0) throw new Error('교육 파일이 1개 이상 있어야 해요.');
  const courses = labelUnknownEmpNos(input);

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
    // 교육 이름이 "__proto__" 같은 특별한 이름이어도 칸이 사라지지 않도록, 기본 속성이 없는 빈 객체를 쓴다
    const cells: MergedRow['courses'] = Object.create(null);
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

/** 부서 한 곳의 인원 집계 (DESIGN.md §7.2 부서별 요약, FR-11) */
export type DeptSummary = {
  /** 결과 엑셀에 적는 부서 이름 (비어 있으면 "(부서 없음)") */
  dept: string;
  total: number;
  completed: number;
  /** 미수료 ("미수료" + "미수료(교육기록 미흡)") */
  notCompleted: number;
  needsCheck: number;
};

/** 결과 엑셀에 적는 부서 이름: 앞뒤 공백을 지우고, 비어 있으면 "(부서 없음)" */
export function deptLabel(dept: string): string {
  const trimmed = dept.trim();
  return trimmed === '' ? NO_DEPT_LABEL : trimmed;
}

/**
 * 부서별로 인원을 센다. 부서는 대상자에 처음 나온 순서대로 보여준다.
 * 엑셀의 COUNTIF가 대소문자를 구분하지 않으므로, 대소문자만 다른 부서("HR"과 "hr")는 같은 부서로 묶는다.
 */
export function summarizeByDept(rows: readonly MergedRow[]): DeptSummary[] {
  const groups = new Map<string, DeptSummary>();
  for (const row of rows) {
    const label = deptLabel(row.dept);
    const key = label.toLowerCase();
    let g = groups.get(key);
    if (!g) {
      g = { dept: label, total: 0, completed: 0, notCompleted: 0, needsCheck: 0 };
      groups.set(key, g);
    }
    g.total++;
    if (row.overall === '수료') g.completed++;
    else if (row.overall === '확인 필요') g.needsCheck++;
    else g.notCompleted++;
  }
  return [...groups.values()];
}

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
