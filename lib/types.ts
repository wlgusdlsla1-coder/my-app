// 설계 참고: DESIGN.md §4.1 — 앱 전체에서 쓰는 데이터 모양

/**
 * 대상자 한 사람.
 * 직원 명단 파일은 없고, 교육 파일들에 한 번이라도 나온 사번 전체가 대상자다 (DESIGN.md D-4)
 */
export type Employee = {
  /** 사번 (앞뒤 공백 제거, 글자로 통일. 앞자리 0은 지우지 않는다) */
  empNo: string;
  /** 이름 */
  name: string;
  /** 부서 */
  dept: string;
};

/** 교육 칸 하나의 판정 결과 (DESIGN.md §4.2) */
export type CourseStatus = '수료' | '미수료' | '교육 기록 없음' | '확인 필요';

/** 전체 수료 여부 (DESIGN.md §4.3) */
export type OverallStatus = '수료' | '미수료' | '미수료(교육기록 미흡)' | '확인 필요';

/** 교육 파일 안의 한 줄 */
export type CourseRecord = {
  empNo: string;
  name: string;
  dept: string;
  /** 원본 수료 여부 값 (앞뒤 공백만 지운 상태) */
  raw: string;
};

/** 교육 파일 하나 */
export type CourseFile = {
  /** 파일 이름에서 확장자를 뺀 부분 */
  courseName: string;
  records: CourseRecord[];
};

/** 통합 시트의 한 줄 */
export type MergedRow = Employee & {
  /** 교육 이름 → 그 교육 칸의 판정 */
  courses: Record<string, CourseStatus>;
  overall: OverallStatus;
};

/** 열 찾기 대상 */
export type Field = 'empNo' | 'name' | 'dept' | 'status';

/** 각 열이 몇 번째 칸(0부터)인지 */
export type ColumnMapping = Partial<Record<Field, number>>;

/**
 * 엑셀 첫 번째 시트를 글자 표로 읽은 결과.
 * 열을 고를 때까지만 브라우저 메모리에 잠깐 두고, 필요한 열만 뽑아낸 뒤에는 버린다.
 */
export type RawSheet = {
  /** 제목 줄의 칸 이름들 */
  headers: string[];
  /** 제목 줄 아래의 데이터 줄들 */
  rows: string[][];
};

/** 오류 코드 (DESIGN.md §8). E07은 직원 명단 파일이 없어져 쓰지 않는다 */
export type ErrorCode = 'E01' | 'E02' | 'E03' | 'E04' | 'E05' | 'E06' | 'E08';
