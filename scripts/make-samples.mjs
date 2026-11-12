// 테스트용 가짜 엑셀을 만드는 스크립트 (설계 참고: DESIGN.md §10.5)
// 실행: npm run samples
// 결과:
//   samples/basic/  — 직원 30명 + 교육 파일 4개 (여러 예외 상황을 일부러 넣은 판) + 정답표.md
//   samples/full/   — 직원 500명 + 교육 파일 20개 (시간 측정용 큰 판) + 정답표.md
// ※ 모든 이름·사번은 지어낸 가짜 값이다. 실제 직원 데이터는 절대 넣지 않는다.

import ExcelJS from 'exceljs';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';

const ROOT = path.resolve(import.meta.dirname, '..', 'samples');

// ── 공통 도구 ─────────────────────────────────────────────

/** 같은 결과가 매번 나오도록 씨앗값을 쓰는 난수 */
function makeRandom(seed) {
  let s = seed;
  return () => {
    s = (s * 1664525 + 1013904223) % 4294967296;
    return s / 4294967296;
  };
}

const FAMILY = ['김', '이', '박', '최', '정', '강', '조', '윤', '장', '임', '한', '오', '서', '신', '권'];
const GIVEN = ['민준', '서연', '도윤', '하은', '지호', '수아', '예준', '지유', '주원', '서윤', '하준', '채원', '시우', '지아', '은우', '다은'];
const DEPTS = ['인사팀', '재무팀', '영업팀', '개발팀', '생산팀', '품질팀', '총무팀', '기획팀'];

/** 가짜 직원 n명 만들기 */
function makeEmployees(n, deptCount, rand) {
  return Array.from({ length: n }, (_, i) => ({
    empNo: `A${String(i + 1).padStart(3, '0')}`,
    name: FAMILY[Math.floor(rand() * FAMILY.length)] + GIVEN[Math.floor(rand() * GIVEN.length)],
    dept: DEPTS[i % deptCount],
  }));
}

/** 표 한 장짜리 엑셀 저장. titleRows는 제목 줄 위에 넣을 안내 줄들 */
async function saveSheet(filePath, headers, rows, titleRows = []) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  for (const t of titleRows) ws.addRow(t);
  ws.addRow(headers);
  for (const r of rows) ws.addRow(r);
  await wb.xlsx.writeFile(filePath);
}

/** 정답 계산: DESIGN.md §4.3 순서 그대로 */
function overallOf(statuses) {
  if (statuses.includes('교육 기록 없음')) return '미수료(교육기록 미흡)';
  if (statuses.includes('미수료')) return '미수료';
  if (statuses.includes('확인 필요')) return '확인 필요';
  return '수료';
}

function countOverall(table) {
  const c = { 수료: 0, 미수료: 0, 미흡: 0, 확인: 0 };
  for (const row of table) {
    if (row.overall === '수료') c.수료++;
    else if (row.overall === '미수료') c.미수료++;
    else if (row.overall === '미수료(교육기록 미흡)') { c.미수료++; c.미흡++; }
    else c.확인++;
  }
  return c;
}

function summaryMarkdown(c, total) {
  return [
    '| 구분 | 인원 |',
    '|---|---|',
    `| 전체 수료 | ${c.수료} |`,
    `| 미수료 | ${c.미수료} |`,
    `| └ 그중 교육기록 미흡 | ${c.미흡} |`,
    `| 확인 필요 | ${c.확인} |`,
    `| **합계** | ${c.수료 + c.미수료 + c.확인} |`,
    `| 직원 명단 인원 | ${total} |`,
  ].join('\n');
}

// ── basic: 예외 상황을 일부러 넣은 작은 판 ───────────────────

async function makeBasic() {
  const dir = path.join(ROOT, 'basic');
  await mkdir(dir, { recursive: true });
  const rand = makeRandom(11);
  const emps = makeEmployees(30, 4, rand);
  // 앞자리 0이 있는 사번도 하나 넣는다 (지워지면 안 됨)
  emps[29].empNo = '00123';

  // 정답표: 직원별 교육 칸 판정
  const answer = new Map(emps.map((e) => [e.empNo, {}]));
  const set = (empNo, course, status) => { answer.get(empNo)[course] = status; };

  // 직원 명단 — 열 제목: 사번·성명·부서명
  await saveSheet(path.join(dir, '직원명단.xlsx'), ['사번', '성명', '부서명'],
    emps.map((e) => [e.empNo, e.name, e.dept]));

  // 교육 1. 성희롱예방교육 — 기본형, 3명 수강(미수료)
  {
    const c = '성희롱예방교육';
    const notDone = new Set(['A003', 'A011', 'A027']);
    const rows = emps.map((e) => {
      const raw = notDone.has(e.empNo) ? '수강' : '수료';
      set(e.empNo, c, raw === '수료' ? '수료' : '미수료');
      return [e.empNo, e.name, raw];
    });
    await saveSheet(path.join(dir, `${c}.xlsx`), ['사번', '이름', '수료여부'], rows);
  }

  // 교육 2. 개인정보보호교육 — 다른 열 제목, 위에 안내 줄, 사번 앞뒤 공백, 이름 불일치 1명
  {
    const c = '개인정보보호교육';
    const notDone = new Set(['A008', 'A022']);
    const rows = emps.map((e) => {
      const raw = notDone.has(e.empNo) ? '수강' : '수료';
      set(e.empNo, c, raw === '수료' ? '수료' : '미수료');
      let empNo = e.empNo;
      if (e.empNo === 'A005') empNo = '  A005';
      if (e.empNo === 'A006') empNo = 'A006  ';
      const name = e.empNo === 'A010' ? '홍길동' : e.name; // 이름 불일치 (D-3)
      return [empNo, name, ` ${raw} `];
    });
    await saveSheet(path.join(dir, `${c}.xlsx`), ['직원번호', '성명', '이수여부'], rows,
      [['2026년 개인정보보호교육 이수 현황'], []]);
  }

  // 교육 3. 산업안전보건교육 — 자동 인식이 안 되는 열 제목 (열 고르기 창 확인용)
  {
    const c = '산업안전보건교육';
    const rows = emps.map((e) => {
      const raw = e.empNo === 'A014' ? '수강' : '수료';
      set(e.empNo, c, raw === '수료' ? '수료' : '미수료');
      return [e.empNo, e.name, raw];
    });
    await saveSheet(path.join(dir, `${c}.xlsx`), ['직원ID', '성명', '진행상태'], rows);
  }

  // 교육 4. 장애인인식개선교육 — 기록 없음 2명, 중복 1명, 이상한 값 1개, 명단에 없는 사번 1개
  {
    const c = '장애인인식개선교육';
    const missing = new Set(['A020', 'A021']);
    const rows = [];
    for (const e of emps) {
      if (missing.has(e.empNo)) { set(e.empNo, c, '교육 기록 없음'); continue; }
      if (e.empNo === 'A015') { // 같은 사번 2줄 → 확인 필요 (D-1)
        rows.push([e.empNo, e.name, '수강'], [e.empNo, e.name, '수료']);
        set(e.empNo, c, '확인 필요');
        continue;
      }
      if (e.empNo === 'A016') { // "수료"/"수강"이 아닌 값 → 확인 필요
        rows.push([e.empNo, e.name, '이수']);
        set(e.empNo, c, '확인 필요');
        continue;
      }
      rows.push([e.empNo, e.name, '수료']);
      set(e.empNo, c, '수료');
    }
    rows.push(['Z999', '가나다', '수료']); // 직원 명단에 없는 사번
    await saveSheet(path.join(dir, `${c}.xlsx`), ['사번', '이름', '수료여부'], rows);
  }

  // 정답표 만들기
  const courses = ['성희롱예방교육', '개인정보보호교육', '산업안전보건교육', '장애인인식개선교육'];
  const table = emps.map((e) => {
    const st = courses.map((c) => answer.get(e.empNo)[c]);
    return { ...e, st, overall: overallOf(st) };
  });
  const special = table.filter((r) => r.overall !== '수료');
  const md = [
    '# 정답표 — samples/basic',
    '',
    '> 이 폴더의 가짜 엑셀을 올렸을 때 나와야 하는 결과다. 스크립트(`scripts/make-samples.mjs`)가 자동으로 만든다.',
    '',
    '## 일부러 넣은 예외 상황',
    '',
    '| 파일 | 넣은 상황 | 기대 결과 |',
    '|---|---|---|',
    '| 직원명단 | 사번 `00123` (앞자리 0) | 0이 지워지지 않음 |',
    '| 개인정보보호교육 | 열 제목이 `직원번호`·`이수여부`, 위에 안내 줄 2줄 | 자동 인식 |',
    '| 개인정보보호교육 | 사번 `  A005`, `A006  ` (공백) | A005·A006과 합쳐짐 |',
    '| 개인정보보호교육 | 수료 여부 값 앞뒤 공백 | 정상 판정 |',
    '| 개인정보보호교육 | A010 이름이 `홍길동` | 이름 불일치 메모 |',
    '| 산업안전보건교육 | 열 제목이 `직원ID`·`진행상태` | 자동 인식 실패 → 열 고르기 |',
    '| 장애인인식개선교육 | A020·A021 기록 없음 | 교육 기록 없음 (노랑) |',
    '| 장애인인식개선교육 | A015 두 줄 | 확인 필요 (회색) |',
    '| 장애인인식개선교육 | A016 값 `이수` | 확인 필요 (회색) |',
    '| 장애인인식개선교육 | 사번 `Z999` | 맨 아래 "확인 필요" 모음 |',
    '',
    '## 요약 시트',
    '',
    summaryMarkdown(countOverall(table), emps.length),
    '',
    '## 수료가 아닌 직원',
    '',
    `| 사번 | ${courses.join(' | ')} | 전체 |`,
    `|---|${courses.map(() => '---').join('|')}|---|`,
    ...special.map((r) => `| ${r.empNo} | ${r.st.join(' | ')} | ${r.overall} |`),
    '',
    '나머지 직원은 모든 교육이 수료이고 전체도 수료다.',
    '',
  ].join('\n');
  await writeFile(path.join(dir, '정답표.md'), md, 'utf8');
  return countOverall(table);
}

// ── full: 시간 측정용 큰 판 ───────────────────────────────

const FULL_COURSES = [
  '성희롱예방교육', '개인정보보호교육', '산업안전보건교육', '장애인인식개선교육', '직장내괴롭힘예방교육',
  '정보보안교육', '청렴교육', '소방안전교육', '응급처치교육', '인권교육',
  '갑질예방교육', '윤리경영교육', '환경안전교육', '보건관리교육', '재난대응교육',
  '공정거래교육', '반부패교육', '안전보건관리교육', '산업재해예방교육', '디지털윤리교육',
];

async function makeFull() {
  const dir = path.join(ROOT, 'full');
  await mkdir(dir, { recursive: true });
  const rand = makeRandom(2026);
  const emps = makeEmployees(500, 8, rand);
  await saveSheet(path.join(dir, '직원명단.xlsx'), ['사번', '이름', '부서'],
    emps.map((e) => [e.empNo, e.name, e.dept]));

  const answer = new Map(emps.map((e) => [e.empNo, []]));
  for (const c of FULL_COURSES) {
    const rows = [];
    for (const e of emps) {
      const r = rand();
      if (r < 0.005) { answer.get(e.empNo).push('교육 기록 없음'); continue; } // 0.5% 기록 없음
      const raw = r < 0.03 ? '수강' : '수료'; // 약 2.5% 미수료
      answer.get(e.empNo).push(raw === '수료' ? '수료' : '미수료');
      rows.push([e.empNo, e.name, raw]);
    }
    await saveSheet(path.join(dir, `${c}.xlsx`), ['사번', '이름', '수료여부'], rows);
  }
  const table = emps.map((e) => ({ overall: overallOf(answer.get(e.empNo)) }));
  const md = [
    '# 정답표 — samples/full',
    '',
    '> 직원 500명 + 교육 파일 20개. 업로드부터 내려받기까지 걸리는 시간(목표 5분 이내)을 잴 때 쓴다.',
    '',
    summaryMarkdown(countOverall(table), emps.length),
    '',
  ].join('\n');
  await writeFile(path.join(dir, '정답표.md'), md, 'utf8');
  return countOverall(table);
}

const basic = await makeBasic();
const full = await makeFull();
console.log('samples/basic 완료:', basic);
console.log('samples/full 완료:', full);
