'use client';

// 설계 참고: DESIGN.md §6.1·§6.3 (nice, FR-10) — 내려받기 전에 통합 결과를 색칠된 표로 미리 보기
// 결과 엑셀의 통합 시트와 같은 열 순서·같은 색을 쓴다. 사람이 많아도 느려지지 않게 처음에는 일부만 보여준다.

import { useState } from 'react';
import { COLORS } from '@/lib/constants';
import { deptLabel, type MergeResult } from '@/lib/merge';
import { COURSE_FILL, OVERALL_FILL, toCssColor } from '@/lib/statusColors';

/** 처음에 보여줄 사람 수, 더 보기를 누를 때마다 늘어나는 수 */
const PAGE_SIZE = 100;

const LEGEND: { color: string; label: string }[] = [
  { color: COLORS.red, label: '미수료' },
  { color: COLORS.yellow, label: '교육 기록 없음 / 미수료(교육기록 미흡)' },
  { color: COLORS.gray, label: '확인 필요' },
];

export default function ResultPreview({ merged }: { merged: MergeResult }) {
  const [limit, setLimit] = useState(PAGE_SIZE);
  const { courseNames, rows } = merged;
  const shown = rows.slice(0, limit);

  const cellStyle = (argb: string | undefined) => {
    const bg = toCssColor(argb);
    return bg ? { backgroundColor: bg } : undefined;
  };

  return (
    <div className="space-y-2">
      <div className="flex flex-wrap items-center gap-x-4 gap-y-1 text-xs text-zinc-600">
        {LEGEND.map((l) => (
          <span key={l.label} className="inline-flex items-center gap-1">
            <span className="inline-block h-3 w-3 rounded-sm border border-zinc-300" style={{ backgroundColor: toCssColor(l.color) }} aria-hidden />
            {l.label}
          </span>
        ))}
      </div>

      <div className="max-h-[28rem] overflow-auto rounded-lg border border-zinc-200 bg-white">
        <table className="w-full border-collapse text-xs">
          <caption className="sr-only">통합 시트 미리보기</caption>
          <thead className="sticky top-0 z-[1] bg-zinc-100 text-left">
            <tr>
              <th scope="col" className="whitespace-nowrap border-b border-zinc-300 px-2 py-1.5">사번</th>
              <th scope="col" className="whitespace-nowrap border-b border-zinc-300 px-2 py-1.5">이름</th>
              <th scope="col" className="whitespace-nowrap border-b border-zinc-300 px-2 py-1.5">부서</th>
              {courseNames.map((c) => (
                <th key={c} scope="col" title={c} className="min-w-[6rem] max-w-[10rem] border-b border-zinc-300 px-2 py-1.5 text-center font-medium break-keep">
                  {c}
                </th>
              ))}
              <th scope="col" className="whitespace-nowrap border-b border-zinc-300 px-2 py-1.5 text-center">전체 수료 여부</th>
            </tr>
          </thead>
          <tbody>
            {shown.map((row) => (
              <tr key={row.empNo} className="border-b border-zinc-100">
                <td className="whitespace-nowrap px-2 py-1 font-mono">{row.empNo}</td>
                <td className="whitespace-nowrap px-2 py-1">{row.name}</td>
                <td className="whitespace-nowrap px-2 py-1">{deptLabel(row.dept)}</td>
                {courseNames.map((c) => (
                  <td key={c} className="whitespace-nowrap px-2 py-1 text-center" style={cellStyle(COURSE_FILL[row.courses[c]])}>
                    {row.courses[c]}
                  </td>
                ))}
                <td className="whitespace-nowrap px-2 py-1 text-center font-semibold" style={cellStyle(OVERALL_FILL[row.overall])}>
                  {row.overall}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="flex flex-wrap items-center gap-3 text-sm text-zinc-600">
        <span>
          {rows.length.toLocaleString()}명 중 {shown.length.toLocaleString()}명을 보여주고 있어요.
        </span>
        {shown.length < rows.length && (
          <>
            <button type="button" onClick={() => setLimit((n) => n + PAGE_SIZE)} className="rounded border border-zinc-300 px-2 py-0.5 hover:bg-zinc-100">
              {PAGE_SIZE}명 더 보기
            </button>
            <button type="button" onClick={() => setLimit(rows.length)} className="rounded border border-zinc-300 px-2 py-0.5 hover:bg-zinc-100">
              전체 보기
            </button>
          </>
        )}
      </div>
    </div>
  );
}
