// 설계 참고: DESIGN.md §6.1·§6.4 — 결과 숫자, 합계 일치 여부, 내려받기 버튼

import type { Summary } from '@/lib/merge';

function Stat({ label, value, sub, tone }: { label: string; value: number; sub?: string; tone?: 'red' | 'gray' | 'green' }) {
  const color = tone === 'red' ? 'text-red-700' : tone === 'gray' ? 'text-zinc-600' : tone === 'green' ? 'text-emerald-700' : '';
  return (
    <div className="rounded-lg border border-zinc-200 bg-white px-4 py-3">
      <div className="text-xs text-zinc-500">{label}</div>
      <div className={`text-2xl font-semibold tabular-nums ${color}`}>{value.toLocaleString()}명</div>
      {sub && <div className="text-xs text-zinc-500">{sub}</div>}
    </div>
  );
}

export default function ResultSummary({
  summary,
  courseCount,
  onDownload,
  downloading,
}: {
  summary: Summary;
  courseCount: number;
  onDownload: () => void;
  downloading: boolean;
}) {
  const matched = summary.total === summary.targets;
  return (
    <div className="space-y-3">
      <div className="grid grid-cols-2 gap-3 sm:grid-cols-4">
        <Stat label="대상자" value={summary.targets} sub={`교육 ${courseCount}개 기준`} />
        <Stat label="전체 수료" value={summary.completed} tone="green" />
        <Stat label="미수료" value={summary.notCompleted} sub={`그중 교육기록 미흡 ${summary.missingRecord}명`} tone="red" />
        <Stat label="확인 필요" value={summary.needsCheck} tone="gray" />
      </div>
      <p className={`text-sm ${matched ? 'text-emerald-700' : 'font-medium text-red-700'}`}>
        {matched ? '✔ 합계가 대상자 수와 일치해요' : '⛔ 합계가 대상자 수와 달라요. 결과를 확인해 주세요.'} (합계 {summary.total.toLocaleString()}명 / 대상자{' '}
        {summary.targets.toLocaleString()}명)
      </p>
      <button
        type="button"
        onClick={onDownload}
        disabled={downloading}
        className="rounded-md bg-emerald-600 px-5 py-2.5 text-base font-medium text-white hover:bg-emerald-700 disabled:cursor-wait disabled:bg-zinc-300"
      >
        {downloading ? '엑셀 만드는 중…' : '엑셀 내려받기'}
      </button>
    </div>
  );
}
