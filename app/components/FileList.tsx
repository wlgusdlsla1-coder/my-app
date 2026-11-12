// 설계 참고: DESIGN.md §6.3 — 올린 파일 이름, 상태(✔/⚠), 빼기·열 고르기 버튼

import type { FileEntry } from '@/lib/pipeline';

export default function FileList({
  entries,
  onRemove,
  onPickColumns,
  disabled,
}: {
  entries: readonly FileEntry[];
  onRemove: (id: string) => void;
  onPickColumns: (id: string) => void;
  disabled?: boolean;
}) {
  if (entries.length === 0) return null;

  return (
    <ul className="divide-y divide-zinc-200 rounded-lg border border-zinc-200 bg-white">
      {entries.map((e) => (
        <li key={e.id} className="flex flex-wrap items-center gap-x-3 gap-y-1 px-3 py-2 text-sm">
          <span aria-hidden className="w-5 text-center">
            {e.status === 'ready' ? '✔' : e.status === 'reading' ? '…' : '⚠'}
          </span>
          <span className="min-w-0 flex-1 break-all font-medium">{e.fileName}</span>

          {e.status === 'ready' && e.course && (
            <span className="text-zinc-500">{e.course.records.length.toLocaleString()}줄</span>
          )}
          {e.status === 'reading' && <span className="text-zinc-500">읽는 중…</span>}
          {e.status === 'needsColumns' && (
            <>
              <span className="text-amber-700">열을 골라 주세요</span>
              <button
                type="button"
                onClick={() => onPickColumns(e.id)}
                disabled={disabled}
                className="rounded border border-amber-400 bg-amber-50 px-2 py-0.5 text-amber-900 hover:bg-amber-100 disabled:opacity-50"
              >
                열 고르기
              </button>
            </>
          )}
          {e.status === 'error' && <span className="text-red-700">열 수 없는 파일</span>}

          <button
            type="button"
            onClick={() => onRemove(e.id)}
            disabled={disabled}
            className="rounded border border-zinc-300 px-2 py-0.5 text-zinc-600 hover:bg-zinc-100 disabled:opacity-50"
            aria-label={`${e.fileName} 빼기`}
          >
            빼기
          </button>
        </li>
      ))}
    </ul>
  );
}
