'use client';

// 설계 참고: DESIGN.md §6.1 — 자동으로 못 찾은 열을 운영자가 직접 고르는 창

import { useEffect, useRef, useState } from 'react';
import { isMappingComplete } from '@/lib/detectColumns';
import { FIELD_LABELS } from '@/lib/messages';
import type { FileEntry } from '@/lib/pipeline';
import type { ColumnMapping, Field } from '@/lib/types';

const FIELDS: Field[] = ['empNo', 'name', 'dept', 'status'];

export default function ColumnPicker({
  entry,
  onConfirm,
  onCancel,
}: {
  entry: FileEntry;
  onConfirm: (mapping: ColumnMapping) => void;
  onCancel: () => void;
}) {
  const sheet = entry.sheet;
  const [mapping, setMapping] = useState<ColumnMapping>(entry.mapping ?? {});
  const dialogRef = useRef<HTMLDivElement>(null);

  // 창이 열리면 첫 칸에 초점을 주고, 닫힐 때 원래 누른 버튼으로 초점을 돌려준다
  useEffect(() => {
    const opener = document.activeElement as HTMLElement | null;
    dialogRef.current?.querySelector<HTMLElement>('select')?.focus();
    return () => opener?.focus();
  }, []);

  // Esc로 닫기, Tab은 창 안에서만 돌게 하기
  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Escape') {
      e.stopPropagation();
      onCancel();
      return;
    }
    if (e.key !== 'Tab') return;
    const items = Array.from(
      dialogRef.current?.querySelectorAll<HTMLElement>('select, button:not([disabled])') ?? [],
    );
    if (items.length === 0) return;
    const first = items[0];
    const last = items[items.length - 1];
    if (e.shiftKey && document.activeElement === first) {
      e.preventDefault();
      last.focus();
    } else if (!e.shiftKey && document.activeElement === last) {
      e.preventDefault();
      first.focus();
    }
  };

  if (!sheet) return null;

  const firstRow = sheet.rows[0] ?? [];
  const chosen = FIELDS.map((f) => mapping[f]).filter((v): v is number => v !== undefined);
  const hasDuplicate = new Set(chosen).size !== chosen.length;
  const canConfirm = isMappingComplete(mapping) && !hasDuplicate;

  return (
    <div
      ref={dialogRef}
      onKeyDown={handleKeyDown}
      className="fixed inset-0 z-10 flex items-center justify-center bg-black/40 p-4"
      role="dialog"
      aria-modal="true"
      aria-labelledby="picker-title"
    >
      <div className="max-h-full w-full max-w-lg overflow-auto rounded-lg bg-white p-5 shadow-xl">
        <h2 id="picker-title" className="mb-1 text-lg font-semibold">
          열을 골라 주세요
        </h2>
        <p className="mb-4 break-all text-sm text-zinc-600">{entry.fileName}</p>

        <div className="space-y-3">
          {FIELDS.map((field) => {
            const idx = mapping[field];
            return (
              <div key={field} className="grid grid-cols-[6rem_1fr] items-center gap-2">
                <label htmlFor={`pick-${field}`} className="text-sm font-medium">
                  {FIELD_LABELS[field]}
                </label>
                <div>
                  <select
                    id={`pick-${field}`}
                    value={idx ?? ''}
                    onChange={(e) =>
                      setMapping((m) => ({ ...m, [field]: e.target.value === '' ? undefined : Number(e.target.value) }))
                    }
                    className="w-full rounded border border-zinc-300 bg-white px-2 py-1.5 text-sm"
                  >
                    <option value="">선택하세요</option>
                    {sheet.headers.map((h, i) => (
                      <option key={i} value={i}>
                        {h || `(제목 없음) ${i + 1}번째 열`}
                      </option>
                    ))}
                  </select>
                  {idx !== undefined && (
                    <p className="mt-0.5 text-xs text-zinc-500">첫 줄 값: {firstRow[idx] || '(비어 있음)'}</p>
                  )}
                </div>
              </div>
            );
          })}
        </div>

        {hasDuplicate && <p className="mt-3 text-sm text-red-700">같은 열을 두 번 고를 수 없어요.</p>}

        <div className="mt-5 flex justify-end gap-2">
          <button type="button" onClick={onCancel} className="rounded border border-zinc-300 px-3 py-1.5 text-sm hover:bg-zinc-100">
            취소
          </button>
          <button
            type="button"
            disabled={!canConfirm}
            onClick={() => onConfirm(mapping)}
            className="rounded bg-blue-600 px-3 py-1.5 text-sm text-white hover:bg-blue-700 disabled:cursor-not-allowed disabled:bg-zinc-300"
          >
            확인
          </button>
        </div>
      </div>
    </div>
  );
}
