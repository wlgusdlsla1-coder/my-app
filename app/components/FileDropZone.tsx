'use client';

// 설계 참고: DESIGN.md §6.1·§6.3 — 끌어다 놓기·클릭으로 교육 파일 고르기 (형식·용량 검사는 상위에서)

import { useRef, useState } from 'react';

export default function FileDropZone({
  onFiles,
  disabled,
  count,
  max,
}: {
  onFiles: (files: File[]) => void;
  disabled?: boolean;
  count: number;
  max: number;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);

  const pick = (list: FileList | null) => {
    if (disabled || !list || list.length === 0) return;
    onFiles(Array.from(list));
    // 같은 파일을 빼고 다시 올릴 수 있도록 선택값을 비운다
    if (inputRef.current) inputRef.current.value = '';
  };

  return (
    <label
      onDragOver={(e) => {
        e.preventDefault();
        if (!disabled) setDragging(true);
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragging(false);
        pick(e.dataTransfer.files);
      }}
      className={`flex cursor-pointer flex-col items-center justify-center gap-1 rounded-lg border-2 border-dashed px-4 py-8 text-center transition-colors focus-within:ring-2 focus-within:ring-blue-500 focus-within:ring-offset-2 ${
        disabled
          ? 'cursor-not-allowed border-zinc-200 bg-zinc-50 text-zinc-400'
          : dragging
            ? 'border-blue-500 bg-blue-50'
            : 'border-zinc-300 bg-white hover:border-blue-400 hover:bg-blue-50/40'
      }`}
    >
      <input
        ref={inputRef}
        type="file"
        multiple
        accept=".xlsx"
        disabled={disabled}
        onChange={(e) => pick(e.target.files)}
        className="sr-only"
        aria-label="교육 파일 고르기"
      />
      <span className="text-base font-medium">여기에 끌어다 놓거나 클릭해서 고르세요 (.xlsx)</span>
      <span className="text-sm text-zinc-500">
        여러 파일을 한 번에 올릴 수 있어요 · {count} / {max}개
      </span>
    </label>
  );
}
