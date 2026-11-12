// 설계 참고: DESIGN.md §6.3 — 오류·주의·완료 안내 메시지 한 줄

export type NoticeKind = 'error' | 'warn' | 'info';

export type NoticeItem = { id: number; kind: NoticeKind; text: string };

const STYLE: Record<NoticeKind, string> = {
  error: 'border-red-300 bg-red-50 text-red-800',
  warn: 'border-amber-300 bg-amber-50 text-amber-900',
  info: 'border-emerald-300 bg-emerald-50 text-emerald-900',
};

const ICON: Record<NoticeKind, string> = { error: '⛔', warn: '⚠', info: '✔' };

export default function Notice({ item, onClose }: { item: NoticeItem; onClose?: () => void }) {
  return (
    <div
      role={item.kind === 'error' ? 'alert' : 'status'}
      className={`flex items-start gap-2 rounded-md border px-3 py-2 text-sm ${STYLE[item.kind]}`}
    >
      <span aria-hidden>{ICON[item.kind]}</span>
      <span className="flex-1">{item.text}</span>
      {onClose && (
        <button type="button" onClick={onClose} className="shrink-0 text-xs underline" aria-label="안내 닫기">
          닫기
        </button>
      )}
    </div>
  );
}
