'use client';

// 설계 참고: DESIGN.md §6.1·§6.2 — 한 페이지 화면: 교육 파일 올리기 → (열 고르기) → 정리하기 → 결과 → 엑셀 내려받기
// ※ 올린 파일은 이 브라우저 안에서만 읽고 처리한다. 서버로 보내거나 저장하지 않는다 (PRD 7번)

import { useEffect, useRef, useState } from 'react';
import { MAX_COURSE_FILES } from '@/lib/constants';
import { downloadBytes, resultFileName } from '@/lib/download';
import { errorMessage } from '@/lib/messages';
import { mergeCourses, summarize, type MergeResult, type Summary } from '@/lib/merge';
import type { FileEntry } from '@/lib/pipeline';
import type { ColumnMapping } from '@/lib/types';
import { checkCourseFiles, courseNameFromFile } from '@/lib/validateFiles';
import ColumnPicker from './ColumnPicker';
import FileDropZone from './FileDropZone';
import FileList from './FileList';
import Notice, { type NoticeItem, type NoticeKind } from './Notice';
import ResultPreview from './ResultPreview';
import ResultSummary from './ResultSummary';

type Result = { merged: MergeResult; summary: Summary };

let nextId = 1;
const newId = () => nextId++;

/** 한 번 양보해서 화면이 먼저 "처리 중"으로 바뀌게 한다 */
const yieldToBrowser = () => new Promise<void>((resolve) => setTimeout(resolve, 0));

export default function Workspace() {
  const [entries, setEntries] = useState<FileEntry[]>([]);
  const [notices, setNotices] = useState<NoticeItem[]>([]);
  const [pickingId, setPickingId] = useState<string | null>(null);
  const [result, setResult] = useState<Result | null>(null);
  const [processing, setProcessing] = useState(false);
  const [downloading, setDownloading] = useState(false);
  /** 내려받기가 끝났을 때 보여주는 완료 안내 */
  const [doneText, setDoneText] = useState<string | null>(null);
  /**
   * 열을 골라야 하는 파일의 원본(File). 열 고르기를 취소하면 읽어 둔 표를 버리고,
   * 다시 열 때 이 원본에서 새로 읽기 위해 둔다. (File은 디스크 파일을 가리킬 뿐 내용을 메모리에 들고 있지 않다)
   */
  const filesRef = useRef(new Map<string, File>());

  // 올리기 칸 밖에 파일을 떨어뜨리면 브라우저가 그 파일을 열면서 페이지가 이동해 목록이 사라진다.
  // 파일을 끌어오는 동작일 때만 기본 동작을 막는다. (올리기 칸 안의 드롭은 FileDropZone이 따로 처리한다)
  useEffect(() => {
    const stop = (e: DragEvent) => {
      if (e.dataTransfer?.types.includes('Files')) e.preventDefault();
    };
    window.addEventListener('dragover', stop);
    window.addEventListener('drop', stop);
    return () => {
      window.removeEventListener('dragover', stop);
      window.removeEventListener('drop', stop);
    };
  }, []);

  const resetResult = () => {
    setResult(null);
    setDoneText(null);
  };

  const reading = entries.some((e) => e.status === 'reading');
  const readyCount = entries.filter((e) => e.status === 'ready').length;
  const allReady = entries.length > 0 && entries.every((e) => e.status === 'ready');
  const busy = reading || processing || downloading;

  const addNotice = (kind: NoticeKind, text: string) =>
    setNotices((prev) => [...prev, { id: newId(), kind, text }]);

  // ── 파일 올리기 ───────────────────────────────────────────
  const handleFiles = async (files: File[]) => {
    resetResult();
    setNotices([]);

    // ① 형식·용량·개수·이름 겹침 검사 (E01·E02·E03·E06)
    const { accepted, rejected } = checkCourseFiles(
      entries.map((e) => ({ name: e.fileName, size: 0 })),
      files,
    );
    for (const { file, code } of rejected) addNotice('error', errorMessage(code, file.name));
    if (accepted.length === 0) return;

    // ② 목록에 "읽는 중" 칸을 먼저 보여준다
    const placeholders: FileEntry[] = accepted.map((file) => ({
      id: String(newId()),
      fileName: file.name,
      courseName: courseNameFromFile(file.name),
      status: 'reading',
    }));
    setEntries((prev) => [...prev, ...placeholders]);

    // ③ 하나씩 읽어서 칸을 갱신한다 (메모리를 아끼려고 순서대로)
    // 엑셀 도구는 처음 필요할 때 불러온다. 불러오지 못하면(예: 새 버전이 배포된 뒤 열려 있던 페이지)
    // "읽는 중"에 멈추지 않도록 모든 칸을 오류로 바꾸고 안내한다.
    let readCourseEntry: typeof import('@/lib/pipeline').readCourseEntry;
    try {
      ({ readCourseEntry } = await import('@/lib/pipeline'));
    } catch {
      const failed = new Map(placeholders.map((p) => [p.id, p]));
      setEntries((prev) =>
        prev.map((e) => (failed.has(e.id) ? { ...e, status: 'error', error: errorMessage('E04', e.fileName) } : e)),
      );
      addNotice('error', '화면 부품을 불러오지 못했어요. 페이지를 새로고침한 뒤 다시 시도해 주세요.');
      return;
    }
    for (let i = 0; i < accepted.length; i++) {
      const placeholder = placeholders[i];
      let entry: FileEntry;
      try {
        entry = await readCourseEntry(placeholder.id, placeholder.fileName, await accepted[i].arrayBuffer());
      } catch {
        entry = { ...placeholder, status: 'error', error: errorMessage('E04', placeholder.fileName) };
      }
      if (entry.error) addNotice(entry.status === 'error' ? 'error' : 'warn', entry.error);
      // 열을 골라야 하는 파일만 원본을 기억해 둔다 (취소 뒤 다시 열 때 새로 읽기 위해)
      if (entry.status === 'needsColumns') filesRef.current.set(entry.id, accepted[i]);
      setEntries((prev) => prev.map((e) => (e.id === entry.id ? entry : e)));
    }
  };

  const handleRemove = (id: string) => {
    filesRef.current.delete(id);
    setEntries((prev) => prev.filter((e) => e.id !== id));
    setNotices([]);
    resetResult();
  };

  // ── 열 고르기 ─────────────────────────────────────────────
  // [열 고르기]를 누르면 창을 연다. 앞서 취소해서 표를 버렸다면 원본 파일에서 새로 읽어 온다.
  const handlePickColumns = async (id: string) => {
    const target = entries.find((e) => e.id === id);
    if (!target) return;
    if (target.sheet) {
      setPickingId(id);
      return;
    }
    const file = filesRef.current.get(id);
    if (!file) {
      addNotice('error', `${target.fileName} 파일을 다시 읽을 수 없어요. 파일을 빼고 다시 올려 주세요.`);
      return;
    }
    setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, status: 'reading' } : e)));
    try {
      const { readCourseEntry } = await import('@/lib/pipeline');
      const fresh = await readCourseEntry(id, target.fileName, await file.arrayBuffer());
      setEntries((prev) => prev.map((e) => (e.id === id ? fresh : e)));
      if (fresh.status === 'needsColumns') setPickingId(id);
    } catch {
      setEntries((prev) => prev.map((e) => (e.id === id ? target : e)));
      addNotice('error', errorMessage('E04', target.fileName));
    }
  };

  // 취소하면 읽어 둔 표를 바로 버린다 (DESIGN.md §9)
  const handleCancelPicker = async () => {
    const id = pickingId;
    setPickingId(null);
    if (!id) return;
    try {
      const { dropSheet } = await import('@/lib/pipeline');
      setEntries((prev) => prev.map((e) => (e.id === id ? dropSheet(e) : e)));
    } catch {
      // 부품을 못 불러와도 창은 이미 닫혔다. 표는 파일을 빼거나 열을 고를 때 정리된다.
    }
  };

  const handleConfirmColumns = async (mapping: ColumnMapping) => {
    const target = entries.find((e) => e.id === pickingId);
    if (!target) return;
    try {
      const { applyMapping } = await import('@/lib/pipeline');
      const updated = applyMapping(target, mapping);
      setEntries((prev) => prev.map((e) => (e.id === updated.id ? updated : e)));
      filesRef.current.delete(target.id); // 열이 정해졌으니 원본 기억도 필요 없다
      setPickingId(null);
      setNotices([]);
      resetResult();
    } catch {
      setPickingId(null);
      addNotice('error', '화면 부품을 불러오지 못했어요. 페이지를 새로고침한 뒤 다시 시도해 주세요.');
    }
  };

  // ── 정리하기 ──────────────────────────────────────────────
  const handleProcess = async () => {
    setProcessing(true);
    setNotices([]);
    setDoneText(null);
    await yieldToBrowser();
    try {
      const courses = entries.flatMap((e) => (e.course ? [e.course] : []));
      const merged = mergeCourses(courses);
      const summary = summarize(merged.rows);
      setResult({ merged, summary });
      if (summary.total !== summary.targets) addNotice('error', errorMessage('E08'));
    } catch (err) {
      addNotice('error', err instanceof Error ? err.message : '정리하는 중 문제가 생겼어요.');
    } finally {
      setProcessing(false);
    }
  };

  // ── 엑셀 내려받기 ─────────────────────────────────────────
  const handleDownload = async () => {
    if (!result) return;
    setDownloading(true);
    await yieldToBrowser();
    try {
      const { buildWorkbook, workbookToBytes } = await import('@/lib/buildWorkbook');
      const fileName = resultFileName();
      downloadBytes(await workbookToBytes(buildWorkbook(result.merged)), fileName);
      setDoneText(`엑셀 파일을 내려받았어요: ${fileName} — 브라우저의 다운로드 폴더를 확인해 주세요.`);
    } catch {
      addNotice('error', '엑셀 파일을 만드는 중 문제가 생겼어요. 다시 시도해 주세요.');
    } finally {
      setDownloading(false);
    }
  };

  const picking = entries.find((e) => e.id === pickingId);

  return (
    <div className="mx-auto w-full max-w-3xl space-y-6 px-4 py-8">
      <header className="space-y-2">
        <h1 className="text-2xl font-bold">의무교육 수료관리 도우미</h1>
        <p className="rounded-md bg-zinc-100 px-3 py-2 text-sm text-zinc-700">
          🔒 올린 파일은 이 PC의 브라우저 안에서만 처리돼요. 서버로 보내거나 저장하지 않고, 페이지를 닫으면 사라져요.
        </p>
      </header>

      {notices.length > 0 && (
        <div className="space-y-2">
          {notices.map((n) => (
            <Notice key={n.id} item={n} onClose={() => setNotices((prev) => prev.filter((x) => x.id !== n.id))} />
          ))}
        </div>
      )}

      <section className="space-y-3" aria-labelledby="step1">
        <h2 id="step1" className="text-lg font-semibold">
          ① 교육별 수료 결과 올리기 <span className="text-sm font-normal text-zinc-500">(교육 플랫폼에서 받은 엑셀, 최대 {MAX_COURSE_FILES}개)</span>
        </h2>
        <FileDropZone onFiles={handleFiles} disabled={busy || entries.length >= MAX_COURSE_FILES} count={entries.length} max={MAX_COURSE_FILES} />
        <FileList entries={entries} onRemove={handleRemove} onPickColumns={handlePickColumns} disabled={busy} />

        <div className="flex flex-wrap items-center gap-3">
          <button
            type="button"
            onClick={handleProcess}
            disabled={!allReady || busy}
            className="rounded-md bg-blue-600 px-5 py-2.5 text-base font-medium text-white hover:bg-blue-700 disabled:cursor-not-allowed disabled:bg-zinc-300"
          >
            {processing ? '정리하는 중…' : '정리하기'}
          </button>
          {entries.length === 0 && <span className="text-sm text-zinc-500">파일을 올리면 누를 수 있어요.</span>}
          {entries.length > 0 && !allReady && !reading && (
            <span className="text-sm text-amber-700">⚠ 표시된 파일은 열을 고르거나 빼 주세요.</span>
          )}
          {reading && <span className="text-sm text-zinc-500">파일을 읽는 중이에요…</span>}
          {allReady && !result && <span className="text-sm text-zinc-500">{readyCount}개 파일을 합칠 준비가 됐어요.</span>}
        </div>
      </section>

      {result && (
        <section className="space-y-3" aria-labelledby="step2">
          <h2 id="step2" className="text-lg font-semibold">
            ② 결과
          </h2>
          <ResultSummary summary={result.summary} courseCount={result.merged.courseNames.length} onDownload={handleDownload} downloading={downloading} />
          {doneText && <Notice item={{ id: 0, kind: 'info', text: doneText }} onClose={() => setDoneText(null)} />}
          <h3 className="pt-2 text-base font-semibold">
            미리보기 <span className="text-sm font-normal text-zinc-500">(내려받는 엑셀의 통합 시트와 같은 내용이에요)</span>
          </h3>
          <ResultPreview merged={result.merged} />
        </section>
      )}

      {picking && <ColumnPicker key={picking.id} entry={picking} onConfirm={handleConfirmColumns} onCancel={handleCancelPicker} />}
    </div>
  );
}
