// 설계 참고: DESIGN.md §6.4 — 결과 엑셀을 운영자 PC로 내려받기 (서버를 거치지 않고 브라우저에서 바로 저장)

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/** 결과 파일 이름: 의무교육_수료현황_YYYYMMDD.xlsx (PC의 오늘 날짜 기준) */
export function resultFileName(date: Date = new Date()): string {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `의무교육_수료현황_${y}${m}${d}.xlsx`;
}

/** 브라우저에서 파일 저장 창을 띄운다 (브라우저 안에서만 호출) */
export function downloadBytes(bytes: Uint8Array, fileName: string): void {
  const blob = new Blob([bytes as BlobPart], { type: XLSX_MIME });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  // 클릭 직후 바로 지우면 일부 브라우저에서 저장이 취소될 수 있어 잠깐 기다린다
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}
