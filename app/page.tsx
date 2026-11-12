// 설계 참고: DESIGN.md §6 — 메인 페이지. 실제 화면 동작은 Workspace(브라우저 쪽 코드)에서 처리한다.

import Workspace from './components/Workspace';

export default function Home() {
  return (
    <main className="flex-1 bg-zinc-50">
      <Workspace />
    </main>
  );
}
