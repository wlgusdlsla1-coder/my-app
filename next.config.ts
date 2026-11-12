import type { NextConfig } from "next";

const isDev = process.env.NODE_ENV === "development";

// 보안 응답 헤더 (DESIGN.md §9). Next.js 가이드의 "Without Nonces" 방식을 따른다.
// 이 앱은 화면이 미리 만들어지는(정적) 페이지라 요청마다 달라지는 nonce를 쓸 수 없어서 script-src에 'unsafe-inline'을 둔다.
// 핵심은 connect-src 'self': 이 사이트의 프로그램이 다른 주소로 데이터를 보내는 연결을 막아,
// "직원 정보를 서버로 보내지 않는다"는 원칙을 브라우저가 대신 지켜 준다.
const contentSecurityPolicy = [
  "default-src 'self'",
  `script-src 'self' 'unsafe-inline'${isDev ? " 'unsafe-eval'" : ""}`,
  "style-src 'self' 'unsafe-inline'",
  "img-src 'self' data: blob:",
  "font-src 'self' data:",
  // 개발 서버의 자동 새로고침(웹소켓)은 개발 중에만 허용한다
  `connect-src 'self'${isDev ? " ws: wss:" : ""}`,
  "worker-src 'self' blob:",
  "object-src 'none'",
  "base-uri 'self'",
  "form-action 'none'",
  "frame-ancestors 'none'",
].join("; ");

const nextConfig: NextConfig = {
  /* config options here */
  cacheComponents: true,
  partialPrefetching: true,
  turbopack: {
    rules: {
      "*.css": {
        loaders: ["@tailwindcss/turbopack"],
        as: "*.css",
      },
    },
  },
  async headers() {
    return [
      {
        source: "/:path*",
        headers: [
          { key: "Content-Security-Policy", value: contentSecurityPolicy },
          // 파일 형식을 브라우저가 멋대로 추측하지 않게 한다
          { key: "X-Content-Type-Options", value: "nosniff" },
          // 다른 사이트 안에 이 화면을 끼워 넣지 못하게 한다 (오래된 브라우저용, CSP의 frame-ancestors와 같은 뜻)
          { key: "X-Frame-Options", value: "DENY" },
          // 다른 사이트로 이동할 때 이 주소를 알려 주지 않는다
          { key: "Referrer-Policy", value: "no-referrer" },
          // 카메라·마이크·위치 같은 기능은 쓰지 않는다
          { key: "Permissions-Policy", value: "camera=(), microphone=(), geolocation=(), payment=(), usb=()" },
        ],
      },
    ];
  },
};

export default nextConfig;
